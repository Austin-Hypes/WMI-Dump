#include <iostream>
#include <windows.h>
#include <comdef.h>
#include <wbemidl.h>
#include <vector>
#include <fstream>
#pragma comment(lib, "wbemuuid")
#pragma comment(lib, "comsupp")

void queryClass(IWbemServices* pSvc, const wchar_t* className, std::wostream& out) {
	out << L"<h2>" << className << L"</h2>\n";

	IEnumWbemClassObject* pEnum = nullptr;
	std::wstring query = L"SELECT * FROM ";
	query += className;
	HRESULT hr = pSvc->ExecQuery(
		bstr_t("WQL"),
		bstr_t(query.c_str()),
		WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
		NULL,
		&pEnum);
	if (FAILED(hr)) {
		out << L"Failed to query " << className << L" (HRESULT=" << std::hex << hr << L")\n";
		return;
	}

	IWbemClassObject* pObj = nullptr;
	ULONG uReturn = 0;

	while (pEnum->Next(WBEM_INFINITE, 1, &pObj, &uReturn) == S_OK) {
		out << L"<table>\n";
		SAFEARRAY* names = NULL;
		hr = pObj->GetNames(NULL, WBEM_FLAG_ALWAYS, NULL, &names);
		if (FAILED(hr)) {
			pObj->Release();
			continue;
		}

		LONG lBound = 0, uBound = 0;
		SafeArrayGetLBound(names, 1, &lBound);
		SafeArrayGetUBound(names, 1, &uBound);

		for (LONG i = 0; i <= uBound; ++i) {
			BSTR propName;
			SafeArrayGetElement(names, &i, &propName);
			VARIANT vt;
			VariantInit(&vt);

			if (SUCCEEDED(pObj->Get(propName, 0, &vt, NULL, NULL))) {
				out << L"<tr><td>" << propName << L"</td><td>";
				switch (vt.vt) {
				case VT_BSTR:
					out << vt.bstrVal;
					break;
				case VT_I4:
				case VT_INT:
					out << vt.intVal;
					break;
				case VT_UI4:
					out << vt.uintVal;
					break;
				case VT_BOOL:
					out << (vt.boolVal ? L"True" : L"False");
					break;
				case VT_NULL:
				case VT_EMPTY:
					out << L"(null)";
					break;
				default:
					out << L"(vt=" << vt.vt << L")";
					break;
				}

				out << L"</td></tr>\n";
				VariantClear(&vt);
			}

			SysFreeString(propName);
		}
		SafeArrayDestroy(names);
		out << "</tables>\n";
		pObj->Release();
	}
	pEnum->Release();
}

void wmiQuery(std::wostream& out) {
	HRESULT hr = CoInitializeEx(0, COINIT_MULTITHREADED);
	if (FAILED(hr)) {
		out << L"CoInitialize failed\n";
		return;
	}

	hr = CoInitializeSecurity(
		NULL, -1, NULL, NULL,
		RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE,
		NULL, EOAC_NONE, NULL);
	if (FAILED(hr)) {
		out << L"CoInitializeSecurity Failed\n";
		CoUninitialize();
		return;
	}

	IWbemLocator* pLoc = nullptr;
	hr = CoCreateInstance(
		CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER,
		IID_IWbemLocator, (LPVOID*)&pLoc);
	if (FAILED(hr)) {
		out << L"CLSID_WbemLocator failed\n";
		CoUninitialize();
		return;
	}
	IWbemServices* pSvc = nullptr;
	hr = pLoc->ConnectServer(
		_bstr_t(L"ROOT\\CIMV2"), NULL, NULL, 0,
		NULL, 0, 0, &pSvc);
	if (FAILED(hr)) {
		out << L"Connect to server failed\n";
		pLoc->Release();
		CoUninitialize();
		return;
	}

	hr = CoSetProxyBlanket(
		pSvc, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, NULL,
		RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE,
		NULL, EOAC_NONE);
	if (FAILED(hr)) {
		out << "CoSetProxyBlanket Failed\n";
		pSvc->Release();
		pLoc->Release();
		CoUninitialize();
		return;
	}

	std::vector<const wchar_t*> classesToQuery = {
		L"Win32_ComputerSystem",
		L"Win32_OperatingSystem",
		L"Win32_Processor",
		L"Win32_PhysicalMemory",
		L"Win32_LogicalDisk",
		L"Win32_BIOS",
		L"Win32_VideoController",
		L"Win32_NetworkAdapterConfiguration",
		L"Win32_UserAccount",
		L"Win32_Service",
		L"Win32_StartupCommand",
		L"Win32_Process"
	};

	for (const auto& cls : classesToQuery) {
		queryClass(pSvc, cls, out);
	}

	pSvc->Release();
	pLoc->Release();
	CoUninitialize();

}

int wmain()
{
	std::wofstream outFile(L"WMI-Info-Dump.html");
	if (!outFile.is_open()) {
		std::wcout << L"Failed to open output file\n";
		return 1;
	}

	outFile << L"<!DOCTYPE html><html><head><meta charset='UTF-8'><title>WMI Dump</title>\n";
	outFile << L"<style>\n"
		L"	body { font-family: Consolas, monospace; background-color: #111; color: #eee; padding: 2em; line-height: 1.5; }\n"
		L"	h1 { color: #6cf; text-align: center; margin-bottom: 2em; }\n"
		L"	h2 { color: #4fc3f7; border-bottom: 1px solid #333; margin-top: 40px; padding-bottom: 5px; }\n"
		L"	table { width: 100%; border-collapse: collapse; margin-bottom: 2em; background-color: #1a1a1a; }\n"
		L"	th, td { padding: 6px 10px; border: 1px solid #333; vertical-align: top; word-break: break-word; }\n"
		L"	tr:nth-child(even) { background-color: #222; }\n"
		L"	tr:nth-child(odd) { background-color: #1a1a1a; }\n"
		L"	td:first-child { font-weight: bold; color: #90caf9; width: 30%; }\n"
		L"	td:last-child { color: #fff; }\n"
		L"	::selection { background: #444; }\n"
		L"	a { color: #03a9f4; }\n"
		L"</style></head><body>\n";
	outFile << L"<h1>WMI Info Dump</h1>\n";
	wmiQuery(outFile);
	outFile << L"</body></html>\n";

	std::wcout << L"WMI Dump complete.\n";
	return 0;
}