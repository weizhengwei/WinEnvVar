
#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#ifndef UNICODE
#define UNICODE
#endif

#ifndef _UNICODE
#define _UNICODE
#endif

#include <windows.h>
#include <windowsx.h>
#include <CommCtrl.h>
#include <tchar.h>
#include <string>
#include <memory>
#include <map>
#include <vector>
#include <fstream> // Required for wofstream
#include "resource.h"

#ifdef _MSC_VER
#pragma comment(lib, "kernel32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "comctl32.lib")
#endif

#define CLASSNAME _T("CLS_ENV_VAR")
#define WIN_TITLE _T("Environment Variable")
#define VER _T("V1.0.0")
#define IDM_REFRESH 100
#define IDM_EXPORT_TO_FILE 101
#define IDM_QUIT 102

#define IDM_OPEN_SYSENV 200

#define IDM_ABOUT 300

#define IDM_CTX_PROPER 1001
#define IDM_CTX_COPY 1002

#define ID_LISTVIEW 1001
HWND hListView;

namespace std {
	using tstring = std::basic_string<TCHAR>;
	typedef std::basic_ofstream<TCHAR> tofstream;
}

std::map<int, int> vecColWidth = {
	{0, 200},
	{1, 700}
};
std::vector<std::wstring> vecCols = {
	L"Key",
	L"Value",
};
void test();
std::map<std::tstring, std::tstring> GetEnvAsMap();
LRESULT CALLBACK WinProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
void CreateMainMenu(HWND hwnd);
void CreateListView(HWND hwnd);
void QuitApp();
void FillListView();
std::tstring GetSelectItemInfo();
std::tstring GetSelectItemInfo_single();
void CopySelectItemInfo(HWND hwnd);
void CopyTextToClipboard(HWND hwnd, const std::tstring& text);
std::tstring GetAllItemInfo();
void ExportToFile(HWND hwnd);
INT_PTR DlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam);

int APIENTRY _tWinMain(HINSTANCE hInstance, HINSTANCE hPreInstance, LPTSTR lpCmdLine, int nShowMode)
{
	LoadLibrary(_T("Riched20.dll"));

	INITCOMMONCONTROLSEX iccex;
	iccex.dwSize = sizeof(INITCOMMONCONTROLSEX);
	iccex.dwICC = ICC_LISTVIEW_CLASSES;
	InitCommonControlsEx(&iccex);
	WNDCLASS wc;
	ZeroMemory(&wc, sizeof(wc));
	wc.style = CS_HREDRAW | CS_VREDRAW;
	wc.lpfnWndProc = WinProc;
	wc.hInstance = hInstance;
	wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);
	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
	wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
	wc.lpszClassName = CLASSNAME;
	RegisterClass(&wc);

	HWND hwnd = CreateWindowEx(WS_EX_APPWINDOW,
		CLASSNAME,
		WIN_TITLE,
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		CW_USEDEFAULT,
		1000,
		800,
		NULL,
		NULL,
		hInstance,
		NULL);
	CreateMainMenu(hwnd);
	CreateListView(hwnd);
	ShowWindow(hwnd, SW_SHOWDEFAULT);
	UpdateWindow(hwnd);

	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	return (int)msg.wParam;
}

void test()
{
	LPTSTR lpEnvStrs = GetEnvironmentStrings();
	OutputDebugString(lpEnvStrs);
	FreeEnvironmentStrings(lpEnvStrs);

	DWORD dwSize = GetEnvironmentVariable(_T("PATH"), NULL, 0);
	if (dwSize > 0)
	{
		std::shared_ptr<TCHAR> szBuffer(new TCHAR[dwSize]);
		GetEnvironmentVariable(_T("PATH"), szBuffer.get(), dwSize);
		TCHAR* context = NULL;
		TCHAR* token = _tcstok_s(szBuffer.get(), L";", &context);
		while (token != NULL)
		{
			OutputDebugString(token);
			token = _tcstok_s(NULL, L";", &context);
			OutputDebugString(L"\n");
		}
		//OutputDebugString(szBuffer.get());
	}

	auto env = GetEnvAsMap();
	for (auto item : env)
	{
		OutputDebugString(item.first.c_str());
		OutputDebugString(_T("="));
		OutputDebugString(item.second.c_str());
		OutputDebugString(_T("\n"));
	}
}

std::map<std::tstring, std::tstring> GetEnvAsMap()
{
	std::map<std::tstring, std::tstring> env;
	auto free = [](TCHAR* p) { FreeEnvironmentStrings(p); };
	const auto envBlock = std::unique_ptr<TCHAR, decltype(free)>{
		GetEnvironmentStrings(),
		free
	};

	for (auto ch = envBlock.get(); *ch != L'\0'; ch++)
	{
		std::tstring key;
		std::tstring value;
		for (; *ch != L'='; ch++)
			key += *ch;
		ch++;
		for (; *ch != L'\0'; ch++)
			value += *ch;
		env[key] = value;
	}
	return env;
}

LRESULT CALLBACK WinProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
	case WM_CREATE:
	{
		break;
	}
	case WM_SIZE:
	{
		RECT rcClient;
		GetClientRect(hwnd, &rcClient);
		InflateRect(&rcClient, -15, -15);
		SetWindowPos(hListView,
			NULL,
			rcClient.left,
			rcClient.top,
			rcClient.right - rcClient.left,
			rcClient.bottom - rcClient.top,
			//SWP_SHOWWINDOW
			SWP_NOACTIVATE | SWP_NOZORDER
		);
		break;
	}
	case WM_COMMAND:
	{
		if (LOWORD(wParam) == IDM_REFRESH)
		{
			ListView_DeleteAllItems(hListView);
			UpdateWindow(hListView);

			FillListView();
		}
		else if (LOWORD(wParam) == IDM_EXPORT_TO_FILE)
		{
			ExportToFile(hwnd);
		}
		else if (LOWORD(wParam) == IDM_QUIT)
		{
			QuitApp();
		}
		else if (LOWORD(wParam) == IDM_OPEN_SYSENV)
		{
			SHELLEXECUTEINFO ShExecInfo;
			ShExecInfo.cbSize = sizeof(SHELLEXECUTEINFO);
			ShExecInfo.fMask = NULL;
			ShExecInfo.hwnd = NULL;
			ShExecInfo.lpVerb = NULL;
			ShExecInfo.lpFile = _T("rundll32");
			ShExecInfo.lpParameters = _T("sysdm.cpl,EditEnvironmentVariables");
			ShExecInfo.lpDirectory = NULL;
			ShExecInfo.nShow = SW_SHOWDEFAULT;
			ShExecInfo.hInstApp = NULL;

			ShellExecuteEx(&ShExecInfo);
		}
		else if (LOWORD(wParam) == IDM_ABOUT)
		{
			TCHAR szTitle[256] = { 0 };
			_stprintf_s(szTitle, ARRAYSIZE(szTitle), _T("%s %s\n\nCopyright(C) 2025 hostzhengwei@gmail.com"), WIN_TITLE, VER);
			MessageBox(hwnd, szTitle, WIN_TITLE, MB_OK);
		}
		else if (LOWORD(wParam) == IDM_CTX_PROPER)
		{
			DialogBox(GetModuleHandle(NULL), MAKEINTRESOURCE(IDD_DIALOG1), hwnd, DlgProc);
		}
		else if (LOWORD(wParam) == IDM_CTX_COPY)
		{
			CopySelectItemInfo(hwnd);
		}
		break;
	}
	case WM_CONTEXTMENU:
	{
		POINT pt;
		pt.x = GET_X_LPARAM(lParam);
		pt.y = GET_Y_LPARAM(lParam);
		ScreenToClient(hListView, &pt);
		LVHITTESTINFO hitTestInfo;
		hitTestInfo.pt = pt;
		ListView_HitTest(hListView, &hitTestInfo);
		if (hitTestInfo.iItem != -1) {
			UINT uCount = ListView_GetSelectedCount(hListView);
			HMENU hMenu = CreatePopupMenu();
			if (uCount == 1)
			{
				AppendMenu(hMenu, MF_STRING, IDM_CTX_PROPER, L"Property");
				AppendMenu(hMenu, MF_SEPARATOR, 0, L"");
				AppendMenu(hMenu, MF_STRING, IDM_CTX_COPY, L"Copy");
			}
			else
			{
				AppendMenu(hMenu, MF_STRING, IDM_CTX_COPY, L"Copy");
			}


			// Determine the menu position
			pt.x = GET_X_LPARAM(lParam);
			pt.y = GET_Y_LPARAM(lParam);
			UINT uFlags = TPM_RIGHTBUTTON; // Use this flag to handle right clicks
			TrackPopupMenuEx(hMenu, uFlags, pt.x, pt.y, hwnd, NULL);
			DestroyMenu(hMenu);
		}
		break;
	}
	case WM_DESTROY:
	{
		QuitApp();
		break;
	}
	default:
		return DefWindowProc(hwnd, msg, wParam, lParam);
	}
	return 0;
}

void QuitApp()
{
	PostQuitMessage(0);
}

void CreateMainMenu(HWND hwnd)
{
	HMENU hMainMenu = CreateMenu();
	HMENU hFileMenu = CreatePopupMenu();
	AppendMenu(hFileMenu, MF_STRING, IDM_REFRESH, _T("Refresh"));
	AppendMenu(hFileMenu, MF_STRING, IDM_EXPORT_TO_FILE, _T("ExportToFile"));
	AppendMenu(hFileMenu, MF_SEPARATOR, NULL, NULL);
	AppendMenu(hFileMenu, MF_STRING, IDM_QUIT, _T("Quit"));


	HMENU hSettingMenu = CreatePopupMenu();
	AppendMenu(hSettingMenu, MF_STRING, IDM_OPEN_SYSENV, _T("OpenSysEnv"));

	HMENU hHelpMenu = CreatePopupMenu();
	AppendMenu(hHelpMenu, MF_STRING, IDM_ABOUT, _T("About"));

	AppendMenu(hMainMenu, MF_POPUP, (UINT_PTR)hFileMenu, _T("File"));
	AppendMenu(hMainMenu, MF_POPUP, (UINT_PTR)hSettingMenu, _T("Setting"));
	AppendMenu(hMainMenu, MF_POPUP, (UINT_PTR)hHelpMenu, _T("Help"));

	SetMenu(hwnd, hMainMenu);
}

void CreateListView(HWND hwnd)
{
	RECT rc;
	GetClientRect(hwnd, &rc);
	InflateRect(&rc, -15, -15);
	hListView = CreateWindowEx(WS_EX_CLIENTEDGE,
		WC_LISTVIEW,
		NULL,
		WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_NOSORTHEADER,
		rc.left,
		rc.top,
		rc.right - rc.left,
		rc.bottom - rc.top,
		hwnd,
		(HMENU)ID_LISTVIEW,
		GetModuleHandle(NULL),
		NULL);
	ListView_SetTextColor(hListView, RGB(10, 10, 160));
	ListView_SetExtendedListViewStyle(hListView, LVS_EX_FULLROWSELECT);


	TCHAR szText[32] = {};
	LVCOLUMN lvCol;
	lvCol.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH;
	lvCol.fmt = LVCFMT_LEFT;
	lvCol.cx = 150;
	lvCol.pszText = szText;

	for (int iCol = 0; iCol < vecCols.size(); iCol++)
	{
		lvCol.cx = vecColWidth.at(iCol);
		_stprintf_s(szText, ARRAYSIZE(szText), _T("%s"), vecCols[iCol].c_str());
		ListView_InsertColumn(hListView, iCol, &lvCol);
	}

	FillListView();
}

void FillListView()
{
	auto env = GetEnvAsMap();
	LVITEM lvItem = {};
	//TCHAR szText[4096] = {};
	auto item = env.begin();
	for(lvItem.iItem = 0; lvItem.iItem < env.size(); lvItem.iItem++)
	{
		lvItem.iSubItem = 0;
		std::vector<TCHAR> szText(item->first.length() +1);
		_stprintf_s(szText.data(), szText.size(), _T("%s"), item->first.c_str());
		lvItem.mask = LVIF_TEXT;
		lvItem.pszText = szText.data();
		ListView_InsertItem(hListView, &lvItem);
		
		szText.resize(item->second.length() + 1);
		lvItem.iSubItem = 1;
		_stprintf_s(szText.data(), szText.size(), _T("%s"), item->second.c_str());
		lvItem.mask = LVIF_TEXT;
		lvItem.pszText = szText.data();
		ListView_SetItem(hListView, &lvItem);
		//_stprintf_s(szText, ARRAYSIZE(szText), _T("%s"), item->first.c_str());
		//lvItem.mask = LVIF_TEXT;
		//lvItem.pszText = szText;
		//ListView_InsertItem(hListView, &lvItem);

		//lvItem.iSubItem = 1;
		//_stprintf_s(szText, ARRAYSIZE(szText), _T("%s"), item->second.c_str());
		//lvItem.mask = LVIF_TEXT;
		//lvItem.pszText = szText;
		//ListView_SetItem(hListView, &lvItem);

		item++;
	}
}

std::tstring GetSelectItemInfo()
{
	UINT itemid = -1;
	itemid = ListView_GetNextItem(hListView, itemid, LVNI_SELECTED);
	std::tstring strInfo;
	while (itemid != -1)
	{
		for (int i = 0; i < 2; i++)
		{
			TCHAR szText[4096] = {};
			ListView_GetItemText(hListView, itemid, i, szText, ARRAYSIZE(szText));
			strInfo += vecCols[i];
			strInfo += _T(" : ");
			strInfo += szText;
			strInfo += _T("\n");
		}
		strInfo += _T("\n");
		itemid = ListView_GetNextItem(hListView, itemid, LVNI_SELECTED);
	}
	return strInfo;
}

std::tstring GetSelectItemInfo_single()
{
	UINT itemid = -1;
	itemid = ListView_GetNextItem(hListView, itemid, LVNI_SELECTED);
	std::tstring strInfo;
	while (itemid != -1)
	{
		for (int i = 0; i < 2; i++)
		{
			TCHAR szText[4096] = {};
			ListView_GetItemText(hListView, itemid, i, szText, ARRAYSIZE(szText));
			strInfo += vecCols[i];
			strInfo += _T(" : ");
			
			if (_tcsstr(szText, _T(";")))
			{
				TCHAR* context = NULL;
				TCHAR* token = _tcstok_s(szText, L";", &context);
				while (token != NULL)
				{
					strInfo += token;
					token = _tcstok_s(NULL, L";", &context);
					strInfo += _T("\n");
				}
			}
			else
			{
				strInfo += szText;
			}
			strInfo += _T("\n");
		}
		strInfo += _T("\n");
		itemid = ListView_GetNextItem(hListView, itemid, LVNI_SELECTED);
	}
	return strInfo;
}

void CopySelectItemInfo(HWND hwnd)
{
	auto strInfo = GetSelectItemInfo();
	CopyTextToClipboard(hwnd, strInfo);
}

void CopyTextToClipboard(HWND hwnd, const std::tstring& text)
{
	BOOL bCopied = FALSE;
	if (OpenClipboard(NULL))
	{
		if (EmptyClipboard())
		{
			size_t len = text.length();
			HGLOBAL hData = GlobalAlloc(GMEM_MOVEABLE | GMEM_DDESHARE, (len + 1) * sizeof(TCHAR));
			if (hData != NULL)
			{
				TCHAR* pszData = static_cast<TCHAR*>(GlobalLock(hData));
				if (pszData != NULL)
				{
					//CopyMemory(pszData, text.c_str(), len*sizeof(TCHAR));
					_tcscpy_s(pszData, len + 1, text.c_str());
					GlobalUnlock(hData);
					if (SetClipboardData(CF_UNICODETEXT, hData) != NULL)
					{
						bCopied = TRUE;;
					}
				}
			}
		}
		CloseClipboard();
	}
}

std::tstring GetAllItemInfo()
{
	int count = ListView_GetItemCount(hListView);
	std::tstring strInfo;
	for (int iCol = 0; iCol < count; iCol++)
	{
		for (int i = 0; i < 2; i++)
		{
			TCHAR szText[4096] = {};
			ListView_GetItemText(hListView, iCol, i, szText, ARRAYSIZE(szText));
			//strInfo += vecCols[i];
			//strInfo += _T(" : ");
			strInfo += szText;
			strInfo += i==0? _T("=") : _T("\n");
		}
	}
	return strInfo;
}

void ExportToFile(HWND hwnd)
{
	OPENFILENAME ofn;
	TCHAR szFileName[MAX_PATH] = {};
	ZeroMemory(&ofn, sizeof(ofn));
	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = hwnd;
	ofn.lpstrFile = szFileName;
	ofn.nMaxFile = sizeof(szFileName);
	ofn.lpstrFilter = _T("Text Files\0*.txt\0All Files\0*.*\0");
	ofn.nFilterIndex = 1;
	ofn.Flags = OFN_PATHMUSTEXIST | OFN_EXPLORER;
	if (GetSaveFileName(&ofn) == TRUE)
	{
		std::tstring strFileName(szFileName);
		strFileName += _T(".txt");
		auto strInfo = GetAllItemInfo();
		std::tofstream output_file(strFileName, std::ios::out);

		if (output_file.is_open())
		{
			output_file << strInfo; // Write the string to the file
			output_file.close();    // Close the file
		}
	}
}

INT_PTR DlgProc(HWND hDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
		case WM_INITDIALOG:
		{
			HWND hEdit = GetDlgItem(hDlg, IDC_RICHEDIT21);
			std::tstring strInfo = GetSelectItemInfo_single();
			SetWindowText(hEdit, strInfo.c_str());
			return (INT_PTR)TRUE;
		}
		case WM_COMMAND:
		{
			switch (LOWORD(wParam)) {
			case IDOK:
			case IDCANCEL:
				EndDialog(hDlg, LOWORD(wParam)); // Close the dialog
				return (INT_PTR)TRUE;
			}
			break;
		}
	}
	return (INT_PTR)FALSE;
}