// grande_latte.cpp : Defines the entry point for the application.
//

#include "stdafx.h"
#include "grande_latte.h"

#include <CommDlg.h> // Open file dialog
#include <Msiquery.h>
#include <map>
#include <string>

using namespace std;

#define MAX_LOADSTRING 100

// List box
#define IDC_LIST 201 
#define CHECKWIDTH 16 
LRESULT OldListProc; 
LRESULT CALLBACK ListProc (HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam); 
static HWND hWndListBox = NULL; 
//

// Languages List
typedef struct pair<int, int> IndexStatusPair;
typedef map<wstring, IndexStatusPair, less<wstring> > LanguagesMap;
LanguagesMap langs;
LanguagesMap::iterator langs_iter;


// Global Variables:
HINSTANCE hInst;								// current instance
TCHAR szTitle[MAX_LOADSTRING];					// The title bar text
TCHAR szWindowClass[MAX_LOADSTRING];			// the main window class name

// Forward declarations of functions included in this code module:
ATOM				MyRegisterClass(HINSTANCE hInstance);
BOOL				InitInstance(HINSTANCE, int);
LRESULT CALLBACK	WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK	About(HWND, UINT, WPARAM, LPARAM);


// MSI
PMSIHANDLE hDatabase = 0;
PMSIHANDLE hView = 0;
PMSIHANDLE hRecord = 0;

void selectAll(BOOL bValue)
{
	for(langs_iter = langs.begin(); langs_iter != langs.end(); langs_iter++)
		langs_iter->second.second = bValue ? 1 : 0;

}

UINT process_error(UINT uiReturn)
{
	PMSIHANDLE hLastErrorRec = MsiGetLastErrorRecord();
	if (!hLastErrorRec) return uiReturn;

	// First we get the buffer size
	DWORD cchExtendedError = 0;
	UINT uiStatus = MsiFormatRecord(NULL, hLastErrorRec, TEXT(""), &cchExtendedError);

	if (ERROR_MORE_DATA != uiStatus) return uiReturn;
    // returned size does not include null terminator.
    cchExtendedError++;

	TCHAR* szExtendedError = new TCHAR[cchExtendedError];
	if (!szExtendedError) return uiReturn;

	uiStatus = MsiFormatRecord(NULL,
							   hLastErrorRec,
                               szExtendedError,
                               &cchExtendedError);

    if (ERROR_SUCCESS == uiStatus)
		MessageBox(NULL, szExtendedError, TEXT("Error"), MB_ICONERROR);

	delete [] szExtendedError;
    szExtendedError = NULL;
	return uiReturn;
}

////// LIST BOX
LRESULT CALLBACK ListProc (HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM 
	lParam) 
{ 
	switch (uMsg) 
	{ 
	case WM_LBUTTONUP : 
		{ 
			POINT pt; 
			int nXPos = LOWORD(lParam); 
			int nYPos = HIWORD(lParam); 
			pt.x = nXPos; 
			pt.y = nYPos; 
			int nIndex = LOWORD(SendMessage(hWnd, LB_ITEMFROMPOINT, 0, (LPARAM)
				MAKELPARAM(nXPos, nYPos))); 
			RECT rect; 
			SendMessage(hWnd, LB_GETITEMRECT, nIndex, (LPARAM)&rect);
			rect.left +=2; 
			rect.right = rect.left + CHECKWIDTH; 
			if (PtInRect(&rect, pt)) 
			{ 
				int nData = SendMessage(hWnd, LB_GETITEMDATA, nIndex, 0);
				SendMessage(hWnd, LB_SETITEMDATA, nIndex, !nData);
				RedrawWindow(hWnd, NULL, 0, RDW_ERASE |RDW_INVALIDATE | RDW_FRAME |
					RDW_UPDATENOW); 
			} 
		} 
		break; 
	} 
	return(CallWindowProc((WNDPROC)OldListProc, hWnd, uMsg, wParam, lParam));
} 
/////////////////////

int APIENTRY _tWinMain(HINSTANCE hInstance,
	HINSTANCE hPrevInstance,
	LPTSTR    lpCmdLine,
	int       nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

	// TODO: Place code here.
	MSG msg;
	HACCEL hAccelTable;

	// Initialize global strings
	LoadString(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadString(hInstance, IDC_GRANDE_LATTE, szWindowClass, MAX_LOADSTRING);
	MyRegisterClass(hInstance);

	// Perform application initialization:
	if (!InitInstance (hInstance, nCmdShow))
	{
		return FALSE;
	}

	hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_GRANDE_LATTE));

	// Main message loop:
	while (GetMessage(&msg, NULL, 0, 0))
	{
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	return (int) msg.wParam;
}



//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
//  COMMENTS:
//
//    This function and its usage are only necessary if you want this code
//    to be compatible with Win32 systems prior to the 'RegisterClassEx'
//    function that was added to Windows 95. It is important to call this function
//    so that the application will get 'well formed' small icons associated
//    with it.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEX wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style			= CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc	= WndProc;
	wcex.cbClsExtra		= 0;
	wcex.cbWndExtra		= 0;
	wcex.hInstance		= hInstance;
	wcex.hIcon			= LoadIcon(hInstance, MAKEINTRESOURCE(IDI_GRANDE_LATTE));
	wcex.hCursor		= LoadCursor(NULL, IDC_ARROW);
	wcex.hbrBackground	= (HBRUSH)(COLOR_WINDOW+1);
	wcex.lpszMenuName	= MAKEINTRESOURCE(IDC_GRANDE_LATTE);
	wcex.lpszClassName	= szWindowClass;
	wcex.hIconSm		= LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_GRANDE_LATTE));

	return RegisterClassEx(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	HWND hWnd;

	hInst = hInstance; // Store instance handle in our global variable

	hWnd = CreateWindow(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, NULL, NULL, hInstance, NULL);

	if (!hWnd)
	{
		return FALSE;
	}

	ShowWindow(hWnd, nCmdShow);
	UpdateWindow(hWnd);

	return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE:  Processes messages for the main window.
//
//  WM_COMMAND	- process the application menu
//  WM_PAINT	- Paint the main window
//  WM_DESTROY	- post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	int wmId, wmEvent;
	PAINTSTRUCT ps;
	HDC hdc;

	switch (message)
	{
	case WM_CREATE:
		// List box
		RECT rcClient;
		GetClientRect(hWnd, &rcClient);

		hWndListBox = CreateWindow(L"LISTBOX", L"List Box Name", 
			WS_BORDER | WS_CHILD | WS_VISIBLE | WS_HSCROLL| WS_VSCROLL 
			|WS_TABSTOP |LBS_OWNERDRAWFIXED | LBS_HASSTRINGS , 
			0, 0, rcClient.right, rcClient.bottom, hWnd, (HMENU)IDC_LIST, hInst, NULL ); 

		OldListProc = SetWindowLong(hWndListBox, GWL_WNDPROC, (LONG) (WNDPROC) ListProc);
		break;
	case WM_COMMAND:
		wmId    = LOWORD(wParam);
		wmEvent = HIWORD(wParam);
		// Parse the menu selections:
		switch (wmId)
		{
		case IDM_FILE_OPEN:
			OPENFILENAME ofn;			// common dialog box structure
			wchar_t szFileName[260];    // buffer for file name

			// Initialize OPENFILENAME
			ZeroMemory(&ofn, sizeof(ofn));
			ofn.lStructSize = sizeof(ofn);
			ofn.hwndOwner = hWnd;
			ofn.lpstrFile = szFileName;
			// Set lpstrFile[0] to '\0' so that GetOpenFileName does not 
			// use the contents of szFile to initialize itself.
			ofn.lpstrFile[0] = '\0';
			ofn.nMaxFile = sizeof(szFileName);
			ofn.lpstrFilter = L"Windows Installer\0*.MSI\0All\0*.*\0";
			ofn.nFilterIndex = 1;
			ofn.lpstrFileTitle = NULL;
			ofn.nMaxFileTitle = 0;
			ofn.lpstrInitialDir = NULL;
			ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

			// Display the Open dialog box. 
			if (GetOpenFileName(&ofn)==TRUE) 
			{
				UINT uiReturn = MsiOpenDatabase(szFileName, MSIDBOPEN_TRANSACT, &hDatabase);
				if (ERROR_SUCCESS != uiReturn)
				{
					// process error
					return process_error(uiReturn);
				}

				uiReturn = MsiDatabaseOpenView(hDatabase, TEXT("SELECT * from Property"), &hView);

				if (ERROR_SUCCESS != uiReturn)
				{
					// process error
					return process_error(uiReturn);
				}

				uiReturn = MsiViewExecute(hView,NULL);

				if (ERROR_SUCCESS != uiReturn)
				{
					// process error
					return process_error(uiReturn);
				}

				while (MsiViewFetch(hView,&hRecord) == ERROR_SUCCESS)
				{
					// Get buffer size
					DWORD dwLength = 0;
					TCHAR* szSizeBuf = new TCHAR[1];
					uiReturn = MsiRecordGetString(hRecord, 1, szSizeBuf, &dwLength);
					delete[] szSizeBuf;
					if (ERROR_MORE_DATA == uiReturn)
					{
						dwLength++;
						TCHAR* szTemp = new TCHAR[dwLength];
						uiReturn = MsiRecordGetString(hRecord, 1, szTemp, &dwLength);
						if (ERROR_SUCCESS == uiReturn)
						{
							if ((wcslen(szTemp) == 6 || wcslen(szTemp) == 7) && wcsncmp(szTemp, L"IS", 2) == 0)
							{
								// Also get the number
								int nStatus = MsiRecordGetInteger(hRecord, 2);
								
								// Get the language name
								int number = 0;
								if(swscanf_s(szTemp, L"IS%5d", &number))
								{
									DWORD dwSize = 255;
									WCHAR* szLanguageName = new WCHAR[dwSize];
									memset(szLanguageName, 0, sizeof(szLanguageName)/sizeof(WCHAR));

									switch(number)
									{
									default:
										GetLocaleInfo(number, LOCALE_SENGLANGUAGE, szLanguageName, dwSize);
									}
									if (wcslen(szLanguageName) == 0)
									{
										MessageBox(NULL, szTemp, TEXT("Unknown Language"), MB_ICONWARNING);
										wcscpy_s(szLanguageName, dwSize, L"(Unknown Language)");
									}
									int nPos = SendMessage(hWndListBox, LB_ADDSTRING, 0, (LPARAM)szLanguageName);
									SendMessage(hWndListBox, LB_SETITEMDATA, nPos, nStatus);

									// Add everything to the languages map
									IndexStatusPair p;
									p.first = nPos;
									p.second = nStatus;
									langs.insert(LanguagesMap::value_type(szTemp,p));
								}
							}
						}
					}
				}

				// If we have elements, enable the menus
				if (SendMessage(hWndListBox, LB_GETCOUNT, 0, 0))
				{
					EnableMenuItem(GetMenu(hWnd), IDM_FILE_SAVE, MF_ENABLED);
					EnableMenuItem(GetMenu(hWnd), IDM_EDIT_SELECTALL, MF_ENABLED);
					EnableMenuItem(GetMenu(hWnd), IDM_EDIT_SELECTNONE,MF_ENABLED);
				}
			}

			break;
		case IDM_FILE_SAVE:
			break;
		case IDM_EDIT_SELECTALL:
			selectAll(TRUE);
			SendMessage(hWndListBox, LB_SETITEMDATA, -1, 1);
			RedrawWindow(hWndListBox, NULL, 0, RDW_ERASE | RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW);
			break;
		case IDM_EDIT_SELECTNONE:
			selectAll(FALSE);
			SendMessage(hWndListBox, LB_SETITEMDATA, -1, 0);
			RedrawWindow(hWndListBox, NULL, 0, RDW_ERASE | RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW);
			break;
		case IDM_ABOUT:
			DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
			break;
		case IDM_EXIT:
			if(hRecord) MsiCloseHandle(hRecord);
			if(hView) MsiCloseHandle(hView);
			if(hDatabase) MsiCloseHandle(hDatabase);
			DestroyWindow(hWnd);
			break;
		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}
		break;
	case WM_MEASUREITEM: 
		{ 
			LPMEASUREITEMSTRUCT lpmis = (LPMEASUREITEMSTRUCT)lParam; 
			TEXTMETRIC tm; 
			HDC hDC = GetDC(hWnd); 
			GetTextMetrics(hDC, &tm); 
			lpmis->itemHeight = tm.tmHeight + 2; 
			ReleaseDC(hWnd, hDC); 
			return TRUE; 
		} 
		break; 

	case WM_DRAWITEM: 
		{ 
			HBRUSH hBrush; 
			wchar_t sBuffer[256]; 
			COLORREF nOldTextColor; 
			if (wParam == IDC_LIST) 
			{ 
				LPDRAWITEMSTRUCT lpdis = (LPDRAWITEMSTRUCT)lParam; 
				if (lpdis->itemID == -1) 
				{ 
					DrawFocusRect(lpdis->hDC, (LPRECT)&lpdis->rcItem); 
					return (TRUE); 
				} 
				if (lpdis->itemAction & ODA_DRAWENTIRE || lpdis->itemAction & ODA_SELECT)
				{ 
					RECT rect; 
					CopyRect(&rect, &lpdis->rcItem); 
					rect.left+=20; 
					SendMessage(lpdis->hwndItem, LB_GETTEXT, lpdis->itemID, (LPARAM)sBuffer);
					PatBlt(lpdis->hDC, lpdis->rcItem.left, lpdis->rcItem.top,
						lpdis->rcItem.right - lpdis->rcItem.left, lpdis->rcItem.bottom - 
						lpdis->rcItem.top, WHITENESS); 
					nOldTextColor = -1; 
					if (lpdis->itemState & ODS_SELECTED) 
					{ 
						hBrush = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT));
						FillRect(lpdis->hDC, (LPRECT)&lpdis->rcItem, hBrush);
						DeleteObject(hBrush); 
						SetBkMode(lpdis->hDC, TRANSPARENT); 
						nOldTextColor = SetTextColor(lpdis->hDC, 
							GetSysColor(COLOR_HIGHLIGHTTEXT)); 
					} 
					else 
					{ 
						hBrush = (HBRUSH)GetStockObject(WHITE_BRUSH); 
						FillRect(lpdis->hDC, (LPRECT)&lpdis->rcItem, hBrush);
					} 
					DrawText(lpdis->hDC, sBuffer, wcslen(sBuffer), &rect, DT_VCENTER |
						DT_SINGLELINE); 
					rect.left=lpdis->rcItem.left + 2; 
					rect.right = rect.left + CHECKWIDTH; 
					if (SendMessage(lpdis->hwndItem, LB_GETITEMDATA, lpdis->itemID, 0))
						DrawFrameControl(lpdis->hDC,&rect,DFC_BUTTON,DFCS_BUTTONCHECK | 
						DFCS_CHECKED); 
					else 
						DrawFrameControl(lpdis->hDC,&rect,DFC_BUTTON,DFCS_BUTTONCHECK); 
					if (lpdis->itemState & ODS_FOCUS) 
						DrawFocusRect(lpdis->hDC, (LPRECT)&lpdis->rcItem);
					if (nOldTextColor != -1) 
						SetTextColor(lpdis->hDC, nOldTextColor); 
				} 
			} 
			return TRUE; 
		} 
		break;

	case WM_PAINT:
		hdc = BeginPaint(hWnd, &ps);
		// TODO: Add any drawing code here...
		EndPaint(hWnd, &ps);
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}
