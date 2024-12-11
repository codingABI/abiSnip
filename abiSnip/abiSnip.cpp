/*+===================================================================
  File:      abiSnip.cpp

  Summary:   Tool to save screenshots as PNG files or copy screenshots to clipboard
             when the "Print screen" key was pressed

			 Supports:
			 - Zooming the mouse pointer area
			 - Area selection
			 - All monitors selection
			 - Single monitor selection
			 - Selections can be adjusted by mouse or keyboard
			 - Folder for PNG files can be set with the context menu of the tray icon
			 - Filename for a PNG file will be set automatically and contains a timestamp, for example "Screenshot 2024-11-24 100706.png"

			 Program should run on Windows 11/10/8.1/2025/2022/2019/2016/2012R2

  License: CC0
  Copyright (c) 2024 codingABI

  Commands:
  Print screen = Start screenshot
  A = Select all monitors
  M = Select next monitor
  Tab = Toggle point A <-> B
  Cursor keys = Move point A/B
  Alt+cursor keys = Fast move point A/B
  Shift+cursor keys = Find color change for point A/B
  Return or left mouse button click = OK, confirm selection
  ESC = Cancel the screenshot
  Insert = Store selection
  Home = Restore selection
  +/- = Increase/decrease selection
  PageUp/PageDown, mouse wheel = Zoom In/Out
  C = Clipboard On/Off
  F = File On/Off
  S = Alternative colors On/Off
  F1 = Display internal information on screen On/Off

  Refs:
  https://learn.microsoft.com/en-us/windows/win32/gdi/capturing-an-image
  https://devblogs.microsoft.com/oldnewthing/20100412-00/?p=14353
  https://www.codeproject.com/Tips/76427/How-to-bring-window-to-top-with-SetForegroundWindo
  https://devblogs.microsoft.com/oldnewthing/20150406-00/?p=44303

  History:
  20241202, Initial version

===================================================================+*/

// For GNU compilers
#if defined (__GNUC__)
#define UNICODE
#define NO_SHLWAPI_STRFCNS
#define _WIN32_WINNT 0x0A00
#define _MAX_ITOSTR_BASE16_COUNT (8 + 1) // Char length for DWORD to hex conversion
#endif

// Includes
#include <windows.h>
#include <WindowsX.h>
#include <wingdi.h>
#include <shlwapi.h>
#include <shlobj.h>
#include <gdiplus.h>
#include <string>
#include <sysinfoapi.h>
#include <vector>
#include "resource.h"

// Library-search records for visual studio
#pragma comment(lib,"msimg32")
#pragma comment(lib,"Shlwapi")
#pragma comment(lib,"Gdiplus.lib")
#pragma comment(lib,"Version.lib")

using namespace Gdiplus;
// Defines
#define REGISTRYPATH L"Software\\CodingABI\\abiSnip" // Regisry path under HKCU to store program settings
#define ZOOMWIDTH 32 // Width of zoomwindow (effective pixel size is ZOOMWIDTH * current zoom scale)
#define ZOOMHEIGHT 32 // Height of zoomwindow (effective pixel size is ZOOMHEIGHT * current zoom scale)
#define MAXZOOMFACTOR 32 // Max zoom scale
#define DEFAULTZOOMSCALE 4 // Default zooom scale when selecting point A or B
#define DEFAULTFONT L"Consolas" // Font
#define DEFAULTSAVETOCLIPBOARD TRUE // TRUE, when screenshot should be saved to clipboard
#define DEFAULTSAVETOFILE TRUE // TRUE, when screenshot should be saved to a PNG file
#define DEFAULTUSEALTERNATIVECOLORS FALSE // TRUE, when alternative colors are enabled
#define DEFAULTSHOWDISPLAYINFORMATION FALSE // TRUE, when drawing internal information on screen is enabled

// Default colors
#define APPCOLOR RGB(245, 167, 66)
#define APPCOLORINV RGB(255,255,255)
// Alternative colors
#define ALTAPPCOLOR RGB(0, 116, 129)
#define ALTAPPCOLORINV RGB(255,255,255)

// Vector to store monitor positions
std::vector<RECT> g_rectMonitor;
DWORD g_selectedMonitor = 0;

// Cursor types
enum BOXTYPE {
	BoxFirstPointA, // Centered cursor for point A
	BoxFinalPointA, // Edge cursor for point A
	BoxFinalPointB // Edge cursor for point B
};

// Program states
enum APPSTATE {
	stateTrayIcon, // Hidden, only tray icon visible
	stateFirstPoint, // Selection of first point A in fullscreen mode
	statePointA, // Modification of point A in fullscreen mode
	statePointB, // Selection/Modification of point B in fullscreen mode
};

// Simple DWORD settings
enum APPDWORDSETTINGS {
	defaultZoomScale,
	saveToClipboard,
	useAlternativeColors,
	displayInternallnformation,
	saveToFile,
	storedSelectionLeft,
	storedSelectionTop,
	storedSelectionRight,
	storedSelectionBottom,
};

// Global Variables:
HINSTANCE g_hInst = NULL; // Current instance
HWND g_hWindow = NULL; // Handle to main window
POINT g_appWindowPos; // SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN when fullscreen was started
HBITMAP g_hBitmap = NULL; // Bitmap for screenshot over all monitors
RECT g_selection = { -1,-1,-1,-1 }; // Selected screenshot area
RECT g_storedSelection = { -1,-1,-1,-1 }; // Stored selection
BOOL g_useAlternativeColors = DEFAULTUSEALTERNATIVECOLORS;
BOOL g_saveToFile = DEFAULTSAVETOFILE;
BOOL g_saveToClipboard = DEFAULTSAVETOCLIPBOARD;
BOOL g_displayInternallnformation = DEFAULTSHOWDISPLAYINFORMATION;
APPSTATE g_appState = stateTrayIcon; // Current program state
HWND g_activeWindow = NULL; // Active window before program starts fullscreen mode
int g_zoomScale = DEFAULTZOOMSCALE; // Zoom scale for mouse cursor
HHOOK g_hHook = NULL; // Handle to hook (We use keyboard hook to start fullscreen mode, when the "Print screen" key was pressed)
NOTIFYICONDATA g_nid; // Tray icon structure

// Function declarations
ATOM                MyRegisterClass(HINSTANCE hInstance);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: KeyboardProc

  Summary:  Keyboard hook procedure

  Args:     int nCode
			WPARAM wParam
			LPARAM lParam

  Returns:  LRESULT

-----------------------------------------------------------------F-F*/
LRESULT CALLBACK KeyboardProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	if (nCode == HC_ACTION)
	{
		KBDLLHOOKSTRUCT* pKeyBoard = (KBDLLHOOKSTRUCT*)lParam;
		if (pKeyBoard->vkCode == VK_SNAPSHOT)
		{
			if (g_appState == stateTrayIcon) SendMessage(g_hWindow, WM_STARTED, 0, 0);
			return 1; // Prevents keypress forwarding
		}
	}
	return CallNextHookEx(g_hHook, nCode, wParam, lParam);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: SetHook

  Summary:  Creates keyboard hook

  Args:

  Returns:

-----------------------------------------------------------------F-F*/
void SetHook()
{
	g_hHook = SetWindowsHookEx(WH_KEYBOARD_LL, KeyboardProc, NULL, 0);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: ReleaseHook

  Summary:  Release the keyboard hook

  Args:

  Returns:

-----------------------------------------------------------------F-F*/
void ReleaseHook()
{
	if (g_hHook != NULL) UnhookWindowsHookEx(g_hHook);
}


/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: SetForegroundWindowInternal

  Summary:  Set window to foreground and set focus to window
			From: https://www.codeproject.com/Tips/76427/How-to-bring-window-to-top-with-SetForegroundWindo

  Args:		HWND hWindow
			  Handle to window

  Returns:

-----------------------------------------------------------------F-F*/
void SetForegroundWindowInternal(HWND hWindow)
{
	if (!::IsWindow(hWindow)) return;

	BYTE keyState[256] = { 0 };
	// To unlock SetForegroundWindow we need to imitate Alt pressing
	if (::GetKeyboardState((LPBYTE)&keyState))
	{
		if (!(keyState[VK_MENU] & 0x80))
		{
			::keybd_event(VK_MENU, 0, KEYEVENTF_EXTENDEDKEY | 0, 0);
		}
	}

	::SetForegroundWindow(hWindow);

	if (::GetKeyboardState((LPBYTE)&keyState))
	{
		if (!(keyState[VK_MENU] & 0x80))
		{
			::keybd_event(VK_MENU, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
		}
	}
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: LoadStringAsWstr

  Summary:  Get resource string as wstring

  Args:     HINSTANCE hInstance
			  Handle to instance module
			UINT uID
			  Resource ID

  Returns:  std::wstring

-----------------------------------------------------------------F-F*/
std::wstring LoadStringAsWstr(HINSTANCE hInstance, UINT uID)
{
	PCWSTR pws;
	int cchStringLength = LoadStringW(hInstance, uID, reinterpret_cast<LPWSTR>(&pws), 0);
	if (cchStringLength > 0) return std::wstring(pws, cchStringLength); else return std::wstring();
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: showProgramInformation

  Summary:   Show program information message box

  Args:     HWND hWindow
			  Handle to main window

  Returns:

-----------------------------------------------------------------F-F*/
void showProgramInformation(HWND hWindow)
{
	std::wstring sTitle(LoadStringAsWstr(g_hInst, IDS_APP_TITLE));
	wchar_t szExecutable[MAX_PATH];
	GetModuleFileName(NULL, szExecutable, MAX_PATH);

	DWORD  verHandle = 0;
	UINT   size = 0;
	LPBYTE lpBuffer = NULL;
	DWORD  verSize = GetFileVersionInfoSize(szExecutable, NULL);

	if (verSize != 0) {
		BYTE* verData = new BYTE[verSize];

		if (GetFileVersionInfo(szExecutable, verHandle, verSize, verData))
		{
			if (VerQueryValue(verData, L"\\", (VOID FAR * FAR*) & lpBuffer, &size))
			{
				if (size)
				{
					VS_FIXEDFILEINFO* verInfo = (VS_FIXEDFILEINFO*)lpBuffer;
					if (verInfo->dwSignature == 0xfeef04bd) {
						sTitle.append(L" ")
							.append(std::to_wstring((verInfo->dwFileVersionMS >> 16) & 0xffff))
							.append(L".")
							.append(std::to_wstring((verInfo->dwFileVersionMS >> 0) & 0xffff))
							.append(L".")
							.append(std::to_wstring((verInfo->dwFileVersionLS >> 16) & 0xffff))
							.append(L".")
							.append(std::to_wstring((verInfo->dwFileVersionLS >> 0) & 0xffff));
					}
				}
			}
		}
		delete[] verData;
	}

	MessageBox(hWindow, LoadStringAsWstr(g_hInst, IDS_PROGINFO).c_str(), sTitle.c_str(), MB_ICONINFORMATION | MB_OK);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: getDWORDSettingFromRegistry

  Summary:   Gets stored DWORD setting from registry

  Args:      APPDWORDSETTINGS setting
			   Setting

  Returns:	DWORD
			  Setting from registry
			  0xFFFFFFF = Error

-----------------------------------------------------------------F-F*/
DWORD getDWORDSettingFromRegistry(APPDWORDSETTINGS setting) {
	DWORD dwResult = 0;

	std::wstring sValueName = L"";
	switch (setting)
	{
	case defaultZoomScale:
		sValueName.assign(L"defaultZoomScale");
		dwResult = DEFAULTZOOMSCALE;
		break;
	case saveToClipboard:
		sValueName.assign(L"saveToClipboard");
		dwResult = DEFAULTSAVETOCLIPBOARD;
		break;
	case saveToFile:
		sValueName.assign(L"saveToFile");
		dwResult = DEFAULTSAVETOFILE;
		break;
	case useAlternativeColors:
		sValueName.assign(L"useAlternativeColors");
		dwResult = DEFAULTUSEALTERNATIVECOLORS;
		break;
	case displayInternallnformation:
		sValueName.assign(L"displayInternallnformation");
		dwResult = DEFAULTSHOWDISPLAYINFORMATION;
		break;
	case storedSelectionLeft:
		sValueName.assign(L"storedSelectionLeft");
		dwResult = -1;
		break;
	case storedSelectionTop:
		sValueName.assign(L"storedSelectionTop");
		dwResult = -1;
		break;
	case storedSelectionRight:
		sValueName.assign(L"storedSelectionRight");
		dwResult = -1;
		break;
	case storedSelectionBottom:
		sValueName.assign(L"storedSelectionBottom");
		dwResult = -1;
		break;
	default:
		OutputDebugString(L"Invalid setting");
		return 0xFFFFFFF;
	}
	if (sValueName.empty())
	{
		OutputDebugString(L"Invalid setting");
		return 0xFFFFFFF;
	}

	HKEY hKey;
	DWORD dwValue;
	DWORD dwSize = sizeof(DWORD);
	LONG lResult;

	// Open registry value
	lResult = RegOpenKeyEx(HKEY_CURRENT_USER, REGISTRYPATH, 0, KEY_READ, &hKey);
	if (lResult == ERROR_SUCCESS) {

		// Read value from registry
		lResult = RegQueryValueEx(hKey, sValueName.c_str(), NULL, NULL, (LPBYTE)&dwValue, &dwSize);
		if (lResult == ERROR_SUCCESS) {
			dwResult = dwValue;
		}
		RegCloseKey(hKey);
	}

	// Check limits
	switch (setting)
	{
	case defaultZoomScale:
		if (dwResult < 1) dwResult = 1;
		if (dwResult > MAXZOOMFACTOR) dwResult = MAXZOOMFACTOR;
		break;
	case saveToClipboard:
	case saveToFile:
	case useAlternativeColors:
	case displayInternallnformation:
		if (dwResult > 1) dwResult = 1;
		break;
	}

	return dwResult;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: storeDWORDSettingInRegistry

  Summary:   Stores DWORD setting in registry

  Args:      APPDWORDSETTINGS setting
			   Setting
             DWORD dwValue
			   Value to write in registry

  Returns:	BOOL
			  TRUE = success
			  FALSE = failure

-----------------------------------------------------------------F-F*/
BOOL storeDWORDSettingInRegistry(APPDWORDSETTINGS setting, DWORD dwValue) {
	BOOL bResult = TRUE;

	std::wstring sValueName = L"";
	switch (setting)
	{
	case defaultZoomScale:
		sValueName.assign(L"defaultZoomScale");
		break;
	case saveToClipboard:
		sValueName.assign(L"saveToClipboard");
		break;
	case saveToFile:
		sValueName.assign(L"saveToFile");
		break;
	case useAlternativeColors:
		sValueName.assign(L"useAlternativeColors");
		break;
	case displayInternallnformation:
		sValueName.assign(L"displayInternallnformation");
		break;
	case storedSelectionLeft:
		sValueName.assign(L"storedSelectionLeft");
		break;
	case storedSelectionTop:
		sValueName.assign(L"storedSelectionTop");
		break;
	case storedSelectionRight:
		sValueName.assign(L"storedSelectionRight");
		break;
	case storedSelectionBottom:
		sValueName.assign(L"storedSelectionBottom");
		break;
	default:
		OutputDebugString(L"Invalid setting");
		return FALSE;
	}
	if (sValueName.empty())
	{
		OutputDebugString(L"Invalid setting");
		return FALSE;
	}

	HKEY hKey;
	LONG lResult;

	// Open registry value
	lResult = RegCreateKeyEx(HKEY_CURRENT_USER, REGISTRYPATH, 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL);
	if (lResult == ERROR_SUCCESS) {

		// Write to registry
		lResult = RegSetValueEx(hKey, sValueName.c_str(), 0, REG_DWORD, (const BYTE*)&dwValue, sizeof(dwValue));
		if (lResult != ERROR_SUCCESS) {
			OutputDebugString(L"Error writing to registry");
			bResult = FALSE;
		}
		RegCloseKey(hKey);
	}
	return bResult;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: getScreenshotPathFromRegistry

  Summary:   Get path of screenshot folder from registry or path to EXE file

  Args:

  Returns:	std::wstring
			  Path of screenshot folder

-----------------------------------------------------------------F-F*/
std::wstring getScreenshotPathFromRegistry()
{
	std::wstring sResult;

	// Get stored path from registry
	wchar_t szValue[MAX_PATH];
	szValue[0] = L'\0';
	DWORD valueSize = 0;
	DWORD keyType = 0;
	// Get size of registry value
	if (RegGetValue(HKEY_CURRENT_USER, REGISTRYPATH, L"screenshotPath", RRF_RT_REG_SZ, &keyType, NULL, &valueSize) == ERROR_SUCCESS)
	{
		if ((valueSize > 0) && (valueSize <= (MAX_PATH + 1) * sizeof(WCHAR))) // Size OK?
		{
			// Get registry value
			valueSize = MAX_PATH * sizeof(WCHAR); // Max size incl. termination
			if (RegGetValue(HKEY_CURRENT_USER, REGISTRYPATH, L"screenshotPath", RRF_RT_REG_SZ | RRF_ZEROONFAILURE, NULL, &szValue, &valueSize) != ERROR_SUCCESS)
			{
				szValue[0] = L'\0';
			}
		}
	}
	if (wcslen(szValue) > 0)
	{
		sResult.assign(szValue);
	}
	else
	{ // No registry value found => Use path of EXE file
		wchar_t szWorkingFolder[MAX_PATH];
		GetModuleFileName(NULL, szWorkingFolder, MAX_PATH);
		PathRemoveFileSpec(szWorkingFolder);
		sResult.assign(szWorkingFolder);
	}
	return sResult;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: changeScreenshotPathAndStorePathToRegistryCallbackProc

  Summary:   Callback function to set the initial folder

  Args:		HWND hwnd
  			UINT uMsg
			LPARAM lParam
			LPARAM lpData

  Returns:	int
  			  Always 0

-----------------------------------------------------------------F-F*/
int CALLBACK changeScreenshotPathAndStorePathToRegistryCallbackProc(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData) {
    if (uMsg == BFFM_INITIALIZED) {
        // Set the initial folder
        SendMessage(hwnd, BFFM_SETSELECTION, TRUE, lpData);
    }
    return 0;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: changeScreenshotPathAndStorePathToRegistry

  Summary:   Let user choose a new path for the screenshot folder and save it to registry

  Args:

  Returns:

-----------------------------------------------------------------F-F*/
void changeScreenshotPathAndStorePathToRegistry()
{
	std::wstring sMessage = LoadStringAsWstr(g_hInst, IDS_SELECTFOLDER);
	WCHAR szStartPath[MAX_PATH];
	LPITEMIDLIST pidl;

	BROWSEINFOW bi = { 0 };
	bi.hwndOwner = g_hWindow;
	bi.lpszTitle = sMessage.c_str();
	bi.ulFlags = BIF_NEWDIALOGSTYLE | BIF_RETURNONLYFSDIRS | BIF_EDITBOX | BIF_VALIDATE;

    // Convert the start path to a PIDL
	_snwprintf_s(szStartPath, MAX_PATH, _TRUNCATE, L"%s",getScreenshotPathFromRegistry().c_str());
    HRESULT hr = SHParseDisplayName(szStartPath, NULL, &pidl, 0, NULL);
    if (SUCCEEDED(hr))
	{
		bi.lpfn = changeScreenshotPathAndStorePathToRegistryCallbackProc;
        bi.lParam = (LPARAM)szStartPath;
	}

	// Browse for folder dialog
	LPITEMIDLIST pidlSelected = SHBrowseForFolder(&bi);
	if (pidlSelected != NULL)
	{
		WCHAR szPath[MAX_PATH];
		if (SHGetPathFromIDList(pidlSelected, szPath))
		{
			RegSetKeyValue(HKEY_CURRENT_USER, REGISTRYPATH, L"screenshotPath", REG_SZ, szPath, (DWORD)(wcslen(szPath) + 1) * sizeof(WCHAR));
		}
		CoTaskMemFree(pidlSelected);
	}
	CoTaskMemFree(pidl);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: GetEncoderClsid

  Summary:   Get encoder CLsid for a mime type string

  Args:     const WCHAR* format
			  Mime type string
			CLSID* pClsid
			  Pointer to encoder CLsid

  Returns:	int
			  >=0 = Ok (Index in encoder array)
			  -1 = Error

-----------------------------------------------------------------F-F*/
int GetEncoderClsid(const WCHAR* format, CLSID* pClsid)
{
	Status status;
	UINT num = 0; // Number of available image encoders
	UINT size = 0; // Total size, in bytes, of the array of ImageCodecInfo objects

	Gdiplus::ImageCodecInfo* pImageCodecInfo = NULL;

	status = Gdiplus::GetImageEncodersSize(&num, &size);
	if (status != Gdiplus::Ok) return -1; // Windows GDI+ not OK
	if (size == 0) return -1; // Error
	if (num == 0) return -1; // No encoders

	pImageCodecInfo = (Gdiplus::ImageCodecInfo*)(malloc(size));
	if (pImageCodecInfo == NULL) return -1;

	status = Gdiplus::GetImageEncoders(num, size, pImageCodecInfo);
	if (status != Gdiplus::Ok) // Windows GDI+ not OK
	{
		free(pImageCodecInfo);
		return -1; // Error
	}

	for (UINT j = 0; j < num; j++)
	{
		// warning C6385: Reading invalid data from 'pImageCodecInfo'
#ifdef _MSC_VER
#pragma warning( suppress: 6385 )
#endif
		if (wcscmp(pImageCodecInfo[j].MimeType, format) == 0)
		{
			*pClsid = pImageCodecInfo[j].Clsid;
			free(pImageCodecInfo);
			return j; // Success
		}
	}
	free(pImageCodecInfo);
	return -1; // Error
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: enterFullScreen

  Summary:   Set program to fullscreen

  Args:     HWND hWindow
			  Handle to window

  Returns:

-----------------------------------------------------------------F-F*/
void enterFullScreen(HWND hWindow)
{
	DWORD dwStyle = GetWindowLong(hWindow, GWL_STYLE);

	int screenX = GetSystemMetrics(SM_XVIRTUALSCREEN);
	int screenY = GetSystemMetrics(SM_YVIRTUALSCREEN);
	int screenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
	int screenHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);

	g_appWindowPos.x = screenX;
	g_appWindowPos.y = screenY;
	SetWindowLong(hWindow, GWL_STYLE, dwStyle & ~WS_OVERLAPPEDWINDOW);
	SetWindowPos(hWindow, HWND_TOPMOST,
		screenX, screenY,
		screenWidth,
		screenHeight,
		SWP_NOOWNERZORDER | SWP_FRAMECHANGED);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: MySetCursorPos

  Summary:   Set mouse cursor position (Variant of SetCursorPos. See comments for "Check position")

  Args:     int posX
			  mouse x position
			int posY
			  mouse y position

  Returns:

-----------------------------------------------------------------F-F*/
void MySetCursorPos(int posX, int posY)
{
	SetCursorPos(posX + g_appWindowPos.x, posY + g_appWindowPos.y);

	/*
	 * Check position
	 *
	 * Under special conditions SetCursorPos does not set correct position, see:
	 * https://stackoverflow.com/questions/65519784/why-does-setcursorpos-reset-the-cursor-position-to-the-left-hand-side-of-the-dis
	 * https://stackoverflow.com/questions/58753372/winapi-setcursorpos-seems-like-not-working-properly-on-multiple-monitors-with-di
	 * As a workaround, try move horizontally first and than repeat whole position
	 */
	POINT mouse;
	GetCursorPos(&mouse);
	if ((mouse.x - g_appWindowPos.x != posX) || (mouse.y - g_appWindowPos.y != posY))
	{
		SetCursorPos(posX + g_appWindowPos.x, mouse.y / 2);
		SetCursorPos(posX + g_appWindowPos.x, posY + g_appWindowPos.y);
	}
}


/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: limitXtoBitmap

  Summary:   Ensures that X is inside screenshot bitmap

  Args:     int X

  Returns:	int
			  X inside the screenshot bitmap

-----------------------------------------------------------------F-F*/
int limitXtoBitmap(int X) {
	if (g_hBitmap == NULL) return X;

	BITMAP bm;
	GetObject(g_hBitmap, sizeof(bm), &bm);

	if (X < 0) return 0;
	if (X > bm.bmWidth - 1) return bm.bmWidth - 1;
	return X;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: limitYtoBitmap

  Summary:   Ensures that Y is inside screenshot bitmap

  Args:     int Y

  Returns:	int
			  Y inside the screenshot bitmap

-----------------------------------------------------------------F-F*/
int limitYtoBitmap(int Y) {
	if (g_hBitmap == NULL) return Y;

	BITMAP bm;
	GetObject(g_hBitmap, sizeof(bm), &bm);

	if (Y < 0) return 0;
	if (Y > bm.bmHeight - 1) return bm.bmHeight - 1;
	return Y;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: SaveBitmapAsPNG

  Summary:   Save bitmap as PNG file

  Args:     HBITMAP hBitmap
			  Bitmap
			const WCHAR* fileName
			  Filename for PNG file

  Returns:	BOOL
			  TRUE = Success
			  FALSE = Error

-----------------------------------------------------------------F-F*/
BOOL SaveBitmapAsPNG(HBITMAP hBitmap, const WCHAR* fileName)
{
	BOOL bRC = TRUE;
	wchar_t szHex[_MAX_ITOSTR_BASE16_COUNT + 2];
	std::wstring sError;
	ULONG_PTR gdiplusToken;
	GdiplusStartupInput gdiplusStartupInput;
	std::wstring sFullPathWorkingFile;

	GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

	// Convert bitmap to GDI+
	Bitmap* bitmap = new Bitmap(hBitmap, NULL);
	if (bitmap != NULL)
	{
		// Create PNG
		CLSID clsidPng;
		GetEncoderClsid(L"image/png", &clsidPng);

		bool bFinished = false;
		do
		{
			// Get Path
			sFullPathWorkingFile.assign(getScreenshotPathFromRegistry()).append(L"\\").append(fileName);

			Status status = bitmap->Save(sFullPathWorkingFile.c_str(), &clsidPng, NULL);
			if (status != Gdiplus::Ok) // Windows GDI+ not OK
			{
				sError.assign(LoadStringAsWstr(g_hInst, IDS_ERRORCREATING).c_str()).append(L"\n").append(sFullPathWorkingFile);

				if (status == Win32Error) {
					size_t size;
					LPWSTR messageBuffer = nullptr;
					size = FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
						NULL, GetLastError(), MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPWSTR)&messageBuffer, 0, NULL);
					if (size > 0)
					{
						StrTrim(messageBuffer, L"\r\n");
						sError.append(L"\n").append(messageBuffer);
					}
					_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT + 2, _TRUNCATE, L"0x%08X", GetLastError());
					LocalFree(messageBuffer);
					sError.append(L" ").append(szHex).append(L"\n").append(LoadStringAsWstr(g_hInst, IDS_CHANGEFOLDER));
					if (MessageBox(g_hWindow, sError.c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OKCANCEL | MB_ICONERROR) != IDCANCEL)
						changeScreenshotPathAndStorePathToRegistry();
					else
						bFinished = true;
				}
				else
				{
					_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT + 2, _TRUNCATE, L"0x%08X", status);
					sError.append(L"\nStatus:").append(szHex);
					MessageBox(g_hWindow, sError.c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONERROR);
					bFinished = true;
				}
				bRC = FALSE;
			}
			else {
				bFinished = true;
			}
		} while (!bFinished);
		delete bitmap;
	}
	else bRC = FALSE;
	GdiplusShutdown(gdiplusToken);
	return bRC;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: saveSelection

  Summary:   Save selected area to clipboard or/and file

  Args:     HWND hWindow
			  Handle to window

  Returns:

-----------------------------------------------------------------F-F*/
void saveSelection(HWND hWindow)
{
	HDC hdcMemFullScreen = NULL;
	HDC hdcMemSelection = NULL;
	HBITMAP hOldBitmapFullScreen = NULL;
	HBITMAP hOldBitmapSelection = NULL;
	HBITMAP hBitmap = NULL;

	RECT finalSelection;

	// Retrieves information the entire screenshot bitmap
	BITMAP bm;
	GetObject(g_hBitmap, sizeof(bm), &bm);

	// Normalize selected rectangle
	if (g_selection.left < g_selection.right) {
		finalSelection.left = g_selection.left;
		finalSelection.right = g_selection.right;
	}
	else
	{
		finalSelection.left = g_selection.right;
		finalSelection.right = g_selection.left;
	}
	if (g_selection.top < g_selection.bottom) {
		finalSelection.top = g_selection.top;
		finalSelection.bottom = g_selection.bottom;
	}
	else
	{
		finalSelection.top = g_selection.bottom;
		finalSelection.bottom = g_selection.top;
	}
	if (finalSelection.left < 0) finalSelection.left = 0;
	if (finalSelection.right > bm.bmWidth - 1) finalSelection.right = bm.bmWidth - 1;
	if (finalSelection.top < 0) finalSelection.top = 0;
	if (finalSelection.bottom > bm.bmHeight - 1) finalSelection.bottom = bm.bmHeight - 1;

	int selectionWidth = finalSelection.right - finalSelection.left + 1;
	int selectionHeight = finalSelection.bottom - finalSelection.top + 1;

	// Copy selected area from entire screenshot to new bitmap
	hdcMemFullScreen = CreateCompatibleDC(NULL);
	if (hdcMemFullScreen == NULL) goto CLEANUP;

	hOldBitmapFullScreen = (HBITMAP)SelectObject(hdcMemFullScreen, g_hBitmap);
	if (hOldBitmapFullScreen == NULL) goto CLEANUP;

	hdcMemSelection = CreateCompatibleDC(NULL);
	if (hdcMemSelection == NULL) goto CLEANUP;

	hBitmap = CreateCompatibleBitmap(hdcMemFullScreen, selectionWidth, selectionHeight);
	if (hBitmap == NULL) goto CLEANUP;

	hOldBitmapSelection = (HBITMAP)SelectObject(hdcMemSelection, hBitmap);
	if (hOldBitmapSelection == NULL) goto CLEANUP;

	if (BitBlt(hdcMemSelection, 0, 0, selectionWidth, selectionHeight,
		hdcMemFullScreen, finalSelection.left, finalSelection.top, SRCCOPY))
	{
		if (g_saveToFile) // Save selected area to file?
		{
			// Create folder
			CreateDirectory(getScreenshotPathFromRegistry().c_str(), NULL);

			// Create file
			SYSTEMTIME tLocal;
			GetLocalTime(&tLocal);
			wchar_t szFileName[MAX_PATH];

#define FILEPATTERN L"Screenshot %04u-%02u-%02u %02d%02d%02d.png"
			if (_snwprintf_s(szFileName, MAX_PATH, _TRUNCATE, FILEPATTERN, tLocal.wYear, tLocal.wMonth, tLocal.wDay, tLocal.wHour, tLocal.wMinute, tLocal.wSecond) >= 0) {
				SaveBitmapAsPNG(hBitmap, szFileName);
			}
		}

		if (g_saveToClipboard) // Save selected area to clipboard?
		{
			if (OpenClipboard(NULL))
			{
				EmptyClipboard(); // Clear clipboard

				// Set bitmap to clipboard
				SetClipboardData(CF_BITMAP, hBitmap);

				CloseClipboard(); // Close clipboard
			}
			else
			{
				// Clipboard error
				MessageBox(hWindow, LoadStringAsWstr(g_hInst, IDS_ERRORCOPYTOCLIPBOARD).c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONERROR);
			}
		}

		if (!g_saveToClipboard && !g_saveToFile) // Clipboard and file was disabled by user => Warning
		{
			MessageBox(hWindow, LoadStringAsWstr(g_hInst, IDS_NOSAVING).c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONWARNING);
		}

	}
	else
	{
		// BitBlt error
		std::wstring sMessage;
		sMessage.assign(L"BitBlt ").append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		MessageBox(hWindow, sMessage.c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONERROR);
	}

CLEANUP:
	// Free resources/Cleanup
	if (hBitmap != NULL) DeleteObject(hBitmap);

	if (hdcMemSelection != NULL)
	{
		if (hOldBitmapSelection != NULL) SelectObject(hdcMemSelection, hOldBitmapSelection);
		DeleteDC(hdcMemSelection);
	}

	if (hdcMemFullScreen != NULL)
	{
		if (hOldBitmapFullScreen != NULL) SelectObject(hdcMemFullScreen, hOldBitmapFullScreen);
		DeleteDC(hdcMemFullScreen);
	}
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: zoomMousePosition

  Summary:   Creates zoom boxes for mouse position or point A/B

  Args:     HDC hdc
			  Handle to drawing context for the window client area
			HDC hdcMemoryDevice
			  Handle to drawing context for the buffer
			BOXTYPE boxType
			  Type of position (BoxFirstPointA, BoxFinalPointA, BoxFinalPointB)

  Returns:

-----------------------------------------------------------------F-F*/
void zoomMousePosition(HDC hdc, HDC hdcMemoryDevice, BOXTYPE boxType)
{
#define MAXSTRDATAZOOM 30
	wchar_t strData[MAXSTRDATAZOOM];
	UINT textFormat = 0;
	POINT textPosition = { 0, 0 };
	RECT rectText{ 0, 0, 0, 0 };
	PLOGFONT plf = (PLOGFONT)LocalAlloc(LPTR, sizeof(LOGFONT));
	HGDIOBJ hfnt = NULL, hfntPrev = NULL;
	HBRUSH hBrush = NULL;
	int zoomCenterX, zoomCenterY, zoomBoxX, zoomBoxY;

	if ((boxType != BoxFirstPointA) && (abs(g_selection.right - g_selection.left) < ZOOMWIDTH * g_zoomScale)) return; // Selection too small
	if ((boxType != BoxFirstPointA) && (abs(g_selection.bottom - g_selection.top) < ZOOMHEIGHT * g_zoomScale)) return; // Selection too small

	// Font
	// Specify a font typeface name and weight.
	if (plf == NULL) goto CLEANUP;
	if (_snwprintf_s(plf->lfFaceName, 12, _TRUNCATE, L"%s", DEFAULTFONT) < 0) goto CLEANUP;
	plf->lfWeight = FW_NORMAL;
	hfnt = CreateFontIndirect(plf);
	if (hfnt == NULL) goto CLEANUP;

	hfntPrev = SelectObject(hdcMemoryDevice, hfnt);
	if (hfntPrev == NULL) goto CLEANUP;

	switch (boxType)
	{
	case BoxFirstPointA:
		zoomCenterX = g_selection.left;
		zoomCenterY = g_selection.top;
		zoomBoxX = zoomCenterX - g_zoomScale * ZOOMWIDTH / 2 - g_zoomScale / 2;
		zoomBoxY = zoomCenterY - g_zoomScale * ZOOMHEIGHT / 2 - g_zoomScale / 2;
		break;
	case BoxFinalPointA:
		if (g_selection.right >= g_selection.left)
		{
			zoomBoxX = g_selection.left;
			zoomCenterX = g_selection.left + ZOOMWIDTH / 2;
		}
		else
		{
			zoomBoxX = g_selection.left - g_zoomScale * ZOOMWIDTH + 1;
			zoomCenterX = g_selection.left - ZOOMWIDTH / 2 + 1;
		}
		if (g_selection.bottom >= g_selection.top)
		{
			zoomBoxY = g_selection.top;
			zoomCenterY = g_selection.top + ZOOMHEIGHT / 2;
		}
		else
		{
			zoomBoxY = g_selection.top - g_zoomScale * ZOOMHEIGHT + 1;
			zoomCenterY = g_selection.top - ZOOMHEIGHT / 2 + 1;
		}
		break;
	case BoxFinalPointB:
		if (g_selection.right < g_selection.left)
		{
			zoomBoxX = g_selection.right;
			zoomCenterX = g_selection.right + ZOOMWIDTH / 2;
		}
		else
		{
			zoomBoxX = g_selection.right - g_zoomScale * ZOOMWIDTH + 1;
			zoomCenterX = g_selection.right - ZOOMWIDTH / 2 + 1;
		}
		if (g_selection.bottom < g_selection.top)
		{
			zoomBoxY = g_selection.bottom;
			zoomCenterY = g_selection.bottom + ZOOMHEIGHT / 2;
		}
		else
		{
			zoomBoxY = g_selection.bottom - g_zoomScale * ZOOMHEIGHT + 1;
			zoomCenterY = g_selection.bottom - ZOOMHEIGHT / 2 + 1;
		}
		break;
	}

	// Zoom bitmap
	if (g_hBitmap != NULL)
	{
		HDC hdcMem = CreateCompatibleDC(hdc);
		if (hdcMem == NULL) return;
		HGDIOBJ hbmOld = SelectObject(hdcMem, g_hBitmap);
		BITMAP bm;
		GetObject(g_hBitmap, sizeof(BITMAP), (PSTR)&bm);
		SetStretchBltMode(hdcMemoryDevice, COLORONCOLOR);

		StretchBlt(hdcMemoryDevice, zoomBoxX, zoomBoxY, ZOOMWIDTH * g_zoomScale, ZOOMHEIGHT * g_zoomScale,
			hdcMem, zoomCenterX - ZOOMWIDTH / 2, zoomCenterY - ZOOMHEIGHT / 2, ZOOMWIDTH, ZOOMHEIGHT, SRCCOPY); // SRCCOPY) ;

		SelectObject(hdcMem, hbmOld);
		DeleteDC(hdcMem);
	}

	// Frame
	RECT outer;
	outer.left = zoomBoxX - 1;
	outer.top = zoomBoxY - 1;
	outer.right = zoomBoxX + ZOOMWIDTH * g_zoomScale + 1;
	outer.bottom = zoomBoxY + ZOOMHEIGHT * g_zoomScale + 1;
	hBrush = CreateSolidBrush(g_useAlternativeColors ? ALTAPPCOLOR : APPCOLOR);
	if (hBrush == NULL) goto CLEANUP;

	if (g_zoomScale > 1) FrameRect(hdcMemoryDevice, &outer, hBrush);

	// Cross
	if (boxType == BoxFirstPointA)
	{
		RECT center;
		center.left = zoomBoxX - 1;
		center.top = zoomBoxY + g_zoomScale * ZOOMHEIGHT / 2 - 1;
		center.right = zoomBoxX + ZOOMWIDTH * g_zoomScale + 1;
		center.bottom = center.top + g_zoomScale + 2;

		FrameRect(hdcMemoryDevice, &center, hBrush);

		center.left = zoomBoxX + g_zoomScale * ZOOMWIDTH / 2 - 1;
		center.top = zoomBoxY - 1;
		center.right = center.left + g_zoomScale + 2;
		center.bottom = zoomBoxY + ZOOMHEIGHT * g_zoomScale + 1;

		FrameRect(hdcMemoryDevice, &center, hBrush);
	}

	// Text position X
	textFormat = 0;
	textPosition = { 0, 0 };
	switch (boxType)
	{
	case BoxFirstPointA:
		textPosition.x = g_selection.left + g_zoomScale / 2 - g_zoomScale / 2;
		textPosition.y = g_selection.top + g_zoomScale * ZOOMHEIGHT / 2 - g_zoomScale / 2;
		textFormat = DT_CENTER;
		_snwprintf_s(strData, MAXSTRDATAZOOM, _TRUNCATE, L"%d", g_selection.left);
		break;
	case BoxFinalPointA:
		if (g_selection.right >= g_selection.left)
		{
			textPosition.x = g_selection.left;
		}
		else
		{
			textPosition.x = g_selection.left + 1;
			textFormat += DT_RIGHT;
		}
		if (g_selection.bottom >= g_selection.top)
		{
			textPosition.y = g_selection.top;
			textFormat += DT_BOTTOM;
		}
		else
		{
			textPosition.y = g_selection.top + 2;
		}
		_snwprintf_s(strData, MAXSTRDATAZOOM, _TRUNCATE, L"%d", g_selection.left);
		break;
	case BoxFinalPointB:
		if (g_selection.right < g_selection.left)
		{
			textPosition.x = g_selection.right;
		}
		else
		{
			textPosition.x = g_selection.right + 1;
			textFormat += DT_RIGHT;
		}
		if (g_selection.bottom < g_selection.top)
		{
			textPosition.y = g_selection.bottom;
			textFormat += DT_BOTTOM;
		}
		else
		{
			textPosition.y = g_selection.bottom + 2;
		}
		_snwprintf_s(strData, MAXSTRDATAZOOM, _TRUNCATE, L"%d", g_selection.right);
		break;
	}

	if ((textFormat & DT_BOTTOM) == DT_BOTTOM)
	{
		rectText.bottom = textPosition.y;
	}
	else
	{
		rectText.top = textPosition.y;
	}
	if ((textFormat & DT_RIGHT) == DT_RIGHT) {
		rectText.right = textPosition.x;
	}
	else
	{
		rectText.left = textPosition.x;
	}
	if ((textFormat & DT_CENTER) == DT_CENTER) {
		rectText.left = textPosition.x;
		rectText.right = rectText.left;
	}

	SetTextColor(hdcMemoryDevice, g_useAlternativeColors ? ALTAPPCOLORINV : APPCOLORINV);
	DrawText(hdcMemoryDevice, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP | textFormat);

	// Text for zoom scale
	textFormat = 0;
	textPosition = { 0, 0 };
	_snwprintf_s(strData, MAXSTRDATAZOOM, _TRUNCATE, L"%dx", g_zoomScale);

	switch (boxType)
	{
	case BoxFirstPointA:
		textPosition.x = g_selection.left - g_zoomScale * ZOOMWIDTH / 2;
		textPosition.y = g_selection.top - g_zoomScale * ZOOMHEIGHT / 2 - 1;
		break;
	case BoxFinalPointA:
		if (g_selection.right >= g_selection.left)
		{
			textPosition.x = g_selection.left + g_zoomScale * ZOOMWIDTH - 1;
			textFormat += DT_RIGHT;
		}
		else
		{
			textPosition.x = g_selection.left - g_zoomScale * ZOOMWIDTH + 2;
		}
		if (g_selection.bottom >= g_selection.top)
		{
			textPosition.y = g_selection.top + g_zoomScale * ZOOMHEIGHT;
			textFormat += DT_BOTTOM;
		}
		else
		{
			textPosition.y = g_selection.top - g_zoomScale * ZOOMHEIGHT + 1;
		}
		break;
	case BoxFinalPointB:
		if (g_selection.right < g_selection.left)
		{
			textPosition.x = g_selection.right + g_zoomScale * ZOOMWIDTH - 1;
			textFormat += DT_RIGHT;
		}
		else
		{
			textPosition.x = g_selection.right - g_zoomScale * ZOOMWIDTH + 2;
		}
		if (g_selection.bottom < g_selection.top)
		{
			textPosition.y = g_selection.bottom + g_zoomScale * ZOOMHEIGHT;
			textFormat += DT_BOTTOM;
		}
		else
		{
			textPosition.y = g_selection.bottom - g_zoomScale * ZOOMHEIGHT + 1;
		}
		break;
	}

	if ((textFormat & DT_BOTTOM) == DT_BOTTOM)
	{
		rectText.bottom = textPosition.y;
	}
	else
	{
		rectText.top = textPosition.y;
	}
	if ((textFormat & DT_RIGHT) == DT_RIGHT) {
		rectText.right = textPosition.x;
	}
	else
	{
		rectText.left = textPosition.x;
	}
	if ((textFormat & DT_CENTER) == DT_CENTER) {
		rectText.left = textPosition.x;
		rectText.right = rectText.left;
	}

	SetBkMode(hdcMemoryDevice, TRANSPARENT);
	SetTextColor(hdcMemoryDevice, g_useAlternativeColors ? ALTAPPCOLOR : APPCOLOR);
	if (g_zoomScale > 1) DrawText(hdcMemoryDevice, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP | textFormat);

	// Pointer selection
	SetTextColor(hdcMemoryDevice, g_useAlternativeColors ? ALTAPPCOLORINV : APPCOLORINV);
	SetBkColor(hdcMemoryDevice, g_useAlternativeColors ? ALTAPPCOLOR : APPCOLOR);
	SetBkMode(hdcMemoryDevice, OPAQUE);

	_snwprintf_s(strData, MAXSTRDATAZOOM, _TRUNCATE, L"");
	textFormat = 0;
	switch (boxType)
	{
	case BoxFirstPointA:
		if (((GetTickCount64() / 1000) & 1) && (g_zoomScale > 1))
		{

			textFormat += DT_RIGHT;
			rectText.right = g_selection.left - g_zoomScale * ZOOMWIDTH / 2 - g_zoomScale / 2 - 2;
			rectText.top = g_selection.top - g_zoomScale * ZOOMHEIGHT / 2 - g_zoomScale / 2 - 1;
			_snwprintf_s(strData, MAXSTRDATAZOOM, _TRUNCATE, L"A");
		}
		break;
	case BoxFinalPointA:
		if (((g_appState != statePointA) || ((GetTickCount64() / 1000) & 1)) && (g_zoomScale > 1))
		{
			if (g_selection.right >= g_selection.left)
			{
				rectText.left = g_selection.left + g_zoomScale * ZOOMWIDTH + 2;
			}
			else
			{
				rectText.right = g_selection.left - g_zoomScale * ZOOMWIDTH - 1;
				textFormat = DT_RIGHT;
			}
			if (g_selection.bottom >= g_selection.top)
			{
				rectText.bottom = g_selection.top + g_zoomScale * ZOOMHEIGHT + 1;
				textFormat += DT_BOTTOM;
			}
			else
			{
				rectText.top = g_selection.top - g_zoomScale * ZOOMHEIGHT;
			}
			_snwprintf_s(strData, MAXSTRDATAZOOM, _TRUNCATE, L"A");
		}
		break;
	case BoxFinalPointB:
		if (((g_appState != statePointB) || ((GetTickCount64() / 1000) & 1)) && (g_zoomScale > 1))
		{
			if (g_selection.right < g_selection.left)
			{
				rectText.left = g_selection.right + g_zoomScale * ZOOMWIDTH + 1;
			}
			else
			{
				rectText.right = g_selection.right - g_zoomScale * ZOOMWIDTH - 1;
				textFormat += DT_RIGHT;
			}
			if (g_selection.bottom < g_selection.top) {
				rectText.bottom = g_selection.bottom + g_zoomScale * ZOOMHEIGHT + 1;
				textFormat += DT_BOTTOM;
			}
			else
			{
				rectText.top = g_selection.bottom - g_zoomScale * ZOOMHEIGHT;
			}
			_snwprintf_s(strData, MAXSTRDATAZOOM, _TRUNCATE, L"B");
		}
		break;
	}

	if (wcslen(strData) > 0) DrawText(hdcMemoryDevice, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP | textFormat);

	// Text position Y

	// Text rotated 90 degree
	SelectObject(hdc, hfntPrev);
	DeleteObject(hfnt);
	plf->lfEscapement = 900; // 90 degree, does not work with lfFaceName "System"
	hfnt = CreateFontIndirect(plf);
	if (hfnt == NULL) goto CLEANUP;

	hfntPrev = SelectObject(hdcMemoryDevice, hfnt);

	textFormat = 0;
	textPosition = { 0, 0 };

	switch (boxType)
	{
	case BoxFirstPointA:
	case BoxFinalPointA:
		_snwprintf_s(strData, MAXSTRDATAZOOM, _TRUNCATE, L"%d", g_selection.top);
		break;
	case BoxFinalPointB:
		_snwprintf_s(strData, MAXSTRDATAZOOM, _TRUNCATE, L"%d", g_selection.bottom);
		break;
	}
	// DT_CALCRECT does not like lfEscapement != 0 => Calculate position later
	DrawText(hdcMemoryDevice, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP | DT_CALCRECT);

	switch (boxType)
	{
	case BoxFirstPointA:
		textPosition.x = g_selection.left + g_zoomScale * ZOOMWIDTH / 2 + 1 - g_zoomScale / 2;
		textPosition.y = g_selection.top + (rectText.right - rectText.left + 1) / 2 - 1;
		break;
	case BoxFinalPointA:
		if (g_selection.right >= g_selection.left)
		{
			textPosition.x = g_selection.left - (rectText.bottom - rectText.top + 1) - 1;
		}
		else
		{
			textPosition.x = g_selection.left + 3;
		}
		if (g_selection.bottom >= g_selection.top)
		{
			textPosition.y = g_selection.top + (rectText.right - rectText.left + 1) - 2;
		}
		else
		{
			textPosition.y = g_selection.top + 1;
		}
		break;
	case BoxFinalPointB:
		if (g_selection.right < g_selection.left)
		{
			textPosition.x = g_selection.right - (rectText.bottom - rectText.top + 1) - 1;
		}
		else
		{
			textPosition.x = g_selection.right + 3;
		}
		if (g_selection.bottom < g_selection.top)
		{
			textPosition.y = g_selection.bottom + (rectText.right - rectText.left + 1) - 2;
		}
		else
		{
			textPosition.y = g_selection.bottom + 1;
		}
		break;
	}

	rectText.left = textPosition.x;
	rectText.top = textPosition.y;

	DrawText(hdcMemoryDevice, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP);

CLEANUP:
	// Free resources/Cleanup
	if (hfntPrev != NULL) SelectObject(hdcMemoryDevice, hfntPrev);

	if (hfnt != NULL) DeleteObject(hfnt);
	if (plf != NULL) LocalFree((LOCALHANDLE)plf);
	if (hBrush != NULL) DeleteObject(hBrush);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: OnMouseMove

  Summary:   Mouse position was moved

  Args:     HWND hWindow
			  Handle to window
			int pixelX
			  Mouse x position
			int pixelY
			  Mouse y position
			DWORD flags

  Returns:

-----------------------------------------------------------------F-F*/
void OnMouseMove(HWND hWindow, int pixelX, int pixelY, DWORD flags) {
	static int lastPixelX = 0xffff;
	static int lastPixelY = 0xffff;

	if ((lastPixelX == pixelX) && (lastPixelY == pixelY)) return;

	switch (g_appState) {
	case stateFirstPoint:
	case statePointA:
		g_selection.left = limitXtoBitmap(pixelX);
		g_selection.top = limitYtoBitmap(pixelY);
		InvalidateRect(hWindow, NULL, TRUE);
		break;
	case statePointB:
	{
		g_selection.right = limitXtoBitmap(pixelX);
		g_selection.bottom = limitYtoBitmap(pixelY);
		InvalidateRect(hWindow, NULL, TRUE);
		break;
	}
	}

	lastPixelX = pixelX;
	lastPixelY = pixelY;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: OnPaint

  Summary:   Redraw main window

  Args:     HWND hWindow
			  Handle to window

  Returns:

-----------------------------------------------------------------F-F*/
void OnPaint(HWND hWindow) {
	PAINTSTRUCT ps;
	RECT rect;
	HDC hdcMemScreenshot = NULL;
	HBITMAP bmBuffer = NULL;
	int iBackupDC = 0;
	HBRUSH hBrushForeground = NULL;
	HBRUSH hBrushBackground = NULL;
#define MAXSTRDATA 128
	wchar_t strData[MAXSTRDATA];
	PLOGFONT plf = (PLOGFONT)LocalAlloc(LPTR, sizeof(LOGFONT));
	HGDIOBJ hfnt = NULL, hfntPrev = NULL;
	RECT inner, outer;
	RECT rectText{ 0, 0, 0, 0 };
	HDC hdc = NULL;
	HDC hdcMemoryDevice = NULL;
	int iWidth = 0;
	int iHeight = 0;
	UINT textFormat = 0;
	std::wstring sDisplayInfos;
	COLORREF color;
	HGDIOBJ hOldGDIObj = NULL;

	hdc = BeginPaint(hWindow, &ps);
	if (hdc == NULL) goto CLEANUP;
	if (plf == NULL) goto CLEANUP;
	if (g_appState == stateTrayIcon) goto CLEANUP;
	if (g_hBitmap == NULL) goto CLEANUP;

	GetClientRect(hWindow, &rect);

	iWidth = rect.right + 1;
	iHeight = rect.bottom + 1;

	// Create buffer memory device (Used as painting target and will be copied after the last painting to the display)
	hdcMemoryDevice = CreateCompatibleDC(hdc);
	if (hdcMemoryDevice == NULL) goto CLEANUP;

	// Bitmap in buffer memory
	bmBuffer = CreateCompatibleBitmap(hdc, iWidth, iHeight);
	if (bmBuffer == NULL) goto CLEANUP;

	// Backup memory device context state
	iBackupDC = SaveDC(hdcMemoryDevice);

	// Font
	// Specify a font typeface name and weight.
	if (_snwprintf_s(plf->lfFaceName, 12, _TRUNCATE, L"%s", DEFAULTFONT) < 0) goto CLEANUP;

	plf->lfWeight = FW_NORMAL;
	hfnt = CreateFontIndirect(plf);
	if (hfnt == NULL) goto CLEANUP;

	hfntPrev = SelectObject(hdcMemoryDevice, hfnt);
	if (hfntPrev == NULL) goto CLEANUP;

	SetTextColor(hdcMemoryDevice, g_useAlternativeColors ? ALTAPPCOLORINV : APPCOLORINV);
	SetBkColor(hdcMemoryDevice, g_useAlternativeColors ? ALTAPPCOLOR : APPCOLOR);

	SelectObject(hdcMemoryDevice, bmBuffer);

	// Get screenshot bitmap information and select bitmap
	BITMAP bm;
	GetObject(g_hBitmap, sizeof(bm), &bm);
	hdcMemScreenshot = CreateCompatibleDC(hdc);
	hOldGDIObj = SelectObject(hdcMemScreenshot, g_hBitmap);
	if (hOldGDIObj == NULL) goto CLEANUP;

	// Use a blen function to get a darker image of the screenshot
	BLENDFUNCTION blendFunc;
	blendFunc.BlendOp = AC_SRC_OVER;
	blendFunc.BlendFlags = 0;
	blendFunc.SourceConstantAlpha = g_useAlternativeColors ? 255 : 50; // Factor to darken the screenshot
	blendFunc.AlphaFormat = 0;

	AlphaBlend(hdcMemoryDevice, 0, 0, bm.bmWidth, bm.bmHeight, hdcMemScreenshot, 0, 0, bm.bmWidth, bm.bmHeight, blendFunc);

	// Restore selected object/Free bitmap
	SelectObject(hdcMemScreenshot, hOldGDIObj);
	hOldGDIObj = NULL;

	hBrushForeground = CreateSolidBrush(g_useAlternativeColors ? ALTAPPCOLOR : APPCOLOR);
	if (hBrushForeground == NULL) goto CLEANUP;
	hBrushBackground = CreateSolidBrush(ALTAPPCOLORINV);
	if (hBrushBackground == NULL) goto CLEANUP;

	// Get inner/outer rects and zoom mouse position
	switch (g_appState)
	{
	case stateFirstPoint:
		inner.left = g_selection.left;
		inner.top = g_selection.top;
		zoomMousePosition(hdc, hdcMemoryDevice, BoxFirstPointA);
		break;
	case statePointA:
	case statePointB:
	{
		if (g_selection.left < g_selection.right)
		{
			inner.left = g_selection.left;
			inner.right = g_selection.right;
		}
		else
		{
			inner.left = g_selection.right;
			inner.right = g_selection.left;
		}
		if (g_selection.top < g_selection.bottom)
		{
			inner.top = g_selection.top;
			inner.bottom = g_selection.bottom;
		}
		else
		{
			inner.top = g_selection.bottom;
			inner.bottom = g_selection.top;
		}
		if (inner.left < 0) inner.left = 0;
		if (inner.right > bm.bmWidth - 1) inner.right = bm.bmWidth - 1;
		if (inner.top < 0) inner.top = 0;
		if (inner.bottom > bm.bmHeight - 1) inner.bottom = bm.bmHeight - 1;

		outer.left = inner.left - 1;
		outer.right = inner.right + 1 + 1; // +1 because GDI the second edge of a rect is not part of the drawn rectangle
		outer.top = inner.top - 1;
		outer.bottom = inner.bottom + 1 + 1; // +1 because GDI the second edge of a rect is not part of the drawn rectangle

		// Show selected area of the darkended background of the screenshot
		HDC hdcMem = CreateCompatibleDC(hdc);
		HGDIOBJ hbmOld = SelectObject(hdcMem, g_hBitmap);
		SetStretchBltMode(hdcMemoryDevice, COLORONCOLOR);
		BitBlt(hdcMemoryDevice, inner.left, inner.top, inner.right - inner.left + 1, inner.bottom - inner.top + 1,
			hdcMem, inner.left, inner.top, SRCCOPY);
		SelectObject(hdcMem, hbmOld);
		DeleteDC(hdcMem);

		// Draw frame
		FrameRect(hdcMemoryDevice, &outer, hBrushForeground);

		// Draw text for selection width
		rectText = { 0, 0, 0, 0 };
		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"%d", inner.right - inner.left + 1);
		DrawText(hdcMemoryDevice, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP | DT_CALCRECT);

		if (inner.top >= (rectText.bottom - rectText.top + 1))
		{ // Enough space for text
			rectText.left = outer.left;
			rectText.right = outer.right;
			rectText.bottom = outer.top;
		}
		else
		{ // Not enough space above frame => go beyond frame
			rectText.left = outer.left;
			rectText.right = outer.right;
			rectText.bottom = outer.top + (rectText.bottom - rectText.top + 1);
		}

		// Text color/background
		SetTextColor(hdcMemoryDevice, g_useAlternativeColors ? ALTAPPCOLORINV : APPCOLORINV);
		SetBkColor(hdcMemoryDevice, g_useAlternativeColors ? ALTAPPCOLOR : APPCOLOR);

		// Draw text, when enough splace
		if (abs(g_selection.right - g_selection.left) >= ZOOMWIDTH * g_zoomScale)
			DrawText(hdcMemoryDevice, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP | DT_CENTER | DT_BOTTOM);

		zoomMousePosition(hdc, hdcMemoryDevice, BoxFinalPointA);
		zoomMousePosition(hdc, hdcMemoryDevice, BoxFinalPointB);

		// Draw text for selection height
		rectText = { 0, 0, 0, 0 };

		// Text rotated 90 degree
		SelectObject(hdc, hfntPrev);
		DeleteObject(hfnt);
		plf->lfEscapement = 900; // 90 degree, does not work with lfFaceName "System"
		hfnt = CreateFontIndirect(plf);
		if (hfnt == NULL) goto CLEANUP;

		hfntPrev = SelectObject(hdcMemoryDevice, hfnt);
		if (hfntPrev == NULL) goto CLEANUP;

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"%d", inner.bottom - inner.top + 1);
		DrawText(hdcMemoryDevice, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP | DT_CALCRECT);

		// DT_CALCRECT does not like lfEscapement != 0 => Calculate position later
		int dX = rectText.right - rectText.left + 1;
		int dY = rectText.bottom - rectText.top + 1;

		if (bm.bmWidth - inner.right >= (rectText.bottom - rectText.top + 1))
		{ // Enough space for text
			rectText.left = outer.right;
			rectText.top = (outer.bottom + outer.top + dX) / 2;
		}
		else
		{ // Not enough space right of frame => use area left of frame
			rectText.left = outer.right - dY;
			rectText.top = (outer.bottom + outer.top + dX) / 2;
		}

		// Text color/background
		SetTextColor(hdcMemoryDevice, g_useAlternativeColors ? ALTAPPCOLORINV : APPCOLORINV);
		SetBkColor(hdcMemoryDevice, g_useAlternativeColors ? ALTAPPCOLOR : APPCOLOR);

		// Draw text, when enough space
		if (abs(g_selection.bottom - g_selection.top) >= ZOOMHEIGHT * g_zoomScale)
			DrawText(hdcMemoryDevice, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP);

		break;
	}
	default:
		OutputDebugString(L"Invalid appState");
		goto CLEANUP;
	}

	// Draw information
	if (g_displayInternallnformation)
	{
		HBRUSH hBrushDisplayForeground = (g_useAlternativeColors ? hBrushBackground : hBrushForeground);
		HBRUSH hBrushDisplayBackground = (g_useAlternativeColors ? hBrushForeground : hBrushBackground);

		RECT rectTextArea = { 0, 0, 0, 0 };
		// Text rotated 0 degree
		SelectObject(hdc, hfntPrev);
		DeleteObject(hfnt);
		plf->lfEscapement = 0;
		hfnt = CreateFontIndirect(plf);
		if (hfnt == NULL) goto CLEANUP;

		hfntPrev = SelectObject(hdcMemoryDevice, hfnt);
		if (hfntPrev == NULL) goto CLEANUP;

		SetBkMode(hdcMemoryDevice, TRANSPARENT);
		if (!g_useAlternativeColors)
			SetTextColor(hdcMemoryDevice, APPCOLOR);
		else
			SetTextColor(hdcMemoryDevice, ALTAPPCOLORINV);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Virtual desktop [%d,%d] %dx%d", GetSystemMetrics(SM_XVIRTUALSCREEN), GetSystemMetrics(SM_YVIRTUALSCREEN), GetSystemMetrics(SM_CXVIRTUALSCREEN), GetSystemMetrics(SM_CYVIRTUALSCREEN));
		sDisplayInfos.assign(strData);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Selection [%d,%d] [%d,%d]", g_selection.left, g_selection.top, g_selection.right, g_selection.bottom);
		sDisplayInfos.append(L"\n").append(strData);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Stored selection [%d,%d] [%d,%d]", g_storedSelection.left, g_storedSelection.top, g_storedSelection.right, g_storedSelection.bottom);
		sDisplayInfos.append(L"\n").append(strData);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Bitmap %dx%d", bm.bmWidth, bm.bmHeight);
		sDisplayInfos.append(L"\n").append(strData);

		POINT mouse;
		GetCursorPos(&mouse);
		hOldGDIObj = SelectObject(hdcMemScreenshot, g_hBitmap);
		if (hOldGDIObj == NULL) goto CLEANUP;
		color = GetPixel(hdcMemScreenshot, mouse.x - g_appWindowPos.x, mouse.y - g_appWindowPos.y);
		SelectObject(hdcMemScreenshot, hOldGDIObj);
		hOldGDIObj = NULL;

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Mouse [%d,%d] RGB %d,%d,%d", mouse.x, mouse.y, GetRValue(color),GetGValue(color),GetBValue(color));
		sDisplayInfos.append(L"\n").append(strData);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Save to file %s", g_saveToFile ? L"On" : L"Off");
		sDisplayInfos.append(L"\n").append(strData);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Save to clipboard %s", g_saveToClipboard ? L"On" : L"Off");
		sDisplayInfos.append(L"\n").append(strData);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Alternative colors %s", g_useAlternativeColors ? L"On" : L"Off");
		sDisplayInfos.append(L"\n").append(strData);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"State %d appWindow [%d,%d] Selected Monitor %d", g_appState, g_appWindowPos.x, g_appWindowPos.y, g_selectedMonitor);
		sDisplayInfos.append(L"\n").append(strData);
		for (int i = 0; i < (int) g_rectMonitor.size(); i++)
		{
			_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Monitor %d [%d,%d] [%d,%d]", i, g_rectMonitor[i].left, g_rectMonitor[i].top, g_rectMonitor[i].right, g_rectMonitor[i].bottom);
			sDisplayInfos.append(L"\n").append(strData);
		}
		sDisplayInfos.append(L"\n\nA = Select all\nM = Select next monitor\nTab = A <-> B\nCursor keys = Move A/B\n")
			.append(L"Alt+cursor keys = Fast move A/B\nShift+cursor keys = Find color change for A/B\nReturn = OK\nESC = Cancel")
			.append(L"\n+/- = Increase/decrease selection")
			.append(L"\nPageUp/PageDown, mouse wheel = Zoom In/Out")
			.append(L"\nInsert = Store selection\nHome = Restore selection")
			.append(L"\nC = Clipboard On/Off\nF = File On/Off\nS = Alternative colors On/Off")
			.append(L"\nF1 = Display information On/Off");

		// Calc text area
		DrawText(hdcMemoryDevice, sDisplayInfos.c_str(), -1, &rectTextArea, DT_NOCLIP | DT_CALCRECT );

		int height = (rectTextArea.bottom - rectTextArea.top + 1);
		int width = (rectTextArea.right - rectTextArea.left + 1) + 1;

		rectTextArea.left = 10;
		rectTextArea.top = 10;
		rectTextArea.bottom = rectTextArea.top + height - 1;
		rectTextArea.right = rectTextArea.left + width - 1;

		textFormat = 0;
		POINT pos = { rectTextArea.left, rectTextArea.top };
		HMONITOR hMonitor = MonitorFromPoint(pos, MONITOR_DEFAULTTONULL);
		if (hMonitor != NULL)
		{
			MONITORINFO mi;
			mi.cbSize = sizeof(mi);
			if (GetMonitorInfo(hMonitor, &mi)) {
				switch (g_appState)
				{
				case stateFirstPoint:
				case statePointA:
					if (g_selection.left < mi.rcMonitor.right / 2) {
						textFormat = DT_RIGHT;
						rectTextArea.left = mi.rcMonitor.right - width - 10;
						rectTextArea.right = rectTextArea.left + width + 1;
					}
					break;
				case statePointB:
					if (g_selection.right < mi.rcMonitor.right / 2) {
						textFormat = DT_RIGHT;
						rectTextArea.left = mi.rcMonitor.right - width - 10;
						rectTextArea.right = rectTextArea.left + width + 1;
					}
					break;
				}
			}
		}

		if (g_useAlternativeColors) FillRect(hdcMemoryDevice, &rectTextArea, hBrushDisplayBackground); // Text area background

		// Draw text in text area
		rectText.left = rectTextArea.left + 1;
		rectText.right = rectTextArea.right - 1;
		rectText.top = rectTextArea.top;
		DrawText(hdcMemoryDevice, sDisplayInfos.c_str(), -1, &rectText, DT_NOCLIP | textFormat);

		// Draw monitor layout
		float scale = (float) (rectTextArea.right - rectTextArea.left) / GetSystemMetrics(SM_CXVIRTUALSCREEN);
		RECT virtualDesktop;
		virtualDesktop.left = rectTextArea.left;
		virtualDesktop.top = rectTextArea.bottom + 10;
		virtualDesktop.right = virtualDesktop.left + (LONG) round((GetSystemMetrics(SM_CXVIRTUALSCREEN) -1) * scale) + 1;
		virtualDesktop.bottom = virtualDesktop.top + (LONG) round((GetSystemMetrics(SM_CYVIRTUALSCREEN) -1) * scale) + 1;
		FrameRect(hdcMemoryDevice, &virtualDesktop, hBrushDisplayForeground); // Virtual desktop frame

		if (g_useAlternativeColors)
		{
			FillRect(hdcMemoryDevice, &virtualDesktop, hBrushDisplayBackground); // Background fill
		}

		for (DWORD i = 0; i < g_rectMonitor.size(); i++)
		{
			RECT monitor;

			monitor.left = virtualDesktop.left + (LONG) round((g_rectMonitor[i].left - g_appWindowPos.x) * scale);
			monitor.top = virtualDesktop.top + (LONG) round((g_rectMonitor[i].top - g_appWindowPos.y) * scale);
			monitor.right = monitor.left + (LONG) round((g_rectMonitor[i].right-g_rectMonitor[i].left - 1) * scale) + 1;
			monitor.bottom = monitor.top + (LONG) round((g_rectMonitor[i].bottom - g_rectMonitor[i].top -1) * scale) + 1;

			FrameRect(hdcMemoryDevice, &monitor, hBrushDisplayForeground);

			_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"%d", i);
			DrawText(hdcMemoryDevice, strData, -1, &monitor, DT_SINGLELINE | DT_NOCLIP | DT_CENTER | DT_VCENTER);

		}
		if ((g_selection.left != -1) && (g_selection.right != -1) && (g_selection.top != -1) && (g_selection.bottom != -1)) {
			RECT selection;
			if (g_selection.right >= g_selection.left) {
				selection.left = virtualDesktop.left + (LONG) round(g_selection.left * scale);
				selection.right = virtualDesktop.left + (LONG) round(g_selection.right * scale) + 1;
			}
			else
			{
				selection.left = virtualDesktop.left + (LONG) round(g_selection.right * scale);
				selection.right = virtualDesktop.left + (LONG) round(g_selection.left * scale) + 1;
			}
			if (g_selection.bottom >= g_selection.top) {
				selection.top = virtualDesktop.top + (LONG) round(g_selection.top * scale);
				selection.bottom = virtualDesktop.top + (LONG) round(g_selection.bottom * scale) + 1;
			}
			else
			{
				selection.top = virtualDesktop.top + (LONG) round(g_selection.bottom * scale);
				selection.bottom = virtualDesktop.top + (LONG) round(g_selection.top * scale) + 1;
			}

			FrameRect(hdcMemoryDevice, &selection, hBrushDisplayForeground);
		}
		else
		{
			if  ((g_selection.left != -1) && (g_selection.top != -1) ) { // Quick and dirty "pixel" at mouse cursor
				RECT pixel;
				pixel.left = virtualDesktop.left + (LONG) round(g_selection.left * scale)-1;
				pixel.top = virtualDesktop.top + (LONG) round(g_selection.top * scale)-1;
				pixel.right = pixel.left+3;
				pixel.bottom = pixel.top+3;
				FrameRect(hdcMemoryDevice, &pixel, hBrushDisplayForeground);
			}
		}
	}

	// Copy memory buffer to display
	BitBlt(hdc, 0, 0, iWidth, iHeight, hdcMemoryDevice, 0, 0, SRCCOPY);

CLEANUP:
	// Free resources/Cleanup
	if (hOldGDIObj != NULL) SelectObject(hdcMemScreenshot, hOldGDIObj);
	if (hdcMemScreenshot != NULL) DeleteDC(hdcMemScreenshot);
	if (hBrushForeground != NULL) DeleteObject(hBrushForeground);
	if (hBrushBackground != NULL) DeleteObject(hBrushBackground);
	if (hdcMemoryDevice != NULL)
	{
		// Restore memory device context statte
		if (iBackupDC != 0) RestoreDC(hdcMemoryDevice, iBackupDC);
		if (hfntPrev != NULL) SelectObject(hdcMemoryDevice, hfntPrev);
	}
	if (hfnt != NULL) DeleteObject(hfnt);
	if (plf != NULL) LocalFree((LOCALHANDLE)plf);

	if (bmBuffer != NULL) DeleteObject(bmBuffer);
	if (hdcMemoryDevice != NULL) DeleteDC(hdcMemoryDevice);

	EndPaint(hWindow, &ps);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: resizeSelection

  Summary:   Increase/decrease the selection dependent on the step size

  Args:     HWND hWindow
			  Handle to window
			int stepSize
			  Step size in pixel

  Returns:

-----------------------------------------------------------------F-F*/
void resizeSelection(HWND hWindow, int stepSize)
{
	if ((g_appState != statePointA) && (g_appState != statePointB)) return;

	if ((stepSize < 0) && (abs(g_selection.right- g_selection.left) < abs(stepSize*2)))
	{
		// Width too small for decreasing by stepsize
		g_selection.left = limitXtoBitmap((g_selection.right + g_selection.left)/2);
		g_selection.right = g_selection.left;
	}
	else
	{

		if (g_selection.left <= g_selection.right)
		{
			g_selection.left = limitXtoBitmap(g_selection.left-stepSize);
			g_selection.right = limitXtoBitmap(g_selection.right+stepSize);
		}
		else
		{
			g_selection.left = limitXtoBitmap(g_selection.left+stepSize);
			g_selection.right = limitXtoBitmap(g_selection.right-stepSize);
		}
	}

	if ((stepSize < 0) && (abs(g_selection.top- g_selection.bottom) < abs(stepSize*2)))
	{
		// Height too small for decreasing by stepsize
		g_selection.top = limitYtoBitmap((g_selection.top + g_selection.top)/2);
		g_selection.bottom = g_selection.top;
	}
	else
	{
		if (g_selection.top <= g_selection.bottom)
		{
			g_selection.top = limitYtoBitmap(g_selection.top-stepSize);
			g_selection.bottom = limitYtoBitmap(g_selection.bottom+stepSize);
		}
		else
		{
			g_selection.top = limitYtoBitmap(g_selection.top+stepSize);
			g_selection.bottom = limitYtoBitmap(g_selection.bottom-stepSize);
		}
	}

	if (g_appState == statePointA) MySetCursorPos(g_selection.left, g_selection.top);
	if (g_appState == statePointB) MySetCursorPos(g_selection.right, g_selection.bottom);

	InvalidateRect(hWindow, NULL, TRUE);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: setBeforeColorChange

  Summary:   Set x/y position in screenshot to the pixel before the next color change
             (direction depends on virtual key code)

  Args:     WPARAM virtualKeyCode
			  Virtual key code
			LONG &x
			  x start position (call by ref)
			LONG &y
			  y start position (call by ref)

  Returns:

-----------------------------------------------------------------F-F*/
void setBeforeColorChange(WPARAM virtualKeyCode, LONG &x, LONG &y)
{
	BITMAP bm;
    int directionX = 0;
    int directionY = 0;
	COLORREF referenceColor;

	if (g_hBitmap == NULL) return;

	GetObject(g_hBitmap, sizeof(bm), &bm);
    HDC hdcMem = CreateCompatibleDC(NULL);
    if (hdcMem == NULL) return;
    HBITMAP hbmOld = (HBITMAP)SelectObject(hdcMem, g_hBitmap);
	if (hbmOld == NULL) goto CLEANUP;

    referenceColor = GetPixel(hdcMem, x, y);
    switch (virtualKeyCode)
    {
    	case VK_UP:
    		directionY = -1;
    		break;
    	case VK_DOWN:
    		directionY = 1;
    		break;
    	case VK_LEFT:
    		directionX = -1;
    		break;
    	case VK_RIGHT:
    		directionX = 1;
    		break;
    	default:
 			OutputDebugString(L"Invalid wParam");
			goto CLEANUP;
	}

	while (true) {
		if (GetPixel(hdcMem, x + directionX, y + directionY) != referenceColor) break;

		if (x+directionX < 0) break;
		if (x+directionX > bm.bmWidth-1) break;
		if (y+directionY < 0) break;
		if (y+directionY > bm.bmHeight-1) break;

		x = limitXtoBitmap(x+directionX);
		y = limitYtoBitmap(y+directionY);
	}

CLEANUP:
    if (hbmOld != NULL) SelectObject(hdcMem, hbmOld);
    if (hdcMem != NULL) DeleteDC(hdcMem);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: checkCursorButtons

  Summary:   Move point A/B by cursor buttons

  Args:     HWND hWindow
			  Handle to window
			WPARAM wParam
			int step
			  Step size in pixels

  Returns:

-----------------------------------------------------------------F-F*/
void checkCursorButtons(HWND hWindow, WPARAM wParam, int step)
{
	switch (wParam)
	{
	case VK_UP:
		switch (g_appState)
		{
		case stateFirstPoint:
		case statePointA:
			if (GetAsyncKeyState(VK_SHIFT) & 0x8000) // Shift pressed?
			{
				setBeforeColorChange(wParam, g_selection.left, g_selection.top);
				MySetCursorPos(g_selection.left, g_selection.top);
			}
			else
			{
				g_selection.top = limitYtoBitmap(g_selection.top - step);
				MySetCursorPos(g_selection.left, g_selection.top);
			}
			break;
		case statePointB:
			if (GetAsyncKeyState(VK_SHIFT) & 0x8000)
			{
				setBeforeColorChange(wParam, g_selection.right, g_selection.bottom);
				MySetCursorPos(g_selection.right, g_selection.bottom);
			}
			else
			{
				g_selection.bottom = limitYtoBitmap(g_selection.bottom - step);
				MySetCursorPos(g_selection.right, g_selection.bottom);
			}
			break;
		}
		InvalidateRect(hWindow, NULL, TRUE);
		break;
	case VK_DOWN:
		switch (g_appState)
		{
		case stateFirstPoint:
		case statePointA:
			if (GetAsyncKeyState(VK_SHIFT) & 0x8000)
			{
				setBeforeColorChange(wParam, g_selection.left, g_selection.top);
				MySetCursorPos(g_selection.left, g_selection.top);
			}
			else
			{
				g_selection.top = limitYtoBitmap(g_selection.top + step);
				MySetCursorPos(g_selection.left, g_selection.top);
			}
			break;
		case statePointB:
			if (GetAsyncKeyState(VK_SHIFT) & 0x8000)
			{
				setBeforeColorChange(wParam, g_selection.right, g_selection.bottom);
				MySetCursorPos(g_selection.right, g_selection.bottom);
			}
			else
			{
				g_selection.bottom = limitYtoBitmap(g_selection.bottom + step);
				MySetCursorPos(g_selection.right, g_selection.bottom);
			}
			break;
		}
		InvalidateRect(hWindow, NULL, TRUE);
		break;
	case VK_LEFT:
		switch (g_appState)
		{
		case stateFirstPoint:
		case statePointA:
			if (GetAsyncKeyState(VK_SHIFT) & 0x8000)
			{
				setBeforeColorChange(wParam, g_selection.left, g_selection.top);
				MySetCursorPos(g_selection.left, g_selection.top);
			}
			else
			{
				g_selection.left = limitXtoBitmap(g_selection.left - step);
				MySetCursorPos(g_selection.left, g_selection.top);
			}
			break;
		case statePointB:
			if (GetAsyncKeyState(VK_SHIFT) & 0x8000)
			{
				setBeforeColorChange(wParam, g_selection.right, g_selection.bottom);
				MySetCursorPos(g_selection.right, g_selection.bottom);
			}
			else
			{
				g_selection.right = limitXtoBitmap(g_selection.right - step);
				MySetCursorPos(g_selection.right, g_selection.bottom);
			}
			break;
		}
		InvalidateRect(hWindow, NULL, TRUE);
		break;
	case VK_RIGHT:
		switch (g_appState)
		{
		case stateFirstPoint:
		case statePointA:
			if (GetAsyncKeyState(VK_SHIFT) & 0x8000)
			{
				setBeforeColorChange(wParam, g_selection.left, g_selection.top);
				MySetCursorPos(g_selection.left, g_selection.top);
			}
			else
			{
				g_selection.left = limitXtoBitmap(g_selection.left + step);
				MySetCursorPos(g_selection.left, g_selection.top);
			}
			break;
		case statePointB:
			if (GetAsyncKeyState(VK_SHIFT) & 0x8000)
			{
				setBeforeColorChange(wParam, g_selection.right, g_selection.bottom);
				MySetCursorPos(g_selection.right, g_selection.bottom);
			}
			else
			{
				g_selection.right = limitXtoBitmap(g_selection.right + step);
				MySetCursorPos(g_selection.right, g_selection.bottom);
			}
			break;
		}
		InvalidateRect(hWindow, NULL, TRUE);
		break;
	}
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: CaptureScreen

  Summary:   Capture screen to a bitmap

  Args:     HWND hWindow
			  Handle to window

  Returns:

-----------------------------------------------------------------F-F*/
int CaptureScreen(HWND hWindow)
{
	HDC hdcScreen = NULL;
	HDC hdcMemDC = NULL;
	HGDIOBJ hbmOld = NULL;

	int screenX = GetSystemMetrics(SM_XVIRTUALSCREEN);
	int screenY = GetSystemMetrics(SM_YVIRTUALSCREEN);
	int screenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
	int screenHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);

	// Retrieve the handle to a display device context for the client
	// area of the window.
	hdcScreen = GetDC(NULL);

	// Create a compatible DC, which is used in a BitBlt from the window DC.
	hdcMemDC = CreateCompatibleDC(hdcScreen);

	if (!hdcMemDC)
	{
		std::wstring sMessage;
		sMessage.assign(L"CreateCompatibleDC ").append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		MessageBox(hWindow, sMessage.c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONERROR);
		goto CLEANUP;
	}

	if (g_hBitmap != NULL) { // Delete previous screenshot
		DeleteObject(g_hBitmap);
		g_hBitmap = NULL;
	}

	// Create a compatible bitmap from the Window DC.
	g_hBitmap = CreateCompatibleBitmap(hdcScreen, screenWidth, screenHeight);
	// Select the compatible bitmap into the compatible memory DC.
	hbmOld = SelectObject(hdcMemDC, g_hBitmap);

	// Bit block transfer into our compatible memory DC.
	if (!BitBlt(hdcMemDC,
		0, 0,
		screenWidth, screenHeight,
		hdcScreen,
		screenX, screenY,
		SRCCOPY))
	{
		std::wstring sMessage;
		sMessage.assign(L"BitBlt ").append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		MessageBox(hWindow, sMessage.c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONERROR);
		goto CLEANUP;
	}

CLEANUP:
	// Free resources/Clean up
	if (hbmOld != NULL) SelectObject(hdcMemDC, hbmOld);
	if (hdcMemDC != NULL) DeleteDC(hdcMemDC);

	ReleaseDC(NULL, hdcScreen);
	return 0;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: MonitorEnumProc

  Summary:   Callback function for EnumDisplayMonitors to add the virtual screen rectangle coordinates for the monitor to a global vector

  Args:     HMONITOR hMonitor
			  Handle to monitor
			HDC hdcMonitor
			  Handle to device context
			LPRECT lprcMonitor
			  Pointer to a RECT for the virtual screen rectangle coordinates
			LPARAM dwData

  Returns:	BOOL
			  Is always TRUE

-----------------------------------------------------------------F-F*/
BOOL CALLBACK MonitorEnumProc(HMONITOR hMonitor, HDC hdcMonitor, LPRECT lprcMonitor, LPARAM dwData) {
	RECT info;
	info = *lprcMonitor;
	g_rectMonitor.push_back(info);
	return TRUE;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: startCaptureGUI

  Summary:   Capture screen and start main window in fullscreen to select an area

  Args:     HWND hWindow
			  Handle to window

  Returns:

-----------------------------------------------------------------F-F*/
void startCaptureGUI(HWND hWindow) {
	// Store current window style
	long prevStyle = GetWindowLong(hWindow, GWL_EXSTYLE);
	COLORREF crKey;
	BYTE     bAlpha;
	DWORD    dwFlags;

	g_activeWindow = GetForegroundWindow();
	GetLayeredWindowAttributes(hWindow, &crKey, &bAlpha, &dwFlags);

	// Hide windows with alpha, because ShowWindow(hWnd,SW_HIDE) is animated and causes artifacts
	SetWindowLong(hWindow, GWL_EXSTYLE, WS_EX_LAYERED);
	SetLayeredWindowAttributes(hWindow, 0, 0, LWA_ALPHA);

	CaptureScreen(hWindow);

	// Stores monitor coordinates
	g_rectMonitor.clear();
	EnumDisplayMonitors(NULL, NULL, MonitorEnumProc, 0);

	// Restore window style
	SetLayeredWindowAttributes(hWindow, crKey, bAlpha, dwFlags);
	SetWindowLong(hWindow, GWL_EXSTYLE, prevStyle);

	// Refresh settings from registry
	g_zoomScale = getDWORDSettingFromRegistry(defaultZoomScale);
	g_saveToFile = getDWORDSettingFromRegistry(saveToFile);
	g_saveToClipboard = getDWORDSettingFromRegistry(saveToClipboard);
	g_useAlternativeColors = getDWORDSettingFromRegistry(useAlternativeColors);
	g_displayInternallnformation = getDWORDSettingFromRegistry(displayInternallnformation);
	g_storedSelection.left = limitXtoBitmap(getDWORDSettingFromRegistry(storedSelectionLeft));
	g_storedSelection.top = limitYtoBitmap(getDWORDSettingFromRegistry(storedSelectionTop));
	g_storedSelection.right = limitXtoBitmap(getDWORDSettingFromRegistry(storedSelectionRight));
	g_storedSelection.bottom = limitYtoBitmap(getDWORDSettingFromRegistry(storedSelectionBottom));

	enterFullScreen(hWindow);
	ShowWindow(hWindow, SW_NORMAL);
	ShowCursor(false);

	POINT mouse;
	GetCursorPos(&mouse);
	g_appState = stateFirstPoint;
	g_selection.left = limitXtoBitmap(mouse.x - g_appWindowPos.x);
	g_selection.top = limitYtoBitmap(mouse.y - g_appWindowPos.y);
	g_selection.right = -1;
	g_selection.bottom = -1;

	// Enable 1s time to show selected point
	SetTimer(hWindow, IDT_TIMER1000MS, 1000, (TIMERPROC)NULL);

	// Force foreground window (Prevents keyboard input focus problems on a second or third monitor)
	SetForegroundWindowInternal(hWindow);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: MyRegisterClass

  Summary:   Register window class

  Args:     HINSTANCE hInstance
			  Handle to the instance that contains the window procedure

  Returns:  ATOM
			  ID of registered class
			  NULL = error

-----------------------------------------------------------------F-F*/
ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEXW wcex;

	wcex.cbSize = sizeof(WNDCLASSEX);

	wcex.style = CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc = WndProc;
	wcex.cbClsExtra = 0;
	wcex.cbWndExtra = 0;
	wcex.hInstance = hInstance;
	wcex.hIcon = NULL;
	wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
	wcex.lpszMenuName = NULL;
	wcex.lpszClassName = L"MainWndClass";
	wcex.hIconSm = NULL;

	return RegisterClassExW(&wcex);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: wWinMain

  Summary:   Process window messages of main window

  Args:     HINSTANCE hInstance
			HINSTANCE hPrevInstance
			LPWSTR lpCmdLine
			int nCmdShow

  Returns:  int
			  0 = success
			  1 = error

-----------------------------------------------------------------F-F*/
#if defined (__GNUC__)
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR pCmdLine, int nCmdShow)
#else
int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
	_In_opt_ HINSTANCE hPrevInstance,
	_In_ LPWSTR    lpCmdLine,
	_In_ int       nCmdShow)
#endif
{
	HANDLE hMutex = CreateMutex(NULL, TRUE, LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str());

	// Prevents concurrent program starts
	if (GetLastError() == ERROR_ALREADY_EXISTS) {
		OutputDebugString(L"Program already startet");
		return 0;
	}

	MyRegisterClass(hInstance);

	g_hInst = hInstance; // Store instance handle in global variable

	g_hWindow = CreateWindow(L"MainWndClass", LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

	if (!g_hWindow) { // Error
		std::wstring sMessage;
		sMessage.assign(L"CreateWindow ").append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		MessageBox(NULL, sMessage.c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONERROR);
		if (hMutex != NULL)
		{
			ReleaseMutex(hMutex);
			CloseHandle(hMutex);
		}
		return 1;
	}

	// Get settings from registry
	g_saveToFile = getDWORDSettingFromRegistry(saveToFile);
	g_saveToClipboard = getDWORDSettingFromRegistry(saveToClipboard);
	if (!g_saveToClipboard && !g_saveToFile) // Clipboard and file was disabled by user => Warning
	{
		MessageBox(g_hWindow, LoadStringAsWstr(g_hInst, IDS_NOSAVING).c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONWARNING);
	}

	// Add tray icon entry
	g_nid.cbSize = sizeof(NOTIFYICONDATA);
	g_nid.hWnd = g_hWindow;
	g_nid.uID = 1;
	g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
	g_nid.uCallbackMessage = WM_TRAYICON;
	g_nid.hIcon = LoadIcon(GetModuleHandle(NULL), MAKEINTRESOURCE(IDI_ICON));
#define MAXTOOLTIPSIZE 64 // Including null termination
	_snwprintf_s(g_nid.szTip, MAXTOOLTIPSIZE, _TRUNCATE, L"%s", LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str());
	Shell_NotifyIcon(NIM_ADD, &g_nid);

	MSG msg;

	SetHook();

	// Main message loop:
	while (GetMessage(&msg, nullptr, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	ReleaseHook();
	// Remove tray icon entry
	Shell_NotifyIcon(NIM_DELETE, &g_nid);

	// Free Mutex to allow next program start
	if (hMutex != NULL)
	{
		ReleaseMutex(hMutex);
		CloseHandle(hMutex);
	}

	return (int)msg.wParam;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: WndProc

  Summary:   Process window messages of main window

  Args:     HWND hWnd
			  Handle to main window
			UINT message
			  Message
			WPARAM wParam
			LPARAM lParam

  Returns:  LRESULT

-----------------------------------------------------------------F-F*/
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
	case WM_MOUSEWHEEL: {
		int delta = GET_WHEEL_DELTA_WPARAM(wParam);
		if (delta > 0)
			SendMessage(hWnd, WM_ZOOMIN, 0, 0);
		else
			SendMessage(hWnd, WM_ZOOMOUT, 0, 0);
		break;
	}
	case WM_ZOOMIN:
		g_zoomScale++;
		if (g_zoomScale > MAXZOOMFACTOR) g_zoomScale = MAXZOOMFACTOR;
		InvalidateRect(hWnd, NULL, TRUE);
		break;
	case WM_ZOOMOUT:
		g_zoomScale--;
		if (g_zoomScale <= 1) g_zoomScale = 1;
		InvalidateRect(hWnd, NULL, TRUE);
		break;
	case WM_SELECTALL: // Select area over all monitors
	{
		if (g_hBitmap == NULL) break;
		BITMAP bm;
		GetObject(g_hBitmap, sizeof(bm), &bm);
		g_appState = statePointB;
		g_selection.left = limitXtoBitmap(0);
		g_selection.top = limitYtoBitmap(0);
		g_selection.right = limitXtoBitmap(bm.bmWidth - 1);
		g_selection.bottom = limitYtoBitmap(bm.bmHeight - 1);
		InvalidateRect(hWnd, NULL, TRUE);
		// Do not SetCursorPos, because this can make trouble on multimonitor systems with different resolutions
		break;
	}
	case WM_STARTED: // Start new capture
	{
		startCaptureGUI(hWnd);
		break;
	}
	case WM_GOTOTRAY: // Hide window and goto tray icon
	{
		KillTimer(hWnd, IDT_TIMER1000MS);
		ShowCursor(true);
		ShowWindow(hWnd, SW_HIDE);
		g_appState = stateTrayIcon;
		SetActiveWindow(g_activeWindow);
		break;
	}
	case WM_NEXTSTATE: // Enter was pressed or left mouse button was clicked => Goto next state
		if (g_appState == stateFirstPoint) // Set point A
		{
			if (wParam != 0) { // Left mouse was clicked
				g_selection.left = limitXtoBitmap(GET_X_LPARAM(lParam));
				g_selection.top = limitYtoBitmap(GET_Y_LPARAM(lParam));
				g_selection.right = g_selection.left;
				g_selection.bottom = g_selection.top;
			}
			else
			{ // Enter was pressed
				g_selection.right = g_selection.left;
				g_selection.bottom = g_selection.top;
			}
			g_appState = statePointB;
			InvalidateRect(hWnd, NULL, TRUE);
			SetTimer(hWnd, IDT_TIMER1000MS, 1000, (TIMERPROC)NULL);
			// Reset zoom
			g_zoomScale = getDWORDSettingFromRegistry(defaultZoomScale);
		}
		else
		{ // Save selection
			if ((g_appState == statePointA) || (g_appState == statePointB))
			{
				if ((g_selection.left != g_selection.right) && (g_selection.top != g_selection.bottom))
				{
					SendMessage(hWnd, WM_GOTOTRAY, 0, 0);
					saveSelection(hWnd);
				}
			}
		}
		break;
	case WM_TRAYICON: // Tray icon messages
		switch (lParam)
		{
		case WM_RBUTTONUP: // Right click on tray icon => Context menu
		{
			POINT pt;
			GetCursorPos(&pt);
			HMENU hMenu = CreatePopupMenu();
			AppendMenu(hMenu, MF_STRING, IDM_CAPTURE, LoadStringAsWstr(g_hInst, IDS_SCREENSHOT).c_str());
			AppendMenu(hMenu, MF_STRING, IDM_OPENFOLDER, LoadStringAsWstr(g_hInst, IDS_OPENFOLDER).c_str());
			AppendMenu(hMenu, MF_STRING, IDM_SETFOLDER, LoadStringAsWstr(g_hInst, IDS_SETFOLDER).c_str());
			AppendMenu(hMenu, MF_STRING | (g_saveToClipboard ? MF_CHECKED : 0), IDM_SAVETOCLIPBOARD, LoadStringAsWstr(g_hInst, IDS_SAVETOCLIPBOARD).c_str());
			AppendMenu(hMenu, MF_STRING | (g_saveToFile ? MF_CHECKED : 0), IDM_SAVETOFILE, LoadStringAsWstr(g_hInst, IDS_SAVETOFILE).c_str());
			AppendMenu(hMenu, MF_STRING, IDM_ABOUT, LoadStringAsWstr(g_hInst, IDS_ABOUT).c_str());
			AppendMenu(hMenu, MF_SEPARATOR | MF_BYPOSITION, 0, NULL);
			AppendMenu(hMenu, MF_STRING, IDM_EXIT, LoadStringAsWstr(g_hInst, IDS_EXIT).c_str());
			SetForegroundWindow(hWnd);
			TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hWnd, NULL);
			DestroyMenu(hMenu);
			break;
		}
		case WM_LBUTTONDBLCLK: // Double click on tray icon => Open screenshot folder
			SendMessage(hWnd, WM_COMMAND, IDM_OPENFOLDER, 0);
			break;
		}
		break;
	case WM_SYSKEYDOWN:
		if (lParam & (1 << 29)) // Alt pressed
		{
			if (wParam == VK_F4) // Alt + F4
				DestroyWindow(hWnd);
			else
				checkCursorButtons(hWnd, wParam, 10); // Move points with cursor buttons with a big step, if a cursor button is pressed
		}
		break;
	case WM_KEYDOWN:
		checkCursorButtons(hWnd, wParam, 1); // Move points with cursor buttons with a one pixel step, if a cursor button is pressed
		switch (wParam)
		{
		case VK_NEXT: // Page down => Zooom out
			SendMessage(hWnd, WM_ZOOMOUT, 0, 0);
			break;
		case VK_PRIOR: // Page up => Zoom in
			SendMessage(hWnd, WM_ZOOMIN, 0, 0);
			break;
		case 'A': // A => Select all
			SendMessage(hWnd, WM_SELECTALL, 0, 0);
			break;
		case 'M': // M => Select next monitor
			if (g_rectMonitor.size() > 0)
			{
				g_selectedMonitor++;
				if (g_selectedMonitor >= g_rectMonitor.size()) g_selectedMonitor = 0;

				g_appState = statePointB;
				g_selection.left = limitXtoBitmap(g_rectMonitor[g_selectedMonitor].left - g_appWindowPos.x);
				g_selection.top = limitYtoBitmap(g_rectMonitor[g_selectedMonitor].top - g_appWindowPos.y);
				g_selection.right = limitXtoBitmap(g_rectMonitor[g_selectedMonitor].right - g_appWindowPos.x - 1);
				g_selection.bottom = limitYtoBitmap(g_rectMonitor[g_selectedMonitor].bottom - g_appWindowPos.y - 1);
				MySetCursorPos(g_selection.right, g_selection.bottom);
				InvalidateRect(hWnd, NULL, TRUE);
			}
			break;
		case 'C': // C => Toggle save to clipboard
			SendMessage(hWnd, WM_COMMAND, IDM_SAVETOCLIPBOARD, 0);
			InvalidateRect(hWnd, NULL, TRUE);
			break;
		case 'F': // F => Toggle save to file
			SendMessage(hWnd, WM_COMMAND, IDM_SAVETOFILE, 0);
			break;
		case 'S': // S => Toogle colors
			SendMessage(hWnd, WM_COMMAND, IDM_ALTERNATIVECOLORS, 0);
			break;
		case VK_INSERT: // Insert => Store selection
			if ((g_selection.left != -1) && (g_selection.right != -1) && (g_selection.top != -1) && (g_selection.bottom != -1))
			{
				g_storedSelection = g_selection;
				storeDWORDSettingInRegistry(storedSelectionLeft, g_selection.left);
				storeDWORDSettingInRegistry(storedSelectionTop, g_selection.top);
				storeDWORDSettingInRegistry(storedSelectionRight, g_selection.right);
				storeDWORDSettingInRegistry(storedSelectionBottom, g_selection.bottom);
				InvalidateRect(hWnd, NULL, TRUE);
			}
			break;
		case VK_HOME: // Home => Restore selection
			if ((g_storedSelection.left != -1) && (g_storedSelection.right != -1) && (g_storedSelection.top != -1) && (g_storedSelection.bottom != -1))
			{
				g_appState = statePointB;
				g_selection.left = limitXtoBitmap(g_storedSelection.left);
				g_selection.right = limitXtoBitmap(g_storedSelection.right);
				g_selection.top = limitYtoBitmap(g_storedSelection.top);
				g_selection.bottom = limitYtoBitmap(g_storedSelection.bottom);
				InvalidateRect(hWnd, NULL, TRUE);
			}
			break;
		case VK_F1: // F1 => Toggle display information
			SendMessage(hWnd, WM_COMMAND, IDM_DISPLAYINFORMATION, 0);
			break;
		case VK_TAB: // Tab => Toggle between points
			if (g_appState == statePointA)
			{
				g_appState = statePointB;
				MySetCursorPos(g_selection.right, g_selection.bottom);
				InvalidateRect(hWnd, NULL, TRUE);
			}
			else
			{
				if (g_appState == statePointB)
				{
					g_appState = statePointA;
					MySetCursorPos(g_selection.left, g_selection.top);
					InvalidateRect(hWnd, NULL, TRUE);
				}
			}
			break;
		}
		break;
	case WM_CHAR:
		switch (wParam)
		{
		case VK_ESCAPE: // ESC => cancel capture
			SendMessage(hWnd, WM_GOTOTRAY, 0, 0);
			break;
		case VK_RETURN: // Return => Confirm/OK
			SendMessage(hWnd, WM_NEXTSTATE, 0, 0);
			break;
		case '+': // Increase selection
		 	resizeSelection(hWnd,1);
			break;
		case '-': // Decrease selection
		 	resizeSelection(hWnd,-1);
			break;
		}
		break;
	case WM_ERASEBKGND: // Skip WM_ERASEBKGND (and prevents flickering), because we fill the hole client area every WM_PAINT
		break;
	case WM_PAINT:
		OnPaint(hWnd);
		break;
	case WM_CLOSE:
		DestroyWindow(hWnd);
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	case WM_LBUTTONDOWN: // Left mouse button => Confirm/OK
		SendMessage(hWnd, WM_NEXTSTATE, wParam, lParam);
		break;
	case WM_RBUTTONUP: // Right mouse button => Context menu
	{
		POINT pt;
		GetCursorPos(&pt);
		ShowCursor(true);
		HMENU hMenu = CreatePopupMenu();
		AppendMenu(hMenu, MF_STRING, IDM_CANCELCAPTURE, LoadStringAsWstr(g_hInst, IDS_CANCELCAPTURE).c_str());
		AppendMenu(hMenu, MF_STRING, IDM_EXIT, LoadStringAsWstr(g_hInst, IDS_EXIT).c_str());
		TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hWnd, NULL);
		DestroyMenu(hMenu);
		ShowCursor(false);
		break;
	}
	case WM_MOUSEMOVE: // Mouse moved
		OnMouseMove(hWnd, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), (DWORD)wParam);
		break;
	case WM_TIMER: // Timer...
		switch (wParam)
		{
		case IDT_TIMER1000MS: // 1s timer
			InvalidateRect(hWnd, NULL, TRUE);
			break;
		case IDT_TIMER5000MS: // Onetime 5s timer
			KillTimer(hWnd, IDT_TIMER5000MS); // Only one time
			SendMessage(hWnd, WM_STARTED, 0, 0);
			break;
		}
		break;
	case WM_COMMAND: // Menu selection
		switch (LOWORD(wParam))
		{
		case IDM_CAPTURE:
			SetTimer(hWnd, IDT_TIMER5000MS, 5000, (TIMERPROC)NULL);
			break;
		case IDM_EXIT:
			PostQuitMessage(0);
			break;
		case IDM_ABOUT:
			showProgramInformation(hWnd);
			break;
		case IDM_OPENFOLDER:
			ShellExecute(hWnd, L"open", getScreenshotPathFromRegistry().c_str(), NULL, NULL, SW_SHOWNORMAL);
			break;
		case IDM_SETFOLDER:
			changeScreenshotPathAndStorePathToRegistry();
			break;
		case IDM_SAVETOCLIPBOARD: // Toggle save to clipboard
			g_saveToClipboard = !g_saveToClipboard;
			storeDWORDSettingInRegistry(saveToClipboard, g_saveToClipboard);
			if (!g_saveToClipboard && !g_saveToFile) // Clipboard and file was disabled by user => Warning
			{
				if (g_appState != stateTrayIcon) ShowCursor(true);
				MessageBox(hWnd, LoadStringAsWstr(g_hInst, IDS_NOSAVING).c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONWARNING);
				if (g_appState != stateTrayIcon) ShowCursor(false);
			}
			InvalidateRect(hWnd, NULL, TRUE);
			break;
		case IDM_SAVETOFILE: // Toggle save to file
			g_saveToFile = !g_saveToFile;
			storeDWORDSettingInRegistry(saveToFile,g_saveToFile);
			if (!g_saveToClipboard && !g_saveToFile) // Clipboard and file was disabled by user => Warning
			{
				if (g_appState != stateTrayIcon) ShowCursor(true);
				MessageBox(hWnd, LoadStringAsWstr(g_hInst, IDS_NOSAVING).c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONWARNING);
				if (g_appState != stateTrayIcon) ShowCursor(false);
			}
			InvalidateRect(hWnd, NULL, TRUE);
			break;
		case IDM_ALTERNATIVECOLORS: // Toggle colors
			g_useAlternativeColors = !g_useAlternativeColors;
			storeDWORDSettingInRegistry(useAlternativeColors, g_useAlternativeColors);
			InvalidateRect(hWnd, NULL, TRUE);
			break;
		case IDM_DISPLAYINFORMATION: // Toggle display information
			g_displayInternallnformation = !g_displayInternallnformation;
			storeDWORDSettingInRegistry(displayInternallnformation, g_displayInternallnformation);
			InvalidateRect(hWnd, NULL, TRUE);
			break;
		case IDM_CANCELCAPTURE: // Cancel screenshot
			SendMessage(hWnd, WM_GOTOTRAY, 0, 0);
			break;
		}
		break;
	case WM_DISPLAYCHANGE:
		// Goto tray icon, when display changed, to prevent problems when connecting/disconnecting monitors
		if (g_appState != stateTrayIcon) SendMessage(hWnd, WM_GOTOTRAY, 0, 0);
		break;
	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}
