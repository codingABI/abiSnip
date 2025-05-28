/*+===================================================================
  File:      abiSnip.cpp

  Summary:   Tool to save screenshots as PNG files or copy screenshots to clipboard
			 when the "Print screen" key was pressed

			 Supports:
			 - Zoom to mouse position
			 - Area selection
			 - All monitors selection
			 - Single monitor selection
			 - Selections can be adjusted by mouse or keyboard
			 - Screenshots will be saved as PNG files and/or copied to clipboard
			 - Folder for the PNG files can be set with the context menu of the tray icon
			 - Filename for a PNG file will be set automatically and contains a timestamp, for example "Screenshot 2024-11-24 100706.png"
			 - Selection can be pixelated
			 - Selection can marked with a colored box
			 - Group policy support

			 Program should run on Windows 11/10/8.1/2025/2022/2019/2016/2012R2

  License: CC0
  Copyright (c) 2024-2025 codingABI

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
  Home = Use stored selection
  Delete = Delete stored and used selection
  +/- = Increase/decrease selection
  PageUp/PageDown, mouse wheel = Zoom In/Out
  C = Save to clipboard On/Off (Can be set/force by GPO)
  F = Save to file On/Off (Can be set/force by GPO)
  S = Alternative colors On/Off
  F1 = Display internal information on screen On/Off (Can be set/force by GPO)
  P = Pixelate selected area
  B = Box around selected area

  Refs:
  https://learn.microsoft.com/en-us/windows/win32/gdi/capturing-an-image
  https://devblogs.microsoft.com/oldnewthing/20100412-00/?p=14353
  https://www.codeproject.com/Tips/76427/How-to-bring-window-to-top-with-SetForegroundWindo
  https://devblogs.microsoft.com/oldnewthing/20150406-00/?p=44303

  History:
  20241202, Initial version 1.0.0.1
  20250102, Add pixelate selection
			Add box around selected area
			Add command line arguments
			Add autostart at logon option
			Version update to 1.0.0.2
  20250103, Fix: No RECT -1,-1,-1,-1 was allowed
			Restore g_storedSelection when starting capture
  20250206, Prevent concurrent actions from tray icon context menu
			Recreate tray icon if explorer crashes
  20250207, Change program info MessageBox to TaskDialog (URL and better high dpi support)
            Version update to 1.0.0.3
  20250423, Close start menu after print screen was pressed (otherwise start menu stays on top)
  20250430, Add FIX01 (perhaps a problem caused by Omnissa Horizon Client)
  20250502, Add tray icon context menu entry for opening the last screenshot
            Version update to 1.0.0.4
  20250515, Initial zoom level can be set by registry
  			Screenshot delay can be set by registry
            Add GPOs (admx/adml) for some settings
  20250523, Check screenshot folder to be a valid folder
            Remove temporary FIX01, because finally fixed
            Enable zoom for mouse position even on small selections
            Add commandline parameters /re /rd to enable/disable to all users run key
			Add Snipping Tool 11 on Windows 11 24H2 computers for "Edit last screenshot..."
			Add option to disable PrintScreenKeyForSnippingEnabled
            Version update to 1.0.0.5

===================================================================+*/

// For GNU compilers
#if defined (__GNUC__)
#define UNICODE
#define _WIN32_WINNT 0x0602
#define _MAX_ITOSTR_BASE16_COUNT (8 + 1) // Char length for DWORD to hex conversion
#define URL_ESCAPE_ASCII_URI_COMPONENT 0x00080000 // Missing in TDM-GCC 9.2.0 shlwapi.h
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
#pragma warning(push)
#pragma warning(disable : 4005)
#include <ntstatus.h>
#pragma warning(pop)
#include "resource.h"

// Library-search records for visual studio
#pragma comment(lib,"msimg32")
#pragma comment(lib,"Shlwapi")
#pragma comment(lib,"Gdiplus")
#pragma comment(lib,"Version")
#pragma comment(lib,"Comctl32")

using namespace Gdiplus;
// Defines
#define REGISTRYSETTINGSPATH L"SOFTWARE\\CodingABI\\abiSnip" // Registry path under HKCU to store program settings
#define REGISTRYGPOPATH L"SOFTWARE\\Policies\\CodingABI\\abiSnip" // GPO path under HKLM/HKCU to force program settings
#define REGISTRYGPODEFAULTSPATH L"SOFTWARE\\Policies\\CodingABI\\abiSnip\\Recommended" // GPO path under HKLM/HKCU for default program settings
#define REGISTRYRUNPATH L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run" // Registry path under HKLM/HKCU to start the program at logon
#define REGISTRYRUNPATHX86 L"SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Run" // Registry path under HKLM on a x64 machine to start the program at logon via the x86 registry
#define ZOOMWIDTH 32 // Width of zoomwindow (effective pixel size is ZOOMWIDTH * current zoom scale)
#define ZOOMHEIGHT 32 // Height of zoomwindow (effective pixel size is ZOOMHEIGHT * current zoom scale)
#define MAXZOOMSCALE 32 // Max zoom scale
#define DEFAULTZOOMSCALE 4 // Default zoom scale when selecting point A or B
#define DEFAULTSCREENSHOTDELAY 5 // Default delay in seconds for a delayed screenshot
#define MAXSCREENSHOTDELAY 60 // Max delay in seconds for a delayed screenshot
#define DEFAULTFONT L"Consolas" // Font
#define DEFAULTSAVETOCLIPBOARD TRUE // TRUE, when screenshot should be saved to clipboard
#define DEFAULTSAVETOFILE TRUE // TRUE, when screenshot should be saved to a PNG file
#define DEFAULTUSEALTERNATIVECOLORS FALSE // TRUE, when alternative colors are enabled
#define DEFAULTSHOWDISPLAYINFORMATION FALSE // TRUE, when drawing internal information on screen is enabled
#define PIXELATEFACTOR 8 // Factor for pixelating an area with key "p"
#define MARKEDWIDTH 3 // Line width when marking selected area
#define MARKEDALPHA 128 // Alpha value when marking selected area
#define UNINITIALIZEDLONG (LONG) 0x80000000 // Value for uninitialized pixel positions

// Default colors
#define APPCOLOR RGB(245, 167, 66)
#define APPCOLORINV RGB(255,255,255)
#define MARKCOLOR RGB(255,0,0)
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
	screenshotDelay,
	saveToClipboard,
	saveToFile,
	useAlternativeColors,
	displayInternalInformation,
	storedSelectionLeft,
	storedSelectionTop,
	storedSelectionRight,
	storedSelectionBottom,
	disablePrintScreenKeyForSnipping,
	DEV
};

// Global Variables:
HINSTANCE g_hInst = NULL; // Current instance
HWND g_hWindow = NULL; // Handle to main window
POINT g_appWindowPos; // SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN when fullscreen was started
HBITMAP g_hBitmap = NULL; // Bitmap for screenshot over all monitors
RECT g_selection = { UNINITIALIZEDLONG,UNINITIALIZEDLONG,UNINITIALIZEDLONG,UNINITIALIZEDLONG }; // Selected screenshot area
RECT g_storedSelection = { UNINITIALIZEDLONG,UNINITIALIZEDLONG,UNINITIALIZEDLONG,UNINITIALIZEDLONG }; // Stored selection
BOOL g_useAlternativeColors = DEFAULTUSEALTERNATIVECOLORS; // TRUE when alternative colors are used
BOOL g_saveToFile = DEFAULTSAVETOFILE; // TRUE when screenshots are saved as files
BOOL g_bSaveToFileGPO = FALSE; // TRUE when saving to file is set by a GPO
BOOL g_saveToClipboard = DEFAULTSAVETOCLIPBOARD; // TRUE when screenshots copied to clipboard
BOOL g_bSaveToClipboardGPO = FALSE; // TRUE when copy to clipboard is set by a GPO
BOOL g_displayInternalInformation = DEFAULTSHOWDISPLAYINFORMATION; // TRUE when internal program data are displayed while selecting a screenshot
BOOL g_bDisplayInternalInformationGPO = FALSE; // TRUE when displaying internal program data is set by GPO
DWORD g_screenshotDelay = DEFAULTSCREENSHOTDELAY; // Delay in seconds for a delayed screenshot startet via tray icon context menu "Screenshot (Xs delayed)"
BOOL g_bScreenshotDelayGPO = FALSE; // TRUE when delay is forced by a GPO
wchar_t g_screenshotPath[MAX_PATH] = L""; // Path for screenshots
BOOL g_bScreenshotPathGPO = FALSE; // TRUE when path for screenshots is set by a GPO
BOOL g_bRunKeyReadOnly = FALSE; // TRUE when automatic run via registry is set in HKLM
BOOL g_onetimeCapture = FALSE; // TRUE in onetimeCapture mode (capture once at program start and exit program afterwards)
APPSTATE g_appState = stateTrayIcon; // Current program state
HWND g_activeWindow = NULL; // Active window before program starts fullscreen mode
DWORD g_zoomScale = DEFAULTZOOMSCALE; // Zoom scale for mouse cursor
BOOL g_bZoomScaleGPO = FALSE; // TRUE when initial zoom scale for mouse cursor is forced by GPO
HHOOK g_hHook = NULL; // Handle to hook (We use keyboard hook to start fullscreen mode, when the "Print screen" key was pressed)
HANDLE g_hSemaphoreModalBlocked = NULL; // Semaphore to ensure modal dialogs (even when started by tray icon menu)
NOTIFYICONDATA g_nid; // Tray icon structure
UINT WM_TASKBARCREATED = 0; // Windows sends this message when the taskbar is created (Needs RegisterWindowMessage)
BOOL g_ignoreNextClick = FALSE; // TRUE when focus was force by a simulated click
std::wstring g_sLastScreenshotFile = L""; // Last used filename (Path + filename + extension)
BOOL g_bDisablePrintScreenKeyForSnipping = FALSE; // TRUE when should bei disabel PrintScreenKeyForSnipping silently
BOOL g_bDEV = FALSE; // TRUE when development functions are enabled (only used temporary)

// Function declarations
ATOM                MyRegisterClass(HINSTANCE hInstance);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);

void enterFullScreen(HWND);

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: IsWindows11_24H2OrNewer

  Summary:  Checks if OS is newer or equal Windows 11 24H2

  Args:

  Returns:  TRUE = Windows 11 24H2 or newer
            FALSE = Older then Windows 11 24H2

-----------------------------------------------------------------F-F*/
typedef NTSTATUS(WINAPI* RtlGetVersionPtr)(PRTL_OSVERSIONINFOW);
bool IsWindows11_24H2OrNewer() {
	HMODULE hMod = GetModuleHandleW(L"ntdll.dll");
	if (hMod) {
		RtlGetVersionPtr fxPtr = (RtlGetVersionPtr)GetProcAddress(hMod, "RtlGetVersion");
		if (fxPtr != nullptr) {
			RTL_OSVERSIONINFOW rovi = { 0 };
			rovi.dwOSVersionInfoSize = sizeof(rovi);
			if (fxPtr(&rovi) == STATUS_SUCCESS) {
				return (rovi.dwMajorVersion > 10) ||
					(rovi.dwMajorVersion == 10 && rovi.dwMinorVersion == 0 && rovi.dwBuildNumber >= 26100);
			}
		}
	}
	return false;
}

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
	if (!IsWindow(hWindow)) return;

	BYTE keyState[256] = { 0 };
	// To unlock SetForegroundWindow we need to imitate Alt pressing
	if (GetKeyboardState((LPBYTE)&keyState))
	{
		if (!(keyState[VK_MENU] & 0x80))
		{
			keybd_event(VK_MENU, 0, KEYEVENTF_EXTENDEDKEY | 0, 0);
		}
	}

	SetForegroundWindow(hWindow);

	if (GetKeyboardState((LPBYTE)&keyState))
	{
		if (!(keyState[VK_MENU] & 0x80))
		{
			keybd_event(VK_MENU, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);
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
  Function: normalizeRectangle

  Summary:   "normalize" rectangle to ensure, that .left is the left side and .top is on the upper side

  Args:     RECT rect
			  Rectangle to be normalized

  Returns:  RECT
			  Normalized rectangle

-----------------------------------------------------------------F-F*/
RECT normalizeRectangle(RECT rect)
{
	RECT result = { 0 };

	if (rect.right >= rect.left)
	{
		result.left = rect.left;
		result.right = rect.right;
	}
	else
	{
		result.left = rect.right;
		result.right = rect.left;
	}
	if (rect.bottom >= rect.top)
	{
		result.top = rect.top;
		result.bottom = rect.bottom;
	}
	else
	{
		result.top = rect.bottom;
		result.bottom = rect.top;
	}
	return result;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: programInformationProc

  Summary:  Callback for program information dialog

  Args:     HWND hWindow
  			UINT uMsg
			WPARAM wParam,
			LPARAM lParam,
			LONG_PTR lpRefData

  Returns:  HRESULT

-----------------------------------------------------------------F-F*/
HRESULT CALLBACK programInformationCallbackProc(HWND hWindow, UINT uMsg, WPARAM wParam, LPARAM lParam, LONG_PTR lpRefData) {
	if (uMsg == TDN_HYPERLINK_CLICKED) {
		// Open URL
		ShellExecute(NULL, L"open", (LPCWSTR)lParam, NULL, NULL, SW_SHOWNORMAL);
	}
	return S_OK;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: showProgramInformation

  Summary:   Show program information dialog

  Args:     HWND hWindow
			  Handle to main window

  Returns:

-----------------------------------------------------------------F-F*/
void showProgramInformation(HWND hWindow)
{
	std::wstring sTitle(LoadStringAsWstr(g_hInst, IDS_APP_TITLE));
	std::wstring sMessage(LoadStringAsWstr(g_hInst, IDS_PROGINFO));
	wchar_t szExecutable[MAX_PATH];
	if (GetModuleFileName(NULL, szExecutable, MAX_PATH) == 0) return;

	DWORD  verHandle = 0;
	UINT   size = 0;
	LPBYTE lpBuffer = NULL;
	DWORD  verSize = GetFileVersionInfoSize(szExecutable, NULL);

	if (verSize != 0) {
		BYTE* verData = new BYTE[verSize];

		// Get verions information from EXE file
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

	// Add flag for DEV enabled
	if (g_bDEV) sTitle.append(L" DEV");

	// Add architecture for binary
	#ifdef _WIN64
	sTitle.append(L" x64");
	#else

	#ifdef _WIN32
	sTitle.append(L" x86");
	#endif

	#endif

	int nButtonPressed = 0;
	TASKDIALOGCONFIG config = { 0 };
	config.cbSize = sizeof(config);
	config.hInstance = g_hInst;
	config.hwndParent = hWindow;
	config.dwCommonButtons = TDCBF_OK_BUTTON;
	config.pszMainIcon = MAKEINTRESOURCE(IDI_ICON);
	config.pszMainInstruction = sTitle.c_str();
	config.pszContent = sMessage.c_str();
	config.pszFooter = L"<A HREF=\"https://github.com/codingABI/abiSnip\">https://github.com/codingABI/abiSnip</A>";
	config.pfCallback = programInformationCallbackProc;
	config.dwFlags = TDF_ENABLE_HYPERLINKS;

	TaskDialogIndirect(&config, &nButtonPressed, NULL, NULL);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: showProgramArguments

  Summary:   Show program arguments message box

  Args:     HWND hWindow
			  Handle to main window

  Returns:

-----------------------------------------------------------------F-F*/
void showProgramArguments(HWND hWindow)
{
	std::wstring sTitle(LoadStringAsWstr(g_hInst, IDS_APP_TITLE));
	std::wstring sContent = L"";
	wchar_t szFullPath[MAX_PATH] = L"";
    if (GetModuleFileName(NULL, szFullPath, MAX_PATH) == 0) return;
	std::wstring sMain = PathFindFileName(szFullPath);
	sMain.append(L" [/af] [/ac] | [/f | /rd | /re | /s | /v | /?]");

	sContent
		.append(L"/ac Create and save screenshot to clipboard\n")
		.append(L"/af Create and save screenshot to file\n")
		.append(L"/f Open screenshot folder\n")
		.append(L"/rd Disable program start at logon for all users\n")
		.append(L"/re Enable program start at logon for all users\n")
		.append(L"/s Open screenshot selection\n")
		.append(L"/v Show version information\n")
		.append(L"/? Show this dialog");

	int nButtonPressed = 0;
	TASKDIALOGCONFIG config = { 0 };
	config.cbSize = sizeof(config);
	config.hInstance = g_hInst;
	config.hwndParent = hWindow;
	config.pszMainIcon = MAKEINTRESOURCE(IDI_ICON);
	config.pszWindowTitle = sTitle.c_str();
	config.dwCommonButtons = TDCBF_OK_BUTTON;
	config.pszMainInstruction = sMain.c_str();
	config.pszContent = sContent.c_str();

	TaskDialogIndirect(&config, &nButtonPressed, NULL, NULL);
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: checkScreenshotTargets

  Summary:   Show warning, if all screenshot targets are disabled

  Args:

  Returns:

-----------------------------------------------------------------F-F*/
void checkScreenshotTargets(HWND hWindow) {
	if (!g_saveToClipboard && !g_saveToFile) // Clipboard and file was disabled by user => Warning
	{
		if (g_appState != stateTrayIcon) ShowCursor(true);
		MessageBox(hWindow, LoadStringAsWstr(g_hInst, IDS_TARGETSDISABLED).c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONWARNING);
		if (g_appState != stateTrayIcon) ShowCursor(false);
	}
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: isSelectionValid

  Summary:   Checks if selected area is valid

  Args:     RECT rect
			  Selection rectangle

  Returns:	BOOL
			  TRUE = valid
			  FALSE = not valid

-----------------------------------------------------------------F-F*/
BOOL isSelectionValid(RECT rect)
{
	if ((rect.left == UNINITIALIZEDLONG) || (rect.right == UNINITIALIZEDLONG) || (rect.top == UNINITIALIZEDLONG) || (rect.bottom == UNINITIALIZEDLONG)) return FALSE;
	return TRUE;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: deleteValueFromRegistry

  Summary:   Deletes a value from registry

  Args:     HKEY hKey
              Registry key used as root
            LPCWSTR lpSubKey
              Registry key
			LPCWSTR lpValueName
			  Registry value

  Returns:	LSTATUS
			  ERROR_SUCCESS = Success
			  Return codes from Winerror.h

-----------------------------------------------------------------F-F*/
LSTATUS deleteValueFromRegistry(HKEY hkRoot, LPCWSTR lpSubKey, LPCWSTR lpValueName)
{
	LSTATUS lsResult = ERROR_SUCCESS;

	if (lpSubKey == NULL) return ERROR_INVALID_PARAMETER;
	if (lpValueName == NULL) return ERROR_INVALID_PARAMETER;

	HKEY hKey = NULL;
	lsResult = RegOpenKeyEx(hkRoot, lpSubKey, 0, KEY_SET_VALUE, &hKey);
	if (lsResult == ERROR_SUCCESS) {
		lsResult = RegDeleteValue(hKey, lpValueName);
		if ((lsResult == ERROR_SUCCESS) || (lsResult == ERROR_FILE_NOT_FOUND)) {
			lsResult = ERROR_SUCCESS;
		}
		RegCloseKey(hKey);
	} else {
		if (lsResult == ERROR_PATH_NOT_FOUND) return ERROR_SUCCESS;
	}
	return lsResult;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: getSZFromRegistry

  Summary:   Gets REG_SZ value from registry

  Args:     HKEY hKey
              Registry key used as root
            LPCWSTR lpSubKey
              Registry key
			LPCWSTR lpValueName
			  Registry value
			wchar_t *pszValue
			  Pointer to wchar_t array where result of registry value will be stored
			DWORD dwMaxValueWChars
			  Max chars for wchar_t array where result of registry value will be stored (including null termination)

  Returns:	BOOL
			  TRUE = Success
			  FALSE = Failed

-----------------------------------------------------------------F-F*/
BOOL getSZFromRegistry(HKEY hKey,LPCWSTR lpSubKey, LPCWSTR lpValueName, wchar_t *pszValue, DWORD dwMaxValueWChars)
{
	BOOL bResult = FALSE;
	if (lpSubKey == NULL) return FALSE;
	if (lpValueName == NULL) return FALSE;
	if (pszValue == NULL) return FALSE;

	*pszValue = L'\0';
	DWORD valueSize = 0;
	DWORD keyType = 0;

	// Get size of registry value
	if (RegGetValue(hKey, lpSubKey, lpValueName, RRF_RT_REG_SZ, &keyType, NULL, &valueSize) == ERROR_SUCCESS)
	{
		if ((valueSize > 0) && (valueSize <= (dwMaxValueWChars + 1) * sizeof(WCHAR))) // Size OK?
		{
			// Get registry value
			valueSize = dwMaxValueWChars * sizeof(WCHAR); // Max size incl. termination
			if (RegGetValue(hKey, lpSubKey, lpValueName, RRF_RT_REG_SZ | RRF_ZEROONFAILURE, NULL, pszValue, &valueSize) == ERROR_SUCCESS)
			{
				bResult = TRUE;
			}
		}
	}
	return bResult;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: getDWORDValueFromRegistry

  Summary:   Gets REG_DWORD value from registry

  Args:     HKEY hKey
              Registry key used as root
            LPCWSTR lpSubKey
              Registry key
			LPCWSTR lpValueName
			  Registry value
			DWORD &dwValue
			  Receives the value's data

  Returns:	LSTATUS
			  ERROR_SUCCESS = Success
			  Return codes from Winerror.h

-----------------------------------------------------------------F-F*/
LSTATUS getDWORDValueFromRegistry(HKEY hkRoot,LPCWSTR lpSubKey, LPCWSTR lpValueName, DWORD &dwValue)
{
	HKEY hKey;
	DWORD dwSize = sizeof(DWORD);
	LSTATUS lsResult = ERROR_SUCCESS;

	if (lpSubKey == NULL) return ERROR_INVALID_PARAMETER;
	if (lpValueName == NULL) return ERROR_INVALID_PARAMETER;

	// Open registry key
	lsResult = RegOpenKeyEx(hkRoot, lpSubKey, 0, KEY_READ, &hKey);
	if (lsResult == ERROR_SUCCESS) {
		// Read value from registry
		lsResult = RegQueryValueEx(hKey, lpValueName, NULL, NULL, (LPBYTE)&dwValue, &dwSize);
		RegCloseKey(hKey);
	}
	return lsResult;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: setDWORDValueToRegistry

  Summary:   Sets REG_DWORD value to registry

  Args:     HKEY hKey
              Registry key used as root
            LPCWSTR lpSubKey
              Registry key
			LPCWSTR lpValueName
			  Registry value
			DWORD dwValue
			  Content of value to set

  Returns:	LSTATUS
			  ERROR_SUCCESS = Success
			  Return codes from Winerror.h

-----------------------------------------------------------------F-F*/
LSTATUS setDWORDValueToRegistry(HKEY hkRoot,LPCWSTR lpSubKey, LPCWSTR lpValueName, DWORD dwValue)
{
	HKEY hKey;
	DWORD dwSize = sizeof(DWORD);
	LSTATUS lsResult = ERROR_SUCCESS;

	if (lpSubKey == NULL) return ERROR_INVALID_PARAMETER;
	if (lpValueName == NULL) return ERROR_INVALID_PARAMETER;

	// Open registry key
	lsResult = RegCreateKeyEx(hkRoot, lpSubKey, 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL);
	if (lsResult == ERROR_SUCCESS) {

		// Write to registry
		lsResult = RegSetValueEx(hKey, lpValueName, 0, REG_DWORD, (const BYTE*)&dwValue, sizeof(dwValue));
		if (lsResult != ERROR_SUCCESS) {
			OutputDebugString(L"Error writing to registry");
		}
		RegCloseKey(hKey);
	}
	return lsResult;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: getDWORDSettingFromRegistry

  Summary:   Gets stored DWORD setting from registry

  Args:      APPDWORDSETTINGS setting
			   Setting

  Returns:	BOOL
			  TRUE = Success
			  FALSE = Error

-----------------------------------------------------------------F-F*/
BOOL getDWORDSettingFromRegistry(APPDWORDSETTINGS setting) {
	DWORD dwValue = 0;
	BOOL bFound = FALSE;

	std::wstring sValueName = L"";
	switch (setting)
	{
		case defaultZoomScale: sValueName.assign(L"defaultZoomScale"); break;
		case screenshotDelay: sValueName.assign(L"screenshotDelay"); break;
		case saveToClipboard: sValueName.assign(L"saveToClipboard"); break;
		case saveToFile: sValueName.assign(L"saveToFile"); break;
		case useAlternativeColors: sValueName.assign(L"useAlternativeColors"); break;
		case displayInternalInformation: sValueName.assign(L"displayInternalInformation"); break;
		case storedSelectionLeft: sValueName.assign(L"storedSelectionLeft"); break;
		case storedSelectionTop: sValueName.assign(L"storedSelectionTop"); break;
		case storedSelectionRight: sValueName.assign(L"storedSelectionRight"); break;
		case storedSelectionBottom: sValueName.assign(L"storedSelectionBottom"); break;
		case disablePrintScreenKeyForSnipping: sValueName.assign(L"disablePrintScreenKeyForSnipping"); break;
		case DEV: sValueName.assign(L"DEV"); break;
		default: OutputDebugString(L"Invalid setting");	return FALSE;
	}
	if (sValueName.empty())
	{
		OutputDebugString(L"Invalid setting");
		return FALSE;
	}

	// Reset GPO flag
	switch (setting)
	{
		case screenshotDelay: g_bScreenshotDelayGPO = FALSE; break;
		case saveToClipboard: g_bSaveToClipboardGPO = FALSE; break;
		case saveToFile: g_bSaveToFileGPO = FALSE; break;
		case displayInternalInformation: g_bDisplayInternalInformationGPO = FALSE; break;
	}

	// Check GPO settings
	switch (setting)
	{
		case defaultZoomScale:
		case screenshotDelay:
		case saveToClipboard:
		case saveToFile:
		case displayInternalInformation:
		case disablePrintScreenKeyForSnipping:
		{
			// Get stored path from GPO or registry
			if (!bFound) bFound = (getDWORDValueFromRegistry(HKEY_CURRENT_USER, REGISTRYGPOPATH, sValueName.c_str(), dwValue) == ERROR_SUCCESS);
			if (!bFound) bFound = (getDWORDValueFromRegistry(HKEY_LOCAL_MACHINE, REGISTRYGPOPATH, sValueName.c_str(), dwValue) == ERROR_SUCCESS);
			break;
		}
	}
	if (bFound) switch (setting)
	{
		case screenshotDelay: g_bScreenshotDelayGPO = TRUE; break;
		case saveToClipboard: g_bSaveToClipboardGPO = TRUE; break;
		case saveToFile: g_bSaveToFileGPO = TRUE; break;
		case displayInternalInformation: g_bDisplayInternalInformationGPO = TRUE; break;
	}

	// User registry value
	if (!bFound) bFound = (getDWORDValueFromRegistry(HKEY_CURRENT_USER, REGISTRYSETTINGSPATH, sValueName.c_str(), dwValue) == ERROR_SUCCESS);

	// Failback to GPO default settings, if not found
	if (!bFound) switch (setting)
	{
		case defaultZoomScale:
		case screenshotDelay:
		case saveToClipboard:
		case saveToFile:
		case displayInternalInformation:
		case disablePrintScreenKeyForSnipping:
		{
			// Get stored path from GPO default settings
			if (!bFound) bFound = (getDWORDValueFromRegistry(HKEY_CURRENT_USER, REGISTRYGPODEFAULTSPATH, sValueName.c_str(), dwValue) == ERROR_SUCCESS);
			if (!bFound) bFound = (getDWORDValueFromRegistry(HKEY_LOCAL_MACHINE, REGISTRYGPODEFAULTSPATH, sValueName.c_str(), dwValue) == ERROR_SUCCESS);
			break;
		}
	}

	// Use program defaults, if not found
	if (!bFound) switch (setting)
	{
		case defaultZoomScale: dwValue = DEFAULTZOOMSCALE; break;
		case screenshotDelay: dwValue = DEFAULTSCREENSHOTDELAY; break;
		case saveToClipboard: dwValue = DEFAULTSAVETOCLIPBOARD; break;
		case saveToFile: dwValue = DEFAULTSAVETOFILE; break;
		case useAlternativeColors: dwValue = DEFAULTUSEALTERNATIVECOLORS; break;
		case displayInternalInformation: dwValue = DEFAULTSHOWDISPLAYINFORMATION; break;
		case storedSelectionLeft: dwValue = UNINITIALIZEDLONG; break;
		case storedSelectionTop: dwValue = UNINITIALIZEDLONG; break;
		case storedSelectionRight: dwValue = UNINITIALIZEDLONG; break;
		case storedSelectionBottom: dwValue = UNINITIALIZEDLONG; break;
		case disablePrintScreenKeyForSnipping: dwValue = FALSE; break;
		case DEV: dwValue = 0; break;
		default: OutputDebugString(L"Invalid setting");	return FALSE;
	}

	// Check limits
	switch (setting)
	{
		case defaultZoomScale:
			if (dwValue < 1) dwValue = 1;
			if (dwValue > MAXZOOMSCALE) dwValue = MAXZOOMSCALE;
			break;
		case screenshotDelay:
			if (dwValue < 1) dwValue = 1;
			if (dwValue > MAXSCREENSHOTDELAY) dwValue = MAXSCREENSHOTDELAY;
			break;
		case saveToClipboard:
		case saveToFile:
		case useAlternativeColors:
		case displayInternalInformation:
			if (dwValue > 1) dwValue = 1;
			break;
		case disablePrintScreenKeyForSnipping:
			if (dwValue > 1) dwValue = 1;
			break;
	}

	switch (setting)
	{
		case defaultZoomScale: g_zoomScale = dwValue; break;
		case screenshotDelay: g_screenshotDelay = dwValue; break;
		case saveToClipboard: g_saveToClipboard = dwValue; break;
		case saveToFile: g_saveToFile = dwValue; break;
		case useAlternativeColors: g_useAlternativeColors = dwValue; break;
		case displayInternalInformation: g_displayInternalInformation = dwValue; break;
		case storedSelectionLeft: g_storedSelection.left = dwValue; break;
		case storedSelectionTop: g_storedSelection.top = dwValue; break;
		case storedSelectionRight: g_storedSelection.right = dwValue; break;
		case storedSelectionBottom: g_storedSelection.bottom = dwValue; break;
		case disablePrintScreenKeyForSnipping: g_bDisablePrintScreenKeyForSnipping = dwValue; break;
		case DEV: g_bDEV = dwValue; break;
		default:
			OutputDebugString(L"Invalid setting");
			return FALSE;
	};
	return TRUE;
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
	std::wstring sValueName = L"";
	switch (setting)
	{
	case saveToClipboard:
		sValueName.assign(L"saveToClipboard");
		break;
	case saveToFile:
		sValueName.assign(L"saveToFile");
		break;
	case useAlternativeColors:
		sValueName.assign(L"useAlternativeColors");
		break;
	case displayInternalInformation:
		sValueName.assign(L"displayInternalInformation");
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
	case disablePrintScreenKeyForSnipping:
		sValueName.assign(L"disablePrintScreenKeyForSnipping");
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

	if (setDWORDValueToRegistry(HKEY_CURRENT_USER,REGISTRYSETTINGSPATH, sValueName.c_str(), dwValue) == ERROR_SUCCESS)
		return TRUE;
	else
		return FALSE;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: setRunKeyRegistryValue

  Summary:   Set or delete registry run value to run the program at logon

  Args:		BOOL enabled
			  TRUE = Enable program start at logon
			  FALSE = Disable program start at logon
			HKEY hkRoot
			  HKEY_CURRENT_USER or HKEY_LOCAL_MACHINE (HKEY_LOCAL_MACHINE needs local admin rights)
  Returns:

-----------------------------------------------------------------F-F*/
void setRunKeyRegistryValue(BOOL enabled, HKEY hkRoot)
{
	wchar_t szProgramPath[MAX_PATH] = L"";
	wchar_t szProgramPathQuoted[MAX_PATH] = L"";

	if ((hkRoot != HKEY_CURRENT_USER) && (hkRoot != HKEY_LOCAL_MACHINE)) return; // Unknown root

	if (!enabled)
	{ // Delete run key value
		deleteValueFromRegistry(hkRoot, REGISTRYRUNPATH, LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str());
		#ifdef _WIN64
		if (hkRoot == HKEY_LOCAL_MACHINE) // Also clear x86 registry on a x64 system
			deleteValueFromRegistry(hkRoot, REGISTRYRUNPATHX86, LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str());
		#endif
	}
	else
	{ // Create run key value
		if (GetModuleFileName(NULL, szProgramPath, MAX_PATH) > 0)
		{
			_snwprintf_s(szProgramPathQuoted, MAX_PATH, _TRUNCATE, L"%c%s%c", L'"', szProgramPath, L'"'); // Quote string
			if (hkRoot != NULL) RegSetKeyValue(hkRoot, REGISTRYRUNPATH, LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), REG_SZ, szProgramPathQuoted, (DWORD)(wcslen(szProgramPathQuoted) + 1) * sizeof(WCHAR));
		}
	}
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: isRunKeyEnabledFromRegistry

  Summary:   Gets registry run value to run the program at logon

  Args:

  Returns:	BOOL
			  TRUE = Program start at logon is enabled
			  FALSE = Program start at logon is disabled

-----------------------------------------------------------------F-F*/
BOOL isRunKeyEnabledFromRegistry()
{
	BOOL bFound = FALSE;
	wchar_t szValue[MAX_PATH] = L"";
	szValue[0] = L'\0';
	DWORD valueSize = 0;
	DWORD keyType = 0;

	g_bRunKeyReadOnly = FALSE;
	wchar_t szProgramPath[MAX_PATH];
	wchar_t szProgramPathQuoted[MAX_PATH];
	if (GetModuleFileName(NULL, szProgramPath, MAX_PATH) == 0) return FALSE;
	_snwprintf_s(szProgramPathQuoted, MAX_PATH, _TRUNCATE, L"%c%s%c", L'"', szProgramPath, L'"'); // Quote string

	// Check HKLM run keys
	if (getSZFromRegistry(HKEY_LOCAL_MACHINE, REGISTRYRUNPATH, LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), szValue, MAX_PATH)) {
		if (_wcsicmp(szValue, szProgramPathQuoted) != 0) { // Ignore invalid registry value
			szValue[0] = L'\0';
		} else 	{
			g_bRunKeyReadOnly = TRUE;
			bFound = TRUE;
			setRunKeyRegistryValue(FALSE,HKEY_CURRENT_USER); // Remove user run key because local machine key is set
		}
	}
	if (!bFound && getSZFromRegistry(HKEY_LOCAL_MACHINE, REGISTRYRUNPATHX86, LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), szValue, MAX_PATH)) {
		if (_wcsicmp(szValue, szProgramPathQuoted) != 0) { // Ignore invalid registry value
			szValue[0] = L'\0';
		} else 	{
			g_bRunKeyReadOnly = TRUE;
			bFound = TRUE;
			setRunKeyRegistryValue(FALSE,HKEY_CURRENT_USER); // Remove user run key because local machine key is set
		}
	}

	if (!bFound && getSZFromRegistry(HKEY_CURRENT_USER, REGISTRYRUNPATH, LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), szValue, MAX_PATH)) {
		if (_wcsicmp(szValue, szProgramPathQuoted) != 0) setRunKeyRegistryValue(TRUE,HKEY_CURRENT_USER); // Fix invalid registry value
	}

	if (wcslen(szValue) > 0)
	{
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: getScreenshotPathFromRegistry

  Summary:   Get path of screenshot folder from GPO, registry or path to EXE file

  Args:

  Returns:

-----------------------------------------------------------------F-F*/
void getScreenshotPathFromRegistry()
{
	BOOL bFound = FALSE;
	wchar_t szRegistryValue[]=L"screenshotPath";
	g_bScreenshotPathGPO = FALSE;
	g_screenshotPath[0] = L'\0';

	// Get stored path from GPO or registry
	bFound = getSZFromRegistry(HKEY_CURRENT_USER, REGISTRYGPOPATH, szRegistryValue, g_screenshotPath, MAX_PATH);
	if (!bFound) bFound = getSZFromRegistry(HKEY_LOCAL_MACHINE, REGISTRYGPOPATH, szRegistryValue, g_screenshotPath, MAX_PATH);
	if (bFound) g_bScreenshotPathGPO = TRUE;
	if (!bFound) bFound = getSZFromRegistry(HKEY_CURRENT_USER, REGISTRYSETTINGSPATH, szRegistryValue, g_screenshotPath, MAX_PATH);

	// Failback value (also check folder to be folder and nothing else, because later we use ShellExecute)
	if (!bFound || !PathIsDirectory(g_screenshotPath))
	{
		// GPO default settings
		bFound = getSZFromRegistry(HKEY_CURRENT_USER, REGISTRYGPODEFAULTSPATH, szRegistryValue, g_screenshotPath, MAX_PATH);
		if (!bFound) bFound = getSZFromRegistry(HKEY_LOCAL_MACHINE, REGISTRYGPODEFAULTSPATH, szRegistryValue, g_screenshotPath, MAX_PATH);
		if (!bFound || !PathIsDirectory(g_screenshotPath)) // Last failback
		{
			// Use path of EXE file
			GetModuleFileName(NULL, g_screenshotPath, MAX_PATH);
			PathRemoveFileSpec(g_screenshotPath);
		}
	}
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
	LPITEMIDLIST pidl;

	BROWSEINFOW bi = { 0 };
	bi.hwndOwner = g_hWindow;
	bi.lpszTitle = sMessage.c_str();
	bi.ulFlags = BIF_NEWDIALOGSTYLE | BIF_RETURNONLYFSDIRS | BIF_EDITBOX | BIF_VALIDATE;

	// Convert the start path to a PIDL
	HRESULT hr = SHParseDisplayName(g_screenshotPath, NULL, &pidl, 0, NULL);
	if (SUCCEEDED(hr))
	{
		bi.lpfn = changeScreenshotPathAndStorePathToRegistryCallbackProc;
		bi.lParam = (LPARAM)g_screenshotPath;
	}

	// Browse for folder dialog
	LPITEMIDLIST pidlSelected = SHBrowseForFolder(&bi);
	if (pidlSelected != NULL)
	{
		WCHAR szPath[MAX_PATH];
		if (SHGetPathFromIDList(pidlSelected, szPath))
		{
			RegSetKeyValue(HKEY_CURRENT_USER, REGISTRYSETTINGSPATH, L"screenshotPath", REG_SZ, szPath, (DWORD)(wcslen(szPath) + 1) * sizeof(WCHAR));
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
  Function: checkPrintScreenKeyForSnipping

  Summary:   Check for PrintScreenKeyForSnippingEnabled registry key (= Windows capture tool is enabled)

  Args:     HWND hWindow
			  Handle to main window

  Returns:

-----------------------------------------------------------------F-F*/
void checkPrintScreenKeyForSnipping(HWND hWindow)
{
	std::wstring sTitle(LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str());
	std::wstring sMain(LoadStringAsWstr(g_hInst, IDS_PRINTKEYWARNINGMAIN).c_str());
	std::wstring sContent(LoadStringAsWstr(g_hInst, IDS_PRINTKEYWARNINGCONTEND).c_str());
	std::wstring sYes(LoadStringAsWstr(g_hInst, IDS_YES).c_str());
	std::wstring sYesAlways(LoadStringAsWstr(g_hInst, IDS_YESALWAYS).c_str());
	std::wstring sNo(LoadStringAsWstr(g_hInst, IDS_NO).c_str());

	int nButtonPressed = 0;

	#define IDYESALWAYS 100
	TASKDIALOG_BUTTON buttons[] = {
		{IDYES,sYes.c_str()},
		{IDYESALWAYS,sYesAlways.c_str()},
		{IDNO,sNo.c_str()}
	};

	TASKDIALOGCONFIG config = { 0 };
	config.cbSize = sizeof(config);
	config.hInstance = g_hInst;
	config.hwndParent = hWindow;
	config.pszWindowTitle = sTitle.c_str();
	config.cButtons = ARRAYSIZE(buttons);
	config.pButtons = buttons;
	config.pszMainIcon = TD_WARNING_ICON;
	config.pszMainInstruction = sMain.c_str();
	config.pszContent = sContent.c_str();
	config.nDefaultButton = IDYES;

	DWORD regValue = 0;
	if (getDWORDValueFromRegistry(HKEY_CURRENT_USER, L"Control Panel\\Keyboard", L"PrintScreenKeyForSnippingEnabled", regValue) == ERROR_SUCCESS) {
		if (regValue == 1) {
			getDWORDSettingFromRegistry(disablePrintScreenKeyForSnipping);
			int rc = IDNO;
			if (!g_bDisablePrintScreenKeyForSnipping) {
				TaskDialogIndirect(&config, &nButtonPressed, NULL, NULL);
				if (nButtonPressed == IDYESALWAYS) {
					storeDWORDSettingInRegistry(disablePrintScreenKeyForSnipping, 1);
				}
			}
			if (g_bDisablePrintScreenKeyForSnipping || (nButtonPressed == IDYES) || (nButtonPressed == IDYESALWAYS)) {
				// Set PrintScreenKeyForSnippingEnabled to 0x0 to disable the Windows builtin capture tool
				setDWORDValueToRegistry(HKEY_CURRENT_USER, L"Control Panel\\Keyboard", L"PrintScreenKeyForSnippingEnabled", 0);
			}
		}
	}
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: forceFocus

  Summary:   Force foreground window (=focus)

  Issue: When start menu is opened and print screen is pressed the start menu stays in top of the abiSnip and holds the keyboard focus
  Workarround: Simulate a mouse click the abiSnip window

  Args:     HWND hWindow
			  Handle to window

  Returns:

-----------------------------------------------------------------F-F*/
void forceFocus(HWND hWindow) {
	POINT mouse;
	GetCursorPos(&mouse); // Stores current mouse position
	SetCursorPos(g_appWindowPos.x, g_appWindowPos.y); // Move mouse to left, upper window corner

	INPUT input = { 0 };
	input.type = INPUT_MOUSE;
	input.mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
	SendInput(1, &input, sizeof(INPUT)); // Send left mouse button down
	input.mi.dwFlags = MOUSEEVENTF_LEFTUP;
	SendInput(1, &input, sizeof(INPUT)); // Send left mouse button up

	SetCursorPos(mouse.x, mouse.y); // Move mouse back to stored position
	g_ignoreNextClick = TRUE; // Ignore last mouse click
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
	if (GetObject(g_hBitmap, sizeof(bm), &bm) != 0)
	{
		if (X < 0) return 0;
		if (X > bm.bmWidth - 1) return bm.bmWidth - 1;
	}
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
	if (GetObject(g_hBitmap, sizeof(bm), &bm) != 0)
	{
		if (Y < 0) return 0;
		if (Y > bm.bmHeight - 1) return bm.bmHeight - 1;
	}
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
			  TRUE = success
			  FALSE = failure

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
			sFullPathWorkingFile.assign(g_screenshotPath).append(L"\\").append(fileName);

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

	if (bRC) g_sLastScreenshotFile = sFullPathWorkingFile;
	GdiplusShutdown(gdiplusToken);
	return bRC;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: saveSelection

  Summary:   Save selected area to clipboard or/and file

  Args:     HWND hWindow
			  Handle to window

  Returns:	BOOL
			  TRUE = success
			  FALSE = failure

-----------------------------------------------------------------F-F*/
BOOL saveSelection(HWND hWindow)
{
	HDC hdcScreenshot = NULL;
	HDC hdcSelection = NULL;
	HGDIOBJ hbmScreenshotOld = NULL;
	HGDIOBJ hbmSelectionOld = NULL;
	HBITMAP hBitmap = NULL;
	BOOL bResult = TRUE;
	RECT finalSelection;
	wchar_t szHex[_MAX_ITOSTR_BASE16_COUNT + 2];
	std::wstring sMessage = L"";
	int selectionWidth = 0;
	int selectionHeight = 0;

	// Retrieves information the entire screenshot bitmap
	BITMAP bm;
	if (GetObject(g_hBitmap, sizeof(bm), &bm) == 0) goto FAIL;

	finalSelection = normalizeRectangle(g_selection);

	if (finalSelection.left < 0) finalSelection.left = 0;
	if (finalSelection.right > bm.bmWidth - 1) finalSelection.right = bm.bmWidth - 1;
	if (finalSelection.top < 0) finalSelection.top = 0;
	if (finalSelection.bottom > bm.bmHeight - 1) finalSelection.bottom = bm.bmHeight - 1;

	selectionWidth = finalSelection.right - finalSelection.left + 1;
	selectionHeight = finalSelection.bottom - finalSelection.top + 1;

	// Copy selected area from entire screenshot to new bitmap
	hdcScreenshot = CreateCompatibleDC(NULL);
	if (hdcScreenshot == NULL) goto FAIL;

	hbmScreenshotOld = SelectObject(hdcScreenshot, g_hBitmap);
	if (hbmScreenshotOld == NULL) goto FAIL;

	hdcSelection = CreateCompatibleDC(NULL);
	if (hdcSelection == NULL) goto FAIL;

	hBitmap = CreateCompatibleBitmap(hdcScreenshot, selectionWidth, selectionHeight);
	if (hBitmap == NULL)
	{
		sMessage.assign(L"CreateCompatibleBitmap@saveSelection ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		goto FAIL;
	}

	hbmSelectionOld = SelectObject(hdcSelection, hBitmap);
	if (hbmSelectionOld == NULL) goto FAIL;

	if (BitBlt(hdcSelection, 0, 0, selectionWidth, selectionHeight,
		hdcScreenshot, finalSelection.left, finalSelection.top, SRCCOPY))
	{
		if (g_saveToFile) // Save selected area to file?
		{
			// Create folder
			CreateDirectory(g_screenshotPath, NULL);

			// Create file
			SYSTEMTIME tLocal;
			GetLocalTime(&tLocal);
			wchar_t szFileName[MAX_PATH] = L"";

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

		checkScreenshotTargets(hWindow);
	}
	else
	{
		// BitBlt error
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT + 2, _TRUNCATE, L"0x%08X", GetLastError());
		sMessage.assign(L"BitBlt@saveSelection ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED))
			.append(L" ")
			.append(szHex);
		goto FAIL;
	}

	goto CLEANUP;
FAIL:
	bResult = FALSE;
	OutputDebugString(L"saveSelection fails");
	if (sMessage.length() == 0)
		sMessage.assign(L"saveSelection ").append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
	MessageBox(hWindow, sMessage.c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONERROR);
CLEANUP:
	// Free resources/Cleanup
	if (hBitmap != NULL) DeleteObject(hBitmap);

	if (hdcSelection != NULL)
	{
		if (hbmSelectionOld != NULL) SelectObject(hdcSelection, hbmSelectionOld);
		DeleteDC(hdcSelection);
	}

	if (hdcScreenshot != NULL)
	{
		if (hbmScreenshotOld != NULL) SelectObject(hdcScreenshot, hbmScreenshotOld);
		DeleteDC(hdcScreenshot);
	}
	return bResult;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: zoomMousePosition

  Summary:   Creates zoom boxes for mouse position or point A/B

  Args:     HDC hdcOutputBuffer
			  Handle to drawing context for the output buffer
			BOXTYPE boxType
			  Type of position (BoxFirstPointA, BoxFinalPointA, BoxFinalPointB)

  Returns:	BOOL
			  TRUE = success
			  TRUE = failure

-----------------------------------------------------------------F-F*/
BOOL zoomMousePosition(HDC hdcOutputBuffer, BOXTYPE boxType)
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
	BOOL bResult = TRUE;
	std::wstring sMessage = L"";

	// Disable zoom for inactive point when selected area is too small
	if (!(((g_appState==statePointA) && (boxType == BoxFinalPointA)) ||
		((g_appState==statePointB) && (boxType == BoxFinalPointB)))) {
		if ((boxType != BoxFirstPointA) && (abs(g_selection.right - g_selection.left) < (long) (ZOOMWIDTH * g_zoomScale))) goto CLEANUP; // Selection too small
		if ((boxType != BoxFirstPointA) && (abs(g_selection.bottom - g_selection.top) < (long) (ZOOMHEIGHT * g_zoomScale))) goto CLEANUP; // Selection too small
	}

	// Font
	// Specify a font typeface name and weight.
	if (plf == NULL) goto FAIL;
	if (_snwprintf_s(plf->lfFaceName, 12, _TRUNCATE, L"%s", DEFAULTFONT) < 0) goto FAIL;
	plf->lfWeight = FW_NORMAL;
	hfnt = CreateFontIndirect(plf);
	if (hfnt == NULL)
	{
		sMessage.assign(L"CreateFontIndirect@zoomMousePosition ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		goto FAIL;
	}

	hfntPrev = SelectObject(hdcOutputBuffer, hfnt);
	if (hfntPrev == NULL) goto FAIL;

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
		BITMAP bm;
		if (GetObject(g_hBitmap, sizeof(BITMAP), (PSTR)&bm) == 0) goto FAIL;
		SetStretchBltMode(hdcOutputBuffer, COLORONCOLOR);

		if (!StretchBlt(hdcOutputBuffer, zoomBoxX, zoomBoxY, ZOOMWIDTH * g_zoomScale, ZOOMHEIGHT * g_zoomScale,
			hdcOutputBuffer, zoomCenterX - ZOOMWIDTH / 2, zoomCenterY - ZOOMHEIGHT / 2, ZOOMWIDTH, ZOOMHEIGHT, SRCCOPY))
		{
			sMessage.assign(L"StretchBlt@zoomMousePosition ")
				.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
			goto FAIL;
		}
	}
	else goto FAIL;

	// Frame
	RECT outer;
	outer.left = zoomBoxX - 1;
	outer.top = zoomBoxY - 1;
	outer.right = zoomBoxX + ZOOMWIDTH * g_zoomScale + 1;
	outer.bottom = zoomBoxY + ZOOMHEIGHT * g_zoomScale + 1;
	hBrush = CreateSolidBrush(g_useAlternativeColors ? ALTAPPCOLOR : APPCOLOR);
	if (hBrush == NULL) goto FAIL;

	if (g_zoomScale > 1) FrameRect(hdcOutputBuffer, &outer, hBrush);

	// Cross
	if (boxType == BoxFirstPointA)
	{
		RECT center;
		center.left = zoomBoxX - 1;
		center.top = zoomBoxY + g_zoomScale * ZOOMHEIGHT / 2 - 1;
		center.right = zoomBoxX + ZOOMWIDTH * g_zoomScale + 1;
		center.bottom = center.top + g_zoomScale + 2;

		FrameRect(hdcOutputBuffer, &center, hBrush);

		center.left = zoomBoxX + g_zoomScale * ZOOMWIDTH / 2 - 1;
		center.top = zoomBoxY - 1;
		center.right = center.left + g_zoomScale + 2;
		center.bottom = zoomBoxY + ZOOMHEIGHT * g_zoomScale + 1;

		FrameRect(hdcOutputBuffer, &center, hBrush);
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

	SetTextColor(hdcOutputBuffer, g_useAlternativeColors ? ALTAPPCOLORINV : APPCOLORINV);
	DrawText(hdcOutputBuffer, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP | textFormat);

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

	SetBkMode(hdcOutputBuffer, TRANSPARENT);
	SetTextColor(hdcOutputBuffer, g_useAlternativeColors ? ALTAPPCOLOR : APPCOLOR);
	if (g_zoomScale > 1) DrawText(hdcOutputBuffer, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP | textFormat);

	// Pointer selection
	SetTextColor(hdcOutputBuffer, g_useAlternativeColors ? ALTAPPCOLORINV : APPCOLORINV);
	SetBkColor(hdcOutputBuffer, g_useAlternativeColors ? ALTAPPCOLOR : APPCOLOR);
	SetBkMode(hdcOutputBuffer, OPAQUE);

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

	if (wcslen(strData) > 0) DrawText(hdcOutputBuffer, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP | textFormat);

	// Text position Y

	// Text rotated 90 degree
	SelectObject(hdcOutputBuffer, hfntPrev);
	DeleteObject(hfnt);
	plf->lfEscapement = 900; // 90 degree, does not work with lfFaceName "System"
	hfnt = CreateFontIndirect(plf);
	if (hfnt == NULL)
	{
		sMessage.assign(L"CreateFontIndirect@zoomMousePosition ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		goto FAIL;
	}

	hfntPrev = SelectObject(hdcOutputBuffer, hfnt);
	if (hfntPrev == NULL) goto FAIL;

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
	DrawText(hdcOutputBuffer, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP | DT_CALCRECT);

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

	DrawText(hdcOutputBuffer, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP);

	goto CLEANUP;
FAIL:
	bResult = FALSE;
	OutputDebugString(L"zoomMousePosition fails");
	if (sMessage.length() == 0)
		sMessage.assign(L"zoomMousePosition ").append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
	MessageBox(g_hWindow, sMessage.c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONERROR);

CLEANUP:
	// Free resources/Cleanup
	if (hfntPrev != NULL) SelectObject(hdcOutputBuffer, hfntPrev);

	if (hfnt != NULL) DeleteObject(hfnt);
	if (plf != NULL) LocalFree((LOCALHANDLE)plf);
	if (hBrush != NULL) DeleteObject(hBrush);

	return bResult;
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

	switch (g_appState)
	{
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
  Function: pixelateScreenshotRect

  Summary:   Pixelate rectangle area on screenshot

  Args:     RECT rect
			  Rectangle area
			DWORD blockSize
			  "Target pixel size"

  Returns:  BOOL
			  TRUE = success
			  FALSE = failure

-----------------------------------------------------------------F-F*/
BOOL pixelateScreenshotRect(RECT rect, DWORD blockSize) {
	HDC hdcScreenshot = NULL;
	HDC hdcPixelated = NULL;
	HBITMAP hBitmapPixelated = NULL;
	RECT rectPixelated = { UNINITIALIZEDLONG, UNINITIALIZEDLONG, UNINITIALIZEDLONG, UNINITIALIZEDLONG };
	HGDIOBJ hbmPixelatedOld = NULL;
	HGDIOBJ hbmScreenshotOld = NULL;
	BOOL bResult = TRUE;
	std::wstring sMessage = L"";

	if (g_hBitmap == NULL) goto FAIL;

	if (!isSelectionValid(rect)) goto FAIL;

	rectPixelated = normalizeRectangle(rect);

	hdcScreenshot = CreateCompatibleDC(NULL);
	if (hdcScreenshot == NULL) goto FAIL;

	hbmScreenshotOld = SelectObject(hdcScreenshot, g_hBitmap);
	if (hbmScreenshotOld == NULL) goto FAIL;

	hdcPixelated = CreateCompatibleDC(hdcScreenshot);
	if (hdcPixelated == NULL) goto FAIL;

	hBitmapPixelated = CreateCompatibleBitmap(hdcScreenshot, (rectPixelated.right - rectPixelated.left + 1) / blockSize, (rectPixelated.bottom - rectPixelated.top + 1) / blockSize);
	if (hBitmapPixelated == NULL)
	{
		sMessage.assign(L"CreateCompatibleBitmap@pixelateScreenshotRect ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		goto FAIL;
	}

	hbmPixelatedOld = SelectObject(hdcPixelated, hBitmapPixelated);
	if (hbmPixelatedOld == NULL) goto FAIL;

	SetStretchBltMode(hdcPixelated, HALFTONE);
	SetStretchBltMode(hdcScreenshot, COLORONCOLOR);

	// Create reduced pixels for selected area
	if (!StretchBlt(hdcPixelated, 0, 0, (rectPixelated.right - rectPixelated.left + 1) / blockSize,
		(rectPixelated.bottom - rectPixelated.top + 1) / blockSize,
		hdcScreenshot, rectPixelated.left, rectPixelated.top, rectPixelated.right - rectPixelated.left + 1,
		rectPixelated.bottom - rectPixelated.top + 1, SRCCOPY))
	{
		sMessage.assign(L"StretchBlt@pixelateScreenshotRect ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		goto FAIL;
	}

	// Copy  back to screenshot => pixelated
	if (!StretchBlt(hdcScreenshot, rectPixelated.left, rectPixelated.top, rectPixelated.right - rectPixelated.left + 1,
		rectPixelated.bottom - rectPixelated.top + 1,
		hdcPixelated, 0, 0, (rectPixelated.right - rectPixelated.left + 1) / blockSize,
		(rectPixelated.bottom - rectPixelated.top + 1) / blockSize, SRCCOPY))
	{
		sMessage.assign(L"StretchBlt@pixelateScreenshotRect ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		goto FAIL;
	}
	goto CLEANUP;
FAIL:
	bResult = FALSE;
	OutputDebugString(L"pixelateScreenshotRect fails");
	if (sMessage.length() == 0)
		sMessage.assign(L"pixelateScreenshotRect ").append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
	MessageBox(g_hWindow, sMessage.c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONERROR);
CLEANUP:
	// Free resources/Cleanup
	if (hbmPixelatedOld != NULL) SelectObject(hdcPixelated, hbmPixelatedOld);
	if (hbmScreenshotOld != NULL) SelectObject(hdcScreenshot, hbmScreenshotOld);

	if (hBitmapPixelated != NULL) DeleteObject(hBitmapPixelated);

	if (hdcPixelated != NULL) DeleteDC(hdcPixelated);
	if (hdcScreenshot != NULL) DeleteDC(hdcScreenshot);
	return bResult;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: markScreenshotRect

  Summary:   Draw line around rectangle area of screenshot

  Args:     RECT rect
			  Rectangle area
			int lineWidth
			  Width of line
			BYTE blendAlpha
			  Alpha value for line

  Returns:  BOOL
			  TRUE = success
			  FALSE = failure

-----------------------------------------------------------------F-F*/
BOOL markScreenshotRect(RECT rect, int lineWidth, BYTE blendAlpha) {
	HDC hdcScreenshot = NULL;
	HDC hdcInner = NULL;
	HDC hdcOuter = NULL;
	HBITMAP hBitmapOuter = NULL;
	HBITMAP hBitmapInner = NULL;
	HGDIOBJ hbmInnerOld = NULL;
	HGDIOBJ hbmOuterOld = NULL;
	HGDIOBJ hbmScreenshotOld = NULL;
	HBRUSH hBrush = NULL;
	RECT frame = { 0 };
	RECT inner{ 0 }, outer{ 0 };
	BLENDFUNCTION blendFunc;
	BOOL bResult = TRUE;
	std::wstring sMessage = L"";
	wchar_t szHex[_MAX_ITOSTR_BASE16_COUNT + 2];

	if (g_hBitmap == NULL) goto FAIL;

	if (lineWidth < 1) goto FAIL;

	if (!isSelectionValid(rect)) goto FAIL;

	inner = normalizeRectangle(rect);
	outer = inner;
	InflateRect(&inner, -(lineWidth / 2 + 1), -(lineWidth / 2 + 1));
	InflateRect(&outer, lineWidth / 2, lineWidth / 2);

	// hdc for screenshot bitmap
	hdcScreenshot = CreateCompatibleDC(NULL);
	if (hdcScreenshot == NULL) goto FAIL;

	hbmScreenshotOld = SelectObject(hdcScreenshot, g_hBitmap);
	if (hbmScreenshotOld == NULL) goto FAIL;

	// Copy inner of marked screenshot area to buffer
	hdcInner = CreateCompatibleDC(hdcScreenshot);
	if (hdcInner == NULL) goto FAIL;

	hBitmapInner = CreateCompatibleBitmap(hdcScreenshot, inner.right - inner.left + 1, inner.bottom - inner.top + 1);
	if (hBitmapInner == NULL)
	{
		sMessage.assign(L"CreateCompatibleBitmap@markScreenshotRect ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		goto FAIL;
	}

	hbmInnerOld = SelectObject(hdcInner, hBitmapInner);
	if (hbmInnerOld == NULL) goto FAIL;

	if (!BitBlt(hdcInner, 0, 0, inner.right - inner.left + 1, inner.bottom - inner.top + 1,
		hdcScreenshot, inner.left, inner.top, SRCCOPY)) goto FAIL;

	// Blend colored frame over outer of marked screenshot area
	hdcOuter = CreateCompatibleDC(hdcScreenshot);
	if (hdcOuter == NULL) goto FAIL;

	hBitmapOuter = CreateCompatibleBitmap(hdcScreenshot, outer.right - outer.left + 1, outer.bottom - outer.top + 1);
	if (hBitmapOuter == NULL)
	{
		sMessage.assign(L"CreateCompatibleBitmap@markScreenshotRect ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		goto FAIL;
	}

	hbmOuterOld = SelectObject(hdcOuter, hBitmapOuter);
	if (hbmOuterOld == NULL) goto FAIL;

	hBrush = CreateSolidBrush(MARKCOLOR);
	if (hBrush == NULL) goto CLEANUP;
	frame.right = outer.right - outer.left + 1;
	frame.bottom = outer.bottom - outer.top + 1;
	FillRect(hdcOuter, &frame, hBrush);

	blendFunc.BlendOp = AC_SRC_OVER;
	blendFunc.BlendFlags = 0;
	blendFunc.SourceConstantAlpha = blendAlpha;
	blendFunc.AlphaFormat = 0;

	if (!AlphaBlend(hdcScreenshot, outer.left, outer.top,
		outer.right - outer.left + 1,
		outer.bottom - outer.top + 1,
		hdcOuter, 0, 0, outer.right - outer.left + 1, outer.bottom - outer.top + 1, blendFunc))
	{
		sMessage.assign(L"AlphaBlend@markScreenshotRect ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		goto FAIL;
	}

	// Copy inner area back from buffer to screenshot
	if (!BitBlt(hdcScreenshot, inner.left, inner.top, inner.right - inner.left + 1, inner.bottom - inner.top + 1,
		hdcInner, 0, 0, SRCCOPY))
	{
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT + 2, _TRUNCATE, L"0x%08X", GetLastError());
		sMessage.assign(L"BitBlt@markScreenshotRect ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED))
			.append(L" ")
			.append(szHex);
		goto FAIL;
	}

	goto CLEANUP;
FAIL:
	bResult = FALSE;
	OutputDebugString(L"markScreenshotRect fails");
	if (sMessage.length() == 0)
		sMessage.assign(L"markScreenshotRect ").append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
	MessageBox(g_hWindow, sMessage.c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONERROR);
CLEANUP:
	// Free resources/Cleanup
	if (hBrush != NULL) DeleteObject(hBrush);

	if (hbmOuterOld != NULL) SelectObject(hdcOuter, hbmOuterOld);
	if (hbmInnerOld != NULL) SelectObject(hdcInner, hbmInnerOld);
	if (hbmScreenshotOld != NULL) SelectObject(hdcScreenshot, hbmScreenshotOld);

	if (hBitmapOuter != NULL) DeleteObject(hBitmapOuter);
	if (hBitmapInner != NULL) DeleteObject(hBitmapInner);

	if (hdcOuter != NULL) DeleteDC(hdcOuter);
	if (hdcInner != NULL) DeleteDC(hdcInner);
	if (hdcScreenshot != NULL) DeleteDC(hdcScreenshot);
	return bResult;
}

/*F+F+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
  Function: OnPaint

  Summary:   Redraw main window

  Args:     HWND hWindow
			  Handle to window

  Returns:  BOOL
			  TRUE = success
			  FALSE = failure

-----------------------------------------------------------------F-F*/
BOOL OnPaint(HWND hWindow) {
	PAINTSTRUCT ps;
	RECT rect;
	HDC hdcScreenshot = NULL;
	HGDIOBJ hbmScreenshotOld = NULL;
	HBITMAP bmOutputBuffer = NULL;
	int iBackupOutputDC = 0;
	HBRUSH hBrushForeground = NULL;
	HBRUSH hBrushBackground = NULL;
#define MAXSTRDATA 128
	wchar_t strData[MAXSTRDATA];
	PLOGFONT plf = (PLOGFONT)LocalAlloc(LPTR, sizeof(LOGFONT));
	HGDIOBJ hfnt = NULL, hfntPrev = NULL;
	RECT inner, outer;
	RECT rectText{ 0, 0, 0, 0 };
	HDC hdc = NULL;
	HDC hdcOutputBuffer = NULL;
	int iWidth = 0;
	int iHeight = 0;
	UINT textFormat = 0;
	std::wstring sDisplayInfos;
	COLORREF color;
	BOOL bResult = TRUE;
	std::wstring sMessage = L"";
	wchar_t szHex[_MAX_ITOSTR_BASE16_COUNT + 2];

	hdc = BeginPaint(hWindow, &ps);
	if (hdc == NULL) goto FAIL;
	if (plf == NULL) goto FAIL;
	if (g_appState == stateTrayIcon) goto FAIL;
	if (g_hBitmap == NULL) goto FAIL;

	GetClientRect(hWindow, &rect);

	iWidth = rect.right + 1;
	iHeight = rect.bottom + 1;

	// Create buffer memory device (Used as painting target and will be copied after the last painting to the display)
	hdcOutputBuffer = CreateCompatibleDC(hdc);
	if (hdcOutputBuffer == NULL) goto FAIL;

	// Bitmap in buffer memory
	bmOutputBuffer = CreateCompatibleBitmap(hdc, iWidth, iHeight);
	if (bmOutputBuffer == NULL)
	{
		sMessage.assign(L"CreateCompatibleBitmap@OnPaint ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		goto FAIL;
	}

	// Backup memory device context state
	iBackupOutputDC = SaveDC(hdcOutputBuffer);
	if (iBackupOutputDC == 0) goto FAIL;

	// Font
	// Specify a font typeface name and weight.
	if (_snwprintf_s(plf->lfFaceName, 12, _TRUNCATE, L"%s", DEFAULTFONT) < 0) goto FAIL;

	plf->lfWeight = FW_NORMAL;
	hfnt = CreateFontIndirect(plf);
	if (hfnt == NULL) {
		sMessage.assign(L"CreateFontIndirect@OnPaint ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		goto FAIL;
	}

	hfntPrev = SelectObject(hdcOutputBuffer, hfnt);
	if (hfntPrev == NULL) goto FAIL;

	SetTextColor(hdcOutputBuffer, g_useAlternativeColors ? ALTAPPCOLORINV : APPCOLORINV);
	SetBkColor(hdcOutputBuffer, g_useAlternativeColors ? ALTAPPCOLOR : APPCOLOR);

	SelectObject(hdcOutputBuffer, bmOutputBuffer);

	// Get screenshot bitmap information and select bitmap
	BITMAP bm;
	if (GetObject(g_hBitmap, sizeof(bm), &bm) == 0) goto FAIL;
	hdcScreenshot = CreateCompatibleDC(hdc);
	if (hdcScreenshot == NULL) goto FAIL;
	hbmScreenshotOld = SelectObject(hdcScreenshot, g_hBitmap);
	if (hbmScreenshotOld == NULL) goto FAIL;

	// Use a blen function to get a darker image of the screenshot
	BLENDFUNCTION blendFunc;
	blendFunc.BlendOp = AC_SRC_OVER;
	blendFunc.BlendFlags = 0;
	blendFunc.SourceConstantAlpha = g_useAlternativeColors ? 255 : 50; // Factor to darken the screenshot
	blendFunc.AlphaFormat = 0;

	if (!AlphaBlend(hdcOutputBuffer, 0, 0, bm.bmWidth, bm.bmHeight, hdcScreenshot, 0, 0, bm.bmWidth, bm.bmHeight, blendFunc))
	{
		sMessage.assign(L"AlphaBlend@OnPaint ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		goto FAIL;
	}

	hBrushForeground = CreateSolidBrush(g_useAlternativeColors ? ALTAPPCOLOR : APPCOLOR);
	if (hBrushForeground == NULL) goto FAIL;
	hBrushBackground = CreateSolidBrush(ALTAPPCOLORINV);
	if (hBrushBackground == NULL) goto FAIL;

	// Get inner/outer rects and zoom mouse position
	switch (g_appState)
	{
	case stateFirstPoint:
		inner.left = g_selection.left;
		inner.top = g_selection.top;
		zoomMousePosition(hdcOutputBuffer, BoxFirstPointA);
		break;
	case statePointA:
	case statePointB:
	{
		inner = normalizeRectangle(g_selection);
		if (inner.left < 0) inner.left = 0;
		if (inner.right > bm.bmWidth - 1) inner.right = bm.bmWidth - 1;
		if (inner.top < 0) inner.top = 0;
		if (inner.bottom > bm.bmHeight - 1) inner.bottom = bm.bmHeight - 1;

		outer.left = inner.left - 1;
		outer.right = inner.right + 1 + 1; // +1 because GDI the second edge of a rect is not part of the drawn rectangle
		outer.top = inner.top - 1;
		outer.bottom = inner.bottom + 1 + 1; // +1 because GDI the second edge of a rect is not part of the drawn rectangle

		// Show selected area on the darkened background of the screenshot
		if (!BitBlt(hdcOutputBuffer, inner.left, inner.top, inner.right - inner.left + 1, inner.bottom - inner.top + 1,
			hdcScreenshot, inner.left, inner.top, SRCCOPY))
		{
			_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT + 2, _TRUNCATE, L"0x%08X", GetLastError());
			sMessage.assign(L"BitBlt@OnPaint ")
				.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED))
				.append(L" ")
				.append(szHex);
			goto FAIL;
		}

		// Draw frame
		FrameRect(hdcOutputBuffer, &outer, hBrushForeground);

		// Draw text for selection width
		rectText = { 0, 0, 0, 0 };
		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"%d", inner.right - inner.left + 1);
		DrawText(hdcOutputBuffer, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP | DT_CALCRECT);

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
		SetTextColor(hdcOutputBuffer, g_useAlternativeColors ? ALTAPPCOLORINV : APPCOLORINV);
		SetBkColor(hdcOutputBuffer, g_useAlternativeColors ? ALTAPPCOLOR : APPCOLOR);

		// Draw text, when enough splace
		if (abs(g_selection.right - g_selection.left) >= (long) (ZOOMWIDTH * g_zoomScale))
			DrawText(hdcOutputBuffer, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP | DT_CENTER | DT_BOTTOM);

		// Draw mouse position
		zoomMousePosition(hdcOutputBuffer, BoxFinalPointA);
		zoomMousePosition(hdcOutputBuffer, BoxFinalPointB);

		// Draw text for selection height
		rectText = { 0, 0, 0, 0 };

		// Text rotated 90 degree
		SelectObject(hdc, hfntPrev);
		DeleteObject(hfnt);
		plf->lfEscapement = 900; // 90 degree, does not work with lfFaceName "System"
		hfnt = CreateFontIndirect(plf);
		if (hfnt == NULL)
		{
			sMessage.assign(L"CreateFontIndirect@OnPaint ")
				.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
			goto FAIL;
		}

		hfntPrev = SelectObject(hdcOutputBuffer, hfnt);
		if (hfntPrev == NULL) goto FAIL;

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"%d", inner.bottom - inner.top + 1);
		DrawText(hdcOutputBuffer, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP | DT_CALCRECT);

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
		SetTextColor(hdcOutputBuffer, g_useAlternativeColors ? ALTAPPCOLORINV : APPCOLORINV);
		SetBkColor(hdcOutputBuffer, g_useAlternativeColors ? ALTAPPCOLOR : APPCOLOR);

		// Draw text, when enough space
		if (abs(g_selection.bottom - g_selection.top) >= (LONG) (ZOOMHEIGHT * g_zoomScale))
			DrawText(hdcOutputBuffer, strData, -1, &rectText, DT_SINGLELINE | DT_NOCLIP);

		break;
	}
	default:
		OutputDebugString(L"Invalid appState");
		goto FAIL;
	}

	// Draw information
	if (g_displayInternalInformation)
	{
		HBRUSH hBrushDisplayForeground = (g_useAlternativeColors ? hBrushBackground : hBrushForeground);
		HBRUSH hBrushDisplayBackground = (g_useAlternativeColors ? hBrushForeground : hBrushBackground);

		RECT rectTextArea = { 0, 0, 0, 0 };
		// Text rotated 0 degree
		SelectObject(hdc, hfntPrev);
		DeleteObject(hfnt);
		plf->lfEscapement = 0;
		hfnt = CreateFontIndirect(plf);
		if (hfnt == NULL)
		{
			sMessage.assign(L"CreateFontIndirect@OnPaint ")
				.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
			goto FAIL;
		}

		hfntPrev = SelectObject(hdcOutputBuffer, hfnt);
		if (hfntPrev == NULL) goto FAIL;

		SetBkMode(hdcOutputBuffer, TRANSPARENT);
		if (!g_useAlternativeColors)
			SetTextColor(hdcOutputBuffer, APPCOLOR);
		else
			SetTextColor(hdcOutputBuffer, ALTAPPCOLORINV);

		int screenX = GetSystemMetrics(SM_XVIRTUALSCREEN);
		int screenY = GetSystemMetrics(SM_YVIRTUALSCREEN);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Virtual desktop [%d,%d] %dx%d", screenX, screenY, GetSystemMetrics(SM_CXVIRTUALSCREEN), GetSystemMetrics(SM_CYVIRTUALSCREEN));
		sDisplayInfos.assign(strData);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Selection [%d,%d] [%d,%d]", g_selection.left, g_selection.top, g_selection.right, g_selection.bottom);
		sDisplayInfos.append(L"\n").append(strData);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Stored selection [%d,%d] [%d,%d]", g_storedSelection.left, g_storedSelection.top, g_storedSelection.right, g_storedSelection.bottom);
		sDisplayInfos.append(L"\n").append(strData);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Bitmap %dx%d", bm.bmWidth, bm.bmHeight);
		sDisplayInfos.append(L"\n").append(strData);

		POINT mouse;
		GetCursorPos(&mouse);
		color = GetPixel(hdcScreenshot, mouse.x - g_appWindowPos.x, mouse.y - g_appWindowPos.y);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Mouse [%d,%d] RGB %d,%d,%d", mouse.x, mouse.y, GetRValue(color), GetGValue(color), GetBValue(color));
		sDisplayInfos.append(L"\n").append(strData);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Save to file %s", g_saveToFile ? L"On" : L"Off");
		sDisplayInfos.append(L"\n").append(strData);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Save to clipboard %s", g_saveToClipboard ? L"On" : L"Off");
		sDisplayInfos.append(L"\n").append(strData);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Alternative colors %s", g_useAlternativeColors ? L"On" : L"Off");
		sDisplayInfos.append(L"\n").append(strData);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"State %d appWindow [%d,%d]", g_appState, g_appWindowPos.x, g_appWindowPos.y);
		sDisplayInfos.append(L"\n").append(strData);

		RECT rectWindow = { 0 };
		GetWindowRect(hWindow, &rectWindow);

		if ((rectWindow.left != screenX) || (rectWindow.top != screenY)) {
			_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L" ([%d,%d]!=[%d,%d])", rectWindow.left, rectWindow.top, screenX, screenY);
			sDisplayInfos.append(strData);
		}

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L" ([%d,%d][%d,%d])", rectWindow.left, rectWindow.top, rectWindow.right, rectWindow.bottom);
		sDisplayInfos.append(strData);

		_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L" Has focus %s Selected Monitor %d", (hWindow == GetForegroundWindow()) ? L"Yes" : L"No", g_selectedMonitor);
		sDisplayInfos.append(strData);

		for (int i = 0; i < (int)g_rectMonitor.size(); i++)
		{
			_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"Monitor %d [%d,%d] [%d,%d]", i, g_rectMonitor[i].left, g_rectMonitor[i].top, g_rectMonitor[i].right, g_rectMonitor[i].bottom);
			sDisplayInfos.append(L"\n").append(strData);
		}

		sDisplayInfos.append(L"\n\nA = Select all\nM = Select next monitor\nTab = A <-> B\nCursor keys = Move A/B\n")
			.append(L"Alt+cursor keys = Fast move A/B\nShift+cursor keys = Find color change for A/B\nReturn = OK\nESC = Cancel")
			.append(L"\n+/- = Increase/decrease selection")
			.append(L"\nPageUp/PageDown, mouse wheel = Zoom In/Out")
			.append(L"\nInsert = Store selection\nHome = Use stored selection\nDelete = Delete stored and used selection\nP = Pixelate selection\nB = Box around selection");
		if (!g_bSaveToClipboardGPO) sDisplayInfos.append(L"\nC = Clipboard On/Off");
		if (!g_bSaveToFileGPO) sDisplayInfos.append(L"\nF = File On/Off");
		sDisplayInfos.append(L"\nS = Alternative colors On/Off");
		if (!g_bDisplayInternalInformationGPO) sDisplayInfos.append(L"\nF1 = Display information On/Off");

		// Calc text area
		DrawText(hdcOutputBuffer, sDisplayInfos.c_str(), -1, &rectTextArea, DT_NOCLIP | DT_CALCRECT);

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

		if (g_useAlternativeColors) FillRect(hdcOutputBuffer, &rectTextArea, hBrushDisplayBackground); // Text area background

		// Draw text in text area
		rectText.left = rectTextArea.left + 1;
		rectText.right = rectTextArea.right - 1;
		rectText.top = rectTextArea.top;
		DrawText(hdcOutputBuffer, sDisplayInfos.c_str(), -1, &rectText, DT_NOCLIP | textFormat);

		// Draw monitor layout
		float scale = (float)(rectTextArea.right - rectTextArea.left) / GetSystemMetrics(SM_CXVIRTUALSCREEN);
		RECT virtualDesktop;
		virtualDesktop.left = rectTextArea.left;
		virtualDesktop.top = rectTextArea.bottom + 10;
		virtualDesktop.right = virtualDesktop.left + (LONG)((GetSystemMetrics(SM_CXVIRTUALSCREEN) - 1) * scale) + 1;
		virtualDesktop.bottom = virtualDesktop.top + (LONG)((GetSystemMetrics(SM_CYVIRTUALSCREEN) - 1) * scale) + 1;
		FrameRect(hdcOutputBuffer, &virtualDesktop, hBrushDisplayForeground); // Virtual desktop frame

		if (g_useAlternativeColors)
		{
			FillRect(hdcOutputBuffer, &virtualDesktop, hBrushDisplayBackground); // Background fill
		}

		for (DWORD i = 0; i < g_rectMonitor.size(); i++)
		{
			RECT monitor;

			monitor.left = virtualDesktop.left + (LONG)((g_rectMonitor[i].left - g_appWindowPos.x) * scale);
			monitor.top = virtualDesktop.top + (LONG)((g_rectMonitor[i].top - g_appWindowPos.y) * scale);
			monitor.right = monitor.left + (LONG)((g_rectMonitor[i].right - g_rectMonitor[i].left - 1) * scale) + 1;
			monitor.bottom = monitor.top + (LONG)((g_rectMonitor[i].bottom - g_rectMonitor[i].top - 1) * scale) + 1;

			FrameRect(hdcOutputBuffer, &monitor, hBrushDisplayForeground);

			_snwprintf_s(strData, MAXSTRDATA, _TRUNCATE, L"%d", i);
			DrawText(hdcOutputBuffer, strData, -1, &monitor, DT_SINGLELINE | DT_NOCLIP | DT_CENTER | DT_VCENTER);

		}
		if (isSelectionValid(g_selection))
		{
			RECT selection;
			if (g_selection.right >= g_selection.left) {
				selection.left = virtualDesktop.left + (LONG)(g_selection.left * scale);
				selection.right = virtualDesktop.left + (LONG)(g_selection.right * scale) + 1;
			}
			else
			{
				selection.left = virtualDesktop.left + (LONG)(g_selection.right * scale);
				selection.right = virtualDesktop.left + (LONG)(g_selection.left * scale) + 1;
			}
			if (g_selection.bottom >= g_selection.top) {
				selection.top = virtualDesktop.top + (LONG)(g_selection.top * scale);
				selection.bottom = virtualDesktop.top + (LONG)(g_selection.bottom * scale) + 1;
			}
			else
			{
				selection.top = virtualDesktop.top + (LONG)(g_selection.bottom * scale);
				selection.bottom = virtualDesktop.top + (LONG)(g_selection.top * scale) + 1;
			}

			FrameRect(hdcOutputBuffer, &selection, hBrushDisplayForeground);
		}
		else
		{
			if ((g_selection.left != UNINITIALIZEDLONG) && (g_selection.top != UNINITIALIZEDLONG)) { // Quick and dirty "pixel" at mouse cursor position
				RECT pixel;
				pixel.left = virtualDesktop.left + (LONG)(g_selection.left * scale) - 1;
				pixel.top = virtualDesktop.top + (LONG)(g_selection.top * scale) - 1;
				pixel.right = pixel.left + 3;
				pixel.bottom = pixel.top + 3;
				FrameRect(hdcOutputBuffer, &pixel, hBrushDisplayForeground);
			}
		}
	}

	// Copy memory buffer to display
	if (!BitBlt(hdc, 0, 0, iWidth, iHeight, hdcOutputBuffer, 0, 0, SRCCOPY))
	{
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT + 2, _TRUNCATE, L"0x%08X", GetLastError());
		sMessage.assign(L"BitBlt@OnPaint ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED))
			.append(L" ")
			.append(szHex);
		goto FAIL;
	}

	goto CLEANUP;
FAIL:
	bResult = FALSE;
	OutputDebugString(L"OnPaint fails");
	if (sMessage.length() == 0)
		sMessage.assign(L"OnPaint ").append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
	MessageBox(hWindow, sMessage.c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONERROR);

CLEANUP:
	// Free resources/Cleanup
	if (hBrushForeground != NULL) DeleteObject(hBrushForeground);
	if (hBrushBackground != NULL) DeleteObject(hBrushBackground);

	if (hbmScreenshotOld != NULL) SelectObject(hdcScreenshot, hbmScreenshotOld);
	if (hdcOutputBuffer != NULL)
	{
		// Restore memory device context state
		if (hfntPrev != NULL) SelectObject(hdcOutputBuffer, hfntPrev);
		if (iBackupOutputDC != 0) RestoreDC(hdcOutputBuffer, iBackupOutputDC);
	}
	if (hfnt != NULL) DeleteObject(hfnt);
	if (plf != NULL) LocalFree((LOCALHANDLE)plf);
	if (bmOutputBuffer != NULL) DeleteObject(bmOutputBuffer);

	if (hdcOutputBuffer != NULL) DeleteDC(hdcOutputBuffer);
	if (hdcScreenshot != NULL) DeleteDC(hdcScreenshot);

	EndPaint(hWindow, &ps);

	return bResult;
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

	if ((stepSize < 0) && (abs(g_selection.right - g_selection.left) < abs(stepSize * 2)))
	{
		// Width too small for decreasing by stepsize
		g_selection.left = limitXtoBitmap((g_selection.right + g_selection.left) / 2);
		g_selection.right = g_selection.left;
	}
	else
	{

		if (g_selection.left <= g_selection.right)
		{
			g_selection.left = limitXtoBitmap(g_selection.left - stepSize);
			g_selection.right = limitXtoBitmap(g_selection.right + stepSize);
		}
		else
		{
			g_selection.left = limitXtoBitmap(g_selection.left + stepSize);
			g_selection.right = limitXtoBitmap(g_selection.right - stepSize);
		}
	}

	if ((stepSize < 0) && (abs(g_selection.top - g_selection.bottom) < abs(stepSize * 2)))
	{
		// Height too small for decreasing by stepsize
		g_selection.top = limitYtoBitmap((g_selection.top + g_selection.top) / 2);
		g_selection.bottom = g_selection.top;
	}
	else
	{
		if (g_selection.top <= g_selection.bottom)
		{
			g_selection.top = limitYtoBitmap(g_selection.top - stepSize);
			g_selection.bottom = limitYtoBitmap(g_selection.bottom + stepSize);
		}
		else
		{
			g_selection.top = limitYtoBitmap(g_selection.top + stepSize);
			g_selection.bottom = limitYtoBitmap(g_selection.bottom - stepSize);
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

  Returns:	BOOL
			  TRUE = success
			  FALSE = failure

-----------------------------------------------------------------F-F*/
BOOL setBeforeColorChange(WPARAM virtualKeyCode, LONG& x, LONG& y)
{
	BITMAP bm;
	int directionX = 0;
	int directionY = 0;
	COLORREF referenceColor;
	BOOL bResult = TRUE;
	HDC hdcScreenshot = NULL;
	HGDIOBJ hbmScreenshotOld = NULL;
	std::wstring sMessage = L"";

	if (g_hBitmap == NULL) goto FAIL;

	if (GetObject(g_hBitmap, sizeof(bm), &bm) == 0) goto FAIL;
	hdcScreenshot = CreateCompatibleDC(NULL);
	if (hdcScreenshot == NULL) goto FAIL;
	hbmScreenshotOld = SelectObject(hdcScreenshot, g_hBitmap);
	if (hbmScreenshotOld == NULL) goto FAIL;

	referenceColor = GetPixel(hdcScreenshot, x, y);
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
		goto FAIL;
	}

	while (true) {
		if (GetPixel(hdcScreenshot, x + directionX, y + directionY) != referenceColor) break;

		if (x + directionX < 0) break;
		if (x + directionX > bm.bmWidth - 1) break;
		if (y + directionY < 0) break;
		if (y + directionY > bm.bmHeight - 1) break;

		x = limitXtoBitmap(x + directionX);
		y = limitYtoBitmap(y + directionY);
	}
	goto CLEANUP;
FAIL:
	bResult = FALSE;
	OutputDebugString(L"setBeforeColorChange fails");
	if (sMessage.length() == 0)
		sMessage.assign(L"setBeforeColorChange ").append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
	MessageBox(g_hWindow, sMessage.c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONERROR);

CLEANUP:
	if (hbmScreenshotOld != NULL) SelectObject(hdcScreenshot, hbmScreenshotOld);
	if (hdcScreenshot != NULL) DeleteDC(hdcScreenshot);

	return bResult;
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

  Returns:	BOOL
			  TRUE = success
			  FALSE = failure

-----------------------------------------------------------------F-F*/
BOOL CaptureScreen(HWND hWindow)
{
	HDC hdcScreen = NULL;
	HDC hdcScreenshot = NULL;
	HGDIOBJ hbmScreenshotOld = NULL;
	BOOL bResult = TRUE;
	std::wstring sMessage = L"";
	wchar_t szHex[_MAX_ITOSTR_BASE16_COUNT + 2];

	int screenX = GetSystemMetrics(SM_XVIRTUALSCREEN);
	int screenY = GetSystemMetrics(SM_YVIRTUALSCREEN);
	int screenWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
	int screenHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);

	// Retrieve the handle to a display device context for the client
	// area of the window.
	hdcScreen = GetDC(NULL);
	if (hdcScreen == NULL) goto FAIL;

	// Create a compatible DC, which is used in a BitBlt from the window DC.
	hdcScreenshot = CreateCompatibleDC(hdcScreen);

	if (!hdcScreenshot)
	{
		sMessage.assign(L"CreateCompatibleDC@CaptureScreen ").append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		goto FAIL;
	}

	if (g_hBitmap != NULL) { // Delete previous screenshot
		DeleteObject(g_hBitmap);
		g_hBitmap = NULL;
	}

	// Create a compatible bitmap from the Window DC.
	g_hBitmap = CreateCompatibleBitmap(hdcScreen, screenWidth, screenHeight);
	if (g_hBitmap == NULL)
	{
		sMessage.assign(L"CreateCompatibleBitmap@CaptureScreen ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
		goto FAIL;
	}
	// Select the compatible bitmap into the compatible memory DC.
	hbmScreenshotOld = SelectObject(hdcScreenshot, g_hBitmap);
	if (hbmScreenshotOld == NULL) goto FAIL;

	// Bit block transfer into our compatible memory DC.
	if (!BitBlt(hdcScreenshot,
		0, 0,
		screenWidth, screenHeight,
		hdcScreen,
		screenX, screenY,
		SRCCOPY))
	{
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT + 2, _TRUNCATE, L"0x%08X", GetLastError());
		sMessage.assign(L"BitBlt@CaptureScreen ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED))
			.append(L" ")
			.append(szHex);
		goto FAIL;
	}
	goto CLEANUP;
FAIL:
	bResult = FALSE;
	OutputDebugString(L"CaptureScreen fails");
	if (sMessage.length() == 0)
		sMessage.assign(L"CaptureScreen ").append(LoadStringAsWstr(g_hInst, IDS_HASFAILED));
	MessageBox(hWindow, sMessage.c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONERROR);
CLEANUP:
	// Free resources/Clean up
	if (hbmScreenshotOld != NULL) SelectObject(hdcScreenshot, hbmScreenshotOld);
	if (hdcScreenshot != NULL) DeleteDC(hdcScreenshot);
	if (hdcScreen != NULL) ReleaseDC(NULL, hdcScreen);
	return bResult;
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

	KillTimer(hWindow, IDT_TIMERSCREENSHOTDELAYED);

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
	getDWORDSettingFromRegistry(defaultZoomScale);
	getDWORDSettingFromRegistry(screenshotDelay);
	getDWORDSettingFromRegistry(saveToFile);
	getDWORDSettingFromRegistry(saveToClipboard);
	getDWORDSettingFromRegistry(useAlternativeColors);
	getDWORDSettingFromRegistry(displayInternalInformation);
	getDWORDSettingFromRegistry(storedSelectionLeft);
	getDWORDSettingFromRegistry(storedSelectionTop);
	getDWORDSettingFromRegistry(storedSelectionRight);
	getDWORDSettingFromRegistry(storedSelectionBottom);
	getScreenshotPathFromRegistry();

	enterFullScreen(hWindow);
	ShowWindow(hWindow, SW_NORMAL);
	ShowCursor(false);

	if (isSelectionValid(g_storedSelection))
	{
		g_appState = statePointB;
		g_selection.left = limitXtoBitmap(g_storedSelection.left);
		g_selection.right = limitXtoBitmap(g_storedSelection.right);
		g_selection.top = limitYtoBitmap(g_storedSelection.top);
		g_selection.bottom = limitYtoBitmap(g_storedSelection.bottom);
		MySetCursorPos(g_selection.right, g_selection.bottom);
	}
	else {
		POINT mouse;
		GetCursorPos(&mouse);
		g_appState = stateFirstPoint;
		g_selection.left = limitXtoBitmap(mouse.x - g_appWindowPos.x);
		g_selection.top = limitYtoBitmap(mouse.y - g_appWindowPos.y);
		g_selection.right = UNINITIALIZEDLONG;
		g_selection.bottom = UNINITIALIZEDLONG;
	}

	// Enable 1s time to show selected point
	SetTimer(hWindow, IDT_TIMER1000MS, 1000, (TIMERPROC)NULL);

	// Final FIX01: Check window position, which is sometimes wrong (perhaps a timing problem or caused by Omnissa Horizon Client)
	RECT rectWindow = { 0 };
	GetWindowRect(hWindow, &rectWindow);
	int screenX = GetSystemMetrics(SM_XVIRTUALSCREEN);
	int screenY = GetSystemMetrics(SM_YVIRTUALSCREEN);

	if ((rectWindow.left != screenX) || (rectWindow.top != screenY)) { // Reapply fullscreen, if window position is not correct
		Sleep(500);
		enterFullScreen(hWindow);
	}

	// Force foreground window (Prevents keyboard input focus problems on a second or third monitor)
	SetForegroundWindowInternal(hWindow);
	Sleep(10);
	if (hWindow != GetForegroundWindow()) forceFocus(hWindow);
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
  Function: checkArguments

  Summary:   Check program arguments

  Args:

  Returns:  BOOL
			  TRUE = Keep program running
			  FALSE = Exit wWinMain afterwards

-----------------------------------------------------------------F-F*/
BOOL checkArguments()
{
	int argc;
	LPWSTR* argv = CommandLineToArgvW(GetCommandLine(), &argc);
	BOOL bExit = FALSE;

	if (argv == NULL) {
		OutputDebugString(L"Argv fails");
		return FALSE;
	}

	bool bAutoSaveToClipboard = FALSE;
	bool bAutoSaveToFile = FALSE;
	getScreenshotPathFromRegistry();

	for (int i = 0; i < argc; i++)
	{
		if (_wcsicmp(argv[i], L"/re") == 0) {
			setRunKeyRegistryValue(TRUE, HKEY_LOCAL_MACHINE);
			bExit = TRUE;
			break;
		}
		if (_wcsicmp(argv[i], L"/rd") == 0) {
			setRunKeyRegistryValue(FALSE, HKEY_LOCAL_MACHINE);
			bExit = TRUE;
			break;
		}
		if (_wcsicmp(argv[i], L"/ac") == 0) bAutoSaveToClipboard = TRUE;
		if (_wcsicmp(argv[i], L"/af") == 0) bAutoSaveToFile = TRUE;
		if (_wcsicmp(argv[i], L"/f") == 0) {
			ShellExecute(NULL, L"open", g_screenshotPath, NULL, NULL, SW_SHOWNORMAL); // Open screenshot folder
			bExit = TRUE;
			break;
		}
		if (_wcsicmp(argv[i], L"/s") == 0) g_onetimeCapture = TRUE; // Enable onetimeCapture mode (Program will exit afterwards automatically)
		if (_wcsicmp(argv[i], L"/v") == 0)
		{
			showProgramInformation(NULL);
			bExit = TRUE;
			break;
		}
		if (_wcsicmp(argv[i], L"/?") == 0)
		{
			showProgramArguments(NULL);
			bExit = TRUE;
			break;
		}
	}

	LocalFree(argv);

	if (bExit) return FALSE; // Exit wWinMain afterwards

	if (bAutoSaveToClipboard || bAutoSaveToFile)
	{
		// Enable only target passed by arguments
		g_saveToFile = FALSE;
		g_saveToClipboard = FALSE;
		if (bAutoSaveToClipboard) g_saveToClipboard = TRUE;
		if (bAutoSaveToFile) g_saveToFile = TRUE;
		CaptureScreen(NULL);
		if (g_hBitmap == NULL) return FALSE; // Error => Exit wWinMain afterwards
		BITMAP bm;
		if (GetObject(g_hBitmap, sizeof(bm), &bm) != 0)
		{
			g_selection.left = limitXtoBitmap(0);
			g_selection.top = limitYtoBitmap(0);
			g_selection.right = limitXtoBitmap(bm.bmWidth - 1);
			g_selection.bottom = limitYtoBitmap(bm.bmHeight - 1);
			saveSelection(NULL);
		}
		return FALSE; // Finished => Exit wWinMain afterwards
	}
	return TRUE; // Keep wWinMain running
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
	std::wstring sMessage = L"";
	wchar_t szHex[_MAX_ITOSTR_BASE16_COUNT + 2];
	MSG msg;

	g_hInst = hInstance; // Store instance handle in global variable

	// Check for temporary fixes
	getDWORDSettingFromRegistry(DEV);

	// Arguments
	if (!checkArguments()) return 0;

	// Semaphore to prevent concurrent actions
	g_hSemaphoreModalBlocked = CreateSemaphore(NULL, 1, 1, NULL);
	if (g_hSemaphoreModalBlocked == NULL)
	{
		OutputDebugString(L"Error creating semaphore");
		return 0;
	}

	HANDLE hMutex = CreateMutex(NULL, TRUE, LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str());

	// Prevents concurrent program starts
	if (!g_onetimeCapture && (GetLastError() == ERROR_ALREADY_EXISTS))
	{
		OutputDebugString(L"Program already startet");
		return 0;
	}

	// Get notification when windows explorer crashes and re-launches the taskbar
	WM_TASKBARCREATED = RegisterWindowMessageW(L"TaskbarCreated");

	// Register class
	MyRegisterClass(hInstance);

	// Create main window
	g_hWindow = CreateWindow(L"MainWndClass", LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

	if (!g_hWindow) { // Error
		_snwprintf_s(szHex, _MAX_ITOSTR_BASE16_COUNT + 2, _TRUNCATE, L"0x%08X", GetLastError());
		sMessage.assign(L"CreateWindow@wWinMain ")
			.append(LoadStringAsWstr(g_hInst, IDS_HASFAILED))
			.append(L" ")
			.append(szHex);
		MessageBox(NULL, sMessage.c_str(), LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str(), MB_OK | MB_ICONERROR);
		if (hMutex != NULL)
		{
			ReleaseMutex(hMutex);
			CloseHandle(hMutex);
		}
		return 1;
	}

	// Get settings from registry
	getDWORDSettingFromRegistry(saveToClipboard);
	getDWORDSettingFromRegistry(saveToFile);
	checkScreenshotTargets(g_hWindow);
	isRunKeyEnabledFromRegistry();

	if (g_onetimeCapture) SendMessage(g_hWindow, WM_STARTED, 0, 0); // onetimeCapture mode
	else
	{
		// Add tray icon entry
		g_nid.cbSize = sizeof(NOTIFYICONDATA);
		g_nid.hWnd = g_hWindow;
		g_nid.uID = 1;
		g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
		g_nid.uCallbackMessage = WM_TRAYICON;
		g_nid.hIcon = LoadIcon(GetModuleHandle(NULL), MAKEINTRESOURCE(IDI_ICON));
#define MAXTOOLTIPSIZE 64 // Including null termination (https://learn.microsoft.com/de-de/windows/win32/api/shellapi/ns-shellapi-notifyicondataw)
		_snwprintf_s(g_nid.szTip, MAXTOOLTIPSIZE, _TRUNCATE, L"%s", LoadStringAsWstr(g_hInst, IDS_APP_TITLE).c_str());
		Shell_NotifyIcon(NIM_ADD, &g_nid);

		checkPrintScreenKeyForSnipping(g_hWindow);

		SetHook();
	}

	// Main message loop:
	while (GetMessage(&msg, nullptr, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	if (!g_onetimeCapture) // If not onetimeCapture mode
	{
		ReleaseHook();
		// Remove tray icon entry
		Shell_NotifyIcon(NIM_DELETE, &g_nid);

		// Free Mutex to allow next program start
		if (hMutex != NULL)
		{
			ReleaseMutex(hMutex);
			CloseHandle(hMutex);
		}
	}

	// Close semaphore handles
	if (g_hSemaphoreModalBlocked != NULL) CloseHandle(g_hSemaphoreModalBlocked);

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
		if (g_zoomScale > MAXZOOMSCALE) g_zoomScale = MAXZOOMSCALE;
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
		if (GetObject(g_hBitmap, sizeof(bm), &bm) != 0)
		{
			g_appState = statePointB;
			g_selection.left = limitXtoBitmap(0);
			g_selection.top = limitYtoBitmap(0);
			g_selection.right = limitXtoBitmap(bm.bmWidth - 1);
			g_selection.bottom = limitYtoBitmap(bm.bmHeight - 1);
			InvalidateRect(hWnd, NULL, TRUE);
			// Do not SetCursorPos, because this can make trouble on multimonitor systems with different resolutions
		}
		break;
	}
	case WM_STARTED: // Start new capture
	{
		// Skip capture, when a modal dialog is running
		if (WaitForSingleObject(g_hSemaphoreModalBlocked, 0) != WAIT_OBJECT_0) break;
		ReleaseSemaphore(g_hSemaphoreModalBlocked, 1, NULL);

		startCaptureGUI(hWnd);
		break;
	}
	case WM_GOTOTRAY: // Hide window and goto tray icon
	{
		if (g_onetimeCapture) DestroyWindow(hWnd); // Exit program in onetimeCapture mode
		KillTimer(hWnd, IDT_TIMER1000MS);
		KillTimer(hWnd, IDT_TIMERSCREENSHOTDELAYED);
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
			getDWORDSettingFromRegistry(defaultZoomScale);
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
			// Skip menu, when a modal dialog is running
			if (WaitForSingleObject(g_hSemaphoreModalBlocked, 0) != WAIT_OBJECT_0)
			{
				SetForegroundWindow(hWnd); // Bring modal dialog in foreground
				break;
			}
			ReleaseSemaphore(g_hSemaphoreModalBlocked, 1, NULL);

			// Update some vars from registry
			getDWORDSettingFromRegistry(saveToFile);
			getDWORDSettingFromRegistry(saveToClipboard);
			getDWORDSettingFromRegistry(screenshotDelay);
			getScreenshotPathFromRegistry();
			BOOL bAutorun = isRunKeyEnabledFromRegistry();

			POINT pt;
			GetCursorPos(&pt);
			HMENU hMenu = CreatePopupMenu();
			wchar_t szMenuEntry[MAX_PATH] = L"";
			_snwprintf_s(szMenuEntry, MAX_PATH, _TRUNCATE, LoadStringAsWstr(g_hInst, IDS_SCREENSHOTDELAYED).c_str(), g_screenshotDelay);
			AppendMenu(hMenu, MF_STRING, IDM_CAPTURE, szMenuEntry);
			if (!g_sLastScreenshotFile.empty() && PathFileExists(g_sLastScreenshotFile.c_str()) ) {
				AppendMenu(hMenu, MF_STRING, IDM_OPENLAST, LoadStringAsWstr(g_hInst, IDS_OPENLAST).c_str());
				if (IsWindows11_24H2OrNewer()) AppendMenu(hMenu, MF_STRING, IDM_EDITLAST, LoadStringAsWstr(g_hInst, IDS_EDITLAST).c_str());
			} else g_sLastScreenshotFile = L"";
			AppendMenu(hMenu, MF_STRING, IDM_OPENFOLDER, LoadStringAsWstr(g_hInst, IDS_OPENFOLDER).c_str());
			AppendMenu(hMenu, MF_STRING | (g_bScreenshotPathGPO ? MF_GRAYED : 0) , IDM_SETFOLDER, LoadStringAsWstr(g_hInst, IDS_SETFOLDER).c_str());
			AppendMenu(hMenu, MF_STRING | (g_saveToClipboard ? MF_CHECKED : 0) | (g_bSaveToClipboardGPO ? MF_GRAYED : 0), IDM_SAVETOCLIPBOARD, LoadStringAsWstr(g_hInst, IDS_SAVETOCLIPBOARD).c_str());
			AppendMenu(hMenu, MF_STRING | (g_saveToFile ? MF_CHECKED : 0) | (g_bSaveToFileGPO ? MF_GRAYED : 0), IDM_SAVETOFILE, LoadStringAsWstr(g_hInst, IDS_SAVETOFILE).c_str());
			AppendMenu(hMenu, MF_SEPARATOR | MF_BYPOSITION, 0, NULL);
			AppendMenu(hMenu, MF_STRING, IDM_ABOUT, LoadStringAsWstr(g_hInst, IDS_ABOUT).c_str());
			AppendMenu(hMenu, MF_STRING | (bAutorun ? MF_CHECKED : 0) | (g_bRunKeyReadOnly ? MF_GRAYED : 0), IDM_AUTORUN, LoadStringAsWstr(g_hInst, IDS_AUTORUN).c_str());
			AppendMenu(hMenu, MF_STRING, IDM_EXIT, LoadStringAsWstr(g_hInst, IDS_EXIT).c_str());
			SetForegroundWindow(hWnd);
			TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hWnd, NULL);
			DestroyMenu(hMenu);
			break;
		}
		case WM_LBUTTONDBLCLK: // Double click on tray icon => Open screenshot folder
			getScreenshotPathFromRegistry();
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
		case VK_NEXT: // Page down => Zoom out
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
		case 'S': // S => Toggle colors
			SendMessage(hWnd, WM_COMMAND, IDM_ALTERNATIVECOLORS, 0);
			break;
		case 'P': // P => Pixelate
			if ((g_appState == statePointB) || (g_appState == statePointA))
			{
				pixelateScreenshotRect(g_selection, PIXELATEFACTOR);
				g_selection.right = UNINITIALIZEDLONG;
				g_selection.bottom = UNINITIALIZEDLONG;
				g_appState = stateFirstPoint;
				MySetCursorPos(g_selection.left, g_selection.top);
				InvalidateRect(hWnd, NULL, TRUE);
			}
			break;
		case 'B': // B => Mark selection
			if ((g_appState == statePointB) || (g_appState == statePointA))
			{
				markScreenshotRect(g_selection, MARKEDWIDTH, MARKEDALPHA);
				g_selection.right = UNINITIALIZEDLONG;
				g_selection.bottom = UNINITIALIZEDLONG;
				g_appState = stateFirstPoint;
				MySetCursorPos(g_selection.left, g_selection.top);
				InvalidateRect(hWnd, NULL, TRUE);
			}
			break;
		case VK_INSERT: // Insert => Store selection
			if (isSelectionValid(g_selection))
			{
				g_storedSelection = g_selection;
				storeDWORDSettingInRegistry(storedSelectionLeft, g_selection.left);
				storeDWORDSettingInRegistry(storedSelectionTop, g_selection.top);
				storeDWORDSettingInRegistry(storedSelectionRight, g_selection.right);
				storeDWORDSettingInRegistry(storedSelectionBottom, g_selection.bottom);
				InvalidateRect(hWnd, NULL, TRUE);
			}
			break;
		case VK_DELETE: // Delete => Clear stored and used selection
		{
			g_storedSelection = { UNINITIALIZEDLONG,UNINITIALIZEDLONG,UNINITIALIZEDLONG,UNINITIALIZEDLONG };
			storeDWORDSettingInRegistry(storedSelectionLeft, UNINITIALIZEDLONG);
			storeDWORDSettingInRegistry(storedSelectionTop, UNINITIALIZEDLONG);
			storeDWORDSettingInRegistry(storedSelectionRight, UNINITIALIZEDLONG);
			storeDWORDSettingInRegistry(storedSelectionBottom, UNINITIALIZEDLONG);

			POINT mouse;
			GetCursorPos(&mouse);
			g_appState = stateFirstPoint;
			g_selection.left = limitXtoBitmap(mouse.x - g_appWindowPos.x);
			g_selection.top = limitYtoBitmap(mouse.y - g_appWindowPos.y);
			g_selection.right = UNINITIALIZEDLONG;
			g_selection.bottom = UNINITIALIZEDLONG;
			InvalidateRect(hWnd, NULL, TRUE);
			break;
		}
		case VK_HOME: // Home => Restore selection
			if (isSelectionValid(g_storedSelection))
			{
				g_appState = statePointB;
				g_selection.left = limitXtoBitmap(g_storedSelection.left);
				g_selection.right = limitXtoBitmap(g_storedSelection.right);
				g_selection.top = limitYtoBitmap(g_storedSelection.top);
				g_selection.bottom = limitYtoBitmap(g_storedSelection.bottom);
				MySetCursorPos(g_selection.right, g_selection.bottom);
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
			resizeSelection(hWnd, 1);
			break;
		case '-': // Decrease selection
			resizeSelection(hWnd, -1);
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
		if (!g_ignoreNextClick) SendMessage(hWnd, WM_NEXTSTATE, wParam, lParam); else g_ignoreNextClick = FALSE;
		break;
	case WM_RBUTTONUP: // Right mouse button => Context menu
	{
		POINT pt;
		GetCursorPos(&pt);
		if (g_appState != stateTrayIcon) ShowCursor(true);
		HMENU hMenu = CreatePopupMenu();
		AppendMenu(hMenu, MF_STRING, IDM_CANCELCAPTURE, LoadStringAsWstr(g_hInst, IDS_CANCELCAPTURE).c_str());
		AppendMenu(hMenu, MF_STRING, IDM_EXIT, LoadStringAsWstr(g_hInst, IDS_EXIT).c_str());
		TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hWnd, NULL);
		DestroyMenu(hMenu);
		if (g_appState != stateTrayIcon) ShowCursor(false);
		break;
	}
	case WM_MOUSEMOVE: // Mouse moved
		OnMouseMove(hWnd, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam), (DWORD)wParam);
		break;
	case WM_TIMER: // Timer...
		switch (wParam)
		{
		case IDT_TIMER1000MS: // 1s timer
			{
				// Check window size
				RECT rectWindow = { 0 };
				GetWindowRect(hWnd, &rectWindow);
				// Reapply fullscreen, if window position/size is not correct
				if ((rectWindow.left != GetSystemMetrics(SM_XVIRTUALSCREEN)) ||
					(rectWindow.top != GetSystemMetrics(SM_YVIRTUALSCREEN)) ||
					(abs(rectWindow.left - rectWindow.left) != GetSystemMetrics(SM_CXVIRTUALSCREEN)) || 
					(abs(rectWindow.bottom - rectWindow.top) != GetSystemMetrics(SM_CYVIRTUALSCREEN))) 
				{
					enterFullScreen(hWnd);
				}

				InvalidateRect(hWnd, NULL, TRUE);
				break;
			}
		case IDT_TIMERSCREENSHOTDELAYED: // Onetime 5s timer
			KillTimer(hWnd, IDT_TIMERSCREENSHOTDELAYED); // Only one time
			SendMessage(hWnd, WM_STARTED, 0, 0);
			break;
		}
	case WM_COMMAND: // Menu selection
		switch (LOWORD(wParam))
		{
			case IDM_CAPTURE:
				if (WaitForSingleObject(g_hSemaphoreModalBlocked, INFINITE) != WAIT_FAILED)
				{
					SetTimer(hWnd, IDT_TIMERSCREENSHOTDELAYED, g_screenshotDelay*1000, (TIMERPROC)NULL);
					ReleaseSemaphore(g_hSemaphoreModalBlocked, 1, NULL);
				}
				break;
			case IDM_EXIT:
				PostQuitMessage(0);
				break;
			case IDM_ABOUT:
				if (WaitForSingleObject(g_hSemaphoreModalBlocked, INFINITE) != WAIT_FAILED)
				{
					showProgramInformation(hWnd);
					ReleaseSemaphore(g_hSemaphoreModalBlocked, 1, NULL);
				}
				break;
			case IDM_OPENFOLDER:
				ShellExecute(hWnd, L"open", g_screenshotPath, NULL, NULL, SW_SHOWNORMAL);
				break;
			case IDM_OPENLAST:
			{
				if (!g_sLastScreenshotFile.empty() && PathFileExists(g_sLastScreenshotFile.c_str()) ) { // Open file if file exists
					ShellExecute(hWnd, L"open", g_sLastScreenshotFile.c_str(), NULL, NULL, SW_SHOWNORMAL);
					break;
				} else g_sLastScreenshotFile = L"";
				// Otherwise open folder
				ShellExecute(hWnd, L"open", g_screenshotPath, NULL, NULL, SW_SHOWNORMAL);
				break;
			}
			case IDM_EDITLAST:
			{
				// It will be deprecated on 05 / 01 / 2025
				// https://learn.microsoft.com/en-us/windows/apps/develop/launch/launch-screen-snipping
				if (!g_sLastScreenshotFile.empty() && PathFileExists(g_sLastScreenshotFile.c_str())) { // Edit file if file exists
					WCHAR outputUrl[MAX_PATH];
					DWORD dwSize = ARRAYSIZE(outputUrl);

					UrlEscape(g_sLastScreenshotFile.c_str(), nullptr, &dwSize, URL_ESCAPE_PERCENT | URL_ESCAPE_ASCII_URI_COMPONENT);
					UrlEscape(g_sLastScreenshotFile.c_str(), outputUrl, &dwSize, URL_ESCAPE_PERCENT | URL_ESCAPE_ASCII_URI_COMPONENT);

					std::wstring sURI = L"ms-screensketch:edit?&filePath=";
					sURI.append(outputUrl);
					ShellExecute(hWnd, L"open", sURI.c_str(), NULL, NULL, SW_SHOWNORMAL);

					break;
				}
				break;
			}
			case IDM_SETFOLDER:
				if (WaitForSingleObject(g_hSemaphoreModalBlocked, INFINITE) != WAIT_FAILED)
				{
					changeScreenshotPathAndStorePathToRegistry();
					ReleaseSemaphore(g_hSemaphoreModalBlocked, 1, NULL);
				}
				break;
			case IDM_SAVETOCLIPBOARD: // Toggle save to clipboard
				if (g_bSaveToClipboardGPO) break;
				g_saveToClipboard = !g_saveToClipboard;
				storeDWORDSettingInRegistry(saveToClipboard, g_saveToClipboard);
				checkScreenshotTargets(hWnd);
				InvalidateRect(hWnd, NULL, TRUE);
				break;
			case IDM_SAVETOFILE: // Toggle save to file
				if (g_bSaveToFileGPO) break;
				g_saveToFile = !g_saveToFile;
				storeDWORDSettingInRegistry(saveToFile, g_saveToFile);
				checkScreenshotTargets(hWnd);
				InvalidateRect(hWnd, NULL, TRUE);
				break;
			case IDM_ALTERNATIVECOLORS: // Toggle colors
				g_useAlternativeColors = !g_useAlternativeColors;
				storeDWORDSettingInRegistry(useAlternativeColors, g_useAlternativeColors);
				InvalidateRect(hWnd, NULL, TRUE);
				break;
			case IDM_DISPLAYINFORMATION: // Toggle display information
				if (g_bDisplayInternalInformationGPO) break;
				g_displayInternalInformation = !g_displayInternalInformation;
				storeDWORDSettingInRegistry(displayInternalInformation, g_displayInternalInformation);
				InvalidateRect(hWnd, NULL, TRUE);
				break;
			case IDM_CANCELCAPTURE: // Cancel screenshot
				SendMessage(hWnd, WM_GOTOTRAY, 0, 0);
				break;
			case IDM_AUTORUN: // Toggle run key value
			{
				BOOL bRunEnabled = isRunKeyEnabledFromRegistry();
				if (!g_bRunKeyReadOnly) setRunKeyRegistryValue(!bRunEnabled,HKEY_CURRENT_USER);
				break;
			}
		}
		break;
	case WM_DISPLAYCHANGE:
		// Goto tray icon, when display changed, to prevent problems when connecting/disconnecting monitors
		if (g_appState != stateTrayIcon) SendMessage(hWnd, WM_GOTOTRAY, 0, 0);
		break;
	default:
		if ((WM_TASKBARCREATED != 0) && (message == WM_TASKBARCREATED)) // Recreate tray icon if explorer was restarted
		{
			Shell_NotifyIcon(NIM_ADD, &g_nid);
			break;
		}
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}