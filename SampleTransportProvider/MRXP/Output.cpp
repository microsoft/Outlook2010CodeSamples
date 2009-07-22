#include "stdafx.h"

BOOL g_bDoFileLogging = false;
BOOL g_bDoDebugLogging = false;
BOOL g_LoggingInitialized = false;

FILE* OpenFile(LPCTSTR szFileName,BOOL bNewFile)
{
	static TCHAR szErr[256]; // buffer for catastrophic failures
	FILE* fOut = NULL;
	LPCTSTR szParams = _T("a+");
	if (bNewFile) szParams = _T("w");
	errno_t iErr = 0;
	iErr = _tfopen_s(&fOut,szFileName,szParams);
	if (!iErr && fOut)
	{
		return fOut;
	}
	else
	{
		// File IO failed
		HRESULT hRes = S_OK;
		LPTSTR szSysErr = NULL;
		FormatMessage(
			FORMAT_MESSAGE_ALLOCATE_BUFFER|FORMAT_MESSAGE_FROM_SYSTEM,
			0,
			iErr,
			0,
			(LPTSTR)&szSysErr,
			0,
			0);

		hRes = StringCchPrintf(szErr,CCH(szErr),
			_T("_tfopen failed, hRes = 0x%08X, iErr = 0x%08X = \"%s\"\n"),
			HRESULT_FROM_WIN32(iErr),
			iErr,
			szSysErr);

		OutputDebugString(szErr);
		LocalFree(szSysErr);
		return NULL;
	}
}

void CloseFile(FILE* fFile)
{
	if (fFile) fclose(fFile);
}

void _Output(FILE* fFile, BOOL bPrintThreadTime, LPCTSTR szMsg)
{
	HRESULT hRes = S_OK;

	if (!szMsg) return; // nothing to print

	// Compute current time and thread for a time stamp
	TCHAR		szThreadTime[MAX_PATH];

	if (bPrintThreadTime)
	{
		SYSTEMTIME	stLocalTime;
		FILETIME	ftLocalTime;
		GetSystemTime(&stLocalTime);
		GetSystemTimeAsFileTime(&ftLocalTime);

		hRes = StringCchPrintf(szThreadTime,
			CCH(szThreadTime),
			_T("0x%04x %02d:%02d:%02d.%03d%s  %02d-%02d-%4d "),
			GetCurrentThreadId(),
			(stLocalTime.wHour <= 12)?stLocalTime.wHour:stLocalTime.wHour-12,
			stLocalTime.wMinute,
			stLocalTime.wSecond,
			stLocalTime.wMilliseconds,
			(stLocalTime.wHour <= 12)?_T("AM"):_T("PM"),
			stLocalTime.wMonth,
			stLocalTime.wDay,
			stLocalTime.wYear);
		if (g_bDoDebugLogging)
			OutputDebugString(szThreadTime);
	}

	if (g_bDoDebugLogging)
		OutputDebugString(szMsg);

	// If there is a file send the output there
	if (g_bDoFileLogging && fFile)
	{
		if (bPrintThreadTime) _fputts(szThreadTime,fFile);
		_fputts(szMsg,fFile);
	}
}

FILE* lpLogFile = NULL;

void InitLogging()
{
	// Check reg keys and enable logging as indicated
	// For example, Break out logfile vs debug output

	g_bDoFileLogging = false;
	g_bDoDebugLogging = false;

	ULONG ulRet = 0;
	HKEY hOptions = NULL;

	ulRet = RegOpenKeyEx(HKEY_CURRENT_USER, RKEY_ROOT, 0, KEY_READ, &hOptions);
	if (ERROR_SUCCESS == ulRet)
	{
		DWORD dwType = 0;
		DWORD dwValue = 0;
		DWORD dwcbValue = sizeof(DWORD);

		ulRet = RegQueryValueEx(hOptions, REG_DEBUG_OPTIONS, NULL, &dwType,
			(LPBYTE)&dwValue, &dwcbValue);
		if (ERROR_SUCCESS == ulRet && dwType == REG_DWORD)
		{
			if (dwValue & DO_FILE_LOGGING)
				g_bDoFileLogging = true;
			if (dwValue & DO_DEBUG_LOGGING)
				g_bDoDebugLogging = true;
		}

		RegCloseKey(hOptions);
	}

	if (g_bDoFileLogging)
		lpLogFile = OpenFile(LOGFILE,false);

	g_LoggingInitialized = true;
}

void DeInitLogging()
{
	if (lpLogFile) CloseFile(lpLogFile);
	lpLogFile = NULL;

	g_LoggingInitialized = false;
}


void Log(BOOL bPrintThreadTime, LPCTSTR szMsg,...)
{
	if (!g_LoggingInitialized)
		InitLogging();

	if (!g_bDoFileLogging && !g_bDoDebugLogging)
		return;

	HRESULT hRes = S_OK;

	va_list argList = NULL;
	va_start(argList, szMsg);

	if (argList)
	{
		TCHAR szDebugString[4096];
		hRes = StringCchVPrintf(szDebugString, CCH(szDebugString), szMsg, argList);
		_Output(lpLogFile, bPrintThreadTime, szDebugString);
	}
	else
		_Output(lpLogFile, bPrintThreadTime, szMsg);
	va_end(argList);
}