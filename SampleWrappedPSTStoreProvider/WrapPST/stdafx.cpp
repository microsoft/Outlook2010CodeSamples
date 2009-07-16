#define INITGUID
#define USES_IID_IExchExt
#define USES_IID_IExchExtCommands
#define USES_IID_IExchExtUserEvents
#define USEU_IID_IExchExtAdvancedCriteria
#define USES_IID_IExchExtSessionEvents
#define USES_IID_IExchExtMessageEvents
#define USES_IID_IExchExtAttachedFileEvents
#define USES_IID_IExchExtPropertySheets

#include <objbase.h>
#include <initguid.h>
#include <mapiguid.h>

#include "stdafx.h"

#define LOGFILE "c:\\temp\\WrapPST.txt"


// macro to get allocated string lengths
#define CCH(string) (sizeof((string))/sizeof(TCHAR))

FILE* OpenFile(LPCTSTR szFileName,BOOL bNewFile)
{
	static TCHAR szErr[256]; // buffer for catastrophic failures
	FILE* fOut = NULL;
	LPCTSTR szParams = _T("a+");
	errno_t myerr = 0;
	if (bNewFile) szParams = _T("w");
	myerr = _tfopen_s(&fOut,szFileName,szParams);
	if (fOut)
	{
		return fOut;
	}
	else
	{
		// File IO failed
		DWORD dwErr = GetLastError();
		HRESULT hRes = HRESULT_FROM_WIN32(dwErr);
		LPTSTR szSysErr = NULL;
		FormatMessage(
			FORMAT_MESSAGE_ALLOCATE_BUFFER|FORMAT_MESSAGE_FROM_SYSTEM,
			0,
			dwErr,
			0,
			(LPTSTR)&szSysErr,
			0,
			0);

		hRes = StringCchPrintf(szErr,CCH(szErr),
			_T("_tfopen failed, hRes = 0x%08X, dwErr = 0x%08X = \"%s\"\n"),
			HRESULT_FROM_WIN32(dwErr),
			dwErr,
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
//		OutputDebugString(szThreadTime);
	}

//	OutputDebugString(szMsg);

	// If there's a file - send the output there
	if (fFile)
	{
		if (bPrintThreadTime) _fputts(szThreadTime,fFile);
		_fputts(szMsg,fFile);
	}
}

FILE* lpLogFile = NULL;

void InitLogging()
{
	lpLogFile = OpenFile(LOGFILE,false);
}

void DeInitLogging()
{
	if (lpLogFile) CloseFile(lpLogFile);
	lpLogFile = NULL;
}


void Log(BOOL bPrintThreadTime, LPCTSTR szMsg,...)
{
	if (!lpLogFile) InitLogging();
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