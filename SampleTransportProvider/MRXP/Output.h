#pragma once

#define LOGFILE _T("C:\\MRXPLOG.txt")

// macro to get allocated string lengths
#define CCH(string) (sizeof((string))/sizeof(TCHAR))
#define CCH_A(string) (sizeof((string))/sizeof(CHAR))

#define RKEY_ROOT _T("SOFTWARE\\Microsoft\\MRXP")
#define REG_DEBUG_OPTIONS _T("DebugFlags")

#define DO_FILE_LOGGING		0x1
#define DO_DEBUG_LOGGING	0x2

void DeInitLogging();
void Log(BOOL bPrintThreadTime, LPCTSTR szMsg,...);