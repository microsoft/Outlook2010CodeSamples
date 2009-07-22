// ImportProcs.h : header file for loading imports from DLLs
//

#pragma once

extern HMODULE	hModMSMAPI;
extern HMODULE	hModMAPI;

HMODULE LoadFromSystemDir(LPTSTR szDLLName);

HMODULE MyLoadLibrary(LPCTSTR lpLibFileName);

HKEY GetMailKey(LPTSTR szClient);
void GetMapiMsiIds(LPTSTR szClient, LPTSTR* lpszComponentID, LPTSTR* lpszAppLCID, LPTSTR* lpszOfficeLCID);
void GetMAPIPath(LPTSTR szClient, LPTSTR szMAPIPath, ULONG cchMAPIPath);
void AutoLoadMAPI();
void UnloadMAPI();
void LoadMAPIFuncs(HMODULE hMod);