#include "stdafx.h"
#include <msi.h>
#include "ImportProcs.h"

LPTSTR GenerateProviderPath()
{
	TCHAR szFilePath[MAX_PATH] = {0};
	DWORD dwDir = NULL;
	HMODULE hMod = NULL;
	hMod = GetModuleHandle("wrppst32.dll");
	if (hMod)
	{
		dwDir = GetModuleFileName(hMod,szFilePath,_countof(szFilePath));
		TCHAR* szExt = _tcsstr(szFilePath, "32.dll");
		if (szExt)
		{
			StringCchPrintf(szExt, 6, ".dll");
			size_t cchStr = 0;
			StringCchLength(szFilePath, MAX_PATH, &cchStr);
			LPTSTR szPath = new TCHAR[cchStr+1];
			if (szPath)
			{
				StringCchCopy(szPath, cchStr+1, szFilePath);
				return szPath;
			}
		}
	}

	return NULL;
}

STDMETHODIMP MergeWithMAPISVC()
{
	Log(true,"MergeWithMAPISVC adding wrapped PST\n");
	Log(true,"DLL install path is %s\n", aWrapPSTServicesIni[2].lpszValue);
	LPTSTR szPath = GenerateProviderPath();
	Log(true,"Better DLL install path is %s\n", szPath);
	aWrapPSTServicesIni[2].lpszValue = szPath;
	aWrapPSTServicesIni[9].lpszValue = szPath;
	HRESULT hRes = HrSetProfileParameters(aWrapPSTServicesIni);
	delete[] szPath;
	return hRes;
}

STDMETHODIMP RemoveFromMAPISVC()
{
	Log(true,"RemoveFromMAPISVC removing wrapped PST\n");
	return HrSetProfileParameters(aREMOVE_WrapPSTServicesIni);
}

///////////////////////////////////////////////////////////////////////////////
// Function name   : GetMAPISVCPath
// Description     : This will get the correct path to the MAPISVC.INF file.
// Return type     : void
// Argument        : LPSTR szMAPIDir - Buffer to hold the path to the MAPISVC file.
//                   ULONG cchMAPIDir - size of the buffer
void GetMAPISVCPath(LPTSTR szMAPIDir, ULONG cchMAPIDir)
{
	Log(true,_T("Enter GetMAPISVCPath\n"));

	GetMAPIPath(_T("Microsoft Outlook"),szMAPIDir,cchMAPIDir);

	// We got the path to msmapi32.dll - need to strip it
	if (szMAPIDir[0] != _T('\0'))
	{
		LPTSTR lpszSlash = NULL;
		LPTSTR lpszCur = szMAPIDir;

		for (lpszSlash = lpszCur; *lpszCur; lpszCur = lpszCur++)
		{
			if (*lpszCur == _T('\\')) lpszSlash = lpszCur;
		}
		*lpszSlash = _T('\0');
	}

	if (szMAPIDir[0] == _T('\0'))
	{
		Log(true,_T("FGetComponentPath failed, loading system directory\n"));
		// Fall back on System32
		UINT uiLen = 0;
		uiLen = GetSystemDirectory(szMAPIDir, cchMAPIDir);
	}

	if (szMAPIDir[0] != _T('\0'))
	{
		Log(true,_T("Using directory: %s\n"),szMAPIDir);
		StringCchPrintf(
			szMAPIDir,
			cchMAPIDir,
			_T("%s\\%s"),
			szMAPIDir,
			_T("MAPISVC.INF"));
	}
}

// $--HrSetProfileParameters----------------------------------------------
// Add values to MAPISVC.INF
// -----------------------------------------------------------------------------
STDMETHODIMP HrSetProfileParameters(SERVICESINIREC *lpServicesIni)
{
	HRESULT	hRes						= S_OK;
	TCHAR	szSystemDir[MAX_PATH+1]		= {0};
	TCHAR	szServicesIni[MAX_PATH+12]	= {0}; // 12 = space for "MAPISVC.INF"
	UINT	n							= 0;
	TCHAR	szPropNum[10]				= {0};

	Log(true,_T("HrSetProfileParameters()\n"));

	if (!lpServicesIni) return MAPI_E_INVALID_PARAMETER;

	GetMAPISVCPath(szServicesIni,CCH(szServicesIni));

	if (!szServicesIni[0])
	{
		UINT uiLen = 0;
		uiLen = GetSystemDirectory(szSystemDir, CCH(szSystemDir));
		if (!uiLen)
			return MAPI_E_CALL_FAILED;

		Log(true,_T("Writing to this directory: \"%s\"\n"),szSystemDir);

		hRes = StringCchPrintf(
			szServicesIni,
			CCH(szServicesIni),
			_T("%s\\%s"),
			szSystemDir,
			_T("MAPISVC.INF"));
	}

	Log(true,_T("Writing to this file: \"%s\"\n"),szServicesIni);

	//
	// Loop through and add items to MAPISVC.INF
	//

	n = 0;

	while(lpServicesIni[n].lpszSection != NULL)
	{
		LPTSTR lpszProp = lpServicesIni[n].lpszKey;
		LPTSTR lpszValue = lpServicesIni[n].lpszValue;

		// Switch the property if necessary

		if ((lpszProp == NULL) && (lpServicesIni[n].ulKey != 0))
		{

			hRes = StringCchPrintf(
				szPropNum,
				CCH(szPropNum),
				_T("%lx"),
				lpServicesIni[n].ulKey);

			if (SUCCEEDED(hRes))
				lpszProp = szPropNum;
		}

		//
		// Write the item to MAPISVC.INF
		//

		WritePrivateProfileString(
			lpServicesIni[n].lpszSection,
			lpszProp,
			lpszValue,
			szServicesIni);
		n++;
	}

	// Flush the information - ignore the return code
	WritePrivateProfileString(NULL, NULL, NULL, szServicesIni);

	return hRes;
}

// result allocated with new, clean up with delete[]
void GetOutlookVersionString(LPTSTR* lppszPath, LPTSTR* lppszVer, LPTSTR* lppszLang)
{
	LPTSTR lpszTempPath = NULL;
	LPTSTR lpszTempVer = NULL;
	LPTSTR lpszTempLang = NULL;

	if (lpszTempPath) *lpszTempPath = NULL;
	if (lpszTempVer) *lpszTempPath = NULL;
	if (lpszTempLang) *lpszTempLang = NULL;

	TCHAR pszaOutlookQualifiedComponents[][MAX_PATH] = {
		_T("{BC174BAD-2F53-4855-A1D5-0D575C19B1EA}"), // O11_CATEGORY_GUID_CORE_OFFICE (retail)
		_T("{BC174BAD-2F53-4855-A1D5-1D575C19B1EA}"), // O11_CATEGORY_GUID_CORE_OFFICE (debug)
		_T("{24AAE126-0911-478F-A019-07B875EB9996}"), // O12_CATEGORY_GUID_CORE_OFFICE (retail)
		_T("{1E77DE88-BCAB-4C37-B9E5-073AF52DFD7A}")  // O14_CATEGORY_GUID_CORE_OFFICE (retail)
	};
	int nOutlookQualifiedComponents = sizeof(pszaOutlookQualifiedComponents)/sizeof(TCHAR*);
	int i = 0;
	DWORD dwValueBuf = 0;
	UINT ret = 0;

	for (i = 0; i < nOutlookQualifiedComponents; i++)
	{
		ret = MsiProvideQualifiedComponent(
			pszaOutlookQualifiedComponents[i],
			_T("outlook.exe"),
			(DWORD) INSTALLMODE_DEFAULT,
			NULL,
			&dwValueBuf);
		if (ERROR_SUCCESS == ret) break;
	}

	if (ERROR_SUCCESS == ret)
	{
		dwValueBuf += 1;
		lpszTempPath = new TCHAR[dwValueBuf];

		if (lpszTempPath != NULL)
		{
			ret = MsiProvideQualifiedComponent(
				pszaOutlookQualifiedComponents[i],
				_T("outlook.exe"),
				(DWORD) INSTALLMODE_EXISTING,
				lpszTempPath,
				&dwValueBuf);

			if (ERROR_SUCCESS == ret)
			{
				lpszTempVer = new TCHAR[MAX_PATH];
				lpszTempLang = new TCHAR[MAX_PATH];
				dwValueBuf = MAX_PATH;
				if (lpszTempVer && lpszTempLang)
				{
					ret = MsiGetFileVersion(lpszTempPath,
						lpszTempVer,
						&dwValueBuf,
						lpszTempLang,
						&dwValueBuf);
					if (ERROR_SUCCESS == ret)
					{
						if (lppszVer)
						{
							*lppszVer = lpszTempVer;
							lpszTempVer = NULL;
						}
					}

					if (lppszPath)
					{
						*lppszPath = lpszTempPath;
						lpszTempPath = NULL;
					}
					if (lppszLang)
					{
						*lppszLang = lpszTempLang;
						lpszTempLang = NULL;
					}
				}
			}
		}
	}

	delete[] lpszTempVer;
	delete[] lpszTempLang;
	delete[] lpszTempPath;
} // GetOutlookVersionString

LPTSTR GetPSTPath()
{
	LPTSTR lpszPath = NULL;
	GetOutlookVersionString(&lpszPath, NULL, NULL);

	// Got a path - strip outlook.exe and append mspst32.dll
	if (lpszPath)
	{
		// We got the path to Outlook.exe - need to strip it
		if (lpszPath[0] != _T('\0'))
		{
			LPTSTR lpszSlash = NULL;
			LPTSTR lpszCur = lpszPath;

			for (lpszSlash = lpszCur; *lpszCur; lpszCur = lpszCur++)
			{
				if (*lpszCur == _T('\\')) lpszSlash = lpszCur;
			}
			if (lpszSlash)
			{
				// A real implementation wouldn't do this, but it works for sample code
				// this only works because outlook.exe and mspst32.dll are both 7.3
				memcpy(lpszSlash+1,_T("mspst32.dll"),11*sizeof(TCHAR));
				return lpszPath;
			}
		}
		// we only get here if we didn't return a path - clean up
		delete[] lpszPath;
	}
	return NULL;
} // GetPSTPath