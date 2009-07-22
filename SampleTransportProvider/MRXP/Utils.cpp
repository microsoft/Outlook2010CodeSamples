#include "stdafx.h"

// HrGetRegMultiSZValueA
// Get a REG_MULTI_SZ registry value - allocating memory using new to hold it.
void HrGetRegMultiSZValueA(
   IN HKEY hKey, // the key.
   IN LPCSTR lpszValue, // value name in key.
   OUT LPVOID* lppData) // where to put the data.
{
   *lppData = NULL;
   DWORD dwKeyType = NULL;
   DWORD cb = NULL;
   LONG lRet = 0;

   // Get its size
   lRet = RegQueryValueExA(
      hKey,
      lpszValue,
      NULL,
      &dwKeyType,
      NULL,
      &cb);

   if (ERROR_SUCCESS == lRet && cb && REG_MULTI_SZ == dwKeyType)
   {
      *lppData = new BYTE[cb];

      if (*lppData)
      {
         // Get the current value
         lRet = RegQueryValueExA(
            hKey,
            lpszValue,
            NULL,
            &dwKeyType,
            (unsigned char*)*lppData,
            &cb);

         if (ERROR_SUCCESS != lRet)
         {
            delete[] *lppData;
            *lppData = NULL;
         }
      }
   }
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

HRESULT CopySBinary(LPSBinary psbDest, const LPSBinary psbSrc, LPVOID lpParent,
					LPALLOCATEBUFFER MyAllocateBuffer, LPALLOCATEMORE MyAllocateMore)
{
	HRESULT	 hRes = S_OK;

	if (!psbDest || !psbSrc || (lpParent && !MyAllocateMore) || (!lpParent && !MyAllocateBuffer))
		return MAPI_E_INVALID_PARAMETER;

	psbDest->cb = psbSrc->cb;

	if (psbSrc->cb)
	{
		if (lpParent)
			hRes = MyAllocateMore(psbSrc->cb, lpParent, (LPVOID *)&psbDest->lpb);
		else
			hRes = MyAllocateBuffer(psbSrc->cb, (LPVOID *)&psbDest->lpb);
		if (S_OK == hRes)
			CopyMemory(psbDest->lpb,psbSrc->lpb,psbSrc->cb);
	}

	return hRes;
}