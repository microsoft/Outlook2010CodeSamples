#include "stdafx.h"
#include "ImportProcs.h"
#include <MAPIX.H>
#include <MAPIUTIL.H>
#include <MAPIFORM.H>
#include <imessage.h>
#include <mapival.h>
#include <mapiwz.h>
#include <tnef.h>
#include <strsafe.h>

// macro to get allocated string lengths
#define CCH(string) (sizeof((string))/sizeof(TCHAR))

// Note: when the user loads MAPI manually, hModMSMAPI and hModMAPI will be the same
HMODULE	hModMSMAPI = NULL; // Address of Outlook's MAPI
HMODULE	hModMAPI = NULL; // Address of MAPI32 in System32
HMODULE	hModMAPIStub = NULL; // Address of the MAPI stub library

// For GetComponentPath
typedef BOOL (STDAPICALLTYPE FGETCOMPONENTPATH)
	(LPSTR szComponent,
	LPSTR szQualifier,
	LPSTR szDllPath,
	DWORD cchBufferSize,
	BOOL fInstall);
typedef FGETCOMPONENTPATH FAR * LPFGETCOMPONENTPATH;

// For all the MAPI funcs we use
typedef SCODE (STDMETHODCALLTYPE HRGETONEPROP)(
					LPMAPIPROP lpMapiProp,
					ULONG ulPropTag,
					LPSPropValue FAR *lppProp);
typedef HRGETONEPROP *LPHRGETONEPROP;

typedef void (STDMETHODCALLTYPE FREEPROWS)(
					LPSRowSet lpRows);
typedef FREEPROWS	*LPFREEPROWS;

typedef	LPSPropValue (STDMETHODCALLTYPE PPROPFINDPROP)(
					LPSPropValue lpPropArray, ULONG cValues,
					ULONG ulPropTag);
typedef PPROPFINDPROP *LPPPROPFINDPROP;

typedef SCODE (STDMETHODCALLTYPE SCDUPPROPSET)(
					int cValues, LPSPropValue lpPropArray,
					LPALLOCATEBUFFER lpAllocateBuffer, LPSPropValue FAR *lppPropArray);
typedef SCDUPPROPSET *LPSCDUPPROPSET;

typedef SCODE (STDMETHODCALLTYPE SCCOUNTPROPS)(
					int cValues, LPSPropValue lpPropArray, ULONG FAR *lpcb);
typedef SCCOUNTPROPS *LPSCCOUNTPROPS;

typedef SCODE (STDMETHODCALLTYPE SCCOPYPROPS)(
					int cValues, LPSPropValue lpPropArray, LPVOID lpvDst,
					ULONG FAR *lpcb);
typedef SCCOPYPROPS *LPSCCOPYPROPS;

typedef SCODE (STDMETHODCALLTYPE OPENIMSGONISTG)(
					LPMSGSESS		lpMsgSess,
					LPALLOCATEBUFFER lpAllocateBuffer,
					LPALLOCATEMORE 	lpAllocateMore,
					LPFREEBUFFER	lpFreeBuffer,
					LPMALLOC		lpMalloc,
					LPVOID			lpMapiSup,
					LPSTORAGE 		lpStg,
					MSGCALLRELEASE FAR *lpfMsgCallRelease,
					ULONG			ulCallerData,
					ULONG			ulFlags,
					LPMESSAGE		FAR *lppMsg );
typedef OPENIMSGONISTG *LPOPENIMSGONISTG;

typedef	LPMALLOC (STDMETHODCALLTYPE MAPIGETDEFAULTMALLOC)(
					void);
typedef MAPIGETDEFAULTMALLOC *LPMAPIGETDEFAULTMALLOC;

typedef	void (STDMETHODCALLTYPE CLOSEIMSGSESSION)(
					LPMSGSESS		lpMsgSess);
typedef CLOSEIMSGSESSION *LPCLOSEIMSGSESSION;

typedef	SCODE (STDMETHODCALLTYPE OPENIMSGSESSION)(
					LPMALLOC		lpMalloc,
					ULONG			ulFlags,
					LPMSGSESS FAR	*lppMsgSess);
typedef OPENIMSGSESSION *LPOPENIMSGSESSION;

typedef SCODE (STDMETHODCALLTYPE HRQUERYALLROWS)(
					LPMAPITABLE lpTable,
					LPSPropTagArray lpPropTags,
					LPSRestriction lpRestriction,
					LPSSortOrderSet lpSortOrderSet,
					LONG crowsMax,
					LPSRowSet FAR *lppRows);
typedef HRQUERYALLROWS *LPHRQUERYALLROWS;

typedef SCODE (STDMETHODCALLTYPE MAPIOPENFORMMGR)(
					LPMAPISESSION pSession, LPMAPIFORMMGR FAR * ppmgr);
typedef MAPIOPENFORMMGR *LPMAPIOPENFORMMGR;

typedef HRESULT (STDMETHODCALLTYPE RTFSYNC) (
	LPMESSAGE lpMessage, ULONG ulFlags, BOOL FAR * lpfMessageUpdated);
typedef RTFSYNC *LPRTFSYNC;

typedef SCODE (STDMETHODCALLTYPE HRSETONEPROP)(
					LPMAPIPROP lpMapiProp,
					LPSPropValue lpProp);
typedef HRSETONEPROP *LPHRSETONEPROP;

typedef void (STDMETHODCALLTYPE FREEPADRLIST)(
					LPADRLIST lpAdrList);
typedef FREEPADRLIST *LPFREEPADRLIST;

typedef SCODE (STDMETHODCALLTYPE PROPCOPYMORE)(
					LPSPropValue		lpSPropValueDest,
					LPSPropValue		lpSPropValueSrc,
					ALLOCATEMORE *	lpfAllocMore,
					LPVOID			lpvObject );
typedef PROPCOPYMORE *LPPROPCOPYMORE;

typedef HRESULT (STDMETHODCALLTYPE WRAPCOMPRESSEDRTFSTREAM) (
	LPSTREAM lpCompressedRTFStream, ULONG ulFlags, LPSTREAM FAR * lpUncompressedRTFStream);
typedef WRAPCOMPRESSEDRTFSTREAM *LPWRAPCOMPRESSEDRTFSTREAM;

typedef SCODE (STDMETHODCALLTYPE HRVALIDATEIPMSUBTREE)(
					LPMDB lpMDB, ULONG ulFlags,
					ULONG FAR *lpcValues, LPSPropValue FAR *lppValues,
					LPMAPIERROR FAR *lpperr);
typedef HRVALIDATEIPMSUBTREE *LPHRVALIDATEIPMSUBTREE;

typedef HRESULT (STDAPICALLTYPE MAPIOPENLOCALFORMCONTAINER)(LPMAPIFORMCONTAINER FAR * ppfcnt);
typedef MAPIOPENLOCALFORMCONTAINER *LPMAPIOPENLOCALFORMCONTAINER;

typedef SCODE (STDAPICALLTYPE HRDISPATCHNOTIFICATIONS)(ULONG ulFlags);
typedef HRDISPATCHNOTIFICATIONS *LPHRDISPATCHNOTIFICATIONS;

typedef HRESULT (STDMETHODCALLTYPE WRAPSTOREENTRYID)(
	ULONG ulFlags,
	__in LPTSTR lpszDLLName,
	ULONG cbOrigEntry,
	LPENTRYID lpOrigEntry,
	ULONG *lpcbWrappedEntry,
	LPENTRYID *lppWrappedEntry);
typedef WRAPSTOREENTRYID *LPWRAPSTOREENTRYID;

typedef SCODE (STDMETHODCALLTYPE CREATEIPROP)(
	LPCIID					lpInterface,
	ALLOCATEBUFFER FAR *	lpAllocateBuffer,
	ALLOCATEMORE FAR *		lpAllocateMore,
	FREEBUFFER FAR *		lpFreeBuffer,
	LPVOID					lpvReserved,
	LPPROPDATA FAR *		lppPropData);
typedef CREATEIPROP *LPCREATEIPROP;

typedef SCODE (STDMETHODCALLTYPE CREATETABLE)(
	LPCIID					lpInterface,
	ALLOCATEBUFFER FAR *	lpAllocateBuffer,
	ALLOCATEMORE FAR *		lpAllocateMore,
	FREEBUFFER FAR *		lpFreeBuffer,
	LPVOID					lpvReserved,
	ULONG					ulTableType,
	ULONG					ulPropTagIndexColumn,
	LPSPropTagArray		lpSPropTagArrayColumns,
	LPTABLEDATA FAR *		lppTableData);
typedef CREATETABLE *LPCREATETABLE;

typedef HRESULT (STDMETHODCALLTYPE HRVALIDATEPARAMETERS)(
	METHODS eMethod,
	LPVOID FAR *ppFirstArg);
typedef HRVALIDATEPARAMETERS *LPHRVALIDATEPARAMETERS;

typedef HRESULT (STDMETHODCALLTYPE BUILDDISPLAYTABLE)(
	LPALLOCATEBUFFER	lpAllocateBuffer,
	LPALLOCATEMORE		lpAllocateMore,
	LPFREEBUFFER		lpFreeBuffer,
	LPMALLOC			lpMalloc,
	HINSTANCE			hInstance,
	UINT				cPages,
	LPDTPAGE			lpPage,
	ULONG				ulFlags,
	LPMAPITABLE *		lppTable,
	LPTABLEDATA	*		lppTblData);
typedef BUILDDISPLAYTABLE *LPBUILDDISPLAYTABLE;

typedef int (STDMETHODCALLTYPE MNLS_LSTRLENW)(
	LPCWSTR	lpString);
typedef MNLS_LSTRLENW *LPMNLS_LSTRLENW;

// All of these get loaded from a MAPI DLL:
LPFGETCOMPONENTPATH			pfnFGetComponentPath = NULL;
LPLAUNCHWIZARDENTRY			pfnLaunchWizard = NULL;
LPMAPIALLOCATEBUFFER		pfnMAPIAllocateBuffer = NULL;
LPMAPIALLOCATEMORE			pfnMAPIAllocateMore = NULL;
LPMAPIFREEBUFFER			pfnMAPIFreeBuffer = NULL;
LPHRGETONEPROP				pfnHrGetOneProp = NULL;
LPMAPIINITIALIZE			pfnMAPIInitialize = NULL;
LPMAPIUNINITIALIZE			pfnMAPIUninitialize = NULL;
LPFREEPROWS					pfnLPFreeProws = NULL;
LPPPROPFINDPROP				pfnPpropFindProp = NULL;
LPSCDUPPROPSET				pfnScDupPropset = NULL;
LPSCCOUNTPROPS				pfnScCountProps = NULL;
LPSCCOPYPROPS				pfnScCopyProps = NULL;
LPOPENIMSGONISTG			pfnOpenIMsgOnIStg = NULL;
LPMAPIGETDEFAULTMALLOC		pfnMAPIGetDefaultMalloc = NULL;
LPOPENTNEFSTREAMEX			pfnOpenTnefStreamEx = NULL;
LPOPENSTREAMONFILE			pfnOpenStreamOnFile = NULL;
LPCLOSEIMSGSESSION			pfnCloseIMsgSession = NULL;
LPOPENIMSGSESSION			pfnOpenIMsgSession = NULL;
LPHRQUERYALLROWS			pfnHrQueryAllRows = NULL;
LPMAPIOPENFORMMGR			pfnMAPIOpenFormMgr = NULL;
LPRTFSYNC					pfnRTFSync = NULL;
LPHRSETONEPROP				pfnHrSetOneProp = NULL;
LPFREEPADRLIST				pfnFreePadrlist = NULL;
LPPROPCOPYMORE				pfnPropCopyMore = NULL;
LPWRAPCOMPRESSEDRTFSTREAM	pfnWrapCompressedRTFStream = NULL;
LPWRAPSTOREENTRYID			pfnWrapStoreEntryID = NULL;
LPCREATEIPROP				pfnCreateIProp = NULL;
LPCREATETABLE				pfnCreateTable = NULL;
LPHRVALIDATEPARAMETERS		pfnHrValidateParameters = NULL;
LPBUILDDISPLAYTABLE			pfnBuildDisplayTable = NULL;
LPMNLS_LSTRLENW				pfnMNLS_lstrlenW = NULL;
LPMAPILOGONEX				pfnMAPILogonEx = NULL;
LPMAPIADMINPROFILES			pfnMAPIAdminProfiles = NULL;
LPHRVALIDATEIPMSUBTREE		pfnHrValidateIPMSubtree = NULL;
LPMAPIOPENLOCALFORMCONTAINER pfnMAPIOpenLocalFormContainer = NULL;
LPHRDISPATCHNOTIFICATIONS	pfnHrDispatchNotifications = NULL;

// Exists to allow some logging
HMODULE MyLoadLibrary(LPCTSTR lpLibFileName)
{
	HMODULE hMod = NULL;
	Log(true,_T("MyLoadLibrary - loading \"%s\"\n"),lpLibFileName);
	hMod = LoadLibrary(lpLibFileName);
	if (hMod)
	{
		Log(true,_T("MyLoadLibrary - \"%s\" loaded at 0x%08X\n"),lpLibFileName,hMod);
	}
	else
	{
		Log(true,_T("MyLoadLibrary - \"%s\" failed to load\n"),lpLibFileName);
	}
	return hMod;
}

void LoadGetComponentPath()
{
	HRESULT hRes = S_OK;

	if (pfnFGetComponentPath) return;

	if (!hModMAPI) hModMAPI = LoadFromSystemDir(_T("mapi32.dll"));
	if (hModMAPI)
	{
		pfnFGetComponentPath = (LPFGETCOMPONENTPATH) GetProcAddress(
			hModMAPI,
			"FGetComponentPath");
		hRes = S_OK;
	}
	if (!pfnFGetComponentPath)
	{
		if (!hModMAPIStub) hModMAPIStub = LoadFromSystemDir(_T("mapistub.dll"));
		if (hModMAPIStub)
		{
			pfnFGetComponentPath = (LPFGETCOMPONENTPATH) GetProcAddress(
				hModMAPIStub,
				"FGetComponentPath");
		}
	}

	Log(true,_T("FGetComponentPath loaded at 0x%08X\n"),pfnFGetComponentPath);
} // LoadGetComponentPath

BOOL GetComponentPath(
					  LPTSTR szComponent,
					  LPTSTR szQualifier,
					  TCHAR* szDllPath,
					  DWORD cchDLLPath)
{
	BOOL bRet = false;
	LoadGetComponentPath();

	if (!pfnFGetComponentPath) return false;
#ifdef UNICODE
	int iRet = 0;
	CHAR szAsciiPath[MAX_PATH] = {0};
	char szAnsiComponent[MAX_PATH] = {0};
	char szAnsiQualifier[MAX_PATH] = {0};
	iRet = WideCharToMultiByte(CP_ACP,0,szComponent,(int) -1,szAnsiComponent,CCH(szAnsiComponent),NULL,NULL);
	iRet = WideCharToMultiByte(CP_ACP,0,szQualifier,(int) -1,szAnsiQualifier,CCH(szAnsiQualifier),NULL,NULL);

	bRet = pfnFGetComponentPath(
		szAnsiComponent,
		szAnsiQualifier,
		szAsciiPath,
		CCH(szAsciiPath),
		TRUE);

	iRet = MultiByteToWideChar(
		CP_ACP,
		0,
		szAsciiPath,
		CCH(szAsciiPath),
		szDllPath,
		cchDLLPath);
#else
	bRet = pfnFGetComponentPath(
		szComponent,
		szQualifier,
		szDllPath,
		cchDLLPath,
		TRUE);
#endif
	return bRet;
} // GetComponentPath

HMODULE LoadFromSystemDir(LPTSTR szDLLName)
{
	if (!szDLLName) return NULL;

	HRESULT	hRes = S_OK;
	HMODULE	hModRet = NULL;
	TCHAR	szDLLPath[MAX_PATH] = {0};
	UINT	uiRet = NULL;

	static TCHAR	szSystemDir[MAX_PATH] = {0};
	static BOOL		bSystemDirLoaded = false;

	Log(true,_T("LoadFromSystemDir - loading \"%s\"\n"),szDLLName);

	if (!bSystemDirLoaded)
	{
		uiRet = GetSystemDirectory(szSystemDir, MAX_PATH);
		bSystemDirLoaded = true;
	}

	hRes = StringCchPrintf(szDLLPath,CCH(szDLLPath),_T("%s\\%s"),szSystemDir,szDLLName);
	Log(true,_T("LoadFromSystemDir - loading from \"%s\"\n"),szDLLPath);
	hModRet = MyLoadLibrary(szDLLPath);

	return hModRet;
}

// $--GetRegistryValue---------------------------------------------------------
// Get a registry value - allocating memory using new to hold it.
// -----------------------------------------------------------------------------
LONG GetRegistryValue(
						   IN HKEY hKey, // the key.
						   IN LPCTSTR lpszValue, // value name in key.
						   OUT DWORD* lpType, // where to put type info.
						   OUT LPVOID* lppData) // where to put the data.
{
	LONG lRes = ERROR_SUCCESS;

	Log(true,_T("GetRegistryValue(%s)\n"),lpszValue);

	*lppData = NULL;
	DWORD cb = NULL;

	// Get its size
	lRes = RegQueryValueEx(
		hKey,
		lpszValue,
		NULL,
		lpType,
		NULL,
		&cb);

	// only handle types we know about - all others are bad
	if (ERROR_SUCCESS == lRes && cb &&
		(REG_SZ == *lpType || REG_DWORD == *lpType || REG_MULTI_SZ == *lpType))
	{
		*lppData = new BYTE[cb];

		if (*lppData)
		{
			// Get the current value
			lRes = RegQueryValueEx(
				hKey,
				lpszValue,
				NULL,
				lpType,
				(unsigned char*)*lppData,
				&cb);

			if (ERROR_SUCCESS != lRes)
			{
				delete[] *lppData;
				*lppData = NULL;
			}
		}
	}
	else lRes = ERROR_INVALID_DATA;

	return lRes;
}

// Opens the mail key for the specified MAPI client, such as 'Microsoft Outlook' or 'ExchangeMAPI'
// Pass NULL to open the mail key for the default MAPI client
HKEY GetMailKey(LPTSTR szClient)
{
	Log(true,_T("Enter GetMailKey(%s)\n"),szClient?szClient:_T("Default"));
	HRESULT hRes = S_OK;
	LONG lRet = S_OK;
	HKEY hMailKey = NULL;
	BOOL bClientIsDefault = false;

	// If szClient is NULL, we need to read the name of the default MAPI client
	if (!szClient)
	{
		HKEY hDefaultMailKey = NULL;
		lRet = RegOpenKeyEx(
			HKEY_LOCAL_MACHINE,
			_T("Software\\Clients\\Mail"),
			NULL,
			KEY_READ,
			&hDefaultMailKey);
		if (hDefaultMailKey)
		{
			DWORD dwKeyType = NULL;
			lRet = GetRegistryValue(
				hDefaultMailKey,
				_T(""), // get the default value
				&dwKeyType,
				(LPVOID*) &szClient);
			Log(true,_T("Default MAPI = %s\n"),szClient?szClient:_T("Default"));
			bClientIsDefault = true;
			RegCloseKey(hDefaultMailKey);
		}
	}

	if (szClient)
	{
		TCHAR szMailKey[256];
		hRes = StringCchPrintf(
			szMailKey,
			CCH(szMailKey),
			_T("Software\\Clients\\Mail\\%s"),
			szClient);

		if (SUCCEEDED(hRes))
		{
			lRet = RegOpenKeyEx(
				HKEY_LOCAL_MACHINE,
				szMailKey,
				NULL,
				KEY_READ,
				&hMailKey);
		}
	}
	if (bClientIsDefault) delete[] szClient;

	return hMailKey;
} // GetMailKey

// Gets MSI IDs for the specified MAPI client, such as 'Microsoft Outlook' or 'ExchangeMAPI'
// Pass NULL to get the IDs for the default MAPI client
// Allocates with new, delete with delete[]
void GetMapiMsiIds(LPTSTR szClient, LPTSTR* lpszComponentID, LPTSTR* lpszAppLCID, LPTSTR* lpszOfficeLCID)
{
	Log(true,_T("GetMapiMsiIds()\n"));
	LONG lRes = S_OK;

	HKEY hKey = GetMailKey(szClient);

	if (hKey)
	{
		DWORD dwKeyType = NULL;

		if (lpszComponentID)
		{
			lRes = GetRegistryValue(
				hKey,
				_T("MSIComponentID"),
				&dwKeyType,
				(LPVOID*) lpszComponentID);
			Log(true,_T("MSIComponentID = %s\n"),*lpszComponentID?*lpszComponentID:_T("<not found>"));
		}

		if (lpszAppLCID)
		{
			lRes = GetRegistryValue(
				hKey,
				_T("MSIApplicationLCID"),
				&dwKeyType,
				(LPVOID*) lpszAppLCID);
			Log(true,_T("MSIApplicationLCID = %s\n"),*lpszAppLCID?*lpszAppLCID:_T("<not found>"));
		}

		if (lpszOfficeLCID)
		{
			lRes = GetRegistryValue(
				hKey,
				_T("MSIOfficeLCID"),
				&dwKeyType,
				(LPVOID*) lpszOfficeLCID);
			Log(true,_T("MSIOfficeLCID = %s\n"),*lpszOfficeLCID?*lpszOfficeLCID:_T("<not found>"));
		}

		RegCloseKey(hKey);
	}
} // GetMapiMsiIds

void GetMAPIPath(LPTSTR szClient, LPTSTR szMAPIPath, ULONG cchMAPIPath)
{
	BOOL bRet = false;

	szMAPIPath[0] = '\0'; // Terminate String at pos 0 (safer if we fail below)

	// Find some strings:
	LPTSTR szComponentID = NULL;
	LPTSTR szAppLCID = NULL;
	LPTSTR szOfficeLCID = NULL;

	GetMapiMsiIds(szClient,&szComponentID,&szAppLCID,&szOfficeLCID);

	if (szComponentID)
	{
		if (szAppLCID)
		{
			bRet = GetComponentPath(
				szComponentID,
				szAppLCID,
				szMAPIPath,
				cchMAPIPath);
		}
		if ((!bRet || szMAPIPath[0] == _T('\0')) && szOfficeLCID)
		{
			bRet = GetComponentPath(
				szComponentID,
				szOfficeLCID,
				szMAPIPath,
				cchMAPIPath);
		}
		if (!bRet  || szMAPIPath[0] == _T('\0'))
		{
			bRet = GetComponentPath(
				szComponentID,
				NULL,
				szMAPIPath,
				cchMAPIPath);
		}
	}

	delete[] szComponentID;
	delete[] szOfficeLCID;
	delete[] szAppLCID;
} // GetMAPIPath

void AutoLoadMAPI()
{
	Log(true,_T("AutoLoadMAPI - loading MAPI exports\n"));

	// Attempt load default MAPI first if we have one
	// This will handle Outlook MAPI
	if (!hModMSMAPI)
	{
		TCHAR szMSMAPI32path[MAX_PATH] = {0};
		GetMAPIPath(NULL,szMSMAPI32path,CCH(szMSMAPI32path));
		if (szMSMAPI32path[0] != NULL)
		{
			hModMSMAPI = MyLoadLibrary(szMSMAPI32path);
		}
	}
	if (hModMSMAPI) LoadMAPIFuncs(hModMSMAPI);

	// In case that fails - load the stub library from system32
	// This will handle rare cases where we don't load Outlook MAPI
	if (!hModMAPI) hModMAPI = LoadFromSystemDir(_T("mapi32.dll"));
	// If hModMAPI is the same as hModMSMAPI, then we don't need to try it again
	if (hModMAPI && hModMAPI != hModMSMAPI) LoadMAPIFuncs(hModMAPI);
} // AutoLoadMAPI

void UnloadMAPI()
{
	Log(true,_T("UnloadMAPI - unloading MAPI exports\n"));
	// Blank out all the functions we've loaded from MAPI DLLs:
	pfnFGetComponentPath = NULL;
	pfnLaunchWizard = NULL;
	pfnMAPIAllocateBuffer = NULL;
	pfnMAPIAllocateMore = NULL;
	pfnMAPIFreeBuffer = NULL;
	pfnHrGetOneProp = NULL;
	pfnMAPIInitialize = NULL;
	pfnMAPIUninitialize = NULL;
	pfnLPFreeProws = NULL;
	pfnPpropFindProp = NULL;
	pfnScDupPropset = NULL;
	pfnScCountProps = NULL;
	pfnScCopyProps = NULL;
	pfnOpenIMsgOnIStg = NULL;
	pfnMAPIGetDefaultMalloc = NULL;
	pfnOpenTnefStreamEx = NULL;
	pfnOpenStreamOnFile = NULL;
	pfnCloseIMsgSession = NULL;
	pfnOpenIMsgSession = NULL;
	pfnHrQueryAllRows = NULL;
	pfnMAPIOpenFormMgr = NULL;
	pfnRTFSync = NULL;
	pfnHrSetOneProp = NULL;
	pfnFreePadrlist = NULL;
	pfnPropCopyMore = NULL;
	pfnWrapCompressedRTFStream = NULL;
	pfnMAPILogonEx = NULL;
	pfnMAPIAdminProfiles = NULL;
	pfnHrValidateIPMSubtree = NULL;
	pfnMAPIOpenLocalFormContainer = NULL;
	pfnHrDispatchNotifications = NULL;
	
	if (hModMSMAPI) FreeLibrary(hModMSMAPI);
	hModMSMAPI = NULL;

	if (hModMAPIStub) FreeLibrary(hModMAPIStub);
	hModMAPIStub = NULL;

	if (hModMAPI) FreeLibrary(hModMAPI);
	hModMAPI = NULL;
}

// Translation - if pfn isn't already set, GetProcAddress for it from hMod
#define GETPROC(pfn,_TYPE, ProcName) \
if (!(pfn)) \
{ \
	(pfn) = (_TYPE) GetProcAddress(hMod, (ProcName)); \
	if (!(pfn)) Log(true,_T("Failed to load \"%hs\" from 0x%08X\n"),(ProcName),(hMod)); \
}

void LoadMAPIFuncs(HMODULE hMod)
{
	Log(true,_T("LoadMAPIFuncs - loading from 0x%08X\n"),hMod);
	if (!hMod) return;

	GETPROC(pfnLaunchWizard,			LPLAUNCHWIZARDENTRY,		LAUNCHWIZARDENTRYNAME)
	GETPROC(pfnMAPIAllocateBuffer,		LPMAPIALLOCATEBUFFER,		"MAPIAllocateBuffer")
	GETPROC(pfnMAPIAllocateMore,		LPMAPIALLOCATEMORE,			"MAPIAllocateMore")
	GETPROC(pfnMAPIFreeBuffer,			LPMAPIFREEBUFFER,			"MAPIFreeBuffer")
	GETPROC(pfnHrGetOneProp,			LPHRGETONEPROP,				"HrGetOneProp@12")
	GETPROC(pfnHrGetOneProp,			LPHRGETONEPROP,				"HrGetOneProp")
	GETPROC(pfnMAPIInitialize,			LPMAPIINITIALIZE,			"MAPIInitialize")
	GETPROC(pfnMAPIUninitialize,		LPMAPIUNINITIALIZE,			"MAPIUninitialize")
	GETPROC(pfnLPFreeProws,				LPFREEPROWS,				"FreeProws@4")
	GETPROC(pfnLPFreeProws,				LPFREEPROWS,				"FreeProws")
	GETPROC(pfnPpropFindProp,			LPPPROPFINDPROP,			"PpropFindProp@12")
	GETPROC(pfnPpropFindProp,			LPPPROPFINDPROP,			"PpropFindProp")
	GETPROC(pfnScDupPropset,			LPSCDUPPROPSET,				"ScDupPropset@16")
	GETPROC(pfnScDupPropset,			LPSCDUPPROPSET,				"ScDupPropset")
	GETPROC(pfnScCountProps,			LPSCCOUNTPROPS,				"ScCountProps@12")
	GETPROC(pfnScCountProps,			LPSCCOUNTPROPS,				"ScCountProps")
	GETPROC(pfnScCopyProps,				LPSCCOPYPROPS,				"ScCopyProps@16")
	GETPROC(pfnScCopyProps,				LPSCCOPYPROPS,				"ScCopyProps")
	GETPROC(pfnOpenIMsgOnIStg,			LPOPENIMSGONISTG,			"OpenIMsgOnIStg@44")
	GETPROC(pfnOpenIMsgOnIStg,			LPOPENIMSGONISTG,			"OpenIMsgOnIStg")
	GETPROC(pfnMAPIGetDefaultMalloc,	LPMAPIGETDEFAULTMALLOC,		"MAPIGetDefaultMalloc@0")
	GETPROC(pfnMAPIGetDefaultMalloc,	LPMAPIGETDEFAULTMALLOC,		"MAPIGetDefaultMalloc")
	GETPROC(pfnOpenTnefStreamEx,		LPOPENTNEFSTREAMEX,			"OpenTnefStreamEx")
	GETPROC(pfnOpenStreamOnFile,		LPOPENSTREAMONFILE,			"OpenStreamOnFile")
	GETPROC(pfnCloseIMsgSession,		LPCLOSEIMSGSESSION,			"CloseIMsgSession@4")
	GETPROC(pfnCloseIMsgSession,		LPCLOSEIMSGSESSION,			"CloseIMsgSession")
	GETPROC(pfnOpenIMsgSession,			LPOPENIMSGSESSION,			"OpenIMsgSession@12")
	GETPROC(pfnOpenIMsgSession,			LPOPENIMSGSESSION,			"OpenIMsgSession")
	GETPROC(pfnHrQueryAllRows,			LPHRQUERYALLROWS,			"HrQueryAllRows@24")
	GETPROC(pfnHrQueryAllRows,			LPHRQUERYALLROWS,			"HrQueryAllRows")
	GETPROC(pfnMAPIOpenFormMgr,			LPMAPIOPENFORMMGR,			"MAPIOpenFormMgr")
	GETPROC(pfnRTFSync,					LPRTFSYNC,					"RTFSync")
	GETPROC(pfnHrSetOneProp,			LPHRSETONEPROP,				"HrSetOneProp@8")
	GETPROC(pfnHrSetOneProp,			LPHRSETONEPROP,				"HrSetOneProp")
	GETPROC(pfnFreePadrlist,			LPFREEPADRLIST,				"FreePadrlist@4")
	GETPROC(pfnFreePadrlist,			LPFREEPADRLIST,				"FreePadrlist")
	GETPROC(pfnPropCopyMore,			LPPROPCOPYMORE,				"PropCopyMore@16")
	GETPROC(pfnPropCopyMore,			LPPROPCOPYMORE,				"PropCopyMore")
	GETPROC(pfnWrapCompressedRTFStream,	LPWRAPCOMPRESSEDRTFSTREAM,	"WrapCompressedRTFStream")
	GETPROC(pfnWrapStoreEntryID,		LPWRAPSTOREENTRYID,			"WrapStoreEntryID@24")
	GETPROC(pfnWrapStoreEntryID,		LPWRAPSTOREENTRYID,			"WrapStoreEntryID")
	GETPROC(pfnCreateIProp,				LPCREATEIPROP,				"CreateIProp@24")
	GETPROC(pfnCreateIProp,				LPCREATEIPROP,				"CreateIProp")
	GETPROC(pfnCreateTable,				LPCREATETABLE,				"CreateTable@36")
	GETPROC(pfnCreateTable,				LPCREATETABLE,				"CreateTable")
	GETPROC(pfnHrValidateParameters,	LPHRVALIDATEPARAMETERS,		"HrValidateParameters@8")
	GETPROC(pfnHrValidateParameters,	LPHRVALIDATEPARAMETERS,		"HrValidateParameters")
	GETPROC(pfnBuildDisplayTable,		LPBUILDDISPLAYTABLE,		"BuildDisplayTable@40")
	GETPROC(pfnBuildDisplayTable,		LPBUILDDISPLAYTABLE,		"BuildDisplayTable")
	GETPROC(pfnMNLS_lstrlenW,			LPMNLS_LSTRLENW,			"MNLS_lstrlenW")
	GETPROC(pfnMAPILogonEx,				LPMAPILOGONEX,				"MAPILogonEx")
	GETPROC(pfnMAPIAdminProfiles,		LPMAPIADMINPROFILES,		"MAPIAdminProfiles")
	GETPROC(pfnHrValidateIPMSubtree,	LPHRVALIDATEIPMSUBTREE,		"HrValidateIPMSubtree@20")
	GETPROC(pfnHrValidateIPMSubtree,	LPHRVALIDATEIPMSUBTREE,		"HrValidateIPMSubtree")
	GETPROC(pfnMAPIOpenLocalFormContainer, LPMAPIOPENLOCALFORMCONTAINER,"MAPIOpenLocalFormContainer")
	GETPROC(pfnHrDispatchNotifications,	LPHRDISPATCHNOTIFICATIONS,	"HrDispatchNotifications@4")
	GETPROC(pfnHrDispatchNotifications,	LPHRDISPATCHNOTIFICATIONS,	"HrDispatchNotifications")
}

#define CHECKLOAD(pfn) \
{ \
	if (!(pfn)) AutoLoadMAPI(); \
	if (!(pfn)) \
	{ \
		Log(true,_T("Function %hs not loaded\n"),#pfn); \
	} \
}

STDMETHODIMP_(SCODE) MAPIAllocateBuffer(ULONG cbSize,
										LPVOID FAR * lppBuffer)
{
	CHECKLOAD(pfnMAPIAllocateBuffer);
	if (pfnMAPIAllocateBuffer) return pfnMAPIAllocateBuffer(cbSize,lppBuffer);
	return MAPI_E_CALL_FAILED;
}

STDMETHODIMP_(SCODE) MAPIAllocateMore(ULONG cbSize,
									  LPVOID lpObject,
									  LPVOID FAR * lppBuffer)
{
	CHECKLOAD(pfnMAPIAllocateMore);
	if (pfnMAPIAllocateMore) return pfnMAPIAllocateMore(cbSize,lpObject,lppBuffer);
	return MAPI_E_CALL_FAILED;
}

STDAPI_(ULONG) MAPIFreeBuffer(LPVOID lpBuffer)
{
	// If we don't have this, the memory's already leaked - don't compound it with a crash
	if (!lpBuffer) return NULL;
	CHECKLOAD(pfnMAPIFreeBuffer);
	if (pfnMAPIFreeBuffer) return pfnMAPIFreeBuffer(lpBuffer);
	return NULL;
}

STDAPI HrGetOneProp(LPMAPIPROP lpMapiProp,
					ULONG ulPropTag,
					LPSPropValue FAR * lppProp)
{
	CHECKLOAD(pfnHrGetOneProp);
	if (pfnHrGetOneProp) return pfnHrGetOneProp(lpMapiProp,ulPropTag,lppProp);
	return MAPI_E_CALL_FAILED;
}

STDMETHODIMP MAPIInitialize(LPVOID lpMapiInit)
{
	CHECKLOAD(pfnMAPIInitialize);
	if (pfnMAPIInitialize) return pfnMAPIInitialize(lpMapiInit);
	return MAPI_E_CALL_FAILED;
}

STDAPI_(void) MAPIUninitialize(void)
{
	CHECKLOAD(pfnMAPIUninitialize);
	if (pfnMAPIUninitialize) pfnMAPIUninitialize();
	return;
}

STDAPI_(void) FreeProws(LPSRowSet lpRows)
{
	CHECKLOAD(pfnLPFreeProws);
	if (pfnLPFreeProws) pfnLPFreeProws(lpRows);
	return;
}

STDAPI_(LPSPropValue) PpropFindProp(LPSPropValue lpPropArray,
									ULONG cValues,
									ULONG ulPropTag)
{
	CHECKLOAD(pfnPpropFindProp);
	if (pfnPpropFindProp) return pfnPpropFindProp(lpPropArray,cValues,ulPropTag);
	return NULL;
}

STDAPI_(SCODE) ScDupPropset(int cValues,
							LPSPropValue lpPropArray,
							LPALLOCATEBUFFER lpAllocateBuffer,
							LPSPropValue FAR *lppPropArray)
{
	CHECKLOAD(pfnScDupPropset);
	if (pfnScDupPropset) return pfnScDupPropset(cValues,lpPropArray,lpAllocateBuffer,lppPropArray);
	return MAPI_E_CALL_FAILED;
}

STDAPI_(SCODE) ScCountProps(int cValues,
							LPSPropValue lpPropArray,
							ULONG FAR *lpcb)
{
	CHECKLOAD(pfnScCountProps);
	if (pfnScCountProps) return pfnScCountProps(cValues,lpPropArray,lpcb);
	return MAPI_E_CALL_FAILED;
}

STDAPI_(SCODE) ScCopyProps(int cValues,
							LPSPropValue lpPropArray,
							LPVOID lpvDst,
							ULONG FAR *lpcb)
{
	CHECKLOAD(pfnScCopyProps);
	if (pfnScCopyProps) return pfnScCopyProps(cValues,lpPropArray,lpvDst,lpcb);
	return MAPI_E_CALL_FAILED;
}

STDAPI_(SCODE) OpenIMsgOnIStg(LPMSGSESS lpMsgSess,
							  LPALLOCATEBUFFER lpAllocateBuffer,
							  LPALLOCATEMORE lpAllocateMore,
							  LPFREEBUFFER lpFreeBuffer,
							  LPMALLOC lpMalloc,
							  LPVOID lpMapiSup,
							  LPSTORAGE lpStg,
							  MSGCALLRELEASE FAR * lpfMsgCallRelease,
							  ULONG ulCallerData,
							  ULONG ulFlags,
							  LPMESSAGE FAR * lppMsg)
{
	CHECKLOAD(pfnOpenIMsgOnIStg);
	if (pfnOpenIMsgOnIStg) return pfnOpenIMsgOnIStg(
		lpMsgSess,
		lpAllocateBuffer,
		lpAllocateMore,
		lpFreeBuffer,
		lpMalloc,
		lpMapiSup,
		lpStg,
		lpfMsgCallRelease,
		ulCallerData,
		ulFlags,
		lppMsg);
	return MAPI_E_CALL_FAILED;
}

STDAPI_(LPMALLOC) MAPIGetDefaultMalloc(VOID)
{
	CHECKLOAD(pfnMAPIGetDefaultMalloc);
	if (pfnMAPIGetDefaultMalloc) return pfnMAPIGetDefaultMalloc();
	return NULL;
}

STDMETHODIMP OpenTnefStreamEx(LPVOID lpvSupport,
							  LPSTREAM lpStream,
							  LPTSTR lpszStreamName,
							  ULONG ulFlags,
							  LPMESSAGE lpMessage,
							  WORD wKeyVal,
							  LPADRBOOK lpAdressBook,
							  LPITNEF FAR * lppTNEF)
{
	CHECKLOAD(pfnOpenTnefStreamEx);
	if (pfnOpenTnefStreamEx) return pfnOpenTnefStreamEx(
		lpvSupport,
		lpStream,
		lpszStreamName,
		ulFlags,
		lpMessage,
		wKeyVal,
		lpAdressBook,
		lppTNEF);
	return MAPI_E_CALL_FAILED;
}

STDAPI_(void) CloseIMsgSession(LPMSGSESS lpMsgSess)
{
	CHECKLOAD(pfnCloseIMsgSession);
	if (pfnCloseIMsgSession) pfnCloseIMsgSession(
		lpMsgSess);
	return;
}

STDAPI_(SCODE) OpenIMsgSession(LPMALLOC lpMalloc,
							   ULONG ulFlags,
							   LPMSGSESS FAR * lppMsgSess)
{
	CHECKLOAD(pfnOpenIMsgSession);
	if (pfnOpenIMsgSession) return pfnOpenIMsgSession(
		lpMalloc,
		ulFlags,
		lppMsgSess);
	return MAPI_E_CALL_FAILED;
}

STDAPI HrQueryAllRows(LPMAPITABLE lpTable,
					  LPSPropTagArray lpPropTags,
					  LPSRestriction lpRestriction,
					  LPSSortOrderSet lpSortOrderSet,
					  LONG crowsMax,
					  LPSRowSet FAR *lppRows)
{
	CHECKLOAD(pfnHrQueryAllRows);
	if (pfnHrQueryAllRows) return pfnHrQueryAllRows(
		lpTable,
		lpPropTags,
		lpRestriction,
		lpSortOrderSet,
		crowsMax,
		lppRows);
	return MAPI_E_CALL_FAILED;
}

STDAPI MAPIOpenFormMgr(LPMAPISESSION pSession, LPMAPIFORMMGR FAR * ppmgr)
{
	CHECKLOAD(pfnMAPIOpenFormMgr);
	if (pfnMAPIOpenFormMgr) return pfnMAPIOpenFormMgr(
		pSession,
		ppmgr);
	return MAPI_E_CALL_FAILED;
}

STDAPI_(HRESULT) RTFSync (LPMESSAGE lpMessage, ULONG ulFlags, BOOL FAR * lpfMessageUpdated)
{
	CHECKLOAD(pfnRTFSync);
	if (pfnRTFSync) return pfnRTFSync(lpMessage,ulFlags,lpfMessageUpdated);
	return MAPI_E_CALL_FAILED;
}

STDAPI HrSetOneProp(LPMAPIPROP lpMapiProp,
					LPSPropValue lpProp)
{
	CHECKLOAD(pfnHrSetOneProp);
	if (pfnHrSetOneProp) return pfnHrSetOneProp(lpMapiProp,lpProp);
	return MAPI_E_CALL_FAILED;
}

STDAPI_(void) FreePadrlist(LPADRLIST lpAdrlist)
{
	CHECKLOAD(pfnFreePadrlist);
	if (pfnFreePadrlist) pfnFreePadrlist(lpAdrlist);
	return;
}

STDAPI_(SCODE) PropCopyMore(LPSPropValue lpSPropValueDest,
							LPSPropValue lpSPropValueSrc,
							ALLOCATEMORE * lpfAllocMore,
							LPVOID lpvObject)
{
	CHECKLOAD(pfnPropCopyMore);
	if (pfnPropCopyMore) pfnPropCopyMore(lpSPropValueDest,lpSPropValueSrc,lpfAllocMore,lpvObject);
	return MAPI_E_CALL_FAILED;
}

STDAPI_(HRESULT) WrapCompressedRTFStream(LPSTREAM lpCompressedRTFStream,
										 ULONG ulFlags,
										 LPSTREAM FAR * lpUncompressedRTFStream)
{
	CHECKLOAD(pfnWrapCompressedRTFStream);
	if (pfnWrapCompressedRTFStream) return pfnWrapCompressedRTFStream(
		lpCompressedRTFStream,
		ulFlags,
		lpUncompressedRTFStream);
	return MAPI_E_CALL_FAILED;
}

STDMETHODIMP MAPILogonEx(ULONG_PTR ulUIParam,
						 LPTSTR lpszProfileName,
						 LPTSTR lpszPassword,
						 ULONG ulFlags,
						 LPMAPISESSION FAR * lppSession)
{
	CHECKLOAD(pfnMAPILogonEx);
	if (pfnMAPILogonEx) return pfnMAPILogonEx(
		ulUIParam,
		lpszProfileName,
		lpszPassword,
		ulFlags,
		lppSession);
	return MAPI_E_CALL_FAILED;
}

STDMETHODIMP MAPIAdminProfiles(ULONG ulFlags,
							   LPPROFADMIN FAR *lppProfAdmin)
{
	CHECKLOAD(pfnMAPIAdminProfiles);
	if (pfnMAPIAdminProfiles) return pfnMAPIAdminProfiles(
		ulFlags,
		lppProfAdmin);
	return MAPI_E_CALL_FAILED;
}

STDAPI HrValidateIPMSubtree(LPMDB lpMDB,
							ULONG ulFlags,
							ULONG FAR *lpcValues,
							LPSPropValue FAR *lppValues,
							LPMAPIERROR FAR *lpperr)
{
	CHECKLOAD(pfnHrValidateIPMSubtree);
	if (pfnHrValidateIPMSubtree) return pfnHrValidateIPMSubtree(
		lpMDB,
		ulFlags,
		lpcValues,
		lppValues,
		lpperr);
	return MAPI_E_CALL_FAILED;
}

STDAPI MAPIOpenLocalFormContainer(LPMAPIFORMCONTAINER FAR * ppfcnt)
{
	CHECKLOAD(pfnMAPIOpenLocalFormContainer);
	if (pfnMAPIOpenLocalFormContainer) return pfnMAPIOpenLocalFormContainer(
		ppfcnt);
	return MAPI_E_CALL_FAILED;
}

STDAPI HrDispatchNotifications(ULONG ulFlags)
{
	CHECKLOAD(pfnHrDispatchNotifications);
	if (pfnHrDispatchNotifications) return pfnHrDispatchNotifications(
		ulFlags);
	return MAPI_E_CALL_FAILED;
}

STDAPI WrapStoreEntryID(ULONG ulFlags, __in LPTSTR lpszDLLName, ULONG cbOrigEntry,
	LPENTRYID lpOrigEntry, ULONG *lpcbWrappedEntry, LPENTRYID *lppWrappedEntry)
{
	CHECKLOAD(pfnWrapStoreEntryID);
	if (pfnWrapStoreEntryID) return pfnWrapStoreEntryID(ulFlags, lpszDLLName, cbOrigEntry, lpOrigEntry, lpcbWrappedEntry, lppWrappedEntry);
	return MAPI_E_CALL_FAILED;
}
STDAPI_(SCODE)
CreateIProp( LPCIID					lpInterface,
			 ALLOCATEBUFFER FAR *	lpAllocateBuffer,
			 ALLOCATEMORE FAR *		lpAllocateMore,
			 FREEBUFFER FAR *		lpFreeBuffer,
			 LPVOID					lpvReserved,
			 LPPROPDATA FAR *		lppPropData )
{
	CHECKLOAD(pfnCreateIProp);
	if (pfnCreateIProp) return pfnCreateIProp(lpInterface, lpAllocateBuffer, lpAllocateMore, lpFreeBuffer, lpvReserved, lppPropData);
	return MAPI_E_CALL_FAILED;
}

STDAPI_(SCODE)
CreateTable( LPCIID					lpInterface,
			 ALLOCATEBUFFER FAR *	lpAllocateBuffer,
			 ALLOCATEMORE FAR *		lpAllocateMore,
			 FREEBUFFER FAR *		lpFreeBuffer,
			 LPVOID					lpvReserved,
			 ULONG					ulTableType,
			 ULONG					ulPropTagIndexColumn,
			 LPSPropTagArray		lpSPropTagArrayColumns,
			 LPTABLEDATA FAR *		lppTableData )
{
	CHECKLOAD(pfnCreateTable);
	if (pfnCreateTable) return pfnCreateTable(
		lpInterface,
		lpAllocateBuffer,
		lpAllocateMore,
		lpFreeBuffer,
		lpvReserved,
		ulTableType,
		ulPropTagIndexColumn,
		lpSPropTagArrayColumns,
		lppTableData);
	return MAPI_E_CALL_FAILED;
}

STDAPI HrValidateParameters( METHODS eMethod, LPVOID FAR *ppFirstArg)
{
	CHECKLOAD(pfnHrValidateParameters);
	if (pfnHrValidateParameters) return pfnHrValidateParameters(
		eMethod,
		ppFirstArg);
	return MAPI_E_CALL_FAILED;
}

STDAPI
BuildDisplayTable(	LPALLOCATEBUFFER	lpAllocateBuffer,
					LPALLOCATEMORE		lpAllocateMore,
					LPFREEBUFFER		lpFreeBuffer,
					LPMALLOC			lpMalloc,
					HINSTANCE			hInstance,
					UINT				cPages,
					LPDTPAGE			lpPage,
					ULONG				ulFlags,
					LPMAPITABLE *		lppTable,
					LPTABLEDATA	*		lppTblData )
{
	CHECKLOAD(pfnBuildDisplayTable);
	if (pfnBuildDisplayTable) return pfnBuildDisplayTable(
		lpAllocateBuffer,
		lpAllocateMore,
		lpFreeBuffer,
		lpMalloc,
		hInstance,
		cPages,
		lpPage,
		ulFlags,
		lppTable,
		lppTblData);
	return MAPI_E_CALL_FAILED;
}

int	WINAPI MNLS_lstrlenW(LPCWSTR lpString)
{
	CHECKLOAD(pfnMNLS_lstrlenW);
	if (pfnMNLS_lstrlenW) return pfnMNLS_lstrlenW(
		lpString);
	return 0;
}