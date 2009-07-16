#include "stdafx.h"
#include <MAPISPI.H>
#include <mspst.h>
#include "edkmdb.h"

#define INITGUID
#define USES_IID_IMAPISup
#define USES_IID_IMSProvider
#define USES_IID_IMAPIProp
#define USES_IID_IMsgStore
#include <initguid.h>
#include <mapiguid.h>

typedef MSPROVIDERINIT FAR *LPMSPROVIDERINIT;

#define MDB_OST_LOGON_UNICODE	((ULONG) 0x00000800)
#define MDB_OST_LOGON_ANSI		((ULONG) 0x00001000)

// OffLineFileInfo
typedef struct {
	ULONG    ulVersion;     // Structure version
	MAPIUID  muidReserved;
	ULONG    ulReserved;
	DWORD    dwAlloc;       // Number of primary source keys
	DWORD    dwNextAlloc;
	LTID     ltidAlloc;
	LTID     ltidNextAlloc;
} OLFI, *POLFI;

#define OLFI_VERSION 0x01

#define GLOB_COUNT_RANGE 0x80000000
#define cbGlobcnt	6
#define FEqLtid( pltid1, pltid2 )       \
	(memcmp(pltid1, pltid2, sizeof(GUID) + cbGlobcnt) == 0)
LTID	ltidZero = { { 0, 0, 0, {0,0,0,0,0,0,0,0} }, { 0,0,0,0,0,0 }, {0} };

// UID of an NST provider - do not change this
const MAPIUID g_muidProvPrvNST =
	{ 0xE9, 0x2F, 0xEB, 0x75, 0x96, 0x50, 0x44, 0x86,
	  0x83, 0xB8, 0x7D, 0xE5, 0x22, 0xAA, 0x49, 0x48 };

#define MS_DLL_NAME_STRING	  "wrppst.dll"
#define PROFILE_USER_STRING "My PR_PROFILE_USER"

LPALLOCATEBUFFER g_lpAllocateBuffer = NULL;
LPALLOCATEMORE g_lpAllocateMore = NULL;
LPFREEBUFFER g_lpFreeBuffer = NULL;

SCODE STDMETHODCALLTYPE MyAllocateBuffer(
										 ULONG			cbSize,
										 LPVOID FAR *	lppBuffer)
{
	if (g_lpAllocateBuffer) return g_lpAllocateBuffer(cbSize,lppBuffer);

	Log(true, "MyAllocateBuffer: g_lpAllocateBuffer not set!\n");
	return MAPI_E_INVALID_PARAMETER;
}

SCODE STDMETHODCALLTYPE MyAllocateMore(
									   ULONG			cbSize,
									   LPVOID			lpObject,
									   LPVOID FAR * lppBuffer)
{
	if (g_lpAllocateMore) return g_lpAllocateMore(cbSize,lpObject,lppBuffer);

	Log(true, "MyAllocateMore: g_lpAllocateMore not set!\n");
	return MAPI_E_INVALID_PARAMETER;
}

ULONG STDAPICALLTYPE MyFreeBuffer(
								  LPVOID			lpBuffer)
{
	if (g_lpFreeBuffer) return g_lpFreeBuffer(lpBuffer);

	Log(true, "MyFreeBuffer: g_lpFreeBuffer not set!\n");
	return (ULONG) MAPI_E_INVALID_PARAMETER;
}

STDINITMETHODIMP MSProviderInit (
								 HINSTANCE				/*hInstance*/,
								 LPMALLOC				lpMalloc,
								 LPALLOCATEBUFFER		lpAllocateBuffer,
								 LPALLOCATEMORE			lpAllocateMore,
								 LPFREEBUFFER			lpFreeBuffer,
								 ULONG					ulFlags,
								 ULONG					ulMAPIVer,
								 ULONG FAR *			lpulProviderVer,
								 LPMSPROVIDER FAR *		lppMSProvider)
{
	Log(true,"MSProviderInit function called\n");
	if (!lppMSProvider || !lpulProviderVer) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	*lppMSProvider = NULL;
	*lpulProviderVer = CURRENT_SPI_VERSION;
	if (ulMAPIVer < CURRENT_SPI_VERSION)
	{
		Log(true,"MSProviderInit: The version of the subsystem cannot handle this version of the provider\n");
		return MAPI_E_VERSION;
	}

	Log(true,"MSProviderInit: saving off memory management routines\n");
	g_lpAllocateBuffer = lpAllocateBuffer;
	g_lpAllocateMore = lpAllocateMore;
	g_lpFreeBuffer = lpFreeBuffer;

	LPTSTR lpszPath = GetPSTPath();
	HMODULE hm = NULL;
	if (lpszPath)
	{
		Log(true, "Got PST path %s\n", lpszPath);

		hm = LoadLibrary(lpszPath);
		Log(true,"LoadLibrary returned 0x%08X\n", hm);
		delete[] lpszPath;
	}

	LPMSPROVIDERINIT pMsProviderInit = NULL;
	if (hm)
	{
		pMsProviderInit = (LPMSPROVIDERINIT)GetProcAddress(hm, "MSProviderInit");
		Log(true,"GetProcAddress returned 0x%08X\n", pMsProviderInit);
	}
	else hRes = E_OUTOFMEMORY;

	if (pMsProviderInit)
	{
		Log(true,"Calling pMsProviderInit\n");
		CMSProvider* pWrappedProvider = NULL;
		LPMSPROVIDER pSyncProviderObj = NULL;

		hRes = pMsProviderInit(
			hm, // not hInstance: first param is handle of module in which the function lives
			lpMalloc,
			lpAllocateBuffer,
			lpAllocateMore,
			lpFreeBuffer,
			ulFlags,
			ulMAPIVer,
			lpulProviderVer,
			&pSyncProviderObj
			);

		Log(true,"pMsProviderInit returned 0x%08X\n", hRes);
		if (SUCCEEDED(hRes))
		{
			pWrappedProvider = new CMSProvider (pSyncProviderObj);
			if (NULL == pWrappedProvider)
			{
				Log(true,"MSProviderInit: Failed to allocate new CMSProvider object\n");
				hRes = E_OUTOFMEMORY;
			}

			// Copy pointer to the allocated object back into the return IMSProvider object pointer
			*lppMSProvider = pWrappedProvider;
		}
	}
	else hRes = E_OUTOFMEMORY;

	return hRes;
}

STDMETHODIMP GetGlobalProfileObject(LPMAPISUP	   lpMAPISupObj,
									LPPROFSECT*    lppGlobalProfileObj)
{
	Log(true,"GetGlobalProfileObject\n");
	if (!lpMAPISupObj || !lppGlobalProfileObj) return MAPI_E_INVALID_PARAMETER;

	HRESULT 	 hRes = S_OK;
	LPSPropValue lpServiceUIDProp = NULL;
	LPPROFSECT	 lpProfileObj = NULL;
	LPPROFSECT	 lpGlobalProfileObj = NULL;

	*lppGlobalProfileObj = NULL;

	// Get information for the provider's profile section
	hRes = lpMAPISupObj->OpenProfileSection(
		(LPMAPIUID) NULL,
		MAPI_MODIFY,
		&lpProfileObj);
	if (FAILED(hRes))
	{
		Log(true,"GetProfileProps: OpenProfileSection(NULL) failed\n");
	}

	if (lpProfileObj)
	{
		hRes = HrGetOneProp(lpProfileObj, PR_SERVICE_UID, &lpServiceUIDProp);
		if (FAILED(hRes))
		{
			if (hRes == MAPI_E_NOT_FOUND)
			{
				Log(true,"GetProfileProps: Returning OpenProfileSection(NULL)\n");
				// this is itself the global profile section
				*lppGlobalProfileObj = lpProfileObj;
				lpProfileObj = NULL;
				hRes = S_OK;
			}
			else
			{
				Log(true,"GetProfileProps: Failed to get PR_SERVICE_UID\n");
			}
		}
		else if (lpServiceUIDProp && PR_SERVICE_UID == lpServiceUIDProp->ulPropTag)
		{
			Log(true,"GetProfileProps: Got PR_SERVICE_UID\n");
			hRes = lpMAPISupObj->OpenProfileSection(
				(LPMAPIUID) lpServiceUIDProp->Value.bin.lpb,
				MAPI_MODIFY,
				&lpGlobalProfileObj);
			if (FAILED(hRes))
			{
				Log(true,"GetProfileProps: OpenProfileSection failed to open global profile section\n");
			}
			else
			{
				Log(true,"GetProfileProps: returning section pointed to by PR_SERVICE_UID\n");
				*lppGlobalProfileObj = lpGlobalProfileObj;
				lpGlobalProfileObj = NULL;
			}
		}
		MyFreeBuffer(lpServiceUIDProp);
	}

	if (lpProfileObj) lpProfileObj->Release();
	if (lpGlobalProfileObj) lpGlobalProfileObj->Release();
	Log(true,"GetGlobalProfileObject returned 0x%08X\n",hRes);
	return hRes;
} // GetGlobalProfileObject

HRESULT SetOfflineStoreProps(LPPROFSECT lpProfSect,ULONG ulUIParam)
{
	Log(true,"SetOfflineStoreProps\n");
	if (!lpProfSect) return MAPI_E_INVALID_PARAMETER;

	HRESULT 		hRes = S_OK;
	LPTSTR			lpszNSTFilePath = NULL;
	ULONG			ulProfProps = 0;
	LPSPropValue	lpProfProps = NULL;

	// See what props there are
	hRes = lpProfSect->GetProps((LPSPropTagArray)&sptClientProps,
								fMapiUnicode,
								&ulProfProps,
								&lpProfProps);
	if (SUCCEEDED(hRes) && lpProfProps && NUM_CLIENT_PROPERTIES == ulProfProps)
	{
		// If there isn't a path to the NST set, go get one
		if (sptClientProps.aulPropTag[CLIENT_PST_PATH] != lpProfProps[CLIENT_PST_PATH].ulPropTag)
		{
			BOOL bRet = false;
			OPENFILENAME OpenFile = {0};

			lpszNSTFilePath = new TCHAR[MAX_PATH+1];
			memset(lpszNSTFilePath, 0, MAX_PATH+1 * sizeof(TCHAR));

			OpenFile.lStructSize = sizeof(OPENFILENAME);
			OpenFile.lpstrFilter = _T("Wrapped PST Offline Store (*.NST)\0*.NST\0\0");
			OpenFile.nFilterIndex = 1;
			OpenFile.lpstrTitle = _T("Select a Wrapped PST Offline Store file");
			OpenFile.Flags = OFN_LONGNAMES | OFN_NOCHANGEDIR | OFN_NONETWORKBUTTON | OFN_HIDEREADONLY;
			OpenFile.lpstrDefExt = "NST";
			OpenFile.lpstrFile = lpszNSTFilePath;
			OpenFile.nMaxFile = MAX_PATH;
			OpenFile.hwndOwner = (HWND)ulUIParam;

			bRet = GetOpenFileName(&OpenFile);

			if (bRet)
			{
				Log(true, "User chose offline file: %s\n", lpszNSTFilePath);
				lpProfProps[CLIENT_PST_PATH].ulPropTag = sptClientProps.aulPropTag[CLIENT_PST_PATH];
				lpProfProps[CLIENT_PST_PATH].Value.lpszA = lpszNSTFilePath;
			}
			else
			{
				hRes = MAPI_E_UNCONFIGURED;
			}
		}

		lpProfProps[CLIENT_MDB_PROVIDER].ulPropTag = sptClientProps.aulPropTag[CLIENT_MDB_PROVIDER];
		lpProfProps[CLIENT_MDB_PROVIDER].Value.bin.cb = sizeof(MAPIUID);
		lpProfProps[CLIENT_MDB_PROVIDER].Value.bin.lpb = (LPBYTE)&g_muidProvPrvNST;
	}

	if (SUCCEEDED(hRes))
	{
		hRes = lpProfSect->SetProps(ulProfProps,lpProfProps,NULL);

		if (SUCCEEDED(hRes))
		{
			hRes = lpProfSect->SaveChanges(KEEP_OPEN_READWRITE);
		}
	}

	if (lpszNSTFilePath) delete lpszNSTFilePath;
	MyFreeBuffer(lpProfProps);

	return hRes;
}

HRESULT GetSourceKey(SKEY* lpSKEY, BYTE docId, BYTE versionId)
{
	memset((void *) lpSKEY, 0x0, sizeof(SKEY));
	memcpy((void *) lpSKEY, (void *) &g_muidProvPrvNST, sizeof(GUID));
	lpSKEY->globcnt[0] = docId;
	lpSKEY->globcnt[1] = versionId;

	return S_OK;
}

STDMETHODIMP CreateMsgEntryID(BYTE   MsgID,
							  ULONG  cbParentEntryID,
							  FEID*  pParentEntryID,
							  MEID** ppMsgEID,
							  ULONG* pcbMsgEID)
{
    Log(true,"CreateMsgEntryID\n");
    HRESULT hRes = S_OK;
    MEID* pMsgEID = NULL;

    *ppMsgEID = 0;
    *pcbMsgEID = NULL;
	hRes = MyAllocateBuffer(sizeof(MEID), (LPVOID *)&pMsgEID);

	if (pMsgEID)
	{
		memset(pMsgEID, 0x0, sizeof(MEID));
		SKEY sKey = {0};

		GetSourceKey(&sKey,7,MsgID);

		memcpy((void *)(pMsgEID) , (void *) pParentEntryID, cbParentEntryID);
		// LTIDMsg
		memcpy((void *)&pMsgEID->ltidMsg, (void *)&sKey, sizeof(sKey));

		*ppMsgEID = pMsgEID;
		*pcbMsgEID = sizeof(MEID);
		pMsgEID = NULL;
	}

    MyFreeBuffer(pMsgEID);
	Log(true,"CreateMsgEntryID returning 0x%08X\n",hRes);
    return hRes;
}

STDMETHODIMP CreateStoreEntryID(
								LPSTR lpstrPath,
								BYTE **ppStoreEID,
								ULONG *pcbStoreEID)
{
	Log(true,"CreateStoreEntryID\n");
	HRESULT hRes = S_OK;
	BYTE * pStoreEID = NULL;
	ULONG cbStoreEID= 0;

	cbStoreEID = 4 + 16 + 1 + (ULONG) strlen(lpstrPath) + 1;
	hRes = MyAllocateBuffer(cbStoreEID, (LPVOID *)&pStoreEID);

	if (SUCCEEDED(hRes))
	{
		memset((void *) pStoreEID, 0x0, cbStoreEID);
		memcpy((void *)(pStoreEID+4), (void *) &g_muidProvPrvNST, 16);
		hRes = StringCbCopy((char *)(pStoreEID + 21),cbStoreEID-21, lpstrPath);

		*ppStoreEID = pStoreEID;
		*pcbStoreEID = cbStoreEID;
		pStoreEID = NULL;
	}

	MyFreeBuffer(pStoreEID);
	Log(true,"CreateStoreEntryID returning 0x%08X\n",hRes);
	return hRes;
}

HRESULT WINAPI InitStoreProps (LPMAPISUP pSupObj,
							   LPPROFSECT pProfObj,
							   LPPROVIDERADMIN pAdminProvObj)
{
	Log(true,"InitStoreProps begin\n");
	if (!pSupObj || !pProfObj || !pAdminProvObj) return MAPI_E_INVALID_PARAMETER;

	HRESULT      hRes = S_OK;
	ULONG        i = 0;
	LPPROFSECT   pProfObjStore = NULL;
	SPropValue   rgProps[10] = {0};
	MAPIUID      uidResource = {0};
	MAPIUID      uidService = {0};
	LPSPropValue lprop = NULL;
	LPSPropValue lpropServiceUID = NULL;
	LPENTRYID    pWrapEID = NULL;
	ULONG        ulWrapEIDLen = 0;
	ULONG        ulVal = 0;
	LPVOID       pTemp = NULL;
	BYTE*        pStoreEID = NULL;
	ULONG        cbStoreEID= 0;

	enum
	{
		_tag_PR_PROFILE_OFFLINE_STORE_PATH = 0,
			_tag_PR_STORE_PROVIDERS,
			_tag_PR_ENTRYID,
			NUM_TAGA
	};

	const static SizedSPropTagArray(NUM_TAGA, sptTaga) =
	{
		NUM_TAGA,
		{
			PR_PROFILE_OFFLINE_STORE_PATH,
				PR_STORE_PROVIDERS,
				PR_ENTRYID
		}
	};

	SizedSPropTagArray (1, tagaServiceUID) =
	{
		1,
		{
			PR_SERVICE_UID
		}
	};

	hRes = pProfObj->GetProps((LPSPropTagArray) &sptTaga, 0, &ulVal, & lprop);
	if (FAILED(hRes)) return hRes;

	// opening profile section for message store
	if (lprop[_tag_PR_STORE_PROVIDERS].ulPropTag == PR_STORE_PROVIDERS &&
		lprop[_tag_PR_ENTRYID].ulPropTag != PR_ENTRYID) // no need to go further if there is an entry ID here
	{
		hRes = pAdminProvObj->OpenProfileSection((LPMAPIUID) lprop[_tag_PR_STORE_PROVIDERS].Value.bin.lpb,
			NULL, MAPI_MODIFY, &pProfObjStore);

		if (SUCCEEDED(hRes) && pProfObjStore)
		{
			hRes = pProfObjStore->GetProps((LPSPropTagArray) &tagaServiceUID, 0, &ulVal, &lpropServiceUID);

			if (lpropServiceUID->Value.bin.cb == sizeof(MAPIUID))
			{
				uidService = *(MAPIUID*)lpropServiceUID->Value.bin.lpb;
				hRes = CreateStoreEntryID(
					lprop[_tag_PR_PROFILE_OFFLINE_STORE_PATH].Value.lpszA,
					&pStoreEID,
					&cbStoreEID);

				if (SUCCEEDED(hRes) && pStoreEID)
				{
					pTemp = MS_DLL_NAME_STRING;

					hRes = WrapStoreEntryID(
						0,
						(LPTSTR)MS_DLL_NAME_STRING,
						cbStoreEID,
						(LPENTRYID)pStoreEID,
						&ulWrapEIDLen,
						&pWrapEID);
					if (S_OK == hRes)
					{
						hRes = pSupObj->NewUID (&uidResource);

						rgProps[i].ulPropTag = PR_ENTRYID;
						rgProps[i].Value.bin.cb = ulWrapEIDLen;
						rgProps[i++].Value.bin.lpb = (LPBYTE)pWrapEID;

						rgProps[i].ulPropTag = PR_STORE_ENTRYID;
						rgProps[i].Value.bin.cb = ulWrapEIDLen;
						rgProps[i++].Value.bin.lpb = (LPBYTE)pWrapEID;

						rgProps[i].ulPropTag = PR_RECORD_KEY;
						rgProps[i].Value.bin.cb = sizeof(MAPIUID);
						rgProps[i++].Value.bin.lpb = (LPBYTE)&uidService;

						rgProps[i].ulPropTag = PR_STORE_RECORD_KEY;
						rgProps[i].Value.bin.cb = sizeof(MAPIUID);
						rgProps[i++].Value.bin.lpb = (LPBYTE)&uidResource;

						rgProps[i].ulPropTag = PR_SEARCH_KEY;
						rgProps[i].Value.bin.cb = sizeof(MAPIUID);
						rgProps[i++].Value.bin.lpb = (LPBYTE)&uidResource;

						rgProps[i].ulPropTag = PR_PROFILE_SECURE_MAILBOX;
						rgProps[i].Value.bin.cb = sizeof(MAPIUID);
						rgProps[i++].Value.bin.lpb = (LPBYTE)&uidResource;

						rgProps[i].ulPropTag = PR_PROFILE_USER;
						rgProps[i++].Value.lpszA = PROFILE_USER_STRING;

						rgProps[i].ulPropTag = PR_PROFILE_CONFIG_FLAGS;
						rgProps[i++].Value.l = NULL;

						rgProps[i].ulPropTag = PR_MDB_PROVIDER;
						rgProps[i].Value.bin.cb = sizeof(MAPIUID);
						rgProps[i++].Value.bin.lpb = (LPBYTE)&g_muidProvPrvNST;

						rgProps[i].ulPropTag = PR_PST_PATH;
						rgProps[i++].Value.lpszA = lprop[_tag_PR_PROFILE_OFFLINE_STORE_PATH].Value.lpszA;

						// Store's profile
						hRes = pProfObjStore->SetProps (i, rgProps, NULL);
						// global profile
						hRes = pProfObj->SetProps (i, rgProps, NULL);
					}
				}
			}
		}
	}

	if (pProfObjStore) pProfObjStore->Release();
	MyFreeBuffer(pStoreEID);
	MyFreeBuffer(lprop);
	MyFreeBuffer(lpropServiceUID);
	MyFreeBuffer(pWrapEID);
	Log(true,"InitStoreProps end\n");
	return hRes;
} // InitStoreProps

// Reads the current OLFI from the message store and checks if it needs to be updated
// As new items are created, blocks of the dwAlloc range will be used up
// When dwAlloc runs out, the PST will move ltidNextAlloc and dwNextAlloc over
// to ltidAlloc and dwAlloc
// This implementation checks when dwAlloc gets small and creates a new ltidNextAlloc/dwNextAlloc
// if needed.
// This function should be called quite often.
HRESULT SetOLFIInOST(LPMDB lpMDB)
{
	Log(true,"SetOLFIInOST\n");
	if (!lpMDB) return MAPI_E_INVALID_PARAMETER;

	HRESULT			hRes = S_OK;
	OLFI			olfi = {0};
	SPropValue		sval	= {0};
	LPSPropValue	pval	= NULL;
	BOOL			fDirty	= FALSE;

	hRes = HrGetOneProp((LPMAPIPROP)lpMDB, PR_OST_OLFI, &pval);
	if ((S_OK == hRes) || (MAPI_E_NOT_FOUND == hRes))
	{
		if (hRes == S_OK)
		{
			olfi = *(POLFI)(pval->Value.bin.lpb);
		}
		else
		{
			memset(&olfi, 0, sizeof(OLFI));
			olfi.ulVersion = OLFI_VERSION;
		}

		// No primary range - create one
		if (FEqLtid(&olfi.ltidAlloc, &ltidZero))
		{
			hRes = CoCreateGuid(&olfi.ltidAlloc.guid);

			olfi.dwAlloc = GLOB_COUNT_RANGE;
			fDirty = TRUE;

			Log(true,"SetOLFIInOST: Updated primary range of EntryIDs");
		}

		// If a good portion of the primary range is used, allocate
		// a secondary range
		if (	(olfi.dwAlloc < (GLOB_COUNT_RANGE / 2))
			&&	(FEqLtid(&olfi.ltidNextAlloc, &ltidZero)))
		{
			hRes = CoCreateGuid(&olfi.ltidNextAlloc.guid);

			olfi.dwNextAlloc = GLOB_COUNT_RANGE;
			fDirty = TRUE;

			Log(true,"SetOLFIInOST: Updated secondary range of EntryIDs");
		}
		if (fDirty)
		{
			Log(true,"SetOLFIInOST writing new/updated OLFI\n",hRes);
			sval.ulPropTag		= PR_OST_OLFI;
			sval.Value.bin.cb	= sizeof(OLFI);
			sval.Value.bin.lpb	= (LPBYTE)&olfi;

			hRes = HrSetOneProp((LPMAPIPROP)lpMDB, &sval);
		}
		MAPIFreeBuffer(pval);
	}

	Log(true,"SetOLFIInOST returning 0x%08X\n",hRes);
	return hRes;
}

HRESULT STDAPICALLTYPE ServiceEntry (
									 HINSTANCE		/*hInstance*/,
									 LPMALLOC		lpMalloc,
									 LPMAPISUP		lpMAPISup,
									 ULONG			ulUIParam,
									 ULONG			ulFlags,
									 ULONG			ulContext,
									 ULONG			cValues,
									 LPSPropValue	lpProps,
									 LPPROVIDERADMIN lpProviderAdmin,
									 LPMAPIERROR FAR *lppMapiError)
{
	// In a real world implementation, an LPMAPIERROR structure should be returned on any error
	if (!lpMAPISup || !lpProviderAdmin) return MAPI_E_INVALID_PARAMETER;
	HRESULT hRes = S_OK;

	Log(true,"ServiceEntry function called\n");
	Log(true,"About to call NSTServiceEntry\n");

	// These contexts are not handled
	if (MSG_SERVICE_INSTALL 		== ulContext ||
		MSG_SERVICE_DELETE			== ulContext ||
		MSG_SERVICE_UNINSTALL		== ulContext ||
		MSG_SERVICE_PROVIDER_CREATE == ulContext ||
		MSG_SERVICE_PROVIDER_DELETE == ulContext)
	{
		// These contexts are not handled
		return MAPI_E_NO_SUPPORT;
	}

	// Get memory routines
	hRes = lpMAPISup->GetMemAllocRoutines(&g_lpAllocateBuffer,&g_lpAllocateMore,&g_lpFreeBuffer);

	LPTSTR lpszPath = GetPSTPath();
	HMODULE hm = NULL;
	if (lpszPath)
	{
		Log(true, "Got PST path %s\n", lpszPath);

		hm = LoadLibrary(lpszPath);
		delete[] lpszPath;
	}

	Log(true, "Got module 0x%08X\n", hm);
	if (!hm) return E_OUTOFMEMORY;

	LPMSGSERVICEENTRY pNSTServiceEntry = (LPMSGSERVICEENTRY)GetProcAddress(hm, "NSTServiceEntry");
	Log(true, "Got procaddress  0x%08X\n", pNSTServiceEntry);
	if (!pNSTServiceEntry) return E_OUTOFMEMORY;

	// Get profile section
	LPPROFSECT lpProfSect = NULL;
	hRes = lpProviderAdmin->OpenProfileSection((LPMAPIUID) NULL, NULL, MAPI_MODIFY, &lpProfSect);

	// Set passed in props into the profile
	if (lpProps && cValues)
	{
		hRes = lpProfSect->SetProps(cValues,lpProps,NULL);
	}

	// Make sure there is a store path set
	if (SUCCEEDED(hRes)) hRes = SetOfflineStoreProps(lpProfSect,ulUIParam);

	ULONG			ulProfProps = 0;
	LPSPropValue	lpProfProps = NULL;

	// Now see what props there are
	hRes = lpProfSect->GetProps((LPSPropTagArray)&sptClientProps,
								fMapiUnicode,
								&ulProfProps,
								&lpProfProps);
	if (SUCCEEDED(hRes))
	{
		CSupport * pMySup = NULL;
		pMySup = new CSupport(lpMAPISup, lpProfSect);
		if (!pMySup)
		{
			Log(true,"MSProviderInit: Failed to allocate new CSupport object\n");
			hRes = E_OUTOFMEMORY;
		}
		if (SUCCEEDED(hRes))
		{
			hRes = pNSTServiceEntry(
				hm,
				lpMalloc,
				pMySup,
				ulUIParam,
				ulFlags,
				ulContext,
				ulProfProps,
				lpProfProps,
				NULL, // pAdminProvObj, // Don't pass this when creating an NST
				lppMapiError);
			if (SUCCEEDED(hRes))
			{
				// finish setting up the profile
				hRes = InitStoreProps(
					lpMAPISup,
					lpProfSect,
					lpProviderAdmin);

				hRes = lpProfSect->SaveChanges(KEEP_OPEN_READWRITE);
			}
		}
	}

	Log(true, "Ran pNSTServiceEntry 0x%08X\n", hRes);

	MyFreeBuffer(lpProfProps);
	if (lpProfSect) lpProfSect->Release();

	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMSProvider::CMSProvider()
//
//    Parameters
//      hInst       Handle to instance of this MS DLL
//
//    Purpose
//      Constructor of the object. Parameters are passed to initialize the
//      data members with the appropiate values.
//
//    Return Value
//      None
//
CMSProvider::CMSProvider (LPMSPROVIDER pPSTMSfwd)
{
	Log(true,"In Constructor of CMSProvider\n");

	m_pPSTMS = pPSTMSfwd;
	m_cRef = 1;
} // CMSProvider::CMSProvider

///////////////////////////////////////////////////////////////////////////////
//    CMSProvider::~CMSProvider()
//
//    Parameters
//      None
//
//    Purpose
//      Release any memory used by the members of this object.
//
//    Return Value
//      None
//
CMSProvider::~CMSProvider()
{
	Log(true,"In Destructor of CMSProvider\n");
} // CMSProvider::~CMSProvider

///////////////////////////////////////////////////////////////////////////////
//    CMSProvider::QueryInterface()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Returns a pointer to a interface requested if the interface is
//      supported and implemented by this object. If it is not supported, it
//      returns NULL.
//
//    Return Value
//      S_OK            If successful. With the interface pointer in *ppvObj
//      E_NOINTERFACE   If interface requested is not supported by this object
//
STDMETHODIMP CMSProvider::QueryInterface(REFIID riid, LPVOID * ppvObj)
{
	if (!m_pPSTMS) return E_NOINTERFACE;
	*ppvObj = NULL;
	// If the interface requested is supported by this object, return a pointer
	// to the provider, with the reference count incremented by one.
	if (riid == IID_IMSProvider  || riid == IID_IUnknown)
	{
		*ppvObj = (LPVOID)this;
		// Increase usage count of this object
		AddRef();
		return S_OK;
	}
	// If the interface is not supported, maybe the PST does.
	return m_pPSTMS->QueryInterface(riid, ppvObj);
} // CMSProvider::QueryInterface

STDMETHODIMP_(ULONG) CMSProvider::AddRef()
{
	if (!m_pPSTMS) return NULL;
	m_cRef++;
	return m_pPSTMS->AddRef();
} // CMSProvider::AddRef

STDMETHODIMP_(ULONG) CMSProvider::Release()
{
	if (!m_pPSTMS) return NULL;
	Log(true,"CMSProvider::Release() called\n");
	ULONG ulRef =  m_pPSTMS->Release();

	m_cRef--;
	if (m_cRef == 0)
	{
		m_pPSTMS = NULL;
		delete this;
		return 0;
	}

	return ulRef;
} // CMSProvider::Release

///////////////////////////////////////////////////////////////////////////////
//    CMSProvider::Shutdown()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      To terminate all session objects attached to this provider object
//
//    Return Value
//      S_OK always.
//
STDMETHODIMP CMSProvider::Shutdown(ULONG * pulFlags)
{
	HRESULT hRes = S_OK;
	Log(true,"CMSProvider::Shutdown\n");
	hRes =m_pPSTMS->Shutdown(pulFlags);
	Log(true,"CMSProvider::Shutdown returned: 0x%08X\n", hRes);
	return hRes ;
} // CMSProvider::Shutdown

///////////////////////////////////////////////////////////////////////////////
//    CMSProvider::Logon()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      This method logs on the store by opening the message store database
//      and allocating the the IMsgStore and IMSLogon objects that are
//      returned to MAPI. During this function the provider gets the
//      configuration properties stored in its corresponding profile section.
//      This infomation include the location of the database file and other
//      options important at logon time.
//      If the client is trying to logon onto a message store database file
//      which has already been open then don't create a new object, but
//      simply return an AddRef()'ed pointer to the IMsgStore logon object
//      and a NULL IMSLogon object. These are the rules of MAPI.
//      If new objects are created they are added to the chain of open
//      message store objects in this IMSProvider.
//
//    Return Value
//      An HRESULT
STDMETHODIMP CMSProvider::Logon(
								LPMAPISUP	  pSupObj,
								ULONG		  ulUIParam,
								LPTSTR		  pszProfileName,
								ULONG		  cbEntryID,
								LPENTRYID	  pEntryID,
								ULONG		  ulFlags,
								LPCIID		  pInterface,
								ULONG * 	  pcbSpoolSecurity,
								LPBYTE *	  ppbSpoolSecurity,
								LPMAPIERROR * ppMAPIError,
								LPMSLOGON *   ppMSLogon,
								LPMDB * 	  ppMDB)
{
	HRESULT hRes = S_OK;
	LPMDB lpPSTMDB = NULL;
	CMsgStore* pWrappedMDB = NULL;

	Log(true,"CMSProvider::Logon Pst logon Called\n");

	LPPROFSECT lpProfSect = NULL;
	CSupport * pMySup = NULL;
	hRes = GetGlobalProfileObject(pSupObj,&lpProfSect);
	pMySup = new CSupport(pSupObj, lpProfSect);
	if (!pMySup)
	{
		Log(true,"CMSProvider::Logon: Failed to allocate new CSupport object\n");
		hRes = E_OUTOFMEMORY;
	}
	if (SUCCEEDED(hRes))
	{
		ulFlags = (ulFlags & ~MDB_OST_LOGON_ANSI) | MDB_OST_LOGON_UNICODE;

		hRes = m_pPSTMS->Logon(
			pMySup,
			ulUIParam,
			pszProfileName,
			cbEntryID,
			pEntryID,
			ulFlags,
			pInterface,
			pcbSpoolSecurity,
			ppbSpoolSecurity,
			ppMAPIError,
			ppMSLogon,
			&lpPSTMDB);
	}
	Log(true,"CMSProvider::Logon returned 0x%08X\n", hRes);

	// Set up the MDB to allow synchronization
	if (SUCCEEDED(hRes))
	{
		hRes = SetOLFIInOST(lpPSTMDB);
		Log(true,"SetOLFIInOST returned 0x%08X\n", hRes);
	}

	// Wrap the outgoing MDB.
	pWrappedMDB = new CMsgStore (lpPSTMDB);
	if (NULL == pWrappedMDB)
	{
		Log(true,"CMSProvider::Logon: Failed to allocate new CMsgStore object\n");
		hRes = E_OUTOFMEMORY;
	}
	// Copy pointer to the allocated object back into the return LPMDB object pointer

	*ppMDB = pWrappedMDB;

	if (lpProfSect) lpProfSect->Release();

	return hRes;
} // CMSProvider::Logon

///////////////////////////////////////////////////////////////////////////////
//    CMSProvider::SpoolerLogon()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Do the same as CMSProvider::Logon() but in the MAPI spooler context
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMSProvider::SpoolerLogon (LPMAPISUP	  pSupObj,
										ULONG		  ulUIParam,
										LPTSTR		  pszProfileName,
										ULONG		  cbEntryID,
										LPENTRYID	  pEntryID,
										ULONG		  ulFlags,
										LPCIID		  pInterface,
										ULONG		  cbSpoolSecurity,
										LPBYTE		  pbSpoolSecurity,
										LPMAPIERROR * ppMAPIError,
										LPMSLOGON *   ppMSLogon,
										LPMDB * 	  ppMDB)
{
	HRESULT hRes = S_OK;

	Log(true,"CMSProvider::SpoolerLogon\n");
	LPPROFSECT lpProfSect = NULL;
	CSupport * pMySup = NULL;
	hRes = GetGlobalProfileObject(pSupObj,&lpProfSect);
	pMySup = new CSupport(pSupObj, lpProfSect);
	if (!pMySup)
	{
		Log(true,"CMSProvider::SpoolerLogon: Failed to allocate new CSupport object\n");
		hRes = E_OUTOFMEMORY;
	}
	if (SUCCEEDED(hRes))
	{
		hRes = m_pPSTMS->SpoolerLogon(
			pMySup, // pSupObj,
			ulUIParam,
			pszProfileName,
			cbEntryID,
			pEntryID,
			ulFlags,
			pInterface,
			cbSpoolSecurity,
			pbSpoolSecurity,
			ppMAPIError,
			ppMSLogon,
			ppMDB);
	}
	Log(true,"CMSProvider::SpoolerLogon returned 0x%08X\n", hRes);

	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMSProvider::CompareStoreIDs()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Give two entry IDs, compare them to find out if they are the
//      same. (i.e. they refer to the same object)
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMSProvider::CompareStoreIDs (ULONG	 cbEntryID1,
										   LPENTRYID pEntryID1,
										   ULONG	 cbEntryID2,
										   LPENTRYID pEntryID2,
										   ULONG	 ulFlags,
										   ULONG *	 pulResult)
{
	HRESULT hRes = S_OK;

	Log(true,"CMSProvider::CompareStoreIDs\n");
	hRes = m_pPSTMS->CompareStoreIDs(cbEntryID1, pEntryID1, cbEntryID2, pEntryID2, ulFlags, pulResult);

	Log(true,"CMSProvider::CompareStoreIDs returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::CMsgStore()
//
//    Parameters
//       pPSTMsgStore : this is the PST's Message Store which is passed through the constructor and
//                      this is stored to further access the Store.
//
//    Purpose
//      Constructor of the object. Parameters are passed to initialize the
//      data members with the appropiate values.
//
//    Return Value
//      None
//
CMsgStore::CMsgStore (LPMDB pPSTMsgStore)
{
	Log(true,"In Constructor of CMsgStore\n");
	m_pPSTMsgStore = pPSTMsgStore;
	m_cRef = 1;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::~CMsgStore()
//
//    Parameters
//      None
//
//    Purpose
//      Destructor of CMsgStore.
//
//    Return Value
//      None
//
CMsgStore::~CMsgStore()
{
	Log(true,"In Destructor of CMsgStore\n");
	m_pPSTMsgStore = NULL;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::QueryInterface()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Returns a pointer to a interface requested if the interface is
//      supported and implemented by this object. If it is not supported, it
//      returns NULL
//
//    Return Value
//      S_OK            If successful. With the interface pointer in *ppvObj
//      E_NOINTERFACE   If interface requested is not supported by this object
//
STDMETHODIMP CMsgStore::QueryInterface (REFIID riid, LPVOID * ppvObj)
{
	*ppvObj = NULL;

	// If the interface requested is supported by this object, return a pointer
	// to the provider, with the reference count incremented by one.
	if (riid == IID_IMsgStore || riid == IID_IMAPIProp || riid == IID_IUnknown)
	{
		*ppvObj = (LPVOID)this;
		// Increase usage count of this object
		AddRef();
		{
			return S_OK;
		}
	}

	return m_pPSTMsgStore->QueryInterface(riid, ppvObj);
}

STDMETHODIMP_(ULONG) CMsgStore::AddRef ()
{
	if (!m_pPSTMsgStore) return NULL;
	m_cRef++;
	m_pPSTMsgStore->AddRef();
	return m_cRef;
}

STDMETHODIMP_(ULONG) CMsgStore::Release ()
{
	if (!m_pPSTMsgStore) return NULL;
	Log(true,"CMsgStore::Release() called\n");
	ULONG ulRef =  m_pPSTMsgStore->Release();

	m_cRef--;
	if (m_cRef == 0)
	{
		m_pPSTMsgStore = NULL;
		delete this;
		return 0;
	}

	return ulRef;
}

///////////////////////////////////////////////////////////////////////////////
// IMAPIProp virtual member functions implementation
//

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::GetLastError()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Returns a MAPIERROR structure with an error description string about
//      the last error that occurred in the object.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::GetLastError(HRESULT	   hError,
									 ULONG		   ulFlags,
									 LPMAPIERROR * ppMAPIError)
{
	Log(true,"CMsgStore::GetLastError\n");
	return m_pPSTMsgStore->GetLastError(
		hError,
		ulFlags,
		ppMAPIError);
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::SaveChanges()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Stub method. The store object itself is not transacted. Changes
//      occur immediately without the need for committing changes.
//
//    Return Value
//      S_OK always
//
STDMETHODIMP CMsgStore::SaveChanges(ULONG ulFlags)
{
	Log(true,"CMsgStore::SaveChanges(ulFlags = 0x%08X\n",ulFlags);
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->SaveChanges(ulFlags);

	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::GetProps()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Return the value of the properties the client especified in the
//      property tag array. If the tag array is NULL, the user wants all
//      the properties.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::GetProps (LPSPropTagArray pPropTagArray,
								  ULONG 		  ulFlags,
								  ULONG *		  pcValues,
								  LPSPropValue *  ppPropArray)
{
	Log(true,"CMsgStore::GetProps\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->GetProps(
		pPropTagArray,
		ulFlags,
		pcValues,
		ppPropArray);

	Log(true,"CMsgStore::GetProps returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::GetPropList()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Duplicate the tag array with ALL the available properties
//      on this object. Properties added with SetProps() will not be returned.
//      The must be explicitly requested in a GetProps() call.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::GetPropList (ULONG			   ulFlags,
									 LPSPropTagArray * ppAllTags)
{
	Log(true,"CMsgStore::GetPropList\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->GetPropList(
		ulFlags,
		ppAllTags);

	Log(true,"CMsgStore::GetPropList returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::OpenProperty()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Stub method. Opening properties directly on the
//      store object is not supported.
//
//    Return Value
//      MAPI_E_NO_SUPPORT always.
//
STDMETHODIMP CMsgStore::OpenProperty (ULONG 	  ulPropTag,
									  LPCIID	  piid,
									  ULONG 	  ulInterfaceOptions,
									  ULONG 	  ulFlags,
									  LPUNKNOWN * ppUnk)
{
	Log(true,"CMsgStore::OpenProperty\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->OpenProperty(
		ulPropTag,
		piid,
		ulInterfaceOptions,
		ulFlags,
		ppUnk);

	Log(true,"CMsgStore::OpenProperty returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::SetProps()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Modify the value of the properties in the object. On IMsgStore
//      objects, changes to the properties are committed immediately.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::SetProps (ULONG 				cValues,
								  LPSPropValue			pPropArray,
								  LPSPropProblemArray * ppProblems)
{
	Log(true,"CMsgStore::SetProps\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->SetProps(
		cValues,
		pPropArray,
		ppProblems);

	Log(true,"CMsgStore::SetProps returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::DeleteProps()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Stub method. Clients are not allowed to delete properties in the
//      IMsgStore object.
//
//    Return Value
//      MAPI_E_NO_SUPPORT always
//
STDMETHODIMP CMsgStore::DeleteProps (LPSPropTagArray	   pPropTagArray,
									 LPSPropProblemArray * ppProblems)
{
	Log(true,"CMsgStore::DeleteProps\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->DeleteProps(
		pPropTagArray,
		ppProblems);

	Log(true,"CMsgStore::DeleteProps returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::CopyTo()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      MAPI 1.0 does not require IMsgStore to support copying itself to
//      other store, so always return MAPI_E_NO_SUPPORT.
//
//    Return Value
//      MAPI_E_NO_SUPPORT always.
//
STDMETHODIMP CMsgStore::CopyTo (ULONG				  ciidExclude,
								LPCIID				  rgiidExclude,
								LPSPropTagArray 	  pExcludeProps,
								ULONG				  ulUIParam,
								LPMAPIPROGRESS		  pProgress,
								LPCIID				  pInterface,
								LPVOID				  pDestObj,
								ULONG				  ulFlags,
								LPSPropProblemArray * ppProblems)
{
	Log(true,"CMsgStore::CopyTo\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->CopyTo(
		ciidExclude,
		rgiidExclude,
		pExcludeProps,
		ulUIParam,
		pProgress,
		pInterface,
		pDestObj,
		ulFlags,
		ppProblems);

	Log(true,"CMsgStore::CopyTo returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::CopyProps()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Stub method. Copying the provider's properties onto
//      another object is not supported.
//
//    Return Value
//      MAPI_E_NO_SUPPORT always
//
STDMETHODIMP CMsgStore::CopyProps (LPSPropTagArray		 pIncludeProps,
								   ULONG				 ulUIParam,
								   LPMAPIPROGRESS		 pProgress,
								   LPCIID				 pInterface,
								   LPVOID				 pDestObj,
								   ULONG				 ulFlags,
								   LPSPropProblemArray * ppProblems)
{
	Log(true,"CMsgStore::CopyProps\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->CopyProps(
		pIncludeProps,
		ulUIParam,
		pProgress,
		pInterface,
		pDestObj,
		ulFlags,
		ppProblems);

	Log(true,"CMsgStore::CopyProps returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::GetNamesFromIDs()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Return the list of names for the identifier list supplied by the
//      caller. This implementation defers to the common IMAPIProp
//      implementation to do the actual work.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::GetNamesFromIDs (LPSPropTagArray * ppPropTags,
										 LPGUID 		   pPropSetGuid,
										 ULONG			   ulFlags,
										 ULONG *		   pcPropNames,
										 LPMAPINAMEID **   pppPropNames)
{
	Log(true,"CMsgStore::GetNamesFromIDs\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->GetNamesFromIDs(
		ppPropTags,
		pPropSetGuid,
		ulFlags,
		pcPropNames,
		pppPropNames);

	Log(true,"CMsgStore::GetNamesFromIDs returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::GetIDsFromNames()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Return a property identifier for each named listed by the client. This
//      IMsgStore implementation defers to the common IMAPIProp
//      implementation to do the actual work.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::GetIDsFromNames (ULONG			   cPropNames,
										 LPMAPINAMEID *    ppPropNames,
										 ULONG			   ulFlags,
										 LPSPropTagArray * ppPropTags)
{
	Log(true,"CMsgStore::GetIDsFromNames\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->GetIDsFromNames(
		cPropNames,
		ppPropNames,
		ulFlags,
		ppPropTags);

	Log(true,"CMsgStore::GetIDsFromNames returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
// IMsgStore virtual member functions implementation
//

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::Advise()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Add a notification node to the list of notification subscribers for
//      the object pointer in the entry ID passed in. In this implementation,
//      use the MAPI notification engine to call MAPI to add a
//      new notification subscriber to the object. The token returned from
//      MAPI is save in the internal list of subscribers along with the
//      notification mask. This is used when the provider's object wants to
//      send notification about them.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::Advise (ULONG			 cbEntryID,
								LPENTRYID		 pEntryID,
								ULONG			 ulEventMask,
								LPMAPIADVISESINK pAdviseSink,
								ULONG * 		 pulConnection)
{
	Log(true,"CMsgStore::Advise\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->Advise(
		cbEntryID,
		pEntryID,
		ulEventMask,
		pAdviseSink,
		pulConnection);

	Log(true,"CMsgStore::Advise returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::Unadvise()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Terminates the advise link for a particular object. The connection
//      number passed in is given to the MAPI notification engine so that it
//      removes the connection that matches it. If the MAPI successfully
//      removes the connection, remove it from the subscription list.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::Unadvise (ULONG ulConnection)
{
	Log(true,"CMsgStore::Unadvise\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->Unadvise(
		ulConnection);

	Log(true,"CMsgStore::Unadvise returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::CompareEntryIDs()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      This function compares the two entry ID structures passed in and if
//      they are pointing to the same object, meaning all the members of the
//      entry ID structure are identical, it return TRUE. Otherwise false.
//
//    Return Value
//      An HRESULT. The comparison result is returned in the
//      pulResult argument.
//
STDMETHODIMP CMsgStore::CompareEntryIDs (ULONG	   cbEntryID1,
										 LPENTRYID pEntryID1,
										 ULONG	   cbEntryID2,
										 LPENTRYID pEntryID2,
										 ULONG	   ulFlags,
										 ULONG *   pulResult)
{
	Log(true,"CMsgStore::CompareEntryIDs\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->CompareEntryIDs(
		cbEntryID1,
		pEntryID1,
		cbEntryID2,
		pEntryID2,
		ulFlags,
		pulResult);

	Log(true,"CMsgStore::CompareEntryIDs returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::OpenEntry()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      This method opens an object that exists in this message store. The object
//      could be a folder or a message in any subfolder.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::OpenEntry (ULONG	   cbEntryID,
								   LPENTRYID   pEntryID,
								   LPCIID	   pInterface,
								   ULONG	   ulFlags,
								   ULONG *	   pulObjType,
								   LPUNKNOWN * ppUnk)
{
	Log(true,"CMsgStore::OpenEntry\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->OpenEntry(
		cbEntryID,
		pEntryID,
		pInterface,
		ulFlags,
		pulObjType,
		ppUnk);

	Log(true,"CMsgStore::OpenEntry returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::SetReceiveFolder()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Associates a message class with a particular folder whose entry ID
//      is specified by the caller. The folder is used as the "Inbox" for
//      all messages with similar or identical message classes.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::SetReceiveFolder (LPTSTR	szMessageClass,
										  ULONG 	ulFlags,
										  ULONG 	cbEntryID,
										  LPENTRYID pEntryID)
{
	Log(true,"CMsgStore::SetReceiveFolder\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->SetReceiveFolder(
		szMessageClass,
		ulFlags,
		cbEntryID,
		pEntryID);

	Log(true,"CMsgStore::SetReceiveFolder returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::GetReceiveFolder()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Get the entry ID a receive folder for a particular message class.
//      This receive folder is used as the "inbox" folder for the particular
//      message class the caller is specifiying.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::GetReceiveFolder (LPTSTR	  szMessageClass,
										  ULONG 	  ulFlags,
										  ULONG *	  pcbEntryID,
										  LPENTRYID * ppEntryID,
										  LPTSTR *	  ppszExplicitClass)
{
	Log(true,"CMsgStore::GetReceiveFolder\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->GetReceiveFolder(
		szMessageClass,
		ulFlags,
		pcbEntryID,
		ppEntryID,
		ppszExplicitClass);

	Log(true,"CMsgStore::GetReceiveFolder returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::GetReceiveFolderTable()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Return the table of message class-to-folder mapping. This table is
//      used to decide into what folder an incoming message should be placed
//      based on its message class.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::GetReceiveFolderTable (ULONG		 ulFlags,
											   LPMAPITABLE * ppTable)
{
	Log(true,"CMsgStore::GetReceiveFolderTable\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->GetReceiveFolderTable(ulFlags, ppTable);

	Log(true,"CMsgStore::GetReceiveFolderTable returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::StoreLogoff()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      This function is used to cleanup and close up any resources owned by
//      the store since the client plans to no longer use the message store.
//      In this implementation, save the flags in the provider's object and return
//      immediately. There is no special cleanup to do.
//
//    Return Value
//      S_OK always
//
STDMETHODIMP CMsgStore::StoreLogoff (ULONG * pulFlags)
{
	Log(true,"CMsgStore::StoreLogoff\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->StoreLogoff(pulFlags);

	Log(true,"CMsgStore::StoreLogoff returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::AbortSubmit()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Stub method. In this implementation, cannot abort a submitted message.
//
//    Return Value
//      MAPI_E_UNABLE_TO_ABORT always
//
STDMETHODIMP CMsgStore::AbortSubmit (ULONG	   cbEntryID,
									 LPENTRYID pEntryID,
									 ULONG	   ulFlags)
{
	Log(true,"CMsgStore::AbortSubmit\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->AbortSubmit(cbEntryID, pEntryID, ulFlags);

	Log(true,"CMsgStore::AbortSubmit returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::GetOutgoingQueue()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Return an IMAPITable interface to the caller. This table represents
//      the outgoing queue of submitted messages. When a message is
//      submitted, a row is added to this table with the appropiate columns.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::GetOutgoingQueue (ULONG ulFlags, LPMAPITABLE * ppTable)
{
	Log(true,"CMsgStore::GetOutgoingQueue\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->GetOutgoingQueue(ulFlags, ppTable);

	Log(true,"CMsgStore::GetOutgoingQueue returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::SetLockState()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      This method is called by the spooler process to lock or unlock a
//      message for the submission process by the transports. While a
//      message is locked, the client processes cannot access this message.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::SetLockState (LPMESSAGE pMessageObj,
									  ULONG 	ulLockState)
{
	Log(true,"CMsgStore::SetLockState\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->SetLockState(pMessageObj,ulLockState);

	Log(true,"CMsgStore::SetLockState returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::FinishedMsg()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      This message is called by the spooler to let the message store do
//      the final processing after the message has been completely sent.
//      During this function, the message fixes the PR_MESSAGE_FLAGS property
//      on the submitted message and lets MAPI do final processing (i.e.
//      Delete the message) before the message is moved out of the outgoing
//      queue table.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::FinishedMsg (ULONG	   ulFlags,
									 ULONG	   cbEntryID,
									 LPENTRYID pEntryID)
{
	Log(true,"CMsgStore::FinishedMsg\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->FinishedMsg(ulFlags, cbEntryID, pEntryID);

	Log(true,"CMsgStore::FinishedMsg returned 0x%08X\n", hRes);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//    CMsgStore::NotifyNewMail()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//      Entry point for the spooler to call into the store and tell it that a
//      new message was placed in a folder. With this information, the
//      spooler-side message store notifies the client-side message store,
//      and it in turn sends a notification to the client.
//
//    Return Value
//      An HRESULT
//
STDMETHODIMP CMsgStore::NotifyNewMail (LPNOTIFICATION pNotification)
{
	Log(true,"CMsgStore::NotifyNewMail\n");
	HRESULT hRes = S_OK;

	hRes = m_pPSTMsgStore->NotifyNewMail(pNotification);

	Log(true,"CMsgStore::NotifyNewMail returned 0x%08X\n", hRes);
	return hRes;
}

HRESULT SetTopSourceKey(LPMDB lpMDB)
{
	Log(true,"SetTopSourceKey\n");
	if (!lpMDB) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	SPropValue propVal[1];
	SKEY mySourceKey;
	GetSourceKey(&mySourceKey, 0, 0);

	propVal[0].ulPropTag = PR_SOURCE_KEY;
	propVal[0].Value.bin.cb = sizeof(SKEY);
	propVal[0].Value.bin.lpb = (LPBYTE ) &mySourceKey;

	hRes = lpMDB->SetProps(1, propVal, NULL);

	Log(true,"SetTopSourceKey returning 0x%08X\n",hRes);
	return hRes;

}

HRESULT STDAPICALLTYPE HrSyncEnd(LPOSTX pOstx, HRESULT hResIn)
{
	HRESULT hRes = S_OK;
	Log(true,"HrSyncEnd(0x%08X)\n",hResIn);

	if (!pOstx) return MAPI_E_INVALID_PARAMETER;

	hRes = pOstx->SetSyncResult(hResIn);
	if (SUCCEEDED(hRes))
	{
		hRes = pOstx->SyncEnd();
		if (FAILED(hRes))
		{
			Log(true,"SyncEnd() returning 0x%08X\n",hRes);
		}
	}
	else
	{
		Log(true,"SetSyncResult(0x%08X) returning 0x%08X\n",hResIn,hRes);
	}

	if (FAILED(hRes)) return hRes;
	else return hResIn;
}

// Download means bring content from the backend server into the PST
HRESULT STDAPICALLTYPE SyncDownloadFolders(LPOSTX pOstx)
{
	Log(true,"SyncDownloadFolders begin\n");
	HRESULT        hRes = S_OK;
	struct SYNC*   pSync = NULL;
	struct DNHIER* pDownHierSync = NULL;
	PXIHC          pxihc = NULL;
	SPropValue     pVals[7];

	SKEY sk1, sk2, sk3, sk4;

	GetSourceKey(&sk1, 1, 0);
	GetSourceKey(&sk2, 0, 0);
	GetSourceKey(&sk3, 2, 0);
	GetSourceKey(&sk4, 1, 1);

	pVals[0].ulPropTag = PR_SOURCE_KEY;
	pVals[0].Value.bin.cb = sizeof(SKEY);
	pVals[0].Value.bin.lpb = (LPBYTE)&sk1;

	pVals[1].ulPropTag = PR_PARENT_SOURCE_KEY;
	pVals[1].Value.bin.cb = sizeof(SKEY);
	pVals[1].Value.bin.lpb = (LPBYTE)&sk2;

	pVals[2].ulPropTag = PR_CHANGE_KEY;
	pVals[2].Value.bin.cb = sizeof(SKEY);
	pVals[2].Value.bin.lpb = (LPBYTE)&sk3;

	pVals[3].ulPropTag = PR_DISPLAY_NAME_W;
	pVals[3].Value.lpszW = L"My Synched Folder";

	pVals[4].ulPropTag = PR_COMMENT_W;
	pVals[4].Value.lpszW = L"Folder synched in with Replication API";

	pVals[5].ulPropTag = PR_PREDECESSOR_CHANGE_LIST ;
	pVals[5].Value.bin.cb = 0;
	pVals[5].Value.bin.lpb = 0;

	pVals[6].ulPropTag = PR_RECORD_KEY;
	pVals[6].Value.bin.cb = sizeof(SKEY);
	pVals[6].Value.bin.lpb = (LPBYTE)&sk4;

	hRes = pOstx->InitSync(SYNC_DOWNLOAD_HIERARCHY);
	Log(true,"InitSync(SYNC_DOWNLOAD_HIERARCHY) returning 0x%08X\n",hRes);
	if (SUCCEEDED(hRes))
	{
		hRes = pOstx->SyncBeg(LR_SYNC, (LPVOID *)&pSync );
		Log(true,"SyncBeg(LR_SYNC, (LPVOID *)&pSync ) returning 0x%08X\n",hRes);
		if (SUCCEEDED(hRes) && pSync)
		{
			pSync->ulFlags = UPS_DNLOAD_ONLY;

			hRes = pOstx->SyncBeg(LR_SYNC_DOWNLOAD_HIERARCHY , (LPVOID *)&pDownHierSync);
			Log(true,"SyncBeg(LR_SYNC_DOWNLOAD_HIERARCHY , (LPVOID *)&pDownHierSync) returning 0x%08X\n",hRes);
			if (SUCCEEDED(hRes) && pDownHierSync && pDownHierSync->pxihc)
			{
				pxihc = pDownHierSync->pxihc;
				hRes = pxihc->Config(NULL, NULL);
				Log(true,"pxihc->Config returning 0x%08X\n",hRes);
				if (SUCCEEDED(hRes))
				{
					// This is where the PST is notified that there are changes for it
					// For demonstration purposes, the change is just a single new folder
					hRes = pxihc->ImportFolderChange(sizeof(pVals)/sizeof(SPropValue), pVals);
					Log(true,"pxihc->ImportFolderChange returning 0x%08X\n",hRes);
				}
			}
			hRes = HrSyncEnd(pOstx,hRes);
		}
		hRes = HrSyncEnd(pOstx,hRes);
	}
	Log(true,"SyncDownloadFolders returning 0x%08X\n",hRes);
	return hRes;
} // SyncDownloadFolders

// Upload means push content from the PST into the backend server
HRESULT STDAPICALLTYPE SyncUploadFolders(LPOSTX pOstx)
{
	Log(true,"SyncUploadFolders begin\n");
	HRESULT        hRes = S_OK;
	struct SYNC*   pSync = NULL;
	struct UPHIER* pUpHierSync = NULL;
	struct UPFLD*  pUpFolderSync = NULL;

	hRes = pOstx->InitSync(SYNC_UPLOAD_HIERARCHY);
	Log(true,"InitSync(SYNC_UPLOAD_HIERARCHY) returning 0x%08X\n",hRes);
	if (SUCCEEDED(hRes))
	{
		hRes = pOstx->SyncBeg(LR_SYNC, (LPVOID *)&pSync);
		Log(true,"SyncBeg(LR_SYNC, (LPVOID *)&pSync) returning 0x%08X\n",hRes);
		if (SUCCEEDED(hRes) && pSync)
		{
			pSync->ulFlags = UPS_UPLOAD_ONLY;

			hRes = pOstx->SyncBeg(LR_SYNC_UPLOAD_HIERARCHY, (LPVOID *)&pUpHierSync);
			Log(true,"SyncBeg(LR_SYNC_UPLOAD_HIERARCHY, (LPVOID *)&pUpHierSync ) returning 0x%08X\n",hRes);
			if (SUCCEEDED(hRes) && pUpHierSync)
			{
				hRes = pOstx->SyncBeg(LR_SYNC_UPLOAD_FOLDER, (LPVOID *)&pUpFolderSync);
				Log(true,"SyncBeg(LR_SYNC_UPLOAD_FOLDER, (LPVOID *)&pUpFolderSync) returning 0x%08X\n",hRes);

				if (SUCCEEDED(hRes) && pUpFolderSync)
				{

				}

				hRes = HrSyncEnd(pOstx,hRes);
			}
			hRes = HrSyncEnd(pOstx,hRes);
		}
		hRes = HrSyncEnd(pOstx,hRes);
	}
	Log(true,"SyncUploadFolders returning 0x%08X\n",hRes);
	return hRes;
} // SyncUploadFolders

// Dummy function to sync a single message into a folder
HRESULT SyncOneContent(LPMDB lpMDB, BYTE bFolderNum,DNTBL *pDownTblSync)
{
	if (!pDownTblSync) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	Log(true,"pDownTblSync->pszName = \"%S\"\n",pDownTblSync->pszName);

	LPMESSAGE	lpMessage = NULL;
	MEID*		lpmeid = NULL;
	ULONG		cbmeid = NULL;
	PXICC		pxicc = pDownTblSync->pxicc;
	SPropValue	pVals[3];
	SKEY		sk1 = {0};

	hRes = pxicc->Config(NULL, NULL);
	Log(true,"pxicc->Config(NULL, NULL) returning 0x%08X\n",hRes);
	if (FAILED(hRes)) return hRes;

	GetSourceKey(&sk1, 0, 0);

	// Not strictly needed because the pszName is passed, but other
	// information needed from the folder can be accessed by doing something similar
	{
		ULONG		ulObjType = NULL;
		LPMESSAGE	lpMessage = NULL;

		hRes = lpMDB->OpenEntry(
			sizeof(FEID),
			(LPENTRYID)&pDownTblSync->feid,
			NULL,
			NULL,
			&ulObjType,
			(LPUNKNOWN*) &lpMessage);
		Log(true,"lpMDB->OpenEntry returning 0x%08X\n",hRes);

		if (lpMessage)
		{
			LPSPropValue lpProp = NULL;
			hRes = HrGetOneProp(lpMessage,PR_DISPLAY_NAME_A,&lpProp);
			Log(true,"lpMDB->OpenEntry returning 0x%08X\n",hRes);

			if (SUCCEEDED(hRes) &&
				lpProp &&
				PR_DISPLAY_NAME_A == lpProp->ulPropTag &&
				NULL != lpProp->Value.lpszA)
			{
				Log(true,"Display Name \"%s\"\n",lpProp->Value.lpszA);
			}

			MyFreeBuffer(lpProp);
			lpMessage->Release();
		}

	}

	hRes = CreateMsgEntryID(bFolderNum,sizeof(FEID),&pDownTblSync->feid,&lpmeid,&cbmeid);
	Log(true,"CreateMsgEntryID returning 0x%08X\n",hRes);
	if (SUCCEEDED(hRes))
	{
		pVals[0].ulPropTag = PR_CHANGE_KEY;
		pVals[0].Value.bin.cb = sizeof(SKEY);
		pVals[0].Value.bin.lpb = (LPBYTE)&sk1;

		pVals[1].ulPropTag = PR_PREDECESSOR_CHANGE_LIST ;
		pVals[1].Value.bin.cb = 0;   // Just the first one
		pVals[1].Value.bin.lpb = NULL;

		pVals[2].ulPropTag = PR_ENTRYID;
		pVals[2].Value.bin.cb = cbmeid;
		pVals[2].Value.bin.lpb = (LPBYTE)lpmeid;

		hRes = pxicc->ImportMessageChange(sizeof(pVals)/sizeof(SPropValue), pVals, 0, &lpMessage);
		Log(true,"pxicc->ImportMessageChange returning 0x%08X\n",hRes);

		if (lpMessage)
		{
			SPropValue pValsToSet[5];

			pValsToSet[0].ulPropTag = PR_BODY_W;
			pValsToSet[0].Value.lpszW = L"Hello Content Body";

			pValsToSet[1].ulPropTag = PR_DISPLAY_NAME_W;
			pValsToSet[1].Value.lpszW = L"Hello Content Display Name";

			pValsToSet[2].ulPropTag = PR_COMMENT_W;
			pValsToSet[2].Value.lpszW = L"Hello Content Comment";

			pValsToSet[3].ulPropTag = PR_SUBJECT_W;
			pValsToSet[3].Value.lpszW = L"Hello Content Subject";

			// Make this a sent message
			pValsToSet[4].ulPropTag = PR_MESSAGE_FLAGS;
			pValsToSet[4].Value.l = MSGFLAG_UNMODIFIED;

			hRes = lpMessage->SetProps(sizeof(pValsToSet)/sizeof(SPropValue),pValsToSet,NULL);
			Log(true,"lpMessage->SetProps returning 0x%08X\n",hRes);

			hRes = lpMessage->SaveChanges(0);
			Log(true,"lpMessage->SaveChanges returning 0x%08X\n",hRes);
			lpMessage->Release();
		}
	}

	if (lpmeid) MyFreeBuffer(lpmeid);

	return hRes;
}

// Download means bring content from the backend server into the PST
HRESULT STDAPICALLTYPE SyncDownloadContent(LPMDB lpMDB, LPOSTX pOstx)
{
	HRESULT hRes = S_OK;
	struct SYNC *pSync = NULL;
	struct SYNCCONT *pSyncContents = NULL;
	int i = 0;

	hRes = pOstx->InitSync(SYNC_DOWNLOAD_CONTENTS);
	Log(true,"InitSync(SYNC_DOWNLOAD_CONTENTS) returning 0x%08X\n",hRes);
	if (SUCCEEDED(hRes))
	{
		hRes = pOstx->SyncBeg(LR_SYNC, (LPVOID *)&pSync);
		Log(true,"SyncBeg(LR_SYNC, (LPVOID *)&pSync ) returning 0x%08X\n",hRes);
		if (SUCCEEDED(hRes))
		{
			pSync->ulFlags = UPS_DNLOAD_ONLY;

			hRes = pOstx->SyncBeg(LR_SYNC_CONTENTS , (LPVOID *)&pSyncContents);
			Log(true,"SyncBeg(LR_SYNC_CONTENTS , (LPVOID *)&pSyncContents ) returning 0x%08X\n",hRes);
			if (SUCCEEDED(hRes))
			{
				for (i = 0; i < (int)pSyncContents->cEnt; i++)
				{
					pSyncContents->iEnt = i;

					// The Replication API assumes that an upload sync will occur prior to a download sync
					// Even if there is no plan to do anything during the upload, this is still needed
					// to fill out the data structures
					// If this is removed, the pszName and feid won't be any good
					struct UPTBL *pUpTblSync = NULL;
					hRes = pOstx->SyncBeg(LR_SYNC_UPLOAD_TABLE , (LPVOID *)&pUpTblSync);
					Log(true,"SyncBeg(LR_SYNC_UPLOAD_TABLE , (LPVOID *)&pUpTblSync ) returning 0x%08X\n",hRes);
					hRes = HrSyncEnd(pOstx, hRes);

					if (SUCCEEDED(hRes))
					{
						struct DNTBL *pDownTblSync = NULL;
						hRes = pOstx->SyncBeg(LR_SYNC_DOWNLOAD_TABLE , (LPVOID *)&pDownTblSync);
						Log(true,"SyncBeg(LR_SYNC_DOWNLOAD_TABLE , (LPVOID *)&pDownTblSync ) returning 0x%08X\n",hRes);
						if (SUCCEEDED(hRes))
						{
							hRes = SyncOneContent(lpMDB, (BYTE)i,pDownTblSync);
							Log(true,"SyncOneContent(pDownTblSync) returning 0x%08X\n",hRes);
						}
						hRes = HrSyncEnd(pOstx, hRes);
					}

					// It's OK if syncing a single folder fails - keep trying the others.
					hRes = S_OK;
				}
			}
			hRes = HrSyncEnd(pOstx, hRes);
		}
		hRes = HrSyncEnd(pOstx, hRes);
	}

	Log(true,"SyncDownloadContent returning 0x%08X\n",hRes);
	return hRes;
}

// DoSync - Test function to demonstrate the Synchronization API
HRESULT STDAPICALLTYPE DoSync(LPMDB lpMDB)
{
	HRESULT hRes = S_OK;
	Log(true,"DoSync begin\n");
	if (!lpMDB)
	{
		Log(true,"DoSync not passed an MDB\n");
		MessageBox(GetActiveWindow(),"Didn't have a store to sync","Sync Error",MB_OK|MB_ICONERROR);
		return S_OK;
	}

	// Ensure there is a valid OLFI - this needs to be called often enough that it never runs out
	hRes = SetOLFIInOST(lpMDB);

	LPPSTX pPstx = NULL;
	LPOSTX pOstx = NULL;

	SetTopSourceKey(lpMDB);

	hRes = lpMDB->QueryInterface(IID_IPSTX, (void**)&pPstx);
	Log(true,"QueryInterface(IID_IPSTX) returning 0x%08X\n",hRes);

	if (SUCCEEDED(hRes) && pPstx)
	{
		hRes = pPstx->GetSyncObject(&pOstx);
		Log(true,"GetSyncObject returning 0x%08X\n",hRes);

		if (SUCCEEDED(hRes) && pOstx)
		{
			hRes = SyncDownloadFolders(pOstx);
			hRes = SyncUploadFolders(pOstx);
			hRes = SyncDownloadContent(lpMDB,pOstx);
		}
	}

	if (pOstx) pOstx->Release();
	if (pPstx) pPstx->Release();

	if (FAILED(hRes))
	{
		Log(true,"DoSync returning 0x%08X\n",hRes);
		MessageBox(GetActiveWindow(),"DoSync failed","Sync Error",MB_OK|MB_ICONERROR);
	}
	else
	{
		Log(true,"DoSync returning S_OK\n");
		MessageBox(GetActiveWindow(),"DoSync succeeded!","Sync Success",MB_OK);
	}
	return hRes;
} // DoSync

// TestDoSync - Test function to demonstrate the Synchronization API
// Logs on to the shared session
// Looks for the all wrapped PST stores
// Runs the test function DoSync on these stores
void _stdcall TestDoSync()
{
	HRESULT hRes = S_OK;
	LPMAPISESSION lpMAPISession = NULL;

	hRes = MAPIInitialize(NULL);
	if (SUCCEEDED(hRes))
	{
		hRes = MAPILogonEx(0, NULL, NULL,
			MAPI_EXTENDED | MAPI_ALLOW_OTHERS, // todo: check flags
			&lpMAPISession);
		if (SUCCEEDED(hRes) && lpMAPISession)
		{
			LPMAPITABLE lpStoresTable = NULL;
			hRes = lpMAPISession->GetMsgStoresTable(0, &lpStoresTable);
			if (SUCCEEDED(hRes) && lpStoresTable)
			{
				SRestriction sres = {0};
				SPropValue   spv = {0};
				LPSRowSet    pRow = NULL;
				enum {EID, NAME, NUM_COLS};
				static SizedSPropTagArray(NUM_COLS,sptCols) = {NUM_COLS, PR_ENTRYID, PR_DISPLAY_NAME};

				sres.rt = RES_PROPERTY;
				sres.res.resProperty.relop = RELOP_EQ;
				sres.res.resProperty.ulPropTag = PR_MDB_PROVIDER;
				sres.res.resProperty.lpProp = &spv;

				spv.ulPropTag = PR_MDB_PROVIDER;
				spv.Value.bin.cb  = sizeof(GUID);
				spv.Value.bin.lpb = (LPBYTE) &g_muidProvPrvNST;

				hRes = HrQueryAllRows(
					lpStoresTable,
					(LPSPropTagArray) &sptCols,
					&sres,
					NULL,
					0,
					&pRow);
				if (SUCCEEDED(hRes) && pRow)
				{
					ULONG ulRowNum = 0;
					for (ulRowNum = 0; ulRowNum < pRow->cRows; ulRowNum++)
					{
						LPMDB lpWrapMDB = NULL;
						hRes = lpMAPISession->OpenMsgStore(
							NULL,
							pRow->aRow[ulRowNum].lpProps[EID].Value.bin.cb,
							(LPENTRYID)pRow->aRow[ulRowNum].lpProps[EID].Value.bin.lpb,
							NULL,
							MAPI_BEST_ACCESS,
							&lpWrapMDB);
						if (SUCCEEDED(hRes) && lpWrapMDB)
						{
							(void) DoSync(lpWrapMDB);
						}
						if (lpWrapMDB) lpWrapMDB->Release();
					}
				}

				FreeProws(pRow);
			}
			if (lpStoresTable) lpStoresTable->Release();

		}

		if (lpMAPISession) lpMAPISession->Release();
		MAPIUninitialize();
	}
} // TestDoSync