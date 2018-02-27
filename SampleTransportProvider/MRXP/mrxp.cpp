#include "stdafx.h"

TCHAR szBlank[] = _T("");
// Static data for the message store provider configuration property sheets
TCHAR szFilter[] = _T("*"); // "*" to allow any character

// Control description structures
DTBLLABEL DtLabel = { sizeof(DTBLLABEL), fMapiUnicode };
DTBLPAGE DtPage = { sizeof(DTBLPAGE), fMapiUnicode };
DTBLEDIT DtCtlUserName = { sizeof(DTBLEDIT), fMapiUnicode, _MAX_PATH, PR_PROFILE_USER };
DTBLEDIT DtCtlInboxPath = { sizeof(DTBLEDIT), fMapiUnicode, _MAX_PATH, PR_PROFILE_MAILBOX };

// Description table for the controls in the first page
#define NUMBER_OF_CONTROLS  5
DTCTL ConfigPage[] =
{
	{DTCT_PAGE,    0,						  NULL, 0, NULL,     0,						&DtPage           },
	{DTCT_EDIT,    DT_REQUIRED | DT_EDITABLE, NULL, 0, szFilter, IDC_EDIT_DISPLAY_NAME, &DtCtlUserName    },
	{DTCT_EDIT,    DT_REQUIRED | DT_EDITABLE, NULL, 0, szFilter, IDC_EDIT_INBOX,		&DtCtlInboxPath   },
	{DTCT_LABEL,   0,						  NULL, 0, NULL,     IDC_LBL_DISPLAY_NAME,  &DtLabel          },
	{DTCT_LABEL,   0,						  NULL, 0, NULL,     IDC_LBL_INBOX,			&DtLabel          },
};

DTPAGE DtPropPage;

BOOL WINAPI DllMain(HINSTANCE /*hinstDLL*/, DWORD fdwReason, LPVOID /*lpvReserved*/)
{
	Log(true, _T("mrxp32.dll in DLLMain, fdwReason = 0x%08X\n"), fdwReason);

	switch (fdwReason)
	{
	case DLL_PROCESS_ATTACH:
		break;

	case DLL_THREAD_ATTACH:
		// Do thread-specific initialization.
		break;

	case DLL_THREAD_DETACH:
		// Do thread-specific cleanup.
		break;

	case DLL_PROCESS_DETACH:
		// Perform any necessary cleanup.
		break;
	}
	return true;
}

/***********************************************************************************************
	XPProviderInit

	- Check that ulMAPIVer is >= CURRENT_SPI_VERSION.  Return MAPI_E_VERSION if it isn't.
	- Create a new instance of the CXPProvider class and return it in the lppXPProvider
	parameter.
	- Return CURRENT_SPI_VERSION in the lpulProviderVer

	Refer to the MSDN documentation for more information.

***********************************************************************************************/

STDINITMETHODIMP XPProviderInit(
	HINSTANCE			hInstance,
	LPMALLOC			/*lpMalloc*/,
	LPALLOCATEBUFFER	lpAllocateBuffer,
	LPALLOCATEMORE 		lpAllocateMore,
	LPFREEBUFFER 		lpFreeBuffer,
	ULONG				/*ulFlags*/,
	ULONG				ulMAPIVer,
	ULONG FAR *			lpulProviderVer,
	LPXPPROVIDER FAR *	lppXPProvider)
{
	Log(true, _T("XPProviderInit function called\n"));

	if (!hInstance || !lpAllocateBuffer || !lpAllocateMore || !lpFreeBuffer
		|| !lpulProviderVer || !lppXPProvider)
		return MAPI_E_INVALID_PARAMETER;

	if (ulMAPIVer < CURRENT_SPI_VERSION)
	{
		Log(true, _T("XPProviderInit: Invalid version passed in ulMAPIVer: %d\n"), ulMAPIVer);
		return MAPI_E_VERSION;
	}

	CXPProvider* pProvider = NULL;
	pProvider = new CXPProvider(hInstance);
	if (!pProvider)
	{
		Log(true, _T("XPProviderInit: failed to create new CXPProvider\n"));
		return MAPI_E_NOT_ENOUGH_MEMORY;
	}

	*lpulProviderVer = CURRENT_SPI_VERSION;

	*lppXPProvider = (LPXPPROVIDER)pProvider;

	return S_OK;
}

/***********************************************************************************************
	ServiceEntry

		- Check what context the function was called in (check the ulContext parameter).
		For simplicity, just return S_OK for MSG_SERVICE_INSTALL, MSG_SERVICE_UNINSTALL, and
		MSG_SERVICE_DELETE, and return MAPI_E_NOT_SUPPORTED for MSG_SERVICE_PROVIDER_CREATE and
		MSG_SERVICE_PROVIDER_DELETE.
		- Get the MAPI alloc routines from the passed in IMAPISupport object
		- Open the provider's profile section.
		- If the function was passed in properties (in the lpProps parameter), go ahead and set them on the
		profile section.
		- Do a GetProps on the profile section for the set of properties that the provider requires for
		configuration.
		- If UI was specified (or if UI is allowed and the caller didn't pass in all of the provider's
		required properties), call DoConfigDialog
		- Validate the returned properties
		- Set them on the profile section if valid

		Refer to the MSDN documentation for more information.

***********************************************************************************************/

HRESULT STDAPICALLTYPE ServiceEntry(
	HINSTANCE		hInstance,
	LPMALLOC		/*lpMalloc*/,
	LPMAPISUP		lpMAPISup,
	ULONG			ulUIParam,
	ULONG			ulFlags,
	ULONG			ulContext,
	ULONG			cValues,
	LPSPropValue	lpProps,
	LPPROVIDERADMIN lpProviderAdmin,
	LPMAPIERROR FAR* /*lppMapiError*/)
{
	Log(true, _T("ServiceEntry function called\n"));

	if (!lpMAPISup || !lpProviderAdmin)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	// "Do-nothing" contexts
	if (MSG_SERVICE_INSTALL == ulContext ||
		MSG_SERVICE_UNINSTALL == ulContext ||
		MSG_SERVICE_DELETE == ulContext)
	{
		Log(true, _T("ServiceEntry: Do-nothing context: 0x%08X\n"), ulContext);
		return hRes;
	}
	else if (MSG_SERVICE_PROVIDER_CREATE == ulContext ||
		MSG_SERVICE_PROVIDER_DELETE == ulContext)
	{
		Log(true, _T("ServiceEntry: Unsupported context: 0x%08X\n"), ulContext);
		return MAPI_E_NO_SUPPORT;
	}

	Log(true, _T("ServiceEntry: Context: 0x%08X\n"), ulContext);

	// Get MAPI alloc routines
	LPALLOCATEBUFFER pAllocateBuffer = NULL;  // MAPIAllocateBuffer function
	LPALLOCATEMORE   pAllocateMore = NULL;  // MAPIAllocateMore function
	LPFREEBUFFER     pFreeBuffer = NULL;  // MAPIFreeBuffer function

	hRes = lpMAPISup->GetMemAllocRoutines(&pAllocateBuffer, &pAllocateMore, &pFreeBuffer);
	if (FAILED(hRes) || !pAllocateBuffer || !pAllocateMore || !pFreeBuffer)
	{
		Log(true, _T("ServiceEntry: GetMemAllocRoutines returned an error: 0x%08X\n"), hRes);
		return hRes;
	}

	// Get the profile section
	LPPROFSECT lpProfSect = NULL;

	hRes = OpenProviderProfileSection(lpProviderAdmin, &lpProfSect);
	if (SUCCEEDED(hRes) && lpProfSect)
	{
		// Set any props that have been passed in
		if (lpProps && cValues)
		{
			hRes = lpProfSect->SetProps(cValues, lpProps, NULL);
			if (FAILED(hRes))
			{
				Log(true, _T("ServiceEntry: SetProps on lpProfSect for passed in properties returned an error: 0x%08X\n"), hRes);
			}
		}

		if (SUCCEEDED(hRes))
		{
			ULONG ulProfProps = 0;
			LPSPropValue lpProfProps = NULL;

			hRes = lpProfSect->GetProps((LPSPropTagArray)&sptClientProps, 0,
				&ulProfProps, &lpProfProps);
			if (SUCCEEDED(hRes) && lpProfProps && ulProfProps)
			{
				HRESULT hResValidConfig = S_OK;

				hResValidConfig = ValidateConfigProps(ulProfProps, lpProfProps);
				if ((SERVICE_UI_ALWAYS & ulFlags) ||
					(SERVICE_UI_ALLOWED & ulFlags && FAILED(hResValidConfig)))
				{
					hRes = DoConfigDialog(hInstance, ulUIParam, &ulProfProps, &lpProfProps,
						(LPSPropTagArray)&sptClientProps, lpMAPISup, MSG_SERVICE_UI_READ_ONLY & ulFlags ? UI_READONLY : 0,
						pAllocateBuffer, pAllocateMore, pFreeBuffer);
					if (SUCCEEDED(hRes))
					{
						hRes = ValidateConfigProps(ulProfProps, lpProfProps);
						if (SUCCEEDED(hRes))
						{
							hRes = lpProfSect->SetProps(ulProfProps, lpProfProps, NULL);
							if (FAILED(hRes))
							{
								Log(true, _T("ServiceEntry: SetProps on lpProfSect for config page's properties returned an error: 0x%08X\n"), hRes);
							}
						}
						else
						{
							Log(true, _T("ServiceEntry: ValidateConfigProps returned an error: 0x%08X\n"), hRes);
						}
					}
					else
					{
						Log(true, _T("ServiceEntry: DoConfigDialog returned an error: 0x%08X\n"), hRes);
					}
				}

				else if (FAILED(hResValidConfig) && !(SERVICE_UI_ALLOWED & ulFlags))
				{
					Log(true, _T("ServiceEntry: UI was not allowed and required properties were not provided!\n"));
					hRes = MAPI_E_UNCONFIGURED;
				}

			}
			else
			{
				Log(true, _T("ServiceEntry: GetProps on lpProfSect for required properties returned an error: 0x%08X\n"), hRes);
			}

			pFreeBuffer(lpProfProps);
		}
	}
	else
	{
		Log(true, _T("ServiceEntry: OpenProfileSection returned an error: 0x%08X\n"), hRes);
	}

	if (lpProfSect)
		lpProfSect->Release();

	return hRes;
}

STDMETHODIMP PreprocessMessage(LPVOID					/*lpvSession*/,
	LPMESSAGE				/*lpMessage*/,
	LPADRBOOK				/*lpAdrBook*/,
	LPMAPIFOLDER				/*lpFolder*/,
	LPALLOCATEBUFFER			/*AllocateBuffer*/,
	LPALLOCATEMORE			/*AllocateMore*/,
	LPFREEBUFFER				/*FreeBuffer*/,
	ULONG FAR *				/*lpcOutbound*/,
	LPMESSAGE FAR * FAR *	/*lpppMessage*/,
	LPADRLIST FAR *			/*lppRecipList*/)
{
	Log(true, _T("PreprocessMessage function called\n"));
	return S_OK;
}

STDMETHODIMP RemovePreprocessInfo(LPMESSAGE /*lpMessage*/)
{
	Log(true, _T("RemovePreprocessInfo function called\n"));
	return S_OK;
}

STDMETHODIMP MergeWithMAPISVC()
{
	return HrSetProfileParameters(aMRXPServicesIni);
}

STDMETHODIMP RemoveFromMAPISVC()
{
	return HrSetProfileParameters(aREMOVE_MRXPServicesIni);
}

HRESULT WINAPI DoConfigDialog(HINSTANCE		hInstance,
	ULONG			ulUIParam,
	ULONG*			pulProfProps,
	LPSPropValue*	ppPropArray,
	LPSPropTagArray	pTags,
	LPMAPISUP		pSupObj,
	ULONG			ulFlags,
	LPALLOCATEBUFFER	pAllocateBuffer,
	LPALLOCATEMORE	pAllocateMore,
	LPFREEBUFFER		pFreeBuffer)
{
	auto hWnd = (HWND)(static_cast<ULONG_PTR>(ulUIParam));
	Log(true, _T("DoConfigDialog function called.\n"));

	if (!hInstance || !pulProfProps || !*pulProfProps || !ppPropArray || !pTags || !pSupObj)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	LPSPropValue pConfigProps = *ppPropArray;
	*ppPropArray = NULL;

	// Internally, DoConfigPropSheet treats the parameters passed there as an ANSI
	// string.  Therefore, force ANSI and then cast to LPTSTR
	char szDialogTitle[128] = { 0 };
	int iRet = 0;

	iRet = LoadStringA(hInstance, IDS_CONFIGSHEET_TITLE, szDialogTitle, CCH_A(szDialogTitle));
	if (!iRet)
		return MAPI_E_NOT_ENOUGH_MEMORY;

	LPPROPDATA pPropDataObj = NULL;

	hRes = CreateIProp(&IID_IMAPIPropData, pAllocateBuffer, pAllocateMore,
		pFreeBuffer, NULL, &pPropDataObj);
	if (SUCCEEDED(hRes) && pPropDataObj)
	{
		hRes = pPropDataObj->SetProps(*pulProfProps, pConfigProps, NULL);
		if (SUCCEEDED(hRes))
		{
			DtPropPage.cctl = NUMBER_OF_CONTROLS;
			DtPropPage.lpszComponent = szBlank;
			DtPropPage.lpszResourceName = MAKEINTRESOURCE(IDD_CONFIGPAGE);
			DtPropPage.lpctl = ConfigPage;

			if (ulFlags & UI_READONLY)
			{
				for (int i = 0; i < NUMBER_OF_CONTROLS; i++)
				{
					if (DtPropPage.lpctl[i].ulCtlFlags & DT_EDITABLE)
						DtPropPage.lpctl[i].ulCtlFlags &= ~DT_EDITABLE;
				}
			}

			LPMAPITABLE pTableObj = NULL;

			hRes = BuildDisplayTable(pAllocateBuffer, pAllocateMore, pFreeBuffer,
				NULL, hInstance, 1, &DtPropPage, fMapiUnicode, &pTableObj, NULL);

			if (SUCCEEDED(hRes) && pTableObj)
			{
				HRESULT hResValidConfig = MAPI_E_UNCONFIGURED;
				while (MAPI_E_UNCONFIGURED == hResValidConfig)
				{
					hResValidConfig = S_OK;

					hRes = pSupObj->DoConfigPropsheet(ulUIParam, 0, (LPTSTR)szDialogTitle,
						pTableObj, pPropDataObj, 0);
					if (SUCCEEDED(hRes))
					{
						LPSPropValue pNewProps = NULL;
						ULONG ulNewProps = 0;

						hRes = pPropDataObj->GetProps(pTags, fMapiUnicode,
							&ulNewProps, &pNewProps);
						if (SUCCEEDED(hRes) && pNewProps && ulNewProps)
						{
							hResValidConfig = ValidateConfigProps(ulNewProps, pNewProps);
							if (SUCCEEDED(hResValidConfig))
							{
								*ppPropArray = pNewProps;
								*pulProfProps = ulNewProps;
							}
							else
							{
								TCHAR szErrCaption[128] = { 0 };
								TCHAR szErrMsg[128] = { 0 };
								iRet = LoadString(hInstance, IDS_ERROR_CAPTION, szErrCaption, CCH(szErrCaption));
								iRet = LoadString(hInstance, IDS_INVALID_CONFIG_ERROR, szErrMsg, CCH(szErrMsg));
								MessageBox(hWnd, szErrMsg, szErrCaption, MB_OK);
							}
						}
						else
						{
							Log(true, _T("DoConfigDialog: GetProps for new values returned an error: 0x%08X\n"), hRes);
						}
					}
					else
					{
						Log(true, _T("DoConfigDialog: DoConfigPropSheet returned an error: 0x%08X\n"), hRes);
					}
				}
			}
			else
			{
				Log(true, _T("DoConfigDialog: BuildDisplayTable returned an error: 0x%08X\n"), hRes);
			}

			if (pTableObj)
				pTableObj->Release();
		}
		else
		{
			Log(true, _T("DoConfigDialog: SetProps on IPropData returned an error: 0x%08X\n"), hRes);
		}
	}
	else
	{
		Log(true, _T("DoConfigDialog: CreateIProp returned an error: 0x%08X\n"), hRes);
	}

	if (pPropDataObj)
		pPropDataObj->Release();

	pFreeBuffer(pConfigProps);

	return hRes;
}

HRESULT	ValidateConfigProps(ULONG ulProps, LPSPropValue pProps)
{
	Log(true, _T("ValidateConfigProps function called\n"));
	if (!ulProps || ulProps < 2 || !pProps)
		return MAPI_E_UNCONFIGURED;

	// Make sure display name is present
	LPSPropValue pVal = NULL;

	pVal = PpropFindProp(pProps, ulProps, PR_PROFILE_USER);
	if (!pVal || PT_ERROR == PROP_TYPE(pVal->ulPropTag))
	{
		Log(true, _T("ValidateConfigProps: PR_PROFILE_USER not present\n"));
		return MAPI_E_UNCONFIGURED;
	}

	// Make sure path to inbox is present and validate it

	pVal = PpropFindProp(pProps, ulProps, PR_PROFILE_MAILBOX);
	if (!pVal || PT_ERROR == PROP_TYPE(pVal->ulPropTag))
	{
		Log(true, _T("ValidateConfigProps: PR_PROFILE_MAILBOX not present\n"));
		return MAPI_E_UNCONFIGURED;
	}

	// Strip off trailing "\" if there is one to make the prop uniform
	size_t cchInboxLen = 0;

	HRESULT hRes = S_OK;
	hRes = StringCchLengthA(pVal->Value.lpszA, STRSAFE_MAX_CCH, &cchInboxLen);
	if (SUCCEEDED(hRes) && cchInboxLen)
	{
		if (pVal->Value.lpszA[cchInboxLen - 1] == _T('\\'))
			pVal->Value.lpszA[cchInboxLen - 1] = _T('\0');
	}

	// Check that the value of this property is a valid
	// path (UNC or local)
	HANDLE hInbox = NULL;

	// FILE_FLAG_BACKUP_SEMANTICS - use this flag when you open
	// a directory.
	hInbox = CreateFileA(pVal->Value.lpszA, GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
		NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, NULL);
	if (INVALID_HANDLE_VALUE == hInbox)
	{
		DWORD dwError = 0;
		dwError = GetLastError();

		Log(true, _T("Value for PR_PROFILE_MAILBOX: \"%hs\" is invalid. CreateFile returned 0x%08X\n"),
			pVal->Value.lpszA, dwError);

		return MAPI_E_UNCONFIGURED;
	}

	CloseHandle(hInbox);

	return S_OK;
}

HRESULT OpenProviderProfileSection(LPPROVIDERADMIN lpAdmin, LPPROFSECT* lppProfSect)
{
	Log(true, _T("OpenProviderProfileSection function called\n"));

	if (!lpAdmin || !lppProfSect)
		return MAPI_E_INVALID_PARAMETER;

	*lppProfSect = NULL;

	HRESULT hRes = S_OK;
	LPMAPITABLE lpProvTable = NULL;

	hRes = lpAdmin->GetProviderTable(0, &lpProvTable);
	if (SUCCEEDED(hRes) && lpProvTable)
	{
		SPropTagArray sptProvider = { 1, {PR_PROVIDER_UID} };

		hRes = lpProvTable->SetColumns(&sptProvider, 0);
		if (SUCCEEDED(hRes))
		{
			hRes = lpProvTable->SeekRow(BOOKMARK_BEGINNING, 0, NULL);
			if (SUCCEEDED(hRes))
			{
				LPSRowSet pRows = NULL;

				hRes = lpProvTable->QueryRows(1, 0, &pRows);
				if (SUCCEEDED(hRes) && pRows && pRows->cRows &&
					pRows->aRow[0].lpProps[0].ulPropTag == PR_PROVIDER_UID)
				{
					hRes = lpAdmin->OpenProfileSection((LPMAPIUID)pRows->aRow[0].lpProps[0].Value.bin.lpb,
						NULL, MAPI_MODIFY, lppProfSect);
					if (FAILED(hRes))
						Log(true, _T("OpenProviderProfileSection: OpenProfileSection returned an error: 0x%08X\n"), hRes);
				}
				else
				{
					Log(true, _T("OpenProviderProfileSection: QueryRows returned an error: 0x%08X\n"), hRes);
				}
			}
			else
			{
				Log(true, _T("OpenProviderProfileSection: SeekRow returned an error: 0x%08X\n"), hRes);
			}
		}
		else
		{
			Log(true, _T("OpenProviderProfileSection: SetColumns returned an error: 0x%08X\n"), hRes);
		}
	}
	else
	{
		Log(true, _T("OpenProviderProfileSection: GetProviderTable returned an error: 0x%08X\n"), hRes);
	}

	return hRes;
}