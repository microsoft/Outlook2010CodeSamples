#include "stdafx.h"

CXPStatus::CXPStatus(CXPLogon* pParent, LPMAPISUP pMAPISup, ULONG cIdentityProps, LPSPropValue pIdentityProps,
	LPALLOCATEBUFFER pAllocBuffer, LPALLOCATEMORE pAllocMore, LPFREEBUFFER pFreeBuffer)
{
	Log(true, _T("CXPStatus object created at address 0x%p\n"), this);

	m_cRef = 1;
	InitializeCriticalSection(&m_csObj);

	m_pParentLogon = pParent;
	if (m_pParentLogon)
		m_pParentLogon->AddRef();

	m_pMAPISup = pMAPISup;
	if (m_pMAPISup)
		m_pMAPISup->AddRef();

	m_cIdentityProps = cIdentityProps;
	m_pIdentityProps = pIdentityProps;
	m_pAllocBuffer = pAllocBuffer;
	m_pAllocMore = pAllocMore;
	m_pFreeBuffer = pFreeBuffer;
}

CXPStatus::~CXPStatus()
{
	Log(true, _T("CXPStatus object destroyed at address 0x%p\n"), this);

	if (m_pMAPISup)
		m_pMAPISup->Release();

	if (m_pParentLogon)
		m_pParentLogon->Release();

	DeleteCriticalSection(&m_csObj);
}

STDMETHODIMP CXPStatus::QueryInterface(REFIID riid, LPVOID* ppvObj)
{
	if (!ppvObj)
		return E_INVALIDARG;

	*ppvObj = NULL;

	if (riid == IID_IMAPIStatus || riid == IID_IMAPIProp || riid == IID_IUnknown)
	{
		*ppvObj = (LPVOID)this;
		AddRef();
		return S_OK;
	}

	return E_NOINTERFACE;
}

// BEGIN IMAPIProp Implementation

/***********************************************************************************************
	CXPStatus::GetLastError

	*Not Supported*

***********************************************************************************************/

STDMETHODIMP CXPStatus::GetLastError(HRESULT /*hResult*/,
	ULONG /*ulFlags*/,
	LPMAPIERROR FAR* /*lppMAPIError*/)
{
	Log(true, _T("CXPStatus::GetLastError function called\n"));

	return MAPI_E_NO_SUPPORT;
}

/***********************************************************************************************
	CXPStatus::SaveChanges

	*Not Supported* (read-only)

***********************************************************************************************/

STDMETHODIMP CXPStatus::SaveChanges(ULONG /*ulFlags*/)
{
	Log(true, _T("CXPStatus::SaveChanges function called\n"));

	return MAPI_E_NO_ACCESS;
}

/***********************************************************************************************
	CXPStatus::GetProps

		- If the caller didn't pass an lpPropTagArray, just use the provider's default set of props
		- Any string props that are PT_UNSPECIFIED are set to the proper type, depending on
		if MAPI_UNICODE is passed in ulFlags
		- Allocate an array of SPropValues to return
		- Copy the values into the allocated SPropValues appropriately, converting strings
		if needed.

***********************************************************************************************/

STDMETHODIMP CXPStatus::GetProps(LPSPropTagArray lpPropTagArray,
	ULONG ulFlags,
	ULONG FAR* lpcValues,
	LPSPropValue FAR* lppPropArray)
{
	Log(true, _T("CXPStatus::GetProps function called\n"));

	if (!lpcValues || !lppPropArray)
		return MAPI_E_INVALID_PARAMETER;

	*lpcValues = 0;
	*lppPropArray = NULL;

	HRESULT hRes = S_OK;
	LPSPropTagArray pPropTags = lpPropTagArray;
	BOOL bGotErrors = false;

	ULONG ulStringPropType = PT_STRING8;

	if (MAPI_UNICODE & ulFlags)
		ulStringPropType = PT_UNICODE;

	if (!pPropTags)
	{
		pPropTags = (LPSPropTagArray)&sptStatusObj;
	}

	hRes = MyAllocateBuffer(pPropTags->cValues * sizeof(SPropValue), (LPVOID*)lppPropArray);
	if (SUCCEEDED(hRes) && *lppPropArray)
	{
		ULONG i = 0;

		for (i = 0; i < pPropTags->cValues; i++)
		{
			HRESULT hResProp = S_OK;
			ULONG ulPropType = PROP_TYPE(pPropTags->aulPropTag[i]);

			switch (PROP_ID(pPropTags->aulPropTag[i]))
			{
			case PROP_ID(PR_RESOURCE_METHODS):
				if (PT_LONG != ulPropType && PT_UNSPECIFIED != ulPropType)
				{
					hResProp = MAPI_E_UNEXPECTED_TYPE;
				}
				else
				{
					(*lppPropArray)[i].ulPropTag = PR_RESOURCE_METHODS;
					(*lppPropArray)[i].Value.ul = MRXP_RESOURCE_METHODS;
				}
				break;
			case PROP_ID(PR_PROVIDER_DISPLAY):
				if (PT_STRING8 != ulPropType && PT_UNICODE != ulPropType && PT_UNSPECIFIED != ulPropType)
				{
					hResProp = MAPI_E_UNEXPECTED_TYPE;
				}
				else
				{
					if (PT_UNSPECIFIED == ulPropType)
						(*lppPropArray)[i].ulPropTag = CHANGE_PROP_TYPE(pPropTags->aulPropTag[i], ulStringPropType);
					else
						(*lppPropArray)[i].ulPropTag = pPropTags->aulPropTag[i];

					hResProp = LoadStringAsType(PROP_TYPE((*lppPropArray)[i].ulPropTag), IDS_PROVIDER_DISPLAY,
						(LPVOID*)&((*lppPropArray)[i].Value.lpszA), *lppPropArray);
				}
				break;
			case PROP_ID(PR_DISPLAY_NAME):
				if (PT_STRING8 != ulPropType && PT_UNICODE != ulPropType && PT_UNSPECIFIED != ulPropType)
				{
					hResProp = MAPI_E_UNEXPECTED_TYPE;
				}
				else
				{
					if (PT_UNSPECIFIED == ulPropType)
						(*lppPropArray)[i].ulPropTag = CHANGE_PROP_TYPE(pPropTags->aulPropTag[i], ulStringPropType);
					else
						(*lppPropArray)[i].ulPropTag = pPropTags->aulPropTag[i];

					hResProp = GetConvertedStringProp((*lppPropArray)[i].ulPropTag, m_pIdentityProps[IDENTITY_EMAIL_ADDRESS],
						&((*lppPropArray)[i]), *lppPropArray);
				}
				break;
			case PROP_ID(PR_IDENTITY_DISPLAY):
				if (PT_STRING8 != ulPropType && PT_UNICODE != ulPropType && PT_UNSPECIFIED != ulPropType)
				{
					hResProp = MAPI_E_UNEXPECTED_TYPE;
				}
				else
				{
					if (PT_UNSPECIFIED == ulPropType)
						(*lppPropArray)[i].ulPropTag = CHANGE_PROP_TYPE(pPropTags->aulPropTag[i], ulStringPropType);
					else
						(*lppPropArray)[i].ulPropTag = pPropTags->aulPropTag[i];

					hResProp = GetConvertedStringProp((*lppPropArray)[i].ulPropTag, m_pIdentityProps[IDENTITY_DISPLAY_NAME],
						&((*lppPropArray)[i]), *lppPropArray);
				}
				break;
			case PROP_ID(PR_IDENTITY_ENTRYID):
				if (PT_BINARY != ulPropType && PT_UNSPECIFIED != ulPropType)
				{
					hResProp = MAPI_E_UNEXPECTED_TYPE;
				}
				else
				{
					(*lppPropArray)[i].ulPropTag = PR_IDENTITY_ENTRYID;
					hResProp = CopySBinary(&((*lppPropArray)[i].Value.bin),
						&(m_pIdentityProps[IDENTITY_ENTRYID].Value.bin), *lppPropArray,
						NULL, m_pAllocMore);
				}
				break;
			case PROP_ID(PR_IDENTITY_SEARCH_KEY):
				if (PT_BINARY != ulPropType && PT_UNSPECIFIED != ulPropType)
				{
					hResProp = MAPI_E_UNEXPECTED_TYPE;
				}
				else
				{
					(*lppPropArray)[i].ulPropTag = PR_IDENTITY_SEARCH_KEY;
					hResProp = CopySBinary(&((*lppPropArray)[i].Value.bin),
						&(m_pIdentityProps[IDENTITY_SEARCH_KEY].Value.bin), *lppPropArray,
						NULL, m_pAllocMore);
				}
				break;
			case PROP_ID(PR_STATUS_CODE):
				if (PT_LONG != ulPropType && PT_UNSPECIFIED != ulPropType)
				{
					hResProp = MAPI_E_UNEXPECTED_TYPE;
				}
				else
				{
					(*lppPropArray)[i].ulPropTag = PR_STATUS_CODE;
					(*lppPropArray)[i].Value.ul = m_pParentLogon->GetStatusBits();
				}
				break;
			case PROP_ID(PR_STATUS_STRING):
				if (PT_STRING8 != ulPropType && PT_UNICODE != ulPropType && PT_UNSPECIFIED != ulPropType)
				{
					hResProp = MAPI_E_UNEXPECTED_TYPE;
				}
				else
				{
					if (PT_UNSPECIFIED == ulPropType)
						(*lppPropArray)[i].ulPropTag = CHANGE_PROP_TYPE(pPropTags->aulPropTag[i], ulStringPropType);
					else
						(*lppPropArray)[i].ulPropTag = pPropTags->aulPropTag[i];

					hResProp = LoadStringAsType(PROP_TYPE((*lppPropArray)[i].ulPropTag), m_pParentLogon->GetStatusStringID(),
						(LPVOID*)&((*lppPropArray)[i].Value.lpszA), *lppPropArray);
				}
				break;
			case PROP_ID(PR_OBJECT_TYPE):
				if (PT_LONG != ulPropType && PT_UNSPECIFIED != ulPropType)
				{
					hResProp = MAPI_E_UNEXPECTED_TYPE;
				}
				else
				{
					(*lppPropArray)[i].ulPropTag = PR_OBJECT_TYPE;
					(*lppPropArray)[i].Value.ul = MAPI_STATUS;
				}
				break;
			default:
				hResProp = MAPI_E_NOT_FOUND;
			}

			if (FAILED(hResProp))
			{
				(*lppPropArray)[i].ulPropTag = CHANGE_PROP_TYPE(pPropTags->aulPropTag[i], PT_ERROR);
				(*lppPropArray)[i].Value.err = hResProp;
				bGotErrors = true;
			}
		}
	}

	if (SUCCEEDED(hRes) && bGotErrors)
		hRes = MAPI_W_ERRORS_RETURNED;

	return hRes;
}

/***********************************************************************************************
	CXPStatus::GetPropList

		- Allocate an array of proptags
		- Copy over the values from the provider's static array
		- Change all PT_UNSPECIFIED's to the proper string type indicated by MAPI_UNICODE in
		ulFlags

***********************************************************************************************/

STDMETHODIMP CXPStatus::GetPropList(ULONG ulFlags,
	LPSPropTagArray FAR* lppPropTagArray)
{
	Log(true, _T("CXPStatus::GetPropList function called\n"));

	if (!lppPropTagArray)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	ULONG ulStringPropType = PT_STRING8;

	if (MAPI_UNICODE & ulFlags)
		ulStringPropType = PT_UNICODE;

	*lppPropTagArray = NULL;

	hRes = MyAllocateBuffer(CbNewSPropTagArray(NUM_STATUS_OBJ_PROPS), (LPVOID*)lppPropTagArray);
	if (SUCCEEDED(hRes) && *lppPropTagArray)
	{
		(*lppPropTagArray)->cValues = NUM_STATUS_OBJ_PROPS;

		memcpy((*lppPropTagArray)->aulPropTag, (LPVOID)sptStatusObj.aulPropTag, sizeof(sptStatusObj.aulPropTag));

		ULONG i = 0;

		for (i = 0; i < NUM_STATUS_OBJ_PROPS; i++)
		{
			if (PROP_TYPE((*lppPropTagArray)->aulPropTag[i]) == PT_UNSPECIFIED)
			{
				(*lppPropTagArray)->aulPropTag[i] = CHANGE_PROP_TYPE((*lppPropTagArray)->aulPropTag[i],
					ulStringPropType);
			}
		}
	}

	return hRes;
}

/***********************************************************************************************
	CXPStatus::OpenProperty

	*Not Supported*

***********************************************************************************************/

STDMETHODIMP CXPStatus::OpenProperty(ULONG /*ulPropTag*/,
	LPCIID /*lpiid*/,
	ULONG /*ulInterfaceOptions*/,
	ULONG /*ulFlags*/,
	LPUNKNOWN FAR* /*lppUnk*/)
{
	Log(true, _T("CXPStatus::OpenProperty function called\n"));

	return MAPI_E_NO_SUPPORT;
}

/***********************************************************************************************
	CXPStatus::SetProps

	*Not Supported* (read-only)

***********************************************************************************************/

STDMETHODIMP CXPStatus::SetProps(ULONG /*cValues*/,
	LPSPropValue /*lpPropArray*/,
	LPSPropProblemArray FAR* /*lppProblems*/)
{
	Log(true, _T("CXPStatus::SetProps function called\n"));

	return MAPI_E_NO_ACCESS;
}

/***********************************************************************************************
	CXPStatus::DeleteProps

	*Not Supported* (read-only)

***********************************************************************************************/

STDMETHODIMP CXPStatus::DeleteProps(LPSPropTagArray /*lpPropTagArray*/,
	LPSPropProblemArray FAR* /*lppProblems*/)
{
	Log(true, _T("CXPStatus::DeleteProps function called\n"));

	return MAPI_E_NO_ACCESS;
}

/***********************************************************************************************
	CXPStatus::CopyTo

	*Not Supported*

***********************************************************************************************/

STDMETHODIMP CXPStatus::CopyTo(ULONG /*ciidExclude*/,
	LPCIID /*rgiidExclude*/,
	LPSPropTagArray /*lpExcludeProps*/,
	ULONG_PTR /*ulUIParam*/,
	LPMAPIPROGRESS /*lpProgress*/,
	LPCIID /*lpInterface*/,
	LPVOID /*lpDestObj*/,
	ULONG /*ulFlags*/,
	LPSPropProblemArray FAR* /*lppProblems*/)
{
	Log(true, _T("CXPStatus::CopyTo function called\n"));

	return MAPI_E_NO_SUPPORT;
}

/***********************************************************************************************
	CXPStatus::CopyProps

	*Not Supported*

***********************************************************************************************/

STDMETHODIMP CXPStatus::CopyProps(LPSPropTagArray /*lpIncludeProps*/,
	ULONG_PTR /*ulUIParam*/,
	LPMAPIPROGRESS /*lpProgress*/,
	LPCIID /*lpInterface*/,
	LPVOID /*lpDestObj*/,
	ULONG /*ulFlags*/,
	LPSPropProblemArray FAR* /*lppProblems*/)
{
	Log(true, _T("CXPStatus::CopyProps function called\n"));

	return MAPI_E_NO_SUPPORT;
}

/***********************************************************************************************
	CXPStatus::GetNamesFromIDs

	*Not Supported*

***********************************************************************************************/

STDMETHODIMP CXPStatus::GetNamesFromIDs(LPSPropTagArray FAR* /*lppPropTags*/,
	LPGUID /*lpPropSetGuid*/,
	ULONG /*ulFlags*/,
	ULONG FAR* /*lpcPropNames*/,
	LPMAPINAMEID FAR* FAR* /*lpppPropNames*/)
{
	Log(true, _T("CXPStatus::GetNamesFromIDs function called\n"));

	return MAPI_E_NO_SUPPORT;
}

/***********************************************************************************************
	CXPStatus::GetIDsFromNames

	*Not Supported*

***********************************************************************************************/

STDMETHODIMP CXPStatus::GetIDsFromNames(ULONG /*cPropNames*/,
	LPMAPINAMEID FAR* /*lppPropNames*/,
	ULONG /*ulFlags*/,
	LPSPropTagArray FAR* /*lppPropTags*/)
{
	Log(true, _T("CXPStatus::GetIDsFromNames function called\n"));

	return MAPI_E_NO_SUPPORT;
}

// END IMAPIProp Implementation

// BEGIN IMAPIStatus Implementation

/***********************************************************************************************
	CXPStatus::ValidateState

		- Check access to the user's inbox (UNC share) and set the provider's status bits appropriately

***********************************************************************************************/

STDMETHODIMP CXPStatus::ValidateState(ULONG_PTR /*ulUIParam*/,
	ULONG /*ulFlags*/)
{
	Log(true, _T("CXPStatus::ValidateState function called\n"));

	HRESULT hRes = S_OK;
	DWORD dwErr = 0;

	dwErr = CheckInboxAvailability(m_pIdentityProps[IDENTITY_EMAIL_ADDRESS].Value.lpszA);
	if (0 != dwErr)
	{
		Log(true, _T("CXPStatus::ValidateState: CheckInboxAvailability returned an error: 0x%08X\n"), dwErr);
		m_pParentLogon->RemoveStatusBits(STATUS_AVAILABLE);
		m_pParentLogon->AddStatusBits(STATUS_FAILURE);
	}
	else
	{
		m_pParentLogon->RemoveStatusBits(STATUS_FAILURE);
		m_pParentLogon->AddStatusBits(STATUS_AVAILABLE);
	}

	hRes = m_pParentLogon->UpdateStatusRow();
	if (FAILED(hRes))
	{
		Log(true, _T("CXPStatus::ValidateState: UpdateStatusRow returned an error: 0x%08X\n"), dwErr);
	}

	return hRes;
}

/***********************************************************************************************
	CXPStatus::SettingsDialog

***********************************************************************************************/

STDMETHODIMP CXPStatus::SettingsDialog(ULONG_PTR /*ulUIParam*/,
	ULONG /*ulFlags*/)
{
	Log(true, _T("CXPStatus::SettingsDialog function called\n"));

	HRESULT hRes = S_OK;

	return hRes;
}

/***********************************************************************************************
	CXPStatus::ChangePassword

	*Not Supported*

***********************************************************************************************/

STDMETHODIMP CXPStatus::ChangePassword(LPTSTR /*lpOldPass*/,
	LPTSTR /*lpNewPass*/,
	ULONG /*ulFlags*/)
{
	Log(true, _T("CXPStatus::ChangePassword function called\n"));

	return MAPI_E_NO_SUPPORT;
}

/***********************************************************************************************
	CXPStatus::FlushQueues

		Since this is already implemented in CXPLogon, just call that implementation.

***********************************************************************************************/

STDMETHODIMP CXPStatus::FlushQueues(ULONG_PTR /*ulUIParam*/,
	ULONG /*cbTargetTransport*/,
	LPENTRYID /*lpTargetTransport*/,
	ULONG ulFlags)
{
	Log(true, _T("CXPStatus::FlushQueues function called\n"));

	HRESULT hRes = S_OK;

	hRes = m_pParentLogon->FlushQueues(0, 0, NULL, ulFlags);
	if (FAILED(hRes))
		Log(true, _T("CXPStatus::FlushQueues: CXPLogon::FlushQueues returned an error: 0x%08X\n"), hRes);

	return hRes;
}

// END IMAPIStatus Implementation

/***********************************************************************************************
	CXPStatus::MyAllocateBuffer

	A wrapper around the provider's member pointer to MAPIAllocateBuffer. Makes sure that
	the function pointer is non-null before calling it.

***********************************************************************************************/

SCODE STDMETHODCALLTYPE CXPStatus::MyAllocateBuffer(ULONG			cbSize,
	LPVOID FAR *	lppBuffer)
{
	if (m_pAllocBuffer) return m_pAllocBuffer(cbSize, lppBuffer);

	Log(true, _T("CXPStatus::MyAllocateBuffer: m_pAllocBuffer not set!\n"));
	return MAPI_E_INVALID_PARAMETER;
}

/***********************************************************************************************
	CXPStatus::MyAllocateMore

	A wrapper around the provider's member pointer to MAPIAllocateMore. Makes sure that
	the function pointer is non-null before calling it.

***********************************************************************************************/

SCODE STDMETHODCALLTYPE CXPStatus::MyAllocateMore(ULONG			cbSize,
	LPVOID		lpObject,
	LPVOID FAR *	lppBuffer)
{
	if (m_pAllocMore) return m_pAllocMore(cbSize, lpObject, lppBuffer);

	Log(true, _T("CXPStatus::MyAllocMore: m_pAllocMore not set!\n"));
	return MAPI_E_INVALID_PARAMETER;
}

/***********************************************************************************************
	CXPStatus::MyFreeBuffer

	A wrapper around provider's member pointer to MAPIFreeBuffer. Makes sure that
	the function pointer is non-null before calling it.

***********************************************************************************************/

ULONG STDAPICALLTYPE CXPStatus::MyFreeBuffer(LPVOID			lpBuffer)
{
	if (m_pFreeBuffer) return m_pFreeBuffer(lpBuffer);

	Log(true, _T("CXPStatus::MyFreeBuffer: m_pFreeBuffer not set!\n"));
	return (ULONG)MAPI_E_INVALID_PARAMETER;
}

/***********************************************************************************************
	CXPStatus::CheckInboxAvailability

	Checks if the provider can open the passed in share

***********************************************************************************************/

DWORD CXPStatus::CheckInboxAvailability(LPSTR lpszInboxPath)
{
	Log(true, _T("CXPStatus::FlushQueues function called\n"));

	if (!lpszInboxPath)
		return (DWORD)-1;

	DWORD dwErr = 0;
	HANDLE hInbox = NULL;

	hInbox = CreateFileA(lpszInboxPath, GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
		NULL, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, NULL);
	if (INVALID_HANDLE_VALUE == hInbox)
	{
		dwErr = ::GetLastError();
	}

	CloseHandle(hInbox);

	return dwErr;
}

/***********************************************************************************************
	CXPStatus::GetConvertedStringProp

	Used by GetProps to handle allocation, copying, and converting of string props.
	Will take the passed in string value, compare it with what is requested, convert if
	necessary.

***********************************************************************************************/

HRESULT CXPStatus::GetConvertedStringProp(ULONG ulRequestedTag, SPropValue spvSource,
	LPSPropValue pspvDestination, LPVOID pParent)
{
	if (!ulRequestedTag || !spvSource.Value.lpszA || !pspvDestination || !pParent)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	if (PROP_TYPE(ulRequestedTag) == PROP_TYPE(spvSource.ulPropTag))
	{
		// No convert
		size_t cbSourceLen = 0;

		if (PT_STRING8 == PROP_TYPE(spvSource.ulPropTag))
		{
			hRes = StringCbLengthA(spvSource.Value.lpszA, STRSAFE_MAX_CCH * sizeof(char), &cbSourceLen);
			if (SUCCEEDED(hRes))
			{
				// NULL terminator
				cbSourceLen += sizeof(char);

				hRes = MyAllocateMore((ULONG)cbSourceLen, pParent, (LPVOID*)&(pspvDestination->Value.lpszA));
				if (SUCCEEDED(hRes) && pspvDestination->Value.lpszA)
				{
					hRes = StringCbCopyA(pspvDestination->Value.lpszA, cbSourceLen, spvSource.Value.lpszA);
				}
			}
		}
		else if (PT_UNICODE == PROP_TYPE(spvSource.ulPropTag))
		{
			hRes = StringCbLengthW(spvSource.Value.lpszW, STRSAFE_MAX_CCH * sizeof(WCHAR), &cbSourceLen);
			if (SUCCEEDED(hRes))
			{
				// NULL terminator
				cbSourceLen += sizeof(WCHAR);

				hRes = MyAllocateMore((ULONG)cbSourceLen, pParent, (LPVOID*)&(pspvDestination->Value.lpszW));
				if (SUCCEEDED(hRes) && pspvDestination->Value.lpszW)
				{
					hRes = StringCbCopyW(pspvDestination->Value.lpszW, cbSourceLen, spvSource.Value.lpszW);
				}
			}
		}
	}
	else if (PT_UNICODE == PROP_TYPE(ulRequestedTag) && PT_STRING8 == PROP_TYPE(spvSource.ulPropTag))
	{
		// ANSI -> UNICODE
		size_t cchStringA = 0;

		hRes = StringCchLengthA(spvSource.Value.lpszA, STRSAFE_MAX_CCH, &cchStringA);
		if (SUCCEEDED(hRes))
		{
			// Null terminator
			cchStringA++;

			int cchWideChars = 0;
			cchWideChars = MultiByteToWideChar(CP_ACP, 0, spvSource.Value.lpszA, (int)cchStringA,
				NULL, 0);
			if (cchWideChars > 0)
			{
				hRes = MyAllocateMore(cchWideChars * sizeof(WCHAR), pParent,
					(LPVOID*)&(pspvDestination->Value.lpszW));
				if (SUCCEEDED(hRes) && pspvDestination->Value.lpszW)
				{
					int iRet = 0;
					iRet = MultiByteToWideChar(CP_ACP, 0, spvSource.Value.lpszA, (int)cchStringA,
						pspvDestination->Value.lpszW, cchWideChars);
					if (iRet <= 0)
					{
						hRes = HRESULT_FROM_WIN32(::GetLastError());
					}
				}
			}
			else
			{
				hRes = HRESULT_FROM_WIN32(::GetLastError());
			}
		}
	}
	else if (PT_STRING8 == PROP_TYPE(ulRequestedTag) && PT_UNICODE == PROP_TYPE(spvSource.ulPropTag))
	{
		// UNICODE -> ANSI
		size_t cchStringW = 0;

		hRes = StringCchLengthW(spvSource.Value.lpszW, STRSAFE_MAX_CCH, &cchStringW);
		if (SUCCEEDED(hRes))
		{
			// Null terminator
			cchStringW++;

			int cchMBChars = 0;

			cchMBChars = WideCharToMultiByte(CP_ACP, 0, spvSource.Value.lpszW,
				(int)cchStringW, NULL, 0, NULL, NULL);

			if (cchMBChars > 0)
			{
				hRes = MyAllocateMore(cchMBChars * sizeof(char), pParent,
					(LPVOID*)&(pspvDestination->Value.lpszA));
				if (SUCCEEDED(hRes))
				{
					int iRet = 0;

					iRet = WideCharToMultiByte(CP_ACP, 0, spvSource.Value.lpszW,
						(int)cchStringW, pspvDestination->Value.lpszA, cchMBChars, NULL, NULL);
					if (iRet <= 0)
					{
						hRes = HRESULT_FROM_WIN32(::GetLastError());
					}
				}
			}
			else
			{
				hRes = HRESULT_FROM_WIN32(::GetLastError());
			}
		}
	}

	return hRes;
}

/***********************************************************************************************
	CXPStatus::LoadStringAsType

	Used by GetProps to load strings out of the string table as a particular type.

***********************************************************************************************/

HRESULT CXPStatus::LoadStringAsType(ULONG ulType, UINT ids, LPVOID* ppBuffer, LPVOID pParent)
{
	if (!ulType || !ids || !ppBuffer)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	int iRet = 0;

	if (PT_STRING8 == ulType)
	{
		hRes = MyAllocateMore(DEF_STATUS_ROW_STRING_LEN * sizeof(char), pParent, ppBuffer);
		if (SUCCEEDED(hRes) && *ppBuffer)
		{
			iRet = LoadStringA(m_pParentLogon->GetHINST(), ids,
				(LPSTR)*ppBuffer, DEF_STATUS_ROW_STRING_LEN);
			if (iRet <= 0)
			{
				*ppBuffer = MRXP_ERROR_STATUS_STRING_A;
			}
		}
	}
	else if (PT_UNICODE == ulType)
	{
		hRes = MyAllocateMore(DEF_STATUS_ROW_STRING_LEN * sizeof(WCHAR), pParent, ppBuffer);
		if (SUCCEEDED(hRes) && *ppBuffer)
		{
			iRet = LoadStringW(m_pParentLogon->GetHINST(), ids,
				(LPWSTR)*ppBuffer, DEF_STATUS_ROW_STRING_LEN);
			if (iRet <= 0)
			{
				*ppBuffer = MRXP_ERROR_STATUS_STRING_W;
			}
		}
	}

	return hRes;
}