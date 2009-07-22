#include "stdafx.h"

CXPLogon::CXPLogon(CXPProvider* pProvider, LPMAPISUP pMAPISup, LPALLOCATEBUFFER pAllocBuffer,
		LPALLOCATEMORE pAllocMore, LPFREEBUFFER pFreeBuffer, ULONG ulProps, LPSPropValue pProps,
		ULONG ulFlags)
{
	Log(true, _T("CXPLogon object created at address 0x%08x\n"), ULONG(this));

	m_cRef = 1;
	InitializeCriticalSection(&m_csObj);

	m_pProvider = pProvider;
	if (m_pProvider)
		m_pProvider->AddRef();

	m_pMAPISup = pMAPISup;

	if (m_pMAPISup)
		m_pMAPISup->AddRef();

	m_pStatusObj = NULL;
	m_pAllocBuffer = pAllocBuffer;
	m_pAllocMore = pAllocMore;
	m_pFreeBuffer = pFreeBuffer;
	m_ulProps = ulProps;
	m_pProps = pProps;
	m_pIdentityProps = NULL;
	m_cchInboxPath = 0;
	m_ulLogonFlags = ulFlags;
	m_pNextLogon = NULL;
	m_ulStatusBits = 0;
	m_pszFileFilter = NULL;
}

CXPLogon::~CXPLogon()
{
	Log(true, _T("CXPLogon object destroyed at address 0x%08x\n"), ULONG(this));

	MyFreeBuffer(m_pszFileFilter);
	MyFreeBuffer(m_pIdentityProps);
	MyFreeBuffer(m_pProps);

	if (m_pStatusObj)
		m_pStatusObj->Release();

	if (m_pMAPISup)
		m_pMAPISup->Release();

	if (m_pProvider)
		m_pProvider->Release();

	DeleteCriticalSection(&m_csObj);
}

STDMETHODIMP CXPLogon::QueryInterface(REFIID riid, LPVOID* ppvObj)
{
	if (!ppvObj)
		return E_INVALIDARG;

	*ppvObj = NULL;

	if (riid == IID_IXPLogon || riid == IID_IUnknown)
	{
		*ppvObj = (LPVOID)this;
		AddRef();
		return S_OK;
	}

	return E_NOINTERFACE;
}

// BEGIN IXPLogon Implementation

/***********************************************************************************************
	CXPLogon::AddressTypes

		- If this is a UNICODE build, return MAPI_UNICODE in the lpulFlags parameter. If not, set
		it to 0.
		- Return a static array (of 1 entry) in lpppAdrTypeArray
		- Return the number of entries (1) in lpcAdrType
		- Return 0 and NULL in lpcMAPIUID and lpppUIDArray, respectively.

		Refer to the MSDN documentation for more information.

***********************************************************************************************/

STDMETHODIMP CXPLogon::AddressTypes(ULONG FAR*				lpulFlags,
									ULONG FAR *				lpcAdrType,
									LPTSTR FAR * FAR *		lpppAdrTypeArray,
									ULONG FAR *				lpcMAPIUID,
									LPMAPIUID FAR * FAR *	lpppUIDArray)
{
	Log(true, _T("CXPLogon::AddressTypes function called\n"));

	if (!lpulFlags || !lpcAdrType || !lpppAdrTypeArray || !lpcMAPIUID || !lpppUIDArray)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	static LPTSTR lpszAddrType = MRXP_ADDRTYPE;

	*lpulFlags = fMapiUnicode;
	*lpcMAPIUID = 0;
	*lpppUIDArray = NULL;

	*lpcAdrType = 1;
	*lpppAdrTypeArray = &lpszAddrType;

	return hRes;
}

/***********************************************************************************************
	CXPLogon::RegisterOptions

	Refer to the MSDN documentation for more information.

***********************************************************************************************/

STDMETHODIMP CXPLogon::RegisterOptions(ULONG FAR *			lpulFlags,
									   ULONG FAR *			lpcOptions,
									   LPOPTIONDATA FAR *	lppOptions)
{
	Log(true, _T("CXPLogon::RegisterOptions function called\n"));

	if (!lpulFlags || !lpcOptions || !lppOptions)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	return hRes;
}

/***********************************************************************************************
	CXPLogon::TransportNotify

		- Function is wrapped in a critical section
		- Add or remove status bits based on what flags were passed
		- Ignore NOTIFY_ABORT_DEFERRED and NOTIFY_CANCEL_MESSAGE

		Refer to the MSDN documentation for more information.

***********************************************************************************************/

STDMETHODIMP CXPLogon::TransportNotify(ULONG FAR *	lpulFlags,
									   LPVOID FAR *	/*lppvData*/)
{
	Log(true, _T("CXPLogon::TransportNotify function called\n"));

	if (!lpulFlags)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	EnterCriticalSection(m_pProvider->GetCS());

	if (*lpulFlags & NOTIFY_BEGIN_INBOUND)
	{
		Log(true, _T("CXPLogon::TransportNotify received NOTIFY_BEGIN_INBOUND\n"));
		AddStatusBits(STATUS_INBOUND_ENABLED);
	}

    if (*lpulFlags & NOTIFY_END_INBOUND)
    {
        Log(true, _T("CXPLogon::TransportNotify received NOTIFY_END_INBOUND\n"));
        RemoveStatusBits(STATUS_INBOUND_ENABLED);
    }

    if (*lpulFlags & NOTIFY_END_INBOUND_FLUSH)
    {
        Log(true, _T("CXPLogon::TransportNotify received NOTIFY_END_INBOUND_FLUSH\n"));
		RemoveStatusBits(STATUS_INBOUND_FLUSH);
    }

    if (*lpulFlags & NOTIFY_BEGIN_OUTBOUND)
    {
        Log(true, _T("CXPLogon::TransportNotify received NOTIFY_BEGIN_OUTBOUND\n"));
		AddStatusBits(STATUS_OUTBOUND_ENABLED);
    }

    if (*lpulFlags & NOTIFY_END_OUTBOUND)
    {
        Log(true, _T("CXPLogon::TransportNotify received NOTIFY_END_OUTBOUND\n"));
		RemoveStatusBits(STATUS_OUTBOUND_ENABLED);
    }

    if (*lpulFlags & NOTIFY_END_OUTBOUND_FLUSH)
    {
        Log(true, _T("CXPLogon::TransportNotify received NOTIFY_END_OUTBOUND_FLUSH\n"));
		RemoveStatusBits(STATUS_OUTBOUND_FLUSH);
    }

    if (*lpulFlags & NOTIFY_ABORT_DEFERRED)
    {
		// Ignore NOTIFY_ABORT_DEFERRED.
        Log(true, _T("CXPLogon::TransportNotify received NOTIFY_ABORT_DEFERRED\n"));
    }

    if (*lpulFlags & NOTIFY_CANCEL_MESSAGE)
    {
		// Ignore NOTIFY_CANCEL_MESSAGE.
        Log(true, _T("CXPLogon::TransportNotify received NOTIFY_CANCEL_MESSAGE\n"));
    }

	hRes = UpdateStatusRow();
	if (FAILED(hRes))
		Log(true, _T("CXPLogon::TransportNotify: UpdateStatusRow returned an error: 0x%08X\n"), hRes);

	LeaveCriticalSection(m_pProvider->GetCS());

	return hRes;
}

/***********************************************************************************************
	CXPLogon::Idle

	Refer to the MSDN documentation for more information.

***********************************************************************************************/

STDMETHODIMP CXPLogon::Idle(ULONG /*ulFlags*/)
{
	HRESULT hRes = S_OK;

	// TODO:
	// 1. Add a profile prop to enable/disable proactive polling on Idle, check it here
	// 2. Check that STATUS_INBOUND_ENABLED is on
	// 3. If it is, call Poll.
	// 4. If Poll returns a non-zero value, call SpoolerNotify with NOTIFY_NEWMAIL

	return hRes;
}

/***********************************************************************************************
	CXPLogon::TransportLogoff

		- Remove this instance from the linked list of CXPLogons in the parent provider object
		- Defer releasing resources until the destructor is called

		Refer to the MSDN documentation for more information.

***********************************************************************************************/

STDMETHODIMP CXPLogon::TransportLogoff(ULONG /*ulFlags*/)
{
	Log(true, _T("CXPLogon::TransportLogoff function called\n"));
	HRESULT hRes = S_OK;

	// Unlink this session from the provider's list
	CXPLogon* pPrev = NULL;
	CXPLogon* pCur = NULL;

	EnterCriticalSection(m_pProvider->GetCS());

	pCur = m_pProvider->GetLogonList();

	while (pCur)
	{
		if (pCur == this)
		{
			if (pPrev == NULL)
				m_pProvider->SetLogonList(pCur->GetNextLogon());
			else
				pPrev->SetNextLogon(pCur->GetNextLogon());
			break;

		}

		pPrev = pCur;
		pCur = pCur->GetNextLogon();
	}

	LeaveCriticalSection(m_pProvider->GetCS());

	return hRes;
}

/***********************************************************************************************
	CXPLogon::SubmitMessage

		- Add STATUS_OUTBOUND_ACTIVE to PR_STATUS_CODE in the provider's status row
		- Set the sender properties on the message
		- Get the recipient table from the passed in message, and check it for any unhandled
		MRXP recipients
		- Actually "send" the message (save it in MSG format to the recipients' shares)
		- Set PR_RESPONSIBILITY to TRUE for the MRXP recipients
		- Strip any duplicate MRXP recips or MRXP recips that are MAPI_P1
		- Pass any MRXP recipients that encountered errors on to IMAPISupport::StatusRecips
		to generate an NDR.
		- Remove STATUS_OUTBOUND_ACTIVE from PR_STATUS_CODE in the status row when the provider is done
		- If there are no MRXP recipients on the message, return MAPI_E_NOT_ME.

		Refer to the MSDN documentation for more information.

***********************************************************************************************/

STDMETHODIMP CXPLogon::SubmitMessage(ULONG			/*ulFlags*/,
									 LPMESSAGE		lpMessage,
									 ULONG_PTR FAR *lpulMsgRef,
									 ULONG_PTR FAR *lpulReturnParm)
{
	Log(true, _T("CXPLogon::SubmitMessage function called\n"));

	if (!lpMessage || !lpulMsgRef || !lpulReturnParm)
		return MAPI_E_INVALID_PARAMETER;

	if (m_ulStatusBits & STATUS_OUTBOUND_ACTIVE)
	{
		Log(true, _T("CXPLogon::SubmitMessage: called while already processing. Returning MAPI_E_BUSY\n"));
		return MAPI_E_BUSY;
	}

	HRESULT hRes = S_OK;

	AddStatusBits(STATUS_OUTBOUND_ACTIVE);
	hRes = UpdateStatusRow();
	if (FAILED(hRes))
	{
		Log(true, _T("CXPLogon::SubmitMessage: UpdateStatusRow returned an error: 0x%08X\n"), hRes);
		RemoveStatusBits(STATUS_OUTBOUND_ACTIVE);
		return hRes;
	}

	LPADRLIST lpMyRecips = NULL;

	hRes = GetRecipientsToProcess(lpMessage, &lpMyRecips);
	if (SUCCEEDED(hRes) && lpMyRecips)
	{
		// Stamp Sender Props

		hRes = StampSender(lpMessage);
		if (FAILED(hRes))
		{
			Log(true, _T("CXPLogon::SubmitMessage: StampSender returned an error: 0x%08X\n"), hRes);
		}

		// Get current system time (for file name)
		// The file name needs to be unique to prevent conflicts
		SYSTEMTIME CurrentTime = {0};
		GetSystemTime(&CurrentTime);

		LPTSTR pszFileName = NULL;
		TCHAR szFileName[MAX_PATH] = {0};

		hRes = StringCchPrintf(szFileName, CCH(szFileName), _T("%02d%02d%4d%02d%02d%02d%03d%s"),
			CurrentTime.wMonth, CurrentTime.wDay, CurrentTime.wYear, CurrentTime.wHour,
			CurrentTime.wMinute, CurrentTime.wSecond, CurrentTime.wMilliseconds, MSG_EXT);
		if (SUCCEEDED(hRes))
		{
			pszFileName = szFileName;
		}
		else
		{
			pszFileName = DEFAULT_FILENAME;
		}

		for (ULONG i = 0; i < lpMyRecips->cEntries; i++)
		{
			// Do the submit
			if (lpMyRecips->aEntries[i].rgPropVals[RECIP_EMAIL_ADDRESS].ulPropTag == PR_EMAIL_ADDRESS)
			{
				LPTSTR pszSaveToDir = NULL;

				pszSaveToDir = lpMyRecips->aEntries[i].rgPropVals[RECIP_EMAIL_ADDRESS].Value.LPSZ;

				TCHAR szFullFilePath[MAX_PATH] = {0};

				hRes = StringCchPrintf(szFullFilePath, CCH(szFullFilePath), _T("%s\\%s"), pszSaveToDir, pszFileName);
				if (SUCCEEDED(hRes))
				{
					Log(true, _T("CXPLogon::SubmitMessage: Delivering message to %s\n"), szFullFilePath);

					hRes = SaveToMSG(lpMessage, szFullFilePath);
					if (FAILED(hRes))
					{
						Log(true, _T("CXPLogon::SubmitMessage: SaveToMSG returned an error: 0x%08X\n"), hRes);
						// Do NDR
						hRes = SetNDRText(lpMyRecips->aEntries[i], hRes, (LPVOID)lpMyRecips);
						if (FAILED(hRes))
						{
							Log(true, _T("CXPLogon::SubmitMessage: SetNDRText returned an error: 0x%08X\n"), hRes);
						}

						hRes = S_OK;
					}
				}
				else
				{
					Log(true, _T("CXPLogon::SubmitMessage: StringCchPrintf returned an error: 0x%08X\n"), hRes);
					// NDR
				}
			}
			else
			{
				Log(true, _T("CXPLogon::SubmitMessage: PR_EMAIL_ADDRESS on a recipient was bad\n"));
				// Error getting the e-mail address - NDR
			}

			// Signal that the provider handled this one
			lpMyRecips->aEntries[i].rgPropVals[RECIP_RESPONSBILITY].Value.b = TRUE;
		}

		LPADRLIST lpBadRecips = NULL;
		hRes = GenerateBadRecipList(lpMyRecips, &lpBadRecips);
		if (SUCCEEDED(hRes) && lpBadRecips)
		{
			hRes = m_pMAPISup->StatusRecips(lpMessage, lpBadRecips);
			if (FAILED(hRes))
			{
				Log(true, _T("CXPLogon::SubmitMessage: StatusRecips returned an error: 0x%08X\n"), hRes);
			}
		}
		else if (FAILED(hRes))
		{
			Log(true, _T("CXPLogon::SubmitMessage: GenerateBadRecipList returned an error: 0x%08X\n"), hRes);
		}

		FreePadrlist(lpBadRecips);

		hRes = lpMessage->ModifyRecipients(MODRECIP_MODIFY, lpMyRecips);
		if (SUCCEEDED(hRes))
		{
			hRes = lpMessage->SaveChanges(0);
			if (FAILED(hRes))
				Log(true, _T("CXPLogon::SubmitMessage: SaveChanges returned an error: 0x%08X\n"), hRes);
		}
		else
		{
			Log(true, _T("CXPLogon::SubmitMessage: ModifyRecipients returned an error: 0x%08X\n"), hRes);
		}
	}
	else
	{
		Log(true, _T("CXPLogon::SubmitMessage: GetRecipientsToProcess returned an error: 0x%08X\n"), hRes);
	}

	FreePadrlist(lpMyRecips);

	RemoveStatusBits(STATUS_OUTBOUND_ACTIVE);
	HRESULT hResTemp = S_OK;
	hResTemp = UpdateStatusRow();
	if (FAILED(hResTemp))
	{
		Log(true, _T("CXPLogon::SubmitMessage: UpdateStatusRow returned an error: 0x%08X\n"), hResTemp);
	}

	if (FAILED(hRes))
	{
		m_pMAPISup->SpoolerNotify(NOTIFY_CRITICAL_ERROR, NULL);
	}

	return hRes;
}

/***********************************************************************************************
	CXPLogon::EndMessage

		Set lpulFlags to 0 and return S_OK.

		Refer to the MSDN documentation for more information.

***********************************************************************************************/

STDMETHODIMP CXPLogon::EndMessage(ULONG_PTR		/*ulMsgRef*/,
								  ULONG FAR *	lpulFlags)
{
	Log(true, _T("CXPLogon::EndMessage function called\n"));

	if (!lpulFlags)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	// 0 indicates message has been sent
	*lpulFlags = 0;

	return hRes;
}

/***********************************************************************************************
	CXPLogon::Poll

		If there are messages waiting, set lpulIncoming to non-zero.
		If there aren't messages waiting, set lpulIncoming to zero.

		Refer to the MSDN documentation for more information.

***********************************************************************************************/

STDMETHODIMP CXPLogon::Poll(ULONG FAR *	lpulIncoming)
{
	Log(true, _T("CXPLogon::Poll function called\n"));

	if (!lpulIncoming)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	// No incoming messages
	*lpulIncoming = 0;

	if (m_ulStatusBits & STATUS_INBOUND_ENABLED)
	{
		hRes = NewMailWaiting(NULL, NULL);
		if (S_FALSE == hRes)
		{
			// No messages waiting
			hRes = S_OK;
		}
		else if (SUCCEEDED(hRes))
		{
			Log(true, _T("CXPLogon::Poll: New messages are available\n"));
			*lpulIncoming = 1;
		}
		else
		{
			Log(true, _T("CXPLogon::Poll: NewMailWaiting returned an error: 0x%08X\n"), hRes);
		}
	}

	return hRes;
}

/***********************************************************************************************
	CXPLogon::StartMessage

		- Add STATUS_INBOUND_ACTIVE to PR_STATUS_CODE in the provider's status row
		- If the provider has any new mail waiting to deliver, set the appropriate properties from the
		incoming mail on the lpMessage parameter.
		- Set the PR_RECEIVED_BY* properties on the message to the provider's identity props.
		- Save changes to the message
		- Remove STATUS_INBOUND_ACTIVE to PR_STATUS_CODE from the provider's status row
		- If the provider has no waiting messages to deliver, then return from the call WITHOUT calling
		SaveChanges on the lpMessage parameter.

		Refer to the MSDN documentation for more information.

***********************************************************************************************/

STDMETHODIMP CXPLogon::StartMessage(ULONG		/*ulFlags*/,
									LPMESSAGE	lpMessage,
									ULONG_PTR FAR *lpulMsgRef)
{
	Log(true, _T("CXPLogon::StartMessage function called\n"));

	if (!lpMessage || !lpulMsgRef || !m_pProps)
		return MAPI_E_INVALID_PARAMETER;

	if (m_ulStatusBits & STATUS_INBOUND_ACTIVE)
	{
		Log(true, _T("CXPLogon::StartMessage: called while already processing. Returning MAPI_E_BUSY\n"));
		return MAPI_E_BUSY;
	}

	HRESULT hRes = S_OK;

	AddStatusBits(STATUS_INBOUND_ACTIVE);
	hRes = UpdateStatusRow();
	if (FAILED(hRes))
	{
		Log(true, _T("CXPLogon::StartMessage: UpdateStatusRow returned an error: 0x%08X\n"), hRes);
		RemoveStatusBits(STATUS_INBOUND_ACTIVE);
		return hRes;
	}

	// See if there's a message to deliver
	LPTSTR pszNewMailFile = NULL;
	FILETIME ftDelivered = {0};

	hRes = NewMailWaiting(&pszNewMailFile, &ftDelivered);
	// Using SUCCEEDED here won't work..S_FALSE will pass the SUCCEEDED macro
	if (S_OK == hRes && pszNewMailFile)
	{
		Log(true, _T("CXPLogon::StartMessage: processing %s\n"), pszNewMailFile);

		hRes = LoadFromMSG(pszNewMailFile, lpMessage);
		if (SUCCEEDED(hRes))
		{
			// Set the provider's Received By properties and delivery time properties.
			SPropValue spvTransportStamped[NUM_XP_SETPROPS] = {0};

			spvTransportStamped[SENDER_DISPLAY_NAME].ulPropTag = PR_RECEIVED_BY_NAME_A;
			spvTransportStamped[SENDER_DISPLAY_NAME].Value.lpszA = m_pIdentityProps[IDENTITY_DISPLAY_NAME].Value.lpszA;

			spvTransportStamped[SENDER_EMAIL_ADDRESS].ulPropTag = PR_RECEIVED_BY_EMAIL_ADDRESS_A;
			spvTransportStamped[SENDER_EMAIL_ADDRESS].Value.lpszA = m_pProps[CLIENT_INBOX].Value.lpszA;

			spvTransportStamped[SENDER_ADDRTYPE].ulPropTag = PR_RECEIVED_BY_ADDRTYPE_A;
			spvTransportStamped[SENDER_ADDRTYPE].Value.lpszA = MRXP_ADDRTYPE_A;

			spvTransportStamped[SENDER_SEARCH_KEY].ulPropTag = PR_RECEIVED_BY_SEARCH_KEY;
			spvTransportStamped[SENDER_SEARCH_KEY].Value.bin = m_pIdentityProps[IDENTITY_SEARCH_KEY].Value.bin;

			spvTransportStamped[SENDER_ENTRYID].ulPropTag = PR_RECEIVED_BY_ENTRYID;
			spvTransportStamped[SENDER_ENTRYID].Value.bin = m_pIdentityProps[IDENTITY_ENTRYID].Value.bin;

			spvTransportStamped[MSG_RECEIVE_TIME].ulPropTag = PR_MESSAGE_DELIVERY_TIME;
			spvTransportStamped[MSG_RECEIVE_TIME].Value.ft = ftDelivered;

			// TODO: Add support to set MSGFLAG_FROMME
			spvTransportStamped[MSG_FLAGS].ulPropTag = PR_MESSAGE_FLAGS;
			spvTransportStamped[MSG_FLAGS].Value.ul = 0;

			hRes = lpMessage->SetProps(NUM_XP_SETPROPS, spvTransportStamped, NULL);
			if (SUCCEEDED(hRes))
			{
				// Pass ITEMPROC_FORCE | NON_EMS_XP_SAVE to force the PST to run rules on this message.
				hRes = lpMessage->SaveChanges(KEEP_OPEN_READWRITE | ITEMPROC_FORCE | NON_EMS_XP_SAVE);
				if (SUCCEEDED(hRes))
				{
					BOOL bRet = DeleteFile(pszNewMailFile);
					if (!bRet)
					{
						Log(true, _T("CXPLogon::StartMessage: DeleteFile returned an error: 0x%08X\n"), GetLastError());
					}
				}
				else
				{
					Log(true, _T("CXPLogon::StartMessage: SaveChanges returned an error: 0x%08X\n"), hRes);
				}
			}
			else
			{
				Log(true, _T("CXPLogon::StartMessage: SetProps returned an error: 0x%08X\n"), hRes);
			}
		}
		else
		{
			Log(true, _T("CXPLogon::StartMessage: LoadFromMSG returned an error: 0x%08X\n"), hRes);
		}
	}
	else if (S_FALSE == hRes)
	{
		Log(true, _T("CXPLogon::StartMessage: No message to deliver...exiting\n"));
	}
	else
	{
		Log(true, _T("CXPLogon::StartMessage: NewMailWaiting returned an error: 0x%08X\n"), hRes);
	}

	MyFreeBuffer(pszNewMailFile);

	// Preserve any error from above to return back to the spooler
	HRESULT hResTemp = S_OK;
	RemoveStatusBits(STATUS_INBOUND_ACTIVE);
	hResTemp = UpdateStatusRow();
	if (FAILED(hResTemp))
	{
		Log(true, _T("CXPLogon::StartMessage: UpdateStatusRow returned an error: 0x%08X\n"), hResTemp);
	}

	if (FAILED(hRes))
	{
		m_pMAPISup->SpoolerNotify(NOTIFY_CRITICAL_ERROR, NULL);

		ULONG ulIncoming = 0;
		Poll(&ulIncoming);
	}

	return hRes;
}

/***********************************************************************************************
	CXPLogon::OpenStatusEntry

		- QI the provider's member CXPStatus object for either the interface passed in or IID_IMAPIStatus
		and return the resulting pointer in lppEntry.

		Refer to the MSDN documentation for more information.

***********************************************************************************************/

STDMETHODIMP CXPLogon::OpenStatusEntry(LPCIID               lpInterface,
									   ULONG                ulFlags,
									   ULONG FAR *          lpulObjType,
									   LPMAPISTATUS FAR *	lppEntry)
{
	Log(true, _T("CXPLogon::OpenStatusEntry function called\n"));

	if (!lpulObjType || ! lppEntry)
		return MAPI_E_INVALID_PARAMETER;

	if (ulFlags & MAPI_MODIFY)
		return MAPI_E_NO_ACCESS;

	HRESULT hRes = S_OK;

	*lpulObjType = 0;
	*lppEntry = NULL;

	if (!m_pStatusObj)
	{
		hRes = CreateStatusRow();
	}

	if (SUCCEEDED(hRes) && m_pStatusObj)
	{
		hRes = m_pStatusObj->QueryInterface(lpInterface ? *lpInterface : IID_IMAPIStatus, (LPVOID*)lppEntry);
		if (SUCCEEDED(hRes) && *lppEntry)
		{
			*lpulObjType = MAPI_STATUS;
		}
	}

	return hRes;
}

/***********************************************************************************************
	CXPLogon::ValidateState

		- Pass this call along to the IMAPIStatus object.

		Refer to the MSDN documentation for more information.

***********************************************************************************************/

STDMETHODIMP CXPLogon::ValidateState(ULONG_PTR ulUIParam,
									 ULONG ulFlags)
{
	Log(true, _T("CXPLogon::ValidateState function called\n"));

	HRESULT hRes = S_OK;

	if (!m_pStatusObj)
	{
		hRes = CreateStatusRow();
	}

	if (SUCCEEDED(hRes) && m_pStatusObj)
	{
		hRes = m_pStatusObj->ValidateState(ulUIParam, ulFlags);
	}

	return hRes;
}

/***********************************************************************************************
	CXPLogon::FlushQueues

		- Update the provider's status row appropriately, depending on what flags are passed.

		Refer to the MSDN documentation for more information.

***********************************************************************************************/

STDMETHODIMP CXPLogon::FlushQueues(ULONG_PTR	/*ulUIParam*/,
								   ULONG		/*cbTargetTransport*/,
								   LPENTRYID	/*lpTargetTransport*/,
								   ULONG		ulFlags)
{
	Log(true, _T("CXPLogon::FlushQueues function called\n"));
	HRESULT hRes = S_OK;

	if (ulFlags & FLUSH_UPLOAD)
		AddStatusBits(STATUS_OUTBOUND_FLUSH);
	if (ulFlags & FLUSH_DOWNLOAD)
		AddStatusBits(STATUS_INBOUND_FLUSH);

	hRes = UpdateStatusRow();
	if (FAILED(hRes))
		Log(true, _T("CXPLogon::FlushQueues: UpdateStatusRow returned an error: 0x%08X\n"), hRes);

	return hRes;
}

// END IXPLogon Implementation

/***********************************************************************************************
	CXPLogon::GenerateIdentity

	To generate the "identity" properties based on the provider's profile configuration. This
	generates an array of SPropValues that is then saved as a member variable.

***********************************************************************************************/

HRESULT CXPLogon::GenerateIdentity()
{
	Log(true, _T("CXPLogon::GenerateIdentity function called\n"));

	if (!m_pProps)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	// Create and setup the provider's identity props
	hRes = MyAllocateBuffer(sizeof(SPropValue) * NUM_IDENTITY_PROPS, (LPVOID*)&m_pIdentityProps);
	if (SUCCEEDED(hRes) && m_pIdentityProps)
	{
		ZeroMemory(m_pIdentityProps, sizeof(SPropValue) * NUM_IDENTITY_PROPS);

		m_pIdentityProps[IDENTITY_DISPLAY_NAME].ulPropTag = PR_SENDER_NAME_A;
		m_pIdentityProps[IDENTITY_DISPLAY_NAME].Value.lpszA = m_pProps[CLIENT_USER_DISPLAY].Value.lpszA;

		m_pIdentityProps[IDENTITY_ADDRTYPE].ulPropTag = PR_SENDER_ADDRTYPE_A;
		m_pIdentityProps[IDENTITY_ADDRTYPE].Value.lpszA = MRXP_ADDRTYPE_A;

		m_pIdentityProps[IDENTITY_EMAIL_ADDRESS].ulPropTag = PR_SENDER_EMAIL_ADDRESS_A;
		m_pIdentityProps[IDENTITY_EMAIL_ADDRESS].Value.lpszA = m_pProps[CLIENT_INBOX].Value.lpszA;

		if (!m_cchInboxPath)
		{
			hRes = StringCchLengthA(m_pProps[CLIENT_INBOX].Value.lpszA, STRSAFE_MAX_CCH, &m_cchInboxPath);
		}

		if (SUCCEEDED(hRes) && m_cchInboxPath)
		{
			size_t cbSearchKey = (m_cchInboxPath + CCH_A(MRXP_ADDRTYPE_A) + 1) * sizeof(char);
			LPSTR pszSearchKey = NULL;

			hRes = MyAllocateMore((ULONG) cbSearchKey, m_pIdentityProps, (LPVOID*)&pszSearchKey);
			if (SUCCEEDED(hRes) && pszSearchKey)
			{
				hRes = StringCbPrintfA(pszSearchKey, cbSearchKey, "%s:%s",
					MRXP_ADDRTYPE_A, m_pProps[CLIENT_INBOX].Value.lpszA);
				if (SUCCEEDED(hRes))
				{
					m_pIdentityProps[IDENTITY_SEARCH_KEY].ulPropTag = PR_SENDER_SEARCH_KEY;
					m_pIdentityProps[IDENTITY_SEARCH_KEY].Value.bin.cb = (ULONG) cbSearchKey;
					m_pIdentityProps[IDENTITY_SEARCH_KEY].Value.bin.lpb = (LPBYTE)pszSearchKey;

					SBinary sbEID = {0};

					hRes = m_pMAPISup->CreateOneOff((LPTSTR)m_pProps[CLIENT_USER_DISPLAY].Value.lpszA,
						(LPTSTR)MRXP_ADDRTYPE_A, (LPTSTR)m_pProps[CLIENT_INBOX].Value.lpszA, 0,
						&(sbEID.cb), (LPENTRYID*)&(sbEID.lpb));
					if (SUCCEEDED(hRes) && sbEID.cb && sbEID.lpb)
					{
						m_pIdentityProps[IDENTITY_ENTRYID].ulPropTag = PR_SENDER_ENTRYID;
						m_pIdentityProps[IDENTITY_ENTRYID].Value.bin.cb = sbEID.cb;
						hRes = CopySBinary(&(m_pIdentityProps[IDENTITY_ENTRYID].Value.bin),
							&sbEID, (LPVOID)m_pIdentityProps, NULL, m_pAllocMore);
						if (FAILED(hRes))
						{
							Log(true, _T("CXPLogon::GenerateIdentity: CopySBinary returned an error: 0x%08X\n"), hRes);
						}
					}
					else
					{
						Log(true, _T("CXPLogon::GenerateIdentity: CreateOneOff returned an error: 0x%08X\n"), hRes);
					}

					MyFreeBuffer(sbEID.lpb);
				}
				else
				{
					Log(true, _T("CXPLogon::GenerateIdentity: StringCbPrintfA returned an error: 0x%08X\n"), hRes);
				}
			}
			else
			{
				Log(true, _T("CXPLogon::GenerateIdentity: MyAllocateMore returned an error: 0x%08X\n"), hRes);
			}
		}
		else
		{
			Log(true, _T("CXPLogon::GenerateIdentity: Error calculating length of inbox path: 0x%08X"), hRes);
		}
	}
	else
	{
		Log(true, _T("CXPLogon::GenerateIdentity: MyAllocateBuffer for identity props returned an error: 0x%08X\n"), hRes);
	}

	if (FAILED(hRes))
	{
		MyFreeBuffer(m_pIdentityProps);
		m_pIdentityProps = NULL;
	}

	return hRes;
}

/***********************************************************************************************
	CXPLogon::CreateStatusRow

	To add a status row for the provider into the session's status table.  Uses the
	properties generated by GenerateIdentity

***********************************************************************************************/

HRESULT CXPLogon::CreateStatusRow()
{
	Log(true, _T("CXPLogon::CreateStatusRow function called\n"));

	if (!m_pProps)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	LPSPropValue lpStatusProps = NULL;

	if (!m_pIdentityProps)
	{
		hRes = GenerateIdentity();
		if (FAILED(hRes))
		{
			Log(true, _T("CXPLogon::CreateStatusRow: GenerateIdentity returned an error: 0x%08X\n"), hRes);
			return hRes;
		}
	}

	// Create the provider's Status object
	m_pStatusObj = new CXPStatus(this, m_pMAPISup, NUM_IDENTITY_PROPS, m_pIdentityProps,
		m_pAllocBuffer, m_pAllocMore, m_pFreeBuffer);
	if (!m_pStatusObj)
		Log(true, _T("CXPLogon::CreateStatusRow: Error creating new CXPStatus object: 0x%08X\n"), GetLastError());

	hRes = MyAllocateBuffer(sizeof(SPropValue) * (NUM_STATUS_ROW_PROPS), (LPVOID *)&lpStatusProps);
	if(SUCCEEDED(hRes) && lpStatusProps)
	{
		int iRet = 0;
		ZeroMemory(lpStatusProps, sizeof(SPropValue) * (NUM_STATUS_ROW_PROPS - 1));

		lpStatusProps[STATUS_RESOURCE_METHODS].ulPropTag = PR_RESOURCE_METHODS;
		lpStatusProps[STATUS_RESOURCE_METHODS].Value.ul = MRXP_RESOURCE_METHODS;

		lpStatusProps[STATUS_PROVIDER_DISPLAY].ulPropTag = PR_PROVIDER_DISPLAY;
		hRes = MyAllocateMore(DEF_STATUS_ROW_STRING_LEN * sizeof(TCHAR), (LPVOID)lpStatusProps,
			(LPVOID*)&(lpStatusProps[STATUS_PROVIDER_DISPLAY].Value.LPSZ));
		if (SUCCEEDED(hRes) && lpStatusProps[STATUS_PROVIDER_DISPLAY].Value.LPSZ)
		{
			iRet = LoadString(m_pProvider->GetHINST(), IDS_PROVIDER_DISPLAY,
				lpStatusProps[STATUS_PROVIDER_DISPLAY].Value.LPSZ, DEF_STATUS_ROW_STRING_LEN);
		}
		if (iRet <= 0)
		{
			lpStatusProps[STATUS_PROVIDER_DISPLAY].Value.LPSZ = MRXP_ERROR_STATUS_STRING;
		}

		lpStatusProps[STATUS_DISPLAY_NAME].ulPropTag = PR_DISPLAY_NAME_A;
		lpStatusProps[STATUS_DISPLAY_NAME].Value.lpszA = m_pProps[CLIENT_INBOX].Value.lpszA;

		lpStatusProps[STATUS_IDENTITY_DISPLAY].ulPropTag = PR_IDENTITY_DISPLAY_A;
		lpStatusProps[STATUS_IDENTITY_DISPLAY].Value.lpszA = m_pProps[CLIENT_USER_DISPLAY].Value.lpszA;

		lpStatusProps[STATUS_STATUS_CODE].ulPropTag = PR_STATUS_CODE;
		lpStatusProps[STATUS_STATUS_CODE].Value.ul = m_ulStatusBits;

		iRet = 0;
		lpStatusProps[STATUS_STATUS_STRING].ulPropTag = PR_STATUS_STRING;
		hRes = MyAllocateMore(DEF_STATUS_ROW_STRING_LEN * sizeof(TCHAR), (LPVOID)lpStatusProps,
			(LPVOID*)&(lpStatusProps[STATUS_STATUS_STRING].Value.LPSZ));
		if (SUCCEEDED(hRes) && lpStatusProps[STATUS_STATUS_STRING].Value.LPSZ)
		{
			UINT ids = 0;

			ids = GetStatusStringID();

			iRet = LoadString(m_pProvider->GetHINST(), ids,
				lpStatusProps[STATUS_STATUS_STRING].Value.LPSZ, DEF_STATUS_ROW_STRING_LEN);
		}
		if (iRet <= 0)
		{
			lpStatusProps[STATUS_STATUS_STRING].Value.LPSZ = MRXP_ERROR_STATUS_STRING;
		}

		lpStatusProps[STATUS_IDENTITY_ENTRYID].ulPropTag = PR_IDENTITY_ENTRYID;
		lpStatusProps[STATUS_IDENTITY_ENTRYID].Value.bin.cb = m_pIdentityProps[IDENTITY_ENTRYID].Value.bin.cb;
		lpStatusProps[STATUS_IDENTITY_ENTRYID].Value.bin.lpb = m_pIdentityProps[IDENTITY_ENTRYID].Value.bin.lpb;

		lpStatusProps[STATUS_IDENTITY_SEARCH_KEY].ulPropTag = PR_IDENTITY_SEARCH_KEY;
		lpStatusProps[STATUS_IDENTITY_SEARCH_KEY].Value.bin.cb = m_pIdentityProps[IDENTITY_SEARCH_KEY].Value.bin.cb;
		lpStatusProps[STATUS_IDENTITY_SEARCH_KEY].Value.bin.lpb = m_pIdentityProps[IDENTITY_SEARCH_KEY].Value.bin.lpb;

		hRes = m_pMAPISup->ModifyStatusRow(NUM_STATUS_ROW_PROPS - 1, lpStatusProps, 0);
		if (FAILED(hRes))
		{
			Log(true, _T("CXPLogon::CreateStatusRow: ModifyStatusRow returned an error: 0x%08X\n"), hRes);
		}
	}
	else
	{
		Log(true, _T("CXPLogon::CreateStatusRow: MAPIAllocateBuffer for status row props returned an error: 0x%08X\n"), hRes);
	}

	MyFreeBuffer(lpStatusProps);

	return hRes;
}

/***********************************************************************************************
	CXPLogon::GetStatusStringID

	Based on the current value of PR_STATUS_CODE, loads the appropriate string
	from the string table.  Used to set PR_STATUS_STRING

***********************************************************************************************/

UINT CXPLogon::GetStatusStringID()
{
	/* Determine the IDS of the status.  Since the status' are bit fields
    of ulStatus, apply this hierarcy to determine the correct string. */

    if (m_ulStatusBits & STATUS_INBOUND_ACTIVE)
        return IDS_STATUS_UPLOADING;
    else if (m_ulStatusBits & STATUS_OUTBOUND_ACTIVE)
        return IDS_STATUS_DOWNLOADING;
    else if (m_ulStatusBits & STATUS_INBOUND_FLUSH)
        return IDS_STATUS_INFLUSHING;
    else if (m_ulStatusBits & STATUS_OUTBOUND_FLUSH)
        return IDS_STATUS_OUTFLUSHING;
    else if ((m_ulStatusBits & STATUS_AVAILABLE) &&
            ((m_ulStatusBits & STATUS_INBOUND_ENABLED) ||
            (m_ulStatusBits & STATUS_OUTBOUND_ENABLED)))
        return IDS_STATUS_ONLINE;
    else if (m_ulStatusBits & STATUS_AVAILABLE)
        return IDS_STATUS_AVAILABLE;
    else
        return IDS_STATUS_OFFLINE;
}

/***********************************************************************************************
	CXPLogon::UpdateStatusRow

	Updates PR_STATUS_CODE and PR_STATUS_STRING in our status row.

***********************************************************************************************/

HRESULT CXPLogon::UpdateStatusRow()
{
	HRESULT hRes = S_OK;
	SPropValue NewStatus[2] = {0};

	NewStatus[0].ulPropTag = PR_STATUS_CODE;
	NewStatus[0].Value.ul = m_ulStatusBits;

	int iRet = 0;
	TCHAR szStatus[128] = {0};

	iRet = LoadString(m_pProvider->GetHINST(), GetStatusStringID(), szStatus, 128);

	NewStatus[1].ulPropTag = PR_STATUS_STRING;
	if (iRet > 0)
		NewStatus[1].Value.LPSZ = szStatus;
	else
		NewStatus[1].Value.LPSZ = MRXP_ERROR_STATUS_STRING;

	hRes = m_pMAPISup->ModifyStatusRow(2, NewStatus, STATUSROW_UPDATE);
	if (FAILED(hRes))
		Log(true, _T("CXPLogon::UpdateStatusRow: ModifyStatusRow returned an error: 0x%08X\n"), hRes);

	return hRes;
}

/***********************************************************************************************
	CXPLogon::GetRecipientsToProcess

	Returns a list of MRXP recipients on the message that the provider needs to handle.  Called
	from SubmitMessage.

***********************************************************************************************/

HRESULT CXPLogon::GetRecipientsToProcess(LPMESSAGE lpMessage, LPADRLIST* lppAdrList)
{
	Log(true, _T("CXPLogon::GetRecipientsToProcess function called\n"));

	if (!lpMessage || !lppAdrList)
		return MAPI_E_INVALID_PARAMETER;

	*lppAdrList = NULL;

	HRESULT hRes = S_OK;

	hRes = StripDupsAndP1(lpMessage);

	LPMAPITABLE lpRecipTable = NULL;

	hRes = lpMessage->GetRecipientTable(fMapiUnicode, &lpRecipTable);
	if (SUCCEEDED(hRes) && lpRecipTable)
	{
		LPSRowSet pRecips = NULL;

		SRestriction rgRestrictions[NUM_RECIP_RES] = {0};

		SPropValue spvAdrType = {0};
		SPropValue spvResponsibility = {0};

		// Must be MRXP address type
		rgRestrictions[0].rt = RES_CONTENT;
		rgRestrictions[0].res.resContent.ulPropTag = PR_ADDRTYPE;
		rgRestrictions[0].res.resContent.ulFuzzyLevel = FL_FULLSTRING | FL_IGNORECASE;
		rgRestrictions[0].res.resContent.lpProp = &spvAdrType;

		spvAdrType.ulPropTag = PR_ADDRTYPE;
		spvAdrType.Value.LPSZ = MRXP_ADDRTYPE;

		// Must not have already been poached by another transport
		rgRestrictions[1].rt = RES_PROPERTY;
		rgRestrictions[1].res.resProperty.ulPropTag = PR_RESPONSIBILITY;
		rgRestrictions[1].res.resProperty.relop = RELOP_EQ;
		rgRestrictions[1].res.resProperty.lpProp = &spvResponsibility;

		spvResponsibility.ulPropTag = PR_RESPONSIBILITY;
		spvResponsibility.Value.b = FALSE;

		SRestriction sresRealMRXPRecips = {0};
		sresRealMRXPRecips.rt = RES_AND;
		sresRealMRXPRecips.res.resAnd.cRes = NUM_RECIP_RES;
		sresRealMRXPRecips.res.resAnd.lpRes = rgRestrictions;

		hRes = HrQueryAllRows(lpRecipTable, (LPSPropTagArray)&sptRecipProps,
			&sresRealMRXPRecips, NULL, 0, &pRecips);
		if (SUCCEEDED(hRes) && pRecips)
		{
			Log(true, _T("CXPLogon::GetRecipientsToProcess: %d %s recipients to process\n"), pRecips->cRows, MRXP_ADDRTYPE);
			if (pRecips->cRows > 0)
			{
				LPADRLIST pAdrList = NULL;

				hRes = MyAllocateBuffer(CbNewADRLIST(pRecips->cRows), (LPVOID*)&pAdrList);
				if (SUCCEEDED(hRes) && pAdrList)
				{
					pAdrList->cEntries = pRecips->cRows;

					for (ULONG i = 0; i < pRecips->cRows; i++)
					{
						pAdrList->aEntries[i].cValues = NUM_RECIP_PROPERTIES;
						pAdrList->aEntries[i].rgPropVals = pRecips->aRow[i].lpProps;

						// Take over the SPropValue array from the SRowSet
						// So it won't be freed when the provider frees the SRowSet
						pRecips->aRow[i].lpProps = NULL;
					}

					*lppAdrList = pAdrList;
				}
				else
				{
					Log(true, _T("CXPLogon::GetRecipientsToProcess: MAPIAllocateBuffer returned an error: 0x%08X\n"), hRes);
				}
			}
			else
			{
				hRes = MAPI_E_NOT_ME;
			}
		}
		else
		{
			Log(true, _T("CXPLogon::GetRecipientsToProcess: QueryRows returned an error: 0x%08X\n"), hRes);
		}

		FreeProws(pRecips);
	}
	else
	{
		Log(true, _T("CXPLogon::GetRecipientsToProcess: GetRecipientTable returned an error: 0x%08X\n"), hRes);
	}

	if (lpRecipTable)
		lpRecipTable->Release();

	return hRes;
}

/***********************************************************************************************
	CXPLogon::StripDupsAndP1

	Called from GetRecipientsToProcess to remove any P1 or duplicate recipients.

***********************************************************************************************/

HRESULT CXPLogon::StripDupsAndP1(LPMESSAGE lpMessage)
{
	Log(true, _T("CXPLogon::StripDupsAndP1 function called\n"));

	if (!lpMessage)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	LPMAPITABLE lpRecipTable = NULL;

	hRes = lpMessage->GetRecipientTable(fMapiUnicode, &lpRecipTable);
	if (SUCCEEDED(hRes) && lpRecipTable)
	{
		SRestriction sresMRXPType = {0};
		SPropValue spvAdrType = {0};

		sresMRXPType.rt = RES_CONTENT;
		sresMRXPType.res.resContent.ulPropTag = PR_ADDRTYPE;
		sresMRXPType.res.resContent.ulFuzzyLevel = FL_FULLSTRING | FL_IGNORECASE;
		sresMRXPType.res.resContent.lpProp = &spvAdrType;

		spvAdrType.ulPropTag = PR_ADDRTYPE;
		spvAdrType.Value.LPSZ = MRXP_ADDRTYPE;

		LPSRowSet pRecips = NULL;

		hRes = HrQueryAllRows(lpRecipTable, (LPSPropTagArray)&sptRecipProps,
			&sresMRXPType, NULL, 0, &pRecips);
		if (SUCCEEDED(hRes) && pRecips && pRecips->cRows)
		{
			ULONG i = 0;

			// Quick loop to clear out the RECIP_REPORT_TEXT values so the provider can
			// use them
			for (i = 0; i < pRecips->cRows; i++)
			{
				pRecips->aRow[i].lpProps[RECIP_REPORT_TEXT].Value.ul = 0;
			}

			for (i = 0; i < pRecips->cRows; i++)
			{
				if (RECIP_DELETE != pRecips->aRow[i].lpProps[RECIP_REPORT_TEXT].Value.ul)
				{
					// To consolidate this so the provider only needs to make one call to
					// ModifyRecipients, use the extra recip prop placeholder
					// for PR_REPORT_TEXT (which right now is just PR_NULL).

					// Use PR_EMAIL_ADDRESS as the duplicate detector
					ULONG j = 0;
					int iRet = 0;
					for (j = i+1; j < pRecips->cRows; j++)
					{
						iRet = lstrcmpi(pRecips->aRow[i].lpProps[RECIP_EMAIL_ADDRESS].Value.LPSZ,
							pRecips->aRow[j].lpProps[RECIP_EMAIL_ADDRESS].Value.LPSZ);
						if (0 == iRet)
						{
							// Duplicate
							if (MAPI_P1 == pRecips->aRow[i].lpProps[RECIP_RECIPIENT_TYPE].Value.ul)
							{
								// Remove this one, update other one
								pRecips->aRow[i].lpProps[RECIP_REPORT_TEXT].Value.ul = RECIP_DELETE;
								pRecips->aRow[j].lpProps[RECIP_REPORT_TEXT].Value.ul = RECIP_UPDATE_RESPONSIBILITY;
							}
							else if (MAPI_P1 == pRecips->aRow[j].lpProps[RECIP_RECIPIENT_TYPE].Value.ul)
							{
								pRecips->aRow[j].lpProps[RECIP_REPORT_TEXT].Value.ul = RECIP_DELETE;
								pRecips->aRow[i].lpProps[RECIP_REPORT_TEXT].Value.ul = RECIP_UPDATE_RESPONSIBILITY;
							}
							else
							{
								pRecips->aRow[j].lpProps[RECIP_REPORT_TEXT].Value.ul = RECIP_DELETE;
							}
						}
					}
				}
			}

			// Now the provider knows which ones to remove and update
			// Loop
			ULONG cDelete = 0;
			ULONG cModify = 0;
			for (i = 0; i < pRecips->cRows; i++)
			{
				if (RECIP_DELETE == pRecips->aRow[i].lpProps[RECIP_REPORT_TEXT].Value.ul)
					cDelete++;
				else if (RECIP_UPDATE_RESPONSIBILITY == pRecips->aRow[i].lpProps[RECIP_REPORT_TEXT].Value.ul)
					cModify++;
			}

			LPADRLIST lpDeleteRecips = NULL;
			LPADRLIST lpModifyRecips = NULL;

			if (cDelete > 0)
			{
				hRes = MyAllocateBuffer(CbNewADRLIST(cDelete), (LPVOID*)&lpDeleteRecips);
				if (SUCCEEDED(hRes) && lpDeleteRecips)
				{
					ZeroMemory((LPVOID)lpDeleteRecips, CbNewADRLIST(cDelete));
					lpDeleteRecips->cEntries = cDelete;
				}
				else
				{
					Log(true, _T("CXPLogon::StripDupsAndP1: MyAllocateBuffer returned an error: 0x%08X\n"), hRes);
				}
			}

			if (cModify > 0)
			{
				hRes = MyAllocateBuffer(CbNewADRLIST(cModify), (LPVOID*)&lpModifyRecips);
				if (SUCCEEDED(hRes) && lpModifyRecips)
				{
					ZeroMemory((LPVOID)lpModifyRecips, CbNewADRLIST(cModify));
					lpModifyRecips->cEntries = cModify;
				}
				else
				{
					Log(true, _T("CXPLogon::StripDupsAndP1: MyAllocateBuffer returned an error: 0x%08X\n"), hRes);
				}
			}

			if ((cDelete > 0 && !lpDeleteRecips) || (cModify > 0 && !lpModifyRecips))
			{
				hRes = MAPI_E_NOT_ENOUGH_MEMORY;
			}

			if (SUCCEEDED(hRes) && (cDelete > 0 || cModify > 0))
			{
				ULONG curDelete = 0;
				ULONG curModify = 0;

				for (i = 0; i < pRecips->cRows; i++)
				{
					if (RECIP_DELETE == pRecips->aRow[i].lpProps[RECIP_REPORT_TEXT].Value.ul)
					{
						lpDeleteRecips->aEntries[curDelete].cValues = pRecips->aRow[i].cValues;
						lpDeleteRecips->aEntries[curDelete].rgPropVals = pRecips->aRow[i].lpProps;

						// Null this out to avoid a double free
						pRecips->aRow[i].lpProps = NULL;

						curDelete++;
					}
					else if (RECIP_UPDATE_RESPONSIBILITY == pRecips->aRow[i].lpProps[RECIP_REPORT_TEXT].Value.ul)
					{
						lpModifyRecips->aEntries[curModify].cValues = pRecips->aRow[i].cValues;
						lpModifyRecips->aEntries[curModify].rgPropVals = pRecips->aRow[i].lpProps;

						// Null this out to avoid a double free
						pRecips->aRow[i].lpProps = NULL;

						// Reset PR_RESPONSIBILITY
						lpModifyRecips->aEntries[curModify].rgPropVals[RECIP_RESPONSBILITY].Value.b = false;
						// Clear out any extra data there is here...
						lpModifyRecips->aEntries[curModify].rgPropVals[RECIP_REPORT_TEXT].Value.ul = 0;

						curModify++;
					}
				}

				if (lpDeleteRecips)
				{
					hRes = lpMessage->ModifyRecipients(MODRECIP_REMOVE, lpDeleteRecips);
					if (FAILED(hRes))
					{
						Log(true, _T("CXPLogon::StripDupsAndP1: ModifyRecipients(REMOVE) returned an error: 0x%08X\n"), hRes);
					}
				}

				if (lpModifyRecips)
				{
					hRes = lpMessage->ModifyRecipients(MODRECIP_MODIFY, lpModifyRecips);
					if (FAILED(hRes))
					{
						Log(true, _T("CXPLogon::StripDupsAndP1: ModifyRecipients(MODIFY) returned an error: 0x%08X\n"), hRes);
					}
				}

				// Deferring saving changes here until the end of SubmitMessage
			}

			FreePadrlist(lpDeleteRecips);
			FreePadrlist(lpModifyRecips);
		}
		else
		{
			Log(true, _T("CXPLogon::StripDupsAndP1: HrQueryAllRows returned an error: 0x%08X\n"), hRes);
		}

		FreeProws(pRecips);
	}
	else
	{
		Log(true, _T("CXPLogon::StripDupsAndP1: GetRecipientTable returned an error: 0x%08X\n"), hRes);
	}

	if (lpRecipTable)
		lpRecipTable->Release();

	return hRes;
}

/***********************************************************************************************
	CXPLogon::NewMailWaiting

	To check if there is new mail waiting in the inbox, and optionally to return
	back the full path to the file and the file creation time.

***********************************************************************************************/

HRESULT CXPLogon::NewMailWaiting(LPTSTR* ppszNewMailFile, LPFILETIME pDeliveredTime)
{
	Log(true, _T("CXPLogon::NewMailWaiting function called\n"));

	if (!m_pProps)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	if (!m_pszFileFilter)
	{
		size_t cbFileFilter = 0;

		if (!m_cchInboxPath)
		{
			hRes = StringCchLengthA(m_pProps[CLIENT_INBOX].Value.lpszA, STRSAFE_MAX_CCH, &m_cchInboxPath);
		}

		if (SUCCEEDED(hRes) && m_cchInboxPath)
		{
			cbFileFilter = (m_cchInboxPath + CCH(MSG_FILTER) + 1) * sizeof(TCHAR);

			hRes = MyAllocateBuffer((ULONG) cbFileFilter, (LPVOID*)&m_pszFileFilter);
			if (SUCCEEDED(hRes) && m_pszFileFilter)
			{
				hRes = StringCbPrintf(m_pszFileFilter, cbFileFilter, _T("%hs\\%s"),
					m_pProps[CLIENT_INBOX].Value.lpszA, MSG_FILTER);
			}
		}
	}

	if (!m_pszFileFilter)
		return MAPI_E_NOT_ENOUGH_MEMORY;

	HANDLE hFindFiles;
	WIN32_FIND_DATA FindData = {0};

	hFindFiles = FindFirstFile(m_pszFileFilter, &FindData);

	if (hFindFiles == INVALID_HANDLE_VALUE)
		hRes = S_FALSE;
	else
	{
		hRes = S_OK;

		if (ppszNewMailFile)
		{
			// Caller wants the file name back.
			size_t cbFileName = 0;

			hRes = StringCbLength(FindData.cFileName, STRSAFE_MAX_CCH * sizeof(TCHAR), &cbFileName);
			if (SUCCEEDED(hRes) && cbFileName)
			{
				// +2 for the Null terminator and the '\' between
				cbFileName += (m_cchInboxPath + 2) * sizeof(TCHAR);
				hRes = MyAllocateBuffer((ULONG) cbFileName, (LPVOID*)ppszNewMailFile);
				if (SUCCEEDED(hRes) && *ppszNewMailFile)
				{
					hRes = StringCbPrintf(*ppszNewMailFile, cbFileName, _T("%hs\\%s"),
						m_pProps[CLIENT_INBOX].Value.lpszA, FindData.cFileName);
				}
			}
		}

		if (pDeliveredTime)
		{
			// Caller wants the file time back
			*pDeliveredTime = FindData.ftCreationTime;
		}

		FindClose(hFindFiles);
	}

	return hRes;
}

/***********************************************************************************************
	CXPLogon::LoadMSGToMessage

	To generate an in-memory LPMESSAGE from an MSG file on disk.

***********************************************************************************************/

HRESULT CXPLogon::LoadMSGToMessage(LPCTSTR szMessageFile, LPMESSAGE* lppMessage)
{
	HRESULT		hRes = S_OK;
	LPSTORAGE	pStorage = NULL;
	LPMALLOC	lpMalloc = NULL;

	// get memory allocation function
	lpMalloc = MAPIGetDefaultMalloc();

	if (lpMalloc)
	{
		// Open the file
#ifdef _UNICODE
		hRes = ::StgOpenStorage(szMessageFile, NULL, STGM_READ | STGM_TRANSACTED,
			NULL, 0, &pStorage);
#else
		{
			LPWSTR	szWideCharStr = NULL;
			hRes = AnsiToUnicode(szMessageFile, &szWideCharStr);
			if (SUCCEEDED(hRes) && szWideCharStr)
			{
				hRes = ::StgOpenStorage(szWideCharStr, NULL, STGM_READ | STGM_TRANSACTED,
					NULL, 0, &pStorage);
				MAPIFreeBuffer(szWideCharStr);
			}
		}
#endif

		if (SUCCEEDED(hRes) && pStorage)
		{
			// Open an IMessage interface on an IStorage object
			hRes = OpenIMsgOnIStg(NULL,
				m_pAllocBuffer,
				m_pAllocMore,
				m_pFreeBuffer,
				lpMalloc,
				NULL,
				pStorage,
				NULL,
				0,
				0,
				lppMessage);
		}
	}

	if (pStorage)
		pStorage->Release();

	return hRes;
}

/***********************************************************************************************
	CXPLogon::LoadFromMSG

	To copy the properties from an MSG file to the passed in LPMESSAGE

***********************************************************************************************/

HRESULT CXPLogon::LoadFromMSG(LPCTSTR szMessageFile, LPMESSAGE lpMessage)
{
	HRESULT				hRes = S_OK;
	LPMESSAGE			pIMsg = NULL;

	hRes = LoadMSGToMessage(szMessageFile,&pIMsg);
	if (SUCCEEDED(hRes) && pIMsg)
	{
		hRes = pIMsg->CopyTo(0, NULL, (LPSPropTagArray)&excludeTags, NULL,
			NULL, &IID_IMessage, lpMessage, 0, NULL);

		// Caller will save changes if needed
	}

	if (pIMsg)
		pIMsg->Release();

	return hRes;
}

/***********************************************************************************************
	CXPLogon::SaveToMSG

	To copy the properties from the passed in LPMESSAGE to a new MSG file

***********************************************************************************************/

HRESULT CXPLogon::SaveToMSG(LPMESSAGE lpMessage, LPCTSTR szFileName)
{
	Log(true, _T("CXPLogon::SaveToMSG function called\n"));

	if (!lpMessage || !szFileName)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	LPMALLOC pMalloc = NULL;

	// get memory allocation function
	pMalloc = MAPIGetDefaultMalloc();

	if (pMalloc)
	{
		LPSTORAGE pStorage = NULL;
		STGOPTIONS myOpts = {0};

		myOpts.usVersion = 1;
		myOpts.ulSectorSize = 4096;

#ifdef _UNICODE
		hRes = ::StgCreateStorageEx(szFileName,
			STGM_READWRITE | STGM_TRANSACTED | STGM_CREATE,
			STGFMT_DOCFILE, 0, &myOpts, 0, __uuidof(IStorage), (LPVOID*)&pStorage);
#else
		{
			// Convert new file name to WideChar
			LPWSTR	lpWideCharStr = NULL;
			hRes = AnsiToUnicode(szFileName, &lpWideCharStr);
			if (SUCCEEDED(hRes) && lpWideCharStr)
			{
				hRes = ::StgCreateStorageEx(lpWideCharStr,
					STGM_READWRITE | STGM_TRANSACTED | STGM_CREATE,
					STGFMT_DOCFILE, 0, &myOpts, 0, __uuidof(IStorage), (LPVOID*)&pStorage);
				MyFreeBuffer(lpWideCharStr);
			}
		}
#endif
		if (SUCCEEDED(hRes) && pStorage)
		{
			LPMSGSESS pMsgSession =  NULL;
			// Open an IMessage session.
			hRes = ::OpenIMsgSession(pMalloc, 0, &pMsgSession);

			if (SUCCEEDED(hRes) && pMsgSession)
			{
				LPMESSAGE pIMsg = NULL;

				// Open an IMessage interface on an IStorage object
				hRes = ::OpenIMsgOnIStg(pMsgSession, m_pAllocBuffer, m_pAllocMore,
					m_pFreeBuffer, pMalloc, NULL, pStorage, NULL, 0, 0, &pIMsg);

				if (SUCCEEDED(hRes) && pIMsg)
				{
					// write the CLSID to the IStorage instance - pStorage. This will
					// only work with clients that support this compound document type
					// as the storage medium. If the client does not support
					// CLSID_MailMessage as the compound document, you will have to use
					// the CLSID that it does support.
					hRes = WriteClassStg(pStorage, CLSID_MailMessage);
					if (SUCCEEDED(hRes))
					{
						// copy message properties to IMessage object opened on top of IStorage.
						hRes = lpMessage->CopyTo(0, NULL, (LPSPropTagArray)&excludeSaveTags,
							NULL, NULL, &IID_IMessage, pIMsg, 0, NULL);
						if (SUCCEEDED(hRes))
						{
							// save changes to IMessage object.
							hRes = pIMsg->SaveChanges(KEEP_OPEN_READWRITE);
							if (SUCCEEDED(hRes))
							{
								// save changes in storage of new doc file
								hRes = pStorage->Commit(STGC_DEFAULT);
								if (FAILED(hRes))
									Log(true, _T("CXPLogon::SaveToMSG: Commit returned an error: 0x%08X\n"), hRes);
							}
							else
							{
								Log(true, _T("CXPLogon::SaveToMSG: SaveChanges returned an error: 0x%08X\n"), hRes);
							}
						}
						else
						{
							Log(true, _T("CXPLogon::SaveToMSG: CopyTo returned an error: 0x%08X\n"), hRes);
						}
					}
					else
					{
						Log(true, _T("CXPLogon::SaveToMSG: WriteClassStg returned an error: 0x%08X\n"), hRes);
					}
				}
				else
				{
					Log(true, _T("CXPLogon::SaveToMSG: OpenIMsgOnIStg returned an error: 0x%08X\n"), hRes);
				}

				if (pIMsg) pIMsg->Release();
			}
			else
			{
				Log(true, _T("CXPLogon::SaveToMSG: OpenIMsgSession returned an error: 0x%08X\n"), hRes);
			}

			if (pMsgSession) CloseIMsgSession(pMsgSession);
		}
		else
		{
			Log(true, _T("CXPLogon::SaveToMSG: StgCreateStorageEx returned an error: 0x%08X\n"), hRes);
		}

		if (pStorage) pStorage->Release();
	}

	return hRes;
}

/***********************************************************************************************
	CXPLogon::StampSender

	To stamp the PR_SENDER and PR_SENT_REPRESENTING* props on a message.  This uses
	the properties generated by the GenerateIdentity function.  This is called from
	SubmitMessage.

***********************************************************************************************/

HRESULT CXPLogon::StampSender(LPMESSAGE lpMessage)
{
	Log(true, _T("CXPLogon::StampSender function called\n"));
	HRESULT hRes = S_OK;

	// Set PR_SENDER* props
	if (m_pIdentityProps)
	{
		hRes = lpMessage->SetProps(NUM_IDENTITY_PROPS, m_pIdentityProps, NULL);
		if (SUCCEEDED(hRes))
		{
			// Now the PR_SENT_REPRESENTING* props
			// Keep the props if the client set them
			// otherwise, use the same values used for
			// the PR_SENDER* props

			ULONG cValues = 0;
			LPSPropValue lpSentRepresentingProps = NULL;

			hRes = lpMessage->GetProps((LPSPropTagArray)&sentRepTags, fMapiUnicode, &cValues,
				&lpSentRepresentingProps);
			if (SUCCEEDED(hRes) && cValues && lpSentRepresentingProps)
			{
				// Validate the props and overwrite if needed
				if (PR_SENT_REPRESENTING_NAME != lpSentRepresentingProps[SENT_REP_NAME].ulPropTag ||
					PR_SENT_REPRESENTING_SEARCH_KEY != lpSentRepresentingProps[SENT_REP_SEARCH_KEY].ulPropTag ||
					PR_SENT_REPRESENTING_ENTRYID != lpSentRepresentingProps[SENT_REP_ENTRYID].ulPropTag ||
					PR_SENT_REPRESENTING_EMAIL_ADDRESS != lpSentRepresentingProps[SENT_REP_EMAIL_ADDRESS].ulPropTag ||
					PR_SENT_REPRESENTING_ADDRTYPE != lpSentRepresentingProps[SENT_REP_ADDRTYPE].ulPropTag)
				{
					// No valid props from the original, use the PR_SENDER* values
					SPropValue spvSentRepresenting[NUM_IDENTITY_PROPS] = {0};

					spvSentRepresenting[IDENTITY_DISPLAY_NAME].ulPropTag = PR_SENT_REPRESENTING_NAME_A;
					spvSentRepresenting[IDENTITY_DISPLAY_NAME].Value = m_pIdentityProps[IDENTITY_DISPLAY_NAME].Value;

					spvSentRepresenting[IDENTITY_SEARCH_KEY].ulPropTag = PR_SENT_REPRESENTING_SEARCH_KEY;
					spvSentRepresenting[IDENTITY_SEARCH_KEY].Value = m_pIdentityProps[IDENTITY_SEARCH_KEY].Value;

					spvSentRepresenting[IDENTITY_ENTRYID].ulPropTag = PR_SENT_REPRESENTING_ENTRYID;
					spvSentRepresenting[IDENTITY_ENTRYID].Value = m_pIdentityProps[IDENTITY_ENTRYID].Value;

					spvSentRepresenting[IDENTITY_EMAIL_ADDRESS].ulPropTag = PR_SENT_REPRESENTING_EMAIL_ADDRESS_A;
					spvSentRepresenting[IDENTITY_EMAIL_ADDRESS].Value = m_pIdentityProps[IDENTITY_EMAIL_ADDRESS].Value;

					spvSentRepresenting[IDENTITY_ADDRTYPE].ulPropTag = PR_SENT_REPRESENTING_ADDRTYPE_A;
					spvSentRepresenting[IDENTITY_ADDRTYPE].Value = m_pIdentityProps[IDENTITY_ADDRTYPE].Value;

					hRes = lpMessage->SetProps(NUM_IDENTITY_PROPS, spvSentRepresenting, NULL);
					if (FAILED(hRes))
					{
						Log(true, _T("CXPLogon::StampSender: SetProps on sent representing props returned an error: 0x%08X\n"), hRes);
					}
				}
			}
			else
			{
				Log(true, _T("CXPLogon::StampSender: GetProps on sent representing props for original values returned an error: 0x%08X\n"), hRes);
			}

			// Defer SaveChanges until the end of SubmitMessage

			MyFreeBuffer(lpSentRepresentingProps);
		}
		else
		{
			Log(true, _T("CXPLogon::StampSender: SetProps on sender props returned an error: 0x%08X\n"), hRes);
		}
	}

	return hRes;
}

/***********************************************************************************************
	CXPLogon::SetNDRText

	To load the proper NDR string from the string table, then set this in the
	"bad" recipient's row.

***********************************************************************************************/

HRESULT CXPLogon::SetNDRText(ADRENTRY adrRecip, HRESULT hResSendError, LPVOID pParent)
{
	Log(true, _T("CXPLogon::SetNDRText function called\n"));

	HRESULT hRes = S_OK;
	UINT uidNDRString = 0;
	int iRet = 0;
	TCHAR szTempNDRText[DEF_NDR_TEXT_LEN] = {0};

	switch (HRESULT_CODE(hResSendError))
	{
	case ERROR_BAD_NETPATH:
		uidNDRString = IDS_NDR_BAD_NETPATH;
		break;
	default:
		uidNDRString = IDS_NDR_DEFAULT_TEXT;
		break;
	}

	adrRecip.rgPropVals[RECIP_REPORT_TEXT].ulPropTag = PR_REPORT_TEXT;

	iRet = LoadString(m_pProvider->GetHINST(), uidNDRString, szTempNDRText,
		DEF_NDR_TEXT_LEN);
	if (iRet >= 0)
	{
		size_t cbLen = 0;

		hRes = StringCbLength(szTempNDRText, STRSAFE_MAX_CCH * sizeof(TCHAR), &cbLen);
		if (SUCCEEDED(hRes) && cbLen)
		{
			LPTSTR szNDRText = NULL;

			cbLen += NDR_ERROR_PAD;

			hRes = MyAllocateMore((ULONG) cbLen, pParent, (LPVOID*)&szNDRText);
			if (SUCCEEDED(hRes) && szNDRText)
			{
				hRes = StringCbPrintf(szNDRText, cbLen,
					_T("%s (0x%08X)"), szTempNDRText, hResSendError);
				if (SUCCEEDED(hRes))
				{
					adrRecip.rgPropVals[RECIP_REPORT_TEXT].Value.LPSZ = szNDRText;
				}
				else
				{
					Log(true, _T("CXPLogon::SetNDRText: StringCchPrintf returned an error: 0x%08X\n"), hRes);
					adrRecip.rgPropVals[RECIP_REPORT_TEXT].Value.LPSZ = ERROR_LOADING_NDR_TEXT;
				}
			}
			else
			{
				Log(true, _T("CXPLogon::SetNDRText: MyAllocateBuffer returned an error: 0x%08X\n"), hRes);
				adrRecip.rgPropVals[RECIP_REPORT_TEXT].Value.LPSZ = ERROR_LOADING_NDR_TEXT;
			}
		}
		else
		{
			Log(true, _T("CXPLogon::SetNDRText: StringCbLength returned an error: 0x%08X\n"), hRes);
			adrRecip.rgPropVals[RECIP_REPORT_TEXT].Value.LPSZ = ERROR_LOADING_NDR_TEXT;
		}
	}
	else
	{
		Log(true, _T("CXPLogon::SetNDRText: LoadString returned an error: 0x%08X\n"), GetLastError());
		adrRecip.rgPropVals[RECIP_REPORT_TEXT].Value.LPSZ = ERROR_LOADING_NDR_TEXT;
	}

	return hRes;
}

/***********************************************************************************************
	CXPLogon::GenerateBadRecipList

	Scans the passed in LPADRLIST and builds a new LPADRLIST with just the NDR'ed
	recipients.  This is called from SubmitMessage to build a list to pass to StatusRecips.

***********************************************************************************************/

HRESULT CXPLogon::GenerateBadRecipList(LPADRLIST lpOriginal, LPADRLIST* lppBadRecips)
{
	Log(true, _T("CXPLogon::GenerateBadRecipList function called\n"));

	if (!lpOriginal || !lppBadRecips)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	ULONG i = 0;
	ULONG cBad = 0;

	*lppBadRecips = NULL;

	// Figure out how many recipients to allocate
	for (i = 0; i < lpOriginal->cEntries; i++)
	{
		if (lpOriginal->aEntries[i].rgPropVals[RECIP_REPORT_TEXT].ulPropTag == PR_REPORT_TEXT)
		{
			cBad++;
		}
	}

	if (cBad > 0)
	{
		hRes = MyAllocateBuffer(CbNewADRLIST(cBad), (LPVOID*)lppBadRecips);
		if (SUCCEEDED(hRes) && *lppBadRecips)
		{
			ZeroMemory(*lppBadRecips, CbNewADRLIST(cBad));

			(*lppBadRecips)->cEntries = cBad;
			ULONG curBad = 0;

			// Loop through again and copy bad recips
			// Also clear out PR_REPORT_TEXT on the original
			for (i = 0; i < lpOriginal->cEntries; i++)
			{
				if (lpOriginal->aEntries[i].rgPropVals[RECIP_REPORT_TEXT].ulPropTag == PR_REPORT_TEXT)
				{
					ULONG cProps = lpOriginal->aEntries[i].cValues;

					(*lppBadRecips)->aEntries[curBad].cValues = cProps;

					hRes = CopySPropValueArray(cProps, lpOriginal->aEntries[i].rgPropVals,
						&((*lppBadRecips)->aEntries[curBad].rgPropVals), (LPVOID)*lppBadRecips);
					if (SUCCEEDED(hRes) && (*lppBadRecips)->aEntries[curBad].rgPropVals)
					{
						// Clear this out so it won't get set on the copy
						// in Sent Items
						// Note: this *looks* like a leak, since the buffer was allocated
						// for the NDR text.  However, it was alloced with MAPIAllocateMore,
						// so it is linked to the ADRLIST.  When it gets released, MAPI will
						// release that memory too.
						lpOriginal->aEntries[i].rgPropVals[RECIP_REPORT_TEXT].ulPropTag = PR_NULL;
						lpOriginal->aEntries[i].rgPropVals[RECIP_REPORT_TEXT].Value.LPSZ = _T("");
					}
					else
					{
						Log(true, _T("CXPLogon::GenerateBadRecipList: CopySPropValueArray returned an error: 0x%08X\n"), hRes);
					}

					curBad++;
				}
			}
		}
		else
		{
			Log(true, _T("CXPLogon::GenerateBadRecipList: MyAllocateBuffer returned an error: 0x%08X\n"), hRes);
		}
	}

	return hRes;
}

/***********************************************************************************************
	CXPLogon::CopySPropValueArray

	To copy an SPropValueArray.  Handles allocating memory if needed.  Note that this
	only handles PT_BOOLEAN, PT_LONG, PT_STRING8, and PT_UNICODE.

***********************************************************************************************/

HRESULT CXPLogon::CopySPropValueArray(ULONG cProps, LPSPropValue lpSourceProps, LPSPropValue* lppTargetProps, LPVOID lpParent)
{
	if (!cProps || !lpSourceProps || !lppTargetProps)
		return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	*lppTargetProps = NULL;

	hRes = MyAllocateMore(cProps * sizeof(SPropValue), lpParent, (LPVOID*)lppTargetProps);
	if (SUCCEEDED(hRes) && *lppTargetProps)
	{
		ULONG i = 0;

		for (i = 0; i < cProps; i++)
		{
			(*lppTargetProps)[i].ulPropTag = lpSourceProps[i].ulPropTag;

			switch(PROP_TYPE(lpSourceProps[i].ulPropTag))
			{
			case PT_BOOLEAN:
				(*lppTargetProps)[i].Value.b = lpSourceProps[i].Value.b;
				break;
			case PT_LONG:
				(*lppTargetProps)[i].Value.ul = lpSourceProps[i].Value.ul;
				break;
			case PT_STRING8:
				{
					size_t cbString = 0;
					hRes = StringCbLengthA(lpSourceProps[i].Value.lpszA, STRSAFE_MAX_CCH * sizeof(char),
						&cbString);
					if (SUCCEEDED(hRes) && cbString)
					{
						cbString++;

						hRes = MyAllocateMore((ULONG) cbString, lpParent, (LPVOID*)&((*lppTargetProps)[i].Value.lpszA));
						if (SUCCEEDED(hRes) && (*lppTargetProps)[i].Value.lpszA)
						{
							hRes = StringCbCopyA((*lppTargetProps)[i].Value.lpszA, cbString,
								lpSourceProps[i].Value.lpszA);
						}
					}
				}
				break;
			case PT_UNICODE:
				{
					size_t cbString = 0;
					hRes = StringCbLengthW(lpSourceProps[i].Value.lpszW, STRSAFE_MAX_CCH * sizeof(WCHAR),
						&cbString);
					if (SUCCEEDED(hRes) && cbString)
					{
						cbString += sizeof(WCHAR);

						hRes = MyAllocateMore((ULONG) cbString, lpParent, (LPVOID*)&((*lppTargetProps)[i].Value.lpszW));
						if (SUCCEEDED(hRes) && (*lppTargetProps)[i].Value.lpszW)
						{
							hRes = StringCbCopyW((*lppTargetProps)[i].Value.lpszW, cbString,
								lpSourceProps[i].Value.lpszW);
						}
					}
				}
				break;
			default:
				// Note: this only handles string types, PT_LONG, and PT_BOOLEAN
				break;
			}

			if (FAILED(hRes))
				break;
		}
	}

	return hRes;
}

/***********************************************************************************************
	CXPLogon::MyAllocateBuffer

	A wrapper around the provider's member pointer to MAPIAllocateBuffer. Makes sure that
	the function pointer is non-null before calling it.

***********************************************************************************************/

SCODE STDMETHODCALLTYPE CXPLogon::MyAllocateBuffer(ULONG			cbSize,
												   LPVOID FAR *		lppBuffer)
{
	if (m_pAllocBuffer) return m_pAllocBuffer(cbSize,lppBuffer);

	Log(true, _T("CXPLogon::MyAllocateBuffer: m_pAllocBuffer not set!\n"));
	return MAPI_E_INVALID_PARAMETER;
}

/***********************************************************************************************
	CXPLogon::MyAllocateMore

	A wrapper around the provider's member pointer to MAPIAllocateMore. Makes sure that
	the function pointer is non-null before calling it.

***********************************************************************************************/

SCODE STDMETHODCALLTYPE CXPLogon::MyAllocateMore(ULONG			cbSize,
												 LPVOID			lpObject,
												 LPVOID FAR *	lppBuffer)
{
	if (m_pAllocMore) return m_pAllocMore(cbSize,lpObject,lppBuffer);

	Log(true, _T("CXPLogon::MyAllocMore: m_pAllocMore not set!\n"));
	return MAPI_E_INVALID_PARAMETER;
}

/***********************************************************************************************
	CXPLogon::MyFreeBuffer

	A wrapper around the provider's member pointer to MAPIFreeBuffer. Makes sure that
	the function pointer is non-null before calling it.

***********************************************************************************************/

ULONG STDAPICALLTYPE CXPLogon::MyFreeBuffer(LPVOID			lpBuffer)
{
	if (m_pFreeBuffer) return m_pFreeBuffer(lpBuffer);

	Log(true, _T("CXPLogon::MyFreeBuffer: m_pFreeBuffer not set!\n"));
	return (ULONG) MAPI_E_INVALID_PARAMETER;
}

/***********************************************************************************************
	CXPLogon::AnsiToUnicode

	To convert a passed in ANSI string to a Unicode string.

***********************************************************************************************/

HRESULT CXPLogon::AnsiToUnicode(LPCSTR pszA, LPWSTR* ppszW)
{
	HRESULT	hRes = S_OK;
	size_t	cbszA = 0;
	size_t	cchszW = 0;
	size_t  cbszW = 0;

	*ppszW = NULL;
	if (NULL == pszA) return S_OK;

	// Determine number of wide characters to be allocated for the Unicode string.
	hRes = StringCbLengthA(pszA, STRSAFE_MAX_CCH * sizeof(char), &cbszA);
	if (SUCCEEDED(hRes))
	{
		cbszA += sizeof(char); // NULL terminator

		hRes = StringCchLengthA(pszA,STRSAFE_MAX_CCH,&cchszW);
		if (SUCCEEDED(hRes))
		{
			cchszW++; // NULL terminator
			cbszW = sizeof(WCHAR) * cchszW;

			hRes = MyAllocateBuffer((ULONG) cbszW, (LPVOID*) ppszW);
			if (SUCCEEDED(hRes) && *ppszW)
			{
				int iRet = NULL;
				// Convert to Unicode.
				iRet = MultiByteToWideChar(CP_ACP, 0, pszA, (int) cbszA,
					*ppszW, (int) cchszW);
				if (iRet <= 0)
				{
					MyFreeBuffer(*ppszW);
					*ppszW = NULL;
				}
			}
		}
	}

	return hRes;
}