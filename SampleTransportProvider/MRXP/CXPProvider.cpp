#include "mrxp.h"

CXPProvider::CXPProvider(HINSTANCE hInstance)
{
	Log(true, _T("CXPProvider object created at address 0x%08x\n"), ULONG(this));
	m_pLogonList = NULL;
	m_hInstance = hInstance;
	m_cRef = 1;
	InitializeCriticalSection(&m_csObj);
}

CXPProvider::~CXPProvider()
{
	Log(true, _T("CXPProvider object destroyed at address 0x%08x\n"), ULONG(this));
	m_hInstance = NULL;
	DeleteCriticalSection(&m_csObj);
}

STDMETHODIMP CXPProvider::QueryInterface(REFIID riid, LPVOID* ppvObj)
{
	if (!ppvObj)
		return E_INVALIDARG;

	*ppvObj = NULL;

	if (riid == IID_IXPProvider || riid == IID_IUnknown)
	{
		*ppvObj = (LPVOID)this;
		AddRef();
		return S_OK;
	}

	return E_NOINTERFACE;
}

// BEGIN IXPProvider Implementation

/***********************************************************************************************
	CXPProvider::Shutdown

		- Walk the linked-list of CXPLogon objects, and call TransportLogoff and Release for each

		Refer to the MSDN documentation for more information.

***********************************************************************************************/

STDMETHODIMP CXPProvider::Shutdown(ULONG FAR* /*lpulFlags*/)
{
	Log(true, _T("CXPProvider::Shutdown function called\n"));
	HRESULT hRes = S_OK;

	EnterCriticalSection(&m_csObj);

	CXPLogon* pCurrent = m_pLogonList;

	while (pCurrent)
	{
		CXPLogon* pNext = pCurrent->GetNextLogon();

		hRes = pCurrent->TransportLogoff(0);
		if (FAILED(hRes))
			Log(true, _T("CXPProvider::Shutdown: TransportLogoff returned an error: 0x%08X\n"), hRes);

		pCurrent->Release();

		pCurrent = pNext;
	}

	LeaveCriticalSection(&m_csObj);

	return S_OK;
}

/***********************************************************************************************
	CXPProvider::TransportLogon

		- Call GetMemAllocRoutines from the passed in IMAPISupport object.  These will be
		saved in the provider's new logon object.
		- Call OpenProfileSection from the IMAPISupport object (pass a NULL MAPIUID) to get
		the provider's profile section, then get the provider's config props.  If they're not there, return
		MAPI_E_UNCONFIGURED.
		- Set the provider's status bits accordingly depending on the passed in flags.
		- Register the provider's preprocessor functions
		- Create a status row and register with the MAPI support object using ModifyStatusRow.
		- Put the provider's logon object into the linked list of logon objects
		- Return the new logon object in lppXPLogon, and set lpulFlags accordingly.

		Refer to the MSDN documentation for more information.

***********************************************************************************************/

STDMETHODIMP CXPProvider::TransportLogon(LPMAPISUP			lpMAPISup,
										 ULONG				/*ulUIParam*/,
										 LPTSTR				/*lpszProfileName*/,
										 ULONG FAR*			lpulFlags,
										 LPMAPIERROR FAR*	/*lppMAPIError*/,
										 LPXPLOGON FAR*		lppXPLogon)
{
	Log(true, _T("CXPProvider::TransportLogon function called\n"));

	if (!lpMAPISup || !lppXPLogon)
		return MAPI_E_INVALID_PARAMETER;

	*lppXPLogon = NULL;

	HRESULT hRes = S_OK;

	LPALLOCATEBUFFER lpAllocBuffer = NULL;
	LPALLOCATEMORE lpAllocMore = NULL;
	LPFREEBUFFER lpFreeBuffer = NULL;

	// Get the MAPI memory routines
	hRes = lpMAPISup->GetMemAllocRoutines(&lpAllocBuffer, &lpAllocMore, &lpFreeBuffer);
	if (SUCCEEDED(hRes) && lpAllocBuffer && lpAllocMore && lpFreeBuffer)
	{
		// Get the profile properties
		LPPROFSECT lpProfSect = NULL;

		hRes = lpMAPISup->OpenProfileSection(NULL, MAPI_MODIFY, &lpProfSect);
		if (SUCCEEDED(hRes) && lpProfSect)
		{
			ULONG ulProfProps = 0;
			LPSPropValue pProfProps = NULL;

			// Don't free the returned props here. The new CXPLogon class
			// will take ownership (unless it fails to create)
			hRes = lpProfSect->GetProps((LPSPropTagArray)&sptClientProps, 0,
				&ulProfProps, &pProfProps);
			if (SUCCEEDED(hRes) && ulProfProps && pProfProps)
			{
				// Validate the properties
				hRes = ValidateConfigProps(ulProfProps, pProfProps);
				if (SUCCEEDED(hRes))
				{
					CXPLogon* pXPLogon = NULL;

					pXPLogon = new CXPLogon(this, lpMAPISup, lpAllocBuffer, lpAllocMore, lpFreeBuffer,
						ulProfProps, pProfProps, *lpulFlags);
					if (!pXPLogon)
					{
						Log(true, _T("CXPProvider::TransportLogon: Failed to create new CXPLogon\n"));
						lpFreeBuffer(pProfProps);
						hRes = MAPI_E_NOT_ENOUGH_MEMORY;
					}

					if(SUCCEEDED(hRes))
					{
						if (*lpulFlags & LOGON_NO_CONNECT)
						{
							pXPLogon->AddStatusBits(STATUS_OFFLINE);
						}
						else
						{
							if (!(*lpulFlags & LOGON_NO_INBOUND))
								pXPLogon->AddStatusBits(STATUS_INBOUND_ENABLED);
							if (!(*lpulFlags & LOGON_NO_OUTBOUND))
								pXPLogon->AddStatusBits(STATUS_OUTBOUND_ENABLED);
						}

						hRes = lpMAPISup->RegisterPreprocessor(NULL, (LPTSTR)MRXP_ADDRTYPE_A, (LPTSTR)MRXP_FULLDLLNAME_A,
							"PreprocessMessage", "RemovePreprocessInfo", 0);
						if (SUCCEEDED(hRes))
						{
							EnterCriticalSection(&m_csObj);

							hRes = pXPLogon->CreateStatusRow();
							if (SUCCEEDED(hRes))
							{
								pXPLogon->SetNextLogon(m_pLogonList);
								m_pLogonList = pXPLogon;

								*lppXPLogon = (LPXPLOGON)pXPLogon;
								if (*lpulFlags & LOGON_NO_CONNECT)
								{
									*lpulFlags = 0;
								}
								else
								{
									// TODO: Add a profile prop to enable/disable proactive polling on Idle, check it here
									// If not set, don't pass LOGON_SP_IDLE
									*lpulFlags = LOGON_SP_IDLE | LOGON_SP_RESOLVE | LOGON_SP_POLL;
								}
							}
							else
							{
								Log(true, _T("CXPProvider::TransportLogon: CreateStatusRow returned an error: 0x%08X\n"), hRes);
								delete pXPLogon;
							}

							LeaveCriticalSection(&m_csObj);
						}
						else
						{
							Log(true, _T("CXPProvider::TransportLogon: RegisterPreprocessor returned an error: 0x%08X\n"), hRes);
							delete pXPLogon;
						}
					}
				}
				else
				{
					Log(true, _T("CXPProvider::TransportLogon: ValidateConfigProps returned an error: 0x%08X\n"), hRes);
				}
			}
			else
			{
				Log(true, _T("CXPProvider::TransportLogon: GetProps on lpProfSect returned an error: 0x%08X\n"), hRes);
			}
		}
		else
		{
			Log(true, _T("CXPProvider::TransportLogon: OpenProfileSection returned an error: 0x%08X\n"), hRes);
		}

		if (lpProfSect)
			lpProfSect->Release();
	}
	else
	{
		Log(true, _T("CXPProvider::TransportLogon: GetMemAllocRoutines returned an error: 0x%08X\n"), hRes);
	}

	return hRes;
}

// END IXPProvider Implementation