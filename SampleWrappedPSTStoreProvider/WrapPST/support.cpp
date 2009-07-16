#include "stdafx.h"

CSupport::CSupport(LPMAPISUP  pMAPISupObj,
				   LPPROFSECT pProfSect)
{
	Log(true,"CSupport::CSupport\n");
	m_cRef = 1;
	if(FAILED(pMAPISupObj->QueryInterface(IID_IMAPISup , (LPVOID*)&m_pMAPISup)))
	{
		m_pMAPISup = NULL;
	}
	m_lpProfSect = pProfSect;
	if (m_lpProfSect) m_lpProfSect->AddRef();
	InitializeCriticalSection (&m_csObj);
	Log(true,"CSupport::CSupport\n");
}

CSupport::~CSupport()
{
	Log(true,"In Destructor of CSupport\n");
	if (m_pMAPISup) m_pMAPISup->Release();
	if (m_lpProfSect) m_lpProfSect->Release();
	DeleteCriticalSection (&m_csObj);
	Log(true,"CSupport::~CSupport\n");
}

STDMETHODIMP CSupport::ModifyStatusRow
(ULONG                      cValues,
 LPSPropValue               pColumnVals,
 ULONG                      ulFlags)
{
	Log(true,"CSupport::ModifyStatusRow()\n");

	// The PST assumes the identity will be an EX address type
	// If these identity props are included in the status row, QueryIdentity can fail later
	// This failure can cause Outlook not to be able to create contacts, appts, ect.
	// This problem disappears after the first time Outlook loads the provider
	// This is really only a problem if no provider sets STATUS_PRIMARY_IDENTITY
	if (pColumnVals)
	{
		ULONG ulCol = 0;
		for (ulCol = 0 ; ulCol < cValues ; ulCol++)
		{
			switch (pColumnVals[ulCol].ulPropTag)
			{
			case(PR_IDENTITY_ENTRYID):
			case(PR_IDENTITY_DISPLAY_A):
			case(PR_IDENTITY_DISPLAY_W):
			case(PR_IDENTITY_SEARCH_KEY):
				Log(true,"Ignoring 0x%08X set by PST provider\n",pColumnVals[ulCol].ulPropTag);
				pColumnVals[ulCol].ulPropTag = PR_NULL;
				break;
			default:
				break;
			}
		}
	}
	return m_pMAPISup->ModifyStatusRow (cValues,
		pColumnVals,
		ulFlags);
}


STDMETHODIMP CSupport::OpenProfileSection
(LPMAPIUID                  lpUid,
 ULONG                      ulFlags,
 LPPROFSECT *               lppProfileObj)
{
	Log(true,"CSupport::OpenProfileSection\n");
	if (lpUid &&
		IsEqualMAPIUID(lpUid, (void *)&pbNSTGlobalProfileSectionGuid) &&
		m_lpProfSect)
	{
		// Allow the opening of the Global Section
		if (m_lpProfSect)
		{
			*lppProfileObj = m_lpProfSect;
			(*lppProfileObj)->AddRef();
			return S_OK;
		}
	}
	return m_pMAPISup->OpenProfileSection(lpUid, ulFlags, lppProfileObj);
}
