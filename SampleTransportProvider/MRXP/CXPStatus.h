#pragma once

class CXPStatus : public IMAPIStatus
{
public:
	// IUnknown
	STDMETHODIMP QueryInterface(REFIID                     riid,
								LPVOID *                   ppvObj);

    inline STDMETHODIMP_(ULONG) AddRef()
	{
		EnterCriticalSection (&m_csObj);
		++m_cRef;
		ULONG ulCount = m_cRef;
		LeaveCriticalSection (&m_csObj);
		return ulCount;
	}

    inline STDMETHODIMP_(ULONG) Release()
	{
		EnterCriticalSection (&m_csObj);
        ULONG ulCount = --m_cRef;
        LeaveCriticalSection (&m_csObj);
        if (!ulCount)
			delete this;
        return ulCount;
	}

	// IMAPIProp
	MAPI_IMAPIPROP_METHODS(IMPL);
	// IMAPIStatus
	MAPI_IMAPISTATUS_METHODS(IMPL);

	// Constructor Destructor
	CXPStatus(CXPLogon* pParent, LPMAPISUP pMAPISup, ULONG cIdentityProps, LPSPropValue pIdentityProps,
		LPALLOCATEBUFFER pAllocBuffer, LPALLOCATEMORE pAllocMore, LPFREEBUFFER pFreeBuffer);
	~CXPStatus();

	// Helper functions
	SCODE STDMETHODCALLTYPE MyAllocateBuffer(ULONG			cbSize,
										     LPVOID FAR *	lppBuffer);

	SCODE STDMETHODCALLTYPE MyAllocateMore(ULONG			cbSize,
										   LPVOID			lpObject,
										   LPVOID FAR *		lppBuffer);

	ULONG STDAPICALLTYPE MyFreeBuffer(LPVOID			lpBuffer);

	DWORD CheckInboxAvailability(LPSTR lpszInboxPath);

	HRESULT GetConvertedStringProp(ULONG ulRequestedTag, SPropValue spvSource,
		LPSPropValue pspvDestination, LPVOID pParent);

	HRESULT LoadStringAsType(ULONG ulType, UINT ids, LPVOID* ppBuffer, LPVOID pParent);

private:
	CXPLogon*			m_pParentLogon;
	LPMAPISUP			m_pMAPISup;
	ULONG				m_cIdentityProps;
	LPSPropValue		m_pIdentityProps;
	LPALLOCATEBUFFER	m_pAllocBuffer;
	LPALLOCATEMORE		m_pAllocMore;
	LPFREEBUFFER		m_pFreeBuffer;
	ULONG				m_cRef;
	CRITICAL_SECTION	m_csObj;
};