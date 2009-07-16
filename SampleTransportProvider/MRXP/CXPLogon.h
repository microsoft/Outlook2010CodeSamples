#pragma once

#define DEF_STATUS_ROW_STRING_LEN 128
#define DEF_NDR_TEXT_LEN 512
#define NDR_ERROR_PAD 14 * sizeof(TCHAR)
#define ERROR_LOADING_NDR_TEXT _T("Error loading NDR Text")
#define MSG_FILTER _T("*.MSG")
#define NUM_RECIP_RES 2

class CXPStatus;

class CXPLogon : public IXPLogon
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

	// IXPLogon
	MAPI_IXPLOGON_METHODS(IMPL);

	// Constructor Destructor
	CXPLogon(CXPProvider* pProvider, LPMAPISUP pMAPISup, LPALLOCATEBUFFER pAllocBuffer,
		LPALLOCATEMORE pAllocMore, LPFREEBUFFER pFreeBuffer, ULONG ulProps, LPSPropValue pProps,
		ULONG ulFlags);
	~CXPLogon();

	// Helper functions

	SCODE STDMETHODCALLTYPE MyAllocateBuffer(ULONG			cbSize,
										     LPVOID FAR *	lppBuffer);

	SCODE STDMETHODCALLTYPE MyAllocateMore(ULONG			cbSize,
										   LPVOID			lpObject,
										   LPVOID FAR *		lppBuffer);

	ULONG STDAPICALLTYPE MyFreeBuffer(LPVOID			lpBuffer);

	inline void AddStatusBits(ULONG ulBitsToAdd)
	{m_ulStatusBits |= ulBitsToAdd;}

	inline void RemoveStatusBits(ULONG ulBitsToRemove)
	{m_ulStatusBits &= ~ulBitsToRemove;}

	inline ULONG GetStatusBits()
	{return m_ulStatusBits;}

	inline HINSTANCE GetHINST()
	{return m_pProvider->GetHINST();}

	inline void SetNextLogon(CXPLogon* pNext)
	{m_pNextLogon = pNext;}

	inline CXPLogon* GetNextLogon()
	{return m_pNextLogon;}

	HRESULT GenerateIdentity();

	HRESULT CreateStatusRow();

	UINT GetStatusStringID();

	HRESULT UpdateStatusRow();

	HRESULT GetRecipientsToProcess(LPMESSAGE lpMessage, LPADRLIST* lppAdrList);

	HRESULT StripDupsAndP1(LPMESSAGE lpMessage);

	HRESULT NewMailWaiting(LPTSTR* ppszNewMailFile, LPFILETIME pDeliveredTime);

	HRESULT LoadFromMSG(LPCTSTR szMessageFile, LPMESSAGE lpMessage);

	HRESULT LoadMSGToMessage(LPCTSTR szMessageFile, LPMESSAGE* lppMessage);

	HRESULT SaveToMSG(LPMESSAGE lpMessage, LPCTSTR szFileName);

	HRESULT StampSender(LPMESSAGE lpMessage);

	HRESULT SetNDRText(ADRENTRY adrRecip, HRESULT hResSendError, LPVOID pParent);

	HRESULT GenerateBadRecipList(LPADRLIST lpOriginal, LPADRLIST* lppBadRecips);

	HRESULT CopySPropValueArray(ULONG cProps, LPSPropValue lpSourceProps, LPSPropValue* lppTargetProps, LPVOID lpParent);

	HRESULT AnsiToUnicode(LPCSTR pszA, LPWSTR* ppszW);

private:
	LPTSTR				m_pszFileFilter;
	CXPProvider*		m_pProvider;
	CXPLogon*			m_pNextLogon;
	CXPStatus*			m_pStatusObj;
	LPMAPISUP			m_pMAPISup;
	LPALLOCATEBUFFER	m_pAllocBuffer;
	LPALLOCATEMORE		m_pAllocMore;
	LPFREEBUFFER		m_pFreeBuffer;
	ULONG				m_ulProps;
	LPSPropValue		m_pProps;
	LPSPropValue		m_pIdentityProps;
	size_t				m_cchInboxPath;
	ULONG				m_ulLogonFlags;
	ULONG				m_ulStatusBits;
	ULONG				m_cRef;
	CRITICAL_SECTION	m_csObj;
};