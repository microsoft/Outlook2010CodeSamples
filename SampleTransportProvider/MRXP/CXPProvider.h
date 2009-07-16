#pragma once

class CXPLogon;

class CXPProvider : public IXPProvider
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

	// IXPProvider
	MAPI_IXPPROVIDER_METHODS(IMPL);

	// Constructor/Destructor
	CXPProvider(HINSTANCE hInstance);
	~CXPProvider();

	// Helper functions
	inline HINSTANCE GetHINST()
	{return m_hInstance;}

	inline CXPLogon* GetLogonList()
	{return m_pLogonList;}

	inline void SetLogonList(CXPLogon* pHead)
	{m_pLogonList = pHead;}

	inline LPCRITICAL_SECTION GetCS()
	{return &m_csObj;}

	// Data
private:
	CXPLogon*			m_pLogonList;
	HINSTANCE			m_hInstance;
	ULONG				m_cRef;
	CRITICAL_SECTION	m_csObj;
};