#include <MAPISPI.H>
#include <mspst.h>

#define	pidProfileMin					0x6600
#define PR_PROFILE_OFFLINE_STORE_PATH	PROP_TAG( PT_STRING8, pidProfileMin+0x10)
#define PR_OST_OLFI PROP_TAG(PT_BINARY, 0x7C07)

enum
{
	CLIENT_PST_PATH = 0,
		CLIENT_MDB_PROVIDER,
		NUM_CLIENT_PROPERTIES
};

const static SizedSPropTagArray(NUM_CLIENT_PROPERTIES, sptClientProps) =
{
	NUM_CLIENT_PROPERTIES,
	{
		PR_PROFILE_OFFLINE_STORE_PATH,
		PR_MDB_PROVIDER
	}
};

class CMSProvider : public IMSProvider
{
public:
///////////////////////////////////////////////////////////////////////////////
// Interface virtual member functions
//
    STDMETHODIMP QueryInterface
                    (REFIID                     riid,
                     LPVOID *                   ppvObj);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();
    MAPI_IMSPROVIDER_METHODS(IMPL);

///////////////////////////////////////////////////////////////////////////////
// Constructors and destructors
//
public :
    CMSProvider     (LPMSPROVIDER pPSTMSfwd);
    ~CMSProvider    ();

///////////////////////////////////////////////////////////////////////////////
// Data members
//
private :
    ULONG              m_cRef;
    LPMSPROVIDER       m_pPSTMS;
public :
};


class CMsgStore : public IMsgStore
{
public:
///////////////////////////////////////////////////////////////////////////////
// Interface virtual member functions
//
    STDMETHODIMP QueryInterface
                    (REFIID                     riid,
                     LPVOID *                   ppvObj);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();
    MAPI_IMAPIPROP_METHODS(IMPL);
    MAPI_IMSGSTORE_METHODS(IMPL);

    CMsgStore(LPMDB pPSTMsgStore);
    ~CMsgStore();

private:
    ULONG m_cRef;
    LPMDB m_pPSTMsgStore;
};

HRESULT STDAPICALLTYPE DoSync(LPMDB lpMDB);

