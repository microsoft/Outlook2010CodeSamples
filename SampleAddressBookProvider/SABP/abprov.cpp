/****************************************************************************
        ABPROV.CPP

        Copyright (c) 2007 Microsoft Corporation

        Implementation of ServiceEntry, ABProviderInit
        and IABProvider, which provides a method to log on to an address book

        provider object and a method to invalidate an address book provider object
*****************************************************************************/

#include "initguid.h"
#define USES_IID_IABProvider            // for use of IID_IABProvider

#include "abp.h"

// Version returned by Unicode or ANSI providers.
#define CURRENT_SPI_VERSION_UNICODE_ANSI          0x00020003L


LPALLOCATEBUFFER    vpfnAllocBuff;      // MAPIAllocateBuffer function
LPALLOCATEMORE      vpfnAllocMore;      // MAPIAllocateMore function

LPFREEBUFFER        vpfnFreeBuff;       // MAPIFreeBuffer function



///////////////////////////////////////////////////////////////////////////////
//
//    ServiceEntry()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//          Called by the profile setup adminitration tool to display the provider
//          configuration properties for this transport provider. Provides

//          exception handler for configuration exceptions.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.

//
HRESULT STDAPICALLTYPE ServiceEntry(HINSTANCE          /* hInst */,
                                    LPMALLOC           pMallocObj,
                                    LPMAPISUP          pMAPISup,
                                    ULONG              ulUIParam,
                                    ULONG              /* ulFlags */,
                                    ULONG              ulContext,
                                    ULONG              /* ulCfgPropCount */,
                                    LPSPropValue       /* pCfgProps */,
                                    LPPROVIDERADMIN    pAdminProv,
                                    LPMAPIERROR     *  /* ppMAPIError */)
{
        if (!pMAPISup || !pAdminProv)
                return MAPI_E_INVALID_PARAMETER;

        HRESULT         hRes      = S_OK;
        LPCConfig       pCfg      = NULL;
        LPPROFSECT      pProfSecn = NULL;


        if (pMallocObj)
                pMallocObj->Release(); // Never use it, so release it.

        // Save the allocation routines
        if (FAILED(hRes = pMAPISup->GetMemAllocRoutines(&vpfnAllocBuff,
                                                        &vpfnAllocMore,
                                                        &vpfnFreeBuff)))
        {
                goto LCleanup;
        }

        switch (ulContext)
        {
        case MSG_SERVICE_DELETE:
                // Open profile for deleting logon props.
                if (FAILED(hRes = pAdminProv->OpenProfileSection((LPMAPIUID)&SABP_UID,
                                                                 NULL,
                                                                 MAPI_MODIFY,
                                                                 &pProfSecn)))
                {
                        goto LCleanup;
                }

                // Delete props from profile.
                if (FAILED(hRes = pProfSecn->DeleteProps((LPSPropTagArray)&vsptLogonTags, NULL)))
                        goto LCleanup;

                break;
        case MSG_SERVICE_INSTALL:
                // Fall through
        case MSG_SERVICE_UNINSTALL:
                break;

        case MSG_SERVICE_CREATE:
                // Fall through
        case MSG_SERVICE_CONFIGURE :
                if (FAILED(hRes = pAdminProv->OpenProfileSection((LPMAPIUID)&SABP_UID,
                                                                 NULL,
                                                                 MAPI_MODIFY,
                                                                 &pProfSecn)))
                {
                        goto LCleanup;
                }

                pCfg = new CConfig((HWND)ulUIParam,
                                   *(LPMAPIUID)&SABP_UID,
                                   pProfSecn);

                if (!pCfg)
                {
                        hRes = MAPI_E_NOT_ENOUGH_MEMORY;
                        goto LCleanup;
                }

                if (FAILED(hRes = pCfg->HrInit()))
                        goto LCleanup;

                if (FAILED(hRes = pCfg->HrDoConfig()))
                        goto LCleanup;

                if (FAILED(hRes = pCfg->HrUpdateProfile()))
                        goto LCleanup;

                break;

        default:
                hRes = MAPI_E_NO_SUPPORT;
                break;
        }

LCleanup:
        if (pCfg)
                pCfg->Release();

        return hRes;

}


///////////////////////////////////////////////////////////////////////////////
//
//    ABProviderInit()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//          This function provides an exception handler for
//          IABProvider object's constructor.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise.

//
STDINITMETHODIMP ABProviderInit(HINSTANCE         hInst,
                                LPMALLOC          pMalloc,
                                LPALLOCATEBUFFER  pAllocBuff,
                                LPALLOCATEMORE    pAllocMore,
                                LPFREEBUFFER      pFreeBuff,
                                ULONG             /* ulFlags */,
                                ULONG             ulMAPIVers,
                                ULONG            *pulABVers,
                                LPABPROVIDER     *ppABProvider)
{
        HRESULT hRes = S_OK;
        LPCABProvider pABProv = NULL;

        *ppABProvider = NULL;

        // Never use this, so release it.
        pMalloc->Release();

        // This must come before first use of new, since the debug
        // version overloads new with MAPI allocators (for leak detection)
        vpfnAllocBuff = pAllocBuff;
        vpfnAllocMore = pAllocMore;
        vpfnFreeBuff  = pFreeBuff;

        // Version check.

        // If the provider is more recent than caller (has a larger
        // version number) and the provider wants to run with the caller, the provider should return

        // its version number back to the caller. If the provider is older (has a smaller

        // version number), the provider returns its version, and lets the caller decide if it

        // can call the provider.
        if (ulMAPIVers < CURRENT_SPI_VERSION)
        {
                *pulABVers = CURRENT_SPI_VERSION;
                return MAPI_E_VERSION;
        }

        *pulABVers = CURRENT_SPI_VERSION_UNICODE_ANSI;

        // Every checking goes right, then create IABProvider.
        pABProv = new CABProvider(hInst);

        if (!pABProv)
        {
                hRes = MAPI_E_NOT_ENOUGH_MEMORY;
                goto LCleanup;
        }

        *ppABProvider = (LPABPROVIDER) pABProv;

LCleanup:
        if (FAILED(hRes))
        {
                if (pABProv)
                        pABProv->Release();
        }
        return hRes;
}


///////////////////////////////////////////////////////////////////////////////
//
//    CABProvider::CABProvider()
//
//    Parameters

//          hInst   [in] module instance handle
//
//    Purpose
//          Class constructor.
//

//    Return Value
//          none.
CABProvider::CABProvider(HINSTANCE hInst)
: m_cRef(1),
  m_hInst(hInst)
{
        InitializeCriticalSection(&m_cs);
}

///////////////////////////////////////////////////////////////////////////////
//
//    ~CABProvider()
//
//    Parameters

//          none.
//
//    Purpose
//          Class destructor.
//

//

//    Return Value

CABProvider::~CABProvider()
{
        DeleteCriticalSection(&m_cs);
}


///////////////////////////////////////////////////////////////////////////////
//
//    QueryInterface()
//
//    Parameters

//          See the OLE documentation for details.
//
//    Purpose
//          Returns an interface if supported.
//

//

//    Return Value
//          NOERROR on success, E_NOINTERFACE if the interface can't be returned.
//
STDMETHODIMP CABProvider::QueryInterface(REFIID riid, LPVOID *ppvObj)
{
        if (IsBadWritePtr(ppvObj, sizeof LPUNKNOWN))
                return MAPI_E_INVALID_PARAMETER;

        *ppvObj = NULL;

        if (riid == IID_IUnknown || riid == IID_IABProvider)
        {
                *ppvObj = static_cast<LPVOID>(this);
                AddRef();
                return NOERROR;
        }

        return E_NOINTERFACE;
}

///////////////////////////////////////////////////////////////////////////////
//
//    CABProvider::AddRef()
//
//    Parameters
//          Refer to OLE Documentation on this method
//
//    Purpose
//          To increase the reference count of this object,
//          and of any contained (or aggregated) object within
//
//    Return Value
//          Usage count for this object
//
STDMETHODIMP_(ULONG) CABProvider::AddRef()
{
        return InterlockedIncrement(&m_cRef);
}

///////////////////////////////////////////////////////////////////////////////
//
//    CABProvider::Release()
//
//    Parameters
//          Refer to OLE Documentation on this method
//
//    Purpose
//          Decreases the usage count of this object. If the count is

//          zero, object deletes itself.
//
//    Return Value
//          Usage of this object
//
STDMETHODIMP_(ULONG) CABProvider::Release()
{
        ULONG ulRef = InterlockedDecrement(&m_cRef);
        if (ulRef == 0)
        {
                delete this;
        }

        return ulRef;
}

///////////////////////////////////////////////////////////////////////////////
//
//    CABProvider::Shutdown()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//          To terminate all sessions attached to this provider.
//
//    Return Value
//      S_OK                    If no error
//      MAPI_E_UNKNOWN_FLAGS    If flags are passed to this method
//
STDMETHODIMP CABProvider::Shutdown(ULONG * /* lpulFlags */)
{
        return S_OK;
}

///////////////////////////////////////////////////////////////////////////////
//
//    CABProvider::Logon()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//          Display the logon dialog to show the options saved in the
//          profile for this provider and allow changes to it. Save

//          new configuration settings back in the profile. Try and

//          open the DB file, or create it if not found. Create a

//          new CABLogon object and return it to the caller. Try
//          to get the logon UID from the profile (the UID is unique
//          to the database file), create a new one if not found. Set
//          the provider UID.   Catch any exceptions thrown in the
//          ABLogon constructor.
//
//    Return Value
//      S_OK                        If no error
//      MAPI_E_NO_ACCESS            If transport is logged with the

//                                  LOGON_NO_DIALOG flag
//      MAPI_E_NOT_ENOUGH_MEMORY    If could not allocate new CXPLogon object
//      hResult                     With any other error code
//
STDMETHODIMP CABProvider::Logon(LPMAPISUP        lpMAPISup,

                                ULONG            /* ulUIParam */,
                                LPTSTR           /* lpszProfileName */,
                                ULONG            /* ulFlags */,
                                ULONG          * /* lpulpcbSecurity */,
                                LPBYTE         * /* lppbSecurity */,
                                LPMAPIERROR    * /* lppMAPIError */,
                                LPABLOGON      * lppABLogon)
{
        HRESULT         hRes      = S_OK;
        LPCABLogon      pLogonObj = NULL;
        LPMAPIUID       pUID      = NULL;
        LPPROFSECT      pProfSecn = NULL;
        ULONG           ulcProps;
        LPSPropValue    pProps    = NULL;
        PCABDatabase    pDB       = NULL;

        CheckParameters_IABProvider_Logon(this,
                                          lpMAPISup,
                                          ulUIParam,
                                          lpszProfileName,
                                          ulFlags,
                                          lpulpcbSecurity,
                                          lppbSecurity,
                                          lppMAPIError,
                                          lppABLogon);

        EnterCriticalSection(&m_cs);

        lpMAPISup->AddRef();

        // Cast SABP_UID(GUID) to LPMAPIUID.
        pUID = (LPMAPIUID)&SABP_UID;

        // Open profile to get database file name.
        if (FAILED(hRes = lpMAPISup->OpenProfileSection(pUID, MAPI_MODIFY, &pProfSecn)))
                goto LCleanup;

        if (FAILED(hRes = pProfSecn->GetProps((LPSPropTagArray)&vsptLogonTags,

                                              fMapiUnicode,
                                              &ulcProps,
                                              &pProps)))
        {
                goto LCleanup;
        }

        // Validate props here.
        if (!pProps || pProps[LOGON_DB_PATH].ulPropTag != PR_SMP_AB_PATH)
        {
                hRes = MAPI_E_UNCONFIGURED;
                goto LCleanup;
        }


        // Open database.
        pDB = new CTxtABDatabase(pProps[LOGON_DB_PATH].Value.LPSZ);
        if (!pDB)
        {
                hRes = MAPI_E_NOT_ENOUGH_MEMORY;
                goto LCleanup;
        }
        if (FAILED(hRes = pDB->Open()))
        {
                goto LCleanup;
        }

        // Initiate CABLogon object.
        pLogonObj = new CABLogon(m_hInst,

                                 lpMAPISup,

                                 *pUID,
                                 pDB);

        if (!pLogonObj)
        {
                hRes = MAPI_E_NOT_ENOUGH_MEMORY;
                goto LCleanup;
        }

        if (FAILED(hRes = lpMAPISup->SetProviderUID(pUID, 0)))
                goto LCleanup;

        // Set status row failure is not fatal, so ignore it.
        pLogonObj->HrSetStatusRow(TRUE, 0);

        *lppABLogon = (LPABLOGON)pLogonObj;

LCleanup:
        LeaveCriticalSection(&m_cs);

        if (FAILED(hRes))
        {
                if (pLogonObj)
                        pLogonObj->Release();

                if (pDB)
                        pDB->Release();
        }

        if (lpMAPISup)
                lpMAPISup->Release();

        return hRes;
}
/* End of the file ABPROV.CPP */