/****************************************************************************
        ABLOGON.CPP

        Copyright (c) 2007 Microsoft Corporation

        Implementation of IABLogon,

        which  is used to access resources in an address book provider.
*****************************************************************************/

#include "initguid.h"
#define USES_IID_IABLogon       // For usage of IID_IABLogon.
#define USES_IID_IABContainer

#define USES_IID_IMAPIContainer

#define USES_IID_IMAPIProp
#define USES_IID_IDistList
#define USES_IID_IMailUser

#include "abp.h"
#include "util.h"

///////////////////////////////////////////////////////////////////////////////
//
//    CABLogon::CABLogon()
//
//    Parameters
//          hInst           [in] module instance handle
//          pSubObj         [in] IMAPISupport object
//          pMAPISup        [in] pointer to the support object
//          pDB             [in] pointer to the database object
//    Purpose
//          Class constructor for CABLogon object
//
//    Return Value
//          none.
//
CABLogon::CABLogon(HINSTANCE    hInst,
                   LPMAPISUP    pSupObj,
                   MAPIUID      &UID,
                   PCABDatabase pDB)
: m_cRef(1),
  m_hInst(hInst),
  m_pSupObj(pSupObj),
  m_pRootCont(NULL),
  m_pDirCont(NULL),
  m_pDB(pDB)
{
        if (m_pSupObj)
                m_pSupObj->AddRef();
        if (m_pDB)
                m_pDB->AddRef();

        memcpy((LPVOID) &m_UID, (LPVOID) &UID, sizeof UID);

        InitializeCriticalSection(&m_cs);
}

///////////////////////////////////////////////////////////////////////////////
//
//    CABLogon::~CABLogon()
//
//    Parameters

//          none.
//
//    Purpose
//          Class destructor.
//
//    Return Value
//          none.
//
CABLogon::~CABLogon()
{
        if (m_pDB)
                m_pDB->Release();
        if (m_pSupObj)
                m_pSupObj->Release();
        if (m_pRootCont)
                m_pRootCont->Release();
        if (m_pDirCont)
                m_pDirCont->Release();

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
//          NOERROR on success, E_NOINTERFACE if the interface
//          can't be returned.
//
STDMETHODIMP CABLogon::QueryInterface(REFIID riid, LPVOID *ppvObj)
{
        if (IsBadWritePtr(ppvObj, sizeof LPUNKNOWN))
                return MAPI_E_INVALID_PARAMETER;

        *ppvObj = NULL;

        if (riid == IID_IUnknown || riid == IID_IABLogon)
        {
                *ppvObj = static_cast<LPVOID>(this);
                AddRef();
                return NOERROR;
        }

        return E_NOINTERFACE;
}

///////////////////////////////////////////////////////////////////////////////
//
//    CABLogon::AddRef()
//
//    Parameters
//          Refer to OLE Documentation on this method
//
//    Purpose
//          To increase the reference count of this object,
//          and of any contained (or aggregated) object within
//
//    Return Value
//      Usage count for this object
//
STDMETHODIMP_ (ULONG) CABLogon::AddRef (void)
{
        return InterlockedIncrement(&m_cRef);
}

///////////////////////////////////////////////////////////////////////////////
//
//    CABLogon::Release()
//
//    Parameters
//          Refer to OLE Documentation on this method
//
//    Purpose
//          Decreases the usage count of this object. If the

//          count is down to zero, object deletes itself.
//
//    Return Value
//      Usage of this object
//
STDMETHODIMP_ (ULONG) CABLogon::Release (void)
{
        ULONG ulRefCount = InterlockedDecrement(&m_cRef);

        // If this object is not being used, delete it
        if (ulRefCount == 0)
        {
                // Destroy the object
                delete this;
        }

        return ulRefCount;
}

///////////////////////////////////////////////////////////////////////////////
//
//    CABLogon::HrSetStatusRow()
//
//	Refer to the MSDN documentation for more information.
//
//    Parameters

//          bAvail      [in] PR_STATUS_CODE
//          ulFlags     [in] see IMAPISupport::ModifyStatusRow
//
//    Purpose
//          Sets the row in the status table for this provider. PR_STATUS_CODE,
//          PR_RESORUCE_METHODS, and PR_PROVIDER_DISPLAY are required columns.
//    Return Value
//          none.
//
STDMETHODIMP CABLogon::HrSetStatusRow(BOOL bAvail, ULONG ulFlags)
{
        const ULONG NUM_STATUS_ROW_PROPS = 9;

        SPropValue rgspvStatusRow[NUM_STATUS_ROW_PROPS] = { 0 };

        ULONG i = 0;
        rgspvStatusRow[i].ulPropTag = PR_PROVIDER_DISPLAY;
        rgspvStatusRow[i++].Value.LPSZ = SZ_PSS_SAMPLE_AB;

        rgspvStatusRow[i].ulPropTag = PR_STATUS_CODE;
        rgspvStatusRow[i++].Value.l = bAvail ? STATUS_AVAILABLE : STATUS_FAILURE;

        rgspvStatusRow[i].ulPropTag = PR_STATUS_STRING;
        rgspvStatusRow[i++].Value.LPSZ = bAvail ? SZ_AVAILABLE : SZ_UNAVAILABLE;

        rgspvStatusRow[i].ulPropTag = PR_DISPLAY_NAME;
        rgspvStatusRow[i++].Value.LPSZ = SZ_PSS_SAMPLE_AB;

        rgspvStatusRow[i].ulPropTag = PR_RESOURCE_FLAGS;
        rgspvStatusRow[i++].Value.l = STATUS_NO_PRIMARY_IDENTITY;

        rgspvStatusRow[i].ulPropTag = PR_RESOURCE_METHODS;
        rgspvStatusRow[i++].Value.l = 0;

        rgspvStatusRow[i].ulPropTag = PR_RESOURCE_TYPE;
        rgspvStatusRow[i++].Value.l = MAPI_AB_PROVIDER;

        rgspvStatusRow[i].ulPropTag = PR_OBJECT_TYPE;
        rgspvStatusRow[i++].Value.l = MAPI_STATUS;

        rgspvStatusRow[i].ulPropTag = PR_PROVIDER_DLL_NAME;
        rgspvStatusRow[i++].Value.LPSZ = SZ_DLL_NAME;

        assert(m_pSupObj);
        return m_pSupObj->ModifyStatusRow(i, rgspvStatusRow, ulFlags);
}

///////////////////////////////////////////////////////////////////////////////
//
//    CABLogon::GetLastError()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//          Gets detailed error information based on hResult.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABLogon::GetLastError(HRESULT         /* hResult */,
                                    ULONG           /* ulFlags */,
                                    LPMAPIERROR   * /* lppMAPIError */)
{
        return MAPI_E_NO_SUPPORT;
}

///////////////////////////////////////////////////////////////////////////////
//
//    CABLogon::Logoff()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise.
STDMETHODIMP CABLogon::Logoff(ULONG /* ulFlags */)
{
        // Unsubscibe registered notifications.
        for (list<ULONG>::iterator it = m_ullstAdviseCnx.begin();
                it != m_ullstAdviseCnx.end();
                it++)
        {
                m_pSupObj -> Unsubscribe(*it);
        }

        return S_OK;
}


///////////////////////////////////////////////////////////////////////////////
//
//    CABLogon::OpenEntry()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CABLogon::OpenEntry(ULONG          cbEntryID,

                                 LPENTRYID      lpEntryID,

                                 LPCIID         lpInterface,

                                 ULONG          /* ulFlags */,
                                 ULONG         *lpulObjType,

                                 LPUNKNOWN     *lppUnk)
{
        CheckParameters_IABLogon_OpenEntry(this,
                                           cbEntryID,

                                           lpEntryID,
                                           lpInterface,
                                           ulFlags,
                                           lpulObjType,
                                           lppUnk);

        HRESULT         hRes            = S_OK;
        PSABP_EID       pEID            = (PSABP_EID)(lpEntryID);
        LPCMailUser     pMailUser       = NULL;

        if (lpInterface	!= NULL
            && (REFIID)*lpInterface != IID_IMAPIContainer
            && (REFIID)*lpInterface != IID_IABContainer
            && (REFIID)*lpInterface != IID_IDistList
            && (REFIID)*lpInterface != IID_IMailUser

            && (REFIID)*lpInterface != IID_IMAPIProp)
        {
                return MAPI_E_INTERFACE_NOT_SUPPORTED;
        }

        // Validate the EID

        ULONG ulEntryType = CheckEID(cbEntryID,lpEntryID);
        EnterCriticalSection(&m_cs);

        switch (ulEntryType)
        {
        case MAPI_ABCONT:
                // Root Container
                if (!pEID || pEID->ID.Num == ROOT_ID_NUM)
                {
                        if (m_pRootCont)
                        {
                                m_pRootCont->AddRef();
                                *lppUnk = static_cast<LPUNKNOWN>(m_pRootCont);
                                *lpulObjType = MAPI_ABCONT;
                        }
                        else
                        {
                                hRes = HrOpenRoot(lpulObjType, lppUnk);
                        }

                        break;
                }

                // Dir Container
                if (pEID->ID.Num == DIR_ID_NUM)
                {
                        if (m_pDirCont)
                        {
                                m_pDirCont->AddRef();
                                *lppUnk = static_cast<LPUNKNOWN>(m_pDirCont);
                                *lpulObjType = MAPI_ABCONT;
                                break;
                        }
                }

                hRes = HrOpenContainer(pEID, lpulObjType, lppUnk);

                break;

        case MAPI_MAILUSER:
                pMailUser = new CMailUser(m_hInst, pEID, this, m_pSupObj, m_pDB);

                if (!pMailUser)
                {
                        hRes = MAPI_E_NOT_ENOUGH_MEMORY;
                        goto LCleanup;
                }

                if (FAILED(hRes = pMailUser->HrInit()))
                {
                        pMailUser->Release();
                        goto LCleanup;
                }

                *lppUnk = (LPUNKNOWN)pMailUser;
                *lpulObjType = MAPI_MAILUSER;


                break;

        default:
                hRes = MAPI_E_INVALID_ENTRYID;
                break;
        }

LCleanup:
        LeaveCriticalSection(&m_cs);

        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrOpenRoot()
//
//    Parameters

//          pulObjType  [in]  pointer to place to return new object's type
//          ppUnk       [out] pointer to pointer to where to write new object
//    Purpose
//          Create a new root container. Sets default properties or any
//          existing in the database. Creates the hierarchy and contents table.
//          Listen for notifications from its children.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CABLogon::HrOpenRoot(ULONG     *pulObjType,
                                  LPUNKNOWN *ppUnk)
{
        if (!pulObjType || !ppUnk)
                return MAPI_E_INVALID_PARAMETER;;

        HRESULT         hRes = S_OK;
        PSABP_EID       pEID = NULL;
        LPCABContainer  pCont = NULL;

        if (FAILED(hRes = HrNewEID(0,
                                   ROOT_ID_NUM,
                                   MAPI_ABCONT,
                                   NULL,
                                   &pEID)))
        {
                goto LCleanup;
        }

        // Create root container.
        pCont = new CABContainer(m_hInst, pEID, this, m_pSupObj, m_pDB);

        if (!pCont)
        {
                hRes = MAPI_E_NOT_ENOUGH_MEMORY;
                goto LCleanup;
        }

        pCont->AddRef(); // For the caller.
        m_pRootCont = pCont;

        hRes = pCont->HrInit();

        *pulObjType = MAPI_ABCONT;
        *ppUnk = static_cast<LPUNKNOWN>(pCont);

LCleanup:
        vpfnFreeBuff(pEID);

        if (FAILED(hRes))
        {
                if (pCont)
                        pCont->Release();
        }

        return hRes;
}


///////////////////////////////////////////////////////////////////////////////
//
//    HrOpenContainer()
//
//    Parameters

//          pEID        [in]  pointer to entry ID of object to open

//          pulObjType  [out] pointer where to write the object's type

//          ppUnk       [out] pointer to pointer where to write the object
//    Purpose
//          Opens a container or DL object. Sets default properties and
//          any already existing in the database. Creates a contents
//          table and subscribes to listen for notifications from its
//          children.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CABLogon::HrOpenContainer(PSABP_EID        pEID,
                                       ULONG            *pulObjType,

                                       LPUNKNOWN        *ppUnk)
{
        if (!pEID || !pulObjType || !ppUnk)
                return MAPI_E_INVALID_PARAMETER;

        HRESULT         hRes    = S_OK;
        LPCABContainer  pCont   = NULL;

        *pulObjType     = 0;
        *ppUnk          = NULL;

        pCont = new CABContainer(m_hInst, pEID, this, m_pSupObj, m_pDB);

        if (!pCont)
                return MAPI_E_NOT_ENOUGH_MEMORY;

        if (FAILED(hRes = pCont->HrInit()))
                goto LCleanup;

        if (pEID->ID.Num == DIR_ID_NUM)
        {
                m_pDirCont = pCont;
        }

        pCont->AddRef(); // For the caller.
        *ppUnk = static_cast<LPUNKNOWN>(pCont);
        *pulObjType = pEID->ID.Type;

LCleanup:
        if (FAILED(hRes) && pCont)
        {
                pCont->Release();
        }

        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    CheckEID()
//
//    Parameters

//          cbEntryID   [in] size of entry ID
//          pEntryID    [in] pointer to an entry ID
//    Purpose
//          Check if the EID is valid, what type of object is its owner, container or mailuser,
//          foreign entry?
//

//    Return Value
//          Code indicating the ID type, or MAPI_E_INVALID_ENTRYID is

//          type can't be determined.
//
ULONG WINAPI CABLogon::CheckEID(ULONG cbEntryID, LPENTRYID pEntryID)
{
        PSABP_EID pEID = (PSABP_EID) pEntryID;

        if (cbEntryID && IsBadReadPtr((LPVOID) pEntryID, cbEntryID))
                return (ULONG) MAPI_E_INVALID_ENTRYID;

        // Root container is NULL
        if (cbEntryID == 0 || pEID == NULL)
                return MAPI_ABCONT;

        if (pEID->muid != m_UID)
                return FOREIGN_ENTRY;

        return pEID->ID.Type;
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrNewEID()
//
//    Parameters

//          ulFlags     [in]  entry ID abflags

//          ulID        [in]  object's database ID number
//          ulEntryType [in]  object's type
//          pParent     [in]  block to link allocation to, if non-
//                            NULL, EID is allocated by MAPIAllocMore
//          ppNewEID    [out] pointer to pointer where to write new EID
//
//    Purpose
//          Creates a new entry ID. The allocation can be linked to another
//          block of memory if pParent is non-NULL (a la MAPIAllocateMore)
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CABLogon::HrNewEID(ULONG           ulFlags,
                                ULONG           ulID,
                                ULONG           ulEntryType,
                                LPVOID          pParent,
                                PSABP_EID      *ppNewEID)
{
        if (!ppNewEID)
                return MAPI_E_INVALID_PARAMETER;

        HRESULT         hRes = S_OK;
        PSABP_EID       pEID = NULL;

        if (pParent)
                hRes = vpfnAllocMore(sizeof SABP_EID, pParent, (LPVOID*) &pEID);
        else
                hRes = vpfnAllocBuff(sizeof SABP_EID,(LPVOID*) &pEID);

        if (SUCCEEDED(hRes))
        {
                ZeroMemory((LPVOID) pEID, sizeof SABP_EID);
                memcpy((LPVOID)&(pEID->muid), (LPVOID)&m_UID, sizeof MAPIUID);
                memcpy((LPVOID)pEID->bFlags, (LPVOID)&ulFlags, sizeof ULONG);

                pEID->ID.Num = ulID;
                pEID->ID.Type = ulEntryType;
                *ppNewEID = pEID;

        }
        return hRes;

}

///////////////////////////////////////////////////////////////////////////////
//
//    CompareEntryIDs()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CABLogon::CompareEntryIDs(ULONG       /* cbEntryID1 */,
                                       LPENTRYID   lpEntryID1,
                                       ULONG       /* cbEntryID2 */,
                                       LPENTRYID   lpEntryID2,
                                       ULONG       /* ulFlags */,
                                       ULONG      *lpulResult)
{
        CheckParameters_IABLogon_CompareEntryIDs(this,
                                                 cbEntryID1,
                                                 lpEntryID1,
                                                 cbEntryID2,
                                                 lpEntryID2,
                                                 ulFlags,

                                                 lpulResult);

        PSABP_EID  lhs = (PSABP_EID) lpEntryID1,
                    rhs = (PSABP_EID) lpEntryID2;

        *lpulResult = (*lhs ==*rhs);

        return S_OK;
}

///////////////////////////////////////////////////////////////////////////////
//
//    Advise()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CABLogon::Advise(ULONG             cbEntryID,
                              LPENTRYID         lpEntryID,
                              ULONG             ulEventMask,
                              LPMAPIADVISESINK  lpAdviseSink,
                              ULONG            *lpulConnection)
{
        HRESULT         hRes = S_OK;
        ULONG           ulCnx;
        ULONG           ulEntryType;
        SizedNOTIFKEY(sizeof SABP_EID, Key);

        CheckParameters_IABLogon_Advise(this,
                                        cbEntryID,
                                        lpEntryID,
                                        ulEventMask,
                                        lpAdviseSink,
                                        lpulConnection);

        EnterCriticalSection(&m_cs);

        // Validate entryID
        ulEntryType = CheckEID(cbEntryID, lpEntryID);
        if (ulEntryType == FOREIGN_ENTRY || ulEntryType == MAPI_E_INVALID_ENTRYID)
        {
                hRes = MAPI_E_INVALID_ENTRYID;
                goto LCleanup;
        }

        // Make sure it's a supported event
        if (ulEventMask & ~SUPPORTED_EVENTS)
        {
                hRes = MAPI_E_NO_SUPPORT;
                goto LCleanup;
        }

        Key.cb = cbEntryID;
        memcpy((LPVOID)Key.ab, (LPVOID)lpEntryID, Key.cb);

        assert(m_pSupObj);
        // let MAPI do the work
        if (FAILED(hRes = m_pSupObj->Subscribe((LPNOTIFKEY)&Key,
                                               ulEventMask,
                                               0,
                                               lpAdviseSink,

                                               & ulCnx)))
        {
                goto LCleanup;
        }

        // save the cnx number
        m_ullstAdviseCnx.push_back(ulCnx);

        *lpulConnection = ulCnx;


LCleanup:
        LeaveCriticalSection(&m_cs);
        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    Unadvise()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CABLogon::Unadvise(ULONG ulConnection)
{
        BOOL bRes = false;

        CheckParameters_IABLogon_Unadvise(this, ulConnection);

        EnterCriticalSection(&m_cs);

        for(list<ULONG>::iterator it = m_ullstAdviseCnx.begin();

            it != m_ullstAdviseCnx.end();
            it++)
        {
                if (ulConnection == *it)
                {
                        m_ullstAdviseCnx.remove(ulConnection);
                        bRes = true;
                        break;
                }
        }

        LeaveCriticalSection(&m_cs);

        if (bRes)
        {
                assert(m_pSupObj);
                m_pSupObj->Unsubscribe(ulConnection);
                return S_OK;
        }

        return MAPI_E_CALL_FAILED;
}

///////////////////////////////////////////////////////////////////////////////
//
//    OpenStatusEntry()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CABLogon::OpenStatusEntry(LPCIID          /* lpInterface */,
                                       ULONG           /* ulFlags */,
                                       ULONG         * /* lpulObjType */,
                                       LPMAPISTATUS  * /* lppEntry */)
{
        return MAPI_E_NO_SUPPORT;
}

///////////////////////////////////////////////////////////////////////////////
//
//    OpenTemplateID()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CABLogon::OpenTemplateID(ULONG          /* cbTemplateID */,
                                      LPENTRYID      /* lpTemplateID */,
                                      ULONG          /* ulTemplateFlags */,
                                      LPMAPIPROP     /* lpMAPIPropData */,
                                      LPCIID         /* lpInterface */,
                                      LPMAPIPROP   * /* lppMAPIPropNew */,
                                      LPMAPIPROP     /* lpMAPIPropSibling */)
{
        return MAPI_E_NO_SUPPORT;
}

///////////////////////////////////////////////////////////////////////////////
//
//    GetOneOffTable()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//          Returns two of the provider's templates: a user and DL template

//          IMAPISuport::GetOneOffTable calls each provider's

//          IABLogon::GetOneOffTable to create the merged oneoff
//          table. Contribute provider's two rows here.

//

//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//

STDMETHODIMP CABLogon::GetOneOffTable(ULONG         /* ulFlags */,
                                      LPMAPITABLE * /* lppTable */)
{
        return MAPI_E_NO_SUPPORT;
}

///////////////////////////////////////////////////////////////////////////////
//
//    PrepareRecips()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CABLogon::PrepareRecips(ULONG           /* ulFlags */,
                                     LPSPropTagArray /* lpPropTagArray */,
                                     LPADRLIST       /* lpRecipList */)
{
        return MAPI_E_NO_SUPPORT;
}

///////////////////////////////////////////////////////////////////////////////
//
//    Notify()
//
//    Parameters

//          ulEvent         [in] notification type
//          pParentEID      [in] pointer to object's parent entry ID
//          pObjectEID      [in] pointer to object's entry ID
//          pPropTags       [in] array of prop tags for changed properties
//
//    Purpose
//          This function creates a notification key (the entry ID of the
//          object for whom notifications were Advise()'d) then creates
//          a single notification structure and calls the support object's
//          Notify to broadcast the notification.
//

//

//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CABLogon::HrNotify(ULONG           ulEvent,

                                PSABP_EID       pParentEID,
                                PSABP_EID       pObjectEID,
                                LPSPropTagArray pPropTags)
{
        ULONG         ulFlags = 0;
        NOTIFICATION  Notif;
        SizedNOTIFKEY(sizeof SABP_EID, Key);

        // make a key
        Key.cb = sizeof SABP_EID;
        memcpy(Key.ab, (LPBYTE) pObjectEID, Key.cb);
        ZeroMemory(&Notif,sizeof Notif);

        Notif.ulEventType             = ulEvent;
        Notif.info.obj.cbEntryID      = sizeof *pObjectEID;
        Notif.info.obj.lpEntryID      = (LPENTRYID) pObjectEID;
        Notif.info.obj.ulObjType      = pObjectEID->ID.Type;
        Notif.info.obj.cbParentID     = pParentEID ? sizeof(*pParentEID) : 0;
        Notif.info.obj.lpParentID     = (LPENTRYID) pParentEID;
        Notif.info.obj.lpPropTagArray = pPropTags;

        assert(m_pSupObj);
        // let MAPI do the notifications
        return m_pSupObj->Notify((LPNOTIFKEY)&Key,
                                 1,
                                 &Notif,
                                 &ulFlags);
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrGetCntsTblRowProps()
//
//    Parameters

//          pRec        [in] pointer to CRecord standing for a content table
//                           row entry.
//          ppProps     [out] pointer to pointer to props for a content row.
//    Purpose
//          This function returns a pointer to props array corresponding to
//          columns of the content table. Putting the func in CABLogon other
//          other than CABContainer, is because that CMailUser also calls it
//          to generate a content row when CMailUser needs to NOTIFY mapi.
//

//

//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CABLogon::HrGetCntsTblRowProps(PCRecord      pRec,

                                            LPSPropValue *ppProps)
{
        if (!pRec || !ppProps)
                return MAPI_E_INVALID_PARAMETER;

        HRESULT         hRes = S_OK;
        PSABP_EID       pEID = NULL;
        LPSPropValue    pProps = NULL;
        LPSPropValue    pPropTmp = NULL;
        ULONG           cValues = pRec->cValues();

        // Allocate memory for pProps.
        if (FAILED(hRes = vpfnAllocBuff(NUM_CTBL_COLS * sizeof SPropValue, (LPVOID * )&pProps)))
                return hRes;

        if (!pProps)
        {
                hRes = MAPI_E_NOT_ENOUGH_MEMORY;
                goto LCleanup;
        }

        ZeroMemory((LPVOID)pProps, NUM_CTBL_COLS * sizeof SPropValue);

        // Generate a new entry ID for an mailuser in content table.
        if (FAILED(hRes = HrNewEID(0,
                                   pRec->RecID(),
                                   MAPI_MAILUSER,
                                   NULL,
                                   &pEID)))
        {
                goto LCleanup;

        }

        // Set regular props.
        pProps[CT_EID].ulPropTag        = PR_ENTRYID;
        pProps[CT_EID].Value.bin.cb     = sizeof(*pEID);
        pProps[CT_EID].Value.bin.lpb    = (LPBYTE)pEID;

        pProps[CT_INST_KEY].ulPropTag   = PR_INSTANCE_KEY;
        pProps[CT_INST_KEY].Value.bin.cb = sizeof(pEID->ID.Num);
        pProps[CT_INST_KEY].Value.bin.lpb = (LPBYTE) & (pEID->ID.Num);

        pProps[CT_REC_KEY].ulPropTag    = PR_RECORD_KEY;
        pProps[CT_REC_KEY].Value.bin.cb = sizeof(pEID->ID.Num);
        pProps[CT_REC_KEY].Value.bin.lpb = (LPBYTE) & (pEID->ID.Num);

        pProps[CT_DISP_TYPE].ulPropTag  = PR_DISPLAY_TYPE;
        pProps[CT_DISP_TYPE].Value.l    = DT_MODIFIABLE;

        pProps[CT_OBJTYPE].ulPropTag    = PR_OBJECT_TYPE;
        pProps[CT_OBJTYPE].Value.l      = MAPI_MAILUSER;

        // Set props whose values are saved in the database.
        for(ULONG i=CT_OBJTYPE + 1; i<NUM_CTBL_COLS; i++)
        {
                pPropTmp = PpropFindProp(pRec->lpProps(), cValues, vsptCntTbl.aulPropTag[i]);

                if (pPropTmp)
                {
                        pProps[i] = *pPropTmp;
                        pPropTmp = NULL;
                }
                else
                {
                        // Set to not found if the prop is missing.
                        HrPropNotFound(pProps[i], vsptCntTbl.aulPropTag[i]);
                }
        }

        *ppProps = pProps;

LCleanup:
        if (FAILED(hRes))
        {
                if (pProps)
                        vpfnFreeBuff(pProps);
        }

        return hRes;
}


///////////////////////////////////////////////////////////////////////////////
//
//    operator !=()
//
//    Return Value
//          TRUE if lhs != rhs, FALSE otherwise
//
inline BOOL operator !=(const MAPIUID & lhs, const MAPIUID & rhs)
{
        return memcmp((LPVOID) & lhs, (LPVOID) & rhs, sizeof MAPIUID);
}

///////////////////////////////////////////////////////////////////////////////
//
//    operator ==()
//
//    Parameters

//          lhs         [in] reference to the left side ID
//          rhs         [in] reference to the right side ID
//    Purpose
//          Equality operator for IDs.
//

//    Return Value
//          TRUE if both are equal, FALSE otherwise.
//

BOOL operator ==(const SABP_ID & lhs, const SABP_ID & rhs)
{

    return (lhs.Num == rhs.Num && lhs.Type == rhs.Type);
}


///////////////////////////////////////////////////////////////////////////////
//
//    operator ==()
//
//    Parameters

//          lhs         [in] reference to the left side entry ID
//          rhs         [in] reference to the right side entry ID
//    Purpose
//          Equality operator for entry IDs.
//

//    Return Value
//          TRUE if both are equal, FALSE otherwise.
//
BOOL operator ==(const SABP_EID & lhs, const SABP_EID & rhs)
{
    return (lhs.ID == rhs.ID
            && (memcmp((LPVOID) & lhs.muid, (LPVOID) & rhs.muid, sizeof MAPIUID)) == 0);
}

/* End of the file of ABLOGON.CPP */