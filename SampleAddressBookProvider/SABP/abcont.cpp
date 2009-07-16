/****************************************************************************
        ABCONT.CPP

        Copyright (c) 2007 Microsoft Corporation

        Implementation of IABContainer,
        which provides access to address book containers.
*****************************************************************************/


#include "initguid.h"
#define USES_IID_IABContainer           // For usage of IID_IABProvider.
#define USES_IID_IMAPIContainer

#define USES_IID_IMAPIProp
#define USES_IID_IMAPIPropData
#define USES_IID_IMAPITableData
#define USES_IID_IMAPITable
#define USES_IID_IMAPIAdviseSink

#include "abp.h"
#include "util.h"
#include <list>

///////////////////////////////////////////////////////////////////////////////
//
//    CABContainer()

//
//    Purpose
//          Class constructor.
//

//    Return Value
//          none.
//
CABContainer::CABContainer(HINSTANCE    hInst,
                           PSABP_EID    pEID,
                           LPCABLogon   pLogon,
                           LPMAPISUP    pSupObj,
                           PCABDatabase pDB)
: m_cRef(1),
  m_hInst(hInst),
  m_pLogon(pLogon),
  m_pSupObj(pSupObj),
  m_pDB(pDB),
  m_pCntsTblData(NULL),
  m_pHierTblData(NULL),
  m_pAdviseSink(NULL),
  m_pPropData(NULL),
  m_fInited(FALSE)
{
        if (m_pLogon)
                m_pLogon->AddRef();
        if (m_pSupObj)
                m_pSupObj->AddRef();
        if (m_pDB)
                m_pDB->AddRef();


        ZeroMemory(&m_EID, sizeof SABP_EID);
        if (pEID)
        {
                memcpy((LPBYTE) &m_EID, (LPBYTE) pEID, sizeof SABP_EID);
        }

        InitializeCriticalSection(&m_cs);
}

STDMETHODIMP CABContainer::HrInit()
{
        HRESULT hRes = S_OK;

        assert(m_pLogon);
        assert(m_pSupObj);

        if (m_fInited)
        {
                return hRes;
        }

        // Create IPROPDATA object.
        if (FAILED(hRes = CreateIProp(&IID_IMAPIPropData,
                                      vpfnAllocBuff,
                                      vpfnAllocMore,
                                      vpfnFreeBuff,
                                      NULL,
                                      &m_pPropData)))
        {
                goto LCleanup;
        }

        m_pPropData->AddRef();


        // Set default container props.
        if (FAILED(hRes = HrSetDefaultContProps(&m_EID)))
                goto LCleanup;

        // Initiate IMAPIAdviseSink object,
        // And subscribe notification.
        m_pAdviseSink = new CTblAdviseSink(this);

        if (! m_pAdviseSink)
                return MAPI_E_NOT_ENOUGH_MEMORY;

        hRes = HrSubscribe();

LCleanup:
        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrSubscribe()
//
//    Parameters

//          none.
//
//    Purpose
//          Asks MAPI to send notifications for a fnevTableModified. This
//          is to update the contents table when a child is added/deleted
//          or changed. Use the entry ID as the key, it's unique, and

//          the child can generate it easily by getting the parent list

//          when the child needs to notify the provider.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::HrSubscribe()
{
        SizedNOTIFKEY(sizeof SABP_EID,Key);
        PSABP_EID pEID = (PSABP_EID)Key.ab;

        Key.cb = sizeof SABP_EID;
        memcpy((LPVOID)pEID, (LPVOID) &m_EID, sizeof SABP_EID);

        // Later when the child, for example, does a notify, it'll construct

        // the entry ID on the fly knowing only the ID number.

        // Don't know or care what the object type is
        pEID->ID.Type = DONT_CARE;

        if(m_pSupObj)
                return m_pSupObj->Subscribe((LPNOTIFKEY)&Key,
                                            fnevTableModified,
                                            0,
                                            m_pAdviseSink,
                                            &m_ulCnx);
        else
                return E_FAIL;
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrSetDefaultContProps()
//
//    Parameters
//
//    Purpose
//          Initialize a list of properties with default values for a container.

//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::HrSetDefaultContProps(PSABP_EID pEID)
{
        enum

        {
                CONT_FLAGS,
                EID,
                INST_KEY,
                OBJ_TYPE,
                DISP_TYPE,
                DISP_NAME,
                DEPTH,
                ACCESS,
                STATUS,
                NUM_PROPS
        };

        SPropValue Props[NUM_PROPS];

        Props[CONT_FLAGS].ulPropTag     = PR_CONTAINER_FLAGS;
        Props[CONT_FLAGS].Value.l       = AB_MODIFIABLE | AB_RECIPIENTS;

        Props[EID].ulPropTag            = PR_ENTRYID;
        Props[EID].Value.bin.cb         = sizeof(*pEID);
        Props[EID].Value.bin.lpb        = (LPBYTE) pEID;

        Props[INST_KEY].ulPropTag       = PR_INSTANCE_KEY;
        Props[INST_KEY].Value.l         = pEID->ID.Num;

        Props[OBJ_TYPE].ulPropTag       = PR_OBJECT_TYPE;
        Props[OBJ_TYPE].Value.l         = MAPI_ABCONT;


        Props[DISP_TYPE].ulPropTag      = PR_DISPLAY_TYPE;
        Props[DISP_TYPE].Value.l        = DT_MODIFIABLE;

        Props[DISP_NAME].ulPropTag      = PR_DISPLAY_NAME;
        Props[DISP_NAME].Value.LPSZ     = m_EID.ID.Num == DIR_ID_NUM

                                          ? SZ_PSS_SAMPLE_AB : NULL_STRING;


        Props[DEPTH].ulPropTag          = PR_DEPTH;
        Props[DEPTH].Value.l            = 0;


        Props[ACCESS].ulPropTag         = PR_ACCESS;
        Props[ACCESS].Value.l           = MAPI_MODIFY;

        Props[STATUS].ulPropTag         = PR_STATUS;
        Props[STATUS].Value.l           = 0;


        return SetProps(NUM_PROPS, (LPSPropValue)&Props, NULL);
}

///////////////////////////////////////////////////////////////////////////////
//
//    ~CABContainer()
//
//    Parameters
//          none.
//
//    Purpose
//          Class destructor. Releases all outstanding objects.
//

//    Return Value
//          none.
CABContainer::~CABContainer()
{
        if (m_pCntsTblData)
                m_pCntsTblData->Release();


        if (m_pHierTblData)
                m_pHierTblData->Release();

        if (m_pPropData)
                m_pPropData->Release();

        if (m_pDB)
                m_pDB->Release();

        if (m_pSupObj)
                m_pSupObj->Release();

        if (m_pAdviseSink)
                m_pAdviseSink->Release();

        if (m_pLogon)
                m_pLogon->Release();

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
STDMETHODIMP CABContainer::QueryInterface(REFIID riid, LPVOID *ppvObj)
{
        if (IsBadWritePtr(ppvObj, sizeof LPUNKNOWN))
                return MAPI_E_INVALID_PARAMETER;

        *ppvObj = NULL;

        if (riid == IID_IUnknown || riid == IID_IMAPIProp

            || riid == IID_IABContainer || riid == IID_IMAPIContainer)
        {
                *ppvObj = static_cast<LPVOID>(this);
                AddRef();
                return NOERROR;
        }

        return E_NOINTERFACE;
}

///////////////////////////////////////////////////////////////////////////////
//
//    AddRef()
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
STDMETHODIMP_ (ULONG) CABContainer::AddRef (void)
{
        return InterlockedIncrement(&m_cRef);
}

///////////////////////////////////////////////////////////////////////////////
//
//    Release()
//
//    Parameters
//          Refer to OLE Documentation on this method
//
//    Purpose
//          Decreases the usage count of this object. If the

//          count is down to zero, object deletes itself.
//
//    Return Value
//          Usage of this object
//
STDMETHODIMP_ (ULONG) CABContainer::Release (void)
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
//    GetLastError()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//          Gets detailed error information based on hResult.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//

STDMETHODIMP CABContainer::GetLastError(HRESULT         hResult,
                                        ULONG           ulFlags,
                                        LPMAPIERROR    *ppMAPIError)
{
        CheckParameters_IMAPIProp_GetLastError(this,
                                               hResult,
                                               ulFlags,
                                               ppMAPIError);

        assert(m_pPropData);
        return m_pPropData ->GetLastError(hResult, ulFlags, ppMAPIError);
}

///////////////////////////////////////////////////////////////////////////////
//
//    SaveChanges()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP  CABContainer::SaveChanges(ULONG /* ulFlags */)
{
        return MAPI_E_NO_SUPPORT;
}


///////////////////////////////////////////////////////////////////////////////
//
//    GetProps()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP  CABContainer::GetProps(LPSPropTagArray    lpPropTagArray,
                                     ULONG              ulFlags,
                                     ULONG             *lpcValues,
                                     LPSPropValue      *lppPropArray)
{
        CheckParameters_IMAPIProp_GetProps(this,
                                           lpPropTagArray,
                                           ulFlags,
                                           lpcValues,
                                           lppPropArray);

        assert(m_pPropData);
        return m_pPropData->GetProps(lpPropTagArray,
                                     ulFlags,
                                     lpcValues,
                                     lppPropArray);
}

///////////////////////////////////////////////////////////////////////////////
//
//    SetProps()
//
//    Refer to the MSDN documentation for more information.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::SetProps(ULONG               cValues,
                                    LPSPropValue        lpPropArray,
                                    LPSPropProblemArray *lppProblems)
{
        CheckParameters_IMAPIProp_SetProps(this,
                                           cValues,
                                           lpPropArray,
                                           lppProblems);

        assert(m_pPropData);
        return m_pPropData->SetProps(cValues,
                                     lpPropArray,
                                     lppProblems);
}

///////////////////////////////////////////////////////////////////////////////
//
//    GetPropList()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP  CABContainer::GetPropList(ULONG ulFlags,
                                        LPSPropTagArray *lppPropTagArray)
{
        CheckParameters_IMAPIProp_GetPropList(this,
                                              ulFlags,
                                              lppPropTagArray);


        assert(m_pPropData);
        return m_pPropData->GetPropList(ulFlags,
                                        lppPropTagArray);
}

///////////////////////////////////////////////////////////////////////////////
//
//    DeleteProps()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::DeleteProps(LPSPropTagArray          lpPropTagArray,
                                       LPSPropProblemArray     *lppProblems)
{
        CheckParameters_IMAPIProp_DeleteProps(this, lpPropTagArray, lppProblems);

        assert(m_pPropData);
        return m_pPropData->DeleteProps(lpPropTagArray, lppProblems);
}

///////////////////////////////////////////////////////////////////////////////
//
//    CopyTo()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::CopyTo(ULONG                 ciidExclude,
                                  LPCIID                rgiidExclude,
                                  LPSPropTagArray       lpExcludeProps,
                                  ULONG                 ulUIParam,
                                  LPMAPIPROGRESS        lpProgress,
                                  LPCIID                lpInterface,
                                  LPVOID                lpDestObj,
                                  ULONG                 ulFlags,
                                  LPSPropProblemArray  *lppProblems)
{
        CheckParameters_IMAPIProp_CopyTo(this,
                                         ciidExclude,
                                         rgiidExclude,
                                         lpExcludeProps,
                                         ulUIParam,
                                         lpProgress,
                                         lpInterface,
                                         lpDestObj,
                                         ulFlags,
                                         lppProblems);

        assert(m_pPropData);
        return m_pPropData->CopyTo(ciidExclude,
                                   rgiidExclude,
                                   lpExcludeProps,
                                   ulUIParam,
                                   lpProgress,
                                   lpInterface,
                                   lpDestObj,
                                   ulFlags,
                                   lppProblems);
}

///////////////////////////////////////////////////////////////////////////////
//
//    CopyProps()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::CopyProps(LPSPropTagArray       lpIncludeProps,
                                     ULONG                 ulUIParam,
                                     LPMAPIPROGRESS        lpProgress,
                                     LPCIID                lpInterface,
                                     LPVOID                lpDestObj,
                                     ULONG                 ulFlags,
                                     LPSPropProblemArray  *lppProblems)

{
        CheckParameters_IMAPIProp_CopyProps(this,
                                            lpIncludeProps,
                                            ulUIParam,
                                            lpProgress,
                                            lpInterface,
                                            lpDestObj,
                                            ulFlags,
                                            lppProblems);

        assert(m_pPropData);
        return m_pPropData->CopyProps(lpIncludeProps,
                                      ulUIParam,
                                      lpProgress,
                                      lpInterface,
                                      lpDestObj,
                                      ulFlags,
                                      lppProblems);
}

///////////////////////////////////////////////////////////////////////////////
//
//    GetNamesFromIDs()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::GetNamesFromIDs(LPSPropTagArray  *lppPropTags,
                                           LPGUID            lpPropSetGuid,
                                           ULONG             ulFlags,
                                           ULONG            *lpcPropNames,
                                           LPMAPINAMEID	   **lpppPropNames)
{
        CheckParameters_IMAPIProp_GetNamesFromIDs(this,
                                                  lppPropTags,
                                                  lpPropSetGuid,
                                                  ulFlags,
                                                  lpcPropNames,
                                                  lpppPropNames);

       assert(m_pPropData);
       return m_pPropData->GetNamesFromIDs(lppPropTags,
                                            lpPropSetGuid,
                                            ulFlags,
                                            lpcPropNames,
                                            lpppPropNames);
}

///////////////////////////////////////////////////////////////////////////////
//
//    GetIDsFromNames()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::GetIDsFromNames(ULONG             cPropNames,
                                           LPMAPINAMEID     *lppPropNames,
                                           ULONG             ulFlags,
                                           LPSPropTagArray  *lppPropTags)
{
        CheckParameters_IMAPIProp_GetIDsFromNames(this,
                                                  cPropNames,
                                                  lppPropNames,
                                                  ulFlags,
                                                  lppPropTags);

        assert(m_pPropData);
        return m_pPropData->GetIDsFromNames(cPropNames,
                                            lppPropNames,
                                            ulFlags,
                                            lppPropTags);
}

///////////////////////////////////////////////////////////////////////////////
//
//    OpenProperty()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//          Returns an object when requested a supported property
//          of type PT_OBJECT.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::OpenProperty(ULONG       ulPropTag,
                                        LPCIID      piid,
                                        ULONG       ulInterfaceOptions,
                                        ULONG       /* ulFlags */,
                                        LPUNKNOWN  *ppUnk)
{
        HRESULT     hRes = S_OK;
        LPMAPITABLE pTbl = NULL;

        CheckParameters_IMAPIProp_OpenProperty(this,
                                               ulPropTag,
                                               piid,
                                               ulInterfaceOptions,
                                               ulFlags,
                                               ppUnk);


        *ppUnk = NULL;

        EnterCriticalSection(&m_cs);

        switch (ulPropTag)
        {
        case PR_CREATE_TEMPLATES:
                if (*piid != IID_IMAPITable)
                {
                        hRes = MAPI_E_INTERFACE_NOT_SUPPORTED;
                        goto LCleanup;
                }

                if (FAILED(hRes = HrOpenTemplates(&pTbl)))
                        goto LCleanup;

                *ppUnk = static_cast<LPUNKNOWN>(pTbl);
                break;

        case PR_CONTAINER_CONTENTS:
                if (*piid != IID_IMAPITable)
                {
                        hRes = MAPI_E_INTERFACE_NOT_SUPPORTED;
                        goto LCleanup;
                }

                if (FAILED(hRes = GetContentsTable(ulInterfaceOptions, &pTbl)))
                        goto LCleanup;

                *ppUnk = static_cast<LPUNKNOWN>(pTbl);
                break;

        case PR_CONTAINER_HIERARCHY:
                if(*piid != IID_IMAPITable)
                {
                        hRes = MAPI_E_INTERFACE_NOT_SUPPORTED;
                        goto LCleanup;
                }
                if (FAILED(hRes = GetHierarchyTable(ulInterfaceOptions, &pTbl)))
                        goto LCleanup;

                *ppUnk = static_cast<LPUNKNOWN>(pTbl);
                break;

        default:
                hRes = MAPI_E_NO_SUPPORT;
                break;
        }

LCleanup:
        LeaveCriticalSection(&m_cs);
        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrOpenTemplates()
//
//    Parameters

//          lppTable    [out] pointer to pointer to hold new oneoff table
//    Purpose
//          Creates a merged oneoff table containing native

//          templates plus all foreign templates in the session.
//
//          Let anyone put in entries, so the oneoff table is a

//          merge of the session table (with the other AB's templates).

//          This is a little tricky. Other ABs call

//          IMAPISupport::GetOneOffTable to get the provider's templates so they

//          can put entries of the provider's type in their databases.
//
//          IABLogon::GetOneOffTable gets called --once-- to

//          create the merged table.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CABContainer::HrOpenTemplates(LPMAPITABLE *lppTable)
{
        if (!lppTable)
                return MAPI_E_INVALID_PARAMETER;

        HRESULT         hRes   = S_OK;
        LPTABLEDATA     pMyTbl = NULL;
        PSABP_EID       pEID   = NULL;
        SRow            sRow   = {0};
        ULONG           ulRow  = 0;

        enum

        {
                NAME,
                EID,
                DEPTH,
                SELECT,
                ADDR_TYPE,
                DISP_TYPE,
                INST_KEY,
                ROWID,
                NUM_ONEOFF_PROPS
        };

        static const SizedSPropTagArray(NUM_ONEOFF_PROPS, rgsptOneOffTags) =
        {
                NUM_ONEOFF_PROPS,
                {
                        PR_DISPLAY_NAME,
                        PR_ENTRYID,
                        PR_DEPTH,
                        PR_SELECTABLE,
                        PR_ADDRTYPE,
                        PR_DISPLAY_TYPE,
                        PR_INSTANCE_KEY,
                        PR_ROWID
                }
        };

        SPropValue rgsPropValue[NUM_ONEOFF_PROPS] = {0};

        if (FAILED(hRes = CreateTable((LPIID)&IID_IMAPITableData,
                                      vpfnAllocBuff,
                                      vpfnAllocMore,
                                      vpfnFreeBuff,
                                      NULL,
                                      TBLTYPE_DYNAMIC,
                                      PR_ROWID,
                                      (LPSPropTagArray) &rgsptOneOffTags,
                                      &pMyTbl)))
                goto LCleanup;


        rgsPropValue[NAME].ulPropTag = PR_DISPLAY_NAME;
        rgsPropValue[NAME].Value.LPSZ = SZ_NEW_CONTACT;

        assert(m_pLogon);
        if (FAILED(hRes = m_pLogon->HrNewEID(0,
                                             ONEOFF_USER_ID,
                                             MAPI_MAILUSER,
                                             NULL,
                                             &pEID)))
                goto LCleanup;

        rgsPropValue[EID].ulPropTag          = PR_ENTRYID;
        rgsPropValue[EID].Value.bin.cb       = sizeof(SABP_EID);
        rgsPropValue[EID].Value.bin.lpb      = (LPBYTE) pEID;

        rgsPropValue[DEPTH].ulPropTag        = PR_DEPTH;
        rgsPropValue[DEPTH].Value.l          = 0;

        rgsPropValue[SELECT].ulPropTag       = PR_SELECTABLE;
        rgsPropValue[SELECT].Value.b         = TRUE;

        rgsPropValue[ADDR_TYPE].ulPropTag    = PR_ADDRTYPE;
        rgsPropValue[ADDR_TYPE].Value.LPSZ   = NULL_STRING;

        rgsPropValue[DISP_TYPE].ulPropTag    = PR_DISPLAY_TYPE;
        rgsPropValue[DISP_TYPE].Value.l      = DT_MAILUSER;

        rgsPropValue[INST_KEY].ulPropTag     = PR_INSTANCE_KEY;
        rgsPropValue[INST_KEY].Value.bin.cb  = sizeof(ULONG);
        rgsPropValue[INST_KEY].Value.bin.lpb = (LPBYTE) &(pEID->ID.Num);

        rgsPropValue[ROWID].ulPropTag        = PR_ROWID;
        rgsPropValue[ROWID].Value.l          = ulRow++;


        sRow.cValues = NUM_ONEOFF_PROPS;
        sRow.lpProps = rgsPropValue;

        if (FAILED(hRes = pMyTbl->HrModifyRow(&sRow)))
                goto LCleanup;

        hRes = HrGetView(pMyTbl, NULL, NULL, NULL, lppTable);

LCleanup:
        vpfnFreeBuff(pEID);

        if (pMyTbl)
                pMyTbl->Release();

        if (FAILED(hRes))
                if (*lppTable)
                {
                        (*lppTable)->Release();
                        *lppTable = NULL;
                }

        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    GetContentsTable()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//          If there is underlying data in the cache, return a view, otherwise
//          create the underlying data, cache it and return a view.

//          Contents tables are sorted ascending on PR_DISPLAY_NAME
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::GetContentsTable(ULONG         /* ulFlags */,
                                            LPMAPITABLE  *ppTable)
{
        HRESULT         hRes    = S_OK;
        LPMAPITABLE     pTable  = NULL;

        CheckParameters_IMAPIContainer_GetContentsTable(this,
                                                        ulFlags,
                                                        ppTable);

        SizedSSortOrderSet(1, sortOrder) =
        {
                1,

                0,

                0,

                {
                        PR_DISPLAY_NAME,

                        TABLE_SORT_ASCEND
                }
        };

        *ppTable = NULL;

        EnterCriticalSection(&m_cs);

        if (!m_pCntsTblData)
        {
                if (FAILED(hRes = HrCreateCntsTable(&m_pCntsTblData)))
                        goto LCleanup;
        }

        if (m_pCntsTblData)
        {
                if (FAILED(hRes = HrGetView(m_pCntsTblData, NULL, NULL, 0, &pTable)))
                        goto LCleanup;
        }

        if (FAILED(hRes = pTable->SortTable((LPSSortOrderSet)&sortOrder, 0)))
                goto LCleanup;

        *ppTable = pTable;

LCleanup:
        LeaveCriticalSection(&m_cs);

        if (FAILED(hRes) && pTable)
                pTable->Release();

        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrCreateCntsTable()
//
//    Parameters

//          ppCTbl     [out] pointer to pointer to the content table to create
//
//    Purpose
//          Create a LPTABLEDATA object to contain the content table rows.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::HrCreateCntsTable(LPTABLEDATA *ppCTbl)
{
        if (!ppCTbl)
                return MAPI_E_INVALID_PARAMETER;

        HRESULT         hRes     = S_OK;
        LPTABLEDATA     pTblData = NULL;
        PCRecord        pRec     = NULL;
        LPSPropValue    pProps   = NULL;

        *ppCTbl = NULL;

        if (FAILED(hRes = CreateTable(&IID_IMAPITableData,
                                      vpfnAllocBuff,
                                      vpfnAllocMore,
                                      vpfnFreeBuff,
                                      0,
                                      TBLTYPE_DYNAMIC,
                                      PR_INSTANCE_KEY,
                                      (LPSPropTagArray)&vsptCntTbl,
                                      &pTblData)))
        {
                goto LCleanup;
        }

        assert(m_pDB);
        m_pDB->GoBegin();
        pRec = m_pDB->Read();

        // Load Rows
        while (pRec)
        {
                assert(m_pLogon);
                if (FAILED(hRes = m_pLogon->HrGetCntsTblRowProps(pRec, &pProps)))
                        goto LCleanup;

                SRow sr = {0};
                sr.cValues = NUM_CTBL_COLS;
                sr.lpProps = pProps;

                if (FAILED(hRes = pTblData->HrModifyRow(&sr)))
                        goto LCleanup;

                // Release the record copy generated by Read().
                pRec->Release();
                pRec = NULL;

                vpfnFreeBuff(pProps);


                pRec = m_pDB->Read();
        }

        *ppCTbl = pTblData;

LCleanup:
        if (FAILED(hRes))
        {
                if (pTblData)
                        pTblData->Release();
        }
        if (pRec)
                pRec->Release(); // if pRec was not released, release it.

        if (pProps)
                vpfnFreeBuff(pProps);

        return hRes;
}


///////////////////////////////////////////////////////////////////////////////
//
//    GetHierarchyTable()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//          Returns a view of
//          the cached hierarchy table data if it exists. If it doesn't
//          the table data is created and cached first.

//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::GetHierarchyTable(ULONG         /* ulFlags */,
                                             LPMAPITABLE  *ppTable)
{
        HRESULT         hRes    = S_OK;
        LPMAPITABLE     pTable  = NULL;

        CheckParameters_IMAPIContainer_GetHierarchyTable(this,
                                                         ulFlags,
                                                         ppTable);

        *ppTable = NULL;

        // This sample provider allows only root container to have subcontainers.
        if (m_EID.ID.Num != ROOT_ID_NUM)
                return MAPI_E_NO_SUPPORT;

        EnterCriticalSection(&m_cs);

        if (!m_pHierTblData)
        {
                // Create root HierTbl.
               if (FAILED(hRes = HrCreateHierTable(&m_pHierTblData)))
                       goto LCleanup;
        }

        if (m_pHierTblData)
        {
                // Get table view.
                if(FAILED(hRes = HrGetView(m_pHierTblData, NULL, NULL, 0, &pTable)))
                        goto LCleanup;
        }

        *ppTable = pTable;

LCleanup:
        LeaveCriticalSection(&m_cs);

        if (FAILED(hRes) && pTable)
                pTable ->Release();

        return hRes;
}


///////////////////////////////////////////////////////////////////////////////
//
//    CreateHierTbl()
//
//    Parameters

//          ppHTbl  [out] pointer to pointer to the hierarchy table to create.
//
//    Purpose
//          Creates the underlying data for the container's

//          hierarchy table, caches it,and returns a view.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::HrCreateHierTable(LPTABLEDATA *ppHTbl)
{
        if (!ppHTbl)
                return MAPI_E_INVALID_PARAMETER;

        HRESULT         hRes     = S_OK;
        LPTABLEDATA     pTblData = NULL;
        SRow            sr       = {0};
        LPSPropValue    pProps = NULL;

        if (FAILED(hRes = CreateTable(&IID_IMAPITableData,
                                      vpfnAllocBuff,
                                      vpfnAllocMore,
                                      vpfnFreeBuff,
                                      0,
                                      TBLTYPE_DYNAMIC,
                                      PR_INSTANCE_KEY,
                                      (LPSPropTagArray)&vsptHierTbl,
                                      &pTblData)))
        {
                goto LCleanup;
        }


        if (FAILED(hRes = HrLoadHierTblProps(&pProps)))
                goto LCleanup;

        sr.cValues = NUM_HTBL_PROPS;
        sr.lpProps = pProps;

        if (FAILED(hRes = pTblData->HrModifyRow(&sr)))
                goto LCleanup;

        *ppHTbl = pTblData;


LCleanup:
        if (FAILED(hRes) && pTblData)
                pTblData->Release();

        vpfnFreeBuff((LPVOID)pProps);

        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrLoadHierTblProps()
//
//    Parameters

//          ppProps  [out] pointer to pointer to SPropValue
//
//    Purpose
//          Load props for a row of hierarchy table.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::HrLoadHierTblProps(LPSPropValue *ppProps)
{
        HRESULT         hRes    = S_OK;
        PSABP_EID       pDirEID = {0};
        LPSPropValue    pProps  = NULL;

        // Generate a EntryID for the hierarchy table row.
        // Since the SABP support only one subcontainer,
        // generate a regular EntryID.
        if (FAILED(hRes = m_pLogon->HrNewEID(0,
                                             DIR_ID_NUM,
                                             MAPI_ABCONT,
                                             NULL,
                                             &pDirEID)))
        {
                goto LCleanup;
        }

        if (FAILED(hRes =

                vpfnAllocBuff(NUM_HTBL_PROPS * sizeof SPropValue, (LPVOID*)&pProps)))
        {
                goto LCleanup;
        }

        if (!pProps)
        {
                hRes = MAPI_E_NOT_ENOUGH_MEMORY;
                goto LCleanup;
        }

        // Props fo a hierarchy table row is defined in the header file abp.h
        memset((LPVOID)pProps, 0, NUM_HTBL_PROPS  * sizeof SPropValue);

        pProps[HT_INST_KEY].ulPropTag     = PR_INSTANCE_KEY;
        pProps[HT_INST_KEY].Value.bin.cb  = sizeof(pDirEID->ID.Num);
        pProps[HT_INST_KEY].Value.bin.lpb = (LPBYTE) & (pDirEID->ID.Num);

        pProps[HT_DEPTH].ulPropTag      = PR_DEPTH;
        pProps[HT_DEPTH].Value.l        = 0;

        pProps[HT_ACCESS].ulPropTag     = PR_ACCESS;
        pProps[HT_ACCESS].Value.l       = MAPI_ACCESS_READ;

        pProps[HT_EID].ulPropTag        = PR_ENTRYID;
        pProps[HT_EID].Value.bin.cb     = sizeof(*pDirEID);
        pProps[HT_EID].Value.bin.lpb    = (LPBYTE) pDirEID;

        pProps[HT_OBJ_TYPE].ulPropTag   = PR_OBJECT_TYPE;
        pProps[HT_OBJ_TYPE].Value.l     = MAPI_ABCONT;

        pProps[HT_DISP_TYPE].ulPropTag  = PR_DISPLAY_TYPE;
        pProps[HT_DISP_TYPE].Value.l    = DT_MODIFIABLE;

        pProps[HT_NAME].ulPropTag       = PR_DISPLAY_NAME;
        pProps[HT_NAME].Value.LPSZ      = SZ_PSS_SAMPLE_AB;

        pProps[HT_CONT_FLAGS].ulPropTag = PR_CONTAINER_FLAGS;
        pProps[HT_CONT_FLAGS].Value.l   = AB_RECIPIENTS | AB_MODIFIABLE;

        pProps[HT_STATUS].ulPropTag     = PR_STATUS;
        pProps[HT_STATUS].Value.l       = 0;

        pProps[HT_COMMENT].ulPropTag    = PR_COMMENT;
        pProps[HT_COMMENT].Value.LPSZ   = SZ_PSS_SAMPLE_AB;

        pProps[HT_PROV_UID].ulPropTag    = PR_AB_PROVIDER_ID;
        pProps[HT_PROV_UID].Value.bin.cb  = sizeof MAPIUID;
        pProps[HT_PROV_UID].Value.bin.lpb = (LPBYTE) & SABP_UID;

        *ppProps = pProps;

LCleanup:
        if (FAILED(hRes))
        {
                if (pProps)
                        vpfnFreeBuff(pProps);
                if (pDirEID)
                        vpfnFreeBuff(pDirEID);
        }
        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    OpenEntry()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          MAPI_E_NO_SUPPORT
//
STDMETHODIMP CABContainer::OpenEntry(ULONG       /* cbEntryID */,
                                     LPENTRYID   /* pEntryID */,
                                     LPCIID      /* pInterface */,

                                     ULONG       /* ulFlags */,
                                     ULONG     * /* pulObjType */,
                                     LPUNKNOWN * /* ppUnk */)
{
        return MAPI_E_NO_SUPPORT;
}


///////////////////////////////////////////////////////////////////////////////
//
//    HrCreateMailUser()
//
//    Parameters

//          pEID        [in] Temp EntryID of the mail user to create.
//
//    Purpose
//

//
//    Return Value
//          MAPI_E_USER_CANCEL for success, so MAPI will do nothing more.
//          others except S_OK for failures.
//          S_OK should never be returned, use MAPI_E_USER_CANCEL instead.
//
STDMETHODIMP CABContainer::HrCreateMailUser(PSABP_EID pEID)
{
        if (!pEID)
                return MAPI_E_INVALID_PARAMETER;

        HRESULT         hRes = S_OK;
        LPMAPITABLE     pTbl = NULL;
        LPCMailUser     pMailUser = NULL;

        pMailUser = new CMailUser(m_hInst,
                                  pEID,
                                  m_pLogon,
                                  m_pSupObj,
                                  m_pDB);
        if (!pMailUser)
                return MAPI_E_NOT_ENOUGH_MEMORY;

        if (FAILED(hRes = pMailUser->HrInit()))
                return hRes;

        // Get LPMAILTABLE from CMailUser's PR_DETAILS_TABLE property
        if (FAILED(hRes = pMailUser->OpenProperty(PR_DETAILS_TABLE,
                                                  (LPCIID)&IID_IMAPITable,
                                                  0,
                                                  0,
                                                  (LPUNKNOWN*)&pTbl)))
        {
                return hRes;
        }

        if (!pTbl)
                return MAPI_E_NOT_ENOUGH_MEMORY;

        // Display prop config dialog.
        hRes = m_pSupObj->DoConfigPropsheet(0, 0, NULL_STRING, pTbl, pMailUser, 0);

        if (FAILED(hRes))
                return hRes;
        else
                return MAPI_E_USER_CANCEL;
}

///////////////////////////////////////////////////////////////////////////////
//
//    CreateEntry()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//          Creates a new entry in this container. A new entry can be a
//          mailuser, a DL (a container), a foreign DL or foreign mailuser.
//          The entry ID tells us whether the entry is foreign or native.

//          Creates a new record in the database, It's in a limbo state,
//          (so add to orphan rec list, until SaveChanges, then

//          insert it in its parent).

//
//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::CreateEntry(ULONG        cbEntryID,
                                       LPENTRYID    lpEntryID,
                                       ULONG        /* ulCreateFlags */,
                                       LPMAPIPROP  *lppMAPIPropEntry)
{
        HRESULT         hRes     = S_OK;
        PSABP_EID       pEID     = (PSABP_EID)lpEntryID;
        ULONG           ulEntryType;

        CheckParameters_IABContainer_CreateEntry(this,
                                                 cbEntryID,
                                                 lpEntryID,
                                                 ulCreateFlags,
                                                 lppMAPIPropEntry);

        EnterCriticalSection(&m_cs);
        assert(m_pLogon);
        ulEntryType = m_pLogon->CheckEID(cbEntryID, lpEntryID);
        *lppMAPIPropEntry = NULL;

        switch(ulEntryType)
	{
        case MAPI_MAILUSER:
                hRes = HrCreateMailUser(pEID);
                break;

        default:
                hRes = MAPI_E_NO_SUPPORT;
                break;
        }

        LeaveCriticalSection(&m_cs);
        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    CopyEntries()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::CopyEntries(LPENTRYLIST        /* lpEntries */,
                                       ULONG              /* ulUIParam */,
                                       LPMAPIPROGRESS     /* lpProgress */,
                                       ULONG              /* ulFlags */)
{
        return MAPI_E_NO_SUPPORT;
}

///////////////////////////////////////////////////////////////////////////////
//
//    DeleteEntries()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//          Converts the
//          lpEntries into a list and deletes each child.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::DeleteEntries(LPENTRYLIST  lpEntries,
                                         ULONG        /* ulFlags */)
{
        HRESULT         hRes = S_OK;
        SPropValue      prop = {0};
        SizedSPropTagArray(1, sptTag) = {1, PR_CONTAINER_CONTENTS};

        assert(m_pDB);

        for(ULONG i=0; i<lpEntries->cValues; i++)
        {
                PSABP_EID pEID = (PSABP_EID)lpEntries->lpbin[i].lpb;

                prop.ulPropTag = PR_INSTANCE_KEY;
                prop.Value.bin.cb = sizeof(pEID->ID.Num);
                prop.Value.bin.lpb = (LPBYTE) & (pEID->ID.Num);

                if (m_pCntsTblData)
                {
                        if (FAILED(m_pCntsTblData->HrDeleteRow(&prop)))
                        {
                                hRes = MAPI_W_PARTIAL_COMPLETION;
                                goto LCleanup;
                        }
                }

                // Remove the entry from database
                if (FAILED(m_pDB->Delete(pEID->ID.Num)))
                {
                        hRes = MAPI_W_PARTIAL_COMPLETION;
                        goto LCleanup;
               }

        }

        hRes = m_pDB->Save();

LCleanup:
        vpfnFreeBuff(&prop);
        vpfnFreeBuff(&sptTag);

        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    DoANR()
//
//    Parameters

//          szTarget        [in]  recipient name to resolve
//          pSearchTbl      [in]  table to look for recipient in
//          psrResolved     [out] resolved ADRENTRY
//          ulNameCol       [in]  index into table rows props where to find
//                                the string property for the comparison
//    Purpose
//          This fcn does ambiguous name resolution on szTarget in
//          pSearchTable. Return a single row in *psrResolved if

//          a match is found and pass back NULL and return MAPI_E_AMBIGUOUS or

//          MAPI_E_NOT_FOUND otherwise.
//
//          If a match is not found,  restrict the container's
//          contents table to rows that have the target string's first token as a

//          substring in the display name. Then run a simple parser on each row
//          to determine if the rest of the tokens match exactly, partially match

//          (ambiguous), or don't match at all
//
//          A match is defined to be one and only one exact match (token for token)
//          or exactly 1 ambiguous (partial) match.
//
//          Do this because MAPI defines an ambiguous match to be > 1 match.

//          The implementation determines what is meant by a match. If there is
//          a single non-exact match, that is, returns a single row in Restrict w/
//          PR_ANR, the CheckNames dialog treats it like a match anyway.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP DoANR(LPTSTR        szTarget,
                   LPMAPITABLE   pSearchTbl,
                   LPSRow        psrResolved,
                   ULONG         ulNameCol)
{
        if (!szTarget || !pSearchTbl || !psrResolved)
                return MAPI_E_INVALID_PARAMETER;

        HRESULT         hRes    = S_OK;
        LPSRowSet       pHitRow = NULL;

        TCHAR           szBuf[MAX_PATH];
        LPTSTR          szTemp  = szBuf;
        ULONG           ulMatches = 0;
        SPropValue      spTemp  = {0};
        SRestriction    SRes    = {0};

        lstrcpyn(szTemp, szTarget, MAX_PATH);
        szTemp[MAX_PATH - 1]    = '\0';
        spTemp.ulPropTag        = ulNameCol;
        spTemp.Value.LPSZ       = GetNextToken(&szTemp);

        // find all rows with the target string's first token
        SRes.rt = RES_CONTENT;
        SRes.res.resContent.ulPropTag    = ulNameCol;
        SRes.res.resContent.lpProp       = &spTemp;
        SRes.res.resContent.ulFuzzyLevel = FL_SUBSTRING | FL_IGNORECASE;

        if (FAILED(hRes = HrQueryAllRows(pSearchTbl,
                                         NULL,
                                         &SRes,
                                         NULL,
                                         0,
                                         &pHitRow)))
        {
                goto LCleanup;
        }

        if (pHitRow)
                ulMatches = pHitRow->cRows;

LCleanup:
        if (ulMatches > 1)
        {
                hRes = MAPI_E_AMBIGUOUS_RECIP;
        }
        else if (ulMatches == 1)
        {
                // pull matching row out of rowset, save it
                *psrResolved = pHitRow->aRow[0];
                pHitRow->aRow[0].cValues = 0;
                pHitRow->aRow[0].lpProps = NULL;
                hRes = S_OK;
        }
        else
        {
                hRes = MAPI_E_NOT_FOUND;
        }
        FreeProws(pHitRow);
        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    CopyAdrEntry()
//
//    Parameters

//

//
//    Purpose
//          Copies an SRow's properties into an ADRENTRY without duplicates,
//          reallocating a larger ADRENTRY if necessary. The old ADRENTRY

//          property array is freed and the new one is plugged into its
//          place.

//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//

STDMETHODIMP CABContainer::HrCopyAdrEntry(LPSPropTagArray pTags,

                                          LPSRow          pRow,
                                          LPADRENTRY      pAdrEntry)
{
        if (!pTags || !pRow || !pAdrEntry)
                return MAPI_E_INVALID_PARAMETER;

        HRESULT         hRes      = S_OK;
        LPSPropValue    pNewProps = NULL;

        list<ULONG>     MergedProps;
        ULONG           cExistingProps;
        ULONG           cTotalProps;
        ULONG           ulPropTag;

        // Make a merged list of existing props on the unresolved AdrEntry,
        // and requested props on the resolved one. Eliminate the duplicates.
        // (that is, pTags and *ppAdrEntry can have tags in common).

        // Want a count of the tags without counting common ones

        for (ULONG i = 0; i < pAdrEntry->cValues; i++)
                MergedProps.push_back(pAdrEntry->rgPropVals[i].ulPropTag);

        // cTotalProps is in case of failure, use it to free
        cExistingProps = MergedProps.size();

        for (ULONG i = 0; i < pTags->cValues; i++)
        {
                BOOL existed = FALSE;
                ULONG ulTag = pTags->aulPropTag[i];
                for(list<ULONG>::iterator it = MergedProps.begin();
                        it != MergedProps.end(); it++)
                {
                        if (*it == ulTag)
                        {
                                existed = TRUE;
                                break;
                        }
                }
                if (!existed)
                        MergedProps.push_back(ulTag);

        }

        cTotalProps = MergedProps.size();


        if (FAILED(hRes = vpfnAllocBuff(cTotalProps  *sizeof SPropValue,
                                        (LPVOID *) & pNewProps)))
        {
                goto LCleanup;
        }

        // if failure, FreeBuff won't complain
        ZeroMemory(pNewProps, cTotalProps  *sizeof SPropValue);

        // copy existing props on the adr entry to new AdrEntry
        for (ULONG i = 0; i < cExistingProps; i++)
        {
                MergedProps.pop_front();  // remove from list

                if (FAILED(hRes = HrCopyProp(&(pNewProps[i]), &(pAdrEntry->rgPropVals[i]),pNewProps)))
                        goto LCleanup;
        }

        // copy props from SRow to new adrentry
        // you can never hit a prop in the row that's already in
        // the adrentry because the merged list has no duplicates
        for (ULONG i = cExistingProps; i < cTotalProps; i++)
        {
                BOOL bFound = FALSE;

                ulPropTag = (ULONG)*(MergedProps.begin());
                for (ULONG j = 0; j < pRow->cValues; j++)
                {
                        if (PROP_ID(ulPropTag) == PROP_ID(pRow->lpProps[j].ulPropTag))
                        {

                                bFound = TRUE;

                                // uses MAPIAllocateMore, so free it later
                                if (FAILED(hRes = HrCopyProp(&(pNewProps[i]),

                                                               &(pRow->lpProps[j]),
                                                               pNewProps)))
                                        goto LCleanup;
                                break;

                        }
                }

                assert(bFound == TRUE);

                MergedProps.pop_front();
        }


LCleanup:

        if (FAILED(hRes))
        {
                vpfnFreeBuff(pNewProps);
                return hRes;
        }

        vpfnFreeBuff(pAdrEntry->rgPropVals);
        pAdrEntry->rgPropVals = pNewProps;
        pAdrEntry->cValues = cTotalProps;
        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    ResolveNames()
//
//    Refer to the MSDN documentation for more information.
//
//    Purpose
//          Resolves names against the container's contents table.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//

STDMETHODIMP CABContainer::ResolveNames(LPSPropTagArray lpPropTagArray,

                                        ULONG           ulFlags,

                                        LPADRLIST       lpAdrList,

                                        LPFlagList      lpFlagList)
{
        HRESULT           hRes     = S_OK;

        LPMAPITABLE       pCntTbl  = NULL;
        SRow              ResolvedRow = {0, 0, NULL};
        LPSPropTagArray   pCols;

        Validate_IABContainer_ResolveNames(this,
                                           lpPropTagArray,
                                           ulFlags,
                                           lpAdrList,
                                           lpFlagList);

        pCols = (!lpPropTagArray ? (LPSPropTagArray) &vsptCntTbl : lpPropTagArray);
        EnterCriticalSection(&m_cs);

        // contents table is also sorted
        if (FAILED(hRes = GetContentsTable(0,&pCntTbl)))
                goto LCleanup;

        if (FAILED(hRes = pCntTbl->SetColumns(pCols, TBL_BATCH)))
                goto LCleanup;


        // if a name is unresolved, try to parse it
        for (ULONG i = 0; i < lpFlagList->cFlags; i++)

        {

                if (MAPI_UNRESOLVED == lpFlagList->ulFlag[i])
                {
                        LPSPropValue pProp = PpropFindProp(lpAdrList->aEntries[i].rgPropVals,
                                lpAdrList->aEntries[i].cValues,
                                PR_DISPLAY_NAME);
                        if (!pProp)
                                continue;

                        hRes = DoANR(pProp->Value.LPSZ,
                                     pCntTbl,

                                     &ResolvedRow,
                                     PR_DISPLAY_NAME);

                        switch (hRes)
                        {
                        case S_OK :

                                // copy props to ADRENTRY
                                lpFlagList->ulFlag[i] = MAPI_RESOLVED;
                                hRes = HrCopyAdrEntry(pCols,
                                                      &ResolvedRow,
                                                      &(lpAdrList->aEntries[i]));
                                break;

                        case MAPI_E_AMBIGUOUS_RECIP:
                                lpFlagList->ulFlag[i] = MAPI_AMBIGUOUS;
                                hRes = S_OK;
                                break;

                        case MAPI_E_NOT_FOUND:
                                hRes = S_OK;
                                break;

                        default:
                                break;
                        }
                }

        }


LCleanup:
        vpfnFreeBuff((LPVOID)ResolvedRow.lpProps);

        if (pCntTbl)
                pCntTbl->Release();

        LeaveCriticalSection(&m_cs);

        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    SetSearchCriteria()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::SetSearchCriteria(LPSRestriction /* pRestriction */,
                                             LPENTRYLIST    /* pContainerList */,
                                             ULONG          /* ulSearchFlags */)
{
        CheckParameters_IMAPIContainer_SetSearchCriteria(this,
                                                         pRestriction,
                                                         pContainerList,
                                                         ulSearchFlags);


        return MAPI_E_NO_SUPPORT;
}


///////////////////////////////////////////////////////////////////////////////
//
//    GetSearchCriteria()
//
//    Refer to the MSDN documentation for more information.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::GetSearchCriteria(ULONG             /* ulFlags */,
                                             LPSRestriction  * /* ppRestriction */,
                                             LPENTRYLIST     * /* ppContainerList */,
                                             ULONG           * /* pulSearchState */)
{
        CheckParameters_IMAPIContainer_GetSearchCriteria(this,
                                                         ulFlags,

                                                         ppRestriction,

                                                         ppContainerList,
                                                         pulSearchState);

        return MAPI_E_NO_SUPPORT;
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrGetView()
//
//    Parameters

//          pTblData         [in]  pointer to underlying table data
//          lpSSortOrderSet  [in]  view sort order
//          lpfCallerRelease [in]  view release callback
//          ulCallerData     [in]  data to call callback with
//          lppMAPITable     [out] pointer to pointer where view is
//                                 returned
//    Purpose
//          Returns a view of the wrapper table implementation.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CABContainer::HrGetView(LPTABLEDATA         pTblData,
                                     LPSSortOrderSet     lpSSortOrderSet,

                                     CALLERRELEASE FAR  *lpfCallerRelease,

                                     ULONG               ulCallerData,

                                     LPMAPITABLE FAR    *lppMAPITable)
{
        HRESULT      hRes = S_OK;
        CMAPITable  *pTbl = NULL;
        LPMAPITABLE  pTblTemp = NULL;

        if (FAILED(hRes = pTblData->HrGetView(lpSSortOrderSet,
                                    lpfCallerRelease,
                                    ulCallerData,
                                    &pTblTemp)))
        {
                return hRes;
        }

        pTbl = new CMAPITable(pTblTemp, m_EID.ID.Num);

        if (!pTbl)
        {
                hRes = MAPI_E_NOT_ENOUGH_MEMORY;
                goto LCleanup;
        }

        *lppMAPITable = (LPMAPITABLE)pTbl;
        return hRes;

LCleanup:

        if (pTblTemp)
                pTblTemp->Release();

        if (pTbl)
                pTbl->Release();

        return hRes;

}

///////////////////////////////////////////////////////////////////////////////
//
//	Implementation of class CTblAdviseSink
//
//

///////////////////////////////////////////////////////////////////////////////
//
//    CTblAdviseSink()

//
//    Purpose
//          Class constructor.
//

//    Return Value
//          none.
CTblAdviseSink::CTblAdviseSink(LPCABContainer pCont)
: m_cRef(1),
  m_pCont(pCont)
{
        InitializeCriticalSection(&m_cs);
}

///////////////////////////////////////////////////////////////////////////////
//
//    ~CTblAdviseSink()

//
//    Purpose
//          Class destructor.
//

//    Return Value
//          none.
CTblAdviseSink::~CTblAdviseSink()
{
        DeleteCriticalSection(&m_cs);
}

///////////////////////////////////////////////////////////////////////////////
//
//    AddRef()
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
STDMETHODIMP_ (ULONG) CTblAdviseSink::AddRef()
{
        return InterlockedIncrement(&m_cRef);
}

///////////////////////////////////////////////////////////////////////////////
//
//    Release()
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
STDMETHODIMP_ (ULONG) CTblAdviseSink::Release()
{
        ULONG ulRefCount = InterlockedDecrement(&m_cRef);

        // If this object is not being, delete it
        if (ulRefCount == 0)
        {
                // Destroy the object
                delete this;

        }

        return ulRefCount;
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
STDMETHODIMP CTblAdviseSink::QueryInterface(const IID &riid, LPVOID *ppvObj)
{
        if (IsBadWritePtr(ppvObj, sizeof LPUNKNOWN))
                return MAPI_E_INVALID_PARAMETER;

        *ppvObj = NULL;

        if (riid == IID_IUnknown || riid == IID_IMAPIAdviseSink)
        {
                *ppvObj = static_cast<LPVOID>(this);
                AddRef();
                return NOERROR;
        }

        return E_NOINTERFACE;
}

///////////////////////////////////////////////////////////////////////////////
//
//    OnNotify()
//
//    Refer to the MSDN documentation for more information.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise
//

STDMETHODIMP_ (ULONG) CTblAdviseSink::OnNotify(ULONG cNotif, LPNOTIFICATION pNotifs)
{
        HRESULT hRes = S_OK;

        if (!pNotifs)
        {
                hRes = MAPI_E_INVALID_PARAMETER;
                return hRes;
        }

        for (ULONG i = 0; i < cNotif; i++)
        {
                switch (pNotifs[i].ulEventType)
                {
                case fnevTableModified:
                        switch (pNotifs[i].info.tab.ulTableEvent)
                        {

                        case TABLE_ROW_ADDED:
                                // If a child was created, changed.Notify
                                m_pCont->m_pLogon->HrNotify(fnevObjectModified,
                                                            NULL,
                                                            &m_pCont->m_EID,
                                                            NULL);
                                // Fall through

                        case TABLE_CHANGED:

                                // Fall through

                        case TABLE_ROW_MODIFIED:
                                // update contents table on any change in children
                                if (m_pCont && m_pCont->m_pCntsTblData)
                                {
                                        hRes = m_pCont->m_pCntsTblData->HrModifyRow(&(pNotifs[i].info.tab.row));
                                }
                                break;

                        default:
                                hRes = MAPI_E_NO_SUPPORT;
                                break;
                        }
                        break;

                case fnevObjectModified:
                        break;

                case fnevObjectCreated:
                        break;

                case fnevObjectDeleted:
                        break;

                default:

                        hRes = MAPI_E_NO_SUPPORT;
                }
        }

    return hRes;
}
/* End of the file ABCONT.CPP */