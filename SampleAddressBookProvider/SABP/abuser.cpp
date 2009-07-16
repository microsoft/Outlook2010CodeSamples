/****************************************************************************
        ABUSER.CPP

        Copyright (c) 2007 Microsoft Corporation

        Implementation of IMailUser,
        which is implemented by messaging user objects.
*****************************************************************************/

#include "initguid.h"
#define USES_IID_IMAPIProp
#define USES_IID_IMailUser
#define USES_IID_IMAPIPropData
#define USES_IID_IMAPITable

#include "abp.h"
#include "resource.h"
#include "util.h"

#define SABP_HLP        _T("SABP Help")
#define ASTERRISK       _T("*")

#define DT_EDIT_FLAGS       (DT_EDITABLE | DT_REQUIRED)

DTBLPAGE     UserPage      = {sizeof(DTBLPAGE)    , fMapiUnicode};
DTBLEDIT     UserName      = {sizeof(DTBLEDIT)    , fMapiUnicode, MAX_PATH , PR_DISPLAY_NAME};
DTBLEDIT     UserAddrType  = {sizeof(DTBLEDIT)    , fMapiUnicode, MAX_PATH , PR_ADDRTYPE};
DTBLEDIT     UserEmailAddr = {sizeof(DTBLEDIT)    , fMapiUnicode, MAX_PATH , PR_EMAIL_ADDRESS};
DTBLEDIT     UserDpt       = {sizeof(DTBLEDIT)    , fMapiUnicode, MAX_PATH , PR_DEPARTMENT_NAME};
DTBLLABEL    DtLabel       = {sizeof(DTBLLABEL)   , fMapiUnicode};

DTCTL  GUserControls[] =

       {
          {DTCT_PAGE    ,0             ,NULL ,0 ,NULL       ,0                  ,&UserPage     },
          {DTCT_EDIT    ,DT_EDIT_FLAGS ,NULL ,0 ,ASTERRISK  ,IDC_EDIT_NAME      ,&UserName     },
          {DTCT_EDIT    ,DT_EDIT_FLAGS ,NULL ,0 ,ASTERRISK  ,IDC_EDIT_ADDRTYPE	,&UserAddrType },
          {DTCT_EDIT    ,DT_EDIT_FLAGS ,NULL ,0 ,ASTERRISK  ,IDC_EDIT_ADDR      ,&UserEmailAddr},
          {DTCT_EDIT    ,DT_EDIT_FLAGS ,NULL ,0 ,ASTERRISK  ,IDC_EDIT_DPT       ,&UserDpt      },
          {DTCT_LABEL   ,0             ,NULL ,0 ,NULL       ,IDC_STATIC_NAME    ,&DtLabel      },
          {DTCT_LABEL   ,0             ,NULL ,0 ,NULL       ,IDC_STATIC_ADDRTYPE,&DtLabel      },
          {DTCT_LABEL   ,0             ,NULL ,0 ,NULL       ,IDC_STATIC_ADDR    ,&DtLabel      },
          {DTCT_LABEL   ,0             ,NULL ,0 ,NULL       ,IDC_STATIC_DPT     ,&DtLabel      },
       };


DTPAGE GUserDetailsPage =

{
        _countof(GUserControls),
        MAKEINTRESOURCE(IDD_USERDETAILS_PAGE),
        SABP_HLP,

        GUserControls
};


///////////////////////////////////////////////////////////////////////////////
//
//    CMailUser()
//
//    Parameters

//
//    Purpose
//          Class constructor.
//

//    Return Value
//          none.
//
CMailUser::CMailUser(HINSTANCE          hInst,
                     PSABP_EID          pEID,
                     LPCABLogon         pLogon,
                     LPMAPISUP          pSupObj,
                     PCABDatabase       pDB)
: m_cRef(1),
  m_hInst(hInst),
  m_pLogon(pLogon),
  m_pSupObj(pSupObj),
  m_pDB(pDB),
  m_pPropData(NULL),
  m_fCreateFlag(FALSE),
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

///////////////////////////////////////////////////////////////////////////////
//
//    ~CMailUser()
//
//    Parameters

//          none.
//
//    Purpose
//          class destructor.
//

//    Return Value
//          none.
//
CMailUser::~CMailUser()
{
        if (m_pPropData)
                m_pPropData ->Release();

        if (m_pDB)
                m_pDB->Release();

        if (m_pSupObj)
                m_pSupObj->Release();

        if (m_pLogon)
                m_pLogon->Release();

        DeleteCriticalSection(&m_cs);
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrInit()
//
//    Parameters

//          none
//
//    Purpose
//          Initializes the new mailuser, the base IProp and sets default

//          properties.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CMailUser::HrInit()
{
        HRESULT hRes = S_OK;
        PCRecord pRec = NULL;

        if (m_fInited)
        {
                return hRes;
        }

        assert(m_pLogon);
        assert(m_pSupObj);

        if (FAILED(hRes = CreateIProp(&IID_IMAPIPropData,
                                      vpfnAllocBuff,
                                      vpfnAllocMore,
                                      vpfnFreeBuff,
                                      NULL,
                                      &m_pPropData)))
        {
                goto LCleanup;
        }

        if (!m_pPropData)
        {
                hRes = MAPI_E_NOT_ENOUGH_MEMORY;
                goto LCleanup;
        }

        m_pPropData->AddRef();

        assert(m_pDB);
        m_pDB->Query(m_EID.ID.Num, &pRec);
        m_fCreateFlag = (pRec == NULL);

        if (FAILED(hRes = HrSetDefaultProps(pRec)))
                goto LCleanup;

LCleanup:
        if (pRec)
                pRec->Release();

        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrSetDefaultProps()
//
//    Parameters

//          pRec        [in]  pointer to a CRecord object
//
//    Purpose
//          Initialize a MailUser's property
//          cache with default values.

//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//

STDMETHODIMP CMailUser::HrSetDefaultProps(PCRecord pRec)
{
        LPSPropValue pPropTmp = NULL;

        enum

        {
                EID,
                REC_KEY,
                SEARCH_KEY,
                OBJ_TYPE,
                DISP_TYPE,

                NAME,
                ADDRTYPE,
                ADDR,
                DPT,
                NUM_PROPS
        };
        const SizedSPropTagArray(NUM_PROPS, rgsptMailUser) =
        {
                NUM_PROPS,
                {
                        PR_ENTRYID,
                        PR_RECORD_KEY,
                        PR_SEARCH_KEY,
                        PR_OBJECT_TYPE,
                        PR_DISPLAY_TYPE,
                        PR_DISPLAY_NAME,
                        PR_ADDRTYPE,
                        PR_EMAIL_ADDRESS,
                        PR_DEPARTMENT_NAME
                }
        };

        SPropValue      rgProps[NUM_PROPS] = {0};

        // Set the regular props.
        rgProps[EID].ulPropTag            = PR_ENTRYID;
        rgProps[EID].Value.bin.cb         = sizeof(m_EID);
        rgProps[EID].Value.bin.lpb        = (LPBYTE) &m_EID;

        rgProps[REC_KEY].ulPropTag        = PR_RECORD_KEY;
        rgProps[REC_KEY].Value.bin.cb     = sizeof(m_EID.ID.Num);
        rgProps[REC_KEY].Value.bin.lpb    = (LPBYTE) & (m_EID.ID.Num);

        rgProps[SEARCH_KEY].ulPropTag     = PR_SEARCH_KEY;
        rgProps[SEARCH_KEY].Value.bin.cb  = sizeof(m_EID.ID.Num);
        rgProps[SEARCH_KEY].Value.bin.lpb = (LPBYTE) & (m_EID.ID.Num);

        rgProps[OBJ_TYPE].ulPropTag       = PR_OBJECT_TYPE;
        rgProps[OBJ_TYPE].Value.l         = MAPI_MAILUSER;

        rgProps[DISP_TYPE].ulPropTag      = PR_DISPLAY_TYPE;
        rgProps[DISP_TYPE].Value.l        = DT_MAILUSER;


        if (pRec)
        {
                // Read the saved props from the record.
                for(ULONG i=DISP_TYPE + 1; i<NUM_PROPS; i++)
                {
                        pPropTmp = PpropFindProp(pRec->lpProps(),
                                                 pRec->cValues(),
                                                 rgsptMailUser.aulPropTag[i]);
                        if (pPropTmp)
                        {
                                rgProps[i] = *pPropTmp;
                                pPropTmp = NULL;
                        }
                        else
                        {
                                HrPropNotFound(rgProps[i], rgsptMailUser.aulPropTag[i]);
                        }
                }
        }
        else
        {
                // Set the props to be default values.
                rgProps[NAME].ulPropTag         = PR_DISPLAY_NAME;
                rgProps[NAME].dwAlignPad        = 0;
                rgProps[NAME].Value.LPSZ        = NULL_STRING;

                rgProps[ADDRTYPE].ulPropTag     = PR_ADDRTYPE;
                rgProps[ADDRTYPE].dwAlignPad    = 0;
                rgProps[ADDRTYPE].Value.LPSZ    = NULL_STRING;

                rgProps[ADDR].ulPropTag         = PR_EMAIL_ADDRESS;
                rgProps[ADDR].dwAlignPad        = 0;
                rgProps[ADDR].Value.LPSZ        = NULL_STRING;

                rgProps[DPT].ulPropTag          = PR_EMAIL_ADDRESS;
                rgProps[DPT].dwAlignPad         = 0;
                rgProps[DPT].Value.LPSZ         = NULL_STRING;
        }

        return SetProps(NUM_PROPS, (LPSPropValue)&rgProps, NULL);
}


///////////////////////////////////////////////////////////////////////////
//    AddRef()
//
//    Parameters
//          See the OLE documentation for details.
//
//    Purpose
//      To increase the reference count of this object, and of any contained

//      (or aggregated) object within
//
//    Return Value
//      Usage count for this object
//
STDMETHODIMP_ (ULONG) CMailUser::AddRef (void)
{
        return InterlockedIncrement(&m_cRef);
}

///////////////////////////////////////////////////////////////////////////////
//
//    Release()
//
//    Parameters
//          See the OLE documentation for details.
//
//    Purpose
//      Decreases the usage count of this object. When the count is zero,
//      the object deletes itself.
//
//    Return Value
//      Usage of this object
//
STDMETHODIMP_ (ULONG) CMailUser::Release (void)
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
//          NOERROR on success, E_NOINTERFACE if the interface can't be
//          returned.
//

STDMETHODIMP CMailUser::QueryInterface(REFIID   riid,

                                       LPVOID  *ppvObj)
{
        if (IsBadWritePtr(ppvObj, sizeof LPUNKNOWN))
                return MAPI_E_INVALID_PARAMETER;

        *ppvObj = NULL;

        if (riid == IID_IUnknown || riid == IID_IMAPIProp || riid == IID_IMailUser)
        {
                *ppvObj = static_cast<LPVOID>(this);
                AddRef();
                return NOERROR;
        }

        return E_NOINTERFACE;
}

///////////////////////////////////////////////////////////////////////////////
//
//    GetLastError()
//
//    Purpose
//          Gets detailed error information based on hResult.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//

STDMETHODIMP CMailUser::GetLastError(HRESULT       hResult,

                                     ULONG         ulFlags,
                                     LPMAPIERROR  *lppMAPIError)
{
        CheckParameters_IMAPIProp_GetLastError(this,
                                               hResult,

                                               ulFlags,

                                               ppMAPIError);

        return m_pPropData->GetLastError(hResult,
                                         ulFlags,
                                         lppMAPIError);
}

///////////////////////////////////////////////////////////////////////////////
//
//    OpenProperty()
//
//    Purpose
//          Returns an object when requested a supported property
//          of type PT_OBJECT.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CMailUser::OpenProperty(ULONG       ulPropTag,
                                     LPCIID      lpiid,

                                     ULONG       /* ulInterfaceOptions */,

                                     ULONG       /* ulFlags */,

                                     LPUNKNOWN  *lppUnk)
{
        CheckParameters_IMAPIProp_OpenProperty(this,
                                               ulPropTag,
                                               lpiid,
                                               ulInterfaceOptions,
                                               ulFlags,
                                               lppUnk);

        HRESULT hRes = S_OK;

        EnterCriticalSection(&m_cs);

        switch (ulPropTag)
        {
        case PR_DETAILS_TABLE:
                if (*lpiid != IID_IMAPITable)
                {
                        hRes = MAPI_E_INTERFACE_NOT_SUPPORTED;
                        break;
                }


                hRes = BuildDisplayTable(vpfnAllocBuff,
                                         vpfnAllocMore,
                                         vpfnFreeBuff,
                                         NULL,
                                         m_hInst,
                                         1,
                                         &GUserDetailsPage,
                                         fMapiUnicode,
                                         (LPMAPITABLE*)lppUnk,
                                         NULL);

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
//    SaveChanges()
//
//    Purpose
//          See the MAPI documentation for details.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise
//

STDMETHODIMP  CMailUser::SaveChanges(ULONG /* ulFlags */)
{
        // The param ulFlags is not used.
        HRESULT         hRes    = S_OK;
        LPSPropValue    pProps  = NULL;
        ULONG           cValues = 0;
        PCRecord        pRec    = NULL;
        // For undo if save failed.
        RECORD_ID_TYPE  recID   = 0;
        const ULONG     ulcSavablePropsNum = 4;

        const SizedSPropTagArray(ulcSavablePropsNum, sptSavableProps) =
        {
                ulcSavablePropsNum,
                {
                        PR_DISPLAY_NAME,
                        PR_ADDRTYPE,
                        PR_EMAIL_ADDRESS,
                        PR_DEPARTMENT_NAME
                }
        };

        EnterCriticalSection(&m_cs);

        if (FAILED(hRes = m_pPropData->GetProps((LPSPropTagArray)&sptSavableProps,
                                                fMapiUnicode,

                                                &cValues,

                                                &pProps)))
        {
                goto LCleanup;
        }

        assert(m_pDB);
        if (m_fCreateFlag)
        {
                // Insert a new entry.
                if (FAILED(hRes = m_pDB->Insert(cValues, pProps, &recID)))
                        goto LCleanup;
        }
        else
        {
                recID = m_EID.ID.Num;
                // Get old props. error ignored.
                m_pDB->Query(m_EID.ID.Num, &pRec);

                // Update an existing entry.
                if (FAILED(hRes = m_pDB->Update(m_EID.ID.Num, cValues, pProps)))
                        goto LCleanup;
        }

        if (FAILED(m_pDB->Save()))
        {
                // Save failed, consider it access denied.
                hRes = MAPI_E_NO_ACCESS;
                // Undo
                if (m_fCreateFlag)
                {
                        // Delete previous inserted record.
                        m_pDB->Delete(recID); // Error ignored.
                }
                else
                {
                        // Restore the updated record.
                        if (pRec)
                                m_pDB->Update(m_EID.ID.Num, pRec->cValues(), pRec->lpProps());
                }
                goto LCleanup;
        }

        if (FAILED(hRes = m_pDB->Query(recID, &pRec)))
                goto LCleanup;

        // Notify parents for the mail user's changes.
        if (pRec)
                HrNotifyParents(pRec, m_fCreateFlag ? TABLE_ROW_ADDED : TABLE_ROW_MODIFIED);

LCleanup:
        LeaveCriticalSection(&m_cs);


        if (pProps)
                vpfnFreeBuff(pProps);

        if (pRec)
                pRec->Release();

        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    GetProps()
//
//    Purpose
//          Delegates to base prop interface.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise
//

STDMETHODIMP  CMailUser::GetProps(LPSPropTagArray       lpPropTagArray,
                                  ULONG                 ulFlags,
                                  ULONG                *lpcValues,
                                  LPSPropValue         *lppPropArray)

{
        CheckParameters_IMAPIProp_GetProps(this,
                                           lpPropTagArray,
                                           ulFlags,
                                           lpcValues,
                                           lppPropArray);

        return m_pPropData->GetProps(lpPropTagArray,
                                     ulFlags,
                                     lpcValues,
                                     lppPropArray);
}

///////////////////////////////////////////////////////////////////////////////
//
//    SetProps()
//
//    Purpose
//          Delegates to base prop interface.
//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CMailUser::SetProps(ULONG                 cValues,
                                 LPSPropValue          lpPropArray,
                                 LPSPropProblemArray  *lppProblems)
{
        CheckParameters_IMAPIProp_SetProps(this,
                                           cValues,
                                           lpPropArray,
                                           lppProblems);

        return m_pPropData->SetProps(cValues,
                                     lpPropArray,
                                     lppProblems);
}

///////////////////////////////////////////////////////////////////////////////
//
//    GetPropList()
//
//    Purpose
//          Delegates to base prop interface.
//    Return Value
//          S_OK on success, a MAPI error code otherwise
//


STDMETHODIMP  CMailUser::GetPropList(ULONG             ulFlags,
                                     LPSPropTagArray  *lppPropTagArray)
{
        CheckParameters_IMAPIProp_GetPropList(this,
                                              ulFlags,
                                              lppPropTagArray);

        return m_pPropData->GetPropList(ulFlags, lppPropTagArray);
}

///////////////////////////////////////////////////////////////////////////////
//
//    DeleteProps()
//
//    Purpose
//          Delegates to base prop interface.
//    Return Value
//          S_OK on success, a MAPI error code otherwise
//

STDMETHODIMP  CMailUser::DeleteProps(LPSPropTagArray       lpPropTagArray,
                                     LPSPropProblemArray  *lppProblems)
{
        CheckParameters_IMAPIProp_DeleteProps(this, lpPropTagArray, lppProblems);

        return m_pPropData->DeleteProps(lpPropTagArray, lppProblems);
}

///////////////////////////////////////////////////////////////////////////////
//
//    CopyTo()
//
//    Purpose
//          Delegates to base prop interface.
//    Return Value
//          S_OK on success, a MAPI error code otherwise
//

STDMETHODIMP  CMailUser::CopyTo(ULONG                 ciidExclude,
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
//    Purpose
//          Delegates to base prop interface.
//    Return Value
//          S_OK on success, a MAPI error code otherwise
//


STDMETHODIMP  CMailUser::CopyProps(LPSPropTagArray       lpIncludeProps,
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
//    Purpose
//          Delegates to base prop interface.
//    Return Value
//          S_OK on success, a MAPI error code otherwise
//

STDMETHODIMP  CMailUser::GetNamesFromIDs(LPSPropTagArray  *lppPropTags,
                                         LPGUID            lpPropSetGuid,
                                         ULONG             ulFlags,
                                         ULONG            *lpcPropNames,
                                         LPMAPINAMEID    **lpppPropNames)
{

        CheckParameters_IMAPIProp_GetNamesFromIDs(this,
                                                  lppPropTags,
                                                  lpPropSetGuid,
                                                  ulFlags,
                                                  lpcPropNames,
                                                  lpppPropNames);

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
//    Purpose
//          Delegates to base prop interface.
//    Return Value
//          S_OK on success, a MAPI error code otherwise
//

STDMETHODIMP  CMailUser::GetIDsFromNames(ULONG             cPropNames,
                                         LPMAPINAMEID     *lppPropNames,
                                         ULONG             ulFlags,
                                         LPSPropTagArray  *lppPropTags)
{
        CheckParameters_IMAPIProp_GetIDsFromNames(this,
                                                  cPropNames,
                                                  lppPropNames,
                                                  ulFlags,
                                                  lppPropTags);

        return m_pPropData->GetIDsFromNames(cPropNames,
                                            lppPropNames,
                                            ulFlags,
                                            lppPropTags);
}

///////////////////////////////////////////////////////////////////////////////
//
//    NotifyParents()
//
//    Parameters

//          pRec        [in] pointer to CRecord obj, that caches props.
//          ulEvent     [in] table event type
//
//    Purpose
//          Notify this container's parents that the object has changed and
//          send them a list of properties they can put in their contents
//          tables for this child.
//
//          Use a table event because then an SRow can be passed for the
//          child's row in contents table to its parent. This is slightly
//          backward: the child isn't a table, the table is in the parent.
//          If an extended notification is used, the row would have to be flattened
//          into a contiguous blob. If a table notification is passed,

//          the row can be passed and let MAPI copy the row data
//          in Notify()
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP  CMailUser::HrNotifyParents(PCRecord       pRec,
                                         ULONG          ulTableEvent)
{
        HRESULT         hRes = S_OK;
        SizedNOTIFKEY(sizeof(SABP_EID), key);
        PSABP_EID       pEID = (PSABP_EID)key.ab;
        NOTIFICATION    Notif;
        ULONG           ulFlags;
        LPSPropValue    pProps  = pRec->lpProps();
        LPSPropValue    pCntTblProps = NULL;
        SRow            srChanged = {0};


        if (FAILED(hRes = m_pLogon->HrGetCntsTblRowProps(pRec, &pCntTblProps)))
                goto LCleanup;

        srChanged.cValues = NUM_CTBL_COLS;
        srChanged.lpProps = pCntTblProps;

        ZeroMemory(&Notif, sizeof(Notif));
        Notif.ulEventType               = fnevTableModified;
        Notif.info.tab.ulTableEvent     = ulTableEvent;
        Notif.info.tab.propIndex        = pProps[CT_INST_KEY];
        Notif.info.tab.row              = srChanged;

        key.cb = sizeof(SABP_EID);
        memcpy((LPVOID)key.ab, &m_EID, key.cb);
        pEID->ID.Type = DONT_CARE;
        pEID->ID.Num  = DIR_ID_NUM;

        ulFlags = 0;

        assert(m_pSupObj);
        hRes = m_pSupObj->Notify((LPNOTIFKEY)&key, 1, &Notif, &ulFlags);

        if (FAILED(hRes))
                goto LCleanup;

LCleanup:

        return hRes;
}

/* End of the file ABUSER.CPP */