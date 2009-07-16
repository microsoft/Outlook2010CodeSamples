/****************************************************************************
        ABDB.CPP

        Copyright (c) 2007 Microsoft Corporation

        Implementation of CRecord and CABDatabase.
*****************************************************************************/


#include "abdb.h"
#include "util.h"

///////////////////////////////////////////////////////////////////////////////
//
//    FCopyOneProp()
//
//    Parameters

//      pspvSrc        [in]  pointer to the source SPropValue
//      pspvDest       [out] pointer the destination SPropValue
//                           if it is NULL, new mem block will be allocated.
//
//    Purpose
//      Copy a SPropValues structure to another.
//      Since the value of this strucure sometimes are pointer, that is, when
//      the prop type is PT_STRING or PT_BINARY,  copy not
//      only the structure, but also the value of the pointer.
//      NOTE: The SPropValue that was copied by this method
//            can be only freed by FreeOneProp().
//
//    Return Value
//      TRUE for success, FALSE for error.
//
BOOL WINAPI FCopyOneProp(LPSPropValue pspvSrc, LPSPropValue pspvDest)
{
        if (pspvSrc == NULL)
                return FALSE;
        if (pspvSrc != NULL && IsBadWritePtr((LPVOID)pspvSrc, sizeof SPropValue))
                return FALSE;

        BOOL            fRes = TRUE;
        ULONG           ulTag;
        ULONG           ulLen;
        LPVOID          pbSrc = NULL;
        LPVOID          pbDest = NULL;
        BOOL            fNeedAlloc = (pspvDest == NULL);

        if (fNeedAlloc)
                // Allocate memory for the dest SPropValue
                pspvDest = (LPSPropValue)malloc(sizeof SPropValue);

        if (pspvDest == NULL)
        {
                fRes = FALSE;
                goto LCleanup;
        }

        ZeroMemory(pspvDest, sizeof SPropValue);

        ulTag = pspvSrc->ulPropTag;
        pspvDest->ulPropTag = ulTag;

        switch(PROP_TYPE(ulTag))
        {
        case PT_BINARY:
                ulLen  = pspvSrc->Value.bin.cb;
                pbSrc  = pspvSrc->Value.bin.lpb;
                if (ulLen == 0)
                {
                        pspvDest->Value.bin.cb = 0;
                        pspvDest->Value.bin.lpb = NULL;
                }
                else
                {
                        pbDest = malloc(ulLen);
                        if (pbDest == NULL)
                        {
                                fRes = FALSE;
                                goto LCleanup;
                        }
                        memcpy((LPVOID)pbDest, (LPVOID)pbSrc, ulLen);
                        pspvDest->Value.bin.cb = ulLen;
                        pspvDest->Value.bin.lpb = static_cast<LPBYTE>(pbDest);
                }
                break;

        case PT_CLSID:
                ulLen = sizeof GUID;
                pbSrc  = (LPBYTE) pspvSrc->Value.lpguid;
                if (ulLen == 0 || pbSrc == NULL)
                {
                        pspvDest->Value.bin.cb = 0;
                        pspvDest->Value.bin.lpb = NULL;
                }
                else
                {
                        pbDest = malloc(ulLen);
                        memcpy((LPVOID)pbDest, (LPVOID)pbSrc, ulLen);
                        pspvDest->Value.bin.cb = ulLen;
                        pspvDest->Value.bin.lpb = static_cast<LPBYTE>(pbDest);
                }
                break;

        case PT_UNICODE:
                ulLen = lstrlenW(pspvSrc->Value.lpszW) + 1;
                pbSrc  = pspvSrc->Value.lpszW;
                pbDest = malloc(ulLen * sizeof WCHAR);
                memcpy((LPVOID)pbDest, (LPVOID)pbSrc, ulLen * sizeof WCHAR);
                pspvDest->Value.lpszW = static_cast<LPWSTR>(pbDest);
                break;

        case PT_STRING8:
                ulLen = lstrlenA(pspvSrc->Value.lpszA) + 1;
                pbSrc  = pspvSrc->Value.lpszA;
                pbDest = malloc(ulLen * sizeof CHAR);
                memcpy((LPVOID)pbDest, (LPVOID)pbSrc, ulLen * sizeof CHAR);
                pspvDest->Value.lpszA = static_cast<LPSTR>(pbDest);
                break;

        default:
                if (ISMV_TAG(ulTag))
                {
                        fRes = FALSE;
                        goto LCleanup;
                }

                pspvDest->Value = pspvSrc->Value;
                break;
        }

LCleanup:

        if (!fRes)
        {
                if (pbDest)
                        free(pbDest);
                if (fNeedAlloc && pspvDest)
                {
                        free(pspvDest);
                        pspvDest = NULL;
                }
        }

        return fRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    FreeOneProp()
//
//    Parameters

//      pspv            [out] pointer to SPropValue to free.
//
//    Purpose
//      Free one SPropValue structure.
//      Since the value of this strucure sometimes are pointer, that is, when
//      the prop type is PT_STRING or PT_BINARY etc, free not
//      only the structure, but also the value of the pointer.
//      NOTE: this function can be only applied to the props whose value was
//            allocated by malloc() or realloc() when it's a pointer
//
//    Return Value
//      none.
//
void FreeOneProp(LPSPropValue pspv)
{
        if (pspv == NULL)
                return;
        if (IsBadWritePtr(pspv, sizeof SPropValue))
                return;

        ULONG ulTag = pspv->ulPropTag;

        switch(PROP_TYPE(ulTag))
        {
        case PT_BINARY:
                // Fall through.

        case PT_CLSID:
                if (pspv->Value.bin.lpb)
                        free(pspv ->Value.bin.lpb);
                break;

        case PT_UNICODE:
                if (pspv->Value.lpszW)
                        free(pspv->Value.lpszW);
                break;

        case PT_STRING8:
                if (pspv->Value.lpszA)
                        free(pspv->Value.lpszA);
                break;

        default:
                if (ISMV_TAG(ulTag))
                        return;
                break;
        }
}

///////////////////////////////////////////////////////////////////////////////
//
//    FreeProps()
//
//    Parameters

//      pspv            [out] pointer to SPropValue to free.
//      cValues         [in]  count of properties.
//
//    Purpose
//      Free SPropValue structure.
//      Since the value of this strucure sometimes are pointer, that is, when
//      the prop type is PT_STRING or PT_BINARY etc, free not
//      only the structure, but also the value of the pointer.
//      NOTE: this function can be only apply the props whose value was
//            allocated by malloc() or realloc() when it's a pointer
//
//    Return Value
//      none.
//
void FreeProps(LPSPropValue pspv, ULONG cValues)
{
        if (pspv == NULL || cValues == 0)
                return;

        for(ULONG i = 0; i < cValues; i++)
        {
                FreeOneProp((LPSPropValue)&pspv[i]);
        }

        // Free LPSPropValue
        free(pspv);
}

///////////////////////////////////////////////////////////////////////////////
//
//    CRecord()
//
//    Purpose
//      Class constructor.
//

//    Return Value
//      none.
//
CRecord::CRecord()
{
        InitMembers();
}

///////////////////////////////////////////////////////////////////////////////
//
//    CRecord()
//
//    Parameters

//      pRec    [in] pointer to record to copy from.
//
//    Purpose
//      Class constructor.
//

//    Return Value
//      none.
//
CRecord::CRecord(PCRecord pRec)
{
        InitMembers();
        if(pRec)
        {
                if (!this->FSetProps(pRec->lpProps(), pRec->cValues()))
                        return;

                m_recID = pRec->RecID();
        }
}

///////////////////////////////////////////////////////////////////////////////
//
//    Init()
//
//    Parameters

//      none.
//
//    Purpose
//      Initialize members.
//

//    Return Value
//      none.
//
void CRecord::InitMembers()
{
        m_cRef = 1;
        m_recID = 0;
        m_cValues = 0;
        m_lpProps = NULL;
}

///////////////////////////////////////////////////////////////////////////////
//
//    ~CRecord()
//
//    Parameters

//       none.
//
//    Purpose
//       Class destructor.
//

//    Return Value
//       none.
//
CRecord::~CRecord()
{
        // Free m_lpProps.
        FreeProps(m_lpProps, m_cValues);
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
STDMETHODIMP CRecord::QueryInterface(REFIID riid, LPVOID *ppvObj)
{
        if (IsBadWritePtr(ppvObj, sizeof LPUNKNOWN))
                return ERROR_INVALID_PARAMETER;

        *ppvObj = NULL;

        if (riid == IID_IUnknown)
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
STDMETHODIMP_ (ULONG) CRecord::AddRef()
{
        return InterlockedIncrement(&m_cRef);
}

///////////////////////////////////////////////////////////////////////////////
//
//    Release()
//
//    Parameters
//      Refer to OLE Documentation on this method
//
//    Purpose
//      Decreases the usage count of this object. If the

//      count is down to zero, object deletes itself.
//
//    Return Value
//      Usage of this object
//
STDMETHODIMP_ (ULONG) CRecord::Release()
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
//    RecID()
//
//    Parameters
//          none.
//
//    Purpose
//          Get ID of the current record.
//
//    Return Value
//          Record ID.
//
RECORD_ID_TYPE CRecord::RecID() const
{
        return m_recID;
}

///////////////////////////////////////////////////////////////////////////////
//
//    cValues()
//
//    Parameters
//      none.
//
//    Purpose
//      Get count of lpProps in the current record.
//
//    Return Value
//      Count of lpProps in the current record.
//
ULONG CRecord::cValues() const
{
        return m_cValues;
}

///////////////////////////////////////////////////////////////////////////////
//
//    lpProps()
//
//    Parameters
//      none.
//
//    Purpose
//      Get pointer to props in the current record.
//
//    Return Value
//      Pointer of lpProps in the current record.
//
LPSPropValue CRecord::lpProps() const
{
        return m_lpProps;
}

///////////////////////////////////////////////////////////////////////////////
//
//    SetRecID()
//
//    Parameters
//      recID   [in] the new record ID.
//
//    Purpose
//      Set ID of the current record.
//
//    Return Value
//      void.
//
void CRecord::SetRecID(RECORD_ID_TYPE recID)
{
        // assert recID should be larger than 0
        // since 0 is used to initilize a new record.
        assert(recID > 0);

        m_recID = recID;
}


///////////////////////////////////////////////////////////////////////////////
//
//    FSetProps()
//
//    Parameters
//      lpProps [in] pointer to SPropValue.
//      cValues [in] count of properties.
//
//    Purpose
//      Set lpProps and cValues in CRecord.
//      To avoid cValues is inconsistant to lpProps's true count,

//      the lpProps and cValues can be changed only by this function.
//
//    Return Value
//      TRUE for success, FALSE for error.
//
BOOL CRecord::FSetProps(LPSPropValue lpProps, ULONG cValues)
{
        if (lpProps == NULL || cValues == 0)
                return FALSE;

        BOOL            fRes = FALSE;
        LPSPropValue    pNewProps = NULL;

        pNewProps = (LPSPropValue)malloc(cValues * sizeof SPropValue);

        if (!pNewProps)
                goto LCleanup;

        // Copy the passed in props to provider's member.
        for(ULONG i = 0; i < cValues; i++)
        {
                if (!FCopyOneProp(&lpProps[i], &pNewProps[i]))
                {
                        fRes = FALSE;
                        goto LCleanup;
                }
        }

        // free the old props after the aboves succeed.
        FreeProps(m_lpProps, m_cValues);

        // set to new props
        m_cValues = cValues;
        m_lpProps = pNewProps;

        fRes = TRUE;

LCleanup:
        if (!fRes)
        {
                if (pNewProps)
                        free(pNewProps);
        }

        return fRes;
}
/* End of the file ABDB.CPP */