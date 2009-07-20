/****************************************************************************
        TABLE.CPP

        Copyright (c) 2007 Microsoft Corporation

        Wrap MAPI's IMAPITable implementation to allow Restrict on PR_ANR.
*****************************************************************************/
#include "abp.h"

///////////////////////////////////////////////////////////////////////////////
//
//    CMAPITable()
//
//    Parameters
//          pTbl        [in] pointer to table to wrap
//          ulParent    [in] ID of object that owns this table
//    Purpose
//          Class constructor. This class wraps an IMAPITable
//          to intercept Restrict on PR_ANR
//
//    Return Value
//          none.
//
CMAPITable::CMAPITable(LPMAPITABLE pTbl, ULONG ulParent)
: m_cRef(1),
  m_pTbl(pTbl),
  m_ParentID(ulParent)
{
        InitializeCriticalSection(&m_cs);
}

///////////////////////////////////////////////////////////////////////////////
//
//    ~CMAPITable()
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
CMAPITable::~CMAPITable()
{
        m_pTbl->Release();

        DeleteCriticalSection(&m_cs);
}

///////////////////////////////////////////////////////////////////////////////
//
//    GetLastError()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::GetLastError(HRESULT             hResult,
                                      ULONG               ulFlags,
                                      LPMAPIERROR FAR    *lppMAPIError)
{
        return m_pTbl->GetLastError(hResult,
                                    ulFlags,
                                    lppMAPIError);

}

///////////////////////////////////////////////////////////////////////////////
//
//    Advise()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::Advise(ULONG             ulEventMask,
                                LPMAPIADVISESINK  lpAdviseSink,
                                ULONG_PTR FAR     *lpulConnection)
{
        return m_pTbl->Advise(ulEventMask,
                              lpAdviseSink,
                              lpulConnection);
}

///////////////////////////////////////////////////////////////////////////////
//
//    Unadvise()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::Unadvise(ULONG_PTR ulConnection)
{
        return m_pTbl->Unadvise(ulConnection);
}

///////////////////////////////////////////////////////////////////////////////
//
//    GetStatus()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::GetStatus(ULONG FAR    *lpulTableStatus,
                                   ULONG FAR    *lpulTableType)
{
        return m_pTbl->GetStatus(lpulTableStatus,
                                 lpulTableType);
}

///////////////////////////////////////////////////////////////////////////////
//
//    SetColumns()
//
//	Refer to the MSDN documentation for more information.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::SetColumns(LPSPropTagArray lpPropTagArray,
                                    ULONG           ulFlags)
{
        return m_pTbl->SetColumns(lpPropTagArray,
                                  ulFlags);
}

///////////////////////////////////////////////////////////////////////////////
//
//    QueryColumns()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::QueryColumns(ULONG                ulFlags,
                                      LPSPropTagArray FAR *lpPropTagArray)
{
        return m_pTbl->QueryColumns(ulFlags,
                                    lpPropTagArray);
}

///////////////////////////////////////////////////////////////////////////////
//
//    GetRowCount()
//
//	Refer to the MSDN documentation for more information.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::GetRowCount(ULONG      ulFlags,
                                     ULONG FAR *lpulCount)
{
        return m_pTbl->GetRowCount(ulFlags,
                                   lpulCount);
}

///////////////////////////////////////////////////////////////////////////////
//
//    SeekRow()
//
//	Refer to the MSDN documentation for more information.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::SeekRow(BOOKMARK   bkOrigin,
                                 LONG       lRowCount,
                                 LONG FAR  *lplRowsSought)
{
        return m_pTbl->SeekRow(bkOrigin,
                               lRowCount,
                               lplRowsSought);
}

///////////////////////////////////////////////////////////////////////////////
//
//    SeekRowApprox()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::SeekRowApprox(ULONG  ulNumerator,
                                       ULONG  ulDenominator)
{
        return m_pTbl->SeekRowApprox(ulNumerator,
                                     ulDenominator);
}

///////////////////////////////////////////////////////////////////////////////
//
//    QueryPosition()
//
//	Refer to the MSDN documentation for more information.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::QueryPosition(ULONG FAR *lpulRow,
                                       ULONG FAR *lpulNumerator,
                                       ULONG FAR *lpulDenominator)
{
        return m_pTbl->QueryPosition(lpulRow,
                                     lpulNumerator,
                                     lpulDenominator);
}

///////////////////////////////////////////////////////////////////////////////
//
//    FindRow()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::FindRow(LPSRestriction  lpRestriction,
                                 BOOKMARK        bkOrigin,
                                 ULONG           ulFlags)
{
        return m_pTbl->FindRow(lpRestriction,
                               bkOrigin,
                               ulFlags);
}

///////////////////////////////////////////////////////////////////////////////
//
//    GetNextToken()
//
//    Parameters
//          pszInput    [in-out] input: buffer to tokenize,
//                               output:buffer with tokenized prefix
//                               removed suitable to pass to GetNextToken
//
//    Purpose
//          Destructively obtains the next non-white space substring
//          in the input buffer. Equivalent to strtok().
//
//    Return Value
//          The next substring if found, NULL if at end of string
//
LPTSTR GetNextToken(LPTSTR *pszInput)
{
        TCHAR *szPtr = *pszInput,
                *szTmp;

        if (!szPtr)
                return NULL;

        // skip white space
        while (szPtr && *szPtr && (unsigned)*szPtr <= 0x20)
                szPtr++;

        szTmp = szPtr;

        // skip blackspace
        while (szPtr && *szPtr && (unsigned)*szPtr > 0x20)
        {
                szPtr++;
        }

        // end of input ?
        if (*szPtr == '\0')
                *pszInput = NULL;
        else
        {
                *szPtr = (TCHAR) 0;
                *pszInput = szPtr + 1;
        }

        return CharUpper(szTmp);
}

///////////////////////////////////////////////////////////////////////////////
//
//    Restrict()
//
//	Refer to the MSDN documentation for more information.
//
//    Purpose
//          This gets called if MAPI_E_AMBIGUOUS is returned for entries in
//          ResolveName, and for the finder. The same initial search
//          algorithm is used as for ResolveName: the first token is found in the target
//          and a content restriction is done for all substrings. This yields many
//          hits. The added step of parsing to filter out the spurious
//          hits is not done because the user can select them from the UI. IAddrBook calls this
//          immediately after ResolveNames if a name is marked as ambiguous, but if
//          a single row is returned here as the ambiguous match, it treats it as if
//          it were resolved.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::Restrict(LPSRestriction  lpRestriction,
                                  ULONG           ulFlags)
{
        HRESULT         hRes = S_OK;
        LPSRestriction  pRes = lpRestriction;
        SRestriction    SRes;
        SPropValue      spTemp = {0};

        CheckParameters_IMAPITable_Restrict(this,lpRestriction,ulFlags);

        if (lpRestriction && (lpRestriction->rt == RES_PROPERTY))
                if (PROP_ID(PR_ANR) == PROP_ID(lpRestriction->res.resProperty.ulPropTag))
                {
                        TCHAR szBuf[MAX_PATH];
                        LPTSTR szTemp = szBuf;

                        if (IsBadReadPtr((LPBYTE) lpRestriction->res.resProperty.lpProp->Value.LPSZ,

                                Cbtszsize(lpRestriction->res.resProperty.lpProp->Value.LPSZ)))
                                return MAPI_E_INVALID_PARAMETER;

                        lstrcpyn(szTemp,
                                lpRestriction->res.resProperty.lpProp->Value.LPSZ,
                                MAX_PATH);
                        szTemp[MAX_PATH - 1] = '\0';

                        spTemp.ulPropTag = PR_DISPLAY_NAME;
                        spTemp.Value.LPSZ = GetNextToken(&szTemp);

                        SRes.rt = RES_CONTENT;
                        SRes.res.resContent.ulPropTag    = PR_DISPLAY_NAME;
                        SRes.res.resContent.lpProp       = &spTemp;
                        SRes.res.resContent.ulFuzzyLevel = FL_SUBSTRING | FL_IGNORECASE;
                        pRes = &SRes;
                }

                hRes = m_pTbl->Restrict(pRes, ulFlags);
                return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    CreateBookmark()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::CreateBookmark(BOOKMARK FAR *lpbkPosition)
{
        return m_pTbl->CreateBookmark(lpbkPosition);
}

///////////////////////////////////////////////////////////////////////////////
//
//    FreeBookmark()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//

STDMETHODIMP CMAPITable::FreeBookmark(BOOKMARK  bkPosition)
{
        return m_pTbl->FreeBookmark(bkPosition);
}

///////////////////////////////////////////////////////////////////////////////
//
//    SortTable()
//
//	Refer to the MSDN documentation for more information.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::SortTable(LPSSortOrderSet  lpSortCriteria,
                                   ULONG            ulFlags)
{
        return m_pTbl->SortTable(lpSortCriteria, ulFlags);
}

///////////////////////////////////////////////////////////////////////////////
//
//    QuerySortOrder()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::QuerySortOrder(LPSSortOrderSet FAR *lppSortCriteria)
{
    return m_pTbl->QuerySortOrder(lppSortCriteria);
}

///////////////////////////////////////////////////////////////////////////////
//
//    QueryRows()
//
//    Refer to the MSDN documentation for more information.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::QueryRows(LONG             lRowCount,
                                   ULONG            ulFlags,
                                   LPSRowSet FAR   *lppRows)
{
        return m_pTbl->QueryRows(lRowCount, ulFlags, lppRows);
}

///////////////////////////////////////////////////////////////////////////////
//
//    Abort()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::Abort()
{
        return m_pTbl->Abort();
}

///////////////////////////////////////////////////////////////////////////////
//
//    ExpandRow()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::ExpandRow(ULONG            cbInstanceKey,
                                   LPBYTE           pbInstanceKey,
                                   ULONG            ulRowCount,
                                   ULONG            ulFlags,
                                   LPSRowSet FAR   *lppRows,
                                   ULONG FAR       *lpulMoreRows)
{
        return m_pTbl->ExpandRow(cbInstanceKey,
                                 pbInstanceKey,
                                 ulRowCount,
                                 ulFlags,
                                 lppRows,
                                 lpulMoreRows);
}

///////////////////////////////////////////////////////////////////////////////
//
//    CollapseRow()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//

STDMETHODIMP CMAPITable::CollapseRow(ULONG       cbInstanceKey,
                                     LPBYTE      pbInstanceKey,
                                     ULONG       ulFlags,
                                     ULONG FAR  *lpulRowCount)
{
        return m_pTbl->CollapseRow(cbInstanceKey,
                                   pbInstanceKey,
                                   ulFlags,
                                   lpulRowCount);
}

///////////////////////////////////////////////////////////////////////////////
//
//    WaitForCompletion()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::WaitForCompletion(ULONG       ulFlags,
                                           ULONG       ulTimeout,
                                           ULONG FAR  *lpulTableStatus)
{
        return m_pTbl->WaitForCompletion(ulFlags,
                                         ulTimeout,
                                         lpulTableStatus);
}

///////////////////////////////////////////////////////////////////////////////
//
//    GetCollapseState()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::GetCollapseState(ULONG        ulFlags,
                                          ULONG        cbInstanceKey,
                                          LPBYTE       lpbInstanceKey,
                                          ULONG FAR   *lpcbCollapseState,
                                          LPBYTE FAR  *lppbCollapseState)
{
        return m_pTbl->GetCollapseState(ulFlags,
                                        cbInstanceKey,
                                        lpbInstanceKey,
                                        lpcbCollapseState,
                                        lppbCollapseState);
}

///////////////////////////////////////////////////////////////////////////////
//
//    SetCollapseState()
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise.
//
STDMETHODIMP CMAPITable::SetCollapseState(ULONG          ulFlags,
                                          ULONG          cbCollapseState,
                                          LPBYTE         pbCollapseState,
                                          BOOKMARK FAR  *lpbkLocation)
{
        return m_pTbl->SetCollapseState(ulFlags,
                                        cbCollapseState,
                                        pbCollapseState,
                                        lpbkLocation);
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

STDMETHODIMP CMAPITable::QueryInterface(REFIID   riid,
                                        LPVOID  *ppvObj)
{
        if (IsBadWritePtr(ppvObj, sizeof LPUNKNOWN))
                return MAPI_E_INVALID_PARAMETER;

        *ppvObj = NULL;

        if (riid == IID_IUnknown || riid == IID_IMAPITable)
        {
                *ppvObj = static_cast<LPVOID>(this);
                AddRef();
                return NOERROR;
        }

        return E_NOINTERFACE;
}

///////////////////////////////////////////////////////////////////////////////
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
STDMETHODIMP_ (ULONG) CMAPITable::AddRef (void)
{
        return InterlockedIncrement(&m_cRef);
}

///////////////////////////////////////////////////////////////////////////////
//    CMAPITable::Release()
//
//    Parameters
//      Refer to OLE Documentation on this method
//
//    Purpose
//      Decreases the usage count of this object and deletes
//      the object when the count reaches 0.
//
//    Return Value
//      Usage of this object
//
STDMETHODIMP_ (ULONG) CMAPITable::Release(void)
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

// End of file for ITABLE.CPP
