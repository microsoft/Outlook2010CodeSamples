/****************************************************************************
        ABCFG.CPP

        Copyright (c) 2007 Microsoft Corporation

        CConfig implementation.
        CConfig implement the methods for configure logon properties,
        such as database file path.
*****************************************************************************/

#include "abp.h"
#include "Commdlg.h"
#include "strsafe.h"

///////////////////////////////////////////////////////////////////////////////
//
//    CConfig()
//
//    Parameters

//
//    Purpose
//          Class constructor.
//

//    Return Value
//          none.
//
CConfig::CConfig(HWND           hWnd,

                 MAPIUID        &UID,
                 LPPROFSECT     pProfSecn)
: m_cRef(1),
  m_hWnd(hWnd),
  m_pProfSecn(pProfSecn),
  m_pProps(NULL),
  m_fInited(FALSE)
{
        memcpy((LPVOID) &m_UID, (LPVOID) &UID, sizeof UID);

        if (m_pProfSecn)
                m_pProfSecn->AddRef();
}


///////////////////////////////////////////////////////////////////////////////
//
//    ~CConfig()
//
//    Parameters

//
//    Purpose
//          Class destructor.
//

//    Return Value
//          none.
//
CConfig::~CConfig()
{
        if (m_pProps)
                vpfnFreeBuff(m_pProps);

        // Release objects
        if (m_pProfSecn)
                m_pProfSecn->Release();
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrInit()
//
//    Parameters

//
//    Purpose
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CConfig::HrInit()
{
        HRESULT hRes = S_OK;
        ULONG   ulcProps = 0;

        if (m_fInited)
        {
                return hRes;
        }

        assert(m_pProfSecn);

        // Load old props from ProfSecn
        hRes = m_pProfSecn->GetProps((LPSPropTagArray)&vsptLogonTags,

                                     fMapiUnicode,
                                     &ulcProps,
                                     &m_pProps);

        if (!m_pProps
                || m_pProps[LOGON_DB_PATH].ulPropTag != PR_SMP_AB_PATH
                || m_pProps[LOGON_DISP_NAME].ulPropTag != PR_DISPLAY_NAME
                || m_pProps[LOGON_COMMENT].ulPropTag != PR_COMMENT)
        {
                // If props loaded from profile failed, or any of them has an error,
                // that may mean it is the first time for users to configure this provider,
                // Set the props to default.
                if (m_pProps)
                        vpfnFreeBuff(m_pProps);

                hRes = HrSetDefaultProps();
        }

        if (FAILED(hRes))
        {
                if(m_pProps)
                {
                        vpfnFreeBuff(m_pProps);
                        m_pProps = NULL;
                }
        }

        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrSetDefaultProps()
//
//    Parameters
//          none.
//
//    Purpose
//          Set default props for config obj.
//
//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CConfig::HrSetDefaultProps()
{
        HRESULT hRes;

        // Allocate buffer for m_pProps.
        if (FAILED(hRes = vpfnAllocBuff(LOGON_NUM_PROPS * sizeof SPropValue, (LPVOID*)&m_pProps)))
                return hRes;

        m_pProps[LOGON_DB_PATH].ulPropTag = PR_SMP_AB_PATH;
        m_pProps[LOGON_DB_PATH].Value.LPSZ = NULL_STRING;

        m_pProps[LOGON_DISP_NAME].ulPropTag = PR_DISPLAY_NAME;
        m_pProps[LOGON_DISP_NAME].Value.LPSZ = SZ_PSS_SAMPLE_AB;

        m_pProps[LOGON_COMMENT].ulPropTag = PR_COMMENT;
        m_pProps[LOGON_COMMENT].Value.LPSZ = NULL_STRING;

        return S_OK;
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
STDMETHODIMP_ (ULONG) CConfig::AddRef()
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
STDMETHODIMP_ (ULONG) CConfig::Release()
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
//    QueryInterface()
//
//    Parameters

//          See the OLE documentation for details.
//
//    Purpose
//          Returns an interface if supported.
//

//    Return Value
//          NOERROR on success, E_NOINTERFACE if the interface
//          can't be returned.
//
STDMETHODIMP CConfig::QueryInterface(REFIID riid, LPVOID *ppvObj)
{
        if (IsBadWritePtr(ppvObj, sizeof LPUNKNOWN))
                return MAPI_E_INVALID_PARAMETER;

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
//    HrDoConfig()
//
//    Parameters

//          none.
//
//    Purpose
//          Do configuration work, such as show config form and let user to
//          select the database file.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CConfig::HrDoConfig()
{
        HRESULT hRes = S_OK;

        if (FAILED(hRes = HrBrowseFile(m_hWnd,

                                       &(m_pProps[LOGON_DB_PATH].Value.LPSZ),
                                       (LPVOID)m_pProps)))
        {
                return hRes;
        }

        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrUpdateProfile()
//
//    Parameters

//          none.
//
//    Purpose
//          Save configue changes to profile.
//

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//
STDMETHODIMP CConfig::HrUpdateProfile()
{
        HRESULT hRes;

        assert(m_pProfSecn);
        if (FAILED(hRes = m_pProfSecn->SetProps(LOGON_NUM_PROPS,
                                                m_pProps,
                                                NULL)))
        {
                return hRes;
        }

        return m_pProfSecn->SaveChanges(0);
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrBrowseFile()
//
//    Parameters
//          hWnd    [in]     handle of main application window
//          pszFile [in/out] in - file name displayed in GetOpenFileName
//                                dialog
//                           out- newly allocated buffer containing

//                                user selected new file name
//          pParen  [in]     memory the allocation is linked to
//    Purpose
//          Wrapper for GetOpenFileName(). Pass in an existing file name
//          string in *pszFile and non-null in pParent, and on output
//          *pszFile has the user's selection in a newly allocated buffer
//          linked to pParent. If pParent is NULL, the selection is copied
//          into *pszFile, which must be allocated by the caller and of
//          sufficient size.
//
//    Return Value
//          TRUE on success, FALSE on CANCEL or error
//

STDMETHODIMP CConfig::HrBrowseFile(HWND hWnd,
                                   LPTSTR *pszFile,
                                   LPVOID pParent)
{
        HRESULT       hRes   = S_OK;
        LPTSTR        szTemp = NULL;
        OPENFILENAME  ofn    = {0,0,0,0,0,0,0,0,0,0,
                                0,0,0,0,0,0,0,0,0,0};

        if (pParent)
        {
                if (FAILED(hRes = vpfnAllocMore(MAX_PATH * sizeof(TCHAR),
                           pParent,
                           (LPVOID*)&szTemp)))
                        return hRes;

        }

        else
        {
                if (FAILED(hRes = vpfnAllocBuff(MAX_PATH * sizeof(TCHAR),
                           (LPVOID*)&szTemp)))
                        return hRes;
        }

        StringCchCopy(szTemp, MAX_PATH, *pszFile);

        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner   = hWnd;
        ofn.lpstrFilter = SZ_OFN_FILTER;
        ofn.lpstrFile   = szTemp;
        ofn.nMaxFile    = MAX_PATH;
        ofn.lpstrTitle  = SZ_PSS_SAMPLE_AB;
        ofn.Flags       = OFN_HIDEREADONLY | OFN_NOCHANGEDIR | OFN_PATHMUSTEXIST;
        ofn.lpstrDefExt = SZ_AB_DEF_EXT;

        if (GetOpenFileName(&ofn))
                // don't need to free old *pszFile, it's linked to
                // pParent when passed to this function, so is szTemp
                *pszFile = szTemp;

        else
                // if this code is reached, user must have canceled browse,
                // don't free szTemp, it's still linked to parent
                hRes = MAPI_E_USER_CANCEL;

        return hRes;
}

/* End of the file CFG.CPP */