/****************************************************************************
        ABP.H

        Copyright (c) 2007 Microsoft Corporation

        Declarations of classes, data structures, function prototypes,
        defines, and so on for this Address Book Provider sample.
*****************************************************************************/

#ifndef _ABP_H_
#define _ABP_H_

#include "stdafx.h"
#include "abdb.h"

#define SZ_PSS_SAMPLE_AB        _T("Sample Address Book")
#define SZ_DLL_NAME             _T("SABP.DLL")
#define SZ_AB_DEF_EXT           _T("*")
#define NULL_STRING             _T("")
#define SZ_OFN_FILTER           _T("All Files  (*.*)")
#define SZ_NEW_CONTACT          _T("New Contact")
#define SZ_AVAILABLE            _T("Available")
#define SZ_UNAVAILABLE          _T("Unavailable")

// GUID to identify this provider.
// {D1D704BA-8FD9-44b8-A8E7-669FFA35A033}
DEFINE_GUID(SABP_UID,
            0xd1d704ba, 0x8fd9, 0x44b8, 0xa8, 0xe7, 0x66, 0x9f, 0xfa, 0x35, 0xa0, 0x33);


#define ROOT_ID_NUM             ((ULONG) 0x00000000)    // Root container's ID num
#define DIR_ID_NUM              ((ULONG) 0x00000001)    // Dir container's ID num

#define FOREIGN_ENTRY           0x1000
#define HAS_TEMPLATEID          0x2000
#define ONEOFF_ENTRY            0x4000
#define ONEOFF_USER_ID          0xfffffffe
#define DONT_CARE               0xffffffff


#define BASE_PROVIDER_ID        0x6600  // From 6600 to 67FF is for Provider-defined
                                        //  internal non-transmittable property
                                        //  Defined for this provider.

#define PR_SMP_AB_PATH          PROP_TAG(PT_TSTRING, (BASE_PROVIDER_ID + 0x0008))

#define SUPPORTED_EVENTS        (fnevObjectModified | \
                                 fnevObjectCopied   | \
                                 fnevObjectDeleted)

// C declarations
extern "C"
{
        STDINITMETHODIMP ABProviderInit(HINSTANCE               hInst,
                                        LPMALLOC                pMalloc,
                                        LPALLOCATEBUFFER        pAllocBuff,
                                        LPALLOCATEMORE          pAllocMore,
                                        LPFREEBUFFER            pFreeBuff,
                                        ULONG                   ulFlags,
                                        ULONG                   ulMAPIVers,
                                        ULONG                   *pulABVers,
                                        LPABPROVIDER            *ppABProvider);

        HRESULT STDAPICALLTYPE ServiceEntry(HINSTANCE           hInstance,
                                            LPMALLOC            pMallocObj,
                                            LPMAPISUP           pMAPISupObj,
                                            ULONG               ulUIParam,
                                            ULONG               ulFlags,
                                            ULONG               ulContext,
                                            ULONG               ulCfgPropCount,
                                            LPSPropValue        pCfgProps,
                                            LPPROVIDERADMIN     pAdminProvObj,
                                            LPMAPIERROR         *ppMAPIError);

}

typedef struct
{
        ULONG   Num;    // Corresponding to DB record number.
        ULONG   Type;   // PR_OBJECT_TYPE | FOREIGN_ENTRY, for example
} SABP_ID, *PSABP_ID;

typedef struct _SABP_EID
{
        BYTE    bFlags[4];      // Bitmask of flags that provide information describing the object.

        MAPIUID muid;           // MAPIUID to indicate the provider's UID.
        ULONG   version;        // put structure on 8-byte boundary.
        SABP_ID ID;             // Corresponding to DB record num and type.
} SABP_EID, *PSABP_EID;


//////////////////////////////////////////////////////////////////////////////
//
// Forward classes definition.
//
class CABProvider;
class CABLogon;
class CABContainer;
class CMailUser;
class CMAPITable;
class CConfig;
class CTblAdviseSink;

//////////////////////////////////////////////////////////////////////////////
//
// typedef.
//
typedef CABProvider     *LPCABProvider;
typedef CABLogon        *LPCABLogon;
typedef CABContainer    *LPCABContainer;
typedef CMailUser       *LPCMailUser;
typedef CMAPITable      *LPCMAPITable;
typedef CConfig	        *LPCConfig;
typedef CTblAdviseSink  *LPCTblAdviseSink;

// MAPI allocation functions
extern LPALLOCATEBUFFER vpfnAllocBuff;

extern LPALLOCATEMORE   vpfnAllocMore;

extern LPFREEBUFFER     vpfnFreeBuff;



class CABProvider : public IABProvider
{
public:
        MAPI_IUNKNOWN_METHODS(IMPL);
        MAPI_IABPROVIDER_METHODS(IMPL);

        CABProvider(HINSTANCE hInst);
        ~CABProvider();

private :
        volatile LONG           m_cRef;

        HINSTANCE               m_hInst;
        CRITICAL_SECTION        m_cs;
};



class CABLogon : public IABLogon
{
public:
        MAPI_IUNKNOWN_METHODS(IMPL);
        MAPI_IABLOGON_METHODS(IMPL);

        CABLogon(HINSTANCE      hInstance,

                 LPMAPISUP      pSupObj,
                 MAPIUID        &UID,
                 PCABDatabase   pDB);
        ~CABLogon();

        STDMETHODIMP    HrSetStatusRow(BOOL bAvail, ULONG ulFlags);

        ULONG WINAPI    CheckEID(ULONG cbEntryID, LPENTRYID pEntryID);

        STDMETHODIMP    HrNewEID(ULONG          ulFlags,
                                 ULONG          ulID,
                                 ULONG          ulEntryType,
                                 LPVOID         pParent,
                                 PSABP_EID      *ppNewEID);

        STDMETHODIMP    HrNotify(ULONG          ulEvent,
                                 PSABP_EID      pParentEID,
                                 PSABP_EID      pObjectEID,
                                 LPSPropTagArray pPropTags);


        STDMETHODIMP    HrGetCntsTblRowProps(PCRecord     pRec,
                                             LPSPropValue *ppProps);

private :

        STDMETHODIMP    HrOpenRoot(ULONG        *pulObjType,
                                   LPUNKNOWN    *ppUnk);


        STDMETHODIMP    HrOpenContainer(PSABP_EID       pEID,
                                        ULONG           *pulObjType,
                                        LPUNKNOWN       *ppUnk);

private:
        volatile LONG           m_cRef;         // Reference count.
        HINSTANCE               m_hInst;        // HINSTANCE.
        MAPIUID                 m_UID;          // MAPIUID corresponding to SABP_UID
        CRITICAL_SECTION        m_cs;           // CRITICAL_SECTION
        LPMAPISUP               m_pSupObj;      // Pointer to IMAPISupport object.
        CABContainer            *m_pRootCont;   // Pointer to the root container.
        CABContainer            *m_pDirCont;    // Pointer to the dir container in root.
        PCABDatabase            m_pDB;          // Pointer to the database for the provider.
        list<ULONG>             m_ullstAdviseCnx; // List<ULONG> to save advise connection.
};


class CTblAdviseSink : public IMAPIAdviseSink
{
public:
        CTblAdviseSink(LPCABContainer m_pCont);
        ~CTblAdviseSink();


        MAPI_IUNKNOWN_METHODS(IMPL);


        STDMETHODIMP_ (ULONG)   OnNotify(ULONG cNotif, LPNOTIFICATION pNotifs);

private:
        volatile LONG           m_cRef;         // Reference count.
        LPCABContainer          m_pCont;        // Pointer to the container to advise.
        CRITICAL_SECTION        m_cs;           // CRITICAL_SECTION.
};

class CABContainer : public IABContainer
{
        friend class CTblAdviseSink;
public:
        MAPI_IUNKNOWN_METHODS(IMPL);
        MAPI_IMAPIPROP_METHODS(IMPL);
        MAPI_IMAPICONTAINER_METHODS(IMPL);
        MAPI_IABCONTAINER_METHODS(IMPL);


        CABContainer(HINSTANCE          hInst,
                     PSABP_EID          pEID,
                     LPCABLogon         pLogon,
                     LPMAPISUP          pSupObj,
                     PCABDatabase       pDB);


        virtual ~CABContainer();


        STDMETHODIMP HrInit();

private:
        STDMETHODIMP    HrSubscribe();


        STDMETHODIMP    HrSetDefaultContProps(PSABP_EID  pEID);


        STDMETHODIMP    HrCreateHierTable(LPTABLEDATA *ppTable);


        STDMETHODIMP    HrCreateCntsTable(LPTABLEDATA *ppTable);


        STDMETHODIMP    HrOpenTemplates(LPMAPITABLE     *lppTable);

        STDMETHODIMP    HrLoadHierTblProps(LPSPropValue *ppProps);


        STDMETHODIMP    HrGetView(LPTABLEDATA           pTblData,
                                  LPSSortOrderSet       lpSSortOrderSet,
                                  CALLERRELEASE FAR     *lpfCallerRelease,
                                  ULONG                 ulCallerData,
                                  LPMAPITABLE FAR       *lppMAPITable);


        STDMETHODIMP    HrCopyAdrEntry(LPSPropTagArray  pTags,
                                       LPSRow           pRow,
                                       LPADRENTRY       pAdrEntry);

        STDMETHODIMP    HrCreateMailUser(PSABP_EID      pEID);

private:
        volatile LONG           m_cRef;	        // Reference count.
        CRITICAL_SECTION        m_cs;           // CRITICAL_SECTION.
        SABP_EID                m_EID;          // EntryID.
        HINSTANCE               m_hInst;
        LPMAPISUP               m_pSupObj;      // Pointer to IMAPISupport object.
        LPCABLogon              m_pLogon;       // Pointer to Logon
        PCABDatabase            m_pDB;          // Pointer to the database for the provider.
        LPTABLEDATA             m_pCntsTblData; // Pointer to the ITABLEDATA object of the provider's content table.
        LPTABLEDATA             m_pHierTblData; // Pointer to the ITABLEDATA object of the provider's hierachy table.
        LPPROPDATA              m_pPropData;    // Pointer to the IPROPDATA object of the provider's props.
        ULONG                   m_ulCnx;        // cnx no. to get notifications on
        LPCTblAdviseSink        m_pAdviseSink;  // Pointer to an IAdviseSink object.
        BOOL                    m_fInited;      // Flag to identify whether CABContainer is initialized.
};

class CMailUser : public IMailUser
{
public:
        MAPI_IUNKNOWN_METHODS(IMPL);
        MAPI_IMAPIPROP_METHODS(IMPL);


        CMailUser(HINSTANCE     hInst,
                  PSABP_EID     pEID,
                  LPCABLogon    pLogon,
                  LPMAPISUP     pSupObj,
                  PCABDatabase  pDB);


        ~CMailUser();


        STDMETHODIMP    HrInit();

private:
        STDMETHODIMP    HrSetDefaultProps(PCRecord pRec);


        STDMETHODIMP    HrNotifyParents(PCRecord pRec, ULONG ulTableEvent);

private:
        volatile LONG           m_cRef;   // Reference count.

        HINSTANCE               m_hInst;
        CRITICAL_SECTION        m_cs;     // CRITICAL_SECTION.
        SABP_EID                m_EID;    // EntryID.
        LPCABLogon              m_pLogon; // Pointer to Logon
        LPMAPISUP               m_pSupObj;// Pointer to IMAPISupport object.
        PCABDatabase            m_pDB;    // Pointer to the database for the provider.
        LPPROPDATA              m_pPropData;// Pointer to the IPROPDATA object of the provider's props.
        BOOL                    m_fCreateFlag;// Flag to indicate whether the CMailUser is initiated for creating new entry.
        BOOL                    m_fInited;// Flg to indicate whether CMailUser is initiated.
};

class CMAPITable : public IMAPITable
{
public:
        MAPI_IUNKNOWN_METHODS(IMPL);
        MAPI_IMAPITABLE_METHODS(IMPL);

        CMAPITable(LPMAPITABLE pTbl,ULONG ulParent);
        ~CMAPITable();

private:
        volatile LONG           m_cRef; // Reference count.
        LPMAPITABLE             m_pTbl; // The IMAPITABLE to wrap.
        CRITICAL_SECTION        m_cs;   // CRITICAL_SECTION
        ULONG                   m_ParentID; //

};

class CConfig : public IUnknown
{
public:
        MAPI_IUNKNOWN_METHODS(IMPL);

        CConfig(HWND            hWnd,
                MAPIUID         &UID,
                LPPROFSECT      pProfSecn);
        ~CConfig();


        STDMETHODIMP    HrInit();

        STDMETHODIMP    HrDoConfig();

        STDMETHODIMP    HrUpdateProfile();

private:
        STDMETHODIMP    HrOpenProfileSection();

        STDMETHODIMP    HrBrowseFile(HWND hWnd,LPTSTR *pszFile, LPVOID pParent);

        STDMETHODIMP    HrSetDefaultProps();

private:
        volatile LONG   m_cRef;         // Reference count.
        HWND            m_hWnd;         // Handle of parent window.
        LPPROFSECT      m_pProfSecn;    // Pointer to profile section.
        LPSPropValue    m_pProps;       // Pointer to property array.
        MAPIUID         m_UID;          // The provider UID.
        BOOL            m_fInited;
};

LPTSTR GetNextToken(LPTSTR *pszInput);

STDMETHODIMP  DoANR(LPTSTR        szTarget,
                    LPMAPITABLE   pSearchTbl,
                    LPSRow        psrResolved,
                    ULONG         ulNameCol);

// Logon props.
typedef enum

{
        LOGON_DB_PATH,
        LOGON_DISP_NAME,
        LOGON_COMMENT,
        LOGON_NUM_PROPS
} LOGON_PROP_IDX;

const SizedSPropTagArray(LOGON_NUM_PROPS, vsptLogonTags) =

{
        LOGON_NUM_PROPS,
        {
                PR_SMP_AB_PATH,
                PR_DISPLAY_NAME,
                PR_COMMENT
        }
};


// Contents table columns.
enum

{
        CT_EID,
        CT_INST_KEY,
        CT_REC_KEY,
        CT_DISP_TYPE,
        CT_OBJTYPE,
        CT_NAME,
        CT_ADDRTYPE,
        CT_ADDR,
        NUM_CTBL_COLS
};
const SizedSPropTagArray(NUM_CTBL_COLS, vsptCntTbl) =
{
        NUM_CTBL_COLS,
        {
                PR_ENTRYID,
                PR_INSTANCE_KEY,
                PR_RECORD_KEY,
                PR_DISPLAY_TYPE,
                PR_OBJECT_TYPE,
                PR_DISPLAY_NAME,
                PR_ADDRTYPE,
                PR_EMAIL_ADDRESS
        }
};

// Hierarchy table columns
enum

{
        HT_INST_KEY,
        HT_DEPTH,
        HT_ACCESS,
        HT_EID,
        HT_OBJ_TYPE,
        HT_DISP_TYPE,
        HT_NAME,
        HT_CONT_FLAGS,
        HT_STATUS,
        HT_COMMENT,
        HT_PROV_UID,
        NUM_HTBL_PROPS
};

const SizedSPropTagArray(NUM_HTBL_PROPS, vsptHierTbl) =
{
        NUM_HTBL_PROPS,
        {
                PR_INSTANCE_KEY,
                PR_DEPTH,
                PR_ACCESS,
                PR_ENTRYID,
                PR_OBJECT_TYPE,
                PR_DISPLAY_TYPE,
                PR_DISPLAY_NAME,
                PR_CONTAINER_FLAGS,
                PR_STATUS,
                PR_COMMENT,
                PR_AB_PROVIDER_ID
        }
};

// Some handy overloads
BOOL operator ==(const SABP_ID & lhs, const SABP_ID & rhs);
BOOL operator ==(const SABP_EID & lhs, const SABP_EID & rhs);
BOOL operator !=(const MAPIUID & lhs, const MAPIUID & rhs);


#endif /* _ABP_H_ */
