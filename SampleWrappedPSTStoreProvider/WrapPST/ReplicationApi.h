


// ===============================================================================================
//	Constants
// ===============================================================================================

#define  pbNSTGlobalProfileSectionGuid "\x85\xED\x14\x23\x9D\xF7\x42\x66\x8B\xF2\xFB\xD4\xA5\x21\x29\x41"

#define	DNH_OK	0x00010000
#define	DNT_OK	0x00010000
#define	HSF_LOCAL	0x00000008
#define	HSF_COPYDESTRUCTIVE	0x00000010
#define	HSF_OK	0x00010000
#define	SS_ACTIVE	0
#define	SS_SUSPENDED	1
#define	SYNC_UPLOAD_HIERARCHY	0x00000001
#define	SYNC_DOWNLOAD_HIERARCHY	0x00000002
#define	SYNC_UPLOAD_CONTENTS	0x00000040
#define	SYNC_DOWNLOAD_CONTENTS	0x00000080
#define	SYNC_OUTGOING_MAIL	0x00000200
#define	SYNC_BACKGROUND	0x00001000
#define	SYNC_THESE_FOLDERS	0x00020000
#define	SYNC_HEADERS	0x02000000
#define	UPC_OK	0x00010000
#define	UPD_ASSOC	0x00000001
#define	UPD_MOV	0x00000002
#define	UPD_OK	0x00010000
#define	UPD_MOVED	0x00020000
#define	UPD_UPDATE	0x00040000
#define	UPD_COMMIT	0x00080000
#define	UPF_NEW	0x00000001
#define	UPF_MOD_PARENT	0x00000002
#define	UPF_MOD_PROPS	0x00000004
#define	UPF_DEL	0x00000008
#define	UPF_OK	0x00010000
#define	UPH_OK	0x00010000
#define	UPM_ASSOC	0x00000001
#define	UPM_NEW	0x00000002
#define	UPM_MOV	0x00000004
#define	UPM_MOD_PROPS	0x00000008
#define	UPM_HEADER	0x00000010
#define	UPM_OK	0x00010000
#define	UPM_MOVED	0x00020000
#define	UPM_COMMIT	0x00040000
#define	UPM_DELETE	0x00080000
#define	UPM_SAVE	0x00100000
#define	UPR_ASSOC	0x00000001
#define	UPR_READ	0x00000002
#define	UPR_OK	0x00010000
#define	UPR_COMMIT	0x00020000
#define	UPS_UPLOAD_ONLY	0x00000001
#define	UPS_DNLOAD_ONLY	0x00000002
#define	UPS_THESE_FOLDERS	0x00000080
#define	UPS_OK	0x00010000
#define	UPT_PUBLIC	0x00000001
#define	UPT_OK	0x00010000
#define	UPV_ERROR	0x00010000
#define	UPV_DIRTY	0x00020000
#define	UPV_COMMIT	0x00040000


// ===============================================================================================
//	Interface Identifiers
// ===============================================================================================

// {4FDEEFF0-0319-11CF-B4CF-00AA0DBBB6E6}
DEFINE_GUID (IID_IPSTX, 0x4FDEEFF0, 0x0319, 0x11CF, 0xB4, 0xCF, 0x00, 0xAA, 0x0D, 0xBB, 0xB6, 0xE6);

// {2067A790-2A45-11D1-EB86-00A0C90DCA6D}
DEFINE_GUID (IID_IPSTX2, 0x2067A790, 0x2A45, 0x11D1, 0xEB, 0x86, 0x00, 0xA0, 0xC9, 0x0D, 0xCA, 0x6D);

// {55f15320-111b-11d2-a999-006008b05aa7}
DEFINE_GUID (IID_IPSTX3, 0x55f15320, 0x111b, 0x11d2, 0xa9, 0x99, 0x00, 0x60, 0x08, 0xb0, 0x5a, 0xa7);

// {aa2e2092-ac08-11d2-a2f9-0060b0ec3d4f}
DEFINE_GUID (IID_IPSTX4, 0xaa2e2092, 0xac08, 0x11d2, 0xa2, 0xf9, 0x00, 0x60, 0xb0, 0xec, 0x3d, 0x4f);

// {55f15322-111b-11d2-a999-006008b05aa7}
DEFINE_GUID (IID_IPSTX5, 0x55f15322, 0x111b, 0x11d2, 0xa9, 0x99, 0x00, 0x60, 0x08, 0xb0, 0x5a, 0xa7);

// {55f15323-111b-11d2-a999-006008b05aa7}
DEFINE_GUID (IID_IPSTX6, 0x55f15323, 0x111b, 0x11d2, 0xa9, 0x99, 0x00, 0x60, 0x08, 0xb0, 0x5a, 0xa7);

// {d2d85db4-840f-49b8-9982-07d2405ec6b7}
DEFINE_GUID (IID_IOSTX, 0xd2d85db4, 0x840f, 0x49b8, 0x99, 0x82, 0x07, 0xd2, 0x40, 0x5e, 0xc6, 0xb7);


// ===============================================================================================
//	Interface Definitions
// ===============================================================================================


#define MAPI_IPSTX_METHODS(IPURE)                                   \
	MAPIMETHOD(GetLastError)											\
		(THIS_	HRESULT						hResult,					\
				ULONG						ulFlags,					\
				LPMAPIERROR FAR *			lppMAPIError) IPURE;		\
    MAPIMETHOD(GetSyncObject)                                           \
        (THIS_  IOSTX **ppostx) IPURE;									 \
    MAPIMETHOD(PlaceHolder1)													\
        () IPURE;														\
    MAPIMETHOD(PlaceHolder2)													\
        () IPURE;														\
    MAPIMETHOD(PlaceHolder3)													\
        () IPURE;														\
    MAPIMETHOD(PlaceHolder4)													\
        () IPURE;														\
	MAPIMETHOD(EmulateSpooler)                                           \
        (THIS_  BOOL fEmulate) IPURE;									 \
    MAPIMETHOD(PlaceHolder5)													\
        () IPURE;														\
    MAPIMETHOD(PlaceHolder6)													\
        () IPURE;														\
    MAPIMETHOD(PlaceHolder7)													\
        () IPURE;														\
    MAPIMETHOD(PlaceHolder8)													\
        () IPURE;														\
    MAPIMETHOD(PlaceHolder9)													\
        () IPURE;														\





  #define MAPI_IOSTX_METHODS(IPURE)                                   \
    MAPIMETHOD(PlaceHolder1)													\
        () IPURE;														\
	MAPIMETHOD(InitSync)											\
		(ULONG ulFlags) IPURE;		\
	MAPIMETHOD(SyncBeg)											\
			(UINT uiSync,							\
			LPVOID *ppv) IPURE;		\
	MAPIMETHOD(SyncEnd)													\
        () IPURE;														\
	MAPIMETHOD(SyncHdrBeg)											\
			(ULONG cbeid,							\
			LPENTRYID lpeid,							\
			LPVOID *ppv) IPURE;		\
	MAPIMETHOD(SyncHdrEnd)											\
			( LPMAPIPROGRESS pprog) IPURE;		\
	MAPIMETHOD(SetSyncResult)											\
			( HRESULT hrSync) IPURE;		\
    MAPIMETHOD(PlaceHolder2)													\
        () IPURE;														\


DECLARE_MAPI_INTERFACE_(IOSTX, IUnknown)
{
    BEGIN_INTERFACE
    MAPI_IUNKNOWN_METHODS(PURE)
    MAPI_IOSTX_METHODS(PURE)
};

DECLARE_MAPI_INTERFACE_PTR(IOSTX, LPOSTX);

DECLARE_MAPI_INTERFACE_(IPSTX, IUnknown)
{
    BEGIN_INTERFACE
    MAPI_IUNKNOWN_METHODS(PURE)
    MAPI_IPSTX_METHODS(PURE)
};

DECLARE_MAPI_INTERFACE_PTR(IPSTX, LPPSTX);




// ===============================================================================================
//	Structure Definitions
// ===============================================================================================

DECLARE_MAPI_INTERFACE_PTR(IExchangeImportHierarchyChanges,	PXIHC);
DECLARE_MAPI_INTERFACE_PTR(IExchangeImportContentsChanges,	PXICC);

struct DNTBLE
{
	UINT	cEntNew;
	UINT	cEntMod;
	UINT	cEntRead;
	UINT	cEntDel;
};

struct LTID
{
	GUID	guid;
	BYTE	globcnt[6];
	WORD	wLevel;
};

#pragma pack(1)

struct FEID
{
	BYTE	abFlags[4];
	MAPIUID	muid;
	WORD	placeholder;
	LTID	ltid;
};

struct MEID
{
	BYTE	abFlags[4];
	MAPIUID	muid;
	WORD	placeholder;
	LTID	ltidFld;
	LTID	ltidMsg;
};

struct SKEY
{
	GUID	guid;
	BYTE	globcnt[6];
};

#pragma pack()

typedef enum {
	LR_SYNC_IDLE = 0,
	LR_SYNC,
	LR_SYNC_UPLOAD_HIERARCHY,
	LR_SYNC_UPLOAD_FOLDER,
	LR_SYNC_CONTENTS,
	LR_SYNC_UPLOAD_TABLE,
	LR_SYNC_UPLOAD_MESSAGE,
	LR_SYNC_UPLOAD_MESSAGE_READ = 8,
	LR_SYNC_UPLOAD_MESSAGE_DEL,
	LR_SYNC_DOWNLOAD_HIERARCHY,
	LR_SYNC_DOWNLOAD_TABLE,
	LR_SYNC_DOWNLOAD_HEADER = 17
} SYNCSTATE;



struct DNHIER
{
	ULONG		ulFlags;
	LPSTREAM	pstmReserved;
	PXIHC		pxihc;
	UINT		cEntNew;
	UINT		cEntMod;
	UINT		cEntDel;
};


struct DNTBL
{
	ULONG		ulFlags;
	LPSTREAM	pstmReserved1;
	LPSTREAM	pstmReserved2;
	LPSTREAM	pstmReserved3;
	LPSTREAM	pstmReserved4;
	PXICC		pxicc;
	PXIHC		pxihc;
	LPSTR		pszName;
	FILETIME	ftLastMod;
	ULONG		ulRights;
	FEID		feid;
	UINT		uintReserved;
	DNTBLE		rgte[2];
	LPSRestriction	psrReserved;
	BOOL		boReserved;
	void*		pReserved1;
	void*		pReserved2;
};

struct UPMSG
{
	ULONG		ulFlags;
	LPMESSAGE	pmsg;
	MEID		meid;
	SBinary		binReserved1;
	SBinary		binReserved2;
	FEID		feid;
	SBinary		binChg;
	SBinary		binPcl;
	SKEY		skeySrc;
};


struct HDRSYNC
{
	UPMSG		*pupmsg;
	FEID		feidPar;
	LPSTREAM	pstmReserved;
	ULONG		ulFlags;
	LPMESSAGE	pmsgFull;
};

struct SYNCCONT
{
	ULONG		ulFlags;
	UINT		iEnt;
	UINT		cEnt;
	LPVOID 		pvReserved;
	LPSPropTagArray	ptagaReserved;
	LPSSortOrderSet	psosReserved;
};

typedef struct UPMOV *PUPMOV;


struct UPMOV
{
	ULONG		ulFlags;
	LPVOID		pReserved;
	LPSTREAM	pstmReserved;
	LPSTR		pszName;
	FEID		feid;
	LPMAPIFOLDER	pfld;
	PXICC		pxicc;
	DWORD		dwReserved;
	PUPMOV		pupmovNext;
	UINT		cEntMov;
};



struct UPDELE
{
	ULONG	ulFlags;
	SKEY	skey;
	DWORD   dwReserved;
	SBinary	binChg;
	SBinary	binPcl;
	SKEY	skeyDst;
	PUPMOV	pupmov;
};

typedef UPDELE *PUPDELE;

struct UPDEL
{
	PUPDELE	pupde;
	UINT	cEnt;
};


struct UPFLD
{
	ULONG		ulFlags;
	LPMAPIFOLDER	pfld;
	FEID		feid;
};

struct UPHIER
{
	ULONG ulFlags;
	LPSTREAM pstmReserved;
	UINT iEnt;
	UINT cEnt;
};


typedef struct UPREAD *PUPREADE;

struct UPREAD
{
	PUPREADE	pupre;
	UINT		cEnt;
};

struct UPREADE
{
	ULONG	ulFlags;
	SKEY	skey;
};

struct SYNC
{
	ULONG		ulFlags;
	LPWSTR		pwzPath;
	FEID		Reserved1;
	MEID		Reserved2;
	LPENTRYLIST	pel;
	ULONG const *	pulFolderOptions;
};

struct UPTBLE
{
	UINT	iEntMod;
	UINT	cEntMod;
	UINT	iEntRead;
	UINT	cEntRead;
	UINT	iEntDel;
	UINT	cEntDel;
};


struct UPTBL
{
	ULONG		ulFlags;
	LPSTREAM	pstmReserved;
	LPSTR		pszName;
	FEID		feid;
	UINT		uintReserved;
	UPTBLE		rgte[2];
	UINT		iEnt;
	UINT		cEnt;
	PUPMOV		pupmovHead;
	void* pReserved;
};










