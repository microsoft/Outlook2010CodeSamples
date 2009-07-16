typedef BOOL (STDAPICALLTYPE FGETCOMPONENTPATH)
	(LPSTR szComponent,
	LPSTR szQualifier,
	LPSTR szDllPath,
	DWORD cchBufferSize,
	BOOL fInstall);
typedef FGETCOMPONENTPATH FAR * LPFGETCOMPONENTPATH;

typedef struct
{
	LPTSTR lpszSection;
	LPTSTR lpszKey;
	ULONG ulKey;
	LPTSTR lpszValue;
} SERVICESINIREC;

static SERVICESINIREC aWrapPSTServicesIni[] =
{
	{_T("Services"),	_T("WRAPPST"), 0L, _T("Wrapped PST")},

	{_T("WRAPPST"),	_T("PR_DISPLAY_NAME"),			0L, _T("Wrapped PST")},
	{_T("WRAPPST"),	_T("PR_SERVICE_DLL_NAME"),		0L, _T("WRPPST.DLL")},
	{_T("WRAPPST"),	_T("PR_SERVICE_ENTRY_NAME"),	0L, _T("ServiceEntry")},
	{_T("WRAPPST"),	_T("PR_RESOURCE_FLAGS"),		0L, _T("SERVICE_NO_PRIMARY_IDENTITY")},
	{_T("WRAPPST"),	_T("Providers"),				0L,	_T("MS_WRAPPST_P")},
	{_T("WRAPPST"),	_T("PR_SERVICE_SUPPORT_FILES"),	0L, _T("WRPPST.DLL")},
	{_T("WRAPPST"),	_T("PR_SERVICE_DELETE_FILES"),	0L,	_T("Wrppst.DLL")},

	{_T("MS_WRAPPST_P"),	_T("PR_RESOURCE_TYPE"),			0L, _T("MAPI_STORE_PROVIDER")},
	{_T("MS_WRAPPST_P"),	_T("PR_PROVIDER_DLL_NAME"),		0L, _T("WRPPST.DLL")},
	{_T("MS_WRAPPST_P"),	_T("PR_RESOURCE_FLAGS"),		0L, _T("STATUS_DEFAULT_STORE")},
	{_T("MS_WRAPPST_P"),	_T("PR_DISPLAY_NAME"),			0L, _T("Wrapped PST")},
	{_T("MS_WRAPPST_P"),	_T("PR_PROVIDER_DISPLAY"),		0L, _T("Wrapped PST Provider")},
//	{_T("MS_WRAPPST_P"),	NULL,							0x67020003, _T("00000010")}, // uncomment to use Unicode PST

	{NULL, NULL, 0L, NULL}
};

static SERVICESINIREC aREMOVE_WrapPSTServicesIni[] =
{
	{_T("Services"),      _T("WRAPPST"), 0L, NULL},

	{_T("WRAPPST"),       NULL,          0L, NULL},

	{_T("MS_WRAPPST_P"),  NULL,          0L, NULL},

	{NULL,                NULL,          0L, NULL}
};


STDMETHODIMP MergeWithMAPISVC();
STDMETHODIMP RemoveFromMAPISVC();
void GetMAPISVCPath(LPSTR szMAPIDir, ULONG ulMAPIDir);
STDMETHODIMP HrSetProfileParameters(SERVICESINIREC *lpServicesIni);

// macro to get allocated string lengths
#define CCH(string) (sizeof((string))/sizeof(TCHAR))

// result allocated with new, clean up with delete[]
void GetOutlookVersionString(LPTSTR* lppszPath, LPTSTR* lppszVer, LPTSTR* lppszLang);

// result allocated with new, clean up with delete[]
LPTSTR GetPSTPath();