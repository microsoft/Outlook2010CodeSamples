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

static SERVICESINIREC aMRXPServicesIni[] =
{
	{_T("Services"),	_T("MRXP"),						0L, _T("MRXP Sample Transport")},

	{_T("MRXP"),		_T("PR_DISPLAY_NAME"),			0L, _T("MRXP Sample Transport")},
	{_T("MRXP"),		_T("PR_SERVICE_DLL_NAME"),		0L, _T("MRXP.DLL")},
	{_T("MRXP"),		_T("PR_SERVICE_ENTRY_NAME"),	0L, _T("ServiceEntry")},
	{_T("MRXP"),		_T("Providers"),				0L,	_T("MRXP_PRV")},
	{_T("MRXP"),		_T("PR_SERVICE_SUPPORT_FILES"),	0L, _T("MRXP.DLL")},
	{_T("MRXP"),		_T("PR_SERVICE_DELETE_FILES"),	0L,	_T("MRXP.DLL")},

	{_T("MRXP_PRV"),	_T("PR_RESOURCE_TYPE"),			0L, _T("MAPI_TRANSPORT_PROVIDER")},
	{_T("MRXP_PRV"),	_T("PR_PROVIDER_DLL_NAME"),		0L, _T("MRXP.DLL")},
	{_T("MRXP_PRV"),	_T("PR_RESOURCE_FLAGS"),		0L, _T("STATUS_PRIMARY_IDENTITY")},
	{_T("MRXP_PRV"),	_T("PR_DISPLAY_NAME"),			0L, _T("MRXP Sample Transport")},
	{_T("MRXP_PRV"),	_T("PR_PROVIDER_DISPLAY"),		0L, _T("MRXP Sample Transport")},

	{NULL, NULL, 0L, NULL}
};

static SERVICESINIREC aREMOVE_MRXPServicesIni[] =
{
	{_T("Services"),	_T("MRXP"), 0L,	NULL},

	{_T("MRXP"),		NULL,	0L, NULL},

	{_T("MRXP_PRV"),	NULL,	0L, NULL},

	{NULL,	NULL, 0L, NULL}
};

void GetMAPISVCPath(LPSTR szMAPIDir, ULONG ulMAPIDir);
STDMETHODIMP HrSetProfileParameters(SERVICESINIREC *lpServicesIni);
HRESULT CopySBinary(LPSBinary psbDest, const LPSBinary psbSrc, LPVOID lpParent,
					LPALLOCATEBUFFER MyAllocateBuffer, LPALLOCATEMORE MyAllocateMore);