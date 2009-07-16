#pragma once

#define INITGUID
#define USES_IID_IMAPIPropData
#define USES_IID_IXPProvider
#define USES_IID_IXPLogon
#define USES_IID_IMessage
#define USES_IID_IMAPIStatus
#define USES_IID_IMAPIProp

#include <mapix.h>
#include <mapispi.h>
#include <mapiutil.h>
#include <stdio.h>
#include <tchar.h>
#include <strsafe.h>
#include <imessage.h>
#include "PropTags.h"
#include "CXPProvider.h"
#include "CXPLogon.h"
#include "CXPStatus.h"
#include "Output.h"
#include "Utils.h"
#include "resource.h"

#define MRXP_ADDRTYPE		_T("MRXP")
#define MRXP_ADDRTYPE_A		"MRXP"
#define MRXP_DLLNAME		_T("MRXP.DLL")
#define MRXP_FULLDLLNAME	_T("MRXP32.DLL")
#define MRXP_FULLDLLNAME_A	"MRXP32.DLL"
#define MRXP_RESOURCE_METHODS	STATUS_FLUSH_QUEUES | STATUS_SETTINGS_DIALOG | STATUS_VALIDATE_STATE
#define MRXP_ERROR_STATUS_STRING _T("<Unknown Error>")
#define MRXP_ERROR_STATUS_STRING_A "<Unknown Error>"
#define MRXP_ERROR_STATUS_STRING_W L"<Unknown Error>"
#define MSG_EXT				_T(".MSG")
#define DEFAULT_FILENAME	_T("TEMP.MSG")

// http://msdn2.microsoft.com/en-us/library/bb820947.aspx
#define STORE_ITEMPROC			((ULONG) 0x00200000)
#define ITEMPROC_FORCE			((ULONG) 0x00000800)
#define NON_EMS_XP_SAVE			((ULONG) 0x00001000)

DEFINE_GUID(CLSID_MailMessage,
			0x00020D0B,
			0x0000, 0x0000, 0xC0, 0x00, 0x0, 0x00, 0x0, 0x00, 0x00, 0x46);

BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved);

STDINITMETHODIMP XPProviderInit(HINSTANCE			hInstance,
								LPMALLOC			lpMalloc,
								LPALLOCATEBUFFER	lpAllocateBuffer,
								LPALLOCATEMORE 		lpAllocateMore,
								LPFREEBUFFER 		lpFreeBuffer,
								ULONG				ulFlags,
								ULONG				ulMAPIVer,
								ULONG FAR *			lpulProviderVer,
								LPXPPROVIDER FAR *	lppXPProvider);

HRESULT STDAPICALLTYPE ServiceEntry (HINSTANCE		hInstance,
									 LPMALLOC		lpMalloc,
									 LPMAPISUP		lpMAPISup,
									 ULONG			ulUIParam,
									 ULONG			ulFlags,
									 ULONG			ulContext,
									 ULONG			cValues,
									 LPSPropValue	lpProps,
									 LPPROVIDERADMIN lpProviderAdmin,
									 LPMAPIERROR FAR *lppMapiError);

STDMETHODIMP PreprocessMessage(LPVOID lpvSession,
							   LPMESSAGE lpMessage,
							   LPADRBOOK lpAdrBook,
							   LPMAPIFOLDER lpFolder,
							   LPALLOCATEBUFFER AllocateBuffer,
							   LPALLOCATEMORE AllocateMore,
							   LPFREEBUFFER FreeBuffer,
							   ULONG FAR * lpcOutbound,
							   LPMESSAGE FAR * FAR * lpppMessage,
							   LPADRLIST FAR * lppRecipList);

STDMETHODIMP RemovePreprocessInfo(LPMESSAGE lpMessage);

STDMETHODIMP MergeWithMAPISVC();
STDMETHODIMP RemoveFromMAPISVC();
HRESULT WINAPI DoConfigDialog (HINSTANCE		hInstance,
							   ULONG			ulUIParam,
							   ULONG*			pulProfProps,
							   LPSPropValue*	ppPropArray,
							   LPSPropTagArray	pTags,
							   LPMAPISUP		pSupObj,
							   ULONG			ulFlags,
							   LPALLOCATEBUFFER	pAllocateBuffer,
							   LPALLOCATEMORE	pAllocateMore,
							   LPFREEBUFFER		pFreeBuffer);
HRESULT	ValidateConfigProps(ULONG ulProps, LPSPropValue pProps);
HRESULT OpenProviderProfileSection(LPPROVIDERADMIN lpAdmin, LPPROFSECT* lppProfSect);