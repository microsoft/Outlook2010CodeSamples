// stdafx.cpp : source file that includes just the standard includes
// Win32App.pch will be the pre-compiled header
// stdafx.obj will contain the pre-compiled type information

#include "stdafx.h"

// TODO: reference any additional headers you need in STDAFX.H
// and not in this file
#define USES_IID_IMAPIPropData
#define USES_IID_IXPProvider
#define USES_IID_IXPLogon
#define USES_IID_IMessage
#define USES_IID_IMAPIStatus
#define USES_IID_IMAPIProp

#include <initguid.h>
#include <Mapiguid.h>

DEFINE_GUID(CLSID_MailMessage,
			0x00020D0B,
			0x0000, 0x0000, 0xC0, 0x00, 0x0, 0x00, 0x0, 0x00, 0x00, 0x46);
