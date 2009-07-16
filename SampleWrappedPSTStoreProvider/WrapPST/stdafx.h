// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#pragma once
#pragma warning(disable:4710)
#include <stdio.h>
#include <WINDOWS.H>
#include <COMMCTRL.H>

#include <MAPIX.H>
#include <MAPIUTIL.H>
#include <MAPIFORM.H>

#include <tchar.h>

#include "resource.h"
#include "utils.h"
#include <strsafe.h>
#include "ReplicationAPI.h"
#include "Provider.h"
#include "support.h"

void DeInitLogging();
void Log(BOOL bPrintThreadTime, LPCTSTR szMsg,...);

void LogREFIID(REFIID riid);

