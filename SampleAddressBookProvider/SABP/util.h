/****************************************************************************
        UTIL.H

        Copyright (c) 2007 Microsoft Corporation

        Declarations of some useful functions.
*****************************************************************************/

#ifndef _UTIL_H_
#define _UTIL_H_

#include "stdafx.h"
#include "abp.h"

#define ISMV_TAG(ulTag)     !!(MV_FLAG & PROP_TYPE((ulTag)))


STDMETHODIMP HrCopyProp(LPSPropValue pspDest, LPSPropValue pspSrc, LPSPropValue pParent);

STDMETHODIMP HrPropNotFound(SPropValue & spv, ULONG ulTag);

#endif /* _UTIL_H_ */

// End of file for UTIL.H
