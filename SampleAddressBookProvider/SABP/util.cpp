/****************************************************************************
        UTIL.CPP

        Copyright (c) 2007 Microsoft Corporation

        Some handy utility functions.
*****************************************************************************/

#include "util.h"

///////////////////////////////////////////////////////////////////////////////
//
//    CopyProp()
//
//    Parameters

//          pspDest   [in/out]  pointer to SPropValue to copy to
//          pspSrc    [in]      pointer to SPropValue to copy from
//          pParent   [in]      parent block to link allocations to
//    Purpose
//          copy an SPropValue from pspSrc to spsDest, including any

//          sub structures. The resulting SPropValue sub structures

//          are linked to pParent so the whole group can be freed by

//          freeing pParent

//    Return Value
//          S_OK on success, a MAPI error code otherwise
//

STDMETHODIMP HrCopyProp(LPSPropValue pspDest, LPSPropValue pspSrc, LPSPropValue pParent)

{
        if (!pspDest || !pspSrc || !pParent)
                return MAPI_E_INVALID_PARAMETER;

        HRESULT         hRes = S_OK;
        ULONG           ulTag = pspSrc->ulPropTag;
        ULONG           ulLen;
        LPBYTE          pbSrc = NULL;
        LPBYTE         *ppbDest = NULL;

        // in case of failure

        ZeroMemory(pspDest, sizeof SPropValue);
        pspDest->ulPropTag = pspSrc->ulPropTag;

        switch(PROP_TYPE(ulTag))
        {
        case PT_BINARY:
                pbSrc  = pspSrc->Value.bin.lpb;
                ulLen  = pspSrc->Value.bin.cb;
                ppbDest = &(pspDest->Value.bin.lpb);
                pspDest->Value.bin.cb = ulLen;
                break;

        case PT_UNICODE:
                ulLen = lstrlenW(pspSrc->Value.lpszW) + 1;
                pbSrc  = (LPBYTE) pspSrc->Value.lpszW;
                ppbDest = (LPBYTE *) &(pspDest->Value.lpszW);
                break;

        case PT_STRING8:
                ulLen = lstrlenA(pspSrc->Value.lpszA) + 1;
                pbSrc  = (LPBYTE) pspSrc->Value.lpszA;
                ppbDest = (LPBYTE *) &(pspDest->Value.lpszA);
                break;

        case PT_CLSID:
                ulLen = sizeof GUID;
                pbSrc  = (LPBYTE) pspSrc->Value.lpguid;
                ppbDest = (LPBYTE *) &(pspDest->Value.lpguid);
                break;

        default:
                if (ISMV_TAG(ulTag))
                        return MAPI_E_NO_SUPPORT;

                ulLen = 0;
                pbSrc = NULL;
                ppbDest = NULL;
                pspDest->Value = pspSrc->Value;
                break;
        }

        if (ulLen && pbSrc && ppbDest)
                if (SUCCEEDED(hRes = vpfnAllocMore(ulLen * sizeof TCHAR,(LPVOID) pParent,(LPVOID *) ppbDest)))
                        memcpy((LPVOID)*ppbDest, (LPVOID)pbSrc, ulLen * sizeof TCHAR);

        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrPropNotFound()
//
//    Parameters

//          pspv      [in/out]  pointer to SPropValue to mark 'not found'
//          ulTag     [in]      property tag of missing property
//    Purpose
//          Marks a property value as NOT FOUND by changing the tag to
//          (ulTag | PT_ERROR) and the Value.err to MAPI_E_NOT_FOUND
//
//    Return Value
//          MAPI_W_ERRORS_RETURNED
//

STDMETHODIMP HrPropNotFound(SPropValue &spv, ULONG ulTag)
{
        spv.ulPropTag = CHANGE_PROP_TYPE(ulTag,PT_ERROR);
        spv.Value.err = MAPI_E_NOT_FOUND;
        return MAPI_W_ERRORS_RETURNED;
}

/* End of the file UTIL.CPP */