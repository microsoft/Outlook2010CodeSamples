/****************************************************************************
        TXTDB.CPP

        Copyright (c) 2007 Microsoft Corporation

        CTxtABDatabase Implementation
*****************************************************************************/

#include "abdb.h"
#include <string.h>

#define LINE_BREAK              _T("\n")
#define ERROR_ENCODING_TEXT     _T("Files are not supported other than those with UNCODE endcoding!\
                                   \nPlease re-configure Sample Address Book Provider.")
#define ERROR_ENCODING_CAPTION  _T("Sample Address Book Provider")

#ifdef UNICODE
#define OPEN_MODE_READONLY      L"r,ccs=UNICODE"
#define OPEN_MODE_WRITE         L"w,ccs=UNICODE"
#else
#define OPEN_MODE_READONLY      "r"
#define OPEN_MODE_WRITE         "w"
#endif

///////////////////////////////////////////////////////////////////////////////
//
//    CTxtABDatabase()
//
//    Parameters

//          filename           [in] database file name
//
//    Purpose
//          Class constructor.
//

//    Return Value
//          none.
//
CTxtABDatabase::CTxtABDatabase(const TCHAR *filename)
{
        if (filename)
                _tcscpy_s(m_filename, MAX_LINE_CHARS, filename);
        InitializeCriticalSection(&m_cs);
}

///////////////////////////////////////////////////////////////////////////////
//
//    ~CTxtABDatabase()
//
//    Parameters

//          none.
//
//    Purpose
//          Class destructor.
//

//    Return Value
//          none.
//
CTxtABDatabase::~CTxtABDatabase()
{
        Close();
        DeleteCriticalSection(&m_cs);

        // Release records stored in record list.
        for(RECLIST::iterator it = m_recList.begin();
                it!= m_recList.end();

                it++)
        {
                (*it)->Release();
        }
}

///////////////////////////////////////////////////////////////////////////////
//
//    IsUnicodeFile()
//
//    Parameters

//          pszFile     [in] pointer to the name of the file to check.
//
//    Purpose
//          Check whether the specified file exists.
//

//    Return Value
//          TRUE for existing, FALSE for non-existing.
//
BOOL FFileExists(TCHAR *pszFile)
{
        if (!pszFile)
                return FALSE;

        errno_t err = 0;
        FILE    *pFile = NULL;

        err = _tfopen_s(&pFile, pszFile, _T("r"));
        if (pFile)
        {
                fclose(pFile);
                return TRUE;
        }
        else
        {
                return FALSE;
        }
}

///////////////////////////////////////////////////////////////////////////////
//
//    IsUnicodeFile()
//
//    Parameters

//          pszFile     [in] pointer to the name of the file to check.
//
//    Purpose
//          Check whether the specified file is encoded in UNICODE.
//

//    Return Value
//          TRUE for UNICODE, FALSE for others.
//
BOOL IsUnicodeFile(TCHAR *pszFile)
{
        if (!pszFile)
                return FALSE;

        BOOL            fRes = FALSE;
        const ULONG     ULNUMS = 2;
        TCHAR           szPre[ULNUMS];
        errno_t         err = 0;
        FILE            *pFile = NULL;

        memset(szPre, 0, (ULNUMS) * sizeof(TCHAR));

        err = _tfopen_s(&pFile, pszFile, _T("r"));
        if (err)
        {
                // File open error.
                goto LCleanup;
        }

        fread(szPre, sizeof(TCHAR), ULNUMS, pFile);


        if (!IsTextUnicode(szPre, ULNUMS, NULL))
        {
                fRes = FALSE;
                goto LCleanup;
        }
        else
        {
                fRes = TRUE;
        }

LCleanup:
        if (pFile)
                fclose(pFile);

        return fRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    Open()
//
//    Parameters

//          none.
//
//    Purpose
//          Open database.
//

//    Return Value
//          S_OK for success, others for failure.
//
STDMETHODIMP CTxtABDatabase::Open()
{
        HRESULT  hRes  = S_OK;
        FILE    *pFile = NULL;

        assert(m_filename);

        // Check whether the file exists.
        if (FFileExists(m_filename))
        {
#ifdef UNICODE
                // Check whether the existed file is UNICODE
                if (!IsUnicodeFile(m_filename))
                {
                        // Not UNICODE file.
                        // Pop up error message
                        MessageBox(0,
                                   ERROR_ENCODING_TEXT,
                                   ERROR_ENCODING_CAPTION,
                                   MB_OK);

                        hRes = MAPI_E_UNCONFIGURED;
                        goto LCleanup;
                }
#endif /* UNICODE */
        }
        else
        {
                // Call Save method to create a new db file.
                Save();
        }

        // Open file for reading records.
        if (FAILED(hRes = HrOpenFile(OPEN_MODE_READONLY, &pFile)))
                goto LCleanup;

        // Load records from the file.
        if (FAILED(hRes = HrLoadRecords(pFile)))
                goto LCleanup;

        m_iter = m_recList.begin();

LCleanup:
        if (pFile)
                if (fclose(pFile))
                        return E_FAIL;

        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrOpenFile()
//
//    Parameters

//          mode        [in] open file mode.
//          ppFile      [out]pointer to pointer to the file opened.
//
//    Purpose
//          Open the specified database file.

//          If it does not exist, create it.
//

//    Return Value
//          S_OK for success, others for failure.
//
STDMETHODIMP CTxtABDatabase::HrOpenFile(TCHAR *mode, FILE **ppFile)
{
        if (!mode || !ppFile)
                return E_INVALIDARG;

        HRESULT hRes = S_OK;
        errno_t err  = 0;

        EnterCriticalSection(&m_cs);

        err = _tfopen_s(ppFile, m_filename, mode);
        if (err)
        {
                // open file failed.
                hRes = E_FAIL;
        }

        LeaveCriticalSection(&m_cs);
        return S_OK;
}

///////////////////////////////////////////////////////////////////////////////
//
//    TrimStr()
//
//    Parameters

//          wcsBuf      [in] chars to trim.
//          nLen        [out] length of the trimmed chars.
//
//    Purpose
//          Trim chars which are smaller than 0x20
//          at the beginning and end of a string.
//

//    Return Value
//          pointer to the trimmed chars.
//
TCHAR * CTxtABDatabase::TrimStr(TCHAR *wcsBuf, int &nLen)
{
        nLen = 0;

        while(wcsBuf != NULL && *wcsBuf != NULL && (unsigned)(*wcsBuf) <= 0x20)
        {
                wcsBuf++;
        }

        nLen = wcsBuf != NULL ? static_cast<int>(_tcslen(wcsBuf)) : 0;
        while(nLen > 0)
        {
                if ((unsigned)(*(wcsBuf + nLen -1)) <= 0x20)
                {
                        nLen--;
                }
                else
                {
                        break;
                }
        }

        return wcsBuf;
}

///////////////////////////////////////////////////////////////////////////////
//
//    FDetectTag()
//
//    Parameters

//          wcsBuf      [in] chars to search in.
//          wcsTag      [in] tag to find.
//
//    Purpose
//          Detect whether the chars contains the specified tag
//

//    Return Value
//          TRUE for found, FALSE for NOT
//
BOOL CTxtABDatabase::FDetectTag(TCHAR *wcsBuf, TCHAR *wcsTag)
{
        if (wcsBuf == NULL || wcsTag == NULL)
                return FALSE;

        while((*wcsBuf) != NULL && (*wcsTag) != NULL)
        {
                if ((*wcsBuf++) != (*wcsTag++))
                        return FALSE;
        }

        return TRUE;
}

///////////////////////////////////////////////////////////////////////////////
//
//    HrLoadRecords()
//
//    Parameters

//          none.
//
//    Purpose
//          Load record from the database file to record list in memory.
//

//    Return Value
//          S_OK for success, others for failure.
//
STDMETHODIMP CTxtABDatabase::HrLoadRecords(FILE *pFile)
{
        if (pFile == NULL)
                return E_INVALIDARG;

        HRESULT         hRes    = S_OK;
        TCHAR           wchLine[MAX_LINE_CHARS + 1];
        list<TSTRING>   lstStrLine;
        TCHAR          *pLine   = NULL;
        int             n_tcslen = 0;

        EnterCriticalSection(&m_cs);

        // Read db file.
        while (_fgetts(wchLine, MAX_LINE_CHARS, pFile))
        {
                pLine = TrimStr(wchLine, n_tcslen);
                if (*pLine == NULL)
                {
                        // Reaching an empty line means a record ends,
                        // so create a new record with
                        // these props in list, and then push it to
                        // the record list.
                        FPushRecord(lstStrLine, m_recList);


                        // Clear the string list for a new record's props.
                        lstStrLine.clear();
                }
                else
                {
                        lstStrLine.push_back(pLine);
                }
        }

        if (lstStrLine.empty())
        {
                goto LCleanup;
        }

        // Push the last CRecord to RECLIST,
        // if the end line of the file is not a empty line.
        FPushRecord(lstStrLine, m_recList);

LCleanup:
        LeaveCriticalSection(&m_cs);
        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    FPushRecord()
//
//    Parameters

//          lstStrProps [in] list of strings containing prop tags and values.
//          recList     [in] record list to push new record in.
//
//    Purpose
//          Create a new CRecord and set its props with props in form of
//          string that are stored in list, and then push the new record
//          to the specified record list.
//

//    Return Value
//          TRUE for success, FALSE for failure.
//
BOOL CTxtABDatabase::FPushRecord(list<TSTRING> &lstStrProps, RECLIST &recList)
{
        if (lstStrProps.empty())
                return FALSE;

        BOOL            fRes = FALSE;
        PCRecord        pRec = NULL;
        LPSPropValue    pProps = NULL;
        RECORD_ID_TYPE  recID = 0;
        size_t          size = 0;

        pRec = new CRecord();
        if (!pRec)
        {
                goto LCleanup;
        }

        recID = NewRecID();
        pRec->SetRecID(recID++);

        size = lstStrProps.size();
        // Allocate pProps.
        pProps = (LPSPropValue)malloc(size * sizeof SPropValue);
        if (!pProps)
        {
                goto LCleanup;
        }

        // Convert the props in form of string to LPSPropValue
        if (!FGetProps(lstStrProps, size, pProps))
                goto LCleanup;

        if (!pRec->FSetProps(pProps, static_cast<ULONG>(size)))
                goto LCleanup;

        // Push the new record to record list.
        recList.push_back(pRec);

        fRes = TRUE;

LCleanup:
        if (pProps && size > 0)
                FreeProps(pProps, size);
        if (!fRes && pRec)
                pRec->Release();

        return fRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    FGetProps()
//
//    Parameters

//          listStrLine [in] list of strings in the form of TAG:VALUE.
//          ppProp      [Out] pointer to pointer to SPropValue structure array.
//
//    Purpose
//          Get props from the string list.
//
//    Return Value
//          TRUE for success, FALSE for failure.
//
BOOL CTxtABDatabase::FGetProps(list<TSTRING>     &listStrLine,
                               ULONG             cPropValues,
                               LPSPropValue      pProps)
{
        if (!cPropValues || !pProps)
                return FALSE;

        BOOL            fRes = FALSE;
        ULONG           index;
        TCHAR          *pszDest = NULL;
        TCHAR          *pszTag = new TCHAR[MAX_LINE_CHARS + 1];
        TCHAR          *pszVal;
        ULONG           ulTag;
        ULONG           i = 0;
        int             n_tcslen = 0;

        memset((LPVOID)pProps, 0, cPropValues * sizeof SPropValue);

        for(list<TSTRING>::iterator iter = listStrLine.begin();
                iter != listStrLine.end() && i < cPropValues;
                iter++)
        {
                // Parse tag value line.
                TCHAR *wStr = const_cast<TCHAR *>((*iter).c_str());
                wStr = TrimStr(wStr, n_tcslen);

                pszDest = _tcsstr(wStr, TAG_VALUE_SEPARATOR);
                if (!pszDest)
                        // This line is of wrong format,

                        // then abandon this record and fall through to the next.
                        goto LCleanup;

                index = pszDest - wStr + 1;

                pszVal  = new TCHAR[MAX_LINE_CHARS + 1];
                if (!pszVal)
                {
                        goto LCleanup;
                }

                _tcsncpy_s(pszTag, MAX_LINE_CHARS, wStr, index - 1);
                _tcsncpy_s(pszVal, MAX_LINE_CHARS, wStr + index, n_tcslen - index);

                // Convert string tag to ulong tag.
                ulTag = _tstol(pszTag);

                pProps[i].ulPropTag     = ulTag;
                pProps[i].Value.LPSZ    = (LPTSTR)pszVal;
                i++;
        }

        fRes = TRUE;

LCleanup:
        if (pszTag)
                delete [] pszTag;

        return fRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    Close()
//
//    Parameters

//          none.
//
//    Purpose
//          Close database
//

//    Return Value
//          S_OK for success, others for failure.
//
STDMETHODIMP CTxtABDatabase::Close()
{
        return S_OK;
}

///////////////////////////////////////////////////////////////////////////////
//
//    RecordCount()
//
//    Parameters

//          none.
//
//    Purpose
//          Get the count of records in the database
//

//    Return Value
//          count of records.
//
ULONG CTxtABDatabase::RecordCount()
{
        return static_cast<ULONG>(m_recList.size());
}

///////////////////////////////////////////////////////////////////////////////
//
//    Read()
//
//    Parameters

//          none.
//
//    Purpose
//          Get the next record for traversing records.
//

//    Return Value
//          Pointer to the next record. NULL when it comes to the end.
//
PCRecord CTxtABDatabase::Read()
{
        PCRecord pRec = NULL;

        if (!m_recList.empty() && m_iter != m_recList.end() && *m_iter != NULL)
        {
                pRec = new CRecord(*m_iter);
                if (!pRec)
                        return NULL;

                pRec->AddRef();

                EnterCriticalSection(&m_cs);
                // Increment iterator
                m_iter++;
                LeaveCriticalSection(&m_cs);
        }

        return pRec;
}

///////////////////////////////////////////////////////////////////////////////
//
//    GoBegin()
//
//    Parameters

//          none.
//
//    Purpose
//          move the iterator to the beginning of the database.
//

//    Return Value
//          none.
//
void CTxtABDatabase::GoBegin()
{
        EnterCriticalSection(&m_cs);
        m_iter = m_recList.begin();
        LeaveCriticalSection(&m_cs);
}

///////////////////////////////////////////////////////////////////////////////
//
//    GoEnd()
//
//    Parameters

//          none.
//
//    Purpose
//          move the iterator to the end of the database.
//

//    Return Value
//          none.
//
void CTxtABDatabase::GoEnd()
{
        EnterCriticalSection(&m_cs);
        m_iter = m_recList.end();
        LeaveCriticalSection(&m_cs);
}

///////////////////////////////////////////////////////////////////////////////
//
//    Query()
//
//    Parameters

//          recID       [in] ID of the record to query.
//          ppRec       [out] the found record.
//
//    Purpose
//          Query a record that matches the specified record ID.
//

//    Return Value
//          S_OK for success, others for failure.
//
STDMETHODIMP CTxtABDatabase::Query(RECORD_ID_TYPE       recID,
                                   PCRecord            *ppRec)
{
        if (!ppRec)
                return E_INVALIDARG;

        PCRecord pRec = NULL;
        *ppRec = NULL;

        RECLIST::iterator it;
        for(it = m_recList.begin(); it!= m_recList.end(); it++)
        {
                if ((*it)->RecID() == recID)
                {
                        pRec = new CRecord(*it);
                        if (!pRec)
                                return ERROR_NOT_ENOUGH_MEMORY;
                        pRec->AddRef();
                        *ppRec = pRec;

                        return S_OK;
                }
        }

        return MAPI_E_NOT_FOUND;
}

///////////////////////////////////////////////////////////////////////////////
//
//    Insert()
//
//    Parameters

//          cValues     [in] count of lpProps
//          lpProps     [out] pointer to SPropValue structure.
//          pRecID      [out] pointer to ID of the new record.
//                            If null, no ID returned.
//
//    Purpose
//          Insert a new record in the database.
//

//    Return Value
//          S_OK for success, others for failure.
//
STDMETHODIMP CTxtABDatabase::Insert(ULONG               cValues,
                                    LPSPropValue        lpProps,
                                    RECORD_ID_TYPE     *pRecID)
{
        if (!cValues || !lpProps || !pRecID)
                return E_INVALIDARG;

        HRESULT hRes = S_OK;

        PCRecord pRec = new CRecord();
        if (!pRec)
                return ERROR_NOT_ENOUGH_MEMORY;

        pRec->SetRecID(NewRecID());
        if (!pRec->FSetProps(lpProps, cValues))
        {
                hRes = E_FAIL;
                goto LCleanup;
        }

        EnterCriticalSection(&m_cs);
        m_recList.push_back(pRec);
        LeaveCriticalSection(&m_cs);
        // Return ID of the new record.
        if (pRecID)
                *pRecID = pRec->RecID();

LCleanup:
        if (FAILED(hRes) && pRec)
                pRec->Release();

        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    NewRecID()
//
//    Parameters

//          none.
//
//    Purpose
//          Generate a new record ID.
//

//    Return Value
//          record ID.
//
RECORD_ID_TYPE CTxtABDatabase::NewRecID()
{
        RECORD_ID_TYPE recID = ulInitRecID;

        if (!m_recList.empty())
        {
                RECLIST::iterator it = m_recList.end();
                recID = (*(--it))->RecID() + 1;
        }

        return recID;
}

///////////////////////////////////////////////////////////////////////////////
//
//    Update()
//
//    Parameters

//          recID       [in] ID of the record to update
//          cValues     [in] count of lpProps
//          lpProps     [in] pointer to SPropValue
//
//    Purpose
//          Update a record in the database.
//

//    Return Value
//          S_OK for success, others for failure.
//
STDMETHODIMP CTxtABDatabase::Update(RECORD_ID_TYPE      recID,
                                    ULONG               cValues,
                                    LPSPropValue        lpProps)
{
        if (!cValues || !lpProps)
                return E_INVALIDARG;

        HRESULT hRes;
        PCRecord pRec = NULL;

        EnterCriticalSection(&m_cs);

        for(RECLIST::iterator it = m_recList.begin();
                it!= m_recList.end();
                it++)
        {
                if ((*it)->RecID() == recID)
                {
                        pRec = *it;
                        break;
                }
        }

        if (!pRec)
        {
                hRes = MAPI_E_NOT_FOUND;
                goto LCleanup;
        }

        if (!pRec->FSetProps(lpProps, cValues))
        {
                hRes = E_FAIL;
                goto LCleanup;
        }

LCleanup:
        LeaveCriticalSection(&m_cs);

        return S_OK;
}

///////////////////////////////////////////////////////////////////////////////
//
//    Delete()
//
//    Parameters

//          recID       [in] ID of the record to update
//
//    Purpose
//          Delete a record from database.
//

//    Return Value
//          S_OK for success, others for failure.
//
STDMETHODIMP CTxtABDatabase::Delete(RECORD_ID_TYPE recID)
{
        RECLIST::iterator it;
        for(it = m_recList.begin(); it!= m_recList.end(); it++)
        {
                if ((*it)->RecID() == recID)
                {
                        EnterCriticalSection(&m_cs);
                        m_recList.remove((*it));
                        LeaveCriticalSection(&m_cs);
                        return S_OK;
                }
        }

        return MAPI_E_NO_SUPPORT;
}

///////////////////////////////////////////////////////////////////////////////
//
//    Save()
//
//    Parameters

//          none.
//
//    Purpose
//          Save records to database permanently.
//

//    Return Value
//          S_OK for success, others for failure.
//
STDMETHODIMP CTxtABDatabase::Save()
{
        HRESULT	 hRes;
        FILE    *pFile = NULL;

        if (FAILED(hRes = HrOpenFile(OPEN_MODE_WRITE, &pFile)))
                goto LCleanup;

        SaveRecords(pFile);

LCleanup:
        if (pFile)
                fclose(pFile);

        return hRes;
}

///////////////////////////////////////////////////////////////////////////////
//
//    SaveRecords()
//
//    Parameters

//          pFile       [in] pointer to the database file
//
//    Purpose
//          Write records to the database file.
//

//    Return Value
//          none.
//
void CTxtABDatabase::SaveRecords(FILE *pFile)
{
        if (!pFile)
                return;

        EnterCriticalSection(&m_cs);

        RECLIST::iterator it;
        for(it = m_recList.begin(); it!= m_recList.end(); it++)
        {
                SaveRecord(*it, pFile);
        }

        LeaveCriticalSection(&m_cs);

        return;
}

///////////////////////////////////////////////////////////////////////////////
//
//    SaveRecord()
//
//    Parameters

//          pRec        [in] pointer to the record to save.
//          pFile       [in] pointer to the database file.
//
//    Purpose
//          Write a record to database file.
//

//    Return Value
//          none.
//
void CTxtABDatabase::SaveRecord(PCRecord pRec, FILE *pFile)
{
        if (!pRec || !pFile)
                return;

        TCHAR *pszProp = NULL;
        TCHAR *pszPropTag = NULL;

        // Write Props
        for(ULONG i = 0; i < pRec->cValues(); i++)
        {
                pszProp = new TCHAR[MAX_LINE_CHARS + 1];
                if (!pszProp)
                        goto LCleanup;

                pszPropTag = new TCHAR[MAX_LINE_CHARS + 1];
                if (!pszPropTag)
                        goto LCleanup;

                // Convert the ulPropTag string to decimal.
                if (_ltot_s(pRec->lpProps()[i].ulPropTag, pszPropTag, MAX_LINE_CHARS, 10))
                        goto LCleanup;

                _tcscpy_s(pszProp, MAX_LINE_CHARS, pszPropTag);
                _tcscat_s(pszProp, MAX_LINE_CHARS, TAG_VALUE_SEPARATOR);
                _tcscat_s(pszProp, MAX_LINE_CHARS, (TCHAR * )pRec->lpProps()[i].Value.LPSZ);
                _tcscat_s(pszProp, MAX_LINE_CHARS, LINE_BREAK);
                _fputts(pszProp, pFile);

                delete [] pszProp;
                pszProp = NULL;
                delete [] pszPropTag;
                pszPropTag = NULL;
        }

        // Write an empty line to the end of the record.
        _fputts(LINE_BREAK, pFile);


LCleanup:
        if (pszProp)
                delete [] pszProp;
        if (pszPropTag)
                delete [] pszPropTag;

}

/* End of the file TXTDB.CPP */