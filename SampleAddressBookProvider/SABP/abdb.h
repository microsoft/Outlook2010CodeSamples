/****************************************************************************
        ABDB.H

        Copyright (c) 2007 Microsoft Corporation

        Declarations of database classes, data structures,

        function prototypes, defines, and so on
*****************************************************************************/


#ifndef _ABDB_H_
#define _ABDB_H_

#include "stdafx.h"
#include <list>

using namespace std;

// macro to verify if ulTag is multi-type.
#define ISMV_TAG(ulTag)         !!(MV_FLAG & PROP_TYPE((ulTag)))

class CABDatabase;
class CTxtABDatabase;
class CRecord;

// type definition for record ID, use ULONG as record ID type.
typedef	ULONG                   RECORD_ID_TYPE;
typedef CABDatabase             *PCABDatabase;
typedef CTxtABDatabase          *PCTxtABDatabase;
typedef CRecord                 *PCRecord;

BOOL WINAPI FCopyOneProp(LPSPropValue pspvSrc, LPSPropValue pspvDest);

void FreeProps(LPSPropValue pspv, ULONG cValues) ;
void FreeOneProp(LPSPropValue pspv);

class CRecord : public IUnknown
{
public:
        CRecord();
        CRecord(PCRecord pRec);
        ~CRecord();


        RECORD_ID_TYPE  RecID() const;
        ULONG           cValues() const;
        LPSPropValue    lpProps() const;


        void            SetRecID(RECORD_ID_TYPE recID);
        BOOL            FSetProps(LPSPropValue lpProps, ULONG cValues);


        MAPI_IUNKNOWN_METHODS(IMPL);

private:
        void            InitMembers();

private:
        volatile LONG   m_cRef;         // Reference count.
        RECORD_ID_TYPE  m_recID;
        ULONG           m_cValues;
        LPSPropValue    m_lpProps;
};

class CABDatabase abstract : public IUnknown
{
public:
        CABDatabase() : m_cRef(1) {};
        virtual ~CABDatabase(){};

        STDMETHODIMP QueryInterface(REFIID riid, LPVOID *ppvObj)
        {
                if (IsBadWritePtr(ppvObj, sizeof LPUNKNOWN))
                        return ERROR_INVALID_PARAMETER;

                *ppvObj = NULL;

                if (riid == IID_IUnknown)
                {
                        *ppvObj = static_cast<LPVOID>(this);
                        AddRef();
                        return NOERROR;
                }


                return E_NOINTERFACE;
        }

        STDMETHODIMP_ (ULONG) AddRef()
        {
                return InterlockedIncrement(&m_cRef);
        }

        STDMETHODIMP_ (ULONG) Release ()
        {
                ULONG ulRefCount = InterlockedDecrement(&m_cRef);

                // If this object is not being used, delete it
                if (ulRefCount == 0)
                {
                        // Destroy the object
                        delete this;
                }

                return ulRefCount;
        }


        virtual STDMETHODIMP Open() = 0;
        virtual STDMETHODIMP Close() = 0;
        virtual ULONG RecordCount() = 0;
        virtual PCRecord Read() = 0;
        virtual void GoBegin() = 0;
        virtual void GoEnd() = 0;
        virtual STDMETHODIMP Query(RECORD_ID_TYPE       recID,
                                   PCRecord             *ppRec) = 0;
        virtual STDMETHODIMP Update(RECORD_ID_TYPE      recID,
                                    ULONG               cValues,
                                    LPSPropValue        lpProps) = 0;
        virtual STDMETHODIMP Insert(ULONG               cValues,
                                    LPSPropValue        lpProps,
                                    RECORD_ID_TYPE     *pRecID) = 0;
        virtual STDMETHODIMP Delete(RECORD_ID_TYPE      recID) = 0;
        virtual STDMETHODIMP Save() = 0;

private:
        volatile LONG   m_cRef; // Reference count.
};


/*----------------------------------------------------------------------------
        Begin definition for CTxtABDatabase
        which is a simple implementation of CABDatabase.
----------------------------------------------------------------------------*/

#define TXTDB_HEADER_ID         _T("[ABDB]")
#define RECORD_ID_TAG           _T("RecID")
#define TAG_VALUE_SEPARATOR     _T(":")
#ifdef  UNICODE
typedef wstring                 TSTRING;
#else
typedef string                  TSTRING;
#endif

#define MAX_LINE_CHARS          255

const RECORD_ID_TYPE ulInitRecID = 0x00003E9;   // 1001

typedef list<PCRecord>  RECLIST;                // List type for store records.

class CTxtABDatabase : public CABDatabase
{
public:
        CTxtABDatabase(const TCHAR *filename);
        ~CTxtABDatabase();

        STDMETHODIMP    Open();
        STDMETHODIMP    Close();
        ULONG                   RecordCount();
        PCRecord                Read();
        void                    GoBegin();
        void                    GoEnd();
        STDMETHODIMP            Query(RECORD_ID_TYPE    recID,
                                      PCRecord          *ppRec);
        STDMETHODIMP            Update(RECORD_ID_TYPE   recID,
                                       ULONG            cValues,
                                       LPSPropValue     lpProps);
        STDMETHODIMP            Insert(ULONG            cValues,
                                       LPSPropValue     lpProps,
                                       RECORD_ID_TYPE     *pRecID);
        STDMETHODIMP            Delete(RECORD_ID_TYPE   recID);
        STDMETHODIMP            Save();

private:
        STDMETHODIMP    HrOpenFile(TCHAR *mode, FILE **ppFile);
        STDMETHODIMP    HrLoadRecords(FILE *pFile);
        void            SaveRecords(FILE *pFile);
        void            SaveRecord(PCRecord pRec, FILE *pFile);
        BOOL            FGetProps(list<TSTRING>  &listStrLine,
                                  ULONG           cPropValues,
                                  LPSPropValue    ppProp);
        TCHAR         * TrimStr(TCHAR *wcsBuf, int &nLen);
        BOOL            FDetectTag(TCHAR *wcsBuf, TCHAR *wcsTag);
        RECORD_ID_TYPE  NewRecID();
        BOOL            FPushRecord(list<TSTRING> &lstStrLine, RECLIST &recList);

private:
        CRITICAL_SECTION  m_cs;          // Critical Section.
        TCHAR             m_filename[MAX_LINE_CHARS];    // CHAR/WCHAR  array of file name
        RECLIST           m_recList;     // List to store records.
        RECLIST::iterator m_iter;        // Iterator of current pointer for Read() method.
};

#endif /* _ABDB_H_ */
