class  CSupport : public IMAPISupport
{

public:
	///////////////////////////////////////////////////////////////////////////////
	// Interface virtual member functions
	//
    inline STDMETHODIMP QueryInterface
		(REFIID                     riid,
		LPVOID *                   ppvObj)
	{
		return m_pMAPISup->QueryInterface(riid, ppvObj);
	};

    inline STDMETHODIMP_(ULONG) AddRef()
	{
		EnterCriticalSection(&m_csObj);
		++m_cRef;
		ULONG ulCount = m_cRef;
		LeaveCriticalSection(&m_csObj);
		return ulCount;
	}
    inline STDMETHODIMP_(ULONG) Release()
	{
		EnterCriticalSection(&m_csObj);
		ULONG ulCount = --m_cRef;
		LeaveCriticalSection(&m_csObj);
		if (!ulCount)
		{
			delete this;
		}
		return ulCount;
	}
    inline STDMETHODIMP GetLastError
		(HRESULT                    hResult,
		ULONG                      ulFlags,
		LPMAPIERROR *              ppMAPIError)
	{
		return m_pMAPISup->GetLastError(hResult, ulFlags, ppMAPIError);
	}

    inline STDMETHODIMP GetMemAllocRoutines
		(LPALLOCATEBUFFER *         pAllocateBuffer,
		LPALLOCATEMORE *           pAllocateMore,
		LPFREEBUFFER *             pFreeBuffer)
	{
		return m_pMAPISup->GetMemAllocRoutines(
			pAllocateBuffer,
			pAllocateMore,
			pFreeBuffer);
	}
    inline STDMETHODIMP Subscribe
		(LPNOTIFKEY                 pKey,
		ULONG                      ulEventMask,
		ULONG                      ulFlags,
		LPMAPIADVISESINK           pAdviseSink,
		ULONG *                    pulConnection)
	{
		return m_pMAPISup->Subscribe(
			pKey,
			ulEventMask,
			ulFlags,
			pAdviseSink,
			pulConnection);
	}
    inline STDMETHODIMP Unsubscribe
		(ULONG                      ulConnection)
	{
		return m_pMAPISup->Unsubscribe(ulConnection);
	}
    inline STDMETHODIMP Notify
		(LPNOTIFKEY                 pKey,
		ULONG                      cNotification,
		LPNOTIFICATION             pNotifications,
		ULONG *                    pulFlags)
	{
		return m_pMAPISup->Notify(
			pKey,
			cNotification,
			pNotifications,
			pulFlags);
	}
    STDMETHODIMP ModifyStatusRow
		(ULONG                      cValues,
		LPSPropValue               pColumnVals,
		ULONG                      ulFlags);


    STDMETHODIMP OpenProfileSection
		(LPMAPIUID                  pUid,
		ULONG                      ulFlags,
		LPPROFSECT *               ppProfileObj);

    inline STDMETHODIMP RegisterPreprocessor
		(LPMAPIUID                  pMuid,
		LPTSTR                     pszAdrType,
		LPTSTR                     pszDLLName,
		LPSTR                      pszPreprocess,
		LPSTR                      pszRemovePreprocessInfo,
		ULONG                      ulFlags)
	{
		return m_pMAPISup->RegisterPreprocessor(
			pMuid,
			pszAdrType,
			pszDLLName,
			pszPreprocess,
			pszRemovePreprocessInfo,
			ulFlags);
	}
    inline STDMETHODIMP NewUID
		(LPMAPIUID                  pMuid)
	{
		return m_pMAPISup->NewUID (pMuid);
	}
    inline STDMETHODIMP MakeInvalid
		(ULONG                      ulFlags,
		LPVOID                     lpObject,
		ULONG                      ulRefCount,
		ULONG                      cMethods)
	{
		return m_pMAPISup->MakeInvalid(
			ulFlags,
			lpObject,
			ulRefCount,
			cMethods);
	}
    inline STDMETHODIMP SpoolerYield
		(ULONG                      ulFlags)
	{
		return m_pMAPISup->SpoolerYield(ulFlags);
	}
    inline STDMETHODIMP SpoolerNotify
		(ULONG                      ulFlags,
		LPVOID                     pvData)
	{
		return m_pMAPISup->SpoolerNotify(ulFlags,pvData);
	}
    inline STDMETHODIMP CreateOneOff
		(LPTSTR                     pszName,
		LPTSTR                     pszAdrType,
		LPTSTR                     pszAddress,
		ULONG                      ulFlags,
		ULONG *                    pcbEntryID,
		LPENTRYID *                ppEntryID)
	{
		return m_pMAPISup->CreateOneOff(
			pszName,
			pszAdrType,
			pszAddress,
			ulFlags,
			pcbEntryID,
			ppEntryID);
	}
    inline STDMETHODIMP SetProviderUID
		(LPMAPIUID                  pProviderID,
		ULONG                      ulFlags)
	{
		return m_pMAPISup->SetProviderUID (
			pProviderID,
			ulFlags);
	}
    inline STDMETHODIMP CompareEntryIDs
		(ULONG                      cbEntry1,
		LPENTRYID                  pEntry1,
		ULONG                      cbEntry2,
		LPENTRYID                  pEntry2,
		ULONG                      ulCompareFlags,
		ULONG *                    pulResult)
	{
		return m_pMAPISup->CompareEntryIDs(
			cbEntry1,
			pEntry1,
			cbEntry2,
			pEntry2,
			ulCompareFlags,
			pulResult);
	}
    inline STDMETHODIMP OpenTemplateID
		(ULONG                      cbTemplateID,
		LPENTRYID                  pTemplateID,
		ULONG                      ulTemplateFlags,
		LPMAPIPROP                 pMAPIPropData,
		LPCIID                     pInterface,
		LPMAPIPROP *               ppMAPIPropNew,
		LPMAPIPROP                 pMAPIPropSibling)
	{
		return m_pMAPISup->OpenTemplateID(
			cbTemplateID,
			pTemplateID,
			ulTemplateFlags,
			pMAPIPropData,
			pInterface,
			ppMAPIPropNew,
			pMAPIPropSibling);
	}
    inline STDMETHODIMP OpenEntry
		(ULONG                      cbEntryID,
		LPENTRYID                  pEntryID,
		LPCIID                     pInterface,
		ULONG                      ulOpenFlags,
		ULONG *                    pulObjType,
		LPUNKNOWN *                ppUnk)
	{
		return m_pMAPISup->OpenEntry(
			cbEntryID,
			pEntryID,
			pInterface,
			ulOpenFlags,
			pulObjType,
			ppUnk);
	}
    inline STDMETHODIMP GetOneOffTable
		(ULONG                      ulFlags,
		LPMAPITABLE *              ppTable)
	{
		return m_pMAPISup->GetOneOffTable(
			ulFlags,
			ppTable);
	}
    inline STDMETHODIMP Address
		(ULONG *                    pulUIParam,
		LPADRPARM                  pAdrParms,
		LPADRLIST *                ppAdrList)
	{
		return m_pMAPISup->Address (
			pulUIParam,
			pAdrParms,
			ppAdrList);
	}
    inline STDMETHODIMP Details
		(ULONG *                    pulUIParam,
		LPFNDISMISS                pfnDismiss,
		LPVOID                     pvDismissContext,
		ULONG                      cbEntryID,
		LPENTRYID                  pEntryID,
		LPFNBUTTON                 pfButtonCallback,
		LPVOID                     pvButtonContext,
		LPTSTR                     pszButtonText,
		ULONG                      ulFlags)
	{
		return m_pMAPISup->Details (
			pulUIParam,
			pfnDismiss,
			pvDismissContext,
			cbEntryID,
			pEntryID,
			pfButtonCallback,
			pvButtonContext,
			pszButtonText,
			ulFlags);
	}
    inline STDMETHODIMP NewEntry
		(ULONG                      ulUIParam,
		ULONG                      ulFlags,
		ULONG                      cbEIDContainer,
		LPENTRYID                  pEIDContainer,
		ULONG                      cbEIDNewEntryTpl,
		LPENTRYID                  pEIDNewEntryTpl,
		ULONG *                    pcbEIDNewEntry,
		LPENTRYID *                ppEIDNewEntry)
	{
		return m_pMAPISup->NewEntry(
			ulUIParam,
			ulFlags,
			cbEIDContainer,
			pEIDContainer,
			cbEIDNewEntryTpl,
			pEIDNewEntryTpl,
			pcbEIDNewEntry,
			ppEIDNewEntry);
	}
    inline STDMETHODIMP DoConfigPropsheet
		(ULONG                      ulUIParam,
		ULONG                      ulFlags,
		LPTSTR                     pszTitle,
		LPMAPITABLE                pDisplayTable,
		LPMAPIPROP                 pCOnfigData,
		ULONG                      ulTopPage)
	{
		return m_pMAPISup->DoConfigPropsheet(
			ulUIParam,
			ulFlags,
			pszTitle,
			pDisplayTable,
			pCOnfigData,
			ulTopPage);
	}
    inline STDMETHODIMP CopyMessages
		(LPCIID                     pSrcInterface,
		LPVOID                      pSrcFolder,
		LPENTRYLIST                 pMsgList,
		LPCIID                      pDestInterface,
		LPVOID                      pDestFolder,
		ULONG                       ulUIParam,
		LPMAPIPROGRESS              pProgress,
		ULONG                       ulFlags)
	{
		return m_pMAPISup->CopyMessages(
			pSrcInterface,
			pSrcFolder,
			pMsgList,
			pDestInterface,
			pDestFolder,
			ulUIParam,
			pProgress,
			ulFlags);
	}
    inline STDMETHODIMP CopyFolder
		(LPCIID                     pSrcInterface,
		LPVOID                     pSrcFolder,
		ULONG                      cbEntryID,
		LPENTRYID                  pEntryID,
		LPCIID                     pDestInterface,
		LPVOID                     pDestFolder,
		LPTSTR                     szNewFolderName,
		ULONG                      ulUIParam,
		LPMAPIPROGRESS             pProgress,
		ULONG                      ulFlags)
	{
		return m_pMAPISup->CopyFolder(
			pSrcInterface,
			pSrcFolder,
			cbEntryID,
			pEntryID,
			pDestInterface,
			pDestFolder,
			szNewFolderName,
			ulUIParam,
			pProgress,
			ulFlags);
	}
    inline STDMETHODIMP DoCopyTo
		(LPCIID                     pSrcInterface,
		LPVOID                     pSrcObj,
		ULONG                      ciidExclude,
		LPCIID                     rgiidExclude,
		LPSPropTagArray            pExcludeProps,
		ULONG                      ulUIParam,
		LPMAPIPROGRESS             pProgress,
		LPCIID                     pDestInterface,
		LPVOID                     pDestObj,
		ULONG                      ulFlags,
		LPSPropProblemArray *      ppProblems)
	{
		return m_pMAPISup->DoCopyTo(
			pSrcInterface,
			pSrcObj,
			ciidExclude,
			rgiidExclude,
			pExcludeProps,
			ulUIParam,
			pProgress,
			pDestInterface,
			pDestObj,
			ulFlags,
			ppProblems);
	}
    inline STDMETHODIMP DoCopyProps
		(LPCIID                     pSrcInterface,
		LPVOID                     pSrcObj,
		LPSPropTagArray            pIncludeProps,
		ULONG                      ulUIParam,
		LPMAPIPROGRESS             pProgress,
		LPCIID                     pDestInterface,
		LPVOID                     pDestObj,
		ULONG                      ulFlags,
		LPSPropProblemArray *      ppProblems)
	{
		return m_pMAPISup->DoCopyProps(
			pSrcInterface,
			pSrcObj,
			pIncludeProps,
			ulUIParam,
			pProgress,
			pDestInterface,
			pDestObj,
			ulFlags,
			ppProblems);
	}
    inline STDMETHODIMP DoProgressDialog
		(ULONG                      ulUIParam,
		ULONG                      ulFlags,
		LPMAPIPROGRESS *           ppProgress)
	{
		return m_pMAPISup->DoProgressDialog(
			ulUIParam,
			ulFlags,
			ppProgress);
	}
    inline STDMETHODIMP ReadReceipt
		(ULONG                      ulFlags,
		LPMESSAGE                  pReadMessage,
		LPMESSAGE *                ppEmptyMessage)
	{
		return m_pMAPISup->ReadReceipt(
			ulFlags,
			pReadMessage,
			ppEmptyMessage);
	}
    inline STDMETHODIMP PrepareSubmit
		(LPMESSAGE                  pMessage,
		ULONG *                    pulFlags)
	{
		return m_pMAPISup->PrepareSubmit(
			pMessage,
			pulFlags);
	}
    inline STDMETHODIMP ExpandRecips
		(LPMESSAGE                  pMessage,
		ULONG *                    pulFlags)
	{
		return m_pMAPISup->ExpandRecips(
			pMessage,
			pulFlags);
	}
    inline STDMETHODIMP UpdatePAB
		(ULONG                      ulFlags,
		LPMESSAGE                  pMessage)
	{
		return m_pMAPISup->UpdatePAB(
			ulFlags,
			pMessage);
	}
    inline STDMETHODIMP DoSentMail
		(ULONG                      ulFlags,
		LPMESSAGE                  pMessage)
	{
		return m_pMAPISup->DoSentMail(
			ulFlags,
			pMessage);
	}
    inline STDMETHODIMP OpenAddressBook
		(LPCIID                     pInterface,
		ULONG                      ulFlags,
		LPADRBOOK *                ppAdrBook)
	{
		return m_pMAPISup->OpenAddressBook(
			pInterface,
			ulFlags,
			ppAdrBook);
	}
    inline STDMETHODIMP Preprocess
		(ULONG                      ulFlags,
		ULONG                      cbEntryID,
		LPENTRYID                  pEntryID)
	{
		return m_pMAPISup->Preprocess(
			ulFlags,
			cbEntryID,
			pEntryID);
	}
    inline STDMETHODIMP CompleteMsg
		(ULONG                      ulFlags,
		ULONG                      cbEntryID,
		LPENTRYID                  pEntryID)
	{
		return m_pMAPISup->CompleteMsg(
			ulFlags,
			cbEntryID,
			pEntryID);
	}
    inline STDMETHODIMP StoreLogoffTransports
		(ULONG *                    pulFlags)
	{
		return m_pMAPISup->StoreLogoffTransports (pulFlags);
	}
    inline STDMETHODIMP StatusRecips
		(LPMESSAGE                  pMessage,
		LPADRLIST                  pRecipList)
	{
		return m_pMAPISup->StatusRecips (pMessage, pRecipList);
	}
    inline STDMETHODIMP WrapStoreEntryID
		(ULONG                      cbOrigEntry,
		LPENTRYID                  pOrigEntry,
		ULONG *                    pcbWrappedEntry,
		LPENTRYID *                ppWrappedEntry)
	{
		return m_pMAPISup->WrapStoreEntryID(
			cbOrigEntry,
			pOrigEntry,
			pcbWrappedEntry,
			ppWrappedEntry);
	}
    inline STDMETHODIMP ModifyProfile
		(ULONG                      ulFlags)
	{
		return m_pMAPISup->ModifyProfile (ulFlags);
	}
    inline STDMETHODIMP IStorageFromStream
		(LPUNKNOWN                  pUnkIn,
		LPCIID                     pInterface,
		ULONG                      ulFlags,
		LPSTORAGE *                ppStorageOut)
	{
		return m_pMAPISup->IStorageFromStream(
			pUnkIn,
			pInterface,
			ulFlags,
			ppStorageOut);
	}
    inline STDMETHODIMP GetSvcConfigSupportObj
		(ULONG                      ulFlags,
		LPMAPISUP *                ppSvcSupport)
	{
		return m_pMAPISup->GetSvcConfigSupportObj(
			ulFlags,
			ppSvcSupport);
	}

	///////////////////////////////////////////////////////////////////////////////
	// Constructors and destructors
	//
public :
    CSupport(LPMAPISUP  pMAPISupObj, LPPROFSECT pProfSect);
    ~CSupport();

	///////////////////////////////////////////////////////////////////////////////
	// Data members
	//
public :
    ULONG            m_cRef;
    LPPROFSECT       m_lpProfSect;
    LPMAPISUP        m_pMAPISup;
    CRITICAL_SECTION m_csObj;
};