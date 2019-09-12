#ifndef _W3PIPOOL_H_INCLUDED
#define _W3PIPOOL_H_INCLUDED
__interface	IThreadPoolWorker
{
	virtual void HandleTheRequest(void *pvParam, OVERLAPPED *pOverlapped) = 0;
};
#include <w3piworker.h>
#include <remapper.h>
#include <IThreadPoolService.h>

//
// CThreadPoolService - A thread pool controller class for W3Pi
//
class CThreadPoolService : public IThreadPoolService
{
public:
	CRemapper *m_premapper = NULL;
	CComPtr<IIsapiExtension> m_spExtension;
	HANDLE m_hRequestQueue;
	CThreadPoolService(void)
	{
		m_critSec.Init();
		m_bInitialized = false;
		m_dwTlsIndex = TlsAlloc();
	}
	~CThreadPoolService(void)
	{
		m_threadPool.Shutdown();
		m_critSec.Term();
		TlsFree(m_dwTlsIndex);
	}
	int	m_numThreads;
	HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv)
	{
		if (!ppv)
			return E_POINTER;
		if (InlineIsEqualGUID(riid, __uuidof(IThreadPoolService)))
		{
			*ppv = static_cast<IUnknown*>(this);
			AddRef();
			return S_OK;
		}
		else
			return m_threadPool.QueryInterface(riid, ppv);
	}
	ULONG STDMETHODCALLTYPE AddRef()
	{
		return 1;
	}
	ULONG STDMETHODCALLTYPE Release()
	{
		return 1;
	}
	IIsapiExtension* GetExtension()
	{
		return m_spExtension;
	}
	bool QueueRequest(const CW3PiWorker::RequestType& request)
	{
		if (0 != m_threadPool.QueueRequest(request))
			return true;
		else
			return false;
	}
	HRESULT Initialize(int iNumThreads, HANDLE hRequestQueue)
	{
		m_hRequestQueue = hRequestQueue;
		m_numThreads = iNumThreads;
		IThreadPoolService *pIFace = static_cast<IThreadPoolService*>(this);
		HRESULT	hResult = S_OK;
		if (!m_bInitialized)
		{
			hResult = m_threadPool.Initialize((LPVOID)pIFace, iNumThreads, 0, hRequestQueue);
			m_bInitialized = SUCCEEDED(hResult);
		}
		return hResult;
	}
	void CancelRequests()
	{
		unsigned long ulResult;
		POSITION myPos = m_jobStatus.GetStartPosition();
		while (myPos != NULL)
		{
			ULONGLONG ulRequestId;
			eJobStatus status;
			m_jobStatus.GetNextAssoc(myPos, ulRequestId, status);
			ulResult = HttpCancelHttpRequest(m_hRequestQueue, ulRequestId, NULL);
			RemoveRequest(ulRequestId);
		}
	}
	eJobStatus CanCancelRequest(ULONGLONG ulRequestId)
	{
		eJobStatus jobStatus;
		bool bRet = m_jobStatus.Lookup(ulRequestId, jobStatus);
		if (!bRet)
			jobStatus = eSTATUS_ERROR;
		return jobStatus;
	}
	eConnectionStatus CanCloseConnection(ULONGLONG ulConnectionId)
	{
		eConnectionStatus connectionStatus;
		bool bRet = m_ConnectionStatus.Lookup(ulConnectionId, connectionStatus);
		if (!bRet)
			connectionStatus = eSTATUS_CLOSED;
		return connectionStatus;
	}
	bool ConnectionHasOutstandingRequests(ULONGLONG ulConnectionId)
	{
		ULONGLONG ulRequestId, ulCnxnId;
		POSITION myPos = m_SessionMap.GetStartPosition();
		while (myPos != NULL)
		{
			m_SessionMap.GetNextAssoc(myPos, ulRequestId, ulCnxnId);
			if (ulConnectionId == ulCnxnId)
			{
				if (eSTATUS_OPEN == CanCloseConnection(ulCnxnId))
				{
					return true;
				}
			}
		}
		return false;
	}
	void RemoveRequest(ULONGLONG ulRequestId)
	{
		m_jobStatus.RemoveKey(ulRequestId);
	}
	void KillConnection(ULONGLONG ulConnectionId)
	{
		//kill connection will remove all outstanding requests on a connection
		ULONGLONG ulRequestId, ulCnxnId;
		POSITION myPos = m_SessionMap.GetStartPosition();
		while (myPos != NULL)
		{
			m_SessionMap.GetNextAssoc(myPos, ulRequestId, ulCnxnId);
			if (ulConnectionId == ulCnxnId)
			{
				if (eSTATUS_OPEN == CanCloseConnection(ulCnxnId))
				{
					HttpCancelHttpRequest(m_hRequestQueue, ulRequestId, NULL);
					RemoveRequest(ulRequestId);
				}
				m_SessionMap.RemoveKey(ulRequestId);
			}
		}
		setConnectionStatus(ulConnectionId, eSTATUS_CLOSED);
	}
	void Shutdown(DWORD dwMaxWait)
	{
		m_threadPool.Shutdown(dwMaxWait);
	}
	void setJobStatus(ULONGLONG ulRequestId, eJobStatus status)
	{
		m_critSec.Lock();
		if (status == eSTATUS_DONE)
			m_jobStatus.RemoveKey(ulRequestId);
		else
			m_jobStatus.SetAt(ulRequestId, status);
		m_critSec.Unlock();
	}
	void setConnectionStatus(ULONGLONG ulConnectionId, eConnectionStatus status)
	{
		m_critSec.Lock();
		if (status == eSTATUS_CLOSED)
		{
			m_ConnectionStatus.RemoveKey(ulConnectionId);
			RemoveSession(ulConnectionId);
		}
		else
			m_ConnectionStatus.SetAt(ulConnectionId, status);
		m_critSec.Unlock();
	}
	void setSessionStatus(ULONGLONG ulRequestId, ULONGLONG ulConnectionId)
	{
		m_critSec.Lock();
		m_SessionMap.SetAt(ulRequestId, ulConnectionId);
		m_critSec.Unlock();
	}
	bool SetThreadWorker(CW3PiWorker *pWorker) throw()
	{
		if (0 != TlsSetValue(m_dwTlsIndex, (void*)pWorker))
			return true;
		else
			return false;
	}
	CW3PiWorker *GetThreadWorker() throw()
	{
		return (CW3PiWorker *)TlsGetValue(m_dwTlsIndex);
	}
	BOOL OnThreadAttach()
	{
		return SUCCEEDED(CoInitializeEx(NULL, COINIT_APARTMENTTHREADED));
	}
	void OnThreadTerminate()
	{
		CoUninitialize();
	}
	void RemoveSession(ULONGLONG ulConnectionId)
	{
		//remove any remaining 
		ULONGLONG ulRequestId, ulCnxnId;
		POSITION myPos = m_SessionMap.GetStartPosition();
		while (myPos != NULL)
		{
			m_SessionMap.GetNextAssoc(myPos, ulRequestId, ulCnxnId);
			if (ulConnectionId == ulCnxnId)
			{
				m_SessionMap.RemoveKey(ulRequestId);
			}
		}
	}
	void setRemapper(CRemapper *premapper)
	{
		m_premapper = premapper;
	}
	CRemapper *GetRemapper() throw()
	{
		return (CRemapper *)&m_premapper;
	}
protected:
	CThreadPool<CW3PiWorker, Win32ThreadTraits> m_threadPool;
	CAtlMap<ULONGLONG, eJobStatus> m_jobStatus;
	CAtlMap<ULONGLONG, eConnectionStatus> m_ConnectionStatus;
	CAtlMap<ULONGLONG, ULONGLONG> m_SessionMap;
	CComCriticalSection	m_critSec;
	DWORD m_dwTlsIndex;
	BOOL m_bInitialized;
}; // CThreadPoolService
inline bool CW3PiWorker::Initialize(void* pvParam)
{
	HRESULT hRet = S_OK;
	CComPtr<IThreadPoolService> spThreadPoolSvc;
	spThreadPoolSvc = static_cast<IThreadPoolService*>(pvParam);
	hRet = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	ATLASSUME(SUCCEEDED(hRet));
	m_hHeap = HeapCreate(HEAP_NO_SERIALIZE, 24576, 0);
	if (!m_hHeap)
		return false;
	return spThreadPoolSvc->SetThreadWorker(this);
}
inline void CW3PiWorker::Execute(RequestType dw, void *pvParam, OVERLAPPED * pOverlapped)
{
	ATLASSUME(0 != dw);
	IThreadPoolWorker* pRequest = reinterpret_cast<IThreadPoolWorker*>(dw);
	pRequest->HandleTheRequest(pvParam, pOverlapped);
}
inline void CW3PiWorker::Terminate(void* pvParam)
{
	if (m_hHeap)
	{
		if (HeapDestroy(m_hHeap))
			m_hHeap = NULL;
		else
		{
			ATLASSERT(FALSE);
		}
	}
	CoUninitialize();
}
#endif