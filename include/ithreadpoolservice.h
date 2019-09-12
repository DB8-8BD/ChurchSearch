__interface ATL_NO_VTABLE __declspec(uuid("746a6287-2fad-45e9-91e7-7c35e3d97804"))
IThreadPoolService : public IUnknown
{
	HRESULT  STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv);
	ULONG STDMETHODCALLTYPE AddRef();
	ULONG STDMETHODCALLTYPE Release();
	HRESULT Initialize(int iNumThreads, HANDLE hRequestQueue);
	bool QueueRequest(const CW3PiWorker::RequestType& request);
	eJobStatus CanCancelRequest(ULONGLONG m_lRequestId);
	eConnectionStatus CanCloseConnection(ULONGLONG ulConnectionId);
	bool ConnectionHasOutstandingRequests(ULONGLONG ulConnectionId);
	void CancelRequests();
	void RemoveRequest(ULONGLONG ulRequestId);
	void RemoveSession(ULONGLONG ulConnectionId);
	void KillConnection(ULONGLONG ulConnectionId);
	void Shutdown(DWORD dwMaxWait);
	void setJobStatus(ULONGLONG ulRequestId, eJobStatus status);
	void setConnectionStatus(ULONGLONG ulConnectionId, eConnectionStatus status);
	void setSessionStatus(ULONGLONG ulRequestId, ULONGLONG ulConnectionId);
	CW3PiWorker *GetThreadWorker();
	bool SetThreadWorker(CW3PiWorker *pWorker);
	IIsapiExtension* GetExtension();
	void setRemapper(CRemapper *premapper);
	CRemapper *GetRemapper();
};