#ifndef _W3PIWORKER_H_INCLUDED
#define _W3PIWORKER_H_INCLUDED
class CW3PiWorker
{
public:
	typedef AtlServerRequest* RequestType;
	HANDLE m_hHeap;
	CW3PiWorker(void){}
	~CW3PiWorker(void){}
	virtual bool Initialize(void* pvParam);
	void Execute(RequestType dw, void *pvParam, OVERLAPPED * pOverlapped) throw();
	virtual void Terminate(void* pvParam);
	virtual bool GetWorkerData(DWORD dwParam, void** ppvData)
	{
		return FALSE;
	}
}; //CW3PiWorker
#endif