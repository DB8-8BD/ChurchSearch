#pragma once
namespace ATL {

	extern "C" __declspec(selectany) const IID IID_IConfigurationCache = { 0xf0637c70, 0xc937, 0x4156,{ 0xac, 0x3, 0x17, 0x9a, 0xee, 0xab, 0x2b, 0x2c } };
	extern "C" __declspec(selectany) const IID IID_IConnectionCache = { 0xe8168a6e, 0xeB6f, 0x44ef,{ 0x82, 0x58, 0x3e, 0xb1, 0xcc, 0xf9, 0x01, 0x51 } };
	extern "C" __declspec(selectany) const IID IID_IAuthCache = { 0xd77d5dba, 0x702b, 0x49f8,{ 0xab, 0x1a, 0x59, 0x27, 0x43, 0xa0, 0x40, 0x11 } };
	extern "C" __declspec(selectany) const IID IID_IHttpCCCache = { 0x5eb409c2, 0xdc9c, 0x4152,{ 0x88, 0x60, 0xfc, 0xe5, 0x27, 0x6f, 0xac, 0xf1 } };

	__interface ATL_NO_VTABLE __declspec(uuid("F0637C70-C937-4156-AC03-179AEEAB2B2C"))
		IConfigurationCache : public IUnknown
	{
		// IConfigurationCache Methods
		STDMETHOD(Add)(LPCSTR szKey, void *pvData, DWORD dwSize,
			FILETIME *pftExpireTime,
			HINSTANCE hInstClient, HCACHEITEM *phEntry,
			IMemoryCacheClient *pClient);

		STDMETHOD(LookupEntry)(LPCSTR szKey, HCACHEITEM * phEntry);
		STDMETHOD(GetData)(const HCACHEITEM hEntry, void **ppvData, DWORD *pdwSize) const;
		STDMETHOD(ReleaseEntry)(const HCACHEITEM hEntry);
		STDMETHOD(RemoveEntry)(const HCACHEITEM hEntry);
		STDMETHOD(RemoveEntryByKey)(LPCSTR szKey);

		STDMETHOD(Flush)();
	};

	__interface ATL_NO_VTABLE __declspec(uuid("E8168A6E-EB6F-44EF-8258-3EB1CCF90151"))
		IConnectionCache : public IUnknown
	{
		// IConnectionCache Methods
		STDMETHOD(Add)(ULONGLONG Key, CConnectionData* Data, DWORD dwSize,
			FILETIME *pftExpireTime,
			HINSTANCE hInstClient, HCACHEITEM *phEntry,
			IMemoryCacheClient *pClient);

		STDMETHOD(LookupEntry)(ULONGLONG szKey, HCACHEITEM * phEntry);
		STDMETHOD(GetData)(const HCACHEITEM hEntry, CConnectionData **ppData,
			DWORD *pdwSize) const;
		STDMETHOD(ReleaseEntry)(const HCACHEITEM hEntry);
		STDMETHOD(RemoveEntry)(const HCACHEITEM hEntry);
		STDMETHOD(RemoveEntryByKey)(ULONGLONG Key);

		STDMETHOD(Flush)();
	};

	__interface ATL_NO_VTABLE __declspec(uuid("D77D5DBA-702B-49F8-AB1A-592743A04011"))
		IAuthCache : public IUnknown
	{
		// IAuthCache Methods
		STDMETHOD(Add)(ULONGLONG Key, void *pvData, DWORD dwSize,
			FILETIME *pftExpireTime,
			HINSTANCE hInstClient, HCACHEITEM *phEntry,
			IMemoryCacheClient *pClient);

		STDMETHOD(LookupEntry)(ULONGLONG Key, HCACHEITEM * phEntry);
		STDMETHOD(GetData)(const HCACHEITEM hEntry, void **ppvData, DWORD *pdwSize) const;
		STDMETHOD(ReleaseEntry)(const HCACHEITEM hEntry);
		STDMETHOD(RemoveEntry)(const HCACHEITEM hEntry);
		STDMETHOD(RemoveEntryByKey)(ULONGLONG Key);

		STDMETHOD(Flush)();
	};

	__interface ATL_NO_VTABLE __declspec(uuid("5EB409C2-DC9C-4152-8860-FCE5276FACF1"))
		IHttpCCCache : public IUnknown
	{
		// IHttpCCCache Methods
		STDMETHOD(Add)(ULONGLONG Key, CHttpCacheControlData *pData, DWORD dwSize,
			FILETIME *pftExpireTime,
			HINSTANCE hInstClient, HCACHEITEM *phEntry,
			IMemoryCacheClient *pClient);

		STDMETHOD(LookupEntry)(ULONGLONG Key, HCACHEITEM * phEntry);
		STDMETHOD(GetData)(const HCACHEITEM hEntry, CHttpCacheControlData **ppData, DWORD *pdwSize) const;
		STDMETHOD(ReleaseEntry)(const HCACHEITEM hEntry);
		STDMETHOD(RemoveEntry)(const HCACHEITEM hEntry);
		STDMETHOD(RemoveEntryByKey)(ULONGLONG Key);

		STDMETHOD(Flush)();
	};

	template <class MonitorClass,
		class StatClass = CStdStatClass,
		class SyncObj = CComCriticalSection,
		class FlushClass = COldFlusher,
		class CullClass = CExpireCuller >
		class CConfigurationCache : public CMemoryCache<void*, StatClass, FlushClass, CFixedStringKey,
		CStringElementTraits<CFixedStringKey >, SyncObj, CullClass>,
		public IConfigurationCache,
		public IMemoryCacheControl,
		public IMemoryCacheStats,
		public IWorkerThreadClient
	{
		typedef CMemoryCache<void*, StatClass, FlushClass, CFixedStringKey,
			CStringElementTraits<CFixedStringKey>, SyncObj, CullClass> cacheBase;

		MonitorClass m_Monitor;

	protected:
		HANDLE m_hTimer;

	public:
		CConfigurationCache() : m_hTimer(NULL)
		{
		}

		HRESULT Initialize(IServiceProvider *pProv)
		{
			HRESULT hr = cacheBase::Initialize(pProv);
			if (FAILED(hr))
				return hr;
			hr = m_Monitor.Initialize();
			if (FAILED(hr))
				return hr;
			return m_Monitor.AddTimer(ATL_BLOB_CACHE_TIMEOUT,
				static_cast<IWorkerThreadClient*>(this), (DWORD_PTR) this, &m_hTimer);
		}

		template <class ThreadTraits>
		HRESULT Initialize(IServiceProvider *pProv, CWorkerThread<ThreadTraits> *pWorkerThread)
		{
			ATLASSERT(pWorkerThread);

			HRESULT hr = cacheBase::Initialize(pProv);
			if (FAILED(hr))
				return hr;

			hr = m_Monitor.Initialize(pWorkerThread);
			if (FAILED(hr))
				return hr;

			return m_Monitor.AddTimer(ATL_BLOB_CACHE_TIMEOUT,
				static_cast<IWorkerThreadClient*>(this), (DWORD_PTR) this, &m_hTimer);
		}

		HRESULT Execute(DWORD_PTR dwParam, HANDLE /*hObject*/)
		{
			CConfigurationCache* pCache = (CConfigurationCache*)dwParam;

			if (pCache)
				pCache->Flush();
			return S_OK;
		}

		HRESULT CloseHandle(HANDLE hObject)
		{
			ATLASSUME(m_hTimer == hObject);
			m_hTimer = NULL;
			::CloseHandle(hObject);
			return S_OK;
		}

		virtual ~CConfigurationCache()
		{
			if (m_hTimer)
			{
				ATLENSURE(SUCCEEDED(m_Monitor.RemoveHandle(m_hTimer)));
			}
		}

		HRESULT Uninitialize()
		{
			HRESULT hrMonitor = S_OK;
			if (m_hTimer)
			{
				hrMonitor = m_Monitor.RemoveHandle(m_hTimer);
				m_hTimer = NULL;
			}
			HRESULT hrShut = m_Monitor.Shutdown();
			HRESULT hrCache = cacheBase::Uninitialize();
			if (FAILED(hrMonitor))
			{
				return hrMonitor;
			}
			if (FAILED(hrShut))
			{
				return hrShut;
			}
			return hrCache;
		}
		// IUnknown methods
		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv)
		{
			HRESULT hr = E_NOINTERFACE;
			if (!ppv)
				hr = E_POINTER;
			else
			{
				if (InlineIsEqualGUID(riid, __uuidof(IUnknown)) ||
					InlineIsEqualGUID(riid, __uuidof(IConfigurationCache)))
				{
					*ppv = (IUnknown *)(IConfigurationCache *) this;
					AddRef();
					hr = S_OK;
				}
				if (InlineIsEqualGUID(riid, __uuidof(IMemoryCacheStats)))
				{
					*ppv = (IUnknown *)(IMemoryCacheStats*)this;
					AddRef();
					hr = S_OK;
				}
				if (InlineIsEqualGUID(riid, __uuidof(IMemoryCacheControl)))
				{
					*ppv = (IUnknown *)(IMemoryCacheControl*)this;
					AddRef();
					hr = S_OK;
				}

			}
			return hr;
		}

		ULONG STDMETHODCALLTYPE AddRef()
		{
			return 1;
		}

		ULONG STDMETHODCALLTYPE Release()
		{
			return 1;
		}

		// IMemoryCache Methods
		HRESULT STDMETHODCALLTYPE Add(LPCSTR szKey, void *pvData, DWORD dwSize,
			FILETIME *pftExpireTime,
			HINSTANCE hInstClient,
			HCACHEITEM *phEntry,
			IMemoryCacheClient *pClient)
		{
			HRESULT hr = E_FAIL;
			//if it's a multithreaded cache monitor we'll let the monitor take care of
			//cleaning up the cache so we don't overflow our configuration settings.
			//if it's not a threaded cache monitor, we need to make sure we don't
			//overflow the configuration settings by adding a new element
			if (m_Monitor.GetThreadHandle() == NULL)
			{
				if (!cacheBase::CanAddEntry(dwSize))
				{
					//flush the entries and check again to see if we can add
					cacheBase::FlushEntries();
					if (!cacheBase::CanAddEntry(dwSize))
						return E_OUTOFMEMORY;
				}
			}
			_ATLTRY
			{
				hr = cacheBase::AddEntry(szKey, pvData, dwSize,
				pftExpireTime, hInstClient, pClient, phEntry);
			return hr;
			}
				_ATLCATCHALL()
			{
				return E_FAIL;
			}
		}

		HRESULT STDMETHODCALLTYPE LookupEntry(LPCSTR szKey, HCACHEITEM * phEntry)
		{
			return cacheBase::LookupEntry(szKey, phEntry);
		}

		HRESULT STDMETHODCALLTYPE GetData(const HCACHEITEM hKey, void **ppvData, DWORD *pdwSize) const
		{
			return cacheBase::GetEntryData(hKey, ppvData, pdwSize);
		}

		HRESULT STDMETHODCALLTYPE ReleaseEntry(const HCACHEITEM hKey)
		{
			return cacheBase::ReleaseEntry(hKey);
		}

		HRESULT STDMETHODCALLTYPE RemoveEntry(const HCACHEITEM hKey)
		{
			return cacheBase::RemoveEntry(hKey);
		}

		HRESULT STDMETHODCALLTYPE RemoveEntryByKey(LPCSTR szKey)
		{
			return cacheBase::RemoveEntryByKey(szKey);
		}

		HRESULT STDMETHODCALLTYPE Flush()
		{
			return cacheBase::FlushEntries();
		}


		HRESULT STDMETHODCALLTYPE SetMaxAllowedSize(DWORD dwSize)
		{
			return cacheBase::SetMaxAllowedSize(dwSize);
		}

		HRESULT STDMETHODCALLTYPE GetMaxAllowedSize(DWORD *pdwSize)
		{
			return cacheBase::GetMaxAllowedSize(pdwSize);
		}

		HRESULT STDMETHODCALLTYPE SetMaxAllowedEntries(DWORD dwSize)
		{
			return cacheBase::SetMaxAllowedEntries(dwSize);
		}

		HRESULT STDMETHODCALLTYPE GetMaxAllowedEntries(DWORD *pdwSize)
		{
			return cacheBase::GetMaxAllowedEntries(pdwSize);
		}

		HRESULT STDMETHODCALLTYPE ResetCache()
		{
			return cacheBase::ResetCache();
		}

		// IMemoryCacheStats methods
		HRESULT STDMETHODCALLTYPE ClearStats()
		{
			m_statObj.ResetCounters();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetHitCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetHitCount();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetMissCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetMissCount();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetMaxAllocSize(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetMaxAllocSize();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetCurrentAllocSize(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetCurrentAllocSize();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetMaxEntryCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetMaxEntryCount();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetCurrentEntryCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetCurrentEntryCount();
			return S_OK;
		}

	}; // CConfigurationCache



	template <class MonitorClass, class StatClass = CStdStatClass, class SyncObj = CComCriticalSection,
		class FlushClass = COldFlusher, class CullClass = CExpireCuller >
		class CConnectionCache : public CMemoryCache<CConnectionData*, StatClass, FlushClass, ULONGLONG,
		CElementTraits<ULONGLONG>, SyncObj, CullClass>, public IConnectionCache, public IMemoryCacheControl,
		public IMemoryCacheStats, public IWorkerThreadClient
	{
		typedef CMemoryCache<CConnectionData*, StatClass, FlushClass, ULONGLONG,
			CElementTraits<ULONGLONG>, SyncObj, CullClass> cacheBase;

		MonitorClass m_Monitor;

	protected:
		HANDLE m_hTimer;

	public:
		CConnectionCache() : m_hTimer(NULL)
		{
		}

		HRESULT Initialize(IServiceProvider *pProv)
		{
			HRESULT hr = cacheBase::Initialize(pProv);
			if (FAILED(hr))
				return hr;
			hr = m_Monitor.Initialize();
			if (FAILED(hr))
				return hr;
			return m_Monitor.AddTimer(ATL_CONNECTION_CACHE_TIMEOUT,
				static_cast<IWorkerThreadClient*>(this),
				(DWORD_PTR) this, &m_hTimer);
		}

		template <class ThreadTraits>
		HRESULT Initialize(IServiceProvider *pProv, CWorkerThread<ThreadTraits> *pWorkerThread)
		{
			ATLASSERT(pWorkerThread);

			HRESULT hr = cacheBase::Initialize(pProv);
			if (FAILED(hr))
				return hr;

			hr = m_Monitor.Initialize(pWorkerThread);
			if (FAILED(hr))
				return hr;

			return m_Monitor.AddTimer(ATL_CONNECTION_CACHE_TIMEOUT,
				static_cast<IWorkerThreadClient*>(this),
				(DWORD_PTR) this, &m_hTimer);
		}

		HRESULT Execute(DWORD_PTR dwParam, HANDLE hObject)
		{
			CConnectionCache* pCache = (CConnectionCache*)dwParam;

			if (pCache)
				pCache->Flush();
			return S_OK;
		}

		HRESULT CloseHandle(HANDLE hObject)
		{
			ATLASSERT(m_hTimer == hObject);
			m_hTimer = NULL;
			::CloseHandle(hObject);
			return S_OK;
		}

		~CConnectionCache()
		{
			if (m_hTimer)
				m_Monitor.RemoveHandle(m_hTimer);
		}

		HRESULT Uninitialize()
		{
			if (m_hTimer)
			{
				m_Monitor.RemoveHandle(m_hTimer);
				m_hTimer = NULL;
			}
			m_Monitor.Shutdown();
			return cacheBase::Uninitialize();
		}
		// IUnknown methods
		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv)
		{
			HRESULT hr = E_NOINTERFACE;
			if (!ppv)
				hr = E_POINTER;
			else
			{
				if (InlineIsEqualGUID(riid, __uuidof(IUnknown)) ||
					InlineIsEqualGUID(riid, __uuidof(IConnectionCache)))
				{
					*ppv = (IUnknown *)(IConnectionCache *) this;
					AddRef();
					hr = S_OK;
				}
				if (InlineIsEqualGUID(riid, __uuidof(IMemoryCacheStats)))
				{
					*ppv = (IUnknown *)(IMemoryCacheStats*)this;
					AddRef();
					hr = S_OK;
				}
				if (InlineIsEqualGUID(riid, __uuidof(IMemoryCacheControl)))
				{
					*ppv = (IUnknown *)(IMemoryCacheControl*)this;
					AddRef();
					hr = S_OK;
				}

			}
			return hr;
		}

		ULONG STDMETHODCALLTYPE AddRef()
		{
			return 1;
		}

		ULONG STDMETHODCALLTYPE Release()
		{
			return 1;
		}

		// IMemoryCache Methods
		HRESULT STDMETHODCALLTYPE Add(ULONGLONG Key, CConnectionData *Data, DWORD dwSize,
			FILETIME *pftExpireTime,
			HINSTANCE hInstClient,
			HCACHEITEM *phEntry,
			IMemoryCacheClient *pClient)
		{
			HRESULT hr = E_FAIL;
			//if it's a multithreaded cache monitor we'll let the monitor 
			//take care of cleaning up the cache so we don't overflow 
			// our configuration settings. if it's not a threaded cache monitor, 
			// we need to make sure we don't overflow the configuration settings 
			// by adding a new element
			if (m_Monitor.GetThreadHandle() == NULL)
			{
				if (!cacheBase::CanAddEntry(dwSize))
				{
					//flush the entries and check again to see if we 
					// can add
					cacheBase::FlushEntries();
					if (!cacheBase::CanAddEntry(dwSize))
						return E_OUTOFMEMORY;
				}
			}
			_ATLTRY
			{
				hr = cacheBase::AddEntry(Key, Data, dwSize,
				pftExpireTime, hInstClient, pClient, phEntry);
			return hr;
			}
				_ATLCATCHALL()
			{
				return E_FAIL;
			}
		}

		// all other methods delegate to CacheBase or m_statObj

		HRESULT STDMETHODCALLTYPE LookupEntry(ULONGLONG Key, HCACHEITEM * phEntry)
		{
			return cacheBase::LookupEntry(Key, phEntry);
		}

		HRESULT STDMETHODCALLTYPE GetData(const HCACHEITEM hKey,
			CConnectionData **ppData, DWORD *pdwSize) const
		{
			return cacheBase::GetEntryData(hKey, ppData, pdwSize);
		}

		HRESULT STDMETHODCALLTYPE ReleaseEntry(const HCACHEITEM hKey)
		{
			return cacheBase::ReleaseEntry(hKey);
		}

		HRESULT STDMETHODCALLTYPE RemoveEntry(const HCACHEITEM hKey)
		{
			return cacheBase::RemoveEntry(hKey);
		}

		HRESULT STDMETHODCALLTYPE RemoveEntryByKey(ULONGLONG Key)
		{
			return cacheBase::RemoveEntryByKey(Key);
		}

		HRESULT STDMETHODCALLTYPE Flush()
		{
			return cacheBase::FlushEntries();
		}


		HRESULT STDMETHODCALLTYPE SetMaxAllowedSize(DWORD dwSize)
		{
			return cacheBase::SetMaxAllowedSize(dwSize);
		}

		HRESULT STDMETHODCALLTYPE GetMaxAllowedSize(DWORD *pdwSize)
		{
			return cacheBase::GetMaxAllowedSize(pdwSize);
		}

		HRESULT STDMETHODCALLTYPE SetMaxAllowedEntries(DWORD dwSize)
		{
			return cacheBase::SetMaxAllowedEntries(dwSize);
		}

		HRESULT STDMETHODCALLTYPE GetMaxAllowedEntries(DWORD *pdwSize)
		{
			return cacheBase::GetMaxAllowedEntries(pdwSize);
		}

		HRESULT STDMETHODCALLTYPE ResetCache()
		{
			return cacheBase::ResetCache();
		}

		// IMemoryCacheStats methods
		HRESULT STDMETHODCALLTYPE ClearStats()
		{
			m_statObj.ResetCounters();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetHitCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetHitCount();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetMissCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetMissCount();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetMaxAllocSize(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetMaxAllocSize();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetCurrentAllocSize(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetCurrentAllocSize();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetMaxEntryCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetMaxEntryCount();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetCurrentEntryCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetCurrentEntryCount();
			return S_OK;
		}

	}; // CConnectionCache

	template <class MonitorClass, class StatClass = CStdStatClass, class SyncObj = CComCriticalSection,
		class FlushClass = COldFlusher, class CullClass = CExpireCuller >
		class CAuthCache : public CMemoryCache<void*, StatClass, FlushClass, ULONGLONG,
		CElementTraits<ULONGLONG>, SyncObj, CullClass>, public IAuthCache, public IMemoryCacheControl,
		public IMemoryCacheStats, public IWorkerThreadClient
	{
		typedef CMemoryCache<void*, StatClass, FlushClass, ULONGLONG,
			CElementTraits<ULONGLONG>, SyncObj, CullClass> cacheBase;

		MonitorClass m_Monitor;

	protected:
		HANDLE m_hTimer;

	public:
		CAuthCache() : m_hTimer(NULL)
		{
		}

		HRESULT Initialize(IServiceProvider *pProv)
		{
			HRESULT hr = cacheBase::Initialize(pProv);
			if (FAILED(hr))
				return hr;
			hr = m_Monitor.Initialize();
			if (FAILED(hr))
				return hr;
			return m_Monitor.AddTimer(ATL_AUTH_CACHE_TIMEOUT,
				static_cast<IWorkerThreadClient*>(this),
				(DWORD_PTR) this, &m_hTimer);
		}

		template <class ThreadTraits>
		HRESULT Initialize(IServiceProvider *pProv, CWorkerThread<ThreadTraits> *pWorkerThread)
		{
			ATLASSERT(pWorkerThread);

			HRESULT hr = cacheBase::Initialize(pProv);
			if (FAILED(hr))
				return hr;

			hr = m_Monitor.Initialize(pWorkerThread);
			if (FAILED(hr))
				return hr;

			return m_Monitor.AddTimer(ATL_AUTH_CACHE_TIMEOUT,
				static_cast<IWorkerThreadClient*>(this),
				(DWORD_PTR) this, &m_hTimer);
		}

		HRESULT Execute(DWORD_PTR dwParam, HANDLE hObject)
		{
			CAuthCache* pCache = (CAuthCache*)dwParam;

			if (pCache)
				pCache->Flush();
			return S_OK;
		}

		HRESULT CloseHandle(HANDLE hObject)
		{
			ATLASSERT(m_hTimer == hObject);
			m_hTimer = NULL;
			::CloseHandle(hObject);
			return S_OK;
		}

		~CAuthCache()
		{
			if (m_hTimer)
				m_Monitor.RemoveHandle(m_hTimer);
		}

		HRESULT Uninitialize()
		{
			if (m_hTimer)
			{
				m_Monitor.RemoveHandle(m_hTimer);
				m_hTimer = NULL;
			}
			m_Monitor.Shutdown();
			return cacheBase::Uninitialize();
		}
		// IUnknown methods
		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv)
		{
			HRESULT hr = E_NOINTERFACE;
			if (!ppv)
				hr = E_POINTER;
			else
			{
				if (InlineIsEqualGUID(riid, __uuidof(IUnknown)) ||
					InlineIsEqualGUID(riid, __uuidof(IAuthCache)))
				{
					*ppv = (IUnknown *)(IAuthCache *) this;
					AddRef();
					hr = S_OK;
				}
				if (InlineIsEqualGUID(riid, __uuidof(IMemoryCacheStats)))
				{
					*ppv = (IUnknown *)(IMemoryCacheStats*)this;
					AddRef();
					hr = S_OK;
				}
				if (InlineIsEqualGUID(riid, __uuidof(IMemoryCacheControl)))
				{
					*ppv = (IUnknown *)(IMemoryCacheControl*)this;
					AddRef();
					hr = S_OK;
				}

			}
			return hr;
		}

		ULONG STDMETHODCALLTYPE AddRef()
		{
			return 1;
		}

		ULONG STDMETHODCALLTYPE Release()
		{
			return 1;
		}

		// IAuthCache Methods
		HRESULT STDMETHODCALLTYPE Add(ULONGLONG Key, void *pvData, DWORD dwSize,
			FILETIME *pftExpireTime,
			HINSTANCE hInstClient,
			HCACHEITEM *phEntry,
			IMemoryCacheClient *pClient)
		{
			HRESULT hr = E_FAIL;
			//if it's a multithreaded cache monitor we'll let the monitor 
			//take care of cleaning up the cache so we don't overflow 
			// our configuration settings. if it's not a threaded cache monitor, 
			// we need to make sure we don't overflow the configuration settings 
			// by adding a new element
			if (m_Monitor.GetThreadHandle() == NULL)
			{
				if (!cacheBase::CanAddEntry(dwSize))
				{
					//flush the entries and check again to see if we 
					// can add
					cacheBase::FlushEntries();
					if (!cacheBase::CanAddEntry(dwSize))
						return E_OUTOFMEMORY;
				}
			}
			_ATLTRY
			{
				hr = cacheBase::AddEntry(Key, pvData, dwSize,
				pftExpireTime, hInstClient, pClient, phEntry);
			return hr;
			}
				_ATLCATCHALL()
			{
				return E_FAIL;
			}
		}

		// all other methods delegate to CacheBase or m_statObj

		HRESULT STDMETHODCALLTYPE LookupEntry(ULONGLONG Key, HCACHEITEM * phEntry)
		{
			return cacheBase::LookupEntry(Key, phEntry);
		}

		HRESULT STDMETHODCALLTYPE GetData(const HCACHEITEM hKey,
			void **ppData, DWORD *pdwSize) const
		{
			return cacheBase::GetEntryData(hKey, ppData, pdwSize);
		}

		HRESULT STDMETHODCALLTYPE ReleaseEntry(const HCACHEITEM hKey)
		{
			return cacheBase::ReleaseEntry(hKey);
		}

		HRESULT STDMETHODCALLTYPE RemoveEntry(const HCACHEITEM hKey)
		{
			return cacheBase::RemoveEntry(hKey);
		}

		HRESULT STDMETHODCALLTYPE RemoveEntryByKey(ULONGLONG Key)
		{
			return cacheBase::RemoveEntryByKey(Key);
		}

		HRESULT STDMETHODCALLTYPE Flush()
		{
			return cacheBase::FlushEntries();
		}


		HRESULT STDMETHODCALLTYPE SetMaxAllowedSize(DWORD dwSize)
		{
			return cacheBase::SetMaxAllowedSize(dwSize);
		}

		HRESULT STDMETHODCALLTYPE GetMaxAllowedSize(DWORD *pdwSize)
		{
			return cacheBase::GetMaxAllowedSize(pdwSize);
		}

		HRESULT STDMETHODCALLTYPE SetMaxAllowedEntries(DWORD dwSize)
		{
			return cacheBase::SetMaxAllowedEntries(dwSize);
		}

		HRESULT STDMETHODCALLTYPE GetMaxAllowedEntries(DWORD *pdwSize)
		{
			return cacheBase::GetMaxAllowedEntries(pdwSize);
		}

		HRESULT STDMETHODCALLTYPE ResetCache()
		{
			return cacheBase::ResetCache();
		}

		// IMemoryCacheStats methods
		HRESULT STDMETHODCALLTYPE ClearStats()
		{
			m_statObj.ResetCounters();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetHitCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetHitCount();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetMissCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetMissCount();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetMaxAllocSize(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetMaxAllocSize();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetCurrentAllocSize(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetCurrentAllocSize();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetMaxEntryCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetMaxEntryCount();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetCurrentEntryCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetCurrentEntryCount();
			return S_OK;
		}

	}; // CAuthCache

	template <class MonitorClass, class StatClass = CStdStatClass, class SyncObj = CComCriticalSection,
		class FlushClass = COldFlusher, class CullClass = CExpireCuller >
		class CHttpCCCache : public CMemoryCache<CHttpCacheControlData*, StatClass, FlushClass, ULONGLONG,
		CElementTraits<ULONGLONG>, SyncObj, CullClass>, public IHttpCCCache, public IMemoryCacheControl,
		public IMemoryCacheStats, public IWorkerThreadClient
	{
		typedef CMemoryCache<CHttpCacheControlData*, StatClass, FlushClass, ULONGLONG,
			CElementTraits<ULONGLONG>, SyncObj, CullClass> cacheBase;

		MonitorClass m_Monitor;

	protected:
		HANDLE m_hTimer;

	public:
		CHttpCCCache() : m_hTimer(NULL)
		{
		}

		HRESULT Initialize(IServiceProvider *pProv)
		{
			HRESULT hr = cacheBase::Initialize(pProv);
			if (FAILED(hr))
				return hr;
			hr = m_Monitor.Initialize();
			if (FAILED(hr))
				return hr;
			return m_Monitor.AddTimer(ATL_HTTP_CONTROL_CACHE_TIMEOUT,
				static_cast<IWorkerThreadClient*>(this),
				(DWORD_PTR)this, &m_hTimer);
		}

		template <class ThreadTraits>
		HRESULT Initialize(IServiceProvider *pProv, CWorkerThread<ThreadTraits> *pWorkerThread)
		{
			ATLASSERT(pWorkerThread);

			HRESULT hr = cacheBase::Initialize(pProv);
			if (FAILED(hr))
				return hr;

			hr = m_Monitor.Initialize(pWorkerThread);
			if (FAILED(hr))
				return hr;

			return m_Monitor.AddTimer(ATL_HTTP_CONTROL_CACHE_TIMEOUT,
				static_cast<IWorkerThreadClient*>(this),
				(DWORD_PTR)this, &m_hTimer);
		}

		HRESULT Execute(DWORD_PTR dwParam, HANDLE hObject)
		{
			CHttpCCCache* pCache = (CHttpCCCache*)dwParam;

			if (pCache)
				pCache->Flush();
			return S_OK;
		}

		HRESULT CloseHandle(HANDLE hObject)
		{
			ATLASSERT(m_hTimer == hObject);
			m_hTimer = NULL;
			::CloseHandle(hObject);
			return S_OK;
		}

		~CHttpCCCache()
		{
			if (m_hTimer)
				m_Monitor.RemoveHandle(m_hTimer);
		}

		HRESULT Uninitialize()
		{
			if (m_hTimer)
			{
				m_Monitor.RemoveHandle(m_hTimer);
				m_hTimer = NULL;
			}
			m_Monitor.Shutdown();
			return cacheBase::Uninitialize();
		}
		// IUnknown methods
		HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **ppv)
		{
			HRESULT hr = E_NOINTERFACE;
			if (!ppv)
				hr = E_POINTER;
			else
			{
				if (InlineIsEqualGUID(riid, __uuidof(IUnknown)) ||
					InlineIsEqualGUID(riid, __uuidof(IHttpCCCache)))
				{
					*ppv = (IUnknown *)(IHttpCCCache *)this;
					AddRef();
					hr = S_OK;
				}
				if (InlineIsEqualGUID(riid, __uuidof(IMemoryCacheStats)))
				{
					*ppv = (IUnknown *)(IMemoryCacheStats*)this;
					AddRef();
					hr = S_OK;
				}
				if (InlineIsEqualGUID(riid, __uuidof(IMemoryCacheControl)))
				{
					*ppv = (IUnknown *)(IMemoryCacheControl*)this;
					AddRef();
					hr = S_OK;
				}

			}
			return hr;
		}

		ULONG STDMETHODCALLTYPE AddRef()
		{
			return 1;
		}

		ULONG STDMETHODCALLTYPE Release()
		{
			return 1;
		}

		// IHttpCCCache Methods
		HRESULT STDMETHODCALLTYPE Add(ULONGLONG Key, CHttpCacheControlData *pData, DWORD dwSize,
			FILETIME *pftExpireTime,
			HINSTANCE hInstClient,
			HCACHEITEM *phEntry,
			IMemoryCacheClient *pClient)
		{
			HRESULT hr = E_FAIL;
			//if it's a multithreaded cache monitor we'll let the monitor 
			//take care of cleaning up the cache so we don't overflow 
			// our configuration settings. if it's not a threaded cache monitor, 
			// we need to make sure we don't overflow the configuration settings 
			// by adding a new element
			if (m_Monitor.GetThreadHandle() == NULL)
			{
				if (!cacheBase::CanAddEntry(dwSize))
				{
					//flush the entries and check again to see if we 
					// can add
					cacheBase::FlushEntries();
					if (!cacheBase::CanAddEntry(dwSize))
						return E_OUTOFMEMORY;
				}
			}
			_ATLTRY
			{
				hr = cacheBase::AddEntry(Key, pData, dwSize,
				pftExpireTime, hInstClient, pClient, phEntry);
			return hr;
			}
				_ATLCATCHALL()
			{
				return E_FAIL;
			}
		}

		// all other methods delegate to CacheBase or m_statObj

		HRESULT STDMETHODCALLTYPE LookupEntry(ULONGLONG Key, HCACHEITEM * phEntry)
		{
			return cacheBase::LookupEntry(Key, phEntry);
		}

		HRESULT STDMETHODCALLTYPE GetData(const HCACHEITEM hKey,
			CHttpCacheControlData **ppData, DWORD *pdwSize) const
		{
			return cacheBase::GetEntryData(hKey, ppData, pdwSize);
		}

		HRESULT STDMETHODCALLTYPE ReleaseEntry(const HCACHEITEM hKey)
		{
			return cacheBase::ReleaseEntry(hKey);
		}

		HRESULT STDMETHODCALLTYPE RemoveEntry(const HCACHEITEM hKey)
		{
			return cacheBase::RemoveEntry(hKey);
		}

		HRESULT STDMETHODCALLTYPE RemoveEntryByKey(ULONGLONG Key)
		{
			return cacheBase::RemoveEntryByKey(Key);
		}

		HRESULT STDMETHODCALLTYPE Flush()
		{
			return cacheBase::FlushEntries();
		}


		HRESULT STDMETHODCALLTYPE SetMaxAllowedSize(DWORD dwSize)
		{
			return cacheBase::SetMaxAllowedSize(dwSize);
		}

		HRESULT STDMETHODCALLTYPE GetMaxAllowedSize(DWORD *pdwSize)
		{
			return cacheBase::GetMaxAllowedSize(pdwSize);
		}

		HRESULT STDMETHODCALLTYPE SetMaxAllowedEntries(DWORD dwSize)
		{
			return cacheBase::SetMaxAllowedEntries(dwSize);
		}

		HRESULT STDMETHODCALLTYPE GetMaxAllowedEntries(DWORD *pdwSize)
		{
			return cacheBase::GetMaxAllowedEntries(pdwSize);
		}

		HRESULT STDMETHODCALLTYPE ResetCache()
		{
			return cacheBase::ResetCache();
		}

		// IMemoryCacheStats methods
		HRESULT STDMETHODCALLTYPE ClearStats()
		{
			m_statObj.ResetCounters();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetHitCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetHitCount();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetMissCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetMissCount();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetMaxAllocSize(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetMaxAllocSize();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetCurrentAllocSize(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetCurrentAllocSize();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetMaxEntryCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetMaxEntryCount();
			return S_OK;
		}

		HRESULT STDMETHODCALLTYPE GetCurrentEntryCount(DWORD *pdwSize)
		{
			if (!pdwSize)
				return E_POINTER;
			*pdwSize = m_statObj.GetCurrentEntryCount();
			return S_OK;
		}

	}; // CHttpCCCache
}