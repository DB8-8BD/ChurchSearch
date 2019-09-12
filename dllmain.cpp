#include "stdafx.h"
#include "dllmain.h"
class CChurchSearchModule : public CAtlDllModuleT<CChurchSearchModule>
{
public:
	BOOL WINAPI DllMain(DWORD dwReason, LPVOID lpReserved) throw()
	{
		return __super::DllMain(dwReason, lpReserved);
	}
};

CChurchSearchModule _AtlModule;

BEGIN_HANDLER_MAP()
	HANDLER_ENTRY("Default", CChurchSearchHandler)
END_HANDLER_MAP()

extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	hInstance;
	return TRUE;
}
