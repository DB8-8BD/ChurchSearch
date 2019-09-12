#pragma once
namespace UTIL
{

	#pragma comment(lib, "Rpcrt4.lib")

	typedef unsigned char VCUE_RPC_CHAR;

	inline bool GuidFromString(const CStringA& strGuid, GUID& guid) throw()
	{
		bool bSuccess = false;
		RPC_STATUS status = UuidFromStringA(reinterpret_cast<VCUE_RPC_CHAR*>(const_cast<LPSTR>(static_cast<LPCSTR>(strGuid))), &guid);
		if (RPC_S_OK == status)
			bSuccess = true;
		return bSuccess;
	}

	inline bool StringFromGuid(const GUID& guid, CStringA& strGuid) throw()
	{
		bool bSuccess = false;
		VCUE_RPC_CHAR* pstr = 0;
		RPC_STATUS status = UuidToStringA(const_cast<GUID*>(&guid), &pstr);
		if (RPC_S_OK == status)
		{
			strGuid = pstr;
			RpcStringFreeA(&pstr);
			bSuccess = true;
		}
		return bSuccess;
	}


}; // namespace UTIL
