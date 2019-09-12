#pragma once
namespace UTIL
{
	/*inline CStringA ToCStringA(LPCWSTR wsz) throw()
	{
		return wsz;
	}

	inline CStringA ToCStringA(const VARIANT& v) throw()
	{
		CComVariant var(v);
		return (SUCCEEDED(var.ChangeType(VT_BSTR))) ? ToCStringA(var.bstrVal) : "";
	}

	inline CStringA ToCStringA(const VARIANT* pv) throw()
	{
		return pv ? ToCStringA(*pv) : "";
	}

	inline CStringA ToCStringA(LONG l) throw()
	{
		CStringA str;
		str.Format("%ld", l);
		return str;
	}*/

	inline WORD StringToWord(const char* p, char** pp) throw()
	{
		return static_cast<WORD>(strtoul(p, pp, 10));
	}


} // namespace UTIL
