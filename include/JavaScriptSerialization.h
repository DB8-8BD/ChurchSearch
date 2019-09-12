// JavaScriptSerialization.h: JavaScript Object Name (De)Serializer classes
//This code and information is provided "as is" without warranty of
//	any kind, either expressed or implied, including but not limited to
//	the implied warranties of merchantability and/or fitness for a
//	particular purpose.
//This is sample code and is freely distributable under the following condition:
//   it is strictly required that all flaws discovered in the code (or comments) of even the
//   most minor severity are reported to Gregory Morse at email address prog321@aol.com
//Authored on 3/25/2010 by Gregory Morse

#pragma once

#define _CRT_SECURE_NO_WARNINGS
#include <stdio.h>
#include <comdef.h>
#include <comutil.h>
#include <atlutil.h>
#include <atlcomtime.h>
#include <atlsafe.h>
#include <atlcoll.h>
#include <regex>
#include <new>

#define V_IS_NUMBER(v) ((V_VT(v) == VT_I4) || (V_VT(v) == VT_I8) || (V_VT(v) == VT_UI8))

//returns 0 on when not checked with is number first though could throw error
#define V_GET_NUMBER(v) ((V_VT(v) == VT_I4) ? V_I4(v) : ((V_VT(v) == VT_I8) ? V_I8(v) : \
														((V_VT(v) == VT_UI8) ? V_UI8(v) : 0)))

class JavaScriptString
{
public:
	JavaScriptString(const char* s)
	{
		_s = s;
		_index = 0;
	}
private:
	static void AppendCharAsUnicode(CStringA & builder, const wchar_t c)
	{
		builder += "\\u";
		builder.AppendFormat("%04x", c);
	}
public:
	CString GetDebugString(const TCHAR* message) const
	{
		CString str;
		str.Format(_T("%s (%lu): %hs"), message, _index, _s);
		return str;
	}
	char GetNextNonEmptyChar()
	{
		while ((_s.GetLength() > _index)) {
			char c = _s[_index++];
			if (!isspace(c)) {
				return c;
			}
		}
		return 0;
	}
	char MoveNext()
	{
		if (_s.GetLength() > _index) {
			return (_s[_index++]);
		}
		return 0;
	}
	CStringA MoveNext(int count)
	{
		CStringA str;
		if (_s.GetLength() >= (_index + count)) {
			str = _s.Mid(_index, count);
			_index += count;
		}
		return str;
	}
	void MovePrev()
	{
		if (_index > 0) {
			_index--;
		}
	}
	void MovePrev(int count)
	{
		while (((_index > 0) && (count > 0))) {
			_index--;
			count--;
		}
	}
	static CStringA QuoteString(LPCWSTR value)
	{
		CStringA builder;
		if (!value) {
			throw;
		}
		int startIndex = 0;
		int count = 0;
		int len = wcslen(value);
		for (int i = 0; (i < len); i++) {
			wchar_t c = value[i];
			if ((((c == '\r') || (c == '\t')) || 
				 ((c == '\"') || (c == '\''))) || 
				((((c == '<') || (c == '>')) || 
				  ((c == '\\') || (c == '\n'))) || 
				 (((c == '\b') || (c == '\f')) || (c < ' ')))) {
				if (count > 0) {
					builder.Append(CStringA(&value[startIndex], count));
				}
				startIndex = (i + 1);
				count = 0;
			}
			switch (c) {
				case '<':
				case '>':
				case '\'':
				{
					AppendCharAsUnicode(builder, c);
					continue;
				}
				case '\\':
				{
					builder.Append("\\\\");
					continue;
				}
				case '\b':
				{
					builder.Append("\\b");
					continue;
				}
				case '\t':
				{
					builder.Append("\\t");
					continue;
				}
				case '\n':
				{
					builder.Append("\\n");
					continue;
				}
				case '\f':
				{
					builder.Append("\\f");
					continue;
				}
				case '\r':
				{
					builder.Append("\\r");
					continue;
				}
				case '\"':
				{
					builder.Append("\\\"");
					continue;
				}
			}
			if (c < ' ') {
				AppendCharAsUnicode(builder, c);
			} else {
				count++;
			}
		}
		if (count > 0) {
			builder.Append(CStringA(&value[startIndex], count));
		}
		return builder;
	}

	static CStringA QuoteString(LPCWSTR value, const bool addQuotes)
	{
		CStringA str = QuoteString(value);
		if (addQuotes) {
			str = "\"" + str + "\"";
		}
		return str;
	}

	CStringA ToString() const
	{
		if (_s.GetLength() > _index) {
			return _s.Mid(_index);
		}
		return CStringA();
	}
private:
	int _index;
	CStringA _s;
};

#define JSON_BadEscape _T("Bad escape")
#define JSON_IllegalPrimitive _T("Illegal Primitive")
#define JSON_StringNotQuoted _T("String Not Quoted")
#define JSON_ExpectedOpenBrace _T("Expected Open Brace")
#define JSON_InvalidMemberName _T("Invalid Member Name")
#define JSON_InvalidObject _T("Invalid Object")
#define JSON_DepthLimitExceeded _T("Depth Limit Exceeded")
#define JSON_InvalidArrayStart _T("Invalid Array Start")
#define JSON_InvalidArrayExtraComma _T("Invalid Array Extra Comma")
#define JSON_InvalidArrayExpectComma _T("Invalid Array Expect Comma")
#define JSON_InvalidArrayEnd _T("Invalid Array End")
#define JSON_InvalidMaxJsonLength _T("Invalid Maximum JSON length exceeded")
#define JSON_InvalidRecursionLimit _T("Invalid Recursion Limit exceeded")
#define JSON_MaxJsonLengthExceeded _T("Maximum JSON length exceeded")
#define JSON_CircularReference _T("Circular Reference")

static __int64 GetDatetimeMinTimeTicks()
{
	SYSTEMTIME st;
	CFileTime FileTime;
	struct tm atm = { 0, 0, 0, 1, 1 - 1, 1970 - 1900, 0, 0, 0 };

	CTime Time(_mkgmtime64(&atm));
	if (!Time.GetAsSystemTime(st) || !SystemTimeToFileTime(&st, &FileTime)) {
		throw; //cannot load without this
	}
	return FileTime.GetTime();
};

static const int c_DefaultMaxJsonLength = 0x200000;
static const int c_DefaultRecursionLimit = 100;
static __int64 DatetimeMinTimeTicks = GetDatetimeMinTimeTicks();

class CTASimpleException
{
public:
	CTASimpleException(const TCHAR* szString) : m_String(szString)
	{
	}
	CString GetString() { return m_String; }
protected:
	CString m_String;
};

class CTANotSupportedException : public CTASimpleException
{
public:
	CTANotSupportedException(const TCHAR* szString) : CTASimpleException(szString)
	{
	}
};

class CTAInvalidArgException : public CTASimpleException
{
public:
	CTAInvalidArgException(const TCHAR* szString) : CTASimpleException(szString)
	{
	}
};

//outside of the reserved variant range for better expandability
#define VT_DICTIONARY 74
#define VT_URI 75

//must not use VariantClear or VariantCopy/VariantCopyInd
//	instead use the destructor or ExtendedVariantClear
class _tavariant_t : public _variant_t
{
public:
	struct DictionaryVariant
	{
		VARTYPE vt;
		CAtlMap<CStringW, VARIANT, CStringElementTraitsI<CStringW> > Dictionary;
	};
	struct FileTimeVariant
	{
		VARTYPE vt;
		FILETIME FileTime;
	};
	struct CLSIDVariant
	{
		VARTYPE vt;
		CLSID clsid;
	};
	struct UriVariant
	{
		VARTYPE vt;
		CUrl Uri;
	};

	_tavariant_t() : _variant_t() {}
	_tavariant_t(const VARIANT& varSrc) : _variant_t(varSrc) {}
    _tavariant_t(const VARIANT* pSrc) : _variant_t(pSrc) {}
    _tavariant_t(const _variant_t& varSrc) : _variant_t(varSrc) {}
    _tavariant_t(VARIANT& varSrc, bool fCopy) : _variant_t(varSrc, fCopy) {}
	virtual ~_tavariant_t()
	{
		ExtendedVariantClear();
	}

	void ExtendedVariantClear()
	{
		ExtendedVariantClear(0);
	}
protected:
	void ExtendedDictionaryClear(CAtlMap<CStringW, VARIANT,
								 CStringElementTraitsI<CStringW> > & DictionaryClear, 
								 int depth)
	{
		POSITION pos = DictionaryClear.GetStartPosition();
		CAtlMap<CStringW, VARIANT, CStringElementTraitsI<CStringW> >::CPair* pPair;
		while (pos && (pPair = DictionaryClear.GetNext(pos))) {
			_tavariant_t* varCurrent = new _tavariant_t();
			CAutoPtr<void> CleanupPtr(varCurrent);
			varCurrent->Attach(VARIANT(pPair->m_value));
			varCurrent->ExtendedVariantClear(depth);
		}
	}

	void ExtendedSafeArrayClear(CComSafeArray<VARIANT> & SafeArrayClear, int depth)
	{
		long lCount;
		long lUpperBound = SafeArrayClear.GetUpperBound();
		if (SafeArrayClear.GetDimensions() != 1) {
			throw new CTAInvalidArgException(JSON_InvalidObject);
		}
		for (lCount = SafeArrayClear.GetLowerBound(); lCount <= lUpperBound; lCount++) {
			_tavariant_t* varCurrent = new _tavariant_t();
			CAutoPtr<void> CleanupPtr(varCurrent);
			varCurrent->Attach(VARIANT(SafeArrayClear.GetAt(lCount)));
			//break circular references
			//must not use SetAt since special VARIANT provided version either calls
			//VariantClear or VariantCopyInd!
			(&SafeArrayClear.GetAt(lCount))->vt = VT_EMPTY;
			varCurrent->ExtendedVariantClear(depth);
		}
	}

	//recursively cleanup special types by going through
	//	all containers - safearrays and custom dictionaries
	void ExtendedVariantClear(int depth)
	{
		HRESULT hr;
		if (++depth > c_DefaultRecursionLimit) {
			throw new CTAInvalidArgException(JSON_DepthLimitExceeded);
		}
		if (vt == (VT_ARRAY | VT_VARIANT)) {
			CComSafeArray<VARIANT> SafeArray;
			if (S_OK == SafeArray.Attach(V_ARRAY(this))) {
				ExtendedSafeArrayClear(SafeArray, depth);
			}
			SafeArray.Detach();
		} else if (	vt == (VT_BYREF | VT_UNKNOWN) &&
					*((VARTYPE*)V_BYREF(this)) == VT_DICTIONARY) {
			*((VARTYPE*)V_BYREF(this)) = VT_EMPTY; //break circular references
			ExtendedDictionaryClear(((DictionaryVariant*)V_BYREF(this))->Dictionary,
									depth);
			delete (DictionaryVariant*)V_BYREF(this);
		} else if (	vt == (VT_BYREF | VT_UNKNOWN) &&
					*((VARTYPE*)V_BYREF(this)) == VT_URI) {
			delete (UriVariant*)V_BYREF(this);
		} else if (	vt == (VT_BYREF | VT_UNKNOWN) &&
					*((VARTYPE*)V_BYREF(this)) == VT_CLSID) {
			delete (CLSIDVariant*)V_BYREF(this);
		} else if (	vt == (VT_BYREF | VT_UNKNOWN) &&
					*((VARTYPE*)V_BYREF(this)) == VT_FILETIME) {
			delete (FileTimeVariant*)V_BYREF(this);
		} else if ( vt == (VT_BYREF | VT_UNKNOWN) &&
					*((VARTYPE*)V_BYREF(this)) == VT_EMPTY) {
			//circular reference breaker
		}
		if (S_OK != (hr = VariantClear(this))) {
			if (hr == DISP_E_ARRAYISLOCKED) {
				//allow and ignore for recursive
			} else
				throw _com_error(hr);
		}
	}
};

template <class T>
static bool AtlArrayToVariant(_tavariant_t & varData, CAtlArray<T> & Array)
{
	HRESULT hr;
	DWORD dwCount;
	CComSafeArray<VARIANT> SafeArray;
	if (S_OK != (hr = SafeArray.Create(Array.GetCount()))) {
		return false;
	}
	for (dwCount = 0; dwCount < Array.GetCount(); dwCount++) {
		if (S_OK != (hr = SafeArray.SetAt(	dwCount,
											_variant_t(_bstr_t(Array.GetAt(dwCount))).Detach(),
											FALSE))) {
			//do not leave empty or hole in SAFEARRAY since behavior would be undefined
			if ((dwCount == 0) || (S_OK != (hr = SafeArray.Resize(dwCount)))) {
				return false;
			} else {
				break;
			}
		}
	}
	varData.vt = VT_ARRAY | VT_VARIANT;
	V_ARRAY(&varData) = SafeArray.Detach();
	return true;
}

class JavaScriptSerializer
{
private:
	enum SerializationFormat
	{
		JavaScript = 1,//construct new date in javascript and no length check
		JSON = 0 //different date format and length check
	};
public:
	JavaScriptSerializer()
	{
		_recursionLimit = c_DefaultRecursionLimit;
		_maxJsonLength = c_DefaultMaxJsonLength;
	}
	
	CStringA Serialize(const _tavariant_t & obj)
	{
		return Serialize(obj, JSON);
	}

	void Serialize(const _tavariant_t & obj, CStringA & output)
	{
		Serialize(obj, output, JSON);
	}

	void Serialize(	const _tavariant_t & obj, CStringA & output,
					const SerializationFormat serializationFormat)
	{
		CAutoPtr< CRBMap<void*, void*> > objectsInUse;
		SerializeValue(	obj, output, 0,
						(CRBMap<void*, void*>* &)objectsInUse,
						serializationFormat);
		if ((serializationFormat == JSON) && 
			(output.GetLength() > _maxJsonLength)) {
			throw new CTANotSupportedException(JSON_MaxJsonLengthExceeded);
		}
	}
	static CStringA SerializeInternal(const _tavariant_t & o)
	{
		CStringA StringA;
		CAutoPtr<JavaScriptSerializer> Serializer(new JavaScriptSerializer());
		Serializer->Serialize(o, StringA);
		return StringA;
	}

	int get_MaxJsonLength()
	{
		return _maxJsonLength;
	}
	void set_MaxJsonLength(int value)
	{
		if (value < 1) {
			throw new CTAInvalidArgException(JSON_InvalidMaxJsonLength);
		}
		_maxJsonLength = value;
	}

	int get_RecursionLimit()
	{
		return _recursionLimit;
	}
	void set_RecursionLimit(int value)
	{
		if (value < 1)
		{
			throw new CTAInvalidArgException(JSON_InvalidRecursionLimit);
		}
		_recursionLimit = value;
	}

private:
	CStringA Serialize(	const _tavariant_t & obj,
						const SerializationFormat serializationFormat)
	{
		CStringA output;
		Serialize(obj, output, serializationFormat);
		return output;
	}
	static void SerializeBoolean(bool o, CStringA & sb)
	{
		if (o) {
			sb.Append("true");
		} else {
			sb.Append("false");
		}
	}
	static void SerializeDateTime(	CFileTime & Time, CStringA & sb,
									const SerializationFormat serializationFormat)
	{
		if (serializationFormat == JSON) {
			sb.Append("\"\\/Date(");
			sb.AppendFormat("%lld", (((Time.GetTime() - DatetimeMinTimeTicks) / 10000)));
			sb.Append(")\\/\"");
		} else {
			sb.Append("new Date(");
			sb.AppendFormat("%lld", (((Time.GetTime() - DatetimeMinTimeTicks) / 10000)));
			sb.Append(")");
		}
	}

	void SerializeDictionary(	CAtlMap<CStringW, VARIANT, 
										CStringElementTraitsI<CStringW> > & Dictionary, 
								CStringA & sb, int depth,
								CRBMap<void*, void*>* & objectsInUse,
								const SerializationFormat serializationFormat)
	{
		sb.AppendChar('{');
		bool flag = true;
		POSITION pos = Dictionary.GetStartPosition();
		CAtlMap<CStringW, VARIANT, CStringElementTraitsI<CStringW> >::CPair* pPair;
		while (pos && (pPair = Dictionary.GetNext(pos))) {
			_tavariant_t* varSerialize = new _tavariant_t();
			CAutoPtr<void> CleanupPtr(varSerialize);
			varSerialize->Attach(VARIANT(pPair->m_value));
			if (!flag) {
				sb.AppendChar(',');
			}
			SerializeDictionaryKeyValue(pPair->m_key, *varSerialize,
										sb, depth, objectsInUse, serializationFormat);
			flag = false;
		}
		sb.AppendChar('}');
	}

	void SerializeDictionaryKeyValue(	const CStringW & key, 
										const _tavariant_t &  value,
										CStringA & sb,
										const int depth,
										CRBMap<void*, void*>* & objectsInUse,
										const SerializationFormat serializationFormat)
	{
		SerializeString(((CStringW)key).GetBuffer(), sb);
		sb.AppendChar(':');
		SerializeValue(value, sb, depth, objectsInUse, serializationFormat);
	}

	void SerializeEnumerable(	SAFEARRAY* pSafeArray, CStringA & sb, int depth,
								CRBMap<void*, void*>* & objectsInUse, 
								const SerializationFormat serializationFormat)
	{
		CComSafeArray<VARIANT>* SafeArray = new CComSafeArray<VARIANT>();
		 //stop destructor from being called since it will use wrong variant clear
		CAutoPtr<void> ArrayCleanup(SafeArray);
		sb.AppendChar('[');
		bool flag = true;
		if ((S_OK != SafeArray->Attach(pSafeArray)) || 
			(S_OK != SafeArrayUnlock(pSafeArray)) ||
			(SafeArray->GetDimensions() != 1)) {
			throw new CTAInvalidArgException(JSON_InvalidObject);
		}
		long lCount;
		long lUpperBound = SafeArray->GetUpperBound();
		for (lCount = SafeArray->GetLowerBound(); lCount <= lUpperBound; lCount++) {
			_tavariant_t* varSerialize = new _tavariant_t();
			CAutoPtr<void> CleanupPtr(varSerialize);
			varSerialize->Attach(VARIANT(SafeArray->GetAt(lCount)));
			if (!flag) {
				sb.AppendChar(',');
			}
			SerializeValue(*varSerialize, sb, depth,
							objectsInUse, serializationFormat);
			flag = false;
		}
		sb.AppendChar(']');
	}

	static void SerializeGuid(const CLSID* pclsid, CStringA & sb)
	{
		LPOLESTR lpOleStr;
		HRESULT hRes = StringFromCLSID(*pclsid, &lpOleStr);
		if (hRes != S_OK) {
			throw new CTAInvalidArgException(JSON_InvalidObject);
		}
		sb.AppendChar('\"');
		sb.Append(CStringA(lpOleStr));
		sb.AppendChar('\"');
		CoTaskMemFree(lpOleStr);
	}

	static void SerializeString(LPCWSTR input, CStringA & sb)
	{
		sb.AppendChar('\"');
		sb.Append(JavaScriptString::QuoteString(input));
		sb.AppendChar('\"');
	}

	static void SerializeUri(const CUrl & uri, CStringA & sb)
	{
		DWORD dwLength;
		CString String;
		sb.AppendChar('\"');
		dwLength = uri.GetUrlLength() + 1;
		//create the url with escaping
		if (!uri.CreateUrl(String.GetBuffer(dwLength), &dwLength, ATL_URL_ESCAPE)) {
			throw new CTAInvalidArgException(JSON_InvalidObject);
		}
		String.ReleaseBuffer(dwLength - 1);
		sb.Append(CStringA(String));
		sb.AppendChar('\"');
	}

	void SerializeValue(const _tavariant_t & o, CStringA & sb, int depth,
						CRBMap<void*, void*>* & objectsInUse,
						const SerializationFormat serializationFormat)
	{
		if (++depth > _recursionLimit) {
			throw new CTAInvalidArgException(JSON_DepthLimitExceeded);
		}
		SerializeValueInternal(o, sb, depth, objectsInUse, serializationFormat);
	}

	void SerializeValueInternal(const _tavariant_t & o, CStringA & sb, int depth,
								CRBMap<void*, void*>* & objectsInUse,
								const SerializationFormat serializationFormat)
	{
		if (o.vt == VT_EMPTY) {
			throw new CTAInvalidArgException(JSON_InvalidObject);
		}
		if (o.vt == VT_NULL) {
			sb.Append("null");
		} else if (o.vt == VT_BSTR) {
			SerializeString(V_BSTR(&o), sb);
		} else if (o.vt == VT_UI1) {
			if (V_UI1(&o) == '\0') {
				sb.Append("null");
			} else {
				SerializeString(CStringW((char)V_UI1(&o)), sb);
			}
		} else if (o.vt == VT_I1) {
			sb.AppendFormat("%d", (int)V_I1(&o));
		} else if (o.vt == VT_UI1) {
			sb.AppendFormat("%u", (unsigned int)V_UI1(&o));
		} else if (o.vt == VT_I2) {
			sb.AppendFormat("%hd", V_I2(&o));
		} else if (o.vt == VT_UI2) {
			sb.AppendFormat("%hu", V_UI2(&o));
		} else if (o.vt == VT_I4) {
			sb.AppendFormat("%ld", V_I4(&o));
		} else if (o.vt == VT_UI4) {
			sb.AppendFormat("%lu", V_UI4(&o));
		} else if (o.vt == VT_I8) {
			sb.AppendFormat("%lld", V_I8(&o));
		} else if (o.vt == VT_UI8) {
			sb.AppendFormat("%llu", V_UI8(&o));
		} else if (o.vt == VT_BOOL) {
			SerializeBoolean(V_BOOL(&o) != 0, sb);
		} else if (	o.vt == (VT_BYREF | VT_UNKNOWN) && 
					*((VARTYPE*)V_BYREF(&o)) == VT_FILETIME) {
					SerializeDateTime(	CFileTime((	(_tavariant_t::FileTimeVariant*)
													V_BYREF(&o))->FileTime),
										sb, serializationFormat);
		} else if (	o.vt == (VT_BYREF | VT_UNKNOWN) &&
					*((VARTYPE*)V_BYREF(&o)) == VT_CLSID) {
			SerializeGuid(&((_tavariant_t::CLSIDVariant*)V_BYREF(&o))->clsid, sb);
		} else if (	o.vt == (VT_BYREF | VT_UNKNOWN) &&
					*((VARTYPE*)V_BYREF(&o)) == VT_URI) {
			SerializeUri(((_tavariant_t::UriVariant*)V_BYREF(&o))->Uri, sb);
		} else if (o.vt == VT_R8) {
			//TODO: verify following 2 precisions are correct
			sb.AppendFormat("%.15g", V_R8(&o));
		} else if (o.vt == VT_R4) {
			sb.AppendFormat("%.7g", (double)V_R4(&o)); //treat as double for formating
		} else if (o.vt == VT_DECIMAL) {
			_bstr_t bstrOut;
			if (S_OK != VarBstrFromDec(	&V_DECIMAL((_tavariant_t*)&o), LOCALE_SYSTEM_DEFAULT, 
										0, bstrOut.GetAddress())) {
				throw new CTAInvalidArgException(JSON_InvalidObject);
			}
			//encoding needed for multilingual?
			sb.Append(JavaScriptString::QuoteString(bstrOut));
		} else {
			//track objects to detect circular references which are not allowed
			//    depth is also safety here since it would cause infinite recursion
			if ((o.vt == (VT_BYREF | VT_UNKNOWN) &&
				 *((VARTYPE*)V_BYREF(&o)) == VT_DICTIONARY) ||
				(o.vt == (VT_ARRAY | VT_VARIANT))) {
				SerializeRecursiveValue(o, sb, depth, objectsInUse, serializationFormat);
			} else {
				throw new CTAInvalidArgException(JSON_InvalidObject);
			}
		}
	}
	class CRBMapRemover
	{
	public:
		CRBMapRemover(CRBMap<void*, void*>* objectsInUse, const _tavariant_t* o) : _o(o)
		{ Assign(objectsInUse); }
		~CRBMapRemover()
		{
			if (_objectsInUse != NULL) {
				_objectsInUse->RemoveKey(V_BYREF(_o));
			}
		}
		void Assign(CRBMap<void*, void*>* objectsInUse)
		{
			_objectsInUse = objectsInUse;
		}
	protected:
		CRBMap<void*, void*>* _objectsInUse;
		const _tavariant_t* _o;
	};
	void SerializeRecursiveValue(	const _tavariant_t & o, CStringA & sb, int depth,
									CRBMap<void*, void*>* & objectsInUse,
									const SerializationFormat serializationFormat)
	{
		CRBMapRemover MapRemover(objectsInUse, &o);
		if (objectsInUse == NULL) {
			MapRemover.Assign(objectsInUse = new CRBMap<void*, void*>());
		} else if (objectsInUse->Lookup(V_BYREF(&o))) {
			throw new CTANotSupportedException(JSON_CircularReference);
		}
		objectsInUse->SetAt(V_BYREF(&o), V_BYREF(&o));
		if (o.vt == (VT_BYREF | VT_UNKNOWN) &&
			*((VARTYPE*)V_BYREF(&o)) == VT_DICTIONARY) {
			SerializeDictionary(((_tavariant_t::DictionaryVariant*)V_BYREF(&o))->Dictionary,
								sb, depth, objectsInUse, serializationFormat);
		} else if (o.vt == (VT_ARRAY | VT_VARIANT)) {
			SerializeEnumerable(V_ARRAY(&o), sb, depth,
								objectsInUse, serializationFormat);
		}
	}

	int _maxJsonLength;
	int _recursionLimit;
};

class JavaScriptObjectDeserializer
{
public:
	JavaScriptObjectDeserializer(const char* input, const int depthLimit)
	{
		_s = new JavaScriptString(input); //possible memory exception
		_depthLimit = depthLimit;
	}
	~JavaScriptObjectDeserializer()
	{
		delete _s;
	}
	static void BasicDeserialize(	_tavariant_t & varDeserialize, const char* input,
									const int depthLimit)
	{
		CAutoPtr<JavaScriptObjectDeserializer>
			deserializer(new JavaScriptObjectDeserializer(input, depthLimit));
		deserializer->DeserializeInternal(varDeserialize, 0);
		char nextNonEmptyChar = deserializer->_s->GetNextNonEmptyChar();
		if (nextNonEmptyChar) {
			throw new CTAInvalidArgException(	CString(JSON_IllegalPrimitive _T(": ")) +
												CString(deserializer->_s->ToString()));
		}
	}
private:
	void AppendCharToBuilder(const char c, CStringW & sb)
	{
		 //possible memory exceptions
		if (((c == '\"') || (c == '\'')) || (c == '/')) {
			sb.AppendChar(c);
		} else if (c == 'b') {
			sb.AppendChar('\b');
		} else if (c == 'f') {
			sb.AppendChar('\f');
		} else if (c == 'n') {
			sb.AppendChar('\n');
		} else if (c == 'r') {
			sb.AppendChar('\r');
		} else if (c == 't') {
			sb.AppendChar('\t');
		} else {
			if (c != 'u') {
				throw new CTAInvalidArgException(_s->GetDebugString(JSON_BadEscape));
			}
			WCHAR wChar;
			int iReadLen;
			if (sscanf(_s->MoveNext(4), "%hx%n", &wChar, &iReadLen) >= 1 &&
				(iReadLen == 4)) {
				sb.AppendChar(wChar); //possible memory exception
			} else {
				throw new CTAInvalidArgException(_s->GetDebugString(JSON_BadEscape));
			}
		}
	}
	char CheckQuoteChar(char c)
	{
		if (c == '\'') {
			return c;
		}
		if (c != '\"') {
			//impossible exception to hit because of lookahead
			throw new CTAInvalidArgException(_s->GetDebugString(JSON_StringNotQuoted));
		}
		return '\"';
	}

	void DeserializeDictionary(_tavariant_t & varDictionary, int depth)
	{
		char nextNonEmptyChar;
		_tavariant_t::DictionaryVariant* dictionary = new _tavariant_t::DictionaryVariant();
		V_VT(&varDictionary) = (VT_BYREF | VT_UNKNOWN);
		dictionary->vt = VT_DICTIONARY;
		V_BYREF(&varDictionary) = dictionary;
		if (_s->MoveNext() != '{') {
			//an impossible exception because of the lookahead
			throw new CTAInvalidArgException(_s->GetDebugString(JSON_ExpectedOpenBrace));
		}
		while (nextNonEmptyChar = _s->GetNextNonEmptyChar()) {
			_s->MovePrev();
			if (nextNonEmptyChar == ':') {
				throw new CTAInvalidArgException(_s->GetDebugString(JSON_InvalidMemberName));
			}
			if (nextNonEmptyChar != '}') {
				CStringW str;
				_tavariant_t varDeserialize;
				str = DeserializeMemberName();
				if (str.IsEmpty()) {
					throw new CTAInvalidArgException(_s->GetDebugString(JSON_InvalidMemberName));
				}
				if (_s->GetNextNonEmptyChar() != ':') {
					throw new CTAInvalidArgException(_s->GetDebugString(JSON_InvalidObject));
				}
				DeserializeInternal(varDeserialize, depth);
				dictionary->Dictionary.SetAt(str, varDeserialize.Detach());
			} else {
				nextNonEmptyChar = _s->GetNextNonEmptyChar();
				break;
			}
			nextNonEmptyChar = _s->GetNextNonEmptyChar();
			if (nextNonEmptyChar != '}') {
				if (nextNonEmptyChar != ',') {
					throw new CTAInvalidArgException(_s->GetDebugString(JSON_InvalidObject));
				}
			} else
				break;
		}
		if (nextNonEmptyChar != '}') {
			throw new CTAInvalidArgException(_s->GetDebugString(JSON_InvalidObject));
		}
	}
	void DeserializeInternal(_tavariant_t & varDeserialize, int depth)
	{
		_tavariant_t varTemp;
		if (++depth > _depthLimit) {
			throw new CTAInvalidArgException(_s->GetDebugString(JSON_DepthLimitExceeded));			
		}
		char nextNonEmptyChar = _s->GetNextNonEmptyChar();
		_s->MovePrev();
		if (IsNextElementDateTime()) {
			DeserializeStringIntoDateTime(varTemp);
		} else if (JavaScriptObjectDeserializer::IsNextElementObject(nextNonEmptyChar)) {
			DeserializeDictionary(varTemp, depth);
		} else if (JavaScriptObjectDeserializer::IsNextElementArray(nextNonEmptyChar)) {
			DeserializeList(varTemp, depth);
		} else if (JavaScriptObjectDeserializer::IsNextElementString(nextNonEmptyChar)) {
			varTemp.Attach(_variant_t(_bstr_t(DeserializeString())).Detach());
		} else {
			DeserializePrimitiveObject(varTemp);
		}
		varDeserialize.Attach(varTemp.Detach());
	}

	void DeserializeList(_tavariant_t & varList, int depth)
	{
		char nextNonEmptyChar;
		CComSafeArray<VARIANT> sa;
		if (S_OK != sa.Create()) {
			throw;
		}
		if (_s->MoveNext() != '[') {
			//impossible to hit because of lookahead
			throw new CTAInvalidArgException(_s->GetDebugString(JSON_InvalidArrayStart));
		}
		bool flag = false;
		while (	(nextNonEmptyChar = _s->GetNextNonEmptyChar()) &&
				(nextNonEmptyChar != ']')) {
			_tavariant_t varDeserialize;
			_s->MovePrev();
			DeserializeInternal(varDeserialize, depth);
			if ((S_OK == sa.Resize(sa.GetUpperBound() + 1 + 1)) &&
				(S_OK == sa.SetAt(sa.GetUpperBound(), varDeserialize.Detach(), FALSE))) {
				flag = false;
				nextNonEmptyChar = _s->GetNextNonEmptyChar();
				if (nextNonEmptyChar && (nextNonEmptyChar != ']')) {
					flag = true;
					if (nextNonEmptyChar != ',') {
						_tavariant_t varCleanup;
						V_VT(&varCleanup) = VT_ARRAY | VT_VARIANT;
						V_ARRAY(&varCleanup) = (LPSAFEARRAY)sa.Detach();
						throw new CTAInvalidArgException
							(_s->GetDebugString(JSON_InvalidArrayExpectComma));
					}
				} else {
					break;
				}
			} else {
				//internal safearray error
				break;
			}
		}
		if (flag) {
			_tavariant_t varCleanup;
			V_VT(&varCleanup) = VT_ARRAY | VT_VARIANT;
			V_ARRAY(&varCleanup) = (LPSAFEARRAY)sa.Detach();
			throw new CTAInvalidArgException(_s->GetDebugString(JSON_InvalidArrayExtraComma));
		}
		if (nextNonEmptyChar != ']') {
			_tavariant_t varCleanup;
			V_VT(&varCleanup) = VT_ARRAY | VT_VARIANT;
			V_ARRAY(&varCleanup) = (LPSAFEARRAY)sa.Detach();
			throw new CTAInvalidArgException(_s->GetDebugString(JSON_InvalidArrayEnd));
		}
		V_VT(&varList) = VT_ARRAY | VT_VARIANT;
		V_ARRAY(&varList) = (LPSAFEARRAY)sa.Detach();
	}
	CStringW DeserializeMemberName()
	{
		char nextNonEmptyChar = _s->GetNextNonEmptyChar();
		_s->MovePrev();
		if (JavaScriptObjectDeserializer::IsNextElementString(nextNonEmptyChar)) {
			return DeserializeString();
		}
		return DeserializePrimitiveToken();
	}
	void DeserializePrimitiveObject(_tavariant_t & varPrimitive)
	{
		double dblTest;
		CStringW s = DeserializePrimitiveToken();
		if (s.Compare(L"null") == 0) {
			V_VT(&varPrimitive) = VT_NULL;
		} else if (s.Compare(L"true") == 0) {
			varPrimitive.Attach(_variant_t(true).Detach());
		} else if (s.Compare(L"false") == 0) {
			varPrimitive.Attach(_variant_t(false).Detach());
		} else {
			bool flag = (s.Find('.') >= 0);
			if (s.ReverseFind('e') < 0 && s.ReverseFind('E') < 0) {
				DECIMAL decTest;
				if (!flag) {
					union {
						long long llTest;
						unsigned long long ullTest;
					};
					//no support for types I1, UI1, I2, UI2, UI4 since they would
					//   make use of the resultant variant more difficult for the caller
					//caller must handle I4 by checking I8 return value
					if (S_OK == VarI8FromStr(s.GetBuffer(), LOCALE_SYSTEM_DEFAULT, 0, &llTest)) {
						varPrimitive.Attach(((llTest & 0xFFFFFFFF80000000) == 0xFFFFFFFF80000000) ||
											((llTest & 0xFFFFFFFF80000000) == 0) ? 
											_variant_t((long)(llTest & 0xFFFFFFFF)).Detach() :
											_variant_t(llTest).Detach());
						return;
					} else if (S_OK == VarUI8FromStr(	s.GetBuffer(), LOCALE_SYSTEM_DEFAULT,
														0, &ullTest)) {
						varPrimitive.Attach(_variant_t(ullTest).Detach());
						return;
					}
				}
				if (S_OK == VarDecFromStr(s.GetBuffer(), LOCALE_SYSTEM_DEFAULT, 0, &decTest)) {
					varPrimitive.Attach(_variant_t(decTest).Detach());
					return;
				}
			}
			//no support for types float or R4 since they would make use of the
			//  resultant variant more difficult for the caller
			if (S_OK != VarR8FromStr(s.GetBuffer(), LOCALE_SYSTEM_DEFAULT, 0, &dblTest)) {
				throw new CTAInvalidArgException(CString(JSON_IllegalPrimitive _T(": ")) + s);
			}
			varPrimitive.Attach(_variant_t(dblTest).Detach());
		}
	}

	CStringW DeserializePrimitiveToken()
	{
		CStringW builder;
		char c;
		while (c = _s->MoveNext()) {
			if ((isalnum(c) ||
				 (c == '.')) ||
				(((c == '-') || (c == '_')) ||
				 (c == '+'))) {
				builder.AppendChar(c); //possible memory exception
			} else {
				_s->MovePrev();
				break;
			}
		}
		return builder;
	}

	CStringW DeserializeString()
	{
		CStringW sb;
		bool flag = false;
		char c = _s->MoveNext();
		char ch = CheckQuoteChar(c);
		while (c = _s->MoveNext()) {
			if (c == '\\') {
				if (flag) {
					sb.AppendChar('\\'); //possible memory exception
					flag = false;
				} else {
					flag = true;
				}
			} else if (flag) {
				AppendCharToBuilder(c, sb); //possible memory exception
				flag = false;
			} else {
				if (c == ch) {
					return sb;
				}
				sb.AppendChar(c); //possible memory exception
			}
		}
		//throw new CTAInvalidArgException(_s->GetDebugString(JSON_StringNotQuoted));
		//incomplete and invalid string - return and be graceful about this error
		//	that way if a single string was the input yet missing a quote it is grabbed
		//	but if the input is part of a dictionary/list
		//  let the object parser throw the exception
		return sb;
	}
	void DeserializeStringIntoDateTime(_tavariant_t & varDateTime)
	{
		__int64 num;
		_tavariant_t::FileTimeVariant* varFileTime;
		std::tr1::regex rx(	"^\"\\\\/Date\\((-?[0-9]+)(?:[a-zA-Z]|(?:\\+|-)[0-9]{4})?\\)\\\\/\"", 
							(std::tr1::regex::flag_type)(std::tr1::regex_constants::ECMAScript));
		std::tr1::cmatch mr; //result capture group 1 is what is wanted
		if (std::tr1::regex_search(_s->ToString().GetBuffer(), mr, rx) &&
			(sscanf(mr[1].str().c_str(), "%lld", &num) == 1)) {
			_s->MoveNext(mr.length());
			varFileTime = new _tavariant_t::FileTimeVariant(); //possible memory exception
			varFileTime->FileTime = CFileTime((num * 10000) + DatetimeMinTimeTicks);
			varFileTime->vt = VT_FILETIME;
			varDateTime.vt = (VT_BYREF | VT_UNKNOWN);
			V_BYREF(&varDateTime) = varFileTime;
		} else {
			varDateTime.Attach(_variant_t(_bstr_t(DeserializeString())).Detach());
		}
	}
	static bool IsNextElementArray(char c)
	{
		return (c == '[');
	}
	bool IsNextElementDateTime()
	{
		const char* DateTimePrefix = "\"\\/Date(";
		CStringA a = _s->MoveNext(DateTimePrefixLength); //possible memory exception
		//empty is returned if not enough characters
		if (!a.IsEmpty()) {
			_s->MovePrev(DateTimePrefixLength);
			return a.Compare(DateTimePrefix) == 0;
		}
		return false;
	}
	static bool IsNextElementObject(char c)
	{
		return (c == '{');
	}
	static bool IsNextElementString(char c)
	{
		return ((c == '\"') || (c == '\''));
	}

	JavaScriptString* _s;
	int _depthLimit;
	static const int DateTimePrefixLength = 8;
};

//TODO: test floating point precision, URLs and GUIDs
#ifdef TEST
static bool JavaScriptSerializationPerformTest()
{
#define DeserializationExceptionTestDepth(	szTest, szGoodTest,\
											szResult, iDefaultRecursionLimit,\
											iGoodDefaultRecursionLimit) \
	try {\
		JavaScriptObjectDeserializer::BasicDeserialize(varData, szTest, iDefaultRecursionLimit);\
		bRet = false;\
	} catch (CTASimpleException* SimpleException) {\
		CAutoPtr<CTASimpleException> CleanupExceptionPtr(SimpleException);\
		bRet &= (SimpleException->GetString().Left\
			(sizeof(szResult) / sizeof(TCHAR) - 1).Compare(szResult) == 0);\
	}\
	varData.ExtendedVariantClear();\
	JavaScriptObjectDeserializer::BasicDeserialize(	varData, szGoodTest,\
													iGoodDefaultRecursionLimit);\
	varData.ExtendedVariantClear();

#define DeserializationExceptionTest(szTest, szGoodTest, szResult) \
	DeserializationExceptionTestDepth(	szTest, szGoodTest, szResult, \
										c_DefaultRecursionLimit, c_DefaultRecursionLimit)

	_tavariant_t varData;
	CStringA StringA;
	CStringA TestStringA;
	JavaScriptSerializer* Serializer;
	CAutoPtr<JavaScriptSerializer> SerializerCleanup;
	bool bRet = true;

	try {
		DeserializationExceptionTestDepth(	"{\"nest\":[{\"\\u3939\":3},1.2839]}",
											"{\"nest\":[{\"\\u3939\":3},1.2839]}",
											JSON_DepthLimitExceeded, 3, 4);
		DeserializationExceptionTest(	
			"[\"\\uDEFG\",8,9,[3,4,[8,[12,[13,{3:[[3,2],2],4:{9:{9:[8,4]}}}],14]]]]",
			"[\"\\uDEFF\",8,9,[3,4,[8,[12,[13,{3:[[3,2],2],4:{9:{9:[8,4]}}}],14]]]]",
			JSON_BadEscape);
		DeserializationExceptionTest(	"{\"nest\":[{\"\\uDEFG\":3},1.2839]}",
										"{\"nest\":[{\"\\uDEFF\":3},1.2839]}",
										JSON_BadEscape);
		DeserializationExceptionTest(	"{\"nest\":[{\"\\q3939\":3},1.2839]}",
										"{\"nest\":[{\"\\u3939\":3},1.2839]}",
										JSON_BadEscape);
		DeserializationExceptionTest(	"{\"nest\":[{\"\\u3939:3},1.2839]}",
										"{\"nest\":[{\"\\u3939\":3},1.2839]}",
										JSON_InvalidObject);
		DeserializationExceptionTest(	"{\"nest\":[{\"3939\":3e3000},1.2839]}",
										"{\"nest\":[{\"3939\":3e300},1.2839]}",
										JSON_IllegalPrimitive);
		DeserializationExceptionTest(	"{\"nest\":[{\"3939\":3e30},1.2839]},1",
										"{\"nest\":[{\"3939\":3e30},1.2839]}",
										JSON_IllegalPrimitive);
		DeserializationExceptionTest(	"{\"nest\":[{:3},1.2839]}",
										"{\"nest\":[{\"member\":3},1.2839]}",
										JSON_InvalidMemberName);
		DeserializationExceptionTest(	"{\"nest\":[{\"\":3},1.2839]}",
										"{\"nest\":[{55.3:3},1.2839]}",
										JSON_InvalidMemberName);
		DeserializationExceptionTest(	"{\"nest\":[{\"test\";3},1.2839]}",
										"{\"nest\":[{\"test\":3},1.2839]}",
										JSON_InvalidObject);
		DeserializationExceptionTest(	"{\"nest\":[{\"test\":3;\"next\":4},1.2839]}",
										"{\"nest\":[{\"test\":3,\"next\":4},1.2839]}",
										JSON_InvalidObject);
		DeserializationExceptionTest(	"{\"nest\":[{\"test\":3,\"next\":4,1.2839]}",
										"{\"nest\":[{\"test\":3,\"next\":4},1.2839]}",
										JSON_InvalidObject);
		DeserializationExceptionTest(	"{\"nest\":[{\"test\":3,\"next\":4};1.2839]}",
										"{\"nest\":[{\"test\":3,\"next\":4},1.2839]}",
										JSON_InvalidArrayExpectComma);
		DeserializationExceptionTest(	"{\"nest\":[{\"test\":3,\"next\":4},1.2839,]}",
										"{\"nest\":[{\"test\":3,\"next\":4},1.2839]}",
										JSON_InvalidArrayExtraComma);
		DeserializationExceptionTest(	"{\"nest\":[{\"test\":3,\"next\":4},1.2839",
										"{\"nest\":[{\"test\":3,\"next\":4},1.2839]}",
										JSON_InvalidArrayEnd);
		SerializerCleanup.Attach(Serializer = new JavaScriptSerializer());
		Serializer->set_MaxJsonLength(14);
		Serializer->set_RecursionLimit(c_DefaultRecursionLimit);
		try {
			Serializer->set_MaxJsonLength(0);
			bRet = false;
		} catch (CTASimpleException* SimpleException) {
			CAutoPtr<CTASimpleException> CleanupExceptionPtr(SimpleException);
			bRet &= (SimpleException->GetString().Left
				(sizeof(JSON_InvalidMaxJsonLength) / sizeof(TCHAR) - 1).Compare
				(JSON_InvalidMaxJsonLength) == 0);
		}
		try {
			Serializer->set_RecursionLimit(0);
			bRet = false;
		} catch (CTASimpleException* SimpleException) {
			CAutoPtr<CTASimpleException> CleanupExceptionPtr(SimpleException);
			bRet &= (SimpleException->GetString().Left
				(sizeof(JSON_InvalidRecursionLimit) / sizeof(TCHAR) - 1).Compare
				(JSON_InvalidRecursionLimit) == 0);
		}
		Serializer->Serialize(_tavariant_t(_variant_t(_bstr_t(L"overflowtest")))); 
		try {
			Serializer->Serialize(_tavariant_t(_variant_t(_bstr_t(L"overflowtest!")))); 
			bRet = false;
		} catch (CTASimpleException* SimpleException) {
			CAutoPtr<CTASimpleException> CleanupExceptionPtr(SimpleException);
			bRet &= (SimpleException->GetString().Left
				(sizeof(JSON_MaxJsonLengthExceeded) / sizeof(TCHAR) - 1).Compare
				(JSON_MaxJsonLengthExceeded) == 0);
		}
		try {
			Serializer->Serialize(_tavariant_t()); 
			bRet = false;
		} catch (CTASimpleException* SimpleException) {
			CAutoPtr<CTASimpleException> CleanupExceptionPtr(SimpleException);
			bRet &= (SimpleException->GetString().Left
				(sizeof(JSON_InvalidObject) / sizeof(TCHAR) - 1).Compare
				(JSON_InvalidObject) == 0);
		}
		try {
			CComSafeArray<VARIANT> SafeArray;
			bRet &= (S_OK == SafeArray.Create(1));
			V_VT(&varData) = VT_ARRAY | VT_VARIANT;
			V_BYREF(&varData) = SafeArray;
			SafeArray.SetAt(0, varData, FALSE);
			SafeArray.Detach();
			Serializer->Serialize(varData); 
			bRet = false;
		} catch (CTASimpleException* SimpleException) {
			CAutoPtr<CTASimpleException> CleanupExceptionPtr(SimpleException);
			bRet &= (SimpleException->GetString().Left
				(sizeof(JSON_CircularReference) / sizeof(TCHAR) - 1).Compare
				(JSON_CircularReference) == 0);
		}
		varData.ExtendedVariantClear();
		try {
			V_VT(&varData) = VT_BYREF | VT_UNKNOWN;
			V_BYREF(&varData) = new _tavariant_t::DictionaryVariant();
			((_tavariant_t::DictionaryVariant*)V_BYREF(&varData))->vt = VT_DICTIONARY;
			((_tavariant_t::DictionaryVariant*)V_BYREF(&varData))->Dictionary.SetAt(L"test",
																					varData);
			Serializer->Serialize(varData); 
			bRet = false;
		} catch (CTASimpleException* SimpleException) {
			CAutoPtr<CTASimpleException> CleanupExceptionPtr(SimpleException);
			bRet &= (SimpleException->GetString().Left
				(sizeof(JSON_CircularReference) / sizeof(TCHAR) - 1).Compare
				(JSON_CircularReference) == 0);
		}
		varData.ExtendedVariantClear();
		SerializerCleanup.Free();

		StringA.Format(
			"{\"list\":[\"test\",null,true,false,11068046444225730969,-7378697629483820647,"
			"\"again\",3,293848833,-1,2147483648,-2147483648,2147483647,8.2,1.2e+128,"
			"[\"in\",\"a\",\"l\",{\"name\":{\"firstinner\":{\"a\":3,\"b\":4,\"c\":5},"
			"\"secondinner\":{\"a\":6,\"b\":7,\"c\":8,\"d\":9,\"e\":10}},"
			" \"tricky\":\"hahacool!<>&\\\\\\\"\\\'\"}],{\"atest\":5,\"test\":\"another\"},"
			"\"last\"],\"date\":\"\\/Date(%lld)\\/\",\"invaldate\":\"\\/Date(notadate)\\/\"}",
					((CFileTime(0).GetTime() - DatetimeMinTimeTicks) / 10000));
		_variant_t varTest;
		varTest = (unsigned char)3;
		TestStringA = JavaScriptSerializer::SerializeInternal(varTest);
		bRet &= (TestStringA.Compare("\"\\u0003\"") == 0);
		varTest.Clear();
		varTest = (char)-127;
		TestStringA = JavaScriptSerializer::SerializeInternal(varTest);
		bRet &= (TestStringA.Compare("-127") == 0);
		varTest.Clear();
		varTest = (unsigned short)33904;
		TestStringA = JavaScriptSerializer::SerializeInternal(varTest);
		bRet &= (TestStringA.Compare("33904") == 0);
		varTest.Clear();
		varTest = (short)-83;
		TestStringA = JavaScriptSerializer::SerializeInternal(varTest);
		bRet &= (TestStringA.Compare("-83") == 0);
		varTest.Clear();
		varTest = (unsigned long)2984292909;
		TestStringA = JavaScriptSerializer::SerializeInternal(varTest);
		bRet &= (TestStringA.Compare("2984292909") == 0);
		varTest.Clear();
		varTest = (float)2984292909.29e28;
		TestStringA = JavaScriptSerializer::SerializeInternal(varTest);
		bRet &= (TestStringA.Compare("2.984293e+037") == 0);

		JavaScriptObjectDeserializer::BasicDeserialize
			(varData, StringA, c_DefaultRecursionLimit);
		StringA = JavaScriptSerializer::SerializeInternal(varData);
		varData.ExtendedVariantClear();
		//normalize by running back through first time to establish dictionary order
		JavaScriptObjectDeserializer::BasicDeserialize
			(varData, StringA, c_DefaultRecursionLimit);
		TestStringA = JavaScriptSerializer::SerializeInternal(varData);
		bRet &= (StringA.Compare(TestStringA) == 0);
	} catch (...) {
		bRet = false; //unexpected exception
	}
	return bRet;
}
#endif