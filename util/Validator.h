#pragma once
#include <atlrx.h>
namespace UTIL
{
#define ATLS_REGEX_PARAM				L"^{['-9a-zA-ZÁ-ÿ ]*}$"
#define ATLS_REGEX_ANSI_PUNCTUATION		L"^{[\\'\\(\\)\\,\\-\\/\\. ]*}$"
//#define ATLS_REGEX_UNICODE_PUNCTUATION	L"^{[\\'\\(\\)\\,\\-\\.\\/\\_\\­\\‘\\’\\‚\\?\\/ ]*}$"
#define ATLS_REGEX_PARAM_FRENCH			L"^{['-9A-zÀ-ÿ ]*}$"
#define ATLS_REGEX_FRENCH_UNICODE		L"^{['-9A-Z^-zÀ-??-? ]*}$"
#define ATLS_REGEX_DBESCAPE	L"{BEGIN}+|{CAST}+|{DROP}+|{INSERT}+|{SELECT}+|{SET}+|{UPDATE}+"

	wchar_t * provinceNames[] = { L"Alberta",L"British Columbia",L"Manitoba",L"Newfoundland",L"New Brunswick",L"Northwest Territories",L"Nova Scotia",L"Nunavut",L"Ontario",L"Prince Edward Island",L"Quebec",L"Saskatchewan",L"Yukon" };
	wchar_t * g_provinceValues[] = { L"AB",L"BC",L"MB",L"NL",L"NB",L"NT",L"NS",L"NU",L"ON",L"PE",L"QC",L"SK",L"YT" };
	typedef unsigned char VCUE_RPC_CHAR;
	//ensre there aren't any control characters or database escape sequences submitted
	inline bool ParamIsAcceptable(LPCWSTR strValue)
	{
		CAtlRegExp<CAtlRECharTraitsW> atlRegExp;
		CAtlREMatchContext<CAtlRECharTraitsW> atlRematchContext;
		if (atlRegExp.Parse(ATLS_REGEX_PARAM) != REPARSE_ERROR_OK)
			return false;
		if (0 != atlRegExp.Match(strValue, &atlRematchContext))
		{
			if (atlRegExp.Parse(ATLS_REGEX_DBESCAPE, 0) != REPARSE_ERROR_OK)
				return false;
			if (0 != atlRegExp.Match(strValue, &atlRematchContext))
			{
				return false;
			}
			else
			{
				return true;
			}
		}
		else
		{
			return false;
		}
	}
	// test for french characters...
	inline bool ParamIsFrench(LPCWSTR strValue)
	{
		CAtlRegExp<CAtlRECharTraitsW> atlRegExp;
		CAtlREMatchContext<CAtlRECharTraitsW> atlRematchContext;
		if (0 != atlRegExp.Parse(ATLS_REGEX_PARAM_FRENCH) != REPARSE_ERROR_OK)
			return false;
		if (0 != atlRegExp.Match(strValue, &atlRematchContext))
		{
			if (atlRegExp.Parse(ATLS_REGEX_DBESCAPE, 0) != REPARSE_ERROR_OK)
				return false;
			if (0 != atlRegExp.Match(strValue, &atlRematchContext))
			{
				return false;
			}
			else
			{
				return true;
			}
		}
		else
		{
			return false;
		}
	}
	//returns  validity of 2 character province
	inline bool ParamIsProvince(LPCWSTR wszValue)
	{
		unsigned int uiLen = wcslen(wszValue);
		if (2 == uiLen)
		{
			for (int i = 0; i < 12; i++)
			{
				if (0 == lstrcmpiW(wszValue, g_provinceValues[i]))
					return true;
			}
		}
		return false;
	}
	//looks up full province name
	inline LPCWSTR GetNamedProvince(const CAtlStringW& strValue)
	{
		CAtlStringW strProvince;
		for (int i=0; i < 12; i++)
		{
			if (0 == lstrcmpiW(strValue, g_provinceValues[i]))
			{
				strProvince = provinceNames[i];
			}
		}
		return ( LPCWSTR ) strProvince;
	}
	//given rowcount and rows per page, calculates number of pages
	inline int CalculatePages(DBCOUNTITEM dbRows = 1, DBCOUNTITEM dbPp = 25)
	{
		int iPages = 0;
		int iRows = 0;
		int iPerPage = 0;
		iRows = dbRows;
		iPerPage = dbPp;
		iPages = abs(iRows / iPerPage);//watchout for divide by zero
		return iPages;                 //guaranteed not to happen in this project
	}
};//namespace UTIL