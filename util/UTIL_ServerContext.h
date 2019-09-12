#pragma once
#include "UTIL_Time.h"
#include <atltime.h>

namespace UTIL
{
	inline bool GetServerVariable(IHttpServerContext* pServerContext, LPCSTR szVariable, CStringA &str) throw()
	{
		bool bSuccess = false;

		if (pServerContext)
		{
			DWORD dwSize = 0;
			pServerContext->GetServerVariable(szVariable, NULL, &dwSize);
			
			if (pServerContext->GetServerVariable(szVariable, str.GetBuffer(dwSize), &dwSize))
				bSuccess = true;

			if (dwSize > 0)
				--dwSize;
			str.ReleaseBuffer(dwSize);
		}

		return bSuccess;
	}

	inline bool ContextWriteClient(IHttpServerContext* pServerContext, LPCSTR pvBuffer, DWORD* pdwBytes) throw()
	{
		return pServerContext->WriteClient(const_cast<LPSTR>(pvBuffer), pdwBytes) ? true : false;
	}

	inline bool ContextWriteClient(IHttpServerContext* pServerContext, LPCSTR pvBuffer, DWORD dwBytes) throw()
	{
		return ContextWriteClient(pServerContext, pvBuffer, &dwBytes);
	}

	inline bool ClientHasRecentContent(IHttpServerContext* pServerContext, COleDateTime tCurrent = GmtTime(), COleDateTimeSpan tsSpan = COleDateTimeSpan(0, 0, 5, 0))
	{
		bool bClientHasRecentContent = false;

		CStringA strLastModified;
		if (GetServerVariable(pServerContext, "HTTP_IF_MODIFIED_SINCE", strLastModified))
		{
			SYSTEMTIME st = {0};
			if (ParseHttpDate(strLastModified, st))
			{
				COleDateTime tCached(st);
				if (tCurrent - tCached <= tsSpan)
					bClientHasRecentContent = true;
			}
		}

		return bClientHasRecentContent;
	}

	inline bool ClientHasRecentContent(IHttpServerContext* pServerContext, unsigned int nMinutes)
	{
		COleDateTime tCurrent((SYSTEMTIME)GmtTime());
		COleDateTimeSpan tsMinutes(0, 0, nMinutes, 0);
		return ClientHasRecentContent(pServerContext, tCurrent, tsMinutes);
	}

	// TODO - implement request functions in terms of server context?
	// Call this function to retrieve the physical folder of the current script
	// Returns true on success, false on failure
	inline bool GetPhysicalScriptFolder(IHttpServerContext* pServerContext, CStringA& strPhysicalFolder)
	{
		bool bSuccess = false;
		strPhysicalFolder = pServerContext->GetScriptPathTranslated();
		if (strPhysicalFolder.GetLength())
		{
			int nFound = strPhysicalFolder.ReverseFind('\\');
			if (nFound != -1)
				strPhysicalFolder = strPhysicalFolder.Left(nFound + 1);
			else
				strPhysicalFolder += '\\';

			// Success!
			bSuccess = true;
		}
		return bSuccess;
	}

	// Return physical script folder directly (or empty string on failure).
	inline CStringA GetPhysicalScriptFolder(IHttpServerContext* pServerContext)
	{
		CStringA strPhysicalScriptFolder;
		if (!GetPhysicalScriptFolder(pServerContext, strPhysicalScriptFolder))
			strPhysicalScriptFolder = "";
		
		return strPhysicalScriptFolder;
	}

	class GetHttpStatusString
	{
	private:
		CHAR m_szStatus[70];

	public:

		GetHttpStatusString(DWORD dwHttpStatus) throw()
		{
			LPCSTR szStatusText = NULL;
			CDefaultErrorProvider::GetErrorText(dwHttpStatus, SUBERR_NONE, &szStatusText, NULL);
			if (szStatusText)
			{
				int nRet = _snprintf(m_szStatus, sizeof(m_szStatus), "%d %s", dwHttpStatus, szStatusText);
				if (nRet == -1)
				{
					strcpy(m_szStatus + sizeof(m_szStatus) - 4, "...");
				}
			}
		}

		operator LPCSTR() const throw()
		{
			return m_szStatus;
		}
	};

	inline CAtlStringA StripQuotes(CAtlStringA strQuotedString)
	{
		CAtlStringA strReturn(strQuotedString);
		int iLength = strReturn.Delete(0);//strip off leading \"
		if (iLength > 0)
		{
			iLength = strReturn.Delete(iLength - 1);//strip off trailing \"
		}
		return strReturn;
	}

}; // namespace UTIL