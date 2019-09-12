#pragma once
namespace UTIL
{
	inline HTTP_CODE SendError(CHttpResponse& response, const CStringA& strError, WORD wHttpStatus = 500)
    {
        // Clear any buffered headers (including cookies) and content.
        response.ClearResponse();

        // Suggest that clients and proxies do not cache this response.
        response.SetCacheControl("no-cache");
        response.SetExpires(0);

        // Set the status code in the response object.
        response.SetStatusCode(wHttpStatus);

        // Build the body of the response.
		//response << "<html><head><title>ATL Server Error</title></head><body><p>" << strError << "</p></body></html>";

        // Return a HTTP_CODE that tells the ATL Server code to discontinue processing of the SRF file.
        return HTTP_ERROR(wHttpStatus, SUBERR_NO_PROCESS);
    }
	//this only works with SQL Server
	inline HTTP_CODE SendDBError(CHttpResponse& response, const CStringA& strError, WORD wHttpStatus = 500)
	{
		response.ClearResponse();
		// Suggest that clients and proxies do not cache this response.
        response.SetCacheControl("no-cache");
        response.SetExpires(0);

        // Set the status code in the response object.
        response.SetStatusCode(wHttpStatus);
		 // Build the body of the response.
		//CDBErrorInfo myErrorInfo;
		//ULONG numRec = 0;
		//BSTR myErrStr,mySource;
		//ISQLErrorInfo *pISQLErrorInfo = NULL; 
		//
		//LCID lcLocale = GetSystemDefaultLCID();
		//myErrorInfo.GetErrorRecords(&numRec);
		//if (numRec)
		//{
		//	//CString strErrorUtf8;
		//	myErrorInfo.GetAllErrorInfo(0,lcLocale,&myErrStr,&mySource);
		//	//UnicodeToUtf8(myErrStr, strErrorUtf8); 
		//	response << "<html><head><title>ATL Server Error</title></head><body><p>" << strError << 
		//		"</p><p>" << CW2A(myErrStr) << "</body></html>";
		//}
		// Return a HTTP_CODE that tells the ATL Server code to discontinue processing of the SRF file.
		return HTTP_ERROR(500, SUBERR_NO_PROCESS);
	}
	
	//class GetErrorString
	//{
	//private:
	//	TCHAR* m_szBuffer;

	//public:
	//	GetErrorString(DWORD dwError = GetLastError()) : m_szBuffer(NULL)
	//	{
	//		FormatMessage( 
	//			FORMAT_MESSAGE_ALLOCATE_BUFFER | 
	//			FORMAT_MESSAGE_FROM_SYSTEM | 
	//			FORMAT_MESSAGE_IGNORE_INSERTS,
	//			NULL,
	//			dwError,
	//			MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
	//			(LPTSTR) &m_szBuffer,
	//			0,
	//			NULL 
	//		);

	//		// Remove CR/LF if necessary
	//		if (m_szBuffer != NULL)
	//		{
	//			int nLen = lstrlen(m_szBuffer);
	//			if (nLen > 1 && m_szBuffer[nLen - 1] == '\n')
	//			{
	//				m_szBuffer[nLen - 1] = 0;
	//				if (m_szBuffer[nLen - 2] == '\r')
	//						m_szBuffer[nLen - 2] = 0;
	//			}
	//		} 
	//	}

	//	operator const TCHAR* () const
	//	{
	//		return m_szBuffer;
	//	}

	//	const TCHAR* operator()() const
	//	{
	//		return m_szBuffer;
	//	}

	//	~GetErrorString()
	//	{
	//		// Free the buffer.
	//		if (m_szBuffer != NULL)
	//			LocalFree(m_szBuffer);
	//	}
	//};
}