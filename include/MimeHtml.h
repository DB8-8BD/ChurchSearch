#pragma once
#include "atlmime.h"

class CMimeTextEx : public CMimeText
{
protected:
	CStringA m_strContentType;
public:
	CMimeTextEx(LPCSTR contentType)
	{
		m_strContentType = contentType;
	}
	virtual inline LPCSTR GetContentType() throw()
	{
		return m_strContentType;
	}
	virtual ATL_NOINLINE CMimeBodyPart* Copy() throw( ... )
	{
		CAutoPtr<CMimeTextEx> pNewText;
		ATLTRY(pNewText.Attach(new CMimeTextEx(this->m_strContentType)));
		if (pNewText)
			*pNewText = *this;
		return pNewText.Detach();
	}
	const CMimeTextEx& operator=(const CMimeTextEx& that) throw( ... )
	{
		if (this != &that)
		{
			*(CMimeText*)this = *(CMimeText*)&that;
			m_strContentType = that.m_strContentType;
		}
		return *this;
	}
	virtual inline BOOL MakeMimeHeader(CStringA& header, LPCSTR szBoundary) throw()
	{
		char szBegin[ATL_MIME_BOUNDARYLEN+8];
		if (*szBoundary)
		{
			memcpy(szBegin, "\r\n\r\n--", 6);
			memcpy(szBegin+6, szBoundary, ATL_MIME_BOUNDARYLEN);
			*(szBegin+(ATL_MIME_BOUNDARYLEN+6)) = '\0';
		}
		else
		{
			memcpy(szBegin, "MIME-Version: 1.0", sizeof("MIME-Version: 1.0"));
		}
		_ATLTRY
		{
			header.Format("%s\r\nContent-Type:%s;\r\n\tcharset=\"%s\"\r\nContent-Transfer-Encoding: 8bit\r\n\r\n", szBegin, m_strContentType, m_szCharset);
			return TRUE;
		}
		_ATLCATCHALL()
		{
			return FALSE;
		}
	}
	//~CMimeHtml(void);
};

class CMimeMessageEx : public CMimeMessage
{
public:
	inline BOOL AddCustomBodyPart(CMimeBodyPart *pPart, BOOL bCopy = TRUE, int nPos = 1) throw()
	{
		BOOL bRet = TRUE;
		if (!pPart)
			return FALSE;
		if (nPos < 1)
		{
			nPos = 1;
		}
		_ATLTRY
		{
			CAutoPtr<CMimeBodyPart> spNewComponent;
			if (bCopy)
				spNewComponent.Attach(pPart->Copy());
			else
				spNewComponent.Attach(pPart);
			if (!spNewComponent)
			{
				return FALSE;
			}
			POSITION currPos = m_BodyParts.FindIndex(nPos-1);
			if (!currPos)
			{
				if (!m_BodyParts.AddTail(spNewComponent))
					bRet = FALSE;
			}
			else
			{
				if (!m_BodyParts.InsertBefore(currPos, spNewComponent))
					bRet = FALSE;
			}
		}
		_ATLCATCHALL()
		{
			bRet = FALSE;
		}
		return bRet;
	}
};