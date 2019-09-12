#pragma once
class CRemapper
{
public:
	CRemapper() throw()
	{
	}
	~CRemapper() throw()
	{
	}
	HRESULT Add(LPCSTR szRX, LPCSTR szRepl) throw()
	{
		CAutoPtr<Mapping> spMapping;
		ATLTRY(spMapping.Attach(new Mapping));
		if (spMapping == NULL)
			return E_OUTOFMEMORY;
		spMapping->m_rx.Parse(szRX);
		_ATLTRY
		{
			spMapping->m_strReplacement = szRepl;
		m_aMappings.Add(spMapping);
		}
			_ATLCATCHALL()
		{
			return E_OUTOFMEMORY;
		}
		return S_OK;
	}
	HRESULT Render(LPCSTR szURL, CStringA& strOut) throw()
	{
		ATLASSERT(szURL);
		if (szURL == NULL)
			return E_POINTER;
		for (size_t i = 0; i < m_aMappings.GetCount(); i++)
		{
			Context context;
			Mapping* pMapping = m_aMappings[i];

			BOOL bRet = pMapping->m_rx.Match(szURL, &context);
			if (bRet)
				return RenderMatch(context, pMapping->m_strReplacement, strOut);
		}
		return E_FAIL;
	}
	/*CRemapper& operator=(const CRemapper& c)
	{
		RemoveAll();
		POSITION begin = c.GetStartPosition();
		int num = c.GetCount();

		for (int i = 0; i < num; i++)
		{
			const CmdResultsMap::CPair *pair = c.GetNext(begin);
			SetAt(pair->m_key, pair->m_value);
		}

		return (*this);
	}*/
private:
	typedef CAtlRegExp<CAtlRECharTraitsA> RegExp;
	struct Mapping
	{
		RegExp m_rx;
		CStringA m_strReplacement;
	};
	typedef CAtlREMatchContext<CAtlRECharTraitsA> Context;

	HRESULT RenderMatch(Context& context, CStringA strReplacement, CStringA& strOut) throw()
	{
		LPCSTR szReplacement = strReplacement;
		_ATLTRY
		{
			int nMatchSize = (int)(context.m_Match.szEnd - context.m_Match.szStart);
		ATLENSURE(strReplacement.GetLength() > 0 && strReplacement.GetLength() < strReplacement.GetLength() + nMatchSize);
		strOut.Preallocate(strReplacement.GetLength() + nMatchSize);
		int nCopied = 0;
		while (nCopied < strReplacement.GetLength())
		{
			// find the next '$'
			int nIndex = strReplacement.Find('$', nCopied);
			if (nIndex == -1)
			{
				// no '$' characters left. append the rest of the replacement text
				strOut.Append(szReplacement + nCopied, strReplacement.GetLength() - nCopied);
				nCopied = strReplacement.GetLength();
			}
			else
			{
				if (nIndex > nCopied)
				{
					// append the static text up to the next '$'
					strOut.Append(szReplacement + nCopied, nIndex - nCopied);
					nCopied = nIndex;
				}
				// get the character after the '$'
				if (++nCopied == strReplacement.GetLength())
					return E_FAIL;
				ATLENSURE(nCopied >= 0 && nCopied < strReplacement.GetLength());
				char ch = strReplacement[nCopied++];
				if (!isdigit(static_cast<unsigned char>(ch)))
				{
					// not a regular expression replacement.
					// treat the character after the '$' as a literal character
					// thus '$$' -> '$'
					strOut += ch;
				}
				else
				{
					// replace the $n with the regular expression match
					// specified by the number n
					UINT nWhich = ch - '0';
					if (nWhich >= context.m_uNumGroups)
						return E_FAIL;
					Context::MatchGroup match;
					context.GetMatch(nWhich, &match);
					strOut.Append(match.szStart, (int)(match.szEnd - match.szStart));
				}
			}
		}
		return S_OK;
		}
			_ATLCATCHALL()
		{
			return E_OUTOFMEMORY;
		}
	}
	CAutoPtrArray<Mapping> m_aMappings;
};