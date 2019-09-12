//Updated: May 7, 2018
//Author: Ryan Beckett
//© 2018
#pragma once
#define MAX_PROVINCE		3
#define MAX_REGION			35
#define MAX_CITY			45
#define MAX_DENOMINATION	100
#define MAX_CHURCHNAME		72
#define MAX_CONFIRM_ID		38
#include "DBCommands.h"
#include "util\Error.h"
#include "util\Validator.h"
#include "util\UTIL_API.h"
#include "util\UTIL_Time.h"
#include "util\UTIL_MemoryCache.h"
#include "util\UTIL_ServerContext.h"
class CChurchSearchHandler : public CRequestHandlerT<CChurchSearchHandler>
{
public:
	CAtlStringW m_strProvince, m_strRegion, m_strCity, m_strDenomination, m_strChurchName, m_strConfirmGuid;
	CAtlStringW m_wstrFirstRow, m_strUserAgent;
	bool m_bProvince, m_bRegion, m_bCity, m_bDenomination, m_bChurchName, m_bPg, m_bConfirmGuid, m_bPp, m_bId, m_bRedirect, m_bConfirmed, m_bFirstRow, m_bBypassParam, m_bOldBrowser, m_bMobile, m_bSsl, m_bHttp2, m_bHttp2Capable, m_bAssetsPushed;
	DBCOUNTITEM m_Rows, m_Pages, m_CurPage, m_CurRow, m_dbPp, m_dbPg, m_Scroll;
	COleDateTime m_LastModified;
	int m_iPp, m_iPg, m_iId;
	CAtlList<CAtlStringA> m_liEtags;
	//default constructor
	CChurchSearchHandler()
	{
		m_bProvince = false;
		m_bRegion = false;
		m_bCity = false;
		m_bDenomination = false;
		m_bChurchName = false;
		m_bPg = false;
		m_bPp = false;
		m_bConfirmGuid = false;
		m_bId = false;
		m_bRedirect = false;
		m_bConfirmed = false;
		m_bFirstRow = false;
		m_bBypassParam = false;
		m_bOldBrowser = false;
		m_bMobile = false;
		m_bSsl = false;
		m_bHttp2 = false;
		m_bHttp2Capable = false;
		m_bAssetsPushed = false;
		m_iPp = 0;
		m_iPg = 0;
		m_iId = 0;
		m_Rows = 0;
		m_dbPp = 10;
		m_dbPg = 1;
		m_Pages = 0;
		m_Scroll = 0;
		m_CurRow = 0;
		m_CurPage = 1;
	}
	//initialize the request handler
	HTTP_CODE ValidateAndExchange()
	{
		m_HttpResponse.SetContentType("text/html");
		//get the file's last modified time (authorship), this may be used to add the last-modified header for caching purposes
		HANDLE hFile;
		hFile = CreateFile(CA2T(m_HttpRequest.GetScriptPathTranslated()), GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
		m_LastModified = UTIL::GetFileModTime(hFile);
		CloseHandle(hFile);
		//connect to the database
		CDataConnection dc;
		if (S_OK != GetDataSource(m_spServiceProvider, _T("Churches"),
			L"Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=churches",
			&dc))
			return UTIL::SendDBError(m_HttpResponse, "Unable to Query Data Source");
		//obtain browser caps service
		m_spServiceProvider->QueryService(__uuidof(IBrowserCapsSvc), &m_spBrowserCapsSvc);
		int iCntQueryParams = 0;//counts number of parameters submitted
		//obtain a pointer the query-params / form-fields, based on the request method
		const CHttpRequestParams *requestParams(NULL);
		if (m_HttpRequest.GetMethod() == CHttpRequest::HTTP_METHOD_POST)
		{
			requestParams = &(this->m_HttpRequest.GetFormVars());
		}
		else
		{
			requestParams = &(this->m_HttpRequest.GetQueryParams());
		}
		if (requestParams->GetCount() == 0)
		{
			// there are no parameters, so the main page will be displayed
		}
		//queryparam or form-field for the province
		if (VALIDATION_SUCCEEDED(requestParams->Validate("Province", static_cast<LPCSTR*>(NULL), 2, MAX_PROVINCE)))
		{
			m_strProvince = requestParams->Lookup("Province");
			if (UTIL::ParamIsProvince(m_strProvince))
			{
				m_bProvince = true;
				iCntQueryParams++;
			}
		}
		//queryparam or form-field for the city
		if (VALIDATION_SUCCEEDED(requestParams->Validate("Region", static_cast<LPCSTR*>(NULL), 2, MAX_REGION)))
		{
			m_strRegion = requestParams->Lookup("Region");
			if (UTIL::ParamIsAcceptable(m_strRegion))
			{
				m_bRegion = true;
				iCntQueryParams++;
			}
		}
		//queryparam or form-field for the city
		if (VALIDATION_SUCCEEDED(requestParams->Validate("City", static_cast<LPCSTR*>(NULL), 2, MAX_CITY)))
		{
			m_strCity = requestParams->Lookup("City");
			if (UTIL::ParamIsAcceptable(m_strCity))
			{
				m_bCity = true;
				iCntQueryParams++;
			}
		}
		//queryparam or form-field for the denomination
		if (VALIDATION_SUCCEEDED(requestParams->Validate("Denomination", static_cast<LPCSTR*>(NULL), 2, MAX_DENOMINATION)))
		{
			m_strDenomination = requestParams->Lookup("Denomination");
			if (UTIL::ParamIsAcceptable(m_strDenomination))
			{
				m_bDenomination = true;
				iCntQueryParams++;
			}
		}
		//queryparam or form-field for the churchname
		if (VALIDATION_SUCCEEDED(requestParams->Validate("ChurchName", static_cast<LPCSTR*>(NULL), 2, MAX_CHURCHNAME)))
		{
			m_strChurchName = requestParams->Lookup("ChurchName");
			if (UTIL::ParamIsAcceptable(m_strChurchName))
			{
				m_bChurchName = true;
				m_strChurchName.Trim();
				m_strChurchName.Insert(0, L'\"');
				m_strChurchName.Append(L"\"");
				iCntQueryParams++;
			}
		}
		//queryparam or form-field for the user's confirmation code
		if (VALIDATION_SUCCEEDED(requestParams->Validate("confirmguid", static_cast<LPCSTR*>(NULL), 2, MAX_CONFIRM_ID)))
		{
			m_strConfirmGuid = requestParams->Lookup("confirmguid");
			if (UTIL::ParamIsAcceptable(m_strConfirmGuid))
			{
				m_bConfirmGuid = true;
			}
		}
		//queryparam or form-field with the per-page number
		if (VALIDATION_S_OK == requestParams->Exchange("Pp", &m_iPp))
		{
			m_bPp = true;
			if (m_iPp > 0)
			{
				m_dbPp = (DBCOUNTITEM)m_iPp;
			}
		}
		//queryparam or form-field with the page number
		if (VALIDATION_S_OK == requestParams->Exchange("Pg", &m_iPg))
		{
			m_bPg = true;
			m_CurPage = (DBCOUNTITEM)m_iPg;
		}
		if (VALIDATION_S_OK == requestParams->Exchange("ChurchID", &m_iId))
		{
			m_bId = true;
		}
		//search for any etags that may have been sent along
		CAtlStringA strIfNoneMatchHeader;
		if (UTIL::GetServerVariable(m_spServerContext, "HTTP_IF_NONE_MATCH", strIfNoneMatchHeader))
		{
			int iTokenPos = 0;
			CAtlStringA strToken = strIfNoneMatchHeader.Tokenize(",", iTokenPos);
			while (!strToken.IsEmpty())
			{
				m_liEtags.AddTail(UTIL::StripQuotes(strToken));
				strToken = strIfNoneMatchHeader.Tokenize(",", iTokenPos);
			}
			if (m_liEtags.IsEmpty())//no multiple etags
			{
				m_liEtags.AddTail(UTIL::StripQuotes(strIfNoneMatchHeader));
			}
		}
		HRESULT hr;
		//if the user has clicked on an individual item
		if (m_bId)
		{
			m_ChurchDetails.m_churchid = (LONG)m_iId;
			if (S_OK != m_ChurchDetails.OpenRowset(dc))
			{
				return UTIL::SendError(m_HttpResponse, "Error with spChurchDetails");
			}
			m_bRedirect = false;
			return HTTP_SUCCESS;
		}
		//allow users to submit confirmation requests when signing up to the CSE application for the first time
		//if true, complete the signup and return the result
		
		if (m_bConfirmGuid)
		{
			CspConfirmUser m_ConfirmUser;
			GUID guid = GUID_NULL;
			CW2A strConfirmGuid(m_strConfirmGuid);
			//use a Windows API to validate the user's confirmation request 
			//(converts a signed pointer to char array to an unsigned one)
			//#pragma lib(rpcrt4.lib)
			if (RPC_S_OK != UuidFromStringA(reinterpret_cast<UTIL::VCUE_RPC_CHAR*>(const_cast<LPSTR>((LPCSTR)strConfirmGuid)), &guid))
				return UTIL::SendError(m_HttpResponse, CAtlStringA("Unable to convert guid value"), 501);
			memcpy(&m_ConfirmUser.m_Guid, &guid, sizeof(GUID));
			m_ConfirmUser.m_dwGuidLength = sizeof(GUID);
			m_ConfirmUser.m_dwGuidStatus = DBSTATUS_S_OK;
			hr = m_ConfirmUser.OpenRowset(dc);
			if (S_OK != hr)
				return UTIL::SendDBError(m_HttpResponse, CAtlStringA("Unable to query database"), 501);
			else
			{
				m_bConfirmed = true;
				m_bId = true;
				return HTTP_SUCCESS;
			}
		}
		//test for https, this is needed for redirect strings, and determining HTTP/2 compatibility
		CAtlStringA strHttpsOn;
		if (m_HttpRequest.GetServerVariable("HTTPS", strHttpsOn))
		{
			if (0 == strHttpsOn.CompareNoCase("ON"))
			{
				m_bSsl = true;
			}
		}
		//test for server http2, although the client might be eligible for HTTP/2, make sure that the desired protocol in-use by the connection is indeed HTTP/2
		CAtlStringA strHttp2On;
		if (m_HttpRequest.GetServerVariable("SERVER_PROTOCOL", strHttp2On))
		{
			if (0 == strHttp2On.CompareNoCase("HTTP/2.0"))
			{
				m_bHttp2 = true;
			}
		}
		//it is usually best to determine browser capabilities on the server.
		//the main purpose of this implementation is to determine whether or not the client is using an old version of Internet Explorer
		//if true, then apply a specific css selector to help fix layout issues as the old trident layout engine is incompatible with newer css 
		if (S_OK == m_spBrowserCapsSvc->GetCaps(m_spServerContext, &m_spBrowserCaps))
		{
			BOOL bMobile = false;
			m_spBrowserCaps->GetBooleanPropertyValue(CComBSTR(L"isMobileDevice"), &bMobile);
			if (0 != bMobile)
				m_bMobile = true;
			CComBSTR bstrTemp;
			if (S_OK == m_spBrowserCaps->GetBrowserName(&bstrTemp))
			{
				m_strUserAgent = bstrTemp;
				if (0 == m_strUserAgent.Compare(_T("IE")))
				{
					CComBSTR bstrVersion;
					if (S_OK == m_spBrowserCaps->GetVersion(&bstrVersion))
					{
						CAtlStringW wstrVersion(bstrVersion);
						wchar_t *pStop;
						unsigned int iTemp = 0;
						iTemp = wcstoul(wstrVersion, &pStop, 10);
						if (iTemp < 8)
						{
							//ie7 or less, m_bOldBrowser affects the implementation of css layout
							m_bOldBrowser = true;
							m_bHttp2Capable = false;
						}
						else
						{
							//search for ie7 compatibility mode
							CComBSTR bstrCompatibilityMode;
							if (S_OK == m_spBrowserCaps->GetPropertyString(CComBSTR("Comment"), &bstrCompatibilityMode))
							{
								CAtlStringW wstrCompatibilityMode(bstrCompatibilityMode);
								if (-1 != wstrCompatibilityMode.Find(L"7.0"))
								{
									//affects the implementation of css layout
									m_bOldBrowser = true;
									m_bHttp2Capable = false;
								}
							}
						}
					}
					//now determine whether this browser is compatible with HTTP/2
					CComBSTR bstrPlatform;
					if (S_OK == m_spBrowserCaps->GetPlatform(&bstrPlatform))
					{
						CAtlStringW wstrPlatform(bstrPlatform);
						if (0 == wstrPlatform.Compare(L"Win10"))
						{
							//HTTP/2 requires TLS
							if (m_bSsl)
							{
								m_bHttp2Capable = true;
							}
						}
					}
				}
				else if (0 == m_strUserAgent.Compare(_T("Edge")))
				{
					//edge is only compatible with HTTP/2 if it's run on Windows 10
					CComBSTR bstrPlatform;
					if (S_OK == m_spBrowserCaps->GetPlatform(&bstrPlatform))
					{
						CAtlStringW wstrPlatform(bstrPlatform);
						if (0 == wstrPlatform.Compare(L"Win10"))
						{
							if (m_bSsl)
							{
								m_bHttp2Capable = true;
							}
						}
					}
				}
				else if (0 == m_strUserAgent.Compare(_T("Chrome")))
				{
					//chrome (on both android and desktop) is compatible with HTTP/2 on versions greater than 65
					CComBSTR bstrVersion;
					if (S_OK == m_spBrowserCaps->GetVersion(&bstrVersion))
					{
						CAtlStringW wstrVersion(bstrVersion);
						wchar_t *pStop;
						unsigned int iTemp = 0;
						iTemp = wcstoul(wstrVersion, &pStop, 10);
						if (iTemp >= 66)
						{
							if (m_bSsl)
							{
								m_bHttp2Capable = true;
							}
						}
					}
				}
				else if (0 == m_strUserAgent.Compare(_T("Firefox")))
				{
					//firefox is compatible with HTTP/2 on versions greater than 58
					CComBSTR bstrVersion;
					if (S_OK == m_spBrowserCaps->GetVersion(&bstrVersion))
					{
						CAtlStringW wstrVersion(bstrVersion);
						wchar_t *pStop;
						unsigned int iTemp = 0;
						iTemp = wcstoul(wstrVersion, &pStop, 10);
						if (iTemp >= 59)
						{
							if (m_bSsl)
							{
								m_bHttp2Capable = true;
							}
						}
					}
				}
				else
				{
					if (m_bSsl)
					{
						m_bHttp2Capable = true;
					}
				}
			}
		}
		//HTTP/2 pushes are only valid for GET methods
		if (m_HttpRequest.GetMethod() == CHttpRequest::HTTP_METHOD_GET)
		{
			if (m_bHttp2Capable)//client is able to use HTTP/2
			{
				if (m_bHttp2)//server has an HTTP/2 connection
				{
					// obtain client's ip address/port combination
					CAtlStringA strConnectionID;
					CAtlStringA strIPAddress;
					m_HttpRequest.GetServerVariable("REMOTE_ADDR", strIPAddress);
					CAtlStringA strPort;
					m_HttpRequest.GetServerVariable("REMOTE_PORT", strPort);
					strIPAddress.Remove('.');
					strPort.Remove(':');
					strIPAddress.Remove('[');
					strIPAddress.Remove(']');
					strConnectionID = strIPAddress.Right(strIPAddress.ReverseFind(':'));
					strConnectionID += strPort;
					ULONGLONG ullConnectionID = 0;
					// convert the string returned from connection-id to an unsigned long long
					char *pEnd;
					ullConnectionID = strtoull(strConnectionID, &pEnd, 10);
					if (0 < ullConnectionID)
					{
						ULARGE_INTEGER uli;
						uli.QuadPart = ullConnectionID;
						CspGetIPCount spGetIPCount;
						spGetIPCount.m_connectionid = uli;
						hr = spGetIPCount.OpenRowset(dc);
						if (S_OK == hr)
						{
							CAtlStringA strBaseImgDir = m_HttpRequest.GetScriptPathTranslated();//full file path of the current script
							CAtlStringA strBaseCssDir = strBaseImgDir.Left(strBaseImgDir.ReverseFind('\\'));//slice off filename
							strBaseImgDir = strBaseCssDir;//reset string to dir
							strBaseImgDir += "\\images\\";//append images dir
							//build a string of GET method headers (HTTP-PUSH = GET method)
							CAtlStringA strGetHeaders;
							strGetHeaders = AssembleHttpHeadersPush();
							CAtlStringA strImgPngGetHeaders(strGetHeaders);
							strImgPngGetHeaders.Append("Accept: image/png\r\n");
							strImgPngGetHeaders.Append("Accept-Encoding: identity\r\n");
							CAtlStringA strImgGifGetHeaders(strGetHeaders);
							strImgGifGetHeaders.Append("Accept: image/gif\r\n");
							strImgGifGetHeaders.Append("Accept-Encoding: identity\r\n");
							hr = spGetIPCount.MoveNext();
							int iCounter = 0;
							if (0 < spGetIPCount.m_Count)//endofrowset means success
							{
								//if the user hasn selected a query parameter, such as a province or city, then push images for rowscroll
								//which they may not have in their browser's cache
								strBaseImgDir += "numbers\\";
								if (iCntQueryParams > 0)
								{
									int iOffsetPos = m_iPg + 10;
									if (iOffsetPos > spGetIPCount.m_Count)
									{
										//we wish to push only images in groups of 10's, and they must be ones that the client hasn't downloaded yet
										iOffsetPos = spGetIPCount.m_Count + 11;
										for (iCounter = spGetIPCount.m_Count; iCounter < iOffsetPos; iCounter++)
										{
											if (iCounter > 100)
												break;//it's not possible to have more than 100 images
											CAtlStringA strImgPath(strBaseImgDir);
											strImgPath.AppendFormat("%i.png", iCounter);
											if (!PushFile(strImgPath, strImgPngGetHeaders))
												break;//if push operation fails, don't do anymore, let the browser request them
										}
										CspUpdateIP updIP;
										updIP.m_connectionid = uli;
										updIP.m_newcount = --iCounter;
										hr = updIP.OpenRowset(dc);//update the database with images pushed so that the server doesn't have to push in vain in the future
										m_bAssetsPushed = true;//cannot cache this page if we have used transmitfile function
									}
								}
							}
							else
							{
								if (iCntQueryParams > 0)
								{
									if (0 == spGetIPCount.m_Count)
									{
										CAtlStringA strImgPrev(strBaseImgDir);
										strImgPrev += "Prev.png";
										PushFile(strImgPrev, strImgPngGetHeaders);
										CAtlStringA strImgNext(strBaseImgDir);
										strImgNext += "Next.png";
										PushFile(strImgNext, strImgPngGetHeaders);
									}
									strBaseImgDir += "numbers\\";
									for (iCounter = 1; iCounter < 11; iCounter++)
									{
										if (iCounter > 100)
											break;
										CAtlStringA strImgPath(strBaseImgDir);
										strImgPath.AppendFormat("%i.png", iCounter);
										if (!PushFile(strImgPath, strImgPngGetHeaders))
											break;
									}
									CspUpdateIP updIP;
									updIP.m_connectionid = uli;
									updIP.m_newcount = --iCounter;
									hr = updIP.OpenRowset(dc);//update the database with images pushed so that the server doesn't have to push in vain in the future
									m_bAssetsPushed = true;//cannot cache this page if we have used transmitfile function
								}
								else
								{
									//first time client has visited this web-app, so push layout related assets
									CAtlStringA strAcceptLang;
									if (m_HttpRequest.GetServerVariable("HTTP_ACCEPT_LANGUAGE", strAcceptLang))
									{
										strGetHeaders.AppendFormat("Accept-Language: %s\r\n", strAcceptLang);
									}
									strBaseCssDir += "\\css\\default.css";
									strGetHeaders.Append("Accept: text/css\r\n");
									strGetHeaders.Append("Accept-Encoding: identity\r\n");
									PushFile(strBaseCssDir, strGetHeaders);//css-file gets pushed
									CAtlStringA strImgCanada(strBaseImgDir);
									strImgCanada += "canada.gif";
									PushFile(strImgCanada, strImgGifGetHeaders);
									CAtlStringA strImgBottom(strBaseImgDir);
									strImgBottom += "Bottom.png";
									PushFile(strImgBottom, strImgPngGetHeaders);
									CAtlStringA strImgBottomRight(strBaseImgDir);
									strImgBottomRight += "BottomRight.png";
									PushFile(strImgBottomRight, strImgPngGetHeaders);
									CAtlStringA strImgLeft(strBaseImgDir);
									strImgLeft += "Left.png";
									PushFile(strImgLeft, strImgPngGetHeaders);
									CAtlStringA strImgTopLeft(strBaseImgDir);
									strImgTopLeft += "TopLeft.png";
									PushFile(strImgTopLeft, strImgPngGetHeaders);
									CAtlStringA strImgBottomLeft(strBaseImgDir);
									strImgBottomLeft += "BottomLeft.png";
									PushFile(strImgBottomLeft, strImgPngGetHeaders);
									CAtlStringA strImgRight(strBaseImgDir);
									strImgRight += "Right.png";
									PushFile(strImgRight, strImgPngGetHeaders);
									CAtlStringA strImgTopRight(strBaseImgDir);
									strImgTopRight += "TopRight.png";
									PushFile(strImgTopRight, strImgPngGetHeaders);
									CAtlStringA strImgTop(strBaseImgDir);
									strImgTop += "Top.png";
									PushFile(strImgTop, strImgPngGetHeaders);
									CspAddIP addIP;
									addIP.m_connectionid = uli;
									addIP.m_newcount = 0;
									hr = addIP.OpenRowset(dc);//add client's ip address to the database.
									                          //count of 0 means that they haven't downloaded any rowscroll pics
									m_bAssetsPushed = true;//cannot cache this page if we have used transmitfile function
								}
							}
						}
					}
				}
			}
		}

		DBCOUNTITEM cRows = 0;//temp variable, used for scrollingin rowset
		if (m_bProvince)
		{
			if (m_bRegion)
			{
				if (m_bCity)
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							// ProvRegCityDenomChurchNameChurches - Churches by ChurchName & Params
							Checked::tcsncpy_s(m_ProvRegCityDenomChurchNameChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegCityDenomChurchNameChurches.m_region, MAX_REGION, m_strRegion, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegCityDenomChurchNameChurches.m_city, MAX_CITY, m_strCity, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegCityDenomChurchNameChurches.m_denomination, MAX_DENOMINATION, m_strDenomination, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegCityDenomChurchNameChurches.m_churchname, MAX_CHURCHNAME, m_strChurchName, _TRUNCATE);
							if (S_OK != m_ProvRegCityDenomChurchNameChurches.OpenRowset(dc))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegCityDenomChurchNameChurches");
							}
							//since the user has submitted a name parameter, turn off the first-row flag, which is used to display headings for table columns
							//because it's possible that there will be only one result, which means the layout of the page will be transformed from a multi-column
							//table to a single-column one. The HasMoreChurches method is used to control how many records are outputted  onto the page
							m_bFirstRow = false;
						}
						else
						{
							Checked::tcsncpy_s(m_ProvRegCityDenomChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegCityDenomChurches.m_region, MAX_REGION, m_strRegion, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegCityDenomChurches.m_city, MAX_CITY, m_strCity, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegCityDenomChurches.m_denomination, MAX_DENOMINATION, m_strDenomination, _TRUNCATE);
							if (S_OK != m_ProvRegCityDenomChurches.OpenRowset(dc))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegCityDenomChurches");
							}
							//use OLEDB (hierarchical database access) to enumerate all the records before actually using them
							hr = m_ProvRegCityDenomChurches.GetApproximatePosition(NULL, NULL, &m_Rows);
							if (S_OK != hr)
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegCityDenomChurches.GetApproximatePosition()");
							}
							//return how many pages there are in the rowset
							m_Pages = (DBCOUNTITEM)UTIL::CalculatePages(m_Rows, m_dbPp);
							if (m_CurPage > m_Pages)
							{
								//end-of-rowset means reset scroll to page 0
								m_CurPage = 0;
							}
							m_Scroll = m_CurPage * m_dbPp;
							//if the user has specified a page #, move to the specified page
							if (m_bPg)
							{
								if (0 != m_iPg)
								{
									//use OLEDB to move the rowset pointer to the calculated scroll position
									if (S_OK != m_ProvRegCityDenomChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegCityDenomChurches.MoveNext()");
									}
									//after the rowset pointer has been scrolled, test for end-of-rowset
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									//layout may be single-column
									m_bFirstRow = false;
								}
								else
								{
									//user has not specified a valid page, so move to their last known position using OLEDB
									if (S_OK != m_ProvRegCityDenomChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegCityDenomChurches.MoveNext()");
									}
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									//layout may not be signle-column
									m_bFirstRow = true;
									//load the first fetched result from the rowset to the status replacement tag handler
									//at this point, we are reasonably certain that the page will have multiple records
									m_wstrFirstRow = m_ProvRegCityDenomChurches.GetFirstRow(m_strProvince, m_strRegion, m_strCity, m_strDenomination);
								}
							}
							else
							{
								//no page number specified, so go to first record (multiple records expected)
								m_bFirstRow = true;
								if (S_OK != m_ProvRegCityDenomChurches.MoveFirst())
								{
									return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegCityDenomChurches.MoveFirst()");
								}
								m_wstrFirstRow = m_ProvRegCityDenomChurches.GetFirstRow(m_strProvince, m_strRegion, m_strCity, m_strDenomination);
							}
						}
						//load denominations selector listbox
						Checked::tcsncpy_s(m_ProvRegCityDenoms.m_city, MAX_CITY, m_strCity, _TRUNCATE);
						Checked::tcsncpy_s(m_ProvRegCityDenoms.m_region, MAX_REGION, m_strRegion, _TRUNCATE);
						Checked::tcsncpy_s(m_ProvRegCityDenoms.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
						if (S_OK != m_ProvRegCityDenoms.OpenRowset(dc))
						{
							return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegCityDenoms");
						}
					}
					else
					{
						Checked::tcsncpy_s(m_ProvRegCityDenoms.m_city, MAX_CITY, m_strCity, _TRUNCATE);
						Checked::tcsncpy_s(m_ProvRegCityDenoms.m_region, MAX_REGION, m_strRegion, _TRUNCATE);
						Checked::tcsncpy_s(m_ProvRegCityDenoms.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
						if (S_OK != m_ProvRegCityDenoms.OpenRowset(dc))
						{
							return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegCityDenoms");
						}
						if (S_OK != m_ProvRegCityDenoms.GetApproximatePosition(NULL, NULL, &cRows))
						{
							return UTIL::SendDBError(m_HttpResponse, "Error with m_ProvRegCitySuburbs.GetApproximatePosition()");
						}
						// If there's only one result, then automatically redirect to that one denomination.
						if (cRows == 1)
						{
							if (S_OK != m_ProvRegCityDenoms.MoveFirst())
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with m_ProvRegCityDenoms.MoveFirst()");
							}
							m_bRedirect = true;
							m_bBypassParam = true;//redirect 'cause there's only one parameter
						}
						else
						{
							// it's possible that m_bRedirect was set to true by a churchid...
							m_bRedirect = false;
							m_bBypassParam = false;
						}
						if (m_bChurchName)
						{
							Checked::tcsncpy_s(m_ProvRegCityChurchNameChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegCityChurchNameChurches.m_region, MAX_REGION, m_strRegion, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegCityChurchNameChurches.m_city, MAX_CITY, m_strCity, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegCityChurchNameChurches.m_churchname, MAX_CHURCHNAME, m_strChurchName, _TRUNCATE);
							if (S_OK != m_ProvRegCityChurchNameChurches.OpenRowset(dc))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegCityChurchNameChurches");
							}
							m_bFirstRow = false;
						}
						else
						{
							Checked::tcsncpy_s(m_ProvRegCityChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegCityChurches.m_region, MAX_REGION, m_strRegion, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegCityChurches.m_city, MAX_CITY, m_strCity, _TRUNCATE);
							if (S_OK != m_ProvRegCityChurches.OpenRowset(dc))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegCityChurches");
							}
							if (S_OK != m_ProvRegCityChurches.GetApproximatePosition(NULL, NULL, &m_Rows))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with m_ProvRegCityChurches.GetApproximatePosition()");
							}
							m_Pages = (DBCOUNTITEM)UTIL::CalculatePages(m_Rows, m_dbPp);
							if (m_CurPage > m_Pages)
							{
								// reset scroll
								m_CurPage = 0;
							}
							m_Scroll = m_CurPage * m_dbPp;
							if (m_bPg)
							{
								if (0 != m_iPg)
								{
									if (S_OK != m_ProvRegCityChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegCityChurches.MoveNext()");
									}
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									m_bFirstRow = false;
								}
								else
								{
									if (S_OK != m_ProvRegCityChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegCityChurches.MoveNext()");
									}
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									m_bFirstRow = true;
									m_wstrFirstRow = m_ProvRegCityChurches.GetFirstRow(m_strProvince, m_strRegion, m_strCity);
								}
							}
							else
							{
								m_bFirstRow = true;
								if (S_OK != m_ProvRegCityChurches.MoveFirst())
								{
									return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegCityChurches.MoveFirst()");
								}
								m_wstrFirstRow = m_ProvRegCityChurches.GetFirstRow(m_strProvince, m_strRegion, m_strCity);
							}
						}
					}
					Checked::tcsncpy_s(m_ProvRegCities.m_region, MAX_REGION, m_strRegion, _TRUNCATE);
					Checked::tcsncpy_s(m_ProvRegCities.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
					if (S_OK != m_ProvRegCities.OpenRowset(dc))
					{
						return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegCities");
					}
				}
				else
				{
					Checked::tcsncpy_s(m_ProvRegCities.m_region, MAX_REGION, m_strRegion, _TRUNCATE);
					Checked::tcsncpy_s(m_ProvRegCities.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
					if (S_OK != m_ProvRegCities.OpenRowset(dc))
					{
						return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegCities");
					}
					if (S_OK != m_ProvRegCities.GetApproximatePosition(NULL, NULL, &cRows))
					{
						return UTIL::SendDBError(m_HttpResponse, "Error with m_ProvRegCities.GetApproximatePosition()");
					}
					if (cRows == 1)
					{
						if (S_OK != m_ProvRegCities.MoveFirst())
						{
							return UTIL::SendDBError(m_HttpResponse, "Error with m_ProvRegCities.MoveFirst()");
						}
						m_bRedirect = true;
						m_bBypassParam = true;
					}
					else
					{
						// it's possible that m_bRedirect was set to true by a churchid...
						m_bRedirect = false;
						m_bBypassParam = false;
					}
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvRegDenomChurchNameChurches
							Checked::tcsncpy_s(m_ProvRegDenomChurchNameChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegDenomChurchNameChurches.m_region, MAX_REGION, m_strRegion, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegDenomChurchNameChurches.m_denomination, MAX_DENOMINATION, m_strDenomination, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegDenomChurchNameChurches.m_churchname, MAX_CHURCHNAME, m_strChurchName, _TRUNCATE);
							if (S_OK != m_ProvRegDenomChurchNameChurches.OpenRowset(dc))
							{
								return UTIL::SendError(m_HttpResponse, "Error with spProvRegDenomChurchNameChurches");
							}
							m_bFirstRow = false;
						}
						else
						{
							//ProvRegDenomChurches
							Checked::tcsncpy_s(m_ProvRegDenomChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegDenomChurches.m_region, MAX_REGION, m_strRegion, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegDenomChurches.m_denomination, MAX_DENOMINATION, m_strDenomination, _TRUNCATE);
							if (S_OK != m_ProvRegDenomChurches.OpenRowset(dc))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegDenomChurches");
							}
							if (S_OK != m_ProvRegDenomChurches.GetApproximatePosition(NULL, NULL, &m_Rows))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with m_ProvRegDenomChurches.GetApproximatePosition()");
							}
							m_Pages = (DBCOUNTITEM)UTIL::CalculatePages(m_Rows, m_dbPp);
							if (m_CurPage > m_Pages)
							{
								// reset scroll
								m_CurPage = 0;
							}
							m_Scroll = m_CurPage * m_dbPp;
							if (m_bPg)
							{
								if (0 != m_iPg)
								{
									if (S_OK != m_ProvRegDenomChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegDenomChurches.MoveNext()");
									}
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									m_bFirstRow = false;
								}
								else
								{
									if (S_OK != m_ProvRegDenomChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegDenomChurches.MoveNext()");
									}
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									m_bFirstRow = true;
									m_wstrFirstRow = m_ProvRegDenomChurches.GetFirstRow(m_strProvince, m_strRegion, m_strDenomination);
								}
							}
							else
							{
								m_bFirstRow = true;
								if (S_OK != m_ProvRegDenomChurches.MoveFirst())
								{
									return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegDenomChurches.MoveFirst()");
								}
								m_wstrFirstRow = m_ProvRegDenomChurches.GetFirstRow(m_strProvince, m_strRegion, m_strDenomination);
							}
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvRegChurchNameChurches
							Checked::tcsncpy_s(m_ProvRegChurchNameChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegChurchNameChurches.m_region, MAX_REGION, m_strRegion, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegChurchNameChurches.m_churchname, MAX_CHURCHNAME, m_strChurchName, _TRUNCATE);
							if (S_OK != m_ProvRegChurchNameChurches.OpenRowset(dc))
							{
								return UTIL::SendError(m_HttpResponse, "Error with spProvRegChurchNameChurches");
							}
							m_bFirstRow = false;
						}
						else
						{
							//ProvRegChurches
							Checked::tcsncpy_s(m_ProvRegChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvRegChurches.m_region, MAX_REGION, m_strRegion, _TRUNCATE);
							if (S_OK != m_ProvRegChurches.OpenRowset(dc))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegChurches");
							}
							if (S_OK != m_ProvRegChurches.GetApproximatePosition(NULL, NULL, &m_Rows))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with m_ProvRegChurches.GetApproximatePosition()");
							}
							m_Pages = (DBCOUNTITEM)UTIL::CalculatePages(m_Rows, m_dbPp);
							if (m_CurPage > m_Pages)
							{
								// reset scroll
								m_CurPage = 0;
							}
							m_Scroll = m_CurPage * m_dbPp;
							if (m_bPg)
							{
								if (0 != m_iPg)
								{
									if (S_OK != m_ProvRegChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegChurches.MoveNext()");
									}
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									m_bFirstRow = false;
								}
								else
								{
									if (S_OK != m_ProvRegChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegChurches.MoveNext()");
									}
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									m_bFirstRow = true;
									m_wstrFirstRow = m_ProvRegChurches.GetFirstRow(m_strProvince, m_strRegion);
								}
							}
							else
							{
								if (S_OK != m_ProvRegChurches.MoveFirst())
								{
									return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegChurches.MoveFirst()");
								}
								m_bFirstRow = true;
								m_wstrFirstRow = m_ProvRegChurches.GetFirstRow(m_strProvince, m_strRegion);
							}
						}
					}
				}
				Checked::tcsncpy_s(m_ProvRegions.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
				// this code is always executed if there is a region - ProvRegions
				if (S_OK != m_ProvRegions.OpenRowset(dc))
				{
					return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegions");
				}
			}
			else
			{
				// this code executes if there isn't a region - ProvRegions
				Checked::tcsncpy_s(m_ProvRegions.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
				if (S_OK != m_ProvRegions.OpenRowset(dc))
				{
					return UTIL::SendDBError(m_HttpResponse, "Error with spProvRegions");
				}
				if (S_OK != m_ProvRegions.GetApproximatePosition(NULL, NULL, &cRows))
				{
					return UTIL::SendDBError(m_HttpResponse, "Error with m_ProvRegions.GetApproximatePosition");
				}
				// If there's only one  result, then automatically redirect to that one region.
				if (cRows == 1)
				{
					if (S_OK != m_ProvRegions.MoveFirst())
					{
						return UTIL::SendDBError(m_HttpResponse, "Error with m_ProvRegCities.MoveFirst()");
					}
					m_bRedirect = true;
					m_bBypassParam = true;
				}
				else
				{
					m_bRedirect = false;
					m_bBypassParam = false;
				}
				if (m_bCity)
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvCityDenomChurchNameChurches
							Checked::tcsncpy_s(m_ProvCityDenomChurchNameChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvCityDenomChurchNameChurches.m_city, MAX_CITY, m_strCity, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvCityDenomChurchNameChurches.m_denomination, MAX_DENOMINATION, m_strDenomination, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvCityDenomChurchNameChurches.m_churchname, MAX_CHURCHNAME, m_strChurchName, _TRUNCATE);
							if (S_OK != m_ProvCityDenomChurchNameChurches.OpenRowset(dc))
							{
								return UTIL::SendError(m_HttpResponse, "Error with spProvCityDenomChurchNameChurches");
							}
							m_bFirstRow = false;
						}
						else
						{
							//ProvCityDenomChurches
							Checked::tcsncpy_s(m_ProvCityDenomChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvCityDenomChurches.m_city, MAX_CITY, m_strCity, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvCityDenomChurches.m_denomination, MAX_DENOMINATION, m_strDenomination, _TRUNCATE);
							if (S_OK != m_ProvCityDenomChurches.OpenRowset(dc))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with spProvCityDenomChurches");
							}
							if (S_OK != m_ProvCityDenomChurches.GetApproximatePosition(NULL, NULL, &m_Rows))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with m_ProvCityDenomChurches.GetApproximatePosition()");
							}
							m_Pages = (DBCOUNTITEM)UTIL::CalculatePages(m_Rows, m_dbPp);
							if (m_CurPage > m_Pages)
							{
								// reset scroll
								m_CurPage = 0;
							}
							m_Scroll = m_CurPage * m_dbPp;
							if (m_bPg)
							{
								if (0 != m_iPg)
								{
									if (S_OK != m_ProvCityDenomChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvCityDenomChurches.MoveNext()");
									}
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									m_bFirstRow = false;
								}
								else
								{
									if (S_OK != m_ProvCityDenomChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvCityDenomChurches.MoveNext()");
									}
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									m_bFirstRow = true;
									m_wstrFirstRow = m_ProvCityDenomChurches.GetFirstRow(m_strProvince, m_strCity, m_strDenomination);
								}
							}
							else
							{
								m_bFirstRow = true;
								if (S_OK != m_ProvCityDenomChurches.MoveFirst())
								{
									return UTIL::SendDBError(m_HttpResponse, "Error with spProvCityDenomChurches.MoveFirst()");
								}
								m_wstrFirstRow = m_ProvCityDenomChurches.GetFirstRow(m_strProvince, m_strCity, m_strDenomination);
							}
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvCityChurchNameChurches
							Checked::tcsncpy_s(m_ProvCityChurchNameChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvCityChurchNameChurches.m_city, MAX_CITY, m_strCity, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvCityChurchNameChurches.m_churchname, MAX_CHURCHNAME, m_strChurchName, _TRUNCATE);
							if (S_OK != m_ProvCityChurchNameChurches.OpenRowset(dc))
							{
								return UTIL::SendError(m_HttpResponse, "Error with spProvCityChurchNameChurches");
							}
							m_bFirstRow = false;
						}
						else
						{
							//ProvCityChurches
							Checked::tcsncpy_s(m_ProvCityChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvCityChurches.m_city, MAX_CITY, m_strCity, _TRUNCATE);
							if (S_OK != m_ProvCityChurches.OpenRowset(dc))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with spProvCityChurches");
							}
							if (S_OK != m_ProvCityChurches.GetApproximatePosition(NULL, NULL, &m_Rows))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with m_ProvCityChurches.GetApproximatePosition()");
							}
							m_Pages = (DBCOUNTITEM)UTIL::CalculatePages(m_Rows, m_dbPp);
							if (m_CurPage > m_Pages)
							{
								// reset scroll
								m_CurPage = 0;
							}
							m_Scroll = m_CurPage * m_dbPp;
							if (m_bPg)
							{
								if (0 != m_iPg)
								{
									if (S_OK != m_ProvCityChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvCityChurches.MoveNext()");
									}
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									m_bFirstRow = false;
								}
								else
								{
									if (S_OK != m_ProvCityChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvCityChurches.MoveNext()");
									}
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									m_bFirstRow = true;
									m_wstrFirstRow = m_ProvCityChurches.GetFirstRow(m_strProvince, m_strCity);
								}
							}
							else
							{
								m_bFirstRow = true;
								if (S_OK != m_ProvCityChurches.MoveFirst())
								{
									return UTIL::SendDBError(m_HttpResponse, "Error with spProvCityChurches.MoveFirst()");
								}
								m_wstrFirstRow = m_ProvCityChurches.GetFirstRow(m_strProvince, m_strCity);
							}
						}
					}
				}
				else
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvDenomChurchNameChurches
							Checked::tcsncpy_s(m_ProvDenomChurchNameChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvDenomChurchNameChurches.m_denomination, MAX_DENOMINATION, m_strDenomination, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvDenomChurchNameChurches.m_churchname, MAX_CHURCHNAME, m_strChurchName, _TRUNCATE);
							if (S_OK != m_ProvDenomChurchNameChurches.OpenRowset(dc))
							{
								return UTIL::SendError(m_HttpResponse, "Error with spProvDenomChurchNameChurches");
							}
							m_bFirstRow = false;
						}
						else
						{
							//ProvDenomChurches
							Checked::tcsncpy_s(m_ProvDenomChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvDenomChurches.m_denomination, MAX_DENOMINATION, m_strDenomination, _TRUNCATE);
							if (S_OK != m_ProvDenomChurches.OpenRowset(dc))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with spProvDenomChurches");
							}
							if (S_OK != m_ProvDenomChurches.GetApproximatePosition(NULL, NULL, &m_Rows))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with m_ProvDenomChurches.GetApproximatePosition()");
							}
							m_Pages = (DBCOUNTITEM)UTIL::CalculatePages(m_Rows, m_dbPp);
							if (m_CurPage > m_Pages)
							{
								// reset scroll
								m_CurPage = 0;
							}
							m_Scroll = m_CurPage * m_dbPp;
							if (m_bPg)
							{
								if (0 != m_iPg)
								{
									if (S_OK != m_ProvDenomChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvDenomChurches.MoveNext()");
									}
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									m_bFirstRow = false;
								}
								else
								{
									if (S_OK != m_ProvDenomChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvDenomChurches.MoveNext()");
									}
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									m_bFirstRow = true;
									m_wstrFirstRow = m_ProvDenomChurches.GetFirstRow(m_strProvince, m_strDenomination);
								}
							}
							else
							{
								if (S_OK != m_ProvDenomChurches.MoveFirst())
								{
									return UTIL::SendDBError(m_HttpResponse, "Error with spProvDenomChurches.MoveFirst()");
								}
								m_bFirstRow = true;
								m_wstrFirstRow = m_ProvDenomChurches.GetFirstRow(m_strProvince, m_strDenomination);
							}
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvChurchNameChurches
							Checked::tcsncpy_s(m_ProvChurchNameChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							Checked::tcsncpy_s(m_ProvChurchNameChurches.m_churchname, MAX_CHURCHNAME, m_strChurchName, _TRUNCATE);
							if (S_OK != m_ProvChurchNameChurches.OpenRowset(dc))
							{
								return UTIL::SendError(m_HttpResponse, "Error with spProvChurchNameChurches");
							}
							m_bFirstRow = false;
						}
						else
						{
							//ProvChurches
							Checked::tcsncpy_s(m_ProvChurches.m_province, MAX_PROVINCE, m_strProvince, _TRUNCATE);
							if (S_OK != m_ProvChurches.OpenRowset(dc))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with spProvChurches");
							}
							if (S_OK != m_ProvChurches.GetApproximatePosition(NULL, NULL, &m_Rows))
							{
								return UTIL::SendDBError(m_HttpResponse, "Error with m_ProvChurches.GetApproximatePosition()");
							}
							m_Pages = (DBCOUNTITEM)UTIL::CalculatePages(m_Rows, m_dbPp);
							if (m_CurPage > m_Pages)
							{
								// reset scroll
								m_CurPage = 0;
							}
							m_Scroll = m_CurPage * m_dbPp;
							if (m_bPg)
							{
								if (0 != m_iPg)
								{
									if (S_OK != m_ProvChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvChurches.MoveNext()");
									}
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									m_bFirstRow = false;
								}
								else
								{
									if (S_OK != m_ProvChurches.MoveNext(m_Scroll, true))
									{
										return UTIL::SendDBError(m_HttpResponse, "Error with spProvChurches.MoveNext()");
									}
									if (m_CurPage >= m_Pages)
										m_CurPage = 0;
									m_bFirstRow = true;
									m_wstrFirstRow = m_ProvChurches.GetFirstRow(m_strProvince);

								}
							}
							else
							{
								if (S_OK != m_ProvChurches.MoveFirst())
								{
									return UTIL::SendDBError(m_HttpResponse, "Error with spProvChurches.MoveFirst()");
								}
								m_bFirstRow = true;
								m_wstrFirstRow = m_ProvChurches.GetFirstRow(m_strProvince);
							}
						}
					}
				}
			}
		}
		if (S_OK != m_AllProvs.OpenRowset(dc))
		{
			return UTIL::SendDBError(m_HttpResponse, "Error with spAllProvs");
		}

		return /*(m_bAssetsPushed) ? HTTP_SUCCESS_NO_CACHE : */HTTP_SUCCESS;
		
	}
	CspAllProvs m_AllProvs;
	CspProvRegions m_ProvRegions;
	CspProvRegCities m_ProvRegCities;
	CspProvRegCityDenoms m_ProvRegCityDenoms;
	CspChurchDetails m_ChurchDetails;
	CspProvRegCityDenomChurchNameChurches m_ProvRegCityDenomChurchNameChurches;
	CspProvRegCityDenomChurches m_ProvRegCityDenomChurches;
	CspProvRegCityChurchNameChurches m_ProvRegCityChurchNameChurches;
	CspProvRegCityChurches m_ProvRegCityChurches;
	CspProvRegDenomChurchNameChurches m_ProvRegDenomChurchNameChurches;
	CspProvRegDenomChurches m_ProvRegDenomChurches;
	CspProvRegChurchNameChurches m_ProvRegChurchNameChurches;
	CspProvRegChurches m_ProvRegChurches;
	CspProvCityDenomChurchNameChurches m_ProvCityDenomChurchNameChurches;
	CspProvCityDenomChurches m_ProvCityDenomChurches;
	CspProvCityChurchNameChurches m_ProvCityChurchNameChurches;
	CspProvDenomChurchNameChurches m_ProvDenomChurchNameChurches;
	CspProvDenomChurches m_ProvDenomChurches;
	CspProvChurchNameChurches m_ProvChurchNameChurches;
	CspProvCityChurches m_ProvCityChurches;
	CspProvChurches m_ProvChurches;
	CspRegCityDenomChurches m_RegCityDenomChurches;
	CspRegCityChurches m_RegCityChurches;
	CspRegDenomChurches m_RegDenomChurches;
	CspRegChurches m_RegChurches;
	CspCityDenomChurches m_CityDenomChurches;
	CspCityChurches m_CityChurches;
	CspDenomChurches m_DenomChurches;
	CspConfirmUser m_ConfirmUser;
	DWORD CChurchSearchHandler::MaxFormSize()
	{
		// Allow at most 271 bytes of data
		return 271;
	}
	BEGIN_REPLACEMENT_METHOD_MAP(CChurchSearchHandler)
		REPLACEMENT_METHOD_ENTRY("Alternatives", OnAlternatives)
		REPLACEMENT_METHOD_ENTRY("ImgMap", OnImgMap)
		REPLACEMENT_METHOD_ENTRY("CSKeywords", OnMetas)
		REPLACEMENT_METHOD_ENTRY("ImgAlt", OnImgAlt)
		REPLACEMENT_METHOD_ENTRY("Img", OnImg)
		REPLACEMENT_METHOD_ENTRY("bImg", OnbImg)
		REPLACEMENT_METHOD_ENTRY("bDetails", OnbDetails)
		REPLACEMENT_METHOD_ENTRY("Rowscroll", OnRowscroll)
		REPLACEMENT_METHOD_ENTRY("Status", OnStatus)
		REPLACEMENT_METHOD_ENTRY("Church", OnChurch)
		REPLACEMENT_METHOD_ENTRY("HasMoreChurches", OnHasMoreChurches)
		REPLACEMENT_METHOD_ENTRY("OneTimeInit", OnOneTimeInit)
		REPLACEMENT_METHOD_ENTRY("Redirect", OnRedirect)
		REPLACEMENT_METHOD_ENTRY("Selector", OnSelector)
		REPLACEMENT_METHOD_ENTRY("bPerPage", OnbPerPage)
		REPLACEMENT_METHOD_ENTRY("bDenomination", OnbDenomination)
		REPLACEMENT_METHOD_ENTRY("bCity", OnbCity)
		REPLACEMENT_METHOD_ENTRY("bRegion", OnbRegion)
		REPLACEMENT_METHOD_ENTRY("bProvince", OnbProvince)
		REPLACEMENT_METHOD_ENTRY("bOldBrowser", OnbOldBrowser)
		REPLACEMENT_METHOD_ENTRY("HasMoreCities", OnHasMoreCities)
		REPLACEMENT_METHOD_ENTRY("HasMoreRegions", OnHasMoreRegions)
		REPLACEMENT_METHOD_ENTRY("HasMoreProvinces", OnHasMoreProvinces)
		REPLACEMENT_METHOD_ENTRY("HasMoreDenoms", OnHasMoreDenoms)
		REPLACEMENT_METHOD_ENTRY("PerPage", OnPerPage)
		REPLACEMENT_METHOD_ENTRY("ChurchName", OnChurchName)
		REPLACEMENT_METHOD_ENTRY("Denomination", OnDenomination)
		REPLACEMENT_METHOD_ENTRY("City", OnCity)
		REPLACEMENT_METHOD_ENTRY("Region", OnRegion)
		REPLACEMENT_METHOD_ENTRY("Province", OnProvince)
		REPLACEMENT_METHOD_ENTRY("title", OnTitle)
		REPLACEMENT_METHOD_ENTRY("Status", OnStatus)
		REPLACEMENT_METHOD_ENTRY("HasMoreChurches", OnHasMoreChurches)
		REPLACEMENT_METHOD_ENTRY("Church", OnChurch)
		REPLACEMENT_METHOD_ENTRY("HeadMethodCatch", OnHeadMethodCatch)
		REPLACEMENT_METHOD_ENTRY("AppendETagCatch", OnAppendETagCatch)
	END_REPLACEMENT_METHOD_MAP()
	/*BOOL CachePage()
	{
	return TRUE;
	}*/
protected:
	CComPtr<IBrowserCaps> m_spBrowserCaps;
	CComPtr<IBrowserCapsSvc> m_spBrowserCapsSvc;
	//implements html <title> tag contents
	HTTP_CODE OnTitle(void)
	{
		if (m_bProvince)
		{
			if (m_bRegion)
			{
				if (m_bCity)
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvRegCityDenomChurchNameChurches
							m_HttpResponse << m_strChurchName << " is a "
								<< m_strDenomination << " church located in "
								<< m_strCity << ", " << m_strRegion << ", "
								<< UTIL::GetNamedProvince(m_strProvince);
						}
						else
						{
							//ProvRegCityDenomChurches
							m_HttpResponse << m_strDenomination << " churches located in "
								<< m_strCity << ", " << m_strRegion << ", "
								<< UTIL::GetNamedProvince(m_strProvince);
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvRegCityChurchNameChurches
							m_HttpResponse << m_strChurchName << " is a church located in "
								<< m_strCity << ", " << m_strRegion << ", "
								<< UTIL::GetNamedProvince(m_strProvince);
						}
						else
						{
							//ProvRegCityChurches
							m_HttpResponse << "Churches located in "
								<< m_strCity << ", " << m_strRegion << ", "
								<< UTIL::GetNamedProvince(m_strProvince);
						}
					}
				}
				else
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvRegDenomChurchNameChurches
							m_HttpResponse << m_strChurchName << " is a "
								<< m_strDenomination << " church located in "
								<< m_strRegion << ", " << UTIL::GetNamedProvince(m_strProvince);
						}
						else
						{
							//ProvRegDenomChurches
							m_HttpResponse << m_strDenomination << " churches located in "
								<< m_strRegion << ", " << UTIL::GetNamedProvince(m_strProvince);
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvRegChurchNameChurches
							m_HttpResponse << m_strChurchName << " is a church located in "
								<< m_strRegion << ", " << UTIL::GetNamedProvince(m_strProvince);
						}
						else
						{
							//ProvRegChurches
							m_HttpResponse << "Churches located in "
								<< m_strRegion << ", " << UTIL::GetNamedProvince(m_strProvince);
						}
					}
				}
			}
			else
			{
				if (m_bCity)
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvCityDenomChurchNameChurches
							m_HttpResponse << m_strChurchName << " is a "
								<< m_strDenomination << " church located in "
								<< m_strCity << ", " << UTIL::GetNamedProvince(m_strProvince);
						}
						else
						{
							//ProvCityDenomChurches
							m_HttpResponse << m_strDenomination << " churches located in "
								<< m_strCity << ", " << UTIL::GetNamedProvince(m_strProvince);
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvCityChurchNameChurches
							m_HttpResponse << m_strChurchName << " is a church located in "
								<< m_strCity << ", " << UTIL::GetNamedProvince(m_strProvince);
						}
						else
						{
							//ProvCityChurches
							m_HttpResponse << "Churches located in "
								<< m_strCity << ", " << UTIL::GetNamedProvince(m_strProvince);
						}
					}
				}
				else
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvDenomChurchNameChurches
							m_HttpResponse << m_strChurchName << " is a "
								<< m_strDenomination << " church located in "
								<< UTIL::GetNamedProvince(m_strProvince);
						}
						else
						{
							//ProvDenomChurches
							m_HttpResponse << m_strDenomination << " churches located in "
								<< UTIL::GetNamedProvince(m_strProvince);
						}
					}
					else
					{
						if (0 != m_bChurchName)
						{
							//ProvChurchNameChurches
							m_HttpResponse << m_strChurchName << " is a church located in "
								<< UTIL::GetNamedProvince(m_strProvince);
						}
						else
						{
							//ProvChurches
							m_HttpResponse << "Churches located in "
								<< UTIL::GetNamedProvince(m_strProvince);
						}
					}
				}
			}
		}
		return HTTP_SUCCESS;
	}
	//whether or not to display a list of provinces in a drop-down listbox
	HTTP_CODE OnbProvince(void)
	{
		if (m_bId)
			return HTTP_S_FALSE;
		if (m_bRedirect)
			return HTTP_S_FALSE;
		return HTTP_SUCCESS;
	}
	//until end-of-rowset reached
	HTTP_CODE OnHasMoreProvinces(void)
	{
		HRESULT hr;
		hr = m_AllProvs.MoveNext();
		return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
	}
	//fill up the listbox if previous conditions are true
	HTTP_CODE OnProvince(void)
	{
		if (m_bProvince)
		{
			if (0 != lstrcmp(m_AllProvs.m_Province, m_strProvince))
			{
				m_HttpResponse << "<option class=\"clsProvince\" ";
				if (4 < m_AllProvs.m_dwProvinceLength)
				{
					// do this for bermuda
					m_HttpResponse << ">" << m_AllProvs.m_Province;
				}
				else
				{
					m_HttpResponse << "value=\"" << m_AllProvs.m_Province << "\">"
						<< UTIL::GetNamedProvince(m_AllProvs.m_Province);
				}
				m_HttpResponse << "</option>";
			}
			else
			{
				m_HttpResponse << "<option class=\"clsProvince\" ";
				if (4 < m_AllProvs.m_dwProvinceLength)
				{
					// bermuda, and other countries
					m_HttpResponse << "selected>"
						<< m_AllProvs.m_Province;
				}
				else
				{
					m_HttpResponse << "selected value=\""
						<< m_AllProvs.m_Province << "\">"
						<< UTIL::GetNamedProvince(m_AllProvs.m_Province);
				}
				m_HttpResponse << "</option>";
			}
		}
		else
		{
			m_HttpResponse << "<option class=\"clsProvince\" ";
			if (4 < m_AllProvs.m_dwProvinceLength)
			{
				// do this for bermuda
				m_HttpResponse << ">" << m_AllProvs.m_Province;
			}
			else
			{
				m_HttpResponse << "value=\"" << m_AllProvs.m_Province << "\">"
					<< UTIL::GetNamedProvince(m_AllProvs.m_Province);
			}
			m_HttpResponse << "</option>";
		}
		return HTTP_SUCCESS;
	}

	HTTP_CODE OnbRegion(void)
	{
		if (m_bId)
			return HTTP_S_FALSE;
		if (m_bRedirect)
			return HTTP_S_FALSE;
		if (m_bProvince)
			return HTTP_SUCCESS;
		else
			return HTTP_S_FALSE;
	}

	HTTP_CODE OnHasMoreRegions(void)
	{
		HRESULT hr = S_FALSE;
		hr = m_ProvRegions.MoveNext();
		return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
	}

	HTTP_CODE OnRegion(void)
	{
		if (0 != m_bRegion)
		{
			if (0 != lstrcmp(m_strRegion, m_ProvRegions.m_Region))
			{
				m_HttpResponse << "<option class=\"clsRegion\">" << m_ProvRegions.m_Region << "</option>";
			}
			else
			{
				m_HttpResponse << "<option class=\"clsRegion\" selected>" << m_ProvRegions.m_Region << "</option>";
			}
		}
		else
		{
			m_HttpResponse << "<option class=\"clsRegion\">" << m_ProvRegions.m_Region << "</option>";
		}
		return HTTP_SUCCESS;
	}

	HTTP_CODE OnbCity(void)
	{
		if (m_bId)
			return HTTP_S_FALSE;
		if (m_bRedirect)
			return HTTP_S_FALSE;
		if (m_bProvince)
		{
			if (m_bRegion)
			{
				return HTTP_SUCCESS;
			}
			else
			{
				return HTTP_S_FALSE;
			}
		}
		else
		{
			return HTTP_S_FALSE;
		}
	}

	HTTP_CODE OnHasMoreCities(void)
	{
		HRESULT hr = S_FALSE;
		hr = m_ProvRegCities.MoveNext();
		return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
	}

	HTTP_CODE OnCity(void)
	{
		if (m_bCity)
		{
			if (0 != lstrcmp(m_strCity, m_ProvRegCities.m_City))
			{
				m_HttpResponse << "<option class=\"clsCity\">" << m_ProvRegCities.m_City << "</option>";
			}
			else
			{
				m_HttpResponse << "<option class=\"clsCity\" selected>" << m_ProvRegCities.m_City << "</option>";
			}
		}
		else
		{
			m_HttpResponse << "<option class=\"clsCity\">" << m_ProvRegCities.m_City << "</option>";
		}
		return HTTP_SUCCESS;
	}
	
	HTTP_CODE OnbDenomination(void)
	{
		if (m_bId)
			return HTTP_S_FALSE;
		if (m_bRedirect)
			return HTTP_S_FALSE;
		if (m_bProvince)
		{
			if (m_bRegion)
			{
				if (m_bCity)
				{
					return HTTP_SUCCESS;
				}
				else
				{
					return HTTP_S_FALSE;
				}
			}
			else
			{
				return HTTP_S_FALSE;
			}
		}
		else
		{
			return HTTP_S_FALSE;
		}
	}

	HTTP_CODE OnHasMoreDenoms(void)
	{
		HRESULT hr = S_FALSE;
		hr = m_ProvRegCityDenoms.MoveNext();
		return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
	}

	HTTP_CODE OnDenomination(void)
	{
		if (m_bDenomination)
		{
			if (0 != lstrcmp(m_ProvRegCityDenoms.m_Denomination, m_strDenomination))
			{
				m_HttpResponse << "<option class=\"clsDenomination\">" << m_ProvRegCityDenoms.m_Denomination << "</option>";
			}
			else
			{
				m_HttpResponse << "<option class=\"clsDenomination\" selected>" << m_ProvRegCityDenoms.m_Denomination << "</option>";
			}
		}
		else
		{
			m_HttpResponse << "<option class=\"clsDenomination\">" << m_ProvRegCityDenoms.m_Denomination << "</option>";
		}
		return HTTP_SUCCESS;
	}

	HTTP_CODE OnbPerPage(void)
	{
		if (m_bId)
			return HTTP_S_FALSE;
		if (m_bRedirect)
			return HTTP_S_FALSE;
		return HTTP_SUCCESS;
	}
	//when outputting records per-page listbox, ensure that any previous captured listbox selected is reflected in the new output
	HTTP_CODE OnPerPage(void)
	{
		if (m_bPp)
		{
			if (m_dbPp == 10)
			{
				m_HttpResponse << OPTION_10;
			}
			if (m_dbPp == 20)
			{
				m_HttpResponse << OPTION_20;
			}
			if (m_dbPp == 25)
			{
				m_HttpResponse << OPTION_25;
			}
			if (m_dbPp == 30)
			{
				m_HttpResponse << OPTION_30;
			}
			if (m_dbPp == 35)
			{
				m_HttpResponse << OPTION_35;
			}
			if (m_dbPp == 40)
			{
				m_HttpResponse << OPTION_40;
			}
			if (m_dbPp == 50)
			{
				m_HttpResponse << OPTION_50;
			}
			if (m_dbPp == 100)
			{
				m_HttpResponse << OPTION_100;
			}
		}
		else
		{
			m_HttpResponse << "<option>10</option><option>20</option><option>25</option>\
<option>30</option><option>35</option><option>40</option><option>50</option><option>100</option>";
		}
		return HTTP_SUCCESS;
	}

	HTTP_CODE OnChurchName(void)
	{
		m_HttpResponse << "<option>" << m_strChurchName << "</option>";
		return HTTP_SUCCESS;
	}
	//replacement tag that contains abunch of client-side script to emulate asp.net's doPostBack functionality
	HTTP_CODE OnSelector(void)
	{
		if (m_bProvince)
		{
			if (m_bRegion)
			{
				if (m_bCity)
				{
					if (m_bDenomination)
					{
						m_HttpResponse << SELECT_PROVINCE_CLR_ALL <<
							SELECT_REGION_CLR_CITY_DENOM << SELECT_CITY_CLR_DENOM <<
							SELECT_DENOM;
						//if (0 != m_bPp)
						m_HttpResponse << SELECT_PP_CLR_DENOM;
						m_HttpResponse << SELECT_NAME_CLR_DENOM;
					}
					else
					{
						m_HttpResponse << SELECT_PROVINCE_CLR_ALL <<
							SELECT_REGION_CLR_CITY_DENOM << SELECT_CITY_CLR_DENOM <<
							SELECT_DENOM;
						//if (0 != m_bPp)
						m_HttpResponse << SELECT_PP_CLR_DENOM;
						m_HttpResponse << SELECT_NAME_CLR_DENOM;
					}
				}
				else
				{
					m_HttpResponse << SELECT_PROVINCE_CLR_CITY_REGION <<
						SELECT_REGION_CLR_CITY << SELECT_CITY;
					//if (0 != m_bPp)
					m_HttpResponse << SELECT_PP_CLR_CITY;
					m_HttpResponse << SELECT_NAME_CLR_CITY;
				}
			}
			else
			{
				m_HttpResponse << SELECT_PROVINCE_CLR_REGION << SELECT_REGION;
				//if (0 != m_bPp)
				m_HttpResponse << SELECT_PP_CLR_REGION;
				m_HttpResponse << SELECT_NAME_CLR_REGION;
			}
		}
		else
		{
			m_HttpResponse << SELECT_PROVINCE;
		}
		return HTTP_SUCCESS;
	}
	//the first page simply displays a map, so give the user the option to set their desired records per-page
	HTTP_CODE OnOneTimeInit(void)
	{
		if (m_bProvince)
			m_HttpResponse << "onchange=\"SelectPp()\"";
		return HTTP_SUCCESS;
	}
	//status bar is actually used to outputting table headers as well as the first row of a table if there are multiple results
	HTTP_CODE OnStatus(void)
	{
		if (m_bId)
		{
			return HTTP_SUCCESS;
		}
		if (m_bRedirect)
		{
			if (m_bBypassParam)
			{
				//do nothing, falls through to the code below
			}
			else
				return HTTP_SUCCESS;
		}
		if (m_bProvince)
		{
			if (m_bRegion)
			{
				if (m_bCity)
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvRegCityDenomChurchNameChurches
							m_HttpResponse << PROV_REG_CITY_DENOM_CHURCHNAME_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
						else
						{
							//ProvRegCityDenomChurches
							m_HttpResponse << PROV_REG_CITY_DENOM_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvRegCityChurchNameChurches
							m_HttpResponse << PROV_REG_CITY_CHURCHNAME_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
						else
						{
							//ProvRegCityChurches
							m_HttpResponse << PROV_REG_CITY_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
					}
				}
				else
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvRegDenomChurchNameChurches
							m_HttpResponse << PROV_REG_DENOM_CHURCHNAME_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
						else
						{
							//ProvRegDenomChurches
							m_HttpResponse << PROV_REG_DENOM_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvRegChurchNameChurches
							m_HttpResponse << PROV_REG_CHURCHNAME_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
						else
						{
							//ProvRegChurches
							m_HttpResponse << PROV_REG_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
					}
				}
			}
			else
			{
				if (m_bCity)
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvCityDenomChurchNameChurches
							m_HttpResponse << PROV_CITY_DENOM_CHURCHNAME_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
						else
						{
							//ProvCityDenomChurches
							m_HttpResponse << PROV_CITY_DENOM_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvCityChurchNameChurches
							m_HttpResponse << PROV_CITY_CHURCHNAME_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
						else
						{
							//ProvCityChurches
							m_HttpResponse << PROV_CITY_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
					}
				}
				else
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvDenomChurchNameChurches
							m_HttpResponse << PROV_DENOM_CHURCHNAME_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
						else
						{
							//ProvDenomChurches
							m_HttpResponse << PROV_DENOM_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvChurchNameChurches
							m_HttpResponse << PROV_CHURCHNAME_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
						else
						{
							//ProvChurches
							m_HttpResponse << PROV_CHURCHES_HEADER;
							if (m_bFirstRow)
							{
								m_HttpResponse << m_wstrFirstRow;
							}
						}
					}
				}
			}
		}
		else
		{
			m_HttpResponse << "Please select a province.";
		}
		return HTTP_SUCCESS;
	}
	//the critical function where most of the application logic in this handler takes place
	//the primary purpose of this method is to determine whether the end-of-rowset has been reached 
	//the secondary purpose is to limit output to the per-page value
	//the third purpose is to redirect users to a details page if there is only one result (m_bRedirect)
	//the fourth purpose is to display the one result after the page has been redirected (m_bId)
	//Not only does it create redirects to view individual churches, but it also will redirect to pages if
	//there is only one selection in any one of the selectors (m_bBypassParam)
	HTTP_CODE OnHasMoreChurches(void)
	{
		HRESULT hr = HTTP_S_FALSE;
		if (m_bId)
		{
			if (!m_bConfirmGuid)
			{
				//it is not possible for the database to contain 2 identical ChurchID's
				hr = m_ChurchDetails.MoveNext();
				//so this will always execute once (and only once)
				return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
			}
			else
			{
				return HTTP_S_FALSE;//break out of loop if this is a user confirmation request
			}
		}
		//begin main program control loop
		if (m_bProvince)
		{
			if (m_bRegion)
			{
				if (m_bCity)
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvRegCityDenomChurchNameChurches
							//Since there may be several rows returned by the db,
							//we need to find exactly the one the user wants, while
							//maintaining the ability to display several rows at once.
							hr = m_ProvRegCityDenomChurchNameChurches.MoveNext();
							if (hr == DB_S_ENDOFROWSET && m_CurRow == 1)
							{
								// if only 1 row is returned, we know we have a specific church to show
								m_bRedirect = true;//redirect to the specific church
								m_bBypassParam = false;
							}
							else if (FAILED(hr))
							{
								m_bRedirect = false;//some other weird error occured, so prevent a redirect
							}
							else if (0 == m_strChurchName.CompareNoCase(m_ProvRegCityDenomChurchNameChurches.m_ChurchName))
							{
								m_bRedirect = true;
								m_bBypassParam = false;
								m_iId = m_ProvRegCityDenomChurchNameChurches.GetChurchID();
								//break out of the loop
								return HTTP_S_FALSE;
							}
							else
							{
								//multiple results
								m_bRedirect = false;
								m_CurRow++;//iteration
								m_iId = m_ProvRegCityDenomChurchNameChurches.GetChurchID();
							}
							//no paging control allowed on full text related searches, this is a limitation of our software
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
						}
						else
						{
							//ProvRegCityDenomChurches
							hr = m_ProvRegCityDenomChurches.MoveNext();
							m_CurRow++;
							if (m_CurRow >= m_dbPp + 1)
								hr = HTTP_S_FALSE;
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvRegCityChurchNameChurches
							hr = m_ProvRegCityChurchNameChurches.MoveNext();
							if (hr == DB_S_ENDOFROWSET && m_CurRow == 1)
							{
								// if only 1 row is returned, we know we have a specific church to show
								m_bRedirect = true;
								m_bBypassParam = false;
							}
							else if (FAILED(hr))
							{
								m_bRedirect = false;
							}
							else if (0 == m_strChurchName.CompareNoCase(m_ProvRegCityChurchNameChurches.m_ChurchName))
							{
								m_bRedirect = true;
								m_bBypassParam = false;
								m_iId = m_ProvRegCityChurchNameChurches.GetChurchID();
								//break out of the loop
								return HTTP_S_FALSE;
							}
							else
							{
								m_bRedirect = false;
								m_CurRow++;
								m_iId = m_ProvRegCityChurchNameChurches.GetChurchID();
							}
							//no paging control allowed on full text related searches
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
						}
						else
						{
							//ProvRegCityChurches
							hr = m_ProvRegCityChurches.MoveNext();
							m_CurRow++;
							if (m_CurRow >= m_dbPp + 1)
								hr = HTTP_S_FALSE;
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
						}
					}
				}
				else
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvRegDenomChurchNameChurches
							hr = m_ProvRegDenomChurchNameChurches.MoveNext();
							if (hr == DB_S_ENDOFROWSET && m_CurRow == 1)
							{
								// if only 1 row is returned, we know we have a specific church to show
								m_bRedirect = true;
								m_bBypassParam = false;
							}
							else if (FAILED(hr))
							{
								m_bRedirect = false;
							}
							else if (0 == m_strChurchName.CompareNoCase(m_ProvRegDenomChurchNameChurches.m_ChurchName))
							{
								m_bRedirect = true;
								m_bBypassParam = false;
								m_iId = m_ProvRegDenomChurchNameChurches.GetChurchID();
								//break out of the loop
								return HTTP_S_FALSE;
							}
							else
							{
								m_bRedirect = false;
								m_CurRow++;
								m_iId = m_ProvRegDenomChurchNameChurches.GetChurchID();
							}
							//no paging control allowed on full text related searches
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;

						}
						else
						{
							//ProvRegDenomChurches
							hr = m_ProvRegDenomChurches.MoveNext();
							m_CurRow++;
							if (m_CurRow >= m_dbPp + 1)
								hr = HTTP_S_FALSE;
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvRegChurchNameChurches
							hr = m_ProvRegChurchNameChurches.MoveNext();
							if (hr == DB_S_ENDOFROWSET && m_CurRow == 1)
							{
								// if only 1 row is returned, we know we have a specific church to show
								m_bRedirect = true;
								m_bBypassParam = false;
							}
							else if (FAILED(hr))
							{
								m_bRedirect = false;
							}
							else if (0 == m_strChurchName.CompareNoCase(m_ProvRegChurchNameChurches.m_ChurchName))
							{
								m_bRedirect = true;
								m_bBypassParam = false;
								m_iId = m_ProvRegChurchNameChurches.GetChurchID();
								//break out of the loop
								return HTTP_S_FALSE;
							}
							else
							{
								m_bRedirect = false;
								m_CurRow++;
								m_iId = m_ProvRegChurchNameChurches.GetChurchID();
							}
							//no paging control allowed on full text related searches
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
						}
						else
						{
							//ProvRegChurches
							hr = m_ProvRegChurches.MoveNext();
							m_CurRow++;
							if (m_CurRow >= m_dbPp + 1)
								hr = HTTP_S_FALSE;
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
						}
					}
				}
			}
			else
			{
				if (m_bCity)
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvCityDenomChurchNameChurches
							hr = m_ProvCityDenomChurchNameChurches.MoveNext();
							if (hr == DB_S_ENDOFROWSET && m_CurRow == 1)
							{
								// if only 1 row is returned, we know we have a specific church to show
								m_bRedirect = true;
								m_bBypassParam = false;
							}
							else if (FAILED(hr))
							{
								m_bRedirect = false;
							}
							else if (0 == m_strChurchName.CompareNoCase(m_ProvCityDenomChurchNameChurches.m_ChurchName))
							{
								m_bRedirect = true;
								m_bBypassParam = false;
								m_iId = m_ProvCityDenomChurchNameChurches.GetChurchID();
								//break out of the loop
								return HTTP_S_FALSE;
							}
							else
							{
								m_bRedirect = false;
								m_CurRow++;
								m_iId = m_ProvCityDenomChurchNameChurches.GetChurchID();
							}
							//no paging control allowed on full text related searches
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
						}
						else
						{
							//ProvCityDenomChurches
							hr = m_ProvCityDenomChurches.MoveNext();
							m_CurRow++;
							if (m_CurRow >= m_dbPp + 1)
								hr = HTTP_S_FALSE;
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvCityChurchNameChurches
							hr = m_ProvCityChurchNameChurches.MoveNext();
							if (hr == DB_S_ENDOFROWSET && m_CurRow == 1)
							{
								// if only 1 row is returned, we know we have a specific church to show
								m_bRedirect = true;
								m_bBypassParam = false;
							}
							else if (FAILED(hr))
							{
								m_bRedirect = false;
							}
							else if (0 == m_strChurchName.CompareNoCase(m_ProvCityChurchNameChurches.m_ChurchName))
							{
								m_bRedirect = true;
								m_bBypassParam = false;
								m_iId = m_ProvCityChurchNameChurches.GetChurchID();
								//break out of the loop
								return HTTP_S_FALSE;
							}
							else
							{
								m_bRedirect = false;
								m_CurRow++;
								m_iId = m_ProvCityChurchNameChurches.GetChurchID();
							}
							//no paging control allowed on full text related searches
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
						}
						else
						{
							//ProvCityChurches
							hr = m_ProvCityChurches.MoveNext();
							m_CurRow++;
							if (m_CurRow >= m_dbPp + 1)
								hr = HTTP_S_FALSE;
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
						}
					}
				}
				else
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvDenomChurchNameChurches
							hr = m_ProvDenomChurchNameChurches.MoveNext();
							if (hr == DB_S_ENDOFROWSET && m_CurRow == 1)
							{
								// if only 1 row is returned, we know we have a specific church to show
								m_bRedirect = true;
								m_bBypassParam = false;
							}
							else if (FAILED(hr))
							{
								m_bRedirect = false;
							}
							else if (0 == m_strChurchName.CompareNoCase(m_ProvDenomChurchNameChurches.m_ChurchName))
							{
								m_bRedirect = true;
								m_bBypassParam = false;
								m_iId = m_ProvDenomChurchNameChurches.GetChurchID();
								//break out of the loop
								return HTTP_S_FALSE;
							}
							else
							{
								m_bRedirect = false;
								m_CurRow++;
								m_iId = m_ProvDenomChurchNameChurches.GetChurchID();
							}
							//no paging control allowed on full text related searches
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
						}
						else
						{
							//ProvDenomChurches
							hr = m_ProvDenomChurches.MoveNext();
							m_CurRow++;
							if (m_CurRow >= m_dbPp + 1)
								hr = HTTP_S_FALSE;
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
						}
					}
					else
					{
						if (m_bChurchName)
						{
							hr = m_ProvChurchNameChurches.MoveNext();
							if (hr == DB_S_ENDOFROWSET && m_CurRow == 1)
							{
								// if only 1 row is returned, we know we have a specific church to show
								m_bRedirect = true;
								m_bBypassParam = false;
							}
							else if (FAILED(hr))
							{
								m_bRedirect = false;
							}
							else if (0 == m_strChurchName.CompareNoCase(m_ProvChurchNameChurches.m_ChurchName))
							{
								m_bRedirect = true;
								m_bBypassParam = false;
								m_iId = m_ProvChurchNameChurches.GetChurchID();
								//break out of the loop
								return HTTP_S_FALSE;
							}
							else
							{
								m_bRedirect = false;
								m_CurRow++;
								m_iId = m_ProvChurchNameChurches.GetChurchID();
							}
							//no paging control allowed on full text related searches
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
						}
						else
						{
							//ProvChurches
							hr = m_ProvChurches.MoveNext();
							m_CurRow++;
							if (m_CurRow >= m_dbPp + 1)
								hr = HTTP_S_FALSE;
							return (hr == S_OK) ? HTTP_SUCCESS : HTTP_S_FALSE;
						}
					}
				}
			}
		}
		return hr;
	}
	//used to extract the data from the current row in the rowset to the html stream
	HTTP_CODE OnChurch(void)
	{
		if (m_bId)
		{
			m_HttpResponse << m_ChurchDetails.GetFirstTable();
			m_strProvince.SetString(m_ChurchDetails.m_Province);
			m_strRegion.SetString(m_ChurchDetails.m_Region);
			m_strCity.SetString(m_ChurchDetails.m_City);
			m_strDenomination.SetString(m_ChurchDetails.m_Denomination);
			m_strChurchName.SetString(m_ChurchDetails.m_ChurchName);
			return HTTP_SUCCESS;
		}
		if (m_bRedirect)
		{
			if (m_bBypassParam)
			{
				//do nothing (fallthrough)
			}
			else
				return HTTP_SUCCESS;
		}
		if (m_bProvince)
		{
			if (m_bRegion)
			{
				if (m_bCity)
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvRegCityDenomChurchNameChurches 
							//PROV_REG_CITY_DENOM_CHURCHNAME_CHURCHES_HEADER
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince << "&Region=" << m_strRegion << "&City="
								<< m_strCity << "&Denomination=" << m_strDenomination
								<< "&ChurchName=" << m_ProvRegCityDenomChurchNameChurches.m_ChurchName
								<< "\">" << m_ProvRegCityDenomChurchNameChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvRegCityDenomChurchNameChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvRegCityDenomChurchNameChurches.m_Address << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvRegCityDenomChurchNameChurches.m_dwPostalCodeStatus)
							{
								//it's possible that mapquest doesn't support utf8 encoded parameters...
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvRegCityDenomChurchNameChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvRegCityDenomChurchNameChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_strCity, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvRegCityDenomChurchNameChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END << "</td><td class=\"chCity\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chCity\">";
							}

							m_HttpResponse << m_strCity << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvRegCityDenomChurchNameChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;//if the database record is new, ensure that the last-modified header reflects this
							}
						}
						else
						{
							//ProvRegCityDenomChurches
							//PROV_REG_CITY_DENOM_CHURCHES_HEADER
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince << "&Region=" << m_strRegion << "&City="
								<< m_strCity << "&Denomination=" << m_strDenomination
								<< "&ChurchName=" << m_ProvRegCityDenomChurches.m_ChurchName
								<< "\">" << m_ProvRegCityDenomChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvRegCityDenomChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvRegCityDenomChurches.m_Address << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvRegCityDenomChurches.m_dwPostalCodeStatus)
							{
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvRegCityDenomChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvRegCityDenomChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_strCity, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvRegCityDenomChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END << "</td><td class=\"chCity\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chCity\">";
							}
							m_HttpResponse << m_strCity << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvRegCityDenomChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;
							}
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvRegCityChurchNameChurches
							//PROV_REG_CITY_CHURCHNAME_CHURCHES_HEADER
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince << "&Region=" << m_strRegion << "&City="
								<< m_strCity << "&ChurchName="
								<< m_ProvRegCityChurchNameChurches.m_ChurchName
								<< "\">" << m_ProvRegCityChurchNameChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvRegCityChurchNameChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvRegCityChurchNameChurches.m_Address << "</td><td class=\"chMail\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chMail\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvRegCityChurchNameChurches.m_dwMailStatus)
							{
								m_HttpResponse << m_ProvRegCityChurchNameChurches.m_Mail << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvRegCityChurchNameChurches.m_dwPostalCodeStatus)
							{
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvRegCityChurchNameChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvRegCityChurchNameChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_strCity, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvRegCityChurchNameChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END << "</td><td class=\"chCity\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chCity\">";
							}
							m_HttpResponse << m_strCity << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvRegCityChurchNameChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;
							}
						}
						else
						{
							//ProvRegCityChurches
							//PROV_REG_CITY_CHURCHES_HEADER
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince << "&Region=" << m_strRegion << "&City="
								<< m_strCity << "&ChurchName=" << m_ProvRegCityChurches.m_ChurchName
								<< "\">" << m_ProvRegCityChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvRegCityChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvRegCityChurches.m_Address << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvRegCityChurches.m_dwPostalCodeStatus)
							{
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvRegCityChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvRegCityChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_strCity, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvRegCityChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END << "</td><td class=\"chDenomination\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chDenomination\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvRegCityChurches.m_dwDenominationStatus)
							{
								m_HttpResponse << m_ProvRegCityChurches.m_Denomination << "</td><td class=\"chDenomination\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chDenomination\">";
							}
							m_HttpResponse << m_strCity << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvRegCityChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;
							}
						}
					}
				}
				else
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvRegDenomChurchNameChurches
							//PROV_REG_DENOM_CHURCHNAME_CHURCHES_HEADER
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince << "&Region=" << m_strRegion;
							if (DBSTATUS_S_ISNULL != m_ProvRegDenomChurchNameChurches.m_dwCityStatus)
							{
								m_HttpResponse << "&City=" << m_ProvRegDenomChurchNameChurches.m_City;
							}
							m_HttpResponse << "&Denomination=" << m_strDenomination
								<< "&ChurchName=" << m_ProvRegDenomChurchNameChurches.m_ChurchName
								<< "\">" << m_ProvRegDenomChurchNameChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvRegDenomChurchNameChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvRegDenomChurchNameChurches.m_Address << "</td><td class=\"chMail\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chMail\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvRegDenomChurchNameChurches.m_dwMailStatus)
							{
								m_HttpResponse << m_ProvRegDenomChurchNameChurches.m_Mail << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvRegDenomChurchNameChurches.m_dwPostalCodeStatus)
							{
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvRegDenomChurchNameChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvRegDenomChurchNameChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_strCity, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvRegDenomChurchNameChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END << "</td><td class=\"chCity\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chCity\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvRegDenomChurchNameChurches.m_dwCityStatus)
							{
								m_HttpResponse << m_ProvRegDenomChurchNameChurches.m_City;
							}
							m_HttpResponse << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvRegDenomChurchNameChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;
							}
						}
						else
						{
							//ProvRegDenomChurches
							//PROV_REG_DENOM_CHURCHES_HEADER
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince << "&Region=" << m_strRegion;
							if (DBSTATUS_S_ISNULL != m_ProvRegDenomChurches.m_dwCityStatus)
							{
								m_HttpResponse << "&City=" << m_ProvRegDenomChurches.m_City;
							}
							m_HttpResponse << "&ChurchName=" << m_ProvRegDenomChurches.m_ChurchName
								<< "\">" << m_ProvRegDenomChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvRegDenomChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvRegDenomChurches.m_Address << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvRegDenomChurches.m_dwPostalCodeStatus)
							{
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvRegDenomChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvRegDenomChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_strCity, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvRegDenomChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END << "</td><td class=\"chCity\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chCity\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvRegDenomChurches.m_dwCityStatus)
							{
								m_HttpResponse << m_ProvRegDenomChurches.m_City;
							}
							m_HttpResponse << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvRegDenomChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;
							}
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvRegChurchNameChurches
							//PROV_REG_CHURCHNAME_CHURCHES_HEADER
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince << "&Region=" << m_strRegion;
							if (DBSTATUS_S_ISNULL != m_ProvRegChurchNameChurches.m_dwCityStatus)
							{
								m_HttpResponse << "&City=" << m_ProvRegChurchNameChurches.m_City;
							}
							m_HttpResponse << "&Denomination=" << m_strDenomination
								<< "&ChurchName=" << m_ProvRegChurchNameChurches.m_ChurchName
								<< "\">" << m_ProvRegChurchNameChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvRegChurchNameChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvRegChurchNameChurches.m_Address << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvRegChurchNameChurches.m_dwPostalCodeStatus)
							{
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvRegChurchNameChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvRegChurchNameChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_ProvRegChurchNameChurches.m_City, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvRegChurchNameChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END << "</td><td class=\"chCity\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chCity\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvRegChurchNameChurches.m_dwCityStatus)
							{
								m_HttpResponse << m_ProvRegChurchNameChurches.m_City;
							}
							m_HttpResponse << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvRegChurchNameChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;
							}
						}
						else
						{
							//ProvRegChurches
							//PROV_REG_CHURCHES_HEADER
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince << "&Region=" << m_strRegion;
							if (DBSTATUS_S_ISNULL != m_ProvRegChurches.m_dwCityStatus)
							{
								m_HttpResponse << "&City=" << m_ProvRegChurches.m_City;
							}
							m_HttpResponse << "&ChurchName=" << m_ProvRegChurches.m_ChurchName
								<< "\">" << m_ProvRegChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvRegChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvRegChurches.m_Address << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvRegChurches.m_dwPostalCodeStatus)
							{
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvRegChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvRegChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_ProvRegChurches.m_City, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvRegChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END << "</td><td class=\"chCity\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chCity\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvRegChurches.m_dwCityStatus)
							{
								m_HttpResponse << m_ProvRegChurches.m_City;
							}
							m_HttpResponse << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvRegChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;
							}
						}
					}
				}
			}
			else
			{
				if (m_bCity)
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvCityDenomChurchNameChurches
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince << "&City=" << m_strCity
								<< "&Denomination=" << m_strDenomination
								<< "&ChurchName=" << m_ProvCityDenomChurchNameChurches.m_ChurchName
								<< "\">" << m_ProvCityDenomChurchNameChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvCityDenomChurchNameChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvCityDenomChurchNameChurches.m_Address << "</td><td class=\"chMail\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chMail\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvCityDenomChurchNameChurches.m_dwMailStatus)
							{
								m_HttpResponse << m_ProvCityDenomChurchNameChurches.m_Mail << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvCityDenomChurchNameChurches.m_dwPostalCodeStatus)
							{
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvCityDenomChurchNameChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvCityDenomChurchNameChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_strCity, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvCityDenomChurchNameChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END;
							}
							m_HttpResponse << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvCityDenomChurchNameChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;
							}
						}
						else
						{
							//ProvCityDenomChurches
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince << "&City=" << m_strCity << "&Denomination=" << m_strDenomination
								<< "&ChurchName=" << m_ProvCityDenomChurches.m_ChurchName
								<< "\">" << m_ProvCityDenomChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvCityDenomChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvCityDenomChurches.m_Address << "</td><td class=\"chMail\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chMail\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvCityDenomChurches.m_dwMailStatus)
							{
								m_HttpResponse << m_ProvCityDenomChurches.m_Mail << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvCityDenomChurches.m_dwPostalCodeStatus)
							{
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvCityDenomChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvCityDenomChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_strCity, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvCityDenomChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END;
							}
							m_HttpResponse << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvCityDenomChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;
							}
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvCityChurchNameChurches
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince << "&City=" << m_strCity
								<< "&ChurchName=" << m_ProvCityChurchNameChurches.m_ChurchName
								<< "\">" << m_ProvCityChurchNameChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvCityChurchNameChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvCityChurchNameChurches.m_Address << "</td><td class=\"chMail\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chMail\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvCityChurchNameChurches.m_dwMailStatus)
							{
								m_HttpResponse << m_ProvCityChurchNameChurches.m_Mail << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvCityChurchNameChurches.m_dwPostalCodeStatus)
							{
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvCityChurchNameChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvCityChurchNameChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_strCity, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvCityDenomChurchNameChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END;
							}
							m_HttpResponse << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvCityChurchNameChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;
							}
						}
						else
						{
							//ProvCityChurches
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince << "&City=" << m_strCity
								<< "&ChurchName=" << m_ProvCityChurches.m_ChurchName
								<< "\">" << m_ProvCityChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvCityChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvCityChurches.m_Address << "</td><td class=\"chMail\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chMail\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvCityChurches.m_dwMailStatus)
							{
								m_HttpResponse << m_ProvCityChurches.m_Mail << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvCityChurches.m_dwPostalCodeStatus)
							{
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvCityChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvCityChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_strCity, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvCityChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END << "</td><td class=\"chDenomination\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chDenomination\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvCityChurches.m_dwDenominationStatus)
							{
								m_HttpResponse << m_ProvCityChurches.m_Denomination;
							}
							m_HttpResponse << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvCityChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;
							}
						}
					}
				}
				else
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//ProvDenomChurchNameChurches
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince;
							if (DBSTATUS_S_ISNULL != m_ProvDenomChurchNameChurches.m_dwCityStatus)
							{
								m_HttpResponse << "&City=" << m_ProvDenomChurchNameChurches.m_City;
							}
							m_HttpResponse << "&ChurchName=" << m_ProvDenomChurchNameChurches.m_ChurchName
								<< "\">" << m_ProvDenomChurchNameChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvDenomChurchNameChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvDenomChurchNameChurches.m_Address << "</td><td class=\"chMail\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chMail\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvDenomChurchNameChurches.m_dwMailStatus)
							{
								m_HttpResponse << m_ProvDenomChurchNameChurches.m_Mail << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvDenomChurchNameChurches.m_dwPostalCodeStatus)
							{
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvDenomChurchNameChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvDenomChurchNameChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_ProvDenomChurchNameChurches.m_City, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvDenomChurchNameChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END << "</td><td class=\"chCity\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chCity\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvDenomChurchNameChurches.m_dwCityStatus)
							{
								m_HttpResponse << m_ProvDenomChurchNameChurches.m_City;
							}
							m_HttpResponse << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvDenomChurchNameChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;
							}
						}
						else
						{
							//ProvDenomChurches
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince;
							if (DBSTATUS_S_ISNULL != m_ProvDenomChurches.m_dwCityStatus)
							{
								m_HttpResponse << "&City=" << m_strCity;
							}
							m_HttpResponse << "&ChurchName=" << m_ProvDenomChurches.m_ChurchName
								<< "\">" << m_ProvDenomChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvDenomChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvDenomChurches.m_Address << "</td><td class=\"chMail\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chMail\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvDenomChurches.m_dwMailStatus)
							{
								m_HttpResponse << m_ProvDenomChurches.m_Mail << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvDenomChurches.m_dwPostalCodeStatus)
							{
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvDenomChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvDenomChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_ProvDenomChurches.m_City, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvDenomChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END;
							}
							m_HttpResponse << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvDenomChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;
							}
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//ProvChurchNameChurches
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince;
							if (DBSTATUS_S_ISNULL != m_ProvChurchNameChurches.m_dwCityStatus)
							{
								m_HttpResponse << "&City=" << m_ProvChurchNameChurches.m_City;
							}
							m_HttpResponse << "&ChurchName=" << m_ProvChurchNameChurches.m_ChurchName
								<< "\">" << m_ProvChurchNameChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvChurchNameChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvChurchNameChurches.m_Address << "</td><td class=\"chMail\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chMail\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvChurchNameChurches.m_dwMailStatus)
							{
								m_HttpResponse << m_ProvChurchNameChurches.m_Mail << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvChurchNameChurches.m_dwPostalCodeStatus)
							{
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvChurchNameChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvChurchNameChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_ProvChurchNameChurches.m_City, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvChurchNameChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END << "</td><td class=\"chCity\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chCity\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvChurchNameChurches.m_dwCityStatus)
							{
								m_HttpResponse << m_ProvChurchNameChurches.m_City;
							}
							m_HttpResponse << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvChurchNameChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;
							}
						}
						else
						{
							//ProvChurches
							m_HttpResponse << "<td class=\"chName\"><a href=\"index.srf?Province="
								<< m_strProvince;
							if (DBSTATUS_S_ISNULL != m_ProvChurches.m_dwCityStatus)
							{
								m_HttpResponse << "&City=" << m_ProvChurches.m_City;
							}
							m_HttpResponse << "&ChurchName=" << m_ProvChurches.m_ChurchName
								<< "\">" << m_ProvChurches.m_ChurchName
								<< "</a></td><td class=\"chAddress\">";
							if (DBSTATUS_S_ISNULL != m_ProvChurches.m_dwAddressStatus)
							{
								m_HttpResponse << m_ProvChurches.m_Address << "</td><td class=\"chMail\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chMail\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvChurches.m_dwMailStatus)
							{
								m_HttpResponse << m_ProvChurches.m_Mail << "</td><td class=\"chPostalCode\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chPostalCode\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvChurches.m_dwPostalCodeStatus)
							{
								m_HttpResponse << MAPQUEST << "title="
									<< CT2A(m_ProvChurches.m_ChurchName, CP_ACP)
									<< "&address=" << CT2A(m_ProvChurches.m_Address, CP_ACP)
									<< "&city=" << CT2A(m_ProvChurches.m_City, CP_ACP) << "&state=" << m_strProvince
									<< "&zipcode=" << CT2A(m_ProvChurches.m_PostalCode, CP_ACP)
									<< MAPQUEST_END << "</td><td class=\"chCity\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chCity\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvChurches.m_dwCityStatus)
							{
								m_HttpResponse << m_ProvChurches.m_City << "</td><td class=\"chRegion\">";
							}
							else
							{
								m_HttpResponse << "</td><td class=\"chRegion\">";
							}
							if (DBSTATUS_S_ISNULL != m_ProvChurches.m_dwRegionStatus)
							{
								m_HttpResponse << m_ProvChurches.m_Region;
							}
							m_HttpResponse << "</td>";
							COleDateTime tModified(UTIL::DbTimeStampToCOleDateTime(m_ProvChurches.m_Modified));
							if (tModified > m_LastModified)
							{
								m_LastModified = tModified;
							}
						}
					}
				}
			}
		}
		return HTTP_SUCCESS;
	}
	//church-details redirect handler. 
	HTTP_CODE OnRedirect(void)
	{
		if (m_bRedirect)
		{
			//can be either churchdetails or bypassparam type redirects
			if (m_bBypassParam)
			{
				ATLTRACE(L"BYPASS");
				if (m_bProvince)
				{
					CAtlString strUrlPath;
					strUrlPath = L"index.srf?Province=";
					strUrlPath += m_strProvince;
					if (0 != m_bRegion)
					{
						strUrlPath += L"&Region=";
						strUrlPath += m_strRegion;
						if (0 != m_bCity)
						{
							strUrlPath += L"&City=";
							strUrlPath += m_strCity;
							if (0 != m_bDenomination)
							{
								strUrlPath += L"&Denomination=";
								strUrlPath += m_strDenomination;
							}
							else
							{
								strUrlPath += L"&Denomination=";
								strUrlPath += m_ProvRegCityDenoms.m_Denomination;
							}
						}
						else
						{
							strUrlPath += L"&City=";
							strUrlPath += m_ProvRegCities.m_City;
						}
					}
					else
					{
						strUrlPath += L"&Region=";
						strUrlPath += m_ProvRegions.m_Region;
					}
					strUrlPath.Replace(L" ", L"%20");
					m_HttpResponse.Redirect(CT2A(strUrlPath));
					return HTTP_SUCCESS;
				}
			}
			else
			{
				ATLTRACE(L"non-BYPASS");
				CAtlString strRedirectURL;
				TCHAR buf[64];
				strRedirectURL = _T("index.srf?ChurchID=");
#if defined(_M_IA64) || defined (_M_AMD64)
				_i64tow(m_ChurchID, buf, 10);
#else
				_ltow(m_iId, buf, 10);
#endif
				strRedirectURL += buf;
				strRedirectURL.FreeExtra();
				if (m_bProvince)
				{
					strRedirectURL += _T("&Province=");
					strRedirectURL += m_strProvince;
					if (m_bRegion)
					{
						strRedirectURL += _T("&Region=");
						strRedirectURL += m_strRegion;
						if (m_bCity)
						{
							strRedirectURL += _T("&City=");
							strRedirectURL += m_strCity;
							if (m_bDenomination)
							{
								strRedirectURL += _T("&Denomination=");
								strRedirectURL += m_strDenomination;
								if (m_bChurchName)
								{
									strRedirectURL += _T("&ChurchName=");
									strRedirectURL += m_strChurchName;
									//ProvRegCityDenomChurchNameChurches
								}
							}
							else
							{
								if (m_bChurchName)
								{
									strRedirectURL += _T("&ChurchName=");
									strRedirectURL += m_strChurchName;
									//ProvRegCityChurchNameChurches
								}
							}
						}
						else
						{
							if (m_bDenomination)
							{
								strRedirectURL += _T("&Denomination=");
								strRedirectURL += m_strDenomination;
								if (0 != m_bChurchName)
								{
									strRedirectURL += _T("&ChurchName=");
									strRedirectURL += m_strChurchName;
									//ProvRegDenomChurchNameChurches
								}
							}
							else
							{
								if (m_bChurchName)
								{
									//ProvRegChurchNameChurches
									strRedirectURL += _T("&ChurchName=");
									strRedirectURL += m_strChurchName;
								}
							}
						}
					}
					else
					{
						if (m_bCity)
						{
							strRedirectURL += _T("&City=");
							strRedirectURL += m_strCity;
							if (m_bDenomination)
							{
								strRedirectURL += _T("&Denomination=");
								strRedirectURL += m_strDenomination;
								if (m_bChurchName)
								{
									//ProvCityDenomChurchNameChurches
									strRedirectURL += _T("&ChurchName=");
									strRedirectURL += m_strChurchName;
								}
							}
							else
							{
								if (m_bChurchName)
								{
									//ProvCityChurchNameChurches
									strRedirectURL += _T("&ChurchName=");
									strRedirectURL += m_strChurchName;
								}
							}
						}
						else
						{
							if (m_bDenomination)
							{
								strRedirectURL += _T("&Denomination=");
								strRedirectURL += m_strDenomination;
								if (0 != m_bChurchName)
								{
									//ProvDenomChurchNameChurches
									strRedirectURL += _T("&ChurchName=");
									strRedirectURL += m_strChurchName;
								}
							}
							else
							{
								if (m_bChurchName)
								{
									//ProvChurchNameChurches
									strRedirectURL += _T("&ChurchName=");
									strRedirectURL += m_strChurchName;
								}
							}
						}
					}
				}
				strRedirectURL.Replace(L" ", L"%20");
				m_HttpResponse.Redirect(CW2A(strRedirectURL));
			}
		}
		return HTTP_SUCCESS;
	}
	//used to output numbers at the bottom of the page, so user's can traverse the rowset easily
	HTTP_CODE OnRowscroll(void)
	{
		if (m_bId)
		{
			return HTTP_SUCCESS;
		}

#if defined(_M_IA64) || defined (_M_AMD64)
		unsigned __int64 i;
#else
		unsigned __int32 i;
#endif
		if (m_Pages > 10)
		{
			m_HttpResponse << "<td " << MOUSEOVER << "><a "
				<< LINKCOLOR << " href=\"index.srf?";
			if (m_bProvince)
			{
				if (m_bRegion)
				{
					if (m_bCity)
					{
						if (m_bDenomination)
						{
							if (m_bChurchName)
							{
								//ProvRegCityDenomChurchNameChurches
							}
							else
							{
								//ProvRegCityDenomChurches
								m_HttpResponse << "Province=" << m_strProvince
									<< "&Region=" << m_strRegion << "&City="
									<< m_strCity << "&Denomination=" << m_strDenomination
									<< "&Pp=" << m_iPp << "&" << "Pg="
									<< 9999 << "\">";
							}
						}
						else
						{
							if (m_bChurchName)
							{
								//ProvRegCityChurchNameChurches
							}
							else
							{
								//ProvRegCityChurches
								m_HttpResponse << "Province=" << m_strProvince
									<< "&Region=" << m_strRegion << "&City="
									<< m_strCity << "&Pp=" << m_iPp
									<< "&" << "Pg=" << 9999 << "\">";
							}
						}
					}
					else
					{
						if (m_bDenomination)
						{
							if (m_bChurchName)
							{
								//ProvRegDenomChurchNameChurches
							}
							else
							{
								//ProvRegDenomChurches
								m_HttpResponse << "Province=" << m_strProvince
									<< "&Region=" << m_strRegion
									<< "&Denomination=" << m_strDenomination
									<< "&Pp=" << m_iPp << "&" << "Pg="
									<< 9999 << "\">";
							}
						}
						else
						{
							if (m_bChurchName)
							{
								//ProvRegChurchNameChurches
							}
							else
							{
								//ProvRegChurches
								m_HttpResponse << "Province=" << m_strProvince
									<< "&Region=" << m_strRegion
									<< "&Pp=" << m_iPp << "&" << "Pg="
									<< 9999 << "\">";
							}
						}
					}
				}
				else
				{
					if (m_bCity)
					{
						if (m_bDenomination)
						{
							if (m_bChurchName)
							{
								//ProvCityDenomChurchNameChurches
							}
							else
							{
								//ProvCityDenomChurches
								m_HttpResponse << "Province=" << m_strProvince
									<< "&City=" << m_strCity
									<< "&Denomination=" << m_strDenomination
									<< "&Pp=" << m_iPp << "&" << "Pg="
									<< 9999 << "\">";
							}
						}
						else
						{
							if (m_bChurchName)
							{
								//ProvCityChurchNameChurches
							}
							else
							{
								//ProvCityChurches
								m_HttpResponse << "Province=" << m_strProvince << "&City="
									<< m_strCity << "&Pp=" << m_iPp << "&Pg=" << 9999 << "\">";
							}
						}
					}
					else
					{
						if (m_bDenomination)
						{
							if (m_bChurchName)
							{
								//ProvDenomChurchNameChurches
							}
							else
							{
								//ProvDenomChurches
								m_HttpResponse << "Province=" << m_strProvince << "&Denomination="
									<< m_strDenomination << "&Pp=" << m_iPp << "&Pg=" << 9999 << "\">";
							}
						}
						else
						{
							if (m_bChurchName)
							{
								//ProvChurchNameChurches
							}
							else
							{
								//ProvChurches
								m_HttpResponse << "Province=" << m_strProvince
									<< "&Pp=" << m_iPp << "&" << "Pg="
									<< 9999 << "\">";
							}
						}
					}
				}
			}
			// output hyperlinks if it's a really old browser
			if (m_bMobile)
			{
				m_HttpResponse << "<<</a></td>";
			}
			else//output corresponding png images
			{
				m_HttpResponse << "<img src=\"images/numbers/prev.png\" alt=\"Next Page\"></a></td>";
			}
			//this is another large control loop, used to iterate possible pages for each rowset
			for (i = m_CurPage; i < m_CurPage + 10; i++)
			{
				if (i == 9999)
				{
					goto jumpScroll;
				}
				else
				{
					if (i >= m_Pages)
						break;
				}
			jumpScroll:
				// this branch generates the page numbers of the scroll
				m_HttpResponse << "<td class=\"spScroll\" " << MOUSEOVER << "><a "
					<< LINKCOLOR << " href=\"index.srf?";
				if (m_bProvince)
				{
					if (m_bRegion)
					{
						if (m_bCity)
						{
							if (m_bDenomination)
							{
								if (m_bChurchName)
								{
									//ProvRegCityDenomChurchNameChurches
								}
								else
								{
									//ProvRegCityDenomChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&Region=" << m_strRegion << "&City="
										<< m_strCity << "&Denomination=" << m_strDenomination
										<< "&Pp=" << m_iPp << "&" << "Pg="
										<< (i - 1) << "\">";
								}
							}
							else
							{
								if (m_bChurchName)
								{
									//ProvRegCityChurchNameChurches
								}
								else
								{
									//ProvRegCityChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&Region=" << m_strRegion << "&City="
										<< m_strCity << "&Pp=" << m_iPp << "&" << "Pg="
										<< (i - 1) << "\">";
								}
							}
						}
						else
						{
							if (m_bDenomination)
							{
								if (m_bChurchName)
								{
									//ProvRegDenomChurchNameChurches
								}
								else
								{
									//ProvRegDenomChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&Region=" << m_strRegion << "&Denomination="
										<< m_strDenomination << "&Pp=" << m_iPp
										<< "&" << "Pg=" << i << "\">";
								}
							}
							else
							{
								if (m_bChurchName)
								{
									//ProvRegChurchNameChurches
								}
								else
								{
									//ProvRegChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&Region=" << m_strRegion
										<< "&Pp=" << m_iPp << "&" << "Pg="
										<< i << "\">";
								}
							}
						}
					}
					else
					{
						if (m_bCity)
						{
							if (m_bDenomination)
							{
								if (m_bChurchName)
								{
									//ProvCityDenomChurchNameChurches
								}
								else
								{
									//ProvCityDenomChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&City=" << m_strCity << "&Denomination="
										<< m_strDenomination << "&Pp=" << m_iPp << "&"
										<< "Pg=" << i << "\">";
								}
							}
							else
							{
								if (m_bChurchName)
								{
									//ProvCityChurchNameChurches
								}
								else
								{
									//ProvCityChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&City=" << m_strCity << "&Pp=" << m_iPp
										<< "&Pg=" << i << "\">";
								}
							}
						}
						else
						{
							if (m_bDenomination)
							{
								if (m_bChurchName)
								{
									//ProvDenomChurchNameChurches
								}
								else
								{
									//ProvDenomChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&City=" << m_strCity << "&Pp=" << m_iPp
										<< "&Pg=" << i << "\">";
								}
							}
							else
							{
								if (m_bChurchName)
								{
									//ProvChurchNameChurches
								}
								else
								{
									//ProvChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&Pp=" << m_iPp << "&" << "Pg="
										<< i << "\">";
								}
							}
						}
					}
				}
				// output corresponding png images
				if (m_bMobile)
				{
					if (i == 0)
					{
						m_HttpResponse << "<<</a> |</td>";
					}
					else
					{
						m_HttpResponse << i << "</a> |</td>";
					}
				}
				else
				{
					m_HttpResponse << "<img src=\"images/numbers/" << i << ".png\" alt=\"Page " << i << "\"></a></td>";
				}
			}
			// this next branching takes care of the 'Next' page hyperlink
			m_HttpResponse << "<td " << MOUSEOVER << "><a "
				<< LINKCOLOR << " href=\"index.srf?";
			if (m_bProvince)
			{
				if (m_bRegion)
				{
					if (m_bCity)
					{
						if (m_bDenomination)
						{
							if (m_bChurchName)
							{
								//ProvRegCityDenomChurchNameChurches
							}
							else
							{
								//ProvRegCityDenomChurches
								m_HttpResponse << "Province=" << m_strProvince
									<< "&Region=" << m_strRegion << "&City="
									<< m_strCity << "&Denomination=" << m_strDenomination
									<< "&Pp=" << m_iPp << "&" << "Pg="
									<< (i - 1) << "\">";
							}
						}
						else
						{
							if (m_bChurchName)
							{
								//ProvRegCityChurchNameChurches
							}
							else
							{
								//ProvRegCityChurches
								m_HttpResponse << "Province=" << m_strProvince
									<< "&Region=" << m_strRegion << "&City="
									<< m_strCity << "&Pp=" << m_iPp
									<< "&" << "Pg=" << (i - 1) << "\">";
							}
						}
					}
					else
					{
						if (m_bDenomination)
						{
							if (m_bChurchName)
							{
								//ProvRegDenomChurchNameChurches
							}
							else
							{
								//ProvRegDenomChurches
								m_HttpResponse << "Province=" << m_strProvince
									<< "&Region=" << m_strRegion
									<< "&Denomination=" << m_strDenomination
									<< "&Pp=" << m_iPp << "&" << "Pg="
									<< i << "\">";
							}
						}
						else
						{
							if (m_bChurchName)
							{
								//ProvRegChurchNameChurches
							}
							else
							{
								//ProvRegChurches
								m_HttpResponse << "Province=" << m_strProvince
									<< "&Region=" << m_strRegion
									<< "&Pp=" << m_iPp << "&" << "Pg="
									<< i << "\">";
							}
						}
					}
				}
				else
				{
					if (m_bCity)
					{
						if (m_bDenomination)
						{
							if (m_bChurchName)
							{
								//ProvCityDenomChurchNameChurches
							}
							else
							{
								//ProvCityDenomChurches
								m_HttpResponse << "Province=" << m_strProvince
									<< "&City=" << m_strCity
									<< "&Denomination=" << m_strDenomination
									<< "&Pp=" << m_iPp << "&" << "Pg="
									<< i << "\">";
							}
						}
						else
						{
							if (m_bChurchName)
							{
								//ProvCityChurchNameChurches
							}
							else
							{
								//ProvCityChurches
								m_HttpResponse << "Province=" << m_strProvince << "&City="
									<< m_strCity << "&Pp=" << m_iPp << "&Pg=" << i << "\">";
							}
						}
					}
					else
					{
						if (m_bDenomination)
						{
							if (m_bChurchName)
							{
								//ProvDenomChurchNameChurches
							}
							else
							{
								//ProvDenomChurches
								m_HttpResponse << "Province=" << m_strProvince << "&Denomination="
									<< m_strDenomination << "&Pp=" << m_iPp << "&Pg=" << i << "\">";
							}
						}
						else
						{
							if (m_bChurchName)
							{
								//ProvChurchNameChurches
							}
							else
							{
								//ProvChurches
								m_HttpResponse << "Province=" << m_strProvince
									<< "&Pp=" << m_iPp << "&" << "Pg="
									<< i << "\">";
							}
						}
					}
				}
			}
			// output corresponding png images
			if (0 != m_bMobile)
			{
				m_HttpResponse << ">></a></td>";
			}
			else
			{
				m_HttpResponse << "<img src=\"images/numbers/next.png\" alt=\"Next Page\"></a></td>";
			}
		}
		else
		{
			//branching which occurs when there are fewer then 10 pages
			for (i = m_CurPage; i < m_CurPage + 10; i++)
			{
				if (i >= m_Pages)
					break;
				m_HttpResponse << "<td " << MOUSEOVER << "><a "
					<< LINKCOLOR << " href=\"index.srf?";
				if (m_bProvince)
				{
					if (m_bRegion)
					{
						if (m_bCity)
						{
							if (m_bDenomination)
							{
								if (m_bChurchName)
								{
									//ProvRegCityDenomChurchNameChurches
								}
								else
								{
									//ProvRegCityDenomChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&Region=" << m_strRegion << "&City="
										<< m_strCity << "&Denomination=" << m_strDenomination
										<< "&Pp=" << m_iPp << "&" << "Pg="
										<< (i - 1) << "\">";
								}
							}
							else
							{
								if (m_bChurchName)
								{
									//ProvRegCityChurchNameChurches
								}
								else
								{
									//ProvRegCityChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&Region=" << m_strRegion << "&City="
										<< m_strCity << "&Pp=" << m_iPp << "&" << "Pg="
										<< (i - 1) << "\">";
								}
							}
						}
						else
						{
							if (m_bDenomination)
							{
								if (m_bChurchName)
								{
									//ProvRegDenomChurchNameChurches
								}
								else
								{
									//ProvRegDenomChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&Region=" << m_strRegion
										<< "&Denomination=" << m_strDenomination
										<< "&Pp=" << m_iPp << "&" << "Pg="
										<< (i - 1) << "\">";
								}
							}
							else
							{
								if (m_bChurchName)
								{
									//ProvRegChurchNameChurches
								}
								else
								{
									//ProvRegChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&Region=" << m_strRegion
										<< "&Pp=" << m_iPp << "&" << "Pg="
										<< (i - 1) << "\">";
								}
							}
						}
					}
					else
					{
						if (m_bCity)
						{
							if (m_bDenomination)
							{
								if (m_bChurchName)
								{
									//ProvCityDenomChurchNameChurches
								}
								else
								{
									//ProvCityDenomChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&City=" << m_strCity
										<< "&Denomination=" << m_strDenomination
										<< "&Pp=" << m_iPp << "&" << "Pg="
										<< (i - 1) << "\">";
								}
							}
							else
							{
								if (m_bChurchName)
								{
									//ProvCityChurchNameChurches
								}
								else
								{
									//ProvCityChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&City=" << m_strCity << "&Pp=" << m_iPp
										<< "&Pg=" << (i - 1) << "\">";
								}
							}
						}
						else
						{
							if (m_bDenomination)
							{
								if (m_bChurchName)
								{
									//ProvDenomChurchNameChurches
								}
								else
								{
									//ProvDenomChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&Denomination=" << m_strDenomination
										<< "&Pp=" << m_iPp << "&" << "Pg="
										<< (i - 1) << "\">";
								}
							}
							else
							{
								if (m_bChurchName)
								{
									//ProvChurchNameChurches
								}
								else
								{
									//ProvChurches
									m_HttpResponse << "Province=" << m_strProvince
										<< "&Pp=" << m_iPp << "&" << "Pg="
										<< (i - 1) << "\">";
								}
							}
						}
					}
				}
				if (m_bMobile)
				{
					m_HttpResponse << i << "</a> |</td>";
				}
				else
				{
					m_HttpResponse << "<img src=\"images/numbers/" << i << ".png\" alt=\"Page " << i << "\"></a></td>";
				}
				if ((i + 1)*m_dbPp < m_Rows)
				{
					if (m_Rows < (i + 2)*m_dbPp)
					{
						//middleman
						m_HttpResponse << "<td " << MOUSEOVER << "><a " << LINKCOLOR
							<< " href=\"index.srf?";
						if (m_bProvince)
						{
							if (m_bRegion)
							{
								if (m_bCity)
								{
									if (m_bDenomination)
									{
										if (m_bChurchName)
										{
											//ProvRegCityDenomChurchNameChurches
										}
										else
										{
											//ProvRegCityDenomChurches
											m_HttpResponse << "Province=" << m_strProvince
												<< "&Region=" << m_strRegion << "&City="
												<< m_strCity << "&Denomination=" << m_strDenomination
												<< "&Pp=" << m_iPp << "&" << "Pg="
												<< i + 1 << "\">";
										}
									}
									else
									{
										if (m_bChurchName)
										{
											//ProvRegCityChurchNameChurches
										}
										else
										{
											//ProvRegCityChurches
											m_HttpResponse << "Province=" << m_strProvince
												<< "&Region=" << m_strRegion << "&City="
												<< m_strCity << "&Pp=" << m_iPp << "&" << "Pg="
												<< i + 1 << "\">";
										}
									}
								}
								else
								{
									if (m_bDenomination)
									{
										if (m_bChurchName)
										{
											//ProvRegDenomChurchNameChurches
										}
										else
										{
											//ProvRegDenomChurches
											m_HttpResponse << "Province=" << m_strProvince
												<< "&Region=" << m_strRegion
												<< "&Denomination=" << m_strDenomination
												<< "&Pp=" << m_iPp << "&" << "Pg="
												<< i + 1 << "\">";
										}
									}
									else
									{
										if (m_bChurchName)
										{
											//ProvRegChurchNameChurches
										}
										else
										{
											//ProvRegChurches
											m_HttpResponse << "Province=" << m_strProvince
												<< "&Region=" << m_strRegion
												<< "&Pp=" << m_iPp << "&" << "Pg="
												<< i + 1 << "\">";
										}
									}
								}
							}
							else
							{
								if (m_bCity)
								{
									if (m_bDenomination)
									{
										if (m_bChurchName)
										{
											//ProvCityDenomChurchNameChurches
										}
										else
										{
											//ProvCityDenomChurches
											m_HttpResponse << "Province=" << m_strProvince
												<< "&City=" << m_strCity
												<< "&Denomination=" << m_strDenomination
												<< "&Pp=" << m_iPp << "&" << "Pg="
												<< i + 1 << "\">";
										}
									}
									else
									{
										if (0 != m_bChurchName)
										{
											//ProvCityChurchNameChurches
										}
										else
										{
											//ProvCityChurches
											m_HttpResponse << "Province=" << m_strProvince
												<< "&City=" << m_strCity << "&Pp="
												<< m_iPp << "&Pg=" << i + 1 << "\">";
										}
									}
								}
								else
								{
									if (m_bDenomination)
									{
										if (m_bChurchName)
										{
											//ProvDenomChurchNameChurches
										}
										else
										{
											//ProvDenomChurches
											m_HttpResponse << "Province=" << m_strProvince
												<< "&Denomination=" << m_strDenomination
												<< "&Pp=" << m_iPp << "&" << "Pg="
												<< i + 1 << "\">";
										}
									}
									else
									{
										if (m_bChurchName)
										{
											//ProvChurchNameChurches
										}
										else
										{
											//ProvChurches
											m_HttpResponse << "Province=" << m_strProvince
												<< "&Pp=" << m_iPp << "&" << "Pg="
												<< i + 1 << "\">";
										}
									}
								}
							}
						}
						if (m_bMobile)
						{
							m_HttpResponse << i + 1 << "</a> |</td>";
						}
						else
						{
							m_HttpResponse << "<img src=\"images/numbers/" << i + 1 << ".png\" alt=\"Page\" " << i + 1 << "\"></a></td>";
						}
					}
				}
			}
		}
		return HTTP_SUCCESS;
	}
	//whether to display the item details or not
	HTTP_CODE OnbDetails(void)
	{
		return (0 != m_bId) ? HTTP_SUCCESS : HTTP_S_FALSE;
	}
	// if we are in details mode, show the churches uploaded picture
	HTTP_CODE OnbImg(void)
	{
		if (0 != m_bId)
		{
			CAtlStringW wstrPath;
			DWORD dwResult = GetModuleFileName(this->GetResourceInstance(), wstrPath.GetBuffer(MAX_PATH), MAX_PATH);
			if (0 != dwResult)
			{
				wstrPath.ReleaseBuffer();
				CAtlStringW wstrImgPath(wstrPath.Left(wstrPath.ReverseFind(L'.')));
				if (0 != wstrImgPath.Replace(_T("cgi-bin"), _T("images\\churches\\")))
				{
#if defined (_M_IA64) || defined (_M_AMD64)
					strPath.m_strPath.AppendFormat(_T("%I64.jpg"), m_ChurchID);
#else
					wstrImgPath.AppendFormat(_T("%i.jpg"), m_iId);
#endif
					// check to see if the file exists, if it doesn't don't render image
					HANDLE hFile = CreateFile(wstrImgPath, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, NULL);
					if (hFile == INVALID_HANDLE_VALUE)
					{
						return HTTP_S_FALSE;//doesn't exist
					}
					else
					{
						CloseHandle(hFile);
						return HTTP_SUCCESS;
					}
				}
			}
			else
			{
				return UTIL::SendError(m_HttpResponse, "Unable to get module file name");
			}
		}
		// we are not in details mode, so display the usual geographic images
		if (m_bProvince)
		{
			if (m_bRegion)
				return HTTP_S_FALSE;
		}
		return HTTP_SUCCESS;
	}
	//if the above code evaluates to true, display the image
	HTTP_CODE OnImg(void)
	{
		if (m_bId)
		{
			wchar_t buf[32];
#if defined(_M_IA64) || defined (_M_AMD64)
			_i64tow(m_ChurchID, buf, 10);
#else
			_ltow(m_iId, buf, 10);
#endif
			m_HttpResponse << "churches/" << buf << ".jpg";
		}
		else
		{
			if (m_bProvince)
				m_HttpResponse << m_strProvince << ".gif";
			else
				m_HttpResponse << "canada.gif";
		}
		return HTTP_SUCCESS;
	}
	//the alt tag
	HTTP_CODE OnImgAlt(void)
	{
		if (m_bId)
		{
			if (m_bChurchName)
			{
				m_HttpResponse << m_strChurchName;
			}
		}
		else
		{
			if (0 != m_bProvince)
			{
				if (0 != m_bRegion)
				{
				}
				else
					m_HttpResponse << "Please choose a region in " << UTIL::GetNamedProvince(m_strProvince);
			}
		}
		return HTTP_SUCCESS;
	}
	//handles the image map
	HTTP_CODE OnImgMap(void)
	{
		if (0 != m_bId)
		{
			m_HttpResponse << "no";
		}
		else
		{
			if (0 != m_bProvince)
			{
				if (0 != m_bRegion)
				{
				}
				else
					m_HttpResponse << m_strProvince;
			}
			else
				m_HttpResponse << "canada";
		}
		return HTTP_SUCCESS;
	}
	//alternatives to be displayed in the churchdetails page.
	HTTP_CODE OnAlternatives(void)
	{
		if (0 != m_bId)
		{
			m_HttpResponse << ALTERNATIVES << m_strProvince << "\">"
				<< UTIL::GetNamedProvince(m_strProvince) << "</a> | "
				<< AHREF << m_strProvince << "&Region=" << m_strRegion << "\">"
				<< m_strRegion << "</a> | " << AHREF << m_strProvince << "&Region="
				<< m_strRegion << "&City=" << m_strCity << "\">" << m_strCity << "</a> | the "
				<< AHREF << m_strProvince << "&Region=" << m_strRegion << "&City="
				<< m_strCity << "&Denomination=" << m_strDenomination << "\">"
				<< m_strDenomination << "</a> Denomination";
		}
		return HTTP_SUCCESS;
	}
	//our web-app is SEO-friendly
	HTTP_CODE OnMetas(void)
	{
		if (m_bProvince)
		{
			if (m_bRegion)
			{
				if (m_bCity)
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							m_HttpResponse << "\t<meta name=\"keywords\" content=\"Church, "
								<< m_strCity << ", " << m_strChurchName
								<< "\">\r\n\t<meta name=\"description\" content=\""
								<< m_strChurchName << " is an " << m_strDenomination
								<< " Church located in " << m_strCity << ", "
								<< m_strProvince << ".\">";
						}
						else
						{
							//CspProvRegCityDenomChurches
							m_HttpResponse << "\t<meta name=\"keywords\" content=\"Church, "
								<< m_strCity << ", " << m_strDenomination
								<< "\">\r\n\t<meta name=\"description\" content=\""
								<< " Churches located in " << m_strCity << ", "
								<< m_strProvince << ".\">";
						}
					}
					else
					{
						if (m_bChurchName)
						{
							m_HttpResponse << "\t<meta name=\"keywords\" content=\"Church, "
								<< m_strCity << ", " << m_strChurchName
								<< "\">\r\n\t<meta name=\"description\" content=\""
								<< m_strChurchName << " is located in " << m_strCity << ", "
								<< m_strProvince << ".\">";
						}
						else
						{
							//CspProvRegCityChurches
							m_HttpResponse << "\t<meta name=\"keywords\" content=\"Church, "
								<< m_strCity << "\">\r\n\t<meta name=\"description\" content=\""
								<< "Churches located in " << m_strCity << ", " << m_strProvince << ".\">";
						}
					}
				}
				else
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//this should be handled by m_bDetails
							m_HttpResponse << "\t<meta name=\"keywords\" content=\"Church, "
								<< m_strChurchName << ", " << m_strDenomination
								<< "\">\r\n\t<meta name=\"description\" content=\""
								<< m_strChurchName << " is a " << m_strDenomination
								<< " Church located in " << m_strProvince << ".\">";
						}
						else
						{
							//CspProvRegDenomChurches
							m_HttpResponse << "\t<meta name=\"keywords\" content=\"Church, "
								<< m_strDenomination << ", " << m_strProvince
								<< "\">\r\n\t<meta name=\"description\" content=\""
								<< m_strDenomination << " Churches located in " << m_strRegion
								<< ", " << m_strProvince << ".\">";
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//this should be handled by m_bDetails
							m_HttpResponse << "\t<meta name=\"keywords\" content=\"Church, "
								<< m_strChurchName << ", " << m_strRegion
								<< "\">\r\n\t<meta name=\"description\" content=\""
								<< m_strChurchName << " is located in " << m_strProvince << ".\">";
						}
						else
						{
							//CspProvRegChurches
							m_HttpResponse << "\t<meta name=\"keywords\" content=\"Church, "
								<< m_strRegion << ", " << m_strProvince
								<< "\">\r\n\t<meta name=\"description\" content=\""
								<< "Churches located in " << m_strRegion << ", " << m_strProvince << ".\">";
						}
					}
				}
			}
			else
			{
				if (m_bDenomination)
				{
					if (m_bChurchName)
					{
						//this should be handled by m_bDetails
						m_HttpResponse << "\t<meta name=\"keywords\" content=\"Church, "
							<< m_strChurchName << ", " << m_strDenomination
							<< "\">\r\n\t<meta name=\"description\" content=\""
							<< m_strChurchName << " is a " << m_strDenomination << " Church located in "
							<< m_strProvince << ".\">";
					}
					else
					{
						//CspProvDenomChurches
						m_HttpResponse << "\t<meta name=\"keywords\" content=\"Church, "
							<< m_strDenomination << ", " << m_strProvince
							<< "\">\r\n\t<meta name=\"description\" content=\""
							<< m_strDenomination << " churches located in " << m_strProvince << ".\">";
					}
				}
				else
				{
					if (m_bChurchName)
					{
						//this should be handled by m_bDetails
						m_HttpResponse << "\t<meta name=\"keywords\" content=\"Church, "
							<< m_strChurchName << ", " << m_strProvince
							<< "\">\r\n\t<meta name=\"description\" content=\""
							<< m_strChurchName << " is a Church located in " << m_strProvince << ".\">";
					}
					else
					{
						//CspProvChurches
						m_HttpResponse << "\t<meta name=\"keywords\" content=\"Church, "
							<< m_strProvince << "\">\r\n\t<meta name=\"description\" content=\""
							<< "Churches located in " << m_strProvince << ".\">";
					}
				}
			}
		}
		else
		{
			if (m_bRegion)
			{
				if (m_bCity)
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//this should be handled by m_bDetails
							m_HttpResponse << "\t<meta name=\"keywords\" content=\"Church, "
								<< m_strRegion << ", " << m_strChurchName << ", " << m_strDenomination
								<< ", " << m_strCity << "\">\r\n\t<meta name=\"description\" content=\""
								<< m_strChurchName << " is a " << m_strDenomination << " Church located in "
								<< m_strCity << ", " << m_strRegion << ".\">";
						}
						else
						{
							//CspRegCityDenomChurches
							m_HttpResponse << "\t<meta name=\"keywords\" content=\"Church, "
								<< m_strRegion << ", " << m_strCity << ", " << m_strDenomination
								<< "\">\r\n\t<meta name=\"description\" content=\""
								<< "Churches located in " << m_strCity << ", " << m_strRegion << ".\">";
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//this should be handled by m_bDetails
						}
						else
						{
							//CspRegCityChurches
						}
					}
				}
				else
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//this should be handled by m_bDetails
						}
						else
						{
							//CspRegDenomChurches
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//this should be handled by m_bDetails
						}
						else
						{
							//CspRegChurches
						}
					}
				}
			}
			else
			{
				if (m_bCity)
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//this should be handled by m_bDetails
						}
						else
						{
							//CspCityDenomChurches
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//this should be handled by m_bDetails
						}
						else
						{
							//CspCityChurches
						}
					}
				}
				else
				{
					if (m_bDenomination)
					{
						if (m_bChurchName)
						{
							//this should be handled by m_bDetails
						}
						else
						{
							//CspDenomChurches
						}
					}
					else
					{
						if (m_bChurchName)
						{
							//this should be handled by m_bDetails
						}
						else
						{
							m_HttpResponse << "\t<meta name=\"keywords\" content=\"ChurchSearch, "
								<< "Church, Churches in Canada, Church Directory, "
								<< "Search for Churches, online churches, church listings "
								<< "Canada\">\r\n\t<meta name=\"description\" content=\" "
								<< "Search for Churches located in Canada.\">";
						}
					}
				}
			}
		}
		return HTTP_SUCCESS;
	}
	//used to modify display of css layout
	HTTP_CODE OnbOldBrowser(void)
	{
		return (m_bOldBrowser) ? HTTP_SUCCESS : HTTP_S_FALSE;
	}
	//used to implement head method
	HTTP_CODE OnHeadMethodCatch(void)
	{
		if (m_HttpRequest.GetMethod() == CHttpRequest::HTTP_METHOD_HEAD)
		{
			DWORD dwLength = m_HttpResponse.m_strContent.GetLength();
			m_HttpResponse.SetWriteToClient(FALSE);
			CHAR szTmp[21];
			Checked::ultoa_s(dwLength, szTmp, _countof(szTmp), 10);
			m_HttpResponse.AppendHeader("content-length", szTmp);
		}
		return HTTP_SUCCESS;
	}
	HTTP_CODE OnAppendETagCatch(void)
	{
		if (m_HttpRequest.GetMethod() != CHttpRequest::HTTP_METHOD_POST)
		{
			SYSTEMTIME st = { 0 };
			m_LastModified.GetAsSystemTime(st);
			CTime myTm(st);
			char szETag[21 + 1];
			Checked::ui64toa_s(myTm.GetTime(), szETag, 22, 10);
			CAtlStringA strEtag('\"');
			strEtag.AppendFormat("%s\"",CAtlStringA(szETag).TrimRight('0'));

			if (!m_liEtags.IsEmpty())//means the user wishes to retrieve a cached response
			{
				POSITION myPos = NULL;
				myPos = m_liEtags.GetHeadPosition();
				do {
					CAtlStringA strEtag;
					strEtag.Append(m_liEtags.GetNext(myPos));
					if (!strEtag.IsEmpty())
					{
						if (0 == strEtag.Compare(CAtlStringA(szETag).TrimRight('0')))
						{
							bool bHttp2 = false;
							CAtlStringA strProtocol;
							if (UTIL::GetServerVariable(m_spServerContext, "SERVER_PROTOCOL", strProtocol))
							{
								if (0 == strProtocol.CompareNoCase("HTTP/2.0"))
									bHttp2 = true;
								else
									bHttp2 = false;
							}
							UTIL::SendError(m_HttpResponse, "Not Modified", 304);
							return HTTP_SUCCESS_NO_PROCESS;//break out
						}
					}

				} while (myPos != NULL);
			}

			m_HttpResponse.AppendHeader("etag", strEtag);
			m_HttpResponse.AppendHeader("Last-Modified", UTIL::COleDateTimeToHttpDate(m_LastModified));
			m_HttpResponse.SetCacheControl("public");
		}
		return HTTP_SUCCESS;
	}
	CAtlStringA AssembleHttpHeadersPush()
	{
		CAtlStringA strGetHeaders;
		CAtlStringA strHost;
		if (m_HttpRequest.GetServerVariable("HTTP_HOST", strHost))
		{
			strGetHeaders.AppendFormat("Host: %s\r\n", strHost);
		}
		CAtlStringA strUserAgent;
		if (m_HttpRequest.GetServerVariable("HTTP_USER_AGENT", strUserAgent))
		{
			strGetHeaders.AppendFormat("User-Agent: %s\r\n", strUserAgent);
		}
		strGetHeaders += "Connection: Keep-Alive\r\n";
		return strGetHeaders;
	}
	//PushFile takes a filename and headers and pushes the resource to a client
	bool PushFile(CAtlStringA strPath, CAtlStringA strHeaders)
	{
		bool bRet = false;
		HANDLE hImgFile = CreateFile(CA2T(strPath), GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
		if (INVALID_HANDLE_VALUE != hImgFile)
		{
			//if the push fails, do not push any more images
			if (m_spServerContext->TransmitFile(hImgFile, NULL, m_spServerContext, NULL, NULL, NULL, (void*)strHeaders.GetString(), strHeaders.GetLength(), NULL, NULL, HSE_IO_DECLARE_PUSH | HSE_IO_SEND_HEADERS))
			{
				bRet = true;
			}
		}
		return bRet;
	}
	
	/*HTTP_CODE GetFlags(LPDWORD pdwStatus)
	{
		*pdwStatus = ATLSRV_INIT_USECACHE;
		return HTTP_SUCCESS;
	}*/
}; // class CShowMeHandler