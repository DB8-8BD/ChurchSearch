/**************************************************************************************
*                                       dbcommands.h                                  *
**************************************************************************************/
/*
	This file contains 56 classes which represent 56 stored procedures in a sql server
	2000 database. The following are arranged according to logic branching in the 
	CChurchSearchHandler::ValidateAndExchange(). Most of this code is wizard generated
*/
/*************************************************************************************/
#include "efx.h"
// CspChurchDetails
class CspChurchDetailsAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Denomination[FIELDSIZE_DENOMINATION];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_City[FIELDSIZE_CITY];
	TCHAR m_Province[FIELDSIZE_PROVINCE];
	TCHAR m_Region[FIELDSIZE_REGION];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_Phone[FIELDSIZE_PHONE];
	TCHAR m_Fax[FIELDSIZE_FAX];
	TCHAR m_ContactPerson[FIELDSIZE_CONTACT_PERSON];
	TCHAR m_Email[FIELDSIZE_EMAIL];
	TCHAR m_Website[FIELDSIZE_WEBSITE];
	TCHAR m_DenomWebsite[FIELDSIZE_DENOM_WEBSITE];
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwDenominationStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwCityStatus;
	DBSTATUS m_dwProvinceStatus;
	DBSTATUS m_dwRegionStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwPhoneStatus;
	DBSTATUS m_dwFaxStatus;
	DBSTATUS m_dwContactPersonStatus;
	DBSTATUS m_dwEmailStatus;
	DBSTATUS m_dwWebsiteStatus;
	DBSTATUS m_dwDenomWebsiteStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwDenominationLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwCityLength;
	DBLENGTH m_dwProvinceLength;
	DBLENGTH m_dwRegionLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwPhoneLength;
	DBLENGTH m_dwFaxLength;
	DBLENGTH m_dwContactPersonLength;
	DBLENGTH m_dwEmailLength;
	DBLENGTH m_dwWebsiteLength;
	DBLENGTH m_dwDenomWebsiteLength;
	LONG m_RETURN_VALUE;
	LONG m_churchid; // input value

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspChurchDetailsAccessor, L"{ ? = CALL dbo.spChurchDetails(?) }")

	BEGIN_ACCESSOR_MAP(CspChurchDetailsAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Denomination, m_dwDenominationLength, m_dwDenominationStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_City, m_dwCityLength, m_dwCityStatus)
			COLUMN_ENTRY_LENGTH_STATUS(6, m_Province, m_dwProvinceLength, m_dwProvinceStatus)
			COLUMN_ENTRY_LENGTH_STATUS(7, m_Region, m_dwRegionLength, m_dwRegionStatus)
			COLUMN_ENTRY_LENGTH_STATUS(8, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(9, m_Phone, m_dwPhoneLength, m_dwPhoneStatus)
			COLUMN_ENTRY_LENGTH_STATUS(10, m_Fax, m_dwFaxLength, m_dwFaxStatus)
			COLUMN_ENTRY_LENGTH_STATUS(11, m_ContactPerson, m_dwContactPersonLength, m_dwContactPersonStatus)
			COLUMN_ENTRY_LENGTH_STATUS(12, m_Email, m_dwEmailLength, m_dwEmailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(13, m_Website, m_dwWebsiteLength, m_dwWebsiteStatus)
			COLUMN_ENTRY_LENGTH_STATUS(14, m_DenomWebsite, m_dwDenomWebsiteLength, m_dwDenomWebsiteStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspChurchDetailsAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_churchid)
	END_PARAM_MAP()
};//CspChurchDetailsAccessor
class CspChurchDetails : public CCommand<CAccessor<CspChurchDetailsAccessor> >
{
public:
	LPCWSTR GetFirstTable()
	{
		CAtlStringW strTable;
		if (DBSTATUS_S_ISNULL != m_dwChurchNameStatus)
		{
			strTable += L"<td class=\"colIndex\">Church Name:</td><td class=\"colContent\">";
			strTable += m_ChurchName;
			strTable += L"</td></tr><tr>";
		}
		if (DBSTATUS_S_ISNULL != m_dwAddressStatus)
		{
			strTable += L"<td class=\"colIndex\">Address:</td><td class=\"colContent\">";
			strTable += m_Address;
			strTable += L"</td></tr><tr>";
		}
		if (DBSTATUS_S_ISNULL != m_dwMailStatus)
		{
			strTable += L"<td class=\"colIndex\">Mail:</td><td class=\"colContent\">";
			strTable += m_Mail;
			strTable += L"</td></tr><tr>";
		}
		if (DBSTATUS_S_ISNULL != m_dwCityStatus)
		{
			strTable += L"<td class=\"colIndex\">City:</td><td class=\"colContent\">";
			strTable += m_City;
			strTable += L"</td></tr><tr>";
		}
		if (DBSTATUS_S_ISNULL != m_dwRegionStatus)
		{
			strTable += L"<td class=\"colIndex\">Region:</td><td class=\"colContent\">";
			strTable += m_Region;
			strTable += L"</td></tr><tr>";
		}
		if (DBSTATUS_S_ISNULL != m_dwProvinceStatus)
		{
			strTable += L"<td class=\"colIndex\">Province:</td><td class=\"colContent\">";
			strTable += m_Province;
			strTable += L"</td></tr><tr>";
		}
		if (DBSTATUS_S_ISNULL != m_dwPostalCodeStatus)
		{
			strTable += L"<td class=\"colIndex\">Postal Code:</td><td class=\"colContent\">";
			strTable += m_PostalCode;
			strTable += L"</td></tr><tr>";
		}
		if (DBSTATUS_S_ISNULL != m_dwContactPersonStatus)
		{
			strTable += L"<td class=\"colIndex\">Contact Person:</td><td class=\"colContent\">";
			strTable += m_ContactPerson;
			strTable += L"</td></tr><tr>";
		}
		if (DBSTATUS_S_ISNULL != m_dwDenominationStatus)
		{
			strTable += L"<td class=\"colIndex\">Denomination:</td><td class=\"colContent\">";
			strTable += m_Denomination;
			strTable += L"</td></tr><tr>";
		}
		if (DBSTATUS_S_ISNULL != m_dwPhoneStatus)
		{
			strTable += L"<td class=\"colIndex\">Phone number:</td><td class=\"colContent\">";
			strTable += m_Phone;
			strTable += L"</td></tr><tr>";
		}
		if (DBSTATUS_S_ISNULL != m_dwFaxStatus)
		{
			strTable += L"<td class=\"colIndex\">Fax number:</td><td class=\"colContent\">";
			strTable += m_Fax;
			strTable += L"</td></tr><tr>";
		}
		if (DBSTATUS_S_ISNULL != m_dwEmailStatus)
		{
			strTable += L"<td class=\"colIndex\">Email Address:</td><td class=\"colContent\"><a href=\"mailto:";
			strTable += m_Email;
			strTable += L"\">";
			strTable += m_Email;
			strTable += L"</a></td></tr><tr>";
		}
		if (DBSTATUS_S_ISNULL != m_dwWebsiteStatus)
		{
			strTable += L"<td class=\"colIndex\">Website Address:</td><td class=\"colContent\"><a href=\"";
			CUrl hyperLink;
			hyperLink.CrackUrl(m_Website);
			if (ATL_URL_SCHEME_HTTP != hyperLink.GetScheme())
			{
				if (0 != hyperLink.GetSchemeNameLength())
				{
					//some other scheme is present (eg. ftp, nntp, mailto)
					strTable += m_Website;
				}
				else
				{
					//add a default http:// scheme
					strTable += L"http://";
					strTable += m_Website;
				}
			}
			else
			{
				strTable += m_Website;
			}
			strTable += L"\">";
			strTable += m_Website;
			strTable += L"</a></td></tr><tr>";
		}
		if (DBSTATUS_S_ISNULL != m_dwDenomWebsiteStatus)
		{
			strTable += L"<td class=\"colIndex\">Denominational Website:</td><td class=\"colContent\"><a href=\"";
			CUrl hyperLink;
			hyperLink.CrackUrl(m_DenomWebsite);
			if (ATL_URL_SCHEME_HTTP != hyperLink.GetScheme())
			{
				if (0 != hyperLink.GetSchemeNameLength())
				{
					//some other scheme is present (eg. ftp, nntp, mailto)
					strTable += m_DenomWebsite;
				}
				else
				{
					//add a default http:// scheme
					strTable += L"http://";
					strTable += m_DenomWebsite;
				}
			}
			else
			{
				strTable += m_DenomWebsite;
			}
			strTable += L"\">";
			strTable += m_DenomWebsite;
			strTable += L"</a></td></tr><tr>";
		}
		return ( LPCWSTR ) strTable;
	}
	LPCWSTR GetMetas()
	{
		CAtlStringW strMetas;
		strMetas = L"<meta name=\"Description\" content=\"";
		if (DBSTATUS_S_ISNULL != m_dwChurchNameStatus)
		{
			strMetas += m_ChurchName;
			if (DBSTATUS_S_ISNULL != m_dwDenominationStatus)
			{
				strMetas += L" is a ";
				strMetas += m_Denomination;
				if (DBSTATUS_S_ISNULL != m_dwCityStatus)
				{
					strMetas += L" located in ";
					strMetas += m_City;
				}
			}
		}
		
			
		return ( LPCWSTR ) strMetas;
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spChurchDetails(?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspChurchDetails
/*************************************************************************************/
// Declaration of the CspProvRegCityDenomChurches
class CspProvRegCityDenomChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_region[FIELDSIZE_REGION];
	TCHAR m_city[FIELDSIZE_CITY];
	TCHAR m_denomination[FIELDSIZE_DENOMINATION];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspProvRegCityDenomChurchesAccessor, L"{ ? = CALL dbo.spProvRegCityDenomChurches(?,?,?,?) }")
	
	BEGIN_ACCESSOR_MAP(CspProvRegCityDenomChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(5, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvRegCityDenomChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_region)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(4, m_city)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(5, m_denomination)
	END_PARAM_MAP()
};//CspProvRegCityDenomChurchesAccessor
class CspProvRegCityDenomChurches : public CCommand<CAccessor<CspProvRegCityDenomChurchesAccessor> >
{
public:
	LPCWSTR GetFirstRow(LPCWSTR wszProvince, LPCWSTR wszRegion, LPCWSTR wszCity, LPCWSTR wszDenomination)
	{
		CAtlStringW strFirstRow;
		strFirstRow = L"<td class=\"chName\"><a href=\"index.srf?Province=";
		strFirstRow += wszProvince;
		strFirstRow += L"&Region=";
		strFirstRow += wszRegion;
		strFirstRow += L"&City=";
		strFirstRow += wszCity;
		strFirstRow += L"&Denomination=";
		strFirstRow += wszDenomination;
		strFirstRow += L"&ChurchName=";
		strFirstRow += m_ChurchName;
		strFirstRow += L"\">";
		strFirstRow += m_ChurchName;
		strFirstRow += L"</a></td><td class=\"chAddress\">";
		if (DBSTATUS_S_ISNULL != m_dwAddressStatus)
		{
			strFirstRow += m_Address;
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwPostalCodeStatus)
		{
			strFirstRow += MAPQUEST;
			strFirstRow += L"title=";
			strFirstRow += m_ChurchName;
			strFirstRow += L"&address";
			strFirstRow += m_Address;
			strFirstRow += L"&city=";
			strFirstRow += wszCity;
			strFirstRow += L"&state=";
			strFirstRow += wszProvince;
			strFirstRow += L"&zipcode=";
			strFirstRow += m_PostalCode;
			strFirstRow += MAPQUEST_END;
			strFirstRow += L"</td><td class=\"chCity\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chCity\">";
		}
		strFirstRow += wszCity;
		strFirstRow += L"</td>";
		return ( LPCWSTR ) strFirstRow;
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
		__if_exists(HasBookmark)
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvRegCityDenomChurches(?,?,?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvRegCityDenomChurches
/*************************************************************************************/
//  Declaration of the CspProvRegCityChurches
class CspProvRegCityChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_Denomination[FIELDSIZE_DENOMINATION];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwDenominationStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwDenominationLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_region[FIELDSIZE_REGION];
	TCHAR m_city[FIELDSIZE_CITY];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspProvRegCityChurchesAccessor, L"{ ? = CALL dbo.spProvRegCityChurches(?,?,?) }")

	BEGIN_ACCESSOR_MAP(CspProvRegCityChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_Denomination, m_dwDenominationLength, m_dwDenominationStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(6, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvRegCityChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_region)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(4, m_city)
	END_PARAM_MAP()
};//CspProvRegCityChurchesAccessor
class CspProvRegCityChurches : public CCommand<CAccessor<CspProvRegCityChurchesAccessor> >
{
public:
	LPCWSTR GetFirstRow(LPCWSTR wszProvince, LPCWSTR wszRegion, LPCWSTR wszCity)
	{
		CAtlStringW strFirstRow;
		strFirstRow = L"<td class=\"chName\"><a href=\"index.srf?Province=";
		strFirstRow += wszProvince;
		strFirstRow += L"&Region=";
		strFirstRow += wszRegion;
		strFirstRow += L"&City=";
		strFirstRow += wszCity;
		strFirstRow += L"&Denomination=";
		strFirstRow += m_Denomination;
		strFirstRow += L"&ChurchName=";
		strFirstRow += m_ChurchName;
		strFirstRow += L"\">";
		strFirstRow += m_ChurchName;
		strFirstRow += L"</a></td><td class=\"chAddress\">";
		if (DBSTATUS_S_ISNULL != m_dwAddressStatus)
		{
			strFirstRow += m_Address;
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwPostalCodeStatus)
		{
			strFirstRow += MAPQUEST;
			strFirstRow += L"title=";
			strFirstRow += m_ChurchName;
			strFirstRow += L"&address";
			strFirstRow += m_Address;
			strFirstRow += L"&city=";
			strFirstRow += wszCity;
			strFirstRow += L"&state=";
			strFirstRow += wszProvince;
			strFirstRow += L"&zipcode=";
			strFirstRow += m_PostalCode;
			strFirstRow += MAPQUEST_END;
			strFirstRow += L"</td><td class=\"chDenomination\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chDenomination\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwDenominationStatus)
		{
			strFirstRow += m_Denomination;
			strFirstRow += L"</td><td class=\"chCity\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chCity\">";
		}
		strFirstRow += wszCity;
		strFirstRow += L"</td>";
		return ( LPCWSTR ) strFirstRow;
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvRegCityChurches(?,?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvRegCityChurches
/*************************************************************************************/
// Declaration of the CspProvRegDenomChurches
class CspProvRegDenomChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_City[FIELDSIZE_CITY];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwCityStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwCityLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_region[FIELDSIZE_REGION];
	TCHAR m_denomination[FIELDSIZE_DENOMINATION];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspProvRegDenomChurchesAccessor, L"{ ? = CALL dbo.spProvRegDenomChurches(?,?,?) }")

	BEGIN_ACCESSOR_MAP(CspProvRegDenomChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_City, m_dwCityLength, m_dwCityStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(5, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvRegDenomChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_region)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(4, m_denomination)
	END_PARAM_MAP()
};//CspProvRegDenomChurchesAccessor
class CspProvRegDenomChurches : public CCommand<CAccessor<CspProvRegDenomChurchesAccessor> >
{
public:
	LPCWSTR GetFirstRow(LPCWSTR wszProvince, LPCWSTR wszRegion, LPCWSTR wszDenomination)
	{
		CAtlStringW strFirstRow;
		strFirstRow = L"<td class=\"chName\"><a href=\"index.srf?Province=";
		strFirstRow += wszProvince;
		strFirstRow += L"&Region=";
		strFirstRow += wszRegion;
		if (DBSTATUS_S_ISNULL != m_dwCityStatus)
		{
			strFirstRow += L"&City=";
			strFirstRow += m_City;
		}
		strFirstRow += L"&Denomination=";
		strFirstRow += wszDenomination;
		strFirstRow += L"&ChurchName=";
		strFirstRow += m_ChurchName;
		strFirstRow += L"\">";
		strFirstRow += m_ChurchName;
		strFirstRow += L"</a></td><td class=\"chAddress\">";
		if (DBSTATUS_S_ISNULL != m_dwAddressStatus)
		{
			strFirstRow += m_Address;
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwPostalCodeStatus)
		{
			strFirstRow += MAPQUEST;
			strFirstRow += L"title=";
			strFirstRow += m_ChurchName;
			strFirstRow += L"&address";
			strFirstRow += m_Address;
			strFirstRow += L"&city=";
			strFirstRow += m_City;
			strFirstRow += L"&state=";
			strFirstRow += wszProvince;
			strFirstRow += L"&zipcode=";
			strFirstRow += m_PostalCode;
			strFirstRow += MAPQUEST_END;
			strFirstRow += L"</td><td class=\"chCity\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chCity\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwCityStatus)
		{
			strFirstRow += m_City;
		}
		strFirstRow += L"</td>";
		return ( LPCWSTR ) strFirstRow;
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvRegDenomChurches(?,?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvRegDenomChurches
/*************************************************************************************/
// Declaration of the CspProvRegChurchNameChurches
class CspProvRegChurchNameChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_City[FIELDSIZE_CITY];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwCityStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwCityLength;
	LONG m_ChurchID;
	DBSTATUS m_dwChurchIDStatus;
	DBLENGTH m_dwChurchIDLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_region[FIELDSIZE_REGION];
	TCHAR m_churchname[FIELDSIZE_CHURCHNAME];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_SERVERCURSOR, true);
		pPropSet->AddProperty(DBPROP_OTHERINSERT, true);
		pPropSet->AddProperty(DBPROP_OTHERUPDATEDELETE, true);
		pPropSet->AddProperty(DBPROP_OWNINSERT, true);
		pPropSet->AddProperty(DBPROP_OWNUPDATEDELETE, true);
	}

	DEFINE_COMMAND_EX(CspProvRegChurchNameChurchesAccessor, L"{ ? = CALL dbo.spProvRegChurchNameChurches(?,?,?) }")
	
	BEGIN_ACCESSOR_MAP(CspProvRegChurchNameChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_City, m_dwCityLength, m_dwCityStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_ChurchID, m_dwChurchIDLength, m_dwChurchIDStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(6, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvRegChurchNameChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_region)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(4, m_churchname)
	END_PARAM_MAP()
};//CspProvRegChurchNameChurchesAccessor
class CspProvRegChurchNameChurches : public CCommand<CAccessor<CspProvRegChurchNameChurchesAccessor> >
{
public:
#if (_MSC_VER == 1310)
	long GetChurchID()
#else
#if defined(_M_IA64) || defined (_M_AMD64)
	__int64 GetChurchID()
#else
	__int32 GetChurchID()
#endif
#endif
	{
#if (_MSC_VER == 1310)
		return m_ChurchID;
#else
#if defined(_M_IA64) || defined (_M_AMD64)
		__int64 iChurchID = 0;
		iChurchID = (__int64)m_ChurchID;
#else
		__int32 iChurchID = 0;
		iChurchID = (__int32)m_ChurchID;
#endif
		return iChurchID;
#endif
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvRegChurchNameChurches(?,?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};
//CspProvRegChurchNameChurches
/*************************************************************************************/
// Declaration of the CspProvRegChurches
class CspProvRegChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_City[FIELDSIZE_CITY];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwCityStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwCityLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_region[FIELDSIZE_REGION];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspProvRegChurchesAccessor, L"{ ? = CALL dbo.spProvRegChurches(?,?) }")

	BEGIN_ACCESSOR_MAP(CspProvRegChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_City, m_dwCityLength, m_dwCityStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(6, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvRegChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_region)
	END_PARAM_MAP()
};//CspProvRegChurchesAccessor
class CspProvRegChurches : public CCommand<CAccessor<CspProvRegChurchesAccessor> >
{
public:
	LPCWSTR GetFirstRow(LPCWSTR wszProvince, LPCWSTR wszRegion)
	{
		//PROV_REG_CHURCHES_HEADER
		CAtlStringW strFirstRow;
		strFirstRow = L"<td class=\"chName\"><a href=\"index.srf?Province=";
		strFirstRow += wszProvince;
		strFirstRow += L"&Region=";
		strFirstRow += wszRegion;
		if (DBSTATUS_S_ISNULL != m_dwCityStatus)
		{
			strFirstRow += L"&City=";
			strFirstRow += m_City;
		}
		strFirstRow += L"&ChurchName=";
		strFirstRow += m_ChurchName;
		strFirstRow += L"\">";
		strFirstRow += m_ChurchName;
		strFirstRow += L"</a></td><td class=\"chAddress\">";
		if (DBSTATUS_S_ISNULL != m_dwAddressStatus)
		{
			strFirstRow += m_Address;
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwPostalCodeStatus)
		{
			strFirstRow += MAPQUEST;
			strFirstRow += L"title=";
			strFirstRow += m_ChurchName;
			strFirstRow += L"&address";
			strFirstRow += m_Address;
			strFirstRow += L"&city=";
			strFirstRow += m_City;
			strFirstRow += L"&state=";
			strFirstRow += wszProvince;
			strFirstRow += L"&zipcode=";
			strFirstRow += m_PostalCode;
			strFirstRow += MAPQUEST_END;
			strFirstRow += L"</td><td class=\"chCity\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chCity\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwCityStatus)
		{
			strFirstRow += m_City;
		}
		strFirstRow += L"</td>";
		return ( LPCWSTR ) strFirstRow;
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvRegChurches(?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvRegChurches
/*************************************************************************************/
// Declaration of the CspProvDenomChurchNameChurches
class CspProvDenomChurchNameChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_City[FIELDSIZE_CITY];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwCityStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwCityLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_ChurchID;
	DBSTATUS m_dwChurchIDStatus;
	DBLENGTH m_dwChurchIDLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_denomination[FIELDSIZE_DENOMINATION];
	TCHAR m_churchname[FIELDSIZE_CHURCHNAME];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_SERVERCURSOR, true);
		pPropSet->AddProperty(DBPROP_OTHERINSERT, true);
		pPropSet->AddProperty(DBPROP_OTHERUPDATEDELETE, true);
		pPropSet->AddProperty(DBPROP_OWNINSERT, true);
		pPropSet->AddProperty(DBPROP_OWNUPDATEDELETE, true);
	}

	DEFINE_COMMAND_EX(CspProvDenomChurchNameChurchesAccessor, L"{ ? = CALL dbo.spProvDenomChurchNameChurches(?,?,?) }")
	
	BEGIN_ACCESSOR_MAP(CspProvDenomChurchNameChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_City, m_dwCityLength, m_dwCityStatus)
			COLUMN_ENTRY_LENGTH_STATUS(6, m_ChurchID, m_dwChurchIDLength, m_dwChurchIDStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(7, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvDenomChurchNameChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_denomination)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(4, m_churchname)
	END_PARAM_MAP()
};//CspProvDenomChurchNameChurchesAccessor
class CspProvDenomChurchNameChurches : public CCommand<CAccessor<CspProvDenomChurchNameChurchesAccessor> >
{
public:
#if (_MSC_VER == 1310)
	long GetChurchID()
#else
#if defined(_M_IA64) || defined (_M_AMD64)
	__int64 GetChurchID()
#else
	__int32 GetChurchID()
#endif
#endif
	{
#if (_MSC_VER == 1310)
		return m_ChurchID;
#else
#if defined(_M_IA64) || defined (_M_AMD64)
		__int64 iChurchID = 0;
		iChurchID = (__int64)m_ChurchID;
#else
		__int32 iChurchID = 0;
		iChurchID = (__int32)m_ChurchID;
#endif
		return iChurchID;
#endif
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvDenomChurchNameChurches(?,?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};// CspProvDenomChurchNameChurches
/*************************************************************************************/
// Declaration of the CspProvDenomChurches
class CspProvDenomChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_City[FIELDSIZE_CITY];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwCityStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwCityLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_denomination[FIELDSIZE_DENOMINATION];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspProvDenomChurchesAccessor, L"{ ? = CALL dbo.spProvDenomChurches(?,?) }")

	BEGIN_ACCESSOR_MAP(CspProvDenomChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_City, m_dwCityLength, m_dwCityStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(6, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvDenomChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_denomination)
	END_PARAM_MAP()
};//CspProvDenomChurchesAccessor
class CspProvDenomChurches : public CCommand<CAccessor<CspProvDenomChurchesAccessor> >
{
public:
	LPCWSTR GetFirstRow(LPCWSTR wszProvince, LPCWSTR wszDenomination)
	{
		//PROV_DENOM_CHURCHES_HEADER
		CAtlStringW strFirstRow;
		strFirstRow = L"<td class=\"chName\"><a href=\"index.srf?Province=";
		strFirstRow += wszProvince;
		if (DBSTATUS_S_ISNULL != m_dwCityStatus)
		{
			strFirstRow += L"&City=";
			strFirstRow += m_City;
		}
		strFirstRow += L"&ChurchName=";
		strFirstRow += m_ChurchName;
		strFirstRow += L"\">";
		strFirstRow += m_ChurchName;
		strFirstRow += L"</a></td><td class=\"chAddress\">";
		if (DBSTATUS_S_ISNULL != m_dwAddressStatus)
		{
			strFirstRow += m_Address;
			strFirstRow += L"</td><td class=\"chMail\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chMail\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwMailStatus)
		{
			strFirstRow += m_Mail;
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwPostalCodeStatus)
		{
			strFirstRow += MAPQUEST;
			strFirstRow += L"title=";
			strFirstRow += m_ChurchName;
			strFirstRow += L"&address";
			strFirstRow += m_Address;
			strFirstRow += L"&city=";
			strFirstRow += m_City;
			strFirstRow += L"&state=";
			strFirstRow += wszProvince;
			strFirstRow += L"&zipcode=";
			strFirstRow += m_PostalCode;
			strFirstRow += MAPQUEST_END;
			strFirstRow += L"</td><td class=\"chCity\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chCity\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwCityStatus)
		{
			strFirstRow += m_City;
		}
		strFirstRow += L"</td>";
		return ( LPCWSTR ) strFirstRow;
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvDenomChurches(?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvDenomChurches
/*************************************************************************************/
// Declaration of the CspProvCityDenomChurches
class CspProvCityDenomChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_city[FIELDSIZE_CITY];
	TCHAR m_denomination[FIELDSIZE_DENOMINATION];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspProvCityDenomChurchesAccessor, L"{ ? = CALL dbo.spProvCityDenomChurches(?,?,?) }")

	BEGIN_ACCESSOR_MAP(CspProvCityDenomChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(5, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvCityDenomChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_city)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(4, m_denomination)
	END_PARAM_MAP()
};//CspProvCityDenomChurchesAccessor
class CspProvCityDenomChurches : public CCommand<CAccessor<CspProvCityDenomChurchesAccessor> >
{
public:
	LPCWSTR GetFirstRow(LPCWSTR wszProvince, LPCWSTR wszCity, LPCWSTR wszDenomination)
	{
		//PROV_REG_CHURCHES_HEADER
		CAtlStringW strFirstRow;
		strFirstRow = L"<td class=\"chName\"><a href=\"index.srf?Province=";
		strFirstRow += wszProvince;
		strFirstRow += L"&City=";
		strFirstRow += wszCity;
		strFirstRow += L"&ChurchName=";
		strFirstRow += m_ChurchName;
		strFirstRow += L"\">";
		strFirstRow += m_ChurchName;
		strFirstRow += L"</a></td><td class=\"chAddress\">";
		if (DBSTATUS_S_ISNULL != m_dwAddressStatus)
		{
			strFirstRow += m_Address;
			strFirstRow += L"</td><td class=\"chMail\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chMail\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwMailStatus)
		{
			strFirstRow += m_Mail;
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwPostalCodeStatus)
		{
			strFirstRow += MAPQUEST;
			strFirstRow += L"title=";
			strFirstRow += m_ChurchName;
			strFirstRow += L"&address";
			strFirstRow += m_Address;
			strFirstRow += L"&city=";
			strFirstRow += wszCity;
			strFirstRow += L"&state=";
			strFirstRow += wszProvince;
			strFirstRow += L"&zipcode=";
			strFirstRow += m_PostalCode;
			strFirstRow += MAPQUEST_END;
		}
		strFirstRow += L"</td>";
		return ( LPCWSTR ) strFirstRow;
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvCityDenomChurches(?,?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvCityDenomChurches
/*************************************************************************************/
// Declaration of the CspProvChurchNameChurches
class CspProvChurchNameChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_City[FIELDSIZE_CITY];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwCityStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwCityLength;
	LONG m_ChurchID;
	DBSTATUS m_dwChurchIDStatus;
	DBLENGTH m_dwChurchIDLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_churchname[FIELDSIZE_CHURCHNAME];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_SERVERCURSOR, true);
		pPropSet->AddProperty(DBPROP_OTHERINSERT, true);
		pPropSet->AddProperty(DBPROP_OTHERUPDATEDELETE, true);
		pPropSet->AddProperty(DBPROP_OWNINSERT, true);
		pPropSet->AddProperty(DBPROP_OWNUPDATEDELETE, true);
	}

	DEFINE_COMMAND_EX(CspProvChurchNameChurchesAccessor, L"{ ? = CALL dbo.spProvChurchNameChurches(?,?) }")

	BEGIN_ACCESSOR_MAP(CspProvChurchNameChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_City, m_dwCityLength, m_dwCityStatus)
			COLUMN_ENTRY_LENGTH_STATUS(6, m_ChurchID, m_dwChurchIDLength, m_dwChurchIDStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(7, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvChurchNameChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_churchname)
	END_PARAM_MAP()
};//CspProvChurchNameChurchesAccessor
class CspProvChurchNameChurches : public CCommand<CAccessor<CspProvChurchNameChurchesAccessor> >
{
public:
#if (_MSC_VER == 1310)
	long GetChurchID()
#else
#if defined(_M_IA64) || defined (_M_AMD64)
	__int64 GetChurchID()
#else
	__int32 GetChurchID()
#endif
#endif
	{
#if (_MSC_VER == 1310)
		return m_ChurchID;
#else
#if defined(_M_IA64) || defined (_M_AMD64)
		__int64 iChurchID = 0;
		iChurchID = (__int64)m_ChurchID;
#else
		__int32 iChurchID = 0;
		iChurchID = (__int32)m_ChurchID;
#endif
		return iChurchID;
#endif
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvChurchNameChurches(?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvChurchNameChurches
/*************************************************************************************/
// Declaration of the CspProvChurches
class CspProvChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_City[FIELDSIZE_CITY];
	TCHAR m_Region[FIELDSIZE_REGION];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwCityStatus;
	DBSTATUS m_dwRegionStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwCityLength;
	DBLENGTH m_dwRegionLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspProvChurchesAccessor, L"{ ? = CALL dbo.spProvChurches(?) }")

	BEGIN_ACCESSOR_MAP(CspProvChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_City, m_dwCityLength, m_dwCityStatus)
			COLUMN_ENTRY_LENGTH_STATUS(6, m_Region, m_dwRegionLength, m_dwRegionStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(7, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
	END_PARAM_MAP()
};//CspProvChurchesAccessor
class CspProvChurches : public CCommand<CAccessor<CspProvChurchesAccessor> >
{
public:
	LPCWSTR GetFirstRow(LPCWSTR wszProvince)
	{
		CAtlStringW strFirstRow;
		strFirstRow = L"<td class=\"chName\"><a href=\"index.srf?Province=";
		strFirstRow += wszProvince;
		if (DBSTATUS_S_ISNULL != m_dwCityStatus)
		{
			strFirstRow += L"&City=";
			strFirstRow += m_City;
		}
		strFirstRow += L"&ChurchName=";
		strFirstRow += m_ChurchName;
		strFirstRow += L"\">";
		strFirstRow += m_ChurchName;
		strFirstRow += L"</a></td><td class=\"chAddress\">";
		if (DBSTATUS_S_ISNULL != m_dwAddressStatus)
		{
			strFirstRow += m_Address;
			strFirstRow += L"</td><td class=\"chMail\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chMail\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwMailStatus)
		{
			strFirstRow += m_Mail;
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwPostalCodeStatus)
		{
			strFirstRow += MAPQUEST;
			strFirstRow += L"title=";
			strFirstRow += m_ChurchName;
			strFirstRow += L"&address";
			strFirstRow += m_Address;
			strFirstRow += L"&city=";
			strFirstRow += m_City;
			strFirstRow += L"&state=";
			strFirstRow += wszProvince;
			strFirstRow += L"&zipcode=";
			strFirstRow += m_PostalCode;
			strFirstRow += MAPQUEST_END;
			strFirstRow += L"</td><td class=\"chCity\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chCity\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwCityStatus)
		{	
			strFirstRow += m_City;
			strFirstRow += L"</td><td class=\"chRegion\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chRegion\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwRegionStatus)
		{
			strFirstRow += m_Region;
		}
		strFirstRow += L"</td>";
		return ( LPCWSTR ) strFirstRow;
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvChurches(?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvChurches
/*************************************************************************************/
// Declaration of the CspRegCityDenomChurches
class CspRegCityDenomChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_region[FIELDSIZE_REGION];
	TCHAR m_city[FIELDSIZE_CITY];
	TCHAR m_denomination[FIELDSIZE_DENOMINATION];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspRegCityDenomChurchesAccessor, L"{ ? = CALL dbo.spRegCityDenomChurches(?,?,?) }")

	BEGIN_ACCESSOR_MAP(CspRegCityDenomChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(5, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspRegCityDenomChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_region)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_city)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(4, m_denomination)
	END_PARAM_MAP()
};//CspRegCityDenomChurchesAccessor
class CspRegCityDenomChurches : public CCommand<CAccessor<CspRegCityDenomChurchesAccessor> >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spRegCityDenomChurches(?,?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspRegCityDenomChurches
/*************************************************************************************/
// Declaration of the CspRegCityChurches
class CspRegCityChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_Denomination[FIELDSIZE_DENOMINATION];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwDenominationStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwDenominationLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_region[FIELDSIZE_REGION];
	TCHAR m_city[FIELDSIZE_CITY];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspRegCityChurchesAccessor, L"{ ? = CALL dbo.spRegCityChurches(?,?) }")

	BEGIN_ACCESSOR_MAP(CspRegCityChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_Denomination, m_dwDenominationLength, m_dwDenominationStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(6, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspRegCityChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_region)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_city)
	END_PARAM_MAP()
};//CspRegCityChurchesAccessor
class CspRegCityChurches : public CCommand<CAccessor<CspRegCityChurchesAccessor> >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spRegCityChurches(?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspRegCityChurches
/*************************************************************************************/
// Declaration of the CspRegDenomChurches
class CspRegDenomChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_City[FIELDSIZE_CITY];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwCityStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwCityLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_region[FIELDSIZE_REGION];
	TCHAR m_denomination[FIELDSIZE_DENOMINATION];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspRegDenomChurchesAccessor, L"{ ? = CALL dbo.spRegDenomChurches(?,?) }")

	BEGIN_ACCESSOR_MAP(CspRegDenomChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_City, m_dwCityLength, m_dwCityStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(6, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspRegDenomChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_region)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_denomination)
	END_PARAM_MAP()
};//CspRegDenomChurchesAccessor
class CspRegDenomChurches : public CCommand<CAccessor<CspRegDenomChurchesAccessor> >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spRegDenomChurches(?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspRegDenomChurches
/*************************************************************************************/
// Declaration of the CspRegChurches
class CspRegChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_City[FIELDSIZE_CITY];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwCityStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwCityLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_region[FIELDSIZE_REGION];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspRegChurchesAccessor, L"{ ? = CALL dbo.spRegChurches(?) }")

	BEGIN_ACCESSOR_MAP(CspRegChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_City, m_dwCityLength, m_dwCityStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(6, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspRegChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_region)
	END_PARAM_MAP()
};//CspRegChurchesAccessor
class CspRegChurches : public CCommand<CAccessor<CspRegChurchesAccessor> >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spRegChurches(?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspRegChurches
/*************************************************************************************/
// Declaration of the CspCityDenomChurches
class CspCityDenomChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_city[FIELDSIZE_CITY];
	TCHAR m_denomination[FIELDSIZE_DENOMINATION];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspCityDenomChurchesAccessor, L"{ ? = CALL dbo.spCityDenomChurches(?,?) }")

	BEGIN_ACCESSOR_MAP(CspCityDenomChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(5, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspCityDenomChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_city)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_denomination)
	END_PARAM_MAP()
};//CspCityDenomChurchesAccessor
class CspCityDenomChurches : public CCommand<CAccessor<CspCityDenomChurchesAccessor> >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spCityDenomChurches(?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspCityDenomChurches
/*************************************************************************************/
// Declaration of the CspCityChurches
class CspCityChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_city[FIELDSIZE_CITY];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspCityChurchesAccessor, L"{ ? = CALL dbo.spCityChurches(?) }")

	BEGIN_ACCESSOR_MAP(CspCityChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(5, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspCityChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_city)
	END_PARAM_MAP()
};//CspCityChurchesAccessor

class CspCityChurches : public CCommand<CAccessor<CspCityChurchesAccessor> >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spCityChurches(?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspCityChurches
/*************************************************************************************/
// Declaration of the CspDenomChurches
class CspDenomChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_City[FIELDSIZE_CITY];
	TCHAR m_Province[FIELDSIZE_PROVINCE];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwCityStatus;
	DBSTATUS m_dwProvinceStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwCityLength;
	DBLENGTH m_dwProvinceLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_denomination[FIELDSIZE_DENOMINATION];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspDenomChurchesAccessor, L"{ ? = CALL dbo.spDenomChurches(?) }")

	BEGIN_ACCESSOR_MAP(CspDenomChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_City, m_dwCityLength, m_dwCityStatus)
			COLUMN_ENTRY_LENGTH_STATUS(6, m_Province, m_dwProvinceLength, m_dwProvinceStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(7, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspDenomChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_denomination)
	END_PARAM_MAP()
};//CspDenomChurchesAccessor

class CspDenomChurches : public CCommand<CAccessor<CspDenomChurchesAccessor> >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spDenomChurches(?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspDenomChurches
/*************************************************************************************/
// Declaration of the CspAllProvs
class CspAllProvsAccessor
{
public:
	TCHAR m_Province[FIELDSIZE_PROVINCE];
	DBSTATUS m_dwProvinceStatus;
	DBLENGTH m_dwProvinceLength;
	LONG m_RETURN_VALUE;

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspAllProvsAccessor, L"{ ? = CALL dbo.spAllProvs }")

	BEGIN_ACCESSOR_MAP(CspAllProvsAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_Province, m_dwProvinceLength, m_dwProvinceStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspAllProvsAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
	END_PARAM_MAP()
};//CspAllProvsAccessor
class CspAllProvs : public CCommand<CAccessor<CspAllProvsAccessor> >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spAllProvs }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspAllProvs
/*************************************************************************************/
// Declaration of the CspProvRegions
class CspProvRegionsAccessor
{
public:
	TCHAR m_Region[FIELDSIZE_REGION];
	DBSTATUS m_dwRegionStatus;
	DBLENGTH m_dwRegionLength;

	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspProvRegionsAccessor, L"{ ? = CALL dbo.spProvRegions(?) }")

	BEGIN_ACCESSOR_MAP(CspProvRegionsAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_Region, m_dwRegionLength, m_dwRegionStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvRegionsAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
	END_PARAM_MAP()
};//CspProvRegionsAccessor
class CspProvRegions : public CCommand<CAccessor<CspProvRegionsAccessor> >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvRegions(?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvRegions
/*************************************************************************************/
// Declaration of the CspProvRegCities
class CspProvRegCitiesAccessor
{
public:
	TCHAR m_City[FIELDSIZE_CITY];
	DBSTATUS m_dwCityStatus;
	DBLENGTH m_dwCityLength;

	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_region[FIELDSIZE_REGION];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspProvRegCitiesAccessor, L"{ ? = CALL dbo.spProvRegCities(?,?) }")

	BEGIN_ACCESSOR_MAP(CspProvRegCitiesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_City, m_dwCityLength, m_dwCityStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvRegCitiesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_region)
	END_PARAM_MAP()
};//CspProvRegCitiesAccessor
class CspProvRegCities : public CCommand<CAccessor<CspProvRegCitiesAccessor> >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvRegCities(?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvRegCities
/*************************************************************************************/
// Declaration of the CspProvRegCityDenoms
class CspProvRegCityDenomsAccessor
{
public:
	TCHAR m_Denomination[FIELDSIZE_DENOMINATION];
	DBSTATUS m_dwDenominationStatus;
	DBLENGTH m_dwDenominationLength;

	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_region[FIELDSIZE_REGION];
	TCHAR m_city[FIELDSIZE_CITY];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspProvRegCityDenomsAccessor, L"{ ? = CALL dbo.spProvRegCityDenoms(?,?,?) }")

	BEGIN_ACCESSOR_MAP(CspProvRegCityDenomsAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_Denomination, m_dwDenominationLength, m_dwDenominationStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvRegCityDenomsAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_region)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(4, m_city)
	END_PARAM_MAP()
};//CspProvRegCityDenomsAccessor
class CspProvRegCityDenoms : public CCommand<CAccessor<CspProvRegCityDenomsAccessor> >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvRegCityDenoms(?,?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvRegCityDenoms
/*************************************************************************************/
// Declaration of the CspProvRegDenomCities
class CspProvRegDenomCitiesAccessor
{
public:
	TCHAR m_City[FIELDSIZE_CITY];
	DBSTATUS m_dwCityStatus;
	DBLENGTH m_dwCityLength;

	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_region[FIELDSIZE_REGION];
	TCHAR m_denomination[FIELDSIZE_DENOMINATION];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspProvRegDenomCitiesAccessor, L"{ ? = CALL dbo.spProvRegDenomCities(?,?,?) }")

	BEGIN_ACCESSOR_MAP(CspProvRegDenomCitiesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_City, m_dwCityLength, m_dwCityStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvRegDenomCitiesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_region)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(4, m_denomination)
	END_PARAM_MAP()
};//CspProvRegDenomCitiesAccessor
class CspProvRegDenomCities : public CCommand<CAccessor<CspProvRegDenomCitiesAccessor> >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvRegDenomCities(?,?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvRegDenomCities
/*************************************************************************************/
// Declaration of the CspProvDenomCities
class CspProvDenomCitiesAccessor
{
public:
	TCHAR m_City[FIELDSIZE_CITY];
	DBSTATUS m_dwCityStatus;
	DBLENGTH m_dwCityLength;

	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_denomination[FIELDSIZE_DENOMINATION];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspProvDenomCitiesAccessor, L"{ ? = CALL dbo.spProvDenomCities(?,?) }")

	BEGIN_ACCESSOR_MAP(CspProvDenomCitiesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_City, m_dwCityLength, m_dwCityStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvDenomCitiesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_denomination)
	END_PARAM_MAP()
};
class CspProvDenomCities : public CCommand<CAccessor<CspProvDenomCitiesAccessor> >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvDenomCities(?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvDenomCities
/*************************************************************************************/
// Declaration of the CspProvCityChurches
class CspProvCityChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_Denomination[FIELDSIZE_DENOMINATION];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwDenominationStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwDenominationLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_city[FIELDSIZE_CITY];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_QUICKRESTART, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspProvCityChurchesAccessor, L"{ ? = CALL dbo.spProvCityChurches(?,?) }")

	BEGIN_ACCESSOR_MAP(CspProvCityChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_Denomination, m_dwDenominationLength, m_dwDenominationStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(6, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvCityChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_city)
	END_PARAM_MAP()
};//CspProvCityChurchesAccessor
class CspProvCityChurches : public CCommand<CAccessor<CspProvCityChurchesAccessor> >
{
public:
	LPCWSTR GetFirstRow(LPCWSTR wszProvince, LPCWSTR wszCity)
	{
		CAtlStringW strFirstRow;
		strFirstRow = L"<td class=\"chName\"><a href=\"index.srf?Province=";
		strFirstRow += wszProvince;
		strFirstRow += L"&City=";
		strFirstRow += wszCity;
		if (DBSTATUS_S_ISNULL != m_dwDenominationStatus)
		{
			strFirstRow += L"&Denomination=";
			strFirstRow += m_Denomination;
		}
		strFirstRow += L"&ChurchName=";
		strFirstRow += m_ChurchName;
		strFirstRow += L"\">";
		strFirstRow += m_ChurchName;
		strFirstRow += L"</a></td><td class=\"chAddress\">";
		if (DBSTATUS_S_ISNULL != m_dwAddressStatus)
		{
			strFirstRow += m_Address;
			strFirstRow += L"</td><td class=\"chMail\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chMail\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwMailStatus)
		{
			strFirstRow += m_Mail;
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwPostalCodeStatus)
		{
			strFirstRow += MAPQUEST;
			strFirstRow += L"title=";
			strFirstRow += m_ChurchName;
			strFirstRow += L"&address";
			strFirstRow += m_Address;
			strFirstRow += L"&city=";
			strFirstRow += wszCity;
			strFirstRow += L"&state=";
			strFirstRow += wszProvince;
			strFirstRow += L"&zipcode=";
			strFirstRow += m_PostalCode;
			strFirstRow += MAPQUEST_END;
			strFirstRow += L"</td><td class=\"chDenomination\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chDenomination\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwDenominationStatus)
		{
			strFirstRow += m_Denomination;
		}
		strFirstRow += L"</td>";
		return ( LPCWSTR ) strFirstRow;
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(HasBookmark) 	
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvCityChurches(?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvCityChurches
/*************************************************************************************/
// Declaration of CChurchNameChurces
class CChurchNameChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_City[FIELDSIZE_CITY];
	TCHAR m_Province[FIELDSIZE_PROVINCE];
	TCHAR m_Denomination[FIELDSIZE_DENOMINATION];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwCityStatus;
	DBSTATUS m_dwProvinceStatus;
	DBSTATUS m_dwDenominationStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwCityLength;
	DBLENGTH m_dwProvinceLength;
	DBLENGTH m_dwDenominationLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_churchname[FIELDSIZE_CHURCHNAME];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_SERVERCURSOR, true);
		pPropSet->AddProperty(DBPROP_OTHERINSERT, true);
		pPropSet->AddProperty(DBPROP_OTHERUPDATEDELETE, true);
		pPropSet->AddProperty(DBPROP_OWNINSERT, true);
		pPropSet->AddProperty(DBPROP_OWNUPDATEDELETE, true);
	}

    DEFINE_COMMAND_EX(CChurchNameChurchesAccessor, L"{ ? = CALL spChurchNameChurches(?) }")
        
    BEGIN_ACCESSOR_MAP(CChurchNameChurchesAccessor, 1)
        BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_City, m_dwCityLength, m_dwCityStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_Province, m_dwProvinceLength, m_dwProvinceStatus)
			COLUMN_ENTRY_LENGTH_STATUS(6, m_Denomination, m_dwDenominationLength, m_dwDenominationStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(7, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
        END_ACCESSOR()
    END_ACCESSOR_MAP()
        
    BEGIN_PARAM_MAP(CChurchNameChurchesAccessor)
        SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
        COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_churchname)
    END_PARAM_MAP()
};//CChurchNameChurchesAccessor
class CChurchNameChurches : public CCommand<CAccessor<CChurchNameChurchesAccessor> >
{
public:
    HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL spChurchNameChurches(?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CChurchNameChurches
/*************************************************************************************/
// Declaration of the CspProvRegCityDenomChurchNameChurches
class CspProvRegCityDenomChurchNameChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_ChurchID;
	DBSTATUS m_dwChurchIDStatus;
	DBLENGTH m_dwChurchIDLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_region[FIELDSIZE_REGION];
	TCHAR m_city[FIELDSIZE_CITY];
	TCHAR m_denomination[FIELDSIZE_DENOMINATION];
	TCHAR m_churchname[FIELDSIZE_CHURCHNAME];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_SERVERCURSOR, true);
		pPropSet->AddProperty(DBPROP_OTHERINSERT, true);
		pPropSet->AddProperty(DBPROP_OTHERUPDATEDELETE, true);
		pPropSet->AddProperty(DBPROP_OWNINSERT, true);
		pPropSet->AddProperty(DBPROP_OWNUPDATEDELETE, true);
	}

	DEFINE_COMMAND_EX(CspProvRegCityDenomChurchNameChurchesAccessor, L"{ ? = CALL dbo.spProvRegCityDenomChurchNameChurches(?,?,?,?,?) }")
	
	BEGIN_ACCESSOR_MAP(CspProvRegCityDenomChurchNameChurchesAccessor, 1)
        BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_ChurchID, m_dwChurchIDLength, m_dwChurchIDStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(6, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
        END_ACCESSOR()
    END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvRegCityDenomChurchNameChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_region)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(4, m_city)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(5, m_denomination)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(6, m_churchname)
	END_PARAM_MAP()
};//CspProvRegCityDenomChurchNameChurchesAccessor
class CspProvRegCityDenomChurchNameChurches : public CCommand<CAccessor<CspProvRegCityDenomChurchNameChurchesAccessor> >
{
public:
#if (_MSC_VER == 1310)
	long GetChurchID()
#else
#if defined(_M_IA64) || defined (_M_AMD64)
	__int64 GetChurchID()
#else
	__int32 GetChurchID()
#endif
#endif
	{
#if (_MSC_VER == 1310)
		return m_ChurchID;
#else
#if defined(_M_IA64) || defined (_M_AMD64)
		__int64 iChurchID = 0;
		iChurchID = (__int64)m_ChurchID;
#else
		__int32 iChurchID = 0;
		iChurchID = (__int32)m_ChurchID;
#endif
		return iChurchID;
#endif
	}
	LPCWSTR GetFirstRow(LPCWSTR wszProvince, LPCWSTR wszRegion, LPCWSTR wszCity, LPCWSTR wszDenomination, LPCWSTR wszChurchName)
	{
		CAtlStringW strFirstRow;
		strFirstRow = L"<td class=\"chName\"><a href=\"index.srf?Province=";
		strFirstRow += wszProvince;
		strFirstRow += L"&Region=";
		strFirstRow += wszRegion;
		strFirstRow += L"&City=";
		strFirstRow += wszCity;
		strFirstRow += L"&Denomination=";
		strFirstRow += wszDenomination;
		strFirstRow += L"&ChurchName=";
		strFirstRow += wszChurchName;
		strFirstRow += L"\">";
		strFirstRow += wszChurchName;
		strFirstRow += L"</a></td><td class=\"chAddress\">";
		if (DBSTATUS_S_ISNULL != m_dwAddressStatus)
		{
			strFirstRow += m_Address;
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chPostalCode\">";
		}
		if (DBSTATUS_S_ISNULL != m_dwPostalCodeStatus)
		{
			strFirstRow += MAPQUEST;
			strFirstRow += L"title=";
			strFirstRow += wszChurchName;
			strFirstRow += L"&address";
			strFirstRow += m_Address;
			strFirstRow += L"&city=";
			strFirstRow += wszCity;
			strFirstRow += L"&state=";
			strFirstRow += wszProvince;
			strFirstRow += L"&zipcode=";
			strFirstRow += m_PostalCode;
			strFirstRow += MAPQUEST_END;
			strFirstRow += L"</td><td class=\"chCity\">";
		}
		else
		{
			strFirstRow += L"</td><td class=\"chCity\">";
		}
		strFirstRow += wszCity;
		strFirstRow += L"</td>";
		return ( LPCWSTR ) strFirstRow;
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvRegCityDenomChurchNameChurches(?,?,?,?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvRegCityDenomChurchNameChurches
/*************************************************************************************/
// Declaration of the CspProvRegCityChurchNameChurches
// uses full text indexing - no scrolling supported
class CspProvRegCityChurchNameChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	LONG m_ChurchID;
	DBSTATUS m_dwChurchIDStatus;
	DBLENGTH m_dwChurchIDLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_region[FIELDSIZE_REGION];
	TCHAR m_city[FIELDSIZE_CITY];
	TCHAR m_churchname[FIELDSIZE_CHURCHNAME];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_SERVERCURSOR, true);
		pPropSet->AddProperty(DBPROP_OTHERINSERT, true);
		pPropSet->AddProperty(DBPROP_OTHERUPDATEDELETE, true);
		pPropSet->AddProperty(DBPROP_OWNINSERT, true);
		pPropSet->AddProperty(DBPROP_OWNUPDATEDELETE, true);
	}

	DEFINE_COMMAND_EX(CspProvRegCityChurchNameChurchesAccessor, L"{ ? = CALL dbo.spProvRegCityChurchNameChurches(?,?,?,?) }")
	
	BEGIN_ACCESSOR_MAP(CspProvRegCityChurchNameChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_ChurchID, m_dwChurchIDLength, m_dwChurchIDStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(6, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvRegCityChurchNameChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_region)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(4, m_city)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(5, m_churchname)
	END_PARAM_MAP()
};//CspProvRegCityChurchNameChurchesAccessor
class CspProvRegCityChurchNameChurches : public CCommand<CAccessor<CspProvRegCityChurchNameChurchesAccessor> >
{
public:
#if (_MSC_VER == 1310)
	long GetChurchID()
#else
#if defined(_M_IA64) || defined (_M_AMD64)
	__int64 GetChurchID()
#else
	__int32 GetChurchID()
#endif
#endif
	{
#if (_MSC_VER == 1310)
		return m_ChurchID;
#else
#if defined(_M_IA64) || defined (_M_AMD64)
		__int64 iChurchID = 0;
		iChurchID = (__int64)m_ChurchID;
#else
		__int32 iChurchID = 0;
		iChurchID = (__int32)m_ChurchID;
#endif
		return iChurchID;
#endif
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvRegCityChurchNameChurches(?,?,?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvRegCityChurchNameChurches
/*************************************************************************************/
// Declaration of the CspProvRegDenomChurchNameChurches
class CspProvRegDenomChurchNameChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	TCHAR m_City[FIELDSIZE_CITY];
	LONG m_ChurchID;
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwCityStatus;
	DBSTATUS m_dwChurchIDStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwCityLength;
	DBLENGTH m_dwChurchIDLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_region[FIELDSIZE_REGION];
	TCHAR m_denomination[FIELDSIZE_DENOMINATION];
	TCHAR m_churchname[FIELDSIZE_CHURCHNAME];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_SERVERCURSOR, true);
		pPropSet->AddProperty(DBPROP_OTHERINSERT, true);
		pPropSet->AddProperty(DBPROP_OTHERUPDATEDELETE, true);
		pPropSet->AddProperty(DBPROP_OWNINSERT, true);
		pPropSet->AddProperty(DBPROP_OWNUPDATEDELETE, true);
	}

	DEFINE_COMMAND_EX(CspProvRegDenomChurchNameChurchesAccessor, L"{ ? = CALL dbo.spProvRegDenomChurchNameChurches(?,?,?,?) }")

	BEGIN_ACCESSOR_MAP(CspProvRegDenomChurchNameChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_City, m_dwCityLength, m_dwCityStatus)
			COLUMN_ENTRY_LENGTH_STATUS(6, m_ChurchID, m_dwChurchIDLength, m_dwChurchIDStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(7, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvRegDenomChurchNameChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_region)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(4, m_denomination)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(5, m_churchname)
	END_PARAM_MAP()
};//CspProvRegDenomChurchNameChurchesAccessor
class CspProvRegDenomChurchNameChurches : public CCommand<CAccessor<CspProvRegDenomChurchNameChurchesAccessor> >
{
public:
#if (_MSC_VER == 1310)
	long GetChurchID()
#else
#if defined(_M_IA64) || defined (_M_AMD64)
	__int64 GetChurchID()
#else
	__int32 GetChurchID()
#endif
#endif
	{
#if (_MSC_VER == 1310)
		return m_ChurchID;
#else
#if defined(_M_IA64) || defined (_M_AMD64)
		__int64 iChurchID = 0;
		iChurchID = (__int64)m_ChurchID;
#else
		__int32 iChurchID = 0;
		iChurchID = (__int32)m_ChurchID;
#endif
		return iChurchID;
#endif
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvRegDenomChurchNameChurches(?,?,?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvRegDenomChurchNameChurches
/*************************************************************************************/
// Declaration of the CspProvCityDenomChurchNameChurches
class CspProvCityDenomChurchNameChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	LONG m_ChurchID;
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwChurchIDStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwChurchIDLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_city[FIELDSIZE_CITY];
	TCHAR m_denomination[FIELDSIZE_DENOMINATION];
	TCHAR m_churchname[74];

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_SERVERCURSOR, true);
		pPropSet->AddProperty(DBPROP_OTHERINSERT, true);
		pPropSet->AddProperty(DBPROP_OTHERUPDATEDELETE, true);
		pPropSet->AddProperty(DBPROP_OWNINSERT, true);
		pPropSet->AddProperty(DBPROP_OWNUPDATEDELETE, true);
	}

	DEFINE_COMMAND_EX(CspProvCityDenomChurchNameChurchesAccessor, L"{ ? = CALL dbo.spProvCityDenomChurchNameChurches(?,?,?,?) }")

	BEGIN_ACCESSOR_MAP(CspProvCityDenomChurchNameChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_ChurchID, m_dwChurchIDLength, m_dwChurchIDStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(6, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvCityDenomChurchNameChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_city)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(4, m_denomination)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(5, m_churchname)
	END_PARAM_MAP()
};//CspProvCityDenomChurchNameChurchesAccessor
class CspProvCityDenomChurchNameChurches : public CCommand<CAccessor<CspProvCityDenomChurchNameChurchesAccessor> >
{
public:
#if (_MSC_VER == 1310)
	long GetChurchID()
#else
#if defined(_M_IA64) || defined (_M_AMD64)
	__int64 GetChurchID()
#else
	__int32 GetChurchID()
#endif
#endif
	{
#if (_MSC_VER == 1310)
		return m_ChurchID;
#else
#if defined(_M_IA64) || defined (_M_AMD64)
		__int64 iChurchID = 0;
		iChurchID = (__int64)m_ChurchID;
#else
		__int32 iChurchID = 0;
		iChurchID = (__int32)m_ChurchID;
#endif
		return iChurchID;
#endif
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvCityDenomChurchNameChurches(?,?,?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvCityDenomChurchNameChurches
/*************************************************************************************/
// Declaration of the CspProvCityChurchNameChurches
class CspProvCityChurchNameChurchesAccessor
{
public:
	TCHAR m_ChurchName[FIELDSIZE_CHURCHNAME];
	TCHAR m_Address[FIELDSIZE_ADDRESS];
	TCHAR m_Mail[FIELDSIZE_MAIL];
	TCHAR m_PostalCode[FIELDSIZE_POSTAL_CODE];
	LONG m_ChurchID;
	DBTIMESTAMP m_Modified;
	DBSTATUS m_dwChurchNameStatus;
	DBSTATUS m_dwAddressStatus;
	DBSTATUS m_dwMailStatus;
	DBSTATUS m_dwPostalCodeStatus;
	DBSTATUS m_dwChurchIDStatus;
	DBSTATUS m_dwModifiedStatus;
	DBLENGTH m_dwChurchNameLength;
	DBLENGTH m_dwAddressLength;
	DBLENGTH m_dwMailLength;
	DBLENGTH m_dwPostalCodeLength;
	DBLENGTH m_dwChurchIDLength;
	DBLENGTH m_dwModifiedLength;
	LONG m_RETURN_VALUE;
	TCHAR m_province[FIELDSIZE_PROVINCE];
	TCHAR m_city[FIELDSIZE_CITY];
	TCHAR m_churchname[FIELDSIZE_CHURCHNAME];
	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		//specify a fast forward cursor (sql server 2000 compatible)
		pPropSet->AddProperty(DBPROP_SERVERCURSOR, true);
		pPropSet->AddProperty(DBPROP_OTHERINSERT, true);
		pPropSet->AddProperty(DBPROP_OTHERUPDATEDELETE, true);
		pPropSet->AddProperty(DBPROP_OWNINSERT, true);
		pPropSet->AddProperty(DBPROP_OWNUPDATEDELETE, true);
	}

	DEFINE_COMMAND_EX(CspProvCityChurchNameChurchesAccessor, L"{ ? = CALL dbo.spProvCityChurchNameChurches(?,?,?) }")

	BEGIN_ACCESSOR_MAP(CspProvCityChurchNameChurchesAccessor, 1)
		BEGIN_ACCESSOR(0, true)
			COLUMN_ENTRY_LENGTH_STATUS(1, m_ChurchName, m_dwChurchNameLength, m_dwChurchNameStatus)
			COLUMN_ENTRY_LENGTH_STATUS(2, m_Address, m_dwAddressLength, m_dwAddressStatus)
			COLUMN_ENTRY_LENGTH_STATUS(3, m_Mail, m_dwMailLength, m_dwMailStatus)
			COLUMN_ENTRY_LENGTH_STATUS(4, m_PostalCode, m_dwPostalCodeLength, m_dwPostalCodeStatus)
			COLUMN_ENTRY_LENGTH_STATUS(5, m_ChurchID, m_dwChurchIDLength, m_dwChurchIDStatus)
		    COLUMN_ENTRY_LENGTH_STATUS(6, m_Modified, m_dwModifiedLength, m_dwModifiedStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspProvCityChurchNameChurchesAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_province)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_city)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(4, m_churchname)
	END_PARAM_MAP()
};//CspProvCityChurchNameChurchesAccessor
class CspProvCityChurchNameChurches : public CCommand<CAccessor<CspProvCityChurchNameChurchesAccessor> >
{
public:
#if (_MSC_VER == 1310)
	long GetChurchID()
#else
#if defined(_M_IA64) || defined (_M_AMD64)
	__int64 GetChurchID()
#else
	__int32 GetChurchID()
#endif
#endif
	{
#if (_MSC_VER == 1310)
		return m_ChurchID;
#else
#if defined(_M_IA64) || defined (_M_AMD64)
		__int64 iChurchID = 0;
		iChurchID = (__int64)m_ChurchID;
#else
		__int32 iChurchID = 0;
		iChurchID = (__int32)m_ChurchID;
#endif
		return iChurchID;
#endif
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spProvCityChurchNameChurches(?,?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspProvCityChurchNameChurches
/*************************************************************************************/
// Declaration of CspConfirmUser
class CspConfirmUserAccessor
{
public:
	LONG m_RETURN_VALUE;
	/*TCHAR m_Username[256];
	BYTE m_Userpwd[50];*/
	GUID m_Guid;
	LONG m_Id;
	/*DBSTATUS m_dwUsernameStatus;
	DBSTATUS m_dwUserpwdStatus;*/
	DBSTATUS m_dwGuidStatus;
	DBSTATUS m_dwIdStatus;
	/*DBLENGTH m_dwUsernameLength;
	DBLENGTH m_dwUserpwdLength;*/
	DBLENGTH m_dwGuidLength;
	DBLENGTH m_dwIdLength;

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_IRowsetScroll, true, DBPROPOPTIONS_REQUIRED);
		pPropSet->AddProperty(DBPROP_IRowsetChange, true, DBPROPOPTIONS_REQUIRED);
	}

	DEFINE_COMMAND_EX(CspConfirmUserAccessor, L"{ CALL dbo.spConfirmUser(?,?) }")

	//BEGIN_ACCESSOR_MAP(CspConfirmUserAccessor, 1)
	//	BEGIN_ACCESSOR(0, true)
	//		/*COLUMN_ENTRY_LENGTH_STATUS(1, m_Username, m_dwUsernameLength, m_dwUsernameStatus)
	//		COLUMN_ENTRY_LENGTH_STATUS(2, m_Userpwd, m_dwUserpwdLength, m_dwUserpwdStatus)*/
	//		COLUMN_ENTRY_LENGTH_STATUS(1, m_Guid, m_dwGuidLength, m_dwGuidStatus)
	//		COLUMN_ENTRY_LENGTH_STATUS(2, m_Id, m_dwIdLength, m_dwIdStatus)
	//		//COLUMN_ENTRY_LENGTH_STATUS(5, m_Id, m_dwIdLength, m_dwIdStatus)
	//	END_ACCESSOR()
	//END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspConfirmUserAccessor)
		/*SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)*/
		/*SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_Username)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_Userpwd)*/
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(1, m_Guid)
		SET_PARAM_TYPE(DBPARAMIO_INPUT | DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(2, m_Id)
	END_PARAM_MAP()
};//CspConfirmUserAccessor
class CspConfirmUser : public CCommand<CAccessor<CspConfirmUserAccessor>, CRowset, CMultipleResults >
{
public:
#if (_MSC_VER == 1310)
	long GetUserID()
#else
#if defined(_M_IA64) || defined (_M_AMD64)
	__int64 GetUserID()
#else
	__int32 GetUserID()
#endif
#endif
	{
#if (_MSC_VER == 1310)
		return m_Id;
#else
#if defined(_M_IA64) || defined (_M_AMD64)
		__int64 iId = 0;
		iId = (__int64)m_Id;
#else
		__int32 iId = 0;
		iId = (__int32)m_Id;
#endif
		return iId;
#endif
	}
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand=NULL)
	{
		DBPROPSET *pPropSet = NULL;
        CDBPropSet propset(DBPROPSET_ROWSET);
		__if_exists(HasBookmark)
        {
            if( HasBookmark() )
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet= &propset;
			}
        }	
        __if_exists(GetRowsetProperties)
        {
            GetRowsetProperties(&propset);
            pPropSet= &propset;
        }
        if (szCommand == NULL)
			szCommand = L"{ CALL dbo.spConfirmUser(?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);
		
		return hr;
	}
};//CspConfirmUser
  /*************************************************************************************/
  // Declaration of the CspGetIPCount
class CspGetIPCountAccessor
{
public:
	LONG m_Count = 0;
	DBSTATUS m_dwCountStatus;
	LONG m_RETURN_VALUE;
	ULARGE_INTEGER m_connectionid;

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspGetIPCountAccessor, L"{ ? = CALL dbo.spGetIPCount(?) }")

	BEGIN_ACCESSOR_MAP(CspGetIPCountAccessor, 1)
		BEGIN_ACCESSOR(0, true)
		COLUMN_ENTRY_STATUS(1, m_Count, m_dwCountStatus)
		END_ACCESSOR()
	END_ACCESSOR_MAP()

	BEGIN_PARAM_MAP(CspGetIPCountAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_connectionid)
	END_PARAM_MAP()
};//CspGetIPCountAccessor
class CspGetIPCount : public CCommand<CAccessor<CspGetIPCountAccessor> >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand = NULL)
	{
		DBPROPSET *pPropSet = NULL;
		CDBPropSet propset(DBPROPSET_ROWSET);
		__if_exists(HasBookmark)
		{
			if (HasBookmark())
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet = &propset;
			}
		}
		__if_exists(GetRowsetProperties)
		{
			GetRowsetProperties(&propset);
			pPropSet = &propset;
		}
		if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spGetIPCount(?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);

		return hr;
	}
};//CspGetIPCount
/*************************************************************************************/
// Declaration of the CspUpdateIP
class CspUpdateIPAccessor
{
public:
	LONG m_RETURN_VALUE;
	LONG m_newcount;
	ULARGE_INTEGER m_connectionid;

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspUpdateIPAccessor, L"{ ? = CALL dbo.spUpdateIP(?,?) }")

	BEGIN_PARAM_MAP(CspUpdateIPAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_connectionid)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_newcount)
	END_PARAM_MAP()
};//CspUpdateIPAccessor
class CspUpdateIP : public CCommand<CAccessor<CspUpdateIPAccessor>, CNoRowset >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand = NULL)
	{
		DBPROPSET *pPropSet = NULL;
		CDBPropSet propset(DBPROPSET_ROWSET);
		__if_exists(HasBookmark)
		{
			if (HasBookmark())
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet = &propset;
			}
		}
		__if_exists(GetRowsetProperties)
		{
			GetRowsetProperties(&propset);
			pPropSet = &propset;
		}
		if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spUpdateIP(?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);

		return hr;
	}
};//CspUpdateIP
/*************************************************************************************/
// Declaration of the CspAddIP
class CspAddIPAccessor
{
public:
	LONG m_RETURN_VALUE;
	LONG m_newcount;
	ULARGE_INTEGER m_connectionid;

	void GetRowsetProperties(CDBPropSet* pPropSet)
	{
		pPropSet->AddProperty(DBPROP_CANFETCHBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
		pPropSet->AddProperty(DBPROP_CANSCROLLBACKWARDS, true, DBPROPOPTIONS_OPTIONAL);
	}

	DEFINE_COMMAND_EX(CspAddIPAccessor, L"{ ? = CALL dbo.spAddIP(?,?) }")

	BEGIN_PARAM_MAP(CspAddIPAccessor)
		SET_PARAM_TYPE(DBPARAMIO_OUTPUT)
		COLUMN_ENTRY(1, m_RETURN_VALUE)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(2, m_connectionid)
		SET_PARAM_TYPE(DBPARAMIO_INPUT)
		COLUMN_ENTRY(3, m_newcount)
	END_PARAM_MAP()
};//CspAddIPAccessor
class CspAddIP : public CCommand<CAccessor<CspAddIPAccessor>, CNoRowset >
{
public:
	HRESULT OpenRowset(const CSession& session, LPCWSTR szCommand = NULL)
	{
		DBPROPSET *pPropSet = NULL;
		CDBPropSet propset(DBPROPSET_ROWSET);
		__if_exists(HasBookmark)
		{
			if (HasBookmark())
			{
				propset.AddProperty(DBPROP_IRowsetLocate, true);
				pPropSet = &propset;
			}
		}
		__if_exists(GetRowsetProperties)
		{
			GetRowsetProperties(&propset);
			pPropSet = &propset;
		}
		if (szCommand == NULL)
			szCommand = L"{ ? = CALL dbo.spAddIP(?,?) }";
		HRESULT hr = Open(session, szCommand, pPropSet);

		return hr;
	}
};//CspAddIP