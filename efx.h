#define     FIELDSIZE_PROVINCE		51
#define     FIELDSIZE_REGION		51
#define     FIELDSIZE_CITY			51
#define     FIELDSIZE_DENOMINATION	101
#define     FIELDSIZE_CHURCHNAME	73
#define		FIELDSIZE_ADDRESS		73
#define		FIELDSIZE_MAIL			129
#define		FIELDSIZE_POSTAL_CODE	11
#define		FIELDSIZE_PHONE			51
#define		FIELDSIZE_FAX			51
#define		FIELDSIZE_CONTACT_PERSON	51
#define		FIELDSIZE_EMAIL			151
#define		FIELDSIZE_WEBSITE		151
#define		FIELDSIZE_DENOM_WEBSITE 151
#define		ALTERNATIVES	"<table align=\"center\" cellspacing=\"0\"><tr class=\"80row\"><td class=\"others\" width=\"80%\">Browse for other churches in...</td><td width=\"20%\"></td></tr><tr class=\"20row\"><td class=\"others\" colspan=\"2\"><a href=\"index.srf?Province="
#define		AHREF				"<a href=\"index.srf?Province="
#define		MOUSEOVER "onmouseover=\"this.style.background='tomato';\" onmouseout=\"this.style.background='transparent';\""
#define		LINKCOLOR "style=\"color:white\""
#define		IMGSIZE	  "function Init()\r\n{\r\n\tvar quarterWidth = Math.round(document.all(\"imgChurch\").width / 4);\r\n\tvar quarterHeight = Math.round(document.all(\"imgChurch\").height / 4);\r\n\tdocument.all(\"imgChurch\").style.pixelLeft = document.all(\"tbl\").offsetWidth + document.all(\"tbl\").offsetLeft + quarterWidth;\r\n\tdocument.all(\"imgChurch\").style.pixelTop = document.all(\"tbl\").offsetTop + quarterHeight;\r\n}"
#define		FORMSUBMIT	"function Init(){}"
#define		MAPQUEST							"<a href=\"http://www.mapquest.com/maps/map.adp?"
#define		MAPQUEST_END						"&country=CA&cid=lfmaplink\">MapQuest</a>"
#define		PROV_REG_CITY_DENOM_CHURCHNAME_CHURCHES_HEADER	"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Map</th><th>City</th></tr>"
#define		PROV_REG_CITY_CHURCHNAME_CHURCHES_HEADER		"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Mail</th><th>Map</th><th>City</th></tr>"
#define		PROV_REG_DENOM_CHURCHNAME_CHURCHES_HEADER		"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Mail</th><th>Map</th><th>City</th></th>"
#define		PROV_CITY_DENOM_CHURCHNAME_CHURCHES_HEADER		"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Mail</th><th>Map</th></tr>"
#define		PROV_REG_CITY_DENOM_CHURCHES_HEADER	"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Map</th><th>City</th></tr>"
#define		PROV_REG_CHURCHNAME_CHURCHES_HEADER	"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Map</th><th>City</th></tr>"
#define		PROV_REG_CITY_CHURCHES_HEADER		"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Map</th><th>Denomination</th><th>City</th></tr>"
#define		PROV_REG_DENOM_CHURCHES_HEADER		"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Map</th><th>City</th></tr>"
#define		PROV_REG_CHURCHES_HEADER			"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Map</th><th>City</th></tr>"
#define		PROV_DENOM_CHURCHES_HEADER			"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Mail</th><th>Map</th><th>City</th></tr>"
#define		PROV_CITY_DENOM_CHURCHES_HEADER		"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Mail</th><th>Map</th></tr>"
#define		PROV_CITY_CHURCHES_HEADER			"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Mail</th><th>Map</th><th>Denomination</th></tr>"
#define		PROV_CITY_CHURCHNAME_CHURCHES_HEADER	"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Mail</th><th>Map</th></tr>"
#define		PROV_DENOM_CHURCHNAME_CHURCHES_HEADER	"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Mail</th><th>Map</th><th>City</th></tr>"
#define		PROV_CHURCHES_HEADER				"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Mail</th><th>Map</th><th>City</th><th>Region</th></tr>"
#define		PROV_CHURCHNAME_CHURCHES_HEADER		"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Mail</th><th>Map</th><th>City</th></tr>"
#define		PROV_CHURCHNAME_HEADER				"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Map</th><th>City</th><th>Denomination</th></tr>"
#define		REG_CITY_DENOM_CHURCHES_HEADER		"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Map</th><th>City</th><th>Denomination</th></tr>"
#define		REG_CITY_CHURCHES_HEADER			"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Map</th><th>Denomination</th></tr>"
#define		REG_DENOM_CHURCHES_HEADER			"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Map</th><th>City</th></tr>"
#define		REG_CHURCHES_HEADER					"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Mail</th><th>Map</th><th>City</th></tr>"
#define		CITY_DENOM_CHURCHES_HEADER			"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Mail</th><th>Map</th><th>City</th><th>Denomination</th></tr>"
#define		CITY_CHURCHES_HEADER				"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Map</th><th>Denomination</th></tr>"
#define		DENOM_CHURCHES_HEADER				"<tr class=\"clsStatus\"><th>Church Name</th><th>Address</th><th>Mail</th><th>Map</th><th>City</th><th>Province</th></tr>"
#define		FTSEARCH	"<table class=\"ftTable\"><tr><th class=\"selHeader\"><label>Church Name:</label></th></tr><tr><td><input class=\"selBox\" type=\"text\" name=\"ftName\" size=\"18\"></td></tr><tr><td><input class=\"selButton\" type=\"submit\" value=\"Search\"></td></tr></table>"
#define		OPTION_10	"<option class=\"clsOption\" selected>10</option><option class=\"clsOption\">20</option><option class=\"clsOption\">25</option><option class=\"clsOption\">30</option><option class=\"clsOption\">35</option><option class=\"clsOption\">40</option><option class=\"clsOption\">50</option><option class=\"clsOption\">100</option>"
#define		OPTION_20	"<option class=\"clsOption\">10</option><option class=\"clsOption\" selected>20</option><option class=\"clsOption\">25</option><option class=\"clsOption\">30</option><option class=\"clsOption\">35</option><option class=\"clsOption\">40</option><option class=\"clsOption\">50</option><option class=\"clsOption\">100</option>"
#define		OPTION_25	"<option class=\"clsOption\">10</option><option class=\"clsOption\">20</option><option class=\"clsOption\" selected>25</option><option class=\"clsOption\">30</option><option class=\"clsOption\">35</option><option class=\"clsOption\">40</option><option class=\"clsOption\">50</option><option class=\"clsOption\">100</option>"
#define		OPTION_30	"<option class=\"clsOption\">10</option><option class=\"clsOption\">20</option><option class=\"clsOption\">25</option><option class=\"clsOption\" selected>30</option><option class=\"clsOption\">35</option><option class=\"clsOption\">40</option><option class=\"clsOption\">50</option><option class=\"clsOption\">100</option>"
#define		OPTION_35	"<option class=\"clsOption\">10</option><option class=\"clsOption\">20</option><option class=\"clsOption\">25</option><option class=\"clsOption\">30</option><option selected class=\"clsOption\">35</option><option class=\"clsOption\">40</option><option class=\"clsOption\">50</option><option class=\"clsOption\">100</option>"
#define		OPTION_40	"<option class=\"clsOption\">10</option><option class=\"clsOption\">20</option><option class=\"clsOption\">25</option><option class=\"clsOption\">30</option><option class=\"clsOption\">35</option><option class=\"clsOption\" selected>40</option><option class=\"clsOption\">50</option><option class=\"clsOption\">100</option>"
#define		OPTION_50	"<option class=\"clsOption\">10</option><option class=\"clsOption\">20</option><option class=\"clsOption\">25</option><option class=\"clsOption\">30</option><option class=\"clsOption\">35</option><option class=\"clsOption\">40</option><option class=\"clsOption\" selected>50</option><option class=\"clsOption\">100</option>"
#define		OPTION_100	"<option class=\"clsOption\">10</option><option class=\"clsOption\">20</option><option class=\"clsOption\">25</option><option class=\"clsOption\">30</option><option class=\"clsOption\">35</option><option class=\"clsOption\">40</option><option class=\"clsOption\">50</option><option class=\"clsOption\" selected>100</option>"
#define		SELECT_PROVINCE			"function SelectProvince()\r\n{\r\n\tdocument.forms[\"sel\"].elements[\"ChurchName\"].value = \"\";\r\n\tdocument.forms[\"sel\"].submit();\r\n}"
#define		SELECT_NAME_CLR_REGION	"function SelectName()\r\n{\r\n\tif (document.all[\"txtBox\"].value != \"\")\r\n\t{\r\n\t\tdocument.forms[\"sel\"].elements[\"Region\"].value = \"\";\r\n\t\tdocument.forms[\"sel\"].submit();\r\n\t}\r\n}\r\n"
#define		SELECT_PP_CLR_REGION	"\r\nfunction SelectPp()\r\n{\r\n\tdocument.forms[\"sel\"].elements[\"Region\"].value = \"\";\r\n\tdocument.forms[\"sel\"].elements[\"ChurchName\"].value = \"\";\r\n\tdocument.forms[\"sel\"].submit();\r\n}"
#define		SELECT_REGION			"function SelectRegion()\r\n{\r\n\tdocument.forms[\"sel\"].elements[\"ChurchName\"].value = \"\";\r\n\tdocument.forms[\"sel\"].submit();\r\n}"
#define		SELECT_PROVINCE_CLR_REGION "function SelectProvince()\r\n{\r\n\tdocument.forms[\"sel\"].elements[\"Region\"].value = \"\";\r\n\tdocument.forms[\"sel\"].elements[\"ChurchName\"].value = \"\";\r\n\tdocument.forms[\"sel\"].submit();\r\n}\r\n"
#define		SELECT_NAME_CLR_CITY	"function SelectName()\r\n{\r\n\tif (document.all[\"txtBox\"].value != \"\")\r\n\t{\r\n\t\tdocument.forms[\"sel\"].elements[\"City\"].value = \"\";\r\n\t\tdocument.forms[\"sel\"].submit();\r\n\t}\r\n}\r\n"
#define		SELECT_PP_CLR_CITY		"\r\nfunction SelectPp()\r\n{\r\n\tdocument.forms[\"sel\"].elements[\"City\"].value = \"\";\r\n\tdocument.forms[\"sel\"].elements[\"ChurchName\"].value = \"\";\r\n\tdocument.forms[\"sel\"].submit();\r\n}"
#define		SELECT_CITY				"function SelectCity()\r\n{\r\n\tdocument.forms[\"sel\"].elements[\"ChurchName\"].value = \"\";\r\n\tdocument.forms[\"sel\"].submit();\r\n}"
#define		SELECT_REGION_CLR_CITY	"function SelectRegion()\r\n{\r\n\tdocument.forms[\"sel\"].elements[\"City\"].value = \"\";\r\n\tdocument.forms[\"sel\"].elements[\"ChurchName\"].value = \"\";\r\n\tdocument.forms[\"sel\"].submit();\r\n}\r\n"
#define		SELECT_PROVINCE_CLR_CITY_REGION	"function SelectProvince()\r\n{\r\n\tdocument.forms[\"sel\"].elements[\"City\"].value = \"\";\r\n\tdocument.forms[\"sel\"].elements[\"Region\"].value = \"\";\r\n\tdocument.forms[\"sel\"].elements[\"ChurchName\"].value = \"\";\r\n\tdocument.forms[\"sel\"].submit();\r\n}\r\n"
#define		SELECT_NAME_CLR_DENOM	"function SelectName()\r\n{\r\n\tif (document.all[\"txtBox\"].value != \"\")\r\n\t{\r\n\t\tdocument.forms[\"sel\"].elements[\"Denomination\"].value = \"\";\r\n\t\tdocument.forms[\"sel\"].submit();\r\n\t}\r\n}\r\n"
#define		SELECT_PP_CLR_DENOM		"\r\nfunction SelectPp()\r\n{\r\n\tdocument.forms[\"sel\"].elements[\"Denomination\"].value = \"\";\r\n\tdocument.forms[\"sel\"].elements[\"ChurchName\"].value = \"\";\r\n\tdocument.forms[\"sel\"].submit();\r\n}"
#define		SELECT_DENOM			"function SelectDenomination()\r\n{\r\n\tdocument.forms[\"sel\"].elements[\"ChurchName\"].value = \"\";\r\n\tdocument.forms[\"sel\"].submit();\r\n}"
#define		SELECT_CITY_CLR_DENOM	"function SelectCity()\r\n{\r\n\tdocument.forms[\"sel\"].elements[\"Denomination\"].value = \"\";\r\n\tdocument.forms[\"sel\"].elements[\"ChurchName\"].value = \"\";\r\n\tdocument.forms[\"sel\"].submit();\r\n}\r\n"
#define		SELECT_REGION_CLR_CITY_DENOM	"function SelectRegion()\r\n{\r\n\tdocument.forms[\"sel\"].elements[\"City\"].value = \"\";\r\n\tdocument.forms[\"sel\"].elements[\"Denomination\"].value = \"\";\r\n\tdocument.forms[\"sel\"].elements[\"ChurchName\"].value = \"\";\r\n\tdocument.forms[\"sel\"].submit();\r\n}\r\n"
#define		SELECT_PROVINCE_CLR_ALL	"function SelectProvince()\r\n{\r\n\tdocument.forms[\"sel\"].elements[\"Denomination\"].value = \"\";\r\n\tdocument.forms[\"sel\"].elements[\"City\"].value = \"\";\r\n\tdocument.forms[\"sel\"].elements[\"Region\"].value = \"\";\r\n\tdocument.forms[\"sel\"].elements[\"ChurchName\"].value = \"\";\r\n\tdocument.forms[\"sel\"].submit();\r\n}\r\n"
#define		TITLE_PROV_REG_CITY_DENOM_NAME	_T("%s is a %s church located in %s, %s, %s")