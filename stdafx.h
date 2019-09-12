// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once
#if defined (_M_ARM)
#define _ARM_
#define _ARM_WINAPI_PARTITION_DESKTOP_SDK_AVAILABLE 1
#elif defined(_M_AMD64)
#define _AMD64_
#elif defined(_M_IX86)
#define _X86_
#endif
#include "targetver.h"
#define _CRT_SECURE_NO_WARNINGS
#define _ATL_USE_WINAPI_FAMILY_PHONE_APP
#define _APISET_MINWIN_VERSION 0x0101
#define _ATL_NO_COMMODULE
#define _ATL_NO_COM_SUPPORT
#define _ATL_NO_WIN_SUPPORT
#define ATL_NO_MLANG
#define ATL_NO_SOAP
#define ATL_NO_MMSYS
#define ATL_CRITICAL_ISAPI_ERROR_LOGONLY
#define ATL_NO_ACLAPI
#define _ATL_PERF_NOXML
#define HSE_IO_DECLARE_PUSH    0x00000080
#include <stdio.h>
#include <tchar.h>
#include <Pathcch.h>
#include <atlbase.h>
#include <atlutil.h>
#include <atlsiface.h>
#include <atlisapi.h>
typedef enum {
	eSTATUS_ERROR,
	eSTATUS_NOT_STARTED,
	eSTATUS_STARTED,
	eSTATUS_DONE
}eJobStatus;
typedef enum {
	eSTATUS_CLOSED,
	eSTATUS_OPEN,
	eSTATUS_EXPIRED
}eConnectionStatus;
enum CONNECTION_TYPE { connectionKeepAlive = 0, connectionStatic };
enum CONNECTION_STATUS { connectionOk = 0, connectionExpired, connectionClose, connectionError };
struct CConnectionData
{
	CONNECTION_TYPE connType;
	CONNECTION_STATUS connStatus;
	char m_szSessionID[128];
	FILETIME m_dwSessionExpiry;
};
typedef enum {
	eFileTypeConfig,
	eFileTypeRemap,
	eFileTypeAuth,
	eFileTypeMime,
	eFileTypeBrowscap
} eConfigFile;
typedef enum
{
	eCachePublic = 0,//forces public cacheing, even if https is used
	eCachePrivate = 1,//private cacheing only, no proxy cacheing
	eCacheNoCache = 2,//forces caches to resubmit request to origin server for validation (ie. if-not-modified)
	eCacheNoStore = 4,//no cacheing allowed
	eCacheMustRevalidate = 8,//overrides any cdn/proxy server caches
	eCacheProxyRevalidate = 16,//overrides proxy caches only
	eCacheImmutable = 64//disable client refresh
} eHttpCacheType;
struct CHttpCacheControlData
{
	DWORD cacheFlags;//one or more of eHttpCacheType
	int maxage;//maximum freshness time in seconds
	int smaxage;//max freshness, but only applies to proxies
	int maxstale;
	int retain;
	__int64 expiry;//expiry header
};
#include <atlcacheex.h>