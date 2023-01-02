Attribute VB_Name = "Wininet"
Option Explicit

Public bRSend As Boolean
Public bError As Boolean

Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3

Public Const INTERNET_DEFAULT_HTTP_PORT = 80
Public Const INTERNET_SERVICE_HTTP = 3
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const ERROR_INSUFFICIENT_BUFFER = 122

'
' additional cache flags
'

Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000      ' don't write this item to the cache
Public Const INTERNET_FLAG_DONT_CACHE = INTERNET_FLAG_NO_CACHE_WRITE
Public Const INTERNET_FLAG_MAKE_PERSISTENT = &H2000000     ' make this item persistent in cache
Public Const INTERNET_FLAG_FROM_CACHE = &H1000000          ' use offline semantics
Public Const INTERNET_FLAG_OFFLINE = INTERNET_FLAG_FROM_CACHE

'
' additional flags
'

Public Const INTERNET_FLAG_SECURE = &H800000               ' use PCT/SSL if applicable (HTTP)
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000      ' use keep-alive semantics
Public Const INTERNET_FLAG_NO_AUTO_REDIRECT = &H200000     ' don't handle redirections automatically
Public Const INTERNET_FLAG_READ_PREFETCH = &H100000        ' do background read prefetch
Public Const INTERNET_FLAG_NO_COOKIES = &H80000            ' no automatic cookie handling
Public Const INTERNET_FLAG_NO_AUTH = &H40000               ' no automatic authentication handling
Public Const INTERNET_FLAG_CACHE_IF_NET_FAIL = &H10000     ' return cache file if net request fails

'
' Security Ignore Flags, Allow HttpOpenRequest to overide
'  Secure Channel (SSL/PCT) failures of the following types.
'

Public Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTP = &H8000       ' ex: https:// to http://
Public Const INTERNET_FLAG_IGNORE_REDIRECT_TO_HTTPS = &H4000      ' ex: http:// to https://
Public Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000      ' expired X509 Cert.
Public Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000        ' bad common name in X509 Cert.

'
' more caching flags
'

Public Const INTERNET_FLAG_RESYNCHRONIZE = &H800           ' asking wininet to update an item if it is newer
Public Const INTERNET_FLAG_HYPERLINK = &H400               ' asking wininet to do hyperlinking semantic which works right for scripts
Public Const INTERNET_FLAG_NO_UI = &H200                   ' no cookie popup
Public Const INTERNET_FLAG_PRAGMA_NOCACHE = &H100          ' asking wininet to add "pragma: no-cache"
Public Const INTERNET_FLAG_CACHE_ASYNC = &H80              ' ok to perform lazy cache-write
Public Const INTERNET_FLAG_FORMS_SUBMIT = &H40             ' this is a forms submit
Public Const INTERNET_FLAG_NEED_FILE = &H10                ' need a file for this request
Public Const INTERNET_FLAG_MUST_CACHE_REQUEST = INTERNET_FLAG_NEED_FILE

'Public Const INTERNET_FLAG_SECURE = &H800000
'Public Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000
'Public Const INTERNET_FLAG_RELOAD = &H80000000

Public Const HTTP_QUERY_STATUS_CODE = 19
Public Const HTTP_QUERY_STATUS_TEXT = 20

'Public Const INTERNET_SERVICE_HTTP = 3
'Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0

Public Const HTTP_QUERY_CONTENT_TYPE = 1
Public Const HTTP_QUERY_CONTENT_LENGTH = 5
Public Const HTTP_QUERY_EXPIRES = 10
Public Const HTTP_QUERY_LAST_MODIFIED = 11
Public Const HTTP_QUERY_PRAGMA = 17
Public Const HTTP_QUERY_VERSION = 18
'Public Const HTTP_QUERY_STATUS_CODE = 19
'Public Const HTTP_QUERY_STATUS_TEXT = 20
Public Const HTTP_QUERY_RAW_HEADERS = 21
Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Public Const HTTP_QUERY_FORWARDED = 30
Public Const HTTP_QUERY_SERVER = 37
Public Const HTTP_QUERY_USER_AGENT = 39
Public Const HTTP_QUERY_SET_COOKIE = 43
Public Const HTTP_QUERY_REQUEST_METHOD = 45

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type


Public Type INTERNET_CACHE_ENTRY_INFO
  dwStructSize As Long
  lpszSourceUrlName As Long
  lpszLocalFileName As Long
  CacheEntryType As Long
  dwUseCount As Long
  dwHitRate As Long
  dwSizeLow As Long
  dwSizeHigh As Long
  LastModifiedTime As FILETIME
  ExpireTime As FILETIME
  LastAccessTime As FILETIME
  LastSyncTime As FILETIME
  lpHeaderInfo As Long
  dwHeaderInfoSize As Long
  lpszFileExtension As Long
  dwExemptDelta As Long
  szExtraMemory As String * 4016      ' MAX_CACHE_ENTRY_INFO_SIZE - SizeOf(INTERNET_CACHE_ENTRY_INFO)(4096 - 80)
End Type

Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function HttpSendRequest Lib "wininet.dll" _
    Alias "HttpSendRequestA" _
    (ByVal hHttpRequest As Long, _
    ByVal sHeaders As String, _
    ByVal lHeadersLength As Long, _
    ByVal sOptional As String, _
    ByVal lOptionalLength As Long) _
    As Integer
    
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" ( _
    ByVal hHttpRequest As Long, _
    ByVal lInfoLevel As Long, _
    ByRef sBuffer As Any, _
    ByRef lBufferLength As Long, _
    ByRef lIndex As Long) As Long

Public Declare Function InternetReadFile Lib "wininet.dll" ( _
    ByVal hFile As Long, _
    ByVal sBuffer As String, _
    ByVal lNumBytesToRead As Long, _
    lNumberOfBytesRead As Long) _
    As Integer

Public Type INTERNET_BUFFERS
    dwStructSize As Long ' used For API versioning. Set To sizeof(INTERNET_BUFFERS)
    Next As Long ' INTERNET_BUFFERS chain of buffers
    lpcszHeader As Long ' pointer To headers (may be NULL)
    dwHeadersLength As Long ' length of headers If Not NULL
    dwHeadersTotal As Long ' size of headers If Not enough buffer
    lpvBuffer As Long ' pointer To data buffer (may be NULL)
    dwBufferLength As Long ' length of data buffer If Not NULL
    dwBufferTotal As Long ' total size of chunk, or content-length If Not chunked
    dwOffsetLow As Long ' used For read-ranges (only used In HttpSendRequest2)
    dwOffsetHigh As Long
End Type

'Declare Function HttpSendRequestEx Lib "wininet.dll" Alias "HttpSendRequestExA" _
' (ByRef hRequest As Long, ByRef lpBuffersIn As INTERNET_BUFFERS, _
' ByRef lpBuffersOut As INTERNET_BUFFERS, ByVal dwFlags As Long, _
' ByVal dwContext As Long) _
' As Long

Public Declare Function InternetWriteFile Lib "wininet.dll" ( _
    ByVal hFile As Long, _
    ByVal sBuffer As String, _
    ByVal lNumberOfBytesToRead As Long, _
    lNumberOfBytesRead As Long) _
    As Integer

Public Declare Function HttpSendRequestEx Lib "wininet.dll" Alias "HttpSendRequestExA" _
    (ByVal hHttpRequest As Long, _
     lpBuffersIn As INTERNET_BUFFERS, _
    ByVal lpBuffersOut As Long, _
    ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Long

Public Declare Function HttpEndRequest Lib "wininet.dll" Alias "HttpEndRequestA" _
    (ByVal hHttpRequest As Long, _
    ByVal lpBuffersOut As Long, _
    ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Long


'Private Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumberOfBytesToRead As Long, lNumberOfBytesRead As Long) As Integer


Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
'Declare Function InternetTimeToSystemTime Lib "wininet.dll" (ByVal lpszTime As String, ByRef pst As SYSTEMTIME, ByVal dwReserved As Long) As Long
'Declare Function GetUrlCacheEntryInfo Lib "wininet.dll" Alias "GetUrlCacheEntryInfoA" (ByVal lpszUrlName As String, lpCacheEntryInfo As INTERNET_CACHE_ENTRY_INFO, lpdwCacheEntryInfoBufferSize As Long) As Long

Public Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias _
      "HttpAddRequestHeadersA" (ByVal hHttpRequest As Long, _
      ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal _
      lModifiers As Long) As Integer
Public Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000
Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000
Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000




