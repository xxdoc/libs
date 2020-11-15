Attribute VB_Name = "modDeclares"
Option Explicit
'this is to (start to) eliminate the need for the vblibcurl.tlb
'typelibs put imports right into our import table, so dependancy dlls
'must be found on startup by windows loader, we can't hunt for them
'during dev this can be annoying unless you put everything in system path
'(and remember to update them if they change in dev)

Public Enum curlioerr
    CURLIOE_OK = 0
    CURLIOE_UNKNOWNCMD = 1
    CURLIOE_FAILRESTART = 2
    CURLIOE_LAST = 3
End Enum

Public Enum curliocmd
    CURLIOCMD_NOP = 0
    CURLIOCMD_RESTARTREAD = 1
    CURLIOCMD_LAST = 2
End Enum

Public Enum curl_infotype
    CURLINFO_TEXT = 0
    CURLINFO_HEADER_IN = 1
    CURLINFO_HEADER_OUT = 2
    CURLINFO_DATA_IN = 3
    CURLINFO_DATA_OUT = 4
    CURLINFO_SSL_DATA_IN = 5
    CURLINFO_SSL_DATA_OUT = 6
    CURLINFO_END = 7
End Enum
 
Public Enum CURLcode
    CURLE_OK = 0
    CURLE_ABORTED_BY_CALLBACK = 42
    CURLE_BAD_CALLING_ORDER = 44
    CURLE_BAD_CONTENT_ENCODING = 61
    CURLE_BAD_DOWNLOAD_RESUME = 36
    CURLE_BAD_FUNCTION_ARGUMENT = 43
    CURLE_BAD_PASSWORD_ENTERED = 46
    CURLE_COULDNT_CONNECT = 7
    CURLE_COULDNT_RESOLVE_HOST = 6
    CURLE_COULDNT_RESOLVE_PROXY = 5
    CURLE_FAILED_INIT = 2
    CURLE_FILESIZE_EXCEEDED = 63
    CURLE_FILE_COULDNT_READ_FILE = 37
    CURLE_FTP_ACCESS_DENIED = 9
    CURLE_FTP_CANT_GET_HOST = 15
    CURLE_FTP_CANT_RECONNECT = 16
    CURLE_FTP_COULDNT_GET_SIZE = 32
    CURLE_FTP_COULDNT_RETR_FILE = 19
    CURLE_FTP_COULDNT_SET_ASCII = 29
    CURLE_FTP_COULDNT_SET_BINARY = 17
    CURLE_FTP_COULDNT_STOR_FILE = 25
    CURLE_FTP_COULDNT_USE_REST = 31
    CURLE_FTP_PORT_FAILED = 30
    CURLE_FTP_QUOTE_ERROR = 21
    CURLE_FTP_SSL_FAILED = 64
    CURLE_FTP_USER_PASSWORD_INCORRECT = 10
    CURLE_FTP_WEIRD_227_FORMAT = 14
    CURLE_FTP_WEIRD_PASS_REPLY = 11
    CURLE_FTP_WEIRD_PASV_REPLY = 13
    CURLE_FTP_WEIRD_SERVER_REPLY = 8
    CURLE_FTP_WEIRD_USER_REPLY = 12
    CURLE_FTP_WRITE_ERROR = 20
    CURLE_FUNCTION_NOT_FOUND = 41
    CURLE_GOT_NOTHING = 52
    CURLE_HTTP_POST_ERROR = 34
    CURLE_HTTP_RANGE_ERROR = 33
    CURLE_HTTP_RETURNED_ERROR = 22
    CURLE_INTERFACE_FAILED = 45
    CURLE_LAST = 67
    CURLE_LDAP_CANNOT_BIND = 38
    CURLE_LDAP_INVALID_URL = 62
    CURLE_LDAP_SEARCH_FAILED = 39
    CURLE_LIBRARY_NOT_FOUND = 40
    CURLE_MALFORMAT_USER = 24
    CURLE_OBSOLETE = 50
    CURLE_OPERATION_TIMEOUTED = 28
    CURLE_OUT_OF_MEMORY = 27
    CURLE_PARTIAL_FILE = 18
    CURLE_READ_ERROR = 26
    CURLE_RECV_ERROR = 56
    CURLE_SEND_ERROR = 55
    CURL_SEND_FAIL_REWIND = 65
    CURLE_SHARE_IN_USE = 57
    CURLE_SSL_CACERT = 60
    CURLE_SSL_CERTPROBLEM = 58
    CURLE_SSL_CIPHER = 59
    CURLE_SSL_CONNECT_ERROR = 35
    CURLE_SSL_ENGINE_INITFAILED = 66
    CURLE_SSL_ENGINE_NOTFOUND = 53
    CURLE_SSL_ENGINE_SETFAILED = 54
    CURLE_SSL_PEER_CERTIFICATE = 51
    CURLE_TELNET_OPTION_SYNTAX = 49
    CURLE_TOO_MANY_REDIRECTS = 47
    CURLE_UNKNOWN_TELNET_OPTION = 48
    CURLE_UNSUPPORTED_PROTOCOL = 1
    CURLE_URL_MALFORMAT = 3
    CURLE_URL_MALFORMAT_USER = 4
    CURLE_WRITE_ERROR = 23
End Enum
 
Public Enum curl_proxytype
    CURLPROXY_HTTP = 0
    CURLPROXY_SOCKS4 = 4
    CURLPROXY_SOCKS5 = 5
End Enum
    
Public Enum curl_httpauth
    CURLAUTH_NONE = 0
    CURLAUTH_BASIC = 1
    CURLAUTH_DIGEST = 2
    CURLAUTH_GSSNEGOTIATE = 4
    CURLAUTH_NTLM = 8
    CURLAUTH_ANY = 15
    CURLAUTH_ANYSAFE = 14
 End Enum
  
Public Enum curl_ftpssl
    CURLFTPSSL_NONE = 0
    CURLFTPSSL_TRY = 1
    CURLFTPSSL_CONTROL = 2
    CURLFTPSSL_ALL = 3
    CURLFTPSSL_LAST = 4
End Enum
   
Public Enum curl_ftpauth
    CURLFTPAUTH_DEFAULT = 0
    CURLFTPAUTH_SSL = 1
    CURLFTPAUTH_TLS = 2
    CURLFTPAUTH_LAST = 3
End Enum
    

Public Enum CURLoption               'https://curl.se/libcurl/c/curl_easy_setopt.html
    CURLOPT_AUTOREFERER = 58
    CURLOPT_BUFFERSIZE = 98
    CURLOPT_CAINFO = 10065
    CURLOPT_CAPATH = 10097
    CURLOPT_CLOSEPOLICY = 72
    CURLOPT_CONNECTTIMEOUT = 78
    CURLOPT_COOKIE = 10022
    CURLOPT_COOKIEFILE = 10031
    CURLOPT_COOKIEJAR = 10082
    CURLOPT_COOKIESESSION = 96
    CURLOPT_CRLF = 27
    CURLOPT_CUSTOMREQUEST = 10036
    CURLOPT_DEBUGDATA = 10095
    CURLOPT_DEBUGFUNCTION = 20094
    CURLOPT_DNS_CACHE_TIMEOUT = 92
    CURLOPT_DNS_USE_GLOBAL_CACHE = 91
    CURLOPT_EDGSOCKET = 10077
    CURLOPT_ENCODING = 10102
    CURLOPT_ERRORBUFFER = 10010
    CURLOPT_FAILONERROR = 45
    CURLOPT_FILETIME = 69
    CURLOPT_FOLLOWLOCATION = 52
    CURLOPT_FORBID_REUSE = 75
    CURLOPT_FRESH_CONNECT = 74
    CURLOPT_FTPACCOUNT = 10134
    CURLOPT_FTPAPPEND = 50
    CURLOPT_FTPLISTONLY = 48
    CURLOPT_FTPPORT = 10017
    CURLOPT_FTPSSLAUTH = 129
    CURLOPT_FTP_CREATE_MISSING_DIRS = 110
    CURLOPT_FTP_RESPONSE_TIMEOUT = 112
    CURLOPT_FTP_SSL = 119
    CURLOPT_FTP_USE_EPRT = 106
    CURLOPT_FTP_USE_EPSV = 85
    CURLOPT_HEADER = 42
    CURLOPT_HEADERDATA = 10029
    CURLOPT_HEADERFUNCTION = 20079
    CURLOPT_HTTP200ALIASES = 10104
    CURLOPT_HTTPAUTH = 107
    CURLOPT_HTTPGET = 80
    CURLOPT_HTTPHEADER = 10023
    CURLOPT_HTTPPOST = 10024
    CURLOPT_HTTPPROXYTUNNEL = 61
    CURLOPT_HTTP_VERSION = 84
    CURLOPT_IOCTLFUNCTION = 20130
    CURLOPT_IOCTLDATA = 10131
    CURLOPT_INFILESIZE = 14
    CURLOPT_INFILESIZE_LARGE = 30115
    CURLOPT_INTERFACE = 10062
    CURLOPT_IPRESOLVE = 113
    CURLOPT_KRB4LEVEL = 10063
    CURLOPT_LASTENTRY = 135
    CURLOPT_LOW_SPEED_LIMIT = 19
    CURLOPT_LOW_SPEED_TIME = 20
    CURLOPT_MAXCONNECTS = 71
    CURLOPT_MAXFILESIZE = 114            'good up to 2gb
    CURLOPT_MAXFILESIZE_LARGE = 30117    'over 2gb pass in curl_off_t struct https://curl.se/libcurl/c/CURLOPT_MAXFILESIZE_LARGE.html
    CURLOPT_MAXREDIRS = 68
    CURLOPT_NETRC = 51
    CURLOPT_NETRC_FILE = 10118
    CURLOPT_NOBODY = 44
    CURLOPT_NOPROGRESS = 43
    CURLOPT_NOSIGNAL = 99
    CURLOPT_PASV_HOST = 126
    CURLOPT_PORT = 3
    CURLOPT_POST = 47
    CURLOPT_POSTFIELDS = 10015
    CURLOPT_POSTFIELDSIZE = 60
    CURLOPT_POSTFIELDSIZE_LARGE = 30120
    CURLOPT_POSTQUOTE = 10039
    CURLOPT_PREQUOTE = 10093
    CURLOPT_PRIVATE = 10103
    CURLOPT_PROGRESSDATA = 10057
    CURLOPT_PROGRESSFUNCTION = 20056
    CURLOPT_PROXY = 10004
    CURLOPT_PROXYAUTH = 111
    CURLOPT_PROXYPORT = 59
    CURLOPT_PROXYTYPE = 101
    CURLOPT_PROXYUSERPWD = 10006
    CURLOPT_PUT = 54
    CURLOPT_QUOTE = 10028
    CURLOPT_RANDOM_FILE = 10076
    CURLOPT_RANGE = 10007
    CURLOPT_READDATA = 10009
    CURLOPT_READFUNCTION = 20012
    CURLOPT_REFERER = 10016
    CURLOPT_RESUME_FROM = 21
    CURLOPT_RESUME_FROM_LARGE = 30116
    CURLOPT_SHARE = 10100
    CURLOPT_SOURCE_HOST = 10122
    CURLOPT_SOURCE_PATH = 10124
    CURLOPT_SOURCE_PORT = 125
    CURLOPT_SOURCE_POSTQUOTE = 10128
    CURLOPT_SOURCE_PREQUOTE = 10127
    CURLOPT_SOURCE_QUOTE = 10133
    CURLOPT_SOURCE_URL = 10132
    CURLOPT_SOURCE_USERPWD = 10123
    CURLOPT_SSLCERT = 10025
    CURLOPT_SSLCERTPASSWD = 10026
    CURLOPT_SSLCERTTYPE = 10086
    CURLOPT_SSLENGINE = 10089
    CURLOPT_SSLENGINE_DEFAULT = 90
    CURLOPT_SSLKEY = 10087
    CURLOPT_SSLKEYPASSWD = 10026
    CURLOPT_SSLKEYTYPE = 10088
    CURLOPT_SSLVERSION = 32
    CURLOPT_SSL_CIPHER_LIST = 10083
    CURLOPT_SSL_CTX_DATA = 10109
    CURLOPT_SSL_CTX_FUNCTION = 20108
    CURLOPT_SSL_VERIFYHOST = 81
    CURLOPT_SSL_VERIFYPEER = 64
    CURLOPT_STDERR = 10037
    CURLOPT_TCP_NODELAY = 121
    CURLOPT_TELNETOPTIONS = 10070
    CURLOPT_TIMECONDITION = 33
    CURLOPT_TIMEOUT = 13
    CURLOPT_TIMEVALUE = 34
    CURLOPT_TRANSFERTEXT = 53
    CURLOPT_UNRESTRICTED_AUTH = 105
    CURLOPT_UPLOAD = 46
    CURLOPT_URL = 10002
    CURLOPT_USERAGENT = 10018
    CURLOPT_USERPWD = 10005
    CURLOPT_VERBOSE = 41
    CURLOPT_WRITEDATA = 10001
    CURLOPT_WRITEFUNCTION = 20011
    CURLOPT_WRITEINFO = 10040
End Enum
    
Public Enum CURL_IPRESOLVE
    CURL_IPRESOLVE_WHATEVER = 0
    CURL_IPRESOLVE_V4 = 1
    CURL_IPRESOLVE_V6 = 2
 End Enum
   
Public Enum CURL_HTTP_VERSION
    CURL_HTTP_VERSION_NONE = 0
    CURL_HTTP_VERSION_1_0 = 1
    CURL_HTTP_VERSION_1_1 = 2
    CURL_HTTP_VERSION_LAST = 3
End Enum
    
Public Enum CURL_NETRC_OPTION
    CURL_NETRC_IGNORED = 0
    CURL_NETRC_OPTIONAL = 1
    CURL_NETRC_REQUIRED = 2
    CURL_NETRC_LAST = 3
End Enum
   
Public Enum CURL_SSLVERSION
    CURL_SSLVERSION_DEFAULT = 0
    CURL_SSLVERSION_TLSv1 = 1
    CURL_SSLVERSION_SSLv2 = 2
    CURL_SSLVERSION_SSLv3 = 3
    CURL_SSLVERSION_LAST = 4
End Enum
    
Public Enum curl_TimeCond
    CURL_TIMECOND_NONE = 0
    CURL_TIMECOND_IFMODSINCE = 1
    CURL_TIMECOND_IFUNMODSINCE = 2
    CURL_TIMECOND_LASTMOD = 3
    CURL_TIMECOND_LAST = 4
End Enum
    
Public Enum CURLFORMcode
    CURL_FORMADD_OK = 0
    CURL_FORMADD_MEMORY = 1
    CURL_FORMADD_OPTION_TWICE = 2
    CURL_FORMADD_NULL = 3
    CURL_FORMADD_UNKNOWN_OPTION = 4
    CURL_FORMADD_INCOMPLETE = 5
    CURL_FORMADD_ILLEGAL_ARRAY = 6
    CURL_FORMADD_DISABLED = 7
    CURL_FORMADD_LAST = 8
End Enum
   
Public Enum CURLformoption
    CURLFORM_ARRAY = 8
    CURLFORM_BUFFER = 11
    CURLFORM_BUFFERLENGTH = 13
    CURLFORM_BUFFERPTR = 12
    CURLFORM_CONTENTHEADER = 15
    CURLFORM_CONTENTSLENGTH = 6
    CURLFORM_CONTENTTYPE = 14
    CURLFORM_COPYCONTENTS = 4
    CURLFORM_COPYNAME = 1
    CURLFORM_END = 17
    CURLFORM_FILE = 10
    CURLFORM_FILECONTENT = 7
    CURLFORM_FILENAME = 16
    CURLFORM_NAMELENGTH = 3
    CURLFORM_NOTHING = 0
    CURLFORM_OBSOLETE = 9
    CURLFORM_OBSOLETE2 = 18
    CURLFORM_PTRCONTENTS = 5
    CURLFORM_PTRNAME = 2
End Enum
   
Public Enum CURLINFO
    CURLINFO_CONNECT_TIME = 3145733
    CURLINFO_CONTENT_LENGTH_DOWNLOAD = 3145743
    CURLINFO_CONTENT_LENGTH_UPLOAD = 3145744
    CURLINFO_CONTENT_TYPE = 1048594
    CURLINFO_EFFECTIVE_URL = 1048577
    CURLINFO_FILETIME = 2097166
    CURLINFO_HEADER_SIZE = 2097163
    CURLINFO_HTTPAUTH_AVAIL = 2097175
    CURLINFO_HTTP_CONNECTCODE = 2097174
    CURLINFO_LASTONE = 28
    CURLINFO_NAMELOOKUP_TIME = 3145732
    CURLINFO_NONE = 0
    CURLINFO_NUM_CONNECTS = 2097178
    CURLINFO_OS_ERRNO = 2097177
    CURLINFO_PRETRANSFER_TIME = 3145734
    CURLINFO_PRIVATE = 1048597
    CURLINFO_PROXYAUTH_AVAIL = 2097176
    CURLINFO_REDIRECT_COUNT = 2097172
    CURLINFO_REDIRECT_TIME = 3145747
    CURLINFO_REQUEST_SIZE = 2097164
    CURLINFO_RESPONSE_CODE = 2097154
    CURLINFO_SIZE_DOWNLOAD = 3145736
    CURLINFO_SIZE_UPLOAD = 3145735
    CURLINFO_SPEED_DOWNLOAD = 3145737
    CURLINFO_SPEED_UPLOAD = 3145738
    CURLINFO_SSL_ENGINES = 4194331
    CURLINFO_SSL_VERIFYRESULT = 2097165
    CURLINFO_STARTTRANSFER_TIME = 3145745
    CURLINFO_TOTAL_TIME = 3145731
End Enum
   
Public Enum curl_closepolicy
    CURLCLOSEPOLICY_NONE = 0
    CURLCLOSEPOLICY_OLDEST = 1
    CURLCLOSEPOLICY_LEAST_RECENTLY_USED = 2
    CURLCLOSEPOLICY_LEAST_TRAFFIC = 3
    CURLCLOSEPOLICY_SLOWEST = 4
    CURLCLOSEPOLICY_CALLBACK = 5
    CURLCLOSEPOLICY_LAST = 6
End Enum
    
Public Enum curl_init_flag
    CURL_GLOBAL_NOTHING = 0
    CURL_GLOBAL_SSL = 1
    CURL_GLOBAL_WIN32 = 2
    CURL_GLOBAL_ALL = 3
    CURL_GLOBAL_DEFAULT = 3
End Enum
   
Public Enum curl_lock_data
    CURL_LOCK_DATA_NONE = 0
    CURL_LOCK_DATA_SHARE = 1
    CURL_LOCK_DATA_COOKIE = 2
    CURL_LOCK_DATA_DNS = 3
    CURL_LOCK_DATA_SSL_SESSION = 4
    CURL_LOCK_DATA_CONNECT = 5
    CURL_LOCK_DATA_LAST = 6
End Enum
   
Public Enum curl_lock_access
    CURL_LOCK_ACCESS_NONE = 0
    CURL_LOCK_ACCESS_SHARED = 1
    CURL_LOCK_ACCESS_SINGLE = 2
    CURL_LOCK_ACCESS_LAST = 3
End Enum
    
Public Enum CURLSHcode
    CURLSHE_OK = 0
    CURLSHE_BAD_OPTION = 1
    CURLSHE_IN_USE = 2
    CURLSHE_INVALID = 3
    CURLSHE_NOMEM = 4
    CURLSHE_LAST = 5
End Enum
    
Public Enum CURLSHoption
    CURLSHOPT_NONE = 0
    CURLSHOPT_SHARE = 1
    CURLSHOPT_UNSHARE = 2
    CURLSHOPT_LOCKFUNC = 3
    CURLSHOPT_UNLOCKFUNC = 4
    CURLSHOPT_USERDATA = 5
    CURLSHOPT_LAST = 6
End Enum
    
Public Enum CURLversion
    CURLVERSION_FIRST = 0
    CURLVERSION_SECOND = 1
    CURLVERSION_THIRD = 2
    CURLVERSION_NOW = 2
End Enum
    
Public Enum CURLversionFeatureBitmask
    CURL_VERSION_IPV6 = 1
    CURL_VERSION_KERBEROS4 = 2
    CURL_VERSION_SSL = 4
    CURL_VERSION_LIBZ = 8
    CURL_VERSION_NTLM = 16
    CURL_VERSION_GSSNEGOTIATE = 32
    CURL_VERSION_DEBUG = 64
    CURL_VERSION_ASYNCHDNS = 128
    CURL_VERSION_SPNEGO = 256
    CURL_VERSION_LARGEFILE = 512
    CURL_VERSION_IDN = 1024
End Enum
    
Public Enum CURLMSG
    CURLMSG_NONE = 0
    CURLMSG_DONE = 1
    CURLMSG_LAST = 2
End Enum
   
Public Enum CURLMcode
    CURLM_CALL_MULTI_PERFORM = -1
    CURLM_OK = 0
    CURLM_BAD_HANDLE = 1
    CURLM_BAD_EASY_HANDLE = 2
    CURLM_OUT_OF_MEMORY = 3
    CURLM_INTERNAL_ERROR = 4
    CURLM_LAST = 5
End Enum



'basic api
'---------------------------------------------------------------------
'[entry(0x60000000), helpstring("Cleanup an easy session")]
'void _stdcall vbcurl_easy_cleanup([in] long easy);
Public Declare Sub vbcurl_easy_cleanup Lib "vblibcurl.dll" (ByVal easy As Long)

'[entry(0x60000001), helpstring("Duplicate an easy handle")]
'long _stdcall vbcurl_easy_duphandle([in] long easy);

'[entry(0x60000002), helpstring("Get information on an easy session")]
'CURLcode _stdcall vbcurl_easy_getinfo(
'                [in] long easy,
'                [in] CURLINFO info,
'                [in] VARIANT* pv);
Public Declare Function vbcurl_easy_getinfo Lib "vblibcurl.dll" ( _
    ByVal easy As Long, _
    ByVal info As CURLINFO, _
    ByRef Value As Variant _
) As CURLcode




'[entry(0x60000003), helpstring("Initialize an easy session")]
'long _stdcall vbcurl_easy_init();
Public Declare Function vbcurl_easy_init Lib "vblibcurl.dll" () As Long
'
'If you did not already call curl_global_init, curl_easy_init does it automatically.
'This may be lethal in multi-threaded cases, since curl_global_init is not thread-safe,
'and it may result in resource problems because there is no corresponding cleanup.
'
'You are strongly advised to not allow this automatic behaviour,
'https://curl.se/libcurl/c/curl_easy_init.html
'
'Note: curl_global_init not exposed by vblibcurl and it only calls curl_easy_init



'[entry(0x60000004), helpstring("Perform an easy transfer")]
'CURLcode _stdcall vbcurl_easy_perform([in] long easy);
Public Declare Function vbcurl_easy_perform Lib "vblibcurl.dll" (ByVal easy As Long) As CURLcode


'[entry(0x60000005), helpstring("Reset an easy handle")]
'void _stdcall vbcurl_easy_reset([in] long easy);

'[entry(0x60000006), helpstring("Set option for easy transfer")]
'CURLcode _stdcall vbcurl_easy_setopt(
'                [in] long easy,
'                [in] CURLoption opt,
'                [in] VARIANT* value);
Public Declare Function vbcurl_easy_setopt Lib "vblibcurl.dll" ( _
    ByVal easy As Long, _
    ByVal opt As CURLoption, _
    ByRef Value As Variant _
) As CURLcode



'[entry(0x60000007), helpstring("Get a string description of an error code")]
'BSTR _stdcall vbcurl_easy_strerror([in] CURLcode err);



'forms
'---------------------------------------------------------------------
'[entry(0x60000008), helpstring("Add two option/value pairs to a form part")]
'CURLFORMcode _stdcall vbcurl_form_add_four_to_part(
'                [in] long part,
'                [in] CURLformoption opt1,
'                [in] VARIANT* val1,
'                [in] CURLformoption opt2,
'                [in] VARIANT* val2);

'[entry(0x60000009), helpstring("Add an option/value pair to a form part")]
'CURLFORMcode _stdcall vbcurl_form_add_pair_to_part(
'                [in] long part,
'                [in] CURLformoption opt,
'                [in] VARIANT* val);

'[entry(0x6000000a), helpstring("Add a completed part to a multi-part form")]
'CURLFORMcode _stdcall vbcurl_form_add_part(
'                [in] long form,
'                [in] long part);

'[entry(0x6000000b), helpstring("Add three option/value pairs to a form part")]
'CURLFORMcode _stdcall vbcurl_form_add_six_to_part(
'                [in] long part,
'                [in] CURLformoption opt1,
'                [in] VARIANT* val1,
'                [in] CURLformoption opt2,
'                [in] VARIANT* val2,
'                [in] CURLformoption opt3,
'                [in] VARIANT* val3);

'[entry(0x6000000c), helpstring("Create a multi-part form handle")]
'long _stdcall vbcurl_form_create();

'[entry(0x6000000d), helpstring("Create a multi-part form-part")]
'long _stdcall vbcurl_form_create_part([in] long form);

'[entry(0x6000000e), helpstring("Free a form and all its parts")]
'void _stdcall vbcurl_form_free([in] long form);




'multi handles
'---------------------------------------------------------------------
'[entry(0x6000000f), helpstring("Add an easy handle")]
'CURLMcode _stdcall vbcurl_multi_add_handle(
'                [in] long multi,
'                [in] long easy);

'[entry(0x60000010), helpstring("Cleanup a multi handle")]
'CURLMcode _stdcall vbcurl_multi_cleanup([in] long multi);

'[entry(0x60000011), helpstring("Call fdset on internal sockets")]
'CURLMcode _stdcall vbcurl_multi_fdset([in] long multi);

'[entry(0x60000012), helpstring("Read per-easy info for a multi handle")]
'long _stdcall vbcurl_multi_info_read(
'                [in] long multi,
'                [in, out] CURLMSG* msg,
'                [in, out] long* easy,
'                [in, out] CURLcode* code);

'[entry(0x60000013), helpstring("Initialize a multi handle")]
'long _stdcall vbcurl_multi_init();

'[entry(0x60000014), helpstring("Read/write the easy handles")]
'CURLMcode _stdcall vbcurl_multi_perform(
'                [in] long multi,
'                [in, out] long* runningHandles);

'[entry(0x60000015), helpstring("Remove an easy handle")]
'CURLMcode _stdcall vbcurl_multi_remove_handle(
'                [in] long multi,
'                [in] long easy);

'[entry(0x60000016), helpstring("Perform select on easy handles")]
'long _stdcall vbcurl_multi_select(
'                [in] long multi,
'                [in] long timeoutMillis);

'[entry(0x60000017), helpstring("Get a string description of an error code")]
'BSTR _stdcall vbcurl_multi_strerror([in] CURLMcode err);





'string lists
'---------------------------------------------------------------------
'[entry(0x60000018), helpstring("Append a string to an slist")]
'void _stdcall vbcurl_slist_append(
'                [in] long slist,
'                [in] BSTR str);

'[entry(0x60000019), helpstring("Create a string list")]
'long _stdcall vbcurl_slist_create();

'[entry(0x6000001a), helpstring("Free a created string list")]
'void _stdcall vbcurl_slist_free([in] long slist);




'escape/unescape
'---------------------------------------------------------------------
'[entry(0x6000001b), helpstring("Escape an URL")]
'BSTR _stdcall vbcurl_string_escape(
'                [in] BSTR url,
'                [in] long len);

'[entry(0x6000001c), helpstring("Unescape an URL")]
'BSTR _stdcall vbcurl_string_unescape(
'                [in] BSTR url,
'                [in] long len);




'version/build info
'---------------------------------------------------------------------
'[entry(0x6000001d), helpstring("Get the underlying libcurl version string")]
'BSTR _stdcall vbcurl_string_version();

'[entry(0x6000001e), helpstring("Age of libcurl version")]
'long _stdcall vbcurl_version_age([in] long ver);

'[entry(0x6000001f), helpstring("ARES version string")]
'BSTR _stdcall vbcurl_version_ares([in] long ver);

'[entry(0x60000020), helpstring("ARES version number")]
'long _stdcall vbcurl_version_ares_num([in] long ver);

'[entry(0x60000021), helpstring("Bitmask of supported features")]
'long _stdcall vbcurl_version_features([in] long ver);

'[entry(0x60000022), helpstring("Info of host on which libcurl was built")]
'BSTR _stdcall vbcurl_version_host([in] long ver);

'[entry(0x60000023), helpstring("Get libcurl version info")]
'long _stdcall vbcurl_version_info([in] CURLversion age);

'[entry(0x60000024), helpstring("Get libidn version, if present")]
'BSTR _stdcall vbcurl_version_libidn([in] long ver);

'[entry(0x60000025), helpstring("Get libz version, if present")]
'BSTR _stdcall vbcurl_version_libz([in] long ver);

'[entry(0x60000026), helpstring("Get numeric version number")]
'long _stdcall vbcurl_version_num([in] long ver);

'[entry(0x60000027), helpstring("Get supported protocols")]
'void _stdcall vbcurl_version_protocols(
'                [in] long ver,
'                [out] SAFEARRAY(BSTR)* ppsa);

'[entry(0x60000028), helpstring("Get SSL version string")]
'BSTR _stdcall vbcurl_version_ssl([in] long ver);

'[entry(0x60000029), helpstring("Get SSL version number")]
'long _stdcall vbcurl_version_ssl_num([in] long ver);

'[entry(0x6000002a), helpstring("Get version string")]
'BSTR _stdcall vbcurl_version_string([in] long ver);



'enum2Text functions for logging...
'--------------------------------------------------------------------
Function info2Text(i As curl_infotype) As String
    
    Dim s As String
    
    If i = CURLINFO_TEXT Then s = "TEXT"                '0
    If i = CURLINFO_HEADER_IN Then s = "HEADER_IN"      '1
    If i = CURLINFO_HEADER_OUT Then s = "HEADER_OUT"    '2
    If i = CURLINFO_DATA_IN Then s = "DATA_IN"          '3
    If i = CURLINFO_DATA_OUT Then s = "DATA_OUT"        '4
    If i = CURLINFO_SSL_DATA_IN Then s = "SSL_IN"       '5
    If i = CURLINFO_SSL_DATA_OUT Then s = "SSL_OUT"     '6
    If i = CURLINFO_END Then s = "END"                  '7
    If Len(s) = 0 Then s = "Unknown " & i
    
    info2Text = s
    
End Function

Function curlCode2Text(X As CURLcode) As String
    
    Dim s As String
    s = "Unknown " & X
    
    If X = 0 Then s = "CURLE_OK"
    If X = 42 Then s = "CURLE_ABORTED_BY_CALLBACK"
    If X = 44 Then s = "CURLE_BAD_CALLING_ORDER"
    If X = 61 Then s = "CURLE_BAD_CONTENT_ENCODING"
    If X = 36 Then s = "CURLE_BAD_DOWNLOAD_RESUME"
    If X = 43 Then s = "CURLE_BAD_FUNCTION_ARGUMENT"
    If X = 46 Then s = "CURLE_BAD_PASSWORD_ENTERED"
    If X = 7 Then s = "CURLE_COULDNT_CONNECT"
    If X = 6 Then s = "CURLE_COULDNT_RESOLVE_HOST"
    If X = 5 Then s = "CURLE_COULDNT_RESOLVE_PROXY"
    If X = 2 Then s = "CURLE_FAILED_INIT"
    If X = 63 Then s = "CURLE_FILESIZE_EXCEEDED"
    If X = 37 Then s = "CURLE_FILE_COULDNT_READ_FILE"
    If X = 9 Then s = "CURLE_FTP_ACCESS_DENIED"
    If X = 15 Then s = "CURLE_FTP_CANT_GET_HOST"
    If X = 16 Then s = "CURLE_FTP_CANT_RECONNECT"
    If X = 32 Then s = "CURLE_FTP_COULDNT_GET_SIZE"
    If X = 19 Then s = "CURLE_FTP_COULDNT_RETR_FILE"
    If X = 29 Then s = "CURLE_FTP_COULDNT_SET_ASCII"
    If X = 17 Then s = "CURLE_FTP_COULDNT_SET_BINARY"
    If X = 25 Then s = "CURLE_FTP_COULDNT_STOR_FILE"
    If X = 31 Then s = "CURLE_FTP_COULDNT_USE_REST"
    If X = 30 Then s = "CURLE_FTP_PORT_FAILED"
    If X = 21 Then s = "CURLE_FTP_QUOTE_ERROR"
    If X = 64 Then s = "CURLE_FTP_SSL_FAILED"
    If X = 10 Then s = "CURLE_FTP_USER_PASSWORD_INCORRECT"
    If X = 14 Then s = "CURLE_FTP_WEIRD_227_FORMAT"
    If X = 11 Then s = "CURLE_FTP_WEIRD_PASS_REPLY"
    If X = 13 Then s = "CURLE_FTP_WEIRD_PASV_REPLY"
    If X = 8 Then s = "CURLE_FTP_WEIRD_SERVER_REPLY"
    If X = 12 Then s = "CURLE_FTP_WEIRD_USER_REPLY"
    If X = 20 Then s = "CURLE_FTP_WRITE_ERROR"
    If X = 41 Then s = "CURLE_FUNCTION_NOT_FOUND"
    If X = 52 Then s = "CURLE_GOT_NOTHING"
    If X = 34 Then s = "CURLE_HTTP_POST_ERROR"
    If X = 33 Then s = "CURLE_HTTP_RANGE_ERROR"
    If X = 22 Then s = "CURLE_HTTP_RETURNED_ERROR"
    If X = 45 Then s = "CURLE_INTERFACE_FAILED"
    If X = 67 Then s = "CURLE_LAST"
    If X = 38 Then s = "CURLE_LDAP_CANNOT_BIND"
    If X = 62 Then s = "CURLE_LDAP_INVALID_URL"
    If X = 39 Then s = "CURLE_LDAP_SEARCH_FAILED"
    If X = 40 Then s = "CURLE_LIBRARY_NOT_FOUND"
    If X = 24 Then s = "CURLE_MALFORMAT_USER"
    If X = 50 Then s = "CURLE_OBSOLETE"
    If X = 28 Then s = "CURLE_OPERATION_TIMEOUTED"
    If X = 27 Then s = "CURLE_OUT_OF_MEMORY"
    If X = 18 Then s = "CURLE_PARTIAL_FILE"
    If X = 26 Then s = "CURLE_READ_ERROR"
    If X = 56 Then s = "CURLE_RECV_ERROR"
    If X = 55 Then s = "CURLE_SEND_ERROR"
    If X = 65 Then s = "CURL_SEND_FAIL_REWIND"
    If X = 57 Then s = "CURLE_SHARE_IN_USE"
    If X = 60 Then s = "CURLE_SSL_CACERT"
    If X = 58 Then s = "CURLE_SSL_CERTPROBLEM"
    If X = 59 Then s = "CURLE_SSL_CIPHER"
    If X = 35 Then s = "CURLE_SSL_CONNECT_ERROR"
    If X = 66 Then s = "CURLE_SSL_ENGINE_INITFAILED"
    If X = 53 Then s = "CURLE_SSL_ENGINE_NOTFOUND"
    If X = 54 Then s = "CURLE_SSL_ENGINE_SETFAILED"
    If X = 51 Then s = "CURLE_SSL_PEER_CERTIFICATE"
    If X = 49 Then s = "CURLE_TELNET_OPTION_SYNTAX"
    If X = 47 Then s = "CURLE_TOO_MANY_REDIRECTS"
    If X = 48 Then s = "CURLE_UNKNOWN_TELNET_OPTION"
    If X = 1 Then s = "CURLE_UNSUPPORTED_PROTOCOL"
    If X = 3 Then s = "CURLE_URL_MALFORMAT"
    If X = 4 Then s = "CURLE_URL_MALFORMAT_USER"
    If X = 23 Then s = "CURLE_WRITE_ERROR"

    curlCode2Text = s
    
End Function

