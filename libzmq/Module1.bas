Attribute VB_Name = "Module1"

Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'/*  This function retrieves the errno as it is known to 0MQ library. The goal */
'/*  of this function is to make the code 100% portable, including where 0MQ   */
'/*  compiled with certain CRT library (on Windows) is linked to an            */
'/*  application that uses different CRT library.                              */
'ZMQ_EXPORT int zmq_errno (void);
Public Declare Function zmq_errno Lib "libzmq.dll" () As Long

'/*  Resolves system errors and 0MQ errors to human-readable string.           */
'ZMQ_EXPORT const char *zmq_strerror (int errnum);
Public Declare Function zmq_strerror Lib "libzmq.dll" (ByVal errNum As Long) As Long

'/*  Run-time API version detection                                            */
'ZMQ_EXPORT void zmq_version (int *major, int *minor, int *patch);
Public Declare Sub zmq_version Lib "libzmq.dll" (ByRef major As Long, ByRef minor As Long, ByRef patch As Long)


'/******************************************************************************/
'/*  0MQ infrastructure (a.k.a. context) initialisation & termination.         */
'/******************************************************************************/
'
'/*  Context options                                                           */
Public Enum zCtxOpt
    ZMQ_IO_THREADS = 1
    ZMQ_MAX_SOCKETS = 2
    ZMQ_SOCKET_LIMIT = 3
    ZMQ_THREAD_PRIORITY = 3
    ZMQ_THREAD_SCHED_POLICY = 4
    ZMQ_MAX_MSGSZ = 5
End Enum
'
'/*  Default for new contexts                                                  */
Const ZMQ_IO_THREADS_DFLT = 1
Const ZMQ_MAX_SOCKETS_DFLT = 1023
Const ZMQ_THREAD_PRIORITY_DFLT = -1
Const ZMQ_THREAD_SCHED_POLICY_DFLT = -1

'ZMQ_EXPORT void *zmq_ctx_new (void);
Public Declare Function zmq_ctx_new Lib "libzmq.dll" () As Long

'ZMQ_EXPORT int zmq_ctx_term (void *context);
Public Declare Function zmq_ctx_term Lib "libzmq.dll" (ByVal ctx As Long) As Long

'ZMQ_EXPORT int zmq_ctx_shutdown (void *context);
Public Declare Function zmq_ctx_shutdown Lib "libzmq.dll" (ByVal ctx As Long) As Long

'ZMQ_EXPORT int zmq_ctx_set (void *context, int option, int optval);
Public Declare Function zmq_ctx_set Lib "libzmq.dll" (ByVal ctx As Long, ByVal opt As Long, ByVal optVal As Long) As Long

'ZMQ_EXPORT int zmq_ctx_get (void *context, int option);
Public Declare Function zmq_ctx_get Lib "libzmq.dll" (ByVal ctx As Long, ByVal opt As Long) As Long

'/*  Old (legacy) API                                                          */
'ZMQ_EXPORT void *zmq_init (int io_threads);
Public Declare Function zmq_init Lib "libzmq.dll" (ByVal iothreads As Long) As Long

'ZMQ_EXPORT int zmq_term (void *context);
Public Declare Function zmq_term Lib "libzmq.dll" (ByVal ctx As Long) As Long

'ZMQ_EXPORT int zmq_ctx_destroy (void *context);
Public Declare Function zmq_ctx_destroy Lib "libzmq.dll" (ByVal ctx As Long) As Long


'/******************************************************************************/
'/*  0MQ message definition.                                                   */
'/******************************************************************************/
'
'typedef struct zmq_msg_t {
'#elif defined (_MSC_VER) && (defined (_M_IX86) || defined (_M_ARM_ARMV7VE))
'    __declspec (align (4)) unsigned char _ [64];
'} zmq_msg_t;

'typedef void (zmq_free_fn) (void *data, void *hint);
'
'ZMQ_EXPORT int zmq_msg_init (zmq_msg_t *msg);
'Public Declare Function zmq_msg_init Lib "libzmq.dll" (ByRef hMsg As Long) As Long

'ZMQ_EXPORT int zmq_msg_init_size (zmq_msg_t *msg, size_t size);
'Public Declare Function zmq_msg_init_size Lib "libzmq.dll" (ByRef hMsg As Long, ByVal sz As Long) As Long


'ZMQ_EXPORT int zmq_msg_init_data (zmq_msg_t *msg, void *data, size_t size, zmq_free_fn *ffn, void *hint);

'ZMQ_EXPORT int zmq_msg_send (zmq_msg_t *msg, void *s, int flags);
'Public Declare Function zmq_msg_send Lib "libzmq.dll" (ByRef hMsg As Long, ByVal s As Long, ByVal flags As Long) As Long

'ZMQ_EXPORT int zmq_msg_recv (zmq_msg_t *msg, void *s, int flags);
'Public Declare Function zmq_msg_recv Lib "libzmq.dll" (ByVal hMsg As Long, ByVal s As Long, ByVal flags As Long) As Long

'ZMQ_EXPORT int zmq_msg_close (zmq_msg_t *msg);
'Public Declare Function zmq_msg_close Lib "libzmq.dll" (ByRef hMsg As Long) As Long

'ZMQ_EXPORT int zmq_msg_move (zmq_msg_t *dest, zmq_msg_t *src);
'ZMQ_EXPORT int zmq_msg_copy (zmq_msg_t *dest, zmq_msg_t *src);
'ZMQ_EXPORT void *zmq_msg_data (zmq_msg_t *msg);
'ZMQ_EXPORT size_t zmq_msg_size (const zmq_msg_t *msg);
'ZMQ_EXPORT int zmq_msg_more (const zmq_msg_t *msg);
'ZMQ_EXPORT int zmq_msg_get (const zmq_msg_t *msg, int property);
'ZMQ_EXPORT int zmq_msg_set (zmq_msg_t *msg, int property, int optval);
'ZMQ_EXPORT const char *zmq_msg_gets (const zmq_msg_t *msg, const char *property);
'
'/******************************************************************************/
'/*  0MQ socket definition.                                                    */
'/******************************************************************************/
'
'/*  Socket types.                                                             */
Public Enum zSckType
    ZMQ_PAIR = 0
    ZMQ_PUB = 1
    ZMQ_SUB = 2
    ZMQ_REQ = 3
    ZMQ_REP = 4
    ZMQ_DEALER = 5
    ZMQ_ROUTER = 6
    ZMQ_PULL = 7
    ZMQ_PUSH = 8
    ZMQ_XPUB = 9
    ZMQ_XSUB = 10
    ZMQ_STREAM = 11
End Enum
'
'/*  Deprecated aliases                                                        */
'const  ZMQ_XREQ ZMQ_DEALER
'const  ZMQ_XREP ZMQ_ROUTER
'
'/*  Socket options.                                                           */
Public Enum zsckOpt
    ZMQ_AFFINITY = 4
    ZMQ_ROUTING_ID = 5
    ZMQ_SUBSCRIBE = 6
    ZMQ_UNSUBSCRIBE = 7
    ZMQ_RATE = 8
    ZMQ_RECOVERY_IVL = 9
    ZMQ_SNDBUF = 11
    ZMQ_RCVBUF = 12
    ZMQ_RCVMORE = 13
    ZMQ_FD = 14
    ZMQ_EVENTS = 15
    ZMQ_TYPE = 16
    ZMQ_LINGER = 17
    ZMQ_RECONNECT_IVL = 18
    ZMQ_BACKLOG = 19
    ZMQ_RECONNECT_IVL_MAX = 21
    ZMQ_MAXMSGSIZE = 22
    ZMQ_SNDHWM = 23
    ZMQ_RCVHWM = 24
    ZMQ_MULTICAST_HOPS = 25
    ZMQ_RCVTIMEO = 27
    ZMQ_SNDTIMEO = 28
    ZMQ_LAST_ENDPOINT = 32
    ZMQ_ROUTER_MANDATORY = 33
    ZMQ_TCP_KEEPALIVE = 34
    ZMQ_TCP_KEEPALIVE_CNT = 35
    ZMQ_TCP_KEEPALIVE_IDLE = 36
    ZMQ_TCP_KEEPALIVE_INTVL = 37
    ZMQ_IMMEDIATE = 39
    ZMQ_XPUB_VERBOSE = 40
    ZMQ_ROUTER_RAW = 41
    ZMQ_IPV6 = 42
    ZMQ_MECHANISM = 43
    ZMQ_PLAIN_SERVER = 44
    ZMQ_PLAIN_USERNAME = 45
    ZMQ_PLAIN_PASSWORD = 46
    ZMQ_CURVE_SERVER = 47
    ZMQ_CURVE_PUBLICKEY = 48
    ZMQ_CURVE_SECRETKEY = 49
    ZMQ_CURVE_SERVERKEY = 50
    ZMQ_PROBE_ROUTER = 51
    ZMQ_REQ_CORRELATE = 52
    ZMQ_REQ_RELAXED = 53
    ZMQ_CONFLATE = 54
    ZMQ_ZAP_DOMAIN = 55
    ZMQ_ROUTER_HANDOVER = 56
    ZMQ_TOS = 57
    ZMQ_CONNECT_ROUTING_ID = 61
    ZMQ_GSSAPI_SERVER = 62
    ZMQ_GSSAPI_PRINCIPAL = 63
    ZMQ_GSSAPI_SERVICE_PRINCIPAL = 64
    ZMQ_GSSAPI_PLAINTEXT = 65
    ZMQ_HANDSHAKE_IVL = 66
    ZMQ_SOCKS_PROXY = 68
    ZMQ_XPUB_NODROP = 69
    ZMQ_BLOCKY = 70
    ZMQ_XPUB_MANUAL = 71
    ZMQ_XPUB_WELCOME_MSG = 72
    ZMQ_STREAM_NOTIFY = 73
    ZMQ_INVERT_MATCHING = 74
    ZMQ_HEARTBEAT_IVL = 75
    ZMQ_HEARTBEAT_TTL = 76
    ZMQ_HEARTBEAT_TIMEOUT = 77
    ZMQ_XPUB_VERBOSER = 78
    ZMQ_CONNECT_TIMEOUT = 79
    ZMQ_TCP_MAXRT = 80
    ZMQ_THREAD_SAFE = 81
    ZMQ_MULTICAST_MAXTPDU = 84
    ZMQ_VMCI_BUFFER_SIZE = 85
    ZMQ_VMCI_BUFFER_MIN_SIZE = 86
    ZMQ_VMCI_BUFFER_MAX_SIZE = 87
    ZMQ_VMCI_CONNECT_TIMEOUT = 88
    ZMQ_USE_FD = 89
End Enum

'
'/*  Message options                                                           */
Public Enum zMsgOpt
    ZMQ_MORE = 1
    ZMQ_SHARED = 3
End Enum

'/*  Send/recv options.                                                        */
Public Enum zSendRecvOpt
    ZMQ_DONTWAIT = 1
    ZMQ_SNDMORE = 2
End Enum

'/*  Security mechanisms                                                       */
Public Enum zSecOpt
    ZMQ_NULL = 0
    ZMQ_PLAIN = 1
    ZMQ_CURVE = 2
    ZMQ_GSSAPI = 3
End Enum

'
'/*  RADIO-DISH protocol                                                       */
'const  ZMQ_GROUP_MAX_LENGTH        15
'
'/*  Deprecated options and aliases                                            */
'const  ZMQ_IDENTITY                ZMQ_ROUTING_ID
'const  ZMQ_CONNECT_RID             ZMQ_CONNECT_ROUTING_ID
'const  ZMQ_TCP_ACCEPT_FILTER       38
'const  ZMQ_IPC_FILTER_PID          58
'const  ZMQ_IPC_FILTER_UID          59
'const  ZMQ_IPC_FILTER_GID          60
'const  ZMQ_IPV4ONLY                31
'const  ZMQ_DELAY_ATTACH_ON_CONNECT ZMQ_IMMEDIATE
'const  ZMQ_NOBLOCK                 ZMQ_DONTWAIT
'const  ZMQ_FAIL_UNROUTABLE         ZMQ_ROUTER_MANDATORY
'const  ZMQ_ROUTER_BEHAVIOR         ZMQ_ROUTER_MANDATORY
'
'/*  Deprecated Message options                                                */
'const  ZMQ_SRCFD 2
'
'/******************************************************************************/
'/*  0MQ socket events and monitoring                                          */
'/******************************************************************************/
'
'/*  Socket transport events (TCP, IPC and TIPC only)                          */
'
Public Enum zEvents
    ZMQ_EVENT_CONNECTED = &H1
    ZMQ_EVENT_CONNECT_DELAYED = &H2
    ZMQ_EVENT_CONNECT_RETRIED = &H4
    ZMQ_EVENT_LISTENING = &H8
    ZMQ_EVENT_BIND_FAILED = &H10
    ZMQ_EVENT_ACCEPTED = &H20
    ZMQ_EVENT_ACCEPT_FAILED = &H40
    ZMQ_EVENT_CLOSED = &H80
    ZMQ_EVENT_CLOSE_FAILED = &H100
    ZMQ_EVENT_DISCONNECTED = &H200
    ZMQ_EVENT_MONITOR_STOPPED = &H400
    ZMQ_EVENT_ALL = &HFFFF
End Enum
'
'ZMQ_EXPORT void *zmq_socket (void *, int type);
Public Declare Function zmq_socket Lib "libzmq.dll" (ByVal ctx As Long, ByVal typ As zSckType) As Long

'ZMQ_EXPORT int zmq_close (void *s);
Public Declare Function zmq_close Lib "libzmq.dll" (ByVal s As Long) As Long

'ZMQ_EXPORT int zmq_setsockopt (void *s, int option, const void *optval,size_t optvallen);
Public Declare Function zmq_setsockopt Lib "libzmq.dll" (ByVal s As Long, ByVal opt As zsckOpt, ByRef optVal As Long, ByVal optValLen As Long) As Long

'ZMQ_EXPORT int zmq_getsockopt (void *s, int option, void *optval, size_t *optvallen);
Public Declare Function zmq_getsockopt Lib "libzmq.dll" (ByVal s As Long, ByVal opt As zsckOpt, ByRef optVal As Long, ByVal optValLen As Long) As Long

'ZMQ_EXPORT int zmq_bind (void *s, const char *addr);
Public Declare Function zmq_bind Lib "libzmq.dll" (ByVal s As Long, ByVal addr As String) As Long

'ZMQ_EXPORT int zmq_connect (void *s, const char *addr);
Public Declare Function zmq_connect Lib "libzmq.dll" (ByVal s As Long, ByVal addr As String) As Long

'ZMQ_EXPORT int zmq_unbind (void *s, const char *addr);
Public Declare Function zmq_unbind Lib "libzmq.dll" (ByVal s As Long, ByVal addr As String) As Long

'ZMQ_EXPORT int zmq_disconnect (void *s, const char *addr);
Public Declare Function zmq_disconnect Lib "libzmq.dll" (ByVal s As Long, ByVal addr As String) As Long

'ZMQ_EXPORT int zmq_send (void *s, const void *buf, size_t len, int flags);
Public Declare Function zmq_send Lib "libzmq.dll" (ByVal s As Long, ByVal buf As String, ByVal leng As Long, ByVal flags As Long) As Long

'ZMQ_EXPORT int zmq_send_const (void *s, const void *buf, size_t len, int flags);

'ZMQ_EXPORT int zmq_recv (void *s, void *buf, size_t len, int flags);
Public Declare Function zmq_recv Lib "libzmq.dll" (ByVal s As Long, ByRef b As Byte, ByVal sz As Long, ByVal flags As Long) As Long

'ZMQ_EXPORT int zmq_socket_monitor (void *s, const char *addr, int events);
'
'
'/******************************************************************************/
'/*  Deprecated I/O multiplexing. Prefer using zmq_poller API                  */
'/******************************************************************************/
Public Enum pollerEvents
    ZMQ_POLLIN = 1
    ZMQ_POLLOUT = 2
    ZMQ_POLLERR = 4
    ZMQ_POLLPRI = 8
End Enum
'
'typedef struct zmq_pollitem_t
'{
'    void *socket;
'    SOCKET fd;
'    short events;
'    short revents;
'} zmq_pollitem_t;
'
'const  ZMQ_POLLITEMS_DFLT 16
'
'ZMQ_EXPORT int  zmq_poll (zmq_pollitem_t *items, int nitems, long timeout);
'
'/******************************************************************************/
'/*  Message proxying                                                          */
'/******************************************************************************/
'
'ZMQ_EXPORT int zmq_proxy (void *frontend, void *backend, void *capture);
'ZMQ_EXPORT int zmq_proxy_steerable (void *frontend, void *backend, void *capture, void *control);
'
'/******************************************************************************/
'/*  Probe library capabilities                                                */
'/******************************************************************************/
'
'const  ZMQ_HAS_CAPABILITIES 1
'ZMQ_EXPORT int zmq_has (const char *capability);
'
'/*  Deprecated aliases */
'const  ZMQ_STREAMER 1
'const  ZMQ_FORWARDER 2
'const  ZMQ_QUEUE 3
'
'/*  Deprecated methods */
'ZMQ_EXPORT int zmq_device (int type, void *frontend, void *backend);
'ZMQ_EXPORT int zmq_sendmsg (void *s, zmq_msg_t *msg, int flags);
Public Declare Function zmq_sendmsg Lib "libzmq.dll" (ByVal s As Long, ByRef msg As Long, ByVal flags As Long) As Long


'ZMQ_EXPORT int zmq_recvmsg (void *s, zmq_msg_t *msg, int flags);
'struct iovec;
'ZMQ_EXPORT int zmq_sendiov (void *s, struct iovec *iov, size_t count, int flags);
'ZMQ_EXPORT int zmq_recviov (void *s, struct iovec *iov, size_t *count, int flags);
'
'/******************************************************************************/
'/*  Encryption functions                                                      */
'/******************************************************************************/
'
'/*  Encode data with Z85 encoding. Returns encoded data                       */
'ZMQ_EXPORT char *zmq_z85_encode (char *dest, const uint8_t *data, size_t size);
'
'/*  Decode data with Z85 encoding. Returns decoded data                       */
'ZMQ_EXPORT uint8_t *zmq_z85_decode (uint8_t *dest, const char *string);
'
'/*  Generate z85-encoded public and private keypair with tweetnacl/libsodium. */
'/*  Returns 0 on success.                                                     */
'ZMQ_EXPORT int zmq_curve_keypair (char *z85_public_key, char *z85_secret_key);
'
'/*  Derive the z85-encoded public key from the z85-encoded secret key.        */
'/*  Returns 0 on success.                                                     */
'ZMQ_EXPORT int zmq_curve_public (char *z85_public_key, const char *z85_secret_key);
'
'/******************************************************************************/
'/*  Atomic utility methods                                                    */
'/******************************************************************************/
'
'ZMQ_EXPORT void *zmq_atomic_counter_new (void);
'ZMQ_EXPORT void zmq_atomic_counter_set (void *counter, int value);
'ZMQ_EXPORT int zmq_atomic_counter_inc (void *counter);
'ZMQ_EXPORT int zmq_atomic_counter_dec (void *counter);
'ZMQ_EXPORT int zmq_atomic_counter_value (void *counter);
'ZMQ_EXPORT void zmq_atomic_counter_destroy (void **counter_p);
'
'
'/******************************************************************************/
'/*  These functions are not documented by man pages -- use at your own risk.  */
'/*  If you need these to be part of the formal ZMQ API, then (a) write a man  */
'/*  page, and (b) write a test case in tests.                                 */
'/******************************************************************************/
'
'/*  Helper functions are used by perf tests so that they don't have to care   */
'/*  about minutiae of time-related functions on different OS platforms.       */
'
'/*  Starts the stopwatch. Returns the handle to the watch.                    */
'ZMQ_EXPORT void *zmq_stopwatch_start (void);
'
'/*  Stops the stopwatch. Returns the number of microseconds elapsed since     */
'/*  the stopwatch was started.                                                */
'ZMQ_EXPORT unsigned long zmq_stopwatch_stop (void *watch_);
'
'/*  Sleeps for specified number of seconds.                                   */
'ZMQ_EXPORT void zmq_sleep (int seconds_);
'
'typedef void (zmq_thread_fn) (void*);
'
'/* Start a thread. Returns a handle to the thread.                            */
'ZMQ_EXPORT void *zmq_threadstart (zmq_thread_fn* func, void* arg);
'
'/* Wait for thread to complete then free up resources.                        */
'ZMQ_EXPORT void zmq_threadclose (void* thread);

'/******************************************************************************/
'/*  These functions are DRAFT and disabled in stable releases, and subject to */
'/*  change at ANY time until declared stable.                                 */
'/******************************************************************************/
'
'#ifdef ZMQ_BUILD_DRAFT_API
'
'/*  DRAFT Socket types.                                                       */
'#define ZMQ_SERVER 12
'#define ZMQ_CLIENT 13
'#define ZMQ_RADIO 14
'#define ZMQ_DISH 15
'#define ZMQ_GATHER 16
'#define ZMQ_SCATTER 17
'#define ZMQ_DGRAM 18
'
'/*  DRAFT Socket options.                                                     */
'#define ZMQ_GSSAPI_PRINCIPAL_NAMETYPE 90
'#define ZMQ_GSSAPI_SERVICE_PRINCIPAL_NAMETYPE 91
'#define ZMQ_BINDTODEVICE 92
'#define ZMQ_ZAP_ENFORCE_DOMAIN 93
'
'/*  DRAFT 0MQ socket events and monitoring                                    */
'/*  Unspecified system errors during handshake. Event value is an errno.      */
'#define ZMQ_EVENT_HANDSHAKE_FAILED_NO_DETAIL   0x0800
'/*  Handshake complete successfully with successful authentication (if        *
' *  enabled). Event value is unused.                                          */
'#define ZMQ_EVENT_HANDSHAKE_SUCCEEDED          0x1000
'/*  Protocol errors between ZMTP peers or between server and ZAP handler.     *
' *  Event value is one of ZMQ_PROTOCOL_ERROR_*                                */
'#define ZMQ_EVENT_HANDSHAKE_FAILED_PROTOCOL    0x2000
'/*  Failed authentication requests. Event value is the numeric ZAP status     *
' *  code, i.e. 300, 400 or 500.                                               */
'#define ZMQ_EVENT_HANDSHAKE_FAILED_AUTH        0x4000
'
'#define ZMQ_PROTOCOL_ERROR_ZMTP_UNSPECIFIED 0x10000000
'#define ZMQ_PROTOCOL_ERROR_ZMTP_UNEXPECTED_COMMAND 0x10000001
'#define ZMQ_PROTOCOL_ERROR_ZMTP_INVALID_SEQUENCE 0x10000002
'#define ZMQ_PROTOCOL_ERROR_ZMTP_KEY_EXCHANGE 0x10000003
'#define ZMQ_PROTOCOL_ERROR_ZMTP_MALFORMED_COMMAND_UNSPECIFIED 0x10000011
'#define ZMQ_PROTOCOL_ERROR_ZMTP_MALFORMED_COMMAND_MESSAGE 0x10000012
'#define ZMQ_PROTOCOL_ERROR_ZMTP_MALFORMED_COMMAND_HELLO 0x10000013
'#define ZMQ_PROTOCOL_ERROR_ZMTP_MALFORMED_COMMAND_INITIATE 0x10000014
'#define ZMQ_PROTOCOL_ERROR_ZMTP_MALFORMED_COMMAND_ERROR 0x10000015
'#define ZMQ_PROTOCOL_ERROR_ZMTP_MALFORMED_COMMAND_READY 0x10000016
'#define ZMQ_PROTOCOL_ERROR_ZMTP_MALFORMED_COMMAND_WELCOME 0x10000017
'#define ZMQ_PROTOCOL_ERROR_ZMTP_INVALID_METADATA 0x10000018
'
'// the following two may be due to erroneous configuration of a peer
'#define ZMQ_PROTOCOL_ERROR_ZMTP_CRYPTOGRAPHIC 0x11000001
'#define ZMQ_PROTOCOL_ERROR_ZMTP_MECHANISM_MISMATCH 0x11000002
'
'#define ZMQ_PROTOCOL_ERROR_ZAP_UNSPECIFIED     0x20000000
'#define ZMQ_PROTOCOL_ERROR_ZAP_MALFORMED_REPLY 0x20000001
'#define ZMQ_PROTOCOL_ERROR_ZAP_BAD_REQUEST_ID 0x20000002
'#define ZMQ_PROTOCOL_ERROR_ZAP_BAD_VERSION 0x20000003
'#define ZMQ_PROTOCOL_ERROR_ZAP_INVALID_STATUS_CODE 0x20000004
'#define ZMQ_PROTOCOL_ERROR_ZAP_INVALID_METADATA 0x20000005
'
'/*  DRAFT Context options                                                     */
'#define ZMQ_MSG_T_SIZE 6
'#define ZMQ_THREAD_AFFINITY_CPU_ADD 7
'#define ZMQ_THREAD_AFFINITY_CPU_REMOVE 8
'#define ZMQ_THREAD_NAME_PREFIX 9
'
'/*  DRAFT Socket methods.                                                     */
'ZMQ_EXPORT int zmq_join (void *s, const char *group);
'ZMQ_EXPORT int zmq_leave (void *s, const char *group);
'
'/*  DRAFT Msg methods.                                                        */
'ZMQ_EXPORT int zmq_msg_set_routing_id(zmq_msg_t *msg, uint32_t routing_id);
'ZMQ_EXPORT uint32_t zmq_msg_routing_id(zmq_msg_t *msg);
'ZMQ_EXPORT int zmq_msg_set_group(zmq_msg_t *msg, const char *group);
'ZMQ_EXPORT const char *zmq_msg_group(zmq_msg_t *msg);
'
'/*  DRAFT Msg property names.                                                 */
'#define ZMQ_MSG_PROPERTY_ROUTING_ID    "Routing-Id"
'#define ZMQ_MSG_PROPERTY_SOCKET_TYPE   "Socket-Type"
'#define ZMQ_MSG_PROPERTY_USER_ID       "User-Id"
'#define ZMQ_MSG_PROPERTY_PEER_ADDRESS  "Peer-Address"
'
'/******************************************************************************/
'/*  Poller polling on sockets,fd and thread-safe sockets                      */
'/******************************************************************************/
'
'#define ZMQ_HAVE_POLLER
'
'typedef struct zmq_poller_event_t
'{
'    void *socket;
'#if defined _WIN32
'    SOCKET fd;
'#Else
'    int fd;
'#End If
'    void *user_data;
'    short events;
'} zmq_poller_event_t;

Type zmq_poller_event
    socket As Long
    fd As Long
    userData As Long
    events As Integer
End Type

'ZMQ_EXPORT void *zmq_poller_new (void);
Public Declare Function zmq_poller_new Lib "libzmq.dll" () As Long

'ZMQ_EXPORT int  zmq_poller_destroy (void **poller_p);
Public Declare Function zmq_poller_destroy Lib "libzmq.dll" (ByRef hPoller As Long) As Long

'ZMQ_EXPORT int  zmq_poller_add (void *poller, void *socket, void *user_data, short events);
Public Declare Function zmq_poller_add Lib "libzmq.dll" (ByVal hPoller As Long, ByVal s As Long, ByRef userData As Long, ByVal events As pollerEvents) As Long

'ZMQ_EXPORT int  zmq_poller_modify (void *poller, void *socket, short events);
'ZMQ_EXPORT int  zmq_poller_remove (void *poller, void *socket);
'ZMQ_EXPORT int  zmq_poller_wait (void *poller, zmq_poller_event_t *event, long timeout);
Public Declare Function zmq_poller_wait Lib "libzmq.dll" (ByVal hPoller As Long, ByRef evt As zmq_poller_event, ByVal timeout As Long) As Long


'ZMQ_EXPORT int  zmq_poller_wait_all (void *poller, zmq_poller_event_t *events, int n_events, long timeout);
'
'#if defined _WIN32
'ZMQ_EXPORT int zmq_poller_add_fd (void *poller, SOCKET fd, void *user_data, short events);
'ZMQ_EXPORT int zmq_poller_modify_fd (void *poller, SOCKET fd, short events);
'ZMQ_EXPORT int zmq_poller_remove_fd (void *poller, SOCKET fd);
'#Else
'ZMQ_EXPORT int zmq_poller_add_fd (void *poller, int fd, void *user_data, short events);
'ZMQ_EXPORT int zmq_poller_modify_fd (void *poller, int fd, short events);
'ZMQ_EXPORT int zmq_poller_remove_fd (void *poller, int fd);
'#End If
'
'ZMQ_EXPORT int zmq_socket_get_peer_state (void *socket,
'                                          const void *routing_id,
'                                          size_t routing_id_size);
'
'/******************************************************************************/
'/*  Scheduling timers                                                         */
'/******************************************************************************/
'
'#define ZMQ_HAVE_TIMERS
'
'typedef void (zmq_timer_fn)(int timer_id, void *arg);
'
'ZMQ_EXPORT void *zmq_timers_new (void);
'ZMQ_EXPORT int   zmq_timers_destroy (void **timers_p);
'ZMQ_EXPORT int   zmq_timers_add (void *timers, size_t interval, zmq_timer_fn handler, void *arg);
'ZMQ_EXPORT int   zmq_timers_cancel (void *timers, int timer_id);
'ZMQ_EXPORT int   zmq_timers_set_interval (void *timers, int timer_id, size_t interval);
'ZMQ_EXPORT int   zmq_timers_reset (void *timers, int timer_id);
'ZMQ_EXPORT long  zmq_timers_timeout (void *timers);
'ZMQ_EXPORT int   zmq_timers_execute (void *timers);
'
'/******************************************************************************/
'/*  GSSAPI definitions                                                        */
'/******************************************************************************/
'
'/*  GSSAPI principal name types                                               */
'#define ZMQ_GSSAPI_NT_HOSTBASED 0
'#define ZMQ_GSSAPI_NT_USER_NAME 1
'#define ZMQ_GSSAPI_NT_KRB5_PRINCIPAL 2
'
'#endif // ZMQ_BUILD_DRAFT_API


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, Source As Any, ByVal length As Long)
Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long

Function zError(code As Long) As String
    Dim lpStr As Long
    lpStr = zmq_strerror(code)
    zError = StringFromPointer(lpStr)
End Function

Function StringFromPointer(buf As Long) As String
    Dim sz As Long
    Dim tmp As String
    Dim b() As Byte
    
    If buf = 0 Then Exit Function
       
    sz = lstrlen(buf)
    If sz = 0 Then Exit Function
    
    ReDim b(sz)
    CopyMemory b(0), ByVal buf, sz
    tmp = StrConv(b, vbUnicode)
    If Right(tmp, 1) = Chr(0) Then tmp = Left(tmp, Len(tmp) - 1)
    
    StringFromPointer = tmp
 
End Function

