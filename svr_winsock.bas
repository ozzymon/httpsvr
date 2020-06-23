Attribute VB_Name = "svr_winsock"
Option Explicit

'winsock minimum version (Miner,Major)
Public Const WS_VERSION_REQD As Long = &H101

'winsock error return
Public Const SOCKET_ERROR As Long = -1

'Address Family
Public Const AF_UNSPEC      As Long = 0
Public Const AF_INET        As Long = 2
Public Const AF_IPX         As Long = 6
Public Const AF_APPLETALK   As Long = 16
Public Const AF_NETBIOS     As Long = 17
Public Const AF_INET6       As Long = 23
Public Const AF_IRDA        As Long = 26
Public Const AF_BTH         As Long = 32

'socket type
Public Const SOCK_STREAM    As Long = 1
Public Const SOCK_DGRAM     As Long = 2
Public Const SOCK_RAW       As Long = 3
Public Const SOCK_RDM       As Long = 4
Public Const SOCK_SEQPACKET As Long = 5

'Protocol
Public Const IPPROTO_IP         As Long = 0
Public Const IPPROTO_ICMP       As Long = 1
Public Const IPPROTO_IGMP       As Long = 2
Public Const BTHPROTO_RFCOMM    As Long = 3
Public Const IPPROTO_TCP        As Long = 6
Public Const IPPROTO_UDP        As Long = 17
Public Const IPPROTO_ICMPV6     As Long = 58
Public Const IPPROTO_RM         As Long = 113

'shutdown type
Public Const SD_RECEIVE         As Integer = 0
Public Const SD_SEND            As Integer = 1
Public Const SD_BOTH            As Integer = 2

Public Const WSADESCRIPTION_LEN     As Integer = 256
Public Const WSASYS_STATUS_LEN      As Integer = 128

'level 4 setsockopt
Public Const SOL_SOCKET         As Long = 0

'optname 4 setsockopt SOL_SOCKET
Public Const SO_DEBUG           As Long = &H1
Public Const SO_ACCEPTCONN      As Long = &H2
Public Const SO_REUSEADDR       As Long = &H4
Public Const SO_KEEPALIVE       As Long = &H8
Public Const SO_DONTROUTE       As Long = &H10
Public Const SO_BROADCAST       As Long = &H20
Public Const SO_USELOOPBACK     As Long = &H40
Public Const SO_LINGER          As Long = &H80
Public Const SO_OOBINLINE       As Long = &H100
Public Const SO_SNDBUF          As Long = &H1001
Public Const SO_RCVBUF          As Long = &H1002
Public Const SO_SNDLOWAT        As Long = &H1003
Public Const SO_RCVLOWAT        As Long = &H1004
Public Const SO_SNDTIMEO        As Long = &H1005
Public Const SO_RCVTIMEO        As Long = &H1006
Public Const SO_ERROR           As Long = &H1007
Public Const SO_TYPE            As Long = &H1008
Public Const SO_BSP_STATE       As Long = &H1009

Public Const SO_RANDOMIZE_PORT          As Long = &H3005
Public Const SO_PORT_SCALABILITY        As Long = &H3006
Public Const SO_REUSE_UNICASTPORT       As Long = &H3007
Public Const SO_REUSE_MULTICASTPORT     As Long = &H3008

Public Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription(WSADESCRIPTION_LEN + 1) As Byte
    szSystemStatus(WSASYS_STATUS_LEN + 1) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public Type hostent
     h_name As LongPtr          'pointer to hostname string
     h_aliases As LongPtr
     h_addrtype As Long         'address type
     h_length As Long           'length of each address
     h_addr_list As LongPtr     'list of addresses (null end)
End Type

'address storage
Public Type sockaddr
    sa_family As Integer
    sa_data(14) As Byte
End Type

'IPv4 address
Public Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero1 As Long
    sin_zero2 As Long
End Type

'---ioctl Constants
Public Const FIONREAD As Long = &H8004667F
Public Const FIONBIO  As Long = &H8004667E
Public Const FIOASYNC As Long = &H8004667D

'-------------------------------------------
' for Server:
'-------------------------------------------
Public Const FD_SETSIZE = 64
'Public Const FIONBIO = 2147772030#
Public Const SOCKADDR_SIZE = 16
Public Const SOCKADDR_IN_SIZE = 16
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Public Const IP_SUCCESS             As Long = 0
Public Const IP_ADD_MEMBERSHIP  As Long = 12
Public Const IP_DROP_MEMBERSHIP As Long = 13
'---network events
Public Const FD_READ                     As Long = &H1&
Public Const FD_WRITE                    As Long = &H2&
Public Const FD_OOB                      As Long = &H4&
Public Const FD_ACCEPT                   As Long = &H8&
Public Const FD_CONNECT                  As Long = &H10&
Public Const FD_CLOSE                    As Long = &H20&
Public Const FD_QOS                      As Long = &H40&
Public Const FD_GROUP_QOS                As Long = &H80&
Public Const FD_ROUTING_INTERFACE_CHANGE As Long = &H100&
Public Const FD_ADDRESS_LIST_CHANGE      As Long = &H200&

Public Const MSG_OOB = &H1                  'process out-of-band data
Public Const MSG_PEEK = &H2                 'peek at incoming message
Public Const MSG_DONTROUTE = &H4            'send without using routing tables
Public Const MSG_WAITALL = &H8              'do not complete until packet is completely filled

Public Const FD_MAX_EVENTS As Integer = 10

Public Type LPWSANETWORKEVENTS
    lNetworkEvents As Long
    iErrorCode(FD_MAX_EVENTS) As Long
End Type

'typedef struct fd_set {
'        u_int fd_count;               /* how many are SET? */
'        SOCKET  fd_array[FD_SETSIZE];   /* an array of SOCKETs */
'} fd_set;

Public Type FD_SET
    fd_count As Long
    fd_array(FD_SETSIZE - 1) As LongPtr
End Type

Public Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

Public Type ip_mreq
     imr_multiaddr As Long
     imr_interface As Long
End Type

Private Q As Boolean

Private Const TOUT = 1
Private Const Port = 8080
Private Const BACKLOG = 20

Private Const RecvSize = 1024
Private Const SendSize = 2048
Private Const BufSize = FD_SETSIZE - 1

Private Type RingBuf
    Fd As LongPtr
    FlgIn As Boolean
    FlgOut As Boolean
    HeadTerm As Integer
    RecvTerm As Integer
    Cntr As Single
    Addr As sockaddr_in
    Recvbyte(RecvSize) As Byte
    Sendbyte(SendSize) As Byte
End Type

'--------------------------------------------------------------------
'- Windows API
'- char     :Byte
'- Int      :Long
'- short    :Integer
'- long     :Long
'- pointer  :LongPtr
'- WORD     :Integer
'- DWORD    :Long
'WSA
Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSAData) As Long
Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
Private Declare PtrSafe Function WSAFDIsSet Lib "ws2_32.dll" Alias "__WSAFDIsSet" (ByVal SOCKET As LongPtr, ByRef Fds As FD_SET) As Long

'Connection
Private Declare PtrSafe Function w_socket Lib "ws2_32.dll" Alias "socket" (ByVal lngAf As Long, ByVal lngType As Long, ByVal lngProtocol As Long) As LongPtr
Private Declare PtrSafe Function w_connect Lib "ws2_32.dll" Alias "connect" (ByVal SOCKET As LongPtr, Name As sockaddr_in, ByVal namelen As Long) As Long
Private Declare PtrSafe Function w_shutdown Lib "ws2_32.dll" Alias "shutdown" (ByVal SOCKET As LongPtr, ByVal how As Long) As Long
Private Declare PtrSafe Function w_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal SOCKET As LongPtr) As Long
Private Declare PtrSafe Function w_select Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, ByVal readFdsptr As LongPtr, ByVal writeFdsptr As LongPtr, ByVal exceptFdsptr As LongPtr, TIMEOUT As timeval) As Long
Private Declare PtrSafe Function w_setsockopt Lib "ws2_32.dll" Alias "setsockopt" (ByVal SOCKET As LongPtr, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare PtrSafe Function w_ioctlsocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal SOCKET As LongPtr, ByVal cmd As Long, argp As Long) As Long

Private Declare PtrSafe Function w_getsockopt Lib "ws2_32.dll" Alias "getsockopt" (ByVal SOCKET As LongPtr, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long

'Transmission
Private Declare PtrSafe Function w_send Lib "ws2_32.dll" Alias "send" (ByVal SOCKET As LongPtr, buf As Any, ByVal Length As Long, ByVal Flags As Long) As Long
Private Declare PtrSafe Function w_sendTo Lib "ws2_32.dll" Alias "sendto" (ByVal SOCKET As LongPtr, buf As Any, ByVal Length As Long, ByVal Flags As Long, remoteAddr As sockaddr_in, ByVal remoteAddrSize As Long) As Long
Private Declare PtrSafe Function w_recv Lib "ws2_32.dll" Alias "recv" (ByVal SOCKET As LongPtr, buf As Any, ByVal Length As Long, ByVal Flags As Long) As Long
Private Declare PtrSafe Function w_recvFrom Lib "ws2_32.dll" Alias "recvfrom" (ByVal SOCKET As LongPtr, buf As Any, ByVal Length As Long, ByVal Flags As Long, fromAddr As sockaddr_in, fromAddrSize As Long) As Long

'svr func
Private Declare PtrSafe Function w_bind Lib "ws2_32.dll" Alias "bind" (ByVal SOCKET As LongPtr, Name As sockaddr, ByVal namelen As Long) As Long
Private Declare PtrSafe Function w_listen Lib "ws2_32.dll" Alias "listen" (ByVal SOCKET As LongPtr, ByVal BACKLOG As Long) As Long
Private Declare PtrSafe Function w_accept Lib "ws2_32.dll" Alias "accept" (ByVal SOCKET As LongPtr, ByRef Addr As sockaddr, ByRef addrlen As Long) As LongPtr

'Utility
Private Declare PtrSafe Function getsockname Lib "ws2_32.dll" (ByVal SOCKET As LongPtr, ByRef Name As sockaddr, ByRef namelen As Long) As Long
'hostnm local
Private Declare PtrSafe Function gethostname Lib "ws2_32.dll" (ByVal host_name As String, ByVal namelen As Integer) As Integer
'addr 2 hostnm
Private Declare PtrSafe Function gethostbyaddr Lib "ws2_32.dll" (ByRef Addr As Long, ByVal Length As Long, ByVal af As Long) As LongPtr
'hostnm 2 addr
Private Declare PtrSafe Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As LongPtr
'addr string to byte
Private Declare PtrSafe Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
'byte order host 2 nw
Private Declare PtrSafe Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Private Declare PtrSafe Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
'byte order nw 2 host
Private Declare PtrSafe Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Private Declare PtrSafe Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer

Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal hpvDest As LongPtr, ByVal hpvSource As LongPtr, ByVal cbCopy As Long)
Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal Buffer As String, ByRef size As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

'error code
Private Const WSABASEERR             As Long = 10000 'No Error
Private Const WSAEINTR               As Long = 10004 'Interrupted by system call
Private Const WSAEBADF               As Long = 10009 'Invalid file handle
Private Const WSAEACCES              As Long = 10013 'access denied
Private Const WSAEFAULT              As Long = 10014 'Invalid buff ptr
Private Const WSAEINVAL              As Long = 10022 'Invalid param
Private Const WSAEMFILE              As Long = 10024 'Too many open files
Private Const WSAEWOULDBLOCK         As Long = 10035 'Operation would block
Private Const WSAEINPROGRESS         As Long = 10036 'Operation now in progress
Private Const WSAEALREADY            As Long = 10037 'Operation already in progress
Private Const WSAENOTSOCK            As Long = 10038 'Socket operation on non-socket
Private Const WSAEDESTADDRREQ        As Long = 10039 '
Private Const WSAEMSGSIZE            As Long = 10040
Private Const WSAEPROTOTYPE          As Long = 10041
Private Const WSAENOPROTOOPT         As Long = 10042
Private Const WSAEPROTONOSUPPORT     As Long = 10043
Private Const WSAESOCKTNOSUPPORT     As Long = 10044
Private Const WSAEOPNOTSUPP          As Long = 10045
Private Const WSAEPFNOSUPPORT        As Long = 10046
Private Const WSAEAFNOSUPPORT        As Long = 10047
Private Const WSAEADDRINUSE          As Long = 10048
Private Const WSAEADDRNOTAVAIL       As Long = 10049
Private Const WSAENETDOWN            As Long = 10050
Private Const WSAENETUNREACH         As Long = 10051
Private Const WSAENETRESET           As Long = 10052
Private Const WSAECONNABORTED        As Long = 10053
Private Const WSAECONNRESET          As Long = 10054
Private Const WSAENOBUFS             As Long = 10055
Private Const WSAEISCONN             As Long = 10056
Private Const WSAENOTCONN            As Long = 10057
Private Const WSAESHUTDOWN           As Long = 10058
Private Const WSAETOOMANYREFS        As Long = 10059
Private Const WSAETIMEDOUT           As Long = 10060
Private Const WSAECONNREFUSED        As Long = 10061
Private Const WSAELOOP               As Long = 10062
Private Const WSAENAMETOOLONG        As Long = 10063
Private Const WSAEHOSTDOWN           As Long = 10064
Private Const WSAEHOSTUNREACH        As Long = 10065
Private Const WSAENOTEMPTY           As Long = 10066
Private Const WSAEPROCLIM            As Long = 10067
Private Const WSAEUSERS              As Long = 10068
Private Const WSAEDQUOT              As Long = 10069
Private Const WSAESTALE              As Long = 10070
Private Const WSAEREMOTE             As Long = 10071
Private Const WSASYSNOTREADY         As Long = 10091
Private Const WSAVERNOTSUPPORTED     As Long = 10092
Private Const WSANOTINITIALISED      As Long = 10093
Private Const WSAEDISCON             As Long = 10101
Private Const WSAENOMORE             As Long = 10102
Private Const WSAECANCELLED          As Long = 10103
Private Const WSAEINVALIDPROCTABLE   As Long = 10104
Private Const WSAEINVALIDPROVIDER    As Long = 10105
Private Const WSAEPROVIDERFAILEDINIT As Long = 10106
Private Const WSASYSCALLFAILURE      As Long = 10107
Private Const WSASERVICE_NOT_FOUND   As Long = 10108
Private Const WSATYPE_NOT_FOUND      As Long = 10109
Private Const WSA_E_NO_MORE          As Long = 10110
Private Const WSA_E_CANCELLED        As Long = 10111
Private Const WSAEREFUSED            As Long = 10112
Private Const WSAHOST_NOT_FOUND      As Long = 11001
Private Const WSATRY_AGAIN           As Long = 11002
Private Const WSANO_RECOVERY         As Long = 11003
Private Const WSANO_DATA             As Long = 11004
Private Const WSANO_ADDRESS          As Long = 11004
Private Const sckInvalidOp           As Long = 40020

'UTF8
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32.dll" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long) As Long

Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32.dll" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cchMultiByte As Long, _
    ByVal lpDefaultChar As LongPtr, _
    ByVal lpUsedDefaultChar As Long) As Long

Private Const CP_UTF8 As Long = 65001

'--------------------------------------------------
' Utilities
'--------------------------------------------------
Private Function IPToText(ByVal IPAddress As Long) As String
    Dim bytes(3) As Byte
    MoveMemory VarPtr(bytes(0)), VarPtr(IPAddress), 4
    IPToText = _
        CStr(bytes(0)) & "." & _
        CStr(bytes(1)) & "." & _
        CStr(bytes(2)) & "." & _
        CStr(bytes(3))
End Function

'errmsg
Private Function ErrorMsg(ByVal ErrorCode) As String
    ErrorMsg = ""
    Select Case ErrorCode
        Case WSAEINTR:           ErrorMsg = "WSAEINTR"
        Case WSAEACCES:          ErrorMsg = "WSAEACCES"
        Case WSAEFAULT:          ErrorMsg = "WSAEFAULT"
        Case WSAEINVAL:          ErrorMsg = "WSAEINVAL"
        Case WSAEMFILE:          ErrorMsg = "WSAEMFILE"
        Case WSAEWOULDBLOCK:     ErrorMsg = "WSAEWOULDBLOCK"
        Case WSAEINPROGRESS:     ErrorMsg = "WSAEINPROGRESS"
        Case WSAEALREADY:        ErrorMsg = "WSAEALREADY"
        Case WSAENOTSOCK:        ErrorMsg = "WSAENOTSOCK"
        Case WSAEDESTADDRREQ:    ErrorMsg = "WSAEDESTADDRREQ"
        Case WSAEMSGSIZE:        ErrorMsg = "WSAEMSGSIZE"
        Case WSAEPROTOTYPE:      ErrorMsg = "WSAEPROTOTYPE"
        Case WSAENOPROTOOPT:     ErrorMsg = "WSAENOPROTOOPT"
        Case WSAEPROTONOSUPPORT: ErrorMsg = "WSAEPROTONOSUPPORT"
        Case WSAESOCKTNOSUPPORT: ErrorMsg = "WSAESOCKTNOSUPPORT"
        Case WSAEOPNOTSUPP:      ErrorMsg = "WSAEOPNOTSUPP"
        Case WSAEPFNOSUPPORT:    ErrorMsg = "WSAEPFNOSUPPORT"
        Case WSAEAFNOSUPPORT:    ErrorMsg = "WSAEAFNOSUPPORT"
        Case WSAEADDRINUSE:      ErrorMsg = "WSAEADDRINUSE"
        Case WSAEADDRNOTAVAIL:   ErrorMsg = "WSAEADDRNOTAVAIL"
        Case WSAENETDOWN:        ErrorMsg = "WSAENETDOWN"
        Case WSAENETUNREACH:     ErrorMsg = "WSAENETUNREACH"
        Case WSAENETRESET:       ErrorMsg = "WSAENETRESET"
        Case WSAECONNABORTED:    ErrorMsg = "WSAECONNABORTED"
        Case WSAECONNRESET:      ErrorMsg = "WSAECONNRESET"
        Case WSAENOBUFS:         ErrorMsg = "WSAENOBUFS"
        Case WSAEISCONN:         ErrorMsg = "WSAEISCONN"
        Case WSAENOTCONN:        ErrorMsg = "WSAENOTCONN"
        Case WSAESHUTDOWN:       ErrorMsg = "WSAESHUTDOWN"
        Case WSAETIMEDOUT:       ErrorMsg = "WSAETIMEDOUT"
        Case WSAECONNREFUSED:    ErrorMsg = "WSAECONNREFUSED"
        Case WSAEHOSTDOWN:       ErrorMsg = "WSAEHOSTDOWN"
        Case WSAEHOSTUNREACH:    ErrorMsg = "WSAEHOSTUNREACH"
        Case WSAEPROCLIM:        ErrorMsg = "WSAEPROCLIM"
        Case WSASYSNOTREADY:     ErrorMsg = "WSASYSNOTREADY"
        Case WSAVERNOTSUPPORTED: ErrorMsg = "WSAVERNOTSUPPORTED"
        Case WSANOTINITIALISED:  ErrorMsg = "WSANOTINITIALISED"
        Case WSAEDISCON:         ErrorMsg = "WSAEDISCON"
        Case WSATYPE_NOT_FOUND:  ErrorMsg = "WSATYPE_NOT_FOUND"
        Case WSAHOST_NOT_FOUND:  ErrorMsg = "WSAHOST_NOT_FOUND"
        Case WSATRY_AGAIN:       ErrorMsg = "WSATRY_AGAIN"
        Case WSANO_RECOVERY:     ErrorMsg = "WSANO_RECOVERY"
        Case WSANO_DATA:         ErrorMsg = "WSANO_DATA"
        Case sckInvalidOp:       ErrorMsg = "sckInvalidOp"
        Case Else:               ErrorMsg = "UNKNOWN"
    End Select
    ErrorMsg = str$(ErrorCode) & " " & ErrorMsg
End Function

'--------------------------------------------------
' UTF8
'--------------------------------------------------

'Byte 2 String
Private Function FromUTF8(ByRef bData() As Byte) As String
    Dim nBufferSize As Long

    If UBound(bData) = -1 Then
        Exit Function
    End If
    
    nBufferSize = MultiByteToWideChar(CP_UTF8, 0, VarPtr(bData(LBound(bData))), UBound(bData) - LBound(bData) + 1, 0, 0)
    
    If bData(UBound(bData)) = 0 Then
        nBufferSize = nBufferSize - 1
    End If
    
    FromUTF8 = String(nBufferSize, vbNullChar)
    
    MultiByteToWideChar CP_UTF8, 0, VarPtr(bData(LBound(bData))), UBound(bData) - LBound(bData) + 1, StrPtr(FromUTF8), nBufferSize
End Function

'String 2 Byte
Private Function ToUTF8(ByRef sData As String) As Byte()
    Dim nBufferSize As Long
    Dim ret() As Byte
    
    If Len(sData) = 0 Then
        ToUTF8 = ""
        Exit Function
    End If
    
    nBufferSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sData), Len(sData), 0, 0, 0, 0)
    ReDim ret(0 To nBufferSize - 1)
    
    WideCharToMultiByte CP_UTF8, 0, StrPtr(sData), Len(sData), VarPtr(ret(0)), nBufferSize, 0, 0
    
    ToUTF8 = ret
End Function

'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
' httpsvr
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------

Public Sub wsTCPListenMpx()
    Dim buf(BufSize) As RingBuf
    Dim Fds As FD_SET
    Dim WSAD As WSAData
    Dim TV As timeval
    Dim ss As LongPtr, sc As LongPtr
    Dim HostAddr As sockaddr, ClientAddr As sockaddr
    Dim HostAddr_in As sockaddr_in
    Dim Addrptr As LongPtr
    Dim flg As Long, ret As Long
    Dim yes As Boolean, NA As Boolean
    Dim i As Integer, j As Integer, cnt As Integer
    Dim wait As Single
    
    Const NULLPTR As LongPtr = 0
        
    'winsock init
    If (WSAStartup(WS_VERSION_REQD, WSAD)) Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If

    'get the socket 4 accept
    ss = w_socket(AF_INET, SOCK_STREAM, IPPROTO_IP)
    If ss < 0 Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If
    
    'setting socket option SO_REUSEADDR
    yes = True
    If w_setsockopt(ss, SOL_SOCKET, SO_REUSEADDR, yes, Len(yes)) = SOCKET_ERROR Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If
    
    'non-block
    flg = 1
    If w_ioctlsocket(ss, FIONBIO, flg) = SOCKET_ERROR Then
        WSACleanup
        Exit Sub
    End If

    'struct 4 accept
    HostAddr_in.sin_family = AF_INET
    HostAddr_in.sin_port = htons(Port)
    HostAddr_in.sin_addr = inet_addr("0.0.0.0") 'INADDR_ANY
    Addrptr = VarPtr(HostAddr_in)
    MoveMemory VarPtr(HostAddr), ByVal Addrptr, SOCKADDR_SIZE

    If w_bind(ss, HostAddr, SOCKADDR_SIZE) = SOCKET_ERROR Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If

    If w_listen(ss, BACKLOG) = SOCKET_ERROR Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If
    
    Debug.Print "wsTCPListenMpx:" & Now()
    Q = False
    TV.tv_sec = 0
    TV.tv_usec = 0
    'buf(0) -> sock 4 accept
    buf(0).Fd = ss
    Debug.Print "opened:" & ss
    'socket 4 transmission init
    For i = 1 To BufSize
        ClearBuf buf(i)
    Next i
    
    Do
        'clean up timeout fd(s)
        wait = Timer
        For i = 1 To BufSize
            If buf(i).Fd <> -1 And wait > (buf(i).Cntr + TOUT) Then
                Debug.Print "clean up:" & buf(i).Fd
                If (w_closesocket(buf(i).Fd) = SOCKET_ERROR) Then
                    Debug.Print ErrorMsg(WSAGetLastError)
                End If
                ClearBuf buf(i)
            End If
        Next i
        'select
        'Original one is macro and make new in VBA
        vbaFD_ZERO Fds

        For i = 0 To BufSize
            If buf(i).Fd <> -1 Then vbaFD_SET buf(i).Fd, Fds
        Next i
        
        cnt = Fds.fd_count
        
        If Q Then Exit Do
        
        If w_select(0, VarPtr(Fds), NULLPTR, NULLPTR, TV) <> 0 Then
            If cnt > 1 Then
                'clean up disconnect fd(s)
                'recieve activated fd(s)
                For i = 1 To BufSize
                    If buf(i).Fd <> -1 And vbaFD_ISSET(buf(i).Fd, Fds) <> 0 Then
                        ret = w_recv(buf(i).Fd, buf(i).Recvbyte(buf(i).RecvTerm), RecvSize - buf(i).RecvTerm, 0)
                        If ret = 0 Then
                            Debug.Print "disconnected:" & buf(i).Fd
                            If (w_closesocket(buf(i).Fd) = SOCKET_ERROR) Then
                                Debug.Print ErrorMsg(WSAGetLastError)
                            End If
                            ClearBuf buf(i)
                        ElseIf ret > 0 Then
                            buf(i).RecvTerm = buf(i).RecvTerm + ret
                            Debug.Print "received:" & buf(i).Fd
                        Else
                            Debug.Print "receive_error:" & buf(i).Fd
                        End If
                    End If
                Next i
                'check buffered fd(s)
                For i = 1 To BufSize
                    If buf(i).Fd <> -1 And (CrLf2(buf(i).Recvbyte, buf(i).HeadTerm)) And _
                            (ContentLength(buf(i).Recvbyte, j, buf(i).HeadTerm)) Then
                        buf(i).RecvTerm = buf(i).HeadTerm + j
                        buf(i).FlgIn = True
                        Debug.Print "checked:" & buf(i).Fd
                    End If
                Next i
                'clean up full-filled fd(s)
                For i = 1 To BufSize
                    If buf(i).Fd <> -1 And buf(i).RecvTerm = RecvSize And buf(i).FlgIn = False Then
                        Debug.Print "full buffered:" & buf(i).Fd
                        If (w_closesocket(buf(i).Fd) = SOCKET_ERROR) Then
                            Debug.Print ErrorMsg(WSAGetLastError)
                        End If
                        ClearBuf buf(i)
                    End If
                Next i
            End If
            'accept
            If vbaFD_ISSET(ss, Fds) <> 0 Then
                j = 1
                Do
                    NA = True
                    For i = j To BufSize
                        If buf(i).Fd = -1 Then
                            NA = False
                            Exit For
                        End If
                    Next i
                    If NA = True Then
                        Exit Do
                    Else
                        j = i
                    End If
                    
                    sc = w_accept(ss, ClientAddr, SOCKADDR_SIZE)

                    If sc = -1 Then
                        Exit Do
                    Else
                        buf(i).Fd = sc
                        buf(i).Cntr = Timer
                        MoveMemory VarPtr(buf(i).Addr), VarPtr(ClientAddr), SOCKADDR_SIZE
                        Debug.Print "accepted:" & buf(i).Fd
                    End If
                Loop
            End If
        End If

        'http io
        For i = 1 To BufSize
            If buf(i).FlgIn = True Then
                Process buf(i).HeadTerm, buf(i).RecvTerm, buf(i).Recvbyte, buf(i).Sendbyte, buf(i).Addr
                buf(i).FlgOut = True
            End If
        Next i
        
        'send
        For i = 1 To BufSize
            If buf(i).FlgOut = True Then
                SendRes buf(i).Fd, buf(i).Sendbyte
                
                If (w_closesocket(buf(i).Fd) = SOCKET_ERROR) Then
                    Debug.Print ErrorMsg(WSAGetLastError)
                End If
                
                Debug.Print "sent:" & buf(i).Fd
                ClearBuf buf(i)
            End If
        Next i
        
    Loop
    
    i = 0
    Do
        If Fds.fd_array(i) <> -1 Then
            If (w_closesocket(Fds.fd_array(i)) = SOCKET_ERROR) Then
                Debug.Print ErrorMsg(WSAGetLastError)
            Else
                Debug.Print "closed:" & Fds.fd_array(i)
            End If
            cnt = cnt - 1
            i = i + 1
        End If
    Loop Until (cnt = 0 Or i = BufSize)
    
    If WSACleanup() = SOCKET_ERROR Then
        MsgBox ErrorMsg(WSAGetLastError)
    End If

End Sub

Public Sub wsTCPListen()
    Dim buf As RingBuf
    Dim WSAD As WSAData
    Dim sockfd As LongPtr 'socket
    Dim HostAddr As sockaddr, ClientAddr As sockaddr
    Dim HostAddr_in As sockaddr_in
    Dim Addrptr As LongPtr
    Dim ret As Long, flg As Long
    Dim yes As Boolean, err As Boolean
    Dim i As Integer, j As Integer
    Dim wait As Single
        
    'winsock init
    If (WSAStartup(WS_VERSION_REQD, WSAD)) Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If

    'get the socket 4 accept
    sockfd = w_socket(AF_INET, SOCK_STREAM, IPPROTO_IP)
    If sockfd < 0 Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If
    
    'setting socket option SO_REUSEADDR
    yes = True
    If w_setsockopt(sockfd, SOL_SOCKET, SO_REUSEADDR, yes, Len(yes)) = SOCKET_ERROR Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If
    
    'struct 4 accept
    HostAddr_in.sin_family = AF_INET
    HostAddr_in.sin_port = htons(Port)
    HostAddr_in.sin_addr = inet_addr("0.0.0.0") 'INADDR_ANY
    Addrptr = VarPtr(HostAddr_in)
    MoveMemory VarPtr(HostAddr), ByVal Addrptr, SOCKADDR_SIZE

    If w_bind(sockfd, HostAddr, SOCKADDR_SIZE) = SOCKET_ERROR Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If

    If w_listen(sockfd, BACKLOG) = SOCKET_ERROR Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If
    
    Debug.Print "wsTCPListen:" & Now()
    Q = False
    
    Do
        err = False
        ClearBuf buf

        'socket 4 transmission
        buf.Fd = w_accept(sockfd, ClientAddr, SOCKADDR_SIZE)
        MoveMemory VarPtr(buf.Addr), VarPtr(ClientAddr), SOCKADDR_SIZE
        If buf.Fd = -1 Then
            Debug.Print ErrorMsg(WSAGetLastError)
            Exit Do
        End If
        
        'non-block
        'block_io don't work if the connection is idle.
        flg = 1
        w_ioctlsocket buf.Fd, FIONBIO, flg

        DoEvents

        wait = Timer
        Do
            ret = w_recv(buf.Fd, buf.Recvbyte(0), RecvSize, 0)
            If Timer > wait + TOUT Then Exit Do
            DoEvents
        Loop While (ret = SOCKET_ERROR)

        If ret = SOCKET_ERROR Then
            Debug.Print "w_recv" & ErrorMsg(WSAGetLastError)
            err = True
        ElseIf ret = 0 Then
            'disconnect
            Debug.Print "w_recv disconnect"
            err = True
        Else
            'number of used bytes
            i = CInt(ret)

            If (CrLf2(buf.Recvbyte, buf.HeadTerm)) And (ContentLength(buf.Recvbyte, j, buf.HeadTerm)) Then
                err = False
                buf.RecvTerm = j + buf.HeadTerm
            Else
                DoEvents
                'loop
                Do
                    ret = w_recv(buf.Fd, buf.Recvbyte(i), RecvSize - i, 0)
                    If ret > 0 Then
                        i = i + ret
                    Else
                        Exit Do
                    End If
                Loop

                If buf.HeadTerm = 0 And (Not CrLf2(buf.Recvbyte, buf.HeadTerm)) Then
                    err = True
                ElseIf Not (ContentLength(buf.Recvbyte, j, buf.HeadTerm)) Then
                    err = True
                Else
                    err = False
                    buf.RecvTerm = j + buf.HeadTerm
                End If
            End If

        End If
        
        If err = False Then
            'Process
            Process buf.HeadTerm, buf.RecvTerm, buf.Recvbyte, buf.Sendbyte, buf.Addr
            'Transmission
            SendRes buf.Fd, buf.Sendbyte
        End If
        
        If (w_closesocket(buf.Fd) = SOCKET_ERROR) Then
            Debug.Print ErrorMsg(WSAGetLastError)
        End If

    Loop While (Not Q)
       
    If (w_closesocket(sockfd) = SOCKET_ERROR) Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If
    
    If WSACleanup() = SOCKET_ERROR Then
        MsgBox ErrorMsg(WSAGetLastError)
    End If

End Sub

Public Sub wsTCPListenOld()
    Dim WSAD As WSAData
    Dim sockfd As LongPtr, Fd As LongPtr 'socket
    Dim HostAddr As sockaddr, ClientAddr As sockaddr
    Dim HostAddr_in As sockaddr_in, ClientAddr_in As sockaddr_in
    Dim Addrptr As LongPtr
    Dim ret As Long, flg As Long
    Dim yes As Boolean, err As Boolean
    Dim i As Integer, j As Integer
    Dim wait As Single
    Dim Recvbyte(RecvSize) As Byte, Sendbyte(SendSize) As Byte
    Dim HeadTerm As Integer, Length As Integer

        
    'winsock init
    If (WSAStartup(WS_VERSION_REQD, WSAD)) Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If

    'get the socket 4 accept
    sockfd = w_socket(AF_INET, SOCK_STREAM, IPPROTO_IP)
    If sockfd < 0 Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If
    
    'setting socket option SO_REUSEADDR
    yes = True
    If w_setsockopt(sockfd, SOL_SOCKET, SO_REUSEADDR, yes, Len(yes)) = SOCKET_ERROR Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If
    
    'struct 4 accept
    HostAddr_in.sin_family = AF_INET
    HostAddr_in.sin_port = htons(Port)
    HostAddr_in.sin_addr = inet_addr("0.0.0.0") 'INADDR_ANY
    Addrptr = VarPtr(HostAddr_in)
    MoveMemory VarPtr(HostAddr), ByVal Addrptr, SOCKADDR_SIZE

    If w_bind(sockfd, HostAddr, SOCKADDR_SIZE) = SOCKET_ERROR Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If

    If w_listen(sockfd, BACKLOG) = SOCKET_ERROR Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If
    
    Debug.Print "wsTCPListenOld:" & Now()
    Q = False
    
    Do
        err = True
        Erase Recvbyte
        HeadTerm = 0
        Length = 0

        'socket 4 transmission
        Fd = w_accept(sockfd, ClientAddr, SOCKADDR_SIZE)
        MoveMemory VarPtr(ClientAddr_in), VarPtr(ClientAddr), SOCKADDR_SIZE
        If Fd = -1 Then
            Debug.Print ErrorMsg(WSAGetLastError)
            Exit Do
        End If
        
        'non-block
        flg = 1
        w_ioctlsocket Fd, FIONBIO, flg

        i = 0
        j = 1
        Do
            ret = w_recv(Fd, Recvbyte(i), RecvSize - i, 0)
            Select Case ret
                Case Is > 0
                    i = i + ret
                'disconnect
                Case 0
                    Exit Do
                Case SOCKET_ERROR
                    If j = 0 Then Exit Do
                    j = j - 1
                    DoEvents
            End Select
        Loop

        If (CrLf2(Recvbyte, HeadTerm)) And (ContentLength(Recvbyte, Length, HeadTerm)) Then
            err = False
        End If
        
        If err = False Then
            'Process
            Process HeadTerm, HeadTerm + Length, Recvbyte, Sendbyte, ClientAddr_in
            'Transmission
            SendRes Fd, Sendbyte
        End If
        
        If (w_closesocket(Fd) = SOCKET_ERROR) Then
            Debug.Print ErrorMsg(WSAGetLastError)
        End If

    Loop While (Not Q)
       
    If (w_closesocket(sockfd) = SOCKET_ERROR) Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If
    
    If WSACleanup() = SOCKET_ERROR Then
        MsgBox ErrorMsg(WSAGetLastError)
    End If

End Sub

'Sample of Process
Private Sub Process(ByVal HeadTerm As Integer, ByVal RecvTerm As Integer, Recvbyte() As Byte, Sendbyte() As Byte, Clntaddr As sockaddr_in)
    Dim i As Integer, j As Integer, k As Integer
    Dim Arg(RecvSize) As Byte, Res() As Byte
    Dim Recvstring As String, SendString As String, Argstring As String, Mthd As String, Resource As String
    
    'On Error GoTo Error_
    
    Recvstring = StrConv(Recvbyte, vbUnicode)
    If Len(Recvstring) < 4 Then
        Mthd = ""
    Else
        Mthd = Left(Recvstring, 4)
    End If
    
    Select Case Mthd
        Case "HEAD"
            SendString = "HTTP/1.0 200 OK" & vbCrLf & _
            "Server: VBA webserver" & vbCrLf & _
            vbCrLf
            
        Case "GET ", "POST"
            'GET (resource)...
            'return registration form
            'POST (resource)...
            'regist & return registration form
            'GET /l -> return lists
            'GET /q -> quit
            Argstring = ""
            If Mthd = "POST" Then
                'resource "POST(space)(resource)(space)HTTP..."
                i = InStr(6, Recvstring, Chr(32))
                If i > 6 Then
                    Resource = Mid(Recvstring, 6, i - 6)
                Else
                    Resource = ""
                End If
                'percent-decode
                j = 0
                Erase Arg
                
                For i = HeadTerm To (RecvTerm - 1)
                    If Recvbyte(i) = 37 And i <= (RecvTerm - 3) Then
                        Arg(j) = CByte("&H" & Chr(Recvbyte(i + 1)) & Chr(Recvbyte(i + 2)))
                        i = i + 2
                    Else
                        Arg(j) = Recvbyte(i)
                    End If
                    j = j + 1
                Next i
                Argstring = FromUTF8(Arg)
            ElseIf Mthd = "GET " Then
                 'resource "GET(space)(resource)(space)HTTP..."
                i = InStr(5, Recvstring, Chr(32))
                If i > 5 Then
                    Resource = Mid(Recvstring, 5, i - 5)
                Else
                    Resource = ""
                End If
            End If
            
            If Resource = "/l" Then
'                Sendstring = "HTTP/1.0 200 OK" & vbCrLf & _
'                "Server: VBA webserver" & vbCrLf & _
'                "Content-Type: application/octet-stream" & vbCrLf & _
'                "Content-Disposition: attachment; filename=OJISAN.csv" & vbCrLf & _
'                "Content-Transfer-Encoding: binary" & _
'                vbCrLf & vbCrLf & _
'                Csv
                SendString = "HTTP/1.0 200 OK" & vbCrLf & _
                "Server: VBA webserver" & vbCrLf & _
                "Content-Type: text/plain; charset=utf-8" & vbCrLf & vbCrLf & _
                Csv
            ElseIf Resource = "/q" Then
                Q = True
                SendString = "HTTP/1.0 200 OK" & vbCrLf & _
                "Server: VBA webserver" & vbCrLf & _
                "Content-Type: text/plain; charset=utf-8" & vbCrLf & _
                vbCrLf & _
                "QUIT"
            ElseIf Resource Like "/[!a-z]*" Then
                SendString = "HTTP/1.0 404 Not Found" & vbCrLf & _
                "Server: VBA webserver" & vbCrLf & _
                vbCrLf
            Else
                SendString = "HTTP/1.0 200 OK" & vbCrLf & _
                "Server: VBA webserver" & vbCrLf & _
                "Content-Type: text/html; charset=utf-8" & vbCrLf & vbCrLf & _
                "<!DOCTYPE html>" & _
                "<html>" & _
                "<head>" & _
                "<meta charset=""UTF-8"">" & _
                "<title>OJISAN</title>" & _
                "</head>" & _
                "<body>" & _
                "<form method=""post"">" & _
                "<p><label for=""ttl"">TITLE:</label>" & _
                "<input type=""hidden"" name=""token"" value=""" & CStr(DateDiff("s", "1970/1/1 9:00", Now)) & """>" & _
                "<input type=""text"" size=""20"" name=""ttl"" required></p>" & _
                "<fieldset><legend>SELECT</legend>" & _
                "<p><label><input type=""radio"" name=""num"" value=""1"" checked>1</label></p>" & _
                "<p><label><input type=""radio"" name=""num"" value=""2"">2</label></p>" & _
                "</fieldset>" & _
                "<p><label for=""char"">CHR:</label><br>" & _
                "<textarea name=""char"" cols=""100"" rows=""3"" maxlength=""1000""></textarea></p>" & _
                "<p><input type=""submit"" value=""SUBMIT""></p>" & _
                "</form>" & _
                "<p>" & Replace(Recvstring, vbCrLf, "<br>") & Argstring & "<br>" & "</p>" & _
                "</body>" & _
                "</html>"
                
                If Argstring <> "" Then
                    Paste IPToText(Clntaddr.sin_addr), Resource, Argstring
                End If
            End If
            
        Case Else
            SendString = "HTTP/1.0 404 Not Found" & vbCrLf & _
            "Server: VBA webserver" & vbCrLf & _
            vbCrLf
    End Select
    
    Res = ToUTF8(SendString)
    Erase Sendbyte
    
    If UBound(Res) < SendSize Then
        k = UBound(Res) + 1
    Else
        k = SendSize
    End If
    
    MoveMemory VarPtr(Sendbyte(0)), VarPtr(Res(0)), k
    
Exit Sub

Error_:
    Erase Sendbyte

End Sub

Private Sub SendRes(Fd As LongPtr, Sendbyte() As Byte)
    Dim i  As Integer, j As Integer
    Dim ret As Long

    For j = SendSize To 0 Step -1
        If Sendbyte(j) <> 0 Then Exit For
    Next j
    
    i = 0
    j = j + 1

    Do
        ret = w_send(Fd, Sendbyte(i), CLng(j), 0)
        If ret > 0 Then
            j = j - ret
            i = i + ret
            DoEvents
        ElseIf ret = SOCKET_ERROR Then
            Debug.Print "w_send " & ErrorMsg(WSAGetLastError)
            Exit Do
        ElseIf ret = 0 Then
            Debug.Print "w_send disconnect"
            Exit Do
        End If
    Loop Until (j = 0)
    
End Sub

Private Function CrLf2(Recv() As Byte, HeadTerm As Integer) As Boolean
    Dim i As Integer, j As Integer
    Dim Crlfcrlf As Variant
    
    Crlfcrlf = Array(13, 10, 13, 10)

    j = 3
    
    For i = RecvSize To 0 Step -1
        If Recv(i) = Crlfcrlf(j) Then
            If j = 0 Then
                i = i + UBound(Crlfcrlf) + 1
                Exit For
            Else
                j = j - 1
            End If
        Else
            j = 3
        End If
    Next i
    
    HeadTerm = i
    
    If i > 0 Then
        CrLf2 = True
    Else
        CrLf2 = False
    End If

End Function

Private Function ContentLength(Recv() As Byte, Length As Integer, ByVal BodyStart As Integer) As Boolean
    Dim i As Integer
    Dim str As String
    
    Const CL As String = "Content-Length:"

    str = StrConv(Recv, vbUnicode)
    i = InStr(str, CL)
    
    If i <> 0 Then
        i = i + Len(CL)
        Length = Val(Mid(str, i + 1))
    Else
        Length = 0
        ContentLength = True
        Exit Function
    End If
    
    For i = RecvSize To BodyStart Step -1
        If Recv(i) <> 0 Then
            If i - BodyStart + 1 = Length Then
                ContentLength = True
            ElseIf i = RecvSize - 1 Then
                ContentLength = True
                Length = RecvSize - BodyStart
            Else
                ContentLength = False
            End If
            Exit For
        End If
    Next i

End Function

Private Sub ClearBuf(buf As RingBuf)

    buf.Fd = -1
    buf.FlgIn = False
    buf.FlgOut = False
    buf.HeadTerm = 0
    buf.RecvTerm = 0
    buf.Cntr = 0
    Erase buf.Recvbyte
    Erase buf.Sendbyte
    ZeroMemory buf.Addr, SOCKADDR_IN_SIZE

End Sub

Private Sub Paste(IP As String, Res As String, Arg As String)
    Dim LastRow As Integer
    
    'escape
    LastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
    ActiveSheet.Cells(LastRow + 1, 1).Value = Now()
    ActiveSheet.Cells(LastRow + 1, 2).Value = IP
    ActiveSheet.Cells(LastRow + 1, 3).Value = Replace(Replace(Res, Chr(34), ""), Chr(44), "")
    ActiveSheet.Cells(LastRow + 1, 4).Value = Replace(Replace(Replace(Replace(Arg, vbCr, ""), vbLf, ""), Chr(34), ""), Chr(44), "")
    
End Sub
Private Function Csv() As String
    Dim Data As Variant
    Dim str As String
    Dim i As Integer
    Const COMMA As String = ","
    
    Data = ActiveSheet.UsedRange
    
    If (IsEmpty(Data)) Then
        Csv = "NO DATA"
    Else
        For i = UBound(Data) To 1 Step -1
             str = str & Format(Data(i, 1), "yyyy/mm/dd hh:mm:ss") & COMMA & _
             Data(i, 2) & COMMA & Data(i, 3) & COMMA & Data(i, 4) & vbCrLf
        Next i
        
        Csv = str
    End If

End Function

Private Sub vbaFD_SET(ByVal Fd As LongPtr, ByRef Fds As FD_SET)
'#define FD_SET(fd, set)
'do {
'    u_int i;
'    for (i = 0; i < set->fd_count; i++) {
'        if (set->fd_array[i] == (fd)) {
'            break;
'        }
'    }
'    if (i == set->fd_count) {
'        if (set->fd_count < FD_SETSIZE) {
'            set->fd_array[i] = (fd);
'            set->fd_count++;
'        }
'    }
'} while(0, 0)

'DESC
'nfds = Fds(0) + 1

    Dim i, j As Integer

    For i = 0 To (Fds.fd_count - 1)
        If Fds.fd_array(i) = Fd Then
            Exit Sub
        ElseIf Fds.fd_array(i) < Fd Then
            Exit For
        End If
    Next i
    
    For j = (Fds.fd_count - 1) To i Step -1
        Fds.fd_array(j + 1) = Fds.fd_array(j)
    Next j
    Fds.fd_array(i) = Fd
    Fds.fd_count = Fds.fd_count + 1

End Sub

Private Sub vbaFD_CLR(ByVal Fd As LongPtr, ByRef Fds As FD_SET)
'#define FD_CLR(fd, set)
'do {
'    u_int i;
'    for (i = 0; i < set->fd_count ; i++) {
'        if (set->fd_array[i] == fd) {
'            while (i < set->fd_count-1) {
'                set->fd_array[i] =
'                    set->fd_array[i+1];
'                i++;
'            }
'            set->fd_count--;
'            break;
'        }
'    }
'} while(0, 0)

    Dim i, j As Integer

    For i = 0 To (Fds.fd_count - 1)
        If Fds.fd_array(i) = Fd Then
            For j = i To ((Fds.fd_count - 1) - 1)
                Fds.fd_array(j) = Fds.fd_array(j + 1)
            Next j
            Fds.fd_array(Fds.fd_count - 1) = 0
            Fds.fd_count = Fds.fd_count - 1
            Exit For
        End If
    Next i

End Sub

Private Sub vbaFD_ZERO(ByRef Fds As FD_SET)
'#define FD_ZERO(set) (set->fd_count=0)

    Fds.fd_count = 0

End Sub

Private Function vbaFD_ISSET(ByVal Fd As LongPtr, ByRef Fds As FD_SET) As Long
'extern int PASCAL FAR __WSAFDIsSet(SOCKET fd, fd_set FAR *);

    vbaFD_ISSET = WSAFDIsSet(Fd, Fds)

End Function


