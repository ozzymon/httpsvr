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
     h_aliases As LongPtr       '
     h_addrtype As Long      'address type
     h_length As Long        'length of each address
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
Private Const PORT = 8080
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
'- 関数宣言
'- char     :Byte
'- Int      :Long
'- short    :Integer
'- long     :Long
'- pointer  :LongPtr
'- WORD     :Integer
'- DWORD    :Long
'WSA関連
Private Declare PtrSafe Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSAData) As Long
Private Declare PtrSafe Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare PtrSafe Function WSAGetLastError Lib "ws2_32.dll" () As Long
Private Declare PtrSafe Function WSAFDIsSet Lib "ws2_32.dll" Alias "__WSAFDIsSet" (ByVal SOCKET As LongPtr, ByRef Fds As FD_SET) As Long

'接続
Private Declare PtrSafe Function w_socket Lib "ws2_32.dll" Alias "socket" (ByVal lngAf As Long, ByVal lngType As Long, ByVal lngProtocol As Long) As LongPtr
Private Declare PtrSafe Function w_connect Lib "ws2_32.dll" Alias "connect" (ByVal SOCKET As LongPtr, Name As sockaddr_in, ByVal namelen As Long) As Long
Private Declare PtrSafe Function w_shutdown Lib "ws2_32.dll" Alias "shutdown" (ByVal SOCKET As LongPtr, ByVal how As Long) As Long
Private Declare PtrSafe Function w_closesocket Lib "ws2_32.dll" Alias "closesocket" (ByVal SOCKET As LongPtr) As Long
Private Declare PtrSafe Function w_select Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, ByVal readFdsptr As LongPtr, ByVal writeFdsptr As LongPtr, ByVal exceptFdsptr As LongPtr, TIMEOUT As timeval) As Long
Private Declare PtrSafe Function w_setsockopt Lib "ws2_32.dll" Alias "setsockopt" (ByVal SOCKET As LongPtr, ByVal level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Private Declare PtrSafe Function w_ioctlsocket Lib "ws2_32.dll" Alias "ioctlsocket" (ByVal SOCKET As LongPtr, ByVal cmd As Long, argp As Long) As Long

Private Declare PtrSafe Function w_getsockopt Lib "ws2_32.dll" Alias "getsockopt" (ByVal SOCKET As LongPtr, ByVal level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long

'送受信
Private Declare PtrSafe Function w_send Lib "ws2_32.dll" Alias "send" (ByVal SOCKET As LongPtr, Buf As Any, ByVal Length As Long, ByVal Flags As Long) As Long
Private Declare PtrSafe Function w_sendTo Lib "ws2_32.dll" Alias "sendto" (ByVal SOCKET As LongPtr, Buf As Any, ByVal Length As Long, ByVal Flags As Long, remoteAddr As sockaddr_in, ByVal remoteAddrSize As Long) As Long
Private Declare PtrSafe Function w_recv Lib "ws2_32.dll" Alias "recv" (ByVal SOCKET As LongPtr, Buf As Any, ByVal Length As Long, ByVal Flags As Long) As Long
Private Declare PtrSafe Function w_recvFrom Lib "ws2_32.dll" Alias "recvfrom" (ByVal SOCKET As LongPtr, Buf As Any, ByVal Length As Long, ByVal Flags As Long, fromAddr As sockaddr_in, fromAddrSize As Long) As Long

'サーバー
Private Declare PtrSafe Function w_bind Lib "ws2_32.dll" Alias "bind" (ByVal SOCKET As LongPtr, Name As sockaddr, ByVal namelen As Long) As Long
Private Declare PtrSafe Function w_listen Lib "ws2_32.dll" Alias "listen" (ByVal SOCKET As LongPtr, ByVal BACKLOG As Long) As Long
Private Declare PtrSafe Function w_accept Lib "ws2_32.dll" Alias "accept" (ByVal SOCKET As LongPtr, ByRef Addr As sockaddr, ByRef addrlen As Long) As LongPtr

'Utility
Private Declare PtrSafe Function getsockname Lib "ws2_32.dll" (ByVal SOCKET As LongPtr, ByRef Name As sockaddr, ByRef namelen As Long) As Long
'ローカルホスト名取得
Private Declare PtrSafe Function gethostname Lib "ws2_32.dll" (ByVal host_name As String, ByVal namelen As Integer) As Integer
'アドレスからホスト名を取得
Private Declare PtrSafe Function gethostbyaddr Lib "ws2_32.dll" (ByRef Addr As Long, ByVal Length As Long, ByVal af As Long) As LongPtr
'ホスト名からIPアドレスを取得
Private Declare PtrSafe Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As LongPtr
'IPをドット形式(x.x.x.x)から内部形式に変更 ※8進と16進に注意
Private Declare PtrSafe Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
'ホストバイトオーダからネットワークバイトオーダに変更
Private Declare PtrSafe Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Private Declare PtrSafe Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
'ネットワークバイトオーダからホストバイトオーダに変更
Private Declare PtrSafe Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Private Declare PtrSafe Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer

Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal hpvDest As LongPtr, ByVal hpvSource As LongPtr, ByVal cbCopy As Long)
Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal Buffer As String, ByRef Size As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

'error code
Private Const WSABASEERR             As Long = 10000 'No Error
Private Const WSAEINTR               As Long = 10004 'Interrupted by system call
Private Const WSAEBADF               As Long = 10009 '無効なファイルハンドルがソケット関数に渡された
Private Const WSAEACCES              As Long = 10013 'access denied
Private Const WSAEFAULT              As Long = 10014 '無効なバッファアドレス
Private Const WSAEINVAL              As Long = 10022 '無効な引数
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

'エラーメッセージを返します。
Private Function ErrorMsg(ByVal ErrorCode) As String
    ErrorMsg = ""
    Select Case ErrorCode
        Case WSAEINTR:           ErrorMsg = "関数呼び出しに割り込みがありました。"
        Case WSAEACCES:          ErrorMsg = "アクセスは拒否されました。"
        Case WSAEFAULT:          ErrorMsg = "アドレスが正しくありません。"
        Case WSAEINVAL:          ErrorMsg = "無効な引数です。"
        Case WSAEMFILE:          ErrorMsg = "開いているファイルが多すぎます。"
        Case WSAEWOULDBLOCK:     ErrorMsg = "ノンブロッキング状態であり、処理が直ちに完了しませんでした。。"
        Case WSAEINPROGRESS:     ErrorMsg = "ブロック操作を実行中です。"
        Case WSAEALREADY:        ErrorMsg = "操作はすでに実行中です。"
        Case WSAENOTSOCK:        ErrorMsg = "記述子がソケットではありません。"
        Case WSAEDESTADDRREQ:    ErrorMsg = "送信先のアドレスが必要です。"
        Case WSAEMSGSIZE:        ErrorMsg = "メッセージが長すぎます。"
        Case WSAEPROTOTYPE:      ErrorMsg = "プロトコルの種類がソケットに対して正しくありません。"
        Case WSAENOPROTOOPT:     ErrorMsg = "プロトコルのオプションが正しくありません。"
        Case WSAEPROTONOSUPPORT: ErrorMsg = "指定したプロトコルがサポートされていません。"
        Case WSAESOCKTNOSUPPORT: ErrorMsg = "サポートされていないプロトコルの種類です。"
        Case WSAEOPNOTSUPP:      ErrorMsg = "要求された操作がソケット上でサポートされていません。"
        Case WSAEPFNOSUPPORT:    ErrorMsg = "プロトコルファミリがサポートされていません。"
        Case WSAEAFNOSUPPORT:    ErrorMsg = "指定されたプロトコルファミリは指定されたアドレスファミリをサポートしていません。"
        Case WSAEADDRINUSE:      ErrorMsg = "アドレスが使用中です。"
        Case WSAEADDRNOTAVAIL:   ErrorMsg = "アドレスをローカルマシンから取得できません。"
        Case WSAENETDOWN:        ErrorMsg = "ネットワークがダウンしています。"
        Case WSAENETUNREACH:     ErrorMsg = "現時点ではこのホストからネットワークにアクセスできません。"
        Case WSAENETRESET:       ErrorMsg = "ネットワークがリセットされたため切断されました。"
        Case WSAECONNABORTED:    ErrorMsg = "タイムアウトその他の不具合で接続処理が中止されました。"
        Case WSAECONNRESET:      ErrorMsg = "ピアによって接続がリセットされました。"
        Case WSAENOBUFS:         ErrorMsg = "使用できるバッファ領域がありません。"
        Case WSAEISCONN:         ErrorMsg = "ソケットは既に接続されています。"
        Case WSAENOTCONN:        ErrorMsg = "ソケットは接続されていません。"
        Case WSAESHUTDOWN:       ErrorMsg = "ソケットは終了しています。"
        Case WSAETIMEDOUT:       ErrorMsg = "接続がタイムアウトになりました。"
        Case WSAECONNREFUSED:    ErrorMsg = "接続が強制的に拒絶されました。"
        Case WSAEHOSTDOWN:       ErrorMsg = "ホストがダウンしています。"
        Case WSAEHOSTUNREACH:    ErrorMsg = "ホストに到達するためのルートがありません。"
        Case WSAEPROCLIM:        ErrorMsg = "プロセスが多すぎます。"
        Case WSASYSNOTREADY:     ErrorMsg = "ネットワークサブシステムが利用できません。"
        Case WSAVERNOTSUPPORTED: ErrorMsg = "要求された Winsock のバージョンはサポートされていません。"
        Case WSANOTINITIALISED:  ErrorMsg = "まず WSAStartup を呼び出す必要があります。"
        Case WSAEDISCON:         ErrorMsg = "正常なシャットダウン処理が進行中です。"
        Case WSATYPE_NOT_FOUND:  ErrorMsg = "この種類のクラスが見つかりません。"
        Case WSAHOST_NOT_FOUND:  ErrorMsg = "ホストが見つかりません。そのようなホストはありません。"
        Case WSATRY_AGAIN:       ErrorMsg = "ホストが見つかりません。DNSサーバーからの応答がありません。"
        Case WSANO_RECOVERY:     ErrorMsg = "回復不能なエラーです。"
        Case WSANO_DATA:         ErrorMsg = "名前は有効ですが、要求された型のデータレコードがありません。"
        Case sckInvalidOp:       ErrorMsg = "現在の状態では不正な操作です。"
        Case Else:               ErrorMsg = "不明なエラー"
    End Select
    ErrorMsg = ErrorMsg & vbCrLf & "エラーコード : " & str$(ErrorCode)
End Function

'--------------------------------------------------
' UTF8関連
'--------------------------------------------------

'UTF8バイト配列をUnicode文字列に変換
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

'Unicode文字列をUTF8バイト配列に変換
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
' サーバプロシージャ
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
Public Sub wsTCPListen()
    Dim Buf As RingBuf
    Dim WSAD As WSAData
    Dim sockfd As LongPtr 'socket
    Dim HostAddr As sockaddr, ClientAddr As sockaddr
    Dim HostAddr_in As sockaddr_in
    Dim Addrptr As LongPtr
    Dim ret As Long, flg As Long
    Dim yes As Boolean, err As Boolean
    Dim i As Integer, j As Integer
    Dim wait As Single
        
    'winsock初期化
    If (WSAStartup(WS_VERSION_REQD, WSAD)) Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If

    'ソケット取得
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
    
    '待受ソケット構造体
    HostAddr_in.sin_family = AF_INET
    HostAddr_in.sin_port = htons(PORT)
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
        ClearBuf Buf

        'クライアントソケット
        Buf.Fd = w_accept(sockfd, ClientAddr, SOCKADDR_SIZE)
        MoveMemory VarPtr(Buf.Addr), VarPtr(ClientAddr), SOCKADDR_SIZE
        If Buf.Fd = -1 Then
            Debug.Print ErrorMsg(WSAGetLastError)
            Exit Do
        End If
        
        'non-block
        '空のコネクションをつかむとblock_ioでは戻ってこなくなる。
        flg = 1
        w_ioctlsocket Buf.Fd, FIONBIO, flg

        DoEvents

        wait = Timer
        Do
            ret = w_recv(Buf.Fd, Buf.Recvbyte(0), RecvSize, 0)
            If Timer > wait + TOUT Then Exit Do
            DoEvents
        Loop While (ret = SOCKET_ERROR)

        If ret = SOCKET_ERROR Then
            Debug.Print "w_recv " & ErrorMsg(WSAGetLastError)
            err = True
        ElseIf ret = 0 Then
            '切断
            Debug.Print "w_recv 切断"
            err = True
        Else
            '使用済みバイト保存
            i = CInt(ret)

            If (CrLf2(Buf.Recvbyte, Buf.HeadTerm)) And (ContentLength(Buf.Recvbyte, j, Buf.HeadTerm)) Then
                err = False
                Buf.RecvTerm = j + Buf.HeadTerm
            Else
                DoEvents
                'loop
                Do
                    ret = w_recv(Buf.Fd, Buf.Recvbyte(i), RecvSize - i, 0)
                    If ret > 0 Then
                        i = i + ret
                    Else
                        Exit Do
                    End If
                Loop

                If Buf.HeadTerm = 0 And (Not CrLf2(Buf.Recvbyte, Buf.HeadTerm)) Then
                    err = True
                ElseIf Not (ContentLength(Buf.Recvbyte, j, Buf.HeadTerm)) Then
                    err = True
                Else
                    err = False
                    Buf.RecvTerm = j + Buf.HeadTerm
                End If
            End If

        End If
        
        If err = False Then
            '加工処理(RecvByteからSendByteを生成)
            Process Buf.HeadTerm, Buf.RecvTerm, Buf.Recvbyte, Buf.Sendbyte
            '送信処理
            SendRes Buf.Fd, Buf.Sendbyte
        End If
        
        If (w_closesocket(Buf.Fd) = SOCKET_ERROR) Then
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

Public Sub wsTCPListenMpx()
    Dim Buf(BufSize) As RingBuf
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
        
    'winsock初期化
    If (WSAStartup(WS_VERSION_REQD, WSAD)) Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If

    'ソケット取得
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

    '待受ソケット構造体
    HostAddr_in.sin_family = AF_INET
    HostAddr_in.sin_port = htons(PORT)
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
    'Buf(0)を親ソケットにする。
    Buf(0).Fd = ss
    Debug.Print "opened:" & ss
    '子ソケット初期化
    For i = 1 To BufSize
        ClearBuf Buf(i)
    Next i
    
    Do
        'clean up timeout fd
        wait = Timer
        For i = 1 To BufSize
            If Buf(i).Fd <> -1 And wait > (Buf(i).Cntr + TOUT) Then
                Debug.Print "clean up:" & Buf(i).Fd
                If (w_closesocket(Buf(i).Fd) = SOCKET_ERROR) Then
                    Debug.Print ErrorMsg(WSAGetLastError)
                End If
                ClearBuf Buf(i)
            End If
        Next i
        'select
        'マクロを使った実装になっているので、その部分をVBA側で
        '置き換える必要がある。
        'Fdsは都度上書きされるらしい
        vbaFD_ZERO Fds

        For i = 0 To BufSize
            If Buf(i).Fd <> -1 Then vbaFD_SET Buf(i).Fd, Fds
        Next i
        
        cnt = Fds.fd_count
        
        If Q Then Exit Do
        
        If w_select(0, VarPtr(Fds), NULLPTR, NULLPTR, TV) <> 0 Then
            If cnt > 1 Then
                'clean up disconnect fd
                'recv activated fd
                For i = 1 To BufSize
                    If Buf(i).Fd <> -1 And vbaFD_ISSET(Buf(i).Fd, Fds) <> 0 Then
                        ret = w_recv(Buf(i).Fd, Buf(i).Recvbyte(Buf(i).RecvTerm), RecvSize - Buf(i).RecvTerm, 0)
                        If ret = 0 Then
                            Debug.Print "disconnected:" & Buf(i).Fd
                            If (w_closesocket(Buf(i).Fd) = SOCKET_ERROR) Then
                                Debug.Print ErrorMsg(WSAGetLastError)
                            End If
                            ClearBuf Buf(i)
                        ElseIf ret > 0 Then
                            Buf(i).RecvTerm = Buf(i).RecvTerm + ret
                            Debug.Print "received:" & Buf(i).Fd
                        Else
                            Debug.Print "receive_error:" & Buf(i).Fd
                        End If
                    End If
                Next i
                'check buffered fd
                For i = 1 To BufSize
                    If Buf(i).Fd <> -1 And (CrLf2(Buf(i).Recvbyte, Buf(i).HeadTerm)) And _
                            (ContentLength(Buf(i).Recvbyte, j, Buf(i).HeadTerm)) Then
                        Buf(i).RecvTerm = Buf(i).HeadTerm + j
                        Buf(i).FlgIn = True
                        Debug.Print "checked:" & Buf(i).Fd
                    End If
                Next i
                'clean up full-filled fd
                For i = 1 To BufSize
                    If Buf(i).Fd <> -1 And Buf(i).RecvTerm = RecvSize And Buf(i).FlgIn = False Then
                        Debug.Print "full buffered:" & Buf(i).Fd
                        If (w_closesocket(Buf(i).Fd) = SOCKET_ERROR) Then
                            Debug.Print ErrorMsg(WSAGetLastError)
                        End If
                        ClearBuf Buf(i)
                    End If
                Next i
            End If
            'accept
            If vbaFD_ISSET(ss, Fds) <> 0 Then
                j = 1
                Do
                    NA = True
                    For i = j To BufSize
                        If Buf(i).Fd = -1 Then
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
                        Buf(i).Fd = sc
                        Buf(i).Cntr = Timer
                        MoveMemory VarPtr(Buf(i).Addr), VarPtr(ClientAddr), SOCKADDR_SIZE
                        Debug.Print "accepted:" & Buf(i).Fd
                    End If
                Loop
            End If
        End If

        'http io
        For i = 1 To BufSize
            If Buf(i).FlgIn = True Then
                Process Buf(i).HeadTerm, Buf(i).RecvTerm, Buf(i).Recvbyte, Buf(i).Sendbyte
                Buf(i).FlgOut = True
            End If
        Next i
        
        'send
        For i = 1 To BufSize
            If Buf(i).FlgOut = True Then
                SendRes Buf(i).Fd, Buf(i).Sendbyte
                
                If (w_closesocket(Buf(i).Fd) = SOCKET_ERROR) Then
                    Debug.Print ErrorMsg(WSAGetLastError)
                End If
                
                Debug.Print "sent:" & Buf(i).Fd
                ClearBuf Buf(i)
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

        
    'winsock初期化
    If (WSAStartup(WS_VERSION_REQD, WSAD)) Then
        MsgBox ErrorMsg(WSAGetLastError)
        WSACleanup
        Exit Sub
    End If

    'ソケット取得
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
    
    '待受ソケット構造体
    HostAddr_in.sin_family = AF_INET
    HostAddr_in.sin_port = htons(PORT)
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

        'クライアントソケット
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
                '切断
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
            '加工処理(RecvByteからSendByteを生成)
            Process HeadTerm, HeadTerm + Length, Recvbyte, Sendbyte
            '送信処理
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

'加工処理サンプル
Private Sub Process(ByVal HeadTerm As Integer, ByVal RecvTerm As Integer, Recvbyte() As Byte, Sendbyte() As Byte)
    Dim i As Integer, j As Integer, k As Integer
    Dim Arg(RecvSize) As Byte, Res() As Byte
    Dim Recvstring As String, Sendstring As String, Argstring As String, Mthd As String, User As String
    
    'On Error GoTo Error_
    
    Recvstring = StrConv(Recvbyte, vbUnicode)
    If Len(Recvstring) < 4 Then
        Mthd = ""
    Else
        Mthd = Left(Recvstring, 4)
    End If
    
    Select Case Mthd
        Case "HEAD"
            Sendstring = "HTTP/1.0 200 OK" & vbCrLf & _
            "Server: VBA webserver" & vbCrLf & _
            vbCrLf
            
        Case "GET ", "POST"
            'GET /(user)...
            'ユーザーが入っているときは登録画面を返す
            'POST /(user)...
            'ユーザーが入っているときは登録画面を返し登録処理をする
            'GET or POST /... -> return csv
            'ユーザーが入っていないときはcsvを返す
            'GET or POST /(userではない文字列)...
            'Not Found
            'GET /q -> 終了
            Argstring = ""
            If Mthd = "POST" Then
                'USER　"POST /"以降の7文字目からスペースまで
                i = InStr(7, Recvstring, Chr(32))
                If i > 7 Then
                    User = Mid(Recvstring, 7, i - 7)
                Else
                    User = ""
                End If
                'メッセージボディからパーセントエンコーディング復元してstringに
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
                 'USER　"GET /"以降の6文字目からスペースまで
                i = InStr(6, Recvstring, Chr(32))
                If i > 6 Then
                    User = Mid(Recvstring, 6, i - 6)
                Else
                    User = ""
                End If
            End If
            
            If User = "" Then
'                Sendstring = "HTTP/1.0 200 OK" & vbCrLf & _
'                "Server: VBA webserver" & vbCrLf & _
'                "Content-Type: application/octet-stream" & vbCrLf & _
'                "Content-Disposition: attachment; filename=OJISAN.csv" & vbCrLf & _
'                "Content-Transfer-Encoding: binary" & _
'                vbCrLf & vbCrLf & _
'                Csv
                Sendstring = "HTTP/1.0 200 OK" & vbCrLf & _
                "Server: VBA webserver" & vbCrLf & _
                "Content-Type: text/plain; charset=Shift_JIS" & vbCrLf & _
                vbCrLf & _
                Csv
            ElseIf User = "q" Then
                Q = True
                Sendstring = "HTTP/1.0 200 OK" & vbCrLf & _
                "Server: VBA webserver" & vbCrLf & _
                "Content-Type: text/plain; charset=Shift_JIS" & vbCrLf & _
                vbCrLf & _
                "QUIT"
            ElseIf User Like "*[!a-z]*" Then
                Sendstring = "HTTP/1.0 404 Not Found" & vbCrLf & _
                "Server: VBA webserver" & vbCrLf & _
                vbCrLf
            Else
                Sendstring = "HTTP/1.0 200 OK" & vbCrLf & _
                "Server: VBA webserver" & vbCrLf & _
                vbCrLf & _
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
                    Paste User, Argstring
                End If
            End If
            
        Case Else
            Sendstring = "HTTP/1.0 404 Not Found" & vbCrLf & _
            "Server: VBA webserver" & vbCrLf & _
            vbCrLf
    End Select
    
    Res = StrConv(Sendstring, vbFromUnicode)
    Erase Sendbyte
    
    If UBound(Res) < SendSize Then
        k = UBound(Res) + 1
    Else
        k = SendSize
    End If
    
    MoveMemory VarPtr(Sendbyte(0)), VarPtr(Res(0)), k
    'Debug.Print Recvstring & vbCrLf & Sendstring
    
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
            Debug.Print "w_send 切断"
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

Private Sub ClearBuf(Buf As RingBuf)

    Buf.Fd = -1
    Buf.FlgIn = False
    Buf.FlgOut = False
    Buf.HeadTerm = 0
    Buf.RecvTerm = 0
    Buf.Cntr = 0
    Erase Buf.Recvbyte
    Erase Buf.Sendbyte
    ZeroMemory Buf.Addr, SOCKADDR_IN_SIZE

End Sub


Private Sub Paste(User As String, Arg As String)
    Dim LastRow As Integer
    
    'エスケープ処理
    LastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
    ActiveSheet.Cells(LastRow + 1, 1).Value = Now()
    ActiveSheet.Cells(LastRow + 1, 2).Value = Replace(Replace(User, Chr(34), ""), Chr(44), "")
    ActiveSheet.Cells(LastRow + 1, 3).Value = Replace(Replace(Replace(Replace(Arg, vbCr, ""), vbLf, ""), Chr(34), ""), Chr(44), "")
    
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
             Data(i, 2) & COMMA & Data(i, 3) & vbCrLf
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

'オリジナルは追加するだけやけど、nfdsが取りやすいように
'降順で追加する。
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


