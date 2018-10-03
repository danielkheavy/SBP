Attribute VB_Name = "Module14"
Option Explicit

' Declarations needed for GetAdaptersInfo & GetIfTable
Private Const MIB_IF_TYPE_OTHER                  As Long = 1

Private Const MIB_IF_TYPE_ETHERNET               As Long = 6

Private Const MIB_IF_TYPE_TOKENRING              As Long = 9

Private Const MIB_IF_TYPE_FDDI                   As Long = 15

Private Const MIB_IF_TYPE_PPP                    As Long = 23

Private Const MIB_IF_TYPE_LOOPBACK               As Long = 24

Private Const MIB_IF_TYPE_SLIP                   As Long = 28

Private Const MIB_IF_ADMIN_STATUS_UP             As Long = 1

Private Const MIB_IF_ADMIN_STATUS_DOWN           As Long = 2

Private Const MIB_IF_ADMIN_STATUS_TESTING        As Long = 3

Private Const MIB_IF_OPER_STATUS_NON_OPERATIONAL As Long = 0

Private Const MIB_IF_OPER_STATUS_UNREACHABLE     As Long = 1

Private Const MIB_IF_OPER_STATUS_DISCONNECTED    As Long = 2

Private Const MIB_IF_OPER_STATUS_CONNECTING      As Long = 3

Private Const MIB_IF_OPER_STATUS_CONNECTED       As Long = 4

Private Const MIB_IF_OPER_STATUS_OPERATIONAL     As Long = 5

Private Const MAX_ADAPTER_DESCRIPTION_LENGTH     As Long = 128

Private Const MAX_ADAPTER_DESCRIPTION_LENGTH_p   As Long = MAX_ADAPTER_DESCRIPTION_LENGTH + 4

Private Const MAX_ADAPTER_NAME_LENGTH            As Long = 256

Private Const MAX_ADAPTER_NAME_LENGTH_p          As Long = MAX_ADAPTER_NAME_LENGTH + 4

Private Const MAX_ADAPTER_ADDRESS_LENGTH         As Long = 8

Private Const DEFAULT_MINIMUM_ENTITIES           As Long = 32

Private Const MAX_HOSTNAME_LEN                   As Long = 128

Private Const MAX_DOMAIN_NAME_LEN                As Long = 128

Private Const MAX_SCOPE_ID_LEN                   As Long = 256

Private Const MAXLEN_IFDESCR                     As Long = 256

Private Const MAX_INTERFACE_NAME_LEN             As Long = MAXLEN_IFDESCR * 2

Private Const MAXLEN_PHYSADDR                    As Long = 8

' Information structure returned by GetIfEntry/GetIfTable
Private Type MIB_IFROW

    wszName(0 To MAX_INTERFACE_NAME_LEN - 1) As Byte    ' MSDN Docs say pointer, but it is WCHAR array
    dwIndex             As Long
    dwType              As Long
    dwMtu               As Long
    dwSpeed             As Long
    dwPhysAddrLen       As Long
    bPhysAddr(MAXLEN_PHYSADDR - 1) As Byte
    dwAdminStatus       As Long
    dwOperStatus        As Long
    dwLastChange        As Long
    dwInOctets          As Long
    dwInUcastPkts       As Long
    dwInNUcastPkts      As Long
    dwInDiscards        As Long
    dwInErrors          As Long
    dwInUnknownProtos   As Long
    dwOutOctets         As Long
    dwOutUcastPkts      As Long
    dwOutNUcastPkts     As Long
    dwOutDiscards       As Long
    dwOutErrors         As Long
    dwOutQLen           As Long
    dwDescrLen          As Long
    bDescr As String * MAXLEN_IFDESCR

End Type

Private Type TIME_t

    aTime As Long

End Type

Private Type IP_ADDRESS_STRING

    IPadrString     As String * 16

End Type

Private Type IP_ADDR_STRING

    AdrNext         As Long
    IpAddress       As IP_ADDRESS_STRING
    IpMask          As IP_ADDRESS_STRING
    NTEcontext      As Long

End Type

' Information structure returned by GetIfEntry/GetIfTable
Private Type IP_ADAPTER_INFO

Next As Long

ComboIndex As Long
AdapterName         As String * MAX_ADAPTER_NAME_LENGTH_p
Description         As String * MAX_ADAPTER_DESCRIPTION_LENGTH_p
MACadrLength        As Long
MACaddress(0 To MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
AdapterIndex        As Long
AdapterType         As Long             ' MSDN Docs say "UInt", but is 4 bytes
DhcpEnabled         As Long             ' MSDN Docs say "UInt", but is 4 bytes
CurrentIpAddress    As Long
IpAddressList       As IP_ADDR_STRING
GatewayList         As IP_ADDR_STRING
DhcpServer          As IP_ADDR_STRING
HaveWins            As Long             ' MSDN Docs say "Bool", but is 4 bytes
PrimaryWinsServer   As IP_ADDR_STRING
SecondaryWinsServer As IP_ADDR_STRING
LeaseObtained       As TIME_t
LeaseExpires        As TIME_t

End Type
     
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                       ByRef Source As Any, _
                                       ByVal numbytes As Long)

Public Declare Function GetAdaptersInfo _
               Lib "iphlpapi.dll" (ByRef pAdapterInfo As Any, _
                                   ByRef pOutBufLen As Long) As Long

Public Declare Function GetNumberOfInterfaces _
               Lib "iphlpapi.dll" (ByRef pdwNumIf As Long) As Long

Public Declare Function GetIfEntry Lib "iphlpapi.dll" (ByRef pIfRow As Any) As Long

Private Declare Function GetIfTable _
                Lib "iphlpapi.dll" (ByRef pIfTable As Any, _
                                    ByRef pdwSize As Long, _
                                    ByVal bOrder As Long) As Long

'-----------------------------------------------------------------------------------
' Get the system's MAC address(es) via GetAdaptersInfo API function (IPHLPAPI.DLL)
'
' Note: GetAdaptersInfo returns information about physical adapters
'-----------------------------------------------------------------------------------
Public Function GetMACs_AdaptInfo() As String

    Dim xbuf     As String

    Dim xbuf1    As String

    Dim AdapInfo As IP_ADAPTER_INFO, bufLen As Long, sts As Long

    Dim retStr   As String, numStructs%, I%, IPinfoBuf() As Byte, srcPtr As Long
    
    ' Get size of buffer to allocate
    sts = GetAdaptersInfo(AdapInfo, bufLen)

    If (bufLen = 0) Then Exit Function
    numStructs = bufLen / Len(AdapInfo)
    retStr = numStructs & " Adapter(s):" & vbcrlf
    'MsgBox retStr
    ' reserve byte buffer & get it filled with adapter information
    ' !!! Don't Redim AdapInfo array of IP_ADAPTER_INFO,
    ' !!! because VB doesn't allocate it contiguous (padding/alignment)
    ReDim IPinfoBuf(0 To bufLen - 1) As Byte
    sts = GetAdaptersInfo(IPinfoBuf(0), bufLen)

    If (sts <> 0) Then Exit Function
    
    ' Copy IP_ADAPTER_INFO slices into UDT structure
    srcPtr = VarPtr(IPinfoBuf(0))
    xbuf = ""
    'MsgBox srcPtr

    'For i = 0 To numStructs - 1
    For I = 0 To 0

        If (srcPtr = 0) Then Exit For
        '        CopyMemory AdapInfo, srcPtr, Len(AdapInfo)
        CopyMemory AdapInfo, ByVal srcPtr, Len(AdapInfo)
        xbuf = MAC2String(AdapInfo.MACaddress)

        ' Extract Ethernet MAC address
        With AdapInfo

            If (.AdapterType = MIB_IF_TYPE_ETHERNET) Then
                retStr = retStr & vbcrlf & "[" & I & "] " & sz2string(.Description) & vbcrlf & vbTab & MAC2String(.MACaddress) & vbcrlf

            End If

        End With

        srcPtr = AdapInfo.Next
    Next I

    'MsgBox xbuf
    ' Return list of MAC address(es)
    'GetMACs_AdaptInfo = retStr
    xbuf1 = ""

    For I = 1 To Len(xbuf)

        If Mid$(xbuf, I, 1) <> "-" Then
            xbuf1 = xbuf1 & Mid$(xbuf, I, 1)

        End If

    Next I

    xbuf = xbuf1
    'MsgBox xbuf
    GetMACs_AdaptInfo = xbuf
    
End Function

'-----------------------------------------------------------------------------------
' Get the system's MAC address(es) via GetIfTable API function (IPHLPAPI.DLL)
'
' Note: GetIfTable returns information also about the virtual loopback adapter
'-----------------------------------------------------------------------------------
Public Function GetMACs_IfTable() As String
    
    Dim NumAdapts As Long, nRowSize As Long, I%, retStr As String

    Dim IfInfo    As MIB_IFROW, IPinfoBuf() As Byte, bufLen As Long, sts As Long
    
    ' Get # of interfaces defined (sometimes 1 more than GetIfTable)
    sts = GetNumberOfInterfaces(NumAdapts)
    
    ' Get size of buffer to allocate
    sts = GetIfTable(ByVal 0&, bufLen, 1)

    If (bufLen = 0) Then Exit Function

    ' reserve byte buffer & get it filled with adapter information
    ReDim IPinfoBuf(0 To bufLen - 1) As Byte
    sts = GetIfTable(IPinfoBuf(0), bufLen, 1)

    If (sts <> 0) Then Exit Function
    
    NumAdapts = IPinfoBuf(0)
    nRowSize = Len(IfInfo)
    retStr = NumAdapts & " Interface(s):" & vbcrlf

    For I = 1 To NumAdapts
        ' copy one IfRow chunk of byte data into an MIB_IFROW structure
        Call CopyMemory(IfInfo, IPinfoBuf(4 + (I - 1) * nRowSize), nRowSize)
        
        ' Take adapter address if correct type
        With IfInfo
            retStr = retStr & vbcrlf & "[" & I & "] " & Left$(.bDescr, .dwDescrLen - 1) & vbcrlf

            If (.dwType = MIB_IF_TYPE_ETHERNET) Then
                retStr = retStr & vbTab & MAC2String(.bPhysAddr) & vbcrlf

            End If

        End With

    Next I

    GetMACs_IfTable = retStr
    
End Function

' Convert a byte array containing a MAC address to a hex string
Private Function MAC2String(AdrArray() As Byte) As String

    Dim aStr As String, hexStr As String, I%
    
    For I = 0 To 5

        If (I > UBound(AdrArray)) Then
            hexStr = "00"
        Else
            hexStr = Hex$(AdrArray(I))

        End If
        
        If (Len(hexStr) < 2) Then hexStr = "0" & hexStr
        aStr = aStr & hexStr

        If (I < 5) Then aStr = aStr & "-"
    Next I
    
    MAC2String = aStr
    
End Function

' Convert a zero-terminated fixed string to a dynamic VB string
Private Function sz2string(ByVal szStr As String) As String
    sz2string = Left$(szStr, InStr(1, szStr, Chr$(0)) - 1)

End Function

Function serial_procesador() As String

    Dim buf As String

    Dim wmi As Object

    Dim mos As Object

    Dim mo  As Object

    buf = ""
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    'Set mos = wmi.ExecQuery("Select * from Win32_Baseboard")
    'Win32_Processor
    Set mos = wmi.ExecQuery("Select * from Win32_Processor")
    
    buf = ""

    For Each mo In mos

        '    Text1 = Text1 & "Serial Number: " & mo.SerialNumber & vbCrLf
        '    Text1 = Text1 & "Manufacturer: " & mo.Manufacturer & vbCrLf
        '    Text1 = Text1 & "Product: " & mo.Product
        buf = buf + mo.ProcessorId
    Next
    serial_procesador = buf

End Function

Function WMIDetect() As Boolean

    On Error GoTo NOWMI

    Dim wmi As Object

    Set wmi = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set wmi = Nothing
    WMIDetect = True
    Exit Function
NOWMI:
    WMIDetect = False

End Function

'mac address

