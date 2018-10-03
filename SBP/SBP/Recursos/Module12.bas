Attribute VB_Name = "Module12"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2004 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
'Determining a Local or Remote MAC Address via SendARP
'Distributor Okan CELEN
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const NO_ERROR = 0

Private Declare Function inet_addr Lib "wsock32.dll" (ByVal s As String) As Long

Private Declare Function SendARP _
                Lib "iphlpapi.dll" (ByVal DestIP As Long, _
                                    ByVal SrcIP As Long, _
                                    pMacAddr As Long, _
                                    PhyAddrLen As Long) As Long

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (dst As Any, _
                                       src As Any, _
                                       ByVal bcount As Long)

Public Function GetRemoteMACAddress(ByVal sRemoteIP As String, _
                                    sRemoteMacAddress As String, _
                                    sDelimiter As String) As Boolean

    Dim dwRemoteIP  As Long

    Dim pMacAddr    As Long

    Dim bpMacAddr() As Byte

    Dim PhyAddrLen  As Long

    dwRemoteIP = ConvertIPtoLong(sRemoteIP)

    'MsgBox dwRemoteIP
    If dwRemoteIP <> 0 Then
        PhyAddrLen = 6
        GetRemoteMACAddress = False

        'MsgBox SendARP(dwRemoteIP, 0&, pMacAddr, PhyAddrLen)
        If SendARP((dwRemoteIP), 0&, pMacAddr, PhyAddrLen) = NO_ERROR Then

            'MsgBox pMacAddr & " " & PhyAddrLen
            If (pMacAddr <> 0) And (PhyAddrLen <> 0) Then
                ReDim bpMacAddr(0 To PhyAddrLen - 1)
                CopyMemory bpMacAddr(0), pMacAddr, ByVal PhyAddrLen
                sRemoteMacAddress = MakeMacAddress(bpMacAddr(), sDelimiter)
                'MsgBox sRemoteMacAddress
                GetRemoteMACAddress = True

            End If

        End If

        'MsgBox "xyz"
    End If

End Function

Public Function ConvertIPtoLong(sIpAddress) As Long
    ConvertIPtoLong = inet_addr(sIpAddress)

End Function

Public Function MakeMacAddress(b() As Byte, sDelim As String) As String

    Dim Cnt  As Long

    Dim buff As String

    On Local Error GoTo MakeMac_error

    If UBound(b) = 5 Then

        For Cnt = 0 To 4
            buff = buff & Right$("00" & Hex(b(Cnt)), 2) & sDelim
        Next
        buff = buff & Right$("00" & Hex(b(5)), 2)

    End If

    MakeMacAddress = buff

MakeMac_exit:
    Exit Function

MakeMac_error:
    MakeMacAddress = "(error MAC address)"

    Resume MakeMac_exit

End Function

