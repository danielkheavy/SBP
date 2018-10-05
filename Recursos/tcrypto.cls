VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tcrypto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*************************************************************************************
'*************************************************************************************
'*** Instructor CCNA - Wagner, Ariel Alejandro  **************************************
'*** Cifrar y Descifrar Cadenas de Texto        **************************************
'*** 2 - 10 - 2008                              **************************************
'*************************************************************************************
'*************************************************************************************

Public Function Encrypt(ByVal Word As String, _
                        ByVal Key As String, _
                        Optional ByVal Mode As Boolean = False) As String

    Dim w  As Long, k As Long, p As Long, j As Long, NuChr As Long

    Dim Cd As String, Kd As String, Rd As String

    ' Cantidad de caracteres...
    w = Len(Word)
    k = Len(Key)

    ' Modalidad de Encripción...
    If Mode = False Then

        For j = 1 To w
            Cd = Mid(Word, j, 1)

            If p = k Then p = 0
            p = p + 1
            Kd = Mid(Key, p, 1)
            NuChr = Asc(Cd) + Asc(Kd)

            If NuChr > 255 Then
                NuChr = NuChr - 255

            End If

            Rd = Rd & Chr(NuChr)
        Next
        Encrypt = Rd
        Exit Function

    End If

    ' Modalidad de Dencripción...
    If Mode = True Then

        For j = 1 To w
            Cd = Mid(Word, j, 1)

            If p = k Then p = 0
            p = p + 1
            Kd = Mid(Key, p, 1)
            NuChr = Asc(Cd) - Asc(Kd)

            If NuChr < 0 Then
                NuChr = NuChr + 255

            End If

            Rd = Rd & Chr(NuChr)
        Next
        Encrypt = Rd
        Exit Function

    End If

End Function

