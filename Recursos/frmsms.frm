VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmsms 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SMS - Mensajes Celular"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmsms.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTelephone 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Buscar mensajes Servidor"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
   Begin VB.ListBox lstEvents 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5325
      Left            =   6600
      TabIndex        =   7
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox txtSend 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox txtMobilenumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Text            =   "+990988493"
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Timer tmrCheckMessage 
      Interval        =   15000
      Left            =   120
      Top             =   600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Habilitar Mensajes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar Mensaje"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   4320
      Width           =   1935
   End
   Begin MSCommLib.MSComm Comm1 
      Left            =   5760
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      ParityReplace   =   32
      RThreshold      =   1
      RTSEnable       =   -1  'True
      BaudRate        =   4800
   End
   Begin VB.Label Label2 
      Caption         =   "Escribir Mensaje :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Recd. Mensaje :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.Menu ldoasli55 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "frmsms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public bEcho            As Boolean

Public bOK              As Boolean

Public bRing            As Boolean

Public bError           As Boolean

Public iRingTime        As Single

Public FirstRun         As Boolean

Public bErrorComm       As Boolean

Public bGreaterSign     As Boolean

Public bMessageStore    As Boolean

Public strMessageBuffer As String

Public FileNumber       As Integer

Dim msgBreak()          As String

Dim msgHeader()         As String

Private Sub comm1_OnComm()

    Static stEvent As String

    Dim stComChar  As String * 1

    Select Case Comm1.CommEvent

        Case comEvReceive

            Do
                stComChar = Comm1.input

                If bMessageStore Then
                    strMessageBuffer = strMessageBuffer & stComChar

                End If

                Select Case stComChar

                    Case ">"
                        bGreaterSign = True
                        lstEvents.AddItem stComChar

                    Case vbLf

                    Case vbCr

                        If Len(stEvent) > 0 Then
                            ProcessEvent stEvent
                            stEvent = ""

                        End If

                    Case Else
                        stEvent = stEvent + stComChar

                End Select

            Loop While Comm1.InBufferCount

    End Select

End Sub

Private Sub Command1_Click()

    If Len(Trim(txtMobilenumber.Text)) = 0 Then
        MsgBox "Ingrese numero antes de enviar ! " & vbCr & "El formato es +996249478", vbInformation + vbOKOnly, "Numero no Valido"
        Exit Sub
    Else
        bGreaterSign = False
        Comm1.Output = "AT+CMGS=" & Chr(34) & Trim(txtMobilenumber.Text) & Chr(34) & vbcrlf

        While Not bGreaterSign

            DoEvents
            Wait
        Wend

        If bGreaterSign Then
            Comm1.Output = Trim(txtSend.Text) & Chr(26) & vbcrlf
            bOK = False
            bError = False

            While Not bOK Or bError

                DoEvents
                Wait
            Wend

            If bOK Then
                MsgBox "Mensaje Enviado", vbInformation + vbOKOnly, "Envio"
            Else
                MsgBox "Mensaje no enviado", vbCritical + vbOKOnly, "No Envio"

            End If

        Else
            MsgBox "Mensaje no puede ser Enviado", vbCritical + vbOKOnly, "No Envio"

        End If

        txtSend.Text = ""
        txtMobilenumber.Text = ""

    End If

End Sub

Private Sub Command2_Click()

    If Comm1.PortOpen = False Then
        Comm1.PortOpen = True
        Comm1.DTREnable = True
        Comm1.RTSEnable = True
        Comm1.RThreshold = 1
        Comm1.InputLen = 1
        bOK = False
        bError = False
        Comm1.Output = "AT" & vbcrlf
        Wait

        If Not bOK Then
            MsgBox "Modem no esta respondiendo"
            Comm1.PortOpen = False
            Exit Sub

        End If

        Command1.Enabled = True
        Command3.Enabled = True
        Command2.Enabled = False
    Else
        MsgBox "Puerto ya abierto !", vbCritical + vbOKOnly, "Error al abrir puerto"

    End If

End Sub

Private Sub ProcessEvent(stEvent As String)

    Dim stNumber As String
  
    lstEvents.AddItem stEvent

    If Mid$(stEvent, 1, 5) = "+CMTI" Then
        txtTelephone.Text = ""
        txtMessage.Text = ""
        strMessageBuffer = ""

        If MsgBox("Nuevo mensaje recibido! " & vbCr & "Desea cargar ?", vbYesNo + vbQuestion, "Por favor Confirmar") = vbYes Then
            stEvent = ""
            Command3_Click

        End If

        bOK = False
        bError = False
        Comm1.Output = "AT+CMGD=1,4" & vbcrlf

        While Not bOK Or bError

            DoEvents
            Wait
        Wend

        If bError Then
            MsgBox "No se puede Borrar"

        End If

        Exit Sub

    End If

    Select Case stEvent

        Case "OK"
            bOK = True

        Case "ERROR"
            bError = True

        Case "RING"

            If bRing = False Then
                bRing = True

            End If

            iRingTime = Timer

        Case Else

            Select Case Left(stEvent, 4)

                Case "TIME"

                Case "DATE"

                Case "NMBR"

                Case "NAME"

            End Select
             
    End Select

End Sub

Private Sub Command3_Click()
    bOK = False
    bError = False
    Comm1.Output = "AT+CMGL=" & Chr(34) & "ALL" & Chr(34) & vbcrlf

    While Not bOK Or bError

        bMessageStore = True
        DoEvents
        Wait
    Wend

    If bOK Then
        ReadMessage

        If InStr(1, UCase(txtMessage.Text), "NOTEPAD", vbTextCompare) <> 0 Then
            Call ExecuteCommand("NotePad.exe")
        ElseIf InStr(1, UCase(txtMessage.Text), "CALC", vbTextCompare) <> 0 Then
            Call ExecuteCommand("Calc.exe")

        End If

    End If

    If bError Then
        txtMessage.Text = "Bad Read"

    End If

End Sub

Private Sub Wait()

    Dim Start

    Start = Timer

    Do While Timer < Start + 8
        DoEvents

        If bOK Then
            Exit Sub

        End If

        If bError Then
            Exit Sub

        End If

    Loop

End Sub

Private Sub WaitLong()

    Dim Start

    Start = Timer

    Do While Timer < Start + 36
        DoEvents

        If bOK Then
            Exit Sub

        End If

        If bError Then
            Exit Sub

        End If

    Loop

End Sub

Private Sub ReadMessage()

    If ParseFile Then
        msgBreak = Split(strMessageBuffer, vbcrlf, , vbTextCompare)
        msgHeader = Split(msgBreak(0), ",", , vbTextCompare)
        txtTelephone.Text = Mid$(Right$(msgHeader(2), 11), 1, 10)
        strMessageBuffer = ""

        For I = 1 To UBound(msgBreak(), 1)
            strMessageBuffer = strMessageBuffer & msgBreak(I) & vbcrlf
        Next I

        txtMessage.Text = strMessageBuffer
    Else
        txtMessage.Text = "Unable to decode Message"

    End If

End Sub

Private Sub Form_Load()
    bMessageStore = False

End Sub

Private Sub lstEvents_DblClick()
    lstEvents.Clear

End Sub

Public Function ParseFile() As Boolean

    Dim FirstOffSet  As Long

    Dim SecondOffSet As Long

    Dim strBuffer1   As String

    Dim strBuffer2   As String

    Dim strBuffer3   As String

    strBuffer1 = strMessageBuffer
    FirstOffSet = InStr(1, strBuffer1, "+CMGL:", vbTextCompare)
    SecondOffSet = InStr(1, strBuffer1, vbcrlf & "OK", vbTextCompare)

    If FirstOffSet <> 0 And SecondOffSet > FirstOffSet Then
        I = FirstOffSet

        While I < SecondOffSet

            strBuffer2 = strBuffer2 & Mid$(strBuffer1, I, 1)
            I = I + 1
        Wend
        ParseFile = True
        strMessageBuffer = strBuffer2
        Exit Function

    End If

    ParseFile = False

End Function

