VERSION 5.00
Begin VB.Form Tsms 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Enviar Correos v1.0"
   ClientHeight    =   8325
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   10635
   DrawMode        =   4  'Mask Not Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   7320
      Width           =   5655
   End
   Begin VB.Frame Frame2 
      Caption         =   "Body"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5250
      Left            =   75
      TabIndex        =   10
      Top             =   1350
      Width           =   10410
      Begin VB.ComboBox txtselecciona 
         Height          =   315
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txthtml 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5040
         MaxLength       =   80
         TabIndex        =   28
         Top             =   960
         Width           =   2715
      End
      Begin VB.TextBox txtTo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         MaxLength       =   80
         TabIndex        =   24
         Top             =   700
         Width           =   2715
      End
      Begin VB.TextBox txtMsg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   1125
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   2400
         Width           =   6615
      End
      Begin VB.TextBox txtAttach 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5025
         MaxLength       =   80
         TabIndex        =   15
         Top             =   650
         Width           =   2715
      End
      Begin VB.TextBox txtFromEmail 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5025
         MaxLength       =   80
         TabIndex        =   14
         Top             =   225
         Width           =   2715
      End
      Begin VB.TextBox txtSubject 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         MaxLength       =   80
         TabIndex        =   13
         Top             =   1920
         Width           =   6615
      End
      Begin VB.TextBox txtFromName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         MaxLength       =   80
         TabIndex        =   12
         Top             =   300
         Width           =   2715
      End
      Begin VB.Label Label5 
         Caption         =   "Seleccionar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   29
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Attach Html"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Attachement"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   4050
         TabIndex        =   22
         Top             =   675
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   150
         TabIndex        =   21
         Top             =   705
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Subject"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   150
         TabIndex        =   20
         Top             =   1920
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   150
         TabIndex        =   19
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mensaje"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   18
         Top             =   2400
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde Email"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   4200
         TabIndex        =   17
         Top             =   225
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SMTP Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   75
      TabIndex        =   1
      Top             =   150
      Width           =   10410
      Begin VB.CheckBox chkSSL 
         Alignment       =   1  'Right Justify
         Caption         =   "Req. SSL"
         Height          =   315
         Left            =   2475
         TabIndex        =   11
         Top             =   675
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.TextBox txtServer 
         Height          =   300
         Left            =   687
         MaxLength       =   80
         TabIndex        =   5
         Text            =   "smtp.gmail.com"
         Top             =   300
         Width           =   1800
      End
      Begin VB.TextBox txtPort 
         Height          =   300
         Left            =   687
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "465"
         Top             =   690
         Width           =   600
      End
      Begin VB.TextBox txtUsername 
         Height          =   300
         Left            =   3321
         MaxLength       =   80
         TabIndex        =   3
         Top             =   300
         Width           =   1800
      End
      Begin VB.TextBox txtPassword 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5925
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   300
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   5178
         TabIndex        =   9
         Top             =   300
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   2544
         TabIndex        =   8
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   225
         TabIndex        =   7
         Top             =   675
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Enviar"
      Height          =   495
      Left            =   6675
      TabIndex        =   0
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seleccione Perfil"
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   7320
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   23
      Top             =   6720
      Width           =   6390
      WordWrap        =   -1  'True
   End
   Begin VB.Menu flo44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "Tsms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Procedure : SendMail
' Author    : coolcurrent4u
' Date      : 4/19/2011
' Purpose   : sends email using the cdo namespace
' TODO      : check for attachment existence before passing it
'           : Pass only number to port textfield or else it will throw an error
' Questions : Please ask in vbforums.com
'---------------------------------------------------------------------------------------
'

Private Sub cmdSend_Click()
    
    Dim retval     As String

    Dim objControl As Control

    'Validate first
    For Each objControl In Me.Controls

        If TypeOf objControl Is TextBox Then
            If Trim$(objControl.Text) = vbNullString And LCase$(objControl.Name) <> "txtattach" Then

                'Label2.Caption = "Error: All fields are required!"
                'Exit Sub
            End If

        End If

    Next
    'Send
    'MsgBox Trim(txtselecciona)
    Frame1.Enabled = False
    Frame2.Enabled = False
    cmdSend.Enabled = False
    Label2.Caption = "Sending..."
    retval = SendMail(Trim$(txtto.Text), Trim$(txtsubject.Text), Trim$(txtfromname.Text) & "<" & Trim$(txtfromemail.Text) & ">", Trim$(txtmsg.Text), Trim$(txtserver.Text), CInt(Trim$(txtport.Text)), Trim$(txtusername.Text), Trim$(txtpassword.Text), Trim$(txtattach.Text), CBool(chkssl.Value), Trim$(txtselecciona), Trim$(txthtml))
    Frame1.Enabled = True
    Frame2.Enabled = True
    cmdSend.Enabled = True
    Label2.Caption = IIf(retval = "ok", "Message sent!", retval)
    
End Sub

'Private Sub txtInfo_GotFocus(Index As Integer)
'    txtInfo(Index).SelStart = 0
'    txtInfo(Index).SelLength = Len(txtInfo(Index))
'End Sub

Private Sub Combo1_Click()

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    If Combo1 = "" Then Exit Sub
    buf = extra_loquesea1(Combo1)
    mytablex.Open "select * from correos where cosms='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        txtserver = Trim("" & mytablex.Fields("txtserver"))
        txtusername = Trim("" & mytablex.Fields("txtusername"))
        txtpassword = Trim("" & mytablex.Fields("txtpassword"))
        txtport = Trim("" & mytablex.Fields("txtport"))
        txtto = Trim("" & mytablex.Fields("txtto"))

        If Trim("" & mytablex.Fields("chkssl")) = "S" Then
            chkssl.Value = 1
        Else
            chkssl.Value = 0

        End If

        txtfromname = Trim("" & mytablex.Fields("txtfromname"))
        txtfromemail = Trim("" & mytablex.Fields("txtfromemail"))
        txtattach = Trim("" & mytablex.Fields("txtattach"))
        txtsubject = Trim("" & mytablex.Fields("txtsubject"))
        txtmsg = Trim("" & mytablex.Fields("txtmsg"))
        txthtml = Trim("" & mytablex.Fields("txthtml"))

        If Trim("" & mytablex.Fields("txtselecciona")) = "T" Then
            txtselecciona.ListIndex = 0

        End If

        If Trim("" & mytablex.Fields("txtselecciona")) = "H" Then
            txtselecciona.ListIndex = 1

        End If

    End If

    mytablex.Close

End Sub

Private Sub flo44_Click()
    Tsms.Hide
    Unload Tsms

End Sub

Private Sub Form_Load()
    txtselecciona.Clear
    txtselecciona.AddItem "T"
    txtselecciona.AddItem "H"
    txtselecciona.ListIndex = 0
    carga_config

End Sub

Sub carga_config()

    Dim mytablex As New ADODB.Recordset

    Combo1.Clear
    Combo1.AddItem ""
    mytablex.Open "select * from correos ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        Combo1.AddItem Trim("" & mytablex.Fields("Descripcio")) & "|" & Trim("" & mytablex.Fields("cosms"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    Combo1.ListIndex = 0

End Sub

