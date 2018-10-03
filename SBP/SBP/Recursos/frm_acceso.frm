VERSION 5.00
Object = "{19BD1EA6-6E36-45BA-AEBD-BCF3093017CC}#11.0#0"; "GorditoButton.ocx"
Begin VB.Form frm_acceso 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame FraClaveDe 
      BackColor       =   &H00808080&
      Caption         =   "Clave de Acceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   45
      Width           =   5100
      Begin VB.TextBox txt_acu 
         Height          =   285
         Left            =   3840
         TabIndex        =   17
         Top             =   975
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txt_estado 
         Height          =   285
         Left            =   3810
         TabIndex        =   16
         Top             =   540
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox clave 
         Alignment       =   2  'Center
         Height          =   585
         IMEMode         =   3  'DISABLE
         Left            =   210
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   870
         Width           =   3180
      End
      Begin GorditoButton.Boton Bot 
         Height          =   750
         Index           =   0
         Left            =   225
         TabIndex        =   2
         Top             =   1590
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "0"
         BackColor       =   4210752
         ResalteColor    =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin GorditoButton.Boton Bot 
         Height          =   750
         Index           =   1
         Left            =   990
         TabIndex        =   3
         Top             =   1590
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "1"
         BackColor       =   4210752
         ResalteColor    =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin GorditoButton.Boton Bot 
         Height          =   750
         Index           =   2
         Left            =   1800
         TabIndex        =   4
         Top             =   1590
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "2"
         BackColor       =   4210752
         ResalteColor    =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin GorditoButton.Boton Bot 
         Height          =   750
         Index           =   3
         Left            =   2595
         TabIndex        =   5
         Top             =   1590
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "3"
         BackColor       =   4210752
         ResalteColor    =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin GorditoButton.Boton Bot 
         Height          =   750
         Index           =   4
         Left            =   225
         TabIndex        =   6
         Top             =   2295
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "4"
         BackColor       =   4210752
         ResalteColor    =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin GorditoButton.Boton Bot 
         Height          =   750
         Index           =   5
         Left            =   1005
         TabIndex        =   7
         Top             =   2295
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "5"
         BackColor       =   4210752
         ResalteColor    =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin GorditoButton.Boton Bot 
         Height          =   750
         Index           =   6
         Left            =   1800
         TabIndex        =   8
         Top             =   2295
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "6"
         BackColor       =   4210752
         ResalteColor    =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin GorditoButton.Boton Bot 
         Height          =   750
         Index           =   7
         Left            =   2595
         TabIndex        =   9
         Top             =   2295
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "7"
         BackColor       =   4210752
         ResalteColor    =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin GorditoButton.Boton Bot 
         Height          =   750
         Index           =   8
         Left            =   240
         TabIndex        =   10
         Top             =   3000
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "8"
         BackColor       =   4210752
         ResalteColor    =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin GorditoButton.Boton Bot 
         Height          =   750
         Index           =   9
         Left            =   1005
         TabIndex        =   11
         Top             =   3000
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "9"
         BackColor       =   4210752
         ResalteColor    =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin GorditoButton.Boton BotCR 
         Height          =   750
         Index           =   10
         Left            =   1785
         TabIndex        =   12
         Top             =   2985
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "CR"
         BackColor       =   4210752
         ResalteColor    =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin GorditoButton.Boton BotOk 
         Height          =   1035
         Left            =   3660
         TabIndex        =   13
         Top             =   1545
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1826
         PicturePosition =   4
         Caption         =   "Ok"
         BackColor       =   4210752
         ResalteColor    =   49152
         PictureDown     =   "frm_acceso.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin GorditoButton.Boton BotX 
         Height          =   1035
         Left            =   3690
         TabIndex        =   14
         ToolTipText     =   "Cancelar"
         Top             =   2715
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1826
         PicturePosition =   4
         Caption         =   "X"
         BackColor       =   4210752
         ResalteColor    =   12632256
         PictureDown     =   "frm_acceso.frx":0E8A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
      End
      Begin VB.Label lblClaveDe 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave de Acceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   225
         TabIndex        =   15
         Top             =   405
         Width           =   3150
      End
   End
End
Attribute VB_Name = "frm_acceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Bot_Click(Index As Integer)

    If Index = 10 Then
        clave.Text = ""
        Exit Sub

    End If

    clave = clave + Bot(Index).Caption

End Sub

Private Sub BotCR_Click(Index As Integer)
    clave = ""

End Sub

Private Sub BotOk_Click()

    On Error GoTo cmd7_err

    Dim found As Integer

    Dim buf   As String

    If Len(clave) = 0 Then
        clave.SetFocus
        Exit Sub

    End If

    found = valida_clave("" & clave)

    If found = 0 Then
        MsgBox "Clave no valida para realizar este proceso ", 48, "Aviso"
        clave = ""
        clave.SetFocus
        Exit Sub
    Else
        txt_estado.Text = "1"
        'explorap.txt_found.Text = "1"
        Unload frm_acceso

    End If

    Exit Sub
cmd7_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    Unload Me
    Exit Sub

End Sub

Private Sub BotX_Click()
    'Unload Me
    'explorap.txt_found.Text = ""
    frm_acceso.Visible = False

End Sub

Private Sub Form_Activate()
    acu = frm_acceso.txt_acu

End Sub

Private Sub Form_Load()
    clave = ""

End Sub
