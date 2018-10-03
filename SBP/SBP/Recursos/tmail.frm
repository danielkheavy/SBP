VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form tmail 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envio de Correo"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   12765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   1200
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   240
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "EnviarCorreo"
      Height          =   735
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7920
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF80&
      Caption         =   "AdjuntarArchivo"
      Height          =   735
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7920
      Width           =   2895
   End
   Begin VB.TextBox mensaje 
      Height          =   4575
      Left            =   240
      MaxLength       =   200
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   3240
      Width           =   12375
   End
   Begin VB.TextBox asunto 
      Height          =   735
      Left            =   3120
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1080
      Width           =   9495
   End
   Begin VB.TextBox direccion 
      Height          =   735
      Left            =   3120
      MaxLength       =   50
      TabIndex        =   1
      Top             =   360
      Width           =   9495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mensaje :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label adjunto 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   5
      Top             =   1800
      Width           =   9495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Adjunto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Asunto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Direccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.Menu fdoo893 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  
Private Sub Command1_Click()

    With CommonDialog1
  
        .ShowOpen
          
        If .FileName = "" Then
            Exit Sub

        End If
      
        adjunto = .FileName
  
    End With

End Sub

Private Sub Command2_Click()

    On Error GoTo cmd9000_err

    'Borramos la ruta
    lblAdjunto = ""
  
    With MAPISession1
        .NewSession = False
        .SignOn

    End With
  
    With MAPIMessages1
        .SessionID = MAPISession1.SessionID
        ' Creamos el mensaje
        .Compose
  
        ' Asunto del mensaje
        .MsgSubject = asunto
  
        ' Mensaje
        .MsgNoteText = mensaje
  
        ' Nombre del Mail del destinatario
        .RecipDisplayName = direccion
  
        ' Archivo Adjunto
        If lblAdjunto <> "" Then
            .AttachmentPathName = adjunto

        End If
          
        ' Enviamos el correo
        .send False
  
    End With
  
    ' Cerramos la sesión abierta del Mapi
    MAPISession1.SignOff
    Exit Sub
cmd9000_err:
    MsgBox "No se puede enviar ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fdoo893_Click()
    tmail.Hide
    Unload tmail

End Sub

Private Sub Form_Load()

    'Me.Caption = " Enviar Email con Mapi "
    'cmdAdjunto.Caption = " Adjuntar archivo "
    'CmdEnviar.Caption = " Enviar mail "
End Sub
