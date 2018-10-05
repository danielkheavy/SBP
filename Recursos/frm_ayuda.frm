VERSION 5.00
Begin VB.Form frm_ayuda 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9090
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
   ScaleHeight     =   2355
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_salir 
      Height          =   555
      Left            =   7110
      Picture         =   "frm_ayuda.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1575
      Width           =   1350
   End
   Begin VB.Label lbl_info3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para buscar presione F1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   285
      TabIndex        =   2
      Top             =   1350
      Width           =   3195
   End
   Begin VB.Label lbl_info2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Transporte,Ubiquese en el cuadro que desea y presion F7."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   195
      TabIndex        =   1
      Top             =   810
      Width           =   7680
   End
   Begin VB.Label lbl_info1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Para Crear codigo Cliente / Proveedor y Vendedor,Producto,"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   210
      TabIndex        =   0
      Top             =   285
      Width           =   8040
   End
End
Attribute VB_Name = "frm_ayuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_salir_Click()
    Unload frm_ayuda

End Sub
