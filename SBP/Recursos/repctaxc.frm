VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form repctaxc 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuentas Corrientes"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   12735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txt_busca 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6240
      MaxLength       =   13
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   2880
      Width           =   3690
   End
   Begin VB.ComboBox cmb_tipoImpresion 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txt_documento 
      Height          =   375
      Left            =   8040
      TabIndex        =   39
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox xtipo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   1920
      Width           =   1575
   End
   Begin VB.ComboBox local1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox nombre 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   31
      Text            =   "%"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   6000
      Width           =   3855
   End
   Begin VB.ComboBox veproducto 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ComboBox verecibo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox tiposaldo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   4800
      Width           =   1575
   End
   Begin VB.ComboBox tipofecha 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   4440
      Width           =   1575
   End
   Begin VB.ComboBox moneda 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox nrolineas 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   8
      Text            =   "45"
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox titulo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   7
      Text            =   "*** DEUDA POR CUENTA ***"
      Top             =   5280
      Width           =   3855
   End
   Begin VB.TextBox fechaf 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox fechai 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox serie 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "%"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox tipo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox numero 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   2
      Text            =   "%"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox codigo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   1
      Text            =   "%"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox vendedor 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   0
      Text            =   "%"
      Top             =   3960
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MSF_busca_cliente 
      Height          =   1695
      Left            =   4080
      TabIndex        =   43
      Top             =   3360
      Visible         =   0   'False
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   12648447
      BackColorFixed  =   14595201
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCliente 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   42
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label lblTipoImpresion 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Impresion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   40
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label xcuentaco1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6240
      TabIndex        =   38
      Top             =   6240
      Width           =   105
   End
   Begin VB.Label xcuentaco 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6240
      TabIndex        =   37
      Top             =   6000
      Width           =   105
   End
   Begin VB.Label Label16 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Creditos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   36
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agrupacion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ver Productos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   28
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ver Recibos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   26
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Saldo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo fecha"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label acu 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lineas x Pagina"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Titulo reporte"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label17 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label21 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Menu eki 
      Caption         =   "&Ejecuta"
   End
   Begin VB.Menu t 
      Caption         =   "&Ticket"
   End
   Begin VB.Menu lso3232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "repctaxc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim suma10           As Double

Dim suma11           As Double

Dim suma12           As Double

Dim suma13           As Double

Dim suma14           As Double

Dim suma15           As Double

Dim suma16           As Double

Dim ssuma10          As Double

Dim ssuma11          As Double

Dim ssuma12          As Double

Dim ssuma13          As Double

Dim ssuma14          As Double

Dim ssuma15          As Double

Dim ssuma16          As Double

Dim my_nombreCliente As String

Private Sub cmdReporteTicket_Click()

End Sub

Private Sub cmb_tipoImpresion_Click()

    If cmb_tipoImpresion.Text = "EXCELL" Or xtipo = "%" Or xtipo <> "%" Then
        cmb_tipoImpresion.ListIndex = 1
        txt_busca.Enabled = True

    End If

End Sub

Private Sub eki_Click()

    Dim found         As Integer

    Dim mytablex      As New ADODB.Recordset

    Dim buf           As String

    Dim salida        As Boolean

    Dim k             As Integer

    Dim my_codcliente As String

    Dim my_busqueda   As String

    Dim donde1        As String

    Dim donde2        As String

    Dim donde3        As String

    Dim donde4        As String

    Dim donde5        As String

    Dim donde6        As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0

    On Error GoTo eki

    If cmb_tipoImpresion = "NORMAL" Then
        found = sql_cuentaxc(mytablex)

        If found = 0 Then
            mytablex.Close
            Exit Sub

        End If

        FileName = globaldir & "\temporal\" & gusuario & ".txt"
        cerrar_archivo
        found = borra_nombre("" & FileName)
        Open FileName For Append As #1
        '------------------------------------
        cabecera_cuentaxc
        cuerpo_programa_cuentaxc mytablex
        '------------------------------------
        Close #1
        cerrar_archivo
        mytablex.Close
        'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
        'genver.Show 1
        found = valida_wordpad(FileName)

    End If

    If cmb_tipoImpresion = "EXCELL" And txt_busca.Visible = False Then

        If veproducto.Text = "N" Then
            Call CuentaXCobrar(my_struc_cuentac(), xcuentaco, fechai, fechaf, tipofecha, local1, tipo, serie, Numero, codigo, nombre, moneda, vendedor, xtipo, tiposaldo, Combo1, salida, k, "N", "")

            If salida = True Then
                Call AbreExcel
                Call carga_imagen
                Call titulo_cuentaXCobrar(fechai, moneda, fechaf)
                'aqui tengo que poner
                Call carga_cuentasXCobrar(my_struc_cuentac(), k)
                donde1 = "CXCobrarD"

            End If

        ElseIf veproducto.Text = "S" Then
            'moneda = "D"
            Call CXC_producto(my_struc_cuentac(), xcuentaco, fechai, fechaf, tipofecha, local1, tipo, serie, Numero, codigo, nombre, moneda, vendedor, xtipo, tiposaldo, Combo1, salida, k, my_codcliente)

            If salida = True Then
                Call AbreExcel
                Call carga_imagen
                Call titulo_CXC_Producto(fechai, moneda, fechaf)
                'aqui tengo que poner
                Call carga_CXC_producto(my_struc_cuentac(), k)
                donde5 = "CXCobrarProductoD"

            End If
    
        End If
  
    End If

    If cmb_tipoImpresion = "EXCELL" And txt_busca.Visible = True Then 'documentos corporativos
        my_busqueda = my_nombreCliente
        'MsgBox "my_busqueda" & my_busqueda
        Call b_codcliente(my_busqueda, my_codcliente, salida)
        Call DocuXCobrar(my_struc_cuentac(), xcuentaco, fechai, fechaf, tipofecha, local1, tipo, serie, Numero, codigo, nombre, moneda, vendedor, xtipo, tiposaldo, Combo1, salida, k, my_codcliente)

        If salida = True Then
            Call AbreExcel
            Call carga_imagen
            Call titulo_cXCobrar_cliente(fechai, moneda, fechaf)
            'aqui tengo que poner
            Call carga_cXCobrar_cliente(my_struc_cuentac(), k)
            donde3 = "CXCDocumentosD"

        End If

        '***en soles
        moneda = "S"
        Call CuentaXCobrar(my_struc_cuentac(), xcuentaco, fechai, fechaf, tipofecha, local1, tipo, serie, Numero, codigo, nombre, moneda, vendedor, xtipo, tiposaldo, Combo1, salida, k, "S", my_codcliente)

        If salida = True Then
     
            Call AbreExcel
            Call carga_imagen
            Call titulo_cuentaXCobrar(fechai, moneda, fechaf)
            'aqui tengo que poner
            Call carga_cuentasXCobrar(my_struc_cuentac(), k)
            donde4 = "CXCDocumentosS"
        Else '
            MsgBox "Datos no encontrados"
            Exit Sub

        End If

    End If

    k = 0
    Call read_caja(my_caja)
    Call Datos_Empresa(my_struc_datos_empresa(), my_caja, salida, 0)

    If salida = True Then
        Call carga_datos_empresa(my_struc_datos_empresa(), k)

    End If

    If donde1 = "CXCobrarD" Then
        Call cerra_excelR("CXCobrarD" & Format(Now, "HHMMSS"))
    ElseIf donde2 = "CXCobrarS" Then
        Call cerra_excelR("CXCobrarS" & Format(Now, "HHMMSS"))
    ElseIf donde3 = "CXCDocumentosD" Then
        Call cerra_excelR("CXCDocumentosD" & Format(Now, "HHMMSS"))
    ElseIf donde4 = "CXCDocumentosS" Then
        Call cerra_excelR("CXCDocumentosS" & Format(Now, "HHMMSS"))
    ElseIf donde5 = "CXCobrarProductoD" Then
        Call cerra_excelR("CXCobrarProductoD" & Format(Now, "HHMMSS"))

    End If

    ' -- guardar el libro
    'Screen.MousePointer = vbNormal

eki:

    Select Case Err.Number

        Case 1004

            Resume Next

        Case Else

            'MsgBox Err.Number & vbcrlf & Err.Description
    End Select

End Sub

Private Sub Form_Activate()

    Dim mytablex As New ADODB.Recordset

    xtipo.Clear
    xtipo.AddItem "CREDITO"
    xtipo.AddItem "ANTICIPO DINERO"
    xtipo.AddItem "DEPOSITO BANCO"
    xtipo.AddItem "ORDEN TRABAJO"
    xtipo.AddItem "%"
    xtipo.ListIndex = 4

    local1.Clear
    local1.AddItem "%"

    tipo.Clear
    tipo.AddItem "%"

    mytablex.Open "SELECT * FROM tipo", cn, adOpenKeyset, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("grupo") = acu Then
            tipo.AddItem "" & mytablex.Fields("tipo")

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    tipo.ListIndex = 0
    mytablex.Open "SELECT * FROM tlocal", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        local1.AddItem "" & mytablex.Fields("codigo")
        mytablex.MoveNext
    Loop
    mytablex.Close
    local1.ListIndex = 0

    If local1.ListCount = 2 Then
        local1.ListIndex = 1

    End If

End Sub

Private Sub Form_Load()
    Combo1.AddItem "Codigo"
    Combo1.AddItem "Vendedor"
    Combo1.AddItem "Zona"
    Combo1.ListIndex = 0

    tipofecha.AddItem "EMISION"
    tipofecha.AddItem "VENCIMIENTO"
    tipofecha.ListIndex = 0
    tiposaldo.AddItem "PENDIENTE"
    tiposaldo.AddItem "CANCELADO"
    tiposaldo.AddItem "%"
    tiposaldo.ListIndex = 2
    moneda.AddItem "%"
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0

    verecibo.AddItem "N"
    verecibo.AddItem "S"
    verecibo.ListIndex = 1

    veproducto.AddItem "N"
    veproducto.AddItem "S"
    veproducto.ListIndex = 0

    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")

    cmb_tipoImpresion.Clear
    cmb_tipoImpresion.AddItem "NORMAL"
    cmb_tipoImpresion.AddItem "EXCELL"
    cmb_tipoImpresion.ListIndex = 0

End Sub

Private Sub lso3232_Click()
    repctaxc.Hide
    Unload repctaxc

End Sub

Sub cabecera_cuentaxc()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 80, 2, 1)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(122, "-")
    found = formateaa(buf, 122, 2, 0)
    found = formateaa("Grp", 4, 0, 0)
    found = formateaa("Tip", 4, 0, 0)
    found = formateaa("Ser", 4, 0, 0)
    found = formateaa("Numero", 12, 0, 0)
    found = formateaa("FechaEmis.", 11, 0, 0)
    found = formateaa("Ct", 3, 0, 0)
    found = formateaa("Vendedor", 12, 0, 0)
    found = formateaa("M", 2, 0, 0)
    found = formateaa("Total ", 11, 0, 1)
    found = formateaa("Abonos ", 11, 0, 1)
    found = formateaa("Interes ", 11, 0, 1)
    found = formateaa("Saldo ", 11, 0, 1)
    'found = formateaa("Vence", 11, 0, 0)
    'found = formateaa("Emision", 11, 0, 0)
    found = formateaa("Dias", 5, 2, 0)
    'found = formateaa("L1 ", 11, 0, 1)
    'found = formateaa("L2 ", 11, 0, 1)
    'found = formateaa("L3 ", 11, 0, 1)
    'found = formateaa("L4 ", 11, 2, 1)
    buf = String(122, "-")
    found = formateaa(buf, 122, 2, 0)

End Sub

Sub cuerpo_programa_cuentaxc(mytablex As ADODB.Recordset)

    Dim Tmp   As String

    Dim tmp1  As String

    Dim sw    As Integer

    Dim buf   As String

    Dim found As Integer

    Dim sdx   As Double

    sdx = 0
    sw = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0
    suma7 = 0
    suma8 = 0
    suma9 = 0
    suma10 = 0
    suma11 = 0
    suma12 = 0
    suma13 = 0
    suma14 = 0
    suma15 = 0
    suma16 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0
    ssuma7 = 0
    ssuma8 = 0
    ssuma9 = 0
    ssuma10 = 0
    ssuma11 = 0
    ssuma12 = 0
    ssuma13 = 0
    ssuma14 = 0
    ssuma15 = 0
    ssuma16 = 0

    tmp1 = ""
    Do

        If mytablex.EOF Then Exit Do
        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("codigo")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If sw = 0 Then
            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombre(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_zona(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("zona")

            End If
   
            sw = 1
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0
            suma9 = 0
            suma10 = 0
            suma11 = 0
            suma12 = 0
            suma13 = 0
            suma14 = 0
            suma15 = 0
            suma16 = 0

        End If

        If Tmp <> tmp1 Then
            imprime_subtotal
   
            buf = String(122, "-")
            found = formateaa(buf, 122, 2, 0)
            nlineas

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombre(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_zona(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("zona")

            End If

            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0
            suma9 = 0
            suma10 = 0
            suma11 = 0
            suma12 = 0
            suma13 = 0
            suma14 = 0
            suma15 = 0
            suma16 = 0

        End If

        buf = "" & mytablex.Fields("grupo")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("tipo")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("serie")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("numero")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("fecha")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("cuota")
        found = formateaa(buf, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("vendedor")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("moneda")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Total")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("abono")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("interes")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("saldo")
        buf = Format(Val(buf), "0.00")
        'If Val(buf) = 0 Then
        '   buf = ""
        'End If
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        'buf = "" & mytablex.Fields("fechav")
        'found = formateaa(buf, 10, 0, 0)
        'found = formateaa("", 1, 0, 0)
        'buf = "" & mytablex.Fields("fecha")
        'found = formateaa(buf, 10, 0, 0)
        'found = formateaa("", 1, 0, 0)
        sdx = 0
   
        '''06/08/2017 kenyo Testing Completo al Sistema
        'If IsDate("" & fechaf) And IsDate("" & mytablex.Fields("fechav")) Then
        'If CVDate(fechaf) > CVDate("" & mytablex.Fields("fechav")) Then
        sdx = mytablex.Fields("dias")
        ' End If
        'End If
        '''06/08/2017 kenyo Testing Completo al Sistema
   
        buf = Format(sdx, "0")
        found = formateaa(buf, 4, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("c1")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("c2")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("c3")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("c4")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)

        nlineas

        'verificar recibe
        If verecibo = "S" Then
            ver_recibos mytablex

        End If

        If veproducto = "S" Then
            busca_producto mytablex

        End If
   
        'verificar productos
        If "" & mytablex.Fields("grupo") = "A" Then
            suma1 = suma1 + Val("" & mytablex.Fields("total"))
            ssuma1 = ssuma1 + Val("" & mytablex.Fields("total"))
            suma2 = suma2 + Val("" & mytablex.Fields("abono"))
            ssuma2 = ssuma2 + Val("" & mytablex.Fields("abono"))
            suma3 = suma3 + Val("" & mytablex.Fields("interes"))
            ssuma3 = ssuma3 + Val("" & mytablex.Fields("interes"))
            suma4 = suma4 + Val("" & mytablex.Fields("saldo"))
            ssuma4 = ssuma4 + Val("" & mytablex.Fields("saldo"))

        End If

        If "" & mytablex.Fields("grupo") = "D" Then
            suma5 = suma5 + Val("" & mytablex.Fields("total"))
            ssuma5 = ssuma5 + Val("" & mytablex.Fields("total"))
            suma6 = suma6 + Val("" & mytablex.Fields("abono"))
            ssuma6 = ssuma6 + Val("" & mytablex.Fields("abono"))
            suma7 = suma7 + Val("" & mytablex.Fields("interes"))
            ssuma7 = ssuma7 + Val("" & mytablex.Fields("interes"))
            suma8 = suma8 + Val("" & mytablex.Fields("saldo"))
            ssuma8 = ssuma8 + Val("" & mytablex.Fields("saldo"))

        End If

        If "" & mytablex.Fields("grupo") = "C" Then
            suma9 = suma9 + Val("" & mytablex.Fields("total"))
            ssuma9 = ssuma9 + Val("" & mytablex.Fields("total"))
            suma10 = suma10 + Val("" & mytablex.Fields("abono"))
            ssuma10 = ssuma10 + Val("" & mytablex.Fields("abono"))
            suma11 = suma11 + Val("" & mytablex.Fields("interes"))
            ssuma11 = ssuma11 + Val("" & mytablex.Fields("interes"))
            suma12 = suma12 + Val("" & mytablex.Fields("saldo"))
            ssuma12 = ssuma12 + Val("" & mytablex.Fields("saldo"))

        End If

        If "" & mytablex.Fields("grupo") = "O" Then
            suma13 = suma13 + Val("" & mytablex.Fields("total"))
            ssuma13 = ssuma13 + Val("" & mytablex.Fields("total"))
            suma14 = suma14 + Val("" & mytablex.Fields("abono"))
            ssuma14 = ssuma14 + Val("" & mytablex.Fields("abono"))
            suma15 = suma15 + Val("" & mytablex.Fields("interes"))
            ssuma15 = ssuma15 + Val("" & mytablex.Fields("interes"))
            suma16 = suma16 + Val("" & mytablex.Fields("saldo"))
            ssuma16 = ssuma16 + Val("" & mytablex.Fields("saldo"))

        End If
   
        'suma5 = suma5 + Val("" & mytablex.Fields("c1"))
        'ssuma5 = ssuma5 + Val("" & mytablex.Fields("c1"))
        'suma6 = suma6 + Val("" & mytablex.Fields("c2"))
        'ssuma6 = ssuma6 + Val("" & mytablex.Fields("c2"))
        'suma7 = suma7 + Val("" & mytablex.Fields("c3"))
        'ssuma7 = ssuma7 + Val("" & mytablex.Fields("c3"))
        'suma8 = suma8 + Val("" & mytablex.Fields("c4"))
        'ssuma8 = ssuma8 + Val("" & mytablex.Fields("c4"))
        mytablex.MoveNext
    Loop
    'found = formateaa("", 62, 0, 0)
    imprime_subtotal
    imprime_total
   
End Sub

Sub imprime_total()

    Dim buf   As String

    Dim found As Integer

    found = formateaa("---------Totales---------------", 41, 2, 0)
    nlineas

    If ssuma1 > 0 Then
        found = formateaa("Adelantos", 52, 0, 0)
        buf = Format(ssuma1, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma2, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma3, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma4, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 2, 1)
        nlineas

    End If

    If ssuma5 > 0 Then
   
        found = formateaa("Depositos", 52, 0, 0)
        buf = Format(ssuma5, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma6, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma7, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma8, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 2, 1)
        nlineas

    End If
   
    If ssuma9 > 0 Then
        found = formateaa("Creditos", 52, 0, 0)
        buf = Format(ssuma9, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma10, "0.00")
        buf = Format(-Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma11, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma12, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 2, 1)
        nlineas

    End If

    If ssuma13 > 0 Then
   
        found = formateaa("OrdenTrabajo", 52, 0, 0)
        buf = Format(ssuma13, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma14, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma15, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma16, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 2, 1)
        nlineas

    End If

End Sub

Sub imprime_subtotal()

    Dim buf   As String

    Dim found As Integer

    If suma1 > 0 Then
        found = formateaa("Adelantos", 52, 0, 0)
        buf = Format(suma1, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma2, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma3, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma4, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 2, 1)
        nlineas

    End If

    If suma5 > 0 Then
   
        found = formateaa("Depositos", 52, 0, 0)
        buf = Format(suma5, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma6, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma7, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma8, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 2, 1)
        nlineas

    End If
   
    If suma9 > 0 Then
        found = formateaa("Creditos", 52, 0, 0)
        buf = Format(suma9, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma10, "0.00")
        buf = Format(-Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma11, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma12, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 2, 1)
        nlineas

    End If

    If suma13 > 0 Then
   
        found = formateaa("OrdenTrabajo", 52, 0, 0)
        buf = Format(suma13, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma14, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma15, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma16, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 2, 1)
        nlineas

    End If
   
End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > Val(nrolineas) Then
        cabecera_cuentaxc

    End If

End Sub

Function busca_nombre(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    Dim buf1     As String

    buf1 = "clientes"

    If acu = "C" Then
        buf1 = "proveedo"

    End If

    mytablex.Open "SELECT * FROM " & buf1 & " where codigo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_nombre = "" & mytablex.Fields("nombre")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_vendedor(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM vendedor where codigo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_vendedor = "" & mytablex.Fields("nombre")

    End If

    mytablex.Close
 
End Function

Function busca_zona(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM zona where zona='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_zona = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close
 
End Function

Function sql_cuentaxc(mytablex As ADODB.Recordset)

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    'buf = "select * from cuentac where "
    'If acu = "C" Then
    'buf = "select * from cuentap where "
    'End If
    buf = "select * from " & xcuentaco & " where "

    If tipofecha = "EMISION" Then
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    End If

    If tipofecha = "VENCIMIENTO" Then
        buf = buf & "  fechav>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fechav<='" & Format(fechaf, "YYYYMMDD") & "' "
   
    End If

    If local1 <> "%" Then
        buf = buf & " and local='" & local1 & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & tipo & "'"

    End If

    If serie <> "%" Then
        buf = buf & " and serie like '" & serie & "'"

    End If

    If Numero <> "%" Then
        buf = buf & " and numero like '" & Numero & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    If xtipo = "CREDITO" Then
        buf = buf & " and grupo='C'"

    End If

    If xtipo = "ANTICIPO DINERO" Then
        buf = buf & " and grupo='A'"

    End If

    If xtipo = "DEPOSITO BANCO" Then
        buf = buf & " and grupo='D'"

    End If

    If xtipo = "ORDEN TRABAJO" Then
        buf = buf & " and grupo='O'"

    End If

    If tiposaldo = "PENDIENTE" Then
        buf = buf & " and (saldo>0 or saldo<0)"

    End If

    If tiposaldo = "CANCELADO" Then
        buf = buf & " and saldo=0"

    End If

    If Combo1 = "Codigo" Then
        buf = buf & " order by codigo,grupo,numero,fechav "

    End If

    If Combo1 = "Vendedor" Then
        buf = buf & " order by Vendedor,grupo,numero,fechav "

    End If

    If Combo1 = "Zona" Then
        buf = buf & " order by Zona,grupo,numero,fechav "

    End If

    'MsgBox buf
    mytablex.Open buf, cn, adOpenKeyset, adLockOptimistic
    sql_cuentaxc = 1

End Function

Sub ver_recibos(mytabley As ADODB.Recordset)  'ojo ver tipo de cliente

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    Dim buf      As String

    mytablex.Open "SELECT * FROM " & XCUENTACO1 & " where local1='" & "" & mytabley.Fields("local") & "' and tipo1='" & "" & mytabley.Fields("tipo") & "' and serie1='" & "" & mytabley.Fields("serie") & "' and numero1='" & "" & mytabley.Fields("numero") & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            '--------------------------------------------------------------
      
            found = formateaa("***", 4, 0, 0)
            buf = "" & mytablex.Fields("tipo")
            found = formateaa(buf, 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("serie")
            found = formateaa(buf, 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("numero")
            found = formateaa(buf, 11, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("fecha")
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 16, 0, 0)
            buf = "" & mytablex.Fields("moneda")
            found = formateaa(buf, 1, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("", 11, 0, 0)
            buf = Format(-Val("" & mytablex.Fields("paga")), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 2, 0)
            nlineas
            '--------------------------------------------------------------
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

End Sub

Sub busca_producto(mytabley As ADODB.Recordset)

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    Dim buf      As String

    '''20/01/2018 kenyo Testing General Sistema
    ' Mues
    'mytablex.Open "SELECT * FROM detalle where  local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "' and numero='" & "" & mytabley.Fields("numero") & "'", cn, adOpenKeyset, adLockOptimistic
    mytablex.Open "SELECT * FROM detalle where dua is null  and local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "' and numero='" & "" & mytabley.Fields("numero") & "'", cn, adOpenKeyset, adLockOptimistic
    '''20/01/2018 kenyo Testing General Sistema

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            '--------------------------------------------------------------
            found = formateaa("++++", 5, 0, 0)
            buf = "" & mytablex.Fields("tipo")
            found = formateaa(buf, 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("producto")
            found = formateaa(buf, 7, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("descripcio")
            found = formateaa(buf, 15, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("unidad")
            found = formateaa(buf, 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("factor")
            found = formateaa(buf, 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("cantidad")
            found = formateaa(buf, 5, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("precio")
            found = formateaa(buf, 7, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("total")
            found = formateaa(buf, 3, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            '--------------------------------------------------------------
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

End Sub

''' /04/2017 kenyo. Formato ticket de cuentas por cobrar y pagar en formato ticket
Sub cabecera_cuentaxcTICKET()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    
    cabecera_tipico2 "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 80, 2, 1)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    ''' /04/2017 kenyo. Formato ticket de cuentas por cobrar y pagar en formato ticket
    'buf = String(48, "-")
    buf = String(40, "-")
    ''' /04/2017 kenyo. Formato ticket de cuentas por cobrar y pagar en formato ticket
     
    found = formateaa(buf, 140, 2, 0)

    found = formateaa("Numero", 7, 0, 0)
    found = formateaa("Fecha", 5, 0, 0)
    ' found = formateaa("Ct", 3, 0, 0)
    ' found = formateaa("Vendedor", 12, 0, 0)
    'found = formateaa("M", 2, 0, 0)
    found = formateaa("Total", 9, 0, 1)
    found = formateaa("Abonos", 9, 0, 1)
    '  found = formateaa("Interes ", 11, 0, 1)
    found = formateaa("Saldo", 9, 0, 1)
    found = formateaa("", 0, 2, 1)
    'found = formateaa("Vence", 11, 0, 0)
    'found = formateaa("Emision", 11, 0, 0)
    'found = formateaa("Dias", 5, 2, 0)
    'found = formateaa("L1 ", 11, 0, 1)
    'found = formateaa("L2 ", 11, 0, 1)
    'found = formateaa("L3 ", 11, 0, 1)
    'found = formateaa("L4 ", 11, 2, 1)
    buf = String(48, "- ") ' AL
    found = formateaa(buf, 40, 2, 0)

End Sub

''' /04/2017 kenyo. Formato ticket de cuentas por cobrar y pagar en formato ticket

Sub cuerpo_programa_cuentaxcTICKET(mytablex As ADODB.Recordset)

    Dim Tmp   As String

    Dim tmp1  As String

    Dim sw    As Integer

    Dim buf   As String

    Dim found As Integer

    Dim sdx   As Double

    sdx = 0
    sw = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0
    suma7 = 0
    suma8 = 0
    suma9 = 0
    suma10 = 0
    suma11 = 0
    suma12 = 0
    suma13 = 0
    suma14 = 0
    suma15 = 0
    suma16 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0
    ssuma7 = 0
    ssuma8 = 0
    ssuma9 = 0
    ssuma10 = 0
    ssuma11 = 0
    ssuma12 = 0
    ssuma13 = 0
    ssuma14 = 0
    ssuma15 = 0
    ssuma16 = 0

    tmp1 = ""
    Do

        If mytablex.EOF Then Exit Do
        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("codigo")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If sw = 0 Then
            If Combo1 = "Codigo" Then
                found = formateaa("Cliente:", 9, 0, 0)
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombre(buf)
                found = formateaa(buf, 26, 2, 0)
                nlineasTICKET
                Tmp = "" & mytablex.Fields("codigo")
   
                'found = formateaa("", 1, 2, 0)
    
            End If
   
            sw = 1
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0
            suma9 = 0
            suma10 = 0
            suma11 = 0
            suma12 = 0
            suma13 = 0
            suma14 = 0
            suma15 = 0
            suma16 = 0

        End If

        If Tmp <> tmp1 Then
            imprime_subtotalTICKET
   
            buf = String(36, "-")
            found = formateaa(buf, 40, 2, 0)
            nlineasTICKET

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_nombre(buf)
                found = formateaa(buf, 31, 2, 0)
                nlineasTICKET
                Tmp = "" & mytablex.Fields("codigo")

            End If

            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0
            suma9 = 0
            suma10 = 0
            suma11 = 0
            suma12 = 0
            suma13 = 0
            suma14 = 0
            suma15 = 0
            suma16 = 0

        End If

        '   buf = "" & mytablex.Fields("grupo")
        '   found = formateaa(buf, 3, 0, 0)
        '   found = formateaa("", 1, 0, 0)
        '
        '   buf = "" & mytablex.Fields("tipo")
        '   found = formateaa(buf, 3, 0, 0)
        '   found = formateaa("", 1, 0, 0)
        '   buf = "" & mytablex.Fields("serie")
        '   found = formateaa(buf, 3, 0, 0)
        '   found = formateaa("", 1, 0, 0)
 
        found = formateaa("", 1, 2, 0)
        found = formateaa("", 1, 2, 0)
        found = formateaa("*", 1, 0, 0)
   
        buf = "" & mytablex.Fields("numero")
        found = formateaa(buf, 6, 0, 0)
        'found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("fecha")
        found = formateaa(buf, 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        'buf = "" & mytablex.Fields("cuota")
        'found = formateaa(buf, 2, 0, 0)
        ' found = formateaa("", 1, 0, 0)
        ' buf = "" & mytablex.Fields("vendedor")
        ' found = formateaa(buf, 11, 0, 0)
        ' found = formateaa("", 1, 0, 0)
        'buf = "" & mytablex.Fields("moneda")
        ' found = formateaa(buf, 1, 0, 0)
        ' found = formateaa("", 1, 0, 0)
   
        buf = " " & mytablex.Fields("Total")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("abono")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
   
        ' buf = "" & mytablex.Fields("interes")
        ' buf = Format(Val(buf), "0.00")
        ' If Val(buf) = 0 Then
        '   buf = ""
        '  End If
        ' found = formateaa(buf, 10, 0, 1)
        ' found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("saldo")
        buf = Format(Val(buf), "0.00")
        'If Val(buf) = 0 Then
        '   buf = ""
        'End If
        found = formateaa(buf, 8, 0, 1)
        'found = formateaa("", 1, 0, 0)
        'buf = "" & mytablex.Fields("fechav")
        'found = formateaa(buf, 10, 0, 0)
        'found = formateaa("", 1, 0, 0)
        'buf = "" & mytablex.Fields("fecha")
        'found = formateaa(buf, 10, 0, 0)
        'found = formateaa("", 1, 0, 0)
        sdx = 0

        If IsDate("" & fechaf) And IsDate("" & mytablex.Fields("fechav")) Then
            If CVDate(fechaf) > CVDate("" & mytablex.Fields("fechav")) Then
                sdx = DateValue("" & fechaf) - DateValue("" & mytablex.Fields("fechav"))

            End If

        End If

        'buf = Format(sdx, "0")
        'found = formateaa(buf, 4, 0, 0)
        '   found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("c1")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("c2")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("c3")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("c4")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
 
        found = formateaa("", 1, 2, 0)
 
        nlineasTICKET

        'verificar recibe
        If verecibo = "S" Then
            ver_recibosTICKET mytablex

        End If

        If veproducto = "S" Then
            found = formateaa(" ", 1, 2, 0)
            busca_productoTICKET mytablex

        End If
   
        'verificar productos
        If "" & mytablex.Fields("grupo") = "A" Then
            suma1 = suma1 + Val("" & mytablex.Fields("total"))
            ssuma1 = ssuma1 + Val("" & mytablex.Fields("total"))
            suma2 = suma2 + Val("" & mytablex.Fields("abono"))
            ssuma2 = ssuma2 + Val("" & mytablex.Fields("abono"))
            suma3 = suma3 + Val("" & mytablex.Fields("interes"))
            ssuma3 = ssuma3 + Val("" & mytablex.Fields("interes"))
            suma4 = suma4 + Val("" & mytablex.Fields("saldo"))
            ssuma4 = ssuma4 + Val("" & mytablex.Fields("saldo"))

        End If

        If "" & mytablex.Fields("grupo") = "D" Then
            suma5 = suma5 + Val("" & mytablex.Fields("total"))
            ssuma5 = ssuma5 + Val("" & mytablex.Fields("total"))
            suma6 = suma6 + Val("" & mytablex.Fields("abono"))
            ssuma6 = ssuma6 + Val("" & mytablex.Fields("abono"))
            suma7 = suma7 + Val("" & mytablex.Fields("interes"))
            ssuma7 = ssuma7 + Val("" & mytablex.Fields("interes"))
            suma8 = suma8 + Val("" & mytablex.Fields("saldo"))
            ssuma8 = ssuma8 + Val("" & mytablex.Fields("saldo"))

        End If

        If "" & mytablex.Fields("grupo") = "C" Then
            suma9 = suma9 + Val("" & mytablex.Fields("total"))
            ssuma9 = ssuma9 + Val("" & mytablex.Fields("total"))
            suma10 = suma10 + Val("" & mytablex.Fields("abono"))
            ssuma10 = ssuma10 + Val("" & mytablex.Fields("abono"))
            suma11 = suma11 + Val("" & mytablex.Fields("interes"))
            ssuma11 = ssuma11 + Val("" & mytablex.Fields("interes"))
            suma12 = suma12 + Val("" & mytablex.Fields("saldo"))
            ssuma12 = ssuma12 + Val("" & mytablex.Fields("saldo"))

        End If

        If "" & mytablex.Fields("grupo") = "O" Then
            suma13 = suma13 + Val("" & mytablex.Fields("total"))
            ssuma13 = ssuma13 + Val("" & mytablex.Fields("total"))
            suma14 = suma14 + Val("" & mytablex.Fields("abono"))
            ssuma14 = ssuma14 + Val("" & mytablex.Fields("abono"))
            suma15 = suma15 + Val("" & mytablex.Fields("interes"))
            ssuma15 = ssuma15 + Val("" & mytablex.Fields("interes"))
            suma16 = suma16 + Val("" & mytablex.Fields("saldo"))
            ssuma16 = ssuma16 + Val("" & mytablex.Fields("saldo"))

        End If
   
        'suma5 = suma5 + Val("" & mytablex.Fields("c1"))
        'ssuma5 = ssuma5 + Val("" & mytablex.Fields("c1"))
        'suma6 = suma6 + Val("" & mytablex.Fields("c2"))
        'ssuma6 = ssuma6 + Val("" & mytablex.Fields("c2"))
        'suma7 = suma7 + Val("" & mytablex.Fields("c3"))
        'ssuma7 = ssuma7 + Val("" & mytablex.Fields("c3"))
        'suma8 = suma8 + Val("" & mytablex.Fields("c4"))
        'ssuma8 = ssuma8 + Val("" & mytablex.Fields("c4"))
        mytablex.MoveNext
    Loop
    'found = formateaa("", 62, 0, 0)
    imprime_subtotalTICKET
    imprime_totalTICKET
   
End Sub

Sub imprime_totalTICKET()

    Dim buf   As String

    Dim found As Integer

    found = formateaa("---------------------------------------", 50, 2, 0)
    found = formateaa("--------------- TOTALES ---------------", 50, 2, 0)

    nlineasTICKET

    'If ssuma1 > 0 Then
    '
    ' found = formateaa("Adelantos", 52, 0, 0)
    '   buf = Format(ssuma1, "0.00")
    '   buf = Format(Val(buf), "0.00")
    '   If Val(buf) = 0 Then
    '      buf = ""
    '   End If
    '
    '   found = formateaa(buf, 10, 0, 1)
    '   found = formateaa("", 1, 0, 0)
    '   buf = Format(ssuma2, "0.00")
    '   buf = Format(Val(buf), "0.00")
    '   If Val(buf) = 0 Then
    '      buf = ""
    '   End If
    '
    '   found = formateaa(buf, 10, 0, 1)
    '   found = formateaa("", 1, 0, 0)
    '   buf = Format(ssuma3, "0.00")
    '   buf = Format(Val(buf), "0.00")
    '   If Val(buf) = 0 Then
    '      buf = ""
    '   End If
    '
    '   found = formateaa(buf, 10, 0, 1)
    '   found = formateaa("", 1, 0, 0)
    '   buf = Format(ssuma4, "0.00")
    '   buf = Format(Val(buf), "0.00")
    '   If Val(buf) = 0 Then
    '      buf = ""
    '   End If
    '   found = formateaa(buf, 10, 2, 1)
    '   nlineas
    '  End If
    '
    '
    '  If ssuma5 > 0 Then
    '   found = formateaa("Depositos", 52, 0, 0)
    '   buf = Format(ssuma5, "0.00")
    '   buf = Format(Val(buf), "0.00")
    '   If Val(buf) = 0 Then
    '      buf = ""
    '   End If
    '
    '   found = formateaa(buf, 8, 0, 1)
    '   found = formateaa("", 1, 0, 0)
    '   buf = Format(ssuma6, "0.00")
    '   buf = Format(Val(buf), "0.00")
    '   If Val(buf) = 0 Then
    '      buf = ""
    '   End If
    '
    '   found = formateaa(buf, 8, 0, 1)
    '   found = formateaa("", 1, 0, 0)
    '   buf = Format(ssuma7, "0.00")
    '   buf = Format(Val(buf), "0.00")
    '   If Val(buf) = 0 Then
    '      buf = ""
    '   End If
    '
    '   found = formateaa(buf, 8, 0, 1)
    '   found = formateaa("", 1, 0, 0)
    '   buf = Format(ssuma8, "0.00")
    '   buf = Format(Val(buf), "0.00")
    '   If Val(buf) = 0 Then
    '      buf = ""
    '   End If
    '   found = formateaa(buf, 8, 2, 1)
    '   nlineas
    '   End If
    '
    If ssuma9 > 0 Then
        found = formateaa("CRED. TOTAL", 12, 0, 0)
        buf = Format(ssuma9, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If
   
        'OKA
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 2, 0, 0)
        buf = Format(ssuma10, "0.00")
        buf = Format(-Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If
   
        'OKA
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma11, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If
      
        found = formateaa(buf, 1, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(ssuma12, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If
   
        found = formateaa(buf, 8, 2, 1)
        nlineasTICKET

    End If

    If ssuma13 > 0 Then
   
        '   found = formateaa("OrdenTrabajo", 8, 0, 0)
        '   buf = Format(ssuma13, "0.00")
        '   buf = Format(Val(buf), "0.00")
        '   If Val(buf) = 0 Then
        '      buf = ""
        '   End If
        '
        '
        '   found = formateaa(buf, 8, 0, 1)
        '   found = formateaa("", 1, 0, 0)
        '   buf = Format(ssuma14, "0.00")
        '   buf = Format(Val(buf), "0.00")
        '   If Val(buf) = 0 Then
        '      buf = ""
        '   End If
        '   found = formateaa(buf, 8, 0, 1)
        '   found = formateaa("", 1, 0, 0)
        '   buf = Format(ssuma15, "0.00")
        '   buf = Format(Val(buf), "0.00")
        '   If Val(buf) = 0 Then
        '      buf = ""
        '   End If
        '   found = formateaa(buf, 8, 0, 1)
        '   found = formateaa("", 1, 0, 0)
        '   buf = Format(ssuma16, "0.00")
        '   buf = Format(Val(buf), "0.00")
        '   If Val(buf) = 0 Then
        '      buf = ""
        '   End If
        '   found = formateaa(buf, 8, 2, 1)
        nlineasTICKET

    End If

End Sub

Sub imprime_subtotalTICKET()

    Dim buf   As String

    Dim found As Integer

    If suma1 > 0 Then
        found = formateaa("Adelantos", 52, 0, 0)
        buf = Format(suma1, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma2, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma3, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma4, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 2, 1)
        nlineasTICKET

    End If

    If suma5 > 0 Then
   
        found = formateaa("Depositos", 52, 0, 0)
        buf = Format(suma5, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma6, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma7, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma8, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 2, 1)
        nlineasTICKET

    End If
   
    If suma9 > 0 Then
        found = formateaa("Cred SubTotal", 12, 0, 0)
        buf = Format(suma9, "0.00")
        buf = Format(Val(buf), "0.00")
   
        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 9, 0, 1)
        'found = formateaa("", 1, 0, 0)
        buf = Format(suma10, "0.00")
        buf = Format(-Val(buf), "0.00")
   
        If Val(buf) = 0 Then
            buf = ""

        End If

        ' found = formateaa(buf, 10, 0, 1)
        '   found = formateaa("", 1, 0, 0)
        '   buf = Format(suma11, "0.00")
        '   buf = Format(Val(buf), "0.00")
        '   If Val(buf) = 0 Then
        '      buf = ""
        '   End If
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 2, 0, 0)
        buf = Format(suma12, "0.00")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 9, 2, 1)
        nlineasTICKET

    End If

    ' If suma13 > 0 Then
   
    '   found = formateaa("OrdenTrabajo", 52, 0, 0)
    '   buf = Format(suma13, "0.00")
    '   buf = Format(Val(buf), "0.00")
    '   If Val(buf) = 0 Then
    '      buf = ""
    '   End If
    '   found = formateaa(buf, 9, 0, 1)
    '   found = formateaa("", 1, 0, 0)
    '   buf = Format(suma14, "0.00")
    '   buf = Format(Val(buf), "0.00")
    '   If Val(buf) = 0 Then
    '      buf = ""
    '   End If
    '   found = formateaa(buf, 9, 0, 1)
    '   found = formateaa("", 1, 0, 0)
    '   buf = Format(suma15, "0.00")
    '   buf = Format(Val(buf), "0.00")
    '   If Val(buf) = 0 Then
    '      buf = ""
    '   End If
    '   found = formateaa(buf, 9, 0, 1)
    '   found = formateaa("", 1, 0, 0)
    '   buf = Format(suma16, "0.00")
    '   buf = Format(Val(buf), "0.00")
    '   If Val(buf) = 0 Then
    '      buf = ""
    '   End If
    '   found = formateaa(buf, 9, 2, 1)
    '   nlineas
    '   End If
   
End Sub

Sub nlineasTICKET()
    contlin = contlin + 1

    If contlin > Val(nrolineas) Then
        cabecera_cuentaxcTICKET

    End If

End Sub

Sub ver_recibosTICKET(mytabley As ADODB.Recordset)  'ojo ver tipo de cliente

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    Dim buf      As String

    mytablex.Open "SELECT * FROM " & XCUENTACO1 & " where local1='" & "" & mytabley.Fields("local") & "' and tipo1='" & "" & mytabley.Fields("tipo") & "' and serie1='" & "" & mytabley.Fields("serie") & "' and numero1='" & "" & mytabley.Fields("numero") & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            '--------------------------------------------------------------
      
            found = formateaa(" R>", 3, 0, 0)
            ' buf = "" & mytablex.Fields("tipo")
            ' found = formateaa(buf, 3, 0, 0)
            ' found = formateaa("", 1, 0, 0)
            'buf = "" & mytablex.Fields("serie")
            ' found = formateaa(buf, 3, 0, 0)
            ' found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("numero")
            found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("fecha")
            found = formateaa(buf, 5, 0, 0)
            found = formateaa("", 8, 0, 0)
            '      buf = "" & mytablex.Fields("moneda")
            '      found = formateaa(buf, 1, 0, 0)
            '      found = formateaa("", 1, 0, 0)
            '      found = formateaa("", 11, 0, 0)
            buf = Format(-Val("" & mytablex.Fields("paga")), "0.00")
            found = formateaa(buf, 8, 0, 1)
            found = formateaa("", 1, 2, 0)
            'found = formateaa("", 1, 2, 0)
            nlineasTICKET
            '--------------------------------------------------------------
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

End Sub

Sub busca_productoTICKET(mytabley As ADODB.Recordset)

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    Dim buf      As String

    '''20/01/2018 kenyo Testing General Sistema
    ' No muestra recetas de producto en VER PRODUCTOS
    'mytablex.Open "SELECT * FROM detalle where local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "' and numero='" & "" & mytabley.Fields("numero") & "'", cn, adOpenKeyset, adLockOptimistic
    mytablex.Open "SELECT * FROM detalle where dua is null  and local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "' and numero='" & "" & mytabley.Fields("numero") & "'", cn, adOpenKeyset, adLockOptimistic
    '''20/01/2018 kenyo Testing General Sistema

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            '--------------------------------------------------------------
            'found = formateaa("+", 1, 0, 0)
            '      buf = "" & mytablex.Fields("tipo")
            '      found = formateaa(buf, 3, 0, 0)
            '      found = formateaa("", 1, 0, 0)

            buf = "" & mytablex.Fields("producto")
            found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("descripcio")
            found = formateaa(buf, 15, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("unidad")
            found = formateaa(buf, 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            '      buf = "" & mytablex.Fields("factor")
            '      found = formateaa(buf, 3, 0, 0)
            '      found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("cantidad")
            found = formateaa(buf, 5, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("precio")
            found = formateaa(buf, 7, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("total")
            found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineasTICKET
            '--------------------------------------------------------------
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

End Sub

Private Sub MSF_busca_cliente_Click()

    Dim mio As Integer

    mio = MSF_busca_cliente.Row - 1

    my_nombreCliente = my_carga_busca_cliente(mio).nombre

End Sub

Private Sub t_Click()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0

    found = sql_cuentaxc(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_cuentaxcTICKET
    cuerpo_programa_cuentaxcTICKET mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)
    
End Sub

Private Sub txt_busca_KeyPress(KeyAscii As Integer)

    Dim found        As Integer

    Dim my_respuesta As String

    Dim salida       As Boolean

    Dim k            As Integer

    Dim tipoclie     As String

    If KeyAscii <> 13 Then Exit Sub
    Call busca_cliente("%" & UCase(txt_busca.Text) & "%", my_carga_busca_cliente(), salida, k)

    If salida = False Then
        MSF_busca_cliente.Visible = False
        MsgBox "Datos No encontrados", 64, "Clientes"
    Else
        MSF_busca_cliente.Visible = True
        MSF_busca_cliente.Visible = True
        Call ini_grid_bus_cliente(MSF_busca_cliente)
        Call carica_busca_cliente(MSF_busca_cliente, my_carga_busca_cliente(), k)

    End If

End Sub
