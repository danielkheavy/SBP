VERSION 5.00
Begin VB.Form repfpago 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes Caja"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox ACU 
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
      MaxLength       =   1
      TabIndex        =   46
      Text            =   "V"
      Top             =   1200
      Width           =   1575
   End
   Begin VB.ComboBox tipoprint 
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
      TabIndex        =   42
      Top             =   4560
      Width           =   1575
   End
   Begin VB.ComboBox subconcepto 
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
      TabIndex        =   39
      Top             =   4200
      Width           =   2055
   End
   Begin VB.ComboBox concepto 
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
      TabIndex        =   38
      Top             =   3840
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
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
      TabIndex        =   36
      Top             =   5400
      Width           =   2055
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
      TabIndex        =   34
      Text            =   "%"
      Top             =   480
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
      TabIndex        =   32
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox caja 
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
      TabIndex        =   28
      Top             =   2880
      Width           =   2055
   End
   Begin VB.ComboBox turno 
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
      TabIndex        =   27
      Top             =   3240
      Width           =   2055
   End
   Begin VB.ComboBox cajero 
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
      TabIndex        =   26
      Top             =   2520
      Width           =   2055
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
      TabIndex        =   25
      Top             =   1560
      Width           =   3855
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
      TabIndex        =   11
      Top             =   3840
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
      TabIndex        =   10
      Text            =   "45"
      Top             =   6240
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
      TabIndex        =   9
      Text            =   "Reporte Formas de Pago"
      Top             =   5880
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
      TabIndex        =   8
      Top             =   3240
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
      TabIndex        =   7
      Top             =   2880
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
      TabIndex        =   6
      Text            =   "%"
      Top             =   840
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
      TabIndex        =   5
      Text            =   "%"
      Top             =   2040
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
      TabIndex        =   4
      Text            =   "%"
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox observa 
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
      TabIndex        =   3
      Text            =   "%"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.ComboBox vfpago 
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
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
   End
   Begin VB.ComboBox estado 
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
      TabIndex        =   1
      Top             =   5040
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
      TabIndex        =   0
      Text            =   "%"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Operación"
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
      TabIndex        =   45
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   "C.ompras V.entas R.ecibos T.odos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3960
      TabIndex        =   44
      Top             =   1200
      Width           =   3210
   End
   Begin VB.Label lblTipoImpresion 
      AutoSize        =   -1  'True
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
      TabIndex        =   43
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Subconcepto"
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
      TabIndex        =   41
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Concepto"
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
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Grupo"
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
      TabIndex        =   37
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label9 
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
      TabIndex        =   35
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label7 
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
      TabIndex        =   33
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caja"
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
      TabIndex        =   31
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label20 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Turno"
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
      TabIndex        =   30
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label22 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cajero"
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
      TabIndex        =   29
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
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
      TabIndex        =   24
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label4 
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
      TabIndex        =   23
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label2 
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
      TabIndex        =   22
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label Label17 
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
      TabIndex        =   21
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label5 
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
      TabIndex        =   20
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label10 
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
      TabIndex        =   19
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label11 
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
      TabIndex        =   18
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label21 
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
      TabIndex        =   17
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label3 
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
      TabIndex        =   16
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observacion"
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
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FormaPago"
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
      TabIndex        =   14
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado"
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
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label15 
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
      TabIndex        =   12
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Menu eju3453 
      Caption         =   "&Ejecutar"
   End
   Begin VB.Menu ldfo232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "repfpago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim omytablex As New ADODB.Recordset

Dim vSQL      As String

Dim mytablex  As New ADODB.Recordset

Function Formato_orden(Num_Campos As Integer, Nombre_Campos() As String) As Boolean

    Dim I As Integer

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Interior.color = RGB(192, 192, 250)

        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
       
        ''03/07/2017 KENYO 'Mejora Reporte de Formas de Pago (TC, Propina adicional
        .columns("A").ColumnWidth = 4
        .columns("B").ColumnWidth = 6
        .columns("C").ColumnWidth = 9
        .columns("D").ColumnWidth = 10
        .columns("E").ColumnWidth = 2
        
        .columns("F").ColumnWidth = 4
        .columns("G").ColumnWidth = 7
        
        .columns("H").ColumnWidth = 7
        .columns("I").ColumnWidth = 6
        .columns("J").ColumnWidth = 10
        .columns("K").ColumnWidth = 7
        .columns("L").ColumnWidth = 2
        
        .columns("M").ColumnWidth = 20
            
        .columns("N").ColumnWidth = 12
        .columns("O").ColumnWidth = 29.5
        .columns("P").ColumnWidth = 8
        .columns("Q").ColumnWidth = 8
        .columns("R").ColumnWidth = 8
        .columns("S").ColumnWidth = 20
        
        ''03/07/2017 KENYO 'Mejora Reporte de Formas de Pago (TC, Propina adicional
  
    End With

End Function

Sub reporte_excellfp(mytablex As ADODB.Recordset)

    Dim found       As Integer

    Dim buf         As String

    Dim buf1        As String

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim Heading(19) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    Dim v           As Long

    Dim h           As Long

    ''03/07/2017 KENYO 'Mejora Reporte de Formas de Pago (TC, Propina adicional
    Dim suma        As Double

    Dim prop        As Double

    suma = 0
    prop = 0
    ''03/07/2017 KENYO 'Mejora Reporte de Formas de Pago (TC, Propina adicional

    'Command1.Visible = True
    ' On Error GoTo cmd6561245_err
    'omytablex.Open "SELECT administrador as Autorizo,observa1 as Motivo,FechaBorra,Salon,Mesa,Vendedor,HoraBorra,Producto,Descripcio,
    'Unidad as Und,Cantidad as Cant,Precio,Total,Caja, Turno FROM logcomanda   " & buf & "  order by fecha,hora", cn, adOpenStatic, adLockOptimistic

    ''03/07/2017 KENYO 'Mejora Reporte de Formas de Pago (TC, Propina adicional
    Heading(1) = "TIPO":
    Heading(2) = "SERIE":
    Heading(3) = "NUMERO":
    Heading(4) = "FECHA":
    Heading(5) = "M"
    Heading(6) = "T/C":
    Heading(7) = "Propina":
    Heading(8) = "TOTAL":
    Heading(9) = "FPAGO":
    Heading(10) = "Descripcion":
    Heading(11) = "ORDEN":
    Heading(12) = "E":
    
    '24/04/2018 Reporte Formas de Pago Mejorado
    Heading(13) = "Obs. FormaPago":
    '24/04/2018 Reporte Formas de Pago Mejorado
    
    Heading(14) = "CÓDIGO":
    Heading(15) = "CLIENTE":
    Heading(16) = "USUARIO":
    Heading(17) = "CAJA":
    Heading(18) = "TURNO"
    
    '24/04/2018 Reporte Formas de Pago Mejorado
    Heading(19) = "Obs.Comprobante":
    '24/04/2018 Reporte Formas de Pago Mejorado
    
    ''03/07/2017 KENYO 'Mejora Reporte de Formas de Pago (TC, Propina adicional
           
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_orden(19, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    objExcel.ActiveSheet.Cells(1, 6) = "     REPORTE DE FORMAS DE PAGO"
    objExcel.ActiveSheet.Cells(1, 6).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 6).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 6).Font.color = RGB(0, 112, 184)
      
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA INICIO  " + fechai
    objExcel.ActiveSheet.Cells(2, 4) = "FECHA FIN  " + fechaf
    v = 4
    h = 1
    sdx1 = 0
    Do

        If mytablex.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("TIPO")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("SERIE")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("NUMERO")
        objExcel.ActiveSheet.Cells(v, h + 3) = "" & mytablex.Fields("FECHA")
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("MONEDA")
        
        ''03/07/2017 KENYO 'Mejora Reporte de Formas de Pago (TC, Propina adicional
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("PARIDAD")
        
        ''''28/09/2017 kenyo Mejora Reportes Forma de pago y Creditos Finanzas (Observa)
        '''27/07/2017 kenyo Testing Completo al Sistema
        'objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("OBSERVA") '
        If mytablex.Fields("FPAGO") <> "3" Then
            If Val("" & mytablex.Fields("DIAS")) > 0 Then
                objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("DIAS")

            End If

        End If
        
        '''27/07/2017 kenyo Testing Completo al Sistema
        ''''28/09/2017 kenyo Mejora Reportes Forma de pago y Creditos Finanzas (Observa)
     
        objExcel.ActiveSheet.Cells(v, h + 7) = "" & mytablex.Fields("TOTAL")
        objExcel.ActiveSheet.Cells(v, h + 8) = "" & mytablex.Fields("FPAGO")
        objExcel.ActiveSheet.Cells(v, h + 9) = "" & mytablex.Fields("DESCRIPCIO") 'nombre de pago
        objExcel.ActiveSheet.Cells(v, h + 10) = "'" & mytablex.Fields("ORDEN")
        objExcel.ActiveSheet.Cells(v, h + 11) = "" & mytablex.Fields("ESTADO")
        objExcel.ActiveSheet.Cells(v, h + 12) = "" & mytablex.Fields("OBSERVA")
        objExcel.ActiveSheet.Cells(v, h + 13) = "'" & mytablex.Fields("CODIGO")
        objExcel.ActiveSheet.Cells(v, h + 14) = "" & mytablex.Fields("NOMBRE")
        objExcel.ActiveSheet.Cells(v, h + 15) = "'" & mytablex.Fields("USUARIO")
        objExcel.ActiveSheet.Cells(v, h + 16) = mytablex.Fields("CAJA")
        objExcel.ActiveSheet.Cells(v, h + 17) = "" & mytablex.Fields("TURNO")
       
        '24/04/2018 Reporte Formas de Pago Mejorado
        objExcel.ActiveSheet.Cells(v, h + 18) = "" & ObtieneObservaComprobante(mytablex.Fields("local"), mytablex.Fields("TIPO"), mytablex.Fields("serie"), mytablex.Fields("numero"))
        '24/04/2018 Reporte Formas de Pago Mejorado
       
        '''27/07/2017 kenyo Testing Completo al Sistema
        
        'prop = prop + objExcel.ActiveSheet.Cells(v, h + 6)
        'objExcel.ActiveSheet.Cells(v + 1, h + 6) = prop
        '
        'suma = suma + objExcel.ActiveSheet.Cells(v, h + 7)
        'objExcel.ActiveSheet.Cells(v + 1, h + 7) = suma
     
        If Val((objExcel.ActiveSheet.Cells(v, h + 6))) > 0 Then
            prop = prop + objExcel.ActiveSheet.Cells(v, h + 6)
            objExcel.ActiveSheet.Cells(v + 1, h + 6) = prop

        End If
        
        suma = suma + objExcel.ActiveSheet.Cells(v, h + 7)
        objExcel.ActiveSheet.Cells(v + 1, h + 7) = suma
                 
        '''27/07/2017 kenyo Testing Completo al Sistema
        ''03/07/2017 KENYO 'Mejora Reporte de Formas de Pago (TC, Propina adicional
        v = v + 1
        'imprime_recetaa mytablex, v, h
   
        mytablex.MoveNext
    Loop

    objExcel.ActiveSheet.Cells(v, h + 5) = "GRAN TOTAL"
 
    Dim k As Integer

    For k = 5 To 8
        objExcel.ActiveSheet.Cells(v, h + k).Font.bold = True
        objExcel.ActiveSheet.Cells(v, h + k).Interior.color = RGB(248, 243, 53)
        objExcel.ActiveSheet.Cells(v, h + k).Interior.color = RGB(248, 243, 53)
        objExcel.ActiveSheet.Cells(v, h + k).Interior.color = RGB(248, 243, 53)
        objExcel.ActiveSheet.Cells(v, h + k).Interior.color = RGB(248, 243, 53)

    Next

    Set objExcel = Nothing

    'Exit Sub
    'cmd6561245_err:
    'MsgBox "Aviso en reporte orden " + error$, 48, "Aviso"
    'Exit Sub
End Sub

Function busca_familia(buf As String) As String

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select * from familia where familia='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_familia = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close
 
End Function

Function busca_paridad() As Double

    Dim sdx As Double

    sdx = 1

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from parame where codigo='01' ", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sdx = Val("" & mytablex.Fields("parivta"))

        If sdx <= 0 Then
            sdx = 1

        End If

    End If

    busca_paridad = sdx
    mytablex.Close

End Function

Function busca_subfamilia(buf As String, buf1 As String) As String

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from subfamil where  familia='" & buf & "' and subfamilia='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_subfamilia = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

Private Sub concepto_Click()

    If concepto = "%" Then Exit Sub
    carga_subconcepto extra_loquesea(concepto)

End Sub

Sub ReporteFormaPagoExcel()

    Dim found As Integer

    found = sql_documento(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    reporte_excellfp mytablex
    mytablex.Close

End Sub

Sub ReporteFormaPago()

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

    found = sql_documento(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento
    cuerpo_programa_documento mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Private Sub eju3453_Click()

    If tipoprint = "Excell" Then
      
        ReporteFormaPagoExcel
        Exit Sub
    Else
        ReporteFormaPago
        Exit Sub
        
    End If

End Sub

Private Sub Form_Activate()

    Dim mytablex As New ADODB.Recordset

    tipo.Clear
    tipo.AddItem "%"
    mytablex.Open "select * from tipo ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("tipodoc") = "W" Or "" & mytablex.Fields("tipodoc") = "V" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "D" Or "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "G" Then
            tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    '------------ otros
    caja.Clear
    caja.AddItem "%"
    mytablex.Open "select * from parameca ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        caja.AddItem "" & mytablex.Fields("caja")
        mytablex.MoveNext
    Loop
    mytablex.Close
    caja.ListIndex = 0

    turno.Clear
    turno.AddItem "%"
    mytablex.Open "select * from turno ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        turno.AddItem "" & mytablex.Fields("turno")
        mytablex.MoveNext
    Loop
    mytablex.Close
    turno.ListIndex = 0

    cajero.Clear
    cajero.AddItem "%"
    mytablex.Open "select * from vendedor ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    cajero.ListIndex = 0
    tipo.ListIndex = 0

    local1.Clear
    local1.AddItem "%"
    mytablex.Open "select * from tlocal ", cn, adOpenStatic, adLockOptimistic

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

    vfpago.Clear
    vfpago.AddItem "%"
    mytablex.Open "select * from fpago ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        vfpago.AddItem "" & mytablex.Fields("fpago") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    vfpago.ListIndex = 0

    concepto.Clear
    concepto.AddItem "%"
    mytablex.Open "select * from concepto ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        concepto.AddItem Trim("" & mytablex.Fields("concepto")) & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    concepto.ListIndex = 0

    subconcepto.Clear
    subconcepto.AddItem "%"
    subconcepto.ListIndex = 0

End Sub

Private Sub Form_Load()

    Dim mytablex As Table

    fechaf = Format(Now, "dd/mm/yyyy")
    fechai = "01" & "/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")

    tipoprint.Clear
    tipoprint.AddItem "Normal"
    tipoprint.AddItem "Excell"
    tipoprint.ListIndex = 1

    estado.AddItem "%"
    estado.AddItem "2"
    estado.AddItem "1"
    estado.AddItem "0"
    estado.ListIndex = 1
    moneda.Clear
    moneda.AddItem "%"
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0

    Combo2.AddItem "FORMAPAGO"
    Combo2.AddItem "VENDEDOR"
    Combo2.AddItem "TIPO"
    Combo2.AddItem "CODIGO"
    Combo2.ListIndex = 0

    ''03/07/2017 KENYO 'Mejora Reporte de Formas de Pago (TC, Propina adicional
    acu.Enabled = True

End Sub

'Reporte de ingresos (Cobranzas) CONTASIS
Private Sub Label17_Click()

    Dim found As Integer

    found = sql_documento(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    reporte_excellfpcontasis mytablex
    mytablex.Close

End Sub

'Reporte de ingresos (Cobranzas) CONTASIS

Private Sub ldfo232_Click()
    repfpago.Hide
    Unload repfpago

End Sub

Function sql_documento(mytablex As ADODB.Recordset)

    Dim buf    As String

    Dim xbuf   As String

    Dim xgrupo As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    buf = "select * from fpagov where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If tipo <> "%" Then
        xbuf = extra_loquesea(tipo)
        buf = buf & " and tipo like '" & xbuf & "'"

    End If

    If vpfago <> "%" Then
        xbuf = extra_loquesea(vfpago)
        buf = buf & " and fpago like '" & xbuf & "'"

    End If

    If local1 <> "%" Then
        buf = buf & " and local like '" & local1 & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    ''03/07/2017 KENYO 'Mejora Reporte de Formas de Pago (TC, Propina adicional 14/09/2017
  
    If acu <> "%" Then
        If acu = "V" Then
            buf = buf & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') "

        End If

        If acu = "C" Then
            buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O') "

        End If

        If acu = "R" Then
            buf = buf & " and (acu='V' or acu='W') "

        End If

    End If

    ''03/07/2017 KENYO 'Mejora Reporte de Formas de Pago (TC, Propina adicional  14/09/2017

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

    If observa <> "%" Then
        buf = buf & " and observa like '" & observa & "'"

    End If

    If cajero <> "%" Then
        xbuf = extra_loquesea(cajero)
        buf = buf & " and usuario like '" & xbuf & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "'"

    End If

    If concepto <> "%" Then
        buf = buf & " and concepto like '" & extra_loquesea(concepto) & "'"

    End If

    If subconcepto <> "%" Then
        buf = buf & " and subconcepto like '" & extra_loquesea(subconcepto) & "'"

    End If

    If estado <> "%" Then
        buf = buf & " and estado like '" & estado & "'"

    End If

    If Combo2 = "FORMAPAGO" Then
        xgrupo = "fpago,fecha"

    End If

    If Combo2 = "CODIGO" Then
        xgrupo = "codigo,fecha"

    End If

    If Combo2 = "VENDEDOR" Then
        xgrupo = "vendedor,fpago,fecha"

    End If

    If Combo2 = "TIPO" Then
        xgrupo = "tipo,fecha"

    End If

    buf = buf & " order by " & xgrupo & ",(numero)"
    'MsgBox buf
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    sql_documento = 1

End Function

Sub cabecera_documento()

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
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(170, "-")
    found = formateaa(buf, 170, 2, 0)
    found = formateaa("E", 2, 0, 0)
    found = formateaa("lo", 3, 0, 0)
    found = formateaa("Fecha", 11, 0, 0)
    found = formateaa("Tip", 4, 0, 0)
    found = formateaa("ser", 5, 0, 0)
    found = formateaa("Numero", 12, 0, 0)
    found = formateaa("Fpago", 10, 0, 0)
    found = formateaa("M", 2, 0, 0)
    found = formateaa("T/C", 6, 0, 0)
    found = formateaa("Prop", 5, 0, 0)
    found = formateaa("Ingreso ", 11, 0, 1)
    found = formateaa("Egreso ", 11, 0, 1)
    found = formateaa("------SA L D O--------", 22, 0, 0)
    found = formateaa("Observa", 12, 0, 0)
    found = formateaa("orden", 6, 0, 0)
    found = formateaa("Codigo", 22, 0, 0)
    found = formateaa("Vend ", 9, 0, 0)
    found = formateaa("Cajero ", 9, 0, 0)
    found = formateaa("Ca", 3, 0, 0)
    found = formateaa("T", 2, 2, 1)
    
    found = formateaa("", 78, 0, 0)
    found = formateaa(dicmoneda, 11, 0, 1)
    found = formateaa("Dolares ", 11, 2, 1)
    
    buf = String(170, "-")
    found = formateaa(buf, 170, 2, 0)

End Sub

Sub cuerpo_programa_documento(mytablex As ADODB.Recordset)

    Dim Tmp     As String

    Dim sw      As Integer

    Dim buf     As String

    Dim found   As Integer

    Dim sdx     As Double

    Dim paridad As Double

    Dim soles   As Double

    Dim dolares As Double

    Dim sdx1    As Double

    Dim xgrupo  As String

    Dim xgrupo1 As String

    sdx = 0
    sdx1 = 0
    sw = 0
    suma1 = 0
    suma2 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma4 = 0
    suma4 = 0
    Do

        If mytablex.EOF Then Exit Do
        If Combo2 = "FORMAPAGO" Then
            xgrupo1 = "" & mytablex.Fields("fpago")

        End If

        If Combo2 = "VENDEDOR" Then
            xgrupo1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo2 = "TIPO" Then
            xgrupo1 = "" & mytablex.Fields("TIPO")

        End If

        If Combo2 = "CODIGO" Then
            xgrupo1 = "" & mytablex.Fields("CODIGO")

        End If

        If sw = 0 Then
            If Combo2 = "FORMAPAGO" Then
                buf = "" & mytablex.Fields("fpago")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_fpago("" & mytablex.Fields("fpago"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                Tmp = "" & mytablex.Fields("fpago")

            End If

            If Combo2 = "VENDEDOR" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor("" & mytablex.Fields("vendedor"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo2 = "CODIGO" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 2, 0)
                'buf = busca_vendedor("" & mytablex.Fields("vendedor"))
                'found = formateaa(buf, 30, 0, 0)
                'found = formateaa("", 1, 2, 0)
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo2 = "TIPO" Then
                buf = "" & mytablex.Fields("tipo")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_tipo("" & mytablex.Fields("tipo"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                Tmp = "" & mytablex.Fields("tipo")

            End If

            nlineas
            sw = 1
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
   
        End If

        If Tmp <> xgrupo1 Then
            found = formateaa("", 50, 0, 0)
            buf = Format(suma4, "0.00")
            found = formateaa(buf, 9, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            buf = Format(suma1, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            buf = Format(suma2, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 2, 0)
   
            'found = formateaa("Soles", 10, 0, 0)
            'buf = Format(suma1, "0.00")
            'found = formateaa(buf, 10, 0, 1)
            'found = formateaa("", 1, 0, 0)
            'found = formateaa("Dolares", 10, 0, 0)
            'buf = Format(suma2, "0.00")
            'found = formateaa(buf, 10, 0, 1)
            'found = formateaa("", 1, 2, 0)
            nlineas
   
            If Combo2 = "FORMAPAGO" Then
                buf = "" & mytablex.Fields("fpago")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_fpago("" & mytablex.Fields("fpago"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                Tmp = "" & mytablex.Fields("fpago")

            End If

            If Combo2 = "VENDEDOR" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor("" & mytablex.Fields("vendedor"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo2 = "TIPO" Then
                buf = "" & mytablex.Fields("tipo")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_tipo("" & mytablex.Fields("tipo"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                Tmp = "" & mytablex.Fields("tipo")

            End If

            If Combo2 = "CODIGO" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                'found = formateaa("", 1, 0, 0)
                'buf = busca_tipo("" & mytablex.Fields("tipo"))
                'found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                Tmp = "" & mytablex.Fields("codigo")

            End If

            nlineas
            suma1 = 0

        End If

        buf = "" & mytablex.Fields("estado")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("local")
        found = formateaa(buf, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("fecha")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("tipo")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("serie")
        found = formateaa(buf, 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("numero")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("fpago") & "" & mytablex.Fields("descripcio")
        found = formateaa(buf, 9, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("moneda")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("paridad")
        found = formateaa(buf, 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Dias") 'propinas
        found = formateaa(buf, 4, 0, 1)
        found = formateaa("", 1, 0, 0)
   
        paridad = Val("" & mytablex.Fields("paridad"))

        If paridad = 0 Then
            paridad = 1

        End If

        soles = 0
        dolares = 0

        If "" & mytablex.Fields("moneda") = "S" Then
            soles = Val("" & mytablex.Fields("total"))
            dolares = Val("" & mytablex.Fields("total")) / paridad

        End If

        If "" & mytablex.Fields("moneda") = "D" Then
            dolares = Val("" & mytablex.Fields("total"))
            soles = Val("" & mytablex.Fields("total")) * paridad

        End If
   
        If "" & mytablex.Fields("acu") = "A" Or "" & mytablex.Fields("acu") = "B" Or "" & mytablex.Fields("acu") = "C" Or "" & mytablex.Fields("acu") = "D" Or "" & mytablex.Fields("acu") = "G" Or "" & mytablex.Fields("acu") = "W" Then
            buf = Format("" & mytablex.Fields("total"), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            found = formateaa("", 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            sdx = sdx + soles
            sdx1 = sdx1 + dolares
            suma4 = suma4 + Val("" & mytablex.Fields("dias"))
            ssuma4 = ssuma4 + Val("" & mytablex.Fields("dias"))
            suma1 = suma1 + Val("" & mytablex.Fields("total"))
            ssuma1 = ssuma1 + Val("" & mytablex.Fields("total"))
      
        End If

        If "" & mytablex.Fields("acu") = "J" Or "" & mytablex.Fields("acu") = "K" Or "" & mytablex.Fields("acu") = "L" Or "" & mytablex.Fields("acu") = "M" Or "" & mytablex.Fields("acu") = "P" Or "" & mytablex.Fields("acu") = "V" Then
            found = formateaa("", 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = Format("" & mytablex.Fields("total"), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            sdx = sdx - soles
            sdx1 = sdx1 - dolares
      
            suma2 = suma1 + Val("" & mytablex.Fields("total"))
            ssuma2 = ssuma1 + Val("" & mytablex.Fields("total"))

        End If
   
        buf = Format(sdx, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(sdx1, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = "" & mytablex.Fields("observa")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("orden")
        found = formateaa(buf, 5, 0, 0)
        found = formateaa("", 1, 0, 0)
      
        buf = "" & mytablex.Fields("codigo") & " " & mytablex.Fields("nombre")
        found = formateaa(buf, 21, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("vendedor")
        found = formateaa(buf, 8, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("usuario")
        found = formateaa(buf, 8, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("caja")
        found = formateaa(buf, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("turno")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas
        mytablex.MoveNext
    Loop
   
    found = formateaa("", 50, 0, 0)
    buf = Format(suma4, "0.00")
    found = formateaa(buf, 9, 0, 1)
    found = formateaa("", 1, 0, 0)
   
    buf = Format(suma1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
   
    buf = Format(suma2, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 50, 0, 0)
    buf = Format(suma4, "0.00")
    found = formateaa(buf, 9, 0, 1)
    found = formateaa("", 1, 0, 0)
   
    buf = Format(suma1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
   
    buf = Format(suma2, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
   
End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > Val(nrolineas) Then
        cabecera_documento

    End If

End Sub

Function busca_tipo(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from tipo where tipo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_tipo = "" & mytablex.Fields("descripcio")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_fpago(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from fpago where fpago=' " & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_fpago = "" & mytablex.Fields("descripcio")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_vendedor(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where codigo=' " & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_vendedor = "" & mytablex.Fields("nombre")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Sub carga_subconcepto(buf As String)

    Dim mytablex As New ADODB.Recordset

    subconcepto.Clear
    subconcepto.AddItem "%"
    mytablex.Open "select * from subconcepto where concepto='" & buf & "'", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        subconcepto.AddItem Trim("" & mytablex.Fields("subconcepto")) & "|" & mytablex.Fields("DESCRIPCIO")
        mytablex.MoveNext
    Loop
    mytablex.Close
    subconcepto.ListIndex = 0

End Sub

'24/04/2018 Reporte Formas de Pago Mejorado
Function ObtieneObservaComprobante(bxlocal As String, _
                                   bxtipo As String, _
                                   bxserie As String, _
                                   bxnumero As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT observa FROM factura   where local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        ObtieneObservaComprobante = ("" & mytablex.Fields("observa"))

    End If

    mytablex.Close

End Function

'24/04/2018 Reporte Formas de Pago Mejorado

'Reporte de ingresos (Cobranzas) CONTASIS
Sub reporte_excellfpcontasis(mytablex As ADODB.Recordset)

    Dim found       As Integer

    Dim buf         As String

    Dim buf1        As String

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim Heading(20) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    Dim v           As Long

    Dim h           As Long

    Dim suma        As Double

    Dim prop        As Double

    suma = 0
    prop = 0
    Heading(1) = "FECHA CANCELACION"
    Heading(2) = "DOCUMENTO"
    Heading(3) = "NUMERO"
    Heading(4) = "CUENTA CONTABLE"
    Heading(5) = "MONEDA"
    Heading(6) = "IMPORTE TOTAL"
    Heading(7) = "T.C"
    
    Heading(8) = "DOCUMENTO"
    Heading(9) = "NUMERO"
    Heading(10) = "FECHA DOCUMENTO"
    Heading(11) = "FECHA VENCIMIENTO"
    
    Heading(12) = "NRO DOC CLIENTE"
    Heading(13) = "APELLIDOS Y NOMBRES,RAZON SOCIAL"
    Heading(14) = "IMPORTE S/"
    Heading(15) = "IMPORTE US$."
    Heading(16) = "CUENTA CONTABLE"
    Heading(17) = "MEDIO DE PAGO"
    Heading(18) = "GLOSA"
    Heading(19) = "CENTRO DE COSTOS 1"
    Heading(20) = "CENTRO DE COSTOS 2"
        
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_ExcelRepCobranzasContasis(20, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    objExcel.ActiveSheet.Cells(1, 4) = "     REPORTE DE FORMAS DE PAGO - CONTASIS"
    objExcel.ActiveSheet.Cells(1, 4).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 4).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 4).Font.color = RGB(0, 112, 184)
      
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA INICIO  " + fechai
    objExcel.ActiveSheet.Cells(2, 4) = "FECHA FIN  " + fechaf
    v = 4
    h = 1
    sdx1 = 0
    Do

        If mytablex.EOF Then Exit Do
        
        objExcel.ActiveSheet.Cells(v, h + 1) = "'"
        objExcel.ActiveSheet.Cells(v, h + 2) = "'"
        objExcel.ActiveSheet.Cells(v, h + 3) = "'"
        objExcel.ActiveSheet.Cells(v, h + 4) = "'"
        objExcel.ActiveSheet.Cells(v, h + 5) = "'"
        objExcel.ActiveSheet.Cells(v, h + 6) = "'"
        objExcel.ActiveSheet.Cells(v, h + 7) = "'" & E_llenar_TipoDocumento(mytablex.Fields("tipo"))
        objExcel.ActiveSheet.Cells(v, h + 8) = "'" & mytablex.Fields("SERIE") & E_llenar_zero(8, mytablex.Fields("numero"))
              
        objExcel.ActiveSheet.Cells(v, h + 9) = "" & mytablex.Fields("FECHA")
        objExcel.ActiveSheet.Cells(v, h + 10) = "" & mytablex.Fields("FECHA")
        
        objExcel.ActiveSheet.Cells(v, h + 11) = "'" & E_llenar_Codigo(mytablex.Fields("CODIGO"), mytablex.Fields("estado"))
        objExcel.ActiveSheet.Cells(v, h + 12) = "'" & E_llenar_TipoRazonSocial(mytablex.Fields("codigo"), mytablex.Fields("nombre"), mytablex.Fields("estado"))
        objExcel.ActiveSheet.Cells(v, h + 13) = "" & mytablex.Fields("TOTAL")

        objExcel.ActiveSheet.Cells(v, h + 15) = "" & busca_CuentaContable(mytablex.Fields("fpago"))
        objExcel.ActiveSheet.Cells(v, h + 17) = "" & mytablex.Fields("OBSERVA")
  
        'objExcel.ActiveSheet.Cells(v, h + 12) = "" & mytablex.Fields("MONEDA")
        
        'objExcel.ActiveSheet.Cells(v, h + 5) = "" & mytablex.Fields("PARIDAD")
        
        '        If mytablex.Fields("FPAGO") <> "3" Then
        '            If Val("" & mytablex.Fields("DIAS")) > 0 Then
        '                objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("DIAS")
        '            End If
        '        End If

        'objExcel.ActiveSheet.Cells(v, h + 7) = "" & mytablex.Fields("TOTAL")
  
        '        objExcel.ActiveSheet.Cells(v, h + 8) = "" & mytablex.Fields("FPAGO")
        '        objExcel.ActiveSheet.Cells(v, h + 9) = "" & mytablex.Fields("DESCRIPCIO") 'nombre de pago
        '        objExcel.ActiveSheet.Cells(v, h + 10) = "'" & mytablex.Fields("ORDEN")
        '        objExcel.ActiveSheet.Cells(v, h + 11) = "" & mytablex.Fields("ESTADO")
        '        objExcel.ActiveSheet.Cells(v, h + 12) = "" & mytablex.Fields("OBSERVA")
        '
        '        objExcel.ActiveSheet.Cells(v, h + 15) = "'" & mytablex.Fields("USUARIO")
        '        objExcel.ActiveSheet.Cells(v, h + 16) = mytablex.Fields("CAJA")
        '        objExcel.ActiveSheet.Cells(v, h + 17) = "" & mytablex.Fields("TURNO")
        '
        '        objExcel.ActiveSheet.Cells(v, h + 18) = "" & ObtieneObservaComprobante(mytablex.Fields("local"), mytablex.Fields("TIPO"), mytablex.Fields("serie"), mytablex.Fields("numero"))
        '
        '        objExcel.ActiveSheet.Cells(v, h + 19) = "" & busca_CuentaContable(mytablex.Fields("fpago"))
        '
        '        If Val((objExcel.ActiveSheet.Cells(v, h + 6))) > 0 Then
        '            prop = prop + objExcel.ActiveSheet.Cells(v, h + 6)
        '            objExcel.ActiveSheet.Cells(v + 1, h + 6) = prop
        '        End If
      
        v = v + 1

        mytablex.MoveNext
    Loop

    Set objExcel = Nothing

End Sub

'Reporte de ingresos (Cobranzas) CONTASIS

'Reporte de ingresos (Cobranzas) CONTASIS
Public Function Formato_ExcelRepCobranzasContasis(Num_Campos As Integer, _
                                                  Nombre_Campos() As String) As Boolean

    With objExcel.ActiveSheet
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, 7)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 7)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, 7)).Interior.color = RGB(192, 192, 250)
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
 
            .columns("A").ColumnWidth = 20
            .columns("B").ColumnWidth = 13
            .columns("C").ColumnWidth = 13
            .columns("D").ColumnWidth = 18
            .columns("E").ColumnWidth = 10
            .columns("F").ColumnWidth = 15
            .columns("G").ColumnWidth = 8
            .columns("H").ColumnWidth = 15
            .columns("I").ColumnWidth = 15
            .columns("J").ColumnWidth = 20
            .columns("K").ColumnWidth = 20
            .columns("L").ColumnWidth = 20
            .columns("M").ColumnWidth = 50
            .columns("N").ColumnWidth = 15
            .columns("O").ColumnWidth = 15
            .columns("P").ColumnWidth = 17
            .columns("Q").ColumnWidth = 18
            .columns("R").ColumnWidth = 25
            .columns("S").ColumnWidth = 20
            .columns("T").ColumnWidth = 20
   
        Next

    End With
    
    With objExcel.ActiveSheet
        .Range(.Cells(3, 1), .Cells(3, 7)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 7)).Font.bold = True
        .Range(.Cells(3, 8), .Cells(3, 20)).Interior.color = RGB(192, 200, 200)

    End With
    
End Function

'Reporte de ingresos (Cobranzas) CONTASIS

'Reporte de ingresos (Cobranzas) CONTASIS
Function E_llenar_zero(hastaCuanto As Integer, myDato As String) As String

    Dim I   As Integer

    Dim max As Integer

    max = Len(myDato)

    For I = 1 To hastaCuanto - max
        myDato = "0" & myDato
    Next
    E_llenar_zero = myDato

End Function

Function E_llenar_TipoDocumento(myDato As String) As String

    If myDato = 1 Then 'Boleta
        E_llenar_TipoDocumento = "03"
    ElseIf myDato = 2 Then 'Factura
        E_llenar_TipoDocumento = "01"
    Else ' Otros
        E_llenar_TipoDocumento = "00"

    End If

End Function

Function E_llenar_TipoPersona(myDato As String) As String

    If Len(Trim(myDato)) = 8 Then  'Dni
        E_llenar_TipoPersona = "1"
    ElseIf Len(Trim(myDato)) = 11 Then  'RUC
        E_llenar_TipoPersona = "6"
    Else ' Otros
        E_llenar_TipoPersona = "0"

    End If

End Function

Function E_llenar_Codigo(myDato As String, myEstado As String) As String

    If myEstado = "1" Then
        E_llenar_Codigo = "99999999999"
    Else

        If E_llenar_TipoPersona(myDato) = 0 Then   'RUC
            E_llenar_Codigo = "00000000"
        Else
            E_llenar_Codigo = myDato

        End If

    End If
    
End Function

Function E_llenar_TipoRazonSocial(codigo As String, _
                                  nombre As String, _
                                  myEstado As String) As String

    If myEstado = "1" Then
        E_llenar_TipoRazonSocial = "VENTA ANULADA"
    Else

        If Len(Trim(codigo)) = 8 Then 'Dni
            E_llenar_TipoRazonSocial = nombre
        ElseIf Len(Trim(codigo)) = 11 Then  'RUC

            If Mid(codigo, 1, 2) = "10" Then
                E_llenar_TipoRazonSocial = busca_nombre_comas("" & nombre)
            Else
                E_llenar_TipoRazonSocial = nombre

            End If

        End If

    End If

End Function

Function busca_nombre_comas(buf As String) As String
    buf = Replace$(buf, " ", ",")
    busca_nombre_comas = buf

End Function

Function busca_CuentaContable(fpago As String) As String

    Dim mytabley As New ADODB.Recordset

    mytabley.Open "SELECT cuentacontable  FROM fpago where fpago='" & "" & fpago & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        busca_CuentaContable = "" & mytabley.Fields("cuentacontable")

    End If

    '------------------------------------- ------------
    mytabley.Close
 
End Function

Function busca_CondicionPago(locall As String, _
                             tipo As String, _
                             serie As String, _
                             Numero As String) As String

    Dim mytabley As New ADODB.Recordset

    mytabley.Open "SELECT top 1 fpago FROM fpagov where local='" & "" & locall & "' and tipo='" & "" & tipo & "' and serie='" & "" & serie & "' and numero='" & "" & Numero & "' order by fpago", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        If ("" & mytabley.Fields("fpago") = "3") Then
            busca_CondicionPago = "CRE"
        Else
            busca_CondicionPago = "CON"

        End If

    End If

    '------------------------------------- ------------
    mytabley.Close
 
End Function

'Reporte de ingresos (Cobranzas) CONTASIS

