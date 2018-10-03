VERSION 5.00
Object = "{D4D26F6B-6564-44F4-A913-03C91CE37740}#2.1#0"; "reportman.ocx"
Begin VB.Form tresegui 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seguimientos"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin reportman.ReportManX Rep 
      Height          =   300
      Left            =   18120
      TabIndex        =   39
      Top             =   9360
      Width           =   1125
      filename        =   ""
      Preview         =   0   'False
      ShowProgress    =   0   'False
      ShowPrintDialog =   0   'False
      Title           =   ""
      Language        =   0
      DoubleBuffered  =   0   'False
      Enabled         =   -1  'True
      Object.Visible         =   -1  'True
      Cursor          =   0
      HelpType        =   0
      HelpKeyword     =   ""
      DefaultPrinter  =   "CAJA"
      AsyncExecution  =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ADMINISTRATIVOS"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6600
      Left            =   7080
      TabIndex        =   28
      Top             =   120
      Width           =   6360
      Begin VB.CommandButton Command39 
         Caption         =   "Consolidado de Comprobantes Electrónicos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   360
         TabIndex        =   49
         Top             =   3600
         Width           =   2655
      End
      Begin VB.CommandButton finanzas 
         Caption         =   "FINANZAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   48
         Top             =   3600
         Width           =   2655
      End
      Begin VB.CommandButton Command23 
         Caption         =   "REPORTE DELIVERY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   47
         Top             =   4320
         Width           =   2655
      End
      Begin VB.CommandButton cmdEntradasSalidas 
         Caption         =   "Entradas / Salidas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   46
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton cmdReporteGeneral 
         Caption         =   "Reporte General Creditos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   45
         Top             =   2000
         Width           =   2655
      End
      Begin VB.CommandButton cmdReporteGeneralTransacciones 
         Caption         =   "Reporte General Transacciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   44
         Top             =   2000
         Width           =   2655
      End
      Begin VB.CommandButton cmdGraficosComparativos 
         Caption         =   "Graficos Comparativos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   43
         Top             =   5900
         Width           =   2655
      End
      Begin VB.CommandButton cmdCPP 
         Caption         =   "Cuentas por Pagar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   42
         Top             =   2800
         Width           =   2655
      End
      Begin VB.CommandButton cmdCPC 
         Caption         =   "Cuentas por Cobrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   41
         Top             =   2800
         Width           =   2655
      End
      Begin VB.CommandButton Command38 
         Caption         =   "Saldo Actual por Mayor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4680
         TabIndex        =   40
         Top             =   6360
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton Command37 
         Caption         =   "PRODUCTOS POR VENCER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         TabIndex        =   38
         Top             =   6240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Reporte Kardex por Producto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   36
         Top             =   400
         Width           =   2655
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Reporte Kardex Sunat"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   35
         Top             =   400
         Width           =   2655
      End
      Begin VB.CommandButton Command28 
         Caption         =   "Comisiones Productos "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   34
         Top             =   4320
         Width           =   2655
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Graficos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   32
         Top             =   5900
         Width           =   2655
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Seguimiento de Productos por Comprobantes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   31
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton Command34 
         Caption         =   "Pedidos/OrdenTrabajo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   30
         Top             =   5040
         Width           =   2655
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Producto Comandadas Eliminada"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   29
         Top             =   5150
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "REPORTE - VENTAS"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   360
      TabIndex        =   19
      Top             =   4665
      Width           =   6210
      Begin VB.CommandButton Command29 
         Caption         =   "Productos Diarios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3195
         TabIndex        =   37
         Top             =   3435
         Width           =   2655
      End
      Begin VB.CommandButton Command32 
         Caption         =   "Servicios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3195
         TabIndex        =   33
         Top             =   2640
         Width           =   2655
      End
      Begin VB.CommandButton Command33 
         Caption         =   "SoloTotales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   27
         Top             =   3435
         Width           =   2655
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Ventas Diarias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   26
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Ventas Semanales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3195
         TabIndex        =   25
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Ventas Mensuales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   24
         Top             =   2640
         Width           =   2655
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Ventas Horas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3195
         TabIndex        =   23
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ventas Dias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   22
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Ventas x Vendedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3195
         TabIndex        =   21
         Top             =   450
         Width           =   2655
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Ventas Caja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   20
         Top             =   450
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "REPORTES PRINCIPALES"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4065
      Left            =   375
      TabIndex        =   10
      Top             =   240
      Width           =   6210
      Begin VB.CommandButton Command20 
         Caption         =   "Vendedor Productos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3195
         TabIndex        =   18
         Top             =   3000
         Width           =   2655
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Clientes Productos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   17
         Top             =   3000
         Width           =   2655
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Seguimiento de Comprobantes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3195
         TabIndex        =   16
         Top             =   2200
         Width           =   2655
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H80000000&
         Caption         =   "Saldo Actual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3180
         TabIndex        =   15
         Top             =   600
         Width           =   2655
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Ranking de Productos Vendidos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   14
         Top             =   600
         Width           =   2655
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Movimiento de Forma Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   13
         Top             =   2200
         Width           =   2655
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Registro de Compras"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3210
         TabIndex        =   12
         Top             =   1400
         Width           =   2655
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Registro de Ventas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   1395
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Centralizar Oficina Central"
      Height          =   615
      Left            =   15750
      TabIndex        =   9
      Top             =   5040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Unidades Vendidas Grupo"
      Height          =   615
      Left            =   15840
      TabIndex        =   8
      Top             =   5265
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Ventas 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Reporte Formato Ticket"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2550
      Left            =   7065
      TabIndex        =   1
      Top             =   6840
      Width           =   6375
      Begin VB.CommandButton Command22 
         Caption         =   "Vendedor Productos Ticket"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   7
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Productos Vs Documentos Tickets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3195
         TabIndex        =   6
         Top             =   1800
         Width           =   2655
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Documentos Emitidos Tickets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   345
         TabIndex        =   5
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Unidades Vendidas Tickets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3195
         TabIndex        =   4
         Top             =   1080
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Copia Cierre Caja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3195
         TabIndex        =   3
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cuadres Generales Tickets"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command35 
      Caption         =   "Clientes/productos"
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   9000
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Menu k88933 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tresegui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCPC_Click()
    repctaxc.xcuentaco = "cuentac"
    repctaxc.XCUENTACO1 = "cuentacd"
    repctaxc.acu = "V"
    repctaxc.Show 1

End Sub

Private Sub cmdCPP_Click()
    repctaxc.xcuentaco = "cuentap"
    repctaxc.XCUENTACO1 = "cuentapd"
    repctaxc.acu = "C"
    repctaxc.Show 1

End Sub

Private Sub cmdEntradasSalidas_Click()
    ''20/07/2017 kenyo reporte de entradas Salidas  en el modulo de reportes
    opcion2 = "44"
    repinv.excell.Visible = True

    repinv.Label15.Visible = True
    repinv.Label16.Visible = True
    repinv.fechai.Visible = True
    repinv.fechaf.Visible = True
    repinv.fechai.Enabled = True
    repinv.excell.Enabled = False

    repinv.Label36.Visible = True
    repinv.vesubfamilia.Visible = True

    repinv.Label33.Visible = True
    repinv.quecosto.Visible = True

    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
    repinv.ChkSaldoInicial.Visible = True
    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat

    repinv.Show 1
    ''20/07/2017 kenyo reporte de entradas Salidas  en el modulo de reportes

End Sub

Private Sub cmdGraficosComparativos_Click()
    FrmCharc.Show 1

End Sub

Private Sub cmdReporteGeneral_Click()
    FrmReporteGeneralCreditos.Show 1

End Sub

Private Sub cmdReporteGeneralTransacciones_Click()
    FrmReporteGeneralTrans.Show 1

End Sub

Private Sub Command1_Click()
    opcion1 = "5"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.tipoexterno.Visible = True
    tcuadrc1.numcuadre.Visible = True
    'tcuadrc1.flagdiario = "1"
    tcuadrc1.cajero = "%"
    tcuadrc1.caja = "%"
    tcuadrc1.turno = "%"
    tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.pantalla = "PANTALLA"

    ''01/07/2017 Kenyo Correcion cierre correlativo al realizar copia de cierre de caja
    tcuadrc1.opcioncierres = "N" ' no hace ciere correlativo
    
    ''01/07/2017 Kenyo Correcion cierre correlativo al realizar copia de cierre de caja

    tcuadrc1.Caption = "COPIA CIERRE DEL DIA"
    tcuadrc1.Show 1

End Sub

Private Sub Command10_Click()
    'gtra5gra_Click
    opcion2 = "10"   'analisis de ventas
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    'repdocum.Label18.Visible = False
    'repdocum.Combo1.Visible = False
    repdocum.vdetalle.Enabled = False
    repdocum.vfpago.Enabled = False
    repdocum.acu = "V"
    repdocum.donde = "HORA"
    repdocum.Combo3.Enabled = False
    repdocum.Show 1

End Sub

Private Sub Command11_Click()
    repfpago.Show 1

End Sub

Private Sub Command12_Click()
    opcion2 = "4"
    repinv.excell.Visible = True
    repinv.Label17.Visible = True
    repinv.Combo1.Visible = True
    repinv.Label25.Visible = True
    repinv.gcanti.Visible = True

    'INICIO 25/04/2017 KENYO VISUALIZA OPCION COSTO
    repinv.quecosto.Visible = True
    repinv.Label33.Visible = True
    'FIN 25/04/2017 KENYO VISUALIZA OPCION COSTO

    ''''20/09/2017 kenyo Reporte de Stock minimo Ticket
    repinv.Combo2.Visible = True
    repinv.Label35.Visible = True
    ''''20/09/2017 kenyo Reporte de Stock minimo Ticket

    repinv.Show 1

End Sub

Private Sub Command13_Click()
    FrmChart.acu = "V"
    FrmChart.Show 1

End Sub

Private Sub Command14_Click()
    opcion2 = "0"
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    repdocum.Label26.Visible = True
    repdocum.comopaga.Visible = True
    repdocum.acu = "V"

    ' ''10/10/2017 Reporte de Seguimiento de facturas En Excel
    repdocum.Combo3 = "EXCELL"
    repdocum.vfpago.BackColor = &H80FFFF
    repdocum.vdetalle.BackColor = &H80FFFF
    ' ''10/10/2017 Reporte de Seguimiento de facturas En Excel

    ''02/11/2017 Reporte de Seguimiento de facturas incluye delivery
    repdocum.Label33.Visible = True
    repdocum.vedelivery.Visible = True
    ''02/11/2017 Reporte de Seguimiento de facturas incluye delivery

    gofpago = "FPAGOV"
    repdocum.Show 1

End Sub

Private Sub Command15_Click()
    opcion2 = "1"
    repraped.acu = "V" 'PEDIDO
    repraped.xdata = "DETALLE"

    ' ''10/10/2017 Reporte de Seguimiento de facturas En Excel
    'repraped.donde = "CLIENTES"
    repraped.donde = "Producto"
    repraped.Combo3 = "EXCELL"
    ' ''10/10/2017 Reporte de Seguimiento de facturas En Excel

    repraped.Show 1

End Sub

Private Sub Command16_Click()
    'gtra5gra_Click
    opcion2 = "10"   'analisis de ventas
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    'repdocum.Label18.Visible = False
    'repdocum.Combo1.Visible = False
    repdocum.vdetalle.Enabled = False
    repdocum.vfpago.Enabled = False
    repdocum.acu = "V"
    repdocum.donde = "Vendedor"
    repdocum.Combo3.Enabled = False
    repdocum.Show 1

End Sub

Private Sub Command17_Click()
    opcion2 = "2"
    ''''13/09/2017 kenyo Mejora Reportes Familias Producto
    repraped.Combo3 = "EXCELL"
    ''''13/09/2017 kenyo Mejora Reportes Familias Producto
    repraped.Label12.Visible = True
    repraped.orden.Visible = True
    repraped.acu = "V" 'PEDIDO
    repraped.xdata = "Detalle"
    repraped.Show 1

End Sub

Private Sub Command18_Click()
    opcion2 = "13"   'analisis de ventas
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    'repdocum.Label18.Visible = False
    'repdocum.Combo1.Visible = False
    repdocum.vdetalle.Enabled = False
    repdocum.vfpago.Enabled = False
    repdocum.acu = "V"
    repdocum.Combo3.Enabled = False
    repdocum.Show 1

End Sub

Private Sub Command19_Click()
    opcion2 = "12"   'analisis de ventas
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    'repdocum.Label18.Visible = False
    'repdocum.Combo1.Visible = False
    repdocum.vdetalle.Enabled = False
    repdocum.vfpago.Enabled = False
    repdocum.acu = "V"
    repdocum.Combo3.Enabled = False
    repdocum.Show 1

End Sub

Private Sub Command2_Click()
    opcion2 = "1"
    opcion1 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True

    usuariopos = gusuario
    tcuadrc1.flagdiario = "1"
    tcuadrc1.cajero = "%"
    tcuadrc1.caja = "%"
    tcuadrc1.turno = "%"
    tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.Caption = "CUADRE PARCIAL DEL DIA"
    'vIMPRIMIR = 1
    tcuadrc1.Show 1

    'If vIMPRIMIR = 1 Then tcuadrc1.Command1_Click
End Sub

Private Sub Command20_Click()
    opcion2 = "2"
    repraped.Label12.Visible = True
    repraped.orden.Visible = True
    repraped.Combo3 = "EXCELL"
    repraped.acu = "V" 'PEDIDO
    repraped.xdata = "Detalle"
    repraped.donde = "VENDEDOR"
    repraped.Show 1

End Sub

Private Sub Command21_Click()
    opcion2 = "2"
    repraped.Label12.Visible = True
    repraped.orden.Visible = True
    repraped.acu = "V" 'PEDIDO
    repraped.Combo3 = "EXCELL"
    repraped.xdata = "Detalle"
    repraped.donde = "CLIENTES"
    repraped.Show 1

End Sub

Private Sub Command22_Click()
    opcion1 = "9"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    tcuadrc1.cajero = "%"
    tcuadrc1.caja = "%"
    tcuadrc1.turno = "%"
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.flagdiario = "1"

    tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
    tcuadrc1.Caption = "VENDEDOR PRODUCTOS"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1

End Sub

Private Sub Command23_Click()
    FrmReporteDelivery.Show 1

End Sub

Private Sub Command24_Click()
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    repdocrv.Label24.Visible = True
    repdocrv.consolidado.Visible = True
    repdocrv.titulo = "REGISTRO DE VENTAS " & dicmoneda
    repdocrv.acu = "C"
    repdocrv.Show 1

End Sub

Private Sub Command25_Click()
    opcion2 = "100"
    repinv.Label15.Visible = True
    repinv.Label16.Visible = True
    repinv.fechai.Visible = True
    repinv.fechaf.Visible = True

    '''09/10/2017 kenyo Testing Reportes
    repinv.fechai.BackColor = 8454143
    repinv.fechaf.BackColor = 8454143
    '''09/10/2017 kenyo Testing Reportes

    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
    repinv.ChkSaldoInicial.Visible = True
    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat

    repinv.Show 1

End Sub

Private Sub Command26_Click()
    logcoma.Show 1

End Sub

Private Sub Command27_Click()
    opcion2 = "14"   'analisis de ventas
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    'repdocum.Label18.Visible = False
    'repdocum.Combo1.Visible = False
    repdocum.vdetalle.Enabled = False
    repdocum.vfpago.Enabled = False
    repdocum.acu = "V"
    repdocum.Combo3.Enabled = False
    repdocum.Show 1

End Sub

Private Sub Command28_Click()
    opcion2 = "1"

    ''''13/09/2017 kenyo Reporte Comisiones Productos en Excel
    repraped.Combo1 = "Vendedor"
    repraped.Combo3 = "EXCELL"
    ''''13/09/2017 kenyo Reporte Comisiones Productos en Excel
    repraped.acu = "V" 'PEDIDO
    repraped.xdata = "DETALLE"
    repraped.Show 1

End Sub

Private Sub Command29_Click()
    opcion2 = "n"
    repraped.Label12.Visible = True
    repraped.orden.Visible = True
    repraped.acu = "V" 'PEDIDO
    repraped.xdata = "Detalle"
    repraped.Combo3.Enabled = False
    repraped.Show 1

End Sub

Private Sub Command3_Click()
    opcion1 = "3"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.flagdiario = "1"
    tcuadrc1.cajero = "%"
    tcuadrc1.caja = "%"
    tcuadrc1.turno = "%"
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
    tcuadrc1.Caption = "UNIDADES VENDIDAS"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    
    ''''13/09/2017 kenyo Mejor Reporte Ticket
    tcuadrc1.check3d1 = 1
    tcuadrc1.check3d2 = 0
    ''''13/09/2017 kenyo Mejor Reporte Ticket
    
    tcuadrc1.Show 1

End Sub

Private Sub Command30_Click()
    opcion1 = "10"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.flagdiario = "1"
    tcuadrc1.cajero = "%"
    tcuadrc1.caja = "%"
    tcuadrc1.turno = "%"
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
    tcuadrc1.Caption = "UNIDADES VENDIDAS"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1

End Sub

Private Sub Command31_Click()
    'If CDate(Format(Now, "dd/mm/yyyy")) < CDate("05/01/2014") Then
    tload.Show 1

    'End If
End Sub

Private Sub Command32_Click()
    opcion2 = "455"   'analisis de ventas
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    'repdocum.Label18.Visible = False
    'repdocum.Combo1.Visible = False
    repdocum.vdetalle.Enabled = False
    repdocum.vfpago.Enabled = False
    repdocum.acu = "V"
    repdocum.Combo3.Enabled = False
    repdocum.Show 1

End Sub

Private Sub Command33_Click()
    opcion2 = "456"   'analisis de ventas
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    'repdocum.Label18.Visible = False
    'repdocum.Combo1.Visible = False
    repdocum.vdetalle.Enabled = False
    repdocum.vfpago.Enabled = False
    repdocum.acu = "V"
    repdocum.Combo3.Enabled = False
    repdocum.Show 1

End Sub

Private Sub Command34_Click()
    cgusuario = "CPEDIDOV"
    dgusuariog = "DPEDIDOV"
    gofpago = "FPAGOV"
    opcion2 = "P"
    repdocum.acu = "I"
    repdocum.Show 1
    Exit Sub
    'opcion2 = "0"
    'cgusuario = "CPEDIDOV"
    'dgusuariog = "DPEDIDOV"
    'repdocum.Label26.Visible = True
    'repdocum.comopaga.Visible = True
    'repdocum.acu = "I"
    'gofpago = "FPAGOV"
    'repdocum.Show 1

End Sub

Private Sub Command35_Click()
    trepocli.Show 1

End Sub

Private Sub Command36_Click()
    Rep.Preview = True
    Rep.FileName = App.path & "\Reportes\Rpt_ProductosStockMinimoAgrupado.rep"

    Rep.SetDatabaseConnectionString "VISIORION", "Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=" & menup.vservidor & ";Database=" & basedatos & ";Uid=sa;pwd=" & clave_servidor & ";Persist Security Info=False"

    Rep.ShowPrintDialog = False

    On Error Resume Next ' SI NO TIENE DATOS

    Rep.Execute

End Sub

Private Sub Command37_Click()
    Rep.Preview = True
    'Establecemos la cadena de conexion de manera dinamica.
    'Rep.SetDatabaseConnectionString "SERVIDOR", "Persist Security Info=True;Driver=MySQL ODBC 5.1 Driver;SERVER=" & cServidor & ";UID=" & cUserName & ";PWD=" & cPassword & ";DATABASE=" & cBDatos & ";PORT=3306"

    'Rep.FileName = ruta & "\informes\Rpt_ProductosStockMinimo.rep"
    Rep.FileName = App.path & "\Reportes\Rpt_ProductosPorVencer.rep"
    'cn.Open "Driver={SQL Server};Server=" & menup.vservidor & ";Database=" & buf & " ;Uid=sa;pwd=" & clave_servidor & ""
    'Rep.SetDatabaseConnectionString ("RESTAURANT6", "Driver={SQL Server};Server=TESTING\SQL2008;Uid=sa;pwd=mastercard;")
    'ok
    'Rep.SetDatabaseConnectionString "VISIORION", "Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=TESTING\SQL2008;Database=RESTAURANT6;Uid=sa;pwd=mastercard;Persist Security Info=False"
    Rep.SetDatabaseConnectionString "VISIORION", "Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=" & menup.vservidor & ";Database=" & basedatos & ";Uid=sa;pwd=" & clave_servidor & ";Persist Security Info=False"

    'Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=TESTING\SQL2008;Database=RESTAURANT6;Uid=sa;pwd=mastercard;Persist Security Info=False
    'Rep.SetDatabaseConnectionString "RESTAURANT6", "Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=TESTING\SQL2008;Uid=sa;pwd=mastercard;Persist Security Info=False"

    Rep.ShowPrintDialog = False
    ' LLamamos al reporte seleccionado.

    ' Ejecutamos el informe
    'IF REP.CalcReport
    On Error Resume Next ' SI NO TIENE DATOS

    Rep.Execute

End Sub

Private Sub Command38_Click()
    Rep.Preview = True
    Rep.FileName = App.path & "\Reportes\Rpt_ProductosStockXUnidad.rep"

    Rep.SetDatabaseConnectionString "VISIORION", "Provider=SQLOLEDB.1;Integrated Security=SSPI;Server=" & menup.vservidor & ";Database=" & basedatos & ";Uid=sa;pwd=" & clave_servidor & ";Persist Security Info=False"

    Rep.ShowPrintDialog = False

    On Error Resume Next ' SI NO TIENE DATOS

    Rep.Execute

End Sub

Private Sub Command39_Click()
    ' Reporte Sunat Facturacion Electronica 28/04/2018
    FrmConsolidadoSunat.ordenado = "Codigo"
    FrmConsolidadoSunat.Show 1

    ' Reporte Sunat Facturacion Electronica 28/04/2018
End Sub

Private Sub Command4_Click()
    'gtra5gra_Clickfr
    opcion2 = "10"   'analisis de ventas
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    'repdocum.Label18.Visible = False
    'repdocum.Combo1.Visible = False
    repdocum.vdetalle.Enabled = False
    repdocum.vfpago.Enabled = False
    repdocum.acu = "V"
    repdocum.donde = "FECHA"
    repdocum.Combo3.Enabled = False
    repdocum.Show 1

End Sub

Private Sub Command5_Click()
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    repdocrv.Label24.Visible = True
    repdocrv.consolidado.Visible = True
    repdocrv.titulo = "REGISTRO DE VENTAS " & dicmoneda
    repdocrv.acu = "V"
    repdocrv.Show 1

End Sub

Private Sub Command6_Click()
    opcion2 = "1"
    repinv.Label15.Visible = True
    repinv.Label16.Visible = True
    repinv.fechai.Visible = True
    repinv.fechaf.Visible = True
    repinv.fechai.Enabled = True

    '''10/08/2017 kenyo Mejor Kardex Producto
    repinv.quecosto.Visible = True
    repinv.Label33.Visible = True

    '''10/08/2017 kenyo Mejor Kardex Producto

    '''09/10/2017 kenyo Testing Reportes
    repinv.fechai.BackColor = 8454143
    repinv.fechaf.BackColor = 8454143

    repinv.excell.Visible = True
    repinv.excell.Enabled = False
    repinv.Label36.Visible = True
    repinv.vesubfamilia.Visible = True

    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat
    repinv.ChkSaldoInicial.Visible = True
    ''' 16/01/2018 Stock Inicial en Kardex Producto y Sunat

    repinv.Show 1

End Sub

Private Sub Command7_Click()
    opcion1 = "2"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.flagdiario = "1"

    tcuadrc1.cajero = "%"
    tcuadrc1.caja = "%"
    tcuadrc1.turno = "%"
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
    tcuadrc1.Caption = "DOCUMENTOS EMITIDOS"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1

End Sub

Private Sub Command8_Click()
    opcion1 = "4"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    tcuadrc1.cajero = "%"
    tcuadrc1.caja = "%"
    tcuadrc1.turno = "%"
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.flagdiario = "1"

    tcuadrc1.fechai = Format(Now, "dd/mm/yyyy")
    tcuadrc1.fechaf = Format(Now, "dd/mm/yyyy")
    tcuadrc1.Caption = "PRODUCTOS VS DOCUMENTOS"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1

End Sub

Private Sub Command9_Click()
    'gtra5gra_Click
    opcion2 = "10"   'analisis de ventas
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    'repdocum.Label18.Visible = False
    'repdocum.Combo1.Visible = False
    repdocum.vdetalle.Enabled = False
    repdocum.vfpago.Enabled = False

    repdocum.Combo3.Enabled = False

    repdocum.acu = "V"
    repdocum.Show 1

End Sub

Private Sub finanzas_Click()

    '24/04/2018 Reporte Finanzas Mejorado
    texplcxc.xcuentaco = "cuentac"
    texplcxc.XCUENTACO1 = "cuentacd"
    texplcxc.xcuentacol = ""
    texplcxc.ldo232.Enabled = False
    texplcxc.mofdi782.Enabled = False
    texplcxc.dj333.Enabled = False
    texplcxc.dj7823.Enabled = False
    texplcxc.ncu773.Enabled = False
    texplcxc.ordenado = "Codigo"

    texplcxc.acu = "V"
    texplcxc.Show 1

    '24/04/2018 Reporte Finanzas Mejorado
End Sub

Private Sub k88933_Click()
    tresegui.Hide
    Unload tresegui

End Sub

