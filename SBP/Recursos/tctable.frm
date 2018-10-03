VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tctable 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Cuentas Contables "
   ClientHeight    =   8910
   ClientLeft      =   150
   ClientTop       =   0
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   12930
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consulta"
      Height          =   8175
      Left            =   30
      TabIndex        =   30
      Top             =   45
      Width           =   12495
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7695
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   13573
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   22
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Descripcio"
            Caption         =   "Descripcio"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Cuenta"
            Caption         =   "Cuenta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "tipocuenta"
            Caption         =   "Tipocuenta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "NivelCuenta"
            Caption         =   "NivelCuenta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "TipoAnalisis"
            Caption         =   "TipoAnalisis"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Moneda"
            Caption         =   "M"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   5940.284
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1425.26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1665.071
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame2"
      Height          =   8535
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   12735
      Begin VB.ComboBox nmoneda 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3480
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox moneda 
         Enabled         =   0   'False
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
         Left            =   4920
         MaxLength       =   1
         TabIndex        =   27
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox ntipoanalisis 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox tipoanalisis 
         Enabled         =   0   'False
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
         Left            =   4920
         MaxLength       =   7
         TabIndex        =   24
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox nivelcuenta 
         Enabled         =   0   'False
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
         Left            =   4920
         MaxLength       =   7
         TabIndex        =   23
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox tipocuenta 
         Enabled         =   0   'False
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
         Left            =   4920
         MaxLength       =   8
         TabIndex        =   22
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox nnivelcuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1680
         Width           =   2535
      End
      Begin VB.ComboBox ntipocuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox cuenta 
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
         MaxLength       =   14
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox descripcio 
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
         MaxLength       =   50
         TabIndex        =   12
         Top             =   600
         Width           =   6495
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   10680
         Picture         =   "tctable.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Imprimir todo"
         Top             =   1320
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   10680
         Picture         =   "tctable.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label3 
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
         TabIndex        =   29
         Top             =   3480
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Analisis"
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
         TabIndex        =   26
         Top             =   2040
         Width           =   4815
      End
      Begin VB.Label Label26 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1-3 Activos 4.Pasivos 5.Capital  6.Gastos 7.Ventas"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   8040
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solo las cuentas de detalle sirven para generar asientos"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   7680
         Width           =   9735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Cuenta"
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
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
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
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo de Cuenta"
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
         Top             =   1320
         Width           =   4815
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nivel de Cuenta"
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
         Top             =   1680
         Width           =   4815
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   12435
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tctable.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox buffer 
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
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10560
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tctable.frx":23A6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Borrar registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tctable.frx":35B8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tctable.frx":47CA
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdAddEntry 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         Picture         =   "tctable.frx":59DC
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Menu ajdu1 
      Caption         =   "&Add"
   End
   Begin VB.Menu f8443 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu bo712 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu fjh433 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
      Begin VB.Menu fk8944 
         Caption         =   "&0.Excell"
      End
      Begin VB.Menu nhreyr 
         Caption         =   "&1.Generico"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu dfju773 
         Caption         =   "&2.Generador"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu cu7833 
      Caption         =   "&CtasPredefinidas"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu ui78232 
         Caption         =   "&1.Utilizar Cuentas Predefindas"
      End
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tctable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txctacon As New ADODB.Recordset

'1.activo y cuentas regualadoras de activo
'  esta constituido por todos los bienes y derechos de la empresa
'1.1 activo  circulante
'lo forman las cuentas disponibles ,exigible,y realizables a un plazo no mayor de un ano o repre-
'sentados por:
'1.1-01-01 caja
'representa en efectivo disponible(monedas,billetes cheques) en un momento determinado.resume a nivel
'general el movimiento de las subcuentas de caja
'2 pasivo obligaciones reales y contingentes de la empresa a favorde terceros
'3 inversion de los accionistas
'4 Ingresos refleja el total de los ingresos percibidos por ventas realizadas al contado y cobros a los clientes
'5 Costos Representa todos los egresos en que se incurre en el proceso productivo en los diferentes
'centros de produccion.en el catalogo de cuentas en el rubro de costos se representan clasificados
'como costos variables de fabricacion y costos fijos de fabricacion
'tambien se incluyen los gastos administrativos del departamento de produccion en la cuenta 5401 y subcuentas
'6 gastos .se registran todos los gastos de las ventas ,distribucion ,administrativos,financieros,de
'mantenimiento ,de recursos humanos ,etc en que se incurre como funcionamiento normal de la empresa
'y que por su naturaleza no se debe aplicar directamente a los costos.

Private Sub ajdu1_Click()

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Nuevo"
    cmdGuardar.Enabled = True
    habilita 1
    cuenta.Enabled = True
    cuenta = ""
    cuenta.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    buf = txctacon.Fields("cuenta")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + txctacon.Fields("cuenta"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    txctacon.Delete
    Command1_Click

    Exit Sub
cmd656_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    Command1_Click

End Sub

Private Sub cmdAddEntry_Click()
    ajdu1_Click

End Sub

Private Sub cmdCerrar_Click()
    dlo132_Click

End Sub

Private Sub cmdDelete_Click()
    bo712_Click

End Sub

Private Sub cmdExit_Click()
    dlo132_Click

End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdGuardar_Click()

    Dim found As Integer

    found = grabar()

End Sub

Private Sub cmdPrint_Click()
    fk8944_Click

    'djuer1_Click
End Sub

Private Sub cmdSave_Click()
    f8443_Click

End Sub

Private Sub cuenta_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(cuenta) = 0 Then Exit Sub
    descripcio.SetFocus

End Sub

Private Sub Command1_Click()
    Frame1.Visible = True
    Frame1.Enabled = True
    'buffer = ""
    opcion1 = "1"
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim cad As String

    If opcion1 = "1" Then  'bodega
        If Len(buffer) = 0 Then
            cad = "SELECT * from cuentas  order by cuenta  "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT *  from cuentas   where  " & Combo1 & " like '" & buffer & "%' order by cuenta"

        End If

        If txctacon.State = 1 Then txctacon.Close
        txctacon.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbGrid1.DataSource = txctacon
        dbGrid1.columns(0).Width = 4000
        dbGrid1.columns(1).Width = 2000

        If txctacon.RecordCount > 0 Then
            dbGrid1.SetFocus

        End If

    End If

End Sub

Private Sub Command2_Click()

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'codcta = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'codcta.SetFocus
        'codcta_KeyPress 13
    End If

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)

    Dim buf  As String

    Dim buf2 As String

    If KeyAscii <> 13 And KeyAscii <> 27 Then
        If KeyAscii = 8 Then
            If Len(buffer) > 0 Then
                buf = Mid$(buffer, 1, Len(buffer) - 1)
                buffer = buf
                KeyAscii = 0
            Else
                KeyAscii = 0
                Exit Sub

            End If

        End If

        buf = Chr(KeyAscii)

        If Chr(KeyAscii) = "*" Then
            buf = ""
            buffer = buf

        End If

        If KeyAscii <> 13 Then
            buffer = buffer + buf

        End If

        buf = buffer
        ejecuta 0
         
    End If

End Sub

Private Sub descripcio_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

    'RUC.SetFocus
End Sub

Private Sub dfju773_Click()

    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "cuentas"
    reporgen.Show 1

End Sub

Private Sub dlo132_Click()

    If Frame2.Visible = True Then
        habilita 0
        Frame2.Visible = False
        dbGrid1.Enabled = True
        Exit Sub

    End If

    tctable.Hide
    Unload tctable

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    buf = txctacon.Fields("cuenta")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Modifica"
    cmdGuardar.Enabled = True
    pone_registro
    habilita 1
    cuenta.Enabled = False
    descripcio.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Sub menu_reporte()

    Dim sdx As String

    On Error GoTo cmd8_err

    sdx = "" & txctacon.Fields("cuenta")
    impresion1
    Exit Sub
cmd8_err:
    MsgBox "Elegir un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    buf = txctacon.Fields("cuenta")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Zoom"
    cmdGuardar.Enabled = False
    pone_registro
    habilita 1
    cuenta.Enabled = False
    descripcio.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fk8944_Click()

    Dim found      As Integer

    Dim I          As Integer

    Dim v          As Long

    Dim R          As Long

    Dim ih         As Integer

    Dim h          As Integer

    Dim cad        As String

    Dim Tmp        As String

    Dim sw         As Integer

    Dim sdx        As Double

    Dim mytablex   As New ADODB.Recordset

    Dim Heading(8) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd45612_err

    If MsgBox("Desea Exportar excel", 1, "Aviso") <> 1 Then Exit Sub
    Heading(1) = "Codigo de Cuenta."
    Heading(2) = "Descripcion de la Cuenta "
    Heading(3) = "Tipo Cuenta"
    Heading(4) = "Nivel Cuenta"
    Heading(5) = dicruc
    Heading(6) = "Ccosto"

    If txctacon.RecordCount = 0 Then Exit Sub
    txctacon.MoveFirst
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    
    '------------------------------------------------
    With objExcel.ActiveSheet
        '.Range(.Cells(1, 1), .Cells(10, 3)).Borders.LineStyle = xlContinuous
        .Range(.Cells(4, 1), .Cells(4, 20)).Borders.LineStyle = xlContinuous
       
        .columns("A").ColumnWidth = 15
        .columns("B").ColumnWidth = 50
        .columns("C").ColumnWidth = 15
        .columns("D").ColumnWidth = 15
    
    End With

    'cabecera
    mytablex.Open "select * from empresa where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        objExcel.ActiveSheet.Cells(2, 1) = "'" & mytablex.Fields("nombre")

    End If

    mytablex.Close
    objExcel.ActiveSheet.Cells(2, 5) = "'" & Format(Now, "dd/mm/yyyy")
    objExcel.ActiveSheet.Cells(3, 2) = "'Lista de Cuentas Contables"
    
    '------------------------------------------------
    objExcel.ActiveSheet.Cells(4, 1) = "'Codigo Cuenta"
    objExcel.ActiveSheet.Cells(4, 2) = "'Descripcion"
    objExcel.ActiveSheet.Cells(4, 3) = "'Tipo Cuenta"
    objExcel.ActiveSheet.Cells(4, 4) = "'Nivel Cuenta"
    objExcel.ActiveSheet.Cells(4, 5) = "'Ruc"
    objExcel.ActiveSheet.Cells(4, 6) = "'Ccosto"
    '------------------------------------------------
    v = 5
    h = 1
    
    Do

        If txctacon.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, 1) = "'" & txctacon.Fields("cuenta")
        objExcel.ActiveSheet.Cells(v, 2) = "'" & txctacon.Fields("Descripcio")
        objExcel.ActiveSheet.Cells(v, 3) = "'" & txctacon.Fields("tipocuenta")
        objExcel.ActiveSheet.Cells(v, 4) = "" & txctacon.Fields("nivel_cta")
        objExcel.ActiveSheet.Cells(v, 5) = "'" & txctacon.Fields("flag_ruc")
        objExcel.ActiveSheet.Cells(v, 6) = "'" & txctacon.Fields("flag_ccosto")
   
        v = v + 1
        txctacon.MoveNext
    Loop
    
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd45612_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    Command1_Click

End Sub

Private Sub Form_Load()

    Dim mytablex As New ADODB.Recordset

    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "cuenta"
    Combo1.ListIndex = 1

    ntipocuenta.Clear
    ntipocuenta.AddItem ""
    mytablex.Open "select * from tipocta", cn, adOpenDynamic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        ntipocuenta.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("tipocta"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    ntipocuenta.ListIndex = 0

    nnivelcuenta.Clear
    nnivelcuenta.AddItem ""
    nnivelcuenta.AddItem "BALANCE|B"
    nnivelcuenta.AddItem "SUBCUENTA|S"
    nnivelcuenta.AddItem "REGISTRO|R"
    nnivelcuenta.ListIndex = 0

    ntipoanalisis.Clear
    ntipoanalisis.AddItem ""
    ntipoanalisis.AddItem "Sin Analisis|"
    ntipoanalisis.AddItem "Por Documentos|S"
    ntipoanalisis.AddItem "Cuenta de banco|B"
    ntipoanalisis.AddItem "SoloDetalle|D"
    ntipoanalisis.ListIndex = 0

    nmoneda.Clear
    nmoneda.AddItem ""
    nmoneda.AddItem "S"
    nmoneda.AddItem "D"
    nmoneda.ListIndex = 0

End Sub

Sub inicializa()
    descripcio = ""
    ntipocuenta.ListIndex = 0
    nnivelcuenta.ListIndex = 0
    tipocuenta = ""
    nivelcuenta = ""
    tipoanalisis = ""
    moneda = ""

End Sub

Sub pone_registro()
    moneda = Trim("" & txctacon.Fields("moneda"))
    tipocuenta = Trim("" & txctacon.Fields("tipocuenta"))
    nivelcuenta = Trim("" & txctacon.Fields("nivelcuenta"))
    cuenta = Trim("" & txctacon.Fields("cuenta"))
    tipoanalisis = Trim("" & txctacon.Fields("tipoanalisis"))
    descripcio = Trim("" & txctacon.Fields("descripcio"))

End Sub

Sub grabando()
    txctacon.Fields("moneda") = Trim(moneda)
    txctacon.Fields("cuenta") = Trim(cuenta)
    txctacon.Fields("descripcio") = Trim(descripcio)
    txctacon.Fields("tipocuenta") = Trim(tipocuenta)
    txctacon.Fields("tipoanalisis") = Trim(tipoanalisis)
    txctacon.Fields("nivelcuenta") = Trim(nivelcuenta)

End Sub

Function grabar()

    Dim found  As Integer

    Dim rbusca As New ADODB.Recordset

    found = valida()

    If found = 0 Then
        MsgBox "Campos invalidos", 48, "Aviso"
        Exit Function

    End If

    If Frame2.Caption = "Nuevo" Then
        If Len(cuenta) = 0 Then
            cuenta.SetFocus
            Exit Function

        End If

        rbusca.Open "select cuenta from cuentas where cuenta='" & cuenta & "'", cn, adOpenStatic, adLockOptimistic

        If rbusca.RecordCount > 0 Then
            rbusca.Close
            MsgBox "Ya existe Codigo Cuenta ", 48, "Aviso"
            Exit Function

        End If

        txctacon.AddNew
        txctacon.Fields("cuenta") = cuenta
        grabando
        txctacon.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        txctacon.Fields("cuenta") = cuenta
        grabando
        txctacon.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    If Len(cuenta) = 0 Then
        cuenta.SetFocus
        Exit Function

    End If

    If Len(descripcio) = 0 Then
        descripcio.SetFocus
        Exit Function

    End If

    'If Len(tipocuenta) = 0 Then
    'tipocuenta.SetFocus
    '   Exit Function
    'End If
    'If Len(nivelcuenta) = 0 Then
    'nivelcuenta.SetFocus
    '   Exit Function
    'End If

    'If Len(tipoanalisis) = 0 Then
    'tipoanalisis.SetFocus
    '   Exit Function
    'End If

    valida = 1

End Function

Sub habilita(sw As Integer)

    If sw = 0 Then

        ajdu1.Enabled = True
        f8443.Enabled = True
        bo712.Enabled = True
        fjh433.Enabled = True
        djuer1.Enabled = True
        djuer1.Enabled = True
        Picture1.Enabled = True
        dbGrid1.Enabled = True
            
    End If

    If sw = 1 Then

        ajdu1.Enabled = False
        f8443.Enabled = False
        bo712.Enabled = False
        fjh433.Enabled = False
        djuer1.Enabled = False
        djuer1.Enabled = False
        Picture1.Enabled = False
        dbGrid1.Enabled = False
        dbGrid1.Enabled = False
            
    End If
      
End Sub

Private Sub Label25_Click()
    Label26.Visible = True

End Sub

Private Sub nhreyr_Click()
    menu_reporte

End Sub

Private Sub impresion1()

    Dim found As Integer

    Dim buf   As String

    If MsgBox("Desea Imprimir", 1, "Aviso") <> 1 Then Exit Sub
    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    'found = ir_primero1()
    
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento
    cuerpo_programa_documento
    '------------------------------------
    Close #1
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

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
    buf = "Tabla de Cuentas Contable  "
    found = formateaa(buf, 90, 2, 0)
       
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)
    '------aqui van los registros----------------------
        
    found = formateaa("Cod Cuenta", 15, 0, 0)
    found = formateaa("Descripcio", 51, 0, 0)
    found = formateaa("Clasif ", 11, 0, 0)
    found = formateaa("Cuenta ", 11, 0, 0)
    found = formateaa(dicruc, 6, 2, 0)
    '--------------------------------------------------
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)

End Sub

Sub cuerpo_programa_documento()

    Dim xdebito  As Double

    Dim xcredito As Double

    Dim buf      As String

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd788_err

    suma1 = 0
    suma2 = 0
    xdebito = 0
    xcredito = 0

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from cuentas where cuenta like '%' order by cuenta", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        'MsgBox "" & mytablex.Fields("producto")
        If mytablex.EOF Then Exit Do
        '-----------------------------------------
        buf = "" & mytablex.Fields("cuenta")
        found = formateaa(buf, 14, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("descripcio")
        found = formateaa(buf, 50, 0, 0)
        found = formateaa("", 1, 0, 0)

        buf = ""

        If "" & mytablex.Fields("tipocuenta") = "A" Then
            buf = "ACTIVO"

        End If

        If "" & mytablex.Fields("tipocuenta") = "P" Then
            buf = "PASIVO"

        End If

        If "" & mytablex.Fields("tipocuenta") = "C" Then
            buf = "CAPITAL"

        End If

        If "" & mytablex.Fields("tipocuenta") = "V" Then
            buf = "VENTAS"

        End If

        If "" & mytablex.Fields("tipocuenta") = "G" Then
            buf = "GASTOS"

        End If

        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = ""

        If Trim("" & mytablex.Fields("nivel_cta")) = "D" Then
            buf = "DETALLE"

        End If

        If Trim("" & mytablex.Fields("nivel_cta")) = "S" Then
            buf = "DEGRUPO"

        End If
    
        'buf = "" & mytablex.Fields("nivel_cta")
    
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 2, 0)

        'buf = "" & mytablex.Fields("ruc")
        'found = formateaa(buf, 5, 0, 0)
        'found = formateaa("", 1, 2, 0)

        nlineas
        mytablex.MoveNext
    Loop
    mytablex.Close
    Exit Sub
cmd788_err:
    Exit Sub

End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > 45 Then
        cabecera_documento

    End If

End Sub

Private Sub nnivelcuenta_Click()
    'If nnivelcuenta <> "" Then
    nivelcuenta = extra_loquesea1(nnivelcuenta)

    'End If
End Sub

Private Sub ntipoanalisis_Click()
    'If ntipoanalisis <> "" Then
    tipoanalisis = extra_loquesea1(ntipoanalisis)

    'End If
End Sub

Private Sub ntipocuenta_Click()
    'If ntipocuenta <> "" Then
    tipocuenta = extra_loquesea1(ntipocuenta)

    'End If
End Sub

Private Sub ui78232_Click()

    On Error GoTo cmd4567_err

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "Select * from cuentas", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        MsgBox "Ya existen cuentas,no se puede copiar ", 48, "Aviso"
        mytablex.Close
        Exit Sub

    End If

    mytablex.Close

    If MsgBox("Desea Copias las cuentas predefinidas ", 1, "Aviso") <> 1 Then Exit Sub
    cn.Execute ("delete from cuentas")
    cn.Execute ("insert into cuentas select  * from ctaseje")
    MsgBox "Proceso Realizado ", 24, "Aviso"
    Command1_Click
    Exit Sub
cmd4567_err:
    MsgBox "Aviso en copiar tablas predefinidas " + error$, 48, "Aviso"

    Exit Sub

End Sub
