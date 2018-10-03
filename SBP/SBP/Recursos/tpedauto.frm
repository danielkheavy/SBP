VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tpedauto 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos Automaticos Proveedores"
   ClientHeight    =   8805
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   14325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   14325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Lista Precios"
      Height          =   3735
      Left            =   2160
      TabIndex        =   41
      Top             =   2040
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton Command8 
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
         Left            =   7440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tpedauto.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Borrar registro"
         Top             =   360
         Width           =   735
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Height          =   3255
         Left            =   120
         OleObjectBlob   =   "tpedauto.frx":1212
         TabIndex        =   43
         Top             =   240
         Width           =   7215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consultas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8775
      Left            =   0
      TabIndex        =   37
      Top             =   720
      Visible         =   0   'False
      Width           =   13935
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
         Height          =   495
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Ejecutar"
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
         Left            =   5400
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7575
         Left            =   0
         TabIndex        =   50
         Top             =   840
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   13361
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   22
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   680
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   14265
      TabIndex        =   31
      Top             =   0
      Width           =   14325
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
         Picture         =   "tpedauto.frx":2275
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Grabar registro"
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
         Picture         =   "tpedauto.frx":3487
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Nuevo registro"
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
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tpedauto.frx":4699
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Generacion de Pedidos Seleccionados"
      Height          =   4695
      Left            =   3120
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox local1 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   46
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox numero 
         Height          =   375
         Left            =   2640
         MaxLength       =   11
         TabIndex        =   26
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox serie 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         MaxLength       =   3
         TabIndex        =   25
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox tipo 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   24
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Procesar"
         Height          =   735
         Left            =   1560
         TabIndex        =   23
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nota. Los campos que no tienen proveedor nose graban"
         Height          =   495
         Left            =   120
         TabIndex        =   48
         Top             =   4080
         Width           =   6375
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         Height          =   375
         Left            =   1080
         TabIndex        =   45
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         Height          =   375
         Left            =   1080
         TabIndex        =   30
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
         Height          =   375
         Left            =   1080
         TabIndex        =   29
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   375
         Left            =   1080
         TabIndex        =   28
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Solo se generan los pedidos de proveedores que existan en la base de datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         TabIndex        =   27
         Top             =   360
         Width           =   4935
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Condicion de Seleccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   1320
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   10935
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cualquiera"
         Height          =   375
         Left            =   3360
         TabIndex        =   47
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton cmdGrabar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tpedauto.frx":58AB
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   9600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tpedauto.frx":6059
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.ComboBox bodega 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   600
         Width           =   4215
      End
      Begin VB.ComboBox familia 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox proveedor 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   12
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1680
         Width           =   3495
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Menor que 0"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Menor Igual que 0"
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   3240
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Menor que stockMinimo Se toma de Tabla Productos"
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   7
         Top             =   2040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   6
         Top             =   2400
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label local2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1560
         TabIndex        =   49
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Almacen"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Familia"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoProveedor"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "StockDisponible"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   2880
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Final"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2400
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Pedidos a Proveedores"
      Height          =   8055
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   14175
      Begin VB.ComboBox Combo2 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   7560
         Width           =   3615
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "tpedauto.frx":6807
         Height          =   7215
         Left            =   120
         OleObjectBlob   =   "tpedauto.frx":681B
         TabIndex        =   0
         Top             =   240
         Width           =   13935
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         Height          =   375
         Left            =   7800
         TabIndex        =   4
         Top             =   7560
         Width           =   1695
      End
      Begin VB.Label Total 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   9480
         TabIndex        =   3
         Top             =   7560
         Width           =   1695
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu fdlo2323 
      Caption         =   "&Nuevo"
   End
   Begin VB.Menu dki23232 
      Caption         =   "&Grabar"
   End
   Begin VB.Menu ldo333 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tpedauto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type campo_precio

    unidad As String
    factor As String
    precio As String
    costo As String
    margen As String
    stock As String

End Type

Dim xproducto         As String

Dim xbodega           As String

Dim campo_precios(12) As campo_precio

Private Sub bodega_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    familia.SetFocus

End Sub

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        ldo333_Click
        Exit Sub

    End If

    Command1_Click

End Sub

Private Sub cmdAddEntry_Click()

    If Frame1.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub
    Frame4.Visible = True
    proveedor.SetFocus

End Sub

Private Sub cmdCancelar_Click()
    Frame4.Visible = False

End Sub

Private Sub cmdExit_Click()
    ldo333_Click

End Sub

Private Sub cmdGrabar_Click()

    If MsgBox("Desea Generar el Pedido", 1, "Aviso") <> 1 Then Exit Sub
    proceso_pedido_automatico

End Sub

Private Sub cmdSave_Click()

    Dim found As Integer

    If Frame1.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub
    found = busca_parame()

    If found = 0 Then
        MsgBox "No existe Parametro", 48, "Aviso"
        Exit Sub
        Exit Sub

    End If

    If Len(tipo) = 0 Then Exit Sub
    found = busca_tipo(0)

    If found = 0 Then
        MsgBox "NO existe tipo " & tipo, 48, "Aviso"
        Exit Sub

    End If

    local1 = local2
    Frame3.Visible = True

End Sub

Private Sub Combo2_Click()
    SQL_pedido Combo2

End Sub

Private Sub Command1_Click()

    Dim buf       As String

    Dim sw        As Integer

    Dim rconsulta As New ADODB.Recordset

    If opcion1 = "1" Or opcion1 = "2" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from Proveedo "
        Else
            buf = "select Nombre,Codigo from Proveedo where " & Combo1 & " like '" & buffer & "%'"

        End If

    End If

    'MsgBox buf
    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True And rconsulta.BOF = True Then
        rconsulta.Close
        buffer.SetFocus
        Exit Sub

    End If

    Set dbGrid1.DataSource = rconsulta
    dbGrid1.columns(0).Width = 4000
    dbGrid1.columns(1).Width = 2000

    If sw = 1 Then
        dbGrid1.SetFocus

    End If

End Sub

Private Sub Command_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
    'sql_pedido1
    ir_inicio
    proceso_generacion
    ldo333_Click

End Sub

Sub proceso_generacion()

    Dim found      As Integer

    Dim Tmp        As String

    Dim sw         As Integer

    Dim xtipo      As String

    Dim xserie     As String

    Dim xnumero    As String

    Dim sdx        As Double

    Dim xproveedor As String

    xnumero = "" & Numero
    'Data2.Refresh
    sw = 0
    Data2.refresh
    Do

        If Data2.Recordset.EOF Then Exit Do
        If Val("" & Data2.Recordset.Fields("cantPEDIDO")) > 0 Then
            xproveedor = "" & Data2.Recordset.Fields("proveedor")

            If sw = 0 Then
                sw = 1
                Tmp = xproveedor
regresa:

                If valida_si_existe() = 1 Then
                    sdx = Val(Numero) + 1
                    Numero = "" & sdx
                    GoTo regresa

                End If

                'MsgBox "X"
                graba_cabecera Numero

            End If

            If Tmp <> xproveedor Then
                Tmp = xproveedor
regresa1:

                If valida_si_existe() = 1 Then
                    sdx = Val(Numero) + 1
                    Numero = "" & sdx
                    GoTo regresa1

                End If

                graba_cabecera Numero

            End If

            'ahora valida si son proveedores
            graba_detalle Numero

        End If

        Data2.Recordset.MoveNext
    Loop
    found = busca_tipo(1)

End Sub

Sub graba_detalle(xnumero As String)

    Dim mytablez   As New ADODB.Recordset

    Dim xproveedor As String

    Dim I          As Integer

    Dim sdx        As Double

    Dim xcant      As Double

    xcant = Val("" & Data2.Recordset.Fields("cantPEDIDO"))

    If xcant <= 0 Then
        xcant = 1

    End If

    xproveedor = "" & Data2.Recordset.Fields("proveedor")

    If Len(xproveedor) = 0 Then
        xproveedor = "999"

    End If

    mytablez.Open "select * from dordenc where local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & xnumero & "'", cn, adOpenStatic, adLockOptimistic
    'If mytablez.RecordCount = 0 Then
    mytablez.AddNew
    mytablez.Fields("local") = local1
    mytablez.Fields("tipo") = tipo
    mytablez.Fields("serie") = serie
    mytablez.Fields("numero") = xnumero
    mytablez.Fields("moneda") = "S"
    mytablez.Fields("vendedor") = ""
    mytablez.Fields("acu") = "R"
    mytablez.Fields("acu1") = ""
    mytablez.Fields("fecha") = Format(Now, "dd/mm/yyyy")
    mytablez.Fields("tipoclie") = "P"
    mytablez.Fields("codigo") = xproveedor
    mytablez.Fields("producto") = "" & Data2.Recordset.Fields("producto")
    mytablez.Fields("descripcio") = "" & Data2.Recordset.Fields("descripcio")
    mytablez.Fields("unidad") = "" & Data2.Recordset.Fields("unidad")
    mytablez.Fields("factor") = Val("" & Data2.Recordset.Fields("factor"))
    mytablez.Fields("cantidad") = xcant
    mytablez.Fields("precio") = Val("" & Data2.Recordset.Fields("costo"))
    mytablez.Fields("total") = Val("" & Data2.Recordset.Fields("total"))
    mytablez.Fields("igv") = Val("" & Data2.Recordset.Fields("igv"))
    mytablez.Update
    'End If
    mytablez.Close

End Sub

Sub ir_inicio()

    On Error GoTo cmd23_err

    Data2.Recordset.MoveFirst
    Exit Sub
cmd23_err:
    Exit Sub

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command8_Click()
    ldo333_Click

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            DBGrid2.columns(1) = dbGrid1.columns(1)
            Frame1.Visible = False
            DBGrid2.SetFocus

        End If

        If opcion1 = "2" Then
            proveedor = dbGrid1.columns(1)
            Frame1.Visible = False
            proveedor.SetFocus

        End If

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
        Command1_Click
         
    End If

End Sub

Private Sub dbgrid2_BeforeColEdit(ByVal ColIndex As Integer, _
                                  ByVal KeyAscii As Integer, _
                                  Cancel As Integer)

    Dim found As Integer

    Select Case ColIndex

        Case 1, 2, 3, 4, 5, 8, 9, 10
            Cancel = True
            Exit Sub

        Case 0

        Case 7, 6

            If "" & DBGrid2.columns(0) <> "X" Then  '
                Cancel = True
                Exit Sub

            End If

    End Select

End Sub

Private Sub dbgrid2_BeforeColUpdate(ByVal ColIndex As Integer, _
                                    OldValue As Variant, _
                                    Cancel As Integer)

    Dim found As Integer

    Dim sdx   As Double

    Select Case ColIndex

        Case 6

            If "" & DBGrid2.columns(0) <> "X" Then  '
                Cancel = True
                Exit Sub

            End If

            If Not IsNumeric(DBGrid2.columns(6)) Then  '
                Cancel = True
                Exit Sub

            End If

            sdx = Val("" & DBGrid2.columns(7)) * Val("" & DBGrid2.columns(6))
            DBGrid2.columns(8) = sdx

        Case 7

            If "" & DBGrid2.columns(0) <> "X" Then  '
                Cancel = True
                Exit Sub

            End If

            If Not IsNumeric(DBGrid2.columns(7)) Then  '
                Cancel = True
                Exit Sub

            End If

            sdx = Val("" & DBGrid2.columns(7)) * Val("" & DBGrid2.columns(6))
            DBGrid2.columns(8) = sdx

    End Select

End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo cmd45err

    If KeyCode = &H70 Then  'f1
        If Len(DBGrid2.columns(2)) > 0 And DBGrid2.Col = 4 Then
            xproducto = "" & DBGrid2.columns(2)
            carga_dbgrid4
            Exit Sub

        End If

    End If

    Exit Sub
cmd45err:
    Exit Sub

End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        If DBGrid2.columns(0) = "X" And DBGrid2.Col = 1 Then
            consulta_proveedor

        End If

    End If

End Sub

Private Sub dflo8922_Click()

End Sub

Private Sub DBGrid4_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        'If opcion3 = "1" Then
        '   Frame5.Visible = False
        '   DBGrid1.SetFocus
        '   Exit Sub
        'End If
        ldo333_Click
        Exit Sub

    End If

    If KeyCode = 13 Then

        'If opcion3 = "1" Then
        '   Frame5.Visible = False
        '   DBGrid1.SetFocus
        '   Exit Sub
        'End If
        'If opcion1 = "8" Then
        If Len("" & DBGrid4.columns(0)) > 0 And Val("" & DBGrid4.columns(1)) > 0 And Len("" & DBGrid4.columns(3)) > 0 Then
            'Data2.Recordset.Edit
            'Data2.Recordset.Fields("unidad") = "" & DBGrid4.Columns(0)
            'Data2.Recordset.Fields("factor") = "" & DBGrid4.Columns(1)
            'Data2.Recordset.Fields("precio") = "" & DBGrid4.Columns(3)
            'Data2.Recordset.Update
            DBGrid2.columns(4) = "" & DBGrid4.columns(0)
            DBGrid2.columns(5) = Val("" & DBGrid4.columns(1))
            DBGrid2.columns(6) = Val("" & DBGrid4.columns(3))
            DBGrid2.refresh
            ldo333_Click
      
        End If

        'End If
    End If

End Sub

Private Sub DBGrid4_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, _
                                    StartLocation As Variant, _
                                    ByVal ReadPriorRows As Boolean)

    Dim dR            As Integer

    Dim row_num       As Integer

    Dim R             As Integer

    Dim rows_returned As Integer

    If ReadPriorRows Then
        dR = -1
    Else
        dR = 1

    End If

    If IsNull(StartLocation) Then
        If ReadPriorRows Then
            row_num = RowBuf.RowCount - 1
            'row_num = 9
        Else
            row_num = 0

        End If

    Else
        row_num = CLng(StartLocation) + dR

    End If

    rows_returned = 0

    For R = 0 To RowBuf.RowCount - 1

        If row_num < 0 Or row_num > 9 Then Exit For
        RowBuf.Value(R, 0) = campo_precios(row_num).unidad
        RowBuf.Value(R, 1) = campo_precios(row_num).factor
        RowBuf.Value(R, 2) = campo_precios(row_num).precio
        RowBuf.Value(R, 3) = campo_precios(row_num).costo
        RowBuf.Value(R, 4) = campo_precios(row_num).margen
        RowBuf.Value(R, 5) = campo_precios(row_num).stock
        RowBuf.Bookmark(R) = row_num
        row_num = row_num + dR
        rows_returned = rows_returned + 1
    Next R

    RowBuf.RowCount = rows_returned

End Sub

Private Sub dki23232_Click()
    cmdSave_Click

End Sub

Private Sub familia_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    proveedor.SetFocus

End Sub

Private Sub fdlo2323_Click()
    cmdAddEntry_Click

End Sub

Private Sub Form_Load()
    Combo3.AddItem "HABITUAL"
    'Combo3.AddItem "MEJOR_PRECIO"
    'Combo3.AddItem "MEJOR_FECHA_REPOSICION"
    Combo3.ListIndex = 0

    Combo2.Clear
    Combo2.AddItem "PROVEEDOR"
    Combo2.AddItem "DESCRIPCIO"
    Combo2.AddItem "val(PRODUCTO)"
    Combo2.ListIndex = 0

    cargar
    local2 = glocal

End Sub

Function busca_parame()

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        tipo = "" & mytablex.Fields("pedauto")
        tipo = "R"
        busca_parame = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_tipo(sw As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    If sw = 0 Then
        serie = ""
        Numero = ""

    End If

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from tipo where tipo='" & tipo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        If sw = 0 Then
            serie = "" & mytablex.Fields("serie")
            sdx = Val("" & mytablex.Fields("numero")) + 1
            Numero = "" & sdx

        End If

        If sw = 0 Then
            'mytablex.Edit
            mytablex.Fields("numero") = Numero
            mytablex.Update

        End If

        busca_tipo = 1

    End If

    mytablex.Close
 
End Function

Private Sub ldo333_Click()

    If Frame5.Visible = True Then
        Frame5.Visible = False
        Exit Sub

    End If

    If Frame3.Visible = True Then
        Frame3.Visible = False
        Exit Sub

    End If

    If Frame1.Visible = True Then
        Frame1.Visible = False

        If opcion1 = "2" Then
            'MsgBox "Hola"
            proveedor.SetFocus
            Exit Sub

        End If

        If opcion1 = "1" Then
            DBGrid2.SetFocus
            Exit Sub

        End If

    End If

    If Frame4.Visible = True Then
        Frame4.Visible = False
        DBGrid2.SetFocus
        Exit Sub

    End If

    tpedauto.Hide
    Unload tpedauto

End Sub

Sub cargar()

    Dim mytablex As New ADODB.Recordset

    bodega.Clear
    familia.Clear
    familia.AddItem "%"

    mytablex.Open "select * from bodega", cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        bodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    '------------------------------------- ------------
    mytablex.Close
    mytablex.Open "select * from familia", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        familia.AddItem "" & mytablex.Fields("familia") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    '------------------------------------- ------------
    mytablex.Close
    bodega.ListIndex = 0
    familia.ListIndex = 0

End Sub

Sub proceso_pedido_automatico()

    Dim found    As Integer

    Dim mytabley As Table

    Dim mysnapx  As New ADODB.Recordset

    Dim sw       As Integer

    Dim buf      As String

    Dim xminimo  As Double

    Dim xsaldo   As Double

    If Check1.Value = 0 Then
        If Len(proveedor) = 0 Then
            MsgBox "Debe existir un proveedor", 48, "Aviso"
            proveedor.SetFocus
            Exit Sub

        End If

    End If

    mydbxglo.Execute "DELETE FROM  " & dgusuario
    Set mytabley = mydbxglo.OpenTable(dgusuario)

    If Check1.Value = 0 Then
        buf = "select producto.unidad,producto.factor,producto.costou ,producto.igv ,Producto.descripcio,Producto.producto,producto.marca,producto.unidad as Und1,producto.Factor as F,Producto.Costou as Precio,producto.monedac as M,producto.familia,producto.Subfamilia,codprov.codigo,producto.minimo,producto.maximo from codprov left join producto on codprov.producto=producto.producto where codprov.codigo='" & proveedor & "'"

    End If

    If Check1.Value = 1 Then
        buf = "select producto.unidad,producto.factor,producto.costou ,producto.igv ,Producto.descripcio,Producto.producto,producto.marca,producto.unidad as Und1,producto.Factor as F,Producto.Costou as Precio,producto.monedac as M,producto.familia,producto.Subfamilia,producto.minimo from producto   "

    End If

    If familia <> "%" Then
        If Check1.Value = 0 Then
            buf = buf & " and  producto.familia='" & extra_loquesea(familia) & "'"

        End If

        If Check1.Value = 1 Then
            buf = buf & "  where producto.familia='" & extra_loquesea(familia) & "'"

        End If

    End If

    mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic

    Do

        If mysnapx.EOF Then Exit Do
        xminimo = Val("" & mysnapx.Fields("minimo"))

        xsaldo = busca_saldo("" & mysnapx.Fields("producto"))
        sw = 0

        If Option1.Value = True Then  'si es menor que cero
            If xsaldo < 0 Then GoTo am1
            GoTo siguiente

        End If

        If Option2.Value = True Then  'si es menor o igual que cero
            If xsaldo <= 0 Then GoTo am1
            GoTo siguiente

        End If

        If Option3.Value = True Then  'si es menor que minimo
            If xsaldo < xminimo Then GoTo am1
            GoTo siguiente

        End If

am1:
        mytabley.AddNew
        mytabley.Fields("Ok") = "X"

        If Check1.Value = 0 Then
            mytabley.Fields("proveedor") = "" & mysnapx.Fields("codigo")
        Else
            mytabley.Fields("proveedor") = "999"

        End If

        'mytabley.Fields("proveedor") = "" & mysnapx.Fields("codigo")
        mytabley.Fields("producto") = "" & mysnapx.Fields("producto")
        mytabley.Fields("descripcio") = "" & mysnapx.Fields("descripcio")
        mytabley.Fields("unidad") = "" & mysnapx.Fields("unidad")
        mytabley.Fields("factor") = Val("" & mysnapx.Fields("factor"))
        mytabley.Fields("costo") = Val("" & mysnapx.Fields("costou"))
        mytabley.Fields("igv") = Val("" & mysnapx.Fields("igv"))
        mytabley.Fields("stockfisic") = 0
        mytabley.Fields("stockdisp") = 0
        mytabley.Fields("cantpedido") = xminimo
        mytabley.Fields("total") = 0
        mytabley.Fields("stockdisp") = xsaldo
        'buscando_ventas
        mytabley.Update
        '------------------------------------------
siguiente:
        mysnapx.MoveNext
    Loop
    mysnapx.Close
 
    Frame4.Visible = False
    DBGrid2.SetFocus
    SQL_pedido "Proveedor"

End Sub

Sub SQL_pedido(buf As String)
    Frame2.Visible = True
    DBGrid2.Enabled = True
    Data2.Connect = "foxpro 2.5;"
    Data2.DatabaseName = globaldir
    Data2.RecordSource = "select * from " & dgusuario & " order by  " & buf
    Data2.refresh
    'DBGrid2.SetFocus

End Sub

Sub sql_pedido1()
    Frame2.Visible = True
    Data2.Connect = "foxpro 2.5;"
    Data2.DatabaseName = globaldir
    Data2.RecordSource = "select * from " & dgusuario & " where OK='X' and cantpedido>0 and len(proveedor)>0 order by PROVEEDOR"
    Data2.refresh
    DBGrid2.SetFocus

End Sub

Sub consulta_proveedor()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "1"
    Command1_Click

End Sub

Sub consulta_proveedor1()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "2"
    Command1_Click

End Sub

Function busca_saldo(buf As String) As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from almacen where local='" & local1 & "' and producto='" & buf & "' and bodega='" & extra_loquesea(bodega) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_saldo = Val("" & mytablex.Fields("saldo"))

        'yminimo = Val("" & mytablex.Fields("minimo"))
    End If

    '------------------------------------- ------------
    mytablex.Close
 
End Function

Private Sub proveedor_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Combo3.SetFocus

End Sub

Private Sub proveedor_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_proveedor1

    End If

End Sub

Sub graba_cabecera(xnumero As String)

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    Dim I        As Integer

    mytablex.Open "select * from cordenc where local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & xnumero & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        pone_registro_compra mytablex, xnumero
        mytablex.Update

    End If

    mytablex.Close
 
End Sub

Sub pone_registro_compra(mytablex As ADODB.Recordset, xnumero As String)

    Dim xproveedor As String

    xproveedor = "" & Data2.Recordset.Fields("proveedor")

    If Len(xproveedor) = 0 Then
        xproveedor = "999"

    End If

    mytablex.Fields("yausado") = "0"
    mytablex.Fields("local") = local1
    mytablex.Fields("tipo") = tipo
    mytablex.Fields("serie") = serie
    mytablex.Fields("numero") = xnumero
    mytablex.Fields("acu") = "R"
    mytablex.Fields("acu1") = ""
    mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("fechae") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("tipoclie") = "P"
    mytablex.Fields("codigo") = xproveedor
    mytablex.Fields("nombre") = ""
    mytablex.Fields("estado") = "0"
    mytablex.Fields("partida") = ""
    mytablex.Fields("destino") = ""
    mytablex.Fields("moneda") = "S"
    mytablex.Fields("vendedor") = ""
    mytablex.Fields("fpago") = ""
    mytablex.Fields("transporte") = ""
    mytablex.Fields("paridad") = 1
    mytablex.Fields("dias") = 1
    mytablex.Fields("bodega") = "01"
    mytablex.Fields("bodegaf") = ""
    mytablex.Fields("observa") = ""
    mytablex.Fields("usuario") = "" & gusuario
    mytablex.Fields("flage") = ""
    mytablex.Fields("hora") = Format(Now, "hh:MM")
    mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")

    mytablex.Fields("total") = 0
    mytablex.Fields("descuento") = 0
    mytablex.Fields("neto") = 0
    mytablex.Fields("impuesto") = 0
    mytablex.Fields("subtotal") = 0

    'mytablex.Fields("c1") = Val(c1)
    'mytablex.Fields("c2") = Val(c2)
    'mytablex.Fields("c3") = Val(c3)
    'mytablex.Fields("c4") = Val(c4)
End Sub

Function valida_si_existe()

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    Dim I        As Integer

    mytablex.Open "select * from cordenc where local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        valida_si_existe = 1

    End If

    mytablex.Close
 
End Function

Sub suma_detalle()

    Dim sdx As Double

    On Error GoTo cmd567_err

    sdx = 0
    total = ""
    Data2.Recordset.MoveFirst
    Do

        If Data2.Recordset.EOF Then Exit Do
        sdx = sdx + Val("" & Data2.Recordset.Fields("total"))
        Data2.Recordset.MoveNext
    Loop
    total = Format(sdx, "0.00")
    Exit Sub
cmd567_err:
    Exit Sub

End Sub

Private Sub total_Click()
    suma_detalle

End Sub

Sub carga_dbgrid4()

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim sw       As Integer

    Dim xbodega  As String

    Dim xsaldo   As Double

    Dim xbuf     As String

    Dim xcosto   As Double

    Dim xmargen  As Double

    For I = 0 To 9
        campo_precios(I).unidad = ""
        campo_precios(I).factor = ""
        campo_precios(I).precio = ""
        campo_precios(I).costo = ""
        campo_precios(I).margen = ""
        campo_precios(I).stock = ""
    Next I

    xbodega = "01"
    xsaldo = 0
    xcosto = 0
    sw = 0
    mytabley.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        xbodega = "" & mytabley.Fields("bodega")

    End If

    mytabley.Close
    mytabley.Open "select * from almacen where local1='" & local1 & "' and producto='" & xproducto & "' and bodega='" & xbodega & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        xsaldo = Val("" & mytabley.Fields("saldo"))

    End If

    mytabley.Close
    mytablex.Open "select * from producto where producto='" & xproducto & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        '----------------------------------------------
        xcosto = Val("" & mytablex.Fields("costou"))
        '----------------
        '----------------
        campo_precios(0).unidad = "" & mytablex.Fields("unidad")
        campo_precios(0).factor = "" & mytablex.Fields("factor")
        campo_precios(0).precio = "" '& mytablex.Fields("costou")
        campo_precios(0).costo = "" & xcosto
        xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor")))
        campo_precios(0).stock = "" & xbuf
        xmargen = 0
        campo_precios(0).margen = "" & xmargen
   
        '----------------------------------------------
        xcosto = 0

        If Val("" & mytablex.Fields("factor1")) > 0 Then
            xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
            xcosto = xcosto * Val("" & mytablex.Fields("factor1"))

        End If

        '----------------
        '----------------
        campo_precios(1).unidad = "" & mytablex.Fields("unidad1")
        campo_precios(1).factor = "" & mytablex.Fields("factor1")
        campo_precios(1).precio = "" & mytablex.Fields("pventa1")
        campo_precios(1).costo = "" & xcosto
        xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor1")))
        campo_precios(1).stock = "" & xbuf
        xmargen = 0

        If xcosto > 0 Then
            xmargen = ((Val("" & mytablex.Fields("pventa1")) - xcosto) * 100) / xcosto

        End If

        campo_precios(1).margen = "" & xmargen
        '--------
   
        '---------
        campo_precios(2).unidad = "" & mytablex.Fields("unidad2")
        campo_precios(2).factor = "" & mytablex.Fields("factor2")
        campo_precios(2).precio = "" & mytablex.Fields("pventa2")
        xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor2")))
        campo_precios(2).stock = "" & xbuf
        xcosto = 0

        If Val("" & mytablex.Fields("factor2")) > 0 Then
            xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
            xcosto = xcosto * Val("" & mytablex.Fields("factor2"))

        End If

        campo_precios(2).costo = "" & xcosto
        xmargen = 0

        If xcosto > 0 Then
            xmargen = ((Val("" & mytablex.Fields("pventa2")) - xcosto) * 100) / xcosto

        End If

        campo_precios(2).margen = "" & xmargen
   
        campo_precios(3).unidad = "" & mytablex.Fields("unidad3")
        campo_precios(3).factor = "" & mytablex.Fields("factor3")
        campo_precios(3).precio = "" & mytablex.Fields("pventa3")
        xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor3")))
        campo_precios(3).stock = "" & xbuf
        xcosto = 0

        If Val("" & mytablex.Fields("factor3")) > 0 Then
            xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
            xcosto = xcosto * Val("" & mytablex.Fields("factor3"))

        End If

        campo_precios(3).costo = "" & xcosto
        xmargen = 0

        If xcosto > 0 Then
            xmargen = ((Val("" & mytablex.Fields("pventa3")) - xcosto) * 100) / xcosto
            campo_precios(3).margen = "" & xmargen

        End If

        campo_precios(3).margen = "" & xmargen
   
        campo_precios(4).unidad = "" & mytablex.Fields("unidad4")
        campo_precios(4).factor = "" & mytablex.Fields("factor4")
        campo_precios(4).precio = "" & mytablex.Fields("pventa4")
        xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor4")))
        campo_precios(4).stock = "" & xbuf
        xcosto = 0

        If Val("" & mytablex.Fields("factor4")) > 0 Then
            xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
            xcosto = xcosto * Val("" & mytablex.Fields("factor4"))

        End If

        campo_precios(4).costo = "" & xcosto
        xmargen = 0

        If xcosto > 0 Then
            xmargen = ((Val("" & mytablex.Fields("pventa4")) - xcosto) * 100) / xcosto

        End If

        campo_precios(4).margen = "" & xmargen
   
        campo_precios(5).unidad = "" & mytablex.Fields("unidad5")
        campo_precios(5).factor = "" & mytablex.Fields("factor5")
        campo_precios(5).precio = "" & mytablex.Fields("pventa5")
        xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor5")))
        campo_precios(5).stock = "" & xbuf
        xcosto = 0

        If Val("" & mytablex.Fields("factor5")) > 0 Then
            xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
            xcosto = xcosto * Val("" & mytablex.Fields("factor5"))

        End If

        campo_precios(5).costo = "" & xcosto
        xmargen = 0

        If xcosto > 0 Then
            xmargen = ((Val("" & mytablex.Fields("pventa5")) - xcosto) * 100) / xcosto

        End If

        campo_precios(5).margen = "" & xmargen
   
        campo_precios(6).unidad = "" & mytablex.Fields("unidad6")
        campo_precios(6).factor = "" & mytablex.Fields("factor6")
        campo_precios(6).precio = "" & mytablex.Fields("pventa6")
        xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor6")))
        campo_precios(6).stock = "" & xbuf
        xcosto = 0

        If Val("" & mytablex.Fields("factor6")) > 0 Then
            xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
            xcosto = xcosto * Val("" & mytablex.Fields("factor6"))

        End If

        campo_precios(6).costo = "" & xcosto
        xmargen = 0

        If xcosto > 0 Then
            xmargen = ((Val("" & mytablex.Fields("pventa5")) - xcosto) * 100) / xcosto
         
        End If

        campo_precios(6).margen = "" & xmargen
   
        campo_precios(7).unidad = "" & mytablex.Fields("unidad7")
        campo_precios(7).factor = "" & mytablex.Fields("factor7")
        campo_precios(7).precio = "" & mytablex.Fields("pventa7")
        xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor7")))
        campo_precios(7).stock = "" & xbuf
        xcosto = 0

        If Val("" & mytablex.Fields("factor7")) > 0 Then
            xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
            xcosto = xcosto * Val("" & mytablex.Fields("factor7"))

        End If

        campo_precios(7).costo = "" & xcosto
        xmargen = 0

        If xcosto > 0 Then
            xmargen = ((Val("" & mytablex.Fields("pventa7")) - xcosto) * 100) / xcosto

        End If

        campo_precios(7).margen = "" & xmargen
        campo_precios(8).unidad = "" & mytablex.Fields("unidad8")
        campo_precios(8).factor = "" & mytablex.Fields("factor8")
        campo_precios(8).precio = "" & mytablex.Fields("pventa8")
        xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor8")))
        campo_precios(8).stock = "" & xbuf
        xcosto = 0

        If Val("" & mytablex.Fields("factor8")) > 0 Then
            xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
            xcosto = xcosto * Val("" & mytablex.Fields("factor8"))

        End If

        campo_precios(8).costo = "" & xcosto
        xmargen = 0

        If xcosto > 0 Then
            xmargen = ((Val("" & mytablex.Fields("pventa8")) - xcosto) * 100) / xcosto

        End If

        campo_precios(8).margen = "" & xmargen
   
        campo_precios(9).unidad = "" & mytablex.Fields("unidad9")
        campo_precios(9).factor = "" & mytablex.Fields("factor9")
        campo_precios(9).precio = "" & mytablex.Fields("pventa9")
        xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor9")))
        campo_precios(9).stock = "" & xbuf
        xcosto = 0

        If Val("" & mytablex.Fields("factor9")) > 0 Then
            xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
            xcosto = xcosto * Val("" & mytablex.Fields("factor9"))

        End If

        campo_precios(9).costo = "" & xcosto
        xmargen = 0

        If xcosto > 0 Then
            xmargen = ((Val("" & mytablex.Fields("pventa9")) - xcosto) * 100) / xcosto
         
        End If

        campo_precios(9).margen = "" & xmargen
        campo_precios(10).unidad = "" & mytablex.Fields("unidad10")
        campo_precios(10).factor = "" & mytablex.Fields("factor10")
        campo_precios(10).precio = "" & mytablex.Fields("pventa10")
        xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor10")))
        campo_precios(10).stock = "" & xbuf
        xcosto = 0

        If Val("" & mytablex.Fields("factor10")) > 0 Then
            xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
            xcosto = xcosto * Val("" & mytablex.Fields("factor10"))

        End If

        campo_precios(10).costo = "" & xcosto
        xmargen = 0

        If xcosto > 0 Then
            xmargen = ((Val("" & mytablex.Fields("pventa10")) - xcosto) * 100) / xcosto

        End If

        campo_precios(10).margen = "" & xmargen
        'margenes
        sw = 1

    End If

    mytablex.Close
 
    DBGrid4.refresh
    Frame5.Visible = True
    DBGrid4.SetFocus

End Sub

