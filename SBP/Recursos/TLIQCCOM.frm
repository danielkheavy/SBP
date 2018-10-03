VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tliqccom 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Comisiones"
   ClientHeight    =   8715
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   15075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   15075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   14865
      TabIndex        =   18
      Top             =   0
      Width           =   14925
      Begin VB.ComboBox vendedor 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   0
         Width           =   1815
      End
      Begin VB.ComboBox caja 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   0
         Width           =   1815
      End
      Begin VB.ComboBox cajero 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6720
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox tipo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9000
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Seleccionar Ventas"
         Height          =   735
         Left            =   11040
         TabIndex        =   28
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   22
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   21
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox producto 
         Height          =   375
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   20
         Text            =   "%"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   4200
         MaxLength       =   11
         TabIndex        =   19
         Text            =   "%"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         Height          =   375
         Left            =   6120
         TabIndex        =   38
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Usuario"
         Height          =   375
         Left            =   6120
         TabIndex        =   37
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   375
         Left            =   8520
         TabIndex        =   36
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Final"
         Height          =   375
         Left            =   0
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FormaPago"
         Height          =   375
         Left            =   0
         TabIndex        =   25
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         Height          =   375
         Left            =   3000
         TabIndex        =   24
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   375
         Left            =   3000
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   15015
      TabIndex        =   14
      Top             =   8220
      Width           =   15075
      Begin VB.CommandButton Command3 
         Caption         =   "Procesar OT"
         Height          =   375
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Producto"
         Height          =   375
         Left            =   9120
         TabIndex        =   17
         Top             =   0
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "TablaVentas"
         Height          =   375
         Left            =   10440
         TabIndex        =   16
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Procesar Liquidacion"
         Height          =   375
         Left            =   11880
         TabIndex        =   15
         Top             =   0
         Width           =   1695
      End
   End
   Begin MSDataGridLib.DataGrid dbGrid1 
      Height          =   5895
      Left            =   120
      TabIndex        =   41
      Top             =   1200
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   10398
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
      ColumnCount     =   18
      BeginProperty Column00 
         DataField       =   "Vendedor"
         Caption         =   "Vendedor"
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
         DataField       =   "Moneda"
         Caption         =   "M"
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
         DataField       =   "total"
         Caption         =   "Total"
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
      BeginProperty Column03 
         DataField       =   "Impuesto"
         Caption         =   "Impuesto"
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
      BeginProperty Column04 
         DataField       =   "Subtotal"
         Caption         =   "Subtotal"
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
      BeginProperty Column05 
         DataField       =   "Comision"
         Caption         =   "Comision"
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
      BeginProperty Column06 
         DataField       =   "Vtaneta"
         Caption         =   "VtaNeta"
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
      BeginProperty Column07 
         DataField       =   "TCosto"
         Caption         =   "TCosto"
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
      BeginProperty Column08 
         DataField       =   "ganancia"
         Caption         =   "Ganancia"
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
      BeginProperty Column09 
         DataField       =   "fecha"
         Caption         =   "Fecha"
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
      BeginProperty Column10 
         DataField       =   "Hora"
         Caption         =   "Hora"
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
      BeginProperty Column11 
         DataField       =   "Local"
         Caption         =   "Local"
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
      BeginProperty Column12 
         DataField       =   "Tipo"
         Caption         =   "Tipo"
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
      BeginProperty Column13 
         DataField       =   "Serie"
         Caption         =   "Serie"
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
      BeginProperty Column14 
         DataField       =   "Numero"
         Caption         =   "Numero"
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
      BeginProperty Column15 
         DataField       =   "Estado"
         Caption         =   "E"
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
      BeginProperty Column16 
         DataField       =   "Codigo"
         Caption         =   "Codigo"
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
      BeginProperty Column17 
         DataField       =   "Nombre"
         Caption         =   "Nombre"
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
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   255.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   750.047
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   2670.236
         EndProperty
      EndProperty
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label txnormal 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1080
      TabIndex        =   31
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Ot"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label txot 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1080
      TabIndex        =   29
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label ganancia 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   12360
      TabIndex        =   13
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ganancia"
      Height          =   375
      Left            =   11400
      TabIndex        =   12
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label costo 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10200
      TabIndex        =   11
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Costo"
      Height          =   375
      Left            =   9240
      TabIndex        =   10
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label vtaneta 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10200
      TabIndex        =   9
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VtaNeta"
      Height          =   375
      Left            =   9240
      TabIndex        =   8
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label comision 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8040
      TabIndex        =   7
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comision"
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label subtotal 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8040
      TabIndex        =   5
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Subtotal"
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label impuesto 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Impuesto"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label total 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   7200
      Width           =   975
   End
   Begin VB.Menu fclo44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tliqccom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rexplorap As New ADODB.Recordset

Private Sub Command1_Click()

    If MsgBox("Procesar Liquidacion", 1, "Aviso") <> 1 Then Exit Sub
    sql_procesa

End Sub

Private Sub Command2_Click()
    sql_mes

End Sub

Private Sub fclo44_Click()
    tliqccom.Hide
    Unload tliqccom

End Sub

Private Sub Form_Activate()
    sql_mes

End Sub

Sub carga_inicial()

    Dim mytablex As New ADODB.Recordset

    vendedor.Clear
    cajero.Clear
    cajero.AddItem "%"
    vendedor.AddItem "%"
    mytablex.Open "select * from vendedor ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    cajero.ListIndex = 0
    vendedor.ListIndex = 0
    tipo.Clear
    tipo.AddItem "TODOS"
    tipo.AddItem "FACTURAS"
    tipo.AddItem "NOTAS"
    tipo.ListIndex = 0
    caja.Clear
    caja.AddItem "%"
    mytablex.Open "select * from parameca ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("terminal") = "C" Then
            caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("DESCRIPCIO")

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    caja.ListIndex = 0

End Sub

Private Sub Form_Load()
    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    'MsgBox fechai
    fechaf = Format(Now, "dd/mm/yyyy")
    carga_inicial

End Sub

Sub sql_mes()

    Dim buf As String

    buf = "select * from factura where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If vendedor <> "%" Then
        buf = buf & " and vendedor='" & extra_loquesea(vendedor) & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja='" & extra_loquesea(caja) & "'"

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario='" & extra_loquesea(cajero) & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If tipo = "FACTURAS" Then
        buf = buf & " and (acu='B' or acu='D' OR acu='A' or acu='C' ) "

    End If

    If tipo = "NOTAS" Then
        buf = buf & " and (acu='G') "

    End If

    If tipo = "TODOS" Then
        buf = buf & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' ) "

    End If

    buf = buf & " and estado='2' order by vendedor,tipo,fecha,str(numero)"

    'MsgBox buf
    If rexplorap.State = 1 Then rexplorap.Close
    rexplorap.Open buf, cn, adOpenStatic, adLockOptimistic

    If rexplorap.EOF = True And rexplorap.BOF = True Then

    End If

    Set dbGrid1.DataSource = rexplorap
    sql_procesa

End Sub

Function procesa_producto()

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    mytablex.Open "select * from producto where producto='" & "" & rexplorap.Fields("producto") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If Option1.Value = True Then
            'rexplorap.Edit
            sdx = Val("" & rexplorap.Fields("subtotal")) * Val("" & mytablex.Fields("comision")) / 100 'comision
            rexplorap.Fields("comision") = sdx
            rexplorap.Fields("tcosto") = Val("" & rexplorap.Fields("cantidad")) * Val("" & mytablex.Fields("costou"))
            rexplorap.Fields("vtaneta") = Val("" & rexplorap.Fields("subtotal")) - Val("" & rexplorap.Fields("comision"))
            rexplorap.Fields("ganancia") = Val("" & rexplorap.Fields("vtaneta")) - Val("" & rexplorap.Fields("tcosto"))
            rexplorap.Update

        End If

        If Option2.Value = True Then  'tabla de comisiones
            'rexplorap.Edit
            sdx = Val("" & rexplorap.Fields("subtotal")) * busca_vendedor() / 100 'comision
            rexplorap.Fields("comision") = sdx
            rexplorap.Fields("tcosto") = Val("" & rexplorap.Fields("cantidad")) * Val("" & mytablex.Fields("costou"))
            rexplorap.Fields("vtaneta") = Val("" & rexplorap.Fields("subtotal")) - Val("" & rexplorap.Fields("comision"))
            rexplorap.Fields("ganancia") = Val("" & rexplorap.Fields("vtaneta")) - Val("" & rexplorap.Fields("tcosto"))
            rexplorap.Update

        End If

    End If

    mytablex.Close

End Function

Sub sql_procesa()

    Dim found As Integer

    Dim vr

    Dim sdx     As Double

    Dim sdx1    As Double

    Dim sdx2    As Double

    Dim sdx3    As Double

    Dim sdx4    As Double

    Dim sdx5    As Double

    Dim sdx6    As Double

    Dim xnormal As Double

    Dim xot     As Double

    xnormal = 0
    xot = 0
    'ir_inicio
    sdx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    sdx4 = 0
    sdx5 = 0
    sdx6 = 0
    Do

        If rexplorap.EOF Then Exit Do
        found = procesa_producto()

        If Val("" & rexplorap.Fields("nroprecio")) = 1 Then
            If Len("" & rexplorap.Fields("tipo1")) = 0 Then
                xnormal = xnormal + Val("" & rexplorap.Fields("total"))

            End If

            If Len("" & rexplorap.Fields("tipo1")) > 0 Then
                xot = xot + Val("" & rexplorap.Fields("total"))

            End If

        End If

        sdx = sdx + Val("" & rexplorap.Fields("total"))
        sdx1 = sdx1 + Val("" & rexplorap.Fields("Impuesto"))
        sdx2 = sdx2 + Val("" & rexplorap.Fields("subtotal"))
        sdx3 = sdx3 + Val("" & rexplorap.Fields("comision"))
        sdx4 = sdx4 + Val("" & rexplorap.Fields("vtaneta"))
        sdx5 = sdx5 + Val("" & rexplorap.Fields("tcosto"))
        sdx6 = sdx6 + Val("" & rexplorap.Fields("ganancia"))
        vr = DoEvents()
        rexplorap.MoveNext
    Loop
 
    total = Format(sdx, "0.00")
    impuesto = Format(sdx1, "0.00")
    subtotal = Format(sdx2, "0.00")
    comision = Format(sdx3, "0.00")
    vtaneta = Format(sdx4, "0.00")
    costo = Format(sdx5, "0.00")
    ganancia = Format(sdx6, "0.00")
    txnormal = Format(xnormal, "0.00")
    txot = Format(xot, "0.00")

End Sub

Sub ir_inicio()

    On Error GoTo cmd1_err

    rexplorap.MoveFirst
    Exit Sub
cmd1_err:
    Exit Sub

End Sub

Function busca_vendedor() As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where codigo='" & "" & rexplorap.Fields("vendedor") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If Val("" & rexplorap.Fields("total")) >= Val("" & mytablex.Fields("ini1")) And Val("" & rexplorap.Fields("total")) <= Val("" & mytablex.Fields("ini1")) Then
            busca_vendedor = Val("" & mytablex.Fields("por1"))
            GoTo am1

        End If

        If Val("" & rexplorap.Fields("total")) >= Val("" & mytablex.Fields("ini2")) And Val("" & rexplorap.Fields("total")) <= Val("" & mytablex.Fields("ini2")) Then
            busca_vendedor = Val("" & mytablex.Fields("por2"))
            GoTo am1

        End If

        If Val("" & rexplorap.Fields("total")) >= Val("" & mytablex.Fields("ini3")) And Val("" & rexplorap.Fields("total")) <= Val("" & mytablex.Fields("ini3")) Then
            busca_vendedor = Val("" & mytablex.Fields("por3"))
            GoTo am1

        End If

        If Val("" & rexplorap.Fields("total")) >= Val("" & mytablex.Fields("ini4")) And Val("" & rexplorap.Fields("total")) <= Val("" & mytablex.Fields("ini4")) Then
            busca_vendedor = Val("" & mytablex.Fields("por4"))
            GoTo am1

        End If

am1:

    End If

    mytablex.Close

End Function
