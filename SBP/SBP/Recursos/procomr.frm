VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form procomr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimiento de Documentos"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   13845
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comprobantes "
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin VB.TextBox FECHAF 
         Height          =   495
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox fechai 
         Height          =   495
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Compras"
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ventas"
         Height          =   375
         Left            =   4080
         TabIndex        =   2
         Top             =   600
         Width           =   1215
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
         Height          =   735
         Left            =   5400
         MaskColor       =   &H00FFFFFF&
         Picture         =   "procomr.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   4695
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   8281
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
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   375
         Left            =   3960
         TabIndex        =   15
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label cantidad 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5640
         TabIndex        =   14
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label producto 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   10560
         TabIndex        =   12
         Top             =   840
         Width           =   2985
      End
      Begin VB.Label dolares 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   9000
         TabIndex        =   11
         Top             =   6480
         Width           =   1695
      End
      Begin VB.Label soles 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   9000
         TabIndex        =   10
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label Label30 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TotalDolares"
         Height          =   375
         Left            =   7320
         TabIndex        =   9
         Top             =   6480
         Width           =   1695
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TotalSoles"
         Height          =   375
         Left            =   7320
         TabIndex        =   8
         Top             =   6120
         Width           =   1695
      End
      Begin VB.Label Label27 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaFinal"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label26 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Menu lo8923 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "procomr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rcompraa As New ADODB.Recordset

Private Sub cmdGrabar_Click()

    Dim found As Integer

    Dim buf   As String

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub

    If Option2 = True Then
        buf = "select detalle.Local,detalle.Tipo,detalle.Serie,detalle.Numero,detalle.Codigo,detalle.fecha,detalle.Moneda,detalle.unidad,detalle.factor,detalle.cantidad,detalle.precio,detalle.Total,detalle.igv,detalle.hora,Usuario,Caja,Turno from detalle where  "
        buf = buf & " (detalle.acu='A' or detalle.acu='B' or detalle.acu='C' or detalle.acu='D' or detalle.acu='G' or detalle.acu='E' or detalle.acu='F' )   "

    End If

    If Option1 = True Then
        buf = "select detalle.Local,detalle.Tipo,detalle.Serie,detalle.Numero,detalle.Codigo,detalle.fecha,detalle.Moneda,detalle.unidad,detalle.factor,detalle.cantidad,detalle.precio,detalle.Total,detalle.igv,detalle.hora,Usuario,Caja,Turno from detalle where  "
        buf = buf & " (detalle.acu='J' or detalle.acu='K' or detalle.acu='L' or detalle.acu='M' or detalle.acu='P' or detalle.acu='N' or detalle.acu='O' )   "

    End If

    buf = buf & " AND  detalle.fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and detalle.fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    buf = buf & " and detalle.producto='" & producto & "'"
    buf = buf & " and detalle.estado='2' "
    buf = buf & " order by detalle.fecha,detalle.local,detalle.tipo,detalle.serie,str(detalle.numero)"

    'MsgBox buf
    If rcompraa.State = 1 Then rcompraa.Close
    rcompraa.Open buf, cn, adOpenStatic, adLockOptimistic
   
    Set dbGrid1.DataSource = rcompraa
    dbGrid1.columns(0).Width = 500
    dbGrid1.columns(1).Width = 500
    dbGrid1.columns(2).Width = 500
    dbGrid1.columns(3).Width = 1000
    dbGrid1.columns(4).Width = 1500
    dbGrid1.columns(5).Width = 1200
    dbGrid1.columns(6).Width = 500
    dbGrid1.columns(7).Width = 700
    dbGrid1.columns(8).Width = 700
    dbGrid1.columns(9).Width = 700
    dbGrid1.columns(10).Width = 700
    dbGrid1.columns(11).Width = 700
    dbGrid1.columns(12).Width = 700
               
    suma_total
    fechai.SetFocus

End Sub

Sub ir_inicio1()

    On Error GoTo cmd12_err

    rcompraa.MoveFirst
    Exit Sub
cmd12_err:
    Exit Sub

End Sub

Sub suma_total()

    Dim sdx  As Double

    Dim sdx1 As Double

    Dim sdx2 As Double

    sdx = 0
    sdx1 = 0
    sdx2 = 0
    ir_inicio1
    Do

        If rcompraa.EOF Then Exit Do
        sdx2 = sdx2 + Val("" & rcompraa.Fields("factor")) * Val("" & rcompraa.Fields("cantidad"))

        If "" & rcompraa.Fields("moneda") = "S" Then
            sdx = sdx + Val("" & rcompraa.Fields("total"))

        End If

        If "" & rcompraa.Fields("moneda") = "D" Then
            sdx1 = sdx1 + Val("" & rcompraa.Fields("total"))

        End If

        rcompraa.MoveNext
    Loop
    cantidad = Format(sdx2, "0.00")
    soles = Format(sdx, "0.00")
    dolares = Format(sdx1, "0.00")

End Sub

Private Sub Form_Load()
    fechai = Format(Now, "dd/mm/yyyy")
    fechaf = Format(Now, "dd/mm/yyyy")

End Sub

Private Sub lo8923_Click()
    procomr.Hide
    Unload procomr

End Sub
