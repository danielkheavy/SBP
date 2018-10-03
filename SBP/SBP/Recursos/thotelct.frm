VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form thotelct 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EstadoCuenta"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox nhabitacion 
      Height          =   315
      Left            =   6000
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox idreserva 
      Height          =   375
      Left            =   9600
      MaxLength       =   10
      TabIndex        =   22
      Top             =   120
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid dbgrid2 
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   2566
      _Version        =   393216
      HeadLines       =   2
      RowHeight       =   29
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "IdcheckIn"
         Caption         =   "IdCheckin"
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
      BeginProperty Column01 
         DataField       =   "habitacion"
         Caption         =   "Habitacion"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "Fpago"
         Caption         =   "Fpago"
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
         DataField       =   "Monto"
         Caption         =   "Monto"
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
         DataField       =   "banco"
         Caption         =   "Banco"
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
      BeginProperty Column07 
         DataField       =   "Observa"
         Caption         =   "Observa"
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
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   3119.811
         EndProperty
      EndProperty
   End
   Begin VB.TextBox habitacion 
      Height          =   375
      Left            =   4320
      MaxLength       =   6
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Idcheckin 
      Height          =   375
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   4471
      _Version        =   393216
      HeadLines       =   2
      RowHeight       =   29
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
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "IdCheckIn"
         Caption         =   "IdCheckIn"
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
      BeginProperty Column01 
         DataField       =   "habitacion"
         Caption         =   "Habitacion"
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
      BeginProperty Column02 
         DataField       =   "Fecha"
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "Producto"
         Caption         =   "Producto"
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
         DataField       =   "Descripcio"
         Caption         =   "Descripcio"
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
         DataField       =   "Unidad"
         Caption         =   "Und"
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
         DataField       =   "Factor"
         Caption         =   "Factor"
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
         DataField       =   "cantidad"
         Caption         =   "Cant"
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
         DataField       =   "precio"
         Caption         =   "Precio"
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
         DataField       =   "Total"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2369.764
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   689.953
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column10 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dbgrid3 
      Height          =   1455
      Left            =   240
      TabIndex        =   18
      Top             =   5880
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   2566
      _Version        =   393216
      HeadLines       =   2
      RowHeight       =   29
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "IdReserva"
         Caption         =   "IdReserva"
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
      BeginProperty Column01 
         DataField       =   "habitacion"
         Caption         =   "Habitacion"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "Fpago"
         Caption         =   "Fpago"
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
         DataField       =   "Monto"
         Caption         =   "Monto"
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
         DataField       =   "banco"
         Caption         =   "Banco"
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
      BeginProperty Column07 
         DataField       =   "Observa"
         Caption         =   "Observa"
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
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   3119.811
         EndProperty
      EndProperty
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "IdReserva"
      Height          =   375
      Left            =   8280
      TabIndex        =   23
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   375
      Left            =   9360
      TabIndex        =   21
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label txtotalader 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10560
      TabIndex        =   20
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Adelantos Reserva"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label txsaldo 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   10560
      TabIndex        =   17
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
      Height          =   735
      Left            =   9360
      TabIndex        =   16
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Label txtotalade 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10560
      TabIndex        =   15
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   375
      Left            =   9360
      TabIndex        =   14
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label txtotal1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10560
      TabIndex        =   13
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   375
      Left            =   9360
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label txtotalco 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Consumos"
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label txtotalha 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Habitaciones"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Adelantos CheckIn"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Habitacion"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CheckIn"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Consumos habitaciones"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Menu flo882 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "thotelct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub flo882_Click()
    thotelct.Hide
    Unload thotelct

End Sub

Private Sub Form_Activate()

    Dim mytablex As New ADODB.Recordset

    nhabitacion.Clear
    nhabitacion.AddItem "%"

    mytablex.Open "select * from hotelcheckin where checkin=" & Val(idcheckin), cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        nhabitacion.AddItem "" & mytablex.Fields("Habitacion")
        mytablex.MoveNext
    Loop
    mytablex.Close
    nhabitacion.ListIndex = 0
 
    sql_cabeza

End Sub

Sub sql_cabeza()

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim sdx2     As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset
   
    mytablex.Open "select * from hotelconsumo where idcheckin=" & Trim("" & idcheckin), cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytablex
   
    mytabley.Open "select * from hotelanticipo where idcheckin=" & Trim("" & idcheckin), cn, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = mytabley
   
    'mytablez.Open "select * from hotelanticipo where idreserva=" & Trim("" & idreserva), cn, adOpenStatic, adLockOptimistic
    'Set dbgrid3.DataSource = mytablez
   
    sdx = 0
    sdx1 = 0
    sdx2 = 0
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("tipo") = "P" Then
            sdx = sdx + Val("" & mytablex.Fields("total"))

        End If

        If "" & mytablex.Fields("tipo") = "H" Then
            sdx1 = sdx1 + Val("" & mytablex.Fields("total"))

        End If

        mytablex.MoveNext
    Loop
    txtotalha = Format(sdx, "0.00")
    txtotalco = Format(sdx1, "0.00")
    txtotal1 = Format(sdx1 + sdx, "0.00")
    sdx = 0
    Do

        If mytabley.EOF Then Exit Do
        sdx = sdx + Val("" & mytabley.Fields("monto"))
        mytabley.MoveNext
    Loop
    txtotalade = Format(sdx, "0.00")
    'sdx = Val(txtotal1) - Val(txtotalade)
    'txsaldo = Format(sdx, "0.00")
   
    sdx = 0
    'Do
    'If mytablez.EOF Then Exit Do
    'sdx = sdx + Val("" & mytablez.Fields("monto"))
    'mytablez.MoveNext
    'Loop
    txtotalader = Format(sdx, "0.00")
    sdx = Val(txtotal1) - (Val(txtotalade) + Val(txtotalader))
    txsaldo = Format(sdx, "0.00")
   
End Sub
