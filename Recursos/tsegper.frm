VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tsegper 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Asistencia"
   ClientHeight    =   9390
   ClientLeft      =   150
   ClientTop       =   750
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8160
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox fechaf 
      DataField       =   "Code"
      DataSource      =   "dcFaculty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4440
      MaxLength       =   10
      TabIndex        =   7
      Top             =   480
      Width           =   2715
   End
   Begin VB.TextBox fechai 
      DataField       =   "Code"
      DataSource      =   "dcFaculty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4440
      MaxLength       =   10
      TabIndex        =   5
      Top             =   120
      Width           =   2715
   End
   Begin VB.TextBox codigo 
      DataField       =   "Code"
      DataSource      =   "dcFaculty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      MaxLength       =   11
      TabIndex        =   3
      Text            =   "%"
      Top             =   120
      Width           =   2715
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10920
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Pasar a Excel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   10440
      TabIndex        =   0
      Top             =   8760
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid dgLogin 
      Bindings        =   "tsegper.frx":0000
      Height          =   7845
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   13838
      _Version        =   393216
      AllowUpdate     =   0   'False
      DefColWidth     =   93
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Final"
      Height          =   195
      Left            =   3480
      TabIndex        =   8
      Top             =   480
      Width           =   825
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Inicio"
      Height          =   195
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   870
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu dflo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tsegper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmAttendanceLog
' DateTime  : 22/10/2006 00:45
' Author    :
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Dim strSQL As String

Dim vValue As Variant

Private Sub dflo33_Click()
    tsegper.Hide
    Unload tsegper

End Sub

Private Sub Form_Activate()
    Combo1.Clear
    Combo1.AddItem "Normal"
    Combo1.AddItem "Minimo Maximo"
    Combo1.ListIndex = 0
    cmdGo_Click

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' DateTime  : 22/10/2006 00:46
' Author    :
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    Dim X As Integer

    On Error GoTo PROC_ERR

    fechai = Format(Now, "dd/mm/yyyy")
    fechaf = Format(Now, "dd/mm/yyyy")
    'load initial result to data grid
    
    'Load fields to combo from datagrid
    'cmbField.AddItem "Codigo"
    'cmbField.AddItem "Fecha"

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & " = " & Err.Description

    Resume PROC_EXIT

End Sub

Private Sub cmdGo_Click()

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    If Combo1 = "Normal" Then
        consulta_normal
    Else
        LoadGrid
   
    End If
    
End Sub

Private Sub cmdPrint_Click()

    Dim mytablex  As New ADODB.Recordset

    Dim xlApp     As Excel.Application

    Dim xlBook    As Excel.Workbook

    Dim xlSheet   As Excel.Worksheet

    Dim sFileName As String

    On Error GoTo PROC_ERR

    'MsgBox "Please format Date column to Date and Time column to time in Excel.", vbInformation, "Message"
    mytablex.Open strSQL, cn, adOpenStatic, adLockOptimistic
    
    sFileName = App.path & "\Time Log as of " & CStr(Format(Now, "mm-dd-yyyy")) & ".xls"

    ExportRecordSetToExcel mytablex, sFileName, "", "TimeLog"

    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(sFileName)
    xlApp.Application.Visible = True

PROC_EXIT:
    Set mytablex = Nothing
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & " = " & Err.Description

    Resume PROC_EXIT

End Sub

Private Sub LoadGrid()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo PROC_ERR

    strSQL = "SELECT a.codigo,b.nombre,a.fecha,MIN(a.TimeIn) as TimeIn,MAX(a.TimeOut) as TimeOut FROM ingper a "
    strSQL = strSQL & "INNER JOIN vendedor b ON a.codigo = b.codigo  WHERE "
    strSQL = strSQL & "  a.fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    strSQL = strSQL & " and a.fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If codigo <> "%" Then
        strSQL = strSQL & "and a.codigo like ' " & codigo & "'"

    End If

    strSQL = strSQL & "GROUP BY a.fecha,a.codigo,b.nombre  ORDER BY a.fecha,a.codigo"
    cmdGo.Caption = "&Resetea"
    'MsgBox strSQL
    'Else
    '    strSQL = "SELECT a.codigo,b.Nombre ,a.fecha,a.TimeIn,a.TimeOut FROM ingper a "
    '    strSQL = strSQL & "INNER JOIN vendedor b ON a.codigo = b.codigo "
    '    strSQL = strSQL & "ORDER BY a.fecha Desc,a.codigo Asc"
    '    cmdGo.Caption = "&Buscar"
    'End If
    'MsgBox strSQL
    mytablex.Open strSQL, cn, adOpenStatic, adLockOptimistic
    Set dgLogin.DataSource = mytablex
    dgLogin.refresh
    Call GridSet

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & " = " & Err.Description

    Resume PROC_EXIT

End Sub

Sub consulta_normal()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo PROC_ERR1

    strSQL = "SELECT a.codigo,b.nombre,a.fecha,TimeIn, TimeOut FROM vendedor b "
    strSQL = strSQL & "inner JOIN ingper a ON b.codigo = a.codigo  WHERE "
    strSQL = strSQL & "  a.fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    strSQL = strSQL & " and a.fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If codigo <> "%" Then
        strSQL = strSQL & "and a.codigo like ' " & codigo & "'"

    End If

    strSQL = strSQL & "order by  a.fecha,a.codigo"
        
    'strSQL = "SELECT a.codigo,b.nombre,a.fecha,TimeIn, TimeOut FROM ingper a,vendedor b where a.codigo=b.codigo "
    'strSQL = strSQL & " order by  a.fecha,a.codigo"
        
    cmdGo.Caption = "&Resetea"
    mytablex.Open strSQL, cn, adOpenStatic, adLockOptimistic
    Set dgLogin.DataSource = mytablex
    dgLogin.refresh
    Call GridSet

PROC_EXIT1:
    Exit Sub

PROC_ERR1:
    MsgBox "Error: " & Err.Number & " = " & Err.Description

    Resume PROC_EXIT1

End Sub

Sub GridSet()

    On Error GoTo PROC_ERR

    dgLogin.columns(0).Width = 1000
    dgLogin.columns(1).Width = 4500
    dgLogin.columns(2).Width = 1200
    dgLogin.columns(3).Width = 1700
    dgLogin.columns(4).Width = 1700
    '
    '    dgLogin.Columns(5).Visible = False
    '    dgLogin.Columns(6).Visible = False
    
PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Error: " & Err.Number & " = " & Err.Description

    Resume PROC_EXIT

End Sub

