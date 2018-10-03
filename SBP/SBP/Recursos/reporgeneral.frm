VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   Caption         =   "Form1"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14550
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
   ScaleHeight     =   10230
   ScaleWidth      =   14550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGenerarExcell 
      Caption         =   "&GenerarExcell"
      Height          =   495
      Left            =   12600
      TabIndex        =   7
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Frame FraCriterioBusqueda 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Criterio Busqueda"
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14535
      Begin VB.TextBox elcriterio 
         Height          =   2415
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   14175
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   615
         Left            =   12840
         TabIndex        =   3
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdEjecutar 
         Caption         =   "Ejecutar"
         Height          =   615
         Left            =   11640
         TabIndex        =   2
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdAyuda 
         Caption         =   "Ayuda"
         Height          =   615
         Left            =   10320
         TabIndex        =   1
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   1455
         Left            =   240
         TabIndex        =   5
         Top             =   2880
         Width           =   9735
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid2 
      Height          =   5655
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   14535
      _ExtentX        =   25638
      _ExtentY        =   9975
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   22
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
            LCID            =   10250
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
            LCID            =   10250
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rreport As New ADODB.Recordset

Dim tdiseno As New ADODB.Recordset

Dim mysnapx As New ADODB.Recordset

Sub casillas(buf As String)

    On Error GoTo cmd9012_err

    If mysnapx.State = 1 Then mysnapx.Close
    mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = mysnapx
    DBGrid2.refresh
    Exit Sub
cmd9012_err:
    MsgBox "Formato Consulta no Valido " + error$, 48, "Aviso"
    Exit Sub
 
End Sub

Private Sub cmdAyuda_Click()
    Label9 = "SELECT [ALL | DISTINCT ]"
    Label9 = Label9 + "             <nombre_campo> [{,<nombre_campo>}]"
    Label9 = Label9 + " FROM <nombre_tabla>|<nombre_vista>"
    Label9 = Label9 + "         [{,<nombre_tabla>|<nombre_vista>}]"
    Label9 = Label9 + " [WHERE <condicion> [{ AND|OR <condicion>}]]"
    Label9 = Label9 + " [GROUP BY <nombre_campo> [{,<nombre_campo >}]]"
    Label9 = Label9 + " [HAVING <condicion>[{ AND|OR <condicion>}]]"
    Label9 = Label9 + " [ORDER BY <nombre_campo>|<indice_campo> [ASC | DESC]"
    Label9 = Label9 + "                 [{,<nombre_campo>|<indice_campo> [ASC | DESC ]}]]"

End Sub

Private Sub cmdClose_Click()
    Form1.Hide
    Unload Form1

End Sub

Private Sub cmdEjecutar_Click()

    If Len(Trim(elcriterio)) = 0 Then Exit Sub
    casillas elcriterio
    'Frame1.Visible = False

End Sub

Private Sub cmdGenerarExcell_Click()

    Dim xlApp     As Excel.Application

    Dim xlBook    As Excel.Workbook

    Dim xlSheet   As Excel.Worksheet

    Dim sFileName As String

    On Error GoTo PROC_ERR

    'MsgBox "Please format Date column to Date and Time column to time in Excel.", vbInformation, "Message"
    If mysnapx.RecordCount = 0 Then
        MsgBox "No existen Datos ", 48, "Aviso"
        Exit Sub

    End If
    
    sFileName = App.path & "\Time Log as of " & CStr(Format(Now, "mm-dd-yyyy")) & ".xls"

    ExportRecordSetToExcel mysnapx, sFileName, "", "TimeLog"

    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(sFileName)
    xlApp.Application.Visible = True

PROC_EXIT:
    Set mysnapx = Nothing
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Exit Sub

PROC_ERR:
    MsgBox "Primero Ejecutar: " & Err.Number & " = " & Err.Description

    Resume PROC_EXIT

End Sub
