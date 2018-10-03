VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form texporta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importacion Datos"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Procesar"
      Height          =   495
      Left            =   6840
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Conectar"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7223
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
   Begin VB.Label registros 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Data Anterior"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "texporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cn1     As New ADODB.Connection

Dim mytablexyz As New ADODB.Recordset

Private Sub Command1_Click()

    On Error GoTo cmd1_error

    If cn1.State = 1 Then
        cn1.Close

    End If

    cn1.Open "Driver={SQL Server};Server=" & menup.vservidor & ";Database=calipso;Uid=sa;pwd="
    consulta_tabla
    Exit Sub
cmd1_error:
    MsgBox "No se puede conectar ,no Existe", 48, "Aviso"
    Exit Sub

End Sub

Sub consulta_tabla()

    Dim buf As String

    If mytablexyz.State = 1 Then mytablexyz.Close
    buf = "SELECT * FROM information_schema.tables"
    mytablexyz.Open buf, cn1, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytablexyz

End Sub

Private Sub Command2_Click()

    Dim buf As String

    Dim clavepaso

    Dim vr       As Long

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd9012_err

    clavepaso = InputBox("Ingrese Clave Paso", buf, tipoletra)

    If clavepaso <> "KALIPOSS" Then Exit Sub
    If mytablexyz.RecordCount = 0 Then Exit Sub

    mytablexyz.MoveFirst
    Do

        If mytablexyz.EOF Then Exit Do
        mytabley.Open "SELECT * FROM " & mytablexyz.Fields("table_name"), cn1, adOpenKeyset, adLockOptimistic

        If mytabley.RecordCount > 0 Then
            procesar_registros mytabley

        End If

        mytabley.Close
        Set mytabley = Nothing
        mytablexyz.MoveNext
    Loop
    Exit Sub
cmd9012_err:
    MsgBox "No se pudo procesar " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub procesar_registros(mytabley As ADODB.Recordset)

    Dim mytablex As New ADODB.Recordset

    Dim vr

    Dim sdx As Double

    On Error GoTo cmd9090_err

    cn.Execute ("delete from " & mytablexyz.Fields("table_name"))
    mytablex.Open "SELECT * FROM " & mytablexyz.Fields("table_name"), cn, adOpenKeyset, adLockOptimistic
    sdx = 0
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        graba_campo mytabley, mytablex
        mytablex.Update
        sdx = sdx + 1
        vr = DoEvents()
        registros = "" & sdx
        mytabley.MoveNext
    Loop
    mytablex.Close
    Set mytablex = Nothing
    Exit Sub
cmd9090_err:
    MsgBox "Aviso en procesar registro " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub graba_campo(mytabley As ADODB.Recordset, mytablex As ADODB.Recordset)

    On Error GoTo cmd7812_err

    Dim I As Integer

    For I = 0 To mytabley.Fields.count - 1
        mytablex.Fields(I) = mytabley.Fields(I)
    Next I

    Exit Sub
cmd7812_err:

End Sub
