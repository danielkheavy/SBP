VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tinicia 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Proceso Inicializacion"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "Borrar Todo"
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "BorrarSeleccion"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Inicializa Movimientos"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   5415
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   120
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
   Begin VB.CommandButton Command1 
      Caption         =   "Inicializa Data Basico"
      Height          =   495
      Left            =   7320
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Menu flo44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tinicia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mytablexyz As New ADODB.Recordset

Private Sub Command1_Click()

    Dim vr

    If Command2.Visible = True Then Exit Sub
    If MsgBox("Desea Borrar", 1, "Aviso") <> 1 Then Exit Sub
    If InputBox("Llave de Paso ", "Control", "") <> "KALIPOS" Then Exit Sub
    If mytablexyz.RecordCount = 0 Then Exit Sub
    mytablexyz.MoveFirst
    Command2.Visible = True
    Do

        If mytablexyz.EOF Then Exit Do
        Command2.Caption = "Borrando:" & UCase("" & mytablexyz.Fields("table_name"))

        If Command2.Visible = False Then Exit Do
        'Select Case UCase("" & mytablexyz.Fields("table_name"))
        '       Case "CUENTAS", ""
        '       Case Else
        borra_linea_tablas
        'End Select
        mytablexyz.MoveNext
    Loop
    crear_vendedor
    Command2.Visible = False
    MsgBox "Proceso Realizado ", 48, "Aviso"
    Exit Sub
    'cmd9012_err:
    'Command2.Visible = False
    'Exit Sub

End Sub

Sub borra_linea_tablas()

    On Error GoTo cmd5644_err

    cn.Execute ("delete from " & mytablexyz.Fields("table_name"))
    Exit Sub
cmd5644_err:
    Exit Sub

End Sub

Private Sub Command2_Click()
    Command2.Visible = False

End Sub

Private Sub Command3_Click()

    Dim vr

    If Command2.Visible = True Then Exit Sub
    If MsgBox("Desea Borrar", 1, "Aviso") <> 1 Then Exit Sub
    If InputBox("Llave de Paso ", "Control", "") <> "KALIPOS" Then Exit Sub
    cn.Execute ("delete from factura")
    cn.Execute ("delete from detalle")
    cn.Execute ("delete from fpagov")
    cn.Execute ("delete from recibo")
    cn.Execute ("delete from cuentacd")
    cn.Execute ("delete from cuentapd")
    cn.Execute ("delete from cuentac")
    cn.Execute ("delete from cuentap")
    MsgBox "Proceso Realizado ", 48, "Aviso"

End Sub

Private Sub Command4_Click()

    On Error GoTo cmd567_err

    cn.Execute ("delete from " & mytablexyz.Fields(2))
    MsgBox "Proceso Realizado ", 48, "Aviso"
    Exit Sub
cmd567_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command5_Click()

    Dim vr

    If Command2.Visible = True Then Exit Sub
    If MsgBox("Desea Borrar", 1, "Aviso") <> 1 Then Exit Sub
    If InputBox("Llave de Paso ", "Control", "") <> "KALIPOSS" Then Exit Sub
    If mytablexyz.RecordCount = 0 Then Exit Sub
    mytablexyz.MoveFirst
    Command2.Visible = True
    Do

        If mytablexyz.EOF Then Exit Do
        Command2.Caption = "Borrando:" & UCase("" & mytablexyz.Fields("table_name"))

        If Command2.Visible = False Then Exit Do
        'If UCase("" & mytablexyz.Fields("table_name")) <> "EMPRESA" Then
        '   vr = DoEvents()
        borra_linea_tablas
        'End If
        mytablexyz.MoveNext
    Loop
    Command2.Visible = False
    MsgBox "Proceso Realizado ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub flo44_Click()

    If Command2.Visible = True Then Exit Sub

    tinicia.Hide
    Unload tinicia

End Sub

Private Sub Form_Load()

    Dim buf As String

    consulta_tabla

End Sub

Sub consulta_tabla()

    Dim buf As String

    If mytablexyz.State = 1 Then mytablexyz.Close
    buf = "SELECT * FROM information_schema.tables"
    mytablexyz.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytablexyz

End Sub

Sub crear_vendedor()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor ", cn, adOpenStatic, adLockOptimistic
    mytablex.AddNew
    mytablex.Fields("CODIGO") = "IDE"
    mytablex.Fields("NOMBRE") = "KALI INGENIERIA"
    mytablex.Fields("CLAVE") = "SCAN"
    mytablex.Fields("veclave") = "S"
    mytablex.Fields("vevend") = "S"
    mytablex.Fields("v1") = "S"
    mytablex.Fields("RW1") = "S"
    mytablex.Fields("local") = "01"
    mytablex.Update
    mytablex.Close
   
    mytablex.Open "select * from empresa ", cn, adOpenStatic, adLockOptimistic
    mytablex.AddNew
    mytablex.Fields("CODIGO") = "01"
    mytablex.Fields("NOMBRE") = "DEMO"
    mytablex.Update
    mytablex.Close

    mytablex.Open "select * from TLOCAL ", cn, adOpenStatic, adLockOptimistic
    mytablex.AddNew
    mytablex.Fields("CODIGO") = "01"
    mytablex.Fields("NOMBRE") = "PRINCIPAL"
    mytablex.Update
    mytablex.Close
   
    mytablex.Open "select * from bodega ", cn, adOpenStatic, adLockOptimistic
    mytablex.AddNew
    mytablex.Fields("CODIGO") = "01"
    mytablex.Fields("NOMBRE") = "PRINCIPAL"
    mytablex.Fields("local") = "01"
    mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
    mytablex.Update
    mytablex.Close
   
    mytablex.Open "select * from parame ", cn, adOpenStatic, adLockOptimistic
    mytablex.AddNew
    mytablex.Fields("CODIGO") = "01"
    mytablex.Fields("igv") = 18
    mytablex.Fields("DESCRIPCIO") = "PRINCIPAL"
    mytablex.Fields("bodega") = "01"
    mytablex.Update
    mytablex.Close
   
End Sub
