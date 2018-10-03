VERSION 5.00
Begin VB.Form tupload 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Centralizacion"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   6240
      TabIndex        =   27
      Top             =   3960
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label validado 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1320
         TabIndex        =   29
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Validando"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   120
      TabIndex        =   15
      Top             =   5520
      Visible         =   0   'False
      Width           =   7815
      Begin VB.TextBox ipmaquina 
         Height          =   735
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   19
         Top             =   0
         Width           =   5055
      End
      Begin VB.TextBox usuario 
         Height          =   735
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   18
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox clave 
         Height          =   735
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   17
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox basedatos 
         Height          =   735
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   16
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "IP MAQUINA REMOTA"
         Height          =   735
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "USUARIO"
         Height          =   735
         Left            =   0
         TabIndex        =   22
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CLAVE"
         Height          =   735
         Left            =   0
         TabIndex        =   21
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BASE DATOS"
         Height          =   735
         Left            =   0
         TabIndex        =   20
         Top             =   2160
         Width           =   2535
      End
   End
   Begin VB.TextBox gerente 
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   2640
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   12
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ComboBox xlocal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   600
      Width           =   5055
   End
   Begin VB.TextBox fechai 
      Height          =   735
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1800
      Width           =   3495
   End
   Begin VB.TextBox fechaf 
      Height          =   735
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   2
      Top             =   2520
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Procesar"
      Height          =   975
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label noselicencia 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10920
      TabIndex        =   26
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label estado 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   2640
      TabIndex        =   25
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   24
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label xlocalx 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6840
      TabIndex        =   14
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USUARIO PERMITIDO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LOCAL"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Registros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Registros Procesados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Label total 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   2640
      TabIndex        =   5
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Label procesados 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   2640
      TabIndex        =   4
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Menu lfo9012 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tupload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim buf1     As String

    Dim found    As Integer

    Dim vr

    On Error GoTo cmd5612_err

    'If Trim(noselicencia) <> "S" Then
    '   MsgBox "Debe habilitar Licencia.. ", 48, "Aviso"
    '   Exit Sub
    'End If

    If Len(Trim(Combo1)) = "%" Then
        MsgBox "Combo sin Datos ", 48, "Aviso"
        Exit Sub

    End If

    If Len(Trim(xlocal)) = 0 Then
        MsgBox "Local sin Datos ", 48, "Aviso"
        Exit Sub

    End If

    If Not IsDate(fechai) Then
        MsgBox "Fecha Erronea ", 48, "Aviso"
        Exit Sub

    End If

    If Not IsDate(fechaf) Then
        MsgBox "Fecha Erronea ", 48, "Aviso"
        Exit Sub

    End If

    If Len(gerente) = 0 Then
        MsgBox "Usuario No existe ", 48, "Aviso"
        gerente.SetFocus
        Exit Sub

    End If

    found = verifica_usuario("" & gerente)

    If found = 0 Then
        MsgBox "Usuario No existe ", 48, "Aviso"
        Exit Sub

    End If

    buf1 = "select * from tlocal  "

    If xlocal <> "%" Then
        buf1 = buf1 & " where codigo='" & extra_loquesea(xlocal) & "'"

    End If

    Command1.Enabled = False
    mytablex.Open buf1, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        MsgBox "Local no encontrado ", 48, "Aviso"
        Command1.Enabled = True
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        xlocalx = "" & mytablex.Fields("codigo")
        ipmaquina = "" & mytablex.Fields("iprecepcion")
        usuario = "" & mytablex.Fields("usuariorecepcion")
        clave = "" & mytablex.Fields("claverecepcion")
        basedatos = "" & mytablex.Fields("basedatos")
        vr = DoEvents()
        found = conectar_remoto()

        If found = 0 Then GoTo sigue1

        If UCase$(Combo1) = UCase$("RecogerProductos") Then
            found = recoger_productos(0)

            If found = 0 Then
                GoTo sigue1

            End If

            'recoger_precios
            recoge_barras
            recoger_familia
            recoger_subfamilia
            recoger_marca
            recoger_categoria
            recoger_seccion

            'recoge_receta
        End If

        If UCase$(Combo1) = UCase$("RecogerClientes") Then
            found = recoger_clientes(0)

        End If

        If UCase$(Combo1) = UCase$("RecogerPersonal") Then
            recoger_personal

        End If

        If UCase$(Combo1) = UCase$("RecogerRecibos") Then
            found = recoger_recibos(0)

            If found = 0 Then
                GoTo sigue1

            End If

            recoger_fpago_recibos 0

        End If

        If UCase$(Combo1) = UCase$("RecogerVentas") Then
            found = recoger_ventas(0)

            If found = 0 Then
                GoTo sigue1

            End If

            recoger_detalle_ventas 0
            recoge_fpago_ventas 0

        End If

        If Combo1 = "RecogerCompras" Then
            found = recoger_ventas(1)

            If found = 0 Then
                GoTo sigue1

            End If

            recoger_detalle_ventas 1
            recoge_fpago_ventas 1

        End If

sigue1:
        mytablex.MoveNext
    Loop
    mytablex.Close
    Command1.Enabled = True
    MsgBox "Proceso Terminado ", 48, "Aviso"
    Exit Sub
cmd5612_err:
    MsgBox "Existe problemas de conexion,llame adm " + error$, 48, "Aviso"
    Command1.Enabled = True
    Exit Sub

End Sub

Function recoger_ventas(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'If MsgBox("Desea Realizar el proceso de recoger Ventas ", 1, "Aviso") <> 1 Then
    '   recoger_ventas = 0
    '   Exit Function
    'End If
    recoger_ventas = 1

    buf = "delete from factura where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If sw = 0 Then
        buf = buf & " and (acu='A' OR acu='B' or acu='C' or acu='D' or acu='G' OR acu='E' or acu='F') "

    End If

    If sw = 1 Then
        buf = buf & " and (acu='J' OR acu='K' or acu='L' or acu='M' or acu='P' OR acu='N' or acu='O') "

    End If

    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    cn.Execute (buf)
    buf = "select * from factura where "
    buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    If sw = 0 Then
        buf = buf & " AND ( acu='A' OR acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') "

    End If

    If sw = 1 Then
        buf = buf & " and (acu='J' OR acu='K' or acu='L' or acu='M' or acu='P' OR acu='N' or acu='O') "

    End If

    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from factura where 1=2 ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 2
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        'If sdx > 5 Then
        '   MsgBox "Demo solamente max 5", 48, "Aviso"
        '   Exit Do
        'End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    'MsgBox "Proceso Terminado ", 48, "Aviso"

End Function

Sub recoger_detalle_ventas(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'If MsgBox("Desea Realizar el proceso de recoger Ventas ", 1, "Aviso") <> 1 Then Exit Sub

    buf = "delete from detalle where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If sw = 0 Then
        buf = buf & " and (acu='A' OR acu='B' or acu='C' or acu='D' or acu='G' OR acu='E' or acu='F') "

    End If

    If sw = 1 Then
        buf = buf & " and (acu='J' OR acu='K' or acu='L' or acu='M' or acu='P' OR acu='N' or acu='O') "

    End If

    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    cn.Execute (buf)
    buf = "select * from detalle where "
    buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    If sw = 0 Then
        buf = buf & " AND ( acu='A' OR acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') "

    End If

    If sw = 1 Then
        buf = buf & " and (acu='J' OR acu='K' or acu='L' or acu='M' or acu='P' OR acu='N' or acu='O') "

    End If

    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from detalle where 1=2 ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 2
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

End Sub

Sub recoge_fpago_ventas(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'If MsgBox("Desea Realizar el proceso de recoger Ventas ", 1, "Aviso") <> 1 Then Exit Sub

    buf = "delete from fpagov where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If sw = 0 Then
        buf = buf & " and (acu='A' OR acu='B' or acu='C' or acu='D' or acu='G' OR acu='E' or acu='F') "

    End If

    If sw = 1 Then
        buf = buf & " and (acu='J' OR acu='K' or acu='L' or acu='M' or acu='P' OR acu='N' or acu='O') "

    End If

    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    cn.Execute (buf)
    buf = "select * from fpagov where "
    buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    If sw = 0 Then
        buf = buf & " AND ( acu='A' OR acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F') "

    End If

    If sw = 1 Then
        buf = buf & " and (acu='J' OR acu='K' or acu='L' or acu='M' or acu='P' OR acu='N' or acu='O') "

    End If

    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from fpagov where 1=2 ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 2
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

    'MsgBox "Proceso Terminado ", 48, "Aviso"
End Sub

Private Sub Form_Activate()
    noselicencia = licencia_remoto

End Sub

Private Sub Form_Load()

    Dim mytablex As New ADODB.Recordset

    'xlocal = glocal
    If cn1.State = 1 Then
        cn1.Close

    End If

    fechai = Format(Now, "dd/mm/yyyy")
    fechaf = Format(Now, "dd/mm/yyyy")
    xlocal.Clear
    xlocal.AddItem ""
    xlocal.AddItem "%"
    mytablex.Open "select * from tlocal ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        xlocal.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    xlocal.ListIndex = 0

    Combo1.Clear

    Combo1.AddItem "%"
    Combo1.AddItem "RecogerVentas"
    Combo1.AddItem "RecogerClientes"
    Combo1.AddItem "RecogerRecibos"
    Combo1.AddItem "RecogerCompras"
    Combo1.AddItem "RecogerProductos"
    Combo1.AddItem "RecogerPersonal"
    Combo1.AddItem "RecogerPedidos"

    Combo1.ListIndex = 0
    estado = "NO CONECTADO"

    If cn1.State = 1 Then
        estado = "CONECTADO"

    End If

    'conectado
End Sub

Public Function conectar_remoto()

    Dim buf As String

    On Error GoTo cmd8912_err

    If cn1.State = 1 Then
        cn1.Close

    End If

    buf = "Provider=SQLOLEDB;Persist Security Info=False;User ID=" & Trim(usuario) & ";Password=" & Trim(clave) & ";Initial Catalog=" & Trim(basedatos) & ";Data Source=" & Trim(ipmaquina)
    estado = "NO CONECTADO"
    cn1.CursorLocation = adUseClient
    cn1.Open buf
  
    If cn1.State = 1 Then
        conectar_remoto = 1
        estado = "CONECTADO"

    End If

    Exit Function
cmd8912_err:
    MsgBox "No se conecta  " + error$, 48, "Aviso"
    Exit Function

End Function

Private Sub lfo9012_Click()

    If Command1.Enabled = False Then Exit Sub
    If cn1.State = 1 Then
        cn1.Close

    End If

    tupload.Hide
    Unload tupload

End Sub

Private Sub xlocal_Click()

    Dim mytablex As New ADODB.Recordset

    ipmaquina = ""
    usuario = ""
    clave = ""
    basedatos = ""
    xlocalx = ""

End Sub

Function recoger_productos(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim xcostop  As Double

    Dim xcostou  As Double

    Dim sdx      As Double

    'on error goto cmd90_errr
    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    If MsgBox("Desea Realizar el proceso de recoger Productos ", 1, "Aviso") <> 1 Then
        recoger_productos = 0
        Exit Function

    End If

    recoger_productos = 1
    'buf = "delete from producto  "
    'cn.Execute (buf)
    buf = "select * from producto  "
    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "No existen Datos ", 48, "Aviso"
        mytablex.Close
        Exit Function

    End If

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    Do

        If mytablex.EOF Then Exit Do
        xcostop = Val("" & mytablex.Fields("costop"))
        xcostou = Val("" & mytablex.Fields("costou"))
denuevox:
        Set mytabley = Nothing
        mytabley.Open "select * from producto where producto='" & Trim("" & mytablex.Fields("producto")) & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount > 1 Then
            mytabley.Delete
            mytabley.Close
            GoTo denuevox

        End If

        If mytabley.RecordCount = 1 Then

            For I = 0 To mytablex.Fields.count - 2
                mytabley.Fields(I) = mytablex.Fields(I)
            Next I

            mytabley.Fields("costou") = xcostou
            mytabley.Fields("costop") = xcostop
            mytabley.Update
            GoTo sigamos

        End If

        If mytabley.RecordCount = 0 Then
            mytabley.AddNew

            For I = 0 To mytablex.Fields.count - 2
                mytabley.Fields(I) = mytablex.Fields(I)
            Next I

            mytabley.Fields("costou") = xcostou
            mytabley.Fields("costop") = xcostop
            mytabley.Update
 
        End If

        mytabley.Close
sigamos:
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    'mytabley.Close
    'MsgBox "Proceso Terminado ", 48, "Aviso"
    'Validamos los borrados
    'abrimos el local
    Frame2.Visible = True
    sdx = 0
    mytabley.Open "select producto from producto ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        If mytablex.State = 1 Then mytablex.Close
        Set mytablex = Nothing
        mytablex.Open "select producto from producto where producto='" & Trim("" & mytabley.Fields("producto")) & "'", cn1, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            MsgBox "" & mytabley.Fields("producto")
            mytabley.Delete

        End If

        mytablex.Close
        sdx = sdx + 1
        validado = "" & sdx
        vr = DoEvents()
        mytabley.MoveNext
    Loop
    mytabley.Close
    Frame2.Visible = False
    'MsgBox "abc"

End Function

Sub recoger_familia()

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    buf = "select * from FAMILIA  "
    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "No existen Datos ", 48, "Aviso"
        mytablex.Close
        Exit Sub

    End If

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    Do

        If mytablex.EOF Then Exit Do

        mytabley.Open "select * from FAMILIA where FAMILIA='" & Trim("" & mytablex.Fields("FAMILIA")) & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount = 0 Then
            mytabley.AddNew

            For I = 0 To mytablex.Fields.count - 1
                mytabley.Fields(I) = mytablex.Fields(I)
            Next I

            mytabley.Update
        Else

            For I = 0 To mytablex.Fields.count - 1
                mytabley.Fields(I) = mytablex.Fields(I)
            Next I

            mytabley.Update

        End If

        mytabley.Close
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Sub recoge_barras()

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    buf = "delete from productb  "
    cn.Execute (buf)
    buf = "select * from productb  "
    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from productb where 1=2 ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

End Sub

Sub recoger_subfamilia()

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    buf = "delete from subfamil  "
    cn.Execute (buf)
    buf = "select * from subfamil "
    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from subfamil where 1=2 ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

End Sub

Sub recoger_marca()

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    buf = "delete from marca  "
    cn.Execute (buf)
    buf = "select * from marca  "
    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from marca where 1=2 ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

End Sub

Sub recoger_seccion()

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    buf = "delete from seccion  "
    cn.Execute (buf)
    buf = "select * from seccion  "
    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from seccion where 1=2 ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

End Sub

Sub recoger_categoria()

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    buf = "delete from categori  "
    cn.Execute (buf)
    buf = "select * from categori  "
    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from categori where 1=2 ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

End Sub

Function recoger_recibos(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'If MsgBox("Desea Realizar el proceso de recoger Recibos ", 1, "Aviso") <> 1 Then
    '   recoger_recibos = 0
    '   Exit Function
    'End If
    recoger_recibos = 1

    buf = "delete from recibo where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If sw = 0 Then
        buf = buf & " and (acu='W' OR acu='V') "

    End If

    If sw = 1 Then
        buf = buf & " and (acu='W' OR acu='V') "

    End If

    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    cn.Execute (buf)
    buf = "select * from recibo where "
    buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    If sw = 0 Then
        buf = buf & " and (acu='W' OR acu='V') "

    End If

    If sw = 1 Then
        buf = buf & " and (acu='W' OR acu='V') "

    End If

    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic
    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from recibo where 1=2 ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 2
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    'MsgBox "Proceso Terminado ", 48, "Aviso"

End Function

Sub recoger_fpago_recibos(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'If MsgBox("Desea Realizar el proceso de recoger Ventas ", 1, "Aviso") <> 1 Then Exit Sub

    buf = "delete from fpagov where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If sw = 0 Then
        buf = buf & " and (acu='W' OR acu='V') "

    End If

    If sw = 1 Then
        buf = buf & " and (acu='W' OR acu='V') "

    End If

    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    cn.Execute (buf)
    buf = "select * from fpagov where "
    buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    If sw = 0 Then
        buf = buf & " and (acu='W' OR acu='V') "

    End If

    If sw = 1 Then
        buf = buf & " and (acu='W' OR acu='V') "

    End If

    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from fpagov where 1=2 ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 2
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

    'MsgBox "Proceso Terminado ", 48, "Aviso"
End Sub

Function verifica_usuario(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where clave='" & buf & "' and conexionremota='S'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        verifica_usuario = 1

    End If

    mytablex.Close

End Function

Function recoger_clientes(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'buf = "delete from clientes  "
    'cn.Execute (buf)
    buf = "select * from clientes  where len(codigo)=8 or len(codigo)=11"
    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    Do

        If mytablex.EOF Then Exit Do
        If mytabley.State = 1 Then
            mytabley.Close

        End If

        Set mytabley = Nothing
        mytabley.Open "select * from clientes where codigo='" & mytablex.Fields("codigo") & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount = 0 Then
            mytabley.AddNew

            For I = 0 To mytablex.Fields.count - 3
                mytabley.Fields(I) = mytablex.Fields(I)
            Next I

            mytabley.Update

        End If

        mytabley.Close
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        'End If
        mytablex.MoveNext
    Loop
    mytablex.Close

End Function

Sub recoger_personal()

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    buf = "delete from vendedor  "
    cn.Execute (buf)
    buf = "select * from vendedor  "
    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from vendedor where 1=2 ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

End Sub

Sub recoger_precios()

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    buf = "delete from precios  "
    cn.Execute (buf)
    buf = "select * from precios  "
    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from precios where 1=2 ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

End Sub

Sub recoge_receta()

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    buf = "delete from receta  "
    cn.Execute (buf)
    buf = "select * from receta  "
    mytablex.Open buf, cn1, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from receta where 1=2 ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 1
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close

End Sub

