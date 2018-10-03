VERSION 5.00
Begin VB.Form tload 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envio Centralizacion"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   120
      TabIndex        =   15
      Top             =   5520
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
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6840
      TabIndex        =   27
      Top             =   1560
      Width           =   105
   End
   Begin VB.Label noselicencia 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10800
      TabIndex        =   26
      Top             =   8280
      Width           =   735
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
      BackColor       =   &H00C0C0C0&
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
Attribute VB_Name = "tload"
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
        ipmaquina = "" & mytablex.Fields("iptransmision")
        usuario = "" & mytablex.Fields("usuariotransmision")
        clave = "" & mytablex.Fields("clavetransmision")
        basedatos = "" & mytablex.Fields("basedatostransmision")
        vr = DoEvents()
        found = conectar_remoto()

        If found = 0 Then GoTo sigue1
        If UCase$(Combo1) = UCase$("EnviarRequerimiento") Then
            found = Enviar_requerimiento(0)

            If found = 0 Then
                'MsgBox "No se pudo realizar..", 48, "Aviso"
                'Exit Sub
                GoTo sigue1

            End If

            Enviar_detalle_requerimiento 0

        End If

        If UCase$(Combo1) = UCase$("EnviarPedidos") Then
            found = Enviar_pedidos(0)

            If found = 0 Then
                'MsgBox "No se pudo realizar..", 48, "Aviso"
                'Exit Sub
                GoTo sigue1

            End If

            Enviar_detalle_pedidos 0

        End If

        If UCase$(Combo1) = UCase$("EnviarGuias") Then
            found = Enviar_guias(0)

            If found = 0 Then
                'MsgBox "No se pudo realizar..", 48, "Aviso"
                'Exit Sub
                GoTo sigue1

            End If

            Enviar_detalle_guias 0

        End If

        If UCase$(Combo1) = UCase$("EnviarProductos") Then
            found = Enviar_productos(0)

            If found = 0 Then
                GoTo sigue1

            End If

            Enviar_familia
            Enviar_subfamilia
            Enviar_marca
            Enviar_categoria
            Enviar_precios1

        End If

        If UCase$(Combo1) = UCase$("EnviarClientes") Then
            found = Enviar_clientes(0)

        End If

        If UCase$(Combo1) = UCase$("EnviarPrecios") Then
            found = Enviar_precios(0)

        End If

        If UCase$(Combo1) = UCase$("EnviarPersonal") Then
            found = Enviar_personal(0)

        End If

        If UCase$(Combo1) = UCase$("EnviarRecibos") Then
            found = Enviar_recibos(0)

            If found = 0 Then
                GoTo sigue1

            End If

            Enviar_fpago_recibos 0

        End If

        If UCase$(Combo1) = UCase$("EnviarVentass") Then
            found = Enviar_ventas(0)

            If found = 0 Then
                GoTo sigue1

            End If

            Enviar_detalle_ventas 0
            recoge_fpago_ventas 0
            found = Enviar_guias(0)

            If found = 0 Then
                GoTo sigue1

            End If

            Enviar_detalle_guias 0
            found = Enviar_recibos(0)

            If found = 0 Then
                GoTo sigue1

            End If

            Enviar_fpago_recibos 0

        End If

        If UCase$(Combo1) = UCase$("EnviarVentas") Then
            found = Enviar_ventas(0)

            If found = 0 Then
                GoTo sigue1

            End If

            Enviar_detalle_ventas 0
            recoge_fpago_ventas 0

        End If

        If Combo1 = "EnviarCompras" Then
            found = Enviar_ventas(1)

            If found = 0 Then
                GoTo sigue1

            End If

            Enviar_detalle_ventas 1
            recoge_fpago_ventas 1

        End If

sigue1:
        mytablex.MoveNext
    Loop
    mytablex.Close
    Command1.Enabled = True
    MsgBox "Envio,Proceso Terminado ", 48, "Aviso"
    tload.Hide
    Unload tload

End Sub

Function Enviar_ventas(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'If MsgBox("Desea Realizar el proceso de Enviar Ventas ", 1, "Aviso") <> 1 Then
    '   Enviar_ventas = 0
    '   Exit Function
    'End If
    Enviar_ventas = 1
    'sw = 0

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

    cn1.Execute (buf)

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

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from factura where 1=2 ", cn1, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 2
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        'If sdx > 5 Then
        '   MsgBox "No existe Licencia,Solo Prueba Solo 5 Transacciones ", 48, "Aviso"
        '   Exit Do
        'End If
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    'MsgBox "Proceso Terminado ", 48, "Aviso"

End Function

Sub Enviar_detalle_ventas(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'If MsgBox("Desea Realizar el proceso de Enviar Ventas ", 1, "Aviso") <> 1 Then Exit Sub

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

    cn1.Execute (buf)
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

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from detalle where 1=2 ", cn1, adOpenStatic, adLockOptimistic
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

    'If MsgBox("Desea Realizar el proceso de Enviar Ventas ", 1, "Aviso") <> 1 Then Exit Sub

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

    cn1.Execute (buf)
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

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from fpagov where 1=2 ", cn1, adOpenStatic, adLockOptimistic
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

    If Label12.Caption = "CIERRE" Then
        Combo1.Clear
        Combo1.AddItem "EnviarVentass"
        Combo1.ListIndex = 0
        Combo1.Enabled = False

    End If

    If Label12.Caption = "REQUERIMIENTO" Then
        Combo1.Clear
        Combo1.AddItem "EnviarRequerimiento"
        Combo1.ListIndex = 0
        Combo1.Enabled = False

    End If

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

    Combo1.AddItem "EnviarVentas"
    Combo1.AddItem "EnviarClientes"
    Combo1.AddItem "EnviarRecibos"
    Combo1.AddItem "EnviarCompras"
    Combo1.AddItem "EnviarProductos"
    Combo1.AddItem "EnviarPrecios"
    Combo1.AddItem "EnviarRequerimiento"
    Combo1.AddItem "EnviarGuias"
    Combo1.AddItem "EnviarPersonal"
    Combo1.AddItem "EnviarPedidos"

    Combo1.AddItem "%"

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
    'MsgBox buf
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

    tload.Hide
    Unload tload

End Sub

Private Sub xlocal_Click()

    Dim mytablex As New ADODB.Recordset

    ipmaquina = ""
    usuario = ""
    clave = ""
    basedatos = ""
    xlocalx = ""

End Sub

Function Enviar_productos(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    If MsgBox("Desea Realizar el proceso de Enviar Productos ", 1, "Aviso") <> 1 Then
        Enviar_productos = 0
        Exit Function

    End If

    Enviar_productos = 1
    buf = "delete from producto  "
    cn1.Execute (buf)
    buf = "select * from producto  "
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from producto where 1=2 ", cn1, adOpenStatic, adLockOptimistic
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

Sub Enviar_familia()

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    buf = "delete from familia  "
    cn1.Execute (buf)
    buf = "select * from familia  "
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from familia where 1=2 ", cn1, adOpenStatic, adLockOptimistic
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

Sub Enviar_subfamilia()

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    buf = "delete from subfamil  "
    cn1.Execute (buf)
    buf = "select * from subfamil "
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from subfamil where 1=2 ", cn1, adOpenStatic, adLockOptimistic
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

Sub Enviar_marca()

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    buf = "delete from marca  "
    cn1.Execute (buf)
    buf = "select * from marca  "
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from marca where 1=2 ", cn1, adOpenStatic, adLockOptimistic
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

Sub Enviar_precios1()

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    buf = "delete from precios  "
    cn1.Execute (buf)
    buf = "select * from precios  "
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from precios where 1=2 ", cn1, adOpenStatic, adLockOptimistic
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

Sub Enviar_categoria()

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    buf = "delete from categori  "
    cn1.Execute (buf)
    buf = "select * from categori  "
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from categori where 1=2 ", cn1, adOpenStatic, adLockOptimistic
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

Function Enviar_recibos(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'If MsgBox("Desea Realizar el proceso de Enviar Recibos ", 1, "Aviso") <> 1 Then
    '   Enviar_recibos = 0
    '   Exit Function
    'End If
    Enviar_recibos = 1

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

    cn1.Execute (buf)
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

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from recibo where 1=2 ", cn1, adOpenStatic, adLockOptimistic
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

Sub Enviar_fpago_recibos(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'If MsgBox("Desea Realizar el proceso de Enviar Ventas ", 1, "Aviso") <> 1 Then Exit Sub

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

    cn1.Execute (buf)
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

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from fpagov where 1=2 ", cn1, adOpenStatic, adLockOptimistic
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

Function Enviar_clientes(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'buf = "delete from clientes  "
    'cn.Execute (buf)
    buf = "select * from clientes  where len(codigo)=8 or len(codigo)=11"
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    Do

        If mytablex.EOF Then Exit Do
        If mytabley.State = 1 Then
            mytabley.Close

        End If

        Set mytabley = Nothing
        mytabley.Open "select * from clientes where codigo='" & mytablex.Fields("codigo") & "'", cn1, adOpenStatic, adLockOptimistic

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

Function Enviar_requerimiento(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'If MsgBox("Desea Realizar el proceso de Enviar Ventas ", 1, "Aviso") <> 1 Then
    '   Enviar_ventas = 0
    '   Exit Function
    'End If
    Enviar_requerimiento = 1
    sw = 0

    buf = "delete from crequisa where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and (acu='Q' ) "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    cn1.Execute (buf)

    buf = "select * from crequisa where "
    buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"
    buf = buf & " AND ( acu='Q') "
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from crequisa where 1=2 ", cn1, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 2
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        'If sdx > 5 Then
        '   MsgBox "No existe Licencia,Solo Prueba Solo 5 Transacciones ", 48, "Aviso"
        '   Exit Do
        'End If
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    'MsgBox "Proceso Terminado ", 48, "Aviso"

End Function

Sub Enviar_detalle_requerimiento(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'If MsgBox("Desea Realizar el proceso de Enviar Ventas ", 1, "Aviso") <> 1 Then Exit Sub

    buf = "delete from drequisa where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and (acu='Q')  "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    cn1.Execute (buf)
    buf = "select * from drequisa where "
    buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"
    buf = buf & " AND ( acu='Q') "

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from drequisa where 1=2 ", cn1, adOpenStatic, adLockOptimistic
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

Function Enviar_guias(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'If MsgBox("Desea Realizar el proceso de Enviar Ventas ", 1, "Aviso") <> 1 Then
    '   Enviar_ventas = 0
    '   Exit Function
    'End If
    Enviar_guias = 1
    sw = 0

    buf = "delete from factura where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and (acu='S' or acu='T' ) "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    cn1.Execute (buf)

    buf = "select * from factura where "
    buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"
    buf = buf & " AND ( acu='S' or acu='T') "
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from factura where 1=2 ", cn1, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 2
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        'If sdx > 5 Then
        '   MsgBox "No existe Licencia,Solo Prueba Solo 5 Transacciones ", 48, "Aviso"
        '   Exit Do
        'End If
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    'MsgBox "Proceso Terminado ", 48, "Aviso"

End Function

Sub Enviar_detalle_guias(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'If MsgBox("Desea Realizar el proceso de Enviar Ventas ", 1, "Aviso") <> 1 Then Exit Sub

    buf = "delete from detalle where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and (acu='S' or acu='T')  "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    cn1.Execute (buf)
    buf = "select * from detalle where "
    buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"
    buf = buf & " AND ( acu='S' or acu='T') "

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from detalle where 1=2 ", cn1, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 3
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

Function Enviar_precios(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    If MsgBox("Desea Realizar el proceso de Enviar Precios ", 1, "Aviso") <> 1 Then
        Enviar_precios = 0
        Exit Function

    End If

    Enviar_precios = 1
    buf = "delete from precios  "
    cn1.Execute (buf)
    buf = "select * from precios  "
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from precios where 1=2 ", cn1, adOpenStatic, adLockOptimistic
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

Function Enviar_personal(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    buf = "delete from vendedor  "
    cn1.Execute (buf)
    buf = "select * from vendedor  "
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from vendedor where 1=2 ", cn1, adOpenStatic, adLockOptimistic
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

End Function

Function Enviar_pedidos(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'If MsgBox("Desea Realizar el proceso de Enviar Ventas ", 1, "Aviso") <> 1 Then
    '   Enviar_ventas = 0
    '   Exit Function
    'End If
    Enviar_pedidos = 1
    sw = 0

    buf = "delete from cpedidov where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and (acu='I' ) "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    cn1.Execute (buf)

    buf = "select * from cpedidov where "
    buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"
    buf = buf & " AND ( acu='I') "
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from cpedidov where 1=2 ", cn1, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew

        For I = 0 To mytablex.Fields.count - 2
            mytabley.Fields(I) = mytablex.Fields(I)
        Next I

        mytabley.Update
        sdx = sdx + 1
        procesados = "" & sdx
        'If sdx > 5 Then
        '   MsgBox "No existe Licencia,Solo Prueba Solo 5 Transacciones ", 48, "Aviso"
        '   Exit Do
        'End If
        vr = DoEvents()
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    'MsgBox "Proceso Terminado ", 48, "Aviso"

End Function

Sub Enviar_detalle_pedidos(sw As Integer)

    Dim I   As Integer

    Dim buf As String

    Dim vr

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    'If MsgBox("Desea Realizar el proceso de Enviar Ventas ", 1, "Aviso") <> 1 Then Exit Sub

    buf = "delete from dpedidov where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and (acu='I')  "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"

    cn1.Execute (buf)
    buf = "select * from dpedidov where "
    buf = buf & " fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    buf = buf & " and local='" & extra_loquesea(xlocalx) & "'"
    buf = buf & " AND ( acu='I') "

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    total = "" & mytablex.RecordCount
    procesados = ""
    sdx = 0
    mytabley.Open "select * from dpedidov where 1=2 ", cn1, adOpenStatic, adLockOptimistic
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

