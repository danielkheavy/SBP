VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form tliqdcom 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Comisiones"
   ClientHeight    =   8715
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   13725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   13725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   13665
      TabIndex        =   19
      Top             =   0
      Width           =   13725
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
         TabIndex        =   40
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Seleccionar Ventas"
         Height          =   735
         Left            =   11040
         TabIndex        =   29
         Top             =   120
         Width           =   1815
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   23
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   22
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox producto 
         Height          =   375
         Left            =   1200
         MaxLength       =   15
         TabIndex        =   21
         Text            =   "*"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   4200
         MaxLength       =   11
         TabIndex        =   20
         Text            =   "*"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         Height          =   375
         Left            =   6120
         TabIndex        =   39
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Usuario"
         Height          =   375
         Left            =   6120
         TabIndex        =   38
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   375
         Left            =   8520
         TabIndex        =   37
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   375
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Final"
         Height          =   375
         Left            =   0
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
         Height          =   375
         Left            =   0
         TabIndex        =   26
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         Height          =   375
         Left            =   3000
         TabIndex        =   25
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   375
         Left            =   3000
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   13665
      TabIndex        =   15
      Top             =   8220
      Width           =   13725
      Begin VB.CommandButton Command3 
         Caption         =   "Procesar OT"
         Height          =   375
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Producto"
         Height          =   375
         Left            =   9120
         TabIndex        =   18
         Top             =   0
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "TablaVentas"
         Height          =   375
         Left            =   10440
         TabIndex        =   17
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Procesar Liquidacion"
         Height          =   375
         Left            =   11880
         TabIndex        =   16
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   13560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "tliqdcom.frx":0000
      Height          =   5775
      Left            =   120
      OleObjectBlob   =   "tliqdcom.frx":0014
      TabIndex        =   0
      Top             =   1320
      Width           =   13455
   End
   Begin VB.Label Label16 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label txnormal 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1080
      TabIndex        =   32
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Ot"
      Height          =   375
      Left            =   120
      TabIndex        =   31
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label txot 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1080
      TabIndex        =   30
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label ganancia 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   12360
      TabIndex        =   14
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ganancia"
      Height          =   375
      Left            =   11400
      TabIndex        =   13
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label costo 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10200
      TabIndex        =   12
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Costo"
      Height          =   375
      Left            =   9240
      TabIndex        =   11
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label vtaneta 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   10200
      TabIndex        =   10
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VtaNeta"
      Height          =   375
      Left            =   9240
      TabIndex        =   9
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label comision 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comision"
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label subtotal 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8040
      TabIndex        =   6
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Subtotal"
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label impuesto 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Impuesto"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label total 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   7200
      Width           =   975
   End
   Begin VB.Menu fclo44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tliqdcom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    If MsgBox("Procesar Liquidacion", 1, "Aviso") <> 1 Then Exit Sub

    sql_procesa

End Sub

Private Sub Command2_Click()
    sql_mes

End Sub

Private Sub fclo44_Click()
    tliqdcom.Hide
    Unload tliqdcom

End Sub

Private Sub Form_Activate()
    sql_mes

End Sub

Sub carga_inicial()

    Dim mytablex As Table

    vendedor.Clear
    cajero.Clear
    cajero.AddItem "%"
    vendedor.AddItem "%"
    Set mytablex = mydbxglo.OpenTable("vendedor")
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
    Set mytablex = mydbxglo.OpenTable("parameca")
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

    buf = "select * from detalle where "
    buf = buf & "   fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"

    If vendedor <> "%" Then
        buf = buf & " and vendedor='" & extra_loquesea(vendedor) & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja='" & extra_loquesea(caja) & "'"

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario='" & extra_loquesea(cajero) & "'"

    End If

    If producto <> "%" Then
        buf = buf & " and producto like '" & producto & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    'If tipo = "BOLETAS" Then
    '   buf = buf & " and (acu='A' or acu='C' ) "
    'End If
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

    Data2.Connect = "foxpro 2.5;"
    Data2.DatabaseName = globaldir
    Data2.RecordSource = buf
    Data2.refresh

End Sub

Function procesa_producto(mytablex As Table)

    Dim sdx As Double

    mytablex.Seek "=", "" & Data2.Recordset.Fields("producto")

    If Not mytablex.NoMatch Then
        If Option1.Value = True Then
            Data2.Recordset.Edit
            sdx = Val("" & Data2.Recordset.Fields("subtotal")) * Val("" & mytablex.Fields("comision")) / 100 'comision
            Data2.Recordset.Fields("comision") = sdx
            Data2.Recordset.Fields("tcosto") = Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & mytablex.Fields("costou"))
            Data2.Recordset.Fields("vtaneta") = Val("" & Data2.Recordset.Fields("subtotal")) - Val("" & Data2.Recordset.Fields("comision"))
            Data2.Recordset.Fields("ganancia") = Val("" & Data2.Recordset.Fields("vtaneta")) - Val("" & Data2.Recordset.Fields("tcosto"))
            Data2.Recordset.Update

        End If

        If Option2.Value = True Then  'tabla de comisiones
            Data2.Recordset.Edit
            sdx = Val("" & Data2.Recordset.Fields("subtotal")) * busca_vendedor() / 100 'comision
            Data2.Recordset.Fields("comision") = sdx
            Data2.Recordset.Fields("tcosto") = Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & mytablex.Fields("costou"))
            Data2.Recordset.Fields("vtaneta") = Val("" & Data2.Recordset.Fields("subtotal")) - Val("" & Data2.Recordset.Fields("comision"))
            Data2.Recordset.Fields("ganancia") = Val("" & Data2.Recordset.Fields("vtaneta")) - Val("" & Data2.Recordset.Fields("tcosto"))
            Data2.Recordset.Update

        End If

    End If

    '------------------------------------- ------------

End Function

Sub sql_procesa()

    Dim mytablex As Table

    Dim mytabley As Table

    Dim found    As Integer

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

    sql_mes
    xnormal = 0
    xot = 0

    Set mytabley = mydbxglo.OpenTable("factura")
    mytabley.Index = "tfactura"

    Set mytablex = mydbxglo.OpenTable("producto")
    mytablex.Index = "producto"
    ir_inicio
    sdx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    sdx4 = 0
    sdx5 = 0
    sdx6 = 0
    Do

        If Data2.Recordset.EOF Then Exit Do
        found = procesa_producto(mytablex)

        If Val("" & Data2.Recordset.Fields("nroprecio")) = 1 Then
            found = procesa_ot(mytabley)

            If found = 1 Then 'normal
                xnormal = xnormal + Val("" & Data2.Recordset.Fields("total"))

            End If

            If found = 2 Then 'ot
                xot = xot + Val("" & Data2.Recordset.Fields("total"))

            End If

        End If

        sdx = sdx + Val("" & Data2.Recordset.Fields("total"))
        sdx1 = sdx1 + Val("" & Data2.Recordset.Fields("Impuesto"))
        sdx2 = sdx2 + Val("" & Data2.Recordset.Fields("subtotal"))
        sdx3 = sdx3 + Val("" & Data2.Recordset.Fields("comision"))
        sdx4 = sdx4 + Val("" & Data2.Recordset.Fields("vtaneta"))
        sdx5 = sdx5 + Val("" & Data2.Recordset.Fields("tcosto"))
        sdx6 = sdx6 + Val("" & Data2.Recordset.Fields("ganancia"))
        vr = DoEvents()
        Data2.Recordset.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
 
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

Function procesa_ot(mytablex As Table)

    Dim sdx As Double

    mytablex.Seek "=", "" & Data2.Recordset.Fields("local"), "" & Data2.Recordset.Fields("tipo"), "" & Data2.Recordset.Fields("serie"), "" & Data2.Recordset.Fields("numero")

    If Not mytablex.NoMatch Then
        If Len("" & mytablex.Fields("tipo1")) > 0 Then
            procesa_ot = 2

        End If

        If Len("" & mytablex.Fields("tipo1")) = 0 Then
            procesa_ot = 1

        End If

    End If

End Function

Sub ir_inicio()

    On Error GoTo cmd1_err

    Data2.Recordset.MoveFirst
    Exit Sub
cmd1_err:
    Exit Sub

End Sub

Function busca_vendedor() As Double

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("vendedor")
    mytablex.Index = "codigo"
    mytablex.Seek "=", "" & Data2.Recordset.Fields("vendedor")

    If Not mytablex.NoMatch Then
        If Val("" & Data2.Recordset.Fields("total")) >= Val("" & mytablex.Fields("ini1")) And Val("" & Data2.Recordset.Fields("total")) <= Val("" & mytablex.Fields("ini1")) Then
            busca_vendedor = Val("" & mytablex.Fields("por1"))
            GoTo am1

        End If

        If Val("" & Data2.Recordset.Fields("total")) >= Val("" & mytablex.Fields("ini2")) And Val("" & Data2.Recordset.Fields("total")) <= Val("" & mytablex.Fields("ini2")) Then
            busca_vendedor = Val("" & mytablex.Fields("por2"))
            GoTo am1

        End If

        If Val("" & Data2.Recordset.Fields("total")) >= Val("" & mytablex.Fields("ini3")) And Val("" & Data2.Recordset.Fields("total")) <= Val("" & mytablex.Fields("ini3")) Then
            busca_vendedor = Val("" & mytablex.Fields("por3"))
            GoTo am1

        End If

        If Val("" & Data2.Recordset.Fields("total")) >= Val("" & mytablex.Fields("ini4")) And Val("" & Data2.Recordset.Fields("total")) <= Val("" & mytablex.Fields("ini4")) Then
            busca_vendedor = Val("" & mytablex.Fields("por4"))
            GoTo am1

        End If

am1:

    End If

    mytablex.Close

End Function
