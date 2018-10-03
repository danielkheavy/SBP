VERSION 5.00
Begin VB.Form trepoctc 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes Ctas "
   ClientHeight    =   6345
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox xtipo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox tituloreporte 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   30
      TabIndex        =   25
      Top             =   5520
      Width           =   3855
   End
   Begin VB.TextBox vendedor 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   11
      TabIndex        =   11
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox codigo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   11
      TabIndex        =   10
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox numero 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   11
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox tipo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox serie 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox fechai 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox fechaf 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.ComboBox moneda 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3240
      Width           =   1575
   End
   Begin VB.ComboBox tipofecha 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4080
      Width           =   1575
   End
   Begin VB.ComboBox tiposaldo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox nombre 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.ComboBox local1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Credito"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   29
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label archivoreporte 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   45
      TabIndex        =   27
      Top             =   5955
      Width           =   135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Titulo reporte"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   24
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label21 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   22
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label17 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label acu 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo fecha"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Saldo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   2175
   End
   Begin VB.Menu fk3883 
      Caption         =   "&Ejecuta"
   End
   Begin VB.Menu flo44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "trepoctc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub fk3883_Click()

    Dim buf As String

    Dim cad As String

    buf = ""
    buf = "{cuentac.fecha} In Date (" & Format(fechai, "yyyy,mm,dd") & ")" & " To Date (" & Format$(fechaf, "yyyy,mm,dd") & ")" & xbuf

    If tipofecha = "EMISION" Then
        buf = "{cuentac.fecha} In Date (" & Format(fechai, "yyyy,mm,dd") & ")" & " To Date (" & Format$(fechaf, "yyyy,mm,dd") & ")" & xbuf

        '    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        '    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    End If

    If tipofecha = "VENCIMIENTO" Then
        buf = "{cuentac.fechaV} In Date (" & Format(fechai, "yyyy,mm,dd") & ")" & " To Date (" & Format$(fechaf, "yyyy,mm,dd") & ")" & xbuf

        'buf = buf & "  fechav>='" & Format(fechai, "YYYYMMDD") & "'"
        'buf = buf & " and fechav<='" & Format(fechaf, "YYYYMMDD") & "' "
    End If

    If Len(codigo) > 0 Then
        buf = buf & "  and left({cuentac.codigo}," & Len(codigo) & ")= " & "'" & Trim("" & codigo) & "'"

    End If

    If Len(nombre) > 0 Then
        buf = buf & "  and left({clientes.nombre}," & Len(nombre) & ")= " & "'" & Trim("" & nombre) & "'"

    End If

    If Len(Trim(tipo)) > 0 Then
        buf = buf & "  and left({cuentac.tipo}," & Len(Trim(extra_loquesea(tipo))) & ")= " & "'" & Trim(extra_loquesea(tipo)) & "'"

    End If

    If Len(serie) > 0 Then
        buf = buf & "  and left({cuentac.serie}," & Len(serie) & ")= " & "'" & Trim("" & serie) & "'"

    End If

    If Len(Numero) > 0 Then
        buf = buf & "  and left({cuentac.numero}," & Len(Numero) & ")= " & "'" & Trim("" & Numero) & "'"

    End If

    If Len(Trim(estado)) > 0 Then
        If estado = "VENDIDO" Then
            buf = buf & "  and {cuentac.estado}='2'"

        End If

        If estado = "ANULADO" Then
            buf = buf & "  and {cuentac.estado}='1'"

        End If

        If estado = "SINGRABAR" Then
            buf = buf & "  and {cuentac.estado}='0'"

        End If

    End If

    If Len(Trim(moneda)) > 0 Then
        buf = buf & "  and left({cuentac.moneda}," & Len(moneda) & ")= " & "'" & Trim("" & moneda) & "'"

    End If

    If Len(Trim(local1)) > 0 Then
        buf = buf & "  and left({cuentac.local}," & Len(Trim(local1)) & ")= " & "'" & Trim("" & local1) & "'"

    End If

    'If Len(Trim(cajero)) > 0 Then
    '   buf = buf & "  and left({cuentac.usuario}," & Len(Trim(extra_loquesea(cajero))) & ")= " & "'" & Trim("" & extra_loquesea(cajero)) & "'"
    'End If
    'If Len(Trim(caja)) > 0 Then
    '   buf = buf & "  and left({cuentac.caja}," & Len(Trim(extra_loquesea(caja))) & ")= " & "'" & Trim("" & extra_loquesea(caja)) & "'"
    'End If
    'If Len(Trim(turno)) > 0 Then
    '   buf = buf & "  and left({cuentac.turno}," & Len(Trim(extra_loquesea(turno))) & ")= " & "'" & Trim("" & extra_loquesea(turno)) & "'"
    'End If
    If acu = "V" Then

        'buf = buf & "  and ({cuentac.acu}='A' or {cuentac.acu}='B' or {cuentac.acu}='C'  or {cuentac.acu}='D' or {cuentac.acu}='G') "
    End If

    If acu = "C" Then

        'buf = buf & "  and ({cuentac.acu}='J' or {cuentac.acu}='K' or {cuentac.acu}='L'  or {cuentac.acu}='M' or {cuentac.acu}='P') "
    End If

    If tiposaldo = "PENDIENTE" Then
        buf = buf & "  and {cuentac.saldo}>0"

    End If

    If tiposaldo = "CANCELADO" Then
        buf = buf & "  and {cuentac.saldo}=0"

    End If

    If xtipo = "CREDITO" Then
        buf = buf & "  and {cuentac.grupo}='C'"

    End If

    If xtipo = "ANTICIPO DINERO" Then
        buf = buf & "  and {cuentac.grupo}='A'"

    End If

    If xtipo = "DEPOSITO BANCO" Then
        buf = buf & "  and {cuentac.grupo}='D'"

    End If

    If xtipo = "ORDEN TRABAJO" Then
        buf = buf & "  and {cuentac.grupo}='O'"

    End If

    'cad = "SELECT * from factura  where   " & buf & " order by idfactura"
    'If txreporte.State = 1 Then
    '   txreporte.Close
    '   Set txreporte = Nothing
    'End If
    'txreporte.Open cad, cn, adOpenStatic, adLockOptimistic
   
    'CrystalReport1.SelectionFormula = "{Ofertas.OferFechaPropuesta} in Date(" & Agno1 & "," & Mes1 & "," & Dia1 & ") to Date(" & Agno2 & "," & Mes2 & "," & Dia2 & ")"
    'Form1.CR1. SelectionFormula ="{Ofertas.FechaEmision} = #" Fecha "# "
    '"{Tabla.CampoFecha} in Date(Año,Mes,Dia) to Date(Año,Mes2,Dia2)"
   
    'tcrystal.archivoreporte = globaldir & "\reportes\registroventa.rpt"
    tcrystal.archivoreporte = "" & archivoreporte
    'tcrystal.condicion = "{cuentac.tipo}=" & "'" & Trim("" & tipo) & "'"
    'xbuf = "  and {cuentac.tipo}=" & "'" & Trim("" & tipo) & "'"
    'tcrystal.condicion = "{cuentac.fecha} In Date (" & Format(fechai, "yyyy,mm,dd") & ")" & " To Date (" & Format$(fechaf, "yyyy,mm,dd") & ")" & xbuf
    'MsgBox buf
    tcrystal.condicion = buf

    If Len(Trim(tituloreporte)) = 0 Then
        tituloreporte = "Fecha Inicio " & fechai & " Fecha Final: " & fechaf

    End If

    tcrystal.xtitulo = tituloreporte
    tcrystal.Show 1

End Sub

Private Sub flo44_Click()
    trepoctc.Hide
    Unload trepoctc

End Sub

Private Sub Form_Activate()

    Dim mytablex As New ADODB.Recordset

    xtipo.Clear
    xtipo.AddItem "CREDITO"
    xtipo.AddItem "ANTICIPO DINERO"
    xtipo.AddItem "DEPOSITO BANCO"
    xtipo.AddItem "ORDEN TRABAJO"
    xtipo.AddItem ""
    xtipo.ListIndex = 0

    local1.Clear
    local1.AddItem ""

    tipo.Clear
    tipo.AddItem ""

    mytablex.Open "SELECT * FROM tipo", cn, adOpenKeyset, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        If acu = "V" Then
            If "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "D" Or "" & mytablex.Fields("tipodoc") = "G" Then
                tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        If acu = "C" Then
            If "" & mytablex.Fields("tipodoc") = "J" Or "" & mytablex.Fields("tipodoc") = "K" Or "" & mytablex.Fields("tipodoc") = "L" Or "" & mytablex.Fields("tipodoc") = "M" Or "" & mytablex.Fields("tipodoc") = "P" Then
                tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    tipo.ListIndex = 0
    mytablex.Open "SELECT * FROM tlocal", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        local1.AddItem "" & mytablex.Fields("codigo")
        mytablex.MoveNext
    Loop
    mytablex.Close
    local1.ListIndex = 0

    If local1.ListCount = 2 Then
        local1.ListIndex = 1

    End If

End Sub

Private Sub Form_Load()

    tipofecha.AddItem "EMISION"
    tipofecha.AddItem "VENCIMIENTO"
    tipofecha.ListIndex = 0
    tiposaldo.AddItem "PENDIENTE"
    tiposaldo.AddItem "CANCELADO"
    tiposaldo.AddItem ""
    tiposaldo.ListIndex = 0
    moneda.AddItem ""
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0

    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")

End Sub
