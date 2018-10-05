VERSION 5.00
Begin VB.Form treporte 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox moneda 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   3120
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
      TabIndex        =   15
      Top             =   4560
      Width           =   3855
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
      TabIndex        =   14
      Top             =   2520
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
      TabIndex        =   13
      Top             =   2160
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
      TabIndex        =   12
      Top             =   360
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
      TabIndex        =   11
      Top             =   0
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
      TabIndex        =   10
      Top             =   720
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
      TabIndex        =   9
      Top             =   1200
      Width           =   1575
   End
   Begin VB.ComboBox bodega 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3840
      Width           =   1575
   End
   Begin VB.ComboBox estado 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   360
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
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox cajero 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.ComboBox caja 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.ComboBox turno 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.ComboBox vendedor 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.ComboBox local1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.ComboBox SERVICIO 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4200
      Width           =   2055
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
      Left            =   2160
      TabIndex        =   35
      Top             =   4920
      Width           =   135
   End
   Begin VB.Label acu 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   34
      Top             =   2400
      Width           =   105
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
      TabIndex        =   33
      Top             =   4560
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
      TabIndex        =   32
      Top             =   2520
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
      TabIndex        =   31
      Top             =   2160
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
      TabIndex        =   30
      Top             =   0
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
      TabIndex        =   29
      Top             =   360
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
      TabIndex        =   28
      Top             =   720
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
      TabIndex        =   27
      Top             =   1200
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
      TabIndex        =   26
      Top             =   3120
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
      TabIndex        =   25
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bodega"
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
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado"
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
      TabIndex        =   23
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label15 
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
      TabIndex        =   22
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label19 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cajero"
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
      TabIndex        =   21
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caja"
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
      TabIndex        =   20
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label22 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Turno"
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
      TabIndex        =   19
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label26 
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
      Left            =   3840
      TabIndex        =   18
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label27 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Servicio"
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
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Menu dkui44 
      Caption         =   "Ejecuta"
   End
   Begin VB.Menu flo44 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "treporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dkui44_Click()

    Dim buf As String

    Dim cad As String

    buf = ""
    buf = "{Factura.fecha} In Date (" & Format(fechai, "yyyy,mm,dd") & ")" & " To Date (" & Format$(fechaf, "yyyy,mm,dd") & ")" & xbuf

    If Len(codigo) > 0 Then
        buf = buf & "  and left({factura.codigo}," & Len(codigo) & ")= " & "'" & Trim("" & codigo) & "'"

    End If

    If Len(nombre) > 0 Then
        buf = buf & "  and left({factura.nombre}," & Len(nombre) & ")= " & "'" & Trim("" & nombre) & "'"

    End If

    If Len(Trim(tipo)) > 0 Then
        buf = buf & "  and left({factura.tipo}," & Len(Trim(extra_loquesea(tipo))) & ")= " & "'" & Trim(extra_loquesea(tipo)) & "'"

    End If

    If Len(serie) > 0 Then
        buf = buf & "  and left({factura.serie}," & Len(serie) & ")= " & "'" & Trim("" & serie) & "'"

    End If

    If Len(Numero) > 0 Then
        buf = buf & "  and left({factura.numero}," & Len(Numero) & ")= " & "'" & Trim("" & Numero) & "'"

    End If

    If Len(Trim(estado)) > 0 Then
        If estado = "VENDIDO" Then
            buf = buf & "  and {factura.estado}='2'"

        End If

        If estado = "ANULADO" Then
            buf = buf & "  and {factura.estado}='1'"

        End If

        If estado = "SINGRABAR" Then
            buf = buf & "  and {factura.estado}='0'"

        End If

    End If

    If Len(Trim(moneda)) > 0 Then
        buf = buf & "  and left({factura.moneda}," & Len(moneda) & ")= " & "'" & Trim("" & moneda) & "'"

    End If

    If Len(Trim(bodega)) > 0 Then
        buf = buf & "  and left({factura.bodega}," & Len(Trim(extra_loquesea(bodega))) & ")= " & "'" & Trim("" & extra_loquesea(bodega)) & "'"

    End If

    If Len(Trim(servicio)) > 0 Then
        'If servicio = "Autoservicio" Then
        buf = buf & "  and {factura.servicio}='" & extra_loquesea(servicio) & "'"

        'End If
        'If estado = "Comanda" Then
        'buf = buf & "  and {factura.servicio}='C'"
        'End If
        'If estado = "Deliveri" Then
        'buf = buf & "  and {factura.servicio}='D'"
        'End If
    End If

    If Len(Trim(local1)) > 0 Then
        buf = buf & "  and left({factura.local}," & Len(Trim(extra_loquesea(local1))) & ")=" & "'" & Trim("" & extra_loquesea(local1)) & "'"

    End If

    If Len(Trim(cajero)) > 0 Then
        buf = buf & "  and left({factura.usuario}," & Len(Trim(extra_loquesea(cajero))) & ")=" & "'" & Trim("" & extra_loquesea(cajero)) & "'"

    End If

    If Len(Trim(caja)) > 0 Then
        buf = buf & "  and left({factura.caja}," & Len(Trim(extra_loquesea(caja))) & ")=" & "'" & Trim("" & extra_loquesea(caja)) & "'"

    End If

    If Len(Trim(turno)) > 0 Then
        buf = buf & "  and left({factura.turno}," & Len(Trim(extra_loquesea(turno))) & ")=" & "'" & Trim("" & extra_loquesea(turno)) & "'"

    End If

    If acu = "V" Then
        buf = buf & " and ({factura.acu}='A'  or {factura.acu}='B' or {factura.acu}='C' or {factura.acu}='D' or {factura.acu}='G') "

    End If

    If acu = "C" Then
        buf = buf & "  and ({factura.acu}='J' or {factura.acu}='K' or {factura.acu}='L'  or {factura.acu}='M' or {factura.acu}='P') "

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
    'tcrystal.archivoreporte = "" & archivoreporte
    'tcrystal.condicion = "{factura.tipo}=" & "'" & Trim("" & tipo) & "'"
    'xbuf = "  and {factura.tipo}=" & "'" & Trim("" & tipo) & "'"
    'tcrystal.condicion = "{Factura.fecha} In Date (" & Format(fechai, "yyyy,mm,dd") & ")" & " To Date (" & Format$(fechaf, "yyyy,mm,dd") & ")" & xbuf
    'MsgBox buf
    tcrystal.condicion = buf

    If Len(Trim(tituloreporte)) = 0 Then
        tituloreporte = "Fecha Inicio " & fechai & " Fecha Final: " & fechaf

    End If

    tcrystal.xtitulo = tituloreporte
    tcrystal.Show 1

End Sub

Private Sub flo44_Click()
    treporte.Hide
    Unload treporte

End Sub

Private Sub Form_Activate()

    Dim mytablex As New ADODB.Recordset

    tipo.Clear
    tipo.AddItem ""
    mytablex.Open "SELECT * FROM tipo ", cn, adOpenKeyset, adLockOptimistic
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

End Sub

Private Sub Form_Load()

    Dim mytablex As New ADODB.Recordset

    servicio.Clear
    servicio.AddItem ""
    mytablex.Open "SELECT * FROM servicio ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        servicio.AddItem "" & mytablex.Fields("servicio") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    servicio.ListIndex = 0
    mytablex.Close

    fechaf = Format(Now, "dd/mm/yyyy")
    fechai = "01" & "/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")

    estado.AddItem ""
    estado.AddItem "VENDIDO"
    estado.AddItem "ANUALDO"
    estado.AddItem "SINGRABAR"
    estado.ListIndex = 1

    bodega.Clear
    bodega.AddItem ""
    mytablex.Open "SELECT * FROM bodega", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        bodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    bodega.ListIndex = 0

    vendedor.Clear
    vendedor.AddItem ""

    cajero.Clear
    cajero.AddItem ""
    mytablex.Open "SELECT * FROM vendedor", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    cajero.ListIndex = 0
    vendedor.ListIndex = 0

    caja.Clear
    caja.AddItem ""
    mytablex.Open "SELECT * FROM parameca", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    caja.ListIndex = 0

    turno.Clear
    turno.AddItem ""
    mytablex.Open "SELECT * FROM turno", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        turno.AddItem "" & mytablex.Fields("turno") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    turno.ListIndex = 0

    local1.Clear
    local1.AddItem ""
    mytablex.Open "SELECT * FROM tlocal", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        local1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("Nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    local1.ListIndex = 0

    If local1.ListCount = 2 Then
        local1.ListIndex = 1

    End If

    moneda.Clear
    moneda.AddItem ""
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0

End Sub
