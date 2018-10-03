VERSION 5.00
Begin VB.Form treinpre 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reimpresion Tickets Auditoria"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox clave 
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
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   5
      PasswordChar    =   "*"
      TabIndex        =   13
      Top             =   360
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   5820
      TabIndex        =   10
      Top             =   330
      Width           =   1575
   End
   Begin VB.ComboBox caja 
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1800
      Width           =   2535
   End
   Begin VB.ComboBox turno 
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox titulo 
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
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   2
      Top             =   2520
      Width           =   3855
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox cajero 
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label fechaf 
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
      Left            =   5820
      TabIndex        =   15
      Top             =   1050
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clave Supervisor"
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
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   2175
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
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha "
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
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label19 
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
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label20 
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
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label22 
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
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Menu flo2323 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "treinpre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    Dim mytablez As Table

    Dim buf      As String

    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    contpag = 0

    If Len(fechai) <> 10 Then Exit Sub
    If Not IsDate(fechai) Then Exit Sub
    fechaf = fechai

    If Len(clave) = 0 Then
        MsgBox "Clave Erronea ", 48, "Aviso"
        clave = ""
        clave.SetFocus
        Exit Sub

    End If

    found = busca_clave("" & clave)

    If found = 0 Then
        MsgBox "Clave Erronea ", 48, "Aviso"
        clave = ""
        clave.SetFocus
        Exit Sub

    End If

    clave = ""
    found = sql_documento(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = borra_nombre("" & FileName)
    cuerpo_programa_documento mytablex
    cerrar_archivo
    mytablex.Close
    '----empieza los cuadres
    Open FileName For Append As #1
    borrar_cuadres
    found = creando_cuadres("" & usuariopos)
    cabecera " HORA " & Format(Now, "hh:mm:ss")
    cuerpo_programa 0
    cerrar_archivo
    
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Function sql_documento(mytablex As ADODB.Recordset)

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    buf = "select * from factura where "
    buf = buf & "  fecha=" & "DateValue('" & fechai & "'" & ")"

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & extra_loquesea("" & caja) & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno='" & extra_loquesea("" & turno) & "'"

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario='" & extra_loquesea("" & cajero) & "'"

    End If

    buf = buf & " and ( acu='C' or acu='D' )"
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    sql_documento = 1

End Function

Sub cuerpo_programa_documento(mytablex As Snapshot)

    Dim sdx As Double

    On Error GoTo cmd34_err

    sdx = 0
    Do

        If mytablex.EOF Then Exit Sub
        'factura_formatox , "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo"), "" & mytablex.Fields("serie"), "" & mytablex.Fields("numero"), "", 0
  
        'C:\WINDOWS\hinhem.scr
  
        mytablex.MoveNext
    Loop
    MsgBox sdx
    Exit Sub
cmd34_err:
    MsgBox "Aviso en cuerpo_programa " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub flo2323_Click()
    treinpre.Hide
    Unload treinpre

End Sub

Private Sub Form_Load()
    fechai = Format(Now, "dd/mm/yyyy")

    Dim mytablex As New ADODB.Recordset

    caja.Clear
    mytablex.Open "select * from parameca ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("terminal") = "C" Then
            caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("descripcio")

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    caja.ListIndex = 0

    turno.Clear
    mytablex.Open "select * from turno ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        turno.AddItem "" & mytablex.Fields("turno") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    turno.ListIndex = 0

    cajero.Clear
    mytablex.Open "select * from vendedor ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    cajero.ListIndex = 0
    local1.Clear
    mytablex.Open "select * from tlocal ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        local1.AddItem "" & mytablex.Fields("codigo") & "|" & "" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    local1.ListIndex = 0

    If local1.ListCount = 2 Then
        local1.ListIndex = 1

    End If
 
End Sub

Sub factura_formatox(bxlocal As String, _
                     bxtipo As String, _
                     bxserie As String, _
                     bxnumero As String, _
                     ascopia As String, _
                     psw As Integer)

    Dim vacu            As String

    Dim mytablex        As New ADODB.Recordset

    Dim mytabley        As New ADODB.Recordset

    Dim mytablez        As New ADODB.Recordset

    Dim found           As Integer

    Dim nro_lineas      As Integer

    Dim contando        As Integer

    Dim faltante        As Integer

    Dim I               As Integer

    Dim archivo_formato As String

    On Error GoTo cmd4500091_err

    gocabeza = "factura"
    godetalle = "detalle"
    gofpago = "fpagov"
    vacu = ""
    'MsgBox "QU"
    nro_lineas = busca_tipo_lineas(bxtipo)
    'MsgBox ""
    'If nro_lineas <= 0 Then
    '   nro_lineas = 10
    'End If
    'MsgBox ""
    contando = 0
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = borra_nombre("" & FileName)
       
    If psw = 2 Then 'si es de orden
        archivo_formato = "orden"
    Else

        'archivo_formato = busca_archivo_formato(bxtipo)
        If Len(archivo_formato) = 0 Then
            MsgBox "No existe archivo formato ", 48, "Aviso"
            Exit Sub

        End If

    End If

    'cabeza
    'proceso_formatos(archivo_formato , mydbx , mytablex , ubicacioni , ubicacionf , basedatos , indice , tipo , numero , ascopia , contando )
       
    mytablex.Open "SELECT * FROM " & gocabeza & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then 'si existe
        mytablex.Close
        Exit Sub

    End If

    'MsgBox ""
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytablex, "{", "}", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytablex, "{", "}", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
              
    vacu = "" & mytablex.Fields("acu")
    'MsgBox ""
    '
    'detalle
    flag_contando = 0

    If "" & mytablex.Fields("observa") = "CONSUMO" Then
        Open FileName For Append As #1
        found = formateaa("1  POR CONSUMO            " & Format(Val("" & mytablex.Fields("total")), "0.00"), 30, 2, 0)
        'found = formateaa("1    POR CONSUMO            ", 30, 2, 0)
        ' found = formateaa("1    COMBUSTIBLE            ", 30, 2, 0)
        contando = contando + 1
        flag_contando = contando + 1
        Close #1

    End If

    If "" & mytablex.Fields("observa") <> "CONSUMO" Then
        mytabley.Open "SELECT * FROM " & godetalle & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

        If mytabley.RecordCount > 0 Then 'si existe
            Do

                If mytabley.EOF Then Exit Do
                If "" & mytabley.Fields("dua") <> "R" Then
                    flag_contando = contando + 1
                    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                    'found = proceso_formatos(archivo_formato, mytabley, "/", "\", godetalle, "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
                    found = proceso_formatos(archivo_formato, mytabley, "/", "\", godetalle, "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
                    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                    contando = contando + 1

                End If

                mytabley.MoveNext
            Loop

        End If

        mytabley.Close

    End If

    '
    'If bxtipo <> "1" And bxtipo <> "2" And bxtipo <> "5" And bxtipo <> "7" Then
    If nro_lineas > 0 Then
        If contando < nro_lineas Then

            For I = contando To nro_lineas
                Open FileName For Append As #1
                found = formateaa("", 1, 2, 0)
                Close #1
            Next I

        End If

    End If

    '----- SUBTOTAL
       
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytablex, "$", "?", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytablex, "$", "?", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
               
    mytablez.Open "SELECT * FROM " & gofpago & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablez.RecordCount > 0 Then 'si existe
        Do

            If mytablez.EOF Then Exit Do
            'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
            'found = proceso_formatos(archivo_formato, mytablez, "<", ">", gofpago, "fpagov", bxlocal, bxtipo,bxserie, bxnumero, ascopia, 0)
            found = proceso_formatos(archivo_formato, mytablez, "<", ">", gofpago, "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
            'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
            mytablez.MoveNext
        Loop

    End If
       
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytablex, "^", "&", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytablex, "^", "&", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
         
    mytablex.Close
    'mytabley.Close
    mytablez.Close
        
    Exit Sub
cmd4500091_err:
    MsgBox "Aviso Proceso en Compilando " & error$, 24, "Aviso"
    'mytablex.Close
    '
    Exit Sub

End Sub

Function busca_tipo_lineas(buf As String)

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM tipo  where  tipo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        busca_tipo_lineas = Val("" & mytablex.Fields("nrolineas"))

    End If

    mytablex.Close

End Function

Sub factura_formatoxx(mytableh As Snapshot, _
                      bxlocal As String, _
                      bxtipo As String, _
                      bxserie As String, _
                      bxnumero As String, _
                      ascopia As String, _
                      psw As Integer)

    Dim vacu            As String

    Dim mytablex        As Table

    Dim mytabley        As Table

    Dim mytablez        As Table

    Dim found           As Integer

    Dim nro_lineas      As Integer

    Dim contando        As Integer

    Dim faltante        As Integer

    Dim I               As Integer

    Dim archivo_formato As String

    On Error GoTo cmd450009_err

    vacu = ""
    nro_lineas = 10
    'Filename = globaldir & "\temporal\" & gusuario & ".txt"
    'found = borra_nombre("" & Filename)
    archivo_formato = busca_archivo_formato(bxtipo, mytableh)
    'MsgBox archivo_formato
    Set mytablex = mydbxglo.OpenTable("factura")
    mytablex.Index = "TFACTURA"
    Set mytabley = mydbxglo.OpenTable("detalle")
    mytabley.Index = "Tdetalle"
    Set mytablez = mydbxglo.OpenTable("fpagov")
    mytablez.Index = "fpagov"
    mytablex.Seek "=", bxlocal, bxtipo, bxserie, bxnumero

    If Not mytablex.NoMatch Then
        'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
        'found = proceso_formatos(archivo_formato, mytablex, "{", "}", "factura", "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
        found = proceso_formatos(archivo_formato, mytablex, "{", "}", "factura", "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
        'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
        vacu = "" & mytablex.Fields("acu")

    End If

    flag_contando = 0
    mytabley.Seek "=", bxlocal, bxtipo, bxserie, bxnumero

    If Not mytabley.NoMatch Then
        Do

            If mytabley.EOF Then Exit Do
            If "" & mytabley.Fields("local") = bxlocal And "" & mytabley.Fields("tipo") = bxtipo And "" & mytabley.Fields("serie") = bxserie And "" & mytabley.Fields("numero") = bxnumero Then
                flag_contando = contando + 1
                'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                'found = proceso_formatos(archivo_formato, mytabley, "/", "\", "detalle", "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
                found = proceso_formatos(archivo_formato, mytabley, "/", "\", "detalle", "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
                'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                  
                contando = contando + 1
            Else
                Exit Do

            End If

            mytabley.MoveNext
        Loop

    End If

    mytablex.Seek "=", bxlocal, bxtipo, bxserie, bxnumero

    If Not mytablex.NoMatch Then
        'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
        'found = proceso_formatos(archivo_formato, mytablex, "$", "?", "factura", "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
        found = proceso_formatos(archivo_formato, mytablex, "$", "?", "factura", "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)

        'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    End If

    mytablez.Seek "=", bxlocal, bxtipo, bxserie, bxnumero

    If Not mytablez.NoMatch Then
        Do

            If mytablez.EOF Then Exit Do
            If "" & mytablez.Fields("local") = bxlocal And "" & mytablez.Fields("tipo") = bxtipo And "" & mytablez.Fields("serie") = bxserie And "" & mytablez.Fields("numero") = bxnumero Then
                'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                'found = proceso_formatos(archivo_formato, mytablez, "<", ">", "fpagov", "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
                found = proceso_formatos(archivo_formato, mytablez, "<", ">", "fpagov", "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
                'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                           
                Else: Exit Do

            End If

            mytablez.MoveNext
        Loop

    End If

    mytablex.Seek "=", bxlocal, bxtipo, bxserie, bxnumero

    If Not mytablex.NoMatch Then
        'inicio 30/05/2017 pll para la parametrizacion nombre consistencia cliente ticketera
        'found = proceso_formatos(archivo_formato, mytablex, "^", "&", "factura", "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
        found = proceso_formatos(archivo_formato, mytablex, "^", "&", "factura", "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)

        'fin 30/05/2017 pll para la parametrizacion nombre consistencia cliente ticketera
    End If

    mytablex.Close
    mytabley.Close
    mytablez.Close
    Exit Sub
cmd450009_err:
    MsgBox "Aviso Proceso en Compilando " & error$, 24, "Aviso"
    Exit Sub

End Sub

Function busca_archivo_formato(bxtipo As String, mytabley As ADODB.Recordset) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from parameca where caja='" & "" & mytabley.Fields("caja") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        Select Case "" & mytabley.Fields("acu")

            Case "Z" 'si es traslado
                busca_archivo_formato = "" & mytablex.Fields("archivo")

            Case "A"
                busca_archivo_formato = "" & mytablex.Fields("archivobm")

            Case "B"
                busca_archivo_formato = "" & mytablex.Fields("archivofm")

            Case "C"
                busca_archivo_formato = "" & mytablex.Fields("archivotb")

            Case "1"
                busca_archivo_formato = "" & mytablex.Fields("archivoexo")

            Case "D"
                busca_archivo_formato = "" & mytablex.Fields("archivotf")

            Case "G"
                busca_archivo_formato = "" & mytablex.Fields("archivonv")

            Case "H"
                busca_archivo_formato = "" & mytablex.Fields("archivope")

            Case "I"
                busca_archivo_formato = "" & mytablex.Fields("archivoot")

        End Select

    End If

    mytablex.Close

End Function

Function busca_clave(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where clave='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_clave = 1

    End If

    mytablex.Close

End Function

Function creando_cuadres(buf2 As String)

    Dim found      As Integer

    Dim globaldat1 As String

    Dim buf        As String

    On Error GoTo cmd56rre_err

    buf = buf2
    globaldat1 = globaldat & "\"
    copiando globaldat1 & "cuadre01.dbf", globaldat1 & buf & "01.dbf"
    copiando globaldat1 & "cuadre01.cdx", globaldat1 & buf & "01.cdx"
    copiando globaldat1 & "cuadre02.dbf", globaldat1 & buf & "02.dbf"
    copiando globaldat1 & "cuadre02.cdx", globaldat1 & buf & "02.cdx"
    copiando globaldat1 & "cuadre03.dbf", globaldat1 & buf & "03.dbf"
    copiando globaldat1 & "cuadre03.cdx", globaldat1 & buf & "03.cdx"
    copiando globaldat1 & "cuadre04.dbf", globaldat1 & buf & "04.dbf"
    copiando globaldat1 & "cuadre04.cdx", globaldat1 & buf & "04.cdx"
    creando_cuadres = 1
    Exit Function
cmd56rre_err:

    MsgBox "Por favor Llame a servicio tecnico", 24, "Aviso"
    Exit Function

End Function

Sub cabecera(bufd As String)

    Dim buf    As String

    Dim titulo As String

    Dim I      As Integer

    Dim sdx    As Double

    Dim found  As Integer

    sdx = graba_cierres("" & caja)
    titulo = "CIERRE DEL DIA NRO: " & Format(Val(numcuadre), "000000")
    buf = titulo
    I = (36 - Len(titulo)) / 2
    found = formateaa(" ", I, 0, 0)
    found = formateaa(titulo, Len(titulo), 2, 0)

    titulo = bufd
    buf = titulo
    I = (36 - Len(titulo)) / 2
    found = formateaa(" ", I, 0, 0)
    found = formateaa(titulo, Len(titulo), 2, 0)
    '-------
    buf = String(35, "-")
    found = formateaa(buf, 35, 2, 0)
    found = formateaa("CAJERO." & extra_loquesea("" & cajero) & " CAJA." & extra_loquesea("" & caja) & " TNO." & extra_loquesea("" & turno), 35, 2, 0)
    found = formateaa("FECHAI." & fechai & " FECHAF." & fechaf, 35, 2, 0)
    buf = String(35, "-")
    found = formateaa(buf, 35, 2, 0)
   
End Sub

Sub cuerpo_programa(sw As Integer)

    Dim buf   As String

    Dim tsw   As Integer

    Dim found As Integer

    Dim I     As Integer

    Dim sdx   As Double

    Dim sdx1  As Double

    Dim sdx2  As Double

    Dim sdx3  As Double

    Dim vr    As Integer

    On Error GoTo cmd23_err

    sum1 = 0
    sum2 = 0
    sum3 = 0
    suma5 = 0
    suma6 = 0
    borrar_cuadres
    fecha = "Poniendo Cajeros"
    visualiza_cajeros
    'buf = String(35, "-")
    'found = formateaa(buf, 35, 2, 0)
    fecha = "Poniendo Igv"
    sdx = busca_igv()
    buf = "T/CAMBIO :" & Format(sdx, "0.000")
    found = formateaa(buf, Len(buf), 2, 0)
    fecha = "Acumulando..espere"
    buf = "SERVICIOS"
    found = formateaa(buf, Len(buf), 2, 0)
    servicio_realizado
    imprime_servicio
    buf = "DOCUMENTOS VALORADOS"
    found = formateaa(buf, Len(buf), 2, 0)
    sum1 = 0
    sum2 = 0
    sum3 = 0
    sum4 = 0
    imprime_doctos 0
    buf = "RESUMEN DE VENTAS"
    found = formateaa(buf, Len(buf), 2, 0)
    imprime_valorv

    If todos = "S" Then
        buf = "OTROS DOCUMENTOS "
        found = formateaa(buf, Len(buf), 2, 0)
        imprime_doctos 1
        found = formateaa("NETO VENTAS", 14, 0, 0)
        buf = Format(sum1, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(sum2, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)

    End If

    'imprime pedidos
    'TOTAL OTROS
    'MsgBox "x"
    buf = "ORDEN TRABAJO "
    found = formateaa(buf, Len(buf), 2, 0)
    imprime_orden_trabajo
    buf = "INGRESOS/EGRESOS"
    found = formateaa(buf, Len(buf), 2, 0)
    imprime_recibos
    
    'buf = "ORDEN TRABAJO-ABONOS "
    'found = formateaa(buf, Len(buf), 2, 0)
    'imprime_ordenes
    
    '
    sdx = busca_igv()

    If sdx = 0 Then
        sdx = 1

    End If
    
    sdx1 = (sum1 + sum3 + suma5) + (sum2 + sum4 + suma6) * sdx
    sdx1 = Format(sdx1, "0.00")
    sdx2 = sdx1 / sdx
    sdx2 = Format(sdx2, "0.00")
    '---------------------------------------------------
    buf = String(35, "-")
    found = formateaa(buf, 35, 2, 0)
    buf = "TOT.EFE.CAJA "
    found = formateaa(buf, 14, 0, 0)
    found = formateaa("", 1, 0, 0)

    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    If busca_config(1) = "N" Then
        sdx2 = 0

    End If

    buf = Format(sdx2, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
    '---------------------------------------------------
    fecha = "POR FAVOR ESPERE ...."
    buf = String(35, "-")
    found = formateaa(buf, 35, 2, 0)
    buf = "FORMA DE PAGO/INGRESOS"
    found = formateaa(buf, Len(buf), 2, 0)
    forma_pago
    imprime_fpago
    tsw = 8

    If sw = 1 Then
        tsw = 2

    End If

    For I = 1 To tsw
        found = formateaa("", 1, 2, 0)
    Next I

    fecha = "TERMINANDO PROCESO ...."
    Exit Sub
cmd23_err:
    MsgBox "Error en cuerpo programa.." & error$, 48, "Aviso"
    Exit Sub
 
End Sub

Sub visualiza_cajeros()

    Dim buf   As String

    Dim buf1  As String

    Dim buf2  As String

    Dim buf3  As String

    Dim found As Integer

    On Error GoTo cmd1_err:

    buf2 = ""
    buf3 = ""
    buf1 = busca_inicio("1", buf2, buf3)
    buf = "TB-I:" & buf2
    found = formateaa(buf, 16, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "TB-F:" & buf3
    found = formateaa(buf, 16, 0, 0)
    found = formateaa("", 1, 2, 0)
    buf2 = ""
    buf3 = ""
    buf1 = busca_inicio("2", buf2, buf3)
    buf = "TF-I:" & buf2
    found = formateaa(buf, 16, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "TF-F:" & buf3
    found = formateaa(buf, 16, 0, 0)
    found = formateaa("", 1, 2, 0)
    buf2 = ""
    buf3 = ""
    buf1 = busca_inicio("8", buf2, buf3)
    buf = "NCTI:" & buf2
    found = formateaa(buf, 16, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "NCTF:" & buf3
    found = formateaa(buf, 16, 0, 0)
    found = formateaa("", 1, 2, 0)

    buf2 = ""
    buf3 = ""
    buf1 = busca_inicio("3", buf2, buf3)
    buf = "BM-I:" & buf2
       
    found = formateaa(buf, 16, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "BM-F:" & buf3
    found = formateaa(buf, 16, 0, 0)
    found = formateaa("", 1, 2, 0)
    buf2 = ""
    buf3 = ""
    buf1 = busca_inicio("4", buf2, buf3)
    buf = "FM-I:" & buf2
    found = formateaa(buf, 16, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "FM-F:" & buf3
    found = formateaa(buf, 16, 0, 0)
    found = formateaa("", 1, 2, 0)
    buf2 = ""
    buf3 = ""
    buf1 = busca_inicio("9", buf2, buf3)
    buf = "NC-I:" & buf2
    found = formateaa(buf, 16, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "NC-F:" & buf3
    found = formateaa(buf, 16, 0, 0)
    found = formateaa("", 1, 2, 0)
    Exit Sub
cmd1_err:
    Exit Sub

End Sub

Function busca_inicio(buf2 As String, buf3 As String, buf4 As String) As String

    Dim mysnapx As Snapshot

    Dim buf     As String

    '-------------------------
    buf = "select * from factura where "
    buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    buf = buf & " and usuario like '" & extra_loquesea("" & cajero) & "%'"
    buf = buf & " and caja like '" & extra_loquesea("" & caja) & "%'"
    buf = buf & " and turno like '" & extra_loquesea("" & turno) & "%'"
    buf = buf & " and tipo ='" & buf2 & "'"
    buf = buf & " order by fecha,str(numero)"
    Set mysnapx = mydbxglo.CreateSnapshot(buf)

    If mysnapx.EOF = True And mysnapx.BOF = True Then
        buf3 = ""
        buf4 = ""
    Else
        buf3 = "" & mysnapx.Fields("numero")
        mysnapx.MoveLast
        buf4 = "" & mysnapx.Fields("numero")

    End If

    mysnapx.Close

End Function

Function busca_igv() As Double

    Dim mytablex As Table

    On Error GoTo cmd666_err

    busca_igv = 0
    Set mytablex = mydbxglo.OpenTable("parame")
    mytablex.Index = "codigo"
    mytablex.Seek "=", "01"

    If Not mytablex.NoMatch Then
        busca_igv = Val("" & mytablex.Fields("parivta"))

    End If

    If mytablex.NoMatch Then
        busca_igv = 1

    End If

    mytablex.Close
    Exit Function
cmd666_err:
    MsgBox "Mensaje,Error en moneda " & error$
    mytablex.Close
    Exit Function

End Function

Sub servicio_realizado()

    Dim found As Integer

    Dim vr, buf, buf1 As String

    On Error GoTo cmd56_err

    Dim sdx      As Double

    Dim mytablex As Table

    Dim mytabley As Table

    Dim signos   As Double

    sum1 = 0
    Set mytablex = mydbxglo.OpenTable(usuariopos & "01")   'cuadre 01
    mytablex.Index = "salon"
    Set mytabley = mydbxglo.OpenTable(usuariopos & "02")  'cuadre 02
    mytabley.Index = "tipo"
    buf = "select * from FACTURA where "
    buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    buf = buf & " and usuario like '" & extra_loquesea("" & cajero) & "%'"
    buf = buf & " and caja like '" & extra_loquesea("" & caja) & "%'"
    buf = buf & " and turno like '" & extra_loquesea("" & turno) & "%'"
    buf = buf & " and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='1') "  'E nota credito
    buf = buf & " order by fecha "
    Set mysnap = mydbxglo.CreateSnapshot(buf)
    Do

        If mysnap.EOF Then Exit Do
        signos = 1

        If "" & mysnap.Fields("acu") = "E" Then  'nota de credito
            signos = -1

        End If

        sum1 = sum1 + 1
        fecha = "CABECERAS ..." & Format(sum1, "00000")
        buf = "" & mysnap.Fields("salon")

        If Len(buf) = 0 Then
            buf = "0"

        End If

        buf1 = "" & mysnap.Fields("servicio")

        If buf1 <> "*" And buf1 <> "D" And buf1 <> "C" Then GoTo a1
        fecha = "" & mysnap.Fields("fecha")

        If "" & mysnap.Fields("acu") <> "S" And "" & mysnap.Fields("acu") <> "T" Then  'entrdas /salidas
            'servicios
            mytablex.Seek "=", buf, "" & mysnap.Fields("servicio")

            If Not mytablex.NoMatch Then
                mytablex.Edit
                suma_contador mytablex, signos
                mytablex.Update

            End If

            If mytablex.NoMatch Then
                mytablex.AddNew
                suma_contador mytablex, signos
                mytablex.Fields("local") = "01"
                mytablex.Update

            End If

        End If

        'documentos
        '--------------
        buf1 = "" & mysnap.Fields("acu")
        mytabley.Seek "=", buf1 & "" & mysnap.Fields("tipo")

        If Not mytabley.NoMatch Then
            mytabley.Edit
            suma_contador1 mytabley, signos
            mytabley.Fields("tipo") = "" & mysnap.Fields("acu") & "" & mysnap.Fields("tipo")
            mytabley.Update

        End If

        If mytabley.NoMatch Then
            mytabley.AddNew

            If opcion1 = "5" Then
                mytabley.Fields("cierre") = busca_cierre(extra_loquesea("" & caja))
                mytabley.Fields("cajero") = extra_loquesea("" & cajero)
                mytabley.Fields("caja") = extra_loquesea("" & caja)
                mytabley.Fields("turno") = extra_loquesea("" & turno)
                mytabley.Fields("fecha") = Format(Now, "dd/mm/yyyy")
                mytabley.Fields("hora") = Format(Now, "hh:mm:ss")

            End If

            suma_contador1 mytabley, signos
            mytabley.Fields("tipo") = "" & mysnap.Fields("acu") & "" & mysnap.Fields("tipo")
            mytabley.Fields("local") = "01"
            mytabley.Update

        End If

        '--------------
        mysnap.MoveNext
    Loop
a1:
    mysnap.Close

    sum1 = 0
    mytablex.Index = "servicio"
    '---------- ingresos /egresos----------------------------------
    buf = "select * from RECIBO where  usuario like '" & extra_loquesea("" & cajero) & "%'"

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    buf = buf & " and caja like '" & extra_loquesea("" & caja) & "%'"
    buf = buf & " and turno like '" & extra_loquesea("" & turno) & "%'"
    buf = buf & " and fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"
    buf = buf & " order by fecha"

    Set mysnap = mydbxglo.CreateSnapshot(buf)

    Do 'Until mysnap.EOF

        If mysnap.EOF Then Exit Do
        signos = 1
        sum1 = sum1 + 1
        fecha = "INGRESOS/EGRESOS ..." & Format(sum1, "00000")
        buf1 = "" & mysnap.Fields("servicio")

        If buf1 <> "W" And buf1 <> "V" Then GoTo a32
        mytablex.Seek "=", buf1 & "" & mysnap.Fields("tipo")

        If Not mytablex.NoMatch Then
            mytablex.Edit
            suma_contador mytablex, signos
            mytablex.Fields("servicio") = "" & mysnap.Fields("servicio") & "" & mysnap.Fields("tipo")
            mytablex.Update

        End If

        If mytablex.NoMatch Then
            mytablex.AddNew
            suma_contador mytablex, signos
            mytablex.Fields("servicio") = "" & mysnap.Fields("servicio") & "" & mysnap.Fields("tipo")
            mytablex.Fields("local") = "01"
            mytablex.Update

        End If

        mysnap.MoveNext
    Loop
a32:
    mysnap.Close
    '--------------------------------------------------------------
    mytablex.Close
    mytabley.Close
    Exit Sub
cmd56_err:

    If Err <> 3260 Then
        MsgBox "1***Mensaje,Error en servicio realizado " & buf & " " & error$, 24, "AVISO"
        mysnap.Close
        '--------------------------------------------------------------
        mytablex.Close
        mytabley.Close
        Exit Sub

    End If

    Resume

End Sub

Sub suma_contador(mytablex As Table, signos As Double)

    Dim sdx As Double

    Dim buf As String

    On Error GoTo cmd57_err

    If Val("" & mysnap.Fields("tipo")) = 5 Then
        If todos <> "S" Then Exit Sub

    End If

    buf = "" & mysnap.Fields("salon")

    If Len(buf) = 0 Then
        buf = "0"

    End If

    mytablex.Fields("servicio") = "" & mysnap.Fields("servicio")
    mytablex.Fields("salon") = buf

    If Val("" & mysnap.Fields("estado")) = 2 Then
        sdx = Val("" & mytablex.Fields("nro")) + 1
        mytablex.Fields("nro") = sdx

        If "" & mysnap.Fields("moneda") = "S" Then
            sdx = Val("" & mytablex.Fields("valors")) + signos * Val("" & mysnap.Fields("total"))
            mytablex.Fields("valors") = sdx

        End If

        If "" & mysnap.Fields("moneda") = "D" Then
            sdx = Val("" & mytablex.Fields("valord")) + signos * Val("" & mysnap.Fields("total"))
            mytablex.Fields("valord") = sdx

        End If

    End If

    If Val("" & mysnap.Fields("estado")) = 1 Then
        sdx = Val("" & mytablex.Fields("nroa")) + 1
        mytablex.Fields("nroa") = sdx

        If "" & mysnap.Fields("moneda") = "S" Then
            sdx = Val("" & mytablex.Fields("valorsa")) + signos * Val("" & mysnap.Fields("total"))
            mytablex.Fields("valorsa") = sdx

        End If

        If "" & mysnap.Fields("moneda") = "D" Then
            sdx = Val("" & mytablex.Fields("valorda")) + signos * Val("" & mysnap.Fields("total"))
            mytablex.Fields("valorda") = sdx

        End If

    End If

    Exit Sub
cmd57_err:
    MsgBox "Error en suma contador " & error$, 24, "AVISO"
    Exit Sub

End Sub

Sub suma_contador1(mytablex As Table, signos As Double)

    Dim sdx As Double

    On Error GoTo cmd54311_err

    If Val("" & mysnap.Fields("tipo")) = 5 Then
        If todos <> "S" Then Exit Sub

    End If

    mytablex.Fields("tipo") = "" & mysnap.Fields("tipo")

    If Val("" & mysnap.Fields("estado")) = 2 Then
        sdx = Val("" & mytablex.Fields("nro")) + 1
        mytablex.Fields("nro") = sdx

        If "" & mysnap.Fields("moneda") = "S" Then
            sdx = Val("" & mytablex.Fields("valors")) + signos * Val("" & mysnap.Fields("total"))
            mytablex.Fields("valors") = sdx
            sdx = Val("" & mytablex.Fields("valorvs")) + signos * Val("" & mysnap.Fields("subtotal"))
            mytablex.Fields("valorvs") = sdx
            sdx = Val("" & mytablex.Fields("igvs")) + signos * Val("" & mysnap.Fields("impuesto"))
            mytablex.Fields("igvs") = sdx
            sdx = Val("" & mytablex.Fields("exos")) + signos * Val("" & mysnap.Fields("gravado"))
            mytablex.Fields("exos") = sdx
            sdx = Val("" & mytablex.Fields("tax1s")) + signos * Val("" & mysnap.Fields("tisc"))
            'sdx = 0
            mytablex.Fields("tax1s") = sdx
            sdx = Val("" & mytablex.Fields("dsctos")) + signos * Val("" & mysnap.Fields("descuento"))
            'sdx = 0
            mytablex.Fields("dsctos") = sdx
            'sdx = Val("" & mytablex.Fields("retes")) + signos * Val("" & mysnap.Fields("tretencion"))
            sdx = 0
            mytablex.Fields("retes") = sdx
            sdx = Val("" & mytablex.Fields("nodsctos")) + signos * Val("" & mysnap.Fields("tivap"))
            'sdx = 0
            mytablex.Fields("nodsctos") = sdx
            sdx = Val("" & mytablex.Fields("brutos")) + signos * Val("" & mysnap.Fields("neto"))
            mytablex.Fields("brutos") = sdx
           
            If "" & mysnap.Fields("dflag") = "" Then
                sdx = Val("" & mytablex.Fields("cdetras")) + signos * Val("" & mysnap.Fields("tdetra"))
                mytablex.Fields("cdetraS") = sdx

            End If

            If "" & mysnap.Fields("dflag") = "1" Then
                sdx = Val("" & mytablex.Fields("ndetras")) + signos * Val("" & mysnap.Fields("tdetra"))
                mytablex.Fields("ndetraS") = sdx

            End If

        End If

        If "" & mysnap.Fields("moneda") = "D" Then
            sdx = Val("" & mytablex.Fields("valord")) + signos * Val("" & mysnap.Fields("total"))
            mytablex.Fields("valord") = sdx
            sdx = Val("" & mytablex.Fields("valorvd")) + signos * Val("" & mysnap.Fields("subtotal"))
            mytablex.Fields("valorvd") = sdx
            sdx = Val("" & mytablex.Fields("igvd")) + signos * Val("" & mysnap.Fields("impuesto"))
            mytablex.Fields("igvd") = sdx
            sdx = Val("" & mytablex.Fields("exod")) + signos * Val("" & mysnap.Fields("gravado"))
            mytablex.Fields("exod") = sdx
            sdx = Val("" & mytablex.Fields("tax1d")) + signos * Val("" & mysnap.Fields("tisc"))
            'sdx = 0
            mytablex.Fields("tax1d") = sdx
            sdx = Val("" & mytablex.Fields("dsctod")) + signos * Val("" & mysnap.Fields("descuento"))
            'sdx = 0
            mytablex.Fields("dsctod") = sdx
            'sdx = Val("" & mytablex.Fields("reted")) + signos * Val("" & mysnap.Fields("tretencion"))
            sdx = 0
            mytablex.Fields("reted") = sdx
            sdx = Val("" & mytablex.Fields("nodsctod")) + signos * Val("" & mysnap.Fields("tivap"))
            sdx = 0
            mytablex.Fields("nodsctod") = sdx
            sdx = Val("" & mytablex.Fields("brutod")) + signos * Val("" & mysnap.Fields("neto"))
            mytablex.Fields("brutod") = sdx

            If "" & mysnap.Fields("dflag") = "" Then
                sdx = Val("" & mytablex.Fields("cdetrad")) + signos * Val("" & mysnap.Fields("tdetra"))
                mytablex.Fields("cdetrad") = sdx

            End If

            If "" & mysnap.Fields("dflag") = "1" Then
                sdx = Val("" & mytablex.Fields("ndetrad")) + signos * Val("" & mysnap.Fields("tdetra"))
                mytablex.Fields("ndetrad") = sdx

            End If

        End If

    End If

    If Val("" & mysnap.Fields("estado")) = 1 Then
        sdx = Val("" & mytablex.Fields("nroa")) + 1
        mytablex.Fields("nroa") = sdx

        If "" & mysnap.Fields("moneda") = "S" Then
            sdx = Val("" & mytablex.Fields("valorsa")) + signos * Val("" & mysnap.Fields("total"))
            mytablex.Fields("valorsa") = sdx

        End If

        If "" & mysnap.Fields("moneda") = "D" Then
            sdx = Val("" & mytablex.Fields("valorda")) + signos * Val("" & mysnap.Fields("total"))
            mytablex.Fields("valorda") = sdx

        End If

    End If

    Exit Sub
cmd54311_err:
    MsgBox "Error en suma_contador 1" + error, 48, "Aviso"
    Exit Sub

End Sub

Sub imprime_servicio()

    Dim buf   As String

    Dim found As Integer

    Dim buf1  As String

    On Error GoTo cmd58_err

    buf = "Servc "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Nro   "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = xxxsoles
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Dolares"
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 2, 0)
    buf1 = "select * from " & usuariopos & "01" & " where  servicio='*' or servicio='C' or servicio='D'" 'cuadre 01
    Set mysnap = mydbxglo.CreateSnapshot(buf1)

    Do 'Until mysnap.EOF

        If mysnap.EOF Then Exit Do
        If "" & mysnap.Fields("servicio") = "*" Then
            buf = "Rapid"

        End If

        If "" & mysnap.Fields("servicio") = "C" Then
            buf = "SA:" & mysnap.Fields("salon")

        End If

        If "" & mysnap.Fields("servicio") = "D" Then
            buf = "Domic"

        End If

        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnap.Fields("nro")
        found = formateaa(buf, 6, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnap.Fields("valors")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnap.Fields("valord")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)
        mysnap.MoveNext
    Loop
    mysnap.Close
    
    Exit Sub
cmd58_err:
    MsgBox "Error en imprime servicio"
    mysnap.Close
    Exit Sub

End Sub

Sub imprime_doctos(sw As Integer)

    Dim soles   As Double

    Dim dolares As Double

    Dim buf     As String

    Dim found   As Integer

    Dim buf2    As String

    Dim xsw     As Integer

    On Error GoTo cmd49_err

    'cabecera "DOCUMENTOS EMITIDOS"
    buf2 = ""
    buf = "Tipo "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Nro   "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = xxxsoles
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Dolares"
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 2, 0)
    soles = 0
    dolares = 0
    Set mysnap = mydbxglo.CreateSnapshot(usuariopos & "02")  'cuadre 02
    xsw = 1
    Do

        If mysnap.EOF Then Exit Do
        buf2 = Mid$("" & mysnap.Fields("tipo"), 2, Len("" & mysnap.Fields("tipo")))

        If sw = 2 Then
            If Mid$("" & mysnap.Fields("tipo"), 1, 1) = "E" Or Mid$("" & mysnap.Fields("tipo"), 1, 1) = "S" Then
                GoTo masvalex
                Else: GoTo masvale

            End If

        End If

        If sw = 0 Or sw = 1 Then
            If Mid$("" & mysnap.Fields("tipo"), 1, 1) = "E" Or Mid$("" & mysnap.Fields("tipo"), 1, 1) = "S" Then
                GoTo masvale

            End If

        End If

masvalex:

        If sw = 0 Then
            xsw = 0

            If Val(buf2) <> 5 Then
                xsw = 1

            End If

        End If

        If sw = 1 Then
            xsw = 0

            If Val(buf2) = 5 Then
                xsw = 1

            End If

        End If

        If xsw = 1 Then
            If Len(buf2) > 0 Then
                buf = "" & mysnap.Fields("tipo")
                found = busca_nombre(buf2)
                found = formateaa("", 1, 0, 0)
                buf = "" & mysnap.Fields("nro")
                found = formateaa(buf, 6, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = "" & mysnap.Fields("valors")
                buf = Format(Val(buf), "0.00")
                found = formateaa(buf, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = "" & mysnap.Fields("valord")
                buf = Format(Val(buf), "0.00")
                found = formateaa(buf, 8, 0, 1)
                found = formateaa("", 1, 2, 0)
                soles = soles + Val("" & mysnap.Fields("valors"))
                dolares = dolares + Val("" & mysnap.Fields("valord"))

                If Val("" & mysnap.Fields("nroa")) > 0 Then
                    '---------------------------------
                    found = formateaa("ANULAD", 6, 0, 0)
                    'buf = "" & mysnap.Fields("tipo")
                    'found = busca_nombre(buf)
                    found = formateaa("", 1, 0, 0)
                    buf = "" & mysnap.Fields("nroa")
                    found = formateaa(buf, 6, 0, 1)
                    found = formateaa("", 1, 0, 0)
                    buf = "" & mysnap.Fields("valorsa")
                    buf = Format(Val(buf), "0.00")
                    found = formateaa(buf, 8, 0, 1)
                    found = formateaa("", 1, 0, 0)
                    buf = "" & mysnap.Fields("valorda")
                    buf = Format(Val(buf), "0.00")
                    found = formateaa(buf, 8, 0, 1)
                    found = formateaa("", 1, 2, 0)

                End If

                '---------------------------------
            End If

        End If

masvale:
        mysnap.MoveNext
    Loop
    mysnap.Close

    buf = "Total "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)

    If sw = 0 Then
        buf = "Ventas"

    End If

    If sw = 1 Then
        buf = "Otros "

    End If

    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(soles, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(dolares, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
    Exit Sub
cmd49_err:
    MsgBox "Error en imprime doctos " & error$
    mysnap.Close
    'mydb.Close

    Exit Sub

End Sub

Sub imprime_valorv()

    Dim sdx      As Double

    Dim asoles   As Double

    Dim adolares As Double

    Dim buf2     As String

    Dim soles    As Double

    Dim dolares  As Double

    Dim solesv   As Double

    Dim dolaresv As Double

    Dim igvs     As Double

    Dim igvd     As Double

    Dim buf      As String

    Dim found    As Integer

    Dim exos     As Double

    Dim exod     As Double

    Dim tax1s    As Double

    Dim tax1d    As Double

    Dim dsctos   As Double

    Dim dsctod   As Double

    Dim brutos   As Double

    Dim brutod   As Double

    Dim tresd    As Double

    Dim tress    As Double

    Dim nodsctos As Double

    Dim nodsctod As Double

    Dim FADX     As Double

    Dim cdetras  As Double

    Dim ndetras  As Double

    Dim cdetrad  As Double

    Dim ndetrad  As Double

    On Error GoTo cmd50_err

    cdetras = 0
    ndetras = 0
    cdetrad = 0
    ndetrad = 0

    nodsctos = 0
    nodsctod = 0
    tresd = 0
    tress = 0
    brutos = 0
    brutod = 0
    dsctos = 0
    dsctod = 0
    asoles = 0
    adolares = 0
    solesv = 0
    dolaresv = 0
    soles = 0
    dolares = 0
    igvs = 0
    tax1s = 0
    tax1d = 0
    igvd = 0
    sum1 = 0
    sum2 = 0

    Set mysnap = mydbxglo.CreateSnapshot(usuariopos & "02")  'cuadre 02

    Do 'Until mysnap.EOF

        If mysnap.EOF Then Exit Do
        buf2 = Mid$("" & mysnap.Fields("tipo"), 2, Len("" & mysnap.Fields("tipo")))

        If Mid$("" & mysnap.Fields("tipo"), 1, 1) = "E" Or Mid$("" & mysnap.Fields("tipo"), 1, 1) = "S" Then
            GoTo masvale2

        End If

        If Val(buf2) <> 5 Then
            cdetras = cdetras + Val("" & mysnap.Fields("cdetras"))
            ndetras = ndetras + Val("" & mysnap.Fields("ndetras"))
            cdetrad = cdetrad + Val("" & mysnap.Fields("cdetrad"))
            ndetrad = ndetrad + Val("" & mysnap.Fields("ndetrad"))

            solesv = solesv + Val("" & mysnap.Fields("valorvs"))
            dolaresv = dolaresv + Val("" & mysnap.Fields("valorvd"))
            igvs = igvs + Val("" & mysnap.Fields("igvs"))
            igvd = igvd + Val("" & mysnap.Fields("igvd"))
            exod = exod + Val("" & mysnap.Fields("exod"))
            exos = exos + Val("" & mysnap.Fields("exos"))
            tax1s = tax1s + Val("" & mysnap.Fields("tax1s"))
            tax1d = tax1d + Val("" & mysnap.Fields("tax1d"))
            soles = soles + Val("" & mysnap.Fields("valors"))
            dolares = dolares + Val("" & mysnap.Fields("valord"))
            dsctos = dsctos + Val("" & mysnap.Fields("dsctos"))
            dsctod = dsctod + Val("" & mysnap.Fields("dsctod"))
            brutos = brutos + Val("" & mysnap.Fields("brutos"))
            brutod = brutod + Val("" & mysnap.Fields("brutod"))
            tress = tress + Val("" & mysnap.Fields("retes"))
            tresd = tresd + Val("" & mysnap.Fields("reted"))
            nodsctos = nodsctos + Val("" & mysnap.Fields("nodsctos"))
            nodsctod = nodsctod + Val("" & mysnap.Fields("nodsctod"))

        Else
            asoles = asoles + Val("" & mysnap.Fields("valors"))
            adolares = adolares + Val("" & mysnap.Fields("valord"))

        End If

masvale2:
        mysnap.MoveNext
    Loop
    mysnap.Close

    buf = "Valor Bruto"
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(brutos, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(brutod, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "Descuentos "
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(dsctos, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(dsctod, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "Valor Venta "
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    'soles v
    FADX = solesv - exos

    buf = Format(FADX, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(dolaresv, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "Impuesto"
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(igvs, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(igvd, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "Imp adicional"
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(tax1s, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(tax1d, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
       
    buf = "Detracc.Cobradas"
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(cdetras, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(cdetrad, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "DetraccNoCobrabas"
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(ndetras, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(ndetrad, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "Exonerado "
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(exos, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(exod, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "Otros Dsctos "
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(tress, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(tresd, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "Imp.Excep. "
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(nodsctos, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(nodsctod, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "Total Ventas"
    found = formateaa(buf, 13, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(soles, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(dolares, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
    sum1 = soles + asoles
    sum2 = dolares + adolares

    '-----------------
    If opcion1 = "5" Then  'si es cierre

        'acumulado hasta la fecha
        '--------------se quito
        'sdx = suma_las_ventas()
        'buf = "ACUMUL VTAS. "
        'found = formateaa(buf, 14, 0, 0)
        'buf = Format(sdx, "0.00")
        'found = formateaa(buf, 8, 0, 1)
        'found = formateaa("", 1, 2, 0)
    End If
       
    '-----------------
    Exit Sub
cmd50_err:
    MsgBox "Error en imprime_valorv" & error$, 24, "Aviso"
    mysnap.Close

    Exit Sub

End Sub

Sub imprime_orden_trabajo()

    Dim buf      As String

    Dim found    As Integer

    Dim mytablex As Snapshot

    Dim js       As Double

    Dim jd       As Double

    Dim jindx    As Double

    Dim xsolesx  As Double

    Dim xdolarx  As Double

    On Error GoTo cmd891213

    jindx = 0
    js = 0
    jd = 0
    buf = "select * from cpedidov where "
    buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    buf = buf & " and usuario like '" & extra_loquesea("" & cajero) & "%'"
    buf = buf & " and caja like '" & extra_loquesea("" & caja) & "%'"
    buf = buf & " and turno like '" & extra_loquesea("" & turno) & "%'"
    buf = buf & " and tipo='6'"
    buf = buf & " order by fecha "
    Set mytablex = mydbxglo.CreateSnapshot(buf)
    Do

        If mytablex.EOF Then Exit Do
        jindx = jindx + 1

        If "" & mytablex.Fields("moneda") = "S" Then
            xsolesx = Val("" & mytablex.Fields("total"))
            js = js + Val("" & mytablex.Fields("total"))

        End If

        If "" & mytablex.Fields("moneda") = "D" Then
            jd = jd + Val("" & mytablex.Fields("total"))
            xdolarx = Val("" & mytablex.Fields("total"))

        End If
       
        found = formateaa("" & mytablex.Fields("nombre"), 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = ""
        found = formateaa(buf, 6, 0, 1)
        found = formateaa("", 1, 0, 0)
       
        buf = "" & xsolesx
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & xdolarx
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)
        mytablex.MoveNext
    Loop
    mytablex.Close
      
    found = formateaa("TotalOrden", 14, 0, 0)
    buf = Format(js, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(jd, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
       
    sum1 = sum1 + js
    sum2 = suma2 + jd
    Exit Sub
cmd891213:
    Exit Sub
   
End Sub

Sub imprime_recibos()

    Dim buf     As String

    Dim buf1    As String

    Dim found   As Integer

    Dim sdx     As Double

    Dim sdx1    As Double

    Dim sdx2    As Double

    Dim sdx3    As Double

    Dim mysnapx As Snapshot

    On Error GoTo cmd87912_err

    buf = "Servc "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Nro   "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = xxxsoles
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Dolares"
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 2, 0)
    sum3 = 0
    sum4 = 0
    sdx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0

    buf1 = "select * from " + usuariopos
    buf1 = buf1 + "01"
    buf1 = buf1 + " where servicio<>null and  mid(servicio,1,1)='W' or mid(servicio,1,1)='V' " 'cuadre 01
    'MsgBox buf1
    Set mysnapx = mydbxglo.CreateSnapshot(buf1)

    Do
      
        If mysnapx.EOF Then Exit Do
        If Mid$("" & mysnapx.Fields("servicio"), 1, 1) = "W" Then
            buf = "Ingreso"
            sum3 = sum3 + Val("" & mysnapx.Fields("valors"))
            sum4 = sum4 + Val("" & mysnapx.Fields("valord"))
            sdx = sdx + Val("" & mysnapx.Fields("valors"))
            sdx1 = sdx1 + Val("" & mysnapx.Fields("valord"))

        End If
       
        If Mid$("" & mysnapx.Fields("servicio"), 1, 1) = "V" Then
            buf = "Egreso"
            sum3 = sum3 - Val("" & mysnapx.Fields("valors"))
            sum4 = sum4 - Val("" & mysnapx.Fields("valord"))
            sdx2 = sdx2 + Val("" & mysnapx.Fields("valors"))
            sdx3 = sdx3 + Val("" & mysnapx.Fields("valord"))

        End If

        buf = busca_tipo(Mid$("" & mysnapx.Fields("servicio"), 2, Len("" & mysnapx.Fields("servicio"))))
        'found = formateaa(Mid$("" & mysnapx.Fields("servicio"), 1, 1) & "*", 2, 0, 0)
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnapx.Fields("nro")
        found = formateaa(buf, 6, 0, 1)
        found = formateaa("", 2, 0, 0)
        buf = "" & mysnapx.Fields("valors")

        If Val(buf) > 0 Then
            buf = Format(Val(buf), "0.00")

        End If

        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnapx.Fields("valord")

        If Val(buf) > 0 Then
            buf = Format(Val(buf), "0.00")

        End If

        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)
        mysnapx.MoveNext
    Loop
    mysnapx.Close
    
    If sdx > 0 Or sdx1 > 0 Then
        buf = "TOT INGRESOS "
        found = formateaa(buf, 15, 0, 0)

        buf = Format(sdx, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = Format(sdx1, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)

    End If

    If sdx2 > 0 Or sdx3 > 0 Then
        buf = "TOT EGRESOS "
        found = formateaa(buf, 15, 0, 0)

        buf = Format(sdx2, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = Format(sdx3, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)

    End If

    Exit Sub
cmd87912_err:
    MsgBox "11.Mensaje en Imprime_recibos " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub forma_pago()

    Dim vr, buf, buf1, buf2 As String

    Dim sdx1     As Double

    Dim sdx      As Double

    Dim asola    As String

    Dim mytablex As Table

    Dim signos   As Double

    On Error GoTo cmd230_err

    sum1 = 0
    sdx1 = 0
    Set mytablex = mydbxglo.OpenTable(usuariopos & "03")  'cuadre 03
    mytablex.Index = "tipo"
    Do

        If mytablex.EOF Then Exit Do
        mytablex.Delete
        mytablex.MoveNext
    Loop
    buf = "select * from fpagov where "
    buf = buf & " fecha>=" & "DateValue('" & fechai & "'" & ")"
    buf = buf & " and fecha<=" & "DateValue('" & fechaf & "'" & ")"

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    buf = buf & " and usuario like '" & extra_loquesea("" & cajero) & "%'"
    buf = buf & " and caja like '" & extra_loquesea("" & caja) & "%'"
    buf = buf & " and turno like '" & extra_loquesea("" & turno) & "%'"
    buf = buf & " and (acu='I' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='W' or acu='V' OR ACU='1') "  'E nota credito
    buf = buf & " and estado='2'"
    buf = buf & " order by fecha"
    Set mysnap = mydbxglo.CreateSnapshot(buf)
    Do

        If mysnap.EOF Then Exit Do
        signos = 1

        If "" & mysnap.Fields("acu") = "E" Then     'nota credito
            signos = -1

        End If

        'formas de pago
        sum1 = sum1 + 1
        fecha = "FORMA DE PAGO ..." & Format(sum1, "00000")
        buf2 = "" & mysnap.Fields("servicio")

        If buf2 = "V" Then
            buf2 = "E"

        End If

        If buf2 <> "E" Then
            buf2 = "I"

        End If

        '----
       
        mytablex.Seek "=", "" & mysnap.Fields("fpago"), buf2

        If Not mytablex.NoMatch Then
            mytablex.Edit
            sdx1 = suma_fpago(buf2, mytablex, signos)
            mytablex.Update

            If mysnap.Fields("moneda") = "D" And sdx1 < 0 Then
                forma_pago1 buf2, sdx1, mytablex

            End If

        End If

        If mytablex.NoMatch Then
            mytablex.AddNew
            sdx1 = suma_fpago(buf2, mytablex, signos)
            mytablex.Fields("local") = "01"
            mytablex.Update

            If mysnap.Fields("moneda") = "D" And sdx1 < 0 Then
                forma_pago1 buf2, sdx1, mytablex

            End If

        End If

        '----
        mysnap.MoveNext
    Loop
    mysnap.Close
    mytablex.Close
    Exit Sub
cmd230_err:
    MsgBox "Error en Forma de Pago1 " & error$, 24, "Aviso"
    mysnap.Close
    mytablex.Close
    Exit Sub

End Sub

Sub borrar_cuadres()

    Dim mytablex As Table

    Dim sw       As String

    On Error GoTo cmd4561_err
   
    sw = "1"
    Set mytablex = mydbxglo.OpenTable(usuariopos & "01")  'cuadre 01
    Do

        If mytablex.EOF Then Exit Do
        mytablex.Delete
        mytablex.MoveNext
    Loop
    mytablex.Close
    sw = "2"
   
    Set mytablex = mydbxglo.OpenTable(usuariopos & "02")  'cuadre 02
    Do

        If mytablex.EOF Then Exit Do
        mytablex.Delete
        mytablex.MoveNext
    Loop
    mytablex.Close
   
    sw = "3"
   
    Set mytablex = mydbxglo.OpenTable(usuariopos & "03")      'cuadre 03
    Do

        If mytablex.EOF Then Exit Do
        mytablex.Delete
        mytablex.MoveNext
    Loop
    mytablex.Close
   
    sw = "4"
   
    Set mytablex = mydbxglo.OpenTable(usuariopos & "04")           'cuadre 04
    Do

        If mytablex.EOF Then Exit Do
        mytablex.Delete
        mytablex.MoveNext
    Loop
    mytablex.Close
   
    Exit Sub
cmd4561_err:
    MsgBox "Error en borra cuadres " & error & " " & sw, 24, "Aviso"
    mytablex.Close
    Exit Sub

End Sub

Function graba_cierres(buf As String) As Double

    Dim mytablex As Table

    On Error GoTo cmd_34emp1

    Dim sdx As Double
   
    Set mytablex = mydbxglo.OpenTable("parameca")
    mytablex.Index = "caja"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then
        sdx = Val("" & mytablex.Fields("cierres")) + 1
        mytablex.Edit
        mytablex.Fields("cierres") = Format(sdx, "00000")
        mytablex.Update
        graba_cierres = sdx

    End If

    mytablex.Close
    Exit Function
cmd_34emp1:

    If Err <> 3260 And Err <> 3186 And Err <> 3187 And Err <> 3158 And Err <> 3046 And Err <> 3202 And Err <> 3164 And Err <> 3188 And Err <> 3218 And Err <> 3006 And Err <> 3197 And Err <> 3189 And Err <> 3022 Then
        mytablex.Close
        Exit Function

    End If

    MsgBox " PRESIONE ENTER Y CONTINUA " & error(Err), 24, "AVISO DE NO ERROR"
    Resume

End Function

Function busca_config(sw As Integer) As String

    Dim mytablex As Table

    On Error GoTo cmd6711_err

    Set mytablex = mydbxglo.OpenTable("parame")
    mytablex.Index = "codigo"
    mytablex.Seek "=", "01"

    If Not mytablex.NoMatch Then

        If sw = 0 Then
            busca_config = "" & mytablex.Fields("centraliza")

        End If

        If sw = 1 Then
            busca_config = "" & mytablex.Fields("vdolar")

        End If

        If sw = 2 Then
            busca_config = "" & mytablex.Fields("tipo5")

        End If

    End If

    mytablex.Close
    Exit Function
  
cmd6711_err:
    mytablex.Close
    MsgBox "Error en busca_config " + error, 48, "Aviso"
    Exit Function
   
End Function

Sub imprime_fpago()

    Dim buf   As String

    Dim found As Integer

    Dim buf1  As String

    Dim sw    As Integer

    Dim isoles, idolares As Double

    Dim esoles, edolares As Double

    Dim ssoles, sdolares As Double

    Dim xsoles, xdolares As Double

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim sdx2     As Double

    Dim sdx3     As Double

    Dim sdx4     As Double

    Dim psoles   As Double

    Dim pdolares As Double

    On Error GoTo cmd9999_err

    Dim pmoneda As String

    buf = "Fpago "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Nro   "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = xxxsoles
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Dolares"
    found = formateaa(buf, 8, 0, 0)
    found = formateaa("", 1, 2, 0)

    isoles = 0
    idolares = 0
    
    Set mysnap = mydbxglo.CreateSnapshot("select * from " & usuariopos & "03" & " where  servicio='I' order by tipo") 'cuadre 03

    Do
    
        If mysnap.EOF Then Exit Do
        sdx = 0
        fecha = "NO TOQUE EL TECLADO..."
        buf = "" & mysnap.Fields("tipo")
        found = busca_fpago(buf, psoles, pdolares)
    
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnap.Fields("nro")
        found = formateaa(buf, 6, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = "" & mysnap.Fields("valors")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf = "" & mysnap.Fields("valord")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)

        '----------------------------------
        If psoles > 0 Or pdolares > 0 Then
            found = formateaa("*DECLARADO ", 14, 0, 0)
            buf = Format(psoles, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa(buf, 8, 0, 1)
            buf = Format(pdolares, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa("", 1, 0, 0)
            found = formateaa(buf, 8, 2, 1)
            found = formateaa("*DIFERENCIA ", 14, 0, 0)
            sdx = psoles - Val("" & mysnap.Fields("valors"))
            buf = Format(sdx, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa(buf, 8, 0, 1)
            sdx = pdolares - Val("" & mysnap.Fields("valord"))
            buf = Format(sdx, "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa("", 1, 0, 0)
            found = formateaa(buf, 8, 2, 1)

        End If

        '----------------------------------
        isoles = isoles + Val("" & mysnap.Fields("valors"))
        idolares = idolares + Val("" & mysnap.Fields("valord"))
        mysnap.MoveNext
    Loop
    mysnap.Close
    
    buf = "Total "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Fpago "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(isoles, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(idolares, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    buf = "FORMA DE PAGO/EGRESOS"
    found = formateaa(buf, Len(buf), 2, 0)

    esoles = 0
    edolares = 0

    Set mysnap = mydbxglo.CreateSnapshot("select * from " & usuariopos & "03" & " where  servicio='E'") 'cuadre 03
    Do

        If mysnap.EOF Then Exit Do
        buf = "" & mysnap.Fields("tipo")
        found = busca_fpago(buf, psoles, pdolares)
        'found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnap.Fields("nro")
        found = formateaa(buf, 6, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnap.Fields("valors")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mysnap.Fields("valord")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)
        esoles = esoles + Val("" & mysnap.Fields("valors"))
        edolares = edolares + Val("" & mysnap.Fields("valord"))
        mysnap.MoveNext
    Loop
    mysnap.Close

    buf = "Total "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Fpago "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(esoles, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    buf = Format(edolares, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
    'saldos
    sw = 0
    ssoles = 0
    sdolares = 0
    xsoles = 0
    xdolares = 0
    buf = "FORMA DE PAGO/SALDOS"
    found = formateaa(buf, Len(buf), 2, 0)

    Set mysnap = mydbxglo.CreateSnapshot("select * from " & usuariopos & "03" & "  order by tipo ") 'cuadre 03

    Do

        If mysnap.EOF Then Exit Do
        If sw = 0 Then
            buf1 = "" & mysnap.Fields("tipo")
            buf = "" & mysnap.Fields("tipo")
            found = busca_fpago(buf, psoles, pdolares)
            'found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 8, 0, 0)
            sw = 1

        End If

        If buf1 <> "" & mysnap.Fields("tipo") Then
            buf = Format(ssoles, "0.00")
            found = formateaa(buf, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(sdolares, "0.00")
            found = formateaa(buf, 8, 0, 1)
            found = formateaa("", 1, 2, 0)

            buf1 = "" & mysnap.Fields("tipo")
            buf = "" & mysnap.Fields("tipo")
            found = busca_fpago(buf, psoles, pdolares)
            'found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 8, 0, 0)
            ssoles = 0
            sdolares = 0

        End If

        If "" & mysnap.Fields("servicio") <> "E" Then
            ssoles = ssoles + Val("" & mysnap.Fields("valors"))
            sdolares = sdolares + Val("" & mysnap.Fields("valord"))
            xsoles = xsoles + Val("" & mysnap.Fields("valors"))
            xdolares = xdolares + Val("" & mysnap.Fields("valord"))

        End If

        If "" & mysnap.Fields("servicio") = "E" Then
            ssoles = ssoles - Val("" & mysnap.Fields("valors"))
            sdolares = sdolares - Val("" & mysnap.Fields("valord"))
            xsoles = xsoles - Val("" & mysnap.Fields("valors"))
            xdolares = xdolares - Val("" & mysnap.Fields("valord"))

        End If

        mysnap.MoveNext
    Loop
    mysnap.Close
    
    'lo puse en el peaje
    'If ssoles > 0 Or sdolares > 0 Then
    buf = Format(ssoles, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(sdolares, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
    'End If
    'Exit Sub
    '-----------------------------------------------------
    'OJO ES TEMPORAL JOHNNY SOLO PARA VICUS
    buf = "Subtotal "
    found = formateaa(buf, 14, 0, 0)
    buf = Format(xsoles, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(xdolares, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)

    '-----------------------------------------------------
    sdx = busca_igv()

    If sdx = 0 Then
        sdx = 1

    End If

    sdx1 = xsoles + xdolares * sdx
    sdx1 = Format(sdx1, "0.00")
    sdx2 = sdx1 / sdx
    sdx2 = Format(sdx2, "0.00")

    buf = String(35, "-")
    found = formateaa(buf, 35, 2, 0)
    'SOLO TEMPORAL------------------
    buf = "Total "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = " "
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)

    If busca_config(1) = "N" Then
        sdx2 = 0

    End If

    buf = Format(sdx2, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 2, 0)
    '---ABACA AQUI---
    Exit Sub
cmd9999_err:
    MsgBox "Error en Imprime Fpago1 " & error$, 24, "Aviso "
    Exit Sub

End Sub

Function busca_cierre(buf As String)

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("parameca")
    mytablex.Index = "caja"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then
        busca_cierre = "" & mytablex.Fields("cierres")

    End If

    mytablex.Close

End Function

Function busca_nombre(buf1 As String)

    Dim mytablex As Table

    Dim buf      As String

    Dim buf3     As String

    Dim found    As Integer

    buf3 = ""
    Set mytablex = mydbxglo.OpenTable("tipo")
    mytablex.Index = "tipo"
    mytablex.Seek "=", buf1

    If Not mytablex.NoMatch Then
        buf3 = Mid$("" & mytablex.Fields("descripcio"), 1, 6)
        busca_nombre = 1

    End If

    mytablex.Close
    found = formateaa(buf3, 6, 0, 0)

End Function

Function busca_tipo(buf As String) As String

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("tipo")
    mytablex.Index = "tipo"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then
        busca_tipo = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

Function suma_fpago(buf As String, mytablex As Table, signos As Double) As Double

    Dim sdx  As Double

    Dim buf1 As String

    On Error GoTo cmd4556_err

    suma_fpago = 0

    If Val("" & mysnap.Fields("tipo")) = 5 Then
        If todos <> "S" Then Exit Function

    End If

    If Val("" & mysnap.Fields("estado")) = 2 Then
        mytablex.Fields("tipo") = "" & mysnap.Fields("fpago")
        mytablex.Fields("servicio") = buf
        sdx = Val("" & mytablex.Fields("nro")) + 1
        mytablex.Fields("nro") = sdx

        If "" & mysnap.Fields("moneda") = "S" Then
            If Val("" & mysnap.Fields("saldos")) <= 0 Then
                mytablex.Fields("valors") = Val("" & mytablex.Fields("valors")) + signos * Val(Format(Val("" & mysnap.Fields("recibe")) + Val("" & mysnap.Fields("saldos")), "0.00"))
            Else
                mytablex.Fields("valors") = Val("" & mytablex.Fields("valors")) + signos * Val(Format(Val("" & mysnap.Fields("recibe"))))

            End If

        End If

        If "" & mysnap.Fields("moneda") = "D" Then
            mytablex.Fields("valord") = Val("" & mytablex.Fields("valord")) + signos * Val(Format(Val("" & mysnap.Fields("recibe")), "0.00"))
            buf1 = Format(Val("" & mysnap.Fields("saldos")), "0.00")
            'mytablex.Fields("valord") = Val("" & mytablex.Fields("valord")) + signos * Val("" & mysnap.Fields("recibed"))
            suma_fpago = Val(buf1)

        End If

    End If

    Exit Function
cmd4556_err:
    MsgBox "Error en Suma Fpago " & error$, 24, "Aviso"
    Exit Function
      
End Function

Sub forma_pago1(buf2 As String, sdx1 As Double, mytablex As Table)

    Dim sdx As Double
      
    If "" & mysnap.Fields("fpago") = "2" Then
        mytablex.Seek "=", "1", buf2  'busco soles  1+servicio

        '---------------
        If mytablex.NoMatch Then
            mytablex.AddNew
            mytablex.Fields("local") = "01"
            mytablex.Fields("tipo") = "1"
            mytablex.Fields("servicio") = buf2
            sdx = Val("" & mytablex.Fields("valors")) + sdx1
            mytablex.Fields("valors") = Val(Format(sdx, "0.00"))
            mytablex.Update

        End If

        If Not mytablex.NoMatch Then
            mytablex.Edit
            
            sdx = Val("" & mytablex.Fields("valors")) + sdx1
            mytablex.Fields("valors") = Format(sdx, "0.00")
            mytablex.Update

        End If

        '---------------------
    End If

End Sub

Function busca_fpago(buf1 As String, sdx As Double, sdx1 As Double)

    Dim buf      As String

    Dim buf3     As String

    Dim found    As Integer

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("fpago")
    mytablex.Index = "fpago"
    mytablex.Seek "=", buf1

    If Not mytablex.NoMatch Then
        buf3 = "" & mytablex.Fields("descripcio")

        If "" & mytablex.Fields("moneda") = "S" Then
            sdx = Val("" & mytablex.Fields("total"))
            sdx1 = 0

        End If

        If "" & mytablex.Fields("moneda") = "D" Then
            sdx1 = Val("" & mytablex.Fields("total"))
            sdx = 0

        End If

        busca_fpago = 1

    End If

    mytablex.Close
   
    found = formateaa(buf3, 6, 0, 0)

End Function

