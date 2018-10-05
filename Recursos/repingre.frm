VERSION 5.00
Begin VB.Form repingre 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes Caja"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox subconcepto 
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   3600
      Width           =   3015
   End
   Begin VB.ComboBox concepto 
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   3240
      Width           =   3015
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
      TabIndex        =   35
      Top             =   120
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2520
      Width           =   1575
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   2880
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   2160
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   480
      Width           =   3855
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox nrolineas 
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
      TabIndex        =   11
      Text            =   "45"
      Top             =   5520
      Width           =   1575
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
      TabIndex        =   10
      Top             =   5160
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   9
      Top             =   3000
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   8
      Top             =   2520
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
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   7
      Text            =   "%"
      Top             =   840
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
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   6
      Text            =   "%"
      Top             =   1680
      Width           =   1575
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
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   5
      Text            =   "%"
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox observa 
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
      MaxLength       =   11
      TabIndex        =   4
      Text            =   "%"
      Top             =   4320
      Width           =   1575
   End
   Begin VB.ComboBox vdetalle 
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox vfpago 
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ComboBox estado 
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
      TabIndex        =   1
      Top             =   4680
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
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   0
      Text            =   "%"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Subconcepto"
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
      Left            =   3960
      TabIndex        =   42
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Concepto"
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
      Left            =   3960
      TabIndex        =   41
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label xcuentaco1 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   6240
      Width           =   105
   End
   Begin VB.Label xcuentaco 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   6000
      Width           =   105
   End
   Begin VB.Label Label7 
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
      TabIndex        =   36
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label19 
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
      Left            =   3960
      TabIndex        =   34
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label20 
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
      Left            =   3960
      TabIndex        =   33
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label22 
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
      Left            =   3960
      TabIndex        =   32
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ver.Documentos"
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
      Left            =   3960
      TabIndex        =   27
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   26
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label acu 
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
      Left            =   7680
      TabIndex        =   25
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lineas x Pagina"
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
      TabIndex        =   24
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label2 
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
      TabIndex        =   23
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label Label17 
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
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label5 
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
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label10 
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
      Left            =   120
      TabIndex        =   20
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label11 
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
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label21 
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
      Left            =   120
      TabIndex        =   18
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label3 
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
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observacion"
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
      TabIndex        =   16
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VerFormaPago"
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
      Left            =   3960
      TabIndex        =   15
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label14 
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
      Left            =   120
      TabIndex        =   14
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label15 
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
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Menu eju3453 
      Caption         =   "&Ejecutar"
   End
   Begin VB.Menu ldfo232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "repingre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub concepto_Click()

    If concepto = "%" Then Exit Sub
    carga_subconcepto "" & extra_loquesea(concepto)

End Sub

Private Sub eju3453_Click()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    contpag = 0
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    found = sql_documento(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    'MsgBox ""
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento
    cuerpo_programa_documento mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Private Sub Form_Activate()

    Dim mytablex As New ADODB.Recordset

    tipo.Clear
    tipo.AddItem "%"
    mytablex.Open "select * from tipo ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do

        'If "" & mytablex.Fields("tipodoc") = "W" Or "" & mytablex.Fields("tipodoc") = "V" Then
        If "" & mytablex.Fields("tipodoc") = acu Then
            tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

        End If

        mytablex.MoveNext
    Loop
    '------------ otros
    mytablex.Close
    caja.Clear
    caja.AddItem "%"
    mytablex.Open "select * from parameca ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        caja.AddItem "" & mytablex.Fields("caja")
        mytablex.MoveNext
    Loop
    mytablex.Close
    caja.ListIndex = 0

    turno.Clear
    turno.AddItem "%"
    mytablex.Open "select * from turno ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        turno.AddItem "" & mytablex.Fields("turno")
        mytablex.MoveNext
    Loop
    mytablex.Close
    turno.ListIndex = 0

    cajero.Clear
    cajero.AddItem "%"
    mytablex.Open "select * from vendedor ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        cajero.AddItem "" & mytablex.Fields("codigo")
        mytablex.MoveNext
    Loop
    mytablex.Close
    cajero.ListIndex = 0
    tipo.ListIndex = 0

    local1.Clear
    local1.AddItem "%"
    mytablex.Open "select * from tlocal ", cn, adOpenStatic, adLockOptimistic
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

    concepto.Clear
    subconcepto.Clear

    concepto.AddItem "%"
    mytablex.Open "select * from concepto ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        concepto.AddItem "" & mytablex.Fields("concepto") & "|" & mytablex.Fields("DESCRIPCIO")
        mytablex.MoveNext
    Loop
    mytablex.Close
    concepto.ListIndex = 0

    subconcepto.AddItem "%"
    subconcepto.ListIndex = 0

End Sub

Private Sub Form_Load()

    Dim mytablex As Table

    fechaf = Format(Now, "dd/mm/yyyy")
    fechai = "01" & "/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")

    vdetalle.AddItem "N"
    vdetalle.AddItem "S"
    vdetalle.ListIndex = 0

    vfpago.AddItem "N"
    vfpago.AddItem "S"
    vfpago.ListIndex = 0

    estado.AddItem "%"
    estado.AddItem "2"
    estado.AddItem "1"
    estado.AddItem "0"
    estado.ListIndex = 0
    moneda.Clear
    moneda.AddItem "%"
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0

End Sub

Private Sub ldfo232_Click()
    repingre.Hide
    Unload repingre

End Sub

Function sql_documento(mytablex As ADODB.Recordset)

    Dim buf  As String

    Dim xbuf As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    buf = "select * from recibo where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If concepto <> "%" Then
        buf = buf & " and concepto='" & extra_loquesea(concepto) & "'"

    End If

    If subconcepto <> "%" Then
        buf = buf & " and subconcepto='" & extra_loquesea(subconcepto) & "'"

    End If

    If tipo <> "%" Then
        xbuf = extra_loquesea(tipo)
        buf = buf & " and tipo like '" & xbuf & "'"

    End If

    If local1 <> "%" Then
        buf = buf & " and local like '" & local1 & "'"

    End If

    If Numero <> "%" Then
        buf = buf & " and numero like '" & Numero & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "'"

    End If

    If estado <> "%" Then
        buf = buf & " and estado like '" & estado & "'"

    End If

    If acu <> "%" Then
        buf = buf & " and acu='" & acu & "'"

    End If

    'If acu = "W" Then
    'buf = buf & " and (acu='A' OR acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F' or acu='W')"
    'End If
    'If acu = "V" Then
    'buf = buf & " and (acu='J' OR acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O' or acu='V')"
    'End If
    buf = buf & " order by concepto,fecha,codigo"
    'MsgBox buf
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    sql_documento = 1

End Function

Sub cabecera_documento()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)
    found = formateaa("E", 2, 0, 0)
    found = formateaa("Fecha", 11, 0, 0)
    found = formateaa("ser", 4, 0, 0)
    found = formateaa("Numero", 12, 0, 0)
    found = formateaa("X", 2, 0, 0)
    found = formateaa("M", 2, 0, 0)
    found = formateaa("Total ", 11, 0, 1)
    found = formateaa("T", 2, 0, 0)
    found = formateaa("Codigo", 12, 0, 0)
    found = formateaa("Nombre", 31, 0, 0)
    found = formateaa("Motivo", 26, 0, 0)
    found = formateaa("Cajero ", 9, 0, 0)
    found = formateaa("Caj", 4, 0, 0)
    found = formateaa("T", 2, 0, 0)
    found = formateaa("Hora", 6, 0, 0)
    found = formateaa("Cob/Ven", 9, 0, 0)
    found = formateaa("Local", 7, 2, 0)

    'cabecera
    If vdetalle = "S" Then

    End If
    
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)

End Sub

Sub cuerpo_programa_documento(mytablex As ADODB.Recordset)

    Dim Tmp   As String

    Dim sw    As Integer

    Dim buf   As String

    Dim found As Integer

    Dim sdx   As Double

    sdx = 0
    sw = 0
    suma1 = 0
    suma2 = 0
    ssuma1 = 0
    ssuma2 = 0
    Do

        If mytablex.EOF Then Exit Do
        If sw = 0 Then
            buf = "" & mytablex.Fields("concepto")
            found = formateaa(buf, 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = busca_tipo("" & mytablex.Fields("concepto"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            sw = 1
            suma1 = 0
            suma2 = 0
            suma3 = 0
            Tmp = "" & mytablex.Fields("concepto")

        End If

        If Tmp <> "" & mytablex.Fields("concepto") Then
            found = formateaa("", 19, 0, 0)
            found = formateaa(dicmoneda, 7, 0, 0)
            buf = Format(suma1, "0.00")
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("Dolares", 8, 0, 0)
            buf = Format(suma2, "0.00")
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
   
            buf = "" & mytablex.Fields("concepto")
            found = formateaa(buf, 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = busca_tipo("" & mytablex.Fields("concepto"))
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas
            Tmp = "" & mytablex.Fields("concepto")
            suma1 = 0

        End If

        buf = "" & mytablex.Fields("estado")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("fecha")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("serie")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("numero")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("acu")

        If buf = "W" Then
            buf = "I"

        End If

        If buf = "V" Then
            buf = "E"

        End If

        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("moneda")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("total")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("tipoclie")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("codigo")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("nombre")
        found = formateaa(buf, 30, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("observa")
        found = formateaa(buf, 25, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("usuario")
        found = formateaa(buf, 8, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("caja")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("turno")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("hora")
        found = formateaa(buf, 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("vendedor")
        found = formateaa(buf, 8, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("Local")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 2, 0)
   
        nlineas

        If "" & mytablex.Fields("moneda") = "S" Then
            suma1 = suma1 + Val("" & mytablex.Fields("total"))
            ssuma1 = ssuma1 + Val("" & mytablex.Fields("total"))

        End If

        If "" & mytablex.Fields("moneda") = "D" Then
            suma2 = suma2 + Val("" & mytablex.Fields("total"))
            ssuma2 = ssuma2 + Val("" & mytablex.Fields("total"))

        End If

        If vdetalle = "S" Then

            'ver_detalle mydbx, mytablex
        End If

        mytablex.MoveNext
    Loop
    found = formateaa("", 19, 0, 0)
    found = formateaa(dicmoneda, 7, 0, 0)
    buf = Format(suma1, "0.00")
    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Dolares", 8, 0, 0)
    buf = Format(suma2, "0.00")
    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 2, 0)
    nlineas
   
    found = formateaa("Total------> ", 19, 0, 1)
    found = formateaa(dicmoneda, 7, 0, 0)
    buf = Format(ssuma1, "0.00")
    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Dolares", 8, 0, 0)
    buf = Format(ssuma2, "0.00")
    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 2, 0)
   
End Sub

Sub ver_detalle(mytabley As ADODB.Recordset)

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    Dim found    As Integer

    mytablex.Open "select * from " & dusuariog & " where tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "' and numero='" & "" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    nlineas
    sw = 0
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("codigo") = "" & mytabley.Fields("codigo") And "" & mytablex.Fields("acu") = "" & mytabley.Fields("acu") Then
            sw = 1
            found = formateaa("%", 1, 0, 0)
            buf = "" & mytablex.Fields("producto")
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("descripcio")
            found = formateaa(buf, 27, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("unidad")
            found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("factor")
            found = formateaa(buf, 4, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("cantidad")
            found = formateaa(buf, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("precio")
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("total")
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        mytablex.MoveNext
    Loop

    If sw = 1 Then
        buf = String(130, "-")
        found = formateaa(buf, 130, 2, 0)
        nlineas

    End If

    mytablex.Close

End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > Val(nrolineas) Then
        cabecera_documento

    End If

End Sub

Function busca_tipo(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from concepto where concepto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_tipo = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

Sub carga_subconcepto(buf As String)

    Dim mytablex As New ADODB.Recordset

    subconcepto.Clear
    subconcepto.AddItem "%"
    mytablex.Open "select * from subconcepto where concepto='" & buf & "'", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        subconcepto.AddItem "" & mytablex.Fields("subconcepto") & "|" & mytablex.Fields("DESCRIPCIO")
        mytablex.MoveNext
    Loop
    mytablex.Close
    subconcepto.ListIndex = 0

End Sub

