VERSION 5.00
Begin VB.Form REPLETRA 
   BackColor       =   &H00808080&
   Caption         =   "Reporte Letras "
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   8775
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
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
      Top             =   1920
      Width           =   1575
   End
   Begin VB.ComboBox rango 
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
      TabIndex        =   27
      Top             =   1560
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
      TabIndex        =   10
      Text            =   "45"
      Top             =   4320
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
      TabIndex        =   9
      Top             =   3960
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
      Left            =   6120
      MaxLength       =   10
      TabIndex        =   12
      Top             =   1080
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
      Left            =   6120
      MaxLength       =   10
      TabIndex        =   11
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox letra 
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
      TabIndex        =   0
      Text            =   "*"
      Top             =   360
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
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox aceptante 
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
      TabIndex        =   1
      Text            =   "*"
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox seccion 
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
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox girador 
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
      TabIndex        =   2
      Text            =   "*"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox banco 
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
      MaxLength       =   6
      TabIndex        =   3
      Text            =   "*"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox agencia 
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
      Text            =   "*"
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox refactura 
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
      Text            =   "*"
      Top             =   3120
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
      MaxLength       =   30
      TabIndex        =   8
      Text            =   "*"
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estados"
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
      TabIndex        =   30
      Top             =   1920
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
      Left            =   3960
      TabIndex        =   28
      Top             =   360
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
      TabIndex        =   26
      Top             =   4320
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
      TabIndex        =   25
      Top             =   3960
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
      Left            =   3960
      TabIndex        =   24
      Top             =   1080
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
      Left            =   3960
      TabIndex        =   23
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro.Letra"
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
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label6 
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
      TabIndex        =   21
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aceptante"
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
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seccion"
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
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Girador"
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
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Banco"
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
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agencia"
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
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Letra,Factura "
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
      TabIndex        =   15
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label14 
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
      TabIndex        =   14
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label21 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
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
      TabIndex        =   13
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Menu dklier 
      Caption         =   "&Ejecuta"
   End
   Begin VB.Menu ldsao2 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "REPLETRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dklier_Click()

    Dim found    As Integer

    Dim mytablex As Snapshot

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

    found = sql_letra(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_letra
    cuerpo_programa_letra mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    found = valida_wordpad(FileName)
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1

End Sub

Private Sub Form_Load()

    Dim mytablex As Table

    Combo1.AddItem "Todos"
    Combo1.AddItem "SoloLetras"
    Combo1.AddItem "Protestados"
    Combo1.AddItem "Renovados"

    Combo1.ListIndex = 1

    fechaf = Format(Now, "dd/mm/yyyy")
    fechai = Format(Now, "dd/mm/yyyy")

    Set mytablex = mydbxglo.OpenTable("carsec")
    seccion.AddItem "%"
    Do

        If mytablex.EOF Then Exit Do
        seccion.AddItem "" & mytablex.Fields("carsec") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
 
    seccion.ListIndex = 0

    moneda.Clear
    moneda.AddItem "%"
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0

    rango.Clear
    rango.AddItem "TODOS"
    rango.AddItem "VENCIDOS"
    rango.AddItem "PENDIENTE"
    rango.AddItem "CANCELADO"
    rango.ListIndex = 1

End Sub

Private Sub ldsao2_Click()
    REPLETRA.Hide
    Unload REPLETRA

End Sub

Function sql_letra(mytablex As Snapshot)

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    If acu = "V" Then
        buf = "select * from letrav where letra like '" & letra & "'"

    End If

    If acu = "C" Then
        buf = "select * from letrac where letra like '" & letra & "'"

    End If

    If rango <> "VENCIDOS" Then
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

        'buf = buf & "  and fechai>=" & "DateValue('" & Fechai & "'" & ")"
        'buf = buf & " and  fechaf<=" & "DateValue('" & fechaf & "'" & ")"
    End If

    If rango = "VENCIDOS" Then
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

        'buf = buf & " and  fechaf>=" & "DateValue('" & Fechai & "'" & ")"
        'buf = buf & " and  fechaf<=" & "DateValue('" & fechaf & "'" & ")"
    End If

    If aceptante <> "%" Then
        buf = buf & " and aceptante like '" & aceptante & "'"

    End If

    If girador <> "%" Then
        buf = buf & " and girador like '" & girador & "'"

    End If

    If banco <> "%" Then
        buf = buf & " and banco like '" & banco & "'"

    End If

    If seccion <> "%" Then
        buf = buf & " and seccion like '" & extra_loquesea(seccion) & "'"

    End If

    If agencia <> "%" Then
        buf = buf & " and agencia like '" & agencia & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If rango = "PENDIENTE" Or rango = "VENCIDOS" Then
        buf = buf & " and saldo>0 "

    End If

    If rango = "CANCELADO" Then
        buf = buf & " and saldo=0 "

    End If

    If Combo1 = "Protestados" Then 'solo protestado
        buf = buf & " and estadop='1' "

    End If

    If Combo1 = "Renovados" Then 'solo Renovados
        buf = buf & " and estado='1' "

    End If

    If Combo1 = "SoloLetras" Then 'solo Renovados
        buf = buf & " and (estado<>'1' and estadop<>'1') "

    End If

    buf = buf & "order by aceptante,Fechaf "
    Set mytablex = mydbxglo.CreateSnapshot(buf)
    sql_letra = 1

End Function

Sub cabecera_letra()

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
    
    buf = String(133, "-")
    found = formateaa(buf, 133, 2, 0)
    'found = formateaa("Aceptante", 12, 2, 0)
    found = formateaa("Letra", 12, 0, 0)
    found = formateaa("Girador", 12, 0, 0)
    found = formateaa("Fechai", 11, 0, 0)
    found = formateaa("Fechaf", 11, 0, 0)
    found = formateaa("M", 2, 0, 0)
    found = formateaa("Bco", 4, 0, 0)
    found = formateaa("Sec", 4, 0, 0)
    found = formateaa("Saldo ", 11, 0, 1)
    found = formateaa("Abono ", 11, 0, 1)
    found = formateaa("Total ", 11, 0, 1)
    found = formateaa("Nro.Unico ", 12, 0, 0)
    found = formateaa("Oct.Dia ", 11, 0, 0)
    found = formateaa("Negociado ", 12, 0, 0)
    found = formateaa("R", 2, 0, 0)
    found = formateaa("P", 2, 2, 0)
    
    buf = String(133, "-")
    found = formateaa(buf, 133, 2, 0)
   
End Sub

Sub cuerpo_programa_letra(mytablex As Snapshot)

    Dim Tmp   As String

    Dim sw    As Integer

    Dim buf   As String

    Dim found As Integer

    Dim sdx   As Double

    sdx = 0
    sw = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0
    suma7 = 0
    suma8 = 0
    suma9 = 0
    Do

        If mytablex.EOF Then Exit Do
        If sw = 0 Then
            buf = "" & mytablex.Fields("aceptante")
            found = formateaa(buf, 11, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = busca_nombre(buf)
            found = formateaa(buf, 60, 2, 0)
            nlineas
            sw = 1
   
            Tmp = "" & mytablex.Fields("aceptante")

        End If

        If Tmp <> "" & mytablex.Fields("aceptante") Then
            found = formateaa("Subneto ", 56, 0, 1)
            buf = Format(suma1, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(suma2, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(suma3, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 2, 0)
            nlineas

            If suma4 > 0 Then
                found = formateaa("P ", 56, 0, 1)
                buf = Format(suma4, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = Format(suma5, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = Format(suma6, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 2, 0)
                nlineas

            End If

            If suma7 > 0 Then
                found = formateaa("R ", 56, 0, 1)
                buf = Format(suma7, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = Format(suma8, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = Format(suma9, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 2, 0)
                nlineas

            End If
   
            buf = "" & mytablex.Fields("aceptante")
            found = formateaa(buf, 11, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = busca_nombre(buf)
            found = formateaa(buf, 60, 2, 0)
            nlineas
            Tmp = "" & mytablex.Fields("aceptante")
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0
            suma9 = 0

        End If

        buf = "" & mytablex.Fields("letra")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("girador")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("fechai")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("fechaf")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("moneda")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("banco")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("seccion")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Saldo")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("abono")
        buf = Format(Val(buf), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        'sdx = Val("" & mytablex.Fields("importe")) + Val("" & mytablex.Fields("interes1")) + Val("" & mytablex.Fields("interes2")) + Val("" & mytablex.Fields("protesto")) + Val("" & mytablex.Fields("otros"))
        buf = "" & mytablex.Fields("importe")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Nrounico")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("ochodia")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = busca_proveedor("" & mytablex.Fields("negociado"))
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = ""

        If "" & mytablex.Fields("estado") = "1" Then
            buf = "R"

        End If

        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = ""

        If "" & mytablex.Fields("estadoP") = "1" Then
            buf = "P"

        End If

        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 2, 0)
   
        'If "" & mytablex.Fields("saldo") <> 1 Then
        If "" & mytablex.Fields("estadoP") = "1" Then
            suma7 = suma7 + Val("" & mytablex.Fields("saldo"))
            suma8 = suma8 + Val("" & mytablex.Fields("abono"))
            suma9 = suma9 + Val("" & mytablex.Fields("importe"))
            ssuma7 = ssuma7 + Val("" & mytablex.Fields("saldo"))
            ssuma8 = ssuma8 + Val("" & mytablex.Fields("abono"))
            ssuma9 = ssuma9 + Val("" & mytablex.Fields("importe"))

        End If

        If "" & mytablex.Fields("estadoP") = "1" Then
            suma4 = suma4 + Val("" & mytablex.Fields("saldo"))
            suma5 = suma5 + Val("" & mytablex.Fields("abono"))
            suma6 = suma6 + Val("" & mytablex.Fields("importe"))
            ssuma4 = ssuma4 + Val("" & mytablex.Fields("saldo"))
            ssuma5 = ssuma5 + Val("" & mytablex.Fields("abono"))
            ssuma6 = ssuma6 + Val("" & mytablex.Fields("importe"))

        End If

        If "" & mytablex.Fields("estado") <> "1" Then
            suma1 = suma1 + Val("" & mytablex.Fields("saldo"))
            suma2 = suma2 + Val("" & mytablex.Fields("abono"))
            suma3 = suma3 + Val("" & mytablex.Fields("importe"))
            ssuma1 = ssuma1 + Val("" & mytablex.Fields("saldo"))
            ssuma2 = ssuma2 + Val("" & mytablex.Fields("abono"))
            ssuma3 = ssuma3 + Val("" & mytablex.Fields("importe"))

        End If
   
        mytablex.MoveNext
    Loop
    'found = formateaa("", 62, 0, 0)
    found = formateaa("Subneto ", 56, 0, 1)
    buf = Format(suma1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma2, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma3, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    nlineas

    If suma4 > 0 Then
        found = formateaa("P ", 56, 0, 1)
        buf = Format(suma4, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma5, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma6, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas

    End If

    If suma7 > 0 Then
        found = formateaa("R ", 56, 0, 1)
        buf = Format(suma7, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma8, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma9, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas

    End If
   
    found = formateaa("Total ", 56, 0, 1)
    buf = Format(ssuma1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(ssuma2, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(ssuma3, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
   
End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > Val(nrolineas) Then
        cabecera_letra

    End If

End Sub

Function busca_nombre(buf As String) As String

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("clientes")

    If acu = "P" Then
        Set mytablex = mydbxglo.OpenTable("proveedo")

    End If

    mytablex.Index = "codigo"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then
        busca_nombre = "" & mytablex.Fields("nombre")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_proveedor(buf As String) As String

    Dim mytablex As Table

    If Len(buf) = 0 Then Exit Function

    Set mytablex = mydbxglo.OpenTable("proveedo")
    mytablex.Index = "codigo"
    mytablex.Seek "=", buf

    If Not mytablex.NoMatch Then
        busca_proveedor = "" & mytablex.Fields("nombre")

    End If

    '------------------------------------- ------------
    mytablex.Close
 
End Function

Sub cabecera_voucher()

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
    
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    found = formateaa("Aceptante", 12, 2, 0)
    found = formateaa("Letra", 12, 0, 0)
    found = formateaa("Girador", 12, 0, 0)
    found = formateaa("Fechai", 11, 0, 0)
    found = formateaa("Fechaf", 11, 0, 0)
    found = formateaa("M", 2, 0, 0)
    found = formateaa("Banco", 7, 0, 0)
    found = formateaa("Seccio", 7, 0, 0)
    found = formateaa("Saldo ", 11, 0, 1)
    found = formateaa("Abono ", 11, 0, 1)
    found = formateaa("Total ", 11, 2, 1)
    
    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)

End Sub

