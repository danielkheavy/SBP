VERSION 5.00
Begin VB.Form tflujoe 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flujo de Entradas Salidas Documentos"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      TabIndex        =   18
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
      Left            =   2160
      MaxLength       =   30
      TabIndex        =   17
      Text            =   "45"
      Top             =   5160
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
      Left            =   2160
      MaxLength       =   30
      TabIndex        =   16
      Top             =   4800
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
      TabIndex        =   15
      Top             =   2880
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
      TabIndex        =   14
      Top             =   2520
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
      TabIndex        =   13
      Text            =   "%"
      Top             =   720
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
      TabIndex        =   12
      Top             =   360
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
      TabIndex        =   11
      Text            =   "%"
      Top             =   1080
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
      Text            =   "%"
      Top             =   1560
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
      Left            =   2160
      MaxLength       =   11
      TabIndex        =   9
      Text            =   "%"
      Top             =   3840
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
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1440
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
      TabIndex        =   7
      Text            =   "%"
      Top             =   1920
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
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3840
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
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4200
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
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.ComboBox horai 
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
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.ComboBox horaf 
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
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2880
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
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.ComboBox servicio 
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
      Top             =   4200
      Width           =   1575
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
      TabIndex        =   38
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
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
      Left            =   0
      TabIndex        =   37
      Top             =   5160
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
      Left            =   0
      TabIndex        =   36
      Top             =   4800
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
      TabIndex        =   35
      Top             =   2880
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
      TabIndex        =   34
      Top             =   2520
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
      TabIndex        =   33
      Top             =   360
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
      TabIndex        =   32
      Top             =   720
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
      TabIndex        =   31
      Top             =   1080
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
      TabIndex        =   30
      Top             =   1560
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
      TabIndex        =   29
      Top             =   3480
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
      TabIndex        =   28
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
      TabIndex        =   27
      Top             =   1440
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
      TabIndex        =   26
      Top             =   1920
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
      Left            =   3840
      TabIndex        =   25
      Top             =   3840
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
      Left            =   3840
      TabIndex        =   24
      Top             =   4200
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
      Left            =   3840
      TabIndex        =   23
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label23 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HoraInicio"
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
      TabIndex        =   22
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label24 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HoraFinal"
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
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label25 
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
      TabIndex        =   20
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
      TabIndex        =   19
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Menu Ki8912 
      Caption         =   "&Ejecuta"
   End
   Begin VB.Menu dl8923 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tflujoe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim xmeses(13)  As Double

Dim xmeses1(13) As Double

Private Sub dl8923_Click()
    tflujoe.Hide
    Unload tflujoe

End Sub

Private Sub Form_Activate()

    Dim mytablex As New ADODB.Recordset

    tipo.Clear
    tipo.AddItem "%"

    mytablex.Open "select * from tipo", cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        'If "" & mytablex.Fields("grupo") = acu Then
        tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")
        'End If
        mytablex.MoveNext
    Loop
    mytablex.Close

    tipo.ListIndex = 0
    caja.Clear
    caja.AddItem "%"
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
    turno.AddItem "%"
    mytablex.Open "select * from turno", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        turno.AddItem "" & mytablex.Fields("turno") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    turno.ListIndex = 0

    cajero.Clear
    cajero.AddItem "%"
    mytablex.Open "select * from vendedor", cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    cajero.ListIndex = 0
 
    local1.Clear
    local1.AddItem "%"
    mytablex.Open "select * from tlocal", cn, adOpenStatic, adLockOptimistic

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

Private Sub Form_Load()

    Dim mytablex As New ADODB.Recordset

    Dim I        As Integer

    servicio.Clear
    servicio.AddItem "%"

    mytablex.Open "SELECT * FROM servicio ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        servicio.AddItem "" & mytablex.Fields("servicio") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    servicio.ListIndex = 0
    mytablex.Close

    horai.Clear
    horai.AddItem "%"
    horaf.AddItem "%"

    For I = 0 To 23
        horai.AddItem Format(I, "00")
        horaf.AddItem Format(I, "00")
    Next I

    horai.ListIndex = 0
    horaf.ListIndex = 0

    fechaf = Format(Now, "dd/mm/yyyy")
    fechai = Format(Now, "dd/mm/yyyy") '"01" & "/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")

    estado.AddItem "%"
    estado.AddItem "2"
    estado.AddItem "1"
    estado.AddItem "0"
    estado.ListIndex = 1

    moneda.Clear
    moneda.AddItem "%"
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0

End Sub

Function sql_documento_meses1(mytablex As ADODB.Recordset)

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    buf = "select Tipo,month(fecha) as xmes,sum(total) as Tot from factura where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & extra_loquesea("" & tipo) & "'"

    End If

    If serie <> "%" Then
        buf = buf & " and serie like '" & serie & "'"

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

    If caja <> "%" Then
        buf = buf & " and caja like '" & extra_loquesea(caja) & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & extra_loquesea(turno) & "'"

    End If

    If horai <> "%" And horaf <> "%" Then
        If Val(horaf) >= Val(horai) Then
            buf = buf & " and hour(hora)>=" & Val(horai)
            buf = buf & " and hour(hora)<=" & Val(horaf)

        End If

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & extra_loquesea(cajero) & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & extra_loquesea(vendedor) & "'"

    End If

    buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P'"
    buf = buf & " OR acu='A' or acu='B' or acu='C' or acu='D' or acu='G')"

    If estado <> "%" Then
        buf = buf & " and estado like '" & estado & "'"

    End If

    buf = buf & " group by Tipo,fecha order by tipo,month(fecha)"
    'MsgBox buf
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sql_documento_meses1 = 1

    End If

End Function

Function sql_documento_meses2(mytablex As ADODB.Recordset)

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    buf = "select Tipo,month(fecha) as xmes,sum(total) as Tot from recibo where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & extra_loquesea("" & tipo) & "'"

    End If

    If serie <> "%" Then
        buf = buf & " and serie like '" & serie & "'"

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

    If caja <> "%" Then
        buf = buf & " and caja like '" & extra_loquesea(caja) & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & extra_loquesea(turno) & "'"

    End If

    If horai <> "%" And horaf <> "%" Then
        If Val(horaf) >= Val(horai) Then
            buf = buf & " and hour(hora)>=" & Val(horai)
            buf = buf & " and hour(hora)<=" & Val(horaf)

        End If

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & extra_loquesea(cajero) & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & extra_loquesea(vendedor) & "'"

    End If

    buf = buf & " and (acu='V' or acu='W' ) "

    If estado <> "%" Then
        buf = buf & " and estado like '" & estado & "'"

    End If

    buf = buf & " group by Tipo,month(fecha) "
    'MsgBox buf
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sql_documento_meses2 = 1

    End If

End Function

Private Sub Ki8912_Click()
    proceso_impresion

End Sub

Sub procesa_compra_venta(mytablex As ADODB.Recordset)

    Dim buf   As String

    Dim sw    As Integer

    Dim Tmp   As String

    Dim found As Integer

    Dim I     As Integer

    Dim sdx1  As Double

    Dim sww   As Integer

    sw = 0
    Do

        If mytablex.EOF Then Exit Do
        If sw = 0 Then
            sw = 1
            buf = busca_tipo("" & mytablex.Fields("Tipo"))
            found = formateaa(buf, 20, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mytablex.Fields("Tipo")

        End If

        If Tmp <> "" & mytablex.Fields("Tipo") Then
            sdx1 = 0

            For I = 1 To 12
                buf = Format(xmeses(I), "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                sdx1 = sdx1 + xmeses(I)
            Next I

            buf = Format(sdx1, "0.00")
            found = formateaa(buf, 10, 0, 1)

            For I = 1 To 12
                xmeses(I) = 0#
            Next I

            found = formateaa("", 1, 2, 0)
            nlineas
   
            buf = busca_tipo("" & mytablex.Fields("Tipo"))
            found = formateaa(buf, 20, 0, 0)
            found = formateaa("", 1, 0, 0)
            Tmp = "" & mytablex.Fields("Tipo")

        End If

        sww = busca_tipo1("" & mytablex.Fields("tipo"))
        xmeses(CInt("" & mytablex.Fields("xmes"))) = xmeses(CInt("" & mytablex.Fields("xmes"))) + sww * Val("" & mytablex.Fields("tot"))
        xmeses1(CInt("" & mytablex.Fields("xmes"))) = xmeses1(CInt("" & mytablex.Fields("xmes"))) + sww * Val("" & mytablex.Fields("tot"))
        mytablex.MoveNext
    Loop
    'mytablex.Close
    sdx1 = 0

    For I = 1 To 12
        buf = Format(xmeses(I), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx1 = sdx1 + xmeses(I)
    Next I

    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    nlineas

    'found = formateaa("", 21, 0, 0)
    'sdx1 = 0
    'For i = 1 To 12
    '    buf = Format(xmeses1(i), "0.00")
    '    found = formateaa(buf, 10, 0, 1)
    '    found = formateaa("", 1, 0, 0)
    '    sdx1 = sdx1 + xmeses1(i)
    'Next i
    'buf = Format(sdx1, "0.00")
    'found = formateaa(buf, 10, 0, 1)
    'found = formateaa("", 1, 2, 0)

End Sub

Sub proceso_impresion()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim I        As Integer

    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    contpag = 0

    For I = 1 To 12
        xmeses(I) = 0
        xmeses1(I) = 0
    Next I

    found = sql_documento_meses1(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento_meses1
    procesa_compra_venta mytablex
    mytablex.Close
    found = sql_documento_meses2(mytablex)

    If found = 1 Then
        procesa_compra_venta mytablex

    End If

    found = formateaa("", 21, 0, 0)
    sdx1 = 0

    For I = 1 To 12
        buf = Format(xmeses1(I), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx1 = sdx1 + xmeses1(I)
    Next I

    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Function busca_tipo(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from tipo where tipo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_tipo = "" & mytablex.Fields("descripcio")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_tipo1(buf As String)

    Dim sw       As Integer

    Dim mytablex As New ADODB.Recordset

    sw = 1
    mytablex.Open "select * from tipo where tipo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        Select Case "" & mytablex.Fields("tipodoc")

            Case "A", "B", "C", "D", "G"
                sw = 1

            Case "J", "K", "L", "M", "P"
                sw = -1

            Case "W"
                sw = 1

            Case "V"
                sw = -1

        End Select

    End If

    '------------------------------------- ------------
    mytablex.Close
    busca_tipo1 = sw

End Function

Sub nlineas()
    contlin = contlin + 1

    If contlin > Val(nrolineas) Then
        cabecera_documento_meses1

    End If

End Sub

Sub cabecera_documento_meses1()

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
    buf = String(164, "-")
    found = formateaa(buf, 164, 2, 0)
    found = formateaa("Descripcio ", 20, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Enero", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Febrero", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Marzo", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Abril", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Mayo", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Junio", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Julio", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Agosto", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Setiembre", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Octubre", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Noviembre", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Diciembre", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Total", 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    
    buf = String(164, "-")
    found = formateaa(buf, 164, 2, 0)

End Sub

