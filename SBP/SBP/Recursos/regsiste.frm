VERSION 5.00
Begin VB.Form regsiste 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar Sistema"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   13290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Generacion de Codigo"
      Height          =   5295
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   11535
      Begin VB.TextBox llave 
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox generar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   2040
         TabIndex        =   17
         Top             =   840
         Width           =   7935
      End
      Begin VB.TextBox generado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   60
         TabIndex        =   16
         Top             =   1440
         Width           =   7935
      End
      Begin VB.TextBox nombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   2040
         TabIndex        =   15
         Top             =   2040
         Width           =   7935
      End
      Begin VB.TextBox empresagen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   2040
         TabIndex        =   14
         Top             =   2640
         Width           =   7935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   4680
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LLave de Paso"
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo generar"
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave generado"
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Empresa"
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Empresa Generado"
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   2640
         Width           =   1935
      End
   End
   Begin VB.TextBox rempresa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      MaxLength       =   50
      TabIndex        =   10
      Top             =   3360
      Width           =   6135
   End
   Begin VB.TextBox copiavasc 
      Height          =   375
      Left            =   3480
      MaxLength       =   100
      TabIndex        =   6
      Top             =   1560
      Width           =   7815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   855
      Left            =   8640
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   855
      Left            =   6720
      TabIndex        =   3
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox answer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3480
      MaxLength       =   50
      TabIndex        =   2
      Top             =   2880
      Width           =   6135
   End
   Begin VB.Label xlicencia 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   4440
      Width           =   45
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   11
      Top             =   3360
      Width           =   3210
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DATOS RECIBIDOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      Height          =   1695
      Left            =   120
      Top             =   2640
      Width           =   11415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DATOS DE ENVIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   4200
      TabIndex        =   8
      Top             =   360
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   1695
      Left            =   120
      Top             =   600
      Width           =   11415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Copia Codigo a Enviar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   3300
   End
   Begin VB.Label vasc 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1080
      Width           =   7815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo Recibido-1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   2880
      Width           =   3210
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo a Enviar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3330
   End
   Begin VB.Menu ldfo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "regsiste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub answer_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H7A Then
        generar = vasc
        Frame1.Visible = True
        llave.SetFocus

    End If

End Sub

Private Sub Command1_Click()

    Dim DefVal, Msg, Title   ' Declare variables.

    Dim mytablex        As New ADODB.Recordset

    Dim cam             As String

    Dim pass            As String

    Dim buf             As String

    Dim I               As Integer

    Dim buf2            As String

    Dim textocodificado As String

    Dim ED              As tcrypto

    Set ED = New tcrypto

    On Error GoTo cmdyu_err

    If Len(Trim(answer)) = 0 Then
        MsgBox "Ingrese Codigo Recibido ", 48, "Aviso"
        Exit Sub

    End If

    If Len(Trim(rempresa)) = 0 Then
        MsgBox "Ingrese Nombre Empresa ", 48, "Aviso"
        Exit Sub

    End If

    cam = serie_disco_duro()
    buf2 = cam

    If xlicencia = "LICENCIA" Then
        pass = "hola"

    End If

    If xlicencia = "LICENCIACENTRALIZADO" Then
        pass = "vicky"

    End If

    'MsgBox xlicencia
    buf = crypt(pass, cam)
    vasc = ""

    For I = 1 To Len(buf)
        vasc = vasc & Asc(Mid$(buf, I, 1))
    Next I

    'MsgBox answer & " " & buf2
    If answer <> buf2 Then
        answer = ""
        MsgBox "Codigo de registro Invalido", 24, "Aviso"
        answer.SetFocus
        Exit Sub

    End If

    'MsgBox globalpath
amk:
    'MsgBox xlicencia
    mytablex.Open "SELECT * FROM  " & xlicencia, cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        mytablex.Close
        cn.Execute ("delete from " & xlicencia)
        GoTo amk

    End If
   
    'MsgBox answer
    mytablex.AddNew
    textocodificado = ED.Encrypt(answer, "FABIANITA", False)
    'MsgBox textocodificado
    mytablex.Fields("serie") = textocodificado
    'textocodificado = ED.Encrypt(textocodificado, "FABIANITA", True)
    'MsgBox textocodificado
    textocodificado = ED.Encrypt(rempresa, "FABIANITA", False)
    mytablex.Fields("nombre") = textocodificado
    mytablex.Update
   
    'Open globalpath & "\serie.txt" For Output As #1
    'Print #1, answer
    'Close #1
    MsgBox "REGISTRO EXITOSO !!!!!", 48, "Aviso"
    ldfo33_Click
    Exit Sub
cmdyu_err:
    MsgBox "Aviso en xxxLixx " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command2_Click()
    ldfo33_Click

End Sub

Private Sub Command3_Click()

    Dim abuf, pass, buf As String

    Dim I    As Integer

    Dim xkey As String

    xkey = "V" + "I" + "S" + "I" + "T" + "E" + "C" + "A" + "A" + "B" + "C"

    If Len(generar) = 0 Then Exit Sub
    'If Len(nombre) = 0 Then Exit Sub

    If llave <> xkey Then Exit Sub
    abuf = ""

    For I = 1 To Len(generar) Step 2
        abuf = abuf & Chr$(Val(Mid$(generar, I, 2)))
        'MsgBox Mid$(generar, i, 2)
    Next I

    If xlicencia = "LICENCIA" Then
        pass = "hola"

    End If

    If xlicencia = "LICENCIACENTRALIZADO" Then
        pass = "vicky"

    End If

    'pass = "hola"
    buf = crypt(pass, abuf)
    generado = buf
      
    cam = serie_disco_duro()
    buf2 = cam

    If xlicencia = "LICENCIA" Then
        pass = "hola"

    End If

    If xlicencia = "LICENCIACENTRALIZADO" Then
        pass = "vicky"

    End If

    'pass = "hola"
    buf = crypt(pass, cam)
      
    abuf = ""

    If xlicencia = "LICENCIA" Then
        pass = "hola"

    End If

    If xlicencia = "LICENCIACENTRALIZADO" Then
        pass = "vicky"

    End If

    'pass = "hola"
    buf = crypt(pass, nombre)
      
    For I = 1 To Len(buf)
        abuf = abuf & Asc(Mid$(buf, I, 1))
    Next I

    empresagen = abuf
    abuf = ""
      
    For I = 1 To Len(empresagen) Step 2
        abuf = abuf & Chr$(Val(Mid$(empresagen, I, 2)))
        'MsgBox Mid$(generar, i, 2)
    Next I

    If xlicencia = "LICENCIA" Then
        pass = "hola"

    End If

    If xlicencia = "LICENCIACENTRALIZADO" Then
        pass = "vicky"

    End If

    'pass = "hola"
    buf = crypt(pass, abuf)
    empresagen = buf

End Sub

Private Sub Form_Activate()

    Dim cam  As String

    Dim pass As String

    Dim buf  As String

    Dim I    As Integer

    Dim buf2 As String

    On Error GoTo cmd9000_err

    'MsgBox "abc"
    If Len(Trim(xlicencia)) = 0 Then
        Exit Sub

    End If

    cam = serie_disco_duro()
    'MsgBox cam
    buf2 = cam

    If xlicencia = "LICENCIA" Then
        pass = "hola"

    End If

    If xlicencia = "LICENCIACENTRALIZADO" Then
        pass = "vicky"

    End If

    'pass = "hola"
    buf = crypt(pass, cam)
    vasc = ""

    For I = 1 To Len(buf)
        vasc = vasc & Asc(Mid$(buf, I, 1))
    Next I

    copiavasc = vasc
    Exit Sub
cmd9000_err:
    MsgBox "Aviso en activi " + error$, 48, "Aviso"
    End
    Exit Sub

End Sub

Private Sub ldfo33_Click()
    regsiste.Hide
    Unload regsiste

End Sub
