VERSION 5.00
Begin VB.Form tiplocal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Servidor De Datos"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   555
   ClientWidth     =   9015
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "Roboto"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   480
      TabIndex        =   7
      Top             =   5160
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CommandButton cmdDefault 
         Caption         =   "Default"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   25
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   24
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton cmdGuardar2 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   23
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox nombre 
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1440
         Width           =   3975
      End
      Begin VB.TextBox defecto 
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   6
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox clave 
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2400
         Width           =   3975
      End
      Begin VB.TextBox ip 
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2880
         Width           =   3975
      End
      Begin VB.TextBox local1 
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MaxLength       =   2
         TabIndex        =   1
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox base 
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1920
         Width           =   3975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Local:"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1215
         TabIndex        =   14
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Defecto:"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1830
         TabIndex        =   13
         Top             =   3360
         Width           =   840
      End
      Begin VB.Label id 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   0
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Id:"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2430
         TabIndex        =   12
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Clave:"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2040
         TabIndex        =   11
         Top             =   2400
         Width           =   630
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero Ip (local):"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   885
         TabIndex        =   10
         Top             =   2880
         Width           =   1785
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Local (01):"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1590
         TabIndex        =   9
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Base Datos:"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1455
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Roboto"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   480
      TabIndex        =   15
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "Borrar"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         TabIndex        =   22
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         TabIndex        =   21
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3570
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   6015
      End
      Begin VB.CommandButton cmdBorraBd 
         Caption         =   "BorraBd"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label nreg 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3240
         TabIndex        =   19
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblNumeroRegistros 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Numero Registros:"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1245
         TabIndex        =   18
         Top             =   360
         Width           =   1890
      End
   End
   Begin VB.Menu flo901 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tiplocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub base_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    clave.SetFocus

End Sub

Private Sub clave_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    ip.SetFocus

End Sub




Private Sub cmdBorraBd_Click()
 On Error GoTo cmd_8912

    Kill globalpath & "\config.txt"
    Exit Sub
cmd_8912:
    MsgBox "Aviso:Puede estar ya Borrado ", 48, "Aviso"
    Exit Sub
End Sub

Private Sub cmdBorrar_Click()
 Dim found As Integer

    If InputBox("Clave de Paso", "Aviso") = "KALIPO" Then
        found = borra_registro()

    End If

End Sub

Private Sub cmdCerrar_Click()
    Frame1.Visible = False
End Sub

Private Sub cmdDefault_Click()
    local1 = "01"
    nombre = "NOMBRE LOCAL"
    base = "CALIPSO"
    clave = ""
    ip = "(LOCAL)"
    defecto = ""

End Sub

Private Sub cmdGuardar2_Click()

    Dim found As Integer

    found = valida()

    If found = 0 Then
        MsgBox "Campos Invalidos ", 48, "Aviso"
        Exit Sub

    End If

    found = graba_registro()

    If found = 0 Then
        MsgBox "No se pudo Grabar ", 48, "Aviso"
        Exit Sub

    End If

    found = cargar()

    Frame1.Visible = False

End Sub

Private Sub cmdModificar_Click()
 Dim found As Integer

    found = lee_registro()

    If found = 0 Then
        Exit Sub

    End If
    Frame1.Top = 120
    Frame1.Caption = "MODIFICA"
    Frame1.Visible = True
    local1.SetFocus

End Sub

Private Sub cmdNuevo_Click()
    inicializa
    Frame1.Top = 120
    Frame1.Caption = "NUEVO"
    Frame1.Visible = True
    local1.SetFocus

End Sub






Private Sub Command3_Click()

    Dim found As Integer

    found = lee_registro()

    If found = 0 Then
        Exit Sub

    End If

    Frame1.Top = 0
    Frame1.Caption = "Modificar"
    Frame1.Visible = True
    local1.SetFocus

End Sub

Private Sub Command4_Click()
    local1 = "01"
    nombre = "NOMBRE LOCAL"
    base = "CALIPSO"
    clave = ""
    ip = "(LOCAL)"
    defecto = ""

End Sub

Private Sub Command5_Click()


End Sub

Private Sub Command6_Click()

    Dim found As Integer

    If InputBox("Clave de Paso", "Aviso") = "KALIPO" Then
        found = borra_registro()

    End If

End Sub

Private Sub Command7_Click()

    On Error GoTo cmd_8912

    Kill globalpath & "\config.txt"
    Exit Sub
cmd_8912:
    MsgBox "Aviso:Puede estar ya Borrado ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub flo901_Click()
    tiplocal.Hide
    Unload tiplocal

End Sub

Sub inicializa()
    nombre = ""
    local1 = ""
    base = ""
    clave = ""
    ip = ""
    defecto = ""

End Sub

Private Sub Form_Activate()

    Dim found As Integer

    found = cargar()

End Sub

Private Sub Form_Load()
    tiplocal.BackColor = RGB(91, 110, 128)
End Sub

Private Sub Label1_Click()
    base = "CALIPSO"

End Sub

Private Sub Label2_Click()
    local1 = "01"

End Sub

Private Sub Label3_Click()
    ip = "(LOCAL)"

End Sub



Private Sub local1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    base.SetFocus

End Sub

Function graba_registro()

    Dim bdip As ipmaquina

    On Error GoTo cmd7823_err

    Dim buf As String

    Dim sdx As Integer

    bdip.local1 = Trim(local1)
    bdip.base = Trim(base)
    bdip.clave = Trim(clave)
    bdip.ip = Trim(ip)
    bdip.nombre = Trim(nombre)
    bdip.defecto = Trim(defecto)
    sdx = 1
    Close
    Open globalpath & "\config.txt" For Random As #4 Len = Len(bdip)

    If Frame1.Caption = "NUEVO" Then
        sdx = (LOF(4) \ Len(bdip)) '
        sdx = sdx + 1

    End If

    If Frame1.Caption = "MODIFICA" Then
        sdx = 1

        If List1.ListIndex <> -1 Then
            sdx = List1.ListIndex + 1

        End If
   
    End If

    Put #4, sdx, bdip
    Close #4
    MsgBox "Proceso Grabado ", 48, "Aviso"
    graba_registro = 1
    Exit Function
cmd7823_err:
    MsgBox "Aviso en grabando " + error$, 48, "Aviso"
    Exit Function

End Function

Function lee_registro()

    Dim found As Integer

    Dim sdx   As Integer

    Dim bdip  As ipmaquina

    On Error GoTo cmd7824_err

    Dim buf As String

    sdx = 1

    If List1.ListIndex <> -1 Then
        sdx = List1.ListIndex + 1

    End If

    Close
    Open globalpath & "\config.txt" For Random As #4 Len = Len(bdip)
    Get #4, sdx, bdip
    local1 = Trim("" & bdip.local1)
    base = Trim("" & bdip.base)
    clave = Trim("" & bdip.clave)
    ip = Trim("" & bdip.ip)
    nombre = Trim("" & bdip.nombre)
    defecto = Trim("" & bdip.defecto)
    lee_registro = 1
    Close #4
    Exit Function
cmd7824_err:
    MsgBox "Aviso en Leer registro " + error$, 48, "Aviso"
    Exit Function

End Function

Function borra_registro()

    Dim found As Integer

    Dim bdip  As ipmaquina

    On Error GoTo cmd37825_err

    Dim buf As String

    Dim filetemp

    Dim j    As Integer

    Dim sdx  As Double

    Dim sdx1 As Double

    Dim I    As Integer

    Close
    sdx = 1
    sdx1 = 0

    If List1.ListIndex <> -1 Then
        sdx = List1.ListIndex + 1

    End If

    If existe_archivo(globalpath & "\temporal.tmp") > 0 Then
        Kill globalpath & "\Temporal.tmp"

    End If

    filetemp = FreeFile
    Open globalpath & "\temporal.tmp" For Random As #filetemp Len = Len(bdip)
    Open globalpath & "\config.txt" For Random As #4 Len = Len(bdip)
    sdx1 = (LOF(4) \ Len(bdip)) '
    j = 1

    For I = 1 To sdx1
        Get #4, I, bdip

        If sdx <> I Then
            Put #filetemp, j, bdip
            j = j + 1

        End If

    Next I

    Close #4
    Close #filetemp
    Close
    Kill globalpath & "\config.txt"
    FileCopy globalpath & "\Temporal.tmp", globalpath & "\config.txt"
    cargar
    Exit Function
cmd37825_err:
    MsgBox "Aviso en cargar registro " + error$, 48, "Aviso"
    Exit Function

End Function

Function cargar()

    Dim found As Integer

    Dim bdip  As ipmaquina

    On Error GoTo cmd7825_err

    Dim buf  As String

    Dim sdx  As Integer

    Dim ind  As Integer

    Dim sdx1 As Double

    Dim I    As Integer

    List1.Clear
    sdx1 = 0
    ind = 1
    Close
    Open globalpath & "\config.txt" For Random As #4 Len = Len(bdip)
    sdx = (LOF(4) \ Len(bdip)) '
    nreg = "" & sdx

    For I = 1 To sdx
        sdx1 = 1
        Get #4, I, bdip
        List1.AddItem Trim("" & bdip.local1) & "|" & Trim("" & bdip.nombre) & "|" & Trim("" & bdip.base) '& "|" & Trim("" & bdip.clave) & "|" & Trim("" & bdip.ip)
        cargar = 1
    Next I

    Close #4

    If sdx1 = 1 Then
        List1.ListIndex = 0

    End If

    Exit Function
cmd7825_err:
    MsgBox "Aviso en cargar registro " + error$, 48, "Aviso"
    Exit Function

End Function

Function valida()

    If Len(local1) = 0 Then
        local1.SetFocus
        Exit Function

    End If

    If Len(nombre) = 0 Then
        nombre.SetFocus
        Exit Function

    End If

    If Len(base) = 0 Then
        base.SetFocus
        Exit Function

    End If

    If Len(ip) = 0 Then
        ip.SetFocus
        Exit Function

    End If

    valida = 1

End Function


