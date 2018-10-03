VERSION 5.00
Begin VB.Form tsisper 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Acceso Personal"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   13860
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   9480
      ScaleHeight     =   4155
      ScaleWidth      =   4155
      TabIndex        =   21
      Top             =   480
      Width           =   4215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H008080FF&
      Caption         =   "Mensajes"
      Height          =   2535
      Left            =   8880
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   6735
      Begin VB.CommandButton Command1 
         Caption         =   "Cerrar"
         Height          =   495
         Left            =   2520
         TabIndex        =   20
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label mensaje 
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   720
         TabIndex        =   19
         Top             =   720
         Width           =   5415
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   8040
      Top             =   1920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Ultimo Acceso Procesado"
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   9255
      Begin VB.Label xhora 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label xfecha 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label nombre 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   960
         Width           =   5655
      End
      Begin VB.Label codigo 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label numero 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hora"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Usuario"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Transaccion"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox clave 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   33
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      IMEMode         =   3  'DISABLE
      Left            =   4080
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Foto"
      Height          =   375
      Left            =   9480
      TabIndex        =   22
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label tipo 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   8040
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label hora 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   6840
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label fecha 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   INGRESE CLAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1440
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
   Begin VB.Menu dkj89232 
      Caption         =   "&Visualizar"
   End
   Begin VB.Menu lfo8923 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tsisper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clave_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
found = busca_vendedor()
If found = 0 Then
   mensaje = "No existe Funcionario"
   Frame2.Visible = True
   codigo = ""
   nombre = ""
   numero = ""
   xfecha = ""
   xhora = ""
   clave = ""
   Command1.SetFocus
   Exit Sub
End If
found = graba_transaccion()
If found = 0 Then
   mensaje = "Error al grabar Funcionario"
   Frame2.Visible = True
   codigo = ""
   nombre = ""
   numero = ""
   xfecha = ""
   xhora = ""
   clave = ""
   Command1.SetFocus
   Exit Sub
End If
   xfecha = Format(Now, "dd/mm/yyyy")
   xhora = Format(Now, "hh:mm:ss")
   clave = ""
   If tipo = "E" Then
   mensaje = "ENTRADA.Satifactoria " & nombre
   End If
   If tipo = "S" Then
   mensaje = "SALIDA.Satifactoria " & nombre
   End If
   Frame2.Visible = True
   Command1.SetFocus
End Sub

Private Sub Command1_Click()
clave = ""
clave.SetFocus
lfo8923_Click
End Sub

Private Sub dkj89232_Click()
segPerso.Show 1
End Sub

Private Sub lfo8923_Click()
If Frame2.Visible = True Then
   Frame2.Visible = False
   Exit Sub
End If
tsisper.Hide
Unload tsisper
End Sub

Private Sub Timer1_Timer()
fecha = Format(Now, "dd/mm/yyyy")
hora = Format(Now, "hh:mm:ss")
End Sub
Function graba_transaccion()
On Error GoTo cmd35_err
Dim mydbx As Database
Dim mytablex As Table
xfecha = ""
xhora = ""
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("sisper")
mytablex.Index = "sisper"
mytablex.Seek "=", fecha, codigo
If mytablex.NoMatch Then
   mytablex.AddNew
   mytablex.Fields("codigo") = "" & codigo
   mytablex.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
   mytablex.Fields("fechag") = Format(fecha, "dd/mm/yyyy")
   
   If Len("" & mytablex.Fields("ef1")) = 0 Then
      mytablex.Fields("ef1") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh1") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af1
   End If
   If Len("" & mytablex.Fields("sf1")) = 0 Then
      mytablex.Fields("sf1") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh1") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af1
   End If
   If Len("" & mytablex.Fields("ef2")) = 0 Then
      mytablex.Fields("ef2") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh2") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af1
   End If
   If Len("" & mytablex.Fields("sf2")) = 0 Then
      mytablex.Fields("sf2") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh2") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af1
   End If
   If Len("" & mytablex.Fields("ef3")) = 0 Then
      mytablex.Fields("ef3") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh3") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af1
   End If
   If Len("" & mytablex.Fields("sf3")) = 0 Then
      mytablex.Fields("sf3") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh3") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af1
   End If
   If Len("" & mytablex.Fields("ef4")) = 0 Then
      mytablex.Fields("ef4") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh4") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af1
   End If
   If Len("" & mytablex.Fields("sf4")) = 0 Then
      mytablex.Fields("sf4") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh4") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af1
   End If
   If Len("" & mytablex.Fields("ef5")) = 0 Then
      mytablex.Fields("ef5") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh5") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af1
   End If
   If Len("" & mytablex.Fields("sf5")) = 0 Then
      mytablex.Fields("sf5") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh5") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af1
   End If
   If Len("" & mytablex.Fields("ef6")) = 0 Then
      mytablex.Fields("ef6") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh6") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af1
   End If
   If Len("" & mytablex.Fields("sf6")) = 0 Then
      mytablex.Fields("sf6") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh6") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af1
   End If
   If Len("" & mytablex.Fields("ef7")) = 0 Then
      mytablex.Fields("ef7") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh7") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af1
   End If
   If Len("" & mytablex.Fields("sf7")) = 0 Then
      mytablex.Fields("sf7") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh7") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af1
   End If
   If Len("" & mytablex.Fields("ef8")) = 0 Then
      mytablex.Fields("ef8") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh8") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af1
   End If
   If Len("" & mytablex.Fields("sf8")) = 0 Then
      mytablex.Fields("sf8") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh8") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af1
   End If
af1:
   mytablex.Update
End If
If Not mytablex.NoMatch Then
   mytablex.Edit
   If Len("" & mytablex.Fields("ef1")) = 0 Then
      mytablex.Fields("ef1") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh1") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af2
   End If
   If Len("" & mytablex.Fields("sf1")) = 0 Then
      mytablex.Fields("sf1") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh1") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af2
   End If
   If Len("" & mytablex.Fields("ef2")) = 0 Then
      mytablex.Fields("ef2") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh2") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af2
   End If
   If Len("" & mytablex.Fields("sf2")) = 0 Then
      mytablex.Fields("sf2") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh2") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af2
   End If
   If Len("" & mytablex.Fields("ef3")) = 0 Then
      mytablex.Fields("ef3") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh3") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af2
   End If
   If Len("" & mytablex.Fields("sf3")) = 0 Then
      mytablex.Fields("sf3") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh3") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af2
   End If
   If Len("" & mytablex.Fields("ef4")) = 0 Then
      mytablex.Fields("ef4") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh4") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af2
   End If
   If Len("" & mytablex.Fields("sf4")) = 0 Then
      mytablex.Fields("sf4") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh4") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af2
   End If
   If Len("" & mytablex.Fields("ef5")) = 0 Then
      mytablex.Fields("ef5") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh5") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af2
   End If
   If Len("" & mytablex.Fields("sf5")) = 0 Then
      mytablex.Fields("sf5") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh5") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af2
   End If
   If Len("" & mytablex.Fields("ef6")) = 0 Then
      mytablex.Fields("ef6") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh6") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af2
   End If
   If Len("" & mytablex.Fields("sf6")) = 0 Then
      mytablex.Fields("sf6") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh6") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af2
   End If
   If Len("" & mytablex.Fields("ef7")) = 0 Then
      mytablex.Fields("ef7") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh7") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af2
   End If
   If Len("" & mytablex.Fields("sf7")) = 0 Then
      mytablex.Fields("sf7") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh7") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af2
   End If
   If Len("" & mytablex.Fields("ef8")) = 0 Then
      mytablex.Fields("ef8") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("eh8") = Format(Now, "hh:mm:ss")
      tipo = "E"
      GoTo af2
   End If
   If Len("" & mytablex.Fields("sf8")) = 0 Then
      mytablex.Fields("sf8") = Format(Now, "dd/mm/yyyy")
      mytablex.Fields("sh8") = Format(Now, "hh:mm:ss")
      tipo = "S"
      GoTo af2
   End If
af2:
   mytablex.Update
End If

mytablex.Close
mydbx.Close
If tipo = "E" Then
   numero = "ENTRADA"
End If
If tipo = "S" Then
   numero = "SALIDA"
End If
graba_transaccion = 1
Exit Function
cmd35_err:
MsgBox "--Error " + error$
mytablex.Close
mydbx.Close
Exit Function
End Function
Function busca_vendedor()
Dim mydbx As Database
Dim mytablex As Table
codigo = ""
nombre = ""
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("vendedor")
mytablex.Index = "clavere"
mytablex.Seek "=", clave
If Not mytablex.NoMatch Then
   codigo = "" & mytablex.Fields("codigo")
   nombre = "" & mytablex.Fields("nombre")
   busca_vendedor = 1
End If

End Function
