VERSION 5.00
Begin VB.Form tomafi 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TomaFisico"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2805
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   1920
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox xclave 
         Height          =   315
         Left            =   120
         MaxLength       =   60
         TabIndex        =   18
         Text            =   "12345"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox xusuario 
         Height          =   315
         Left            =   120
         MaxLength       =   60
         TabIndex        =   16
         Text            =   "sa"
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox xvservidor 
         Height          =   315
         Left            =   120
         MaxLength       =   60
         TabIndex        =   14
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Conectar"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Usuario"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Servidor"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Ingreso Codigo Barras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox cantidad 
         Height          =   375
         Left            =   840
         MaxLength       =   8
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox producto 
         Height          =   375
         Left            =   840
         MaxLength       =   15
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Graba"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acumulado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.Label cantidada 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Label conectado 
      BackColor       =   &H00FFFF80&
      Caption         =   "N"
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Conectar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BorrarTodo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TomaFisica"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "tomafi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim mytablex As New ADODB.Recordset
If Len(producto) = 0 Then
   producto.SetFocus
   Exit Sub
End If
If Not IsNumeric(cantidad) Then
   cantidad = ""
   cantidad.SetFocus
   Exit Sub
End If
 mytablex.Open "SELECT *  FROM pocket where producto='" & producto & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.AddNew
      mytablex.Fields("producto") = Trim(producto)
      mytablex.Fields("cantidad") = Val(cantidad) + Val(cantidada)
      mytablex.Update
      Else
      mytablex.Fields("cantidad") = Val(cantidad) + Val(cantidada)
      mytablex.Update
   End If
mytablex.Close
producto = ""
cantidada = ""
cantidad = ""
producto.SetFocus
End Sub

Private Sub Command2_Click()
Frame1.Visible = False
End Sub

Private Sub Form_Load()
globaldir = App.path & "\001d\06"
globaldir = "\ORION.V5\001D\06"
globalpath = "\ORION.V5"
carga_servidor
End Sub
Sub carga_servidor()
Dim found As Integer
Dim buf As String
On Error GoTo cmd169999_err
    vservidor = ""
    buf = ""
    If Dir$(globalpath & "\server.txt") <> "" Then
       Close
       Open globalpath & "\server.txt" For Input As #1
       Input #1, buf
       Close #1
       xvservidor = buf
    End If
    '------------------------------------------------
    Exit Sub
cmd169999_err:
   Close
   Exit Sub
End Sub
Public Function conectarpo()
Dim dbuser As String
Dim dbpassword As String
Dim dbname As String
Dim dbserver As String
On Error GoTo cmd1_error

 cn.CursorLocation = adUseClient
 cn.CommandTimeout = 1024
 cn.Open "Driver={SQL Server};Server=" & xvservidor & ";Database=calipso;Uid=" & xusuario & ";pwd=" & xclave
 'cn.Open "Driver={SQL Server};Server=" & vservidor & ";Database=calipso ;Uid=sa "
 conectarpo = 1
 conectado = "S"
 Exit Function
cmd1_error:
 MsgBox " " & error$, 48, "Aviso"
 Exit Function
 End Function

Private Sub Label10_Click()
Frame2.Visible = False
End Sub

Private Sub Label11_Click()
If conectado = "S" Then Exit Sub
Frame2.Visible = True
End Sub

Private Sub Label12_Click()
conectarpo
End Sub

Private Sub Label3_Click()
If conectado <> "S" Then Exit Sub
Codigo = ""
cantidada = ""
cantidad = ""
Frame1.Visible = True
producto.SetFocus
End Sub

Sub inicializa_codigo()
   cantidada = ""
   cantidad = ""

End Sub

Private Sub Label5_Click()
End
End Sub

Private Sub Label6_Click()
If conectado <> "S" Then Exit Sub
If MsgBox("Desea Borrar Todo ", 1, "Aviso") <> 13 Then Exit Sub
cn.Execute ("delete from pocket")
End Sub

Private Sub producto_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(producto) = 0 Then
   producto.SetFocus
   Exit Sub
End If
inicializa_codigo
found = busca_codigo()
cantidad.SetFocus
End Sub
Function busca_codigo()
Dim found As Integer
Dim mytablex As New ADODB.Recordset
 cantidada = ""
 mytablex.Open "SELECT *  FROM pocket where producto='" & producto & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      busca_codigo = 1
      cantidada = "" & mytablex.Fields("cantidad")
End If
mytablex.Close
End Function
