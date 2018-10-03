VERSION 5.00
Begin VB.Form menucaja 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso Caja registradora"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Cambio de Paridad"
      Height          =   6015
      Left            =   3480
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox compra 
         Height          =   615
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   20
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox venta 
         Height          =   615
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   19
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "menucaja.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Salir"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "menucaja.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Grabar registro"
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COMPRA"
         Height          =   615
         Left            =   360
         TabIndex        =   22
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "VENTA"
         Height          =   615
         Left            =   360
         TabIndex        =   21
         Top             =   1560
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FF80&
      Caption         =   "Mensajes del Sistema...."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   10095
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9120
         MaskColor       =   &H00E0E0E0&
         Picture         =   "menucaja.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Salir"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label mensaje_error 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.TextBox terminal 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MaxLength       =   3
      TabIndex        =   0
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox turno 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   7
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox cajero 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      MaxLength       =   5
      TabIndex        =   6
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton image10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Tipo/Cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      Picture         =   "menucaja.frx":3636
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Apertura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      Picture         =   "menucaja.frx":4ED8
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BorraTemporal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      Picture         =   "menucaja.frx":677A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   680
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   12300
      TabIndex        =   1
      Top             =   0
      Width           =   12360
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "menucaja.frx":8934
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Nota. El Numero de Caja/Terminal no debe estar en uso en otra Maquina !!!"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   5400
      Width           =   5535
   End
   Begin VB.Shape Shape2 
      Height          =   2175
      Left            =   240
      Top             =   3120
      Width           =   5895
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "CONTROL DE ACCESO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   5295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CAJA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TURNO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   9
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CAJERO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5175
      Left            =   6360
      Picture         =   "menucaja.frx":9B46
      Stretch         =   -1  'True
      Top             =   840
      Width           =   5820
   End
   Begin VB.Menu lo333 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "menucaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fechavta As String
Private Sub cmdExit_Click()

lo333_Click
End Sub

Private Sub Command1_Click()
tapertur.Show 1
End Sub

Private Sub Command3_Click()
Frame3.Visible = False
terminal.SetFocus
End Sub

Private Sub Form_Load()
cajero = dgusuario
End Sub

Private Sub lo333_Click()
If Frame3.Visible = True Then
   Frame3.Visible = False
   terminal.SetFocus
   Exit Sub
End If
menucaja.Hide
Unload menucaja

End Sub

Private Sub terminal_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
turno.SetFocus
End Sub

Private Sub turno_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If existe_apertura() = 0 Then
  buf = "LA CAJA NRO:" & terminal & Chr$(10) & Chr$(13)
   buf = buf & "TURNO          :" & turno & Chr$(10) & Chr$(13)
   buf = buf & "CAJERO         :" & cajero & Chr$(10) & Chr$(13)
   buf = buf & "NO SE ENCUENTRA APERTURADO                           " & Chr$(10) & Chr$(13)
   buf = buf & "PARA EL DIA " & Format(Now, "dd/mm/yyyy")
   mensaje_error = buf
   Frame3.Visible = True
   Command3.SetFocus
   Exit Sub
End If

End Sub
Function existe_apertura()
Dim rs1 As New ADODB.Recordset
Dim buf As String
fechavta = ""
buf = "SELECT fechai FROM apertura where cajero='" & Trim(cajero) & "' and caja='" & Trim(terminal) & "' and turno='" & Trim(turno) & "'"
buf = buf & " and fechai>='" & Format(Now, "YYYYMMDD") & "' and fechaf<='" & Format(Now, "YYYYMMDD") & "'"
'MsgBox buf
existe_apertura = 1
If rs1.State = 1 Then rs1.Close
   rs1.Open buf, cn, adOpenDynamic, adLockReadOnly
   If rs1.EOF Then
      existe_apertura = 0
      Else
      fechavta = rs1.Fields("fechai")
   End If
   rs1.Close
   Set rs1 = Nothing

End Function

