VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tfamilia 
   BackColor       =   &H00FFFF00&
   Caption         =   "Tabla de Familias"
   ClientHeight    =   8580
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   13905
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   9360
      TabIndex        =   28
      Top             =   3480
      Visible         =   0   'False
      Width           =   11895
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Ejecutar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8280
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox buffer 
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6735
         Left            =   240
         TabIndex        =   32
         Top             =   1200
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   11880
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   22
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox TIPO 
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
      MaxLength       =   1
      TabIndex        =   27
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox margen5 
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
      TabIndex        =   25
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox margen4 
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
      TabIndex        =   23
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox margen3 
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
      TabIndex        =   21
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox margen2 
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
      TabIndex        =   19
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox margen1 
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
      TabIndex        =   15
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox maxdscto 
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
      TabIndex        =   14
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox red 
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
      MaxLength       =   15
      TabIndex        =   3
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox puerto 
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
      MaxLength       =   15
      TabIndex        =   2
      Top             =   2280
      Width           =   3255
   End
   Begin VB.TextBox descripcio 
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
      TabIndex        =   1
      Top             =   1320
      Width           =   5775
   End
   Begin VB.TextBox codigo 
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
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
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
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tfamilia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Grabar registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdAddEntry 
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
      Picture         =   "tfamilia.frx":1212
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Nuevo registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
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
      Left            =   2160
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tfamilia.frx":2424
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprimir"
      Top             =   0
      Width           =   735
   End
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
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tfamilia.frx":3636
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
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
      Left            =   720
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tfamilia.frx":4848
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Borrar registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   465
      Left            =   3720
      Picture         =   "tfamilia.frx":5A5A
      Stretch         =   -1  'True
      Top             =   840
      Width           =   600
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Margen Defecto5"
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
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Margen Defecto4"
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
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Margen Defecto3"
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
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Margen Defecto2"
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
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.ControlInventrario 1.NoControl"
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
      Left            =   3600
      TabIndex        =   18
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Margen Defecto1"
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
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Max Dscto"
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
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Puerto Red"
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
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Puerto Local"
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
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   11
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
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
      TabIndex        =   10
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
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
      TabIndex        =   9
      Top             =   840
      Width           =   2175
   End
   Begin VB.Menu ajdu1 
      Caption         =   "&Add"
   End
   Begin VB.Menu bo712 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu grba1 
      Caption         =   "&Grabar"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tfamilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ajdu1_Click()
If Frame1.Visible = True Then Exit Sub
inicializa
codigo = ""
codigo.SetFocus
End Sub

Private Sub bo712_Click()
Dim found As Integer
If Frame1.Visible = True Then Exit Sub
found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
found = borra_registro()
If found = 0 Then Exit Sub
MsgBox "Ok,Registro Borrado", 48, "Aviso"
codigo = ""
inicializa
codigo.SetFocus
End Sub




Private Sub buffer_DblClick()
Command1_Click
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   Frame1.Visible = False
   Frame1.Enabled = False
   codigo.SetFocus
   Exit Sub
End If
Command1_Click
End Sub

Private Sub cmdAddEntry_Click()
ajdu1_Click
End Sub

Private Sub cmdDelete_Click()
bo712_Click
End Sub

Private Sub cmdExit_Click()
dlo132_Click
End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdPrint_Click()
djuer1_Click
End Sub

Private Sub cmdSave_Click()
grba1_Click
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   dlo132_Click
   Exit Sub
End If
If Len(codigo) = 0 Then Exit Sub
found = busca_registro()
If found = 0 Then
   inicializa
End If
descripcio.SetFocus
End Sub

Private Sub codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   opcion1 = "1"
   consulta_datos
End If
End Sub
Sub consulta_datos()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT * FROM familia  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      rconsulta.Close
      Exit Sub
   End If
   

Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.ListIndex = 0
opcion1 = 1
Frame1.Enabled = True
Frame1.Visible = True
buffer = ""
buffer.SetFocus
Command1_Click


End Sub





Private Sub Command1_Click()
ejecuta 1
End Sub
Sub ejecuta(sw As Integer)
Dim rconsulta As New ADODB.Recordset
Dim cad As String
If opcion1 = 1 Then  'bodega
   If Len(buffer) = 0 Then
      cad = "SELECT Descripcio,Familia,Puerto,Red from familia    "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Descripcio,Familia,Puerto,Red from familia   where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      rconsulta.Close
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid1.DataSource = rconsulta
   dbGrid1.Columns(0).Width = 4000
   dbGrid1.Columns(1).Width = 2000
   dbGrid1.Columns(2).Width = 4000
   
   If sw = 1 Then
      dbGrid1.SetFocus
   End If
   Exit Sub
End If


End Sub


Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   codigo = dbGrid1.Columns(1)
   Frame1.Enabled = False
   Frame1.Visible = False
   codigo.SetFocus
   codigo_KeyPress 13
End If

End Sub


Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
Dim buf As String
Dim buf2 As String
If KeyAscii <> 13 And KeyAscii <> 27 Then
         If KeyAscii = 8 Then
            If Len(buffer) > 0 Then
               buf = Mid$(buffer, 1, Len(buffer) - 1)
               buffer = buf
               KeyAscii = 0
               Else
               KeyAscii = 0
               Exit Sub
            End If
         End If
         buf = Chr(KeyAscii)
         If Chr(KeyAscii) = "*" Then
            buf = ""
            buffer = buf
         End If
         If KeyAscii <> 13 Then
            buffer = buffer + buf
         End If
         buf = buffer
         ejecuta 0
End If

End Sub

Private Sub descripcio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
tipo.SetFocus

End Sub



Private Sub descripcio_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   codigo.SetFocus
   Exit Sub
End If
End Sub

Private Sub djuer1_Click()
If Frame1.Visible = True Then Exit Sub
reporgen.NAMETABLA = "familia"
reporgen.Show 1

End Sub

Private Sub dlo132_Click()
If Frame1.Visible = True Then
   Frame1.Visible = False
   Exit Sub
End If
tfamilia.Hide
Unload tfamilia
End Sub



Sub inicializa()
descripcio = ""
'TIPO.ListIndex = 0
puerto = ""
tipo = ""
red = ""
maxdscto = ""
margen1 = ""
margen2 = ""
margen3 = ""
margen4 = ""
margen5 = ""
End Sub
Function borra_registro()
On Error GoTo cmd56_err

cn.Execute ("DELETE   FROM FAMILIA WHERE familia='" & Trim(codigo) & "'")
borra_registro = 1
Exit Function
cmd56_err:
MsgBox "Aviso en borra " + error$, 48, "Aviso"
Exit Function



End Function
Function busca_registro()
Dim rsexiste As New ADODB.Recordset
   rsexiste.Open "SELECT * FROM familia where  familia='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount > 0 Then  'si existe
      pone_registro rsexiste
      busca_registro = 1
   End If

End Function
Sub pone_registro(mytablex As ADODB.Recordset)
codigo = "" & mytablex.Fields("familia")
descripcio = "" & mytablex.Fields("descripcio")
tipo = "" & mytablex.Fields("tipo")
puerto = "" & mytablex.Fields("puerto")
red = "" & mytablex.Fields("red")
margen1 = "" & mytablex.Fields("margen1")
margen2 = "" & mytablex.Fields("margen2")
margen3 = "" & mytablex.Fields("margen3")
margen4 = "" & mytablex.Fields("margen4")
margen5 = "" & mytablex.Fields("margen5")
maxdscto = "" & mytablex.Fields("maxdscto")
End Sub
Sub grabando(sw As Integer)

Dim cad As String


If sw = 0 Then
   cad = "INSERT INTO familia VALUES('" & Trim(codigo) & "','"
   cad = cad & Trim(descripcio) & "','"
   cad = cad & Trim(tipo) & "','"
   cad = cad & Trim(puerto) & "','"
   cad = cad & Trim(red) & "','"
   cad = cad & Trim(maxdscto) & "','"
   cad = cad & Trim(margen1) & "','"
   cad = cad & Trim(margen2) & "','"
   cad = cad & Trim(margen3) & "','"
   cad = cad & Trim(margen4) & "','"
   cad = cad & Trim(margen5) & "')"
   cn.Execute (cad)
   MsgBox "Adicion exitosa", 48, "Aviso"
   codigo.SetFocus
   Exit Sub
End If

If sw = 1 Then
   cad = "UPDATE familia SET "
   cad = cad & "familia = '" & Trim(codigo) & "'"
   cad = cad & ",descripcio='" & Trim(descripcio) & "'"
   cad = cad & ",tipo='" & Trim(tipo) & "'"
   cad = cad & ",puerto='" & Trim(puerto) & "'"
   cad = cad & ",red='" & Trim(red) & "'"
   cad = cad & ",maxdscto=" & Val(maxdscto) & ""
   cad = cad & ",margen1=" & Val(margen1) & ""
   cad = cad & ",margen2=" & Val(margen2) & ""
   cad = cad & ",margen3=" & Val(margen3) & ""
   cad = cad & ",margen4=" & Val(margen4) & ""
   cad = cad & ",margen5=" & Val(margen5) & ""
   cad = cad & " WHERE  familia='" & Trim(codigo) & "'"
   cn.Execute (cad)
   MsgBox "Rescripcion exitosa", 48, "Aviso"
   codigo.SetFocus
   Exit Sub
End If


End Sub

Private Sub grba1_Click()
Dim found As Integer
If Frame1.Visible = True Then Exit Sub
found = grabar()
If found = 0 Then Exit Sub
codigo.SetFocus

End Sub

Private Sub Image1_Click()
opcion1 = "1"
consulta_datos
End Sub




Function grabar()
Dim found As Integer
Dim cad As String
Dim rsexiste As New ADODB.Recordset


found = valida()
If found = 0 Then
   MsgBox "Campos invalidos", 48, "Aviso"
   Exit Function
End If

   rsexiste.Open "SELECT * FROM familia where  familia='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount > 0 Then  'si existe
      If MsgBox("Desea Reescribir? ", 1, "Aviso") <> 1 Then
         codigo.SetFocus
         Exit Function
      End If
      grabando 1
      Exit Function
   End If
   If MsgBox("Desea Adicionar ? ", 1, "Aviso") <> 1 Then
      codigo.SetFocus
      Exit Function
   End If
   grabando 0

End Function

Function valida()
If Len(codigo) = 0 Then
   codigo.SetFocus
   Exit Function
End If
If Len(descripcio) = 0 Then
   descripcio.SetFocus
   Exit Function
End If
valida = 1
End Function


Private Sub margen1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   red.SetFocus
   Exit Sub
End If

End Sub

Private Sub maxdscto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   'margen.SetFocus
   Exit Sub
End If

End Sub

Private Sub puerto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
red.SetFocus

End Sub

Private Sub puerto_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   tipo.SetFocus
   Exit Sub
End If

End Sub

Private Sub red_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
margen1.SetFocus

End Sub

Private Sub red_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   puerto.SetFocus
   Exit Sub
End If

End Sub

Private Sub tipo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
puerto.SetFocus
End Sub

Private Sub tipo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   descripcio.SetFocus
   Exit Sub
End If
End Sub
