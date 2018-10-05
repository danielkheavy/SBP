VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tficha 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla Clientes"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10530
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Consulta"
      Height          =   8415
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   10455
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
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
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
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
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
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbGrid2 
         Height          =   6855
         Left            =   240
         TabIndex        =   39
         Top             =   1200
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   12091
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   840
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   9135
      Begin VB.TextBox ruc 
         Height          =   495
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   40
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox correo 
         Height          =   495
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   30
         Top             =   5160
         Width           =   4575
      End
      Begin VB.TextBox telefonom 
         Height          =   495
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   28
         Top             =   4680
         Width           =   1575
      End
      Begin VB.TextBox anexo 
         Height          =   495
         Left            =   3840
         MaxLength       =   11
         TabIndex        =   26
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox telefonot 
         Height          =   495
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   24
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox telefonop 
         Height          =   495
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   22
         Top             =   3720
         Width           =   1575
      End
      Begin VB.TextBox fechanac 
         Height          =   495
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   20
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox distrito 
         Height          =   495
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   18
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox direccion 
         Height          =   495
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   16
         Top             =   2280
         Width           =   4575
      End
      Begin VB.TextBox dni 
         Height          =   495
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   14
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tficha.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tficha.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox nombre 
         Height          =   495
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   10
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox codigo 
         Height          =   495
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RUC"
         Height          =   495
         Left            =   120
         TabIndex        =   41
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Correo E."
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefono Movil"
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Anexo"
         Height          =   495
         Left            =   3240
         TabIndex        =   27
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefono Trabajo"
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefono Particular"
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaNac"
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distrito"
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DNI"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFF00&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   10470
      TabIndex        =   1
      Top             =   0
      Width           =   10530
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tficha.frx":0F5C
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox xcliente 
         Height          =   375
         Left            =   4440
         MaxLength       =   11
         TabIndex        =   33
         Text            =   "*"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdHelp 
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
         Picture         =   "tficha.frx":170A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ayuda"
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
         Picture         =   "tficha.frx":291C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Borrar registro"
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
         Picture         =   "tficha.frx":3B2E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
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
         Picture         =   "tficha.frx":4D40
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Imprimir"
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
         Picture         =   "tficha.frx":5F52
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cliente"
         Height          =   375
         Left            =   3720
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   13150
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
            LCID            =   10250
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
            LCID            =   10250
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
   Begin VB.Menu ahyy1 
      Caption         =   "&Add"
   End
   Begin VB.Menu dmi22 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu dfj8221 
      Caption         =   "&Borra"
   End
   Begin VB.Menu dk281 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu fdo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tficha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rt As New ADODB.Recordset
Dim rsf As New ADODB.Recordset

Private Sub sql()
On Error GoTo cmd5_err
Dim cad As String
cad = "SELECT Nombre,cliente,Dni,Direccion,Telefonop,Telefonot,anexot,Telefonoc,Distrito,correo,FechaNac,ruc FROM cliente "
If xcliente <> "*" Then
   cad = cad & " where cliente='" & xcliente & "'"
End If
If rsf.State = 1 Then rsf.Close
rsf.Open cad, cn, adOpenStatic, adLockOptimistic
Set dbgrid1.DataSource = rsf
dbgrid1.Columns(0).Width = 5000
dbgrid1.Columns(1).Width = 1000
Exit Sub
cmd5_err:
MsgBox "Aviso en sql " + Error, 48, "Aviso"
Exit Sub
End Sub

Private Sub ahyy1_Click()
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub

Frame1.Visible = True
Frame1.Caption = "NUEVO"
codigo = ""
codigo.Enabled = True
inicializa
codigo.SetFocus
End Sub
Sub inicializa()
ruc = ""
nombre = ""
dni = ""
fechanac = ""
correo = ""
telefonop = ""
telefonot = ""
anexo = ""
telefonom = ""
direccion = ""
distrito = ""

End Sub

Private Sub anexo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
telefonom.SetFocus

End Sub

Private Sub buffer_DblClick()
Command2_Click
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   If opcion1 = 1 Then
      Frame2.Visible = False
      xcliente.SetFocus
      Exit Sub
   End If
End If
Command2_Click

End Sub

Private Sub cmdAddEntry_Click()
ahyy1_Click
End Sub

Private Sub cmdDelete_Click()
dfj8221_Click
End Sub

Private Sub cmdExit_Click()
fdo33_Click
End Sub

Private Sub cmdHelp_Click()
dmi22_Click
End Sub

Private Sub cmdPrint_Click()
dk281_Click
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
nombre.SetFocus
End Sub

Private Sub Command1_Click()
sql
dbgrid1.SetFocus

End Sub

Private Sub Command2_Click()
ejecuta 1
End Sub

Private Sub Command3_Click()
Dim found As Integer
Dim rsexiste As New ADODB.Recordset
Dim cad As String
On Error GoTo cmd2_err
If Len(codigo) = 0 Then
   codigo.SetFocus
   Exit Sub
End If
If Len(nombre) = 0 Then
   nombre.SetFocus
   Exit Sub
End If

If Frame1.Caption = "NUEVO" Then
   rsexiste.Open "SELECT cliente FROM cliente where cliente='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount > 0 Then  'si existe
      MsgBox "Ya existe codigo ", 48, "Aviso"
      codigo.SetFocus
      Exit Sub
   End If
   cad = "INSERT INTO cliente VALUES('" & Trim(codigo) & "','" & Trim(dni) & "','" & Trim(nombre) & "','" & Trim(direccion) & "','" & Trim(distrito) & "','" & Trim(fechanac) & "','" & Trim(telefonop) & "','" & Trim(telefonot) & "','" & Trim(anexo) & "','" & Trim(telefonom) & "','" & Trim(correo) & "','" & Trim(ruc) & "')"
   cn.Execute (cad)
   sql

dbgrid1.SetFocus

   fdo33_Click
End If
If Frame1.Caption = "MODIFICA" Then
   cad = "UPDATE cliente SET dni = '" & Trim(dni) & "', nombre= '" & Trim(nombre) & "', direccion= '" & Trim(direccion) & "', distrito= '" & Trim(distrito) & "', fechanac= '" & Trim(fechanac) & "', telefonop= '" & Trim(telefonop) & "', telefonot= '" & Trim(telefonot) & "', anexot= '" & Trim(anexo) & "', telefonoc= '" & Trim(telefonom) & "', correo= '" & Trim(correo) & "',ruc='" & Trim(ruc) & "' WHERE cliente = '" & Trim(codigo) & "'"
   cn.Execute (cad)
   sql
   dbgrid1.SetFocus
   fdo33_Click
End If


Exit Sub
cmd2_err:
MsgBox "Aviso en command3 " + Error$, 48, "Aviso"
Exit Sub

End Sub

Private Sub Command4_Click()
fdo33_Click
End Sub

Private Sub correo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub


End Sub

Private Sub dbGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   If opcion1 = 1 Then
      xcliente = Trim(dbGrid2.Columns(1))
      Frame2.Visible = False
      Frame2.Enabled = False
      xcliente.SetFocus
      Exit Sub
   End If

End If

End Sub

Private Sub dbGrid2_KeyPress(KeyAscii As Integer)
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

Private Sub dfj8221_Click()
Dim buf As String
On Error GoTo cmd4_err
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
buf = Trim(dbgrid1.Columns(1))
If MsgBox("Desea Borrar " + dbgrid1.Columns(1), 1, "Aviso") = 1 Then
   cn.Execute ("DELETE   FROM cliente WHERE cliente ='" & Trim(dbgrid1.Columns(1)) & "'")
   rsf.Requery
   sql
   dbgrid1.SetFocus
End If
dbgrid1.SetFocus
Exit Sub
cmd4_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
dbgrid1.SetFocus
Exit Sub

End Sub

Private Sub direccion_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
distrito.SetFocus

End Sub

Private Sub distrito_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
fechanac.SetFocus

End Sub

Private Sub dk281_Click()
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
If rt.State = 1 Then rt.Close
rt.Open "SELECT * FROM cliente ", cn, adOpenKeyset, adLockOptimistic
Set trepcli1.DataSource = rt
trepcli1.Show 1

End Sub

Private Sub dmi22_Click()
On Error GoTo cmd3_err
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
codigo = Trim(dbgrid1.Columns(1))
nombre = Trim(dbgrid1.Columns(0))
dni = Trim(dbgrid1.Columns(2))
direccion = Trim(dbgrid1.Columns(3))
telefonop = Trim(dbgrid1.Columns(4))
telefonot = Trim(dbgrid1.Columns(5))
anexo = Trim(dbgrid1.Columns(6))
telefonom = Trim(dbgrid1.Columns(7))
distrito = Trim(dbgrid1.Columns(8))
correo = Trim(dbgrid1.Columns(9))
fechanac = Trim(dbgrid1.Columns(10))
ruc = Trim(dbgrid1.Columns(11))
Frame1.Visible = True
Frame1.Caption = "MODIFICA"
codigo.Enabled = False
nombre.SetFocus
Exit Sub
cmd3_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub
End Sub

Private Sub dni_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
ruc.SetFocus

End Sub

Private Sub fdo33_Click()
If Frame2.Visible = True Then
   buffer_KeyPress 27
   Exit Sub
End If
If Frame1.Visible = True Then
   If Frame1.Caption = "NUEVO" Then
      Frame1.Visible = False
      dbgrid1.SetFocus
   End If
   If Frame1.Caption = "MODIFICA" Then
      Frame1.Visible = False
      dbgrid1.SetFocus
   End If
   Exit Sub
End If
tficha.Hide
Unload tficha
End Sub

Private Sub fechanac_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
telefonop.SetFocus

End Sub

Private Sub Form_Load()
sql

End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
dni.SetFocus
End Sub

Private Sub ruc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
direccion.SetFocus

End Sub

Private Sub telefonom_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
correo.SetFocus

End Sub

Private Sub telefonop_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
telefonot.SetFocus

End Sub

Private Sub telefonot_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
anexo.SetFocus

End Sub

Private Sub xcliente_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
sql
dbgrid1.SetFocus
End Sub

Private Sub xcliente_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_xcliente
End If

End Sub
Sub consulta_xcliente()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT * FROM cliente  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      Exit Sub
   End If

Frame2.Visible = True
Frame2.Enabled = True
buffer = ""
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Cliente"
Combo1.ListIndex = 0
opcion1 = 1
buffer.SetFocus
Command2_Click

End Sub
Sub ejecuta(sw As Integer)
Dim rconsulta As New ADODB.Recordset
Dim cad As String
If opcion1 = 1 Then  'clientes
   If Len(buffer) = 0 Then
      cad = "SELECT Nombre,Cliente FROM cliente  "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Nombre,Cliente FROM cliente where  " & Combo1 & " like '" & buffer & "%'"
   End If
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True And rconsulta.BOF = True Then
      buffer.SetFocus
      Exit Sub
   End If
   Set dbGrid2.DataSource = rconsulta
   dbGrid2.Columns(0).Width = 5000
   dbGrid2.Columns(1).Width = 1000
   If sw = 1 Then
      dbGrid2.SetFocus
   End If
End If






End Sub


