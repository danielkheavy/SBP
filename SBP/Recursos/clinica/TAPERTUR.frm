VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tapertur 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla Aperturas"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   12030
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
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
         TabIndex        =   23
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
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
      End
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
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbGrid2 
         Height          =   6735
         Left            =   240
         TabIndex        =   24
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Apertura "
      Height          =   5175
      Left            =   1440
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   8415
      Begin VB.TextBox fechaf 
         Height          =   495
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   16
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox turno 
         Height          =   495
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   14
         Top             =   1320
         Width           =   615
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
         Left            =   6720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TAPERTUR.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   480
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
         Left            =   6720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TAPERTUR.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox fechai 
         Height          =   495
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   10
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox caja 
         Height          =   495
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   0
         Top             =   840
         Width           =   615
      End
      Begin VB.Label cajero 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1680
         TabIndex        =   19
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajero"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Final"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFF00&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11970
      TabIndex        =   2
      Top             =   0
      Width           =   12030
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
         Picture         =   "TAPERTUR.frx":0F5C
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
         Picture         =   "TAPERTUR.frx":216E
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "TAPERTUR.frx":3380
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "TAPERTUR.frx":4592
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "TAPERTUR.frx":57A4
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid1 
      Height          =   7455
      Left            =   0
      TabIndex        =   1
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
Attribute VB_Name = "tapertur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public rt As New ADODB.Recordset
Public rsede As New ADODB.Recordset
Private Sub sql()
On Error GoTo cmd5_err
Dim cad As String
cad = "SELECT Cajero,Caja,Turno,Fechai,Fechaf FROM apertura where cajero='" & Trim(dgusuario) & "'"
If rsede.State = 1 Then rsede.Close
rsede.Open cad, cn, adOpenStatic, adLockOptimistic
Set dbgrid1.DataSource = rsede
dbgrid1.Columns(0).Width = 1000
dbgrid1.Columns(1).Width = 1000
dbgrid1.Columns(2).Width = 1000
Exit Sub
cmd5_err:
MsgBox "Aviso en sql " + Error, 48, "Aviso"
Exit Sub
End Sub

Private Sub ahyy1_Click()
If Frame1.Visible = True Then Exit Sub
Frame1.Visible = True
Frame1.Caption = "NUEVO"
cajero = dgusuario
caja = ""
turno = ""
fechai = ""
fechaf = ""
caja.Enabled = True
turno.Enabled = True
inicializa
caja.SetFocus
End Sub
Sub inicializa()
fechai = Format(Now, "dd/mm/yyyy")
fechaf = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   If opcion1 = 1 Then
      Frame2.Visible = False
      caja.SetFocus
      Exit Sub
   End If
   If opcion1 = 2 Then
      Frame2.Visible = False
      turno.SetFocus
      Exit Sub
   End If

End If

End Sub

Private Sub caja_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
turno.SetFocus
End Sub

Private Sub caja_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_caja
End If

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


Private Sub Command1_Click()
ejecuta 1

End Sub

Private Sub Command3_Click()
Dim found As Integer
Dim rsexiste As New ADODB.Recordset
Dim cad As String
On Error GoTo cmd2_err
If Len(caja) = 0 Then
   caja.SetFocus
   Exit Sub
End If
If Len(turno) = 0 Then
   turno.SetFocus
   Exit Sub
End If
If Not IsDate(fechai) Then
   fechai.SetFocus
   Exit Sub
End If
If Not IsDate(fechaf) Then
   fechaf.SetFocus
   Exit Sub
End If

If Frame1.Caption = "NUEVO" Then
   rsexiste.Open "SELECT caja FROM apertura where cajero='" & Trim(cajero) & "' and caja='" & Trim(caja) & "' and turno='" & Trim(turno) & "'", cn, adOpenKeyset, adLockOptimistic
   If rsexiste.RecordCount > 0 Then  'si existe
      MsgBox "Ya existe caja aperturada ", 48, "Aviso"
      caja.SetFocus
      Exit Sub
   End If
   cad = "INSERT INTO apertura VALUES('" & Trim(cajero) & "','" & Trim(caja) & "','" & Trim(turno) & "','" & Trim(fechai) & "','" & Trim(fechaf) & "')"
   cn.Execute (cad)
   sql
   fdo33_Click
End If
If Frame1.Caption = "MODIFICA" Then
   cad = "UPDATE apertura SET fechai = '" & Trim(fechai) & "',fechaf='" & Trim(fechaf) & "' WHERE cajero = '" & Trim(cajero) & "' and caja='" & Trim(caja) & "' and turno='" & Trim(turno) & "'"
   cn.Execute (cad)
   sql
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

Private Sub dbGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   If opcion1 = 1 Then
      caja = Trim(dbGrid2.Columns(1))
      Frame2.Visible = False
      Frame2.Enabled = False
      caja.SetFocus
      Exit Sub
   End If
   If opcion1 = 2 Then
      turno = Trim(dbGrid2.Columns(1))
      Frame2.Visible = False
      Frame2.Enabled = False
      turno.SetFocus
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
If Frame1.Visible = True Then Exit Sub
buf = Trim(dbgrid1.Columns(1))
If MsgBox("Desea Borrar " + dbgrid1.Columns(1), 1, "Aviso") = 1 Then
   cn.Execute ("DELETE   FROM apertura WHERE cajero ='" & Trim(dbgrid1.Columns(0)) & "' and caja='" & Trim(dbgrid1.Columns(1)) & "' and turno='" & Trim(dbgrid1.Columns(2)) & "'")
   rsede.Requery
   sql
End If
dbgrid1.SetFocus
Exit Sub
cmd4_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
dbgrid1.SetFocus
Exit Sub

End Sub

Private Sub dk281_Click()
If Frame1.Visible = True Then Exit Sub
If rt.State = 1 Then rt.Close
rt.Open "SELECT * FROM apertura ", cn, adOpenKeyset, adLockOptimistic
Set trepcli1.DataSource = rt
trepcli1.Show 1

End Sub

Private Sub dmi22_Click()
On Error GoTo cmd3_err
If Frame1.Visible = True Then Exit Sub
cajero = Trim(dbgrid1.Columns(0))
caja = Trim(dbgrid1.Columns(1))
turno = Trim(dbgrid1.Columns(2))
fechai = Trim(dbgrid1.Columns(3))
fechaf = Trim(dbgrid1.Columns(4))
Frame1.Visible = True
Frame1.Caption = "MODIFICA"
caja.Enabled = False
turno.Enabled = False
fechai.SetFocus
Exit Sub
cmd3_err:
MsgBox "Seleccione un dato ", 48, "Aviso"
Exit Sub
End Sub

Private Sub fdo33_Click()

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
tapertur.Hide
Unload tapertur
End Sub

Private Sub fechaf_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub fechai_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
fechaf.SetFocus

End Sub

Private Sub Form_Load()
sql

End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
End Sub

Private Sub turno_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
fechai.SetFocus

End Sub

Private Sub turno_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_turno
End If

End Sub
Sub consulta_caja()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT * FROM caja  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      Exit Sub
   End If

Frame2.Visible = True
Frame2.Enabled = True
buffer = ""
Combo1.Clear
Combo1.AddItem "caja"
Combo1.ListIndex = 0

opcion1 = 1
buffer.SetFocus
Command1_Click

End Sub
Sub consulta_turno()
Dim cad As String
   Dim rconsulta As New ADODB.Recordset
   cad = "SELECT * FROM turno  "
   If rconsulta.State = 1 Then rconsulta.Close
   rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic
   If rconsulta.EOF = True Or rconsulta.BOF = True Then
      Exit Sub
   End If

Frame2.Visible = True
Frame2.Enabled = True
buffer = ""
Combo1.Clear
Combo1.AddItem "turno"
Combo1.ListIndex = 0

opcion1 = 2
buffer.SetFocus
Command1_Click

End Sub
Sub ejecuta(sw As Integer)
Dim rconsulta As New ADODB.Recordset
Dim cad As String
If opcion1 = 1 Then  'clientes
   If Len(buffer) = 0 Then
      cad = "SELECT Nombre,Caja FROM caja  "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Nombre,Caja FROM Caja where  " & Combo1 & " like '" & buffer & "%'"
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
If opcion1 = 2 Then  'clientes
   If Len(buffer) = 0 Then
      cad = "SELECT Nombre,Turno FROM Turno  "
   End If
   If Len(buffer) > 0 Then
      cad = "SELECT Nombre,Turno FROM Turno where  " & Combo1 & " like '" & buffer & "%'"
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


