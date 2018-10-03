VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form tapertur 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apertura del Dia"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11190
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11190
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   4800
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CR"
         Height          =   855
         Index           =   10
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "9"
         Height          =   855
         Index           =   9
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "8"
         Height          =   855
         Index           =   8
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "7"
         Height          =   855
         Index           =   7
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "6"
         Height          =   855
         Index           =   6
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "5"
         Height          =   855
         Index           =   5
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "4"
         Height          =   855
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "3"
         Height          =   855
         Index           =   3
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "2"
         Height          =   855
         Index           =   2
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1"
         Height          =   855
         Index           =   1
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   855
         Index           =   11
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ENTER"
         Height          =   855
         Index           =   0
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton btnsalir 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cerrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   4200
         Picture         =   "tapertur.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Imprimir todo"
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox fechaf 
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
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   11
      Top             =   4200
      Width           =   1815
   End
   Begin VB.TextBox fechai 
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
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   9
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox cajero 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      MaxLength       =   11
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox turno 
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
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.TextBox CAJA 
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
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   680
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   11130
      TabIndex        =   1
      Top             =   0
      Width           =   11190
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
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tapertur.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Grabar registro"
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
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tapertur.frx":1ADC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   4440
      TabIndex        =   30
      Top             =   3720
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
      _Version        =   393216
      Format          =   70385665
      CurrentDate     =   38186
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   495
      Left            =   4440
      TabIndex        =   31
      Top             =   4200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
      _Version        =   393216
      Format          =   70385665
      CurrentDate     =   38186
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA- COMPUTADORA"
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   960
      TabIndex        =   15
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label fechasis 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   4320
      TabIndex        =   14
      Top             =   2640
      Width           =   135
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8400
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8400
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nota. Fecha Inicio debe coincidir con la fecha de la Caja Registradora"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   4920
      Width           =   8295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA TERMINO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA INICIO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CAJERO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TURNO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label CAJAX 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CAJA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Menu dju2323 
      Caption         =   "&Grabar"
   End
   Begin VB.Menu flo423 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tapertur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnsalir_Click()
Frame1.Visible = False
End Sub

Private Sub CAJA_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(CAJA) = 0 Then
   CAJA.SetFocus
   Exit Sub
End If
turno.SetFocus
End Sub

Private Sub CAJAX_Click()
Frame1.Caption = "CAJA"
Frame1.Visible = True


End Sub

Private Sub cmdExit_Click()
flo423_Click
End Sub

Private Sub cmdSave_Click()
Dim found As Integer
Dim mytablex As New ADODB.Recordset

On Error GoTo cmd23_err
If Frame1.Visible = True Then Exit Sub
found = valida()
If found = 0 Then
   MsgBox "Campos invalidos", 48, "Aviso"
   Exit Sub
End If

mytablex.Open "SELECT * FROM apertura where  cajero='" & cajero & "' and caja='" & CAJA & "' and turno='" & turno & "'", cn, adOpenKeyset, adLockOptimistic
If mytablex.RecordCount = 0 Then  'si existe
   mytablex.AddNew
   grabando mytablex
   mytablex.Update
   MsgBox "La caja ha sido aperturado", 48, "Aviso"
Else
   If MsgBox("Desea Reescribir?", 1, "Aviso") = 1 Then
   'mytablex.Edit
   grabando mytablex
   mytablex.Update
   End If
End If
'------------------------------------- ------------
mytablex.Close
 

flo423_Click
Exit Sub
cmd23_err:
MsgBox "Error " & error$, 48, "Aviso"
Exit Sub
End Sub
Sub grabando(mytablex As ADODB.Recordset)
mytablex.Fields("cajero") = cajero
mytablex.Fields("caja") = CAJA
mytablex.Fields("turno") = turno
mytablex.Fields("fechai") = Format(fechai, "dd/mm/yyyy")
mytablex.Fields("fechaf") = Format(fechaf, "dd/mm/yyyy")
End Sub

Private Sub Command1_Click(Index As Integer)
If Index = 10 Then
   If Frame1.Caption = "CAJERO" Then
      cajero = ""
   End If
   If Frame1.Caption = "CAJA" Then
      CAJA = ""
   End If
   If Frame1.Caption = "TURNO" Then
      turno = ""
   End If
    Exit Sub
End If
If Index = 0 Then 'enter
If Frame1.Caption = "CAJERO" Then
      cajero.SetFocus
   End If
   If Frame1.Caption = "CAJA" Then
      CAJA.SetFocus
   End If
   If Frame1.Caption = "TURNO" Then
      turno.SetFocus
   End If
   Frame1.Visible = False
   Exit Sub
End If

   If Frame1.Caption = "CAJERO" Then
      cajero = cajero & Command1(Index).Caption
   End If
   If Frame1.Caption = "CAJA" Then
      CAJA = CAJA & Command1(Index).Caption
   End If
   If Frame1.Caption = "TURNO" Then
      turno = turno & Command1(Index).Caption
   End If

End Sub

Private Sub Command2_Click()

End Sub

Private Sub dju2323_Click()
If Frame1.Visible = True Then Exit Sub
flo423_Click
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
fechai = DTPicker1.Value
End Sub

Private Sub DTPicker1_Change()
fechai = DTPicker1.Value
End Sub

Private Sub DTPicker2_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
fechaf = DTPicker2.Value
End Sub

Private Sub DTPicker2_Change()
fechaf.Text = DTPicker2.Value
End Sub

Private Sub fechaf_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(CAJA) = 0 Then
   CAJA.SetFocus
   Exit Sub
End If
If Len(turno) = 0 Then
   turno.SetFocus
   Exit Sub
End If
If Len(fechaf) = 0 Then
   fechaf = Format(Now, "dd/mm/yyyy")
End If
cmdSave_Click
End Sub

Private Sub fechai_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(CAJA) = 0 Then
   CAJA.SetFocus
   Exit Sub
End If
If Len(turno) = 0 Then
   turno.SetFocus
   Exit Sub
End If
If Len(fechai) = 0 Then
   fechai = Format(Now, "dd/mm/yyyy")
End If
found = busca_turno()
If found = 1 Then  'si es el mismo dia
   fechaf = Format(Now, "dd/mm/yyyy")
End If

If found = 2 Then 'si es otro dia sumar 1
   fechaf = Format(Now + 1, "dd/mm/yyyy")
End If
fechaf.SetFocus

End Sub

Private Sub flo423_Click()
If Frame1.Visible = True Then
   Frame1.Visible = False
   Exit Sub
End If
tapertur.Hide
Unload tapertur
End Sub

Private Sub Form_Load()
fechasis = Format(Now, "dd/mm/yyyy")
End Sub
Function valida()
Dim found As Integer
If Len(CAJA) = 0 Then
   CAJA.SetFocus
   Exit Function
End If

found = busca_caja()
If found = 0 Then
   CAJA = ""
   CAJA.SetFocus
   Exit Function
End If
If Len(turno) = 0 Then
   turno.SetFocus
   Exit Function
End If
found = busca_turno()
If found = 0 Then
   turno = ""
   turno.SetFocus
   Exit Function
End If
If valida_fecha("" & fechai) = 0 Then
   fechai = ""
   fechai.SetFocus
   Exit Function
End If
If valida_fecha("" & fechaf) = 0 Then
   fechaf = ""
   fechaf.SetFocus
   Exit Function
End If
valida = 1
End Function
Function busca_caja()

Dim mytablex As New ADODB.Recordset

mytablex.Open "SELECT * FROM parameca where  caja='" & CAJA & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
    busca_caja = 1
   End If
'------------------------------------- ------------
mytablex.Close
 
End Function
Function busca_turno()
Dim mytablex As New ADODB.Recordset
mytablex.Open "SELECT * FROM turno where  turno='" & turno & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
   Select Case "" & mytablex.Fields("flag")
   Case "1"
        busca_turno = 1
   Case "2"
        busca_turno = 2
   End Select
End If
mytablex.Close
 
End Function
Function busca_apertura()

Dim mytablex As New ADODB.Recordset
Dim fechag As String

mytablex.Open "SELECT * FROM apertura where  cajero='" & cajero & "' and caja='" & CAJA & "' and turno='" & turno & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
   busca_apertura = 1
   fechai = Format("" & mytablex.Fields("fechai"), "dd/mm/yyyy")
   fechaf = Format("" & mytablex.Fields("fechaf"), "dd/mm/yyyy")
End If
'------------------------------------- ------------
mytablex.Close
 
End Function

Private Sub Label1_Click()
Frame1.Caption = "TURNO"
Frame1.Visible = True


End Sub

Private Sub Label2_Click()
Exit Sub
Frame1.Caption = "CAJERO"
Frame1.Visible = True
teclado = ""


End Sub

Private Sub turno_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(CAJA) = 0 Then
   CAJA.SetFocus
   Exit Sub
End If
If Len(turno) = 0 Then
   turno.SetFocus
   Exit Sub
End If
found = busca_apertura()
fechai.SetFocus

End Sub

