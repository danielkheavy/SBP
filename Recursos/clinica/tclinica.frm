VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tclinica 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla Clinicas"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10275
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Frame1"
      Height          =   5175
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox codigo 
         Height          =   495
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox nombre 
         Height          =   495
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   10
         Top             =   840
         Width           =   4575
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
         Left            =   4920
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tclinica.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2520
         UseMaskColor    =   -1  'True
         Width           =   1215
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
         Left            =   4920
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tclinica.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   10215
      TabIndex        =   1
      Top             =   0
      Width           =   10275
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
         Picture         =   "tclinica.frx":0F5C
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "tclinica.frx":216E
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
         Picture         =   "tclinica.frx":3380
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
         Picture         =   "tclinica.frx":4592
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
         Picture         =   "tclinica.frx":57A4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Nuevo registro"
         Top             =   0
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
Attribute VB_Name = "tclinica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SQL()

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd5_err

    Dim cad As String

    cad = "SELECT Nombre,Clinica FROM clinica "

    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = mytablex
    dbGrid1.columns(0).Width = 5000
    dbGrid1.columns(1).Width = 1000
    Exit Sub
cmd5_err:
    MsgBox "Aviso en sql " + error, 48, "Aviso"
    Exit Sub

End Sub

Private Sub ahyy1_Click()

    If Frame1.Visible = True Then Exit Sub
    Frame1.Visible = True
    Frame1.Caption = "NUEVO"
    codigo = ""
    codigo.Enabled = True
    inicializa
    codigo.SetFocus

End Sub

Sub inicializa()
    nombre = ""

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

Private Sub Command3_Click()

    Dim found    As Integer

    Dim rsexiste As New ADODB.Recordset

    Dim cad      As String

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
        rsexiste.Open "SELECT clinica FROM clinica where clinica='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic

        If rsexiste.RecordCount > 0 Then  'si existe
            MsgBox "Ya existe codigo ", 48, "Aviso"
            codigo.SetFocus
            Exit Sub

        End If

        cad = "INSERT INTO clinica VALUES('" & Trim(codigo) & "','" & Trim(nombre) & "')"
        cn.Execute (cad)
        SQL
        fdo33_Click

    End If

    If Frame1.Caption = "MODIFICA" Then
        cad = "UPDATE clinica SET nombre = '" & Trim(nombre) & "' WHERE clinica = '" & Trim(codigo) & "'"
        cn.Execute (cad)
        SQL
        fdo33_Click

    End If

    Exit Sub
cmd2_err:
    MsgBox "Aviso en command3 " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command4_Click()
    fdo33_Click

End Sub

Private Sub dfj8221_Click()

    Dim buf As String

    On Error GoTo cmd4_err

    If Frame1.Visible = True Then Exit Sub
    buf = Trim(dbGrid1.columns(1))

    If MsgBox("Desea Borrar " + dbGrid1.columns(1), 1, "Aviso") = 1 Then
        cn.Execute ("DELETE   FROM clinica WHERE clinica ='" & Trim(dbGrid1.columns(1)) & "'")
        cn.Requery
        SQL

    End If

    dbGrid1.SetFocus
    Exit Sub
cmd4_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    dbGrid1.SetFocus
    Exit Sub

End Sub

Private Sub dk281_Click()
    'Public rt As New ADODB.Recordset
    'If Frame1.Visible = True Then Exit Sub
    'If rt.State = 1 Then rt.Close
    'rt.Open "SELECT * FROM clinica ", cn, adOpenKeyset, adLockOptimistic
    'Set trepcli1.DataSource = rt
    'trepcli1.Show 1

End Sub

Private Sub dmi22_Click()

    On Error GoTo cmd3_err

    If Frame1.Visible = True Then Exit Sub
    codigo = Trim(dbGrid1.columns(1))
    nombre = Trim(dbGrid1.columns(0))
    Frame1.Visible = True
    Frame1.Caption = "MODIFICA"
    codigo.Enabled = False
    nombre.SetFocus
    Exit Sub
cmd3_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fdo33_Click()

    If Frame1.Visible = True Then
        If Frame1.Caption = "NUEVO" Then
            Frame1.Visible = False
            dbGrid1.SetFocus

        End If

        If Frame1.Caption = "MODIFICA" Then
            Frame1.Visible = False
            dbGrid1.SetFocus

        End If

        Exit Sub

    End If

    tclinica.Hide
    Unload tclinica

End Sub

Private Sub Form_Load()
    SQL

End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub
