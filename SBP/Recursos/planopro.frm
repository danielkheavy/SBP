VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form planopro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plano de Produccion"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   13620
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   2340
      TabIndex        =   7
      Top             =   1980
      Visible         =   0   'False
      Width           =   8055
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancelar 
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
         Left            =   6360
         MaskColor       =   &H00FFFFFF&
         Picture         =   "planopro.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton cmdGrabar 
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
         Left            =   6360
         MaskColor       =   &H00FFFFFF&
         Picture         =   "planopro.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   135
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7830
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "planopro.frx":0F5C
      Height          =   6015
      Left            =   0
      OleObjectBlob   =   "planopro.frx":0F70
      TabIndex        =   6
      Top             =   600
      Width           =   12735
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   680
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   13560
      TabIndex        =   0
      Top             =   0
      Width           =   13620
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "planopro.frx":218B
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Consulta"
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
         Picture         =   "planopro.frx":339D
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Picture         =   "planopro.frx":45AF
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "planopro.frx":57C1
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Picture         =   "planopro.frx":69D3
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Menu snnue1 
      Caption         =   "&Nuevo"
   End
   Begin VB.Menu bo7823 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu modi3321 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu imp9021 
      Caption         =   "&Imprime"
   End
   Begin VB.Menu zom812 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu ccomi12 
      Caption         =   "&Consulta"
   End
   Begin VB.Menu flo343 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "planopro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bo7823_Click()

    On Error GoTo cmd456_err

    If MsgBox("Desea Borrar " & "" & Data2.Recordset.Fields("numero"), 1, "Aviso") <> 1 Then Exit Sub

    mydbxglo.Execute "DELETE FROM DPRODUCC where numero='" & "" & Data2.Recordset.Fields("numero") & "'"
 
    Data2.Recordset.Delete
    Exit Sub
cmd456_err:
    Exit Sub

End Sub

Private Sub ccomi12_Click()

    If Frame2.Visible = True Then Exit Sub
    Frame2.Visible = True
    fechai.SetFocus

End Sub

Private Sub cmdCancelar_Click()
    flo343_Click

End Sub

Private Sub cmdExit_Click()
    flo343_Click

End Sub

Private Sub cmdGrabar_Click()
    sql_cabeza
    flo343_Click

End Sub

Private Sub flo343_Click()

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Exit Sub

    End If

    planopro.Hide
    Unload planopro

End Sub

Sub sql_cabeza()

    On Error GoTo cmd37_err

    Dim buf As String

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub

    'buf = "select * from cproduCc where "
    'buf = buf & "  fecha>=" & "DateValue('" & Fechai & "'" & ")"
    'buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
    'buf = buf & " order by val(numero)"
    '               Data2.Connect = "foxpro 2.5;"
    '               Data2.DatabaseName = globaldir
    '               Data2.RecordSource = buf
    '               Data2.Refresh
    '               DBGrid2.SetFocus

    txempre.Open cad, cn, adOpenStatic, adLockOptimistic
    Exit Sub
cmd37_err:
    MsgBox "Error en select " & error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fechaf = Format(Now, "dd/mm/yyyy")
    sql_cabeza

End Sub

Private Sub modi3321_Click()

    Dim found As Integer

    On Error GoTo cmd67_err

    If Frame2.Visible = True Then Exit Sub
    found = copiar_tmprodu()

    If found = 0 Then
        MsgBox "Error al copiar temporal ", 48, "Aviso"
        Exit Sub

    End If

    tproducc.Numero = "" & Data2.Recordset.Fields("numero")
    tproducc.bandera = "Modifica"
    tproducc.Numero.Enabled = False
    tproducc.Show 1
    Exit Sub
cmd67_err:
    Exit Sub

End Sub

Private Sub snnue1_Click()

    Dim found As Integer

    If Frame2.Visible = True Then Exit Sub
    found = copiar_tmprodu()

    If found = 0 Then
        MsgBox "Error al copiar temporal ", 48, "Aviso"
        Exit Sub

    End If

    tproducc.bandera = "Nuevo"
    tproducc.Show 1

End Sub

