VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form planillv 
   BackColor       =   &H80000009&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Consultas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   10695
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
         Left            =   8160
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "planillv.frx":0000
         Height          =   5775
         Left            =   120
         OleObjectBlob   =   "planillv.frx":0014
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   840
         Width           =   10335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Consulta PLanilla"
      Height          =   7335
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   10575
      Begin VB.TextBox diatraba 
         Height          =   375
         Left            =   7680
         MaxLength       =   11
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox horatraba 
         Height          =   375
         Left            =   7680
         MaxLength       =   11
         TabIndex        =   22
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox horaextr 
         Height          =   375
         Left            =   7680
         MaxLength       =   11
         TabIndex        =   21
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Recalculo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   20
         Top             =   6000
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Imprime"
         Height          =   375
         Left            =   8880
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "planillv.frx":09DF
         Height          =   1935
         Left            =   120
         OleObjectBlob   =   "planillv.frx":09F3
         TabIndex        =   24
         Top             =   1800
         Width           =   4935
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Bindings        =   "planillv.frx":157E
         Height          =   1935
         Left            =   5160
         OleObjectBlob   =   "planillv.frx":1592
         TabIndex        =   25
         Top             =   1800
         Width           =   4935
      End
      Begin MSDBGrid.DBGrid DBGrid5 
         Bindings        =   "planillv.frx":211D
         Height          =   1935
         Left            =   120
         OleObjectBlob   =   "planillv.frx":2131
         TabIndex        =   26
         Top             =   4320
         Width           =   4935
      End
      Begin VB.Label tipopla 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   47
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipopla"
         Height          =   375
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label xnombre 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   45
         Top             =   600
         Width           =   3735
      End
      Begin VB.Label xcodigo 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingresos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   1440
         Width           =   4935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aportes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   3960
         Width           =   4935
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descuentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5160
         TabIndex        =   41
         Top             =   1440
         Width           =   4935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Totales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5160
         TabIndex        =   40
         Top             =   3960
         Width           =   4935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Ingresos"
         Height          =   375
         Left            =   5160
         TabIndex        =   39
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Descuentos"
         Height          =   375
         Left            =   5160
         TabIndex        =   38
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TotalAportes"
         Height          =   375
         Left            =   5160
         TabIndex        =   37
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TotalCobrar"
         Height          =   375
         Left            =   5160
         TabIndex        =   36
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label totingreso 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6840
         TabIndex        =   35
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label totdscto 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6840
         TabIndex        =   34
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label totaporte 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6840
         TabIndex        =   33
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label totalcobrar 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   6840
         TabIndex        =   32
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dias Trabajados"
         Height          =   375
         Left            =   6120
         TabIndex        =   31
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Horas Trabajados"
         Height          =   375
         Left            =   6120
         TabIndex        =   30
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Horas Extras"
         Height          =   375
         Left            =   6120
         TabIndex        =   29
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Division"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label xdivision 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   680
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   10650
      TabIndex        =   14
      Top             =   0
      Width           =   10710
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "planillv.frx":2CBC
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Consulta"
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
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "planillv.frx":3ECE
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "planillv.frx":50E0
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.TextBox periodo 
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Consulta"
      Height          =   3255
      Left            =   1440
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   8055
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
         Left            =   6600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "planillv.frx":62F2
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   1335
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
         Left            =   6600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "planillv.frx":6AA0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   6
         Text            =   "*"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox NOMBRE 
         Height          =   375
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   5
         Text            =   "*"
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox division 
         Height          =   375
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "*"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Division"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.ComboBox tipopla1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.Data Data5 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data4 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data3 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   9840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "planillv.frx":724E
      Height          =   5775
      Left            =   0
      OleObjectBlob   =   "planillv.frx":7262
      TabIndex        =   0
      Top             =   1560
      Width           =   10575
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Periodo"
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Planilla"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.Menu plai8232 
      Caption         =   "&Planilla"
   End
   Begin VB.Menu dki3243 
      Caption         =   "&Trabajador"
   End
   Begin VB.Menu dflo3433 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "planillv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub buffer_DblClick()
Command1_Click
End Sub



Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
dflo3433_Click
Exit Sub
End If
Command1_Click

End Sub

Private Sub cmdCancelar_Click()
dflo3433_Click
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
NOMBRE.SetFocus
End Sub

Private Sub Command1_Click()
Dim buf As String
If opcion1 = "1" Then
   If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from Vendedor "
   Else
   buf = "select Nombre,Codigo from vendedor where " & Combo1 & " like '" & buffer & "%'"
   End If
End If
If opcion1 = "2" Or opcion1 = "3" Or opcion1 = "4" Then
   If Len(buffer) = 0 Then
   buf = "select Descripcio,tipo from tplanico "
   Else
   buf = "select Descripcio,Tipo from tplanico where " & Combo1 & " like '" & buffer & "%'"
   End If
End If

               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldir
               Data1.RecordSource = buf
               Data1.Refresh
               If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
                  Data1.Recordset.Close
                  buffer.SetFocus
                  Exit Sub
               End If
               DBGrid1.Columns(0).Width = 4000
               DBGrid1.Columns(1).Width = 2000
               DBGrid1.SetFocus

End Sub

Private Sub Command2_Click()
'recalculo_basico
'recalculo_descuento
'recalculo_aportacion
suma_total
End Sub

Private Sub Command4_Click()
consulta_sql
dflo3433_Click
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim found As Integer
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   Exit Sub
   If opcion1 = "1" Then
   codigo = DBGrid1.Columns(1)
   Frame1.Visible = False
   codigo.SetFocus
   codigo_KeyPress 13
   End If
   If opcion1 = "2" Then
      found = consulta_repitencia2("" & DBGrid1.Columns(1)) 'ingreso
      If found = 1 Then
         MsgBox "Codigo ya registrado ", 48, "Aviso"
         Exit Sub
      End If
      found = graba_ingreso("" & DBGrid1.Columns(1), 1)
      If found = 0 Then Exit Sub
      Frame1.Visible = False
      DBGrid3.SetFocus
   End If
   If opcion1 = "3" Then
      found = consulta_repitencia3("" & DBGrid1.Columns(1)) 'aportacion
      If found = 1 Then
         MsgBox "Codigo ya registrado ", 48, "Aviso"
         Exit Sub
      End If
      found = graba_ingreso("" & DBGrid1.Columns(1), 2)
      If found = 0 Then Exit Sub
      Frame1.Visible = False
      DBGrid4.SetFocus
   End If
   If opcion1 = "4" Then
      found = consulta_repitencia4("" & DBGrid1.Columns(1)) 'descuento
      If found = 1 Then
         MsgBox "Codigo ya registrado ", 48, "Aviso"
         Exit Sub
      End If
      found = graba_ingreso("" & DBGrid1.Columns(1), 3)
      If found = 0 Then Exit Sub
      Frame1.Visible = False
      DBGrid5.SetFocus
   End If

End If
End Sub

Private Sub DBGrid2_DblClick()
Dim found As Integer
On Error GoTo cmd34_err
If tipopla1 = "%" Then
   MsgBox "No ha seleccionado un tipo de planilla", 48, "Aviso"
   Exit Sub
End If
If Frame1.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
Frame2.Visible = True
tipopla = tipopla1
xcodigo = "" & Data2.Recordset.Fields("codigo")
xnombre = "" & Data2.Recordset.Fields("nombre")
xdivision = "" & Data2.Recordset.Fields("division")
consulta_planilla
suma_total
found = ver_sisper("" & xcodigo)
DBGrid3.SetFocus
Exit Sub
cmd34_err:
Exit Sub
End Sub

Private Sub dbgrid3_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex <> 1 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
       Case 1
            If Len("" & DBGrid3.Columns(0)) = 0 Then
               Cancel = True
               Exit Sub
            End If
End Select
End Sub

Private Sub dbgrid3_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
If ColIndex <> 1 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
       Case 1
            'MsgBox "" & DBGrid3.Columns(1)
            If Not IsNumeric("" & DBGrid3.Columns(1)) Then
               Cancel = True
               Exit Sub
            End If
End Select

End Sub

Private Sub DBGrid3_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo cmd56_err
If KeyCode = &H2E Then  'borrar linea
   If MsgBox("Desea Borrar ", 1, "Borrar") <> 1 Then Exit Sub
   Data3.Recordset.Delete
   Data3.Refresh
   Exit Sub
End If
Exit Sub
cmd56_err:
Data3.Refresh
Exit Sub

End Sub

Private Sub dbgrid4_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex <> 1 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
       Case 1
            If Len("" & DBGrid4.Columns(0)) = 0 Then
               Cancel = True
               Exit Sub
            End If
End Select

End Sub

Private Sub DBGrid4_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
If ColIndex <> 1 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
       Case 1
            'MsgBox "" & DBGrid3.Columns(1)
            If Not IsNumeric("" & DBGrid4.Columns(1)) Then
               Cancel = True
               Exit Sub
            End If
End Select

End Sub

Private Sub DBGrid4_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo cmd57_err
If KeyCode = &H2E Then  'borrar linea
   If MsgBox("Desea Borrar ", 1, "Borrar") <> 1 Then Exit Sub
   Data5.Recordset.Delete
   Data5.Refresh
   Exit Sub
End If
Exit Sub
cmd57_err:
Data5.Refresh
Exit Sub

End Sub

Private Sub dbgrid5_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex <> 1 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
       Case 1
            If Len("" & DBGrid5.Columns(0)) = 0 Then
               Cancel = True
               Exit Sub
            End If
End Select

End Sub

Private Sub DBGrid5_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
If ColIndex <> 1 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
       Case 1
            'MsgBox "" & DBGrid3.Columns(1)
            If Not IsNumeric("" & DBGrid5.Columns(1)) Then
               Cancel = True
               Exit Sub
            End If
End Select

End Sub

Private Sub DBGrid5_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo cmd59_err
If KeyCode = &H2E Then  'borrar linea
   If MsgBox("Desea Borrar ", 1, "Borrar") <> 1 Then Exit Sub
   Data4.Recordset.Delete
   Data4.Refresh
   Exit Sub
End If
Exit Sub
cmd59_err:
Data4.Refresh
Exit Sub

End Sub

Private Sub dflo3433_Click()
If Frame3.Visible = True Then
   Frame3.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If

If Frame1.Visible = True Then
   If opcion1 = "1" Then
   Frame1.Visible = False
   codigo.SetFocus
   Exit Sub
   End If
   If opcion1 = "2" Then
   Frame1.Visible = False
   DBGrid3.SetFocus
   Exit Sub
   End If
   If opcion1 = "3" Then
   Frame1.Visible = False
   DBGrid4.SetFocus
   Exit Sub
   End If
   If opcion1 = "4" Then
   Frame1.Visible = False
   DBGrid4.SetFocus
   Exit Sub
   End If
End If

If Frame2.Visible = True Then
   Frame2.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If
planillv.Hide
Unload planillv
End Sub

Private Sub dki3243_Click()
If Frame1.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
Frame3.Visible = True
codigo.SetFocus
End Sub

Private Sub Form_Activate()
periodo = Format(Month(Now), "00") + Format(Year(Now), "0000")
carga_inicial
consulta_sql
End Sub

Sub consulta_sql()
Dim buf As String
buf = "select * from vendedor where "
buf = buf & "   codigo like '" & codigo & "'"
If NOMBRE <> "%" Then
buf = buf & "  and  nombre like  '" & NOMBRE & "'"
End If
If division <> "%" Then
buf = buf & "  and  division like  '" & division & "'"
End If

               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = buf
               Data2.Refresh
               DBGrid2.SetFocus
End Sub

Private Sub Image1_Click()

End Sub
Sub consulta_planilla()
Dim buf As String
buf = "select * from remune02 where "
buf = buf & " tipopla='" & tipopla & "'"
buf = buf & " and codigo='" & xcodigo & "'"
buf = buf & " and periodo='" & periodo & "'"
               Data3.Connect = "foxpro 2.5;"
               Data3.DatabaseName = globaldir
               Data3.RecordSource = buf
               Data3.Refresh

buf = "select * from descue02 where "
buf = buf & " tipopla='" & tipopla & "'"
buf = buf & " and codigo='" & xcodigo & "'"
buf = buf & " and periodo='" & periodo & "'"
               Data4.Connect = "foxpro 2.5;"
               Data4.DatabaseName = globaldir
               Data4.RecordSource = buf
               Data4.Refresh
               

buf = "select * from aporta02 where "
buf = buf & " tipopla='" & tipopla & "'"
buf = buf & " and codigo='" & xcodigo & "'"
buf = buf & " and periodo='" & periodo & "'"
               Data5.Connect = "foxpro 2.5;"
               Data5.DatabaseName = globaldir
               Data5.RecordSource = buf
               Data5.Refresh
               
               


End Sub

Private Sub Label1_Click()
consulta_vendedor
End Sub
Sub consulta_vendedor()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Codigo"
Combo1.ListIndex = 0

Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "1"
Command1_Click

End Sub
Sub consulta_1()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Tipo"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "2"
Command1_Click

End Sub
Sub consulta_2()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Tipo"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "3"
Command1_Click

End Sub
Sub consulta_3()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Tipo"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "4"
Command1_Click

End Sub
Function graba_ingreso(buf As String, sw As Integer)

Dim mytablex As Table
Dim found As Integer

Set mytablex = mydbxglo.OpenTable("tplanico")
mytablex.Index = "tplanico"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
      If sw = 1 Then
      Data3.Recordset.AddNew
      Data3.Recordset.Fields("codigo") = "" & xcodigo
      Data3.Recordset.Fields("tipo") = "" & mytablex.Fields("tipo")
      Data3.Recordset.Fields("concepto") = "" & mytablex.Fields("descripcio")
      'Data3.Recordset.Fields("porcentaje") = "" & mytablex.Fields("porcentaje")
      Data3.Recordset.Update
      End If
      If sw = 2 Then
      Data4.Recordset.AddNew
      Data4.Recordset.Fields("codigo") = "" & xcodigo
      Data4.Recordset.Fields("tipo") = "" & mytablex.Fields("tipo")
      Data4.Recordset.Fields("concepto") = "" & mytablex.Fields("descripcio")
      'Data4.Recordset.Fields("porcentaje") = "" & mytablex.Fields("porcentaje")
      Data4.Recordset.Update
      End If
      If sw = 3 Then
      Data5.Recordset.AddNew
      Data5.Recordset.Fields("codigo") = "" & xcodigo
      Data5.Recordset.Fields("tipo") = "" & mytablex.Fields("tipo")
      Data5.Recordset.Fields("concepto") = "" & mytablex.Fields("descripcio")
      'Data5.Recordset.Fields("porcentaje") = "" & mytablex.Fields("porcentaje")
      Data5.Recordset.Update
      End If
      graba_ingreso = 1
End If
'------------------------------------- ------------
mytablex.Close
 
End Function

Private Sub Label3_Click()
'consulta_1
End Sub

Private Sub Label4_Click()
'consulta_2
End Sub

Private Sub Label5_Click()
'consulta_3
End Sub
Sub recalculo_descuento()

Dim mytablex As Table
Dim mytabley As Table
Dim v_imp_con As Double
Dim v_importe As Double

   Set mytablex = mydbxglo.OpenTable("tplanic1")
   mytablex.Index = "tplanic"
   Set mytabley = mydbxglo.OpenTable("remune01")
   mytabley.Index = "remune01"
   v_importe = 0
   ir_inicio 5
Do
   If Data5.Recordset.EOF Then Exit Do
   mytablex.Seek "=", "" & Data5.Recordset.Fields("tipo")
   If Not mytablex.NoMatch Then
      
      Do
      If mytablex.EOF Then Exit Do
      If "" & mytablex.Fields("tipo") = "" & Data5.Recordset.Fields("tipo") Then
         'buscar en remuneraciones----------------------------------------------------
         mytabley.Seek "=", "" & xcodigo, "" & mytablex.Fields("codigo")
          If Not mytabley.NoMatch Then
            v_imp_con = (Val("" & mytablex.Fields("porcentaje")) / 100#) * Val("" & mytabley.Fields("importe"))
            v_importe = v_importe + v_imp_con
         End If
         '------------------------------------------------------
         Else
         Exit Do
      End If
      mytablex.MoveNext
      Loop
      If v_importe > 0 Then
         Data5.Recordset.Edit
         Data5.Recordset.Fields("importe") = Val(Format(v_importe, "0.00"))
         Data5.Recordset.Update
      End If
   End If
   Data5.Recordset.MoveNext
Loop
   mytabley.Close
   mytablex.Close
    


End Sub
Sub recalculo_aportacion()

Dim mytablex As Table
Dim mytabley As Table
Dim v_imp_con As Double
Dim v_importe As Double

   Set mytablex = mydbxglo.OpenTable("tplanic1")
   mytablex.Index = "tplanic"
   Set mytabley = mydbxglo.OpenTable("remune01")
   mytabley.Index = "remune01"
   v_importe = 0
   ir_inicio 4
Do
   If Data4.Recordset.EOF Then Exit Do
   mytablex.Seek "=", "" & Data4.Recordset.Fields("tipo")
   If Not mytablex.NoMatch Then
      
      Do
      If mytablex.EOF Then Exit Do
      If "" & mytablex.Fields("tipo") = "" & Data4.Recordset.Fields("tipo") Then
         'buscar en remuneraciones----------------------------------------------------
         mytabley.Seek "=", "" & xcodigo, "" & mytablex.Fields("codigo")
         If Not mytabley.NoMatch Then
            v_imp_con = (Val("" & mytablex.Fields("porcentaje")) / 100#) * Val("" & mytabley.Fields("importe"))
            v_importe = v_importe + v_imp_con
         End If
         '------------------------------------------------------
         Else
         Exit Do
      End If
      mytablex.MoveNext
      Loop
      If v_importe > 0 Then
         Data4.Recordset.Edit
         Data4.Recordset.Fields("importe") = Val(Format(v_importe, "0.00"))
         Data4.Recordset.Update
      End If
   End If
   Data4.Recordset.MoveNext
Loop
   mytabley.Close
   mytablex.Close
    


End Sub
Sub recalculo_basico()

Dim mytablex As Table
Dim mytabley As Table
Dim v_imp_con As Double
Dim v_importe As Double

   Set mytablex = mydbxglo.OpenTable("tplanic1")
   mytablex.Index = "tplanic"
   Set mytabley = mydbxglo.OpenTable("remune01")
   mytabley.Index = "remune01"
   v_importe = 0
   ir_inicio 3
Do
   If Data3.Recordset.EOF Then Exit Do
   mytablex.Seek "=", "" & Data3.Recordset.Fields("tipo")
   If Not mytablex.NoMatch Then
      Do
      If mytablex.EOF Then Exit Do
      If "" & mytablex.Fields("tipo") = "" & Data3.Recordset.Fields("tipo") Then
         'buscar en remuneraciones----------------------------------------------------
         mytabley.Seek "=", "" & xcodigo, "" & mytablex.Fields("codigo")
         If Not mytabley.NoMatch Then
            v_imp_con = (Val("" & mytablex.Fields("porcentaje")) / 100#) * Val("" & mytabley.Fields("importe"))
            v_importe = v_importe + v_imp_con
         End If
         '------------------------------------------------------
         Else
         Exit Do
      End If
      mytablex.MoveNext
      Loop
      If v_importe > 0 Then
         Data3.Recordset.Edit
         Data3.Recordset.Fields("importe") = Val(Format(v_importe, "0.00"))
         Data3.Recordset.Update
      End If
   End If
   Data3.Recordset.MoveNext
Loop
   mytabley.Close
   mytablex.Close
    

End Sub
Sub ir_inicio(buf As Integer)
On Error GoTo cmd99_err
Select Case buf
       Case 5
            Data5.Recordset.MoveFirst
       Case 4
          Data4.Recordset.MoveFirst
       Case 3
          Data3.Recordset.MoveFirst
End Select
Exit Sub
cmd99_err:
Exit Sub
End Sub
Function consulta_repitencia2(buf As String) 'ingreso
Dim mytabley As Table


   Set mytabley = mydbxglo.OpenTable("remune01")
   mytabley.Index = "remune01"
   mytabley.Seek "=", "" & xcodigo, buf
   If Not mytabley.NoMatch Then
      consulta_repitencia2 = 1
   End If
   mytabley.Close
    
End Function
Function consulta_repitencia3(buf As String) 'aporta
Dim mytabley As Table


   Set mytabley = mydbxglo.OpenTable("aporta01")
   mytabley.Index = "aporta01"
   mytabley.Seek "=", "" & xcodigo, buf
   If Not mytabley.NoMatch Then
      consulta_repitencia3 = 1
   End If
   mytabley.Close
    
End Function
Function consulta_repitencia4(buf As String) 'descue
Dim mytabley As Table

   
   Set mytabley = mydbxglo.OpenTable("descue01")
   mytabley.Index = "descue01"
   mytabley.Seek "=", "" & xcodigo, buf
   If Not mytabley.NoMatch Then
      consulta_repitencia4 = 1
   End If
   mytabley.Close
    
End Function
Sub suma_ingreso()
Dim sdx As Double
ir_inicio 3
sdx = 0
Do
If Data3.Recordset.EOF Then Exit Do
sdx = sdx + Val("" & Data3.Recordset.Fields("importe"))
Data3.Recordset.MoveNext
Loop
totingreso = Format(sdx, "0.00")

End Sub
Sub suma_descuento()
Dim sdx As Double
ir_inicio 5
sdx = 0
Do
If Data5.Recordset.EOF Then Exit Do
sdx = sdx + Val("" & Data5.Recordset.Fields("importe"))
Data5.Recordset.MoveNext
Loop
totdscto = Format(sdx, "0.00")

End Sub
Sub suma_aporta()
Dim sdx As Double
ir_inicio 4
sdx = 0
Do
If Data4.Recordset.EOF Then Exit Do
sdx = sdx + Val("" & Data4.Recordset.Fields("importe"))
Data4.Recordset.MoveNext
Loop
totaporte = Format(sdx, "0.00")

End Sub
Sub suma_total()
Dim sdx As Double
suma_descuento
suma_ingreso
suma_aporta
sdx = Val(totingreso) - Val(totdscto) - Val(totaporte)
totalcobrar = Format(sdx, "0.00")
End Sub
Function ver_sisper(buf As String)
Dim mytabley As Table

   horatraba = ""
   horaextr = ""
   diatraba = ""
   
   Set mytabley = mydbxglo.OpenTable("sisper01")
   mytabley.Index = "sisper01"
   mytabley.Seek "=", tipopla, buf
   If Not mytabley.NoMatch Then
      horatraba = "" & mytabley.Fields("horatraba")
      horaextr = "" & mytabley.Fields("horaextr")
      diatraba = "" & mytabley.Fields("diatraba")
   End If
   mytabley.Close
    
End Function
Sub carga_inicial()
Dim mytablex As Table

tipopla1.Clear
tipopla1.AddItem "%"

   Set mytablex = mydbxglo.OpenTable("tipopla")
   Do
   If mytablex.EOF Then Exit Do
   tipopla1.AddItem "" & mytablex.Fields("tipopla")
   mytablex.MoveNext
   Loop
   mytablex.Close
    
   tipopla1.ListIndex = 0
End Sub





Private Sub plai8232_Click()
DBGrid2_DblClick
End Sub
