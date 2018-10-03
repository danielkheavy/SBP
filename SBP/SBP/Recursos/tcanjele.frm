VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tcanjele 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Canje Letras"
   ClientHeight    =   9210
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   9615
      Left            =   8280
      TabIndex        =   36
      Top             =   5880
      Visible         =   0   'False
      Width           =   13215
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6735
         Left            =   240
         TabIndex        =   40
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ingreso de Letras"
      Height          =   5175
      Left            =   480
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   12255
      Begin VB.CommandButton Command2 
         Caption         =   "Sumar"
         Height          =   735
         Left            =   11400
         TabIndex        =   35
         Top             =   2760
         Width           =   735
      End
      Begin VB.CommandButton cmdAddEntry 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Limpia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11400
         Picture         =   "tcanjele.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Nuevo registro"
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Graba"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11400
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tcanjele.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Grabar registro"
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11400
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tcanjele.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Borrar registro"
         Top             =   2040
         Width           =   735
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "tcanjele.frx":3636
         Height          =   3015
         Left            =   120
         OleObjectBlob   =   "tcanjele.frx":364A
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1200
         Width           =   11055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
      Begin VB.Label codigo 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   30
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   855
      End
      Begin VB.Label nombre 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   28
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label total 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5760
         TabIndex        =   26
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   25
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label moneda 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   24
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Letra"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   23
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label letrato 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6960
         TabIndex        =   22
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label vendedor 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   21
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "F1 Seccion,Aceptante,Girador"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   4200
         Width           =   4935
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Generar"
      BeginProperty Font 
         Name            =   "Verdana"
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
      Picture         =   "tcanjele.frx":4EFD
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   1335
   End
   Begin VB.TextBox numeroi 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   4320
      MaxLength       =   11
      TabIndex        =   14
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox cant 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3480
      MaxLength       =   3
      TabIndex        =   12
      Top             =   1080
      Width           =   855
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   13680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   13680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox numero 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      MaxLength       =   11
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox serie 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   960
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox tipo 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NroLetraInI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Loc"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cuo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   720
      Width           =   495
   End
   Begin VB.Label local1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label cuota 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ojo.La factura deben encontrarse en cartera"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label acu 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8040
      TabIndex        =   6
      Top             =   720
      Width           =   105
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro Documento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin VB.Menu loasdlo23 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcanjele"
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
   loasdlo23_Click
   Exit Sub
End If
Command1_Click

End Sub

Private Sub cant_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
numeroi.SetFocus
End Sub

Private Sub cmdAddEntry_Click()
Command4_Click
End Sub

Private Sub cmdDelete_Click()
Frame2.Visible = False
End Sub

Private Sub cmdSave_Click()
Dim buf As String
Dim sw As Integer
Dim i As Integer
Dim mytablex As New adodb.Recordset
sumar_letra
If Val(letrato) = 0 Then
   MsgBox "Total debe ser >0", 48, "Aviso"
   dbgrid2.SetFocus
   Exit Sub
End If
If Val(letrato) <> Val(total) Then
   MsgBox "Total debe ser Igual al total factura ", 48, "Aviso"
   dbgrid2.SetFocus
   Exit Sub
End If
found = valida_tmp()
If found = 1 Then
   MsgBox "Ingreso de letra no Valido ", 48, "Aviso"
   Exit Sub
End If
If MsgBox("Desea grabar", 1, "Aviso") <> 1 Then Exit Sub
sw = 0
If acu = "V" Then
buf = "letrav"
'Set mytablex = mydbxglo.OpenTable("letrav")
End If
If acu = "C" Then
buf = "letrac"
'Set mytablex = mydbxglo.OpenTable("letrac")
End If
mytablex.Index = "letra"
Data2.Refresh
Set rs = Data2.Recordset.Clone
Do
If rs.EOF Then Exit Do
   If Len("" & rs.Fields("letra")) > 0 And Val("" & rs.Fields("importe")) > 0 Then
   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "select * from " & buf & " where letra='" & rs.Fields("letra") & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
   mytablex.AddNew
   mytablex.Fields("letra") = "" & rs.Fields("letra")
   mytablex.Fields("fechai") = "" & rs.Fields("fechai")
   mytablex.Fields("fechaf") = "" & rs.Fields("fechaf")
   mytablex.Fields("importe") = Val("" & rs.Fields("importe"))
   mytablex.Fields("saldo") = Val("" & rs.Fields("importe"))
   mytablex.Fields("moneda") = "" & rs.Fields("moneda")
   mytablex.Fields("observa") = "" & rs.Fields("observa")
   mytablex.Fields("abono") = 0
   mytablex.Fields("paridad") = 1
   mytablex.Fields("aceptante") = codigo
   mytablex.Fields("girador") = "" & rs.Fields("girador")
   mytablex.Fields("vendedor") = vendedor
   mytablex.Fields("banco") = "" & rs.Fields("banco")
   mytablex.Fields("seccion") = "" & rs.Fields("seccion")
   mytablex.Fields("agencia") = ""
   mytablex.Fields("refactura") = serie & "-" & numero
   mytablex.Fields("estado") = "0"
   mytablex.Fields("estador") = "0"
   mytablex.Fields("estadop") = "0"
   mytablex.Fields("nombreg") = nombre
   mytablex.Fields("nombrea") = nombre
   mytablex.Fields("nrounico") = "" & rs.Fields("nrounico")
   mytablex.Fields("ochodia") = rs.Fields("ochodia")
   mytablex.Update
   sw = 1
   End If
   End If
   rs.MoveNext
Loop
mytablex.Close
If sw = 1 Then
   Borrar_cuentac
End If
MsgBox "Proceso Terminado", 48, "Aviso"
Frame2.Visible = False
tipo.SetFocus

End Sub
Function valida_tmp()
Dim sdx As Integer
Dim found As Integer
Data2.Refresh
sdx = 0
Set rs = Data2.Recordset.Clone
Do
If rs.EOF Then Exit Do
   found = busca_letra("" & rs.Fields("letra"))
   If found = 1 Then
      sdx = 1
   End If
   If Len("" & rs.Fields("letra")) = 0 Then
      sdx = 1
   End If
   If Val("" & rs.Fields("importe")) <= 0 Then
      sdx = 1
   End If
   If Not IsDate("" & rs.Fields("fechai")) Then
      sdx = 1
   End If
   If Not IsDate("" & rs.Fields("fechaf")) Then
      sdx = 1
   End If
   rs.MoveNext
Loop
valida_tmp = sdx

End Function

Private Sub Command1_Click()
Dim mytablex As New adodb.Recordset
Dim buf As String
Dim buf3 As String
Dim buf2 As String
Dim buf4 As String
Dim quiebre As String
buf4 = ""
If acu = "C" Then
   quiebre = "  (tipodoc='J' or tipodoc='K' or tipodoc='L' or tipodoc='M' or tipodoc='P') "
End If
If acu = "V" Then
   quiebre = " (tipodoc='A' or tipodoc='B' or tipodoc='C' or tipodoc='D' or tipodoc='G') "
End If

If opcion1 = "4" Or opcion1 = "5" Then
   If acu = "C" Then
      buf4 = "proveedo"
   End If
   If acu = "V" Then
     buf4 = "CLIENTES"
   End If
End If
If opcion1 = "2" Then
   If acu = "C" Then
      buf2 = "cuentap"
   End If
   If acu = "V" Then
     buf2 = "cuentac"
   End If
End If
If opcion1 = "2" Then
   buf3 = " tipo='" & tipo & "'"
End If
   If opcion1 = "4" Then
      If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from  " & buf4
      Else
      buf = "select Nombre,Codigo from  " & buf4 & " where " & Combo1 & " like '" & buffer & "%'"
      End If
   End If
   If opcion1 = "5" Then
      If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from  " & buf4
      Else
      buf = "select Nombre,Codigo from  " & buf4 & " where " & Combo1 & " like '" & buffer & "%'"
      End If
   End If
   If opcion1 = "6" Then
      If Len(buffer) = 0 Then
      buf = "select Descripcio,Banco from  banco"
      Else
      buf = "select Banco,Banco from  banco where " & Combo1 & " like '" & buffer & "%'"
      End If
   End If
   If opcion1 = "3" Then
      If Len(buffer) = 0 Then
      buf = "select Descripcio,Carsec from   carsec"
      Else
      buf = "select Descripcio,Carsec from carsec where " & Combo1 & " like '" & buffer & "%'"
      End If
   End If
   
   If opcion1 = "1" Then
      If Len(buffer) = 0 Then
      buf = "select Descripcio,Tipo from   Tipo where " & quiebre
      Else
      buf = "select Descripcio,Tipo from tipo where " & quiebre & " and  " & Combo1 & " like '" & buffer & "%'"
      End If
   End If
   If opcion1 = "2" Then
      If Len(buffer) = 0 Then
      buf = "select Tipo,Serie,Numero,Cuota,Local,Fecha,Codigo,Nombre,Saldo,Total,Estado from  " & buf2 & " where " & buf3
      Else
      buf = "select Tipo,Serie,Numero,Cuota,Local,Fecha,Codigo,Nombre,Saldo,Total,Estado from  " & buf2 & " where " & buf3 & " and " & Combo1 & " like '" & buffer & "%'"
      End If
   End If
   'MsgBox buf
   
   mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
   Set dbGrid1.DataSource = mytablex
   If mytablex.RecordCount = 0 Then
      mytablex.Close
      buffer.SetFocus
      Exit Sub
   End If
   
   
               If opcion1 = "2" Then
               dbGrid1.columns(0).Width = 600
               dbGrid1.columns(1).Width = 600
               dbGrid1.columns(2).Width = 1500
               dbGrid1.columns(3).Width = 700
               dbGrid1.columns(4).Width = 700
               dbGrid1.columns(5).Width = 1400
               dbGrid1.columns(6).Width = 1500
               dbGrid1.columns(7).Width = 2500
               dbGrid1.columns(8).Width = 1500
               dbGrid1.columns(9).Width = 1500
               dbGrid1.columns(10).Width = 700
               End If
               If opcion1 = "1" Then
               dbGrid1.columns(0).Width = 4000
               dbGrid1.columns(1).Width = 2000
               End If
               dbGrid1.SetFocus

End Sub

Private Sub Command2_Click()
sumar_letra
End Sub

Private Sub Command4_Click()
Dim found As Integer
Dim sdx As Double
Dim i As Integer
found = busca_tipo()
If found = 0 Then
   tipo.SetFocus
   Exit Sub
End If
If Len(numero) = 0 Then
   tipo.SetFocus
   Exit Sub
End If
If Len(cant) = 0 Or Not IsNumeric(cant) Then
   cant.SetFocus
   Exit Sub
End If
If Len(numeroi) = 0 Or Not IsNumeric(numeroi) Then
   numeroi.SetFocus
   Exit Sub
End If
found = busca_factura()
If found = 0 Then
   serie.SetFocus
   Exit Sub
End If
If found = 2 Then
   MsgBox "Saldo debe ser Mayor a Cero", 48, "Aviso"
   serie.SetFocus
   Exit Sub
End If

borrar_temporal
sql_temporal
sdx = Val(numeroi)
For i = 1 To Val(cant)
    found = busca_letra("" & sdx)
    Data2.Recordset.AddNew
    Data2.Recordset.Fields("letra") = Format(sdx, "000")
    Data2.Recordset.Update
    sdx = sdx + 1
Next i
Frame2.Visible = True
dbgrid2.SetFocus
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim found As Integer
Dim buf As String
Dim xtemp As Variant
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
If opcion1 = "6" Then
   dbgrid2.columns(8) = dbGrid1.columns(1)
   Frame1.Visible = False
   Frame1.Enabled = False
   dbgrid2.SetFocus
End If
If opcion1 = "5" Then
   dbgrid2.columns(7) = dbGrid1.columns(1)
   Frame1.Visible = False
   Frame1.Enabled = False
   dbgrid2.SetFocus
End If
If opcion1 = "4" Then
   dbgrid2.columns(6) = dbGrid1.columns(1)
   Frame1.Visible = False
   Frame1.Enabled = False
   dbgrid2.SetFocus
End If
If opcion1 = "3" Then
   dbgrid2.columns(5) = dbGrid1.columns(1)
   Frame1.Visible = False
   Frame1.Enabled = False
   dbgrid2.SetFocus
End If
If opcion1 = "1" Then
   tipo = dbGrid1.columns(1)
   Frame1.Visible = False
   Frame1.Enabled = False
   tipo.SetFocus
   tipo_KeyPress 13
End If
If opcion1 = "2" Then
   serie = dbGrid1.columns(1)
   numero = dbGrid1.columns(2)
   cuota = dbGrid1.columns(3)
   local1 = dbGrid1.columns(4)
   Frame1.Visible = False
   Frame1.Enabled = False
   serie.SetFocus
   serie_KeyPress 13
End If
End If

End Sub

Private Sub DBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
Select Case ColIndex
       Case 4
          If Len("" & dbgrid2.columns(0)) = 0 Then
             Cancel = True
             Exit Sub
          End If
          If Len("" & dbgrid2.columns(5)) = 0 Then
             Cancel = True
             Exit Sub
          End If
          dbgrid2.Col = 0
          dbgrid2.Row = dbgrid2.VisibleRows - 1
          dbgrid2.SetFocus
End Select
End Sub

Private Sub DBGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Dim found As Integer
If KeyAscii = 27 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
     Case 5, 6, 7, 8, 10
          Cancel = True
          Exit Sub
     Case 1
          If Len("" & dbgrid2.columns(0)) = 0 Then
             Cancel = True
             Exit Sub
          End If
     Case 2
          If Len("" & dbgrid2.columns(0)) = 0 Then
             Cancel = True
             Exit Sub
          End If
          If Len("" & dbgrid2.columns(1)) = 0 Then
             Cancel = True
             Exit Sub
          End If
     Case 3
          If Len("" & dbgrid2.columns(0)) = 0 Then
             Cancel = True
             Exit Sub
          End If
     Case 4
          If Len("" & dbgrid2.columns(0)) = 0 Then
             Cancel = True
             Exit Sub
          End If
     Case 5
          If Len("" & dbgrid2.columns(0)) = 0 Then
             Cancel = True
             Exit Sub
          End If
     Case 6
          If Len("" & dbgrid2.columns(0)) = 0 Then
             Cancel = True
             Exit Sub
          End If
     Case 9
          If Len("" & dbgrid2.columns(0)) = 0 Then
             Cancel = True
             Exit Sub
          End If
          
          
          
          
End Select
          
End Sub

Private Sub DBGrid2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim found As Integer
Select Case ColIndex
     Case 0
          If Len("" & dbgrid2.columns(0)) = 0 Then
             Cancel = True
             Exit Sub
          End If
          found = busca_tmpcanle("" & dbgrid2.columns(0))
          If found = 1 Then
             MsgBox "Letra ya Existente", 48, "Aviso"
             Cancel = True
             Exit Sub
          End If
          
          
          
          found = busca_letra("" & dbgrid2.columns(0))
          If found = 1 Then
             MsgBox "Letra ya Existente", 48, "Aviso"
             Cancel = True
             Exit Sub
          End If
          
     Case 1
          If Len("" & dbgrid2.columns(1)) = 0 Then
             Cancel = True
             Exit Sub
          End If
          
          If Len("" & dbgrid2.columns(1)) <> 10 Then
             Cancel = True
             Exit Sub
          End If
          If Val(Mid$("" & dbgrid2.columns(1), 1, 2)) < 0 Or Val(Mid$("" & dbgrid2.columns(1), 1, 2)) > 31 Then
             Cancel = True
             Exit Sub
          End If
          If Val(Mid$("" & dbgrid2.columns(1), 3, 2)) < 0 Or Val(Mid$("" & dbgrid2.columns(1), 3, 2)) > 12 Then
             Cancel = True
             Exit Sub
          End If
          If Not IsDate("" & dbgrid2.columns(1)) Then
             Cancel = True
             Exit Sub
          End If
    Case 2
          
          If Len("" & dbgrid2.columns(2)) <> 10 Then
             Cancel = True
             Exit Sub
          End If
          If Val(Mid$("" & dbgrid2.columns(2), 1, 2)) < 0 Or Val(Mid$("" & dbgrid2.columns(2), 1, 2)) > 31 Then
             Cancel = True
             Exit Sub
          End If
          If Val(Mid$("" & dbgrid2.columns(2), 3, 2)) < 0 Or Val(Mid$("" & dbgrid2.columns(2), 3, 2)) > 12 Then
             Cancel = True
             Exit Sub
          End If
          If Not IsDate("" & dbgrid2.columns(2)) Then
             Cancel = True
             Exit Sub
          End If
    Case 3
        If "" & dbgrid2.columns(3) <> "S" And "" & dbgrid2.columns(3) <> "D" Then
           Cancel = True
           Exit Sub
        End If
   Case 4
        If Val("" & dbgrid2.columns(4)) = 0 Then
           Cancel = True
           Exit Sub
        End If
        If Not IsNumeric(dbgrid2.columns(4)) Then
           Cancel = True
           Exit Sub
        End If
        'DBGrid2.Columns(4) = Format(Val("" & DBGrid2.Columns(4)), "0.00")
   Case 9
        If Len("" & dbgrid2.columns(9)) > 0 Then
           dbgrid2.columns(10) = "" & (CVDate(dbgrid2.columns(2)) + 8)
        End If
End Select

End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H2E Then  'borrar linea
End If


If KeyCode = &H70 And dbgrid2.Col = 5 Then 'f1
   If Len(dbgrid2.columns(0)) = 0 Then Exit Sub
      consulta_cartera
End If
If KeyCode = &H70 And dbgrid2.Col = 6 Then 'f1 aceptante
   If Len(dbgrid2.columns(0)) = 0 Then Exit Sub
      consulta_aceptante
End If
If KeyCode = &H70 And dbgrid2.Col = 7 Then 'f1 girador
   If Len(dbgrid2.columns(0)) = 0 Then Exit Sub
      consulta_girador
End If
If KeyCode = &H70 And dbgrid2.Col = 8 Then 'f1 banco
   If Len(dbgrid2.columns(0)) = 0 Then Exit Sub
      consulta_banco
End If



End Sub

Private Sub Form_Load()
borrar_temporal
sql_temporal

End Sub

Private Sub Label6_Click()
sumar_letra
End Sub

Private Sub letrato_Click()
sumar_letra
End Sub

Private Sub loasdlo23_Click()
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then
   If opcion1 = "1" Then
      Frame1.Visible = False
      Frame1.Enabled = False
      serie.SetFocus
      Exit Sub
   End If
   If opcion1 = "2" Or opcion1 = "1" Or opcion1 = "13" Or opcion1 = "3" Or opcion1 = "4" Or opcion1 = "5" Or opcion1 = "6" Then
      Frame1.Visible = False
      Frame1.Enabled = False
      'dbgrid2.SetFocus
      Exit Sub
   End If

End If
cerrar_datas 1
cerrar_datas 2
tcanjele.Hide
Unload tcanjele
End Sub
Sub consulta_tipo()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Tipo"
Combo1.ListIndex = 0
Frame1.Enabled = True
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "1"
Command1_Click
End Sub
Sub consulta_documento()
Combo1.Clear

Combo1.AddItem "Tipo"
Combo1.AddItem "Serie"
Combo1.AddItem "Numero"
Combo1.AddItem "Nombre"
Combo1.ListIndex = 0
Frame1.Enabled = True
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "2"
Command1_Click
End Sub
Sub consulta_cartera()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.ListIndex = 0
Frame1.Enabled = True
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "3"
Command1_Click

End Sub
Sub consulta_aceptante()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.ListIndex = 0
Frame1.Enabled = True
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "4"
Command1_Click

End Sub
Sub consulta_girador()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.ListIndex = 0
Frame1.Enabled = True
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "5"
Command1_Click
End Sub
Sub consulta_banco()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.ListIndex = 0
Frame1.Enabled = True
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "6"
Command1_Click

End Sub




Private Sub numeroi_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub serie_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then
   serie.SetFocus
   Exit Sub
End If
found = busca_factura()
If found = 0 Then
   serie.SetFocus
   Exit Sub
End If
If found = 2 Then
   MsgBox "Saldo debe ser Mayor a Cero", 48, "Aviso"
   serie.SetFocus
   Exit Sub
End If
cant.SetFocus
End Sub

Private Sub serie_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_documento
End If

End Sub

Private Sub tipo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
found = busca_tipo()
If found = 0 Then Exit Sub
serie.SetFocus
End Sub

Private Sub tipo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_tipo
End If
End Sub
Function busca_tipo()
Dim mytablex As New adodb.Recordset

mytablex.Open "select * from tipo where tipo='" & tipo & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If acu = "V" Then
      If "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "D" Or "" & mytablex.Fields("tipodoc") = "G" Then
         busca_tipo = 1
      End If
   End If
   If acu = "C" Then
      If "" & mytablex.Fields("tipodoc") = "J" Or "" & mytablex.Fields("tipodoc") = "K" Or "" & mytablex.Fields("tipodoc") = "L" Or "" & mytablex.Fields("tipodoc") = "M" Or "" & mytablex.Fields("tipodoc") = "P" Then
         busca_tipo = 1
      End If
   End If
End If
mytablex.Close
End Function
Sub habilita(sw As Integer)
Dim xsw
If sw = 0 Then
   xsw = True
End If
If sw = 1 Then
   xsw = False
End If
tipo.Enabled = xsw
serie.Enabled = xsw
End Sub
Function busca_factura()
Dim buf As String
Dim mytablex As New adodb.Recordset
If acu = "V" Then
buf = "cuentac"

End If
If acu = "C" Then
buf = "cuentap"
End If

mytablex.Open "select * from " & buf & " where local='" & local1 & "' and tipo='" & tipo & "' and serie='" & serie & "' and numero='" & numero & "' and cuota='" & cuota & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   busca_factura = 2
   If Val("" & mytablex.Fields("saldo")) > 0 Then
      busca_factura = 1
      codigo = "" & mytablex.Fields("codigo")
      nombre = "" & mytablex.Fields("nombre")
      total = "" & mytablex.Fields("saldo")
      moneda = "" & mytablex.Fields("moneda")
      vendedor = "" & mytablex.Fields("vendedor")
   End If
End If
'------------------------------------- ------------
mytablex.Close
 

End Function
Sub borrar_temporal()
mydbxglo.Execute "DELETE FROM tmpcanle "

End Sub
Sub sql_temporal()
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = "select * from tmpcanle"
               Data2.Refresh
End Sub
Function busca_tmpcanle(buf As String)
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("tmpcanle")
mytablex.Index = "tmpcanle"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_tmpcanle = 1
End If
mytablex.Close
End Function
Function busca_letra(buf As String)
Dim buf1 As String
Dim mytablex As New adodb.Recordset
If acu = "V" Then
buf1 = "letrav"
'Set mytablex = mydbxglo.OpenTable("letrav")
End If
If acu = "C" Then
buf1 = "letrac"
'Set mytablex = mydbxglo.OpenTable("letrac")
End If
mytablex.Open "select * from " & buf1 & " where letra='" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   busca_letra = 1
End If
'------------------------------------- ------------
mytablex.Close
 
End Function
Sub sumar_letra()
Dim fila As Integer
Dim suma As Double
On Error GoTo cmd4_err
suma = 0
For fila = 0 To Data2.Recordset.RecordCount - 1
dbgrid2.Row = fila    'El índice de la primera fila empieza en 0.
suma = suma + Val("" & dbgrid2.columns(4).Value)
Next
letrato = Format(suma, "0.00")
Exit Sub
cmd4_err:
Exit Sub
End Sub
Sub Borrar_cuentac()
Dim mytablex As New adodb.Recordset
Dim buf As String
If acu = "V" Then
buf = "cuentac"
'Set mytablex = mydbxglo.OpenTable("cuentac")
End If
If acu = "C" Then
buf = "cuentap"
'Set mytablex = mydbxglo.OpenTable("cuentap")
End If
mytablex.Open "select * from " & buf & " where local='" & local1 & "' and serie='" & serie & "' and numero='" & numero & "' and cuota='" & cuota & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   'mytablex.Edit
   mytablex.Fields("abono") = Val("" & mytablex.Fields("abono")) + Val(letrato)
   mytablex.Fields("saldo") = Val("" & mytablex.Fields("total")) + Val("" & mytablex.Fields("interes")) - Val("" & mytablex.Fields("abono"))
   mytablex.Update
End If
mytablex.Close
End Sub
Sub cerrar_datas(sw As Integer)
On Error GoTo cmd1_err
Select Case sw
  Case 1
       Data1.Recordset.Close
  Case 2
       Data2.Recordset.Close
End Select
Exit Sub
cmd1_err:
Exit Sub
       
End Sub


