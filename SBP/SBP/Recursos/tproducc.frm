VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form tproducc 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mapa de Produccion"
   ClientHeight    =   11445
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11445
   ScaleWidth      =   14865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Ingreso de Lineas"
      Height          =   3255
      Left            =   3600
      TabIndex        =   22
      Top             =   6840
      Visible         =   0   'False
      Width           =   7335
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
         Left            =   5400
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tproducc.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Borrar registro"
         Top             =   2400
         Width           =   735
      End
      Begin VB.CommandButton Command2 
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
         Left            =   6240
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tproducc.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Grabar registro"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox t16 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t15 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t14 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t13 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t12 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t11 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t10 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t9 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t8 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t7 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t6 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t5 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t4 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t3 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t2 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t1 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.Label linea 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4440
         TabIndex        =   63
         Top             =   360
         Width           =   855
      End
      Begin VB.Label nt16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   62
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   61
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   60
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   59
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   58
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   57
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   56
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   55
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label30 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   54
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label29 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   53
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   52
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   51
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   50
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   49
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   48
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   47
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   46
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   45
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   44
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tallas"
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   840
         Width           =   975
      End
      Begin VB.Label nlinea 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   42
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linea"
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   975
      End
   End
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
      Height          =   6495
      Left            =   11640
      TabIndex        =   17
      Top             =   6240
      Visible         =   0   'False
      Width           =   11415
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
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   8160
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "tproducc.frx":2424
         Height          =   5535
         Left            =   120
         OleObjectBlob   =   "tproducc.frx":2438
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   840
         Width           =   11055
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
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
      Height          =   375
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   15
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox numero 
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
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox area 
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
      Left            =   5640
      MaxLength       =   60
      TabIndex        =   12
      Top             =   1440
      Width           =   3255
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "tproducc.frx":2E03
      Height          =   3375
      Left            =   120
      OleObjectBlob   =   "tproducc.frx":2E17
      TabIndex        =   11
      Top             =   1920
      Width           =   11055
   End
   Begin VB.TextBox bodega 
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
      Left            =   5640
      MaxLength       =   6
      TabIndex        =   9
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox bodegai 
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
      Left            =   5640
      MaxLength       =   6
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox fecha 
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
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
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
      Left            =   720
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tproducc.frx":4D62
      Style           =   1  'Graphical
      TabIndex        =   3
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
      Picture         =   "tproducc.frx":5F74
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Nuevo registro"
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
      Picture         =   "tproducc.frx":7186
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   735
   End
   Begin VB.Label bandera 
      BackColor       =   &H00FFFF00&
      Height          =   375
      Left            =   4560
      TabIndex        =   64
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
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
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ingresar Insumos+Mano Obra"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5280
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Observacion"
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
      Left            =   3360
      TabIndex        =   13
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Almacen Destino"
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
      Left            =   3360
      TabIndex        =   10
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Almacen Insumo"
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
      Left            =   3360
      TabIndex        =   8
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
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
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero"
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
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.Menu ajdu1 
      Caption         =   "&Add"
   End
   Begin VB.Menu grba1 
      Caption         =   "&Grabar"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tproducc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ajdu1_Click()
If Frame1.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
inicializa
numero = ""
sql_detalle
numero.Enabled = True
numero.SetFocus
End Sub

Private Sub area_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
DBGrid2.SetFocus
End Sub

Private Sub area_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   bodega.SetFocus
   Exit Sub
End If

End Sub


Private Sub bo712_Click()

End Sub

Private Sub bodega_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
found = busca_bodega("" & bodega)
If found = 0 Then
   bodega = ""
   bodega.SetFocus
   Exit Sub
End If
area.SetFocus

End Sub

Private Sub bodega_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_bodega
End If

If KeyCode = &H26 Then
   bodegai.SetFocus
   Exit Sub
End If

End Sub

Private Sub bodegai_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
found = busca_bodega("" & bodegai)
If found = 0 Then
   bodegai = ""
   bodegai.SetFocus
   Exit Sub
End If
bodega.SetFocus


End Sub

Private Sub bodegai_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_bodegai
End If

If KeyCode = &H26 Then
   fechaf.SetFocus
   Exit Sub
End If

End Sub

Private Sub buffer_DblClick()
Command1_Click
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   dlo132_Click
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

Private Sub cmdPrint_Click()
djuer1_Click
End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdSave_Click()
grba1_Click
End Sub


Private Sub cmdSort_Click()

End Sub

Private Sub Command1_Click()
Dim buf As String
If opcion1 = "7" Then
   If Len(buffer) = 0 Then
   buf = "select Tipo,Serie,Codigo,Numero,Fecha,Moneda,Total,Nombre,Vendedor from cpedidov "
   Else
   buf = "select Tipo,Serie,Codigo,Numero,Fecha,Moneda,Total,Nombre,Vendedor from cpedidov where " & Combo1 & " like '" & buffer & "%'"
   End If
End If

If opcion1 = "1" Then
   If Len(buffer) = 0 Then
   buf = "select Codigo,Numero,Fecha,Bodegai,Bodega,Observa from cproducc "
   Else
   buf = "select Codigo,Numero,Fecha,Bodegai,Bodega,Observa from cproducc where " & Combo1 & " like '" & buffer & "%'"
   End If
End If
If opcion1 = "2" Then
   If Len(buffer) = 0 Then
   buf = "select Nombre,Codigo from bodega "
   Else
   buf = "select Nombre,Codigo from bodega where " & Combo1 & " like '" & buffer & "%'"
   End If
End If
If opcion1 = "3" Then
   If Len(buffer) = 0 Then
   buf = "select Nombre,Codigo from bodega "
   Else
   buf = "select Nombre,Codigo from bodega where " & Combo1 & " like '" & buffer & "%'"
   End If
End If

If opcion1 = "4" Then
   If Len(buffer) = 0 Then
   buf = "select Nombre,Codigo from vendedor "
   Else
   buf = "select Nombre,Codigo from Vendedor where " & Combo1 & " like '" & buffer & "%'"
   End If
End If
If opcion1 = "6" Then
   If Len(buffer) = 0 Then
   buf = "select Nombre,Codigo from CLIENTES "
   Else
   buf = "select Nombre,Codigo from CLIENTES where " & Combo1 & " like '" & buffer & "%'"
   End If
End If
If opcion1 = "5" Then
   If Len(buffer) = 0 Then
   buf = "select Descripcio,Producto,Unidad,Factor,Costou,Linea from producto "
   Else
   buf = "select Descripcio,Producto,Unidad,Factor,Costou,Linea from producto where " & Combo1 & " like '" & buffer & "%'"
   End If
End If
               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldir
               Data1.RecordSource = buf
               Data1.refresh
               If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
                  Data1.Recordset.Close
                  buffer.SetFocus
                  Exit Sub
               End If
               If opcion1 = "2" Or opcion1 = "3" Or opcion1 = "4" Then
               DBGrid1.columns(0).Width = 4000
               DBGrid1.columns(1).Width = 2000
               End If
               If opcion1 = "1" Then
               DBGrid1.columns(0).Width = 1500
               DBGrid1.columns(1).Width = 1500
               DBGrid1.columns(2).Width = 1500
               DBGrid1.columns(3).Width = 1500
               DBGrid1.columns(4).Width = 1500
               DBGrid1.columns(5).Width = 4500
               End If
               If opcion1 = "7" Then
               DBGrid1.columns(0).Width = 700
               DBGrid1.columns(1).Width = 700
               DBGrid1.columns(2).Width = 1500
               DBGrid1.columns(3).Width = 1500
               DBGrid1.columns(4).Width = 1500
               DBGrid1.columns(5).Width = 1500
               DBGrid1.columns(6).Width = 1500
               End If
               If opcion1 = "5" Then
               DBGrid1.columns(0).Width = 4500
               DBGrid1.columns(1).Width = 1500
               DBGrid1.columns(2).Width = 1000
               DBGrid1.columns(3).Width = 1000
               DBGrid1.columns(4).Width = 1000
               End If
               DBGrid1.SetFocus
               

End Sub



Private Sub Command2_Click()
Dim sdx As Double
Data2.Recordset.Edit
Data2.Recordset.Fields("t1") = Val("" & t1)
Data2.Recordset.Fields("t2") = Val("" & t2)
Data2.Recordset.Fields("t3") = Val("" & t3)
Data2.Recordset.Fields("t4") = Val("" & t4)
Data2.Recordset.Fields("t5") = Val("" & t5)
Data2.Recordset.Fields("t6") = Val("" & t6)
Data2.Recordset.Fields("t7") = Val("" & t7)
Data2.Recordset.Fields("t8") = Val("" & t8)
Data2.Recordset.Fields("t9") = Val("" & t9)
Data2.Recordset.Fields("t10") = Val("" & t10)
Data2.Recordset.Fields("t11") = Val("" & t11)
Data2.Recordset.Fields("t12") = Val("" & t12)
Data2.Recordset.Fields("t13") = Val("" & t13)
Data2.Recordset.Fields("t14") = Val("" & t14)
Data2.Recordset.Fields("t15") = Val("" & t15)
Data2.Recordset.Fields("t16") = Val("" & t16)
sdx = Val("" & t1) + Val("" & t2) + Val("" & t3) + Val("" & t4) + Val("" & t5) + Val("" & t6) + Val("" & t7) + Val("" & t8) + Val("" & t9) + Val("" & t10) + Val("" & t11) + Val("" & t12) + Val("" & t13) + Val("" & t14) + Val("" & t15) + Val("" & t16)
Data2.Recordset.Fields("cantidad") = sdx
Data2.Recordset.Update
dlo132_Click
End Sub

Private Sub Command3_Click()
dlo132_Click
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim found As Integer
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   If opcion1 = "1" Then
   numero = DBGrid1.columns(1)
   Frame1.Visible = False
   numero.SetFocus
   numero_KeyPress 13
   End If
   If opcion1 = "2" Then
   bodegai = DBGrid1.columns(1)
   Frame1.Visible = False
   bodegai.SetFocus
   bodegai_KeyPress 13
   End If
If opcion1 = "3" Then
   bodega = DBGrid1.columns(1)
   Frame1.Visible = False
   bodega.SetFocus
   bodega_KeyPress 13
   End If
If opcion1 = "6" Then
   DBGrid2.columns(10) = DBGrid1.columns(1)
   Frame1.Visible = False
   DBGrid2.SetFocus
   End If
If opcion1 = "4" Then
   DBGrid2.columns(11) = DBGrid1.columns(1)
   Frame1.Visible = False
   DBGrid2.SetFocus
   End If
   
   

If opcion1 = "5" Then
   'xtarjeta = poner_tarjeta()
   'found = verifica_doble("" & xtarjeta, "" & Data1.Recordset.Fields("producto"))
   'If found = 1 Then
   '   MsgBox "Tarjeta+Producto ya seleccionado", 48, "Aviso"
   '   DBGrid1.SetFocus
   '   Exit Sub
   'End If
   ir_ultimo
   Data2.Recordset.AddNew
   Data2.Recordset.Fields("tarjeta") = "00"
   Data2.Recordset.Fields("numero") = numero
   Data2.Recordset.Fields("bodega") = bodega
   Data2.Recordset.Fields("fecha") = fecha
   Data2.Recordset.Fields("producto") = "" & Data1.Recordset.Fields("producto")
   Data2.Recordset.Fields("descripcio") = "" & Data1.Recordset.Fields("descripcio")
   Data2.Recordset.Fields("unidad") = "" & Data1.Recordset.Fields("unidad")
   Data2.Recordset.Fields("factor") = Val("" & Data1.Recordset.Fields("factor"))
   Data2.Recordset.Fields("linea") = "" & Data1.Recordset.Fields("linea")
   Data2.Recordset.Fields("nro") = "1"
   Data2.Recordset.Update
   'Data2.Refresh
   Frame1.Visible = False
   DBGrid2.SetFocus
   End If
End If


End Sub
Sub ir_ultimo()
On Error GoTo cmd6_err
Data2.Recordset.MoveLast
Exit Sub
cmd6_err:
Exit Sub
End Sub


Private Sub descripcio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
End Sub



Private Sub descripcio_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   'codigo.SetFocus
   Exit Sub
End If

End Sub

Private Sub DBGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Dim found As Integer
Select Case ColIndex
Case 1, 2, 3, 6, 7, 8, 9, 10, 11, 12, 13, 14
     Cancel = True
     Exit Sub
Case 0
     If Len("" & DBGrid2.columns(1)) = 0 Then  '
        MsgBox "Selecciona producto Primero,", 48, "Aviso"
        Cancel = True
        Exit Sub
     End If

Case 5
     If Len("" & DBGrid2.columns(0)) = 0 Then  '
        Cancel = True
        Exit Sub
     End If
     If Len("" & DBGrid2.columns(8)) > 0 Then  'ojo no se puede poner si es talla
        Cancel = True
        Exit Sub
     End If
Case 9
     If Len("" & DBGrid2.columns(0)) = 0 Then  '
        Cancel = True
        Exit Sub
     End If
End Select

End Sub

Private Sub DBGrid2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim found As Integer
Select Case ColIndex
       Case 5
        If Len(DBGrid2.columns(0)) = 0 Then
           Cancel = True
           Exit Sub
        End If
        If Not IsNumeric(DBGrid2.columns(5)) Then  '
           Cancel = True
           Exit Sub
        End If
        Case 9
        If Len(DBGrid2.columns(0)) = 0 Then
           Cancel = True
           Exit Sub
        End If
        
        
End Select
End Sub

Private Sub dbgrid2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 27 Then Exit Sub
area.SetFocus
End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo cmd46_err
If KeyCode = &H2E Then  'borrar linea
   Data2.Recordset.Delete
   Data2.refresh
   Exit Sub
End If
If KeyCode = &H76 Then  'f7
   'tprodup.Show 1
End If
If KeyCode = &H70 And DBGrid2.Col = 10 Then 'f1
   consulta_codigo
   Exit Sub
End If
If KeyCode = &H70 And DBGrid2.Col = 11 Then 'f1
   consulta_codigo1
   Exit Sub
End If


If KeyCode = &H70 Then  'f1
   'If Len(DBGrid2.Columns(0)) > 0 Then Exit Sub
   consulta_producto
End If

If KeyCode = &H71 Then  'f2
   If Len(DBGrid2.columns(1)) > 0 And Len(DBGrid2.columns(8)) > 0 Then
      ingreso_tallas "" & DBGrid2.columns(8)
   End If
End If
Exit Sub
cmd46_err:
Exit Sub
End Sub

Private Sub djuer1_Click()

End Sub

Private Sub dlo132_Click()
If Frame2.Visible = True Then
   Frame2.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If

If opcion1 = "1" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      numero.SetFocus
      Exit Sub
   End If
End If
If opcion1 = "6" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      DBGrid2.SetFocus
      Exit Sub
   End If
End If
If opcion1 = "4" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      DBGrid2.SetFocus
      Exit Sub
   End If
End If


If opcion1 = "2" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      bodegai.SetFocus
      Exit Sub
   End If
End If

If opcion1 = "3" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      bodega.SetFocus
      Exit Sub
   End If
End If
If opcion1 = "5" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      DBGrid2.SetFocus
      Exit Sub
   End If
End If
tproducc.Hide
Unload tproducc
End Sub



Private Sub fecha_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(fecha) = 0 Then
   fecha = Format(Now, "dd/mm/yyyy")
End If
If Not IsDate(fecha) Then Exit Sub
fechaf.SetFocus
End Sub

Private Sub fecha_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   If numero.Enabled = True Then
   numero.SetFocus
   End If
   Exit Sub
End If


End Sub

Private Sub fechaf_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(fechaf) = 0 Then
   fechaf = Format(Now, "dd/mm/yyyy")
End If
If Not IsDate(fechaf) Then Exit Sub
bodegai.SetFocus

End Sub

Private Sub fechaf_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
  fecha.SetFocus
  Exit Sub
End If

End Sub

Sub inicializa()
fechaf = ""
area = ""
fecha = ""
bodegai = ""
bodega = ""
borrar_todo
End Sub
Function verifica_existe()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("cproducc")
mytablex.Index = "cproducc"
mytablex.Seek "=", numero
If Not mytablex.NoMatch Then
   verifica_existe = 1
End If
mytablex.Close
 

End Function
Function busca_registro()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("cproducc")
mytablex.Index = "cproducc"
mytablex.Seek "=", numero
If Not mytablex.NoMatch Then
   pone_registro mytablex
   busca_registro = 1
End If
carga_detalle
'------------------------------------- ------------
mytablex.Close
 
End Function
Sub carga_detalle()

Dim mytablex As Table
Dim i As Integer
borrar_todo

Set mytablex = mydbxglo.OpenTable("dproducc")
mytablex.Index = "dproducc"
mytablex.Seek "=", numero
If Not mytablex.NoMatch Then
   Do
   If mytablex.EOF Then Exit Do
   If "" & mytablex.Fields("numero") = numero Then
      Data2.Recordset.AddNew
      For i = 0 To mytablex.Fields.count - 1
      Data2.Recordset.Fields(i) = mytablex.Fields(i)
      Next i
      Data2.Recordset.Update
      Else: Exit Do
   End If
   mytablex.MoveNext
   Loop
End If
mytablex.Close
 
Data2.refresh
End Sub
Sub pone_registro(mytablex As Table)
Dim found As Integer
numero = "" & mytablex.Fields("numero")
fecha = "" & mytablex.Fields("fecha")
fechaf = "" & mytablex.Fields("fechaf")
area = "" & mytablex.Fields("observa")
bodegai = "" & mytablex.Fields("bodegai")
bodega = "" & mytablex.Fields("bodega")
End Sub
Sub grabando(mytablex As Table)
mytablex.Fields("numero") = numero
mytablex.Fields("observa") = area
mytablex.Fields("fecha") = fecha
mytablex.Fields("bodegai") = bodegai
mytablex.Fields("bodega") = bodega
mytablex.Fields("fechaf") = fechaf
mytablex.Fields("estado") = "0"
End Sub

Private Sub Form_Activate()
Dim found As Integer
sql_detalle
If bandera = "Modifica" Then
   found = busca_registro()
   sql_detalle
End If

End Sub

Private Sub grba1_Click()
Dim found As Integer
If Frame1.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
If bandera = "Nuevo" Then
   found = verifica_existe()
   If found = 1 Then
      MsgBox "Ya existe Numero,elija Otro Numero"
      numero = ""
      numero.SetFocus
      Exit Sub
   End If
End If
found = grabar()
If found = 0 Then Exit Sub
If bandera = "Nuevo" Then
numero.SetFocus
End If
End Sub

Private Sub Label1_Click()
cmdSort_Click
End Sub


Function grabar()
Dim found As Integer
Dim mytablex As Table

Dim xx As String

found = valida()
If found = 0 Then
   MsgBox "Campos invalidos", 48, "Aviso"
   Exit Function
End If

Set mytablex = mydbxglo.OpenTable("cproducc")
mytablex.Index = "cproducc"
mytablex.Seek "=", numero
If mytablex.NoMatch Then
   mytablex.AddNew
   grabando mytablex
   mytablex.Update
   xx = busca_config(1)
   grabar = 1
End If
If Not mytablex.NoMatch Then
   If MsgBox("Desea Reescribir?", 1, "Aviso") = 1 Then
   mytablex.Edit
   grabando mytablex
   mytablex.Update
   grabar = 1
   End If
End If
'------------------------------------- ------------
mytablex.Close
 
'ahora grabando detalle de productos-------
grabando_detalle
End Function
Sub grabando_detalle()
Dim i As Integer

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("dproducc")
mytablex.Index = "dproducc"
akm12:
mytablex.Seek "=", numero
If Not mytablex.NoMatch Then
   mytablex.Delete
   GoTo akm12
End If
'ahora adicionado
ir_inicio
Do
If Data2.Recordset.EOF Then Exit Do
mytablex.AddNew
For i = 0 To Data2.Recordset.Fields.count - 1
    mytablex.Fields(i) = Data2.Recordset.Fields(i)
Next i
mytablex.Fields("numero") = numero
'mytablex.Fields("estado") = "0"
mytablex.Update
Data2.Recordset.MoveNext
Loop
mytablex.Close
 
End Sub

Function valida()
Dim found As Integer
If Len(numero) = 0 Then
   numero.SetFocus
   Exit Function
End If
If Not IsDate(fecha) Then
   fecha.SetFocus
   Exit Function
End If
If Not IsDate(fechaf) Then
   fechaf.SetFocus
   Exit Function
End If

If Len(bodegai) = 0 Then
   bodegai.SetFocus
   Exit Function
End If
If Len(bodega) = 0 Then
   bodega.SetFocus
   Exit Function
End If
found = busca_bodega("" & bodegai)
If found = 0 Then
   bodegai.SetFocus
   Exit Function
End If
found = busca_bodega("" & bodega)
If found = 0 Then
   bodega.SetFocus
   Exit Function
End If
valida = 1
End Function
Function busca_bodega(buf As String)

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("bodega")
mytablex.Index = "codigo"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_bodega = 1
End If
'------------------------------------- ------------
mytablex.Close
 

End Function
Function busca_codigo1(buf As String)

Dim mytablex As Table


Set mytablex = mydbxglo.OpenTable("clientes")
mytablex.Index = "codigo"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_codigo1 = 1
   
End If
'------------------------------------- ------------
mytablex.Close
 
End Function




Private Sub Label8_Click()
On Error GoTo cmd45_err
If Len("" & Data2.Recordset.Fields("producto")) > 0 Then
   If Len("" & Data2.Recordset.Fields("nro")) = 0 Then
      MsgBox "Ingresar el numero de Formula", 48, "Aviso"
      Exit Sub
   End If
   'ingresar insumos
   rproducc.numero = "" & numero
   'rproducc.nro = "" & Data2.Recordset.Fields("nro")
   'rproducc.producto = "" & Data2.Recordset.Fields("producto")
   'rproducc.descripcio = "" & Data2.Recordset.Fields("descripcio")
   
   
   rproducc.nro = "" & Data2.Recordset.Fields("nro")
   rproducc.tarjeta = "" & Data2.Recordset.Fields("tarjeta")
   rproducc.cantidad = "" & Data2.Recordset.Fields("cantidad")
   rproducc.xlinea = "" & Data2.Recordset.Fields("linea")
   rproducc.producto = "" & Data2.Recordset.Fields("producto")
   rproducc.descripcio = "" & Data2.Recordset.Fields("descripcio")
   rproducc.xt1 = "" & Data2.Recordset.Fields("t1")
   rproducc.xt2 = "" & Data2.Recordset.Fields("t2")
   rproducc.xt3 = "" & Data2.Recordset.Fields("t3")
   rproducc.xt4 = "" & Data2.Recordset.Fields("t4")
   rproducc.xt5 = "" & Data2.Recordset.Fields("t5")
   rproducc.xt6 = "" & Data2.Recordset.Fields("t6")
   rproducc.xt7 = "" & Data2.Recordset.Fields("t7")
   rproducc.xt8 = "" & Data2.Recordset.Fields("t8")
   rproducc.xt9 = "" & Data2.Recordset.Fields("t9")
   rproducc.xt10 = "" & Data2.Recordset.Fields("t10")
   rproducc.xt11 = "" & Data2.Recordset.Fields("t11")
   rproducc.xt12 = "" & Data2.Recordset.Fields("t12")
   rproducc.xt13 = "" & Data2.Recordset.Fields("t13")
   rproducc.xt14 = "" & Data2.Recordset.Fields("t14")
   rproducc.xt15 = "" & Data2.Recordset.Fields("t15")
   rproducc.xt16 = "" & Data2.Recordset.Fields("t16")
   
   rproducc.Show 1
End If
Exit Sub
cmd45_err:
Exit Sub
End Sub

Private Sub numero_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   dlo132_Click
   Exit Sub
End If
If Len(numero) = 0 Then
   numero = busca_config(0)
   Exit Sub
End If
fecha.SetFocus
End Sub
Sub consulta_bodegai()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Codigo"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "2"
Command1_Click


End Sub
Sub consulta_bodega()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Codigo"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "3"
Command1_Click


End Sub

Sub consulta_codigo()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Codigo"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "6"
Command1_Click
End Sub
Sub consulta_codigo1()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Codigo"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "4"
Command1_Click
End Sub
Sub consulta_pedido()
Combo1.Clear
Combo1.AddItem "Numero"
Combo1.AddItem "Codigo"
Combo1.AddItem "Nombre"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "7"
Command1_Click
End Sub

Sub consulta_producto()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Producto"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "5"
Command1_Click

End Sub




Function busca_config(sw As Integer) As String

Dim mytablex As Table
Dim sdx As Double
If sw = 1 Then
   If Not IsNumeric(numero) Then
      busca_config = ""
      Exit Function
   End If
End If

Set mytablex = mydbxglo.OpenTable("parame")
mytablex.Index = "codigo"
mytablex.Seek "=", "01"
If Not mytablex.NoMatch Then
   If sw = 0 Then
      sdx = Val("" & mytablex.Fields("produccion")) + 1
      busca_config = "" & sdx
   End If
   If sw = 1 Then
      mytablex.Edit
      mytablex.Fields("produccion") = numero
      mytablex.Update
   End If
End If
'------------------------------------- ------------
mytablex.Close
 

End Function
Sub sql_detalle()
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = "select * from " & "_p" & gusuario & " order by str(tarjeta)"
               Data2.refresh
End Sub
Function verifica_doble(buf As String, buf1 As String)

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("_p" & gusuario)
mytablex.Index = "dproducc1"
mytablex.Seek "=", buf, buf1
If Not mytablex.NoMatch Then
   verifica_doble = 1
End If
'------------------------------------- ------------
mytablex.Close
 

End Function
Sub ingreso_tallas(buf As String)
Dim found As Integer
linea = buf
found = busca_linea(buf)
If found = 0 Then Exit Sub
pone_tallas
Frame2.Visible = True
t1.SetFocus
End Sub
Function busca_linea(buf As String)

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("linea")
mytablex.Index = "linea"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_linea = 1
   nlinea = "" & mytablex.Fields("descripcio")
   nt1 = "" & mytablex.Fields("t1")
   nt2 = "" & mytablex.Fields("t2")
   nt3 = "" & mytablex.Fields("t3")
   nt4 = "" & mytablex.Fields("t4")
   nt5 = "" & mytablex.Fields("t5")
   nt6 = "" & mytablex.Fields("t6")
   nt7 = "" & mytablex.Fields("t7")
   nt8 = "" & mytablex.Fields("t8")
   nt9 = "" & mytablex.Fields("t9")
   nt10 = "" & mytablex.Fields("t10")
   nt11 = "" & mytablex.Fields("t11")
   nt12 = "" & mytablex.Fields("t12")
   nt13 = "" & mytablex.Fields("t13")
   nt14 = "" & mytablex.Fields("t14")
   nt15 = "" & mytablex.Fields("t15")
   nt16 = "" & mytablex.Fields("t16")
End If
'------------------------------------- ------------
mytablex.Close
 
End Function
Sub pone_tallas()
t1 = "" & Data2.Recordset.Fields("t1")
t2 = "" & Data2.Recordset.Fields("t2")
t3 = "" & Data2.Recordset.Fields("t3")
t4 = "" & Data2.Recordset.Fields("t4")
t5 = "" & Data2.Recordset.Fields("t5")
t6 = "" & Data2.Recordset.Fields("t6")
t7 = "" & Data2.Recordset.Fields("t7")
t8 = "" & Data2.Recordset.Fields("t8")
t9 = "" & Data2.Recordset.Fields("t9")
t10 = "" & Data2.Recordset.Fields("t10")
t11 = "" & Data2.Recordset.Fields("t11")
t12 = "" & Data2.Recordset.Fields("t12")
t13 = "" & Data2.Recordset.Fields("t13")
t14 = "" & Data2.Recordset.Fields("t14")
t15 = "" & Data2.Recordset.Fields("t15")
t16 = "" & Data2.Recordset.Fields("t16")
End Sub



Private Sub pedido_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   KeyAscii = 0
   Exit Sub
End If
fecha.SetFocus
End Sub

Private Sub pedido_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   numero.Enabled = True
   numero.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
consulta_pedido
End If


End Sub

Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t2.SetFocus
End Sub

Private Sub t10_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t11.SetFocus

End Sub

Private Sub t11_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t12.SetFocus

End Sub

Private Sub t12_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t13.SetFocus

End Sub

Private Sub t13_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t14.SetFocus

End Sub

Private Sub t14_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t15.SetFocus

End Sub

Private Sub t15_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t16.SetFocus

End Sub

Private Sub t16_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
End Sub

Private Sub t2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t3.SetFocus

End Sub

Private Sub t3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t4.SetFocus

End Sub

Private Sub t4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t5.SetFocus

End Sub

Private Sub t5_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t6.SetFocus

End Sub

Private Sub t6_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t7.SetFocus

End Sub

Private Sub t7_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t8.SetFocus

End Sub

Private Sub t8_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t9.SetFocus

End Sub

Private Sub t9_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
t10.SetFocus

End Sub

Sub borrar_todo()
On Error GoTo cmd15_err
ir_inicio
sigue1:
Data2.Recordset.Delete
Data2.refresh
GoTo sigue1
Exit Sub
cmd15_err:
Exit Sub
End Sub
Sub ir_inicio()
On Error GoTo cmd16_err
Data2.Recordset.MoveFirst
Exit Sub
cmd16_err:
Exit Sub

End Sub
