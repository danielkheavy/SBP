VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form matedoc 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento Documentos"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
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
      Height          =   6975
      Left            =   0
      TabIndex        =   67
      Top             =   0
      Visible         =   0   'False
      Width           =   13455
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
         TabIndex        =   71
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
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
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
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "matedoc.frx":0000
         Height          =   5295
         Left            =   120
         OleObjectBlob   =   "matedoc.frx":0014
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   1080
         Width           =   12975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Generando Documentos"
      Height          =   3135
      Left            =   2400
      TabIndex        =   56
      Top             =   1920
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox xtipo 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   64
         Top             =   360
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
         Left            =   5880
         MaskColor       =   &H00E0E0E0&
         Picture         =   "matedoc.frx":09DF
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Borrar registro"
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command1 
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
         Left            =   5880
         MaskColor       =   &H00E0E0E0&
         Picture         =   "matedoc.frx":1BF1
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Grabar registro"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox xserie 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   58
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox xnumero 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1800
         MaxLength       =   11
         TabIndex        =   57
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label xdescripcio 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1800
         TabIndex        =   66
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
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
         Left            =   840
         TabIndex        =   65
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000009&
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
         Left            =   840
         TabIndex        =   61
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
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
         Left            =   840
         TabIndex        =   60
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000009&
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
         Left            =   840
         TabIndex        =   59
         Top             =   1560
         Width           =   975
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Detalle"
      Height          =   7095
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   13575
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   12720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "matedoc.frx":2E03
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "Consulta"
         Top             =   1200
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
         Left            =   12720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "matedoc.frx":4015
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         ToolTipText     =   "Borrar registro"
         Top             =   480
         Width           =   735
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ingreso de Pedidos x Local"
         Height          =   3975
         Left            =   360
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   9495
         Begin VB.TextBox precio 
            Height          =   375
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   3240
            Width           =   1935
         End
         Begin VB.TextBox proveedorp 
            Height          =   375
            Left            =   1920
            MaxLength       =   11
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   2520
            Width           =   1935
         End
         Begin VB.TextBox l1 
            Height          =   375
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox l2 
            Height          =   375
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox l3 
            Height          =   375
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox l4 
            Height          =   375
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1575
         End
         Begin VB.CommandButton Command4 
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
            Left            =   8640
            MaskColor       =   &H00E0E0E0&
            Picture         =   "matedoc.frx":5227
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Grabar registro"
            Top             =   1080
            Width           =   735
         End
         Begin VB.CommandButton Command5 
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
            Left            =   8640
            MaskColor       =   &H00E0E0E0&
            Picture         =   "matedoc.frx":6439
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Borrar registro"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox observa4 
            Height          =   375
            Left            =   3480
            MaxLength       =   20
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox observa3 
            Height          =   375
            Left            =   3480
            MaxLength       =   20
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   1560
            Width           =   3495
         End
         Begin VB.TextBox observa2 
            Height          =   375
            Left            =   3480
            MaxLength       =   20
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   1200
            Width           =   3495
         End
         Begin VB.TextBox observa1 
            Height          =   375
            Left            =   3480
            MaxLength       =   20
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   840
            Width           =   3495
         End
         Begin VB.TextBox lx4 
            Height          =   375
            Left            =   6960
            MaxLength       =   10
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox lx3 
            Height          =   375
            Left            =   6960
            MaxLength       =   10
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox lx2 
            Height          =   375
            Left            =   6960
            MaxLength       =   10
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox lx1 
            Height          =   375
            Left            =   6960
            MaxLength       =   10
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Precio"
            Height          =   375
            Left            =   720
            TabIndex        =   55
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label nproveedor 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1920
            TabIndex        =   51
            Top             =   2880
            Width           =   5175
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre"
            Height          =   375
            Left            =   720
            TabIndex        =   50
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label Label10 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Proveedor"
            Height          =   375
            Left            =   720
            TabIndex        =   49
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label Label22 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Local"
            Height          =   375
            Left            =   720
            TabIndex        =   47
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label tl1 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   720
            TabIndex        =   46
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label tl2 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   720
            TabIndex        =   45
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label tl3 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   720
            TabIndex        =   44
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label tl4 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   720
            TabIndex        =   43
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label Label34 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1920
            TabIndex        =   42
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label35 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1920
            TabIndex        =   41
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label36 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1920
            TabIndex        =   40
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label20 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cantidad"
            Height          =   375
            Left            =   1920
            TabIndex        =   39
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label21 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Observaciones"
            Height          =   375
            Left            =   3480
            TabIndex        =   38
            Top             =   480
            Width           =   3495
         End
         Begin VB.Label Label23 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Can.Recibido"
            Height          =   375
            Left            =   6960
            TabIndex        =   37
            Top             =   480
            Width           =   1575
         End
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "matedoc.frx":764B
         Height          =   6375
         Left            =   120
         OleObjectBlob   =   "matedoc.frx":765F
         TabIndex        =   21
         Top             =   480
         Width           =   12495
      End
   End
   Begin VB.ComboBox orden 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   840
      Width           =   1575
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "matedoc.frx":9DD6
      Height          =   5775
      Left            =   120
      OleObjectBlob   =   "matedoc.frx":9DEA
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1440
      Width           =   13455
   End
   Begin VB.ComboBox tipoclie 
      BackColor       =   &H00C0FFFF&
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
      Left            =   9960
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox codigo 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   9960
      MaxLength       =   11
      TabIndex        =   7
      Text            =   "*"
      Top             =   480
      Width           =   1575
   End
   Begin VB.ComboBox estado 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox fechaf 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox fechai 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   6840
      MaxLength       =   10
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox numero 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3720
      MaxLength       =   11
      TabIndex        =   3
      Text            =   "*"
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox serie 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "*"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox tipo 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "*"
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Orden"
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
      TabIndex        =   19
      Top             =   840
      Width           =   975
   End
   Begin VB.Label acu 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   11640
      TabIndex        =   17
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000009&
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
      Left            =   9000
      TabIndex        =   16
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000009&
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
      Left            =   9000
      TabIndex        =   15
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado"
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
      TabIndex        =   14
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaFinal"
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
      Left            =   5640
      TabIndex        =   13
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaInicio"
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
      Left            =   5640
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
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
      Left            =   2760
      TabIndex        =   11
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
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
      Left            =   2760
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
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
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu djueju782 
      Caption         =   "&Ejecutar"
   End
   Begin VB.Menu menup12 
      Caption         =   "&Generar"
   End
   Begin VB.Menu ldso232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "matedoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xacu As String
Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   ldso232_Click
   Exit Sub
End If
Command3_Click
End Sub
'------------------------------------- ------------


Private Sub cmdDelete_Click()
If Frame3.Visible = True Then Exit Sub
Frame1.Visible = False
DBGrid2.SetFocus
End Sub

Private Sub cmdSort_Click()
If Frame3.Visible = True Then Exit Sub
xxbusca_locales
Frame3.Visible = True
precio = "" & Data3.Recordset.Fields("precio")
l1 = "" & Data3.Recordset.Fields("l1")
l2 = "" & Data3.Recordset.Fields("l2")
l3 = "" & Data3.Recordset.Fields("l3")
l4 = "" & Data3.Recordset.Fields("l4")
observa1 = "" & Data3.Recordset.Fields("observa1")
observa2 = "" & Data3.Recordset.Fields("observa2")
observa3 = "" & Data3.Recordset.Fields("observa3")
observa4 = "" & Data3.Recordset.Fields("observa4")
lx1 = "" & Data3.Recordset.Fields("lx1")
lx2 = "" & Data3.Recordset.Fields("lx2")
lx3 = "" & Data3.Recordset.Fields("lx3")
lx4 = "" & Data3.Recordset.Fields("lx4")
proveedorp = "" & Data3.Recordset.Fields("proveedorp")
l1.SetFocus

End Sub

Private Sub Command1_Click()
Dim found As Integer
If Len(xtipo) = 0 Then
   xtipo.SetFocus
   Exit Sub
End If
found = busca_tipo1("" & xtipo)
If found = 0 Then
   MsgBox "Tipo Documento no existe", 48, "Aviso"
   xtipo.SetFocus
   Exit Sub
End If
If Len(xserie) = 0 Then
   xserie.SetFocus
   Exit Sub
End If
If Len(xnumero) = 0 Then
   xnumero.SetFocus
   Exit Sub
End If
found = busca_numero("" & xnumero)
If found = 1 Then
   MsgBox "Ya existe numero ", 48, "Aviso"
   Exit Sub
End If
found = proceso_ordenes()
If found = 0 Then
   MsgBox "Proceso Realizado ", 48, "Aviso"
   Exit Sub
End If
found = busca_tipo("" & xtipo, 1)
End Sub

Private Sub Command2_Click()
ldso232_Click
End Sub

Private Sub Command3_Click()
If opcion1 = "1" Then
      If Len(buffer) = 0 Then
      buf = "select Descripcio,Tipo from Tipo "
      Else
      buf = "select Descripcio,Tipo from Tipo where " & Combo1 & " like '" & buffer & "*'"
      End If
   End If
   If Combo2.ListIndex = 1 Then
      buf = buf & " order by " & Combo1
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
               If opcion1 = "1" Then
                  DBGrid1.Columns(0).Width = 4000
                  DBGrid1.Columns(1).Width = 2000
               End If
               DBGrid1.SetFocus

End Sub

Private Sub Command4_Click()
Dim sdx As Double
Data3.Recordset.Edit
Data3.Recordset.Fields("precio") = Val(precio)
Data3.Recordset.Fields("l1") = Val(l1)
Data3.Recordset.Fields("l2") = Val(l2)
Data3.Recordset.Fields("l3") = Val(l3)
Data3.Recordset.Fields("l4") = Val(l4)
sdx = Val(l1) + Val(l2) + Val(l3) + Val(l4)
Data3.Recordset.Fields("cantidad") = sdx

Data3.Recordset.Fields("lx1") = Val(lx1)
Data3.Recordset.Fields("lx2") = Val(lx2)
Data3.Recordset.Fields("lx3") = Val(lx3)
Data3.Recordset.Fields("lx4") = Val(lx4)

Data3.Recordset.Fields("observa1") = observa1
Data3.Recordset.Fields("observa2") = observa2
Data3.Recordset.Fields("observa3") = observa3
Data3.Recordset.Fields("observa4") = observa4
Data3.Recordset.Fields("total") = Val(Format(Val("" & Data3.Recordset.Fields("cantidad")) * Val("" & Data3.Recordset.Fields("precio")), "0.00"))
Data3.Recordset.Update
Command5_Click

End Sub

Private Sub Command5_Click()
Frame3.Visible = False
DBGrid3.SetFocus
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
If opcion1 = "1" Then
   xtipo = DBGrid1.Columns(1)
   Frame4.Visible = False
   xtipo.SetFocus
End If
End If

End Sub

Private Sub DBGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = &H71 Then  'f2
   marca_nop
End If

If KeyCode = &H72 Then  'f3
   visualiza_grid3
End If

End Sub
Sub marca_nop()
On Error GoTo cmd33_err
If "" & Data2.Recordset.Fields("nop") = "S" Then
   Data2.Recordset.Edit
   Data2.Recordset.Fields("nop") = "N"
   Data2.Recordset.Update
   Exit Sub
End If
If "" & Data2.Recordset.Fields("nop") = "N" Or Len("" & Data2.Recordset.Fields("nop")) = 0 Then
   Data2.Recordset.Edit
   Data2.Recordset.Fields("nop") = "S"
   Data2.Recordset.Update
   Exit Sub
End If
Exit Sub
cmd33_err:
MsgBox Error$
Exit Sub
End Sub
Private Sub DBGrid3_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Frame1.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If
End Sub

Private Sub djueju782_Click()
If Frame3.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
sql_cabecera
DBGrid2.SetFocus

End Sub

Sub sql_cabecera()
Dim buf As String
buf = "select * from  " & cgusuario
buf = buf & "  where fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
buf = buf & " and acu='" & acu & "'"
If tipo <> "*" Then
buf = buf & " and tipo like '" & tipo & "'"
End If
If serie <> "*" Then
buf = buf & " and serie like '" & serie & "'"
End If
If numero <> "*" Then
buf = buf & " and numero like '" & numero & "'"
End If
If tipoclie <> "*" Then
buf = buf & " and tipoclie like '" & tipoclie & "'"
End If
If codigo <> "*" Then
buf = buf & " and codigo like '" & codigo & "'"
End If
If orden <> "*" Then
buf = buf & " order by " & orden
End If
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = buf
               Data2.Refresh

End Sub
Sub sql_detalle()
Dim buf As String
buf = "select * from  " & dgusuariog
buf = buf & " where  tipo='" & DBGrid2.Columns(2) & "'"
buf = buf & " and serie='" & DBGrid2.Columns(3) & "'"
buf = buf & " and numero='" & DBGrid2.Columns(4) & "'"

               Data3.Connect = "foxpro 2.5;"
               Data3.DatabaseName = globaldir
               Data3.RecordSource = buf
               Data3.Refresh

End Sub

Private Sub Form_Load()
orden.Clear
orden.AddItem "*"
orden.AddItem "Codigo"
orden.AddItem "Fecha"
orden.AddItem "Serie"
orden.AddItem "Numero"
orden.AddItem "tipoclie"
orden.ListIndex = 0


estado.Clear
estado.AddItem "*"
estado.AddItem "0"
estado.AddItem "1"
estado.ListIndex = 0

tipoclie.Clear
tipoclie.AddItem "*"
tipoclie.AddItem "P"
tipoclie.AddItem "C"
tipoclie.AddItem "I"
tipoclie.ListIndex = 0

fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
fechaf = Format(Now, "dd/mm/yyyy")


End Sub

Private Sub geny545_Click()

End Sub

Private Sub julo121_Click()

End Sub

Private Sub ldso232_Click()
If Frame4.Visible = True Then
   If opcion1 = "1" Then
    Frame4.Visible = False
    xtipo.SetFocus
    Exit Sub
   End If
End If
If Frame2.Visible = True Then
   Frame2.Visible = False
   Exit Sub
End If

If Frame3.Visible = True Then
   Frame3.Visible = False
   Exit Sub
End If

If Frame1.Visible = True Then
   Frame1.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If
matedoc.Hide
Unload matedoc
End Sub
Sub visualiza_grid3()
Frame1.Visible = True
sql_detalle
DBGrid3.SetFocus
End Sub
Sub xxbusca_locales()
Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("empresa")
mytablex.Index = "codigo"
mytablex.Seek "=", menup.gempresa
If Not mytablex.NoMatch Then
   tl1 = "" & mytablex.Fields("l1")
   tl2 = "" & mytablex.Fields("l2")
   tl3 = "" & mytablex.Fields("l3")
   tl4 = "" & mytablex.Fields("l4")
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close
End Sub
Function proceso_ordenes()
Dim mydbx As Database
Dim mytablex As Table
Dim mytabley As Table
Dim mytablez As Table
Dim sw As Integer
Dim vr
sw = 0
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytabley = mydbx.OpenTable(xarchivo)
mytabley.Index = "tfactura"

Set mytablez = mydbx.OpenTable(xarchivo1)
mytablez.Index = "cuerpo1"

Set mytablex = mydbx.OpenTable(dgusuariog)
mytablex.Index = "tdetalle"
ir_inicio
Do
If Data2.Recordset.EOF Then Exit Do
   sw = 0
   mytablex.Seek "=", "" & Data2.Recordset.Fields("tipo"), "" & Data2.Recordset.Fields("serie"), "" & Data2.Recordset.Fields("numero")
   If Not mytablex.NoMatch Then
      Do
      If mytablex.EOF Then GoTo salix
      If "" & mytablex.Fields("tipo") = "" & Data2.Recordset.Fields("tipo") And "" & mytablex.Fields("serie") = "" & Data2.Recordset.Fields("serie") And "" & mytablex.Fields("numero") = "" & Data2.Recordset.Fields("numero") Then
         If "" & Data2.Recordset.Fields("nop") <> "S" Then
               hacer_cabeza mytablex, mytabley
               hacer_detalle mytablex, mytabley, mytablez
               sw = 1
         End If
         Else: GoTo salix
      End If
      mytablex.MoveNext
      Loop
salix:
   End If
   If sw = 1 Then
   Data2.Recordset.Edit
   Data2.Recordset.Fields("nop") = "S"
   Data2.Recordset.Update
   End If
Data2.Recordset.MoveNext
Loop
proceso_ordenes = sw
mytablex.Close
mytabley.Close
mytablez.Close
mydbx.Close
End Function
Sub hacer_cabeza(mytablex As Table, mytabley As Table)
mytabley.Seek "=", xtipo, xserie, xnumero
If mytabley.NoMatch Then
   mytabley.AddNew
   For i = 0 To Data2.Recordset.Fields.Count - 1
          mytabley.Fields(i) = "" & Data2.Recordset.Fields(i)
   Next i
   mytabley.Fields("tipo") = xtipo
   mytabley.Fields("serie") = xserie
   mytabley.Fields("numero") = xnumero
   mytabley.Fields("acu") = xacu
   mytabley.Fields("acu1") = acu
   mytabley.Fields("nop") = "N"
   mytabley.Fields("fecha") = Format(Now, "dd/mm/yyyy")
   mytabley.Fields("fechae") = Format(Now, "dd/mm/yyyy")
   'If opcion1 = 9 Then   'si es generacion orden de compra
   '   mytabley.Fields("tipoclie") = "P"
   '   mytabley.Fields("codigo") = "" & mytablex.Fields("proveedorp")
   'End If
   mytabley.Update
End If
End Sub
Sub hacer_cabeza1()
End Sub
Sub hacer_detalle(mytablex As Table, mytabley As Table, mytablez As Table)
Dim i As Integer
Dim sdx As Double
mytablez.Seek "=", xtipo, xserie, xnumero, "" & mytablex.Fields("producto"), "" & mytablex.Fields("proveedorp")
If mytablez.NoMatch Then
   mytablez.AddNew
   For i = 0 To mytablex.Fields.Count - 1
       mytablez.Fields(i) = "" & mytablex.Fields(i)
   Next i
   mytablez.Fields("tipo") = xtipo
   mytablez.Fields("serie") = xserie
   mytablez.Fields("numero") = xnumero
   mytablez.Fields("acu") = xacu
   mytablez.Fields("acu1") = acu
   mytablez.Fields("fecha") = Format(Now, "dd/mm/yyyy")
   'mytablez.Fields("fechae") = Format(Now, "dd/mm/yyyy")
   mytablez.Update
End If
If Not mytablez.NoMatch Then
   mytablez.Edit
   mytablez.Fields("l1") = Val("" & mytablez.Fields("l1")) + Val("" & mytablex.Fields("l1"))
   mytablez.Fields("l2") = Val("" & mytablez.Fields("l2")) + Val("" & mytablex.Fields("l2"))
   mytablez.Fields("l3") = Val("" & mytablez.Fields("l3")) + Val("" & mytablex.Fields("l3"))
   mytablez.Fields("l4") = Val("" & mytablez.Fields("l4")) + Val("" & mytablex.Fields("l4"))
   mytablez.Fields("cantidad") = Val("" & mytablez.Fields("l1")) + Val("" & mytablez.Fields("l2")) + Val("" & mytablez.Fields("l3")) + Val("" & mytablez.Fields("l4"))
   mytablez.Update
End If
End Sub
Sub consulta_tipo()
Combo1.Clear
Combo1.AddItem "Tipo"
Combo1.AddItem "Descripcio"
Combo1.ListIndex = 0
Frame4.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "1"
Command3_Click

End Sub

Private Sub menup12_Click()

If Frame3.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
Frame2.Visible = True
xtipo = ""
xserie = ""
xnumero = ""
xtipo.SetFocus

End Sub

Private Sub xnumero_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Command1.SetFocus

End Sub

Private Sub xserie_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
xnumero.SetFocus

End Sub

Private Sub xtipo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
found = busca_tipo("" & xtipo, 0)
If found = 0 Then
   MsgBox "No existe Tipo Documento ", 48, "Aviso"
   Exit Sub
End If
xserie.SetFocus
End Sub

Private Sub xtipo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_tipo
End If

End Sub
Sub ir_inicio()
On Error GoTo cmd5_err
Data2.Recordset.MoveFirst
Exit Sub
cmd5_err:
Exit Sub
End Sub
Function busca_numero(buf As String)
Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("cordenc")
mytablex.Index = "tfactura"
mytablex.Seek "=", xtipo, xserie, buf
If Not mytablex.NoMatch Then
   busca_numero = 1
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close
End Function
Function busca_tipo(buf As String, sw As Integer)
Dim mydbx As Database
Dim mytablex As Table
xdescripcio = ""
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("tipo")
mytablex.Index = "tipo"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_tipo = 1
   xdescripcio = "" & mytablex.Fields("descripcio")
   If sw = 1 Then
      If IsNumeric(xnumero) Then
         mytablex.Edit
         mytablex.Fields("numero") = xnumero
         mytablex.Update
      End If
   End If
   If sw = 0 Then
      If Len(xserie) = 0 Then
         xserie = "" & mytablex.Fields("serie")
      End If
      If Len(xnumero) = 0 Then
         sdx = Val("" & mytablex.Fields("numero")) + 1
         xnumero = "" & sdx
      End If
   End If
   '--------------------------------
End If
mytablex.Close
mydbx.Close
End Function
Function busca_tipo1(buf As String)
Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("tipo")
mytablex.Index = "tipo"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
      Select Case "" & mytablex.Fields("tipodoc")
          Case "A", "B", "C", "D", "G", "E", "F"  'VENTAS
               xarchivo = "FACTURA"
               xarchivo1 = "DETALLE"
               xacu = "V"
          Case "J", "K", "L", "M", "P", "N", "O"  'COMPRAS
               xarchivo = "FACTURA"
               xarchivo1 = "DETALLE"
               xacu = "C"
          Case "H"  'COTIZACION VENTAS
               xarchivo = "CCOTIZAV"
               xarchivo1 = "DCOTIZAV"
               xacu = "" & mytablex.Fields("tipodoc")
          Case "I"  'PEDIDO VENTAS
               xarchivo = "CPEDIDOV"
               xarchivo1 = "DPEDIDOV"
               xacu = "" & mytablex.Fields("tipodoc")
          Case "Q"  'REQUISICION COMPRAS
               xarchivo = "CREQUISA"
               xarchivo1 = "DREQUISA"
               xacu = "" & mytablex.Fields("tipodoc")
          Case "R"  'ORDEN COMPRA
               xarchivo = "CORDENC"
               xarchivo1 = "DORDENC"
               xacu = "" & mytablex.Fields("tipodoc")
          
          Case "T", "S" 'GUIA REMISION
               xarchivo = "FACTURA"
               xarchivo1 = "DETALLE"
               xacu = "" & mytablex.Fields("tipodoc")
      End Select
      
      busca_tipo1 = 1

End If
mytablex.Close
mydbx.Close
End Function
