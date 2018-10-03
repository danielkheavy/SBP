VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form changepr 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio Precios Rapidos"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   13155
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
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
      Height          =   7695
      Left            =   0
      TabIndex        =   103
      Top             =   0
      Visible         =   0   'False
      Width           =   12975
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
         Left            =   6120
         TabIndex        =   106
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
         TabIndex        =   105
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
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin MSDBGrid.DBGrid DBGrid11 
         Bindings        =   "changepr.frx":0000
         Height          =   5295
         Left            =   120
         OleObjectBlob   =   "changepr.frx":0014
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   1080
         Width           =   12735
      End
   End
   Begin VB.TextBox fechavence 
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
      Left            =   3240
      MaxLength       =   10
      TabIndex        =   101
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   12720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox codigo 
      Enabled         =   0   'False
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
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   98
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox descripcio 
      Enabled         =   0   'False
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
      Left            =   3000
      MaxLength       =   60
      TabIndex        =   96
      Top             =   0
      Width           =   5535
   End
   Begin VB.TextBox monedav 
      Enabled         =   0   'False
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
      Left            =   3000
      MaxLength       =   6
      TabIndex        =   95
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox monedac 
      Enabled         =   0   'False
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
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   94
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9960
      MaskColor       =   &H00FFFFFF&
      Picture         =   "changepr.frx":09E0
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   2295
   End
   Begin VB.TextBox unidad 
      Enabled         =   0   'False
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
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   69
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox factor 
      Enabled         =   0   'False
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
      Left            =   1080
      MaxLength       =   6
      TabIndex        =   68
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox costou 
      Enabled         =   0   'False
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
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   67
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox costop 
      Enabled         =   0   'False
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
      Left            =   1080
      MaxLength       =   10
      TabIndex        =   66
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox margen15 
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
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   65
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox pventa15 
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
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   64
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox maximo15 
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
      Left            =   6120
      MaxLength       =   5
      TabIndex        =   63
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox minimo15 
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
      Left            =   5640
      MaxLength       =   5
      TabIndex        =   62
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox margen14 
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
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   61
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox pventa14 
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
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   60
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox maximo14 
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
      Left            =   6120
      MaxLength       =   5
      TabIndex        =   59
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox minimo14 
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
      Left            =   5640
      MaxLength       =   5
      TabIndex        =   58
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox margen13 
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
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   57
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox pventa13 
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
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   56
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox maximo13 
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
      Left            =   6120
      MaxLength       =   5
      TabIndex        =   55
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox minimo13 
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
      Left            =   5640
      MaxLength       =   5
      TabIndex        =   54
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox margen12 
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
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   53
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox pventa12 
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
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   52
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox maximo12 
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
      Left            =   6120
      MaxLength       =   5
      TabIndex        =   51
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox minimo12 
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
      Left            =   5640
      MaxLength       =   5
      TabIndex        =   50
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox unidad10 
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
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   49
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox factor10 
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
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   48
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox pventa10 
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
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   47
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox margen10 
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   46
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox unidad9 
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
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   45
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox factor9 
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
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   44
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox pventa9 
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
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   43
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox margen9 
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   42
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox unidad8 
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
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   41
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox factor8 
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
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   40
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox pventa8 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   39
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox margen8 
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   38
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox unidad7 
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
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   37
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox factor7 
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
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   36
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox pventa7 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   35
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox margen7 
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   34
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox unidad6 
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
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   33
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox factor6 
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
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   32
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox pventa6 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   31
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox margen6 
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   30
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox unidad5 
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
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   29
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox factor5 
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
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   28
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox pventa5 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   27
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox margen5 
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   26
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox unidad4 
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
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   25
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox factor4 
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
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   24
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox pventa4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   23
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox margen4 
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   22
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox unidad3 
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
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   21
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox factor3 
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
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   20
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox pventa3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   19
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox margen3 
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   18
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox unidad2 
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
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   17
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox factor2 
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
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   16
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox pventa2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   15
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox margen2 
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   14
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox dscto 
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
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   13
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox minimo11 
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
      Left            =   5640
      MaxLength       =   5
      TabIndex        =   12
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox maximo11 
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
      Left            =   6120
      MaxLength       =   5
      TabIndex        =   11
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox pventa11 
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
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox margen11 
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
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox fechafd 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   8
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox fechaid 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox fechaf11 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox fechai11 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox margen1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox pventa1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox factor1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox unidad1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.ComboBox local2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid dbgrid4 
      Height          =   2295
      Left            =   0
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   5400
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777152
      HeadLines       =   1
      RowHeight       =   13
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
         Name            =   "MS Serif"
         Size            =   6.75
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
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaVence"
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
      Left            =   2160
      TabIndex        =   102
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
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
      Left            =   0
      TabIndex        =   99
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
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
      Left            =   2160
      TabIndex        =   97
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mon.Costo"
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
      Left            =   0
      TabIndex        =   92
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unidad"
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
      Left            =   0
      TabIndex        =   91
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label28 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Factor"
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
      Left            =   0
      TabIndex        =   90
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label29 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CostoUlt."
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
      Left            =   0
      TabIndex        =   89
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image foto 
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Left            =   0
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label fotonombre 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   255
      Left            =   120
      TabIndex        =   88
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label35 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CostoProm."
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
      Left            =   0
      TabIndex        =   87
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label57 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Foto"
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
      Left            =   0
      TabIndex        =   86
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label54 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "Oferta.precio=0 acepta"
      Height          =   195
      Left            =   5640
      TabIndex        =   85
      Top             =   4800
      Width           =   1635
   End
   Begin VB.Label Label49 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dscto Autom."
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
      Left            =   5640
      TabIndex        =   84
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label48 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaFinal"
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
      Left            =   5640
      TabIndex        =   83
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label47 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaInicio"
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
      Left            =   5640
      TabIndex        =   82
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label46 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaFinal"
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
      Left            =   5640
      TabIndex        =   81
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label45 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaInicio"
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
      Left            =   5640
      TabIndex        =   80
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label44 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%Margen"
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
      Left            =   7680
      TabIndex        =   79
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label43 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pvta"
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
      Left            =   6600
      TabIndex        =   78
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label42 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Max"
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
      Left            =   6120
      TabIndex        =   77
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label41 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Min"
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
      Left            =   5640
      TabIndex        =   76
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label40 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%Margen"
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
      Left            =   4680
      TabIndex        =   75
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label39 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pvta"
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
      Left            =   3720
      TabIndex        =   74
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label38 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Factor"
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
      Left            =   3000
      TabIndex        =   73
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label37 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Und"
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
      Left            =   2160
      TabIndex        =   72
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label36 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mon.Pvta"
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
      Left            =   2160
      TabIndex        =   71
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label33 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ListaNro"
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
      Left            =   3720
      TabIndex        =   70
      Top             =   360
      Width           =   975
   End
   Begin VB.Menu lo23 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "changepr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   Frame1.Visible = False
   'ccosto.SetFocus
   Exit Sub
End If
Command1_Click

End Sub


Private Sub cmdGrabar_Click()
Dim found As Integer
Dim mytablex As New ADODB.Recordset
If MsgBox("Esta seguro", 1, "Aviso") <> 1 Then Exit Sub

mytablex.Open "SELECT * FROM userlocal where codigo='" & gusuario & "' and local='" & local2 & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.Close
      MsgBox "Usuario No autorizado,utilizar este local ", 48, "Aviso"
      Exit Sub
   End If
   mytablex.Close
graba_fecha "" & codigo
found = graba_precios("" & codigo)
If found = 1 Then
   changepr.Hide
   Unload changepr
End If
End Sub

Private Sub Command1_Click()
ejecuta 1
End Sub
Sub ejecuta(sw As Integer)
Dim buf As String
If Len(buffer) = 0 Then
buf = "select Descripcio,ccosto from ccosto "
Else
buf = "select Descripcio,CCosto from ccosto where " & Combo1 & " like '" & buffer & "%'"
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
               dbGrid1.columns(0).Width = 4000
               dbGrid1.columns(1).Width = 2000
If sw = 1 Then
               dbGrid1.SetFocus
End If
End Sub


Private Sub Command2_Click()

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
   'ccosto = DBGrid1.Columns(1)
   'Frame1.Visible = False
   'ccosto.SetFocus
End If

End Sub

Private Sub factor1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
pventa1.SetFocus

End Sub

Private Sub factor2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
pventa2.SetFocus

End Sub

Private Sub factor3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
pventa3.SetFocus

End Sub

Private Sub factor4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
pventa4.SetFocus

End Sub

Private Sub factor5_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
factor5.SetFocus

End Sub

Private Sub Form_Activate()
busca_fecha "" & codigo
pone_fotos
carga_precios "" & codigo
consulta_precios "" & codigo
If Len(unidad1) = 0 Then
   unidad1 = "UND"
End If
If Len(factor1) = 0 Then
   factor1 = "1"
End If
End Sub
Sub pone_fotos()
foto = LoadPicture()
'fotonombre = "" & mytablex.Fields("fotonombre")
If Len(fotonombre) > 0 Then
If existe_archivo(fotonombre) > 0 Then
   foto = LoadPicture(fotonombre)
End If
End If
End Sub

Private Sub Form_Load()
local2.Clear
local2.AddItem "01"
local2.AddItem "02"
local2.AddItem "03"
local2.AddItem "04"
local2.ListIndex = 0


End Sub
Function busca_cambio() As Double
Dim mytablex As New ADODB.Recordset
Dim sdx As Double
sdx = 1
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "Select * from parame where codigfo='01'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   sdx = Val("" & mytablex.Fields("paricomp"))
   If Val("" & mytablex.Fields("paricomp")) <= 0 Then
      sdx = 1
   End If
End If
busca_cambio = sdx
mytablex.Close
End Function

Sub carga_precios(buf As String)
Dim mytablex As New ADODB.Recordset
inicializa_precios
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "Select * from precios where producto='" & buf & "' and local='" & local2 & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   pone_xprecio mytablex
   calcula_margenes
End If
mytablex.Close
'pventa1.SetFocus
End Sub
Sub pone_xprecio(mytablex As ADODB.Recordset)
unidad1 = "" & mytablex.Fields("unidad1")
unidad2 = "" & mytablex.Fields("unidad2")
unidad3 = "" & mytablex.Fields("unidad3")
unidad4 = "" & mytablex.Fields("unidad4")
unidad5 = "" & mytablex.Fields("unidad5")
unidad6 = "" & mytablex.Fields("unidad6")
unidad7 = "" & mytablex.Fields("unidad7")
unidad8 = "" & mytablex.Fields("unidad8")
unidad9 = "" & mytablex.Fields("unidad9")
unidad10 = "" & mytablex.Fields("unidad10")
factor1 = "" & mytablex.Fields("factor1")
factor2 = "" & mytablex.Fields("factor2")
factor3 = "" & mytablex.Fields("factor3")
factor4 = "" & mytablex.Fields("factor4")
factor5 = "" & mytablex.Fields("factor5")
factor6 = "" & mytablex.Fields("factor6")
factor7 = "" & mytablex.Fields("factor7")
factor8 = "" & mytablex.Fields("factor8")
factor9 = "" & mytablex.Fields("factor9")
factor10 = "" & mytablex.Fields("factor10")
pventa1 = "" & mytablex.Fields("pventa1")
pventa2 = "" & mytablex.Fields("pventa2")
pventa3 = "" & mytablex.Fields("pventa3")
pventa4 = "" & mytablex.Fields("pventa4")
pventa5 = "" & mytablex.Fields("pventa5")
pventa6 = "" & mytablex.Fields("pventa6")
pventa7 = "" & mytablex.Fields("pventa7")
pventa8 = "" & mytablex.Fields("pventa8")
pventa9 = "" & mytablex.Fields("pventa9")
pventa10 = "" & mytablex.Fields("pventa10")
margen1 = "" & mytablex.Fields("margen1")
margen2 = "" & mytablex.Fields("margen2")
margen3 = "" & mytablex.Fields("margen3")
margen4 = "" & mytablex.Fields("margen4")
margen5 = "" & mytablex.Fields("margen5")
margen6 = "" & mytablex.Fields("margen6")
margen7 = "" & mytablex.Fields("margen7")
margen8 = "" & mytablex.Fields("margen8")
margen9 = "" & mytablex.Fields("margen9")
margen10 = "" & mytablex.Fields("margen10")
minimo11 = "" & mytablex.Fields("minimo11")
minimo12 = "" & mytablex.Fields("minimo12")
minimo13 = "" & mytablex.Fields("minimo13")
minimo14 = "" & mytablex.Fields("minimo14")
minimo15 = "" & mytablex.Fields("minimo15")
maximo11 = "" & mytablex.Fields("maximo11")
maximo12 = "" & mytablex.Fields("maximo12")
maximo13 = "" & mytablex.Fields("maximo13")
maximo14 = "" & mytablex.Fields("maximo14")
maximo15 = "" & mytablex.Fields("maximo15")
pventa11 = "" & mytablex.Fields("pventa11")
pventa12 = "" & mytablex.Fields("pventa12")
pventa13 = "" & mytablex.Fields("pventa13")
pventa14 = "" & mytablex.Fields("pventa14")
pventa15 = "" & mytablex.Fields("pventa15")
margen11 = "" & mytablex.Fields("margen11")
margen12 = "" & mytablex.Fields("margen12")
margen13 = "" & mytablex.Fields("margen13")
margen14 = "" & mytablex.Fields("margen14")
margen15 = "" & mytablex.Fields("margen15")
fechai11 = "" & mytablex.Fields("fechai11")
fechaf11 = "" & mytablex.Fields("fechaf11")
fechaid = "" & mytablex.Fields("fechaid")
fechafd = "" & mytablex.Fields("fechafd")
dscto = "" & mytablex.Fields("dscto")
If Len(unidad1) = 0 Then
   unidad1 = "UND"
End If
If Len(factor1) = 0 Then
   factor1 = "1"
End If
End Sub
Sub inicializa_precios()
unidad1 = "UND"
unidad2 = ""
unidad3 = ""
unidad4 = ""
unidad5 = ""
unidad6 = ""
unidad7 = ""
unidad8 = ""
unidad9 = ""
unidad10 = ""
'saldoini = ""

factor1 = "1"
factor2 = ""
factor3 = ""
factor4 = ""
factor5 = ""
factor6 = ""
factor7 = ""
factor8 = ""
factor9 = ""
factor10 = ""

pventa1 = ""
pventa2 = ""
pventa3 = ""
pventa4 = ""
pventa5 = ""
pventa6 = ""
pventa7 = ""
pventa8 = ""
pventa9 = ""
pventa10 = ""

margen1 = ""
margen2 = ""
margen3 = ""
margen4 = ""
margen5 = ""
margen6 = ""
margen7 = ""
margen8 = ""
margen9 = ""
margen10 = ""
minimo11 = ""
minimo12 = ""
minimo13 = ""
minimo14 = ""
minimo15 = ""

maximo11 = ""
maximo12 = ""
maximo13 = ""
maximo14 = ""
maximo15 = ""

pventa11 = ""
pventa12 = ""
pventa13 = ""
pventa14 = ""
pventa15 = ""
margen11 = ""
margen12 = ""
margen13 = ""
margen14 = ""
margen15 = ""
fechai11 = ""
fechaf11 = ""
fechaid = ""
fechafd = ""
dscto = ""
'ccosto = ""
End Sub
Sub calcula_margenes()
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double
Dim acostou As String
On Error GoTo cmd786_err
sdx = Val(costou) + Val(flete)
acostou = "" & sdx
If Val(factor) <= 0 Then
   factor = "1"
End If
If Val(factor1) <= 0 Then
   factor1 = "1"
End If
If Val(costou) = 0 And Val(costop) = 0 Then
   margen1 = "0"
   margen2 = "0"
   margen3 = "0"
   margen4 = "0"
   margen5 = "0"
   margen6 = "0"
   margen7 = "0"
   margen8 = "0"
   margen9 = "0"
   margen10 = "0"
   
   margen11 = "0"
   margen12 = "0"
   margen13 = "0"
   margen14 = "0"
   margen15 = "0"
   Exit Sub
End If

          If monedac = "S" Then
             If monedav = "D" Then
                sdx = Val(acostou) / busca_cambio()
                If sdx <= 0 Then
                   sdx = 1
                End If
                acostou = "" & sdx
             End If
          End If
          If monedac = "D" Then
             If monedav = "S" Then
                sdx = Val(acostou) * busca_cambio()
                If sdx <= 0 Then
                   sdx = 1
                End If
                acostou = "" & sdx
             End If
          End If
       If Val(acostou) > 0 And Val(pventa1) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa1) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen1 = Format(sdx2, "0.00")
          GoTo siguiente1
       End If
       If Val(margen1) > 0 And Val(pventa1) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen1) / 100
          pventa1 = Format(sdx, "0.00")
          GoTo siguiente1
       End If
       If Val(acostou) <= 0 And Val(pventa1) > 0 And Val(margen1) > 0 Then
          sdx = Val(pventa1) / (1 + (Val(margen1) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente1
       End If
       
siguiente1:
       If Val(acostou) > 0 And Val(pventa2) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor2)
          sdx1 = Val(pventa2) '/ Val(factor1)
          sdx2 = (Val(sdx1) - sdx) * 100 / sdx
          margen2 = Format(sdx2, "0.00")
          GoTo siguiente2
       End If
       If Val(margen2) > 0 And Val(pventa2) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen2) / 100
          pventa2 = Format(sdx, "0.00")
          GoTo siguiente2
       End If
       If Val(acostou) <= 0 And Val(pventa2) > 0 And Val(margen2) > 0 Then
          sdx = Val(pventa2) / (1 + (Val(margen2) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente2
       End If
siguiente2:
       If Val(acostou) > 0 And Val(pventa3) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor3)
          sdx1 = Val(pventa3) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen3 = Format(sdx2, "0.00")
          GoTo siguiente3
       End If
       If Val(margen3) > 0 And Val(pventa3) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen3) / 100
          pventa3 = Format(sdx, "0.00")
          GoTo siguiente3
       End If
       If Val(acostou) <= 0 And Val(pventa3) > 0 And Val(margen3) > 0 Then
          sdx = Val(pventa3) / (1 + (Val(margen3) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente2
       End If
siguiente3:
If Val(acostou) > 0 And Val(pventa4) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor4)
          sdx1 = Val(pventa4) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen4 = Format(sdx2, "0.00")
          GoTo siguiente4
       End If
       If Val(margen4) > 0 And Val(pventa4) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen4) / 100
          pventa4 = Format(sdx, "0.00")
          GoTo siguiente4
       End If
       If Val(acostou) <= 0 And Val(pventa4) > 0 And Val(margen4) > 0 Then
          sdx = Val(pventa4) / (1 + (Val(margen4) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente4
       End If
siguiente4:
If Val(acostou) > 0 And Val(pventa5) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor5)
          sdx1 = Val(pventa5) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen5 = Format(sdx2, "0.00")
          GoTo siguiente5
       End If
       If Val(margen5) > 0 And Val(pventa5) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen5) / 100
          pventa5 = Format(sdx, "0.00")
          GoTo siguiente5
       End If
       If Val(acostou) <= 0 And Val(pventa5) > 0 And Val(margen5) > 0 Then
          sdx = Val(pventa5) / (1 + (Val(margen5) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente5
       End If
siguiente5:
If Val(acostou) > 0 And Val(pventa6) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor6)
          sdx1 = Val(pventa6) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen6 = Format(sdx2, "0.00")
          GoTo siguiente6
       End If
       If Val(margen6) > 0 And Val(pventa6) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen6) / 100
          pventa6 = Format(sdx, "0.00")
          GoTo siguiente6
       End If
       If Val(acostou) <= 0 And Val(pventa6) > 0 And Val(margen6) > 0 Then
          sdx = Val(pventa6) / (1 + (Val(margen6) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente6
       End If
siguiente6:
If Val(acostou) > 0 And Val(pventa7) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor7)
          sdx1 = Val(pventa7) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen7 = Format(sdx2, "0.00")
          GoTo siguiente7
       End If
       If Val(margen7) > 0 And Val(pventa7) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen7) / 100
          pventa7 = Format(sdx, "0.00")
          GoTo siguiente7
       End If
       If Val(acostou) <= 0 And Val(pventa7) > 0 And Val(margen7) > 0 Then
          sdx = Val(pventa7) / (1 + (Val(margen7) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente7
       End If
siguiente7:
If Val(costou) > 0 And Val(pventa8) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor8)
          sdx1 = Val(pventa8) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen8 = Format(sdx2, "0.00")
          GoTo siguiente8
       End If
       If Val(margen8) > 0 And Val(pventa8) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen8) / 100
          pventa8 = Format(sdx, "0.00")
          GoTo siguiente8
       End If
       If Val(acostou) <= 0 And Val(pventa8) > 0 And Val(margen8) > 0 Then
          sdx = Val(pventa8) / (1 + (Val(margen8) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente8
       End If
siguiente8:
If Val(acostou) > 0 And Val(pventa9) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor9)
          sdx1 = Val(pventa9) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen9 = Format(sdx2, "0.00")
          GoTo siguiente9
       End If
       If Val(margen9) > 0 And Val(pventa9) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen9) / 100
          pventa9 = Format(sdx, "0.00")
          GoTo siguiente9
       End If
       If Val(acostou) <= 0 And Val(pventa9) > 0 And Val(margen9) > 0 Then
          sdx = Val(pventa9) / (1 + (Val(margen9) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente9
       End If
siguiente9:
If Val(acostou) > 0 And Val(pventa10) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor10)
          sdx1 = Val(pventa10) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen10 = Format(sdx2, "0.00")
          GoTo siguiente10
       End If
       If Val(margen10) > 0 And Val(pventa10) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen10) / 100
          pventa2 = Format(sdx, "0.00")
          GoTo siguiente10
       End If
       If Val(acostou) <= 0 And Val(pventa10) > 0 And Val(margen10) > 0 Then
          sdx = Val(pventa10) / (1 + (Val(margen10) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente10
       End If
siguiente10:
If Val(acostou) > 0 And Val(pventa11) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa11) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen11 = Format(sdx2, "0.00")
          GoTo siguiente11
       End If
       If Val(margen11) > 0 And Val(pventa11) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen11) / 100
          pventa11 = Format(sdx, "0.00")
          GoTo siguiente11
       End If
       If Val(acostou) <= 0 And Val(pventa11) > 0 And Val(margen11) > 0 Then
          sdx = Val(pventa11) / (1 + (Val(margen11) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente11
       End If
siguiente11:
If Val(acostou) > 0 And Val(pventa12) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa12) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen12 = Format(sdx2, "0.00")
          GoTo siguiente12
       End If
       If Val(margen12) > 0 And Val(pventa12) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen12) / 100
          pventa12 = Format(sdx, "0.00")
          GoTo siguiente12
       End If
       If Val(acostou) <= 0 And Val(pventa12) > 0 And Val(margen12) > 0 Then
          sdx = Val(pventa12) / (1 + (Val(margen12) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente12
       End If
siguiente12:
If Val(acostou) > 0 And Val(pventa13) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa13) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen13 = Format(sdx2, "0.00")
          GoTo siguiente13
       End If
       If Val(margen13) > 0 And Val(pventa13) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen13) / 100
          pventa13 = Format(sdx, "0.00")
          GoTo siguiente13
       End If
       If Val(acostou) <= 0 And Val(pventa13) > 0 And Val(margen13) > 0 Then
          sdx = Val(pventa13) / (1 + (Val(margen13) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente13
       End If
siguiente13:
If Val(acostou) > 0 And Val(pventa14) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa14) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen14 = Format(sdx2, "0.00")
          GoTo siguiente14
       End If
       If Val(margen14) > 0 And Val(pventa14) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen14) / 100
          pventa14 = Format(sdx, "0.00")
          GoTo siguiente14
       End If
       If Val(acostou) <= 0 And Val(pventa14) > 0 And Val(margen14) > 0 Then
          sdx = Val(pventa14) / (1 + (Val(margen14) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente14
       End If
siguiente14:
If Val(acostou) > 0 And Val(pventa15) > 0 Then  'calculando margenes
          sdx = (Val(acostou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa15) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen15 = Format(sdx2, "0.00")
          GoTo siguiente15
       End If
       If Val(margen15) > 0 And Val(pventa15) <= 0 And Val(acostou) > 0 Then
          sdx = Val(acostou) + Val(acostou) * Val(margen15) / 100
          pventa15 = Format(sdx, "0.00")
          GoTo siguiente10
       End If
       If Val(acostou) <= 0 And Val(pventa15) > 0 And Val(margen15) > 0 Then
          sdx = Val(pventa15) / (1 + (Val(margen15) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente15
       End If
siguiente15:
       Exit Sub
cmd786_err:
MsgBox "Error en calcula margenes", 48, "Aviso"
Exit Sub
End Sub
Function graba_precios(buf As String)
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "Select * from precios where producto='" & codigo & "' and local='" & local2 & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   'mytablex.Edit
   graba_xprecio mytablex
   mytablex.Update
   graba_precios = 1
Else
   mytablex.AddNew
   mytablex.Fields("producto") = codigo
   mytablex.Fields("local") = local2
   graba_xprecio mytablex
   mytablex.Update
   graba_precios = 1
End If
mytablex.Close
End Function
Sub graba_xprecio(mytablex As ADODB.Recordset)
mytablex.Fields("unidad1") = unidad1
mytablex.Fields("unidad2") = unidad2
mytablex.Fields("unidad3") = unidad3
mytablex.Fields("unidad4") = unidad4
mytablex.Fields("unidad5") = unidad5
mytablex.Fields("unidad6") = unidad6
mytablex.Fields("unidad7") = unidad7
mytablex.Fields("unidad8") = unidad8
mytablex.Fields("unidad9") = unidad9
mytablex.Fields("unidad10") = unidad10
mytablex.Fields("factor1") = Val(factor1)
mytablex.Fields("factor2") = Val(factor2)
mytablex.Fields("factor3") = Val(factor3)
mytablex.Fields("factor4") = Val(factor4)
mytablex.Fields("factor5") = Val(factor5)
mytablex.Fields("factor6") = Val(factor6)
mytablex.Fields("factor7") = Val(factor7)
mytablex.Fields("factor8") = Val(factor8)
mytablex.Fields("factor9") = Val(factor9)
mytablex.Fields("factor10") = Val(factor10)
mytablex.Fields("pventa1") = Val(pventa1)
mytablex.Fields("pventa2") = Val(pventa2)
mytablex.Fields("pventa3") = Val(pventa3)
mytablex.Fields("pventa4") = Val(pventa4)
mytablex.Fields("pventa5") = Val(pventa5)
mytablex.Fields("pventa6") = Val(pventa6)
mytablex.Fields("pventa7") = Val(pventa7)
mytablex.Fields("pventa8") = Val(pventa8)
mytablex.Fields("pventa9") = Val(pventa9)
mytablex.Fields("pventa10") = Val(pventa10)
mytablex.Fields("margen1") = Val(margen1)
mytablex.Fields("margen2") = Val(margen2)
mytablex.Fields("margen3") = Val(margen3)
mytablex.Fields("margen4") = Val(margen4)
mytablex.Fields("margen5") = Val(margen5)
mytablex.Fields("margen6") = Val(margen6)
mytablex.Fields("margen7") = Val(margen7)
mytablex.Fields("margen8") = Val(margen8)
mytablex.Fields("margen9") = Val(margen9)
mytablex.Fields("margen10") = Val(margen10)
mytablex.Fields("minimo11") = Val(minimo11)
mytablex.Fields("minimo12") = Val(minimo12)
mytablex.Fields("minimo13") = Val(minimo13)
mytablex.Fields("minimo14") = Val(minimo14)
mytablex.Fields("minimo15") = Val(minimo15)
mytablex.Fields("maximo11") = Val(maximo11)
mytablex.Fields("maximo12") = Val(maximo12)
mytablex.Fields("maximo13") = Val(maximo13)
mytablex.Fields("maximo14") = Val(maximo14)
mytablex.Fields("maximo15") = Val(maximo15)
mytablex.Fields("pventa11") = Val(pventa11)
mytablex.Fields("pventa12") = Val(pventa12)
mytablex.Fields("pventa13") = Val(pventa13)
mytablex.Fields("pventa14") = Val(pventa14)
mytablex.Fields("pventa15") = Val(pventa15)
mytablex.Fields("margen11") = Val(margen11)
mytablex.Fields("margen12") = Val(margen12)
mytablex.Fields("margen13") = Val(margen13)
mytablex.Fields("margen14") = Val(margen14)
mytablex.Fields("margen15") = Val(margen15)
If IsDate(fechai11) Then
mytablex.Fields("fechai11") = fechai11
End If
If IsDate(fechaf11) Then
mytablex.Fields("fechaf11") = fechaf11
End If
If IsDate(fechaid) Then
mytablex.Fields("fechaid") = fechaid
End If
If IsDate(fechafd) Then
mytablex.Fields("fechafd") = fechafd
End If
mytablex.Fields("dscto") = Val(dscto)
End Sub




Private Sub lo23_Click()
If Frame1.Visible = True Then
   'ccosto.SetFocus
   Exit Sub
End If
changepr.Hide
Unload changepr
End Sub

Private Sub local2_Click()
'MsgBox "xx"
carga_precios "" & codigo
If Len(unidad1) = 0 Then
   unidad1 = "UND"
End If
If Len(factor1) = 0 Then
   factor1 = "1"
End If

End Sub

Private Sub margen1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
unidad2.SetFocus

End Sub

Private Sub margen2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
unidad3.SetFocus

End Sub

Private Sub margen3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
unidad4.SetFocus

End Sub

Private Sub margen4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
unidad5.SetFocus

End Sub

Private Sub margen5_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
factor6.SetFocus

End Sub

Private Sub pventa1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   lo23_Click
   Exit Sub
End If
calcula_margenes
margen1.SetFocus

End Sub

Private Sub pventa2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes
margen2.SetFocus

End Sub

Private Sub pventa3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes
margen3.SetFocus

End Sub

Private Sub pventa4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes
margen4.SetFocus

End Sub

Private Sub pventa5_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
calcula_margenes
margen5.SetFocus

End Sub

Private Sub unidad1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
factor1.SetFocus
End Sub

Private Sub unidad2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
factor2.SetFocus

End Sub

Private Sub unidad3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
factor3.SetFocus

End Sub

Private Sub unidad4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
factor4.SetFocus

End Sub

Private Sub unidad5_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
factor5.SetFocus

End Sub
Sub consulta_precios(buf1 As String)
Dim buf As String
On Error GoTo cmdo0_err
Dim mytablex As New ADODB.Recordset
   buf = "select Local AS Lista,Unidad1 as Und1,Pventa1,"
   buf = buf & " Unidad2 as Und2,Pventa2,"
   buf = buf & " Unidad3 as Und3,Pventa3,"
   buf = buf & " Unidad4 as Und4,Pventa4,"
   buf = buf & " Unidad5 as Und5,Pventa5,"
   buf = buf & " Unidad6 as Und6,Pventa6,"
   buf = buf & " Unidad7 as Und7,Pventa7,"
   buf = buf & " Unidad8 as Und8,Pventa8,"
   buf = buf & " Unidad9 as Und9,Pventa9,"
   buf = buf & " Unidad10 as Und10,Pventa10"
   buf = buf & "  from precios where producto='" & buf1 & "'"
   If mytablex.State = 1 Then
      mytablex.Close
   End If
   mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
   If mytablex.EOF = True And mytablex.BOF = True Then
   End If
   Set dbgrid4.DataSource = mytablex
   dbgrid4.columns(0).Width = 600
   dbgrid4.columns(1).Width = 600
   dbgrid4.columns(2).Width = 600
   dbgrid4.columns(3).Width = 800
   dbgrid4.columns(4).Width = 600
   dbgrid4.columns(5).Width = 800
   dbgrid4.columns(6).Width = 600
   dbgrid4.columns(7).Width = 800
   dbgrid4.columns(8).Width = 600
   dbgrid4.columns(9).Width = 800
   dbgrid4.columns(10).Width = 600
   dbgrid4.columns(11).Width = 800
   dbgrid4.columns(12).Width = 600
   dbgrid4.columns(13).Width = 800
   dbgrid4.columns(14).Width = 600
   dbgrid4.columns(15).Width = 800
   dbgrid4.columns(16).Width = 600
   dbgrid4.columns(17).Width = 800
   dbgrid4.columns(18).Width = 600
   dbgrid4.columns(19).Width = 800
   dbgrid4.columns(20).Width = 600
   Exit Sub
cmdo0_err:
   Exit Sub
End Sub
Sub busca_fecha(buf As String)
Dim mytablex As New ADODB.Recordset
fechavence = ""
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "Select * from producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   fechavence = "" & mytablex.Fields("fechavence")
End If
mytablex.Close

End Sub
Sub graba_fecha(buf As String)
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "Select * from producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If IsDate(fechavence) Then
      mytablex.Fields("fechavence") = Format(fechavence, "dd/mm/yyyy")
      mytablex.Update
   End If
End If
mytablex.Close

End Sub
