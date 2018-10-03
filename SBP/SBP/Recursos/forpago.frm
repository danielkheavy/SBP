VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form forpago 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   9990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
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
      Height          =   8415
      Left            =   120
      TabIndex        =   49
      Top             =   1440
      Visible         =   0   'False
      Width           =   9855
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
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   240
         Width           =   1935
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
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   6240
         TabIndex        =   51
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
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7335
         Left            =   0
         TabIndex        =   54
         Top             =   840
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   12938
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Letras"
      Height          =   4695
      Left            =   120
      TabIndex        =   35
      Top             =   1320
      Visible         =   0   'False
      Width           =   9615
      Begin VB.CommandButton Command5 
         BackColor       =   &H008080FF&
         Caption         =   "Close"
         Height          =   735
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FF80&
         Caption         =   "Acepta"
         Height          =   735
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Nuevo registro"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox numerole 
         Height          =   405
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox fechafle 
         Height          =   405
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox fechaile 
         Height          =   405
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro Letra"
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
         TabIndex        =   41
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         TabIndex        =   39
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label7 
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
         TabIndex        =   37
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Entrega"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   9615
      Begin VB.CommandButton Command9 
         BackColor       =   &H008080FF&
         Caption         =   "Close"
         Height          =   735
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0080FF80&
         Caption         =   "Acepta"
         Height          =   735
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   360
         Width           =   975
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
         Height          =   735
         Left            =   8400
         Picture         =   "forpago.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Nuevo registro"
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox campo1 
         Height          =   405
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox campo2 
         Height          =   405
         Left            =   1560
         MaxLength       =   60
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox campo3 
         Height          =   405
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox campo4 
         Height          =   405
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox campo5 
         Height          =   405
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label fpago 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5160
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label descripcio1 
         BackColor       =   &H00FFFFC0&
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
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label descripcio2 
         BackColor       =   &H00FFFFC0&
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
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label descripcio3 
         BackColor       =   &H00FFFFC0&
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
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label descripcio4 
         BackColor       =   &H00FFFFC0&
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
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label descripcio5 
         BackColor       =   &H00FFFFC0&
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
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label moneda 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5160
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
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
      Height          =   735
      Left            =   4680
      Picture         =   "forpago.frx":1212
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Nuevo registro"
      Top             =   7080
      Width           =   975
   End
   Begin VB.Frame Framefp 
      BackColor       =   &H00808080&
      Caption         =   "FormaPago"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   8295
      Left            =   15
      TabIndex        =   1
      Top             =   45
      Width           =   9855
      Begin VB.CommandButton Command7 
         BackColor       =   &H008080FF&
         Caption         =   "Close"
         Height          =   735
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   6960
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H0080FF80&
         Caption         =   "Acepta"
         Height          =   735
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   6960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Data Data9 
         Caption         =   "Data2"
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
         Top             =   7680
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBGrid.DBGrid DBGrid9 
         Bindings        =   "forpago.frx":2424
         Height          =   4335
         Left            =   120
         OleObjectBlob   =   "forpago.frx":2438
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1560
         Width           =   4335
      End
      Begin MSDataGridLib.DataGrid dbgrid10 
         Height          =   4215
         Left            =   4560
         TabIndex        =   34
         Top             =   1680
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   7435
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COMO PAGA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TIPOS DE PAGO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4560
         TabIndex        =   31
         Top             =   1200
         Width           =   5175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label tipoclie 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5160
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
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
         Left            =   5880
         TabIndex        =   15
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label fechadia 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   7200
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FALTA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   5880
         Width           =   2055
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.Label txtotals 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "US$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   600
         Width           =   495
      End
      Begin VB.Label txtotald 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2640
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T/Cambio"
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
         Left            =   5880
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label paridad 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   7200
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   5880
         Width           =   495
      End
      Begin VB.Label stxtotals 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2640
         TabIndex        =   4
         Top             =   5880
         Width           =   1815
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "US$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   6240
         Width           =   495
      End
      Begin VB.Label stxtotald 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2640
         TabIndex        =   2
         Top             =   6240
         Width           =   1815
      End
   End
   Begin VB.Menu ldo343 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "forpago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        ldo343_Click
        Exit Sub

    End If

    Command2_Click

End Sub

Private Sub campo1_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Frame2.Visible = False
        Framefp.Enabled = True
        dbgrid10.Enabled = True
        dbgrid10.SetFocus
        Exit Sub

    End If

    If Len(CAMPO1) > 0 Then
        found = busca_codigo("" & CAMPO1)

    End If

    CAMPO2.SetFocus

End Sub

Private Sub campo1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        Frame2.Visible = False
        dbgrid10.Enabled = True
        dbgrid10.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_cliente

    End If

End Sub

Private Sub campo2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    campo3.SetFocus

End Sub

Private Sub campo2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        CAMPO1.SetFocus
        Exit Sub

    End If

End Sub

Private Sub campo3_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    campo4.SetFocus

End Sub

Private Sub campo3_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        CAMPO2.SetFocus
        Exit Sub

    End If

End Sub

Private Sub campo4_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    campo5.SetFocus

End Sub

Private Sub campo4_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        campo3.SetFocus
        Exit Sub

    End If

End Sub

Private Sub campo5_KeyPress(KeyAscii As Integer)

    Dim sdx As Double

    Dim ind As Integer

    If KeyAscii <> 13 Then Exit Sub
    Data9.Recordset.AddNew
    Data9.Recordset.Fields("descripcio") = "" & dbgrid10.columns(0)
    Data9.Recordset.Fields("fpago") = "" & dbgrid10.columns(1)
    Data9.Recordset.Fields("moneda") = "" & dbgrid10.columns(2)
    Data9.Recordset.Fields("codigo") = CAMPO1
    Data9.Recordset.Fields("nombre") = CAMPO2
    Data9.Recordset.Fields("orden") = campo3
    Data9.Recordset.Fields("observa") = campo4
    Data9.Recordset.Fields("dias") = campo5
    Data9.Recordset.Update
    Frame2.Visible = False
    Framefp.Enabled = True
    ind = dbgrid9.Row

    If ind < 0 Then
        ind = 0

    End If

    dbgrid9.Row = ind
    dbgrid9.Col = 2
    dbgrid9.SetFocus

End Sub

Private Sub campo5_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        campo4.SetFocus
        Exit Sub

    End If

End Sub

Private Sub cmdDelete_Click()

End Sub

Private Sub cmdAddEntry_Click()

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    borra_pagos

End Sub

Private Sub Command2_Click()

    Dim rconsulta As New ADODB.Recordset

    Dim buf       As String

    Dim buf1      As String

    If tipoclie = "C" Then
        buf1 = "clientes"

    End If

    If tipoclie = "P" Then
        buf1 = "proveedo"

    End If

    If opcion1 = "1" Or opcion1 = "199" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from  " & buf1
        Else
            buf = "select Nombre,Codigo from " & buf1 & "'  where " & Combo1 & " like '" & buffer & "%'"

        End If

    End If
  
    If Combo2.ListIndex >= 1 Then
        buf = buf & " order by " & Combo2

    End If

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = rconsulta
   
    If rconsulta.EOF = True And rconsulta.BOF = True Then
        rconsulta.Close
        buffer.SetFocus
        Exit Sub

    End If
   
    If opcion1 = "1" Or opcion1 = "199" Then
        dbGrid1.columns(0).Width = 4000
        dbGrid1.columns(1).Width = 2000

    End If

    dbGrid1.SetFocus

End Sub

Private Sub Command3_Click()
    inicializa_letra
    numerole.SetFocus

End Sub

Private Sub Command4_Click()

    Dim ind   As Integer

    Dim found As Integer

    If Len(numerole) = 0 Then
        numerole.SetFocus
        Exit Sub

    End If

    If Len(fechaile) <> 10 And Not IsDate(fechaile) Then
        fechaile = ""
        Exit Sub

    End If

    If Len(fechafle) <> 10 And Not IsDate(fechafle) Then
        fechafle = ""
        Exit Sub

    End If

    Data9.Recordset.AddNew
    Data9.Recordset.Fields("descripcio") = "" & dbgrid10.columns(0)
    Data9.Recordset.Fields("fpago") = "" & dbgrid10.columns(1)
    Data9.Recordset.Fields("moneda") = "" & dbgrid10.columns(2)
    'Data9.Recordset.Fields("codigo") = codigo
    'Data9.Recordset.Fields("nombre") = nombre
    Data9.Recordset.Fields("orden") = fechaile
    Data9.Recordset.Fields("observa") = fechafle
    Data9.Recordset.Fields("dias") = numerole
    Data9.Recordset.Update
    Frame3.Visible = False
    ind = dbgrid9.Row

    If ind < 0 Then
        ind = 0

    End If

    dbgrid9.Row = ind
    dbgrid9.Col = 2
    dbgrid9.SetFocus

End Sub

Private Sub Command5_Click()
    ldo343_Click

End Sub

Private Sub Command7_Click()
    ldo343_Click

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            CAMPO1 = dbGrid1.columns(1)
            Frame1.Visible = False
            CAMPO1.SetFocus
            campo1_KeyPress 13

        End If

        If opcion1 = "199" Then
            'codigole = Trim(dbGrid1.columns(1))
            'nombrele = Trim(dbGrid1.columns(0))
            Frame1.Visible = False

            'codigole.SetFocus
            'codigole_KeyPress 13
        End If

    End If

End Sub

Private Sub DBGrid10_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim ind As Integer

    If KeyCode <> 13 And KeyCode <> 27 Then Exit Sub
    If KeyCode = 27 Then
        ldo343_Click
        Exit Sub

    End If

    If Val(stxtotals) <= 0 Then
        If MsgBox("Desea Grabar", 1, "Aviso") <> 1 Then Exit Sub
        opcion2 = 10000
        forpago.Hide
        Unload forpago
        Exit Sub

    End If

    Frame2.Caption = "" & dbgrid10.columns(0)
    fpago = "" & dbgrid10.columns(1)
    moneda = "" & dbgrid10.columns(2)

    If "" & dbgrid10.columns(3) = "A" Or "" & dbgrid10.columns(3) = "B" Or "" & dbgrid10.columns(3) = "E" Then  'efectivo,dolares,euros
        'recibe.SetFocus
        ind = dbgrid9.Row

        If ind < 0 Then
            ind = 0

        End If

        macro_inserta_registro
        dbgrid9.Row = ind
        dbgrid9.Col = 2
        dbgrid10.Enabled = False
        dbgrid9.SetFocus
        Exit Sub

    End If

    If "" & dbgrid10.columns(3) = "C" Then   'credito
        Framefp.Enabled = False
        macro_credito 0
        CAMPO1.SetFocus

    End If

    If "" & dbgrid10.columns(3) = "D" Then   'tarejta credito
        Framefp.Enabled = False
        macro_credito 0
        CAMPO1.SetFocus

    End If

    If "" & dbgrid10.columns(3) = "F" Then   'TARJETA DEBITO
        Framefp.Enabled = False
        macro_credito 0
        CAMPO1.SetFocus

    End If

    If "" & dbgrid10.columns(3) = "G" Then   'letra
        Framefp.Enabled = False
        inicializa_letra
        fechaile = Format(Now, "dd/mm/yyyy")
        fechafle = Format(Now, "dd/mm/yyyy")
        Frame3.Visible = True
        numerole.SetFocus

    End If

    If "" & dbgrid10.columns(3) = "H" Then   'bancos
        Framefp.Enabled = False
        macro_credito 1
        CAMPO1.SetFocus

    End If

    If "" & dbgrid10.columns(3) = "I" Then   'CHEQUES
        Framefp.Enabled = False
        macro_credito 1
        CAMPO1.SetFocus

    End If

End Sub

Sub inicializa_letra()
    'codigole = ""
    'nombrele = ""
    numerole = ""
    fechaile = ""
    fechafle = ""

End Sub

Sub macro_inserta_registro()
    Data9.Recordset.AddNew
    Data9.Recordset.Fields("descripcio") = "" & dbgrid10.columns(0)
    Data9.Recordset.Fields("fpago") = "" & dbgrid10.columns(1)
    Data9.Recordset.Fields("moneda") = "" & dbgrid10.columns(2)
    Data9.Recordset.Update
    Data9.refresh

End Sub

Private Sub DBGrid9_AfterColUpdate(ByVal ColIndex As Integer)

    Select Case ColIndex

        Case 2
            suma_fpagov

            If Label2.Caption = "Vuelto" Then
                If MsgBox("Desea Grabar", 1, "Aviso") <> 1 Then Exit Sub
                opcion2 = 10000
                forpago.Hide
                Unload forpago

            End If

    End Select

End Sub

Private Sub DBGrid9_BeforeColEdit(ByVal ColIndex As Integer, _
                                  ByVal KeyAscii As Integer, _
                                  Cancel As Integer)

    If ColIndex <> 2 Then
        Cancel = True
        Exit Sub

    End If

    Select Case ColIndex

        Case 2

            If Len("" & dbgrid9.columns(0)) = 0 Then
                Cancel = True
                Exit Sub

            End If
            
    End Select

End Sub

Sub valida_ingresado()

    Dim sdx      As Double

    Dim xsoles   As Double

    Dim xdolares As Double

    Dim xfaltas  As Double

    Dim xfaltad  As Double

    Dim xvueltos As Double

    Dim xvueltod As Double

    Dim sdx1     As Double

    Dim sdx2     As Double

    xsoles = 0
    xdolares = 0
    xfaltas = 0
    xfaltad = 0
    xvueltos = 0
    xvueltod = 0

    If "" & dbgrid9.columns(1) = "S" Then
        xsoles = Val("" & dbgrid9.columns(2))
        xdolares = Val(Val("" & dbgrid9.columns(2))) / Val(paridad)

    End If

    If "" & dbgrid9.columns(1) = "D" Then
        xdolares = Val("" & dbgrid9.columns(2))
        xsoles = Val("" & dbgrid9.columns(2)) * Val(paridad)

    End If

    Data9.Recordset.Edit
    Data9.Recordset.Fields("recibes") = xsoles
    Data9.Recordset.Fields("recibed") = xdolares
    sdx1 = Val(stxtotals) - xsoles
    sdx2 = Val(stxtotald) - xdolares
    Data9.Recordset.Fields("saldos") = sdx1
    Data9.Recordset.Fields("saldod") = sdx2
    'Data9.Recordset.Fields("codigo") = campo1
    'Data9.Recordset.Fields("nombre") = campo2
    'Data9.Recordset.Fields("orden") = campo3
    'Data9.Recordset.Fields("observa") = campo4
    'Data9.Recordset.Fields("dias") = campo5
    Data9.Recordset.Update

    If sdx1 > 0 Then
        'DBGrid9.Col = 0
        Label2.Caption = "Falta"
        dbgrid10.Enabled = True
        dbgrid10.SetFocus

    End If

    If sdx1 <= 0 Then
        Label2.Caption = "Vuelto"

    End If

    stxtotals = Format(sdx1, "0.00")
    stxtotald = Format(sdx2, "0.00")

End Sub

Private Sub DBGrid9_BeforeColUpdate(ByVal ColIndex As Integer, _
                                    OldValue As Variant, _
                                    Cancel As Integer)

    Select Case ColIndex

        Case 2
            opcion2 = 0
            valida_ingresado
            
    End Select

End Sub

Private Sub DBGrid9_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Data9.Recordset.Delete
        dbgrid10.Enabled = True
        dbgrid10.SetFocus
        Exit Sub

    End If

End Sub

Private Sub DBGrid9_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error GoTo cmd8912_err

    If KeyCode = &H2E Then  'borrar linea
        If dbgrid9.Row = -1 Then
            Exit Sub

        End If

        Data9.Recordset.Delete

        If Data9.Recordset.EOF = True And Data9.Recordset.BOF = True Then
            Exit Sub

        End If

        Exit Sub

    End If

    Exit Sub
cmd8912_err:
    Exit Sub

End Sub

Private Sub fechafle_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fechafle) <> 10 And Not IsDate(fechafle) Then
        fechafle = ""
        Exit Sub

    End If

End Sub

Private Sub fechafle_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fechaile.SetFocus
        Exit Sub

    End If

End Sub

Private Sub fechaile_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(fechaile) <> 10 And Not IsDate(fechaile) Then
        fechaile = ""
        Exit Sub

    End If

    fechafle.SetFocus

End Sub

Private Sub fechaile_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        numerole.SetFocus
        Exit Sub

    End If

End Sub

Private Sub Form_Load()
    Framefp.Top = 10: Framefp.Left = 10
    Frame1.Top = 10: Frame1.Left = 10
    Frame3.Top = 10: Frame3.Left = 10

    Dim found As Integer

    borra_pagos
    sql_formapago
    sql_pagos

End Sub

Sub sql_formapago()

    Dim rconsulta As New ADODB.Recordset

    Dim buf       As String

    buf = ""

    If anticipoo = "S" Then
        buf = " where tipo='A'  or tipo='B'  or tipo='D' or tipo='E' OR tipo='F' or tipo='G' or tipo='H' or tipo='V' "

    End If

    If anticipoo = "B" Then  'deposito a bancos
        buf = " where tipo='H' "

    End If

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open "select Descripcio,Fpago,Moneda,Tipo,Dias from fpago " & buf, cn, adOpenStatic, adLockOptimistic
    Set dbgrid10.DataSource = rconsulta
    dbgrid10.columns(0).Width = 4000
    dbgrid10.columns(1).Width = 600
    dbgrid10.columns(2).Width = 600
    dbgrid10.columns(3).Width = 600
    dbgrid10.columns(4).Width = 600
               
End Sub

Sub sql_pagos()

    Data9.Connect = "foxpro 2.5;"
    Data9.DatabaseName = globaldat
    Data9.RecordSource = "select * from  " & fpusuarior
    Data9.refresh

End Sub

Function macro_credito(sw As Integer)
    Frame2.Visible = True
    descripcio1.Visible = True
    descripcio2.Visible = True
    descripcio3.Visible = True
    descripcio4.Visible = True
    descripcio5.Visible = True
    CAMPO1.MaxLength = 11
    CAMPO2.MaxLength = 60
    campo3.MaxLength = 15
    campo4.MaxLength = 30
    campo5.MaxLength = 3
    'campo1 = ""
    'campo2 = ""
    campo3 = ""
    campo4 = ""
    campo5 = ""
    CAMPO1.Visible = True
    CAMPO2.Visible = True
    campo3.Visible = True
    campo4.Visible = True
    campo5.Visible = True
   
    descripcio1 = "Codigo"
    descripcio2 = "Nombre"
    descripcio3 = "Orden"
    descripcio4 = "Observacion"
    descripcio5 = "NroDias"

    If sw = 1 Then
        descripcio3 = "Banco"
        descripcio4 = "Cuenta"
        descripcio5 = "Numero"
        campo3.MaxLength = 6
        campo4.MaxLength = 20
        campo5.MaxLength = 10

    End If
   
End Function

Private Sub ldo343_Click()

    If Frame1.Visible = True Then
        If opcion1 = "1" Then
            Frame1.Visible = False
            CAMPO1.SetFocus
            Exit Sub

        End If

    End If

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Framefp.Enabled = True
        dbgrid10.Enabled = True
        dbgrid10.SetFocus
        Exit Sub

    End If

    If Frame3.Visible = True Then
        Frame3.Visible = False
        dbgrid10.Enabled = True
        dbgrid10.SetFocus
        Exit Sub

    End If

    forpago.Hide
    Unload forpago

End Sub

Private Sub recibe_KeyPress(KeyAscii As Integer)

End Sub

Private Sub recibe_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        campo5.SetFocus
        Exit Sub

    End If

End Sub

Sub borra_pagos()
    
    mydbxglo.Execute "DELETE FROM " & fpusuarior
     
    Data9.refresh
    Label2.Caption = "Falta"
    stxtotals = txtotals
    stxtotald = txtotald
    
End Sub

Function busca_codigo(buf As String)

    Dim mytablex As New ADODB.Recordset

    Dim buf1     As String

    If tipoclie = "C" Then
        buf1 = "clientes"

    End If

    If tipoclie = "P" Then
        buf1 = "proveedo"

    End If

    If tipoclie = "V" Then
        buf1 = "vendedor"

    End If

    mytablex.Open "select * from " & buf1 & " where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        Exit Function

    End If

    busca_codigo = 1
    mytablex.Close
 
End Function

Sub consulta_cliente()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "1"
    Command2_Click

End Sub

Sub consulta_letra()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "199"
    Command2_Click

End Sub

Sub suma_fpagov()

    Dim sdxs As Double

    Dim sdxd As Double

    Dim sdx  As Double

    Dim sdx1 As Double

    On Error GoTo cmd7812_err

    Label2.Caption = "Falta"
    sdxs = Val(txtotals)  'saldoa
    stxtotals = Format(sdxs, "0.00")
    'Data9.Recordset.MoveFirst
    Data9.refresh
    Do

        If Data9.Recordset.EOF Then Exit Do
        Data9.Recordset.Edit
        sdx = Val("" & Data9.Recordset.Fields("recibe"))

        If "" & Data9.Recordset.Fields("moneda") = "D" Then
            sdx = sdx * Val(paridad) 'Val("" & Data9.Recordset.Fields("cambio"))
            sdx = Val(Format(sdx, "0.00"))

        End If

        If sdx >= sdxs Then
            sdx1 = -sdx + sdxs
            sdx1 = Val(Format(sdx1, "0.00"))
            Data9.Recordset.Fields("total") = sdxs
            Data9.Recordset.Fields("saldos") = sdx1
            stxtotals = Format(sdx1, "0.00")
            sdxs = 0
            GoTo conmuta

        End If

        If sdxs > sdx Then
            sdx1 = sdxs - sdx
            sdx1 = Val(Format(sdx1, "0.00"))
            Data9.Recordset.Fields("total") = sdx
            Data9.Recordset.Fields("saldos") = 0
            stxtotals = Format(sdx1, "0.00")
            sdxs = sdx1

        End If

conmuta:
        Data9.Recordset.Update
        Data9.Recordset.MoveNext
    Loop
    stxtotald = Format(0, "0.00")

    If Val(paridad) > 0 Then
        sdx = Val(stxtotals) / Val(paridad)
        stxtotald = Format(sdx, "0.00")

    End If

    If stxtotals <= 0 Then
        Label2.Caption = "Vuelto"

        'DBGrid10.SetFocus
    End If

    Exit Sub
cmd7812_err:
    MsgBox "Error en " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub numerole_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(numerole) = 0 Then
        numerole.SetFocus
        Exit Sub

    End If

    fechaile.SetFocus

End Sub

Private Sub numerole_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        Exit Sub

    End If

End Sub
