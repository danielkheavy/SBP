VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PARKEs 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salida de vehiculo"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   14790
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
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
      Height          =   7695
      Left            =   45
      TabIndex        =   84
      Top             =   45
      Visible         =   0   'False
      Width           =   14655
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   6615
         Left            =   120
         TabIndex        =   88
         Top             =   960
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   11668
         _Version        =   393216
         HeadLines       =   2
         RowHeight       =   23
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
      Begin VB.CommandButton Command8 
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
         TabIndex        =   87
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
         TabIndex        =   86
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
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8640
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ProcesoCobrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   2520
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   8535
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Credito"
         Height          =   375
         Left            =   3480
         TabIndex        =   77
         Top             =   6000
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contado"
         Height          =   375
         Left            =   1560
         TabIndex        =   76
         Top             =   6000
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000FFFF&
         Caption         =   "SALIR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   6840
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NOTA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   5760
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1035
         ScaleWidth      =   3795
         TabIndex        =   61
         Top             =   1440
         Width           =   3855
         Begin VB.Label fecha 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1080
            TabIndex        =   63
            Top             =   120
            Width           =   2535
         End
         Begin VB.Image Image2 
            BorderStyle     =   1  'Fixed Single
            Height          =   1455
            Left            =   0
            Picture         =   "parkes.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   915
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Entrada"
            Height          =   375
            Left            =   1080
            TabIndex        =   62
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1035
         ScaleWidth      =   3795
         TabIndex        =   58
         Top             =   2640
         Width           =   3855
         Begin VB.Label hora 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1080
            TabIndex        =   60
            Top             =   120
            Width           =   2535
         End
         Begin VB.Image Image3 
            BorderStyle     =   1  'Fixed Single
            Height          =   1335
            Left            =   0
            Picture         =   "parkes.frx":0A4B
            Stretch         =   -1  'True
            Top             =   0
            Width           =   915
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hora Entrada"
            Height          =   375
            Left            =   1080
            TabIndex        =   59
            Top             =   600
            Width           =   2535
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   240
         ScaleHeight     =   1035
         ScaleWidth      =   7635
         TabIndex        =   55
         Top             =   240
         Width           =   7695
         Begin VB.Image Image1 
            BorderStyle     =   1  'Fixed Single
            Height          =   1455
            Left            =   0
            Picture         =   "parkes.frx":1AE7
            Stretch         =   -1  'True
            Top             =   0
            Width           =   915
         End
         Begin VB.Label placas 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1080
            TabIndex        =   57
            Top             =   120
            Width           =   5415
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Placa"
            Height          =   375
            Left            =   1080
            TabIndex        =   56
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   4080
         ScaleHeight     =   1155
         ScaleWidth      =   3795
         TabIndex        =   53
         Top             =   3840
         Width           =   3855
         Begin VB.TextBox total 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   915
            Left            =   960
            MaxLength       =   10
            TabIndex        =   68
            Top             =   60
            Width           =   2775
         End
         Begin VB.Image Image4 
            BorderStyle     =   1  'Fixed Single
            Height          =   975
            Left            =   0
            Picture         =   "parkes.frx":2532
            Stretch         =   -1  'True
            Top             =   0
            Width           =   915
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Total Cobrar"
            Height          =   375
            Left            =   1080
            TabIndex        =   54
            Top             =   960
            Width           =   2535
         End
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   4080
         ScaleHeight     =   1035
         ScaleWidth      =   3795
         TabIndex        =   50
         Top             =   1440
         Width           =   3855
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fecha Salida"
            Height          =   375
            Left            =   1080
            TabIndex        =   52
            Top             =   600
            Width           =   2535
         End
         Begin VB.Image Image5 
            BorderStyle     =   1  'Fixed Single
            Height          =   1455
            Left            =   0
            Picture         =   "parkes.frx":7966
            Stretch         =   -1  'True
            Top             =   0
            Width           =   915
         End
         Begin VB.Label fechas 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1080
            TabIndex        =   51
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   4080
         ScaleHeight     =   1035
         ScaleWidth      =   3795
         TabIndex        =   47
         Top             =   2640
         Width           =   3855
         Begin VB.Label Label9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Hora Salida"
            Height          =   375
            Left            =   1080
            TabIndex        =   49
            Top             =   600
            Width           =   2535
         End
         Begin VB.Image Image6 
            BorderStyle     =   1  'Fixed Single
            Height          =   1335
            Left            =   0
            Picture         =   "parkes.frx":83B1
            Stretch         =   -1  'True
            Top             =   0
            Width           =   915
         End
         Begin VB.Label horas 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1080
            TabIndex        =   48
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "BOLETA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FACTURA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   5160
         Width           =   1335
      End
      Begin VB.TextBox RUC 
         Height          =   375
         Left            =   840
         MaxLength       =   11
         TabIndex        =   44
         Top             =   5160
         Width           =   2535
      End
      Begin VB.TextBox NOMBRE 
         Height          =   375
         Left            =   840
         MaxLength       =   60
         TabIndex        =   43
         Top             =   5520
         Width           =   4335
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1155
         ScaleWidth      =   3795
         TabIndex        =   41
         Top             =   3840
         Width           =   3855
         Begin VB.TextBox valor 
            Height          =   495
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   82
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox cantidad 
            Height          =   495
            Left            =   2280
            MaxLength       =   10
            TabIndex        =   81
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PrecioxHora"
            Height          =   495
            Left            =   1080
            TabIndex        =   83
            Top             =   600
            Width           =   1215
         End
         Begin VB.Image Image7 
            BorderStyle     =   1  'Fixed Single
            Height          =   1335
            Left            =   0
            Picture         =   "parkes.frx":944D
            Stretch         =   -1  'True
            Top             =   0
            Width           =   915
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "NroHorasCobrar"
            Height          =   495
            Left            =   1080
            TabIndex        =   42
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Label xtipo 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3000
         TabIndex        =   75
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label acu 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3000
         TabIndex        =   74
         Top             =   7080
         Width           =   1455
      End
      Begin VB.Label serie 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   73
         Top             =   7080
         Width           =   1335
      End
      Begin VB.Label numero 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1680
         TabIndex        =   72
         Top             =   6720
         Width           =   1335
      End
      Begin VB.Label xhora 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   70
         Top             =   7080
         Width           =   1455
      End
      Begin VB.Label xfecha 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   69
         Top             =   6720
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "RUC"
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE"
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   5520
         Width           =   735
      End
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFF80&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "BUSCA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox PLACA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      MaxLength       =   20
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Z"
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
      Index           =   35
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Y"
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
      Index           =   34
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "X"
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
      Index           =   33
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "W"
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
      Index           =   32
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "U"
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
      Index           =   31
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "T"
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
      Index           =   30
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "S"
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
      Index           =   29
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "R"
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
      Index           =   28
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Q"
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
      Index           =   27
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "P"
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
      Index           =   26
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "O"
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
      Index           =   25
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "N"
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
      Index           =   24
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "M"
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
      Index           =   23
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "L"
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
      Index           =   22
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "K"
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
      Index           =   21
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "J"
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
      Index           =   20
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "I"
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
      Index           =   19
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "H"
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
      Index           =   18
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "G"
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
      Index           =   17
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "F"
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
      Index           =   16
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "E"
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
      Index           =   15
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "D"
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
      Index           =   14
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "C"
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
      Index           =   13
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "B"
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
      Index           =   12
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "A"
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
      Index           =   11
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "0"
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
      Index           =   9
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "9"
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
      Index           =   8
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "8"
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
      Index           =   7
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "7"
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
      Index           =   6
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "6"
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
      Index           =   5
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "5"
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
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "4"
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
      Index           =   3
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "3"
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
      Index           =   2
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "2"
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
      Index           =   1
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "1"
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
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label cajero 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   80
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label TURNO 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   79
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label CAJA 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   78
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Image Image9 
      Height          =   855
      Left            =   1440
      Picture         =   "parkes.frx":A4E9
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label tipo 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RUC"
      Height          =   375
      Left            =   120
      TabIndex        =   71
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  'Fixed Single
      Height          =   5655
      Left            =   4560
      Picture         =   "parkes.frx":BC93
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   8490
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PLACA NRO."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      TabIndex        =   37
      Top             =   240
      Width           =   1935
   End
   Begin VB.Menu ffsa 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "PARKEs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        ffsa_Click
        Exit Sub

    End If

    Command8_Click

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)

    Dim sdx As Double

    If KeyAscii <> 13 Then Exit Sub
    If Val(cantidad) = 0 Then
        cantidad = "1"

    End If

    sdx = Val(valor) * Val(cantidad)
    total = Format(sdx, "0.00")
    valor.SetFocus

End Sub

Private Sub Command1_Click(Index As Integer)

    If tipo = "RUC" Then
        RUC = RUC + Command1(Index).Caption

    End If

    If tipo = "NOMBRE" Then
        nombre = nombre + Command1(Index).Caption

    End If

    If tipo = "PLACA" Then
        placa = placa + Command1(Index).Caption

    End If

End Sub

Private Sub Command2_Click()

    Dim sdx   As Double

    Dim found As Integer

    If Val(cantidad) <= 0 Then
        MsgBox "Cantidad Minimo 1", 48, "Aviso"
        Exit Sub

    End If

    serie = "" & mytable11.Fields("serienv")
    sdx = Val("" & mytable11.Fields("numeronv")) + 1
    Numero = "" & sdx
    xtipo = "5"
    acu = "C"
    found = graba_parqueo()
    found = graba_producto()
    found = graba_fpagov()
    'mytable11.Edit
    mytable11.Fields("numeronv") = Numero
    mytable11.Fields("uvueltos") = 0
    mytable11.Fields("uvueltod") = 0
    mytable11.Update
    proceso_impresion11 xtipo, serie, Numero, 1, "1"
    found = borra_parqueo()
    placa = ""
    Command3_Click

End Sub

Private Sub Command3_Click()
    tipo = "PLACA"
    Frame1.Visible = False
    placa.SetFocus

End Sub

Private Sub Command4_Click()

    Dim found As Integer

    RUC = ""
    nombre = ""
    found = busca_parqueo()

    If found = 1 Then
        Frame1.Visible = True
        tipo = "RUC"
        RUC.SetFocus
        Exit Sub

    End If

    placa.SetFocus

End Sub

Private Sub Command5_Click()

    Dim sdx   As Double

    Dim found As Integer

    If Val(cantidad) <= 0 Then
        MsgBox "Cantidad Minimo 1", 48, "Aviso"
        Exit Sub

    End If

    serie = Trim("" & mytable11.Fields("serietb"))
    sdx = Val("" & mytable11.Fields("numerotb")) + 1
    Numero = "" & sdx
    xtipo = "1"
    acu = "C"
    found = graba_parqueo()
    found = graba_producto()
    found = graba_fpagov()
    'mytable11.Edit
    mytable11.Fields("numerotb") = Numero
    mytable11.Fields("uvueltos") = 0
    mytable11.Fields("uvueltod") = 0
    mytable11.Update
    proceso_impresion11 xtipo, serie, Numero, 1, "1"
         
    found = borra_parqueo()
    placa = ""
    Command3_Click

End Sub

Private Sub Command6_Click()

    Dim sdx   As Double

    Dim found As Integer

    If Val(cantidad) <= 0 Then
        MsgBox "Cantidad Minimo 1", 48, "Aviso"
        Exit Sub

    End If

    If Len(RUC) <> 11 Then
        MsgBox dicruc & " No Valido", 48, "Aviso"
        RUC.SetFocus
        Exit Sub

    End If

    If Len(nombre) = 0 Then
        MsgBox "Nombre no Valido", 48, "Aviso"
        nombre.SetFocus
        Exit Sub

    End If

    serie = "" & mytable11.Fields("serietf")
    sdx = Val("" & mytable11.Fields("numerotf")) + 1
    Numero = "" & sdx
    xtipo = "2"
    acu = "D"
    found = graba_parqueo()
    found = graba_producto()
    found = graba_fpagov()
    'mytable11.Edit
    mytable11.Fields("numerotf") = Numero
    mytable11.Fields("uvueltos") = 0
    mytable11.Fields("uvueltod") = 0
    mytable11.Update
    proceso_impresion11 xtipo, serie, Numero, 1, "1"
         
    found = borra_parqueo()
    placa = ""

    Command3_Click

End Sub

Private Sub Command7_Click()
    ffsa_Click

End Sub

Private Sub Command8_Click()
    ejecuta 1

End Sub

Private Sub ffsa_Click()

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Exit Sub

    End If

    PARKEs.Hide
    Unload PARKEs

End Sub

Private Sub Form_Load()
    Frame2.Top = 10: Frame2.Left = 10

    tipo = "PLACA"
    Combo1.Clear
    Combo1.AddItem "Placa"
    Combo1.ListIndex = 0

End Sub

Private Sub Label1_Click()
    placa = ""

End Sub

Function graba_parqueo()

    On Error GoTo cmd4512_err

    Dim mytabley As New ADODB.Recordset

    Dim sdx      As Double

vuelve01:
    Set mytabley = Nothing
    mytabley.Open "select * from factura where local='" & "" & mytable11.Fields("local") & "' and tipo='" & xtipo & "' and serie='" & serie & "' and numero='" & Numero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        sdx = Val("" & Numero) + 1
        Numero = "" & sdx
        mytabley.Close
        GoTo vuelve01
    Else
        mytabley.AddNew
        graba_registro mytabley
        mytabley.Update
        mytabley.Close
        graba_parqueo = 1

    End If

    Exit Function
cmd4512_err:
    MsgBox "Aviso en Graba Parqueo " + error$, 48, "Aviso"
    Exit Function

End Function

Function busca_parqueo()

    Dim mytablex As New ADODB.Recordset

    fecha = ""
    hora = ""
    fechas = ""
    horas = ""
    placas = ""

    mytablex.Open "select * from movplaca where placa='" & placa & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        placas = "" & mytablex.Fields("placa")
        fecha = "" & mytablex.Fields("fecha")
        hora = "" & mytablex.Fields("hora")
        fechas = Format(Now, "dd/mm/yyyy")
        horas = Format(Now, "hh:mm:ss")
        busca_parqueo = 1
        cobrar_fechas

    End If

    mytablex.Close

End Function

Sub cobrar_fechas()

    Dim sdx  As Double

    Dim sdx1 As Double

    Dim horasta

    Dim t0, t1

    sdx1 = 0
    sdx = 0
    t0 = DateValue(fechas)
    t1 = DateValue(fecha)
    sdx1 = (t0 - t1) * 24
    xfecha = "" & sdx
    'MsgBox xfecha
    'horasta = DateDiff("s", fecha & " " & Format(hora, "hh:mm:ss"), fechas & " " & Format(horas, "hh:mm:ss"))
    'MsgBox TimeSerial()
    'MsgBox Format(timserial(0, Val(horasta), 0), "h:Nn")
    t0 = Format(hora, "hh:mm:ss")
    t1 = Format(horas, "hh:mm:ss")
    xhora = Format(TimeValue(t1) - TimeValue(t0), "hh:mm:ss")

    If Val(Mid$(xhora, 1, 2)) > 0 Then
        sdx = sdx + Val(Mid$(xhora, 1, 2))

    End If

    If Val(Mid$(xhora, 4, 2)) >= 1 Then
        sdx = sdx + 1

    End If

    sdx = sdx + sdx1
    cantidad = "" & sdx
    total = sdx * Val(valor)

End Sub

Private Sub Label12_Click()
    tipo = "RUC"
    RUC = ""

End Sub

Private Sub Label13_Click()
    tipo = "NOMBRE"
    nombre = ""

End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub PLACA_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Command4_Click

End Sub

Private Sub PLACA_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_codigo

    End If

End Sub

Private Sub RUC_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(RUC) > 0 Then
        found = busca_codigo()

    End If

    nombre.SetFocus

End Sub

Sub graba_registro(mytablex As ADODB.Recordset)

    Dim I As Integer

    On Error GoTo cmd19012_err

    mytablex.Fields("observa") = ""
    mytablex.Fields("tivap") = 0
    mytablex.Fields("tipo1") = ""
    mytablex.Fields("serie1") = ""
    mytablex.Fields("numero1") = ""
    'mytablex.Fields("observa") = observa
    mytablex.Fields("adetotal") = 0
    mytablex.Fields("acuenta") = 0
    mytablex.Fields("retipo1") = ""
    mytablex.Fields("renumero1") = ""
    mytablex.Fields("renumero2") = ""
    mytablex.Fields("renumero3") = ""
    mytablex.Fields("retotal1") = 0
    mytablex.Fields("retotal2") = 0
    mytablex.Fields("retotal3") = 0
    mytablex.Fields("retotal") = 0
    mytablex.Fields("zona") = ""
    mytablex.Fields("nombre") = Trim("" & nombre)
    mytablex.Fields("estado") = "2"
    mytablex.Fields("tipoclie") = "C"
    mytablex.Fields("tipo") = Trim("" & xtipo)
    mytablex.Fields("serie") = Trim("" & serie)
    mytablex.Fields("numero") = Trim("" & Numero)
    mytablex.Fields("codigo") = Trim("" & RUC)
    mytablex.Fields("partida") = ""
    mytablex.Fields("destino") = ""
    mytablex.Fields("yausado") = "0"
    mytablex.Fields("nro_items") = 1

    mytablex.Fields("fecha") = Format("" & dia, "dd/mm/yyyy")
    mytablex.Fields("fechae") = Format(Now, "dd/mm/yyyy")

    mytablex.Fields("moneda") = "S"
    mytablex.Fields("vendedor") = "" 'Trim("" & Vendedor)
    mytablex.Fields("fpago") = ""
    mytablex.Fields("transporte") = Mid$(Trim("" & placa), 1, 11)
    mytablex.Fields("paridad") = 3#
    mytablex.Fields("dias") = 1
    mytablex.Fields("bodega") = Trim("" & mytable11.Fields("bodega"))
    mytablex.Fields("bodegaf") = ""
    mytablex.Fields("observa") = Mid$(Trim("[" & fecha & " " & hora & " " & fechas & " " & horas & "]"), 1, 60)
    mytablex.Fields("usuario") = Trim("" & gusuario)
    mytablex.Fields("caja") = Trim("" & caja)
    mytablex.Fields("turno") = Trim("" & turno)
    mytablex.Fields("acu") = Trim("" & acu)
    mytablex.Fields("acu1") = ""
    mytablex.Fields("flage") = ""
    mytablex.Fields("telefono") = ""
    mytablex.Fields("hora") = Format(Now, "hh:MM:ss")
    mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("gravado") = 0
    mytablex.Fields("total") = Val("" & total)
    mytablex.Fields("redondeo") = 0
    mytablex.Fields("descuento") = 0
    mytablex.Fields("neto") = 0
    mytablex.Fields("impuesto") = Val("" & total) - Val("" & total) / 1.19
    mytablex.Fields("subtotal") = Val("" & total) / 1.19

    'mytablex.Fields("tipo1") = ""
    'mytablex.Fields("serie1") = ""
    mytablex.Fields("serie2") = ""
    mytablex.Fields("serie3") = ""
    mytablex.Fields("serie4") = ""
    mytablex.Fields("serie5") = ""
    mytablex.Fields("serie6") = ""
    mytablex.Fields("serie7") = ""

    'mytablex.Fields("numero1") = ""
    mytablex.Fields("numero2") = ""
    mytablex.Fields("numero3") = ""
    mytablex.Fields("numero4") = ""
    mytablex.Fields("numero5") = ""
    mytablex.Fields("numero6") = ""
    mytablex.Fields("numero7") = ""
    mytablex.Fields("c1") = 0
    mytablex.Fields("c2") = 0
    mytablex.Fields("c3") = 0
    mytablex.Fields("c4") = 0
    mytablex.Fields("c5") = 0
    mytablex.Fields("c6") = 0
    mytablex.Fields("c7") = 0
    mytablex.Fields("c8") = 0
    mytablex.Fields("c9") = 0
    mytablex.Fields("local") = Trim("" & mytable11.Fields("local"))
    mytablex.Fields("montopagar") = 0
    mytablex.Fields("ruc") = Trim("" & RUC)
    mytablex.Fields("TDOCDELI") = ""
    mytablex.Fields("servicio") = "A"
    'MsgBox "abcd"
    Exit Sub
cmd19012_err:
    MsgBox "Aviso en GRaba Registro " & I & " " & error$, 48, "Aviso"
    Exit Sub

End Sub

Function borra_parqueo()
    cn.Execute ("delete from movplaca where placa='" & placa & "'")

End Function

Function busca_codigo()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from clientes where codigo='" & RUC & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        nombre = "" & mytablex.Fields("nombre")
        busca_codigo = 1

    End If

    mytablex.Close

End Function

Function graba_fpagov()

    Dim mytabley As New ADODB.Recordset

    Dim found    As Integer

    On Error GoTo cdm4411_err

    '---------- validando si es cuenta corriente

    mytabley.Open "select * from fpagov where 1=2", cn, adOpenStatic, adLockOptimistic
    mytabley.AddNew
    grabar_registro_fpagov mytabley
    mytabley.Update
    mytabley.Close
    Exit Function
cdm4411_err:
    MsgBox "Error en graba_fpagov " + error$, 48, "Aviso"
    Exit Function

End Function

Function graba_producto()

    Dim mytabley As New ADODB.Recordset

    Dim found    As Integer

    On Error GoTo cdm44113_err

    mytabley.Open "select * from detalle where 1=2", cn, adOpenDynamic, adLockOptimistic
    mytabley.AddNew
    grabar_producto mytabley
    mytabley.Update
    mytabley.Close
    Exit Function
cdm44113_err:
    MsgBox "Error en graba_detalle " + error$, 48, "Aviso"
    Exit Function

End Function

Sub grabar_registro_fpagov(mytabley As ADODB.Recordset)
    mytabley.Fields("paridad") = 3#
    mytabley.Fields("codigo") = "" & RUC
    mytabley.Fields("nombre") = "" & nombre
    mytabley.Fields("tipo") = "" & xtipo
    mytabley.Fields("serie") = "" & serie
    mytabley.Fields("numero") = "" & Numero
    mytabley.Fields("tipoclie") = "C"
    'mytabley.Fields("codigo") = "" & xruc
    mytabley.Fields("fecha") = Format("" & dia, "dd/mm/yyyy")
    mytabley.Fields("moneda") = "S"
    mytabley.Fields("total") = Val(total)
    mytabley.Fields("caja") = "" & caja
    mytabley.Fields("turno") = "" & turno
    mytabley.Fields("usuario") = "" & cajero
    'mytabley.Fields("vendedor") = "" & cajero
   
    mytabley.Fields("total") = Val(total)
    mytabley.Fields("cambio") = 0
    mytabley.Fields("recibe") = Val(total)
    mytabley.Fields("recibes") = 0 ' Val(total)
    mytabley.Fields("recibed") = 0
    mytabley.Fields("saldos") = 0
    mytabley.Fields("saldod") = 0
    mytabley.Fields("nombre") = nombre
    mytabley.Fields("orden") = ""
    mytabley.Fields("observa") = ""
    mytabley.Fields("dias") = "10"
    mytabley.Fields("fpago") = "1"
    mytabley.Fields("acufp") = "A"
    mytabley.Fields("descripcio") = dicmoneda
    mytabley.Fields("acu") = "" & acu
    mytabley.Fields("local") = "" & mytable11.Fields("local")
    mytabley.Fields("servicio") = "*"
    mytabley.Fields("estado") = "2"
   
End Sub

Function grabar_producto(mytablez As ADODB.Recordset)
    mytablez.Fields("estado") = "2"
    mytablez.Fields("tipo") = "" & xtipo
    mytablez.Fields("serie") = "" & serie
    mytablez.Fields("numero") = Numero
    mytablez.Fields("tipoclie") = "C"
    mytablez.Fields("acu") = acu
    mytablez.Fields("codigo") = "" & RUC
    mytablez.Fields("acu1") = ""
    mytablez.Fields("fecha") = Format(dia, "dd/mm/yyyy")
    mytablez.Fields("moneda") = "S"
    mytablez.Fields("producto") = "PARQUEO"
    mytablez.Fields("FAMILIA") = "PLAYA"
    mytablez.Fields("descripcio") = "SERVICIO PARQUEO"
    mytablez.Fields("unidad") = "UND"
    mytablez.Fields("factor") = 1
    mytablez.Fields("cantidad") = Val(cantidad)
    mytablez.Fields("precio") = Val(valor)
    mytablez.Fields("igv") = 19
    mytablez.Fields("neto") = Val(total) / 1.19
    mytablez.Fields("descuento") = 0
    mytablez.Fields("subtotal") = Val(total) / 1.19
    mytablez.Fields("impuesto") = Val(total) - Val(total) / 1.19
    mytablez.Fields("total") = Val(total)
    mytablez.Fields("estado") = "2"
    mytablez.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
    mytablez.Fields("hora") = Format(Now, "hh:mm:ss")
    mytablez.Fields("vendedor") = "" '& Vendedor
    mytablez.Fields("bodega") = "01"
    mytablez.Fields("bodegaf") = ""
    mytablez.Fields("deslipo") = 0
    mytablez.Fields("flage") = ""
    mytablez.Fields("linea") = ""
    mytablez.Fields("t1") = 0
    mytablez.Fields("t2") = 0
    mytablez.Fields("t3") = 0
    mytablez.Fields("t4") = 0
    mytablez.Fields("t5") = 0
    mytablez.Fields("t6") = 0
    mytablez.Fields("t7") = 0
    mytablez.Fields("t8") = 0
    mytablez.Fields("t9") = 0
    mytablez.Fields("t10") = 0
    mytablez.Fields("t11") = 0
    mytablez.Fields("t12") = 0
    mytablez.Fields("t13") = 0
    mytablez.Fields("t14") = 0
    mytablez.Fields("t15") = 0
    mytablez.Fields("t16") = 0
    mytablez.Fields("l1") = ""
    mytablez.Fields("l2") = ""
    mytablez.Fields("l3") = ""
    mytablez.Fields("l4") = ""
    mytablez.Fields("local") = ""
    mytablez.Fields("proveedorp") = ""
    mytablez.Fields("observa1") = ""
    mytablez.Fields("observa2") = ""
    mytablez.Fields("observa3") = ""
    mytablez.Fields("observa4") = ""
    mytablez.Fields("zona") = ""
    mytablez.Fields("isc") = 0
    mytablez.Fields("tax") = 0
    mytablez.Fields("vtaneta") = 0
    mytablez.Fields("tcosto") = 0
    mytablez.Fields("ganancia") = 0
    mytablez.Fields("comision") = 0
    mytablez.Fields("usuario") = cajero
    mytablez.Fields("cajero") = cajero
    mytablez.Fields("caja") = caja
    mytablez.Fields("turno") = turno
    mytablez.Fields("servicio") = "*"
    mytablez.Fields("comanda") = ""
    mytablez.Fields("mesa") = ""
    mytablez.Fields("salon") = ""
    mytablez.Fields("mesero") = ""
    'mytablez.Fields("codigop") = ""
    mytablez.Fields("local") = "" & mytable11.Fields("local")

End Function

Sub consulta_codigo()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.ListIndex = 0
    Frame2.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "1"
    Command8_Click

End Sub

Sub ejecuta(sw As Integer)

    Dim buf       As String

    Dim rconsulta As New ADODB.Recordset

    If opcion1 = "1" Then
        If Len(buffer) = 0 Then
            buf = "select Placa,Fecha,Hora,Vendedor from movplaca"
        Else
            buf = "select Placa,Fecha,Hora,Vendedor from movplaca where  " & Combo1 & " like '" & buffer & "*'"

        End If

    End If

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = rconsulta

    If rconsulta.RecordCount = 0 Then
        buffer = ""
        rconsulta.Close
        buffer.SetFocus
        Exit Sub

    End If

    If opcion1 = "20" Then
        DBGrid2.columns(0).Width = 2000
        DBGrid2.columns(1).Width = 1300

    End If

    If sw = 1 Then
        DBGrid2.SetFocus

    End If
         
End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo cmd56_err

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = &H2E Then
        If MsgBox("Desea Borra ", 1, "Aviso") <> 1 Then Exit Sub
        Data1.Recordset.Delete
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            placa = DBGrid2.columns(0)
            Frame2.Visible = False
            placa.SetFocus
            PLACA_KeyPress 13

        End If

    End If

    Exit Sub
cmd56_err:
    Exit Sub

End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim buf  As String

    Dim buf2 As String

    Dim sw   As Integer

    If KeyCode <> 13 And KeyCode <> 27 Then

        'MsgBox KeyCode
        If KeyCode >= 48 And KeyCode <= 57 Then
            GoTo sigue9

        End If

        If KeyCode >= 65 And KeyCode <= 90 Then
            GoTo sigue9

        End If

        If KeyCode >= 97 And KeyCode <= 122 Then
            GoTo sigue9

        End If

        If KeyCode = 8 And Chr(KeyCode) = "*" Then
            GoTo sigue9

        End If

        Exit Sub
sigue9:

        If KeyCode = 8 Then
            If Len(buffer) > 0 Then
                buf = Mid$(buffer, 1, Len(buffer) - 1)
                buffer = buf
                KeyCode = 0
            Else
                KeyCode = 0
                Exit Sub

            End If

        End If

        buf = Chr(KeyCode)

        If Chr(KeyCode) = "*" Then
            buf = ""
            buffer = buf

        End If

        If KeyCode <> 13 Then
            buffer = buffer + buf

        End If

        buf = buffer
        ejecuta 0

    End If

End Sub

Sub proceso_impresion11(bxtipo As String, _
                        bxserie As String, _
                        bxnumero As String, _
                        sw As Integer, _
                        ascopia As String)

    Dim found    As Integer

    Dim archivot As String

    On Error GoTo cmd6_err:

    'MsgBox ""
    cerrar_archivo

    If sw = 0 Then   'si es posible
        found = abre_puerto("" & mytable11.Fields("capuerto"), Val("" & mytable11.Fields("catipo")), "" & mytable11.Fields("gavetacola"))

    End If

    'verificamos si es puerto LPT para no hacer formato impresion
    found = control_impresion(bxtipo, 10)

    If found = 10 And sw <> 2 Then
        Exit Sub

    End If

    'MsgBox "proceso impresion"
    factura_formatox Trim("" & mytable11.Fields("local")), "" & bxtipo, "" & bxserie, "" & bxnumero, ascopia, sw
    cerrar_archivo
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = control_impresion(bxtipo, sw)
    'genver.File = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    Exit Sub
cmd6_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$, 48, "Aviso"
    Exit Sub

End Sub

Function control_impresion(bxtipo As String, psw As Integer)

    Dim found      As Integer

    Dim sFile      As String

    Dim mytablex   As New ADODB.Recordset

    Dim sw         As String

    Dim xcolax     As String

    Dim xxpuerto   As String

    Dim oldprinter As String

    On Error GoTo cmd67111_err

    sw = ""
    xcolax = ""
    xxpuerto = "X_"

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM tipo where  tipo='" & bxtipo & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe

        Select Case "" & mytablex.Fields("tipodoc")

            Case "A"
                xxpuerto = "" & mytable11.Fields("puertobm")
                sw = "" & mytable11.Fields("ibm")
                xcolax = "" & mytable11.Fields("cbm")

            Case "B"
                xxpuerto = "" & mytable11.Fields("puertofm")
                sw = "" & mytable11.Fields("ifm")
                xcolax = "" & mytable11.Fields("cfm")

            Case "C"
                xxpuerto = "" & mytable11.Fields("puertotb")
                sw = "" & mytable11.Fields("itb")
                xcolax = "" & mytable11.Fields("ctb")

            Case "D"
                xxpuerto = "" & mytable11.Fields("puertotf")
                sw = "" & mytable11.Fields("itf")
                xcolax = "" & mytable11.Fields("ctf")

            Case "G"
                xxpuerto = "" & mytable11.Fields("puertonv")
                sw = "" & mytable11.Fields("inv")
                xcolax = "" & mytable11.Fields("cnv")

            Case "H"
                xxpuerto = "" & mytable11.Fields("puertope")
                sw = "" & mytable11.Fields("ipe")
                xcolax = "" & mytable11.Fields("cpe")

            Case "I"  'pedidos
       
                xxpuerto = "" & mytable11.Fields("puertope")
                sw = "" & mytable11.Fields("ipe")
                xcolax = "" & mytable11.Fields("cpro")
       
            Case "T"
                xxpuerto = "" & mytable11.Fields("puertoot")
                sw = "" & mytable11.Fields("iot")
                xcolax = "" & mytable11.Fields("cpro")

            Case "1"
                xxpuerto = "" & mytable11.Fields("puertoexo")
                sw = "" & mytable11.Fields("iexo")
                xcolax = "" & mytable11.Fields("cexo")

        End Select

    End If

    mytablex.Close

    If psw = 10 Then  'solo es para ver si es LPT
        control_impresion = 11

        If xxpuerto = "LPT" Then
            control_impresion = 10

        End If

        Exit Function

    End If

    'found = abre_puerto("" & mytable11.Fields("capuerto"), Val("" & mytable11.Fields("catipo")))
    'ahora validamos los parametros de impresion

    If psw = 2 Then  'si  es orden de despacho
   
        If "" & mytable11.Fields("odcola") = "S" Then
      
            oldprinter = Printer.DeviceName
            selecciona_impresoras ("" & mytable11.Fields("odpuerto"))
            sFile = globaldir & "\temporal\" & gusuario & ".txt"
            found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))
            selecciona_impresoras (oldprinter)

        End If

        If "" & mytable11.Fields("odcola") <> "S" Then
            'MsgBox "" & mytable11.Fields("odpuerto")
            found = star_sp342("" & mytable11.Fields("odpuerto"), 0)
            found = corte_papel("" & mytable11.Fields("odpuerto"), Val("" & mytable11.Fields("catipo")))

        End If

        control_impresion = found
        Exit Function

    End If

    If sw = "S" Then
        If MsgBox("Desea Imprimir", 1 + 256, "Aviso") <> 1 Then
            control_impresion = 1
            Exit Function

        End If

    End If

    If xcolax = "S" Then
        oldprinter = Printer.DeviceName
        selecciona_impresoras (xxpuerto)
        sFile = globaldir & "\temporal\" & gusuario & ".txt"
        found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))
        selecciona_impresoras (oldprinter)

    End If

    If xcolax <> "S" Then
        found = star_sp342(xxpuerto, 0)
        found = corte_papel(xxpuerto, Val("" & mytable11.Fields("catipo")))

    End If

    control_impresion = found
    Exit Function
cmd67111_err:
    MsgBox "Aviso en control impresion " + error$, 48, "Aviso"
    Exit Function

End Function

Sub factura_formatox(bxlocal As String, _
                     bxtipo As String, _
                     bxserie As String, _
                     bxnumero As String, _
                     ascopia As String, _
                     psw As Integer)

    Dim vacu            As String

    Dim mytablex        As New ADODB.Recordset

    Dim mytabley        As New ADODB.Recordset

    Dim mytablez        As New ADODB.Recordset

    Dim found           As Integer

    Dim nro_lineas      As Integer

    Dim contando        As Integer

    Dim faltante        As Integer

    Dim I               As Integer

    Dim archivo_formato As String

    On Error GoTo cmd450009_err

    vacu = ""
    'MsgBox "QU"
       
    nro_lineas = busca_tipo_lineas(bxtipo)
    'MsgBox ""
    'If nro_lineas <= 0 Then
    '   nro_lineas = 10
    'End If
    'MsgBox ""
    contando = 0
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = borra_nombre("" & FileName)
       
    If psw = 2 Then 'si es de orden
        archivo_formato = "orden"
    Else
        'MsgBox bxtipo
        archivo_formato = busca_archivo_formato(bxtipo)

        If Len(archivo_formato) = 0 Then
            MsgBox "No existe archivo formato ", 48, "Aviso"
            'MsgBox ""
            Exit Sub

        End If

    End If

    'cabeza
    'proceso_formatos(archivo_formato , mydbx , mytablex , ubicacioni , ubicacionf , basedatos , indice , tipo , numero , ascopia , contando )
    mytablex.Open "SELECT * FROM " & gocabeza & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then 'si existe
        mytablex.Close
        Exit Sub

    End If

    'MsgBox ""
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytablex, "{", "}", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytablex, "{", "}", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    vacu = "" & mytablex.Fields("acu")
    'MsgBox ""
    '
    'detalle
    flag_contando = 0
    mytabley.Open "SELECT * FROM " & godetalle & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytabley.RecordCount > 0 Then 'si existe
        Do

            If mytabley.EOF Then Exit Do
            If "" & mytabley.Fields("dua") <> "R" Then
                flag_contando = contando + 1
                'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                'found = proceso_formatos(archivo_formato, mytabley, "/", "\", godetalle, "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
                found = proceso_formatos(archivo_formato, mytabley, "/", "\", godetalle, "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
                'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                   
                contando = contando + 1

            End If
          
            mytabley.MoveNext
        Loop

    End If

    'mytabley.Close
    '
    If nro_lineas > 0 Then

        'If bxtipo <> "1" And bxtipo <> "2" And bxtipo <> "5" And bxtipo <> "7" Then
        If contando < nro_lineas Then

            For I = contando To nro_lineas
                Open FileName For Append As #1
                found = formateaa("", 1, 2, 0)
                Close #1
            Next I

        End If

    End If

    '----- SUBTOTAL
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytablex, "$", "?", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytablex, "$", "?", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
           
    mytablez.Open "SELECT * FROM " & gofpago & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablez.RecordCount > 0 Then 'si existe
        Do

            If mytablez.EOF Then Exit Do
            'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
            'found = proceso_formatos(archivo_formato, mytablez, "<", ">", gofpago, "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
            found = proceso_formatos(archivo_formato, mytablez, "<", ">", gofpago, "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
            'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                    
            mytablez.MoveNext
        Loop

    End If

    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytablex, "^", "&", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytablex, "^", "&", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    mytablex.Close
    mytabley.Close
    mytablez.Close
    Exit Sub
cmd450009_err:
    MsgBox "Aviso Proceso en Compilando " & error$, 24, "Aviso"
    'mytablex.Close
    Exit Sub

End Sub

Function busca_archivo_formato(bxtipo As String) As String

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    'MsgBox bxtipo
    mytablex.Open "SELECT * FROM tipo where  tipo='" & bxtipo & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe

        Select Case "" & mytablex.Fields("tipodoc")

            Case "Z" 'si es traslado
                busca_archivo_formato = "" & mytablex.Fields("archivo")

            Case "A"
                busca_archivo_formato = "" & mytable11.Fields("archivobm")

            Case "B"
                busca_archivo_formato = "" & mytable11.Fields("archivofm")

            Case "C"
                busca_archivo_formato = "" & mytable11.Fields("archivotb")

            Case "1"
                busca_archivo_formato = "" & mytable11.Fields("archivoexo")

            Case "D"
                busca_archivo_formato = "" & mytable11.Fields("archivotf")

            Case "G"
                busca_archivo_formato = "" & mytable11.Fields("archivonv")

            Case "H"
                busca_archivo_formato = "" & mytable11.Fields("archivope")

            Case "T"
                busca_archivo_formato = "" & mytable11.Fields("archivoot")

            Case "I"
                busca_archivo_formato = "" & mytable11.Fields("archivope")

                'MsgBox ""
        End Select

        'MsgBox ""
    End If

    mytablex.Close
 
End Function

Function busca_parame1(buf As String, sw As Integer) As String

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM parame where  codigo='01'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe

        If sw = 2 Then
      
        End If

        If sw = 0 Then
            sdx = Val("" & mytablex.Fields("clientes")) + 1
            busca_parame1 = "" & sdx

        End If

        If sw = 1 Then
            'mytablex.Edit
            mytablex.Fields("clientes") = buf
            mytablex.Update

        End If

    End If

    mytablex.Close

End Function

Function busca_tipo_lineas(buf As String)

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM tipo  where  tipo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        busca_tipo_lineas = Val("" & mytablex.Fields("nrolineas"))

        'MsgBox ""
    End If

    mytablex.Close

End Function

Private Sub total_KeyPress(KeyAscii As Integer)

    Dim sdx As Double

    If KeyAscii <> 13 Then Exit Sub
    If Val(total) > 0 And Val(cantidad) > 0 Then
        sdx = Val(total) / Val(cantidad)
        valor = Format(sdx, "0.00")

    End If

    RUC.SetFocus

End Sub

Private Sub valor_KeyPress(KeyAscii As Integer)

    Dim sdx As Double

    If KeyAscii <> 13 Then Exit Sub
    sdx = Val(valor) * Val(cantidad)
    total = Format(sdx, "0.00")
    total.SetFocus

End Sub

