VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tncr1 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta Precios"
   ClientHeight    =   10530
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   14715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Consulta"
      Height          =   5175
      Left            =   7920
      TabIndex        =   54
      Top             =   9240
      Visible         =   0   'False
      Width           =   13455
      Begin MSDataGridLib.DataGrid dbgrid15 
         Height          =   4815
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   8493
         _Version        =   393216
         HeadLines       =   2
         RowHeight       =   19
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
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
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
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   48
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   40
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   8880
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Height          =   1935
      Left            =   11160
      TabIndex        =   51
      Top             =   9240
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Error en Busqueda, Intente de Nuevo"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   5520
         TabIndex        =   52
         Top             =   1440
         Width           =   5295
      End
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H000000FF&
      Caption         =   "LIMPIAR"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   39
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   9720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   38
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   37
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   36
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   35
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   34
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   33
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   32
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   31
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   30
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   29
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ESPACIADOR"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   28
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   9720
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ENTER"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   27
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   9720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   26
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   25
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   24
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   23
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   22
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   21
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ñ"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   20
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   19
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   18
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   17
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   16
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   15
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   14
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   13
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   12
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   11
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   10200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   10
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   9
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   7
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   6
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton teclas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Index           =   0
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000003&
      Caption         =   "LIMPIAR"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox producto 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      MaxLength       =   15
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label precio 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   4
      Left            =   5400
      TabIndex        =   61
      Top             =   6120
      Width           =   6135
   End
   Begin VB.Label precio 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   3
      Left            =   5400
      TabIndex        =   60
      Top             =   5040
      Width           =   6135
   End
   Begin VB.Label unidad 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   4
      Left            =   1080
      TabIndex        =   59
      Top             =   6120
      Width           =   4335
   End
   Begin VB.Label unidad 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   3
      Left            =   1080
      TabIndex        =   58
      Top             =   5040
      Width           =   4335
   End
   Begin VB.Label unidad 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   2
      Left            =   1080
      TabIndex        =   57
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Label precio 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   2
      Left            =   5400
      TabIndex        =   56
      Top             =   3960
      Width           =   6135
   End
   Begin VB.Label precio 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   1
      Left            =   5400
      TabIndex        =   9
      Top             =   3000
      Width           =   6135
   End
   Begin VB.Label precio 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   0
      Left            =   5400
      TabIndex        =   8
      Top             =   1920
      Width           =   6135
   End
   Begin VB.Label unidad 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label unidad 
      Alignment       =   2  'Center
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Index           =   0
      Left            =   1080
      TabIndex        =   6
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label moneda 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9360
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label local1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   13320
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   12840
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label descripcio 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   12255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ingrese Código Barras"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Menu b1212 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tncr1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'' 03/07/2018 Mejora Visor de Precios
' REEMPLAZAR TODO EL FORMULARIO CON DISEÑO
'' 03/07/2018 Mejora Visor de Precios

Private Type campo_precio

    unidad As String
    factor As String
    precio As String

End Type

Dim campo_precios(12) As campo_precio

Private Sub b1212_Click()
    tncr1.Hide
    Unload tncr1

End Sub

Function busca_equiva(buf As String) As Integer

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM productb where barras='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = "" & mytablex.Fields("producto")
        busca_equiva = 1
        mytablex.Close
        Exit Function

    End If

    mytablex.Close
   
    mytablex.Open "SELECT * FROM producto where barras='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = "" & mytablex.Fields("producto")
        busca_equiva = 1

    End If

    mytablex.Close

End Function

Private Sub Command1_Click()
    Frame2.Visible = False
    producto = ""
    descripcio = ""
    precio(0).Caption = ""
    precio(1).Caption = ""
    precio(2).Caption = ""
   
    precio(3).Caption = ""
    precio(4).Caption = ""
   
    unidad(0).Caption = ""
    unidad(1).Caption = ""
    unidad(2).Caption = ""
    unidad(3).Caption = ""
    unidad(4).Caption = ""
   
    producto.SetFocus

    If producto.Enabled = True Then

        'producto.SetFocus
    End If

End Sub

Private Sub dbgrid15_DblClick()

    On Error GoTo cmd8978_err

    producto = Trim("" & dbgrid15.columns(1))
    Frame2.Visible = False
    producto_KeyPress 13
    Exit Sub
cmd8978_err:
    producto = ""
    Exit Sub

End Sub

Private Sub Form_Load()

    If Len(Trim(local1)) = 0 Then
        local1 = "01"

    End If

End Sub

Private Sub producto_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub

    Dim vr

    If Len(producto) = 0 Then
        Exit Sub

    End If

    found = busca_producto("" & producto)

    If found = 0 Then
        consulta_producto
   
        'Frame1.Visible = True
        'vr = DoEvents()
        'Sleep (3000)
        'vr = DoEvents()
        'Command1_Click
        'Frame1.Visible = False
        'producto.SetFocus
        Exit Sub

    End If

    vr = DoEvents()
    Sleep (5000)
    vr = DoEvents()
    Command1_Click

End Sub

Sub consulta_producto()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT Descripcio,Producto,Unidad,factor FROM producto where descripcio like '%" & Trim(producto) & "%'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Set dbgrid15.DataSource = mytablex
    dbgrid15.columns(0).Width = 4000
    dbgrid15.columns(1).Width = 2000
    Frame2.Visible = True
    dbgrid15.SetFocus
   
End Sub

Function busca_producto(buf As String)

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    I = 0
    found = 0

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        found = busca_equiva(buf) 'busca en la table codigo barras

        If found = 0 Then
            Exit Function

        End If

        mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Function

        End If

    End If

    If "" & mytablex.Fields("estado") = "N" Then  'si no esta activo
        MsgBox "Producto no activo ", 48, "Aviso"
        mytablex.Close
        Exit Function

    End If
   
    If mytabley.State = 1 Then mytabley.Close
    mytabley.Open "SELECT * FROM precios where producto='" & buf & "' and local='" & local1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        mytabley.Close
        mytablex.Close
        Exit Function

    End If

    busca_producto = 1
    descripcio = "" & mytablex.Fields("descripcio")
      
    precio(0).Caption = ""
    precio(1).Caption = ""
    precio(2).Caption = ""
      
    precio(3).Caption = ""
    precio(4).Caption = ""
      
    unidad(0).Caption = ""
    unidad(1).Caption = ""
    unidad(2).Caption = ""
    unidad(3).Caption = ""
    unidad(4).Caption = ""
      
    If Len("" & mytabley.Fields("unidad1")) > 0 Then
        unidad(0).Caption = "" & mytabley.Fields("unidad1")

    End If

    If Len("" & mytabley.Fields("unidad2")) > 0 Then
        unidad(1).Caption = "" & mytabley.Fields("unidad2")

    End If
      
    If Val("" & mytabley.Fields("pventa1")) > 0 Then
        precio(0).Caption = "S/. " & Format(Val("" & mytabley.Fields("pventa1")), "0.00")

    End If

    If Val("" & mytabley.Fields("pventa2")) > 0 Then
        precio(1).Caption = "S/. " & Format(Val("" & mytabley.Fields("pventa2")), "0.00")

    End If
      
    If Len("" & mytabley.Fields("unidad3")) > 0 Then
        unidad(2).Caption = "" & mytabley.Fields("unidad3")

    End If

    If Val("" & mytabley.Fields("pventa3")) > 0 Then
        precio(2).Caption = "S/. " & Format(Val("" & mytabley.Fields("pventa3")), "0.00")

    End If
      
    If Len("" & mytabley.Fields("unidad4")) > 0 Then
        unidad(3).Caption = "" & mytabley.Fields("unidad4")

    End If

    If Val("" & mytabley.Fields("pventa4")) > 0 Then
        precio(3).Caption = "S/. " & Format(Val("" & mytabley.Fields("pventa4")), "0.00")

    End If
      
    If Len("" & mytabley.Fields("unidad5")) > 0 Then
        unidad(4).Caption = "" & mytabley.Fields("unidad5")

    End If

    If Val("" & mytabley.Fields("pventa5")) > 0 Then
        precio(4).Caption = "S/. " & Format(Val("" & mytabley.Fields("pventa5")), "0.00")

    End If

    'Command1_Click

End Function

Public Sub SlowDown(MilliSeconds As Long)

    Dim lngTickStore As Long

    lngTickStore = GetTickCount()

    Do While lngTickStore + MilliSeconds > GetTickCount()
        DoEvents
    Loop

End Sub

Private Sub teclas_Click(Index As Integer)

    On Error GoTo cmd6732_err

    If Index = 39 Then
        Command1_Click
        Exit Sub

    End If

    If Index = 27 Then
        If Frame2.Visible = True Then
            producto = Trim("" & dbgrid15.columns(1))

        End If

        Frame2.Visible = False
          
        producto_KeyPress 13
          
        Exit Sub

    End If

    If Index = 28 Then
        producto = producto & " "
        Exit Sub

    End If

    If Index = 40 Then
        If Len(producto) = 0 Then Exit Sub
        producto = Mid$(producto, 1, Len(producto) - 1)
               
        Exit Sub

    End If

    producto = producto & teclas(Index).Caption
    producto_KeyPress 13
    Exit Sub
cmd6732_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

