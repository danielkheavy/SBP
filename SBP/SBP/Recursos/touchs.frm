VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form touchs 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema Orion Tpv"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   14775
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command19 
      BackColor       =   &H0080FF80&
      Caption         =   "Copia"
      Height          =   855
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Anula"
      Height          =   855
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "Descon-  gela"
      Height          =   855
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C0FFFF&
      Caption         =   "  Comen- Tarios"
      Height          =   615
      Left            =   11400
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   4080
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3480
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   13200
      Top             =   0
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H0080FF80&
      Caption         =   "Ingreso"
      Height          =   855
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Egreso"
      Height          =   855
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H000000FF&
      Caption         =   "Congela"
      Height          =   855
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00FF8080&
      Caption         =   "Auto Servicio"
      Height          =   855
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cobra Mesa"
      Height          =   855
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0000C0C0&
      Caption         =   "Domicilios"
      Height          =   855
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Modifica"
      Height          =   615
      Left            =   11400
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Borra Linea"
      Height          =   615
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000C0&
      Caption         =   "Pedido Comanda"
      Height          =   855
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Borra Pedido"
      Height          =   615
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "touchs.frx":0000
      Height          =   3615
      Left            =   3120
      OleObjectBlob   =   "touchs.frx":0014
      TabIndex        =   3
      Top             =   480
      Width           =   8295
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   5055
      Left            =   3120
      ScaleHeight     =   4995
      ScaleWidth      =   11595
      TabIndex        =   2
      Top             =   4680
      Width           =   11655
      Begin VB.CommandButton Command14 
         BackColor       =   &H00C0E0FF&
         Height          =   735
         Left            =   10440
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Control Personal"
         Height          =   735
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00C0E0FF&
         Caption         =   " Abrir     Gaveta"
         Height          =   735
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   23
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   22
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   21
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   20
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   19
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   18
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   17
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   16
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   15
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   14
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   13
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   12
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   11
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   10
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   9
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   8
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   7
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   6
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   5
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   4
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   3
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   2
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   1
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFF00&
         Height          =   975
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T/C."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   76
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label paridadfp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5040
         TabIndex        =   75
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label paridad 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4320
         TabIndex        =   74
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UltimoVuelto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   73
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label uvueltos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   72
         Top             =   4560
         Width           =   1305
      End
      Begin VB.Label uvueltod 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   71
         Top             =   4560
         Width           =   1305
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "         PRODUCTOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   5040
         TabIndex        =   6
         Top             =   0
         Width           =   1455
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   6480
         Picture         =   "touchs.frx":108B
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1200
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   3840
         Picture         =   "touchs.frx":3031
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1200
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   1320
         TabIndex        =   5
         Top             =   0
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   735
         Left            =   2280
         Picture         =   "touchs.frx":4C03
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   705
         Left            =   120
         Picture         =   "touchs.frx":6685
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000FF00&
      Height          =   5055
      Left            =   0
      ScaleHeight     =   4995
      ScaleWidth      =   2955
      TabIndex        =   1
      Top             =   4680
      Width           =   3015
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   17
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   16
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   15
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   14
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   13
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   12
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   11
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   10
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   9
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   8
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   7
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   6
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   5
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   4
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   3
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   2
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   1
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton command8 
         BackColor       =   &H0000FFFF&
         Height          =   615
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   720
         Width           =   975
      End
      Begin VB.Image Image6 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   0
         Picture         =   "touchs.frx":80FB
         Stretch         =   -1  'True
         Top             =   0
         Width           =   960
      End
      Begin VB.Image Image5 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   1920
         Picture         =   "touchs.frx":9CCD
         Stretch         =   -1  'True
         Top             =   0
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "    FAMILIA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   960
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   3615
      Left            =   0
      OleObjectBlob   =   "touchs.frx":BC73
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   11400
      Picture         =   "touchs.frx":C646
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1080
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   11400
      Picture         =   "touchs.frx":C9A3
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label flag_carga 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   1680
      TabIndex        =   70
      Top             =   9720
      Width           =   45
   End
   Begin VB.Label horasis 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7920
      TabIndex        =   69
      Top             =   0
      Width           =   975
   End
   Begin VB.Label fechasis 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6960
      TabIndex        =   68
      Top             =   0
      Width           =   975
   End
   Begin VB.Label cajero 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5040
      TabIndex        =   67
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label turno 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6600
      TabIndex        =   66
      Top             =   0
      Width           =   375
   End
   Begin VB.Label caja 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6120
      TabIndex        =   65
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   61
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8160
      TabIndex        =   60
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5055
   End
   Begin VB.Menu dlo232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "touchs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mfamcod(15000) As String
Dim mfampag As Integer
Dim mfamtop As Integer

Dim mprodcod(15000) As String
Dim wprodcod(15000) As String
Dim wwprodcod(30) As String
Dim mprodpag As Integer
Dim mprodtop As Integer
Dim exisdev As Integer


Private Sub Command7_Click(Index As Integer)
Dim found As Integer
Dim buf As String
If Val(Label3) = 0 Then
   MsgBox "Cantidad=0", 48, "Aviso"
   Exit Sub
End If
If Len(wwprodcod(Index)) = 0 Then
   Exit Sub
End If
      buf = "" & wwprodcod(Index)
      found = busca_producto(buf, 0, Val(Label3))
      If found = 0 Then
         MsgBox "No existe Producto Buscado ", 48, "Aviso"
         Exit Sub
      End If
      found = ir_final1()
         
End Sub

Private Sub Command8_Click(Index As Integer)
menu_carga_producto Command8(Index).Caption
menu_producto "INI"

End Sub

Private Sub Form_Activate()
Dim found As Integer
On Error GoTo cmd3411_err
If flag_carga <> "S" Then
   found = busca_paridad()
   sql_detalle
   flag_carga = "S"
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus
End If
uvueltos = "S/.:" & Format(Val("" & mytable11.Fields("uvueltos")), "0.00")
uvueltod = "US$:" & Format(Val("" & mytable11.Fields("uvueltod")), "0.00")
If "" & mytable11.Fields("terminal") = "T" Then
End If
Exit Sub
cmd3411_err:
Exit Sub

End Sub

Private Sub Form_Load()
carga_inicial
End Sub
Sub carga_inicial()
carga_familia

End Sub
Sub carga_producto()

End Sub
Sub carga_familia()
Dim mytablex As Table
Dim i As Integer
For i = 0 To 14999
    mfamcod(i) = ""
Next i

i = -1
Set mytablex = mydbxglo.OpenTable("familia")
Do
If mytablex.EOF Then Exit Do
i = i + 1
mfamcod(i) = "" & mytablex.Fields("familia")
mytablex.MoveNext
Loop
mfamtop = i
mytablex.Close
mfampag = 0
menu_familia "INI"

End Sub
Sub menu_carga_producto(buf As String)
Dim mytablex As Table
Dim i As Integer
For i = 0 To 29
   wwprodcod(i) = ""
Next i
For i = 0 To 14999
    mprodcod(i) = ""
    wprodcod(i) = ""
Next i

i = -1
Set mytablex = mydbxglo.OpenTable("producto")
mytablex.Index = "familia"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("familia") = buf Then
   i = i + 1
   mprodcod(i) = "" & mytablex.Fields("descripcio")
   wprodcod(i) = "" & mytablex.Fields("producto")
   Else: Exit Do
End If
mytablex.MoveNext
Loop
End If
mytablex.Close
mprodtop = i
mprodpag = 0

End Sub
Sub menu_producto(buf As String)
Dim i As Integer
Dim j As Integer

Select Case buf
       Case "INI"
            mprodpag = 0
       Case "SIG"
            mprodpag = mprodpag + 23
            If mprodpag > 102 Then
               mprodpag = 0
            End If
       Case "ANT"
            mprodpag = mprodpag - 23
            If mprodpag < 0 Then
               mprodpag = 0
            End If
End Select
j = -1
For i = mprodpag To 23 + mprodpag
    j = j + 1
    Command7(j).Caption = mprodcod(i)
    wwprodcod(j) = wprodcod(i)
Next i


End Sub
Sub menu_familia(buf As String)
Dim i As Integer
Dim j As Integer
Select Case buf
       Case "INI"
            mfampag = 0
       Case "SIG"
            mfampag = mfampag + 17
            If mfampag > 102 Then
               mfampag = 0
            End If
       Case "ANT"
            mfampag = mfampag - 17
            If mfampag < 0 Then
               mfampag = 0
            End If
End Select
j = -1
For i = mfampag To 17 + mfampag
    j = j + 1
    Command8(j).Caption = mfamcod(i)
Next i

End Sub

Private Sub Image1_Click()
Dim sdx As Double
sdx = Val(Label3) - 1
Label3 = "" & sdx

End Sub

Private Sub image2_Click()
Dim sdx As Double
sdx = Val(Label3) + 1
Label3 = "" & sdx
End Sub

Private Sub image3_Click()
menu_producto "ANT"
End Sub

Private Sub Image4_Click()
menu_producto "SIG"
End Sub

Private Sub image5_Click()
menu_familia "SIG"
End Sub

Private Sub Image6_Click()
menu_familia "ANT"
End Sub

Private Sub Image8_Click()
If Data2.Recordset.RecordCount = 0 Then
  Data2.Recordset.MoveNext
Else
  Data2.Recordset.MoveNext
  If Data2.Recordset.EOF Then
    Data2.Recordset.MoveLast
  End If
End If
End Sub

Private Sub Label3_Click()
Label3 = "1"
End Sub

Private Sub Label5_Click()
menu_familia "INI"
End Sub
Function busca_equiva(buf As String) As Integer
Dim mytablex As Table
Set mytablex = mydbxglo.OpenTable("productb")
mytablex.Index = "productb"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   buf = "" & mytablex.Fields("producto")
   busca_equiva = 1
End If
mytablex.Close
 
End Function

Function busca_producto(buf As String, sw As Integer, canti As String)
Dim mytablex As Table
Dim mytabley As Table
Dim buf1 As String
Dim i As Integer
Dim ssw As Integer
Dim found As Integer

Set mytablex = mydbxglo.OpenTable("producto")
Set mytabley = mydbxglo.OpenTable("precios")
mytabley.Index = "tprecios"
'------------------------------------
found = busca_equiva(buf)
If found = 1 Then
   
End If
i = 0
ssw = 0
Set mytablex = mydbxglo.OpenTable("producto")
buf1 = buf
mytablex.Index = "producto"
a11d:
If i > 4 Then
   mytablex.Close
   Exit Function
End If
mytablex.Seek "=", buf1
If mytablex.NoMatch And ssw = 0 Then
   mytablex.Index = "barras"
   ssw = 1
   i = i + 1
   GoTo a11d
End If
'------------------------------------
If Not mytablex.NoMatch Then
   mytabley.Seek "=", "" & mytablex.Fields("producto"), "" & mytable11.Fields("listap")
   If mytabley.NoMatch Then  'si no existe la tabla creada
      mytabley.Close
      mytablex.Close
      Exit Function
   End If
   If Val("" & mytabley.Fields("pventa1")) <= 0 Then
      If "" & mytable11.Fields("noprecio") = "S" Then
         MsgBox "" & mytablex.Fields("descripcio") & "  Precio <=0 No Permitido ", 48, "Aviso"
         mytablex.Close
         busca_producto = 2
         Exit Function
      End If
      If "" & mytablex.Fields("oferta") <> "S" Then
         'MsgBox "" & mytablex.Fields("descripcio") & "  Precio <=0", 48, "Aviso"
         'mytablex.Close
         'busca_producto = 2
         'Exit Function
      End If
   End If
   'End If
   'canti = ""
   buf = ""
   '----------- verfica a forzar la balanza
   If Val(canti) <= 0 Then
   If "" & mytable11.Fields("actbala") = "S" Then
     If "" & mytablex.Fields("peso") = "S" Then
ajk91:
        buf = puerto_balanza1()
        If Val(buf) = 0 Then
           If MsgBox("Balanza No leido,Continua Leyendo? ", 1, "Aviso") = 1 Then
              GoTo ajk91
              '------
              Else
              MsgBox "No leido ", 48, "Aviso"
              busca_producto = 2
              mytablex.Close
             Exit Function
           End If
        End If
     End If
   End If
   canti = Format(Val(buf), "0.00")
   'canti = buf
   End If
   If Val(canti) <= 0 Then
      canti = "1"
   End If
   busca_producto = 1
   '---------------------------------------
   If sw = 0 Or sw = 2 Then
      graba_temporald mytablex, sw, canti, mytabley
      
   End If
End If
mytablex.Close
mytabley.Close
End Function
Sub calcula_igv(sw As Integer)
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double
Dim tdscto As Double
Dim tdscto1 As Double
Dim found As Integer
Dim xtivap As Double
On Error GoTo cmd4567_err

xtivap = Val(Data2.Recordset.Fields("total")) * Val(Data2.Recordset.Fields("ivap")) / 100
Data2.Recordset.Fields("tivap") = xtivap
tdscto = Val(Data2.Recordset.Fields("total")) * Val(Data2.Recordset.Fields("deslipo")) / 100       'calcular descuento
Data2.Recordset.Fields("descuento") = tdscto  'total descuento
Data2.Recordset.Fields("total") = Val(Data2.Recordset.Fields("total")) - Val(Data2.Recordset.Fields("descuento")) 'cobrar
Data2.Recordset.Fields("subtotal") = Val(Data2.Recordset.Fields("total")) 'subtotal
Data2.Recordset.Fields("impuesto") = 0
Data2.Recordset.Fields("neto") = Val(Data2.Recordset.Fields("subtotal")) + Val(Data2.Recordset.Fields("descuento"))
If Val(Data2.Recordset.Fields("total")) > 0 And Val(Data2.Recordset.Fields("igv")) > 0 Then
   sdx2 = 1 + Val(Data2.Recordset.Fields("igv")) / 100
   sdx1 = Val(Data2.Recordset.Fields("total")) / sdx2
   Data2.Recordset.Fields("subtotal") = sdx1  'subtotal
   sdx = Val(Data2.Recordset.Fields("total")) - Val(Data2.Recordset.Fields("subtotal"))
   Data2.Recordset.Fields("impuesto") = sdx  'impuesto
   Data2.Recordset.Fields("descuento") = tdscto
   Data2.Recordset.Fields("neto") = Val(Data2.Recordset.Fields("subtotal")) + Val(Data2.Recordset.Fields("descuento"))
End If
Exit Sub
cmd4567_err:
MsgBox "Error en Calcula Igv " + error$, 48, "Aviso"
Exit Sub
End Sub
Sub graba_temporald(mytablex As Table, sw As Integer, canti As String, mytabley As Table)
Dim found As Integer
Dim xxca As String
Dim sdx As Double
Dim xpreciox As Double
On Error GoTo cmd3461_err
Data2.Recordset.AddNew
xxca = "1"
If Val(canti) > 0 Then
   xxca = "" & canti
End If
xpreciox = 0
If Val(paridad) <= 0 Then
   paridad = "1"
End If
If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
   xpreciox = Val("" & mytabley.Fields("pventa1"))
   If "" & mytablex.Fields("monedav") = "D" Then
      xpreciox = Val("" & mytabley.Fields("pventa1")) * Val(paridad)
   End If
End If
If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
   xpreciox = Val("" & mytabley.Fields("pventa1"))
   If "" & mytablex.Fields("monedav") = "S" Then
      xpreciox = Val("" & mytabley.Fields("pventa1")) / Val(paridad)
   End If
End If
Data2.Recordset.Fields("producto") = "" & mytablex.Fields("producto")
Data2.Recordset.Fields("proveedorp") = "" & mytablex.Fields("proveedor1")
Data2.Recordset.Fields("tipo") = ""
Data2.Recordset.Fields("serie") = ""
Data2.Recordset.Fields("numero") = ""
Data2.Recordset.Fields("vendedor") = ""
Data2.Recordset.Fields("descripcio") = "" & mytablex.Fields("descripcio")
Data2.Recordset.Fields("cantidad") = Val(Format(Val(xxca), "0.000"))
Data2.Recordset.Fields("descuento") = Val("" & mytablex.Fields("isc")) 'ojo verificar
Data2.Recordset.Fields("unidad") = "" & mytabley.Fields("unidad1")
Data2.Recordset.Fields("factor") = Val("" & mytabley.Fields("factor1"))
Data2.Recordset.Fields("precio") = xpreciox
Data2.Recordset.Fields("total") = xpreciox
Data2.Recordset.Fields("subtotal") = xpreciox
'DBGrid2.Columns(13) = Val("" & mytablex.Fields("tax"))
Data2.Recordset.Fields("unidad") = "" & mytabley.Fields("unidad1")
Data2.Recordset.Fields("factor") = Val("" & mytabley.Fields("factor1"))
Data2.Recordset.Fields("precio") = xpreciox
Data2.Recordset.Fields("total") = xpreciox
Data2.Recordset.Fields("subtotal") = xpreciox
Data2.Recordset.Fields("deslipo") = 0
Data2.Recordset.Fields("tax") = 0
Data2.Recordset.Fields("isc") = 0
Data2.Recordset.Fields("impuesto") = 0
Data2.Recordset.Fields("igv") = Val("" & mytablex.Fields("igv"))
Data2.Recordset.Fields("linea") = "" & mytablex.Fields("linea")
Data2.Recordset.Fields("descuento") = 0
Data2.Recordset.Fields("neto") = 0
Data2.Recordset.Fields("ccosto") = "" & mytabley.Fields("ccosto")  'ojo si no es por local
Data2.Recordset.Fields("familia") = "" & mytablex.Fields("Familia")
Data2.Recordset.Fields("subfamilia") = "" & mytablex.Fields("subFamilia")
Data2.Recordset.Fields("marca") = "" & mytablex.Fields("marca")
Data2.Recordset.Fields("total") = Val(Data2.Recordset.Fields("cantidad")) * Val(Data2.Recordset.Fields("precio"))
Data2.Recordset.Fields("ivap") = Val("" & mytablex.Fields("ivap"))
calcula_igv 0
Data2.Recordset.Update
Exit Sub
cmd3461_err:
MsgBox "Error en Graba Temporales,Realizar proceso borrar Todo ", 48, "Aviso"
Exit Sub
End Sub
Function sumar_detalle()
Dim found As Integer
'found = ir_final1()
End Function


Private Sub Timer1_Timer()
fechasis = Format(Now, "dd/mm/yyyy")
horasis = Format(Now, "HH:MM:SS")
End Sub
Function busca_paridad()

Dim mytablex As Table
Dim found As Integer
paridad = "1"
paridadfp = "1"
Set mytablex = mydbxglo.OpenTable("parame")
mytablex.Index = "codigo"
mytablex.Seek "=", "01"
If Not mytablex.NoMatch Then
   paridad = "" & mytablex.Fields("parivta")
   paridadfp = "" & mytablex.Fields("parivta")
   If Val(paridad) = 0 Then
      paridad = "1"
   End If
   If Val(paridadfp) = 0 Then
      paridadfp = "1"
   End If
   busca_paridad = 1
End If
'------------------------------------- ------------
mytablex.Close
 
End Function
Sub sql_detalle()
Dim buf As String
Dim found As Integer
On Error GoTo cmd34_err
buf = "select * from " & dgusuario
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = buf
               Data2.Refresh
               DBGrid2.Refresh
               found = sumar_detalle()
               'DBGrid2.Row = DBGrid2.VisibleRows - 2
               'DBGrid2.Col = 0
               'DBGrid2.SetFocus
Exit Sub
cmd34_err:
MsgBox "Error en select " & error$, 48, "Aviso"
Exit Sub
End Sub
Function puerto_balanza1() As String
On Error GoTo cmd6712_err
Dim i As Long
Dim d As Integer
Dim buffers As String
    Select Case "" & mytable11.Fields("portbala")
           Case "COM1"
           d = 1
           Case "COM2"
           d = 2
           Case "COM3"
           d = 3
           Case "COM4"
           d = 4
           Case "COM5"
           d = 5
           
End Select

MSComm1.CommPort = d
MSComm1.Settings = "9600,n,8,1"
MSComm1.InputLen = 10
MSComm1.PortOpen = True
MSComm1.Output = Chr$(80)
buffers = ""
'For i = 1 To 9000
'Next i
i = 0
Do
'DoEvents
buffers = buffers & MSComm1.Input
i = i + 1
If i > 15000 Then
   Exit Do
End If
Loop Until Len(buffers) >= 10
cerrar_balanza
puerto_balanza1 = buffers
Exit Function
cmd6712_err:
cerrar_balanza
Exit Function
End Function
Sub cerrar_balanza()
On Error GoTo cmd892_err
MSComm1.PortOpen = False
Exit Sub
cmd892_err:
Exit Sub
End Sub
Function busca_unidad(buf As String)

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("producto")
mytablex.Index = "producto"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   If "" & mytablex.Fields("vtaund") = "S" Then
      busca_unidad = 1
   End If
   
End If
'------------------------------------- ------------
mytablex.Close
 

End Function
Function ir_primero1()
On Error GoTo cmd771222_err
Data2.Recordset.MoveFirst
'Data2.Refresh
ir_primero1 = 1
Exit Function
cmd771222_err:
Exit Function
End Function
Function ir_final1()
On Error GoTo cmd7712221_err
Data2.Recordset.MoveLast
'Data2.Refresh
ir_final1 = 1
Exit Function
cmd7712221_err:
Exit Function

End Function



