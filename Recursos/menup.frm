VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{19BD1EA6-6E36-45BA-AEBD-BCF3093017CC}#11.0#0"; "GorditoButton.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.2#0"; "Codejock.SkinFramework.v13.2.1.ocx"
Begin VB.Form menup 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Smart Business Point"
   ClientHeight    =   9720
   ClientLeft      =   150
   ClientTop       =   615
   ClientWidth     =   16980
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "menup.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   16980
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   15600
      ScaleHeight     =   615
      ScaleWidth      =   1095
      TabIndex        =   51
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSCommLib.MSComm visorcl 
      Left            =   14880
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   5
      DTREnable       =   -1  'True
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6720
      Top             =   9840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9720
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Acceso al Sistema"
      BeginProperty Font 
         Name            =   "Roboto"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9135
      Left            =   0
      TabIndex        =   25
      Top             =   1320
      Width           =   14775
      Begin VB.PictureBox Picture2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8655
         Left            =   6720
         ScaleHeight     =   8595
         ScaleWidth      =   7875
         TabIndex        =   28
         Top             =   1200
         Width           =   7935
         Begin VB.TextBox clave 
            Height          =   495
            Left            =   2160
            TabIndex        =   0
            Top             =   2640
            Width           =   2775
         End
         Begin VB.ComboBox gempresa 
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            ItemData        =   "menup.frx":0CCA
            Left            =   2160
            List            =   "menup.frx":0CCC
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   2160
            Width           =   4335
         End
         Begin VB.PictureBox Picture3 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            ScaleHeight     =   555
            ScaleWidth      =   7395
            TabIndex        =   29
            Top             =   960
            Width           =   7455
            Begin VB.Label Label8 
               Caption         =   "CONTROL DE ACCESO"
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2160
               TabIndex        =   30
               Top             =   120
               Width           =   3495
            End
         End
         Begin GorditoButton.Boton Comando 
            Height          =   750
            Index           =   0
            Left            =   720
            TabIndex        =   52
            Top             =   3315
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1323
            PicturePosition =   0
            Caption         =   "0"
            UseGif          =   -1  'True
            BackColor       =   4210752
            ResalteColor    =   12632256
            PlayGif         =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
         End
         Begin GorditoButton.Boton Comando 
            Height          =   750
            Index           =   1
            Left            =   1515
            TabIndex        =   53
            Top             =   3315
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1323
            PicturePosition =   0
            Caption         =   "1"
            UseGif          =   -1  'True
            BackColor       =   4210752
            ResalteColor    =   12632256
            PlayGif         =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
         End
         Begin GorditoButton.Boton Comando 
            Height          =   750
            Index           =   2
            Left            =   2310
            TabIndex        =   54
            Top             =   3315
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1323
            PicturePosition =   0
            Caption         =   "2"
            UseGif          =   -1  'True
            BackColor       =   4210752
            ResalteColor    =   12632256
            PlayGif         =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
         End
         Begin GorditoButton.Boton Comando 
            Height          =   750
            Index           =   3
            Left            =   3105
            TabIndex        =   55
            Top             =   3315
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1323
            PicturePosition =   0
            Caption         =   "3"
            UseGif          =   -1  'True
            BackColor       =   4210752
            ResalteColor    =   12632256
            PlayGif         =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
         End
         Begin GorditoButton.Boton Comando 
            Height          =   750
            Index           =   4
            Left            =   735
            TabIndex        =   56
            Top             =   4020
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1323
            PicturePosition =   0
            Caption         =   "4"
            UseGif          =   -1  'True
            BackColor       =   4210752
            ResalteColor    =   12632256
            PlayGif         =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
         End
         Begin GorditoButton.Boton Comando 
            Height          =   750
            Index           =   5
            Left            =   1515
            TabIndex        =   57
            Top             =   4020
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1323
            PicturePosition =   0
            Caption         =   "5"
            UseGif          =   -1  'True
            BackColor       =   4210752
            ResalteColor    =   12632256
            PlayGif         =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
         End
         Begin GorditoButton.Boton Comando 
            Height          =   750
            Index           =   6
            Left            =   2310
            TabIndex        =   58
            Top             =   4020
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1323
            PicturePosition =   0
            Caption         =   "6"
            UseGif          =   -1  'True
            BackColor       =   4210752
            ResalteColor    =   12632256
            PlayGif         =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
         End
         Begin GorditoButton.Boton Comando 
            Height          =   750
            Index           =   7
            Left            =   3105
            TabIndex        =   59
            Top             =   4020
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1323
            PicturePosition =   0
            Caption         =   "7"
            UseGif          =   -1  'True
            BackColor       =   4210752
            ResalteColor    =   12632256
            PlayGif         =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
         End
         Begin GorditoButton.Boton Comando 
            Height          =   750
            Index           =   8
            Left            =   750
            TabIndex        =   60
            Top             =   4725
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1323
            PicturePosition =   0
            Caption         =   "8"
            UseGif          =   -1  'True
            BackColor       =   4210752
            ResalteColor    =   12632256
            PlayGif         =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
         End
         Begin GorditoButton.Boton Comando 
            Height          =   750
            Index           =   9
            Left            =   1515
            TabIndex        =   61
            Top             =   4725
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1323
            PicturePosition =   0
            Caption         =   "9"
            UseGif          =   -1  'True
            BackColor       =   4210752
            ResalteColor    =   12632256
            PlayGif         =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
         End
         Begin GorditoButton.Boton Comando 
            Height          =   750
            Index           =   10
            Left            =   2310
            TabIndex        =   62
            Top             =   4725
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   1323
            PicturePosition =   0
            Caption         =   "CR"
            UseGif          =   -1  'True
            BackColor       =   4210752
            ResalteColor    =   12632256
            PlayGif         =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
         End
         Begin GorditoButton.Boton cmdIngresar 
            Height          =   750
            Left            =   6630
            TabIndex        =   63
            ToolTipText     =   "Ingresar al sistema"
            Top             =   1680
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   1323
            PicturePosition =   4
            Caption         =   "Ok"
            BackColor       =   4210752
            ResalteColor    =   12632256
            PictureDown     =   "menup.frx":0CCE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PERU"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2160
            TabIndex        =   39
            Top             =   1680
            Width           =   4335
         End
         Begin VB.Label Label3 
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " PAIS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   720
            TabIndex        =   38
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "001D"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4920
            TabIndex        =   35
            Top             =   2760
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Denomina 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2145
            Left            =   360
            TabIndex        =   33
            Top             =   5520
            Width           =   6000
         End
         Begin VB.Label Label6 
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " CLAVE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   720
            TabIndex        =   32
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " EMPRESA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   720
            TabIndex        =   31
            Top             =   2160
            Width           =   1455
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "menup.frx":1B58
         Height          =   7815
         Left            =   7920
         OleObjectBlob   =   "menup.frx":1B6C
         TabIndex        =   1
         Top             =   1320
         Width           =   5055
      End
      Begin VB.Image Image11 
         Height          =   480
         Left            =   1785
         Picture         =   "menup.frx":2547
         Top             =   8505
         Width           =   480
      End
      Begin VB.Image Image16 
         BorderStyle     =   1  'Fixed Single
         Height          =   945
         Left            =   13080
         Picture         =   "menup.frx":287D
         Stretch         =   -1  'True
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Image Image20 
         BorderStyle     =   1  'Fixed Single
         Height          =   960
         Left            =   13080
         Picture         =   "menup.frx":4027
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   1200
      End
      Begin VB.Image Image19 
         BorderStyle     =   1  'Fixed Single
         Height          =   960
         Left            =   13080
         Picture         =   "menup.frx":4F6D
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   1200
      End
      Begin VB.Image Image18 
         Height          =   960
         Left            =   13080
         Picture         =   "menup.frx":D492
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   1200
      End
      Begin VB.Image Image17 
         Height          =   960
         Left            =   13080
         Picture         =   "menup.frx":EBEF
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   34
         Top             =   600
         Width           =   7935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Configuraci�n del Servidor --->"
         BeginProperty Font 
            Name            =   "Roboto"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   8520
         Width           =   1575
      End
      Begin VB.Label vservidor 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9120
         TabIndex        =   26
         Top             =   600
         Width           =   5535
      End
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ingreso Dinero"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12240
      MaskColor       =   &H8000000E&
      Picture         =   "menup.frx":16D35
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton btnsalir 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   12240
      Picture         =   "menup.frx":18BBF
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Imprimir todo"
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "RelojPersonal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8640
      Picture         =   "menup.frx":19489
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10440
      Picture         =   "menup.frx":19A72
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&CambiaUsuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12240
      Picture         =   "menup.frx":19F6F
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "RankCompra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10440
      Picture         =   "menup.frx":1BEC5
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&ReciboEgreso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12240
      Picture         =   "menup.frx":1DA09
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ranking Venta Productos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8640
      Picture         =   "menup.frx":1DD13
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton image7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Traslado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8640
      Picture         =   "menup.frx":1FECD
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton image14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ord.Compra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12240
      MaskColor       =   &H8000000E&
      Picture         =   "menup.frx":209BE
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton image10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Requerimiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12240
      Picture         =   "menup.frx":20FA7
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton image12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Estadistica Ventas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12240
      Picture         =   "menup.frx":22849
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton image2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Productos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10440
      Picture         =   "menup.frx":2453B
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton image1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Tienda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6240
      Picture         =   "menup.frx":24AA7
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton image3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cotizaci�n"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8640
      Picture         =   "menup.frx":24EEB
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton image6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&GuiaEntrada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10440
      Picture         =   "menup.frx":254D4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton image5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&GuiaSalida"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8640
      Picture         =   "menup.frx":2801E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton image4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Fac.Compras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10440
      MaskColor       =   &H8000000E&
      Picture         =   "menup.frx":2A1D8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2385
      Width           =   1695
   End
   Begin VB.CommandButton image8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Fac.Ventas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10440
      Picture         =   "menup.frx":2A74E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Width           =   1695
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   14880
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   512
      SThreshold      =   1
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   11760
      Top             =   360
      _Version        =   851970
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label Servicio2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Servicio de Fact.Electr�nica:"
      BeginProperty Font 
         Name            =   "Roboto"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   120
      TabIndex        =   65
      Top             =   240
      Width           =   3180
   End
   Begin VB.Label Servicio 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Roboto"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3360
      TabIndex        =   64
      Top             =   240
      Width           =   1905
   End
   Begin VB.Label estasen 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   15240
      TabIndex        =   44
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Image22 
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Left            =   14520
      Picture         =   "menup.frx":2BEAC
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label ipmaquina 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   0
      TabIndex        =   36
      Top             =   9720
      Width           =   75
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  'Fixed Single
      Height          =   3375
      Left            =   120
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   5655
   End
   Begin VB.Label guactivo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2400
      TabIndex        =   22
      Top             =   720
      Width           =   120
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario Activo:"
      BeginProperty Font 
         Name            =   "Roboto"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   8640
      TabIndex        =   3
      Top             =   8400
      Width           =   5295
   End
   Begin VB.Label xxempresa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Licenciado a:"
      BeginProperty Font 
         Name            =   "Roboto"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   240
      TabIndex        =   2
      Top             =   9240
      Width           =   14775
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Refresca!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   16200
      TabIndex        =   50
      Top             =   4680
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "Egresos :0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   16200
      TabIndex        =   49
      Top             =   3960
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "Ingresos :0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   16200
      TabIndex        =   48
      Top             =   4320
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "ComprasDia :0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   16200
      TabIndex        =   47
      Top             =   3120
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "Ventas Dia   :0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   16200
      TabIndex        =   46
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Refresca!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   16200
      TabIndex        =   45
      Top             =   3480
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Refresca!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   16200
      TabIndex        =   43
      Top             =   2280
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label cuentacp 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "CuentaxCobrar:0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   16200
      TabIndex        =   42
      Top             =   1920
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label cuentacc 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "CuentaxCobrar:0.00"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   16200
      TabIndex        =   41
      Top             =   1560
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "Recuerda.."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   16200
      TabIndex        =   40
      Top             =   1200
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Menu dlotablas 
      Caption         =   "&Tablas"
   End
   Begin VB.Menu tienda1 
      Caption         =   "&Tienda"
   End
   Begin VB.Menu fdjvtas 
      Caption         =   "&Ventas"
   End
   Begin VB.Menu hiscli 
      Caption         =   "&HistoriaClinica"
   End
   Begin VB.Menu jui34hote 
      Caption         =   "&Ho&tel"
   End
   Begin VB.Menu dki444 
      Caption         =   "Activo&Fijo"
   End
   Begin VB.Menu comjk1 
      Caption         =   "&Compras"
   End
   Begin VB.Menu djk82221 
      Caption         =   "&Distribucion"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu alki1 
      Caption         =   "&Almacen"
   End
   Begin VB.Menu prod1 
      Caption         =   "&Finanzas"
      Begin VB.Menu dju33 
         Caption         =   "0.Prueba"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu dk8ct3 
         Caption         =   "&1.Cuentas x Cobrar"
      End
      Begin VB.Menu cuy7812 
         Caption         =   "&2.Cuentas por Pagar"
      End
      Begin VB.Menu b889 
         Caption         =   "&3.Bancos"
         Visible         =   0   'False
      End
      Begin VB.Menu bca88 
         Caption         =   "&3.Caja"
      End
      Begin VB.Menu DocuXCobrar 
         Caption         =   "&4.Documentos x cobrar"
      End
   End
   Begin VB.Menu bando1 
      Caption         =   "&Produccion"
   End
   Begin VB.Menu dlo1211 
      Caption         =   "Plani&Llas"
   End
   Begin VB.Menu cont121 
      Caption         =   "&Contable"
   End
   Begin VB.Menu li8931utyi 
      Caption         =   "&Util"
   End
   Begin VB.Menu ldossali4 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "menup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
'"Insert into [SERVICIOS].dbo.Pedidos select * from [BASED].dbo.Pedidos ."
' 003 AXES
' 004 GROGUI
'CONTADORES ACUMULABLES EN UNA TABLA
'EMPIEZA AQUI
'BEGIN TRANSACTION
'UPDATE Contadores SET0 ultimoNumeroUnaTabla = ultimoNumeroUnaTabla + 1
'n = [SELECT 12 FROM Contadores]
'INSERT INTO unaTabla (Numero, ...) VALUES ([n], ...)
'COMMIT TRANSACTION
'TERMINAL AQUI

'PARA ACTUALIZAR EN 02 TABLAS ALGUN DARO
'Update FPAGOV
'Set FPAGOV.estado = FACTURA.estadom
'FROM         FPAGOV INNER JOIN
'                      FACTURA ON FACTURA.LOCAL = FPAGOV.LOCAL AND FACTURA.TIPO = FPAGOV.TIPO AND FACTURA.SERIE = FPAGOV.SERIE AND
'                      FACTURA.numero = FPAGOV.numero
'
'
'
'
'
' nononoo-------
'---------------
'---------------

'Update Facturas, Cobros
'Set Facturas.Pagados = Facturas.Pagados + Cobros.importe
'Where Facturas.numero = Cobros.Factura

'ivap
'Update factura
'SET              TIVAP =
'                          (SELECT     SUM(detalle.tivap)
'                            From detalle
'                            WHERE      detalle.local = factura.local AND detalle.tipo = factura.tipo AND detalle.serie = factura.serie AND detalle.numero = factura.numero)

'ojo

'Update fpagov
'SET              ESTADO =
'                          (SELECT     factura.estado
'                            From factura
'                            WHERE      factura.local = fpagov.local AND factura.tipo = fpagov.tipo AND factura.serie = fpagov.serie AND factura.numero = fpagov.numero)
'ojo2

'INSERT INTO SUBFAMIL
'                      (SUBFAMILIA, DESCRIPCIO, FAMILIA)
'SELECT DISTINCT SUBFAMILIA, FAMILIA, FAMILIA AS Expr1
'From producto

Option Explicit

Dim ultimo_costo       As Double

Dim rtn                As Long

'ojo poner el tipoimp='C' si es compras

'Facturacion Electronica Servicio 26/03/2018
Dim WithEvents objServ As servicios
Attribute objServ.VB_VarHelpID = -1

Dim rptaServicio       As String

'Facturacion Electronica Servicio 26/03/2018

'Testing Proyecto Facturacion Electronica 05/04/2018
Dim rpta               As String

'Testing Proyecto Facturacion Electronica 05/04/2018

Private Sub abnonoer2_Click()

End Sub

Private Sub alki1_Click()
    treev.Show 1

End Sub

Private Sub alo112_Click()
    ttlocal.Show 1

End Sub

Private Sub b889_Click()
    treevbco.Show 1

End Sub

Private Sub bando1_Click()
    treevpro.Show 1

End Sub

Private Sub bandoker_Click()
    tbanco.Show 1

End Sub

Private Sub bca88_Click()
    treevcaj.Show 1

End Sub

Private Sub btnsalir_Click()
    ldossali4_Click

End Sub

Private Sub caje12_Click()
    tcategor.Show 1

End Sub

Private Sub cargtw_Click()
    tcase.Show 1

End Sub

Private Sub clasifo33_Click()
    tclasifi.Show 1

End Sub

Sub visualiza_mesa()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        glomesa = Trim("" & mytablex.Fields("mesaseccion"))

    End If

    mytablex.Close

    If Len(Trim(glomesa)) = 0 Then
        glomesa = "Mesa"

    End If

End Sub

Sub palabra_bienvenida()

End Sub

Sub palabra_bienvenida1()

End Sub

Function abrir_usuario_global()

    On Error GoTo cmd78888_err

    Set mytablexxx = mydbxglo.OpenTable("__" & gusuario & ".dbf")
    abrir_usuario_global = 1
    Exit Function
cmd78888_err:
    MsgBox "Error en abrir usuario " + error, 48, "Aviso"
    Exit Function

End Function

Function busca_clave(buf As String)

    Dim buf1 As String

    Dim texi As New ADODB.Recordset

    gusuario = ""
    image2.Enabled = False  'tablas
    hiscli.Enabled = False
    image1.Enabled = False   'tienda1
    image8.Enabled = False
    Command2.Enabled = False
    image4.Enabled = False
    image14.Enabled = False
    Command4.Enabled = False
    image5.Enabled = False
    image6.Enabled = False
    image7.Enabled = False
    Command8.Enabled = False
    Command3.Enabled = False
    image10.Enabled = False
    image12.Enabled = False
    image3.Enabled = False
    Command7.Enabled = False
    Command6.Enabled = False
    hiscli.Enabled = False
    dki444.Enabled = False

    texi.Open "SELECT * FROM vendedor where clave='" & Trim(buf) & "'", cn, adOpenKeyset, adLockOptimistic

    If texi.RecordCount = 0 Then  'si existe
        texi.Close
        Exit Function

    End If

    busca_clave = 1
    gusuario = "" & texi.Fields("codigo")
    guactivo = gusuario & "-" & "" & texi.Fields("nombre")
    ngusuario = "" & texi.Fields("nombre")
    'buf1 = "" & texi.Fields("visibles")

    '  If Trim("" & texi.Fields("v14")) = "S" Then
    '     'MsgBox "" & texi.Fields("v14")
    '     jui34hote.Visible = True
    '     jui34hote.Enabled = True
    '   End If
    '
    If "" & texi.Fields("v13") = "S" Then
        dki444.Visible = True
        dki444.Enabled = True

    End If

    If "" & texi.Fields("v1") = "S" Then
        dlotablas.Visible = True
        dlotablas.Enabled = True
        image2.Enabled = True

    End If

    If "" & texi.Fields("v2") = "S" Then
        tienda1.Visible = True
        tienda1.Enabled = True
        image1.Enabled = True
        image3.Enabled = True

    End If

    If "" & texi.Fields("v3") = "S" Then
        fdjvtas.Visible = True
        fdjvtas.Enabled = True
        image8.Enabled = True
        Command2.Enabled = True
        Command6.Enabled = True

    End If

    If "" & texi.Fields("v4") = "S" Then
        comjk1.Visible = True
        comjk1.Enabled = True
        image4.Enabled = True
        image14.Enabled = True
        Command4.Enabled = True

    End If

    If "" & texi.Fields("v5") = "S" Then
        alki1.Visible = True
        alki1.Enabled = True
        image5.Enabled = True
        image6.Enabled = True
        image10.Enabled = True
        image7.Enabled = True

    End If

    If "" & texi.Fields("v6") = "S" Then
        prod1.Visible = True
        prod1.Enabled = True
        Command8.Enabled = True
        Command3.Enabled = True
      
    End If

    If "" & texi.Fields("v7") = "S" Then
   
        ' produccion
        '''27/07/2017 kenyo Testing Completo al Sistema
        'bando1.Visible = True
        'bando1.Enabled = True
        '''27/07/2017 kenyo Testing Completo al Sistema
    
        'Image10.Enabled = True
    End If

    If "" & texi.Fields("v8") = "S" Then
        dlo1211.Visible = True
        dlo1211.Enabled = True
        Command7.Enabled = True

        'Command6.Enabled = True
    End If

    If "" & texi.Fields("v9") = "S" Then
      
        ' produccion
        '''27/07/2017 kenyo Testing Completo al Sistema
        'cont121.Visible = True
        'cont121.Enabled = True
        '''27/07/2017 kenyo Testing Completo al Sistema
      
    End If

    If "" & texi.Fields("v10") = "S" Then
        'i893432.Visible = True
        'i893432.Enabled = True
        image12.Enabled = True

    End If

    If "" & texi.Fields("v11") = "S" Then
        hiscli.Visible = True
        hiscli.Enabled = True

    End If

    If "" & texi.Fields("v12") = "S" Then

        'IM89ds.Visible = True
        'IM89ds.Enabled = True
    End If

    Command7.Enabled = False

    If "" & texi.Fields("VRELOJ") = "S" Then
        Command7.Enabled = True

    End If

    If "" & texi.Fields("tienda") = "S" Then
        image1.Enabled = True

    End If

    If "" & texi.Fields("PRODUCTOS") = "S" Then
        image2.Enabled = True

    End If

    If "" & texi.Fields("terminal") = "S" Then
        image3.Enabled = True

    End If

    If "" & texi.Fields("minireporte") = "S" Then
        Command6.Enabled = True

    End If
   
    texi.Close

End Function

Function cambia_directorio()

    On Error GoTo cmd21_err

    'ChDir ("c:\rp_orion\001d\" & glocal)
    ChDir (globaldir)
    cambia_directorio = 1
    Exit Function
cmd21_err:
    MsgBox "Error " & error$, 24, "Aviso"
    Exit Function

End Function

Function busca_empresa(buf As String)

    On Error GoTo cmd543_err

    Dim texi As New ADODB.Recordset

    'MsgBox buf

    texi.Open "SELECT nombre FROM empresa where codigo='" & Trim(buf) & "'", cn, adOpenKeyset, adLockOptimistic

    If texi.RecordCount = 0 Then  'si existe
        texi.Close
        Exit Function

    End If

    'glocal = "" & texi.Fields("localdefecto")
    'MsgBox "abc"
    'Label1.Caption = "" & texi.Fields("nombre")
    busca_empresa = 1
    texi.Close
    Exit Function
cmd543_err:
    MsgBox "Aviso en busca Empresa " + error$, 48, "Aviso"
    busca_empresa = 0
    Exit Function

End Function

Function busca_local()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from tlocal where defecto='S'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_local = 1
        glocal = Trim("" & mytablex.Fields("codigo"))

    End If

    mytablex.Close

End Function

Private Sub clet11_Click()

End Sub

Private Sub clet112_Click()

End Sub

Private Sub clet11pag_Click()

    'explreci.Caption = "EGRESO DINERO"
    'explreci.afecta = "L"
    'explreci.acu = "V"
    'explreci.Show 1
    'Dim found As Integer
    'found = copiar_recibos()
    'If found = 0 Then
    '   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
    '   End
    '   Exit Sub
    'End If
    'fgusuario = "_r" & gusuario
    'trecaja.Caption = "EGRESO DINERO"
    'trecaja.afecta = "L"
    'trecaja.acu = "V"
    'trecaja.Show 1
End Sub

Private Sub clave_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        'gempresa.SetFocus
        Exit Sub

    End If

End Sub

Private Sub cli94_Click()
    'ttnclie.show 1
    'explocli.Show 1
    tnclie.DBPROV = "clientes"
    tnclie.Show 1

End Sub

Private Sub cmom9934_Click()

    'tprodup.Caption = "Tabla de productos Insumos"
    'tprodup.insumo.Value = 1
    'tprodup.Show 1
End Sub

Private Sub cnju44_Click()
    tcaja.Show 1

End Sub

Private Sub cmdIngresar_Click()
 
    clave_KeyPress 13

End Sub

Private Sub clave_KeyPress(KeyAscii As Integer)

    Dim found   As Integer

    Dim xcampo1 As String

    Dim xcampo2 As String

    Dim xcampo3 As String

    Dim salida  As Boolean
 
    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Exit Sub

    End If
 
    'If IsValidIPAddress(vservidor) = False Then
    '   MsgBox "Ingrese Ip Valida"
    '   clave.SetFocus
    '   Exit Sub
    'End If
    clave = UCase(clave)

    If Len(clave) = 0 Then
        clave.SetFocus
        Exit Sub

    End If

    '----------------------
    cerrar_base
    glocal = "01"

    If Len(Trim("" & gempresa)) = 0 Then
        clave = ""
        clave.SetFocus
        Exit Sub

    End If

    'MsgBox "abc"

    'If Len(vservidor) = 0 Then
    '   MsgBox "No existe Vendedor ", 48, "Aviso"
    '   clave.SetFocus
    '   Exit Sub
    'End If
    menup.Caption = nombre_sistema & " Empresa:" & gempresa
    found = extraer_campos1(gempresa, xcampo1, xcampo2, xcampo3)
    valida_conec xcampo2

    If Len(Trim(vservidor)) = 0 Then
        clave = ""
        clave.SetFocus
        Exit Sub

    End If

    'If Len(Trim(clave_servidor)) = 0 Then
    '   clave = ""
    '   clave.SetFocus
    '   Exit Sub
    'End If
    'MsgBox "abc"
    'MsgBox extra_loquesea1(gempresa)
    'MsgBox xcampo2
    found = conectar(xcampo2)

    If found = 0 Then
        MsgBox "Error de Conexion Sql Server ", 48, "Aviso"
        clave.SetFocus
        Exit Sub

    End If

    'found = conectara()
    'If found = 0 Then
    '   MsgBox "Error de Conexion Sql Server ", 48, "Aviso"
    '   clave.SetFocus
    '   Exit Sub
    'End If
    'SETEAR GLOGALDIR- SETEO
    globaldir = globalpath & "\001d\06"
    globaldat = globalpath & "\001d\06"
    globalcont = globalpath & "\001d\contable"
    globalpri = globalpath & "\001d"
    globalweb = globalpath & "\001d\06\web"
    orionv4 = "\orion.v4\001d\01"
    'MsgBox gempresa
    empresapos = "01"
    'found = busca_empresa(extra_loquesea(gempresa))
    'If found = 0 Then
    '   MsgBox "Empresa No existe", 48, "Aviso"
    '   clave = ""
    '   clave.SetFocus
    '   Exit Sub
    'End If
    globalemp = extra_loquesea(gempresa)

    'Label1 = Label9
    found = busca_clave("" & clave)

    If found = 0 Then
        MsgBox "No existe Clave asignado a ningun Funcionario", 48, "Aviso"
        clave = ""
        clave.SetFocus
        Exit Sub

    End If

    ' 26/07/2018 Desactivar Facturacion Electronica
    V_EstadoSistema = Obtiene_EstadoSistema()
    ' 26/07/2018 Desactivar Facturacion Electronica

    'Testing Proyecto Facturacion Electronica 05/04/2018
    If V_EstadoSistema = "FE BYH" Then
        Servicio.Visible = True
        Servicio2.Visible = True
        Set objServ = New servicios
    
        rptaServicio = objServ.ObtenerEstado("facturador-local")

        If rptaServicio = "El servicio est� detenido" Then
            Servicio = "Servicio Detenido"
            Servicio.BackColor = &HFF&
        ElseIf rptaServicio = "El servicio est� activo" Then
            Servicio = "Servicio Activo"
            Servicio.BackColor = &HFF8080

        End If
    
        If Servicio = "Servicio Detenido" Then
            Shell ("C:\BYH\iniciar.bat"), vbNormalFocus
            Servicio = "Servicio Activo"
            Servicio.BackColor = &HFF8080

        End If

    ElseIf V_EstadoSistema = "CONINT" Then
        Servicio.Visible = False
        Servicio2.Visible = False

    End If

    ' 26/07/2018 Desactivar Facturacion Electronica

    'MsgBox "abc"
    'found = busca_local()
    'If found = 0 Then
    If Len(Trim(glocal)) = 0 Then
        glocal = "01"

    End If

    If verificador_datos("" & gusuario) = "S" Then
        Frame2.Visible = False
        tncr1.Show 1
        barra_herramienta 1
        cerrar_base
        End
        Exit Sub

    End If

    'MsgBox "abc"

    'Set mydbzglo = OpenDatabase(globalcont, False, False, "foxpro 2.5;")
    'MsgBox "xxxx"
    Set mydbxglo = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
    'MsgBox "xxx"
    found = cargas_iniciales()

    'MsgBox "xxxx"
    If found = 0 Then

    End If

    'found = verifica_licenciaremoto()
    '--------aqui verifica que no ingrese otro con el mismo usuario
    found = copiar_temporalxxx()

    If found = 0 Then

        'MsgBox "Usuario ya Activo ", 48, "Aviso"
        'clave = ""
        'clave.SetFocus
        'Exit Sub
    End If

    'MsgBox "abc"
    found = abrir_usuario_global()

    If found = 0 Then
        MsgBox "Error al abrir control ", 48, "Aviso"
        clave = ""
        clave.SetFocus
        Exit Sub

    End If

    'fdk9833.Visible = True
    palabra_bienvenida1
    globalocal = glocal
    'recordar
    'menup.Caption = "SISTEMA ORION V5.0 " '& Label1
    Frame2.Visible = False
    recordar
    'visualizar_recordar
    'visualizar_ingresos
    visualiza_mesa
    'kenyo 20/04/2017
    'opcion2 = "4"
    'repinv.excell.Visible = True
    'repinv.Label17.Visible = True
    'repinv.Combo1.Visible = True
    'repinv.Label25.Visible = True
    'repinv.gcanti.Visible = True
    'repinv.excell.Value = 1
    'repinv.Combo1.Text = "TODOS"
    'repinv.CargaInicial
    'repinv.Show 1
    'repinv.Hide
    'inicio 28/11/2017 pll para las configuracion monedas
    Call configura_moneda(my_vdolar, salida)

    If salida = False Then
        MsgBox "Ingresar el dato en Tablas -->Parametros Generales/ve Dolar "

    End If

    'obtenemos estado del lfacturador sunat 02/02/2018
    'MsgBox objServ.ObtenerEstado("facturador-local") & " - Facturador Electronico Sunat", vbInformation
  
    'Testing Proyecto Facturacion Electronica 05/04/2018
    '        Dim rpta As String
    '
    '        Call valida_facturacionElectronica
    '        If rpta = "" Then
    '            respuesta.Text = "TODO OK"
    '        Else
    '            rptaa.Visible = True
    '            respuesta.Text = rpta
    '        End If
    'Testing Proyecto Facturacion Electronica 05/04/2018
  
    'fi 28/11/2017 pll
End Sub


Private Sub Comando_Click(Index As Integer)

    If Index = 10 Then
        clave = ""
        Exit Sub

    End If

    clave = clave & Comando(Index).Caption

End Sub

Private Sub comjk1_Click()
    treevc.Show 1

End Sub

Private Sub Command2_Click()
    opcion2 = "2"
    'inicio 07/09/2017 pll
    'repraped.Label12.Visible = True
    'repraped.orden.Visible = True
    'fin 07/09/2017 pll
    repraped.acu = "V" 'PEDIDO
    repraped.xdata = "Detalle"
    repraped.Show 1

End Sub

Private Sub Command3_Click()
    explreci.xcuentaco = "Cuentap"
    explreci.Caption = "EGRESO DINERO"
    'explreci.afecta = "P"  'proveedor
    explreci.acu = "V"
    explreci.Show 1

End Sub

Private Sub Command4_Click()
    opcion2 = "2"
    'inicio 07/09/2017 pll
    'repraped.Label12.Visible = True
    'repraped.orden.Visible = True
    'fin 07/09/2017 pll
    repraped.acu = "C" 'PEDIDO
    repraped.xdata = "DETALLE"
    repraped.Show 1

End Sub

Private Sub Command5_Click()
    'visualizar_recordar
    esconde_todo
    esconde_menu
    cerrar_usuario_global
    'activar1
    activar2
    cerrar_base
    Frame2.Visible = True
    vservidor = ""
    clave_servidor = ""

    clave = ""
    clave.SetFocus

End Sub

Sub cerrar_usuario_global()

    On Error GoTo cmd543311_err

    mytablexxx.Close
    Exit Sub
cmd543311_err:
    Exit Sub

End Sub

Private Sub Command6_Click()
    flag_clave1 = 0
    tconcla.X = "MINIREPORTE"
    tconcla.Show 1

    If flag_clave1 <> 1 Then  'si es descongela
        Exit Sub

    End If

    tresegui.Show 1

End Sub

Private Sub Command7_Click()
    tingper.Show 1

End Sub

Private Sub Command8_Click()
    explreci.xcuentaco = "Cuentac"
    explreci.Caption = "INGRESO DINERO"
    'explreci.afecta = "C"
    explreci.acu = "W"
    explreci.Show 1

End Sub

Private Sub comu734_Click()
    repingre.acu = "W"
    repingre.Show 1

End Sub

Private Sub d892321_Click()
    menucaja.Label3 = "USUARIO"
    menucaja.acu = "T"
    menucaja.tipoterminal = "TOUCH"
    menucaja.Show 1

End Sub

Private Sub cont121_Click()
    treevcon.Show 1

End Sub

Private Sub cuentacc_Click()

    If prod1.Visible = True Then
        texplcxc.xcuentaco = "cuentac"
        texplcxc.XCUENTACO1 = "cuentacd"
        texplcxc.xcuentacol = ""
        'texplcxc.tipoclie = "C"
        texplcxc.ldo232.Enabled = False
        texplcxc.mofdi782.Enabled = False
        texplcxc.dj333.Enabled = False
        texplcxc.dj7823.Enabled = False
        texplcxc.ncu773.Enabled = False
        texplcxc.acu = "V"
        texplcxc.Show 1

    End If

End Sub

Private Sub cuentacp_Click()

    If prod1.Visible = True Then
        texplcxc.xcuentaco = "cuentaP"
        texplcxc.XCUENTACO1 = "cuentapd"
        texplcxc.xcuentacol = ""
        texplcxc.ldo232.Enabled = False
        texplcxc.mofdi782.Enabled = False
        texplcxc.dj333.Enabled = False
        texplcxc.dj7823.Enabled = False
        texplcxc.ncu773.Enabled = False

        'texplcxc.tipoclie = "C"
        texplcxc.acu = "C"
        texplcxc.Show 1

    End If

End Sub

Private Sub cuy7812_Click()
    treevfp.Show 1

End Sub

Private Sub dclocor_Click()

    tncolor.Show 1

End Sub

Private Sub dfju3434_Click()

End Sub

Private Sub dfki83434_Click()

End Sub

Private Sub dfk88222_Click()

    Dim found As Integer

    Dim buf   As String

    found = copiar_tcxcre()

    If found = 0 Then
        MsgBox "No se realizar este proceso ", 48, "Aviso"
        Exit Sub

    End If

    '"_b" & gusuario
    expfpa1.Show 1

End Sub

Private Sub dflo87744_Click()
    tplaconc.Show 1

    'tplanico.Show 1
End Sub

Private Sub djnwuew2_Click()

End Sub

Private Sub di88232_Click()

End Sub

Private Sub DirectSS1_ClickIn(ByVal X As Long, ByVal Y As Long)

End Sub

Private Sub dj7722_Click()

End Sub

Private Sub dfkvt623_Click()
    cgusuario = "FACTURA"
    dgusuario = "_d" & gusuario
    fgusuario = "_f" & gusuario
    dgusuariog = "DETALLE"
    tcomvta.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    tcomvta.fechaf = Format(Now, "dd/mm/yyyy")
    tcomvta.fechai = Format(Now, "dd/mm/yyyy")
    tcomvta.Caption = "Documentos Facturacion Compras Ventas"
    tcomvta.tipoclie = "%"
    tcomvta.acu = "%"
    tcomvta.Show 1

End Sub

Private Sub dhy72323_Click()

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then

        'Image19_Click
    End If

End Sub

Private Sub dj7822_Click()

End Sub

Private Sub dj7823_Click()
    opcion2 = "2"
    trepasis.Show 1

End Sub

Private Sub dj78231_Click()
    opcion2 = "12"   'analisis de ventas
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    'repdocum.Label18.Visible = False
    'repdocum.Combo1.Visible = False
    repdocum.vdetalle.Enabled = False
    repdocum.vfpago.Enabled = False
    repdocum.acu = "C"
    repdocum.Show 1

End Sub

Private Sub dj78232_Click()

End Sub

Private Sub dj782322_Click()

End Sub

Private Sub dj78t562_Click()

End Sub

Private Sub dj88obra_Click()

    'tobrasc.Show 1
End Sub

Private Sub djcont44_Click()

End Sub

Private Sub dji33_Click()

    'thiscli.Show 1
End Sub

Private Sub dj8232_Click()

End Sub

Private Sub dj82343_Click()
    tinterfa.Show 1

End Sub

Private Sub dj83auto_Click()

End Sub

Private Sub dj883355_Click()

End Sub

Private Sub djjue_Click()
    tdocumen.Show 1
    'ttipodoc.Show 1

End Sub

Private Sub djk7722_Click()
    tinterfa.Show 1

End Sub

Private Sub djret900_Click()

End Sub

Private Sub dju33_Click()
    Texcell.Show 1

End Sub

Private Sub dju333_Click()
    tlolfar.Show 1

End Sub

Private Sub dju748_Click()
    tfuentec.Show 1

End Sub

Private Sub dju7823_Click()

End Sub

Private Sub dju782321_Click()

End Sub

Private Sub dju78343_Click()
    tconecta.Show 1

End Sub

Private Sub dju823_Click()

    opcion2 = "1"
    repraped.acu = "C" 'PEDIDO
    repraped.xdata = "DETALLE"
    repraped.Show 1

End Sub

Private Sub dju8823_Click()

End Sub

Private Sub djuctee_Click()

End Sub

Private Sub djucuagre_Click()

End Sub

Private Sub djuranh3_Click()

End Sub

Private Sub dki22di3_Click()

End Sub

Private Sub dk78232_Click()

End Sub

Private Sub djurpre9_Click()
    tctadef.Show 1

End Sub

Private Sub dk782323_Click()
    opcion2 = "1"
    trepasis.Show 1

End Sub

Private Sub dk8822_Click()

    Dim buf As String

    buf = extra_loquesea(menup.gempresa)

    'If buf = "003D" Or buf = "004D" Then
    '   tacvta.Show 1
    'End If
End Sub

Private Sub dk882299_Click()

End Sub

Private Sub dk8823_Click()

    trepprov.Show 1

End Sub

Private Sub dk890oo1_Click()

End Sub

Private Sub dk8923_Click()
    tmargen.Show 1

End Sub

Private Sub dk89232_Click()

End Sub

Private Sub dk892323_Click()
    opcion2 = "3"
    'inicio 07/0972017 pll
    'repraped.Label12.Visible = True
    'repraped.orden.Visible = True
    'fin 07/0972017 pll
    repraped.acu = "Q" 'PEDIDO
    repraped.xdata = "DREQUISA"
    repraped.Show 1

End Sub

Private Sub dk8344_Click()
    tdatarep.Show 1

End Sub

Private Sub dk8ct3_Click()
    treevfi.Show 1

End Sub

Private Sub dk9232321_Click()
    tplape.Show 1

End Sub

Private Sub dk93331_Click()
    tempresa.Show 1

End Sub

Private Sub dki232311_Click()
    explreci.Caption = "INGRESO DINERO"
    'explreci.afecta = "C"
    explreci.acu = "W"
    explreci.Show 1

End Sub

Private Sub dki3434_Click()
    opcion2 = "2"
    'inicio 07/09/2017 pll
    'repraped.Label12.Visible = True
    'repraped.orden.Visible = True
    'inicio 07/09/2017 pll
    repraped.acu = "C" 'PEDIDO
    repraped.xdata = "DETALLE"
    repraped.Show 1

End Sub

Private Sub dki34882_Click()
    tubica.Show 1

End Sub

Private Sub dkicoir_Click()

    'explreci.Caption = "INGRESO DINERO"
    'explreci.afecta = "C"  'clientes
    'explreci.acu = "W"
    'explreci.Show 1
End Sub

Private Sub dkicon6_Click()
    tplaconc.Show 1

End Sub

Private Sub dkiegtr44_Click()
    'If Forms.Count > 1 Then
    '   Forms(1).SetFocus
    '   Exit Sub
    'End If
    repingre.acu = "V"
    repingre.Show 1

End Sub

Private Sub dkifor4_Click()
    teditor.Show 1

End Sub

Private Sub dkiglo_Click()
    ttparame.Show 1

End Sub

Private Sub dkizoind_Click()
    tnzona.Show 1

End Sub

Private Sub dkl78232_Click()
    'tsisper.Show 1
    tingper.Show 1

End Sub

Private Sub dkoi7823_Click()
    tturno.Show 1

End Sub

Private Sub dl8923_Click()

    'barracod.Show 1
    'tprodup.Show 1
End Sub

Private Sub dlkpago3_Click()
    tfpago.Show 1

End Sub

Private Sub dlo1_Click()

    If MsgBox("Desea Salir del Sistema", 1, "Aviso") <> 1 Then Exit Sub
    End

End Sub

Private Sub dki444_Click()
    treevaf.Show 1

End Sub

Private Sub dlo1211_Click()
    treevpla.Show 1

End Sub

Private Sub dlo232_Click()
    tccosto.Show 1

End Sub

Private Sub dlo4341_Click()
    talmacen.Show 1

End Sub

Private Sub dlo745_Click()
    tplagepe.Show 1

End Sub

Private Sub dlor23_Click()
    xprodet.Show 1

End Sub

Private Sub dlowe2_Click()
    'Dim found As Integer
    'found = copiar_recibos()
    'If found = 0 Then
    '   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
    '   End
    '   Exit Sub
    'End If
    'fgusuario = "_r" & gusuario
    'trecaja.Caption = "EGRESO DINERO"
    'trecaja.afecta = "L"
    'trecaja.acu = "V"
    'trecaja.Show 1

End Sub

Private Sub dlsd_Click()

    Dim buf As String

    Dim op  As String

    Dim I   As Integer

    For I = 1 To 10
        'print_com "HOLA MUDNO COMO ESTAS" + Chr$(10) + Chr$(13)
        'print_com "HOLA MUDNO COMO ESTAS" + Chr$(10) + Chr$(13)
    Next I

    Exit Sub

    If InputBox("clave", op, "") <> "JELC" Then Exit Sub
    'trecisc.Show 1

    'serversq.Show 1
    Exit Sub
    buf = Chr$(27) + "i"  'epson
    buf = Chr$(27) + "p" + Chr$(0) + Chr$(25) + Chr$(250) 'EPSON
    'impresion_codbar1 buf

End Sub

Private Sub dmat64_Click()
    opcion2 = "94"
    'inicio 07/09/2017 pll
    'repinv.excell.Visible = True
    'repinv.Label17.Visible = True
    'repinv.Combo1.Visible = True
    'repinv.Label25.Visible = True
    'repinv.gcanti.Visible = True
    'fin 07/09/2017 pll
    repinv.Show 1

End Sub

Private Sub ehyer343_Click()
    explreci.Caption = "EGRESO DINERO"
    'explreci.afecta = "P"  'proveedor
    explreci.acu = "V"
    explreci.Show 1

    'explreci.Caption = "EGRESO DINERO"
    'explreci.afecta = "C"
    'explreci.acu = "V"
    'explreci.Show 1

    'Dim found As Integer
    'found = copiar_recibos()
    'If found = 0 Then
    '   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
    '   End
    '   Exit Sub
    'End If
    'fgusuario = "_r" & gusuario'

    'trecaja.Caption = "EGRESO DINERO"
    'trecaja.afecta = "C"
    'trecaja.acu = "V"
    'trecaja.Show 1

End Sub

Private Sub exmilo9232_Click()

End Sub

Private Sub emo8923_Click()

End Sub

Private Sub em882_Click()

End Sub

Private Sub emi8tra_Click()

End Sub

Private Sub er299_Click()
    repfpago.Show 1

End Sub

Private Sub exopro2_Click()

    'pocket.Show 1
End Sub

Private Sub facrier_Click()

End Sub

Private Sub facu4545_Click()

    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    repdocum.acu = "C"
    repdocum.Show 1

End Sub

Private Sub facy633_Click()

End Sub

Private Sub facyu7222_Click()

End Sub

Private Sub fami34_Click()

    ttfamilia.Show 1

End Sub

Private Sub fdkmnau64_Click()

End Sub

Private Sub fd9944_Click()

End Sub

Private Sub fdju3431_Click()

End Sub

Private Sub fdju8834_Click()
    opcion2 = "1"
    repraped.acu = "I" 'PEDIDO
    repraped.xdata = "dpedidov"
    repraped.Show 1

End Sub

Private Sub dlotablas_Click()
    treevtab.Show 1

End Sub

Private Sub DocuXCobrar_Click()
    treevfi.Show 1

End Sub

Private Sub fdjvtas_Click()
    treevv.Show 1

End Sub

Private Sub fdk882_Click()
    tximp.Show 1

End Sub

Private Sub fdk883_Click()
    tgrupopl.Show 1

End Sub

Private Sub fdki44_Click()

End Sub

Private Sub fdkji89343_Click()

End Sub

Private Sub fdlo12_Click()

End Sub

Private Sub fdorm8823_Click()

    'explocli.Show 1
End Sub

Private Sub fdlo4_Click()

End Sub

Private Sub fdo833t_Click()

End Sub

Private Sub fk34322_Click()

End Sub

Private Sub fk66344_Click()

End Sub

Private Sub fk8923_Click()
    ingclie.Show 1

End Sub

Private Sub fk8934_Click()
    TUNIFLEX.Show 1

    'tprntw.Show 1
End Sub

Private Sub fk8944_Click()

End Sub

Private Sub fki44_Click()

End Sub

Private Sub fl494_Click()

End Sub

Private Sub flkpro1_Click()
    tnclie.DBPROV = "proveedo"
    tnclie.Caption = "Tabla de Proveedores "
    tnclie.Show 1

    'tnprov.Show 1
    'tnprov.show 1
End Sub

Private Sub flopro992_Click()
    planopro.Show 1

End Sub

Private Sub fmi7343_Click()

End Sub

Sub abrir_global_principal()

    On Error GoTo cmd323232_err

    Set mydbxxx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
    Exit Sub
cmd323232_err:
    Exit Sub

End Sub

Private Sub gac76222_Click()

End Sub

Private Sub Form_Activate()

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    On Error GoTo cmd89555_err

    'glomesa = "Mesa"
    'visualiza_mesa
    'MsgBox GetMACAddress2
    codigohuella = ""

    If Len(Trim(ipmaquina)) = 0 Then
        ipmaquina = "" & Winsock1.LocalIP

    End If

    'ipmaquina = Trim(RecuperarIP)
    'MsgBox GetMACs_AdaptInfo()
    carga_bases
    'visualizar_recordar
    'recordar
    found = leer_visor()

    '''''LEONARDO 15-12-2016--
    Dim I As Integer

    For I = 0 To 10
        'comando(i).Sound = App.path & "\Sonido\click.wav": comando(i).PlaySound = InClick
        ' comando(i).Sound = "C:\Windows\click.wav": comando(i).PlaySound = InClick
    Next
    'cmdIngresar.Sound = App.path & "\Sonido\cash.wav": cmdIngresar.PlaySound = InClick
    'cmdIngresar.Sound = "C:\Windows\cash.wav": cmdIngresar.PlaySound = InClick

    '--------------------------------
    Exit Sub
cmd89555_err:
    MsgBox "Aviso en Activate " + error$, 48, "Aviso"
    End
    Exit Sub

End Sub

Private Sub Form_Load()
    SkinFramework1.LoadSkin App.path & "/Skins/Gilouche.cjstyles" & "", ""
    SkinFramework1.ApplyWindow Me.hwnd
    SkinFramework1.ApplyOptions = SkinFramework1.ApplyOptions Or xtpSkinApplyMetrics

    Frame2.BackColor = RGB(91, 110, 128)
    menup.BackColor = RGB(91, 110, 128)
    Frame2.Top = 0
    
    'xxempresa.ForeColor = RGB()
    Label7.BackColor = RGB(151, 213, 201)
    
    
    

    Dim found     As Integer

    Dim xtelefono As String

    Dim X         As Integer

    On Error GoTo cmd78321_err

    swptovta = "0" '0 normal  1 ptovta
    flag_denisse = "0"   '1 denise 0 normal
    vservidor = ""
    clave_servidor = ""

    ''01/07/2017 Edici�n de Razon Social
    'nombre_sistema = "CasVisiOrion                                                                                     Visitec S.A.C"
    'nombre_sistema1 = "CasVisiOrion Visitec S.A.C"

    ' 26/07/2018 Desactivar Facturacion Electronica

    ' 02/08/2018 Edicion Nombre de Sistema
    nombre_sistema1 = "Smart Business Point"
    ' 02/08/2018 Edicion Nombre de Sistema

    nombre_sistema = nombre_sistema1
    ' 26/07/2018 Desactivar Facturacion Electronica

    ''01/07/2017 Edici�n de Razon Social

    If App.PrevInstance = True Then
        MsgBox "Ya existe " & nombre_sistema & " Ejecutandose...", vbExclamation
        End
        Exit Sub

    End If

    If Trim(Label10) = "ARGENTINA" Then
        estasen = "ARGENTINA"
        dicargentina

    End If

    If Trim(Label10) = "PERU" Then
        estasen = "PERU"
        dicperu

    End If

    xxxsoles = dicmoneda
    menup.Caption = nombre_sistema
    'Ajusta Me
    'Reajusta Me
    'dji33.Visible = False
    'i893432.Visible = False
    esconde_menu
    'gempresa.Clear
    'gempresa.AddItem "01|CALIPSO"
    'gempresa.ListIndex = 0
    '
    'globalpath = "\\192.168.1.300\D\orion.v5\"
    'globalpath = "d:\orion.v5"
    'xxempresa = "Sistema Orion V5"
    globalpath = App.path ' esta es la ruta principal
    'globalpath = "d:\orion.v5"
    'globalpath = "d:\ORION.V5" ' esta es la ruta alterna

    'carga_empresas
    cargar_grafico1
    cargar_grafico2
    'carga_tienda
    xtelefono = "" & Chr$(57) + "" & Chr$(57) + "" & Chr$(54) + "" & Chr$(50) + "" & Chr$(52) + "" & Chr$(54) + "" & Chr$(52) + "" & Chr$(55) + "" & Chr$(56)
    Denomina = "" & Chr$(10) + Chr$(13)
    Denomina = Denomina & "" & Chr$(10) + Chr$(13)

    ' 26/07/2018 Desactivar Facturacion Electronica
    Denomina = Denomina & nombre_sistema1 & Chr$(10) + Chr$(13)
    ' 26/07/2018 Desactivar Facturacion Electronica

    ''01/07/2017 Edici�n de Razon Social
    'Denomina = Denomina & "Av. Venezuela 1179  " & Chr$(10) + Chr$(13)
    'Denomina = Denomina & "Oficina 301 - Bre�a - Lima - Per�" & Chr$(10) + Chr$(13)
    'Denomina = Denomina & "Lima - Per�" & Chr$(10) + Chr$(13)
    ''01/07/2017 Edici�n de Razon Social

    'Denomina = Denomina & "  2018 - Derechos Reservados (R)" + Chr$(10) + Chr$(13)
    'Denomina = Denomina & "  Version Internacional " + Chr$(10) + Chr$(13)

    activar2
    Call CambiarCR
    carga_bases

    'carga_servidor
    'leer_camino
    'carga_disco_duro
    palabra_bienvenida
    cerrar_puertosmscomm
    'recordar_ventass
    'barra_herramienta 0
    'MsgBox objServ.ObtenerEstado("facturador-local"), vbInformation
    Exit Sub
cmd78321_err:
    MsgBox "Aviso en Load " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub barra_herramienta(sw As Integer)

    On Error GoTo cmd7812_err

    If sw = 0 Then
        rtn = FindWindow("Shell_traywnd", "") 'get the Window
        Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW) 'hide the Tasbar

    End If

    If sw = 1 Then
        rtn = FindWindow("Shell_traywnd", "") 'get the Window
        Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW) 'show the Taskbar

    End If

    Exit Sub
cmd7812_err:
    Exit Sub

End Sub

Sub activar2()
    'IM89ds.Visible = False
    'dji33.Visible = False
    jui34hote.Enabled = False
    dlotablas.Enabled = False
    tienda1.Enabled = False
    fdjvtas.Enabled = False
    comjk1.Enabled = False
    alki1.Enabled = False
    prod1.Enabled = False
    bando1.Enabled = False
    dlo1211.Enabled = False
    cont121.Enabled = False
    'i893432.Enabled = False
    jui34hote.Visible = False
    'i893432.Visible = False
    cont121.Visible = False
    dlo1211.Visible = False
    bando1.Visible = False
    prod1.Visible = False
    alki1.Visible = False
    comjk1.Visible = False
    fdjvtas.Visible = False
    tienda1.Visible = False
    dlotablas.Visible = False
    jui34hote.Visible = False

End Sub

Private Sub Form_Terminate()

    'MsgBox "Hola"
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo cmd333_err

    cerrar_usuario_global
    barra_herramienta 1
    'MsgBox "abc"
    cerrar_base
    Exit Sub
cmd333_err:
    Exit Sub

End Sub

Private Sub form7343_Click()

End Sub

Private Sub gempresa_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    clave.SetFocus

End Sub

'Private Sub glocal_KeyPress(KeyAscii As Integer)
'If KeyAscii <> 13 Then Exit Sub
'clave.SetFocus
'End Sub

Private Sub graty711_Click()

    FrmChart.acu = "C"
    FrmChart.Show 1

End Sub

Private Sub gtra5gra_Click()

End Sub

Private Sub gui231_Click()

End Sub

Private Sub gui8343_Click()

    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    repdocum.acu = "S"
    repdocum.Show 1

End Sub

Private Sub guia834312_Click()

End Sub

Private Sub guiremid1_Click()

End Sub

Private Sub gyiy6333_Click()

End Sub

Private Sub guiua43_Click()

End Sub

Private Sub hgeneri3_Click()

End Sub

Private Sub hiscli_Click()
    treevcli.Show 1

End Sub

Private Sub Hnhen7734_Click()
    opcion2 = "1"
    planilag.Show 1

End Sub

Private Sub huieny1_Click()

    cgusuario = "FACTURA"
    dgusuario = "_d" & gusuario
    fgusuario = "_f" & gusuario
    dgusuariog = "DETALLE"
    explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    explorap.fechaf = Format(Now, "dd/mm/yyyy")
    explorap.fechai = Format(Now, "dd/mm/yyyy")
    explorap.Caption = "Documentos Guia Remision Entrada"
    explorap.tipoclie = "C"
    explorap.acu = "S"
    explorap.Show 1

End Sub

Private Sub hyreg55_Click()

End Sub

Private Sub IM89ds_Click()
    treeipm.Show 1

End Sub

Private Sub Image1_Click()

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    mytablex.Open "select * from vendedor where codigo='" & Trim(gusuario) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If "" & mytablex.Fields("escajero") = "N" Then
            MsgBox "Usuario no permitido para hacer caja ", 48, "Aviso"
            mytablex.Close
            Exit Sub

        End If

    End If

    mytablex.Close
    mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If "" & mytablex.Fields("touch") = "1" Then
            menucaja.Label3 = "CAJERO"
            menucaja.acu = "C"
            menucaja.Label1.Visible = True
            menucaja.turno.Visible = True
            menucaja.Label5 = "CAJA"
            menucaja.tipoterminal = "TOUCH"
            mytablex.Close
            menucaja.Show 1
            GoTo amipasa

        End If

        If "" & mytablex.Fields("touch") = "2" Then
            menucaja.Label3 = "CAJERO"
            menucaja.acu = "C"
            menucaja.Label1.Visible = True
            menucaja.turno.Visible = True
            menucaja.Label5 = "CAJA"
            menucaja.tipoterminal = "TOUCH2"
            mytablex.Close
            menucaja.Show 1
            GoTo amipasa

        End If

        If "" & mytablex.Fields("touch") <> "2" And "" & mytablex.Fields("touch") <> "1" Then
            menucaja.Label3 = "CAJERO"
            menucaja.acu = "C"
            menucaja.Label1.Visible = True
            menucaja.turno.Visible = True
            menucaja.Label5 = "CAJA"
            menucaja.tipoterminal = "NORMAL"
            mytablex.Close
            menucaja.Show 1
            GoTo amipasa

        End If

    End If

amipasa:
    'mytablex.Close
    'cn.Close
    Command5_Click

End Sub

Private Sub Image10_Click()
    'tmovcheq.Show 1
    cgusuario = "cREQUISA"
    dgusuario = "_d" & gusuario
    fgusuario = "_f" & gusuario
    dgusuariog = "dREQUISA"
    explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    explorap.fechaf = Format(Now, "dd/mm/yyyy")
    explorap.Caption = "Requerimientos de Almacenes"
    explorap.acu = "Q"
    explorap.tipoclie = "V"
    explorap.Show 1

End Sub

Private Sub Image11_Click()
    tiplocal.Show 1
    Exit Sub
    tcongta.Show 1
    carga_servidor
    leer_camino

End Sub

Private Sub image12_Click()
    'gtra5gra_Click
    opcion2 = "10"   'analisis de ventas
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    'repdocum.Label18.Visible = False
    'repdocum.Combo1.Visible = False
    repdocum.vdetalle.Enabled = False
    repdocum.vfpago.Enabled = False
    repdocum.acu = "V"
    repdocum.Show 1

End Sub

Private Sub image14_Click()
    cgusuario = "cordenc"
    dgusuario = "_d" & gusuario
    fgusuario = "_f" & gusuario
    dgusuariog = "dordenc"
    explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    explorap.fechaf = Format(Now, "dd/mm/yyyy")
    explorap.Caption = "Documentos Orden de Compra"
    explorap.tipoclie = "P"
    explorap.acu = "R"
    explorap.Show 1

End Sub

Private Sub Image2_Click()
    'xprodet.producto = "ZZZZ" '
    xprodet.Show 1

    'tprodup.Show 1
End Sub

Private Sub Image22_Click()
    TRUCLINE.Show 1

End Sub

Private Sub Image3_Click()
    cgusuario = "ccotizav"
    dgusuario = "_d" & gusuario
    fgusuario = "_f" & gusuario
    dgusuariog = "dcotizav"
    explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    explorap.fechaf = Format(Now, "dd/mm/yyyy")
    explorap.Caption = "Documentos Cotizacion Ventas"
    explorap.tipoclie = "C"
    explorap.acu = "H"
    explorap.Show 1

End Sub

Private Sub Image4_Click()
    cgusuario = "FACTURA"
    dgusuario = "_d" & gusuario
    fgusuario = "_f" & gusuario
    dgusuariog = "DETALLE"
    explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    explorap.fechaf = Format(Now, "dd/mm/yyyy")
    explorap.Caption = "Documentos Facturacion Compras"
    explorap.tipoclie = "P"
    explorap.acu = "C"
    explorap.importacion = "COMERCIAL"
    explorap.Show 1

End Sub

Private Sub Image5_Click()
    cgusuario = "FACTURA"
    dgusuario = "_d" & gusuario
    fgusuario = "_f" & gusuario
    dgusuariog = "DETALLE"
    explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    explorap.fechaf = Format(Now, "dd/mm/yyyy")
    explorap.Caption = "Documentos Guia Remision Ventas"
    explorap.tipoclie = "C"
    explorap.acu = "T"
    explorap.Show 1

End Sub

Private Sub Image6_Click()
    cgusuario = "FACTURA"
    dgusuario = "_d" & gusuario
    fgusuario = "_f" & gusuario
    dgusuariog = "DETALLE"
    explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    explorap.fechaf = Format(Now, "dd/mm/yyyy")
    explorap.Caption = "Documentos Guia Remision Compra"
    explorap.tipoclie = "P"
    explorap.acu = "S"
    explorap.Show 1

End Sub

Private Sub Image7_Click()
    cgusuario = "ctraslad"
    'cgusuario = "FACTURA"
    dgusuario = "_d" & gusuario
    fgusuario = "_f" & gusuario
    dgusuariog = "dtraslad"
    'dgusuariog = "DETALLE"
    'inicio 07/09/2017 pll
    'explorap.fk4844.Visible = False
    'fin 07/09/2017 pll
    explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    explorap.fechaf = Format(Now, "dd/mm/yyyy")
    'inicio 05/09/2017 pll
    'explorap.fechai = Format(Now, "dd/mm/yyyy")
    explorap.Caption = "Traslado entre almacen de un mismo establecimiento"
    explorap.tipoclie = "V"
    explorap.acu = "Z"
    explorap.Show 1

End Sub

Private Sub Image8_Click()
    cgusuario = "FACTURA"
    dgusuario = "_d" & gusuario
    fgusuario = "_f" & gusuario
    dgusuariog = "DETALLE"

    explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")

    explorap.fechaf = Format(Now, "dd/mm/yyyy")
    explorap.Caption = "Documentos Facturacion Ventas"

    ' Testing Proyecto Facturacion Electronica 05/04/2018
    explorap.DetalleSunat.Visible = True
    explorap.DarBaja.Visible = True
    ' Testing Proyecto Facturacion Electronica 05/04/2018

    explorap.tipoclie = "C"
    explorap.acu = "V"
    explorap.importacion = "COMERCIAL"
    explorap.Show 1

End Sub

Private Sub ineret23_Click()

End Sub

Private Sub interfase23_Click()

    'genconta.Show 1
End Sub

Private Sub kaiu34_Click()

End Sub

Private Sub ki78341_Click()

End Sub

Private Sub kie64_Click()
    tletra.acu = "V"
    tletra.Show 1

End Sub

Private Sub j8845_Click()

End Sub

Private Sub jatr3ka_Click()

End Sub

Private Sub ki98822_Click()

End Sub

Private Sub kfi883433_Click()

End Sub

Private Sub kier23_Click()

End Sub

Private Sub kiet345_Click()

End Sub

Private Sub juer834_Click()

End Sub

Private Sub klinyer_Click()

End Sub

Private Sub jui34hote_Click()
    treevho.Show 1

End Sub

Private Sub Label12_Click()

    If Frame2.Visible = True Then
        clave.SetFocus
        Exit Sub

    End If

    recordar

End Sub

Private Sub Label13_Click()
    recordar_ventas

End Sub

Private Sub Label18_Click()
    recordar_ingresos

End Sub

Private Sub Label9_Click()

    'dlor23_Click
End Sub

Private Sub ldirer1_Click()

End Sub

Private Sub kai7822_Click()

End Sub

Private Sub kdora_Click()

End Sub

Private Sub ki93431_Click()

End Sub

Private Sub jk8893_Click()
    treevtab.Show 1

End Sub

Private Sub kicaju2_Click()

    'tcaja.Show 1
    tcajado.Show 1

End Sub

Private Sub kie6411_Click()

End Sub

Private Sub Label2_Click()

    Dim wmi As Object

    Dim buf As String

    Dim mos As Object

    Dim mo  As Object

    Exit Sub
    
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set mos = wmi.ExecQuery("Select * from Win32_Baseboard")
    
    buf = ""

    For Each mo In mos

        buf = buf & "Serial Number: " & mo.SerialNumber & vbcrlf
        buf = buf & "Manufacturer: " & mo.Manufacturer & vbcrlf
        buf = buf & "Product: " & mo.Product
        MsgBox buf
    Next

End Sub

Private Sub Label5_Click()
    treinpre.Show 1

End Sub

Private Sub ldossali4_Click()

    If MsgBox("Desea Salir del Sistema", vbOKCancel + vbQuestion + vbDefaultButton2, "Aviso") <> 1 Then Exit Sub
    barra_herramienta 1
    cerrar_base
    End

End Sub

Sub cerrar_base()

    On Error GoTo cmd8912_err

    cn.Close
    Exit Sub
cmd8912_err:
    Exit Sub

End Sub

Private Sub letra5x8_Click()
    'explreci.Caption = "INGRESO DINERO"
    'explreci.afecta = "L"
    'explreci.acu = "W"
    'explreci.Show 1

    'Dim found As Integer
    'found = copiar_recibos()
    'If found = 0 Then
    '   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
    '   End
    '   Exit Sub
    'End If
    'fgusuario = "_r" & gusuario
    'trecaja.Caption = "INGRESO DINERO"
    'trecaja.afecta = "L"
    'trecaja.acu = "W"
    'trecaja.Show 1

End Sub

Private Sub letra6222_Click()

End Sub

Private Sub letrax1_Click()

End Sub

Private Sub lei7741_Click()

End Sub

Private Sub letytra734_Click()

    REPLETRA.titulo = "Letras por Pagar"
    REPLETRA.acu = "C"
    REPLETRA.Show 1

End Sub

Private Sub ley7812_Click()

    'tletra.acu = "C"
    'tletra.Show 1

End Sub

Private Sub leypaj61_Click()

End Sub

Private Sub li8923212_Click()

End Sub

Private Sub loacy734act_Click()

End Sub

Private Sub limnbro4_Click()

End Sub

Private Sub lincaj2_Click()

End Sub

Private Sub libj734_Click()

    'tlbrodia.Show 1
End Sub

Private Sub ljuwe_Click()

End Sub

Private Sub lmu7343_Click()

    'rptlibma.Show 1
End Sub

Private Sub lo3dua4_Click()

End Sub

Private Sub lo9934311_Click()

    If Forms.count > 1 Then
        Forms(1).SetFocus
        Exit Sub

    End If

    cgusuario = "CREQUISA"
    dgusuariog = "DREQUISA"
    repdocum.acu = "Q"
    repdocum.Show 1

End Sub

Private Sub lo8911_Click()
    'texplcxc.acu = "V"
    'texplcxc.Show 1

End Sub

Private Sub lo89114_Click()

End Sub

Private Sub loa911_Click()

End Sub

Private Sub loe231_Click()

    'tabbanco.Show 1
End Sub

Private Sub loe32pro91_Click()

End Sub

Private Sub loepgare_Click()

End Sub

Private Sub loet12_Click()

End Sub

Private Sub lp034_Click()

End Sub

Private Sub lolin55_Click()

End Sub

Private Sub loproces0001_Click()

End Sub

Private Sub loropropar1_Click()

End Sub

Private Sub loteru1_Click()

End Sub

Private Sub loti8845_Click()

End Sub

Private Sub lpro8523_Click()

End Sub

Private Sub lro343_Click()

End Sub

Private Sub lrovta52_Click()

End Sub

Private Sub many343_Click()

End Sub

Private Sub manyeye11_Click()

End Sub

Private Sub lropro92_Click()

End Sub

Private Sub lropro923_Click()

End Sub

Private Sub lro343proc_Click()

End Sub

Private Sub maori34_Click()

End Sub

Private Sub lo9034_Click()
    treevcon.Show 1

End Sub

Private Sub menceju34_Click()
    'If Forms.Count > 1 Then
    '   Forms(1).SetFocus
    '   Exit Sub
    'End If
    frmsms.Show 1

End Sub

Private Sub mifo934_Click()
    trecitot.Caption = "COMPROBANTES INGRESOS EGRESOS"
    'explreci.acu = "W"
    trecitot.Show 1
    Exit Sub

End Sub

Private Sub min454_Click()

    'proseccio.Show 1
End Sub

Private Sub mkl991_Click()

    tpersona.Show 1

End Sub

Private Sub mo88445_Click()
    tplamo.Show 1

    'planilla.Show 1
End Sub

Private Sub ninery23_Click()

End Sub

Private Sub o973434_Click()

    cgusuario = "CORDENC"
    dgusuariog = "DORDENC"
    repdocum.acu = "R"
    repdocum.Show 1

End Sub

Private Sub nomsal1_Click()

    cgusuario = "FACTURA"
    dgusuario = "_d" & gusuario
    fgusuario = "_f" & gusuario
    dgusuariog = "DETALLE"
    explorap.fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    explorap.fechaf = Format(Now, "dd/mm/yyyy")
    explorap.Caption = "Documentos Guia Remision Salida"
    explorap.tipoclie = "P"
    explorap.acu = "T"
    explorap.Show 1

End Sub

Private Sub notacre1_Click()

End Sub

Private Sub notadeg1_Click()

End Sub

Private Sub notcre_Click()

End Sub

Private Sub notdebu_Click()

End Sub

Private Sub oiorde82_Click()

    cgusuario = "CORDENC"
    dgusuariog = "DORDENC"
    repdocum.acu = "R"
    repdocum.Show 1

End Sub

Private Sub old733_Click()
    tcambio.Show 1

End Sub

Private Sub opmi882_Click()

    toperaco.Show 1

End Sub

Private Sub oriki723_Click()

    'torigen.Show 1
End Sub

Private Sub pahy62321_Click()

End Sub

Private Sub orlo3422_Click()

    repprodu.titulo = "Reportes de Produccion"
    repprodu.Show 1

End Sub

Private Sub pak83434_Click()

    repctaxc.acu = "C"
    repctaxc.Show 1

End Sub

Private Sub pedoc232_Click()

End Sub

Private Sub paremuer1_Click()

End Sub

Private Sub patimer_Click()

    'tpacaja.Show 1
End Sub

Private Sub pedore11_Click()

End Sub

Private Sub peo8al1_Click()

    cgusuario = "CREQUISA"
    dgusuariog = "DREQUISA"
    repdocum.acu = "Q"
    repdocum.Show 1

End Sub

Private Sub pero8855_Click()

End Sub

Private Sub pero912_Click()

    If busca_clave1(gusuario) <> "S" Then
        MsgBox "No tiene Permiso", 48, "Aviso"
        Exit Sub

    End If

    tpersona.Show 1

    'tpersona.show 1
End Sub

Function busca_clave1(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_clave1 = Trim("" & mytablex.Fields("vevend"))

    End If

    mytablex.Close

End Function

Private Sub pero99404_Click()

    'tpartes.Show 1
End Sub

Private Sub plu8234_Click()
    tctable.Show 1

End Sub

Private Sub pro534_Click()

End Sub

Private Sub proce896666_Click()

End Sub

Private Sub proer_Click()

End Sub

Private Sub prokiwe1_Click()

End Sub

Private Sub prolo834_Click()

End Sub

Private Sub pro8923_Click()
    
End Sub

Private Sub promi773_Click()

End Sub

Private Sub recajasd_Click()

    'Dim found As Integer
    'found = copiar_recibos()
    'If found = 0 Then
    '   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
    '   End
    '   Exit Sub
    'End If
    'fgusuario = "_r" & gusuario
    'trecaja.Caption = "INGRESO DINERO"
    'trecaja.afecta = "L"
    'trecaja.acu = "W"
    'trecaja.Show 1

End Sub

Private Sub recov1_Click()

    'Dim found As Integer
    'found = copiar_recibos()
    'If found = 0 Then
    '   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
    '   End
    '   Exit Sub
    'End If
    'fgusuario = "_r" & gusuario

    'trecaja.Caption = "INGRESO DINERO"
    'trecaja.afecta = "C"
    'trecaja.acu = "W"
    'trecaja.Show 1

End Sub

Private Sub pro9022_Click()
    treevpro.Show 1

End Sub

Private Sub regy773_Click()
    regsiste.xlicencia = "LICENCIA"
    regsiste.Show 1

End Sub

Private Sub rehunt75_Click()

    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    repdocrv.titulo = "REGISTRO DE COMPRAS " & dicmoneda
    repdocrv.acu = "C"
    repdocrv.Show 1

End Sub

Private Sub reki12_Click()

End Sub

Private Sub rekicaje_Click()

    'Dim found As Integer
    'found = copiar_recibos()
    'If found = 0 Then
    '   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
    '   End
    '   Exit Sub
    'End If
    'fgusuario = "_r" & gusuario

    'trecaja.Caption = "EGRESO DINERO"
    'trecaja.afecta = "P"
    'trecaja.acu = "V"
    'trecaja.Show 1
End Sub

Private Sub relonova_Click()

End Sub

Private Sub relou734_Click()

End Sub

Private Sub reo9454_Click()

End Sub

Private Sub renfg34_Click()

    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    repdocrv.titulo = "REGISTRO DE VENTAS " & dicmoneda
    repdocrv.acu = "V"
    repdocrv.Show 1

End Sub

Private Sub rewnom12_Click()

End Sub

Private Sub rh88341_Click()
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    repdocrv.titulo = "REGISTRO DE COMPRAS " & dicmoneda
    repdocrv.acu = "C"
    repdocrv.Show 1

End Sub

Private Sub rimeti8833_Click()

End Sub

Private Sub rki3434_Click()

    txpoclie.Show 1

End Sub

Private Sub s90111_Click()

End Sub

Private Sub salo993in_Click()

End Sub

Private Sub saloaty4_Click()

End Sub

Private Sub saloiin911_Click()

End Sub

Private Sub saloinid_Click()

End Sub

Private Sub saloinidwer_Click()

End Sub

Private Sub sdhyeer_Click()

End Sub

Private Sub segui4_Click()

End Sub

Private Sub sdju7733_Click()

End Sub

Private Sub segu8344_Click()

End Sub

Private Sub sejcu1_Click()

    tseccion.Show 1

End Sub

Private Sub ski23_Click()

    ttranspo.Show 1

End Sub

Private Sub smatrq_Click()

    tnmarca.Show 1

End Sub

Private Sub srhyere_Click()

End Sub

Private Sub tabe1em_Click()

End Sub

Private Sub t55_Click()

End Sub

Private Sub tacj871_Click()

End Sub

Private Sub tio9343_Click()

    'cotipodo.Show 1
End Sub

Private Sub tupodo2_Click()

End Sub

Private Sub vcki34qw_Click()

End Sub

Private Sub univt5612_Click()
    
End Sub

Private Sub tr6665_Click()

End Sub

Private Sub li8931utyi_Click()
    treevuti.Show 1

End Sub

Private Sub Picture1_Click()

    Dim objs

    Dim OBJ

    Dim wmi

    Dim strMBD

    Set wmi = GetObject("WinMgmts:")
    Set objs = wmi.InstancesOf("Win32_BaseBoard")

    For Each OBJ In objs

        strMBD = "MotherBoard Number:  " & OBJ.SerialNumber
    Next
    MsgBox strMBD

End Sub

Private Sub MSComm1_OnComm()

    Dim vr

    'Dim tiempo
    ' declaramos una variable donde quedaran los datos recibidos
    Static strData As String

    Select Case MSComm1.CommEvent

        Case comEventBreak
            MsgBox "Error", "1comEventBreak"

        Case comEventFrame
            MsgBox "Error", "1comEventFrame"

        Case comEventOverrun
            MsgBox "Error", "1comEventOverrun"

        Case comEventRxOver
            MsgBox "Error", "1comEventRxOver"

        Case comEventRxParity
            MsgBox "Error", "1comEventRxParity"

        Case comEventTxFull
            MsgBox "Error", "1comEventTxFull"

        Case comEventDCB
            MsgBox "Error", "1comEventDCB"

        Case comEvSend
            'vr = DoEvents()
            'vr = DoEvents()
            vr = DoEvents()

            'genver.Caption = genver.Caption & MSComm1.Input
            'MsgBox "Transmitiendo"
            'Sleep (0.2)
        Case comEvReceive

            'strData = strData & MSComm1.Input
            'MsgBox strData
            'strData = ""
        Case Else: 'MsgBox MSComm1.CommEvent
            'Sleep (1)
            'GoTo iniciamos
            MsgBox "Presione enter para Continuar..Comm Event " & MSComm1.CommEvent, 48, "Aviso"
    
    End Select

End Sub



Private Sub tienda1_Click()
    treevpt.Show 1

End Sub

Private Sub u8933_Click()
    treevpla.Show 1

End Sub

Private Sub vki6343_Click()

    'tasiento.Show 1
End Sub

Private Sub w444_Click()

    'trepvta.Show 1
End Sub

Private Sub vki8494_Click()

End Sub

Private Sub txtText1_Change()

End Sub

Private Sub txtText1_KeyPress(KeyAscii As Integer)

    Dim found   As Integer

    Dim xcampo1 As String

    Dim xcampo2 As String

    Dim xcampo3 As String

    Dim salida  As Boolean
 
    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Exit Sub

    End If
 
    'If IsValidIPAddress(vservidor) = False Then
    '   MsgBox "Ingrese Ip Valida"
    '   clave.SetFocus
    '   Exit Sub
    'End If
    clave = UCase(clave)

    If Len(clave) = 0 Then
        clave.SetFocus
        Exit Sub

    End If

    '----------------------
    cerrar_base
    glocal = "01"

    If Len(Trim("" & gempresa)) = 0 Then
        clave = ""
        clave.SetFocus
        Exit Sub

    End If

    'MsgBox "abc"

    'If Len(vservidor) = 0 Then
    '   MsgBox "No existe Vendedor ", 48, "Aviso"
    '   clave.SetFocus
    '   Exit Sub
    'End If
    menup.Caption = nombre_sistema & " Empresa:" & gempresa
    found = extraer_campos1(gempresa, xcampo1, xcampo2, xcampo3)
    valida_conec xcampo2

    If Len(Trim(vservidor)) = 0 Then
        clave = ""
        clave.SetFocus
        Exit Sub

    End If

    'If Len(Trim(clave_servidor)) = 0 Then
    '   clave = ""
    '   clave.SetFocus
    '   Exit Sub
    'End If
    'MsgBox "abc"
    'MsgBox extra_loquesea1(gempresa)
    'MsgBox xcampo2
    found = conectar(xcampo2)

    If found = 0 Then
        MsgBox "Error de Conexion Sql Server ", 48, "Aviso"
        clave.SetFocus
        Exit Sub

    End If

    'found = conectara()
    'If found = 0 Then
    '   MsgBox "Error de Conexion Sql Server ", 48, "Aviso"
    '   clave.SetFocus
    '   Exit Sub
    'End If
    'SETEAR GLOGALDIR- SETEO
    globaldir = globalpath & "\001d\06"
    globaldat = globalpath & "\001d\06"
    globalcont = globalpath & "\001d\contable"
    globalpri = globalpath & "\001d"
    globalweb = globalpath & "\001d\06\web"
    orionv4 = "\orion.v4\001d\01"
    'MsgBox gempresa
    empresapos = "01"
    'found = busca_empresa(extra_loquesea(gempresa))
    'If found = 0 Then
    '   MsgBox "Empresa No existe", 48, "Aviso"
    '   clave = ""
    '   clave.SetFocus
    '   Exit Sub
    'End If
    globalemp = extra_loquesea(gempresa)

    'Label1 = Label9
    found = busca_clave("" & clave)

    If found = 0 Then
        MsgBox "No existe Clave asignado a ningun Funcionario", 48, "Aviso"
        clave = ""
        clave.SetFocus
        Exit Sub

    End If

    ' 26/07/2018 Desactivar Facturacion Electronica
    V_EstadoSistema = Obtiene_EstadoSistema()
    ' 26/07/2018 Desactivar Facturacion Electronica

    'Testing Proyecto Facturacion Electronica 05/04/2018
    If V_EstadoSistema = "FE BYH" Then
        Servicio.Visible = True
        Servicio2.Visible = True
        Set objServ = New servicios
    
        rptaServicio = objServ.ObtenerEstado("facturador-local")

        If rptaServicio = "El servicio est� detenido" Then
            Servicio = "Servicio Detenido"
            Servicio.BackColor = &HFF&
        ElseIf rptaServicio = "El servicio est� activo" Then
            Servicio = "Servicio Activo"
            Servicio.BackColor = &HFF8080

        End If
    
        If Servicio = "Servicio Detenido" Then
            Shell ("C:\BYH\iniciar.bat"), vbNormalFocus
            Servicio = "Servicio Activo"
            Servicio.BackColor = &HFF8080

        End If

    ElseIf V_EstadoSistema = "CONINT" Then
        Servicio.Visible = False
        Servicio2.Visible = False

    End If

    ' 26/07/2018 Desactivar Facturacion Electronica

    'MsgBox "abc"
    'found = busca_local()
    'If found = 0 Then
    If Len(Trim(glocal)) = 0 Then
        glocal = "01"

    End If

    If verificador_datos("" & gusuario) = "S" Then
        Frame2.Visible = False
        tncr1.Show 1
        barra_herramienta 1
        cerrar_base
        End
        Exit Sub

    End If

    'MsgBox "abc"

    'Set mydbzglo = OpenDatabase(globalcont, False, False, "foxpro 2.5;")
    'MsgBox "xxxx"
    Set mydbxglo = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
    'MsgBox "xxx"
    found = cargas_iniciales()

    'MsgBox "xxxx"
    If found = 0 Then

    End If

    'found = verifica_licenciaremoto()
    '--------aqui verifica que no ingrese otro con el mismo usuario
    found = copiar_temporalxxx()

    If found = 0 Then

        'MsgBox "Usuario ya Activo ", 48, "Aviso"
        'clave = ""
        'clave.SetFocus
        'Exit Sub
    End If

    'MsgBox "abc"
    found = abrir_usuario_global()

    If found = 0 Then
        MsgBox "Error al abrir control ", 48, "Aviso"
        clave = ""
        clave.SetFocus
        Exit Sub

    End If

    'fdk9833.Visible = True
    palabra_bienvenida1
    globalocal = glocal
    'recordar
    'menup.Caption = "SISTEMA ORION V5.0 " '& Label1
    Frame2.Visible = False
    recordar
    'visualizar_recordar
    'visualizar_ingresos
    visualiza_mesa
    'kenyo 20/04/2017
    'opcion2 = "4"
    'repinv.excell.Visible = True
    'repinv.Label17.Visible = True
    'repinv.Combo1.Visible = True
    'repinv.Label25.Visible = True
    'repinv.gcanti.Visible = True
    'repinv.excell.Value = 1
    'repinv.Combo1.Text = "TODOS"
    'repinv.CargaInicial
    'repinv.Show 1
    'repinv.Hide
    'inicio 28/11/2017 pll para las configuracion monedas
    Call configura_moneda(my_vdolar, salida)

    If salida = False Then
        MsgBox "Ingresar el dato en Tablas -->Parametros Generales/ve Dolar "

    End If

    'obtenemos estado del lfacturador sunat 02/02/2018
    'MsgBox objServ.ObtenerEstado("facturador-local") & " - Facturador Electronico Sunat", vbInformation
  
    'Testing Proyecto Facturacion Electronica 05/04/2018
    '        Dim rpta As String
    '
    '        Call valida_facturacionElectronica
    '        If rpta = "" Then
    '            respuesta.Text = "TODO OK"
    '        Else
    '            rptaa.Visible = True
    '            respuesta.Text = rpta
    '        End If
    'Testing Proyecto Facturacion Electronica 05/04/2018
  
    'fi 28/11/2017 pll

End Sub

Private Sub vservidor_Click()
    tiplocal.Show 1
    Exit Sub
    tcongta.Show 1
    carga_servidor
    leer_camino

End Sub

Private Sub xj7811_Click()

    opcion2 = "11"   'analisis de ventas
    cgusuario = "FACTURA"
    dgusuariog = "DETALLE"
    'repdocum.Label18.Visible = False
    'repdocum.Combo1.Visible = False
    repdocum.vdetalle.Enabled = False
    repdocum.vfpago.Enabled = False
    repdocum.acu = "C"
    repdocum.Show 1

End Sub

Private Sub xliki23_Click()

    tlinea.Show 1

End Sub

Private Sub xlo8923_Click()

End Sub

Function cargas_iniciales()

    Dim ybuf  As String

    Dim found As Integer

    On Error GoTo cmd169999_err

    'Exit Function
    'If CDate(Format(Now, "dd/mm/yyyy")) > CDate("30/09/2013") Then
    'MsgBox "xx"
    '   found = numero_registro()
    '   Exit Function
    '
    'End If
    'Exit Function

    ybuf = "" & menup.vservidor

    If Trim(UCase$("" & menup.vservidor)) <> "(LOCAL)" Then
        If IsValidIPAddress(menup.vservidor) = False Then

            'MsgBox "Ip no Esta bien Definido ", 48, "Aviso"
            'Exit Function
        End If

    Else  'SI ES LOCAL
        ybuf = GetMACs_AdaptInfo()  'si es local

    End If
    
    'If Mid$(ybuf, 1, 3) <> "169" And Mid$(ybuf, 1, 3) <> "127" And Mid$(ybuf, 1, 3) <> "192" Then
    ' MsgBox "Es externo"
    'End If
    
    If IsValidIPAddress(ybuf) = True Then
        If verifica_mac(ybuf, "LICENCIA") = 0 Then
            found = numero_registro()

        End If

        Exit Function

    End If

    If verifica_disco_duro("LICENCIA") = 0 Then
        found = numero_registro()

    End If

    Exit Function
cmd169999_err:
    MsgBox "No se cargas Iniciales " + error$, 48, "Aviso"
    Close
    Exit Function

End Function

Function verifica_mac(ybuf As String, xlicencia As String)

    Dim discoduro As String

    Dim sdx       As Integer

    Dim buf       As String

    Dim ED        As tcrypto

    Set ED = New tcrypto

    Dim mytablex As New ADODB.Recordset

    On Error GoTo rutina_error13

    licencia_remoto = ""
    sdx = 0
    mytablex.Open "SELECT * FROM  " & xlicencia, cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then 'si no existe
        mytablex.Close
        verifica_mac = 0
        Exit Function

    End If

    'MsgBox "xxxx"
    buf = Trim("" & mytablex.Fields("serie"))
    discoduro = ED.Encrypt(buf, "FABIANITA", True)
    'MsgBox discoduro & " 68611015"
    buf = ""
    buf = Trim(serie_mac(ybuf))

    'MsgBox buf
    'MsgBox "Mac:" & buf & " Disco Duro:" & Mid$(discoduro, 1, Len(buf))
    'MsgBox discoduro
    If buf = Mid$(discoduro, 1, Len(buf)) And Len(buf) > 0 Then

        'MsgBox ""
        If xlicencia = "LICENCIA" Then
            xxempresa = "ESTA LICENCIA PERTENECE A:" & ED.Encrypt(Trim("" & mytablex.Fields("nombre")), "FABIANITA", True)

        End If
       
        sdx = 1

    End If

    If sdx = 1 Then
        If xlicencia = "LICENCIACENTRALIZADO" Then
            licencia_remoto = "S"

        End If

    End If

    verifica_mac = sdx
    Exit Function
rutina_error13:
    verifica_mac = 0
    MsgBox "Error al leer archivo CAM " + error$, 24, "Aviso"
    Exit Function

    'serie_mac
End Function

Function verifica_disco_duro(xlicencia As String)

    Dim discoduro As String

    Dim DATO      As String

    Dim sdx       As Integer

    Dim buf       As String

    Dim ED        As tcrypto

    Set ED = New tcrypto

    Dim mytablex As New ADODB.Recordset

    On Error GoTo rutina_error1

    sdx = 0

    'If Dir$(globalpath & "\SERIE.TXT") <> "" Then
    '   Open globalpath & "\SERIE.TXT" For Input As #1
    '   Input #1, discoduro
    '   Close #1
    '   buf = serie_disco_duro()
   
    '---------------lo nuevo-------------
    licencia_remoto = ""
    mytablex.Open "SELECT * FROM  " & xlicencia, cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then 'si no existe
        mytablex.Close
        verifica_disco_duro = 0
        Exit Function

    End If

    DATO = Trim("" & mytablex.Fields("serie"))
    discoduro = ED.Encrypt(DATO, "FABIANITA", True)
    buf = Trim("" & mytablex.Fields("nombre"))

    If xlicencia = "LICENCIA" Then
        xxempresa = "ESTA LICENCIA PERTENECE A:" & ED.Encrypt(buf, "FABIANITA", True)

    End If
   
    'MsgBox discoduro
    mytablex.Close
    buf = serie_disco_duro()

    If buf = discoduro And (Len(buf) > 0 And Len(discoduro) > 0) Then
        sdx = 1

    End If

    If sdx = 1 Then
        If xlicencia = "LICENCIACENTRALIZADO" Then
            licencia_remoto = "S"

        End If

    End If

    verifica_disco_duro = sdx
    Exit Function
rutina_error1:
    verifica_disco_duro = 0
    'MsgBox "Error al leer archivo " + error$, 24, "Aviso"
    Exit Function

End Function

Function numero_registro() As Integer

    Dim sdx

    Dim texi As New ADODB.Recordset

    texi.Open "SELECT local FROM factura where tipo LIKE '%' ", cn, adOpenKeyset, adLockOptimistic

    If texi.RecordCount = 0 Then  'si existe
        texi.Close
        Exit Function

    End If

    sdx = 0

    'MsgBox texi.RecordCount
    If texi.RecordCount > 30 Then
        sdx = 50 - texi.RecordCount
        MsgBox "Licencia Demo Finalizado : " & "Le Quedan  " & sdx & " Transacciones Para su Desactivacion " & Chr$(10) & Chr$(13) ' & "Comunicarse Urgente con su proveedor o " & Chr$(10) & Chr$(13) & "Llamar al 4751670/989313650"

        If sdx < 1 Then
            MsgBox "Licencia Demo Finalizado-Desactivado " & Chr$(10) & Chr$(13) & "Comunicarse con su proveedor o " & Chr$(10) & Chr$(13) '& " Llamar al 989313651/4751670 wwww.kalipos.com "
            texi.Close
            regsiste.xlicencia = "LICENCIA"
            regsiste.Show 1 '1
            End
            Exit Function

        End If

    End If

    numero_registro = 1
    texi.Close

End Function

Sub ir_ultimo(mytablex As Table)

    On Error GoTo cmd5678_err

    mytablex.MoveLast
    Exit Sub
cmd5678_err:
    Exit Sub
    Exit Sub

End Sub

Sub cargar_grafico1()

    On Error GoTo cmd7779_err

    Image9.Picture = LoadPicture(globalpath & "\ico\leon.jpg")
    'danielkheavy
    'Image11.Picture = LoadPicture(globalpath & "\ico\tpv.jpg")
    'danielkheavy
    Exit Sub
cmd7779_err:
    MsgBox " Carga Grafico:" & error$
    Exit Sub

End Sub

Sub cargar_grafico2()

    On Error GoTo cmd77790_err

    image1.Picture = LoadPicture(globalpath & "\ico\icono.jpg")
    Exit Sub
cmd77790_err:

    'MsgBox " Carga Grafico:" & error$
End Sub

Public Function IsLoadForm(ByVal FormCaption As String, _
                           Optional active As Variant) As Boolean

    Dim rtn As Integer, I As Integer

    Dim Name

    rtn = False
    I = 1
    Name = LCase(FormCaption)

    Do Until I > Forms.count - 1 Or rtn

        If LCase(Forms(I).Caption) = Name Then rtn = True
        I = I + 1
    Loop

    If rtn Then
        If Not IsMissing(active) Then
            If active Then
                Forms(I - 1).WindowState = vbNormal

            End If

        End If

    End If

    IsLoadForm = rtn

End Function

Function ver_abiertos_form() As Integer
    ver_abiertos_form = 0

    'MsgBox Forms.Count
    If Forms.count > 1 Then
        ver_abiertos_form = 1
        Forms(2).WindowState = vbNormal

    End If

End Function

Sub carga_disco_duro()

    Dim found As Integer

    On Error GoTo cmd69999_err

    serial_number = ""

    If Dir$("c:\xyz.txt") <> "" Then
        Close
        Open "C:\XYZ.TXT" For Input As #1
        Input #1, serial_number
        Close #1

    End If

    '------------------------------------------------
    Exit Sub
cmd69999_err:
    Close
    'MsgBox "Error en activaciones " & Error$, 1, "Aviso"
    Exit Sub

End Sub

Sub pasar_claves()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor ", cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Sub
        'mytablex.Edit
        mytablex.Fields("v1") = Mid$("" & mytablex.Fields("visibles"), 1, 1)
        mytablex.Fields("v2") = Mid$("" & mytablex.Fields("visibles"), 2, 1)
        mytablex.Fields("v3") = Mid$("" & mytablex.Fields("visibles"), 3, 1)
        mytablex.Fields("v4") = Mid$("" & mytablex.Fields("visibles"), 4, 1)
        mytablex.Fields("v5") = Mid$("" & mytablex.Fields("visibles"), 5, 1)
        mytablex.Fields("v6") = Mid$("" & mytablex.Fields("visibles"), 6, 1)
        mytablex.Fields("v7") = Mid$("" & mytablex.Fields("visibles"), 7, 1)
        mytablex.Fields("v8") = Mid$("" & mytablex.Fields("visibles"), 8, 1)
        mytablex.Fields("v9") = Mid$("" & mytablex.Fields("visibles"), 9, 1)
        mytablex.Fields("v10") = Mid$("" & mytablex.Fields("visibles"), 10, 1)
        mytablex.Fields("v11") = Mid$("" & mytablex.Fields("visibles"), 11, 1)

        mytablex.Fields("rw1") = Mid$("" & mytablex.Fields("permiso"), 1, 1)
        mytablex.Fields("rw2") = Mid$("" & mytablex.Fields("permiso"), 2, 1)
        mytablex.Fields("rw3") = Mid$("" & mytablex.Fields("permiso"), 3, 1)
        mytablex.Fields("rw4") = Mid$("" & mytablex.Fields("permiso"), 4, 1)
        mytablex.Fields("rw5") = Mid$("" & mytablex.Fields("permiso"), 5, 1)
        mytablex.Fields("rw6") = Mid$("" & mytablex.Fields("permiso"), 6, 1)
        mytablex.Fields("rw7") = Mid$("" & mytablex.Fields("permiso"), 7, 1)
        mytablex.Fields("rw8") = Mid$("" & mytablex.Fields("permiso"), 8, 1)
        mytablex.Fields("rw9") = Mid$("" & mytablex.Fields("permiso"), 9, 1)
        mytablex.Fields("rw10") = Mid$("" & mytablex.Fields("permiso"), 10, 1)
        mytablex.Fields("rw11") = Mid$("" & mytablex.Fields("permiso"), 11, 1)
        mytablex.Update
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Sub carga_servidor()

    Dim found As Integer

    Dim buf   As String

    On Error GoTo cmd169999_err

    Exit Sub
    vservidor = ""
    buf = ""

    If Dir$(globalpath & "\server.txt") <> "" Then
        Close
        Open globalpath & "\server.txt" For Input As #1
        Input #1, buf
        Close #1
        vservidor = buf
       
    End If

    '------------------------------------------------
    Exit Sub
cmd169999_err:
    Close
    Exit Sub

End Sub

Sub leer_camino()

    Dim found As Integer

    Dim buf   As String

    On Error GoTo cmd00169999_err

    Exit Sub
    clave_servidor = ""
    buf = ""

    If Dir$(globalpath & "\camino.txt") <> "" Then
        Close
        Open globalpath & "\camino.txt" For Input As #1
        Input #1, buf
        Close #1
        clave_servidor = buf

    End If

    '------------------------------------------------
    Exit Sub
cmd00169999_err:
    Close
    Exit Sub

End Sub

Sub CERRAR_VARGLO()

End Sub

Function verificador_datos(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM vendedor where codigo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        verificador_datos = "" & mytablex.Fields("verificador")

    End If

    mytablex.Close

End Function

Private Sub y6663_Click()

End Sub

Sub cuadrar_igv()

    'cn.Execute ("update factura set subtotal=total/1.18")
    'cn.Execute ("update factura set impuesto=total-subtotal")
    'cn.Execute ("update detalle set subtotal=total/1.18")
    'cn.Execute ("update detalle set impuesto=total-subtotal")
End Sub

Sub carga_config()

    Dim found As Integer

    Dim buf   As String

    Dim ind   As Integer

    On Error GoTo cmd699992_err

    Exit Sub
    gempresa.Clear
    ind = 0

    If Dir$(globalpath & "\config") <> "" Then
        Close
        '------------------------------------
        Open globalpath & "\config" For Input As #2
        Do

            If EOF(2) Then Exit Do
            Line Input #2, buf
            gempresa.AddItem buf
            ind = 1
        Loop
        Close

        '------------------------------------
    End If

    If ind = 0 Then
        gempresa.Clear
        gempresa.AddItem "01|CALIPSO"

    End If

    gempresa.ListIndex = 0
    '------------------------------------------------
    Exit Sub
cmd699992_err:
    Close
    Exit Sub

End Sub

Sub hacer_sunat()

    Dim buf      As String

    Dim buf1     As String

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim sdx      As Double

    Dim sdx1     As Double

    cn.Execute ("delete from sisunat")
    'Exit Sub

    'vemos el saldo inicial del almacen
    'mytablex.Open "Select * from saldoini where local='01' and bodega='01' and  producto='" & "" & rrproducto.Fields("producto") & "' and bodega='" & extra_loquesea(subodega) & "' and fecha='" & "01/" & sumes & "/" & suano & "'", cn, adOpenStatic, adLockOptimistic
    '   If mytablex.RecordCount > 0 Then
    '   xcantet = Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
    '   End If
    '--------------------------
    mytablex.Close

    buf1 = "SELECT * FROM detalle where year(fecha)=2012 and "
    buf1 = buf1 & " (acu='A' or acu='B' or acu='C' or acu='D' or acu='J' OR acu='K' OR acu='L' OR acu='M' or acu='G' OR ACU='T' OR ACU='P' OR ACU='S') "
    mytablex.Open buf1, cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do

        If mytabley.State = 1 Then
            mytabley.Close
            Set mytabley = Nothing

        End If

        buf = "select * from sisunat where "
        buf = buf & " local='" & "" & mytablex.Fields("local") & "'"
        buf = buf & " and bodega='" & "" & mytablex.Fields("bodega") & "'"
        buf = buf & " and producto='" & "" & mytablex.Fields("producto") & "'"
        buf = buf & " and mes='" & Format(Month(mytablex.Fields("fecha")), "00") & "'"
        buf = buf & " and anno='" & Format(Year(mytablex.Fields("fecha")), "0000") & "'"

        mytabley.Open buf, cn, adOpenKeyset, adLockOptimistic

        If mytabley.RecordCount = 0 Then
            mytabley.AddNew
            acumular_sunat mytablex, mytabley
            mytabley.Update
        Else
            acumular_sunat mytablex, mytabley
            mytabley.Update

        End If

        mytabley.Close
        mytablex.MoveNext
    Loop
    mytablex.Close

    'MsgBox "Acabe"
End Sub

Sub acumular_sunat(mytablex As ADODB.Recordset, mytabley As ADODB.Recordset)

    Dim sdx As Double

    sdx = 0
    mytabley.Fields("local") = "" & mytablex.Fields("local")
    mytabley.Fields("bodega") = "" & mytablex.Fields("bodega")
    mytabley.Fields("producto") = "" & mytablex.Fields("producto")
    mytabley.Fields("mes") = "" & Format(Month(mytablex.Fields("fecha")), "00")
    mytabley.Fields("anno") = "" & Format(Year(mytablex.Fields("fecha")), "0000")
   
    Select Case "" & mytablex.Fields("acu")

        Case "A", "B", "C", "D", "G", "T"
            sdx = Val("" & mytabley.Fields("cantidad")) - Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))

        Case "J", "K", "L", "M", "P", "S"
            sdx = Val("" & mytabley.Fields("cantidad")) + Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
            ultimo_costo = Val("" & mytablex.Fields("precio"))

    End Select

    mytabley.Fields("costo") = ultimo_costo
    mytabley.Fields("cantidad") = sdx

End Sub

Sub carga_tienda()

    On Error GoTo cmd127779_err

    image1.Picture = LoadPicture(globalpath & "\ico\tienda.jpg")
    Exit Sub
cmd127779_err:
    MsgBox "Carga Tienda:" & error$
    Exit Sub

End Sub

Function leer_visor()

    Dim found    As Integer

    Dim bdvisor1 As bdvisor

    On Error GoTo cmd7824_err

    Dim buf As String

    Open globalpath & "\visor.txt" For Random As #4 Len = Len(bdvisor1)
    Get #4, 1, bdvisor1
    found = envio_visor(bdvisor1.ppuerto, bdvisor1.vvelocidad, bdvisor1.mmensaje1, bdvisor1.mmensaje2)
    leer_visor = 1
    Close #4
    Exit Function
cmd7824_err:
    MsgBox "Aviso en Leer visor " + error$, 48, "Aviso"
    Exit Function

End Function

Sub carga_bases()

    Dim found As Integer

    Dim bdip  As ipmaquina

    On Error GoTo cmd78241_err

    Dim sdx1 As Double

    Dim buf  As String

    Dim sdx  As Integer

    Dim I    As Integer

    Dim sdx2 As Integer

    sdx1 = 0
    sdx2 = -1
    gempresa.Clear
    Open globalpath & "\config.txt" For Random As #4 Len = Len(bdip)
    sdx = (LOF(4) \ Len(bdip)) '

    For I = 1 To sdx
        Get #4, I, bdip
        gempresa.AddItem Trim("" & bdip.local1) & "|" & Trim("" & bdip.base) & "|" & Trim("" & bdip.nombre)
        sdx1 = 1
    
        If Trim("" & bdip.defecto) = "S" Then
            gempresa.ListIndex = I - 1
            sdx2 = gempresa.ListIndex

        End If

    Next I

    If sdx1 = 1 Then
        gempresa.ListIndex = 0

    End If

    If sdx2 <> -1 Then
        gempresa.ListIndex = sdx2

    End If

    Close #4
    Exit Sub
cmd78241_err:
    MsgBox "Aviso en cargar bases " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub valida_conec(buf As String)

    Dim found As Integer

    Dim bdip  As ipmaquina

    Dim sdx   As Integer

    On Error GoTo cmd7824_err

    vservidor = ""
    clave_servidor = ""
    Open globalpath & "\config.txt" For Random As #4 Len = Len(bdip)
    sdx = (LOF(4) \ Len(bdip)) '

    sdx = 1

    If gempresa.ListIndex <> -1 Then
        sdx = gempresa.ListIndex + 1

    End If

    Get #4, sdx, bdip
    vservidor = Trim("" & bdip.ip)
    clave_servidor = Trim("" & bdip.clave)
    glocal = Trim("" & bdip.local1)
    Close #4
    Exit Sub
cmd7824_err:
    MsgBox "Aviso en Leer registro " + error$, 48, "Aviso"

End Sub

Sub esconde_menu()
    hiscli.Visible = False
    dki444.Visible = False
    cont121.Visible = False
    dlo1211.Visible = False
    bando1.Visible = False
    prod1.Visible = False
    alki1.Visible = False
    comjk1.Visible = False
    fdjvtas.Visible = False
    tienda1.Visible = False
    dlotablas.Visible = False
    jui34hote.Visible = False

    hiscli.Visible = False

End Sub

Sub recordar()

    Dim sdx As Double

    On Error GoTo cmd67890_err

    Dim mytablex As New ADODB.Recordset

    Label11.Visible = False
    Label12.Visible = False
    cuentacc.Visible = False
    cuentacp.Visible = False

    If Frame2.Visible = True Then
        clave.SetFocus
        Exit Sub

    End If

    Label11.Visible = True
    Label12.Visible = True
    cuentacc.Visible = True
    cuentacp.Visible = True
    
    Label11.ForeColor = RGB(91, 110, 128)
    Label11.BackColor = RGB(91, 110, 128)
    
    Label12.ForeColor = RGB(91, 110, 128)
    Label12.BackColor = RGB(91, 110, 128)
    
    cuentacc.ForeColor = RGB(91, 110, 128)
    Label12.BackColor = RGB(91, 110, 128)
    
    cuentacp.ForeColor = RGB(91, 110, 128)
    Label12.BackColor = RGB(91, 110, 128)
    
    

    cuentacc = "CuentaxCobrar :" & dicmoneda & " 0.00 "
    cuentacp = "CuentaxPagar  :" & dicmoneda & " 0.00 "
    mytablex.Open "SELECT sum(saldo) as xsaldo from cuentac where saldo>0 ", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    sdx = 0
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + Val("" & mytablex.Fields("xsaldo"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    cuentacc = "CuentaxCobrar :" & dicmoneda & Format(sdx, "0.00")

    mytablex.Open "SELECT sum(saldo) as xsaldo from cuentap where saldo>0 ", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    sdx = 0
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + Val("" & mytablex.Fields("xsaldo"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    cuentacp = "CuentaxPagar  :" & dicmoneda & Format(sdx, "0.00")
    Exit Sub
cmd67890_err:
    MsgBox "Aviso en Recordar " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub recordar_ventas()

    Dim buf      As String

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim xfechai  As String

    Dim mytablex As New ADODB.Recordset

    xfechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    xfechai = Format(Now, "dd/mm/yyyy")
    buf = buf & "  fecha>='" & Format(xfechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(Now, "YYYYMMDD") & "' "

    mytablex.Open "SELECT acu,count(numero) as nrof,sum(total) as xtotal from factura where " & buf & " and estado='2' and (acu='A' or acu='B' or acu='C' or acu='D' or acu='G' OR acu='J' or acu='K' or acu='C' or acu='L' or acu='M' or acu='P' ) group by acu", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    sdx = 0
    sdx1 = 0
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("acu") = "A" Or "" & mytablex.Fields("acu") = "B" Or "" & mytablex.Fields("acu") = "C" Or "" & mytablex.Fields("acu") = "D" Or "" & mytablex.Fields("acu") = "G" Then
            sdx = sdx + Val("" & mytablex.Fields("xtotal"))

        End If

        If "" & mytablex.Fields("acu") = "J" Or "" & mytablex.Fields("acu") = "K" Or "" & mytablex.Fields("acu") = "L" Or "" & mytablex.Fields("acu") = "M" Or "" & mytablex.Fields("acu") = "P" Then
            sdx1 = sdx1 + Val("" & mytablex.Fields("xtotal"))

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    Label14 = "Ventas  Dia :" & dicmoneda & Format(sdx, "0.00")
    Label15 = "Compras Dia :" & dicmoneda & Format(sdx1, "0.00")

End Sub

Sub visualizar_recordar()

    On Error GoTo cmd8000_err

    Label13.Visible = False
    Label14.Visible = False
    Label15.Visible = False

    visualizar_ingresos

    Label14 = "Ventas  Dia :" & dicmoneda & " 0.00 "
    Label15 = "Compras Dia :" & dicmoneda & " 0.00 "

    If Frame2.Visible = True Then
        clave.SetFocus
        Exit Sub

    End If

    Label13.Visible = True
    Label14.Visible = True
    Label15.Visible = True
    Exit Sub
cmd8000_err:
    MsgBox "Aviso en visualiza recordar " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub visualizar_ingresos()
    Label16.Visible = False
    Label17.Visible = False
    Label18.Visible = False

    Label16 = "Ingresos  Dia :" & dicmoneda & " 0.00 "
    Label17 = "Egresos   Dia :" & dicmoneda & " 0.00 "

    If Frame2.Visible = True Then
        clave.SetFocus
        Exit Sub

    End If

    Label16.Visible = True
    Label17.Visible = True
    Label18.Visible = True

End Sub

Sub recordar_ingresos()

    Dim buf      As String

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim xfechai  As String

    Dim mytablex As New ADODB.Recordset

    xfechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    xfechai = Format(Now, "dd/mm/yyyy")
    buf = buf & "  fecha>='" & Format(xfechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(Now, "YYYYMMDD") & "' "

    mytablex.Open "SELECT acu,SUM(total) AS xtotal from recibo where " & buf & " and estado='2' and (acu='V' or acu='W' ) group by acu ", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    sdx = 0
    sdx1 = 0
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("acu") = "W" Then
            sdx = sdx + Val("" & mytablex.Fields("xtotal"))

        End If

        If "" & mytablex.Fields("acu") = "V" Then
            sdx1 = sdx1 + Val("" & mytablex.Fields("xtotal"))

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    Label16 = "Ingresos Dia :" & dicmoneda & Format(sdx, "0.00")
    Label17 = "Egresos Dia  :" & dicmoneda & Format(sdx1, "0.00")

End Sub

Sub esconde_todo()
    Label13.Visible = False
    Label14.Visible = False
    Label15.Visible = False

    Label16.Visible = False
    Label17.Visible = False
    Label18.Visible = False

    Label11.Visible = False
    Label12.Visible = False
    cuentacc.Visible = False
    cuentacp.Visible = False

End Sub

Function verifica_licenciaremoto()

    Dim ybuf  As String

    Dim found As Integer

    On Error GoTo cmd99169999_err

    'Exit Function
    'If CDate(Format(Now, "dd/mm/yyyy")) > CDate("30/09/2013") Then
    'MsgBox "xx"
    '   found = numero_registro()
    '   Exit Function
    '
    'End If
    'Exit Function

    ybuf = "" & menup.vservidor

    If Trim(UCase$("" & menup.vservidor)) <> "(LOCAL)" Then
        If IsValidIPAddress(menup.vservidor) = False Then
            MsgBox "Ip no Esta bien Definido ", 48, "Aviso"
            Exit Function

        End If

    Else  'SI ES LOCAL
        ybuf = GetMACs_AdaptInfo()  'si es local

    End If
    
    'If Mid$(ybuf, 1, 3) <> "169" And Mid$(ybuf, 1, 3) <> "127" And Mid$(ybuf, 1, 3) <> "192" Then
    ' MsgBox "Es externo"
    'End If
    
    If IsValidIPAddress(ybuf) = True Then
        If verifica_mac(ybuf, "LICENCIACENTRALIZADO") = 0 Then
            found = numero_registro()

        End If

        Exit Function

    End If

    If verifica_disco_duro("LICENCIACENTRALIZADO") = 0 Then
        found = numero_registro()

    End If

    Exit Function
cmd99169999_err:
    MsgBox "No se cargas Iniciales " + error$, 48, "Aviso"
    Close
    Exit Function

End Function

'
'
''Testing Proyecto Facturacion Electronica 05/04/2018
'Public Function valida_facturacionElectronica()
'Dim mytable As New ADODB.Recordset
'Dim mysql As String
'Dim k As Integer
'Dim salida As String
'rpta = ""
'
' mysql = ""
' mysql = "SELECT  codigo1  from Tlocal"
' mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
' If mytable.RecordCount > 0 Then  'si existe
'    Do
'    If mytable.EOF Then Exit Do
'        If mytable.Fields("CODIGO1") = "" Then
'          rpta = "* FALTA AGREGAR RUC (CODIGO1) A LOCAL"
'          Exit Do
'        End If
'    mytable.MoveNext
'    Loop
' End If
' mytable.Close
'
' mysql = ""
' mysql = "SELECT CAJA,DESCRIPCIO,SERIETB,SERIETF from PARAMECA WHERE CAJA <>00"
' mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
' If mytable.RecordCount > 0 Then  'si existe
'    Do
'    If mytable.EOF Then Exit Do
'        If Len(mytable.Fields("serietb")) <> 0 Then
'          If Len(mytable.Fields("serietb")) <> 4 And Len(mytable.Fields("serietb")) = 0 Then
'            rpta = rpta & " * VERIFICAR SERIE DE DOCUMENTO"
'            Exit Do
'          End If
'        End If
'    mytable.MoveNext
'    Loop
' End If
' mytable.Close
'
' If Servicio = "Servicio Detenido" Then
'    Shell ("C:\BYH\iniciar.bat"), vbNormalFocus
'    Servicio = "Servicio Activo"
' End If
'
' mysql = ""
' mysql = "SELECT  codigo1  from Tlocal where codigo='01' "
' mytable.Open mysql, cn, adOpenStatic, adLockOptimistic
' If mytable.RecordCount > 0 Then  'si existe
'          RucSql = mytable.Fields("codigo1")
' End If
' mytable.Close
'
' Call lee_conf_RUC("C:\BYH\SERVICIO\VISITEC\application.yml", "A")
'
' If RucYml <> RucSql Then
'   rpta = rpta & " * RUC de Yml DIFERENTE A RUC de Local"
' End If
'
' Const ATTR_DIRECTORY = 16
' If Dir$("D:\ce_output", ATTR_DIRECTORY) = "" Then
'    rpta = rpta & " * Carpeta D:\ce_output NO EXISTE"
' End If
'
' If Dir$("D:\ce_output\CREA", ATTR_DIRECTORY) = "" Then
'    rpta = rpta & " * Carpeta D:\ce_output\CREA NO EXISTE"
' End If
' If Dir$("D:\ce_output\ERROR", ATTR_DIRECTORY) = "" Then
'    rpta = rpta & " * Carpeta D:\ce_output\ERROR NO EXISTE"
' End If
' If Dir$("D:\ce_output\FIRMADO", ATTR_DIRECTORY) = "" Then
'    rpta = rpta & " * Carpeta D:\ce_output\FIRMADO NO EXISTE"
' End If
' If Dir$("D:\ce_output\PROCESADO", ATTR_DIRECTORY) = "" Then
'    rpta = rpta & " * Carpeta D:\ce_output\PROCESADO NO EXISTE"
' End If
'
'Exit Function
'End Function
''Testing Proyecto Facturacion Electronica 05/04/2018

