VERSION 5.00
Object = "{19BD1EA6-6E36-45BA-AEBD-BCF3093017CC}#11.0#0"; "GorditoButton.ocx"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form tcomanda 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Comandas"
   ClientHeight    =   9945
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "tcomanda.frx":0000
   ScaleHeight     =   9945
   ScaleMode       =   0  'User
   ScaleWidth      =   19872.86
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox panelpersonas 
      BackColor       =   &H80000004&
      Height          =   2790
      Left            =   0
      ScaleHeight     =   2730
      ScaleWidth      =   3165
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   0
      Width           =   3228
      Begin VB.TextBox xcomando 
         Alignment       =   2  'Center
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
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   53
         Top             =   0
         Width           =   1230
      End
      Begin GorditoButton.Boton Comando 
         Height          =   750
         Index           =   0
         Left            =   15
         TabIndex        =   54
         Top             =   615
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "0"
         BackColor       =   4210752
         ResalteColor    =   12632256
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
         Left            =   780
         TabIndex        =   55
         Top             =   600
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "1"
         BackColor       =   4210752
         ResalteColor    =   12632256
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
         Left            =   1560
         TabIndex        =   56
         Top             =   600
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "2"
         BackColor       =   4210752
         ResalteColor    =   12632256
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
         Left            =   2370
         TabIndex        =   57
         Top             =   600
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "3"
         BackColor       =   4210752
         ResalteColor    =   12632256
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
         Left            =   0
         TabIndex        =   58
         Top             =   1305
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "4"
         BackColor       =   4210752
         ResalteColor    =   12632256
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
         Left            =   780
         TabIndex        =   59
         Top             =   1305
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "5"
         BackColor       =   4210752
         ResalteColor    =   12632256
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
         Left            =   1575
         TabIndex        =   60
         Top             =   1305
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "6"
         BackColor       =   4210752
         ResalteColor    =   12632256
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
         Left            =   2370
         TabIndex        =   61
         Top             =   1305
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "7"
         BackColor       =   4210752
         ResalteColor    =   12632256
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
         Left            =   15
         TabIndex        =   62
         Top             =   2010
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "8"
         BackColor       =   4210752
         ResalteColor    =   12632256
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
         Left            =   780
         TabIndex        =   63
         Top             =   2010
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "9"
         BackColor       =   4210752
         ResalteColor    =   12632256
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
         Left            =   1575
         TabIndex        =   64
         Top             =   2010
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   1323
         PicturePosition =   0
         Caption         =   "CR"
         BackColor       =   4210752
         ResalteColor    =   12632256
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
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "# de Personas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   540
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Width           =   1950
      End
   End
   Begin VB.CommandButton Command3d1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   7
      Left            =   -1200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1650
      Width           =   855
   End
   Begin ChamaleonButton.ChameleonBtn groupsalon 
      Height          =   855
      Index           =   0
      Left            =   3600
      TabIndex        =   15
      Top             =   1200
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421504
      MPTR            =   1
      MICON           =   "tcomanda.frx":2479C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupsalon 
      Height          =   855
      Index           =   1
      Left            =   5565
      TabIndex        =   16
      Top             =   1200
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421504
      MPTR            =   1
      MICON           =   "tcomanda.frx":2479DE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupsalon 
      Height          =   855
      Index           =   2
      Left            =   7530
      TabIndex        =   17
      Top             =   1200
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421504
      MPTR            =   1
      MICON           =   "tcomanda.frx":2479FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupsalon 
      Height          =   855
      Index           =   3
      Left            =   9495
      TabIndex        =   18
      Top             =   1200
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421504
      MPTR            =   1
      MICON           =   "tcomanda.frx":247A16
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   0
      Left            =   1260
      TabIndex        =   19
      Top             =   3390
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247A32
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   1
      Left            =   1260
      TabIndex        =   20
      Top             =   4545
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247A4E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   2
      Left            =   1260
      TabIndex        =   21
      Top             =   5700
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247A6A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   3
      Left            =   1260
      TabIndex        =   22
      Top             =   6855
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247A86
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   5
      Left            =   3375
      TabIndex        =   23
      Top             =   3390
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247AA2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   6
      Left            =   3375
      TabIndex        =   24
      Top             =   4545
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247ABE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   7
      Left            =   3375
      TabIndex        =   25
      Top             =   5700
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247ADA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   8
      Left            =   3375
      TabIndex        =   26
      Top             =   6855
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247AF6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   10
      Left            =   5505
      TabIndex        =   27
      Top             =   3390
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247B12
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   11
      Left            =   5505
      TabIndex        =   28
      Top             =   4545
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247B2E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   12
      Left            =   5505
      TabIndex        =   29
      Top             =   5700
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247B4A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   13
      Left            =   5505
      TabIndex        =   30
      Top             =   6855
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247B66
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   15
      Left            =   7620
      TabIndex        =   31
      Top             =   3390
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247B82
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   16
      Left            =   7620
      TabIndex        =   32
      Top             =   4545
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247B9E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   17
      Left            =   7620
      TabIndex        =   33
      Top             =   5700
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247BBA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   18
      Left            =   7620
      TabIndex        =   34
      Top             =   6855
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247BD6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   20
      Left            =   9735
      TabIndex        =   35
      Top             =   3390
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247BF2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   21
      Left            =   9735
      TabIndex        =   36
      Top             =   4545
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247C0E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   22
      Left            =   9735
      TabIndex        =   37
      Top             =   5700
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247C2A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   23
      Left            =   9735
      TabIndex        =   38
      Top             =   6855
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247C46
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   25
      Left            =   11865
      TabIndex        =   39
      Top             =   3390
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247C62
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   26
      Left            =   11865
      TabIndex        =   40
      Top             =   4545
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247C7E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   27
      Left            =   11865
      TabIndex        =   41
      Top             =   5700
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247C9A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   28
      Left            =   11865
      TabIndex        =   42
      Top             =   6855
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247CB6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Label5 
      Height          =   1215
      Left            =   12360
      TabIndex        =   43
      Top             =   240
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   2143
      BTYPE           =   3
      TX              =   "Graba Comanda"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   192
      BCOLO           =   8421631
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421504
      MPTR            =   1
      MICON           =   "tcomanda.frx":247CD2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Label7 
      Height          =   855
      Left            =   12360
      TabIndex        =   44
      Top             =   1665
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Cerrar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421504
      BCOLO           =   8421504
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421504
      MPTR            =   1
      MICON           =   "tcomanda.frx":247CEE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   4
      Left            =   1260
      TabIndex        =   45
      Top             =   8010
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247D0A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   9
      Left            =   3375
      TabIndex        =   46
      Top             =   8010
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247D26
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   14
      Left            =   5505
      TabIndex        =   47
      Top             =   8010
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247D42
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   19
      Left            =   7620
      TabIndex        =   48
      Top             =   8010
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247D5E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   24
      Left            =   9735
      TabIndex        =   49
      Top             =   8010
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247D7A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn groupmesa 
      Height          =   1125
      Index           =   29
      Left            =   11865
      TabIndex        =   50
      Top             =   8010
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   1984
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   13160660
      BCOLO           =   13160660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "tcomanda.frx":247D96
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label xindex 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Left            =   11865
      TabIndex        =   14
      Top             =   5700
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GRABA COMANDA PRECUENTA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   13590
      TabIndex        =   13
      Top             =   2475
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   840
      Left            =   3840
      Picture         =   "tcomanda.frx":247DB2
      Stretch         =   -1  'True
      Top             =   2535
      Width           =   1440
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   9960
      Picture         =   "tcomanda.frx":249DC0
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1440
   End
   Begin VB.Image Image6 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   5400
      Picture         =   "tcomanda.frx":24C21A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1320
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   8760
      Picture         =   "tcomanda.frx":24E228
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1395
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FOTO VENDEDOR"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13770
      TabIndex        =   12
      Top             =   2820
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   165
      Left            =   13800
      Top             =   2730
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LIMPIA PANTALLA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   14130
      TabIndex        =   11
      Top             =   3405
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COMANDA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13305
      TabIndex        =   10
      Top             =   2820
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VENDEDOR >"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   4560
      TabIndex        =   9
      Top             =   9240
      Width           =   1965
   End
   Begin VB.Label xcomanda2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   13440
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label comanda 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   14265
      TabIndex        =   7
      Top             =   3495
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label mesero 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   8970
      TabIndex        =   6
      Top             =   9240
      Width           =   2175
   End
   Begin VB.Label nmesero 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   6600
      TabIndex        =   5
      Top             =   9240
      Width           =   2370
   End
   Begin VB.Label mesa 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Left            =   10980
      TabIndex        =   4
      Top             =   5700
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label salon 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   13485
      TabIndex        =   3
      Top             =   2775
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MESA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5280
      TabIndex        =   2
      Top             =   2520
      Width           =   4687
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALON"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6720
      TabIndex        =   1
      Top             =   480
      Width           =   2040
   End
   Begin VB.Menu dlo922 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcomanda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msalcod(100)    As String

Dim msalpag         As Integer

Dim msaltop         As Integer

Dim mmesacod(15000) As String

Dim wmesacod(15000) As String

Dim wwmesacod(30)   As String

Dim mmesapag        As Integer

Dim mmesatop        As Integer

Private Sub Comando_Click(Index As Integer)

    If Index = 10 Then
        xcomando.Text = ""
        Exit Sub

    End If

    xcomando = xcomando + Comando(Index).Caption

End Sub

Private Sub dlo922_Click()
    tcomanda.Hide
    Unload tcomanda

End Sub

Private Sub Form_Activate()

    Dim I As Integer

    For I = 0 To 2

        If Trim(groupsalon(I).Caption) = Trim("" & mytable11.Fields("salon")) Then
            groupsalon_Click I
            Exit For

        End If

    Next I

    Label2.Caption = "" & glomesa
    xindex = ""

End Sub

Private Sub Form_Load()
    carga_salon

    ''' kenyo 21/09/2017 Mayor numero de mesas (30)
    If mytable11.Fields("obligapersonas") = "S" And xcomando.Text = "" Then
        panelpersonas.Visible = True
    Else
        panelpersonas.Visible = False

    End If

    ''' kenyo 21/09/2017 Mayor numero de mesas (30)

End Sub

Sub carga_salon()

    Dim mytablex As New ADODB.Recordset

    Dim I        As Integer

    For I = 0 To 99
        msalcod(I) = ""
    Next I

    I = -1
    mytablex.Open "select * from salon ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        I = I + 1
        msalcod(I) = "" & mytablex.Fields("salon")
        mytablex.MoveNext
    Loop
    msaltop = I
    mytablex.Close
    msalpag = 0
    menu_salon "INI"

End Sub

Sub menu_salon(buf As String)

    Dim I As Integer

    Dim j As Integer

    Select Case buf

        Case "INI"
            msalpag = 0

        Case "SIG"
            msalpag = msalpag + 3

            If msalpag > 102 Then
                msalpag = 0

            End If

        Case "ANT"
            msalpag = msalpag - 3

            If msalpag < 0 Then
                msalpag = 0

            End If

    End Select

    j = -1

    For I = msalpag To 3 + msalpag
        j = j + 1
        groupsalon(j).Caption = msalcod(I)
    Next I

End Sub

Function eselmeseroo(buf As String, buf1 As String, buf2 As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from dcomanda where salon='" & buf & "' and mesa='" & buf1 & "' and vendedor='" & buf2 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        eselmeseroo = 1

    End If

    mytablex.Close

End Function

Function verifica_permisom(buf As String, buf1 As String, buf2 As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from usermesa where salon='" & buf & "' and mesa='" & buf1 & "' and codigo='" & buf2 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        verifica_permisom = 1

    End If

    mytablex.Close

End Function

Private Sub groupmesa_Click(Index As Integer)

    Dim I       As Integer

    Dim k       As Integer

    Dim found   As Integer

    Dim mmesero As String

    'MsgBox Index
    'MsgBox Trim("" & groupmesa(Index).Caption)
    mesa = ""

    If phabilita_mesero() = "S" Then  'si el flag esta habilitado
        found = verifica_permisom(Trim("" & salon), Trim("" & groupmesa(Index).Caption), Trim("" & mesero))

        If found = 0 Then
            MsgBox "Mesa No habilitado para el Mesero,hable con su administrador ", 48, "Aviso"
            refresca_mesa
            Exit Sub

        End If

    End If

    'MsgBox Index
    mmesero = mismo_mesero()

    If groupmesa(Index).BackColor = &HFF00& Then
        If Len(Trim("" & salon)) > 0 And Len(groupmesa(Index).Caption) > 0 Then
            found = eselmeseroo(Trim("" & salon), Trim("" & groupmesa(Index).Caption), Trim("" & mesero))

            If found = 0 Then
                If Trim(mmesero) = "S" Or Trim(mmesero) = "" Then
                    MsgBox "Esta mesa no fue Aperturado por este codigo Mesero ", 48, "Aviso"
                    refresca_mesa
                    Exit Sub

                End If

            End If

        End If

    End If

    'MsgBox Index
    'MsgBox groupmesa(Index).Caption

    ''' kenyo 31/08/2017 Mayor numero de mesas (30)
    'For i = 0 To 23
    For I = 0 To 29
        ''' kenyo 31/08/2017 Mayor numero de mesas (30)

        groupmesa(I).BackColor = &HFFFFFF
    Next I

    'refresca_mesa
    If Len(groupmesa(Index).Caption) = 0 Then Exit Sub
    groupmesa(Index).BackColor = &HFF&
    groupmesa(Index).BackOver = &HFF&

    '---------------------------------------
    If groupsalon(0).BackColor = &HFFFFFF And groupsalon(1).BackColor = &HFFFFFF And groupsalon(2).BackColor = &HFFFFFF And groupsalon(3).BackColor = &HFFFFFF Then
        'refresca_mesa
        Exit Sub

    End If

    k = 0

    If groupsalon(0).BackColor = &HFF& Then
        k = 0

    End If

    If groupsalon(1).BackColor = &HFF& Then
        k = 1

    End If

    If groupsalon(2).BackColor = &HFF& Then
        k = 2

    End If

    If groupsalon(3).BackColor = &HFF& Then
        k = 3

    End If

    mesa = groupmesa(Index).Caption

    'groupsalon_Click (k)
    'refresca_mesa
End Sub

Private Sub groupsalon_Click(Index As Integer)

    Dim I As Integer

    If Len(groupsalon(Index).Caption) = 0 Then Exit Sub

    For I = 0 To 3
        groupsalon(I).BackColor = &HFFFFFF
    Next I

    ''' kenyo 31/08/2017 Mayor numero de mesas (30)
    'For i = 0 To 23
    For I = 0 To 29
        ''' kenyo 31/08/2017 Mayor numero de mesas (30)
        groupmesa(I).BackColor = &HFFFFFF

    Next I

    groupsalon(Index).BackColor = &HFF&
    menu_carga_mesa groupsalon(Index).Caption
    menu_mesa "INI", groupsalon(Index).Caption
    salon = groupsalon(Index).Caption
    pone_nombre_salon "" & groupsalon(Index).Caption
    mesa = ""
    xindex = "" & Index

End Sub

Sub refresca_mesa()

    Dim I     As Integer

    Dim Index As Integer

    Index = Val(xindex)

    If Len(Trim(salon)) = 0 Then Exit Sub

    For I = 0 To 3
        groupsalon(I).BackColor = &HFFFFFF
    Next I

    ''' kenyo 31/08/2017 Mayor numero de mesas (30)
    'For i = 0 To 23
    For I = 0 To 29
        ''' kenyo 31/08/2017 Mayor numero de mesas (30)

        groupmesa(I).BackColor = &HFFFFFF
    Next I

    groupsalon(Index).BackColor = &HFF&
    menu_carga_mesa groupsalon(Index).Caption
    menu_mesa "INI", groupsalon(Index).Caption
    salon = groupsalon(Index).Caption
    pone_nombre_salon "" & groupsalon(Index).Caption
    mesa = ""

End Sub

Sub verifica_mesas(indx As Integer, buf As String, buf1 As String)

    '''' 25/07/2018 Delivery y Para Llevar desde mozo
    'Dim mytablex As New ADODB.Recordset
    'groupmesa(indx).BackColor = &HFFFFFF
    'If Len(buf1) > 0 And Len(buf) > 0 Then
    '   mytablex.Open "select * from dcomanda where salon='" & buf & "' and mesa='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic
    '   If mytablex.RecordCount > 0 Then
    '      groupmesa(indx).BackColor = &HFF00&
    '   End If
    '   mytablex.Close
    'End If
    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    groupmesa(indx).BackColor = &HFFFFFF

    If Len(buf1) > 0 And Len(buf) > 0 Then
        mytablex.Open "select * from dcomanda where salon='" & Trim(buf) & "' and mesa='" & Trim(buf1) & "'", cn, adOpenStatic, adLockOptimistic

        'mytablex.Open "select * from MESA where salon='" & buf & "' and mesa='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic
        If mytablex.RecordCount > 0 Then
            groupmesa(indx).BackColor = &HFF00&
            mytabley.Open "select * from MESA where salon='" & Trim(buf) & "' and mesa='" & Trim(buf1) & "'", cn, adOpenStatic, adLockOptimistic

            If mytabley.RecordCount > 0 Then
                If Trim("" & mytabley.Fields("estado")) = "2" Then
                    groupmesa(indx).BackColor = &HFFFF&     '&HFF&

                End If

            End If

            mytabley.Close

        End If

        mytablex.Close

    End If

    '''' 25/07/2018 Delivery y Para Llevar desde mozo

End Sub

Sub menu_carga_mesa(buf As String)

    Dim mytablex As New ADODB.Recordset

    Dim I        As Integer

    For I = 0 To 29
        wwmesacod(I) = ""
    Next I

    For I = 0 To 14999
        mmesacod(I) = ""
        wmesacod(I) = ""
    Next I

    I = -1

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM mesa where salon='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("salon") = buf Then
   
            I = I + 1

            ' Mesas Nombre 21/05/2018
        
            ' mmesacod(I) = "" & mytablex.Fields("mesa")
            ' wmesacod(I) = "" & mytablex.Fields("mesa")
   
            Dim rpta As String

            rpta = ""
            mmesacod(I) = " " & mytablex.Fields("mesa")
            wmesacod(I) = "" & mytablex.Fields("mesa")
   
            If Len(Trim(mytablex.Fields("dnombre"))) > 0 Then
                '24/08/2018  Delivery por mesa
                'rpta = busca_clienteDelivery_mesa(mytablex.Fields("codigo"), buf)
                rpta = "" & Mid$(mytablex.Fields("dnombre"), 1, 25)
  
                '24/08/2018  Delivery por mesa
            End If
   
            If Len(rpta) > 0 Then
                mmesacod(I) = mmesacod(I) & "| " & rpta

            End If
   
            ' Mesas Nombre 21/05/2018
   
            Else: Exit Do

        End If

        mytablex.MoveNext
    Loop

    mytablex.Close
    mmesatop = I
    mmesapag = 0

End Sub

Sub menu_mesa(buf As String, buf1 As String)

    Dim I As Integer

    Dim j As Integer

    Select Case buf

        Case "INI"
            mmesapag = 0

        Case "SIG"
       
            ''' kenyo 31/08/2017 Mayor numero de mesas (30)
            'mmesapag = mmesapag + 23
            mmesapag = mmesapag + 29
            ''' kenyo 31/08/2017 Mayor numero de mesas (30)
            
            If mmesapag > 102 Then
                mmesapag = 0

            End If

        Case "ANT"
       
            ''' kenyo 31/08/2017 Mayor numero de mesas (30)
            'mmesapag = mmesapag - 23
            mmesapag = mmesapag - 29
            ''' kenyo 31/08/2017 Mayor numero de mesas (30)
            
            If mmesapag < 0 Then
                mmesapag = 0

            End If

    End Select

    j = -1

    ''' kenyo 31/08/2017 Mayor numero de mesas (30)
    'For i = mmesapag To 23 + mmesapag
    For I = mmesapag To 29 + mmesapag
        ''' kenyo 31/08/2017 Mayor numero de mesas (30)
    
        j = j + 1
        groupmesa(j).Caption = mmesacod(I)

        '''' 25/07/2018 Delivery y Para Llevar desde mozo
        ' Mesas Nombre 21/05/2018
        'verifica_mesas j, buf1, groupmesa(j).Caption
        verifica_mesas j, buf1, extra_loquesea(groupmesa(j).Caption)
        ' Mesas Nombre 21/05/2018
        '''' 25/07/2018 Delivery y Para Llevar desde mozo
    Next I

End Sub

Private Sub Image2_Click()

    Dim I As Integer

    ''' kenyo 31/08/2017 Mayor numero de mesas (30)
    'For i = 0 To 23
    For I = 0 To 29
        ''' kenyo 31/08/2017 Mayor numero de mesas (30)
    
        groupmesa(I).BackColor = &HFFFFFF
        mesa = ""
    Next I

    menu_mesa "SIG", salon

End Sub

Private Sub Image3_Click()

    Dim I As Integer

    ''' kenyo 31/08/2017 Mayor numero de mesas (30)
    'For i = 0 To 23
    For I = 0 To 29
        ''' kenyo 31/08/2017 Mayor numero de mesas (30)
    
        groupmesa(I).BackColor = &HFFFFFF
        mesa = ""
    Next I

    menu_mesa "ANT", salon

End Sub

Private Sub Image5_Click()

    Dim I As Integer

    For I = 0 To 3
        groupsalon(I).BackColor = &HFFFFFF
        salon = ""
        mesa = ""
    Next I

    ''' kenyo 31/08/2017 Mayor numero de mesas (30)
    'For i = 0 To 23
    For I = 0 To 29
        ''' kenyo 31/08/2017 Mayor numero de mesas (30)
    
        groupmesa(I).Caption = ""
        groupmesa(I).BackColor = &HFFFFFF
    Next I

    menu_salon "SIG"

End Sub

Private Sub Image6_Click()

    Dim I As Integer

    For I = 0 To 3
        groupsalon(I).BackColor = &HFFFFFF
        salon = ""
        mesa = ""
    Next I

    ''' kenyo 31/08/2017 Mayor numero de mesas (30)
    'For i = 0 To 23
    For I = 0 To 29
        ''' kenyo 31/08/2017 Mayor numero de mesas (30)
    
        groupmesa(I).Caption = ""
        groupmesa(I).BackColor = &HFFFFFF
    Next I

    menu_salon "ANT"

End Sub

Private Sub Label5_Click()

    Dim found As Integer

    If Len(comanda) = 0 Then
        comanda = "1"
        xcomanda2 = "1"

    End If

    xcomanda2 = busca_nroc()
    comanda = xcomanda2

    If Len(Trim(salon)) = 0 Then
        MsgBox "No ha seleccionado Salon", 24, "Aviso"
        Exit Sub

    End If

    If Len(Trim(mesa)) = 0 Then
        MsgBox "No ha seleccionado Mesa", 24, "Aviso"
        Exit Sub

    End If

    If Len(Trim(mesero)) = 0 Then
        MsgBox "No ha seleccionado Mesero", 24, "Aviso"
        Exit Sub

    End If

    If Len(Trim(comanda)) = 0 Then
        MsgBox "No ha seleccionado comanda", 24, "Aviso"
        Exit Sub

    End If
 
    If mytable11.Fields("obligapersonas") = "S" And xcomando.Text = "" Then
        MsgBox "Ingrese número de personas", 24, "Aviso"
        xcomando.SetFocus
        Exit Sub

    End If
                                       
    '''' 25/07/2018 Delivery y Para Llevar desde mozo
    V_EstadoMesa = "" & busca_tipoSalon(Trim(salon))

    If V_EstadoMesa = "C" Then
        If Len(tptovta.dcodigo) > 0 Or Len(tptovta.dcodigo) Then
            MsgBox ("Borre datos de cliente para MESA ")
            Exit Sub

        End If

    ElseIf V_EstadoMesa = "D" Then
    
        '24/08/2018  Delivery por mesa
        If busca_EstadoXMesa(Trim(salon), Trim(extra_loquesea(mesa))) = "L" Then
            If Len(tptovta.dcodigo) = 0 Then
                MsgBox ("Agregue datos de cliente para DELIVERY")
                Exit Sub

            End If

        End If
        
        '24/08/2018  Delivery por mesa
        
    ElseIf V_EstadoMesa = "L" Then

        '24/08/2018  Delivery por mesa
        If Len(tptovta.dcodigo) > 0 Then
            MsgBox ("No Considerar Datos de DELIVERY")
            Exit Sub

        End If
        
        If busca_EstadoXMesa(Trim(salon), Trim(extra_loquesea(mesa))) = "L" Then
            If Len(tptovta.nombre) = 0 Then
                MsgBox ("Agregue datos de cliente para LLEVAR")
                Exit Sub

            End If

        End If
        
        '24/08/2018  Delivery por mesa

        '        If Len(tptovta.dcodigo) > 0 Then
        '            MsgBox ("No Considerar Datos de DELIVERY")
        '            Exit Sub
        '        End If
        '
        '        If Len(tptovta.nombre) > 0 Then
        '            graba_cliente_tipo (tptovta.codigo)
        '        End If
        '
        '        If Len(tptovta.codigo) = 0 Then
        '            MsgBox ("Agregue datos de cliente para LLEVAR")
        '            Exit Sub
        '        End If

    End If
    
    '''' 25/07/2018 Delivery y Para Llevar desde mozo
                        
    found = grabar_comanda()

    If found = 0 Then
        MsgBox "No existen datos ", 48, "Aviso"
        Exit Sub

    End If

    'imprimio comanda
    'orden13012015
    '------------------------
    tptovta.salon.Caption = salon
    tptovta.mesa.Caption = mesa
    tptovta.mesero = mesero
    flag_comanda = 1
    'despacho_orden salon, mesa, Trim(tptovta.caja)
    'Exit Sub
    tptovta.comanda = comanda
    tcomanda.Hide
    Unload tcomanda
                        
    'tservicio = "C"
    'found = orden_despacho()
    'mytable11.Edit
    'mytable11.Fields("comandanro") = xcomanda
    'mytable11.Update
    'si todo sale bien debe inicializar en el load
    'del pedido todo en cero
                        
    Exit Sub

End Sub

Function grabar_comanda()

    On Error GoTo cmd788_err

    Dim buf       As String

    Dim sdx       As Double

    Dim mytablezx As New ADODB.Recordset

    Dim mytablex  As New ADODB.Recordset

    Dim I         As Integer

    Dim sw        As Integer

    sdx = 0
    'sum2 = 0
    sw = 0

    '-----------
    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM dcomanda where caja='" & Trim(comanda) & "'", cn, adOpenDynamic, adLockOptimistic
    '  ------- grabar chicas --------
    tptovta.Data2.refresh
    'MsgBox tptovta.Data2.Recordset.Fields.count - 1
    Do

        If tptovta.Data2.Recordset.EOF Then Exit Do
        If Len("" & tptovta.Data2.Recordset.Fields("producto")) > 0 And (Val("" & tptovta.Data2.Recordset.Fields("cantidad")) > 0 Or Val("" & tptovta.Data2.Recordset.Fields("cantidad")) < 0) Then
            mytablex.AddNew

            For I = 0 To tptovta.Data2.Recordset.Fields.count - 1
                mytablex.Fields(I) = tptovta.Data2.Recordset.Fields(I)
            Next I
                 
            'MsgBox glocal & " " & gusuario
            mytablex.Fields("local") = Trim(glocal)
            mytablex.Fields("usuario") = Trim(gusuario)
            mytablex.Fields("caja") = Trim(tptovta.caja)
            mytablex.Fields("turno") = Trim(tptovta.turno)
            mytablex.Fields("vendedor") = Trim(mesero)
                 
            mytablex.Fields("salon") = Trim(salon)
                    
            '24/08/2018  Delivery por mesa
            ' mytablex.Fields("mesa") = Trim(mesa)
            
            mytablex.Fields("mesa") = Trim(extra_loquesea(mesa))
            '24/08/2018  Delivery por mesa
            
            mytablex.Fields("comanda") = Trim(comanda)
            mytablex.Fields("numero") = Trim(comanda)
            mytablex.Fields("estado") = "0"
            mytablex.Fields("servicio") = "C"
            mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
            mytablex.Fields("Hora") = Format(Now, "hh:mm:ss")
                 
            '24/08/2018  Delivery por mesa
            ' 25/07/2018 Delivery y Para Llevar desde mozo
            If V_EstadoMesa = "C" Then
                mytablex.Fields("codigo") = ""
                mytablex.Fields("nombre") = ""
            ElseIf V_EstadoMesa = "D" Then

                If busca_EstadoXMesa(Trim(salon), Trim(extra_loquesea(mesa))) = "L" Then
                    mytablex.Fields("codigo") = Trim(tptovta.dcodigo)
                    mytablex.Fields("nombre") = Trim(tptovta.dnombre)
                Else
                    mytablex.Fields("codigo") = extrae_datos_comanda(1, Trim(extra_loquesea(salon)), Trim(extra_loquesea(mesa)))
                    mytablex.Fields("nombre") = extrae_datos_comanda(2, Trim(extra_loquesea(salon)), Trim(extra_loquesea(mesa)))
                   
                End If
                
            ElseIf V_EstadoMesa = "L" Then

                If busca_EstadoXMesa(Trim(salon), Trim(extra_loquesea(mesa))) = "L" Then
         
                    mytablex.Fields("codigo") = Trim(tptovta.codigo)
                    mytablex.Fields("nombre") = Trim(tptovta.nombre)
                    
                Else
                    mytablex.Fields("codigo") = extrae_datos_comanda(1, Trim(extra_loquesea(salon)), Trim(extra_loquesea(mesa)))
                    mytablex.Fields("nombre") = extrae_datos_comanda(2, Trim(extra_loquesea(salon)), Trim(extra_loquesea(mesa)))
                
                End If

            End If

            ' 25/07/2018 Delivery y Para Llevar desde mozo
            '24/08/2018  Delivery por mesa
 
            'MsgBox "b"
            mytablex.Update
            sw = 1

        End If
       
        '------------- finalizamos el combo
               
        tptovta.Data2.Recordset.MoveNext
             
    Loop
    mytablex.Close
    'MsgBox "abc"
       
    '''' 25/07/2018 Delivery y Para Llevar desde mozo
    ' Mesas Nombre 21/05/2018
    '  If Len(tptovta.dcodigo) > 0 Then
    '   cn.Execute ("update MESA set CODIGO='" & tptovta.dcodigo & "' where mesa='" & Trim(mesa) & "' and salon='" & Trim(salon) & "'")
    '  End If
    ' Mesas Nombre 21/05/2018
    '''' 25/07/2018 Delivery y Para Llevar desde mozo
       
    If sw = 1 Then 'si se ha gravado bien graba en mesa salon la fecha y hora
           
        '''' 25/07/2018 Delivery y Para Llevar desde mozo
        ' Mesas Nombre 21/05/2018
        'If tptovta.dcodigo > 0 Then
        
        If Len(tptovta.dcodigo) > 0 Then
            cn.Execute ("update MESA set CODIGO='" & Trim(tptovta.dcodigo) & "'  where mesa='" & Trim(mesa) & "' and salon='" & Trim(salon) & "'")
            cn.Execute ("update MESA set dnombre='" & Trim(tptovta.dnombre) & "'  where mesa='" & Trim(mesa) & "' and salon='" & Trim(salon) & "'")
       
            '24/08/2018  Delivery por mesa
            cn.Execute ("update MESA set telefono='" & Trim(tptovta.telefono) & "'  where mesa='" & Trim(mesa) & "' and salon='" & Trim(salon) & "'")
            cn.Execute ("update MESA set ddireccion='" & Trim(tptovta.ddireccion) & "'  where mesa='" & Trim(mesa) & "' and salon='" & Trim(salon) & "'")
            cn.Execute ("update MESA set referencia='" & Trim(tptovta.referencia) & "'  where mesa='" & Trim(mesa) & "' and salon='" & Trim(salon) & "'")
            '24/08/2018  Delivery por mesa
        
        End If
        
        If Len(tptovta.nombre) > 0 Then
            cn.Execute ("update MESA set codigo='" & Trim(tptovta.codigo) & "'  where mesa='" & Trim(mesa) & "' and salon='" & Trim(salon) & "'")
            cn.Execute ("update MESA set dnombre='" & Trim(tptovta.nombre) & "'  where mesa='" & Trim(mesa) & "' and salon='" & Trim(salon) & "'")
           
        End If
        
        ' Mesas Nombre 21/05/2018
        '''' 25/07/2018 Delivery y Para Llevar desde mozo
           
        mytablex.Open "SELECT * FROM mesa where salon='" & Trim(salon) & "' and mesa='" & Trim(mesa) & "'", cn, adOpenDynamic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
            mytablex.Fields("Hora") = Format(Now, "hh:mm:ss")

            'mesas kenyo
            If Len(Trim(xcomando)) = 0 Then
                mytablex.Fields("capacidad") = "0"
            Else
                mytablex.Fields("capacidad") = xcomando

            End If
      
            mytablex.Update

        End If

        mytablex.Close
        
    End If

    grabar_comanda = sw
    Exit Function
cmd788_err:
    MsgBox "Aviso en tcomanda " + error$ + " " & I, 48, "Aviso"
    Exit Function

End Function

Function busca_nroc() As String

    Dim sdx      As Double

    Dim mytablex As Table

    sdx = Val("" & mytable11.Fields("comandanro")) + 1
    xcomanda2 = "" & sdx
    'mytable11.Edit
    mytable11.Fields("comandanro") = xcomanda2
    mytable11.Update
    busca_nroc = xcomanda2

End Function

Private Sub Label7_Click()
    dlo922_Click

End Sub

Function verifica_combo(buf As String)

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM _c" & gusuario & " where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        verifica_combo = 1

    End If

    mytablex.Close

End Function

Private Sub Label9_Click()

    Dim found As Integer

    If Len(comanda) = 0 Then
        comanda = "1"
        xcomanda2 = "1"

    End If

    xcomanda2 = busca_nroc()
    comanda = xcomanda2

    If Len(Trim(salon)) = 0 Then
        MsgBox "No ha seleccionado Salon", 24, "Aviso"
        Exit Sub

    End If

    If Len(Trim(mesa)) = 0 Then
        MsgBox "No ha seleccionado Mesa", 24, "Aviso"
        Exit Sub

    End If

    If Len(Trim(mesero)) = 0 Then
        MsgBox "No ha seleccionado Mesero", 24, "Aviso"
        Exit Sub

    End If

    If Len(Trim(comanda)) = 0 Then
        MsgBox "No ha seleccionado comanda", 24, "Aviso"
        Exit Sub

    End If
                        
    found = grabar_comanda()

    If found = 0 Then
        MsgBox "No existen datos ", 48, "Aviso"
        Exit Sub

    End If

    tptovta.salon.Caption = salon
    tptovta.mesa.Caption = mesa
    tptovta.mesero = mesero
    flag_comanda = 1
    tptovta.comanda = comanda
    imprime_precuenta
    tcomanda.Hide
    Unload tcomanda
    'tservicio = "C"
    'found = orden_despacho()
    'mytable11.Edit
    'mytable11.Fields("comandanro") = xcomanda
    'mytable11.Update
    'si todo sale bien debe inicializar en el load
    'del pedido todo en cero
                        
    Exit Sub

End Sub

Sub imprime_precuenta()

    Dim found As Integer

    If Len(Trim(salon)) = 0 Then
        MsgBox "No ha seleccionado Salon", 24, "Aviso"
        Exit Sub

    End If

    If Len(Trim(mesa)) = 0 Then
        MsgBox "No ha seleccionado Mesa", 24, "Aviso"
        Exit Sub

    End If

    found = sumar_destadocuenta("" & salon, "" & mesa)

    If found = 0 Then
        MsgBox "No se pudo imprimir ", 48, "Aviso"
        Exit Sub

    End If

    formato_precuenta "" & salon, "" & mesa

End Sub

Sub pone_nombre_salon(buf As String)

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close

    mytablex.Open "SELECT * FROM salon where salon='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe

        'Label2.Caption = "" & mytablex.Fields("descripcio")
    End If

    mytablex.Close

End Sub

Function mismo_mesero() As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM parame where  codigo='01'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        mismo_mesero = Trim("" & mytablex.Fields("mesavendedor"))

    End If

    mytablex.Close

End Function

Function phabilita_mesero() As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM parame where  codigo='01'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        phabilita_mesero = Trim("" & mytablex.Fields("selemesa"))

    End If

    mytablex.Close

End Function

'aqui viene lo nuevo de comandas
Sub orden13012015()

    Dim found           As Integer

    Dim archivo_formato As String

    Dim archivo_orden   As String

    Dim mytablex        As New ADODB.Recordset

    Dim mytabley        As New ADODB.Recordset

    Dim Tmp             As String

    Dim sw              As Integer

    Dim xprinter        As String

    sw = 0
    Tmp = ""
    'borrar_todo
    'sql_detalle9
    archivo_formato = globaldir & "\formatos\despacho"
    archivo_orden = globaldir & "\temporal\" & gusuario & ".txt"
    borrar_archivo archivo_orden
    mytabley.Open "SELECT * FROM dcomanda where salon='" & "" & salon & "' and mesa='" & mesa & "' and numero='" & xcomanda2 & "'", cn, adOpenKeyset, adLockOptimistic
    xprinter = ""
    Do

        If mytabley.EOF Then Exit Do
        If "" & mytabley.Fields("dua") = "R" Then GoTo kkordenx
        If sw = 0 Then
            found = proceso_formatoso(mytabley, archivo_formato, archivo_orden, "{", "}")
            Tmp = "" & mytabley.Fields("familia")
            xprinter = busca_printer_familia(mytabley)
            sw = 1

        End If

        If Tmp <> "" & mytabley.Fields("familia") Then
            found = imprime_archivo_ordenc(mytabley, xprinter)
            xprinter = busca_printer_familia(mytabley)
            Tmp = "" & mytabley.Fields("familia")
            borrar_archivo archivo_orden
            found = proceso_formatoso(mytabley, archivo_formato, archivo_orden, "{", "}")

        End If

        found = proceso_formatoso(mytabley, archivo_formato, archivo_orden, "/", "\")
kkordenx:
        mytabley.MoveNext
    Loop
    found = imprime_archivo_ordenc(mytabley, xprinter)
    mytabley.Close

End Sub

Function busca_printer_familia(mytablex As ADODB.Recordset) As String

    Dim mytabley As New ADODB.Recordset

    mytabley.Open "SELECT * FROM familia where familia='" & Trim("" & mytablex.Fields("familia")) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        busca_printer_familia = Trim("" & mytabley.Fields("red"))

    End If

    mytabley.Close

End Function

Function imprime_archivo_ordenc(mytablex As ADODB.Recordset, xbuf1 As String)

    Dim sFile As String

    Dim found As Integer

    Dim oldprinter

    Dim mytabley As New ADODB.Recordset

    sFile = globaldir & "\temporal\" & gusuario & ".txt"
    oldprinter = Printer.DeviceName
    selecciona_impresoras (Trim(xbuf1))
    found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanorden"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))
    selecciona_impresoras (Trim(oldprinter))

End Function

'''' 25/07/2018 Delivery y Para Llevar desde mozo
Function graba_cliente_tipo(buf As String)

    Dim mytablex  As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    Dim sdx       As Double

    Dim buf1      As String

    Dim codigogen As String

    On Error GoTo cmdd7812_err

    If Len(tptovta.nombre) > 0 Then 'no no tiene codigo
   
        mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Function

        End If

        sdx = Val("" & mytablex.Fields("clientes")) + 1
        codigogen = "" & sdx
        mytablex.Close
sigueb1:
        mytablex.Open "select * from clientes where codigo='" & codigogen & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            mytablex.Close
            sdx = sdx + 1
            codigogen = "" & sdx
            GoTo sigueb1

        End If

        tptovta.codigo = codigogen
        mytablex.AddNew
        mytablex.Fields("codigo") = "" & tptovta.codigo
        mytablex.Fields("tipo") = "O"
        mytablex.Fields("nombre") = "" & tptovta.nombre
        mytablex.Fields("direccion") = "" & tptovta.xdireccion
        mytablex.Fields("correo") = Trim("" & tptovta.correo)
        mytablex.Fields("estado") = "A"
        mytablex.Fields("moneda") = "S"
        mytablex.Update
        tptovta.codigo = "" & mytablex.Fields("codigo")
        mytablex.Close
        Exit Function

    End If

    If Len(tptovta.nombre) > 0 Then
        mytablex.Open "SELECT * FROM clientes  where  codigo='" & tptovta.codigo & "'", cn, adOpenDynamic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            mytablex.Fields("nombre") = Trim("" & tptovta.nombre)

            If Len(Trim("" & tptovta.xdireccion)) > 0 Then
                mytablex.Fields("direccion") = Trim("" & tptovta.xdireccion)

            End If

            If Len(Trim("" & tptovta.correo)) > 0 Then
                mytablex.Fields("correo") = Trim("" & tptovta.correo)

            End If

            mytablex.Update
        Else
            mytablex.AddNew
            mytablex.Fields("nombre") = "" & tptovta.nombre
            mytablex.Fields("codigo") = "" & tptovta.codigo

            If tptovta.xtipo = "2" Or tptovta.xtipo = "4" Then
                mytablex.Fields("tipo") = "J"
            Else
                mytablex.Fields("tipo") = "O"

            End If

            If Len("" & tptovta.xdireccion) > 0 Then
                mytablex.Fields("direccion") = "" & tptovta.xdireccion

            End If

            If Len(Trim("" & tptovta.correo)) > 0 Then
                mytablex.Fields("correo") = Trim("" & tptovta.correo)

            End If

            mytablex.Update

        End If

        mytablex.Close

    End If

    Exit Function
cmdd7812_err:
    MsgBox "Aviso en graba cliente tipo " + error$, 48, "Aviso"
    Exit Function

End Function

'''' 25/07/2018 Delivery y Para Llevar desde mozo

'24/08/2018  Delivery por mesa
Function extrae_datos_comanda(tipo As Integer, salon As String, mesa As String)

    Dim mytablex As New ADODB.Recordset

    If Len(salon) > 0 And Len(mesa) > 0 Then
        mytablex.Open "select top 1 * from dcomanda where salon='" & salon & "' and mesa='" & mesa & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            If tipo = 1 Then
                extrae_datos_comanda = Trim("" & mytablex.Fields("codigo"))
            ElseIf tipo = 2 Then
                extrae_datos_comanda = Trim("" & mytablex.Fields("nombre"))

            End If
      
        End If

        mytablex.Close

    End If

End Function

'24/08/2018  Delivery por mesa

