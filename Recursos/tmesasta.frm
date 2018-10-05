VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form tmesasta 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Salones y Mesas"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "tmesasta.frx":0000
   ScaleHeight     =   10200
   ScaleWidth      =   14460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton image4 
      BackColor       =   &H00808080&
      Caption         =   "&PedidoMesa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   11400
      MaskColor       =   &H80000007&
      Picture         =   "tmesasta.frx":57F1E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   300
      Width           =   1995
   End
   Begin ChamaleonButton.ChameleonBtn groupsalon 
      Height          =   855
      Index           =   0
      Left            =   3360
      TabIndex        =   9
      Top             =   840
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
      MICON           =   "tmesasta.frx":58494
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
      Left            =   5350
      TabIndex        =   10
      Top             =   825
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
      MICON           =   "tmesasta.frx":584B0
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
      Left            =   7350
      TabIndex        =   11
      Top             =   840
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
      MICON           =   "tmesasta.frx":584CC
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
      Left            =   9360
      TabIndex        =   12
      Top             =   840
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
      MICON           =   "tmesasta.frx":584E8
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
      Left            =   825
      TabIndex        =   13
      Top             =   2925
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
      MICON           =   "tmesasta.frx":58504
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
      Left            =   825
      TabIndex        =   14
      Top             =   4080
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
      MICON           =   "tmesasta.frx":58520
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
      Left            =   825
      TabIndex        =   15
      Top             =   5235
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
      MICON           =   "tmesasta.frx":5853C
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
      Left            =   825
      TabIndex        =   16
      Top             =   6390
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
      MICON           =   "tmesasta.frx":58558
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
      Left            =   2940
      TabIndex        =   17
      Top             =   2925
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
      MICON           =   "tmesasta.frx":58574
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
      Left            =   2940
      TabIndex        =   18
      Top             =   4080
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
      MICON           =   "tmesasta.frx":58590
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
      Left            =   2940
      TabIndex        =   19
      Top             =   5235
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
      MICON           =   "tmesasta.frx":585AC
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
      Left            =   2940
      TabIndex        =   20
      Top             =   6390
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
      MICON           =   "tmesasta.frx":585C8
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
      Left            =   5070
      TabIndex        =   21
      Top             =   2925
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
      MICON           =   "tmesasta.frx":585E4
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
      Left            =   5070
      TabIndex        =   22
      Top             =   4080
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
      MICON           =   "tmesasta.frx":58600
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
      Left            =   5070
      TabIndex        =   23
      Top             =   5235
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
      MICON           =   "tmesasta.frx":5861C
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
      Left            =   5070
      TabIndex        =   24
      Top             =   6390
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
      MICON           =   "tmesasta.frx":58638
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
      Left            =   7185
      TabIndex        =   25
      Top             =   2925
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
      MICON           =   "tmesasta.frx":58654
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
      Left            =   7185
      TabIndex        =   26
      Top             =   4080
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
      MICON           =   "tmesasta.frx":58670
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
      Left            =   7185
      TabIndex        =   27
      Top             =   5235
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
      MICON           =   "tmesasta.frx":5868C
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
      Left            =   7185
      TabIndex        =   28
      Top             =   6390
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
      MICON           =   "tmesasta.frx":586A8
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
      Left            =   9300
      TabIndex        =   29
      Top             =   2925
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
      MICON           =   "tmesasta.frx":586C4
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
      Left            =   9300
      TabIndex        =   30
      Top             =   4080
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
      MICON           =   "tmesasta.frx":586E0
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
      Left            =   9300
      TabIndex        =   31
      Top             =   5235
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
      MICON           =   "tmesasta.frx":586FC
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
      Left            =   9300
      TabIndex        =   32
      Top             =   6390
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
      MICON           =   "tmesasta.frx":58718
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
      Left            =   11430
      TabIndex        =   33
      Top             =   2925
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
      MICON           =   "tmesasta.frx":58734
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
      Left            =   11430
      TabIndex        =   34
      Top             =   4080
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
      MICON           =   "tmesasta.frx":58750
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
      Left            =   11430
      TabIndex        =   35
      Top             =   5235
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
      MICON           =   "tmesasta.frx":5876C
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
      Left            =   11400
      TabIndex        =   36
      Top             =   6360
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
      MICON           =   "tmesasta.frx":58788
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn CmdCerrar 
      Height          =   855
      Left            =   11445
      TabIndex        =   38
      Top             =   1290
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
      MICON           =   "tmesasta.frx":587A4
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
      Left            =   840
      TabIndex        =   39
      Top             =   7560
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
      MICON           =   "tmesasta.frx":587C0
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
      Left            =   2955
      TabIndex        =   40
      Top             =   7560
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
      MICON           =   "tmesasta.frx":587DC
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
      Left            =   5085
      TabIndex        =   41
      Top             =   7560
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
      MICON           =   "tmesasta.frx":587F8
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
      Left            =   7200
      TabIndex        =   42
      Top             =   7560
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
      MICON           =   "tmesasta.frx":58814
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
      Left            =   9315
      TabIndex        =   43
      Top             =   7560
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
      MICON           =   "tmesasta.frx":58830
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
      Left            =   11445
      TabIndex        =   44
      Top             =   7560
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
      MICON           =   "tmesasta.frx":5884C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Left            =   11040
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label xindex 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Left            =   10605
      TabIndex        =   37
      Top             =   1950
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "              LIBRE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8055
      TabIndex        =   6
      Top             =   9525
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   8055
      Top             =   9045
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "              OCUPADO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6735
      TabIndex        =   5
      Top             =   9525
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   6735
      Top             =   9045
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRECUENTA EMITIDA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5415
      TabIndex        =   4
      Top             =   9525
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   5415
      Top             =   9045
      Width           =   1335
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   3960
      Picture         =   "tmesasta.frx":58868
      Stretch         =   -1  'True
      Top             =   2025
      Width           =   1320
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   9180
      Picture         =   "tmesasta.frx":5A876
      Stretch         =   -1  'True
      Top             =   2025
      Width           =   1320
   End
   Begin VB.Image Image6 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   4560
      Picture         =   "tmesasta.frx":5CCD0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1680
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   8400
      Picture         =   "tmesasta.frx":5ECDE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label mesa 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Left            =   1815
      TabIndex        =   3
      Top             =   1350
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label salon 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Left            =   885
      TabIndex        =   2
      Top             =   1350
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ESTADO DE LAS MESAS"
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
      Height          =   735
      Left            =   5265
      TabIndex        =   1
      Top             =   2025
      Width           =   3945
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALON"
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
      Height          =   735
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   2160
   End
   Begin VB.Menu dlo922 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tmesasta"
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

Private Sub Command3d1_Click(Index As Integer)

    If Index = 10 Then
        'xcomanda = ""
        Exit Sub

    End If

    'xcomanda = xcomanda & Command3d1(Index).Caption
    'comanda = xcomanda

End Sub

Private Sub cmdCerrar_Click()
    dlo922_Click

End Sub

Private Sub Command8_Click()

End Sub

Private Sub dlo922_Click()
    tmesasta.Hide
    Unload tmesasta

End Sub

Private Sub Form_Load()

    Dim I As Integer

    carga_salon

    For I = 0 To 2

        If Trim(groupsalon(I).Caption) = Trim("" & mytable11.Fields("salon")) Then
            groupsalon_Click I
            Exit For

        End If

    Next I

    '
    'groupsalon_Click 0
    Label2 = " Estado " & glomesa

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

Private Sub groupmesa_Click(Index As Integer)

    Dim I As Integer

    Dim k As Integer

    'For i = 0 To 23
    '    groupmesa(i).BackColor = &HFFFFFF
    'Next i
    If Len(groupmesa(Index).Caption) = 0 Then Exit Sub
    'groupmesa(Index).BackColor = &HFF&

    '---------------------------------------
    ''''20/01/2018 kenyo Testing General Sistema
    'If groupsalon(0).BackColor = &HFFFFFF And groupsalon(1).BackColor = &HFFFFFF And groupsalon(2).BackColor = &HFFFFFF Then
    If groupsalon(0).BackColor = &HFFFFFF And groupsalon(1).BackColor = &HFFFFFF And groupsalon(2).BackColor = &HFFFFFF And groupsalon(3).BackColor = &HFFFFFF Then
        ''''20/01/2018 kenyo Testing General Sistema
        Exit Sub

    End If

    k = 0

    If groupsalon(0).BackColor = &HFF& Then
        k = 0

    End If

    If groupsalon(1).BackColor = &HFF& Then
        k = 1

    End If

    'MsgBox "Hola"

    '''' 25/07/2018 Delivery y Para Llevar desde mozo
    ' Mesas Nombre 21/05/2018
    'mesa = groupmesa(Index).Caption
    mesa = extra_loquesea(groupmesa(Index).Caption)
    ' Mesas Nombre 21/05/2018
    '''' 25/07/2018 Delivery y Para Llevar desde mozo

    If Len(salon) = 0 Then
        MsgBox "Seleccione un Salon ", 48, "Aviso"
        Exit Sub

    End If

    If Len(mesa) = 0 Then
        MsgBox "Seleccione una mesa ", 48, "Aviso"
        Exit Sub

    End If

    If image4.BackColor = &HFF8080 Then
        'MsgBox "Selecciono mesa " & mesa
        tptovta.salon.Caption = Trim(salon)
        tptovta.mesa.Caption = Trim(mesa)
        tmesasta.Hide
        Unload tmesasta
        Exit Sub

    End If

    '''' 25/07/2018 Delivery y Para Llevar desde mozo
    ' Mesas Nombre 21/05/2018
    'TBRCOMA.salon = salon
    'TBRCOMA.mesa = mesa
    TBRCOMA.salon = Trim(salon)
    TBRCOMA.mesa = Trim(mesa)
    ' Mesas Nombre 21/05/2018
    '''' 25/07/2018 Delivery y Para Llevar desde mozo

    TBRCOMA.Show 1
    tmesasta.Hide
    Unload tmesasta

    'carga_salon
    'groupsalon_Click 0

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

    mesa = ""

End Sub

Sub verifica_mesas(indx As Integer, buf As String, buf1 As String)

    '''' 25/07/2018 Delivery y Para Llevar desde mozo
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
   
            '''' 25/07/2018 Delivery y Para Llevar desde mozo
            ' Mesas Nombre 21/05/2018
            'mmesacod(I) = "" & mytablex.Fields("mesa")
            'wmesacod(I) = "" & mytablex.Fields("mesa")
            Dim rpta As String

            rpta = ""
            mmesacod(I) = " " & mytablex.Fields("mesa")
            wmesacod(I) = "" & mytablex.Fields("mesa")
   
            '24/08/2018  Delivery por mesa
            If IsNull(Trim(mytablex.Fields("codigo"))) = False Then
                'rpta = busca_clienteDelivery_mesa(mytablex.Fields("codigo"), buf)
                rpta = "" & Mid$(mytablex.Fields("dnombre"), 1, 25)

            End If

            '24/08/2018  Delivery por mesa

            If Len(rpta) > 0 Then
                mmesacod(I) = mmesacod(I) & "| " & rpta

            End If
   
            ' Mesas Nombre 21/05/2018
            '''' 25/07/2018 Delivery y Para Llevar desde mozo
   
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
        ' verifica_mesas j, buf1, groupmesa(j).Caption
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

Private Sub Image4_Click()

    If image4.BackColor = &HFFFFFF Then
        image4.BackColor = &HFF8080
        Exit Sub

    End If

    If image4.BackColor = &HFF8080 Then
        image4.BackColor = &HFFFFFF
        Exit Sub

    End If

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

Private Sub lblVENDEDOR_Click()

End Sub

