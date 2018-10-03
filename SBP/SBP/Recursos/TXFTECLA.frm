VERSION 5.00
Begin VB.Form TXFTECLA 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignacion Teclas"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   16290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   16290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label codigo 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   74
      Top             =   6480
      Width           =   105
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   36
      Left            =   8760
      TabIndex        =   73
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   35
      Left            =   7320
      TabIndex        =   72
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   34
      Left            =   5880
      TabIndex        =   71
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   33
      Left            =   4440
      TabIndex        =   70
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   32
      Left            =   3000
      TabIndex        =   69
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   31
      Left            =   1560
      TabIndex        =   68
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   30
      Left            =   120
      TabIndex        =   67
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   29
      Left            =   13080
      TabIndex        =   66
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   28
      Left            =   11640
      TabIndex        =   65
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   27
      Left            =   10200
      TabIndex        =   64
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   26
      Left            =   8760
      TabIndex        =   63
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   25
      Left            =   7320
      TabIndex        =   62
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   24
      Left            =   5880
      TabIndex        =   61
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   23
      Left            =   4440
      TabIndex        =   60
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   22
      Left            =   3000
      TabIndex        =   59
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   21
      Left            =   1560
      TabIndex        =   58
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   20
      Left            =   120
      TabIndex        =   57
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   19
      Left            =   13080
      TabIndex        =   56
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   18
      Left            =   11640
      TabIndex        =   55
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   17
      Left            =   10200
      TabIndex        =   54
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   16
      Left            =   8760
      TabIndex        =   53
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   15
      Left            =   7320
      TabIndex        =   52
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   14
      Left            =   5880
      TabIndex        =   51
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   13
      Left            =   4440
      TabIndex        =   50
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   12
      Left            =   3000
      TabIndex        =   49
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   11
      Left            =   1560
      TabIndex        =   48
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   10
      Left            =   120
      TabIndex        =   47
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   9
      Left            =   13080
      TabIndex        =   46
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   8
      Left            =   11640
      TabIndex        =   45
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   7
      Left            =   10200
      TabIndex        =   44
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   6
      Left            =   8760
      TabIndex        =   43
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   5
      Left            =   7320
      TabIndex        =   42
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   4
      Left            =   5880
      TabIndex        =   41
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   3
      Left            =   4440
      TabIndex        =   40
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   2
      Left            =   3000
      TabIndex        =   39
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   1
      Left            =   1560
      TabIndex        =   38
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label DESCRIPCIO 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   36
      Left            =   8760
      TabIndex        =   36
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   35
      Left            =   7320
      TabIndex        =   35
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   34
      Left            =   5880
      TabIndex        =   34
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   33
      Left            =   4440
      TabIndex        =   33
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   32
      Left            =   3000
      TabIndex        =   32
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   31
      Left            =   1560
      TabIndex        =   31
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   30
      Left            =   120
      TabIndex        =   30
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ñ"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   29
      Left            =   13080
      TabIndex        =   29
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   28
      Left            =   11640
      TabIndex        =   28
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   27
      Left            =   10200
      TabIndex        =   27
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   26
      Left            =   8760
      TabIndex        =   26
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   25
      Left            =   7320
      TabIndex        =   25
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   24
      Left            =   5880
      TabIndex        =   24
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   23
      Left            =   4440
      TabIndex        =   23
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   22
      Left            =   3000
      TabIndex        =   22
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   21
      Left            =   1560
      TabIndex        =   21
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   20
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   19
      Left            =   13080
      TabIndex        =   19
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   18
      Left            =   11640
      TabIndex        =   18
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   17
      Left            =   10200
      TabIndex        =   17
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   16
      Left            =   8760
      TabIndex        =   16
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   15
      Left            =   7320
      TabIndex        =   15
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   5880
      TabIndex        =   14
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   13
      Left            =   4440
      TabIndex        =   13
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   12
      Left            =   3000
      TabIndex        =   12
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   11
      Left            =   1560
      TabIndex        =   11
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   13080
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   11640
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   10200
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   8760
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   7320
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   5880
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label tecla 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu flo444 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "TXFTECLA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub flo444_Click()
    TXFTECLA.Hide
    Unload TXFTECLA

End Sub

Private Sub Form_Load()
    carga_valores

End Sub

Sub carga_valores()

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    For I = 0 To 36
        descripcio(I).Caption = ""
    Next I

    For I = 0 To 36

        If mytablex.State = 1 Then
            mytablex.Close

        End If

        mytablex.Open "select descripcio from producto where tecla='" & tecla(I) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            descripcio(I).Caption = Trim("" & mytablex.Fields("descripcio"))

        End If

        mytablex.Close
    Next I

End Sub
