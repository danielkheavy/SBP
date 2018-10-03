VERSION 5.00
Begin VB.Form HOTOUCH 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Habitaciones"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   13320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Menu de Opciones"
      Height          =   8175
      Left            =   120
      TabIndex        =   41
      Top             =   120
      Visible         =   0   'False
      Width           =   13095
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFF80&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H008080FF&
         Caption         =   "Abonos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H000080FF&
         Caption         =   "Servicios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CheckOut"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FF00&
         Caption         =   "Facturacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0000FF00&
         Caption         =   "Precuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "CheckIn"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Reserva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label habitacion 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6480
         TabIndex        =   48
         Top             =   4080
         Width           =   105
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7575
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   6015
      End
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0FFFF&
      Height          =   855
      Index           =   8
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7440
      Width           =   1095
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0FFFF&
      Height          =   855
      Index           =   7
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0FFFF&
      Height          =   855
      Index           =   6
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   29
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   28
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   27
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   26
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   25
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   24
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   23
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   22
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   21
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   20
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   19
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   18
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   17
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   16
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   15
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   14
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   13
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   12
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   11
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   10
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   9
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   8
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   7
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   6
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   5
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   4
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   3
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0FFFF&
      Height          =   615
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0FFFF&
      Height          =   615
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0FFFF&
      Height          =   615
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Index           =   3
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0FFFF&
      Height          =   855
      Index           =   4
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0FFFF&
      Height          =   855
      Index           =   5
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   12120
      Picture         =   "hotouch.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   12720
      Picture         =   "hotouch.frx":1BD2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HABITACIONES"
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
      Left            =   1200
      TabIndex        =   25
      Top             =   120
      Width           =   10935
   End
   Begin VB.Image Image6 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   120
      Picture         =   "hotouch.frx":3B78
      Stretch         =   -1  'True
      Top             =   840
      Width           =   600
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   720
      Picture         =   "hotouch.frx":574A
      Stretch         =   -1  'True
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PISO"
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
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   1095
   End
   Begin VB.Menu fl9444 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "HOTOUCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mfamcod(15000) As String
Dim mfampag As Integer
Dim mfamtop As Integer

Dim mprodcod(15000) As String
Dim wprodcod(15000) As String
Dim wwprodcod(30) As String
Dim mprodpag As Integer
Dim mprodtop As Integer

Private Sub Command1_Click()
treserva.buffer = habitacion
treserva.buffer.Enabled = False
treserva.Show 1
End Sub

Private Sub Command2_Click()
Dim mytablex As New adodb.Recordset
mytablex.Open "select * from habitacion where habitacion='" & Trim(habitacion) & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      If Trim("" & mytablex.Fields("estado")) = "OCUPADO" Then
         MsgBox "Habitacion Ocupada ", 48, "Aviso"
         mytablex.Close
         Exit Sub
      End If
      Else
      Exit Sub
   End If
tcheckin.buffer = Trim(habitacion)
mytablex.Close
tcheckin.buffer.Enabled = False
tcheckin.Show 1
fl9444_Click
End Sub

Private Sub Command3_Click()
Dim mytablex As New adodb.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM habitacion where habitacion='" & Trim(habitacion) & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If Trim("" & mytablex.Fields("estado")) <> "OCUPADO" Then
      MsgBox "Habitacion libre ", 48, "Aviso"
      mytablex.Close
      Exit Sub
   End If
   Else
   Exit Sub
End If
tcheckin.buffer = Trim("" & habitacion)
mytablex.Close

tcheckin.ajdu1.Enabled = False
tcheckin.f8443.Enabled = False
tcheckin.bo712.Enabled = False
tcheckin.dk8833.Visible = True
tcheckin.Reporte = "PRECUENTA"
tcheckin.buffer.Enabled = False
tcheckin.Show 1

End Sub

Private Sub Command4_Click()
Dim mytablex As New adodb.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM HABITACION where habitacion='" & Trim(habitacion) & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount = 0 Then Exit Sub
   If Trim("" & mytablex.Fields("estado")) <> "OCUPADO" Then
      MsgBox "Habitacion libre,No se puede facturar ", 48, "Aviso"
      mytablex.Close
      Exit Sub
   End If
   mytablex.Close
mytablex.Open "SELECT * FROM checkin where habitacion='" & Trim(habitacion) & "' and estado<>'2'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   mytablex.Close
   MsgBox "No existe Check IN ", 48, "Aviso"
   Exit Sub
End If
tfactura.ridcheckin = Trim("" & mytablex.Fields("idcheckin"))
mytablex.Close
tfactura.rhabita = Trim(habitacion)
tfactura.rhabita.Enabled = False
tfactura.Show 1
End Sub

Private Sub Command5_Click()
Dim mytablex As New adodb.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM checkin where habitacion='" & Trim(habitacion) & "' and estado<>'2'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   MsgBox "No existe Check In ", 48, "Aviso"
   mytablex.Close
   Exit Sub
End If
tcheckin.buffer = Trim(habitacion)
mytablex.Close

tcheckin.ajdu1.Enabled = False
tcheckin.f8443.Enabled = False
tcheckin.bo712.Enabled = False
tcheckin.dk8833.Visible = True
tcheckin.Caption = "Salida Huesped(Check Out) "
'tcheckin.Frame5.Visible = True
tcheckin.Reporte = "CHECKOUT"
tcheckin.buffer.Enabled = False
tcheckin.dk8833.Caption = "CheckOut"
tcheckin.Show 1
fl9444_Click



End Sub

Private Sub Command6_Click()
Dim mytablex As New adodb.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM checkin where habitacion='" & Trim(habitacion) & "' and estado<>'2'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   MsgBox "No existe Check In ", 48, "Aviso"
   mytablex.Close
   Exit Sub
End If
tcaresta.buffer = "" & mytablex.Fields("idcheckin")
mytablex.Close
tcaresta.buffer.Enabled = False
tcaresta.Show 1
End Sub

Private Sub Command7_Click()
Dim mytablex As New adodb.Recordset


If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM checkin where habitacion='" & Trim(habitacion) & "' and estado<>'2'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   MsgBox "No existe Check In ", 48, "Aviso"
   mytablex.Close
   Exit Sub
End If
trecibo.buffer = "" & mytablex.Fields("idcheckin")
mytablex.Close
trecibo.buffer.Enabled = False
trecibo.Show 1
End Sub

Private Sub Command8_Click()
fl9444_Click
End Sub

Private Sub fl9444_Click()
If Frame1.Visible = True Then
   Frame1.Visible = False
   'hotouch
   'Exit Sub
End If
HOTOUCH.Hide
Unload HOTOUCH
End Sub

Private Sub Form_Load()
carga_familia
End Sub
Sub carga_familia()
Dim mytablex As New adodb.Recordset
Dim i As Integer
For i = 0 To 14999
    mfamcod(i) = ""
Next i

i = -1
mytablex.Open "select * from piso ", cn, adOpenStatic, adLockOptimistic

Do
If mytablex.EOF Then Exit Do
i = i + 1
mfamcod(i) = "" & mytablex.Fields("piso")
mytablex.MoveNext
Loop
mfamtop = i
mytablex.Close
mfampag = 0
menu_familia "INI"

End Sub

Sub menu_familia(buf As String)
Dim i As Integer
Dim j As Integer
Select Case buf
       Case "INI"
            mfampag = 0
       Case "SIG"
            mfampag = mfampag + 7
            If mfampag > 102 Then
               mfampag = 0
            End If
       Case "ANT"
            mfampag = mfampag - 7
            If mfampag < 0 Then
               mfampag = 0
            End If
End Select
j = -1
For i = mfampag To 7 + mfampag
    j = j + 1
    zfamilia(j).Caption = mfamcod(i)
Next i

End Sub

Private Sub Image1_Click()
menu_producto "SIG"
End Sub

Private Sub Image2_Click()
menu_producto "ANT"
End Sub

Private Sub Image5_Click()
menu_familia "SIG"
End Sub

Private Sub Image6_Click()
menu_familia "ANT"
End Sub

Private Sub zfamilia_Click(Index As Integer)
menu_carga_producto zfamilia(Index).Caption
menu_producto "INI"

End Sub
Sub menu_carga_producto(buf As String)
Dim mytablex As New adodb.Recordset
Dim buf1 As String

Dim i As Integer
For i = 0 To 28
   wwprodcod(i) = ""
Next i
For i = 0 To 14999
    mprodcod(i) = ""
    wprodcod(i) = ""
Next i

i = -1

If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM habitacion where piso='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("piso") = buf Then
   i = i + 1
   buf1 = "LIBRE"
   If Trim("" & mytablex.Fields("estado")) = "OCUPADO" Then
   buf1 = "OCUPADO"
   End If
   mprodcod(i) = "" & mytablex.Fields("habitacion") & " " & Chr$(10) & Chr$(13) & "" & mytablex.Fields("tipohabitacion") & Chr$(10) & Chr$(13) & buf1 & Chr$(10) & Chr$(13) & "S/." & mytablex.Fields("precio")
   wprodcod(i) = "" & mytablex.Fields("habitacion")
   Else: Exit Do
End If
mytablex.MoveNext
Loop

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
            mprodpag = mprodpag + 28
            If mprodpag > 102 Then
               mprodpag = 0
            End If
       Case "ANT"
            mprodpag = mprodpag - 28
            If mprodpag < 0 Then
               mprodpag = 0
            End If
End Select
j = -1
For i = mprodpag To 28 + mprodpag
    j = j + 1
    zproducto(j).Caption = mprodcod(i)
    wwprodcod(j) = wprodcod(i)
Next i


End Sub


Private Sub zproducto_Click(Index As Integer)
Dim buf As String
   buf = "" & wwprodcod(Index)
   habitacion = Trim("" & wwprodcod(Index))
   If Len(buf) = 0 Then Exit Sub
   Frame1.Visible = True
   buf = "" & mprodcod(Index)
   Label3 = buf
End Sub
