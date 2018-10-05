VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form TBRCOMA 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visualizar Comandas"
   ClientHeight    =   10245
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   15435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Height          =   8295
      Left            =   15360
      TabIndex        =   43
      Top             =   2280
      Visible         =   0   'False
      Width           =   15090
      Begin VB.ComboBox salonf 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   99
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   4440
         Picture         =   "TBRCOMA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   6000
         Width           =   1470
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Seleccione Cantidad a Mover"
         Height          =   4935
         Left            =   7440
         TabIndex        =   69
         Top             =   720
         Visible         =   0   'False
         Width           =   6255
         Begin VB.Label mcantidadm 
            BackColor       =   &H00E0E0E0&
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
            Height          =   615
            Left            =   1800
            TabIndex        =   88
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            TabIndex        =   87
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   1
            Left            =   1320
            TabIndex        =   86
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   2
            Left            =   2520
            TabIndex        =   85
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   3
            Left            =   3720
            TabIndex        =   84
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   4
            Left            =   120
            TabIndex        =   83
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   5
            Left            =   1320
            TabIndex        =   82
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   6
            Left            =   2520
            TabIndex        =   81
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   7
            Left            =   3720
            TabIndex        =   80
            Top             =   3360
            Width           =   1215
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   8
            Left            =   120
            TabIndex        =   79
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   9
            Left            =   1320
            TabIndex        =   78
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "BR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   10
            Left            =   2520
            TabIndex        =   77
            Top             =   4080
            Width           =   1215
         End
         Begin VB.Label Label15 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Producto"
            Height          =   495
            Left            =   120
            TabIndex        =   76
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label17 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Descripcio"
            Height          =   495
            Left            =   120
            TabIndex        =   75
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label Label19 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cantidad"
            Height          =   495
            Left            =   120
            TabIndex        =   74
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label20 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cantidad Mover"
            Height          =   615
            Left            =   120
            TabIndex        =   73
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label mproducto 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   1800
            TabIndex        =   72
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label mdescripcio 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   1800
            TabIndex        =   71
            Top             =   840
            Width           =   4335
         End
         Begin VB.Label mcantidad 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   1800
            TabIndex        =   70
            Top             =   1320
            Width           =   1695
         End
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   3720
         Width           =   855
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   4440
         Width           =   855
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   5160
         Width           =   855
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   3720
         Width           =   855
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   4440
         Width           =   855
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   5160
         Width           =   855
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   10
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
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
         TabIndex        =   57
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   12
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   13
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   14
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   15
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   16
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   17
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   18
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   19
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   20
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   21
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3720
         Width           =   975
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   22
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   23
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   975
         Left            =   5880
         Picture         =   "TBRCOMA.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Imprimir todo"
         Top             =   6000
         Width           =   1470
      End
      Begin VB.Label mesaf 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   5880
         TabIndex        =   98
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "INICIO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   360
         TabIndex        =   97
         Top             =   720
         Width           =   3375
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   360
         Picture         =   "TBRCOMA.frx":1194
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1320
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   360
         Picture         =   "TBRCOMA.frx":313A
         Stretch         =   -1  'True
         Top             =   3720
         Width           =   1320
      End
      Begin VB.Label saloni 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   2280
         TabIndex        =   96
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label mesai 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   2280
         TabIndex        =   95
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FINAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   3960
         TabIndex        =   94
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SALON"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3960
         TabIndex        =   93
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MESA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3960
         TabIndex        =   92
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SALON"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   91
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MESA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   90
         Top             =   2160
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   5055
      Left            =   3600
      TabIndex        =   27
      Top             =   2280
      Visible         =   0   'False
      Width           =   8865
      Begin VB.TextBox MOTIVO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         IMEMode         =   3  'DISABLE
         Left            =   1890
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   100
         Top             =   1005
         Width           =   3420
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   10
         Left            =   3870
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   9
         Left            =   2910
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   8
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3780
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   7
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2940
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   6
         Left            =   3870
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2940
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   2910
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2940
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   4
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2940
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2100
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   3870
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2100
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   2910
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2100
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   11
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2100
         Width           =   975
      End
      Begin VB.TextBox CLAVE 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1920
         MaxLength       =   11
         PasswordChar    =   "*"
         TabIndex        =   30
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   6705
         Picture         =   "TBRCOMA.frx":4D0C
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Imprimir todo"
         Top             =   3555
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Index           =   12
         Left            =   6555
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   345
         Width           =   1515
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Motivo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   510
         Left            =   120
         TabIndex        =   101
         Top             =   1020
         Width           =   1785
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CLAVE"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   42
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Cantidad que va a Cobrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4425
      Left            =   5280
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   10815
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   5730
         Picture         =   "TBRCOMA.frx":55D6
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2520
         Width           =   1245
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   975
         Left            =   7035
         Picture         =   "TBRCOMA.frx":5EA0
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Imprimir todo"
         Top             =   2520
         Width           =   1245
      End
      Begin VB.TextBox cantdev 
         Height          =   735
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   19
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label DESCRIPCIO 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   2760
         TabIndex        =   26
         Top             =   1320
         Width           =   7695
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PRODUCTO"
         Height          =   735
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label cantidad 
         BackColor       =   &H00E0E0E0&
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
         Height          =   735
         Left            =   2760
         TabIndex        =   24
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CANTIDAD"
         Height          =   735
         Left            =   240
         TabIndex        =   23
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CANTIDAD"
         Height          =   735
         Left            =   240
         TabIndex        =   22
         Top             =   2760
         Width           =   2535
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Descuentos por Producto"
      Height          =   2400
      Left            =   3000
      TabIndex        =   11
      Top             =   2040
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   3810
         Picture         =   "TBRCOMA.frx":676A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   915
         Width           =   1470
      End
      Begin VB.CommandButton Command13 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   5415
         Picture         =   "TBRCOMA.frx":7034
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Imprimir todo"
         Top             =   945
         Width           =   1455
      End
      Begin VB.TextBox desporcentaje 
         Height          =   375
         Left            =   480
         MaxLength       =   10
         TabIndex        =   14
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
         Height          =   375
         Left            =   2520
         TabIndex        =   17
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label desprecio 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label desdescripcio 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   480
         Width           =   6375
      End
   End
   Begin VB.CommandButton btnsalir 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
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
      Height          =   855
      Left            =   14280
      Picture         =   "TBRCOMA.frx":78FE
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimir todo"
      Top             =   9120
      Width           =   1155
   End
   Begin MSDataGridLib.DataGrid table2 
      Height          =   7335
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   12938
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   3
      RowHeight       =   21
      FormatLocked    =   -1  'True
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
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   18
      BeginProperty Column00 
         DataField       =   "Salon"
         Caption         =   "Salon"
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
         DataField       =   "Mesa"
         Caption         =   "Mesa"
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
      BeginProperty Column02 
         DataField       =   "Vendedor"
         Caption         =   "Mozo"
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
      BeginProperty Column03 
         DataField       =   "Comanda"
         Caption         =   "Comanda"
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
      BeginProperty Column04 
         DataField       =   "Precio"
         Caption         =   "Precio"
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
      BeginProperty Column05 
         DataField       =   "cantidad"
         Caption         =   "Cant"
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
      BeginProperty Column06 
         DataField       =   "Cantdev"
         Caption         =   "Cant1"
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
      BeginProperty Column07 
         DataField       =   "Total"
         Caption         =   "Total"
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
      BeginProperty Column08 
         DataField       =   "Descripcio"
         Caption         =   "Descripcio"
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
      BeginProperty Column09 
         DataField       =   "Unidad"
         Caption         =   "Und"
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
      BeginProperty Column10 
         DataField       =   "factor"
         Caption         =   "Fac"
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
      BeginProperty Column11 
         DataField       =   "Deslipo"
         Caption         =   "Deslipo"
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
      BeginProperty Column12 
         DataField       =   "Estado"
         Caption         =   "E"
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
      BeginProperty Column13 
         DataField       =   "Dua"
         Caption         =   "Flag"
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
      BeginProperty Column14 
         DataField       =   "fecha"
         Caption         =   "Fecha"
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
      BeginProperty Column15 
         DataField       =   "Caja"
         Caption         =   "Caja"
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
      BeginProperty Column16 
         DataField       =   "Hora"
         Caption         =   "Hora"
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
      BeginProperty Column17 
         DataField       =   "Destopo"
         Caption         =   "Destopo"
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
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   524.976
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   659.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   4740.095
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   705.26
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   675.213
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column15 
         EndProperty
         BeginProperty Column16 
         EndProperty
         BeginProperty Column17 
         EndProperty
      EndProperty
   End
   Begin ChamaleonButton.ChameleonBtn cmdReimprimirComanda 
      Height          =   855
      Index           =   0
      Left            =   2280
      TabIndex        =   102
      Top             =   9120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Reimprimir Comanda"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "TBRCOMA.frx":81C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn cmdMoviendoProducto 
      Height          =   855
      Index           =   1
      Left            =   3435
      TabIndex        =   103
      Top             =   9120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Moviendo Producto"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "TBRCOMA.frx":81E4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Command9 
      Height          =   855
      Left            =   4590
      TabIndex        =   104
      Top             =   9120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Moviendo Mesa"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "TBRCOMA.frx":8200
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Command6 
      Height          =   855
      Left            =   5730
      TabIndex        =   105
      Top             =   9120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Juntando Mesa"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "TBRCOMA.frx":821C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Command5 
      Height          =   855
      Left            =   6870
      TabIndex        =   106
      Top             =   9120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Descuento Global"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   32768
      BCOLO           =   49152
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421504
      MPTR            =   1
      MICON           =   "TBRCOMA.frx":8238
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Command15 
      Height          =   855
      Left            =   8010
      TabIndex        =   107
      Top             =   9120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Descuento Producto"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "TBRCOMA.frx":8254
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Command2 
      Height          =   855
      Left            =   9150
      TabIndex        =   108
      Top             =   9120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Avance Cuenta"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "TBRCOMA.frx":8270
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn Command3 
      Height          =   855
      Left            =   11910
      TabIndex        =   109
      Top             =   9120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "Anula Producto"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      MICON           =   "TBRCOMA.frx":828C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
      Height          =   855
      Left            =   13080
      TabIndex        =   111
      Top             =   9120
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "ANULA TODO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   8421504
      MPTR            =   1
      MICON           =   "TBRCOMA.frx":82A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox referencia 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5640
      MaxLength       =   120
      TabIndex        =   112
      Top             =   1200
      Width           =   7455
   End
   Begin VB.TextBox ddireccion 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5640
      MaxLength       =   200
      TabIndex        =   113
      Top             =   720
      Width           =   7455
   End
   Begin VB.TextBox dnombre 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      MaxLength       =   60
      TabIndex        =   114
      Top             =   1200
      Width           =   5415
   End
   Begin VB.TextBox telefono 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   120
      MaxLength       =   11
      TabIndex        =   115
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox dcodigo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5640
      MaxLength       =   60
      TabIndex        =   116
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   120
      TabIndex        =   110
      Top             =   9300
      Width           =   915
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label cfecha 
      BackColor       =   &H00E0E0E0&
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
      Height          =   495
      Left            =   6240
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hora"
      Height          =   495
      Left            =   8040
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label chora 
      BackColor       =   &H00E0E0E0&
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
      Height          =   495
      Left            =   9240
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label txtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   840
      Left            =   960
      TabIndex        =   5
      Top             =   9120
      Width           =   1275
   End
   Begin VB.Label MESA 
      BackColor       =   &H00E0E0E0&
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
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MESA"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label SALON 
      BackColor       =   &H00E0E0E0&
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
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALON"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  'Fixed Single
      Height          =   870
      Left            =   11100
      Picture         =   "TBRCOMA.frx":82C4
      Stretch         =   -1  'True
      Top             =   9105
      Width           =   840
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  'Fixed Single
      Height          =   870
      Left            =   10290
      Picture         =   "TBRCOMA.frx":98B2
      Stretch         =   -1  'True
      Top             =   9105
      Width           =   840
   End
   Begin VB.Menu fl9923 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "TBRCOMA"
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

Dim tbrmytablex     As New ADODB.Recordset

Dim logclave        As String

Private Sub btnsalir_Click()
    fl9923_Click

End Sub

Private Sub ChameleonBtn1_Click()

    '' 22/12/2017 Anulacion completa de comanda
    Frame1.Visible = True
    Frame1.Caption = "ANULA TODO"
    clave = ""
    clave.SetFocus
    '' 22/12/2017 Anulacion completa de comanda

End Sub

Private Sub clave_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    'On Error GoTo cmd891_err
    If KeyAscii <> 13 Then Exit Sub
    If Len(clave) = 0 Then
        clave.SetFocus
        Exit Sub

    End If

    'MsgBox "abc"
    If Frame1.Caption = "JUNTANDO MESA" Then
        found = busca_clave(20) 'si puede anular  borra comanda

        If found = 0 Then
            clave = ""
            clave.SetFocus
            Exit Sub

        End If

        'MsgBox "abc"
        Frame1.Visible = False
        saloni = salon
        mesai = mesa
        salonf.Clear
        salonf.AddItem salon
        'MsgBox "abc"
        carga_salones
        mesaf = ""
        'MsgBox "abc"

        'MsgBox "abc"
        Frame4.Visible = False
        Frame2.Visible = True
        Frame2.Caption = "JUNTAR MESAS"
        menu_carga_mesa Trim("" & saloni)
        menu_mesa "INI", Trim("" & saloni)
        mesaf = ""

        Exit Sub

    End If

    If Frame1.Caption = "MOVIENDO MESA" Then
        found = busca_clave(21) 'si puede anular  borra comanda

        If found = 0 Then
            clave = ""
            clave.SetFocus
            Exit Sub

        End If

        Frame1.Visible = False

        saloni = salon
        mesai = mesa
        salonf.Clear
        salonf.AddItem salon
        carga_salones
        mesaf = ""
        Frame4.Visible = False
        Frame2.Visible = True
        Frame2.Caption = "MOVIENDO MESAS"

        menu_carga_mesa Trim("" & saloni)
        menu_mesa "INI", Trim("" & saloni)
        mesaf = ""

        Exit Sub

    End If

    If Frame1.Caption = "MOVIENDO PRODUCTO" Then
        found = busca_clave(22) 'si puede anular  borra comanda

        If found = 0 Then
            clave = ""
            clave.SetFocus
            Exit Sub

        End If

        Frame1.Visible = False
        saloni = salon
        mesai = mesa
        salonf.Clear
        salonf.AddItem salon
        carga_salones
        mesaf = ""
        Frame4.Visible = True
        Frame2.Visible = True
        'MsgBox "abc"
        menu_carga_mesa Trim("" & saloni)
        menu_mesa "INI", Trim("" & saloni)
        mesaf = ""

        mproducto = "" & tbrmytablex.Fields("producto")
        mdescripcio = "" & tbrmytablex.Fields("descripcio")
        mcantidad = "" & tbrmytablex.Fields("cantidad")
        Frame2.Caption = "MOVIENDO MESAS PRODUCTOS"

        Exit Sub

    End If

    If Frame1.Caption = "ANULA PRODUCTO" Then
        found = busca_clave(3) 'si puede anular  borra comanda
        
        If found = 0 Then
            clave = ""
            clave.SetFocus
            Exit Sub

        End If
    
        If MOTIVO.Text = "" Then
      
            MOTIVO.SetFocus
            Exit Sub

        End If

        If tbrmytablex.RecordCount = 0 Then Exit Sub
        If MsgBox("Desea Borrar " & tbrmytablex.Fields("descripcio"), 1, "Aviso") <> 1 Then Exit Sub
       
        found = graba_logcomanda()
       
        ''12/07/2017 kenyo anular multicomandas
        ' found = imprimir_orden_anula(0)
    
        If mytable11.Fields("multicomanda") = "S" Then
            nroimpresion = 0
            found = imprimir_orden_anula(0)
        
            nroimpresion = 1
            found = imprimir_orden_anula(0)
         
            nroimpresion = 2
            found = imprimir_orden_anula(0)
        
            nroimpresion = 3
            found = imprimir_orden_anula(0)
       
        Else
       
            nroimpresion = 0
            found = imprimir_orden_anula(0)

        End If

        ''12/07/2017 kenyo anular multicomandas
       
        tbrmytablex.Delete
        tbrmytablex.Requery
       
        'table2.Refresh
        'suma_table2
        clave = ""

        ' 25/07/2018 Delivery y Para Llevar desde mozo
        If tbrmytablex.RecordCount = 0 Then
            If V_EstadoMesa = "D" Or V_EstadoMesa = "L" Then
                cn.Execute ("update MESA set dnombre='',codigo='',ddireccion='',telefono='',referencia='' where mesa='" & mesa & "' and salon='" & salon & "'")
                tptovta.codigo = ""
                tptovta.nombre = ""
                dcodigo = ""
                dnombre = ""
                ddireccion = ""
                referencia = ""
                telefono = ""
                
            End If

        End If
       
        ' 25/07/2018 Delivery y Para Llevar desde mozo
       
        Frame1.Visible = False
        suma_comandas
        'table2.SetFocus
        'MsgBox ""
       
        Exit Sub

    End If

    '' 22/12/2017 Anulacion completa de comanda
    If Frame1.Caption = "ANULA TODO" Then
        found = busca_clave(3)
        
        If found = 0 Then
            clave = ""
            clave.SetFocus
            Exit Sub

        End If
    
        If MOTIVO.Text = "" Then
            MOTIVO.SetFocus
            Exit Sub

        End If
       
        If tbrmytablex.RecordCount = 0 Then Exit Sub
        If MsgBox("Desea Borrar Todo", 1, "Aviso") <> 1 Then Exit Sub
       
        '' 22/12/2017 Anulacion completa de comanda
        'found = graba_logcomanda()
        '        Do
        '           If tbrmytablex.EOF Then Exit Do
        '           found = graba_logcomanda()
        '           tbrmytablex.MoveNext
        '        Loop
        '' 22/12/2017 Anulacion completa de comanda
     
        '' 22/12/2017 Anulacion completa de comanda
         
        If mytable11.Fields("multicomanda") = "S" Then
            nroimpresion = 0
            found = imprimir_orden_anulaTodo(0)
            
            nroimpresion = 1
            found = imprimir_orden_anulaTodo(0)
             
            nroimpresion = 2
            found = imprimir_orden_anulaTodo(0)
            
            nroimpresion = 3
            found = imprimir_orden_anulaTodo(0)
        Else
            '            tbrmytablex.MoveFirst
            nroimpresion = 0
            found = imprimir_orden_anulaTodo(0)

        End If
     
        '' 22/12/2017 Anulacion completa de comanda
     
        If tbrmytablex.RecordCount = 0 Then Exit Sub
    
        Dim I As Integer

        cn.Execute ("delete from dcomanda where salon='" & Trim(salon) & "' and mesa='" & Trim(mesa) & "'")
        tbrmytablex.Requery
       
        clave = ""
        Frame1.Visible = False
        suma_comandas
 
        Exit Sub

    End If

    '' 22/12/2017 Anulacion completa de comanda

    If Frame1.Caption = "REIMPRIME" Then
        found = busca_clave(3) 'si puede anular  borra comanda

        If found = 0 Then
            clave = ""
            clave.SetFocus
            Exit Sub

        End If

        If tbrmytablex.RecordCount = 0 Then Exit Sub
        If MsgBox("Desea reimprimir " & tbrmytablex.Fields("descripcio"), 1, "Aviso") <> 1 Then Exit Sub
        found = imprimir_orden_anula(1)
        'found = graba_logcomanda()
        'tbrmytablex.Delete
        tbrmytablex.Requery
        'table2.Refresh
        'suma_table2
        clave = ""
        Frame1.Visible = False
        'table2.SetFocus
        'MsgBox ""
        Exit Sub

    End If

    If Frame1.Caption = "AVANCE CUENTA" Then  'avance cuenta
        found = busca_clave(1)

        If found = 0 Then
            clave = ""
            clave.SetFocus
            Exit Sub

        End If

        'estado_cuenta
        '----------------- preparando temporales estado cuenta
        'buf = "INSERT INTO kardex from dcomanda where salon='" & SALON & "' and mesa='" & MESA & "'"
        'cn.Execute (buf)
        suma_comandas
        found = sumar_destadocuenta("" & salon, "" & mesa)

        If found = 0 Then
            MsgBox "No se pudo imprimir ", 48, "Aviso"
            Exit Sub

        End If

        '    fin temporales estado cuenta
        formato_precuenta "" & salon, "" & mesa
    
        If tbrmytablex.RecordCount > 0 Then
            tbrmytablex.MoveFirst
            table2.refresh
            table2.SetFocus

        End If

        clave = ""
        Frame1.Visible = False
        fl9923_Click
        Exit Sub

    End If

    If Frame1.Caption = "DESCUENTO" Then  'DESCUENTO
        found = busca_clave(2)

        If found = 0 Then
            clave = ""
            clave.SetFocus
            Exit Sub

        End If

        Tredscto.total = txtotal
        Tredscto.Show 1
        grabar_descto
        'menu_descuento
        clave = ""
        Frame1.Visible = False

    End If

    If Frame1.Caption = "DESCUENTO PRODUCTO" Then  'DESCUENTO
        found = busca_clave(2)

        If found = 0 Then
            clave = ""
            clave.SetFocus
            Exit Sub

        End If

        Frame1.Visible = False
        'menu_descuento
        desporcentaje = ""
        clave = ""
        desdescripcio = "" & tbrmytablex.Fields("descripcio")
        desprecio = "" & tbrmytablex.Fields("precio")
        Frame5.Visible = True
        desporcentaje.SetFocus

    End If

    Exit Sub
cmd891_err:
    MsgBox "Aviso en Clave " + error$, 48, "Aviso"
    Exit Sub

End Sub

'' 22/12/2017 Anulacion completa de comanda
Function graba_logcomandaTodo()

    Dim mytablex As New ADODB.Recordset

    Dim I        As Integer

    If tbrmytablex.RecordCount = 0 Then Exit Function
   
    mytablex.Open "SELECT * FROM logcomanda where tipo='A'", cn, adOpenDynamic, adLockOptimistic

    mytablex.AddNew

    For I = 0 To tbrmytablex.Fields.count - 2
        mytablex(I) = tbrmytablex(I)
    Next I

    mytablex.Fields("fechaborra") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("horaborra") = Format(Now, "hh:mm:ss")
    mytablex.Fields("administrador") = Trim("" & logclave)
    mytablex.Fields("observa1") = Trim$(MOTIVO.Text)
    mytablex.Update

End Function

'' 22/12/2017 Anulacion completa de comanda

Private Sub cmdGuardar_Click()
    'MsgBox "x"

    If Val(cantdev) > Val(cantidad) Then
        MsgBox "Intente de Nuevo ", 48, "Aviso"
        cantdev.SetFocus
        Exit Sub

    End If

    tbrmytablex.Fields("cantdev") = Val(cantdev)
    tbrmytablex.Update
    table2.refresh
    Frame3.Visible = False
    Exit Sub

End Sub

Private Sub cmdMoviendoProducto_Click(Index As Integer)

    If tbrmytablex.RecordCount = 0 Then Exit Sub
    Frame1.Visible = True
    Frame1.Caption = "MOVIENDO PRODUCTO"
    clave = ""
    clave.SetFocus

End Sub

Private Sub cmdReimprimirComanda_Click(Index As Integer)
    Frame1.Visible = True
    Frame1.Caption = "REIMPRIME"
    clave = ""
    clave.SetFocus

End Sub

Private Sub Command1_Click(Index As Integer)

    If Index = 12 Then
        clave_KeyPress 13
        Exit Sub

    End If

    If Index = 10 Then
        clave = ""
        Exit Sub

    End If

    clave = clave & Command1(Index).Caption

End Sub

Function eselmeseroo(buf As String, buf1 As String, buf2 As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from dcomanda where salon='" & buf & "' and mesa='" & buf1 & "' and vendedor='" & buf2 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        eselmeseroo = 1

    End If

    mytablex.Close

End Function

Private Sub Command11_Click()

    If Len(mesaf) = 0 Then Exit Sub

    If Frame4.Visible = True Then

        If Not IsNumeric(mcantidadm) Then
            MsgBox "Debe ingresar una cantidad numerico ", 48, "Aviso"
            Exit Sub

        End If
   
        If Val(mcantidadm) <= 0 Then
            MsgBox "Debe ingresar una cantidad numerico ", 48, "Aviso"
   
            ''28/06/2017 'CORRECCION  Y VALIDACION DE CANTIDAD AL MOVER PRODUCTOS MESA
            mcantidadm = ""
            ''28/06/2017 'CORRECCION  Y VALIDACION DE CANTIDAD AL MOVER PRODUCTOS MESA
            Exit Sub

        End If
   
        If Val(mcantidadm) > Val(mcantidad) Then
            MsgBox "Cantidad Mayor ", 48, "Aviso"
            mcantidadm = ""
            Exit Sub

        End If
   
    End If

    If Frame2.Caption = "JUNTAR MESAS" Then
        junta_mesas

    End If

    If Frame2.Caption = "MOVIENDO MESAS" Then
        moviendo_salon

    End If

    If Frame2.Caption = "MOVIENDO MESAS PRODUCTOS" Then
        moviendo_salonproducto

    End If
   
    Frame2.Visible = False
    fl9923_Click

End Sub

Private Sub Command13_Click()
    Frame5.Visible = False

End Sub

Private Sub Command14_Click()

    If Not IsNumeric(desporcentaje) Then
        MsgBox "Intente de Nuevo ", 48, "Aviso"
        desporcentaje.SetFocus
        Exit Sub

    End If

    tbrmytablex.Fields("deslipo") = Val(desporcentaje)
    resuma_comanda tbrmytablex, 0
    'suma_linea
    tbrmytablex.Update
    table2.refresh
    suma_comandas
    Frame5.Visible = False
    Exit Sub

End Sub

Private Sub Command15_Click()

    If tbrmytablex.RecordCount = 0 Then Exit Sub
    Frame1.Visible = True
    Frame1.Caption = "DESCUENTO PRODUCTO"
    clave = ""
    clave.SetFocus

End Sub

Private Sub Command2_Click()
    Frame1.Visible = True
    Frame1.Caption = "AVANCE CUENTA"
    clave = ""
    clave.SetFocus

End Sub

Private Sub Command3_Click()
    Frame1.Visible = True
    Frame1.Caption = "ANULA PRODUCTO"
    clave = ""
    clave.SetFocus

End Sub

Private Sub Command4_Click()
    Frame1.Visible = False

End Sub

Private Sub Command5_Click()

    If tbrmytablex.RecordCount = 0 Then Exit Sub
    Frame1.Visible = True
    Frame1.Caption = "DESCUENTO"
    clave = ""
    clave.SetFocus

End Sub

Private Sub Command6_Click()

    If tbrmytablex.RecordCount = 0 Then Exit Sub
    Frame1.Visible = True
    Frame1.Caption = "JUNTANDO MESA"
    clave = ""
    clave.SetFocus

End Sub

Private Sub Command7_Click()
    Frame3.Visible = False

End Sub

Private Sub Command8_Click()
    Frame2.Visible = False

End Sub

Private Sub Command9_Click()

    If tbrmytablex.RecordCount = 0 Then Exit Sub
    Frame1.Visible = True
    Frame1.Caption = "MOVIENDO MESA"
    clave = ""
    clave.SetFocus

End Sub

Private Sub fl9923_Click()
    TBRCOMA.Hide
    Unload TBRCOMA

End Sub

Private Sub Form_Activate()

    '24/08/2018  Delivery por mesa
    'Me.Width = 15420: Me.Height = 8955
    Me.Width = 15420: Me.Height = 10965
    '24/08/2018  Delivery por mesa

    Frame1.Top = 1680: Frame1.Left = 3360
    Frame2.Top = 10: Frame2.Left = 10
    Frame3.Top = 2055: Frame3.Left = 1860
    Frame5.Top = 1680: Frame5.Left = 3360

    If "" & mytable11.Fields("descuento") = "N" Then  'activa autoservicio
        Command5.Enabled = False

    End If

    If "" & mytable11.Fields("precuenta") = "N" Then  'activa autoservicio
        Command2.Enabled = False

    End If

    If tbrmytablex.State = 1 Then
        tbrmytablex.Close

    End If

    Set tbrmytablex = Nothing
    mira_fechas
    actualiza_comanda
    consulta_comanda
    suma_comandas
    Label2.Caption = "" & glomesa
    Command9.Caption = "MOVIENDO " & glomesa
    Command6.Caption = "JUNTANDO " & glomesa
    Label5.Caption = "" & glomesa
    Label7.Caption = "" & glomesa
    table2.columns(1).Caption = "" & glomesa
    '-------------------inicializo frames ---------------
    Frame1.Top = 1575: Frame1.Left = 3465
    Frame2.Top = 10: Frame2.Left = 75
    Frame3.Top = 0: Frame3.Left = 2280

    Frame5.Top = 1680: Frame5.Left = 3360

    '24/08/2018  Delivery por mesa
    Call extrae_datos_mesa(salon, mesa)
    '24/08/2018  Delivery por mesa

End Sub

Sub mira_fechas()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * from mesa where salon='" & salon & "' and mesa='" & mesa & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        cfecha = Format("" & mytablex.Fields("fecha"), "dd/mm/yyyy")
        chora = Format("" & mytablex.Fields("hora"), "hh:mm:ss")

    End If

    mytablex.Close

End Sub

Sub consulta_comanda()

    If tbrmytablex.State = 1 Then tbrmytablex.Close
    Set tbrmytablex = Nothing
    tbrmytablex.Open "SELECT * from dcomanda where len(salon)>0 and len(mesa)>0 and len(numero)>0  and salon='" & salon & "' and mesa='" & mesa & "'", cn, adOpenDynamic, adLockOptimistic
    Set table2.DataSource = tbrmytablex
    table2.refresh

End Sub

Private Sub groupmesa_Click(Index As Integer)
    mesaf = Trim("" & groupmesa(Index).Caption)

End Sub

Private Sub Image2_Click()
    menu_mesa "SIG", saloni

End Sub

Private Sub Image3_Click()
    menu_mesa "ANT", saloni

End Sub

Private Sub Image7_Click()

    If tbrmytablex.EOF = False Then
        tbrmytablex.MoveNext 'movemos al siguiente registro
        Exit Sub

    End If

    If Not tbrmytablex.BOF Then
        tbrmytablex.MoveLast
        Exit Sub

    End If

End Sub

Private Sub Image8_Click()

    If tbrmytablex.BOF = False Then
        tbrmytablex.MovePrevious 'movemos al registro anterior
        Exit Sub

    End If

    'dbvarios.MovePrevious
    If Not tbrmytablex.EOF Then
        tbrmytablex.MoveFirst
        Exit Sub

    End If

End Sub

Function busca_clave(sw As Integer)

    Dim mytablex As New ADODB.Recordset

    logclave = ""

    If sw = 0 Then
        mytablex.Open "SELECT * from vendedor where clave='" & clave & "' and anula='S'", cn, adOpenDynamic, adLockOptimistic

    End If

    If sw = 1 Then
        mytablex.Open "SELECT * from vendedor where clave='" & clave & "' and precuenta='S'", cn, adOpenDynamic, adLockOptimistic

    End If

    If sw = 2 Then
        mytablex.Open "SELECT * from vendedor where clave='" & clave & "' and descuento='S'", cn, adOpenDynamic, adLockOptimistic

    End If

    If sw = 3 Then
        mytablex.Open "SELECT * from vendedor where clave='" & clave & "' and borra_comanda='S'", cn, adOpenDynamic, adLockOptimistic

    End If

    If sw = 20 Then
        mytablex.Open "SELECT * from vendedor where clave='" & clave & "' and juntamesa='S'", cn, adOpenDynamic, adLockOptimistic

    End If

    If sw = 21 Then
        mytablex.Open "SELECT * from vendedor where clave='" & clave & "' and muevemesa='S'", cn, adOpenDynamic, adLockOptimistic

    End If

    If sw = 22 Then
        mytablex.Open "SELECT * from vendedor where clave='" & clave & "' and mueveproducto='S'", cn, adOpenDynamic, adLockOptimistic

    End If

    If mytablex.RecordCount > 0 Then
        busca_clave = 1
        logclave = "" & mytablex.Fields("codigo")

    End If

    mytablex.Close
    Set mytablex = Nothing

End Function

Function imprimir_orden_anula(sw As Integer)

    Dim found   As Integer

    Dim puertox As String

    Dim buf     As String

    Dim Puerto  As String

    Dim puertos As String

    Dim puertod As String

    Dim oldprinter

    Dim cola As String

    On Error GoTo cmd891212_err

    'puertos="LPT"
    puertos = "oc"
    cerrar_archivo
    found = busca_familia_orden("" & tbrmytablex.Fields("producto"), Puerto, puertod, cola)

    If found = 0 Then
        Puerto = puertos

    End If

    If Len(Puerto) = 0 Then
        Puerto = "LPT"

    End If
   
    ''12/07/2017 kenyo anular multicomandas
    'MsgBox "Presione enter para continuar " & puertox
   
    If mytable11.Fields("multicomanda") = "S" Then
        If nroimpresion = 3 Then
            MsgBox "Presione enter para continuar " & puertox

        End If

    Else
        MsgBox "Presione enter para continuar " & puertox

    End If

    ''12/07/2017 kenyo anular multicomandas
   
    '--------------------------------------
    'guardando en un archivo
   
    FileName = tptovta.caja & Puerto
    found = borra_nombre("" & FileName)
    ncanal = FreeFile
    'MsgBox filename
    Open FileName For Append As #ncanal

    If sw = 0 Then
   
        IMPRIME_LINEA_ANULA
   
    End If

    If sw = 1 Then
        IMPRIME_LINEA_REIMPRIME

    End If

    Close #ncanal
    cerrar_archivo

    'ahora la impresion
    'found = star_sp342orden(puerto, 0)
    'MsgBox puertod
    If cola = "S" Then
        oldprinter = Printer.DeviceName
        selecciona_impresoras (Trim(puertod))
        'MsgBox FileName
        'MsgBox puerto
        'MsgBox puertod
        'End
        found = Imprime_archivojj(FileName, 0, "" & mytable11.Fields("tamanorden"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))
        'found = Imprime_archivojj(puertod, 0, "" & mytable11.Fields("tamanorden"), "" & mytable11.Fields("nombrefont"))
        'MsgBox "xx"
        'found = Imprime_archivojj(xbuf0, 0, "" & mytable11.Fields("tamanorden"), "" & mytable11.Fields("nombrefont"))
        selecciona_impresoras (oldprinter)
    Else

        'found = star_sp342(xxpuerto, 0)
        'found = star_sp342(xbuf1, ticketera_cajon)
    End If

    cerrar_archivo
    found = borra_nombre("" & FileName)
    Exit Function
cmd891212_err:
    MsgBox "Error Anular Orden 1 producto " & error$, 48, "Aviso"
    Exit Function

End Function

'' 22/12/2017 Anulacion completa de comanda

Function imprimir_orden_anulaTodo(sw As Integer)

    Dim found   As Integer

    Dim puertox As String

    Dim buf     As String

    Dim Puerto  As String

    Dim puertos As String

    Dim puertod As String

    Dim oldprinter

    Dim cola As String

    On Error GoTo cmd891212_err

    'puertos="LPT"
    puertos = "oc"
    cerrar_archivo
   
    Do
        found = busca_familia_orden("" & tbrmytablex.Fields("producto"), Puerto, puertod, cola)

        If found = 0 Then
            Puerto = puertos

        End If

        If Len(Puerto) = 0 Then
            Puerto = "LPT"

        End If

        ''12/07/2017 kenyo anular multicomandas
        'MsgBox "Presione enter para continuar " & puertox
   
        If mytable11.Fields("multicomanda") = "S" Then
            If nroimpresion = 3 Then
                MsgBox "Presione enter para continuar " & puertox

            End If

        Else
            MsgBox "Presione enter para continuar " & puertox

        End If

        ''12/07/2017 kenyo anular multicomandas
  
        '--------------------------------------
        'guardando en un archivo
   
        FileName = tptovta.caja & Puerto
        found = borra_nombre("" & FileName)
        ncanal = FreeFile
        'MsgBox filename
        Open FileName For Append As #ncanal

        If sw = 0 Then
            IMPRIME_LINEA_ANULA

        End If

        If sw = 1 Then
            IMPRIME_LINEA_REIMPRIME

        End If

        Close #ncanal
        cerrar_archivo

        'ahora la impresion
        'found = star_sp342orden(puerto, 0)
        'MsgBox puertod
        If cola = "S" Then
            oldprinter = Printer.DeviceName
            selecciona_impresoras (Trim(puertod))
            'MsgBox FileName
            'MsgBox puerto
            'MsgBox puertod
            'End
            found = Imprime_archivojj(FileName, 0, "" & mytable11.Fields("tamanorden"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))
            'found = Imprime_archivojj(puertod, 0, "" & mytable11.Fields("tamanorden"), "" & mytable11.Fields("nombrefont"))
            'MsgBox "xx"
            'found = Imprime_archivojj(xbuf0, 0, "" & mytable11.Fields("tamanorden"), "" & mytable11.Fields("nombrefont"))
            selecciona_impresoras (oldprinter)
        Else

            'found = star_sp342(xxpuerto, 0)
            'found = star_sp342(xbuf1, ticketera_cajon)
        End If
         
        tbrmytablex.MoveNext
    Loop
         
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Exit Function
cmd891212_err:
    MsgBox "Error Anular Orden 1 producto " & error$, 48, "Aviso"
    Exit Function

End Function

'' 22/12/2017 Anulacion completa de comanda

Sub IMPRIME_LINEA_ANULA()

    Dim buf   As String

    Dim found As Integer

    Dim I     As Integer

    On Error GoTo cmd23_err

    '----- formato nuevo
    
    found = formateaa("      **** ANULACIÓN     ****", 28, 2, 0)
    buf = String(0, "-")
    found = formateaa(buf, 0, 2, 0)
   
    buf = "Salon:" & tbrmytablex.Fields("salon") & " Mesa:" & tbrmytablex.Fields("mesa") & " Mesero:" & tbrmytablex.Fields("vendedor")
    found = formateaa(buf, 28, 2, 0)
    
    buf = "Comanda:" & tbrmytablex.Fields("comanda")
    found = formateaa(buf, 28, 2, 0)
    
    '' 10/07/2018 Edicion Comanda
    If formatocomanda = "G" Then
        buf = String(28, "-")
        found = formateaa(buf, 28, 2, 0)
    Else
        buf = String(60, "-")
        found = formateaa(buf, 65, 2, 0)
    
    End If

    '' 10/07/2018 Edicion Comanda

    If formatocomanda = "D" Then
        If tipocomanda = "DL" Then
            
            If ("" & mytable11.Fields("tamanorden")) = 8 Then
            
                buf = "*" & Mid$("" & tbrmytablex.Fields("DESCRIPCIO"), 1, 35)
                found = formateaa(buf, 36, 2, 0)

                If Len("" & tbrmytablex.Fields("DESCRIPCIO")) > 35 Then
                    buf = Mid$("" & tbrmytablex.Fields("DESCRIPCIO"), 36, 38)
                    found = formateaa(buf, 38, 2, 0)

                End If

                If Len("" & tbrmytablex.Fields("DESCRIPCIO")) > 73 Then
                    buf = Mid$("" & tbrmytablex.Fields("DESCRIPCIO"), 74, 42)
                    found = formateaa(buf, 38, 2, 0)

                End If

                If Len("" & tbrmytablex.Fields("DESCRIPCIO")) > 121 Then
                    buf = Mid$("" & tbrmytablex.Fields("DESCRIPCIO"), 108, 42)
                    found = formateaa(buf, 38, 2, 0)

                End If
        
            Else
                buf = "*" & Mid$("" & tbrmytablex.Fields("DESCRIPCIO"), 1, 23)
                found = formateaa(buf, 23, 2, 0)

                If Len("" & tbrmytablex.Fields("DESCRIPCIO")) > 22 Then
                    buf = Mid$("" & tbrmytablex.Fields("DESCRIPCIO"), 22, 44)
                    found = formateaa(buf, 28, 2, 0)

                End If

                If Len("" & tbrmytablex.Fields("DESCRIPCIO")) > 46 Then
                    buf = Mid$("" & tbrmytablex.Fields("DESCRIPCIO"), 50, 66)
                    found = formateaa(buf, 28, 2, 0)

                End If

                If Len("" & tbrmytablex.Fields("DESCRIPCIO")) > 90 Then
                    buf = Mid$("" & tbrmytablex.Fields("DESCRIPCIO"), 78, 88)
                    found = formateaa(buf, 28, 2, 0)

                End If

                If Len("" & tbrmytablex.Fields("DESCRIPCIO")) > 114 Then
                    buf = Mid$("" & tbrmytablex.Fields("DESCRIPCIO"), 106, 110)
                    found = formateaa(buf, 28, 2, 0)

                End If

            End If

        ElseIf tipocomanda = "CO" Then
            buf = "*" & Mid$("" & tbrmytablex.Fields("PRODUCTO"), 1, 30)
            found = formateaa(buf, 31, 2, 0)
        ElseIf tipocomanda = "DC" Then

            Dim mytablepr As New ADODB.Recordset

            mytablepr.Open "SELECT DESCORTO FROM PRODUCTO where producto='" & "" & tbrmytablex.Fields("producto") & "'", cn, adOpenDynamic, adLockOptimistic

            If mytablepr.RecordCount > 0 Then
                buf = "*" & Mid$("" & mytablepr.Fields("DESCORTO"), 1, 30)
                found = formateaa(buf, 31, 2, 0)
            Else
                buf = "*" & Mid$("" & mytablepr.Fields("descorto"), 1, 30)
                found = formateaa(buf, 31, 2, 0)

            End If
            
            mytablepr.Close

        End If

    End If
    
    '' 10/07/2018 Edicion Comanda

    '    If Len("" & tbrmytablex.Fields("descripcio")) > 20 Then
    '       buf = "" & Mid$("" & tbrmytablex.Fields("descripcio"), 1, 20)
    '       found = formateaa(buf, 20, 0, 0)
    '       found = formateaa(" ", 1, 2, 0)
    '       buf = "" & Mid$("" & tbrmytablex.Fields("descripcio"), 21, 20)
    '       found = formateaa(buf, 20, 0, 0)
    '       found = formateaa(" ", 1, 0, 0)
    '    Else
    '    buf = "" & Mid$("" & tbrmytablex.Fields("descripcio"), 1, 20)
    '    found = formateaa(buf, 20, 0, 0)
    '    found = formateaa(" ", 1, 0, 0)
    '    End If

    buf = "-" & tbrmytablex.Fields("cantidad")
    found = formateaa(buf, 7, 2, 1)
    found = formateaa("", 1, 2, 0)

    If Len("" & tbrmytablex.Fields("observa1")) > 0 Then
        buf = "*" & tbrmytablex.Fields("observa1")
        found = formateaa(buf, 28, 2, 0)

    End If

    If Len("" & tbrmytablex.Fields("observa2")) > 0 Then
        buf = "*" & tbrmytablex.Fields("observa2")
        found = formateaa(buf, 28, 2, 0)

    End If

    If Len("" & tbrmytablex.Fields("observa3")) > 0 Then
        buf = "*" & tbrmytablex.Fields("observa3")
        found = formateaa(buf, 28, 2, 0)

    End If

    For I = 1 To 5
        found = formateaa("", 1, 2, 0)
    Next I

    Exit Sub
cmd23_err:
    MsgBox "Aviso en imprime linea anula " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub IMPRIME_LINEA_REIMPRIME()

    Dim buf   As String

    Dim found As Integer

    Dim I     As Integer

    On Error GoTo cmd923_err

    '----- formato nuevo
    found = formateaa("****REIMPRESION****", 28, 2, 0)
    buf = "Salon:" & tbrmytablex.Fields("salon") & " Mesa:" & tbrmytablex.Fields("mesa") & " Mesero:" & tbrmytablex.Fields("vendedor")
    found = formateaa(buf, 28, 2, 0)
    buf = "Comanda:" & tbrmytablex.Fields("comanda")
    found = formateaa(buf, 28, 2, 0)
    buf = String(28, "-")
    found = formateaa(buf, 28, 2, 0)

    If Len("" & tbrmytablex.Fields("descripcio")) > 20 Then
        buf = "" & Mid$("" & tbrmytablex.Fields("descripcio"), 1, 20)
        found = formateaa(buf, 20, 0, 0)
        found = formateaa(" ", 1, 2, 0)
        buf = "" & Mid$("" & tbrmytablex.Fields("descripcio"), 21, 20)
        found = formateaa(buf, 20, 0, 0)
        found = formateaa(" ", 1, 0, 0)
    Else
        buf = "" & Mid$("" & tbrmytablex.Fields("descripcio"), 1, 20)
        found = formateaa(buf, 20, 0, 0)
        found = formateaa(" ", 1, 0, 0)

    End If
    
    buf = "" & tbrmytablex.Fields("cantidad")
    found = formateaa(buf, 7, 2, 1)
    found = formateaa("", 1, 2, 0)

    If Len("" & tbrmytablex.Fields("observa1")) > 0 Then
        buf = "*" & tbrmytablex.Fields("observa1")
        found = formateaa(buf, 28, 2, 0)

    End If

    If Len("" & tbrmytablex.Fields("observa2")) > 0 Then
        buf = "*" & tbrmytablex.Fields("observa2")
        found = formateaa(buf, 28, 2, 0)

    End If

    If Len("" & tbrmytablex.Fields("observa3")) > 0 Then
        buf = "*" & tbrmytablex.Fields("observa3")
        found = formateaa(buf, 28, 2, 0)

    End If

    For I = 1 To 5
        found = formateaa("", 1, 2, 0)
    Next I

    Exit Sub
cmd923_err:
    MsgBox "Aviso en imprime linea REIMPRESION " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function busca_familia_orden(buf1 As String, _
                             Puerto As String, _
                             puertod As String, _
                             cola As String)

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd90fam_err

    mytablex.Open "SELECT * FROM producto where producto='" & buf1 & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        ''12/07/2017 kenyo anular multicomandas
        '
        '            Puerto = "" & mytablex.Fields("grupoimpresion")
        '            puertod = "" & mytablex.Fields("puertoimpresion")
        '            cola = "" & mytablex.Fields("cola")
        '            busca_familia_orden = 1

        If nroimpresion = 0 Then
            Puerto = "" & mytablex.Fields("grupoimpresion")
            puertod = "" & mytablex.Fields("puertoimpresion")
            cola = "" & mytablex.Fields("cola")
            busca_familia_orden = 1

            If Len(Trim(puertod)) = 0 Then
                busca_familia_orden = 0

            End If

        End If
  
        If nroimpresion = 1 Then
            Puerto = "" & mytablex.Fields("grupoimpresion")
            puertod = "" & mytablex.Fields("puertoimpresion1")
            cola = "" & mytablex.Fields("cola")
            busca_familia_orden = 1
          
            If Len(Trim(mytablex.Fields("puertoimpresion1"))) = "0" Then
           
                puertod = ""
                Puerto = ""
                cola = ""
                busca_familia_orden = 0
           
            End If
           
            If Len(Trim(puertod)) = 0 Then
                busca_familia_orden = 0

            End If
            
        End If

        If nroimpresion = 2 Then
            Puerto = "" & mytablex.Fields("grupoimpresion")
            puertod = "" & mytablex.Fields("puertoimpresion2")
            cola = "" & mytablex.Fields("cola")
            busca_familia_orden = 1
            
            If Len(Trim(mytablex.Fields("puertoimpresion2"))) = "0" Then
           
                puertod = ""
                Puerto = ""
                cola = ""
                busca_familia_orden = 0
           
            End If
           
            If Len(Trim(puertod)) = 0 Then
                busca_familia_orden = 0

            End If

        End If
  
        If nroimpresion = 3 Then
            Puerto = "" & mytablex.Fields("grupoimpresion")
            puertod = "" & mytablex.Fields("puertoimpresion3")
            cola = "" & mytablex.Fields("cola")
            busca_familia_orden = 1
            
            If Len(Trim(mytablex.Fields("puertoimpresion3"))) = "0" Then
           
                puertod = ""
                Puerto = ""
                cola = ""
                busca_familia_orden = 0
           
            End If
           
            If Len(Trim(puertod)) = 0 Then
                busca_familia_orden = 0

            End If

        End If

        ''12/07/2017 kenyo anular multicomandas
       
    End If

    mytablex.Close
    Exit Function
cmd90fam_err:
    MsgBox "Aviso en Busca Familia orden " + error$, 48, "Aviso"
    Exit Function

End Function

'Function busca_familia_orden(buf1 As String, puerto As String, puertod As String, cola As String)
'Dim mytablex As New ADODB.Recordset
'mytablex.Open "SELECT * FROM producto where producto='" & buf1 & "'", cn, adOpenDynamic, adLockOptimistic
'If mytablex.RecordCount > 0 Then
'      puerto = "" & mytablex.Fields("grupoimpresion")
'      puertod = "" & mytablex.Fields("puertoimpresion")
'      cola = "" & mytablex.Fields("cola")
'      busca_familia_orden = 1
'End If
'mytablex.Close
'   Exit Function

'End Function

Sub estado_cuenta()

    Dim Puerto As String

    Dim found  As Integer

    Dim I      As Integer

    Dim oldprinter

    'panel3d13.Visible = True
    'label60 = "ESTADO CUENTA"
    '-----OJO ESTO SE ADICIONO..
   
    cerrar_archivo
    found = estado_mesas("" & salon, "" & mesa, "2")
    FileName = gusuario
    found = borra_nombre("" & FileName)
    ncanal = FreeFile
    Open FileName For Append As #ncanal
    Puerto = "" & mytable11.Fields("ecpuerto")  'impresora precuenta
    '----ojo es temporal
    'If MsgBox("Desea Imprimir ", 1, "Aviso") <> 1 Then Exit Sub
    cabecera_pedido
    cuerpo_cuenta

    If "" & mytable11.Fields("eccola") <> "S" Then

        For I = 1 To 10
            found = formateaa("", 1, 2, 0)
        Next I

    End If

    Close #ncanal
    cerrar_archivo
   
    If "" & mytable11.Fields("eccola") <> "S" Then
        '------------------------------------
        found = star_sp342("" & mytable11.Fields("ecpuerto"), 0)
        found = corte_papel("" & mytable11.Fields("ecpuerto"), 1)

        '------------------------------------
    End If

    If "" & mytable11.Fields("eccola") = "S" Then
        oldprinter = Printer.DeviceName
        selecciona_impresoras ("" & mytable11.Fields("ecpuerto"))
        found = Imprime_archivojj(FileName, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))
        selecciona_impresoras (oldprinter)

    End If
   
    cerrar_archivo

End Sub

Sub cabecera_pedido()

    Dim found    As Integer

    Dim buf      As String

    Dim buf2     As String

    Dim btipo    As String

    Dim xmozo    As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd4111_err

    buf = String(45, "-")
    found = formateaa(buf, 45, 2, 0)
    buf = "       AVANCE DE CUENTA"
    found = formateaa(buf, 32, 2, 0)
    buf = "       COMPROBANTE NO AUTORIZADO"
    found = formateaa(buf, 32, 2, 0)
    'buf = "    Cajero:" & USUARIO & " Caja:" & caja & " Turno:" & TURNO
    'found = formateaa(buf, 36, 2, 0)
    buf = "  Fecha:" & Format(Now, "dd/mm/yyyy") & "  Hora:" & Format(Now, "hh:mm:ss")
    found = formateaa(buf, 32, 2, 0)
    'If tservicio = "*" Then
    '   found = formateaa(" *** RAPIDO    ***", 25, 2, 0)
    'End If
    'If tservicio = "C" Then
    xmozo = ""
    buf = "   Salon :" & "" & salon & " Mesa:" & "" & mesa
      
    mytablex.Open "SELECT * FROM dcomanda where salon='" & salon & "' and mesa='" & mesa & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        xmozo = mytablex.Fields("vendedor")

    End If

    mytablex.Close
    mytablex.Open "SELECT * FROM vendedor where codigo='" & xmozo & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        xmozo = Mid$(mytablex.Fields("nombre"), 1, 7)

    End If

    mytablex.Close

    buf = buf & " Mozo:" & xmozo & " Pers:" & "1"
    found = formateaa(buf, 32, 2, 0)
    'End If
    'If tservicio = "D" Then
    '   found = formateaa(" *** DOMICILIO ***", 36, 2, 0)
    '   found = formateaa(buf, 36, 2, 0)
    '   imprime_cliente_delivery "" & codigocli
    '
    'End If
    found = formateaa("Cant  Producto        P.U.     Tot", 32, 2, 0)
    buf = String(45, "-")
    found = formateaa(buf, 45, 2, 0)
    Exit Sub
cmd4111_err:
    MsgBox "Mensaje,Error en cabecera Pedido " & error$, 48, "Aviso"
    Exit Sub

End Sub

Sub cuerpo_cuenta()

    Dim buf     As String

    Dim found   As Integer

    Dim zmoneda As String

    Dim I       As Integer

    Dim buf2    As String

    Dim buf3    As String

    Dim sw      As Integer

    On Error GoTo cmd3999_err

    If Len(Trim(salon)) = 0 Then Exit Sub
    If Len(Trim(mesa)) = 0 Then Exit Sub

    buf2 = "" & salon
    buf3 = "" & mesa

    If "" & mytable11.Fields("moneda") = "S" Then
        zmoneda = dicmoneda

    End If

    If "" & mytable11.Fields("moneda") = "D" Then
        zmoneda = "US$"

    End If

    suma1 = 0
    suma2 = 0
    
    tbrmytablex.MoveFirst
    Do

        If tbrmytablex.EOF Then Exit Do
        imprime_estado_cuenta
        tbrmytablex.MoveNext
    Loop
    
    'For i = 1 To 5
    'found = formateaa("", 1, 2, 0)
    'Next i
    buf = String(45, "-")
    found = formateaa(buf, 45, 2, 0)

    buf = "    NroUnidades "
    found = formateaa(buf, 20, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma1, "0.00")
    found = formateaa(buf, 7, 2, 1)
    buf = "****TOTAL       "
    found = formateaa(buf, 17, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa(zmoneda, 3, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma2, "0.00")
    found = formateaa(buf, 9, 2, 1)

    found = formateaa("", 1, 2, 0)
    found = formateaa("RUC No:____________________________", 32, 2, 0)
    found = formateaa("", 1, 2, 0)
    found = formateaa("RAZON SOCIAL:______________________", 32, 2, 0)
    found = formateaa("", 1, 2, 0)
    found = formateaa("___________________________________", 32, 2, 0)

    For I = 1 To 1
        found = formateaa("", 1, 2, 0)
    Next I

    Exit Sub
cmd3999_err:
    MsgBox "Mensaje, Error en cuerpo Cuenta " & error$
    Exit Sub

End Sub

Sub imprime_estado_cuenta()

    Dim buf   As String

    Dim found As Integer

    On Error GoTo cmd32156_err

    buf = "" & tbrmytablex.Fields("cantidad")
    found = formateaa(buf, 3, 0, 0)
    found = formateaa(" ", 1, 0, 0)
    buf = Mid$("" & tbrmytablex.Fields("descripcio"), 1, 15)
    found = formateaa(buf, 15, 0, 0)
    found = formateaa(" ", 1, 0, 0)
    
    buf = "" & tbrmytablex.Fields("precio")
    found = formateaa(buf, 5, 0, 1)
    found = formateaa(" ", 1, 0, 0)

    buf = "" & tbrmytablex.Fields("total")
    buf = Format(Val(buf), "0.00")
    found = formateaa(buf, 6, 2, 1)

    suma1 = suma1 + Val("" & tbrmytablex.Fields("cantidad"))
    suma2 = suma2 + Val("" & tbrmytablex.Fields("total"))
    Exit Sub
cmd32156_err:
    MsgBox "Aviso en imprime estado cuenta ", 48, "Aviso"
    Exit Sub

End Sub

Sub grabar_descto()

    On Error GoTo cmd6543_err

    Dim found As Integer

    Dim a     As Double

    If tbrmytablex.RecordCount = 0 Then Exit Sub
    tbrmytablex.MoveFirst
    Do

        If tbrmytablex.EOF Then Exit Do
        If (Val("" & tbrmytablex.Fields("cantidad")) > 0 Or Val("" & tbrmytablex.Fields("cantidad")) < 0) And Val("" & tbrmytablex.Fields("precio")) > 0 Then

            'tbrmytablex.Edit
            'MsgBox tipodescuento
            If tipodescuento = "2" Then
                tbrmytablex.Fields("destopo") = 0

            End If

            If tipodescuento = "0" Then
                tbrmytablex.Fields("destopo") = Val(valordescuento)

            End If

            If tipodescuento = "1" Then
                a = (Val(valordescuento) * 100) / Val(txtotal)
                tbrmytablex.Fields("destopo") = a

            End If

            If tipodescuento = "3" Then   '----recargos
                tbrmytablex.Fields("destopo") = 0
                tbrmytablex.Fields("precio") = Val("" & tbrmytablex.Fields("precio")) + Val("" & tbrmytablex.Fields("precio")) * valordescuento / 100
                tbrmytablex.Fields("TOTAL") = Val("" & tbrmytablex.Fields("precio")) * Val("" & tbrmytablex.Fields("cantidad"))

            End If

            'suma_linea
            resuma_comanda tbrmytablex, 0
            tbrmytablex.Update

        End If

        tbrmytablex.MoveNext
    Loop
    'sql_detalle
    suma_comandas
    Exit Sub
cmd6543_err:
    MsgBox "Aviso " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub suma_comandas()

    Dim sdx  As Double

    Dim sdx1 As Double

    txtotal = ""
    sdx = 0
    sdx1 = 0

    If tbrmytablex.RecordCount = 0 Then Exit Sub
    tbrmytablex.MoveFirst
    Do

        If tbrmytablex.EOF Then Exit Do
        sdx = sdx + Val(tbrmytablex.Fields("cantidad"))
        sdx1 = sdx1 + Val(tbrmytablex.Fields("total"))
        tbrmytablex.MoveNext
    Loop
    txtotal = Format(sdx1, "0.00")

    If tbrmytablex.RecordCount > 0 Then
        tbrmytablex.MoveFirst
        table2.refresh
        table2.SetFocus

    End If

End Sub

Sub resuma_comanda(mytablex As ADODB.Recordset, xpercepcion As Double)

    Dim xtivap      As Double

    Dim tdscto      As Double

    Dim sdx2        As Double

    Dim sdx1        As Double

    Dim xtisc       As Double

    Dim X           As Double

    Dim Y           As Double

    Dim sdx         As Double

    Dim ypercepcion As Double

    Dim xneto       As Double

    On Error GoTo cmd94534_err

    ypercepcion = 0

    mytablex.Fields("percepcion") = xpercepcion

    If busca_tipoprecio() = "N" Then
        mytablex.Fields("neto") = Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("precio"))
        mytablex.Fields("descuento") = Val("" & mytablex.Fields("neto")) * Val("" & mytablex.Fields("deslipo")) / 100 + Val("" & mytablex.Fields("neto")) * Val("" & mytablex.Fields("destopo")) / 100     'calcular descuento
        mytablex.Fields("subtotal") = Val("" & mytablex.Fields("neto")) - Val("" & mytablex.Fields("descuento")) 'cobrar
        mytablex.Fields("impuesto") = Val("" & mytablex.Fields("subtotal")) * Val("" & mytablex.Fields("igv")) / 100  'calcular descuento
        mytablex.Fields("total") = Val("" & mytablex.Fields("subtotal")) + Val("" & mytablex.Fields("impuesto")) 'cobrar
        mytablex.Fields("tivap") = Val("" & mytablex.Fields("total")) * Val("" & mytablex.Fields("ivap")) / 100
        mytablex.Fields("tpercepcio") = 0

        If "" & mytablex.Fields("l1") = "S" Then
            mytablex.Fields("tpercepcio") = Val("" & mytablex.Fields("total")) * Val("" & mytablex.Fields("percepcion")) / 100    'calcular descuento
            mytablex.Fields("total") = Val("" & mytablex.Fields("total")) + Val("" & mytablex.Fields("tpercepcio")) 'cobrar

        End If

        mytablex.Fields("servicioco") = Val("" & mytablex.Fields("subtotal")) * Val("" & mytablex.Fields("serviciopo")) / 100     'calcular descuento
        mytablex.Fields("total") = Val("" & mytablex.Fields("total")) + Val("" & mytablex.Fields("servicioco")) 'cobrar

    Else
        mytablex.Fields("neto") = Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("precio"))
        mytablex.Fields("descuento") = Val("" & mytablex.Fields("neto")) * Val("" & mytablex.Fields("deslipo")) / 100 + Val("" & mytablex.Fields("neto")) * Val("" & mytablex.Fields("destopo")) / 100
        mytablex.Fields("total") = Val("" & mytablex.Fields("neto")) - Val("" & mytablex.Fields("descuento")) 'cobrar
        mytablex.Fields("subtotal") = Val("" & mytablex.Fields("total")) / (1 + Val("" & mytablex.Fields("igv")) / 100) 'calcular descuento
        mytablex.Fields("impuesto") = Val("" & mytablex.Fields("total")) - Val("" & mytablex.Fields("subtotal")) 'cobrar
        xtivap = Val("" & mytablex.Fields("total")) * Val("" & mytablex.Fields("ivap")) / 100
        mytablex.Fields("tivap") = xtivap
        mytablex.Fields("tpercepcio") = 0

        If "" & mytablex.Fields("l1") = "S" Then
            mytablex.Fields("tpercepcio") = Val("" & mytablex.Fields("total")) * Val("" & mytablex.Fields("percepcion")) / 100   'calcular descuento
            mytablex.Fields("total") = Val("" & mytablex.Fields("total")) + Val("" & mytablex.Fields("tpercepcio")) 'cobrar

        End If

        mytablex.Fields("servicioco") = Val("" & mytablex.Fields("subtotal")) * Val("" & mytablex.Fields("serviciopo")) / 100      'calcular descuento

    End If

    Exit Sub
cmd94534_err:
    MsgBox "Aviso en resuma_precios ", 48, "Aviso"
    Exit Sub

End Sub

Sub junta_mesas()

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    On Error GoTo cmd20069_err

    mytablex.Open "SELECT * FROM mesa where salon='" & salonf & "' and mesa='" & mesaf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "No existe mesa ", 48, "Aviso"
        mytablex.Close
        Exit Sub

    End If

    mytablex.Close
    'MsgBox "abc"
    'tbrmytablex.Requery
    'consulta_comanda

    If tbrmytablex.RecordCount = 0 Then Exit Sub
    tbrmytablex.MoveFirst
    Do

        If tbrmytablex.EOF Then Exit Do
        'MsgBox tbrmytablex.Fields("salon")
        tbrmytablex.Fields("salon") = Trim(salonf)
        tbrmytablex.Fields("mesa") = Trim(mesaf)
        tbrmytablex.Update
        tbrmytablex.MoveNext
    Loop
    Exit Sub
cmd20069_err:
    MsgBox "Mensaje,Error en junta salon " & error$
    Exit Sub

End Sub

Sub moviendo_salon()

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    Dim buf      As String

    On Error GoTo cmd2006_err

    mytablex.Open "SELECT * FROM mesa where salon='" & salonf & "' and mesa='" & mesaf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "No existe mesa ", 48, "Aviso"
        mytablex.Close
        Exit Sub

    End If

    mytablex.Close

    mytablex.Open "SELECT * FROM dcomanda where salon='" & salonf & "' and mesa='" & mesaf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        MsgBox "Salon y Mesa esta ocupado ", 48, "Aviso"
        mytablex.Close
        Exit Sub

    End If

    mytablex.Close
    cn.Execute ("update dcomanda set salon='" & salonf & "',mesa='" & mesaf & "' where mesa='" & mesai & "' and salon='" & saloni & "'")
    tbrmytablex.Requery
    Exit Sub
    sw = 0

    If tbrmytablex.RecordCount = 0 Then Exit Sub
    tbrmytablex.MoveFirst

    Do

        If tbrmytablex.EOF Then Exit Do
        tbrmytablex.Fields("salon") = Trim(salonf)
        tbrmytablex.Fields("mesa") = Trim(mesaf)
        tbrmytablex.Update
        sw = 1
        tbrmytablex.MoveNext
    Loop
       
    Exit Sub
cmd2006_err:
    MsgBox "Mensaje,Error en reemplaza salon " & error$
    Exit Sub

End Sub

Sub moviendo_salonproducto()

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    Dim buf      As String

    Dim I        As Integer

    Dim sdx      As Double

    On Error GoTo cmd20066_err

    If tbrmytablex.RecordCount = 0 Then Exit Sub
    mytablex.Open "SELECT * FROM dcomanda where salon='" & salonf & "' and mesa='" & mesaf & "'", cn, adOpenDynamic, adLockOptimistic
    mytablex.AddNew
      
    ''28/06/2017 'CORRECCION  Y VALIDACION DE CANTIDAD AL MOVER PRODUCTOS MESA
    'For i = 0 To tbrmytablex.Fields.count - 1
    For I = 0 To tbrmytablex.Fields.count - 2
        ''28/06/2017 'CORRECCION  Y VALIDACION DE CANTIDAD AL MOVER PRODUCTOS MESA
      
        mytablex(I) = tbrmytablex(I)
    Next I

    mytablex.Fields("salon") = Trim(salonf)
    mytablex.Fields("mesa") = Trim(mesaf)
    mytablex.Fields("cantidad") = Val("" & mcantidadm)
    mytablex.Fields("total") = Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("precio"))
    mytablex.Update
    mytablex.Update
    mytablex.Close
    sdx = Val("" & tbrmytablex.Fields("cantidad")) - Val("" & mcantidadm)

    If sdx = 0 Then 'si es cero borrarlo
        tbrmytablex.Delete
    Else
        tbrmytablex.Fields("cantidad") = sdx
        tbrmytablex.Fields("total") = Val("" & tbrmytablex.Fields("cantidad")) * Val("" & tbrmytablex.Fields("precio"))
        tbrmytablex.Update

    End If
    
    Exit Sub
cmd20066_err:
    MsgBox "Mensaje,Error en reemplaza saalon Productos" & error$
    Exit Sub

End Sub

Private Sub Label11_Click(Index As Integer)

    If Index = 10 Then
        mcantidadm = ""
        Exit Sub

    End If

    mcantidadm = mcantidadm + Label11(Index).Caption
 
    ''28/06/2017 'CORRECCION  Y VALIDACION DE CANTIDAD AL MOVER PRODUCTOS MESA
    If Not IsNumeric(mcantidadm) Then
        MsgBox "Debe ingresar una cantidad numerico ", 48, "Aviso"
        Exit Sub

    End If
    
    If Val(mcantidadm) <= 0 Then
        MsgBox "Debe ingresar una cantidad numerico ", 48, "Aviso"
        mcantidadm = ""
        Exit Sub

    End If
    
    If Val(mcantidadm) > Val(mcantidad) Then
        MsgBox "Cantidad Mayor ", 48, "Aviso"
        mcantidadm = ""
        Exit Sub

    End If

    ''28/06/2017 'CORRECCION  Y VALIDACION DE CANTIDAD AL MOVER PRODUCTOS MESA
   
End Sub

Private Sub salonf_Click()
    'MsgBox salonf
    menu_carga_mesa salonf
    menu_mesa "INI", Trim("" & salonf)

End Sub

Private Sub table2_DblClick()

    If tbrmytablex.RecordCount = 0 Then Exit Sub
    descripcio = "" & tbrmytablex.Fields("descripcio")
    cantidad = "" & tbrmytablex.Fields("cantidad")
    cantdev = "" & tbrmytablex.Fields("cantdev")
    Frame3.Visible = True
    cantdev.SetFocus

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

    mytablex.Open "SELECT * FROM mesa where salon='" & Trim(buf) & "'", cn, adOpenDynamic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        I = I + 1
        mmesacod(I) = "" & mytablex.Fields("mesa")
        wmesacod(I) = "" & mytablex.Fields("mesa")
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
            mmesapag = mmesapag + 23

            If mmesapag > 102 Then
                mmesapag = 0

            End If

        Case "ANT"
            mmesapag = mmesapag - 23

            If mmesapag < 0 Then
                mmesapag = 0

            End If

    End Select

    j = -1

    For I = mmesapag To 23 + mmesapag
        j = j + 1
        groupmesa(j).Caption = mmesacod(I)
        verifica_mesas j, buf1, groupmesa(j).Caption
    Next I

End Sub

Sub verifica_mesas(indx As Integer, buf As String, buf1 As String)

    Dim mytablex As New ADODB.Recordset

    groupmesa(indx).BackColor = &HFFFFFF

    If Len(buf1) > 0 And Len(buf) > 0 Then
        mytablex.Open "select * from dcomanda where salon='" & buf & "' and mesa='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            groupmesa(indx).BackColor = &HFF00&

        End If

        mytablex.Close

    End If

End Sub

Function graba_logcomanda()

    Dim mytablex As New ADODB.Recordset

    Dim I        As Integer

    If tbrmytablex.RecordCount = 0 Then Exit Function
    'mytablex.Open "SELECT * FROM logcomanda where tipo='A'", cn, adOpenDynamic, adLockOptimistic 'LEONARDO: ANULAR COMANDA
    mytablex.Open "SELECT * FROM logcomanda where tipo='A'", cn, adOpenDynamic, adLockOptimistic
    mytablex.AddNew

    'For i = 0 To tbrmytablex.Fields.count - 1
    For I = 0 To tbrmytablex.Fields.count - 2 'para que no salga error por los numeros de campo
        'mytablex(i) = tbrmytablex(i)
        mytablex(I) = tbrmytablex(I)
    Next I

    mytablex.Fields("fechaborra") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("horaborra") = Format(Now, "hh:mm:ss")
    mytablex.Fields("administrador") = Trim("" & logclave)
    mytablex.Fields("observa1") = Trim$(MOTIVO.Text)
    mytablex.Update

End Function

Sub carga_salones()

    Dim mytablex As New ADODB.Recordset

    'salonf.Clear
    mytablex.Open "SELECT * from salon", cn, adOpenDynamic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        salonf.AddItem "" & mytablex.Fields("salon")
        mytablex.MoveNext
    Loop
    mytablex.Close
    salonf.ListIndex = 0

End Sub

Sub actualiza_comanda()

    On Error GoTo cmd9089_err

    Dim mytablex As New ADODB.Recordset

    Set mytablex = Nothing
    'cn.CursorLocation = adUseClient
    'mytablex.CursorLocation = adUseClient
    mytablex.Open "SELECT * from dcomanda where len(salon)>0 and len(mesa)>0 and len(numero)>0  and salon='" & salon & "' and mesa='" & mesa & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        resuma_comanda mytablex, 0
        mytablex.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    Exit Sub
cmd9089_err:
    'MsgBox "Aviso en actualiza Comanda " + error$, 48, "Aviso"
    Exit Sub

End Sub

'24/08/2018  Delivery por mesa
Function extrae_datos_mesa(salon As String, mesa As String)

    Dim mytablex As New ADODB.Recordset

    If Len(salon) > 0 And Len(mesa) > 0 Then
        mytablex.Open "select * from mesa where salon='" & salon & "' and mesa='" & mesa & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            dcodigo = Trim("" & mytablex.Fields("codigo"))
            dnombre = Trim("" & mytablex.Fields("dnombre"))
            telefono = Trim("" & mytablex.Fields("telefono"))
            ddireccion = Trim("" & mytablex.Fields("ddireccion"))
            referencia = Trim("" & mytablex.Fields("referencia"))

        End If

        mytablex.Close

    End If

End Function

'24/08/2018  Delivery por mesa

