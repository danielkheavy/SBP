VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form explobco 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorador de Movimientos Bancos"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   14805
   StartUpPosition =   1  'CenterOwner
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   6615
      Left            =   0
      OleObjectBlob   =   "explobco.frx":0000
      TabIndex        =   20
      Top             =   1200
      Width           =   14655
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFF00&
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   14745
      TabIndex        =   0
      Top             =   0
      Width           =   14805
      Begin VB.CommandButton cmdPrint 
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
         Height          =   615
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "explobco.frx":09D3
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdExit 
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
         Height          =   615
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "explobco.frx":1BE5
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdSort 
         BackColor       =   &H00C0C0C0&
         Height          =   615
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "explobco.frx":2DF7
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Consulta"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
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
         Height          =   615
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "explobco.frx":4009
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Borrar registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.ComboBox caja 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox fechaf 
         Height          =   495
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox fechai 
         Height          =   495
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   6
         Top             =   0
         Width           =   1815
      End
      Begin VB.TextBox nombre 
         Height          =   495
         Left            =   6840
         MaxLength       =   11
         TabIndex        =   5
         Text            =   "%"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox codigo 
         Height          =   495
         Left            =   6840
         MaxLength       =   11
         TabIndex        =   4
         Text            =   "%"
         Top             =   0
         Width           =   1215
      End
      Begin VB.ComboBox cajero 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
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
         Left            =   13320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "explobco.frx":521B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.ComboBox tipo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   11160
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         Height          =   495
         Left            =   8040
         TabIndex        =   19
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   495
         Left            =   3000
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         Height          =   495
         Left            =   3000
         TabIndex        =   17
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   495
         Left            =   6000
         TabIndex        =   16
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   495
         Left            =   6000
         TabIndex        =   15
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Usuario"
         Height          =   495
         Left            =   8040
         TabIndex        =   14
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   495
         Left            =   10560
         TabIndex        =   13
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Menu ki2323 
      Caption         =   "&Nuevo"
   End
   Begin VB.Menu dmo33 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu djk232 
      Caption         =   "&Borra"
   End
   Begin VB.Menu dk2323 
      Caption         =   "&Zomm"
   End
   Begin VB.Menu fdkrep0 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu fl2323 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "explobco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

