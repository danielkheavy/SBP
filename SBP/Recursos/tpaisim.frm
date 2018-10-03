VERSION 5.00
Begin VB.Form tpaisim 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Costos x Proveedor"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   10905
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command9 
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
      Height          =   855
      Left            =   9480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tpaisim.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Salir"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command10 
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
      Height          =   855
      Left            =   9480
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tpaisim.frx":1212
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Grabar registro"
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox proveedor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   120
      Width           =   5895
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6720
      MaxLength       =   1
      TabIndex        =   15
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6720
      MaxLength       =   1
      TabIndex        =   13
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6720
      MaxLength       =   1
      TabIndex        =   11
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox monedac 
      Height          =   375
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proveedor"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CategoriaArancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha FOB"
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaCompra"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ParteArancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Precio FOB"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PrecioCompra"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mon.FOB"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda Compra"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
   Begin VB.Menu dlo232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tpaisim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command9_Click()
    dlo232_Click

End Sub

Private Sub dlo232_Click()
    tpaisim.Hide
    Unload tpaisim

End Sub
