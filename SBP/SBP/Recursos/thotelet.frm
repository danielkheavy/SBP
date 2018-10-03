VERSION 5.00
Begin VB.Form thotelet 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estado Habitaciones"
   ClientHeight    =   8820
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   16935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   16935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Menu de Opciones"
      Height          =   8655
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   2775
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Salida"
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
         Left            =   120
         Picture         =   "thotelet.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   2160
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Entrada"
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
         Left            =   120
         Picture         =   "thotelet.frx":0444
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CommandButton image1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Reservas"
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
         Left            =   120
         Picture         =   "thotelet.frx":0888
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Caracteristicas Habitacion"
      Height          =   8295
      Left            =   13200
      TabIndex        =   40
      Top             =   7920
      Visible         =   0   'False
      Width           =   13215
      Begin VB.TextBox vpiso 
         Height          =   615
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   52
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cerrar"
         Height          =   855
         Left            =   10320
         TabIndex        =   51
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox vcapacidad 
         Height          =   615
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   49
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox vtipo 
         Height          =   615
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   47
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox vprecio 
         Height          =   615
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   45
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox vdescripcio 
         Height          =   615
         Left            =   2040
         MaxLength       =   120
         TabIndex        =   41
         Top             =   1080
         Width           =   5655
      End
      Begin VB.TextBox vhabitacion 
         Height          =   615
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   43
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Piso"
         Height          =   615
         Left            =   120
         TabIndex        =   53
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Image foto 
         BorderStyle     =   1  'Fixed Single
         Height          =   3615
         Left            =   120
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   5655
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Capacidad"
         Height          =   615
         Left            =   120
         TabIndex        =   50
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label12 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   615
         Left            =   120
         TabIndex        =   48
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label11 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio"
         Height          =   615
         Left            =   120
         TabIndex        =   46
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label10 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
         Height          =   615
         Left            =   120
         TabIndex        =   44
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Habitacion"
         Height          =   615
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "EstadoHabitacion"
      Height          =   4095
      Left            =   15960
      TabIndex        =   29
      Top             =   7200
      Visible         =   0   'False
      Width           =   7335
      Begin VB.TextBox ehabitacion 
         Enabled         =   0   'False
         Height          =   495
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   37
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   495
         Left            =   6240
         TabIndex        =   36
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Grabar"
         Height          =   495
         Left            =   6240
         TabIndex        =   35
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Mantenimiento"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   1920
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Sucio"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Ocupado"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Libre"
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Habitacion"
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   0
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   1
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   2
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   3
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   4
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   5
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   6
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   7
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   8
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   9
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   10
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   11
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   12
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   13
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   14
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   15
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   16
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   17
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   18
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   19
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   20
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   21
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   22
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton groupmesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   23
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Limpieza"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15360
      TabIndex        =   39
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Libre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15360
      TabIndex        =   38
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mantenimiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15360
      TabIndex        =   28
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sucio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15360
      TabIndex        =   27
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label habitacion 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4080
      TabIndex        =   26
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ocupado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15360
      TabIndex        =   25
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HABITACIONES"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   24
      Top             =   0
      Width           =   8055
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   13440
      Picture         =   "thotelet.frx":0CCC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1320
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   12120
      Picture         =   "thotelet.frx":2C72
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1320
   End
   Begin VB.Menu flo994 
      Caption         =   "&Salir"
   End
   Begin VB.Menu dk4rep 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu procesos 
      Caption         =   "&Procesos"
      Visible         =   0   'False
      Begin VB.Menu checkin 
         Caption         =   "&1.CheckIn"
      End
      Begin VB.Menu habitax 
         Caption         =   "&2.Habitacion-Consumos"
      End
      Begin VB.Menu checkout 
         Caption         =   "&3.CheckOut"
      End
      Begin VB.Menu estadoc 
         Caption         =   "&4.EstadoCuenta"
      End
      Begin VB.Menu estadohabita 
         Caption         =   "&5.EstadoHabitacion"
      End
   End
End
Attribute VB_Name = "thotelet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mmesacod(15000) As String
Dim wmesacod(15000) As String
Dim wwmesacod(30) As String
Dim mmesapag As Integer
Dim mmesatop As Integer

Dim msalcod(100) As String
Dim msalpag As Integer
Dim msaltop As Integer
Option Explicit

Private Sub checkin_Click()
   tcheckin.xhabitacion = "" & habitacion
   tcheckin.Show 1
   'menu_carga_mesa
   'menu_mesa "INI"

End Sub

Private Sub checkout_Click()
 
 tcheckin.xhabitacion = "" & habitacion
 tcheckin.xsw = "SALIDA"
 tcheckin.Show 1
 'menu_carga_mesa
 'menu_mesa "INI"

End Sub

Private Sub Command1_Click()
If Len(Trim(ehabitacion)) = 0 Then Exit Sub
guarda_habitacion ehabitacion
Frame1.Visible = False
menu_mesa "INI"
End Sub

Private Sub Command2_Click()
Frame1.Visible = False
End Sub

Private Sub Command3_Click()
Frame2.Visible = False
End Sub


Private Sub dk4rep_Click()
If Frame2.Visible = True Then Exit Sub
trepohotel.Show 1

End Sub

Private Sub flo994_Click()
If Frame2.Visible = True Then
   Frame2.Visible = False
   Exit Sub
End If
thotelet.Hide
Unload thotelet
End Sub

Sub menu_carga_mesa(buf As String)
Dim mytablex As New ADODB.Recordset
Dim i As Integer
For i = 0 To 29
   wwmesacod(i) = ""
Next i
For i = 0 To 14999
    mmesacod(i) = ""
    wmesacod(i) = ""
Next i
i = -1
If mytablex.State = 1 Then mytablex.Close
If buf = "TODOS" Then
mytablex.Open "SELECT * FROM habitacion ", cn, adOpenDynamic, adLockOptimistic
Else
mytablex.Open "SELECT * FROM habitacion where piso='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

End If
Do
If mytablex.EOF Then Exit Do
   i = i + 1
   mmesacod(i) = "" & mytablex.Fields("habitacion")
   wmesacod(i) = "" & mytablex.Fields("Descripcio")
  
mytablex.MoveNext
Loop

mytablex.Close
mmesatop = i
mmesapag = 0

End Sub
Sub menu_mesa(buf As String)
Dim i As Integer
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
For i = mmesapag To 23 + mmesapag
    j = j + 1
    groupmesa(j).Caption = wmesacod(i) 'mmesacod(i)
    verifica_habitacion j, groupmesa(j).Caption
    'verifica_mesas j, groupmesa(j).Caption
Next i

End Sub
Sub verifica_habitacion(indx As Integer, buf1 As String)
Dim mytablex As New ADODB.Recordset
If Len(Trim(buf1)) = 0 Then Exit Sub
groupmesa(indx).BackColor = &H80FF80
mytablex.Open "select * from habitacion where habitacion='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      If "" & mytablex.Fields("estado") = "0" Then  'libre
      groupmesa(indx).BackColor = &H80FF80     'verde
      End If
      If "" & mytablex.Fields("estado") = "1" Then  'ocupado
      groupmesa(indx).BackColor = &H80FFFF     'amarillo
      End If
      If "" & mytablex.Fields("estado") = "2" Then  'sucio
      groupmesa(indx).BackColor = &HFF&         'rojo
      End If
      If "" & mytablex.Fields("estado") = "3" Then  'mantenimiento
      groupmesa(indx).BackColor = &H80FF&   'naranja
      End If
      If "" & mytablex.Fields("estado") = "4" Then  'Limpieza
      groupmesa(indx).BackColor = &HFFFF80    'cyan   '
      End If
      
   End If
End Sub


Private Sub Form_Load()
Dim i As Integer
For i = 0 To 23
  groupmesa(i).BackColor = &H80FF80
Next i

menu_carga_mesa "TODOS"
menu_mesa "INI"
End Sub
Sub carga_salon()
Dim mytablex As New ADODB.Recordset
Dim i As Integer
For i = 0 To 99
    msalcod(i) = ""
Next i
i = -1
mytablex.Open "select * from habitacionpiso ", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
i = i + 1
msalcod(i) = "" & mytablex.Fields("piso")
mytablex.MoveNext
Loop
msaltop = i
mytablex.Close
msalpag = 0
menu_salon "INI"

End Sub
Sub menu_salon(buf As String)
Dim i As Integer
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
For i = msalpag To 3 + msalpag
    j = j + 1
    groupsalon(j).Caption = msalcod(i)
Next i

End Sub

Sub verifica_mesas(indx As Integer, buf1 As String)
Dim mytablex As New ADODB.Recordset
Dim buf As String
'groupmesa(indx).BackColor = &HFFFFFF

buf = "select * from hotelcheckin where  HABITACION='" & buf1 & "'"
buf = buf & "  and arribofecha<='" & Format(Now, "YYYYMMDD") & "'"
buf = buf & "  and estado='0'"

'buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

If Len(buf1) > 0 Then
   mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      groupmesa(indx).BackColor = &HFF00&
   End If
   mytablex.Close
End If
End Sub

Private Sub groupmesa_Click(Index As Integer)
visualiza_habitacion "" & groupmesa(Index).Caption
End Sub

Private Sub groupmesa_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   habitacion = "" & groupmesa(Index).Caption
   PopupMenu procesos
End If
End Sub
Private Sub estadohabita_Click()
Dim mytablex As New ADODB.Recordset
ehabitacion = Trim("" & habitacion)
If Len(ehabitacion) = 0 Then Exit Sub
Frame1.Visible = True
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
mytablex.Open "select * from habitacion where habitacion='" & ehabitacion & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
      If "" & mytablex.Fields("estado") = "0" Then
      Option1.Value = True
      End If
     If "" & mytablex.Fields("estado") = "1" Then
     Option2.Value = True
      End If
      If "" & mytablex.Fields("estado") = "2" Then
      Option3.Value = True
      End If
      If "" & mytablex.Fields("estado") = "3" Then  'mantenimiento
      Option4.Value = True
      End If
      End If
      mytablex.Close
End Sub
Sub guarda_habitacion(buf As String)
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from habitacion where habitacion='" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If Option1.Value = True Then
      mytablex.Fields("estado") = "0"
   End If
   If Option2.Value = True Then
      mytablex.Fields("estado") = "1"
   End If
If Option3.Value = True Then
      mytablex.Fields("estado") = "2"
   End If
If Option4.Value = True Then
      mytablex.Fields("estado") = "3"
   End If
   mytablex.Update
End If
mytablex.Close
End Sub
Private Sub estadoc_Click()
tcheckin.xhabitacion = "" & habitacion
tcheckin.xsw = "PRECUENTA"
   tcheckin.Show 1
End Sub

Private Sub groupsalon_Click(Index As Integer)
Dim i As Integer
If Len(groupsalon(Index).Caption) = 0 Then Exit Sub
For i = 0 To 3
  groupsalon(i).BackColor = &HFFFFFF
Next i
For i = 0 To 23
  'groupmesa(i).BackColor = &HFFFFFF
  groupmesa(i).BackColor = &H80FF80
Next i
groupsalon(Index).BackColor = &HFF&
menu_carga_mesa groupsalon(Index).Caption
menu_mesa "INI" ', groupsalon(Index).Caption
piso = groupsalon(Index).Caption
'mesa = ""
'xindex = "" & Index

End Sub

Private Sub habitax_Click()
tcheckin.xhabitacion = "" & habitacion
tcheckin.xsw = "CONSUMO"
   tcheckin.Show 1
End Sub


Private Sub image2_Click()
Dim i As Integer
For i = 0 To 23
    groupmesa(i).BackColor = &H80FF80
    'mesa = ""
Next i

menu_mesa "SIG"

End Sub

Private Sub image3_Click()
Dim i As Integer
For i = 0 To 23
    groupmesa(i).BackColor = &H80FF80
    'mesa = ""
Next i

menu_mesa "ANT" ', salon

End Sub
Sub reporte()
Dim found As Integer
FileName = globaldir & "\temporal\" & gusuario & ".txt"
    borrar_archivo FileName
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento1
    cuerpo_programa_documento1
    '------------------------------------
    Close #1
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Sub cabecera_documento1()
Dim buf As String
Dim i As Integer
Dim found As Integer
    If contlin > 0 Then
       buf = Chr$(12)
       found = formateaa(buf, Len(buf), 0, 0)
    End If
    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = "Reporte de Habitaciones  "
    found = formateaa(buf, 90, 2, 0)
    
    found = formateaa("Estado", 8, 2, 0)
    found = formateaa("hab", 7, 0, 0)
    found = formateaa("Nombre", 20, 2, 0)
   
    
    buf = String(30, "-")
    found = formateaa(buf, 30, 2, 0)
    

End Sub
Sub cuerpo_programa_documento1()
Dim buf As String
Dim mytablex As New ADODB.Recordset
Dim found As Integer
Dim sdx As Double
Dim Tmp As String
Dim sw As Integer
Dim sdx1 As Double
Dim sdx2 As Double

On Error GoTo cmd78812_err
sw = 0
Tmp = ""
sdx = 0
sdx1 = 0
sdx2 = 0
mytablex.Open "select * from habitacion where estado<>'0' order by estado", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
If sw = 0 Then
      Tmp = "" & mytablex.Fields("estado")
      buf = "+" & mytablex.Fields("estado")
      found = formateaa(buf, 7, 0, 0)
      found = formateaa("", 1, 0, 0)
      If "" & mytablex.Fields("estado") = "1" Then
      found = formateaa("Ocupado", 10, 0, 0)
      End If
      If "" & mytablex.Fields("estado") = "2" Then
      found = formateaa("Sucio", 10, 0, 0)
      End If
      If "" & mytablex.Fields("estado") = "3" Then
      found = formateaa("Mantenimiento", 10, 0, 0)
      End If
      found = formateaa("", 1, 2, 0)
      nlineas
sw = 1
End If
If Tmp <> "" & mytablex.Fields("estado") Then
Tmp = "" & mytablex.Fields("estado")
      buf = "+" & mytablex.Fields("estado")
      found = formateaa(buf, 7, 0, 0)
      found = formateaa("", 1, 0, 0)
      If "" & mytablex.Fields("estado") = "1" Then
      found = formateaa("Ocupado", 10, 0, 0)
      End If
      If "" & mytablex.Fields("estado") = "2" Then
      found = formateaa("Sucio", 10, 0, 0)
      End If
      If "" & mytablex.Fields("estado") = "3" Then
      found = formateaa("Mantenimiento", 10, 0, 0)
      End If
      found = formateaa("", 1, 2, 0)
      nlineas
End If
If "" & mytablex.Fields("estado") = "1" Then
sdx = sdx + 1
End If
If "" & mytablex.Fields("estado") = "2" Then
sdx1 = sdx1 + 1
End If
If "" & mytablex.Fields("estado") = "3" Then
sdx2 = sdx2 + 1
End If

buf = "-" & mytablex.Fields("habitacion")
      found = formateaa(buf, 7, 0, 0)
      found = formateaa("", 1, 0, 0)
      buf = "" & mytablex.Fields("descripcio")
      found = formateaa(buf, 20, 0, 0)
      found = formateaa("", 1, 2, 0)
      nlineas
      mytablex.MoveNext
Loop
mytablex.Close
buf = "Ocupados      :" & sdx
found = formateaa(buf, 25, 2, 0)
nlineas
buf = "Sucios        :" & sdx1
found = formateaa(buf, 25, 2, 0)
nlineas
buf = "Mantenimiento :" & sdx
found = formateaa(buf, 25, 2, 0)
nlineas

Exit Sub
cmd78812_err:
MsgBox "Aviso en cuerpo " + error$, 48, "Aviso"
Exit Sub
End Sub
Sub nlineas()
    contlin = contlin + 1
    If contlin > 45 Then
       cabecera_documento1
    End If
End Sub

Private Sub Image5_Click()
Dim i As Integer
For i = 0 To 3
    groupsalon(i).BackColor = &HFFFFFF
    piso = ""
    'mesa = ""
Next i
For i = 0 To 23
    groupmesa(i).Caption = ""
    groupmesa(i).BackColor = &H80FF80
Next i

menu_salon "SIG"

End Sub

Private Sub Image6_Click()
Dim i As Integer
For i = 0 To 3
    groupsalon(i).BackColor = &HFFFFFF
    piso = ""
    'mesa = ""
Next i
For i = 0 To 23
    groupmesa(i).Caption = ""
    groupmesa(i).BackColor = &H80FF80
Next i

menu_salon "ANT"

End Sub

Sub visualiza_habitacion(buf As String)
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from habitacion where habitacion='" & buf & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   mytablex.Close
   Exit Sub
End If
vhabitacion = Trim("" & mytablex.Fields("habitacion"))
vdescripcio = Trim("" & mytablex.Fields("descripcio"))
vtipo = Trim("" & mytablex.Fields("tipohabitacion"))
vcapacidad = Trim("" & mytablex.Fields("capacidad"))
vpiso = Trim("" & mytablex.Fields("piso"))
vprecio = Trim("" & mytablex.Fields("precio"))
pone_fotonombre mytablex
mytablex.Close
Frame2.Visible = True
vhabitacion.SetFocus
End Sub
Sub pone_fotonombre(mytablex As ADODB.Recordset)
Dim buf As String
On Error GoTo cm897888_err
foto = LoadPicture()
buf = Trim("" & mytablex.Fields("habitacion"))
If Len(buf) > 0 Then
      viewBMP mytablex, buf
      If existe_archivo(globaldir & "\grafico\" & buf & ".jpg") > 0 Then
         foto = LoadPicture(globaldir & "\grafico\" & buf & ".jpg")
      End If
End If
Exit Sub
cm897888_err:
MsgBox "Aviso en pone_foto " + error$, 48, "Aviso"
Exit Sub
End Sub

