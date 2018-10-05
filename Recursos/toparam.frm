VERSION 5.00
Begin VB.Form toparam 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametros Colores Comandos"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10890
   ForeColor       =   &H8000000A&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   10890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCommand2 
      Caption         =   "Command1"
      Height          =   360
      Left            =   1920
      TabIndex        =   45
      Top             =   360
      Width           =   990
   End
   Begin VB.ComboBox OPCION 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Grabar"
      Height          =   360
      Left            =   6360
      TabIndex        =   41
      Top             =   240
      Width           =   990
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Auto Servicio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pedido Delivery"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1815
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1065
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cuenta Separada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Delivery"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descto Pedido Actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Limpia Pedido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cierre X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Copia Ventas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Anula Venta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Borra Linea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Abre Gaveta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Des congela"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Congela"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "GRABA COMAN DAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   14
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comen tario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ver mesa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Master card"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ingreso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Egreso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cierre Z"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cash"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Visa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cortesia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Credito Clientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ve Delivery"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   25
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ruc Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   26
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Control Personal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   28
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   27
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Graba OT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   31
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Abono OT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   32
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Entrega OT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   33
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tablet Delivery"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   34
      Left            =   2640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Edición"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   3840
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   4425
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   360
         Left            =   1920
         TabIndex        =   42
         Top             =   1800
         Width           =   990
      End
      Begin VB.CommandButton Label1 
         Caption         =   "Guardar"
         Height          =   360
         Left            =   1920
         TabIndex        =   40
         Top             =   1440
         Width           =   990
      End
      Begin VB.TextBox estaxx 
         Height          =   375
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   4
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Habilitado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   195
         TabIndex        =   6
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label posicion 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OcultaMesas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   30
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6330
      Width           =   1215
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "TotMesas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   29
      Left            =   195
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6330
      Width           =   975
   End
   Begin VB.Label lblFamilia 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Opción"
      Height          =   375
      Left            =   3480
      TabIndex        =   44
      Top             =   240
      Width           =   810
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   120
      Picture         =   "toparam.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   840
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   120
      Picture         =   "toparam.frx":035D
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   840
   End
   Begin VB.Label caja 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Menu flo444 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "toparam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCerrar_Click()
    Frame1.Visible = False

End Sub

Sub OpcionTodos(FLAG As String, color As String)

    Dim a        As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from paramecacolor where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew

    End If

    mytablex.Fields("caja") = Trim(caja)

    For I = 0 To 34
        'estado de cada opcion
        a = "F" & I
        mytablex.Fields(a) = FLAG
          
        '    a = "C" & i
        '    I1 = Val("" & mytablex.Fields(a))
        '    a = "d" & i
        '    I2 = Val("" & mytablex.Fields(a))
        '    a = "e" & i
        '    I3 = Val("" & mytablex.Fields(a))
        If color = "D" Then xopciones(I).BackColor = RGB(80, 80, 80)
        If color = "H" Then xopciones(I).BackColor = RGB(219, 219, 219)
    Next I

    ''27/10/2017 Parametros de caja (Market, Restaurant)

    mytablex.Fields("colorfamilia1") = 180
    mytablex.Fields("colorfamilia2") = 210
    mytablex.Fields("colorfamilia3") = 255
    
    mytablex.Fields("colorproducto1") = 137
    mytablex.Fields("colorproducto2") = 135
    mytablex.Fields("colorproducto3") = 137
    ''27/10/2017 Parametros de caja (Market, Restaurant)

    mytablex.Update
    mytablex.Close

End Sub

'''27/10/2017 Parametros de caja (Market, Restaurant)
Sub OpcionCajaRestaurant()

    Dim I1       As Integer

    Dim I2       As Integer

    Dim I3       As Integer

    Dim I4       As Integer
    
    Dim a        As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from paramecacolor where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew

    End If
    
    mytablex.Fields("caja") = Trim(caja)
    
    For I = 0 To 34
        'estado de cada opcion
        a = "F" & I
        mytablex.Fields(a) = "S"
        xopciones(I).BackColor = RGB(80, 80, 80)
         
        If a = "F12" Or a = "F13" Or a = "F28" Or a = "F23" Or a = "F31" Or a = "F32" Or a = "F33" Or a = "F34" Then
            xopciones(I).BackColor = RGB(219, 219, 219)
            mytablex.Fields(a) = "N"

        End If

    Next I
    
    mytablex.Fields("colorfamilia1") = 180
    mytablex.Fields("colorfamilia2") = 210
    mytablex.Fields("colorfamilia3") = 255
    
    mytablex.Fields("colorproducto1") = 137
    mytablex.Fields("colorproducto2") = 135
    mytablex.Fields("colorproducto3") = 137
    
    mytablex.Update
    mytablex.Close

End Sub

Sub OpcionCajaMarket()

    Dim I1       As Integer

    Dim I2       As Integer

    Dim I3       As Integer

    Dim I4       As Integer
    
    Dim a        As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from paramecacolor where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew

    End If
    
    mytablex.Fields("caja") = Trim(caja)
    
    For I = 0 To 34
        'estado de cada opcion
        a = "F" & I
        mytablex.Fields(a) = "S"
        xopciones(I).BackColor = RGB(80, 80, 80)
         
        If a = "F28" Or a = "F2" Or a = "F14" Or a = "F15" Or a = "F16" Or a = "F23" Or a = "F31" Or a = "F32" Or a = "F33" Or a = "F34" Then
            xopciones(I).BackColor = RGB(219, 219, 219)
            mytablex.Fields(a) = "N"

        End If

    Next I
    
    mytablex.Fields("colorfamilia1") = 180
    mytablex.Fields("colorfamilia2") = 210
    mytablex.Fields("colorfamilia3") = 255
    
    mytablex.Fields("colorproducto1") = 137
    mytablex.Fields("colorproducto2") = 135
    mytablex.Fields("colorproducto3") = 137
    
    mytablex.Update
    mytablex.Close

End Sub

'''27/10/2017 Parametros de caja (Market, Restaurant)

Sub OpcionMozo()

    Dim I1       As Integer

    Dim I2       As Integer

    Dim I3       As Integer

    Dim I4       As Integer

    Dim a        As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from paramecacolor where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew

    End If

    mytablex.Fields("caja") = Trim(caja)

    For I = 0 To 34
        'estado de cada opcion
        a = "F" & I
        mytablex.Fields(a) = "N"
        xopciones(I).BackColor = RGB(219, 219, 219)
     
        If a = "F6" Or a = "F10" Or a = "F14" Or a = "F15" Or a = "F16" Or a = "F27" Or a = "F29" Then
            xopciones(I).BackColor = RGB(80, 80, 80)
            mytablex.Fields(a) = "S"

        End If

    Next I

    mytablex.Fields("colorfamilia1") = 180
    mytablex.Fields("colorfamilia2") = 210
    mytablex.Fields("colorfamilia3") = 255
    
    mytablex.Fields("colorproducto1") = 137
    mytablex.Fields("colorproducto2") = 135
    mytablex.Fields("colorproducto3") = 137

    mytablex.Update
    mytablex.Close

End Sub

Sub OpcionPedido()

    Dim a        As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from paramecacolor where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew

    End If

    mytablex.Fields("caja") = Trim(caja)

    For I = 0 To 34
        'estado de cada opcion
        a = "F" & I
        mytablex.Fields(a) = "N"
        xopciones(I).BackColor = RGB(219, 219, 219)
     
        If a = "F6" Or a = "F10" Or a = "F12" Or a = "F13" Or a = "F15" Or a = "F26" Or a = "F27" Or a = "F15" Or a = "F31" Or a = "F33" Then
            xopciones(I).BackColor = RGB(80, 80, 80)
            mytablex.Fields(a) = "S"

        End If
   
    Next I
    
    mytablex.Fields("colorfamilia1") = 180
    mytablex.Fields("colorfamilia2") = 210
    mytablex.Fields("colorfamilia3") = 255
    
    mytablex.Fields("colorproducto1") = 137
    mytablex.Fields("colorproducto2") = 135
    mytablex.Fields("colorproducto3") = 137

    mytablex.Update
    mytablex.Close

End Sub

Private Sub cmdCommand1_Click()

    If OPCION.ListIndex = 0 Then
        MsgBox ("Seleccione una opción") ' No Ha seleccionado nada
        Exit Sub

    End If

    If MsgBox("Desea Grabar???", vbExclamation + vbYesNo, "Eliminar") = vbNo Then
        Exit Sub
        
    End If

    '''27/10/2017 Parametros de caja (Market, Restaurant)
    If OPCION.ListIndex = 1 Then Call OpcionTodos("N", "H") ' Deshabilita Todo
    If OPCION.ListIndex = 2 Then Call OpcionTodos("S", "D") ' habilita Todo
    If OPCION.ListIndex = 3 Then Call OpcionCajaRestaurant
    If OPCION.ListIndex = 4 Then Call OpcionCajaMarket
    If OPCION.ListIndex = 5 Then Call OpcionMozo
    If OPCION.ListIndex = 6 Then Call OpcionPedido
    '''27/10/2017 Parametros de caja (Market, Restaurant)

End Sub

Private Sub cmdCommand2_Click()

    Dim I        As Integer

    Dim a        As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from paramecacolor where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    For I = 0 To 34
           
        a = "F" & I
        
        If Mid(mytablex.Fields(a), 1, 1) = "S" Then xopciones(I).BackColor = RGB(80, 80, 80)
        If Mid(mytablex.Fields(a), 1, 1) = "N" Then xopciones(I).BackColor = RGB(219, 219, 219)
          
    Next I

    mytablex.Update
    mytablex.Close

End Sub

Private Sub flo444_Click()
    toparam.Hide
    Unload toparam

End Sub

Private Sub Form_Activate()

    Dim I        As Integer

    Dim a        As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from paramecacolor where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    For I = 0 To 34
           
        a = "F" & I

        If Mid(mytablex.Fields(a), 1, 1) = "S" Then
            xopciones(I).BackColor = RGB(80, 80, 80)
        Else
            xopciones(I).BackColor = RGB(219, 219, 219)

        End If

    Next I

    mytablex.Update
    mytablex.Close

End Sub

Private Sub Form_Load()
    'HS1.Min = 0
    'HS1.max = 255
    'HS1.LargeChange = 25
    'HS1.SmallChange = 5
    '
    'hs2.Min = 0
    'hs2.max = 255
    'hs2.LargeChange = 25
    'hs2.SmallChange = 5
    '
    'hs3.Min = 0
    'hs3.max = 255
    'hs3.LargeChange = 25
    'hs3.SmallChange = 5

    OPCION.Clear
    OPCION.AddItem "SELECCIONE OPCIÓN"
    OPCION.AddItem "Deshabilitar Todos"
    OPCION.AddItem "Habilitar Todos"
    OPCION.AddItem "Parametros Caja Restaurant"
    OPCION.AddItem "Parametros Caja Market"
    OPCION.AddItem "Parametros Mozo"
    OPCION.AddItem "Parametros Pedido"
    OPCION.ListIndex = 0

End Sub

Private Sub Label1_Click()

    Dim a        As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from paramecacolor where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew

    End If

    mytablex.Fields("caja") = Trim(caja)
    a = "c" & posicion
    'mytablex.Fields(a) = "" & HS1.Value
    a = "d" & posicion
    'mytablex.Fields(a) = "" & hs2.Value
    a = "e" & posicion
    'mytablex.Fields(a) = "" & hs3.Value
    a = "f" & posicion
    mytablex.Fields(a) = estaxx
    mytablex.Update
    mytablex.Close
    Frame1.Visible = False
    Form_Activate

End Sub

Private Sub xopciones_Click(Index As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim a        As String

    Frame1.Visible = True
    posicion = Index
    mytablex.Open "select * from paramecacolor where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        'MsgBox Index
        a = "c" & Index
        'HS1.Value = Val("" & mytablex.Fields(a))
        a = "d" & Index
        'hs2.Value = Val("" & mytablex.Fields(a))
        a = "e" & Index
        'hs3.Value = Val("" & mytablex.Fields(a))
        a = "f" & Index
        estaxx = "" & mytablex.Fields(a)

    End If

    mytablex.Close

End Sub

