VERSION 5.00
Begin VB.Form cuadrege 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuadres Diarios"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text38 
      Height          =   375
      Left            =   7560
      MaxLength       =   10
      TabIndex        =   71
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text37 
      Height          =   375
      Left            =   6360
      MaxLength       =   10
      TabIndex        =   70
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text36 
      Height          =   375
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   69
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text35 
      Height          =   375
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   68
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text34 
      Height          =   375
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   66
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text33 
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   65
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text32 
      Height          =   375
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   63
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text31 
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   62
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text47 
      Height          =   375
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   53
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text46 
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   52
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text45 
      Height          =   375
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   50
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text44 
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   49
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text30 
      Height          =   375
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   47
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text29 
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   46
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text28 
      Height          =   375
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   44
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text27 
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   43
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text26 
      Height          =   375
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   41
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text25 
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   40
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Picture         =   "cuadrege.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.ComboBox grupos 
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
      Left            =   4080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text24 
      Height          =   375
      Left            =   7560
      MaxLength       =   10
      TabIndex        =   37
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text23 
      Height          =   375
      Left            =   6360
      MaxLength       =   10
      TabIndex        =   36
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text22 
      Height          =   375
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   35
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text21 
      Height          =   375
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   34
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text20 
      Height          =   375
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   32
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text19 
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   31
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      Height          =   375
      Left            =   7560
      MaxLength       =   10
      TabIndex        =   30
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Height          =   375
      Left            =   6360
      MaxLength       =   10
      TabIndex        =   29
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   7560
      MaxLength       =   10
      TabIndex        =   28
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Height          =   375
      Left            =   6360
      MaxLength       =   10
      TabIndex        =   27
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Height          =   375
      Left            =   7560
      MaxLength       =   10
      TabIndex        =   26
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   6360
      MaxLength       =   10
      TabIndex        =   24
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   23
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   22
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   21
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   20
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   19
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   17
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   13
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   12
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   9
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   7
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox fecha 
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   375
      Left            =   120
      TabIndex        =   67
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   375
      Left            =   120
      TabIndex        =   64
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Soles"
      Height          =   375
      Left            =   6360
      TabIndex        =   61
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dolares"
      Height          =   375
      Left            =   7560
      TabIndex        =   60
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Soles"
      Height          =   375
      Left            =   3840
      TabIndex        =   59
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dolares"
      Height          =   375
      Left            =   5040
      TabIndex        =   58
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Soles"
      Height          =   375
      Left            =   1320
      TabIndex        =   57
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dolares"
      Height          =   375
      Left            =   2520
      TabIndex        =   56
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      Height          =   375
      Left            =   120
      TabIndex        =   55
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Otros"
      Height          =   375
      Left            =   120
      TabIndex        =   54
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Egresos"
      Height          =   375
      Left            =   120
      TabIndex        =   51
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ingresos"
      Height          =   375
      Left            =   120
      TabIndex        =   48
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Compras"
      Height          =   375
      Left            =   120
      TabIndex        =   45
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ventas"
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Documentos"
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Otros"
      Height          =   375
      Left            =   120
      TabIndex        =   33
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Resultado"
      Height          =   375
      Left            =   6360
      TabIndex        =   25
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pasivo"
      Height          =   375
      Left            =   3840
      TabIndex        =   18
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Soles"
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dolares"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tarjetas"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Credito"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Efectivo"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Grupo"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Activo"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3720
      Width           =   3615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Menu fdlo34 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "cuadrege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fdlo34_Click()
cuadrege.Hide
Unload cuadrege
End Sub

Private Sub Form_Activate()
Dim mydbx As Database
Dim mytablex As Table
grupos.Clear
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("grupos")
mytablex.Index = "grupos"
mytablex.Seek "=", "01"
If Not mytablex.NoMatch Then
   grupos.AddItem "" & mytablex.Fields("l1")
   grupos.AddItem "" & mytablex.Fields("l2")
   grupos.AddItem "" & mytablex.Fields("l3")
   grupos.AddItem "" & mytablex.Fields("l4")
   grupos.ListIndex = 0
End If
mytablex.Close
mydbx.Close
End Sub

Private Sub Form_Load()
fecha = Format(Now, "dd/mm/yyyy")
End Sub
