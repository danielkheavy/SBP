VERSION 5.00
Begin VB.Form repdocum 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Documentos"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox vedelivery 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   66
      Top             =   1690
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   64
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox mesa 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   63
      Text            =   "%"
      Top             =   6360
      Width           =   855
   End
   Begin VB.ComboBox salon 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   60
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox tipofecha 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   58
      Text            =   "E"
      Top             =   2760
      Width           =   375
   End
   Begin VB.ComboBox tipores 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   5640
      Width           =   1575
   End
   Begin VB.ComboBox servicio 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   53
      Top             =   7920
      Width           =   1575
   End
   Begin VB.ComboBox comopaga 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox local1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox horaf 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   3480
      Width           =   1575
   End
   Begin VB.ComboBox horai 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   3120
      Width           =   1575
   End
   Begin VB.ComboBox cajero 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   4080
      Width           =   1575
   End
   Begin VB.ComboBox turno 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   4800
      Width           =   1575
   End
   Begin VB.ComboBox caja 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   4440
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   7560
      Width           =   3855
   End
   Begin VB.ComboBox grupos 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox nombre 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   33
      Text            =   "%"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.ComboBox estado 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   2070
      Width           =   1575
   End
   Begin VB.ComboBox vfpago 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   1330
      Width           =   1575
   End
   Begin VB.ComboBox vdetalle 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   970
      Width           =   1575
   End
   Begin VB.ComboBox bodega 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox fpago 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   23
      Text            =   "%"
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox transporte 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   21
      Text            =   "%"
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox vendedor 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   19
      Text            =   "%"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox codigo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   8
      Text            =   "%"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox numero 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   11
      TabIndex        =   7
      Text            =   "%"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox tipo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox serie 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   3
      TabIndex        =   5
      Text            =   "%"
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox fechai 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox fechaf 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   3
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox titulo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   2
      Top             =   6840
      Width           =   3855
   End
   Begin VB.TextBox nrolineas 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   1
      Text            =   "45"
      Top             =   7200
      Width           =   1575
   End
   Begin VB.ComboBox moneda 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ver Datos Delivery"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   67
      Top             =   1700
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Impresion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   65
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label31 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Salon"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   62
      Top             =   6000
      Width           =   2175
   End
   Begin VB.Label Label30 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mesa/Seccion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   61
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label29 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha (E)mision (V)encimiento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   59
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label28 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   57
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label donde 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4200
      TabIndex        =   55
      Top             =   7920
      Width           =   105
   End
   Begin VB.Label Label27 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Servicio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   54
      Top             =   7920
      Width           =   2175
   End
   Begin VB.Label Label26 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ver como Paga"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   52
      Top             =   5280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label25 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   50
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label24 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HoraFinal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   48
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label23 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HoraInicio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   47
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label22 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cajero"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   44
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Turno"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   42
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label19 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   40
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label18 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Agrupacion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Grupos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   36
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   34
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   32
      Top             =   2070
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VerFormaPago"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   30
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VerDetalle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   28
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Almacen"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FormaPago"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Transportista"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label21 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label17 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Titulo reporte"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lineas x Pagina"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label acu 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   600
      Width           =   255
   End
   Begin VB.Menu ejui23 
      Caption         =   "&Ejecutar"
   End
   Begin VB.Menu dlo232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "repdocum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dlo232_Click()
    repdocum.Hide
    Unload repdocum

End Sub

Sub sumar_percepcion(mytabley As ADODB.Recordset, _
                     sdx As Double, _
                     sdx1 As Double, _
                     sdx2 As Double)

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    Dim found    As Integer

    sdx = 0
    sdx1 = 0
    sdx2 = 0

    mytablex.Open "select * from " & dgusuariog & " where local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "' and numero='" & "" & mytabley.Fields("numero") & "' and percepcion>0", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + Val("" & mytablex.Fields("total"))
        sdx1 = sdx1 + Val("" & mytablex.Fields("tpercepcio"))
        sdx2 = sdx2 + Val("" & mytablex.Fields("total")) + Val("" & mytablex.Fields("tpercepcio"))
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Sub menu_percepcion()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    contpag = 0
    found = sql_documento(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    'MsgBox filename
    cerrar_archivo
    found = borra_nombre("" & FileName)
    'MsgBox filename
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento
    cuerpo_programa_documento mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    found = valida_wordpad(FileName)

End Sub

Sub menu_comision()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    contpag = 0
    found = sql_documento(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    'MsgBox filename
    cerrar_archivo
    found = borra_nombre("" & FileName)
    'MsgBox filename
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento
    cuerpo_programa_documento mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    found = valida_wordpad(FileName)

End Sub

Private Sub ejui23_Click()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    'MsgBox opcion2
    If opcion2 = "455" Then
        menu_codigo
        Exit Sub

    End If

    If opcion2 = "456" Then
        menu_totales
        Exit Sub

    End If

    If opcion2 = "100" Then   'si quiere estadisticas x meses
        menu_percepcion
        Exit Sub

    End If

    If opcion2 = "4000" Then   'si quiere estadisticas x meses
        menu_comision
        Exit Sub

    End If

    If opcion2 = "10" Or opcion2 = "11" Then   'si quiere estadisticas x meses
        proceso_venta_diario
        Exit Sub

    End If

    If opcion2 = "12" Then   'si quiere estadisticas x meses
        menu_meses
        Exit Sub

    End If

    If opcion2 = "14" Then   'si quiere estadisticas x meses
        menu_dias
        Exit Sub

    End If

    If opcion2 = "13" Then  'semanal
        menu_semanas
        Exit Sub

    End If

    If opcion2 = "45" Then  'semanal
        menu_codigo
        Exit Sub

    End If

    If Combo1 = "Fecha" Then
        MsgBox "Ya el reporte sale por fechas elija otro ", 48, "Aviso"
        Exit Sub

    End If

    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    contpag = 0

    '''10/10/2017 Reporte de Seguimiento de facturas En Excel
    If Combo3 = "NORMAL" Then
        SeguimientoPantalla

    End If
 
    If Combo3 = "EXCELL" Then
        SeguimientoExcel

    End If

    '''10/10/2017 Reporte de Seguimiento de facturas En Excel

End Sub

' ''10/10/2017 Reporte de Seguimiento de facturas En Excel
Sub SeguimientoPantalla()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    found = sql_documento(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    'MsgBox filename
    cerrar_archivo
    found = borra_nombre("" & FileName)
    'MsgBox filename
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento
    cuerpo_programa_documento mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
    found = valida_wordpad(FileName)
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
End Sub

Sub SeguimientoExcel()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    found = sql_documento(mytablex)

    If found = 0 Then
        mytablex.Close
        Exit Sub

    End If

    cuerpo_programa_seguimientoExcel mytablex

End Sub

Function ObtieneFormaPago(ByRef ftipo As String, resultado As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT descripcio FROM fpago  where  fpago='" & ftipo & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        resultado = ("" & mytablex.Fields("descripcio"))
        mytablex.MoveNext

    End If

    mytablex.Close
   
End Function

Sub cuerpo_programa_seguimientoExcel(mytablex As ADODB.Recordset)

    Dim vr

    Dim sw1          As Integer

    Dim Tmp          As String

    Dim tmp1         As String

    Dim sw           As Integer

    Dim buf          As String

    Dim found        As Integer

    Dim sdx          As Double

    Dim mytabley     As New ADODB.Recordset

    Dim sdx1         As Double

    Dim sdxtmp       As Double

    Dim v            As Long

    Dim h            As Integer

    Dim vprecios(10) As String

    Dim Heading(31)  As String

    Dim Heading2(13) As String

    Dim resultado    As String
    
    h = 1
    sdx1 = 0
    sdx = 0
    sw = 0
    v = 4
    h = 1
    suma1 = 0
    ssuma1 = 0
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    
    Heading(1) = Combo1
    Heading(2) = "Fecha"
    Heading(3) = "Tipo"
    Heading(4) = "Serie"
    Heading(5) = "Numero"
    Heading(6) = "Cod.Cliente"
    Heading(7) = "Nombre"
    Heading(8) = "M"
    
    Heading(9) = "Vendedor"
    Heading(10) = "Cajero"
    Heading(11) = "Total"
    Heading(12) = "Hora"
    Heading(13) = "Servicio"
    Heading(14) = "Caja"
    Heading(15) = "Turno"
    Heading(16) = "Estado"
    Heading(17) = "Mesa"
    Heading(18) = "Pers."
    Heading(19) = "Desc."
    Heading(20) = "Observación"

    If vfpago = "S" And vedelivery = "S" Then
        ObtieneFormaPago 1, resultado
        Heading(21) = resultado

        ObtieneFormaPago 2, resultado
        Heading(22) = resultado

        ObtieneFormaPago 3, resultado
        Heading(23) = resultado

        ObtieneFormaPago 4, resultado
        Heading(24) = resultado

        ObtieneFormaPago 5, resultado
        Heading(25) = resultado
        Heading(26) = "OTROS"
    
        Heading(28) = "Teléfono"
        Heading(29) = "Dirección"
        Heading(30) = "Referencia"

    End If

    If vfpago = "S" And vedelivery = "N" Then
        ObtieneFormaPago 1, resultado
        Heading(21) = resultado

        ObtieneFormaPago 2, resultado
        Heading(22) = resultado

        ObtieneFormaPago 3, resultado
        Heading(23) = resultado

        ObtieneFormaPago 4, resultado
        Heading(24) = resultado

        ObtieneFormaPago 5, resultado
        Heading(25) = resultado
        Heading(26) = "OTROS"

    End If

    If vfpago = "N" And vedelivery = "S" Then
        Heading(21) = "Teléfono"
        Heading(22) = "Dirección"
        Heading(23) = "Referencia"

    End If
   
    If vdetalle = "N" And vfpago = "N" And vedelivery = "N" Then
        Call Formato_ExcelSeguimiento(20, Heading())
    Else
        
        If vdetalle = "S" Then
            Heading2(6) = "*Cod.Producto"
            Heading2(7) = "Descripción"
            Heading2(8) = "Und"
            Heading2(9) = "Precio"
            Heading2(10) = "Cantidad"
            Heading2(11) = "Total"
            Heading2(12) = "** Comentario                                                 "
            
            '08/05/2018 Reporte Pedidos Orden de Trabajo en Excel
            If vfpago = "N" Then Call Formato_ExcelSeguimiento(20, Heading())
            '08/05/2018 Reporte Pedidos Orden de Trabajo en Excel
            
            If vfpago = "S" Then Call Formato_ExcelSeguimiento(26, Heading())
            
            Call Formato_ExcelSeguimientoDetalle(16, Heading2())
            v = v + 1

        End If
          
        If vfpago = "S" And vedelivery = "S" Then
            Call Formato_ExcelSeguimiento(30, Heading())

        End If
        
        If vfpago = "S" And vedelivery = "N" Then
            If vdetalle <> "S" Then Call Formato_ExcelSeguimiento(26, Heading())

        End If
        
        If vfpago = "N" And vedelivery = "S" Then
            If vdetalle <> "S" Then Call Formato_ExcelSeguimiento(23, Heading())

        End If

    End If
    
    objExcel.ActiveSheet.Cells(1, 6) = "     SEGUIMIENTO DE COMPROBANTES"
    objExcel.ActiveSheet.Cells(1, 6).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 6).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 6).Font.color = RGB(0, 112, 184)
    
    objExcel.ActiveSheet.Cells(2, 1) = "FECHA INICIO  " + fechai
    objExcel.ActiveSheet.Cells(2, 4) = "FECHA FIN  " + fechaf
    tmp1 = ""

    Do

        If mytablex.EOF Then Exit Do
        vr = DoEvents()

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("Usuario")

        End If

        If Combo1 = "Caja" Then
            tmp1 = "" & mytablex.Fields("Caja")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("Turno")

        End If

        If Combo1 = "Ccosto" Then
            tmp1 = "" & mytablex.Fields("Ccosto")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("bodega")

        End If

        If Combo1 = "TipoDocumento" Then
            tmp1 = "" & mytablex.Fields("tipo")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("codigo")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("usuario")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("turno")

        End If

        If Combo1 = "Local" Then
            tmp1 = "" & mytablex.Fields("local")

        End If

        If Combo1 = "Fecha" Then
            tmp1 = "" & mytablex.Fields("Fecha")

        End If

        If Combo1 = "Hora" Then
            tmp1 = "" & mytablex.Fields("Hora1")

        End If

        If Combo1 = "Servicio" Then
            tmp1 = "" & mytablex.Fields("Servicio")

        End If

        If Combo1 = "Mesa" Then
            tmp1 = "" & mytablex.Fields("Mesa")

        End If

        If sw = 0 Then
  
            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                objExcel.ActiveSheet.Cells(v, h) = "Caja: " & buf
                v = v + 1
                nlineas
                Tmp = "" & mytablex.Fields("Caja")

            End If
   
            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " " & busca_bodega(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("bodega")

            End If
     
            If Combo1 = "TipoDocumento" Then
                buf = "" & mytablex.Fields("tipo")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " " & busca_tipo(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("tipo")

            End If
   
            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("Codigo")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_cliente(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Codigo")

            End If
   
            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("Vendedor")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Vendedor")

            End If
   
            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " " & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("usuario")

            End If
   
            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("turno")
                objExcel.ActiveSheet.Cells(v, h) = "Turno: " & buf & " "
                v = v + 1
                Tmp = "" & mytablex.Fields("turno")

            End If
   
            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " " & busca_localx(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Local")

            End If
   
            If Combo1 = "Fecha" Then
                buf = "" & mytablex.Fields("Fecha")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                v = v + 1
                Tmp = "" & mytablex.Fields("Fecha")

            End If
   
            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("Hora1")
                objExcel.ActiveSheet.Cells(v, h) = "Hora: " & buf & " "
                v = v + 1
                Tmp = "" & mytablex.Fields("Hora1")

            End If
   
            If Combo1 = "Servicio" Then
                buf = "" & mytablex.Fields("Servicio")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                v = v + 1
                Tmp = "" & mytablex.Fields("Servicio")

            End If
   
            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("Mesa")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                v = v + 1
                Tmp = "" & mytablex.Fields("Mesa")

            End If
  
            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                v = v + 1
                Tmp = "" & mytablex.Fields("Zona")

            End If
     
            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                objExcel.ActiveSheet.Cells(v, h) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("vendedor")

            End If
   
            objExcel.ActiveSheet.Cells(v - 1, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v - 1, h).Font.color = RGB(62, 95, 138)
   
            objExcel.ActiveSheet.Cells(v - 1, h + 1).Font.bold = True
            objExcel.ActiveSheet.Cells(v - 1, h + 1).Font.color = RGB(62, 95, 138)
   
            sw = 1
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0

        End If

        If Tmp <> tmp1 Then

            buf = Format(suma1, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 10) = "" & buf
            objExcel.ActiveSheet.Cells(v, h + 10).Font.bold = True
  
            v = v + 1
 
            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                objExcel.ActiveSheet.Cells(v, h) = "Caja: " & buf
                v = v + 1
                nlineas
                Tmp = "" & mytablex.Fields("Caja")

            End If
   
            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " " & busca_bodega(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("bodega")

            End If
     
            If Combo1 = "TipoDocumento" Then
                buf = "" & mytablex.Fields("tipo")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " " & busca_tipo(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("tipo")

            End If
   
            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("Codigo")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & ""
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_cliente(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Codigo")

            End If
   
            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("Vendedor")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Vendedor")

            End If
   
            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                objExcel.ActiveSheet.Cells(v, h + 1) = "'" & busca_vendedor(buf)
   
                v = v + 1
                Tmp = "" & mytablex.Fields("usuario")

            End If
   
            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("turno")
                objExcel.ActiveSheet.Cells(v, h) = "Turno: " & buf & " "
                v = v + 1
                Tmp = "" & mytablex.Fields("turno")

            End If
   
            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " " & busca_localx(buf)
                v = v + 1
                Tmp = "" & mytablex.Fields("Local")

            End If
   
            If Combo1 = "Fecha" Then
                buf = "" & mytablex.Fields("Fecha")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                v = v + 1
                Tmp = "" & mytablex.Fields("Fecha")

            End If
   
            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("Hora1")
                objExcel.ActiveSheet.Cells(v, h) = "Hora: " & buf & " "
                v = v + 1
                Tmp = "" & mytablex.Fields("Hora1")

            End If
   
            If Combo1 = "Servicio" Then
                buf = "" & mytablex.Fields("Servicio")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                v = v + 1
                Tmp = "" & mytablex.Fields("Servicio")

            End If
   
            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("Mesa")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                v = v + 1
                Tmp = "" & mytablex.Fields("Mesa")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                objExcel.ActiveSheet.Cells(v, h) = "'" & buf & " "
                v = v + 1
                Tmp = "" & mytablex.Fields("Zona")

            End If

            objExcel.ActiveSheet.Cells(v - 1, h).Font.bold = True
            objExcel.ActiveSheet.Cells(v - 1, h).Font.color = RGB(62, 95, 138)
   
            objExcel.ActiveSheet.Cells(v - 1, h + 1).Font.bold = True
            objExcel.ActiveSheet.Cells(v - 1, h + 1).Font.color = RGB(62, 95, 138)
 
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0

        End If
        
        buf = "'" & mytablex.Fields("fecha")
        objExcel.ActiveSheet.Cells(v, h + 1) = "" & buf
        
        buf = "" & mytablex.Fields("tipo")
        objExcel.ActiveSheet.Cells(v, h + 2) = "" & buf

        buf = "" & mytablex.Fields("serie")
        objExcel.ActiveSheet.Cells(v, h + 3) = "" & buf
        
        buf = "" & mytablex.Fields("numero")
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & buf
        
        buf = "'" & mytablex.Fields("codigo")
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & buf
        
        buf = "'" & mytablex.Fields("nombre")
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & buf

        buf = "" & mytablex.Fields("moneda")
        objExcel.ActiveSheet.Cells(v, h + 7) = "" & buf
           
        buf = "'" & mytablex.Fields("vendedor")
        objExcel.ActiveSheet.Cells(v, h + 8) = "" & buf

        buf = "'" & mytablex.Fields("usuario")
        objExcel.ActiveSheet.Cells(v, h + 9) = "" & buf
                
        buf = "" & mytablex.Fields("total")
        objExcel.ActiveSheet.Cells(v, h + 10) = "" & buf
        
        suma1 = suma1 + Val("" & mytablex.Fields("total"))
        ssuma1 = ssuma1 + Val("" & mytablex.Fields("total"))
   
        buf = "" & mytablex.Fields("hora")
        objExcel.ActiveSheet.Cells(v, h + 11) = "" & buf
        
        buf = "" & mytablex.Fields("servicio")
        objExcel.ActiveSheet.Cells(v, h + 12) = "" & buf
              
        buf = "'" & mytablex.Fields("caja")
        objExcel.ActiveSheet.Cells(v, h + 13) = "" & buf
        
        buf = "" & mytablex.Fields("turno")
        objExcel.ActiveSheet.Cells(v, h + 14) = "" & buf
        
        buf = "" & mytablex.Fields("estado")
        objExcel.ActiveSheet.Cells(v, h + 15) = "" & buf
       
        buf = "'" & mytablex.Fields("mesa")
        objExcel.ActiveSheet.Cells(v, h + 16) = "" & buf
        
        '08/05/2018 Reporte Pedidos Orden de Trabajo en Excel
        'buf = "" & mytablex.Fields("personas")
        'objExcel.ActiveSheet.Cells(v, h + 17) = "" & buf
            
        If mytablex.Fields("acu") <> "I" Then
            buf = "" & mytablex.Fields("personas")
            objExcel.ActiveSheet.Cells(v, h + 17) = "" & buf

        End If

        '08/05/2018 Reporte Pedidos Orden de Trabajo en Excel
        
        buf = "" & mytablex.Fields("descuento")
        objExcel.ActiveSheet.Cells(v, h + 18) = "" & buf
             
        '08/05/2018 Reporte Pedidos Orden de Trabajo en Excel
        objExcel.ActiveSheet.Cells(v, h + 19) = "" & mytablex.Fields("observa")
        '08/05/2018 Reporte Pedidos Orden de Trabajo en Excel
             
        If vfpago = "S" Then
            sumar_como_pagaExcel mytablex, v, h + 20

        End If
      
        ''02/11/2017 Reporte de Seguimiento de facturas incluye delivery
        If vedelivery = "S" And vfpago = "S" Then
            buf = mytablex.Fields("codigo")
            objExcel.ActiveSheet.Cells(v, h + 27) = "'" & busca_datosdelivery(buf, 0) 'Telefono
            objExcel.ActiveSheet.Cells(v, h + 28) = "'" & busca_datosdelivery(buf, 1) 'Direccion
            objExcel.ActiveSheet.Cells(v, h + 29) = "'" & busca_datosdelivery(buf, 2) 'Referencia

        End If
       
        If vedelivery = "S" And vfpago = "N" Then
            buf = mytablex.Fields("codigo")
            objExcel.ActiveSheet.Cells(v, h + 20) = "'" & busca_datosdelivery(buf, 0) 'Telefono
            objExcel.ActiveSheet.Cells(v, h + 21) = "'" & busca_datosdelivery(buf, 1) 'Direccion
            objExcel.ActiveSheet.Cells(v, h + 22) = "'" & busca_datosdelivery(buf, 2) 'Referencia

        End If
       
        ''02/11/2017 Reporte de Seguimiento de facturas incluye delivery

        If vdetalle = "S" Then

            Dim I As Integer

            Dim m As Integer
        
            If vfpago = "S" Then
                m = 25
            Else
                m = 18

            End If

            For I = 1 To m
                objExcel.ActiveSheet.Cells(v, I).Font.bold = True
                objExcel.ActiveSheet.Cells(v, I).Interior.color = RGB(232, 232, 232)
            Next
        
            ver_detalleExcel mytablex, v, h
            v = fin
            v = v + 1

        End If
 
        sdxtmp = 0

        v = v + 1

seguy13:
        mytablex.MoveNext
    Loop

    sw1 = 0
   
    buf = Format(suma1, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 10) = "" & buf
    objExcel.ActiveSheet.Cells(v, h + 10).Font.bold = True
     
    v = v + 1
        
    objExcel.ActiveSheet.Cells(v, h + 8) = "Gran Total"
          
    buf = Format(ssuma1, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 10) = "" & buf

    Dim k As Integer

    For k = 9 To 11
        objExcel.ActiveSheet.Cells(v, k).Font.bold = True
        objExcel.ActiveSheet.Cells(v, k).Interior.color = RGB(248, 243, 53)
    Next
  
    v = v + 1
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto

End Sub

Sub ver_detalleExcel(mytabley As ADODB.Recordset, ByVal a As Integer, ByVal b As Integer)

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    Dim found    As Integer

    fin = 0

    mytablex.Open "select * from " & dgusuariog & " where acu<>'T' and local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "' and numero='" & "" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    sw = 0
    Do

        If mytablex.EOF Then Exit Do
        sw = 1
        a = a + 1
   
        objExcel.ActiveSheet.Cells(a, b + 5) = "*" & mytablex.Fields("PRODUCTO")
        objExcel.ActiveSheet.Cells(a, b + 6) = "" & mytablex.Fields("DESCRIPCIO")
        objExcel.ActiveSheet.Cells(a, b + 7) = "" & mytablex.Fields("UNIDAD")
        objExcel.ActiveSheet.Cells(a, b + 8) = "" & mytablex.Fields("PRECIO")
        objExcel.ActiveSheet.Cells(a, b + 9) = "" & mytablex.Fields("CANTIDAD")
        objExcel.ActiveSheet.Cells(a, b + 10) = "" & mytablex.Fields("total")
 
        ''' kenyo 09/11/2017 Mejora grupo comentario
        objExcel.ActiveSheet.Cells(a, b + 11) = "** " & mytablex.Fields("observa1")
        ''' kenyo 09/11/2017 Mejora grupo comentario

        mytablex.MoveNext
    Loop
    fin = a
    mytablex.Close

End Sub

Sub sumar_como_pagaExcel(mytabley As ADODB.Recordset, _
                         ByVal a As Integer, _
                         ByVal b As Integer)

    Dim found    As Integer

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim sdx2     As Double

    Dim sdx3     As Double

    Dim sdx4     As Double

    Dim sdx5     As Double

    sdx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    sdx4 = 0
    sdx5 = 0

    mytablex.Open "select * from fpagov where local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "' and numero='" & "" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            If "" & mytablex.Fields("local") = "" & mytabley.Fields("local") And "" & mytablex.Fields("tipo") = "" & mytabley.Fields("tipo") And "" & mytablex.Fields("serie") = "" & mytabley.Fields("serie") And "" & mytablex.Fields("numero") = "" & mytabley.Fields("numero") Then

                Select Case "" & mytablex.Fields("FPAGO") ' acufp

                        'reemplzamos total por recibe
                    Case "1"   'A Efectivo
                        sdx = sdx + Val("" & mytablex.Fields("recibe"))
                        suma6 = suma6 + Val("" & mytablex.Fields("recibe"))

                    Case "2"   'B Dolares
                        sdx1 = sdx1 + Val("" & mytablex.Fields("recibe"))
                        suma7 = suma7 + Val("" & mytablex.Fields("recibe"))

                    Case "3"   'C Credito
                        sdx2 = sdx2 + Val("" & mytablex.Fields("recibe"))
                        suma8 = suma8 + Val("" & mytablex.Fields("recibe"))

                    Case "4"   'D TARJET DE CRED // visa
                        sdx4 = sdx4 + Val("" & mytablex.Fields("recibe"))
                        suma10 = suma10 + Val("" & mytablex.Fields("recibe"))

                    Case "5"   ' MASTERCARD
                        sdx5 = sdx5 + Val("" & mytablex.Fields("recibe"))
                
                    Case Else
                        sdx3 = sdx3 + Val("" & mytablex.Fields("recibe"))
                        suma9 = suma9 + Val("" & mytablex.Fields("recibe"))

                End Select

            Else
                Exit Do

            End If

            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
   
    buf = Format(sdx, "0.00") 'EFECTIVO
    objExcel.ActiveSheet.Cells(a, b) = "" & buf
   
    buf = Format(sdx1, "0.00") 'DOLARES
    objExcel.ActiveSheet.Cells(a, b + 1) = "" & buf

    buf = Format(sdx2, "0.00") 'CREDITO
    objExcel.ActiveSheet.Cells(a, b + 2) = "" & buf
   
    buf = Format(sdx4, "0.00") 'VISA
    objExcel.ActiveSheet.Cells(a, b + 3) = "" & buf
     
    buf = Format(sdx5, "0.00") 'MASTERCARD
    objExcel.ActiveSheet.Cells(a, b + 4) = "" & buf
   
    buf = Format(sdx3, "0.00") ' OTROS
    objExcel.ActiveSheet.Cells(a, b + 5) = "" & buf
   
End Sub

' ''10/10/2017 Reporte de Seguimiento de facturas En Excel

Private Sub Form_Activate()
    ReDim xmeses(13) As Double
    ReDim xmeses1(13) As Double

    Dim mytablex As New ADODB.Recordset

    If donde = "FECHA" Then
        Combo1.ListIndex = 9

    End If

    If donde = "HORA" Then
        Combo1.ListIndex = 10

    End If

    If donde = "Vendedor" Then
        Combo1.ListIndex = 4

    End If

    salon.Clear
    salon.AddItem "%"
    mytablex.Open "select * from salon", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        salon.AddItem "" & mytablex.Fields("salon") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    salon.ListIndex = 0

    tipo.Clear
    tipo.AddItem "%"

    mytablex.Open "select * from tipo", cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("grupo") = acu Then
            tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close

    tipo.ListIndex = 0
    caja.Clear
    caja.AddItem "%"
    mytablex.Open "select * from parameca ", cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("terminal") = "C" Then
            caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("descripcio")

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    caja.ListIndex = 0

    turno.Clear
    turno.AddItem "%"
    mytablex.Open "select * from turno", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        turno.AddItem "" & mytablex.Fields("turno") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    turno.ListIndex = 0

    cajero.Clear
    cajero.AddItem "%"
    mytablex.Open "select * from vendedor", cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    cajero.ListIndex = 0
 
    local1.Clear
    local1.AddItem "%"
    mytablex.Open "select * from tlocal", cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        local1.AddItem "" & mytablex.Fields("codigo") & "|" & "" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    local1.ListIndex = 0

    If local1.ListCount = 2 Then
        local1.ListIndex = 1

    End If

    bodega.Clear
    bodega.AddItem "%"
    mytablex.Open "select * from bodega", cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        bodega.AddItem "" & mytablex.Fields("codigo") & "|" & "" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    bodega.ListIndex = 0

End Sub

Private Sub Form_Load()

    Dim mytablex As New ADODB.Recordset

    Dim I        As Integer

    servicio.Clear
    servicio.AddItem "%"
    mytablex.Open "SELECT * FROM servicio ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        servicio.AddItem "" & mytablex.Fields("servicio") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    servicio.ListIndex = 0
    mytablex.Close

    comopaga.Clear
    comopaga.AddItem "N"
    comopaga.AddItem "S"
    comopaga.ListIndex = 0

    horai.Clear
    horai.AddItem "%"
    horaf.AddItem "%"

    For I = 0 To 23
        horai.AddItem Format(I, "00")
        horaf.AddItem Format(I, "00")
    Next I

    horai.ListIndex = 0
    horaf.ListIndex = 0

    Combo1.AddItem "Caja"
    Combo1.AddItem "Almacen"
    Combo1.AddItem "TipoDocumento"
    Combo1.AddItem "Codigo"
    Combo1.AddItem "Vendedor"
    Combo1.AddItem "Zona"
    Combo1.AddItem "Cajero"
    Combo1.AddItem "Turno"
    Combo1.AddItem "Local"
    Combo1.AddItem "Fecha"
    Combo1.AddItem "Hora"
    Combo1.AddItem "Servicio"
    Combo1.AddItem "Mesa"

    Combo1.ListIndex = 0

    tipores.Clear
    tipores.AddItem "Monto"
    tipores.AddItem "Cantidad"
    tipores.ListIndex = 0

    fechaf = Format(Now, "dd/mm/yyyy")
    fechai = Format(Now, "dd/mm/yyyy") '"01" & "/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")

    grupos.AddItem "N"
    grupos.AddItem "S"
    grupos.ListIndex = 0

    vdetalle.AddItem "N"
    vdetalle.AddItem "S"
    vdetalle.ListIndex = 0

    vfpago.AddItem "N"
    vfpago.AddItem "S"
    vfpago.ListIndex = 0

    estado.AddItem "%"
    estado.AddItem "2"
    estado.AddItem "1"
    estado.AddItem "0"
    estado.ListIndex = 1

    ''02/11/2017 Reporte de Seguimiento de facturas incluye delivery
    vedelivery.Clear
    vedelivery.AddItem "N"
    vedelivery.AddItem "S"
    vedelivery.ListIndex = 0
    ''02/11/2017 Reporte de Seguimiento de facturas incluye delivery

    bodega.AddItem "%"

    mytablex.Open "select * from bodega", cn, adOpenStatic, adLockOptimistic

    Do

        If mytablex.EOF Then Exit Do
        bodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("Nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    bodega.ListIndex = 0

    moneda.Clear
    moneda.AddItem "%"
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0

    '''09/10/2017 kenyo Testing Reportes
    Combo3.Clear
    Combo3.AddItem "NORMAL"
    Combo3.AddItem "EXCELL"
    Combo3.ListIndex = 0
    '''09/10/2017 kenyo Testing Reportes

End Sub

Function sql_documento(mytablex As ADODB.Recordset)

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    buf = "select * from " & cgusuario & " where "

    If Combo1.Text = "Hora" Then

        ''10/10/2017 Reporte de Seguimiento de facturas En Excel
        'buf = "select left(hora,2) as Hora1,hora,Fecha,Serie,Numero,Codigo,Nombre,Moneda,Total,Bodega,Hora,Servicio,Vendedor,Usuario,Caja,Turno,Estado,Horae from " & cgusuario & " where "
        buf = "select left(hora,2) as Hora1,local,hora,Fecha,tipo,Serie,Numero,Codigo,Nombre,Moneda,Total,Bodega,Hora,Servicio,Vendedor,Usuario,Caja,Turno,Estado,Horae,mesa,descuento,personas from " & cgusuario & " where "
        ''10/10/2017 Reporte de Seguimiento de facturas En Excel

    End If

    If tipofecha = "E" Then
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    Else
        buf = buf & "  fechae>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fechae<='" & Format(fechaf, "YYYYMMDD") & "' "

    End If

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & extra_loquesea("" & tipo) & "'"

    End If

    If serie <> "%" Then
        buf = buf & " and serie like '" & serie & "'"

    End If

    If Numero <> "%" Then
        buf = buf & " and numero like '" & Numero & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If salon <> "%" Then
        buf = buf & " and salon like '" & extra_loquesea(salon) & "'"

    End If

    If mesa <> "%" Then
        buf = buf & " and mesa like '" & mesa & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & extra_loquesea("" & caja) & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & extra_loquesea("" & turno) & "'"

    End If

    If servicio <> "%" Then
        buf = buf & " and  servicio='" & extra_loquesea(servicio) & "'"

    End If

    If opcion2 = "100" Then 'si es percecion solo los percepciones
        buf = buf & " and percepcion>0"

    End If

    If opcion2 = "4000" Then 'comiion

        'buf = buf & " and percepcion>0"
    End If

    If horai <> "%" And horaf <> "%" Then
        If Val(horaf) >= Val(horai) Then
            buf = buf & " and hour(hora)>=" & Val(horai)
            buf = buf & " and hour(hora)<=" & Val(horaf)

        End If

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & extra_loquesea("" & cajero) & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    If transporte <> "%" Then
        buf = buf & " and transporte like '" & transporte & "'"

    End If

    If fpago <> "%" Then
        buf = buf & " and fpago like '" & fpago & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    If acu <> "C" And acu <> "V" Then
        buf = buf & " and acu='" & acu & "'"

    End If

    If acu = "C" Then
        buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P')"

    End If

    If acu = "V" Then

        '19/06/2017 kenyo NOTA DE CREDITO
        'buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G')"
        buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G'  or acu='E')"
        '19/06/2017 kenyo NOTA DE CREDITO
   
    End If

    If estado <> "%" Then
        buf = buf & " and estado like '" & estado & "'"

    End If

    If Combo1 = "Hora" Then
        buf = buf & "order by left(hora,2),FECHA,str(numero)"

    End If

    If Combo1 = "TipoDocumento" Then
        buf = buf & "order by tipo,fecha,str(numero)"

    End If

    If Combo1 = "Codigo" Then
        buf = buf & "order by Codigo,FECHA,str(numero)"

    End If

    If Combo1 = "Vendedor" Then
        buf = buf & "order by vendedor,fecha,str(numero)"

    End If

    If Combo1 = "Zona" Then
        buf = buf & "order by Zona,fecha,str(numero)"

    End If

    If Combo1 = "Cajero" Then
        buf = buf & "order by Usuario,fecha,str(numero)"

    End If

    If Combo1 = "Caja" Then
        buf = buf & "order by Caja,fecha,str(numero)"

    End If

    If Combo1 = "Mesa" Then
        buf = buf & "order by Mesa,fecha,str(numero)"

    End If

    If Combo1 = "Turno" Then
        buf = buf & "order by Turno,fecha,str(numero)"

    End If

    If Combo1 = "Local" Then
        buf = buf & "order by Local,fecha,str(numero)"

    End If

    If Combo1 = "Almacen" Then
        buf = buf & "order by Local,fecha,Bodega"

    End If

    If Combo1 = "Servicio" Then
        buf = buf & "order by Servicio,fecha,str(numero)"

    End If

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    'MsgBox buf
    sql_documento = 1

End Function

Sub cabecera_documento()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    buf = String(155, "-")
    found = formateaa(buf, 155, 2, 0)

    'MsgBox opcion2
    If opcion2 = "P" Then
        found = formateaa("Local", 6, 0, 0)
        found = formateaa("Fecha", 11, 0, 0)
        found = formateaa("FechaE", 11, 0, 0)
        found = formateaa("HoraE", 11, 0, 0)
        found = formateaa("Tip", 4, 0, 0)
        found = formateaa("Ser", 4, 0, 0)
        found = formateaa("Numero", 12, 0, 0)
        found = formateaa("Codigo", 12, 0, 0)
        found = formateaa("Nombre", 41, 0, 0)
        found = formateaa("M", 2, 0, 0)
        found = formateaa("Total ", 11, 0, 1)
        found = formateaa("Acuenta ", 11, 0, 1)
        found = formateaa("Saldo ", 11, 0, 1)
        found = formateaa("E", 1, 2, 0)
        buf = String(155, "-")
        found = formateaa(buf, 155, 2, 0)
        Exit Sub

    End If
    
    If opcion2 = "100" Then
        found = formateaa("Local", 6, 0, 0)
        found = formateaa("Fecha", 11, 0, 0)
        found = formateaa("Tip", 4, 0, 0)
        found = formateaa("Ser", 4, 0, 0)
        found = formateaa("Numero", 12, 0, 0)
        found = formateaa("Codigo", 12, 0, 0)
        found = formateaa("Nombre", 41, 0, 0)
        found = formateaa("M", 2, 0, 0)
        found = formateaa("SubTotal ", 11, 0, 1)
        found = formateaa("Percepcio ", 11, 0, 1)
        found = formateaa("Total ", 11, 0, 1)
        found = formateaa("E", 2, 2, 0)
        buf = String(135, "-")
        found = formateaa(buf, 135, 2, 0)
        Exit Sub

    End If

    If opcion2 = "4000" Then
        found = formateaa("Local", 6, 0, 0)
        found = formateaa("Fecha", 11, 0, 0)
        found = formateaa("Tip", 4, 0, 0)
        found = formateaa("Ser", 4, 0, 0)
        found = formateaa("Numero", 12, 0, 0)
        found = formateaa("Codigo", 12, 0, 0)
        found = formateaa("Nombre", 41, 0, 0)
        found = formateaa("M", 2, 0, 0)
        found = formateaa("SubTotal ", 11, 0, 1)
        found = formateaa("Comision ", 11, 0, 1)
        found = formateaa("TotCom ", 11, 0, 1)
        found = formateaa("E", 2, 2, 0)
        buf = String(135, "-")
        found = formateaa(buf, 135, 2, 0)
        Exit Sub

    End If
    
    If opcion2 = "900" Then
        found = formateaa("", 65, 0, 0)
        found = formateaa("-------RECIBO DE PAGO----------", 40, 2, 0)
        found = formateaa("Lo", 3, 0, 0)
        found = formateaa("Tp", 3, 0, 0)
        found = formateaa("Srie", 5, 0, 0)
        found = formateaa("Numero", 12, 0, 0)
        found = formateaa("Fecha", 11, 0, 0)
        found = formateaa("Nombre", 31, 0, 0)
        found = formateaa("Tip Srie Numero ", 20, 0, 0)
        found = formateaa("Fecha", 11, 0, 0)
        found = formateaa("Total ", 11, 0, 1)
        found = formateaa("Abono ", 11, 0, 1)
        found = formateaa("Saldo", 11, 2, 1)
        buf = String(135, "-")
        found = formateaa(buf, 135, 2, 0)
        Exit Sub

    End If

    found = formateaa("Fecha", 11, 0, 0)
    found = formateaa("Ser", 4, 0, 0)
    found = formateaa("Numero", 12, 0, 0)
    found = formateaa("Codigo", 12, 0, 0)
    found = formateaa("Nombre", 31, 0, 0)
    found = formateaa("M", 2, 0, 0)
    found = formateaa("Total ", 11, 0, 1)
    found = formateaa("Bo", 3, 0, 0)
    found = formateaa("Hora", 11, 0, 0)
    found = formateaa("S", 2, 0, 0)

    'aqui debe aparecer a quien le pertenece
    If grupos = "S" Then
        found = formateaa("L1 ", 9, 0, 1)
        found = formateaa("L2 ", 9, 0, 1)
        found = formateaa("L3 ", 9, 0, 1)
        found = formateaa("L4 ", 9, 0, 1)

    End If

    If comopaga = "S" Then
        found = formateaa("EFECTIVO ", 9, 0, 1)
        found = formateaa("DOLARES ", 9, 0, 1)
        found = formateaa("CREDITO ", 9, 0, 1)
        found = formateaa("OTROS ", 9, 0, 1)
        GoTo arime1

    End If

    If grupos = "N" Then
        found = formateaa("Vended", 7, 0, 0)
        found = formateaa("Cajero", 7, 0, 0)
        found = formateaa("Caja", 3, 0, 0)
        found = formateaa("T", 2, 0, 0)
        found = formateaa("E", 2, 0, 0)
        found = formateaa("Msa", 4, 0, 0)
        found = formateaa("Pers", 5, 0, 0)
    
        '''19/09/2017 KENYO Descuento en Seguimiento de Facturas
        found = formateaa("Dscto", 5, 0, 0)
        '''19/09/2017 KENYO Descuento en Seguimiento de Facturas
   
    End If

arime1:
    found = formateaa("", 1, 2, 0)

    'cabecera
    If vdetalle = "S" Then
        found = formateaa("%", 1, 0, 0)
        buf = "Producto"
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "Descripcio"
        found = formateaa(buf, 27, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "Unid"
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "Fx"
        found = formateaa(buf, 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "Cant."
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "Precio"
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "Total"
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
     
        buf = "L1"
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "L2"
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "L3"
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "L4"
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 2, 0)
     
    End If
    
    buf = String(135, "-")
    found = formateaa(buf, 135, 2, 0)

End Sub

Sub cuerpo_programa_documento(mytablex As ADODB.Recordset)

    Dim Tmp   As String

    Dim sw    As Integer

    Dim buf   As String

    Dim found As Integer

    Dim sdx   As Double

    Dim sdx1  As Double

    Dim sdx2  As Double

    Dim xdx1  As Double

    Dim xdx2  As Double

    Dim xdx3  As Double

    Dim xdx4  As Double

    Dim tmp1  As String

    suma5 = 0
    suma6 = 0
    suma7 = 0
    suma8 = 0
    suma9 = 0
    sdx = 0
    sw = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    xdx1 = 0
    xdx2 = 0
    xdx3 = 0
    xdx4 = 0
    tmp1 = ""
    'MsgBox vfpago
    Do

        If mytablex.EOF Then Exit Do
        If Combo1 = "TipoDocumento" Then
            tmp1 = "" & mytablex.Fields("tipo")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("codigo")

        End If

        If Combo1 = "Hora" Then
            tmp1 = "" & mytablex.Fields("Hora1")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("usuario")

        End If

        If Combo1 = "Caja" Then
            tmp1 = "" & mytablex.Fields("Caja")

        End If

        If Combo1 = "Mesa" Then
            tmp1 = "" & mytablex.Fields("Mesa")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("Turno")

        End If

        If Combo1 = "Local" Then
            tmp1 = "" & mytablex.Fields("Local")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("Bodega")

        End If

        If Combo1 = "Servicio" Then
            tmp1 = "" & mytablex.Fields("Servicio")

        End If

        If sw = 0 Then
            If Combo1 = "TipoDocumento" Then
                buf = "" & mytablex.Fields("tipo")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_tipo("" & mytablex.Fields("tipo"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("tipo")
                sw = 1

            End If

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_codigo("" & mytablex.Fields("codigo"), "" & mytablex.Fields("tipoclie"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("codigo")
                sw = 1

            End If

            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("Hora1")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Hora1")
                sw = 1

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor("" & mytablex.Fields("vendedor"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")
                sw = 1

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_zona("" & mytablex.Fields("Zona"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("zona")
                sw = 1

            End If

            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor("" & mytablex.Fields("usuario"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("usuario")
                sw = 1

            End If

            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("caja")
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = "" & busca_caja("" & mytablex.Fields("caja"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("caja")
                sw = 1

            End If

            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("Mesa")
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = "" & mytablex.Fields("Mesa")
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("mesa")
                sw = 1

            End If

            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = "" & busca_bodega("" & mytablex.Fields("bodega"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("bodega")
                sw = 1

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("turno")
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = ""
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("turno")
                sw = 1

            End If

            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = ""
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Local")
                sw = 1

            End If

            If Combo1 = "Servicio" Then
                buf = "" & mytablex.Fields("Servicio")
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = ""
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Servicio")
                sw = 1

            End If

            suma1 = 0
            suma2 = 0
            suma3 = 0
            xdx1 = 0
            xdx2 = 0
            xdx3 = 0
            xdx4 = 0
   
        End If

        If Tmp <> tmp1 Then
            If opcion2 = "P" Then
                found = formateaa("", 114, 0, 0)
                buf = Format(suma1, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
      
                buf = Format(suma2, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
      
                buf = Format(suma3, "0.00")
                found = formateaa(buf, 10, 0, 1)
                GoTo akk1

            End If

            If opcion2 = "100" Then
                found = formateaa("", 92, 0, 0)
                buf = Format(suma1, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
      
                buf = Format(suma2, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
      
                buf = Format(suma3, "0.00")
                found = formateaa(buf, 10, 0, 1)
      
                suma1 = 0
                suma2 = 0
                suma3 = 0
                GoTo akk1
   
            End If

            If opcion2 = "4000" Then
                found = formateaa("", 92, 0, 0)
                buf = Format(suma1, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
      
                found = formateaa("", 10, 0, 1)
                found = formateaa("", 1, 0, 0)
      
                buf = Format(suma2, "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 2, 0)
      
                suma1 = 0
                suma2 = 0
                suma3 = 0
                GoTo akk1
   
            End If

            found = formateaa("", 39, 0, 0)
            found = formateaa(dicmoneda, 7, 0, 0)
            buf = Format(suma1, "0.00")
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("Dolares", 8, 0, 0)
            buf = Format(suma2, "0.00")
            found = formateaa(buf, 10, 0, 0)

            If grupos = "S" Then
                found = formateaa("", 22, 0, 0)
                buf = Format(xdx1, "0.00")
                found = formateaa(buf, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = Format(xdx2, "0.00")
                found = formateaa(buf, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = Format(xdx3, "0.00")
                found = formateaa(buf, 8, 0, 1)
                found = formateaa("", 1, 0, 0)
                buf = Format(xdx4, "0.00")
                found = formateaa(buf, 8, 0, 1)
                found = formateaa("", 1, 0, 0)

            End If

akk1:
            found = formateaa("", 1, 2, 0)
            nlineas

            If Combo1 = "TipoDocumento" Then
                buf = "" & mytablex.Fields("tipo")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_tipo("" & mytablex.Fields("tipo"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("tipo")

            End If

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_codigo("" & mytablex.Fields("codigo"), "" & mytablex.Fields("tipoclie"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("Hora1")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Hora1")

            End If
   
            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("vendedor")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor("" & mytablex.Fields("vendedor"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 3, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_zona("" & mytablex.Fields("Zona"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("zona")

            End If

            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("usuario")
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor("" & mytablex.Fields("usuario"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("usuario")

            End If

            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("caja")
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_caja("" & mytablex.Fields("caja"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("caja")

            End If

            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("Mesa")
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = "" & mytablex.Fields("caja")
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("mesa")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = ""
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Turno")

            End If

            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = "" & busca_bodega("" & mytablex.Fields("bodega"))
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = ""
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("Local")
                sw = 1

            End If

            If Combo1 = "Servicio" Then
                buf = "" & mytablex.Fields("servicio")
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = ""
                found = formateaa(buf, 30, 0, 0)
                found = formateaa("", 1, 2, 0)
                nlineas
                Tmp = "" & mytablex.Fields("servicio")
                sw = 1

            End If

            suma1 = 0
            suma2 = 0
            suma3 = 0

        End If

        If opcion2 = "900" Then
            buf = "" & mytablex.Fields("LOCAL")
            found = formateaa(buf, 2, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("tipo")
            found = formateaa(buf, 2, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("serie")
            found = formateaa(buf, 4, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("numero")
            found = formateaa(buf, 11, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("fecha")
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("nombre")
            found = formateaa(buf, 30, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("", 23, 0, 0)
            found = formateaa("", 8, 0, 0)
            buf = Format(Val("" & mytablex.Fields("total")), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            suma1 = 0
            imprime_cuentag mytablex
            imprime_cuentacd mytablex
            GoTo p900

        End If

        If opcion2 = "100" Then  'PERCEPCION
            buf = "" & mytablex.Fields("LOCAL")
            found = formateaa(buf, 5, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("Fecha")
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("tipo")
            found = formateaa(buf, 2, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("serie")
            found = formateaa(buf, 4, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("numero")
            found = formateaa(buf, 11, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("codigo")
            found = formateaa(buf, 11, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("nombre")
            found = formateaa(buf, 40, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("moneda")
            found = formateaa(buf, 1, 0, 0)
            found = formateaa("", 1, 0, 0)
      
            sumar_percepcion mytablex, sdx, sdx1, sdx2
      
            buf = Format(sdx, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(sdx1, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(sdx2, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("ESTADO")
            found = formateaa(buf, 1, 0, 0)
            found = formateaa("", 1, 2, 0)
      
            If "" & mytablex.Fields("moneda") = "S" And "" & mytablex.Fields("ESTADO") = "2" Then
                suma1 = suma1 + sdx
                suma2 = suma2 + sdx1
                suma3 = suma3 + sdx2
                ssuma1 = ssuma1 + sdx
                ssuma2 = ssuma2 + sdx1
                ssuma3 = ssuma3 + sdx2

            End If

            If "" & mytablex.Fields("moneda") = "D" And "" & mytablex.Fields("ESTADO") = "2" Then
                suma1 = suma1 + sdx * 2.8
                suma2 = suma2 + sdx1 * 2.8
                suma3 = suma3 + sdx2 * 2.8
                ssuma1 = ssuma1 + sdx * 2.8
                ssuma2 = ssuma2 + sdx1 * 2.8
                ssuma3 = ssuma3 + sdx2 * 2.8

            End If

            GoTo p900

        End If

        If opcion2 = "4000" Then  'COMISION
            buf = "" & mytablex.Fields("LOCAL")
            found = formateaa(buf, 5, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("Fecha")
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("tipo")
            found = formateaa(buf, 2, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("serie")
            found = formateaa(buf, 4, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("numero")
            found = formateaa(buf, 11, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("codigo")
            found = formateaa(buf, 11, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("nombre")
            found = formateaa(buf, 40, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("moneda")
            found = formateaa(buf, 1, 0, 0)
            found = formateaa("", 1, 0, 0)
      
            buf = "" & mytablex.Fields("subtotal")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
      
            buf = "" & busca_vendedorco(mytablex)
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            sdx = Val(buf) * Val("" & mytablex.Fields("subtotal")) / 100
            buf = Format(sdx, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
      
            buf = "" & mytablex.Fields("ESTADO")
            found = formateaa(buf, 1, 0, 0)
            found = formateaa("", 1, 2, 0)
      
            If "" & mytablex.Fields("moneda") = "S" And "" & mytablex.Fields("ESTADO") = "2" Then
                suma1 = suma1 + Val("" & mytablex.Fields("subtotal"))
                suma2 = suma2 + sdx
         
                ssuma1 = ssuma1 + Val("" & mytablex.Fields("subtotal"))
                ssuma2 = ssuma2 + sdx
         
            End If

            If "" & mytablex.Fields("moneda") = "D" And "" & mytablex.Fields("ESTADO") = "2" Then
                suma1 = suma1 + Val("" & mytablex.Fields("subtotal")) * 2.8
                suma2 = suma2 + sdx * 2.8
                ssuma1 = ssuma1 + Val("" & mytablex.Fields("subtotal")) * 2.8
                ssuma2 = ssuma2 + sdx * 2.8

            End If

            GoTo p900

        End If

        If opcion2 = "P" Then  'COMENTARIOS
            buf = "" & mytablex.Fields("LOCAL")
            found = formateaa(buf, 5, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("Fecha")
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("Fechae")
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("Horae")
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("tipo")
            found = formateaa(buf, 2, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("serie")
            found = formateaa(buf, 4, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("numero")
            found = formateaa(buf, 11, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("codigo")
            found = formateaa(buf, 11, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("nombre")
            found = formateaa(buf, 40, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("moneda")
            found = formateaa(buf, 1, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = Format(Val("" & mytablex.Fields("TOTAL")), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(Val("" & mytablex.Fields("acuenta")), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            sdx = Val("" & mytablex.Fields("total")) - Val("" & mytablex.Fields("acuenta"))
            buf = Format(sdx, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("ESTADO")
            found = formateaa(buf, 1, 0, 0)
            found = formateaa("", 1, 2, 0)

            If "" & mytablex.Fields("moneda") = "S" And "" & mytablex.Fields("ESTADO") = "2" Then
                suma1 = suma1 + Val("" & mytablex.Fields("total"))
                suma2 = suma2 + Val("" & mytablex.Fields("acuenta"))
                suma3 = suma3 + (Val("" & mytablex.Fields("total")) - Val("" & mytablex.Fields("acuenta")))
                ssuma1 = ssuma1 + Val("" & mytablex.Fields("total"))
                ssuma2 = ssuma2 + Val("" & mytablex.Fields("acuenta"))
                ssuma3 = ssuma3 + (Val("" & mytablex.Fields("total")) - Val("" & mytablex.Fields("acuenta")))

            End If

            If "" & mytablex.Fields("moneda") = "D" And "" & mytablex.Fields("ESTADO") = "2" Then
                suma1 = suma1 + Val("" & mytablex.Fields("total"))
                suma2 = suma2 + Val("" & mytablex.Fields("acuenta"))
                suma3 = suma3 + (Val("" & mytablex.Fields("total")) - Val("" & mytablex.Fields("acuenta")))
                ssuma1 = ssuma1 + Val("" & mytablex.Fields("total"))
                ssuma2 = ssuma2 + Val("" & mytablex.Fields("acuenta"))
                ssuma3 = ssuma3 + (Val("" & mytablex.Fields("total")) - Val("" & mytablex.Fields("acuenta")))

            End If

            If vdetalle = "S" Then
                ver_detalle mytablex

            End If

            If vfpago = "S" Then
                ver_fpagov mytablex

            End If

            GoTo p900

        End If

        buf = "" & mytablex.Fields("fecha")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("serie")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("numero")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("codigo")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("nombre")
        found = formateaa(buf, 30, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("moneda")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("total")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("bodega")
        found = formateaa(buf, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("hora")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("servicio")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        buf = "" & mytablex.Fields("mesa")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        If comopaga = "S" Then
            sumar_como_paga mytablex
            GoTo arime

        End If

        If grupos = "S" Then
            buf = Format(Val("" & mytablex.Fields("C1")), "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa(buf, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            buf = Format(Val("" & mytablex.Fields("C2")), "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa(buf, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            buf = Format(Val("" & mytablex.Fields("C3")), "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa(buf, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            buf = Format(Val("" & mytablex.Fields("C4")), "0.00")

            If Val(buf) = 0 Then
                buf = ""

            End If

            found = formateaa(buf, 8, 0, 1)
            found = formateaa("", 1, 0, 0)
      
            'xdx1 = xdx1 + Val(Format(Val("" & mytablex.Fields("C1")), "0.00"))
            'xdx2 = xdx2 + Val(Format(Val("" & mytablex.Fields("C2")), "0.00"))
            'xdx3 = xdx3 + Val(Format(Val("" & mytablex.Fields("C3")), "0.00"))
            'xdx4 = xdx4 + Val(Format(Val("" & mytablex.Fields("C4")), "0.00"))
        End If

        If grupos = "N" Then
            buf = "" & mytablex.Fields("vendedor")
            found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("usuario")
            found = formateaa(buf, 6, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("caja")
            found = formateaa(buf, 2, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("Turno")
            found = formateaa(buf, 1, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("estado")
            found = formateaa(buf, 1, 0, 0)
            found = formateaa("", 1, 0, 0)
   
            buf = "" & mytablex.Fields("mesa")
            found = formateaa(buf, 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("personas")
            found = formateaa(buf, 2, 0, 0)
            found = formateaa("", 3, 0, 0)
            '''19/09/2017 KENYO Descuento en Seguimiento de Facturas
            buf = "" & mytablex.Fields("descuento")
            found = formateaa(buf, 7, 0, 0)
            found = formateaa("", 1, 0, 0)
            '''19/09/2017 KENYO Descuento en Seguimiento de Facturas

        End If

arime:
        found = formateaa("", 1, 2, 0)
        nlineas

        If "" & mytablex.Fields("moneda") = "S" Then
            If tipores = "Monto" Then
                suma1 = suma1 + Val("" & mytablex.Fields("total"))
                ssuma1 = ssuma1 + Val("" & mytablex.Fields("total"))

            End If

            If tipores = "Cantidad" Then
                suma1 = suma1 + 1
                ssuma1 = ssuma1 + 1

            End If

        End If

        If "" & mytablex.Fields("moneda") = "D" Then
            If tipores = "Monto" Then
                suma2 = suma2 + Val("" & mytablex.Fields("total"))
                ssuma2 = ssuma2 + Val("" & mytablex.Fields("total"))

            End If

            If tipores = "Cantidad" Then
                suma2 = suma2 + 1
                ssuma2 = ssuma2 + 1

            End If

        End If

        If vdetalle = "S" Then
            ver_detalle mytablex

        End If

        If vfpago = "S" Then
            ver_fpagov mytablex

        End If

p900:
        mytablex.MoveNext
    Loop

    If opcion2 = "100" Then
        found = formateaa("", 92, 0, 0)
        buf = Format(suma1, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
      
        buf = Format(suma2, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
      
        buf = Format(suma3, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
      
        found = formateaa("", 92, 0, 0)
        buf = Format(ssuma1, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
      
        buf = Format(ssuma2, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
      
        buf = Format(ssuma3, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
      
        Exit Sub
   
    End If

    If opcion2 = "4000" Then
        found = formateaa("", 92, 0, 0)
        buf = Format(suma1, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
      
        found = formateaa("", 10, 0, 1)
        found = formateaa("", 1, 0, 0)
    
        buf = Format(suma2, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
      
        found = formateaa("", 92, 0, 0)
        buf = Format(ssuma1, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
      
        found = formateaa("", 10, 0, 1)
        found = formateaa("", 1, 0, 0)
    
        buf = Format(ssuma2, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
        Exit Sub
   
    End If

    If opcion2 = "P" Then
        found = formateaa("", 114, 0, 0)
        buf = Format(suma1, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
      
        buf = Format(suma2, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
      
        buf = Format(suma3, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
      
        found = formateaa("", 114, 0, 0)
        buf = Format(ssuma1, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
      
        buf = Format(ssuma2, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
      
        buf = Format(ssuma3, "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas
        Exit Sub

    End If

    If opcion2 = "900" Then
        GoTo otro900
   
    End If

    found = formateaa("", 39, 0, 0)
    found = formateaa(dicmoneda, 7, 0, 0)
    buf = Format(suma1, "0.00")
    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Dolares:", 8, 0, 0)
    buf = Format(suma2, "0.00")
    found = formateaa(buf, 10, 0, 0)
   
    If comopaga = "S" Then
        found = formateaa("", 24, 0, 0)
        buf = Format(suma6, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma7, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma8, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(suma9, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 2, 0)
        GoTo amida2

    End If

    If grupos = "S" Then
        found = formateaa("", 22, 0, 0)
        buf = Format(xdx1, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(xdx2, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(xdx3, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = Format(xdx4, "0.00")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)

    End If

amida2:
    found = formateaa("", 1, 2, 0)
    nlineas
   
    found = formateaa("Gran Total ", 39, 0, 1)
    found = formateaa(dicmoneda, 7, 0, 0)
    buf = Format(ssuma1, "0.00")
    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Dolares:", 8, 0, 0)
    buf = Format(ssuma2, "0.00")
    found = formateaa(buf, 10, 0, 0)
    found = formateaa("", 1, 2, 0)
otro900:
   
End Sub

Sub ver_detalle(mytabley As ADODB.Recordset)

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    Dim found    As Integer

    'MsgBox dgusuariog
    mytablex.Open "select * from " & dgusuariog & " where local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "' and numero='" & "" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    nlineas
    sw = 0
    Do

        If mytablex.EOF Then Exit Do
        'If "" & mytablex.Fields("codigo") = "" & mytabley.Fields("codigo") And "" & mytablex.Fields("acu") = "" & mytabley.Fields("acu") Then
        sw = 1
        found = formateaa("%", 1, 0, 0)
        buf = "" & mytablex.Fields("producto")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("descripcio")
        found = formateaa(buf, 27, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("unidad")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("factor")
        found = formateaa(buf, 4, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("cantidad")
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("precio")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("total")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        nlineas

        If Len(Trim("" & mytablex.Fields("observa1"))) > 0 Then
            buf = "" & mytablex.Fields("observa1")
            found = formateaa(buf, 50, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If Len(Trim("" & mytablex.Fields("observa2"))) > 0 Then
            buf = "" & mytablex.Fields("observa2")
            found = formateaa(buf, 50, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If

        If Len(Trim("" & mytablex.Fields("observa3"))) > 0 Then
            buf = "" & mytablex.Fields("observa3")
            found = formateaa(buf, 50, 0, 0)
            found = formateaa("", 1, 2, 0)
            nlineas

        End If
   
        mytablex.MoveNext
    Loop

    If sw = 1 Then
        buf = String(130, "-")
        found = formateaa(buf, 130, 2, 0)
        nlineas

    End If

    mytablex.Close

End Sub

Sub ver_fpagov(mytabley As ADODB.Recordset)

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    Dim found    As Integer

    On Error GoTo cmd6712_err

    If Len(gofpago) = 0 Then Exit Sub

    mytablex.Open "select * from " & gofpago & " where local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "' and numero='" & "" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then

        mytablex.Close
        Exit Sub

    End If

    buf = String(130, "-")
    found = formateaa(buf, 130, 2, 0)
    nlineas
    sw = 0
    Do

        If mytablex.EOF Then Exit Do
        'If "" & mytablex.Fields("codigo") = "" & mytabley.Fields("codigo") And "" & mytablex.Fields("acu") = "" & mytabley.Fields("acu") Then
        sw = 1
        found = formateaa(">", 1, 0, 0)
        buf = "" & mytablex.Fields("numero")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
     
        buf = "" & mytablex.Fields("fpago")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
     
        buf = "" & mytablex.Fields("descripcio")
        found = formateaa(buf, 15, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("recibe")
        buf = Format(Val(buf), "0.00")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas
        mytablex.MoveNext
    Loop

    If sw = 1 Then
        buf = String(130, "-")
        found = formateaa(buf, 130, 2, 0)
        nlineas

    End If

    mytablex.Close
    Exit Sub
cmd6712_err:
    MsgBox "Error en " & error$, 48, "Aviso"
    mytablex.Close
    Exit Sub

End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > Val(nrolineas) Then
        If opcion2 = "13" Then
            cabecera_documento_semanas1

        End If

        If opcion2 = "10" Or opcion2 = "11" Then
            cabecera_documento

        End If

        If opcion2 = "12" Then
            cabecera_documento_meses1

        End If

    End If

End Sub

Function busca_tipo(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from tipo where tipo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_tipo = "" & mytablex.Fields("descripcio")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_servicio(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from servicio where servicio='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_servicio = "" & mytablex.Fields("descripcio")
    Else
        busca_servicio = " "

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_cliente(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from clientes where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_cliente = "" & mytablex.Fields("nombre")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_vendedor(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_vendedor = "" & mytablex.Fields("nombre")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_zona(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from zona where zona='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        busca_zona = "" & mytablex.Fields("descripcio")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_caja(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from parameca where caja='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        busca_caja = "" & mytablex.Fields("descripcio")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_bodega(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from bodega where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        busca_bodega = "" & mytablex.Fields("nombre")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_localx(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from tlocal where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_localx = "" & mytablex.Fields("nombre")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

''02/11/2017 Reporte de Seguimiento de facturas incluye delivery
Function busca_datosdelivery(buf As String, sw As Integer) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from DELIVERI where codigo='" & "" & buf & "'  ", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If sw = 0 Then
            busca_datosdelivery = "" & mytablex.Fields("telefono")
        ElseIf sw = 1 Then
            busca_datosdelivery = "" & mytablex.Fields("direccion")
        ElseIf sw = 2 Then
            busca_datosdelivery = "" & mytablex.Fields("referencia")

        End If

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

''02/11/2017 Reporte de Seguimiento de facturas incluye delivery

Function busca_codigo(buf As String, sw As String) As String

    Dim mytablex As New ADODB.Recordset

    Dim buf1     As String

    buf1 = "vendedor"

    If sw = "C" Then
        buf1 = "clientes"

    End If

    If sw = "P" Then
        buf1 = "proveedo"

    End If

    mytablex.Open "select * from " & buf1 & " where codigo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_codigo = "" & mytablex.Fields("nombre")

    End If

    '------------------------------------- ------------
    mytablex.Close
 
End Function

Function valida_flag(buf As String)

    Select Case buf

        Case "T", "A", "B", "C", "D", "G", "E", "F"
            valida_flag = 1

        Case "S", "J", "K", "L", "M", "P", "N", "O"
            valida_flag = 2

    End Select
       
End Function

Sub proceso_venta_diario()  'ventas diarias

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    contpag = 0

    found = sql_documento_meses(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento_meses
    cuerpo_programa_documento_meses mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Function sql_documento_meses(mytablex As ADODB.Recordset)

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function

    If Combo1 = "Fecha" Then
        buf = "select Fecha,Tipo,Moneda as M,sum(subtotal) as subTot,sum(impuesto) as impto,sum(gravado) as grava,sum(total) as Tot,sum(tivap) as xivap, count(tipo) as Nrod,sum(tisc) as xisc from " & cgusuario & " where "

    End If

    If Combo1 = "Hora" Then
        buf = "select left(hora,2) AS hora,Tipo,Moneda as M,sum(subtotal) as subTot,sum(impuesto) as impto,sum(gravado) as grava,sum(total) as Tot,sum(tivap) as xivap, count(tipo) as Nrod,sum(tisc) as xisc from " & cgusuario & " where "

    End If

    If Combo1 = "TipoDocumento" Then
        buf = "select Tipo,Tipo,Moneda as M,sum(subtotal) as subTot,sum(impuesto) as impto,sum(gravado) as grava,sum(total) as Tot,sum(tivap) as xivap,count(tipo) as Nrod,sum(tisc) as xisc from " & cgusuario & " where "

    End If

    If Combo1 = "Codigo" Then
        buf = "select Codigo,Tipo,Moneda as M,sum(subtotal) as subTot,sum(impuesto) as impto,sum(gravado) as grava,sum(total) as Tot,sum(tivap) as xivap,count(tipo) as Nrod,sum(tisc) as xisc from " & cgusuario & " where "

    End If

    If Combo1 = "Vendedor" Then
        buf = "select Vendedor,Tipo,Moneda as M,sum(subtotal) as subTot,sum(impuesto) as impto,sum(gravado) as grava,sum(total) as Tot,sum(tivap) as xivap,count(tipo) as Nrod,sum(tisc) as xisc from " & cgusuario & " where "

    End If

    If Combo1 = "Zona" Then
        buf = "select zona,Tipo,Moneda as M,sum(subtotal) as subTot,sum(impuesto) as impto,sum(gravado) as grava,sum(total) as Tot,sum(tivap) as xivap,count(tipo) as Nrod,sum(tisc) as xisc from " & cgusuario & " where "

    End If

    If Combo1 = "Cajero" Then
        buf = "select Usuario,Tipo,Moneda as M,sum(subtotal) as subTot,sum(impuesto) as impto,sum(gravado) as grava,sum(total) as Tot,sum(tivap) as xivap,count(tipo) as Nrod,sum(tisc) as xisc from " & cgusuario & " where "

    End If

    If Combo1 = "Caja" Then
        buf = "select Caja,Tipo,Moneda as M,sum(subtotal) as subTot,sum(impuesto) as impto,sum(gravado) as grava,sum(total) as Tot,sum(tivap) as xivap,count(tipo) as Nrod,sum(tisc) as xisc from " & cgusuario & " where "

    End If

    If Combo1 = "Mesa" Then
        buf = "select Mesa,Tipo,Moneda as M,sum(subtotal) as subTot,sum(impuesto) as impto,sum(gravado) as grava,sum(total) as Tot,sum(tivap) as xivap,count(tipo) as Nrod,sum(tisc) as xisc from " & cgusuario & " where "

    End If

    If Combo1 = "Turno" Then
        buf = "select Turno,Tipo,Moneda as M,sum(subtotal) as subTot,sum(impuesto) as impto,sum(gravado) as grava,sum(total) as Tot,sum(tivap) as xivap,count(tipo) as Nrod,sum(tisc) as xisc from " & cgusuario & " where "

    End If

    If Combo1 = "Local" Then
        buf = "select Local,Tipo,Moneda as M,sum(subtotal) as subTot,sum(impuesto) as impto,sum(gravado) as grava,sum(total) as Tot,sum(tivap) as xivap,count(tipo) as Nrod,sum(tisc) as xisc from " & cgusuario & " where "

    End If

    If Combo1 = "Servicio" Then
        buf = "select Servicio,Tipo,Moneda as M,sum(subtotal) as subTot,sum(impuesto) as impto,sum(gravado) as grava,sum(total) as Tot,sum(tivap) as xivap,count(tipo) as Nrod,sum(tisc) as xisc from " & cgusuario & " where "

    End If

    If Combo1 = "Almacen" Then
        buf = "select Bodega,Tipo,Moneda as M,sum(subtotal) as subTot,sum(impuesto) as impto,sum(gravado) as grava,sum(total) as Tot,sum(tivap) as xivap,count(tipo) as Nrod,sum(tisc) as xisc from " & cgusuario & " where "

    End If

    If tipofecha = "E" Then
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    Else
        buf = buf & "  fechae>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fechae<='" & Format(fechaf, "YYYYMMDD") & "' "

    End If

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & extra_loquesea("" & tipo) & "'"

    End If

    If serie <> "%" Then
        buf = buf & " and serie like '" & serie & "'"

    End If

    If Numero <> "%" Then
        buf = buf & " and numero like '" & Numero & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If salon <> "%" Then
        buf = buf & " and salon like '" & extra_loquesea(salon) & "'"

    End If

    If mesa <> "%" Then
        buf = buf & " and mesa like '" & mesa & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & extra_loquesea(caja) & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "'"

    End If

    If horai <> "%" And horaf <> "%" Then
        If Val(horaf) >= Val(horai) Then
            buf = buf & " and hour(hora)>=" & Val(horai)
            buf = buf & " and hour(hora)<=" & Val(horaf)

        End If

    End If

    If servicio <> "%" Then
        'If servicio = "Deliveri" Then
        buf = buf & " and  servicio='" & extra_loquesea(servicio) & "'"

        'End If
        'If servicio = "Comanda" Then
        '   buf = buf & " and  servicio='C' "
        'End If
        'If servicio = "Autoservicio" Then
        '   buf = buf & " and  servicio='*' "
        'End If
    End If

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    If transporte <> "%" Then
        buf = buf & " and transporte like '" & transporte & "'"

    End If

    If fpago <> "%" Then
        buf = buf & " and fpago like '" & fpago & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    If acu <> "C" And acu <> "V" Then
        buf = buf & " and acu='" & acu & "'"

    End If

    If acu = "C" Then
        buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P')"

    End If

    If acu = "V" Then
        '19/06/2017 kenyo NOTA DE CREDITO
        ' buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G')"
        buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E')"

        '19/06/2017 kenyo NOTA DE CREDITO
    End If

    If estado <> "%" Then
        buf = buf & " and estado like '" & estado & "'"

    End If

    If Combo1 = "Fecha" Then
        buf = buf & " group by fecha,tipo,moneda  order by fecha"

    End If

    If Combo1 = "Hora" Then
        buf = buf & " group by left(hora,2),tipo,moneda  order by hora"

    End If

    If Combo1 = "TipoDocumento" Then
        buf = buf & " group by Tipo,tipo,moneda order by tipo"

    End If

    If Combo1 = "Codigo" Then
        buf = buf & " group by Codigo,tipo,moneda order by codigo"

    End If

    If Combo1 = "Vendedor" Then
        buf = buf & " group by Vendedor,tipo,moneda order by vendedor"

    End If

    If Combo1 = "Zona" Then
        buf = buf & " group by Zona,tipo,moneda order by zona"

    End If

    If Combo1 = "Cajero" Then
        buf = buf & " group by Usuario,tipo,moneda order by usuario"

    End If

    If Combo1 = "Caja" Then
        buf = buf & " group by Caja,tipo,moneda order by caja"

    End If

    If Combo1 = "Mesa" Then
        buf = buf & " group by Mesa,tipo,moneda order by mesa"

    End If

    If Combo1 = "Turno" Then
        buf = buf & " group by Turno,tipo,moneda order by turno"

    End If

    If Combo1 = "Local" Then
        buf = buf & " group by Local,tipo,moneda order by local"

    End If

    If Combo1 = "Servicio" Then
        buf = buf & " group by Servicio,tipo,moneda order by Servicio"

    End If

    If Combo1 = "Almacen" Then
        buf = buf & " group by Bodega,Tipo,moneda order by bodega"

    End If

    'MsgBox buf

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    sql_documento_meses = 1

End Function

Sub cabecera_documento_meses()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    buf = String(155, "-")
    found = formateaa(buf, 155, 2, 0)
    found = formateaa("Tipo", 3, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Descripcio", 10, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("NroDo", 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("M", 1, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("SubTotal", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Impuesto", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Exonerad", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    
    found = formateaa("Isc", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    
    found = formateaa("Ivap", 10, 0, 1)
    found = formateaa("", 1, 0, 0)

    found = formateaa("Total", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    
    found = formateaa("VtaAcum", 10, 0, 1)
    found = formateaa("", 1, 2, 0)
        
    buf = String(155, "-")
    found = formateaa(buf, 155, 2, 0)

End Sub

Sub cuerpo_programa_documento_meses(mytablex As ADODB.Recordset)

    Dim Tmp   As String

    Dim sw    As Integer

    Dim buf   As String

    Dim sdx1  As Double

    Dim found As Integer

    Dim tmp1  As String

    sdx1 = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0
    suma7 = 0
    suma8 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0
    ssuma7 = 0
    ssuma8 = 0

    Do

        If mytablex.EOF Then Exit Do
        If Combo1 = "Fecha" Then
            tmp1 = "" & mytablex.Fields("fecha")

        End If

        If Combo1 = "Hora" Then
            tmp1 = "" & mytablex.Fields("hora")

        End If

        If Combo1 = "TipoDocumento" Then
            tmp1 = "" & mytablex.Fields("Tipo")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("bodega")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("Codigo")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("Vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("usuario")

        End If

        If Combo1 = "Caja" Then
            tmp1 = "" & mytablex.Fields("Caja")

        End If

        If Combo1 = "Mesa" Then
            tmp1 = "" & mytablex.Fields("mesa")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("Turno")

        End If

        If Combo1 = "Local" Then
            tmp1 = "" & mytablex.Fields("Local")

        End If

        If Combo1 = "Servicio" Then
            tmp1 = "" & mytablex.Fields("Servicio")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("bodega")

        End If
   
        If sw = 0 Then
            If Combo1 = "Fecha" Then
                buf = "" & mytablex.Fields("Fecha")
                found = formateaa(buf, 10, 2, 0)
                Tmp = "" & mytablex.Fields("Fecha")

            End If

            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("hora")
                found = formateaa(buf, 10, 2, 0)
                Tmp = "" & mytablex.Fields("hora")

            End If

            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                found = formateaa(buf, 10, 0, 0)
                buf = busca_bodega("" & mytablex.Fields("bodega"))
                found = formateaa(buf, 30, 2, 0)
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "TipoDocumento" Then
                buf = "" & mytablex.Fields("Tipo")
                found = formateaa(buf, 10, 0, 0)
                buf = busca_tipo("" & mytablex.Fields("tipo"))
                found = formateaa(buf, 30, 2, 0)
                Tmp = "" & mytablex.Fields("Tipo")

            End If

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_cliente("" & mytablex.Fields("codigo"))
                found = formateaa(buf, 30, 2, 0)
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("Vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor("" & mytablex.Fields("Vendedor"))
                found = formateaa(buf, 30, 2, 0)
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 10, 2, 0)
                Tmp = "" & mytablex.Fields("Zona")

            End If

            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("Usuario")
                found = formateaa(buf, 10, 0, 0)
                buf = busca_vendedor("" & mytablex.Fields("usuario"))
                found = formateaa(buf, 30, 2, 0)
                Tmp = "" & mytablex.Fields("Usuario")

            End If

            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                found = formateaa(buf, 10, 0, 0)
                buf = busca_caja("" & mytablex.Fields("caja"))
                found = formateaa(buf, 30, 2, 0)
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("Mesa")
                found = formateaa(buf, 10, 0, 0)
                buf = "" & mytablex.Fields("mesa")
                found = formateaa(buf, 30, 2, 0)
                Tmp = "" & mytablex.Fields("mesa")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 10, 2, 0)
                Tmp = "" & mytablex.Fields("Turno")

            End If

            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                found = formateaa(buf, 10, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_localx("" & mytablex.Fields("Local"))
                found = formateaa(buf, 10, 2, 0)
                Tmp = "" & mytablex.Fields("Local")

            End If

            If Combo1 = "Servicio" Then
                buf = "" & mytablex.Fields("Servicio")
                found = formateaa(buf, 10, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_servicio("" & mytablex.Fields("servicio"))
                found = formateaa(buf, 10, 2, 0)
                Tmp = "" & mytablex.Fields("Servicio")

            End If

            nlineas
            sw = 1

        End If

        If Tmp <> tmp1 Then
            found = formateaa("", 13, 0, 0)
            buf = Format(suma8, "0.00")
            found = formateaa(buf, 10, 0, 1)
            buf = Format(suma1, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(suma2, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(suma3, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(suma4, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(suma7, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(suma5, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(ssuma6, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 2, 0)
            nlineas
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0

            If Combo1 = "Fecha" Then
                buf = "" & mytablex.Fields("Fecha")
                found = formateaa(buf, 10, 2, 0)
                Tmp = "" & mytablex.Fields("Fecha")

            End If

            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("Hora")
                found = formateaa(buf, 10, 2, 0)
                Tmp = "" & mytablex.Fields("Hora")

            End If

            If Combo1 = "TipoDocumento" Then
                buf = "" & mytablex.Fields("Tipo")
                found = formateaa(buf, 10, 0, 0)
                buf = busca_tipo("" & mytablex.Fields("tipo"))
                found = formateaa(buf, 30, 2, 0)
                Tmp = "" & mytablex.Fields("Tipo")

            End If

            If Combo1 = "Almacen" Then
                buf = "" & mytablex.Fields("bodega")
                found = formateaa(buf, 10, 0, 0)
                buf = busca_bodega("" & mytablex.Fields("bodega"))
                found = formateaa(buf, 30, 2, 0)
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Codigo" Then
                buf = "" & mytablex.Fields("codigo")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_cliente("" & mytablex.Fields("codigo"))
                found = formateaa(buf, 30, 2, 0)
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Vendedor" Then
                buf = "" & mytablex.Fields("Vendedor")
                found = formateaa(buf, 11, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_vendedor("" & mytablex.Fields("Vendedor"))
                found = formateaa(buf, 30, 2, 0)
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 10, 2, 0)
                Tmp = "" & mytablex.Fields("Zona")

            End If

            If Combo1 = "Cajero" Then
                buf = "" & mytablex.Fields("Usuario")
                found = formateaa(buf, 10, 0, 0)
                buf = busca_vendedor("" & mytablex.Fields("usuario"))
                found = formateaa(buf, 30, 2, 0)
                Tmp = "" & mytablex.Fields("Usuario")

            End If

            If Combo1 = "Caja" Then
                buf = "" & mytablex.Fields("Caja")
                found = formateaa(buf, 10, 0, 0)
                buf = busca_caja("" & mytablex.Fields("caja"))
                found = formateaa(buf, 30, 2, 0)
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("Mesa")
                found = formateaa(buf, 10, 0, 0)
                buf = "" & mytablex.Fields("Mesa")
                found = formateaa(buf, 30, 2, 0)
                Tmp = "" & mytablex.Fields("mesa")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 10, 2, 0)
                Tmp = "" & mytablex.Fields("Turno")

            End If

            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                found = formateaa(buf, 10, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_localx("" & mytablex.Fields("Local"))
                found = formateaa(buf, 10, 2, 0)
                Tmp = "" & mytablex.Fields("Local")

            End If

            If Combo1 = "Servicio" Then
                buf = "" & mytablex.Fields("Servicio")
                found = formateaa(buf, 10, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = busca_servicio("" & mytablex.Fields("servicio"))
                found = formateaa(buf, 10, 2, 0)
                Tmp = "" & mytablex.Fields("servicio")

            End If
   
            nlineas

        End If
   
        buf = "" & mytablex.Fields("Tipo")
        found = formateaa(buf, 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = busca_tipo("" & mytablex.Fields("Tipo"))
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        suma8 = suma8 + Val("" & mytablex.Fields("nrod"))
        ssuma8 = ssuma8 + Val("" & mytablex.Fields("nrod"))
        buf = "" & mytablex.Fields("nrod")
        found = formateaa(buf, 5, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("M")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        sdx1 = Val("" & mytablex.Fields("tot")) - Val("" & mytablex.Fields("grava")) - Val("" & mytablex.Fields("impto"))
        'buf = Format(Val("" & mytablex.Fields("subTot")), "0.00")
        buf = Format(sdx1, "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        suma1 = suma1 + Val("" & mytablex.Fields("subtot"))
        ssuma1 = ssuma1 + Val("" & mytablex.Fields("subtot"))
   
        buf = Format(Val("" & mytablex.Fields("Impto")), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        suma2 = suma2 + Val("" & mytablex.Fields("impto"))
        ssuma2 = ssuma2 + Val("" & mytablex.Fields("impto"))
   
        buf = Format(Val("" & mytablex.Fields("grava")), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        suma3 = suma3 + Val("" & mytablex.Fields("grava"))
        ssuma3 = ssuma3 + Val("" & mytablex.Fields("grava"))
   
        buf = Format(Val("" & mytablex.Fields("xisc")), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        suma4 = suma4 + Val("" & mytablex.Fields("xisc"))
        ssuma4 = ssuma4 + Val("" & mytablex.Fields("xisc"))
     
        buf = Format(Val("" & mytablex.Fields("xivap")), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        suma7 = suma7 + Val("" & mytablex.Fields("xivap"))
        ssuma7 = ssuma7 + Val("" & mytablex.Fields("xivap"))
   
        buf = Format(Val("" & mytablex.Fields("tot")), "0.00")

        If Val(buf) = 0 Then
            buf = ""

        End If

        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 2, 0)
        suma5 = suma5 + Val("" & mytablex.Fields("tot"))
        ssuma5 = ssuma5 + Val("" & mytablex.Fields("tot"))
        ssuma6 = ssuma6 + Val("" & mytablex.Fields("tot"))
        nlineas

        mytablex.MoveNext
    Loop
    found = formateaa("", 13, 0, 0)
    buf = Format(suma8, "0.00")
    found = formateaa(buf, 10, 0, 1)
   
    buf = Format(suma1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma2, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma3, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma4, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma7, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma5, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(ssuma6, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 13, 0, 0)
    buf = Format(ssuma8, "0.00")
    found = formateaa(buf, 10, 0, 1)
    buf = Format(ssuma1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(ssuma2, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(ssuma3, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(ssuma4, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(ssuma7, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(ssuma5, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)

End Sub

Sub menu_meses()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    If Combo1 = "Fecha" Then Exit Sub
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    contpag = 0

    found = sql_documento_meses1(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento_meses1
    cuerpo_programa_documento_meses1 mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Function sql_documento_meses1(mytablex As ADODB.Recordset)

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    If Combo1 = "TipoDocumento" Then
        buf = "select Tipo,month(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Codigo" Then
        buf = "select Codigo,Month(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Hora" Then
        buf = "select left(hora,2) as Hora,Month(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Vendedor" Then
        buf = "select Vendedor,Month(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Zona" Then
        buf = "select Zona,Month(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Cajero" Then
        buf = "select Usuario,Month(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Caja" Then
        buf = "select Caja,Month(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Mesa" Then
        buf = "select Mesa,Month(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Turno" Then
        buf = "select Turno,Month(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Local" Then
        buf = "select Local,Month(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Servicio" Then
        buf = "select Servicio,Month(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Almacen" Then
        buf = "select Bodega,Month(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If tipofecha = "E" Then
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    Else
        buf = buf & "  fechae>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fechae<='" & Format(fechaf, "YYYYMMDD") & "' "

    End If

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & extra_loquesea("" & tipo) & "'"

    End If

    If serie <> "%" Then
        buf = buf & " and serie like '" & serie & "'"

    End If

    If Numero <> "%" Then
        buf = buf & " and numero like '" & Numero & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If salon <> "%" Then
        buf = buf & " and salon like '" & extra_loquesea(salon) & "'"

    End If

    If mesa <> "%" Then
        buf = buf & " and mesa like '" & mesa & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "'"

    End If

    If horai <> "%" And horaf <> "%" Then
        If Val(horaf) >= Val(horai) Then
            buf = buf & " and hour(hora)>=" & Val(horai)
            buf = buf & " and hour(hora)<=" & Val(horaf)

        End If

    End If

    If servicio <> "%" Then
        buf = buf & " and  servicio='" & extra_loquesea(servicio) & "'"

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    If transporte <> "%" Then
        buf = buf & " and transporte like '" & transporte & "'"

    End If

    If fpago <> "%" Then
        buf = buf & " and fpago like '" & fpago & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    If acu <> "C" And acu <> "V" Then
        buf = buf & " and acu='" & acu & "'"

    End If

    If acu = "C" Then
        buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P')"

    End If

    If acu = "V" Then
        '19/06/2017 kenyo NOTA DE CREDITO
        '  buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G')"
        buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E')"

        '19/06/2017 kenyo NOTA DE CREDITO
    End If

    If estado <> "%" Then
        buf = buf & " and estado like '" & estado & "'"

    End If

    If Combo1 = "TipoDocumento" Then
        buf = buf & " group by Tipo,month(Fecha) order by tipo"

    End If

    If Combo1 = "Codigo" Then
        buf = buf & " group by Codigo,month(Fecha) order by codigo "

    End If

    If Combo1 = "Hora" Then
        buf = buf & " group by left(hora,2),month(Fecha) order by left(hora,2)"

    End If

    If Combo1 = "Vendedor" Then
        buf = buf & " group by Vendedor,month(Fecha)  order by vendedor"

    End If

    If Combo1 = "Zona" Then
        buf = buf & " group by Zona,month(Fecha) order by zona"

    End If

    If Combo1 = "Cajero" Then
        buf = buf & " group by Usuario,month(Fecha) order by usuario"

    End If

    If Combo1 = "Caja" Then
        buf = buf & " group by Caja,month(Fecha) order by caja"

    End If

    If Combo1 = "Mesa" Then
        buf = buf & " group by Mesa,month(Fecha) order by mesa"

    End If

    If Combo1 = "Turno" Then
        buf = buf & " group by Turno,month(Fecha) order by turno"

    End If

    If Combo1 = "Local" Then
        buf = buf & " group by Local,month(Fecha) order by local"

    End If

    If Combo1 = "Servicio" Then
        buf = buf & " group by Servicio,month(Fecha) order by Servicio"

    End If

    If Combo1 = "Almacen" Then
        buf = buf & " group by Bodega,month(Fecha) order by bodega"

    End If

    'MsgBox buf

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    sql_documento_meses1 = 1

End Function

Sub cabecera_documento_meses1()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    buf = String(164, "-")
    found = formateaa(buf, 164, 2, 0)
    found = formateaa("Descripcio ", 20, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Enero", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Febrero", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Marzo", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Abril", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Mayo", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Junio", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Julio", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Agosto", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Setiembre", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Octubre", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Noviembre", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Diciembre", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Total", 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    
    buf = String(164, "-")
    found = formateaa(buf, 164, 2, 0)

End Sub

Sub cuerpo_programa_documento_meses1(mytablex As ADODB.Recordset)

    Dim Tmp   As String

    Dim sw    As Integer

    Dim buf   As String

    Dim sdx1  As Double

    Dim found As Integer

    Dim tmp1  As String

    ReDim xmeses(13) As Double
    ReDim xmeses1(13) As Double

    Dim I As Integer

    sdx1 = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0
    sw = 0

    For I = 1 To 12
        xmeses(I) = 0#
        xmeses1(I) = 0#
    Next I

    tmp1 = ""
    Do

        If mytablex.EOF Then Exit Do
   
        If Combo1 = "TipoDocumento" Then
            tmp1 = "" & mytablex.Fields("Tipo")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("Codigo")

        End If

        If Combo1 = "Hora" Then
            tmp1 = "" & mytablex.Fields("Hora")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("Vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("usuario")

        End If

        If Combo1 = "Caja" Then
            tmp1 = "" & mytablex.Fields("Caja")

        End If

        If Combo1 = "Mesa" Then
            tmp1 = "" & mytablex.Fields("Mesa")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("Turno")

        End If

        If Combo1 = "Local" Then
            tmp1 = "" & mytablex.Fields("Local")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("Bodega")

        End If

        If Combo1 = "Servicio" Then
            tmp1 = "" & mytablex.Fields("Servicio")

        End If

        If sw = 0 Then
            If Combo1 = "TipoDocumento" Then
                buf = busca_tipo("" & mytablex.Fields("Tipo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Tipo")

            End If

            If Combo1 = "Codigo" Then
                buf = busca_cliente("" & mytablex.Fields("codigo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Hora" Then
                buf = busca_cliente("" & mytablex.Fields("Hora"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Hora")

            End If

            If Combo1 = "Vendedor" Then
                buf = busca_vendedor("" & mytablex.Fields("Vendedor"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Zona")

            End If

            If Combo1 = "Cajero" Then
                buf = busca_vendedor("" & mytablex.Fields("Usuario"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Usuario")

            End If

            If Combo1 = "Caja" Then
                buf = busca_caja("" & mytablex.Fields("Caja"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("mesa")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("mesa")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Turno")

            End If

            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Local")

            End If

            If Combo1 = "Almacen" Then
                buf = busca_bodega("" & mytablex.Fields("bodega"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Servicio" Then
                buf = busca_servicio("" & mytablex.Fields("servicio"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("servicio")

            End If

            nlineas
            sw = 1

        End If

        If Tmp <> tmp1 Then
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            sdx1 = 0

            For I = 1 To 12
                buf = Format(xmeses(I), "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                sdx1 = sdx1 + xmeses(I)
            Next I

            buf = Format(sdx1, "0.00")
            found = formateaa(buf, 10, 0, 1)

            For I = 1 To 12
                xmeses(I) = 0#
            Next I

            found = formateaa("", 1, 2, 0)
            nlineas

            If Combo1 = "TipoDocumento" Then
                buf = busca_tipo("" & mytablex.Fields("Tipo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Tipo")

            End If

            If Combo1 = "Codigo" Then
                buf = busca_cliente("" & mytablex.Fields("codigo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("Hora")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Hora")

            End If

            If Combo1 = "Vendedor" Then
                buf = busca_vendedor("" & mytablex.Fields("Vendedor"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Zona")

            End If

            If Combo1 = "Cajero" Then
                buf = busca_vendedor("" & mytablex.Fields("Usuario"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Usuario")

            End If

            If Combo1 = "Caja" Then
                buf = busca_caja("" & mytablex.Fields("Caja"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("mesa")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("mesa")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Turno")

            End If

            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Local")

            End If

            If Combo1 = "Almacen" Then
                buf = busca_bodega("" & mytablex.Fields("Bodega"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Servicio" Then
                buf = busca_servicio("" & mytablex.Fields("servicio"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("servicio")

            End If

            nlineas

        End If

        '-----------------aqui se debe sumar los meses-----

        If tipores = "Monto" Then
            xmeses(CInt("" & mytablex.Fields("xmes"))) = xmeses(CInt("" & mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("tot"))
            xmeses1(CInt("" & mytablex.Fields("xmes"))) = xmeses1(CInt("" & mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("tot"))

        End If

        If tipores = "Cantidad" Then
            xmeses(CInt("" & mytablex.Fields("xmes"))) = xmeses(CInt("" & mytablex.Fields("xmes"))) + 1
            xmeses1(CInt("" & mytablex.Fields("xmes"))) = xmeses1(CInt("" & mytablex.Fields("xmes"))) + 1

        End If
   
        '--------------------------------------------------
        mytablex.MoveNext
    Loop
    sdx1 = 0

    For I = 1 To 12
        buf = Format(xmeses(I), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx1 = sdx1 + xmeses(I)
    Next I

    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 21, 0, 0)
    sdx1 = 0

    For I = 1 To 12
        buf = Format(xmeses1(I), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx1 = sdx1 + xmeses1(I)
    Next I

    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)

End Sub

Sub sumar_como_paga(mytabley As ADODB.Recordset)

    Dim found    As Integer

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim sdx2     As Double

    Dim sdx3     As Double

    sdx = 0
    sdx1 = 0
    sdx2 = 0
    sdx3 = 0
    mytablex.Open "select * from fpagov where local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "' and numero='" & "" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            If "" & mytablex.Fields("local") = "" & mytabley.Fields("local") And "" & mytablex.Fields("tipo") = "" & mytabley.Fields("tipo") And "" & mytablex.Fields("serie") = "" & mytabley.Fields("serie") And "" & mytablex.Fields("numero") = "" & mytabley.Fields("numero") Then

                Select Case "" & mytablex.Fields("acufp")

                    Case "A"
                        sdx = sdx + Val("" & mytablex.Fields("recibe"))
                        suma6 = suma6 + Val("" & mytablex.Fields("recibe"))

                    Case "B"
                        sdx1 = sdx1 + Val("" & mytablex.Fields("recibe"))
                        suma7 = suma7 + Val("" & mytablex.Fields("recibe"))

                    Case "C"
                        sdx2 = sdx2 + Val("" & mytablex.Fields("recibe"))
                        suma8 = suma8 + Val("" & mytablex.Fields("recibe"))

                    Case Else
                        sdx3 = sdx3 + Val("" & mytablex.Fields("recibe"))
                        suma9 = suma9 + Val("" & mytablex.Fields("recibe"))

                End Select

            Else
                Exit Do

            End If

            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    buf = Format(sdx, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(sdx2, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(sdx3, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 1, 0, 0)
   
End Sub

Sub imprime_cuentag(mytabley As ADODB.Recordset)

    Dim mytablex As New ADODB.Recordset

    Dim fcredito As Double

    Dim fsoles   As Double

    Dim buf      As String

    Dim found    As Integer

    Dim sdx      As Double

    sdx = 0
    fcredito = 0
    fsoles = 0
    mytablex.Open "select * from " & gofpago & " where local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "' and numero='" & "" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("local") = "" & mytabley.Fields("local") And "" & mytablex.Fields("tipo") = "" & mytabley.Fields("tipo") And "" & mytablex.Fields("serie") = "" & mytabley.Fields("serie") And "" & mytablex.Fields("numero") = "" & mytabley.Fields("numero") Then
            If "" & mytablex.Fields("acufp") = "C" Then   'credito
                fcredito = fcredito + Val("" & mytablex.Fields("recibe"))
            Else
                fsoles = fsoles + Val("" & mytablex.Fields("recibe"))

            End If

        Else
            Exit Do

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    buf = Format(fsoles, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    sdx = Val("" & mytabley.Fields("total")) - fsoles
    buf = Format(sdx, "0.00")
    found = formateaa(buf, 10, 2, 1)
    nlineas
    suma1 = Val(buf)

End Sub

Sub imprime_cuentacd(mytabley As ADODB.Recordset)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim found    As Integer

    Dim sdx      As Double

    sdx = 0
    mytablex.Open "select * from cuentacd where local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "' and numero='" & "" & mytabley.Fields("numero") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("local1") = "" & mytabley.Fields("local") And "" & mytablex.Fields("tipo1") = "" & mytabley.Fields("tipo") And "" & mytablex.Fields("serie1") = "" & mytabley.Fields("serie") And "" & mytablex.Fields("numero1") = "" & mytabley.Fields("numero") Then
            found = formateaa("", 65, 0, 0)
            buf = "" & mytablex.Fields("tipo")
            found = formateaa(buf, 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("serie")
            found = formateaa(buf, 4, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("numero")
            found = formateaa(buf, 11, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("fecha")
            found = formateaa(buf, 10, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("", 10, 0, 0)
            buf = Format(Val("" & mytablex.Fields("paga")), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            suma1 = suma1 - Val("" & mytablex.Fields("paga"))
            buf = Format(suma1, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 2, 0)
            nlineas
        Else
            Exit Do

        End If

        mytablex.MoveNext
    Loop

End Sub

'ventas semanales
Sub menu_semanas()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    If Combo1 = "Fecha" Then Exit Sub
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    contpag = 0

    found = sql_documento_semanas1(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento_semanas1
    cuerpo_programa_documento_semanas1 mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Function sql_documento_semanas1(mytablex As ADODB.Recordset)

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function

    '("" & mysnap.Fields("fecha"))
    If Combo1 = "TipoDocumento" Then
        buf = "select Tipo,(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Codigo" Then
        buf = "select Codigo,(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Hora" Then
        buf = "select left(hora,2) as Hora,(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Vendedor" Then
        buf = "select Vendedor,(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Zona" Then
        buf = "select Zona,(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Cajero" Then
        buf = "select Usuario,(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Caja" Then
        buf = "select Caja,(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Mesa" Then
        buf = "select Mesa,(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Turno" Then
        buf = "select Turno,(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Local" Then
        buf = "select Local,(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Almacen" Then
        buf = "select Bodega,(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Servicio" Then
        buf = "select Servicio,(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If tipofecha = "E" Then
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    Else
        buf = buf & "  fechae>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fechae<='" & Format(fechaf, "YYYYMMDD") & "' "

    End If

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & extra_loquesea("" & tipo) & "'"

    End If

    If serie <> "%" Then
        buf = buf & " and serie like '" & serie & "'"

    End If

    If Numero <> "%" Then
        buf = buf & " and numero like '" & Numero & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "'"

    End If

    If horai <> "%" And horaf <> "%" Then
        If Val(horaf) >= Val(horai) Then
            buf = buf & " and hour(hora)>=" & Val(horai)
            buf = buf & " and hour(hora)<=" & Val(horaf)

        End If

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    If transporte <> "%" Then
        buf = buf & " and transporte like '" & transporte & "'"

    End If

    If fpago <> "%" Then
        buf = buf & " and fpago like '" & fpago & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    If acu <> "C" And acu <> "V" Then
        buf = buf & " and acu='" & acu & "'"

    End If

    If acu = "C" Then
        buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P')"

    End If

    If acu = "V" Then
        '19/06/2017 kenyo NOTA DE CREDITO
        ' buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G')"
        buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E')"

        '19/06/2017 kenyo NOTA DE CREDITO
    End If

    If servicio <> "%" Then
        buf = buf & " and  servicio='" & extra_loquesea(servicio) & "'"

    End If

    If estado <> "%" Then
        buf = buf & " and estado like '" & estado & "'"

    End If

    If Combo1 = "TipoDocumento" Then
        buf = buf & " group by Tipo,(Fecha) order by tipo"

    End If

    If Combo1 = "Codigo" Then
        buf = buf & " group by Codigo,(Fecha) order by codigo "

    End If

    If Combo1 = "Hora" Then
        buf = buf & " group by left(hora,2),(Fecha) order by left(hora,2)"

    End If

    If Combo1 = "Vendedor" Then
        buf = buf & " group by Vendedor,(Fecha)  order by vendedor"

    End If

    If Combo1 = "Zona" Then
        buf = buf & " group by Zona,(Fecha) order by zona"

    End If

    If Combo1 = "Cajero" Then
        buf = buf & " group by Usuario,(Fecha) order by usuario"

    End If

    If Combo1 = "Caja" Then
        buf = buf & " group by Caja,(Fecha) order by caja"

    End If

    If Combo1 = "Mesa" Then
        buf = buf & " group by Mesa,(Fecha) order by mesa"

    End If

    If Combo1 = "Turno" Then
        buf = buf & " group by Turno,(Fecha) order by turno"

    End If

    If Combo1 = "Local" Then
        buf = buf & " group by Local,(Fecha) order by local"

    End If

    If Combo1 = "Almacen" Then
        buf = buf & " group by Bodega,(Fecha) order by bodega"

    End If

    If Combo1 = "Servicio" Then
        buf = buf & " group by Servicio,(Fecha) order by Servicio"

    End If

    'MsgBox buf

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    sql_documento_semanas1 = 1

End Function

Sub cabecera_documento_semanas1()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    buf = String(164, "-")
    found = formateaa(buf, 164, 2, 0)
    found = formateaa("Descripcio ", 20, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Domingo", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Lunes", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Martes", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Mierco", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Jueves", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Viernes", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Sabado", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Total", 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    
    buf = String(164, "-")
    found = formateaa(buf, 164, 2, 0)

End Sub

Sub cuerpo_programa_documento_semanas1(mytablex As ADODB.Recordset)

    Dim Tmp   As String

    Dim sw    As Integer

    Dim buf   As String

    Dim sdx1  As Double

    Dim found As Integer

    Dim tmp1  As String

    ReDim xmeses(13) As Double
    ReDim xmeses1(13) As Double

    Dim I As Integer

    sdx1 = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0
    sw = 0

    For I = 1 To 12
        xmeses(I) = 0#
        xmeses1(I) = 0#
    Next I

    tmp1 = ""
    Do

        If mytablex.EOF Then Exit Do
   
        If Combo1 = "TipoDocumento" Then
            tmp1 = "" & mytablex.Fields("Tipo")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("Codigo")

        End If

        If Combo1 = "Hora" Then
            tmp1 = "" & mytablex.Fields("Hora")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("Vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("usuario")

        End If

        If Combo1 = "Caja" Then
            tmp1 = "" & mytablex.Fields("Caja")

        End If

        If Combo1 = "Mesa" Then
            tmp1 = "" & mytablex.Fields("Mesa")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("Turno")

        End If

        If Combo1 = "Local" Then
            tmp1 = "" & mytablex.Fields("Local")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("Bodega")

        End If

        If Combo1 = "Servicio" Then
            tmp1 = "" & mytablex.Fields("Servicio")

        End If

        If sw = 0 Then
            If Combo1 = "TipoDocumento" Then
                buf = busca_tipo("" & mytablex.Fields("Tipo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Tipo")

            End If

            If Combo1 = "Codigo" Then
                buf = busca_cliente("" & mytablex.Fields("codigo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Hora" Then
                buf = busca_cliente("" & mytablex.Fields("Hora"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Hora")

            End If

            If Combo1 = "Vendedor" Then
                buf = busca_vendedor("" & mytablex.Fields("Vendedor"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Zona")

            End If

            If Combo1 = "Cajero" Then
                buf = busca_vendedor("" & mytablex.Fields("Usuario"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Usuario")

            End If

            If Combo1 = "Caja" Then
                buf = busca_caja("" & mytablex.Fields("Caja"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("mesa")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("mesa")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Turno")

            End If

            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Local")

            End If

            If Combo1 = "Almacen" Then
                buf = busca_bodega("" & mytablex.Fields("bodega"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Servicio" Then
                buf = busca_servicio("" & mytablex.Fields("servicio"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("servicio")

            End If

            nlineas
            sw = 1

        End If

        If Tmp <> tmp1 Then
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            sdx1 = 0

            For I = 1 To 7
                buf = Format(xmeses(I), "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                sdx1 = sdx1 + xmeses(I)
            Next I

            buf = Format(sdx1, "0.00")
            found = formateaa(buf, 10, 0, 1)

            For I = 1 To 12
                xmeses(I) = 0#
            Next I

            found = formateaa("", 1, 2, 0)
            nlineas

            If Combo1 = "TipoDocumento" Then
                buf = busca_tipo("" & mytablex.Fields("Tipo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Tipo")

            End If

            If Combo1 = "Codigo" Then
                buf = busca_cliente("" & mytablex.Fields("codigo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("Hora")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Hora")

            End If

            If Combo1 = "Vendedor" Then
                buf = busca_vendedor("" & mytablex.Fields("Vendedor"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Zona")

            End If

            If Combo1 = "Cajero" Then
                buf = busca_vendedor("" & mytablex.Fields("Usuario"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Usuario")

            End If

            If Combo1 = "Caja" Then
                buf = busca_caja("" & mytablex.Fields("Caja"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("mesa")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("mesa")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Turno")

            End If

            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Local")

            End If

            If Combo1 = "servicio" Then
                buf = busca_servicio("" & mytablex.Fields("servicio"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("servicio")

            End If

            If Combo1 = "Almacen" Then
                buf = busca_bodega("" & mytablex.Fields("Bodega"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("bodega")

            End If

            nlineas

        End If

        '-----------------aqui se debe sumar los meses-----
        If tipores = "Monto" Then
            xmeses(Weekday(mytablex.Fields("xmes"))) = xmeses(Weekday(mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("tot"))
            xmeses1(Weekday(mytablex.Fields("xmes"))) = xmeses1(Weekday(mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("tot"))

        End If

        If tipores = "Cantidad" Then
            xmeses(Weekday(mytablex.Fields("xmes"))) = xmeses(Weekday(mytablex.Fields("xmes"))) + 1
            xmeses1(Weekday(mytablex.Fields("xmes"))) = xmeses1(Weekday(mytablex.Fields("xmes"))) + 1

        End If

        '--------------------------------------------------
        mytablex.MoveNext
    Loop
    sdx1 = 0

    For I = 1 To 7
        buf = Format(xmeses(I), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx1 = sdx1 + xmeses(I)
    Next I

    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 21, 0, 0)
    sdx1 = 0

    For I = 1 To 7
        buf = Format(xmeses1(I), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx1 = sdx1 + xmeses1(I)
    Next I

    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)

End Sub

'rutinas de dias
Sub menu_dias()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    If Combo1 = "Fecha" Then Exit Sub
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    contpag = 0

    found = sql_documento_dias1(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento_dias1
    cuerpo_programa_documento_dias1 mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Function sql_documento_dias1(mytablex As ADODB.Recordset)

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    If Combo1 = "TipoDocumento" Then
        buf = "select Tipo,day(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Codigo" Then
        buf = "select Codigo,day(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Hora" Then
        buf = "select left(hora,2) as Hora,day(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Vendedor" Then
        buf = "select Vendedor,day(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Zona" Then
        buf = "select Zona,day(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Cajero" Then
        buf = "select Usuario,day(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Caja" Then
        buf = "select Caja,day(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Mesa" Then
        buf = "select Mesa,day(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Turno" Then
        buf = "select Turno,day(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Local" Then
        buf = "select Local,day(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Servicio" Then
        buf = "select Servicio,day(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Almacen" Then
        buf = "select Bodega,day(fecha) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If tipofecha = "E" Then
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    Else
        buf = buf & "  fechae>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fechae<='" & Format(fechaf, "YYYYMMDD") & "' "

    End If

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & extra_loquesea("" & tipo) & "'"

    End If

    If serie <> "%" Then
        buf = buf & " and serie like '" & serie & "'"

    End If

    If Numero <> "%" Then
        buf = buf & " and numero like '" & Numero & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If salon <> "%" Then
        buf = buf & " and salon like '" & extra_loquesea(salon) & "'"

    End If

    If mesa <> "%" Then
        buf = buf & " and mesa like '" & mesa & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "'"

    End If

    If horai <> "%" And horaf <> "%" Then
        If Val(horaf) >= Val(horai) Then
            buf = buf & " and hour(hora)>=" & Val(horai)
            buf = buf & " and hour(hora)<=" & Val(horaf)

        End If

    End If

    If servicio <> "%" Then
        buf = buf & " and  servicio='" & extra_loquesea(servicio) & "'"

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    If transporte <> "%" Then
        buf = buf & " and transporte like '" & transporte & "'"

    End If

    If fpago <> "%" Then
        buf = buf & " and fpago like '" & fpago & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    If acu <> "C" And acu <> "V" Then
        buf = buf & " and acu='" & acu & "'"

    End If

    If acu = "C" Then
        buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P')"

    End If

    If acu = "V" Then
        '19/06/2017 kenyo NOTA DE CREDITO
        'buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G')"
        buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E')"

        '19/06/2017 kenyo NOTA DE CREDITO
    End If

    If estado <> "%" Then
        buf = buf & " and estado like '" & estado & "'"

    End If

    If Combo1 = "TipoDocumento" Then
        buf = buf & " group by Tipo,day(Fecha) order by tipo"

    End If

    If Combo1 = "Codigo" Then
        buf = buf & " group by Codigo,day(Fecha) order by codigo "

    End If

    If Combo1 = "Hora" Then
        buf = buf & " group by left(hora,2),day(Fecha) order by left(hora,2)"

    End If

    If Combo1 = "Vendedor" Then
        buf = buf & " group by Vendedor,day(Fecha)  order by vendedor"

    End If

    If Combo1 = "Zona" Then
        buf = buf & " group by Zona,day(Fecha) order by zona"

    End If

    If Combo1 = "Cajero" Then
        buf = buf & " group by Usuario,day(Fecha) order by usuario"

    End If

    If Combo1 = "Caja" Then
        buf = buf & " group by Caja,day(Fecha) order by caja"

    End If

    If Combo1 = "Mesa" Then
        buf = buf & " group by Mesa,day(Fecha) order by mesa"

    End If

    If Combo1 = "Turno" Then
        buf = buf & " group by Turno,day(Fecha) order by turno"

    End If

    If Combo1 = "Local" Then
        buf = buf & " group by Local,day(Fecha) order by local"

    End If

    If Combo1 = "Servicio" Then
        buf = buf & " group by Servicio,day(Fecha) order by Servicio"

    End If

    If Combo1 = "Almacen" Then
        buf = buf & " group by Bodega,day(Fecha) order by bodega"

    End If

    'MsgBox buf

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    sql_documento_dias1 = 1

End Function

Sub cabecera_documento_dias1()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    buf = String(250, "-")
    found = formateaa(buf, 250, 2, 0)
    found = formateaa("Descripcio ", 20, 0, 0)
    found = formateaa("", 1, 0, 0)
    
    For I = 1 To 31
        found = formateaa(Format(I, "00"), 10, 0, 1)
        found = formateaa("", 1, 0, 0)
    Next I

    found = formateaa("", 1, 2, 0)
    
    buf = String(250, "-")
    found = formateaa(buf, 250, 2, 0)

End Sub

Sub cuerpo_programa_documento_dias1(mytablex As ADODB.Recordset)

    Dim Tmp   As String

    Dim sw    As Integer

    Dim buf   As String

    Dim sdx1  As Double

    Dim found As Integer

    Dim tmp1  As String

    ReDim xmeses(32) As Double
    ReDim xmeses1(32) As Double

    Dim I As Integer

    sdx1 = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0
    sw = 0

    For I = 1 To 12
        xmeses(I) = 0#
        xmeses1(I) = 0#
    Next I

    tmp1 = ""
    Do

        If mytablex.EOF Then Exit Do
   
        If Combo1 = "TipoDocumento" Then
            tmp1 = "" & mytablex.Fields("Tipo")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("Codigo")

        End If

        If Combo1 = "Hora" Then
            tmp1 = "" & mytablex.Fields("Hora")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("Vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("usuario")

        End If

        If Combo1 = "Caja" Then
            tmp1 = "" & mytablex.Fields("Caja")

        End If

        If Combo1 = "Mesa" Then
            tmp1 = "" & mytablex.Fields("Mesa")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("Turno")

        End If

        If Combo1 = "Local" Then
            tmp1 = "" & mytablex.Fields("Local")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("Bodega")

        End If

        If Combo1 = "Servicio" Then
            tmp1 = "" & mytablex.Fields("Servicio")

        End If

        If sw = 0 Then
            If Combo1 = "TipoDocumento" Then
                buf = busca_tipo("" & mytablex.Fields("Tipo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Tipo")

            End If

            If Combo1 = "Codigo" Then
                buf = busca_cliente("" & mytablex.Fields("codigo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Hora" Then
                buf = busca_cliente("" & mytablex.Fields("Hora"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Hora")

            End If

            If Combo1 = "Vendedor" Then
                buf = busca_vendedor("" & mytablex.Fields("Vendedor"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Zona")

            End If

            If Combo1 = "Cajero" Then
                buf = busca_vendedor("" & mytablex.Fields("Usuario"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Usuario")

            End If

            If Combo1 = "Caja" Then
                buf = busca_caja("" & mytablex.Fields("Caja"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("mesa")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("mesa")

            End If
   
            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Turno")

            End If

            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Local")

            End If

            If Combo1 = "Almacen" Then
                buf = busca_bodega("" & mytablex.Fields("bodega"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Servicio" Then
                buf = busca_servicio("" & mytablex.Fields("servicio"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("servicio")

            End If

            nlineas
            sw = 1

        End If

        If Tmp <> tmp1 Then
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            sdx1 = 0

            For I = 1 To 31
                buf = Format(xmeses(I), "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                sdx1 = sdx1 + xmeses(I)
            Next I

            buf = Format(sdx1, "0.00")
            found = formateaa(buf, 10, 0, 1)

            For I = 1 To 31
                xmeses(I) = 0#
            Next I

            found = formateaa("", 1, 2, 0)
            nlineas

            If Combo1 = "TipoDocumento" Then
                buf = busca_tipo("" & mytablex.Fields("Tipo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Tipo")

            End If

            If Combo1 = "Codigo" Then
                buf = busca_cliente("" & mytablex.Fields("codigo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("Hora")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Hora")

            End If

            If Combo1 = "Vendedor" Then
                buf = busca_vendedor("" & mytablex.Fields("Vendedor"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Zona")

            End If

            If Combo1 = "Cajero" Then
                buf = busca_vendedor("" & mytablex.Fields("Usuario"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Usuario")

            End If

            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("mesa")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("mesa")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Turno")

            End If

            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Local")

            End If

            If Combo1 = "Almacen" Then
                buf = busca_bodega("" & mytablex.Fields("Bodega"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Servicio" Then
                buf = busca_servicio("" & mytablex.Fields("servicio"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("servicio")

            End If

            nlineas

        End If

        '-----------------aqui se debe sumar los meses-----
        If tipores = "Monto" Then
            xmeses(CInt("" & mytablex.Fields("xmes"))) = xmeses(CInt("" & mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("tot"))
            xmeses1(CInt("" & mytablex.Fields("xmes"))) = xmeses1(CInt("" & mytablex.Fields("xmes"))) + Val("" & mytablex.Fields("tot"))

        End If

        If tipores = "Cantidad" Then
            xmeses(CInt("" & mytablex.Fields("xmes"))) = xmeses(CInt("" & mytablex.Fields("xmes"))) + 1
            xmeses1(CInt("" & mytablex.Fields("xmes"))) = xmeses1(CInt("" & mytablex.Fields("xmes"))) + 1

        End If

        '--------------------------------------------------
        mytablex.MoveNext
    Loop
    sdx1 = 0

    For I = 1 To 31
        buf = Format(xmeses(I), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx1 = sdx1 + xmeses(I)
    Next I

    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 21, 0, 0)
    sdx1 = 0

    For I = 1 To 31
        buf = Format(xmeses1(I), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx1 = sdx1 + xmeses1(I)
    Next I

    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)

End Sub

Sub menu_codigo()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    If Combo1 = "Fecha" Then Exit Sub
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    contpag = 0

    found = sql_codigo(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_codigo
    cuerpo_programa_codigo mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Function sql_codigo(mytablex As ADODB.Recordset)

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function

    '("" & mysnap.Fields("fecha"))
    If Combo1 = "TipoDocumento" Then
        buf = "select Tipo,(servicio) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Codigo" Then
        buf = "select Codigo,(servicio) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Hora" Then
        buf = "select left(hora,2) as Hora,(servicio) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Vendedor" Then
        buf = "select Vendedor,(servicio) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Zona" Then
        buf = "select Zona,(servicio) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Cajero" Then
        buf = "select Usuario,(servicio) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Caja" Then
        buf = "select Caja,(servicio) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Mesa" Then
        buf = "select Mesa,(servicio) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Turno" Then
        buf = "select Turno,(servicio) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Local" Then
        buf = "select Local,(servicio) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Almacen" Then
        buf = "select Bodega,(servicio) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Servicio" Then
        buf = "select Servicio,(servicio) as xmes,sum(total) as Tot from " & cgusuario & " where "

    End If

    If tipofecha = "E" Then
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    Else
        buf = buf & "  fechae>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fechae<='" & Format(fechaf, "YYYYMMDD") & "' "

    End If

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    If salon <> "%" Then
        buf = buf & " and salon like '" & extra_loquesea(salon) & "'"

    End If

    If mesa <> "%" Then
        buf = buf & " and mesa like '" & mesa & "'"

    End If

    If salon <> "%" Then
        buf = buf & " and salon like '" & extra_loquesea(salon) & "'"

    End If

    If mesa <> "%" Then
        buf = buf & " and mesa like '" & mesa & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & extra_loquesea("" & tipo) & "'"

    End If

    If serie <> "%" Then
        buf = buf & " and serie like '" & serie & "'"

    End If

    If Numero <> "%" Then
        buf = buf & " and numero like '" & Numero & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "'"

    End If

    If horai <> "%" And horaf <> "%" Then
        If Val(horaf) >= Val(horai) Then
            buf = buf & " and hour(hora)>=" & Val(horai)
            buf = buf & " and hour(hora)<=" & Val(horaf)

        End If

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    If transporte <> "%" Then
        buf = buf & " and transporte like '" & transporte & "'"

    End If

    If fpago <> "%" Then
        buf = buf & " and fpago like '" & fpago & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    If acu <> "C" And acu <> "V" Then
        buf = buf & " and acu='" & acu & "'"

    End If

    If acu = "C" Then
        buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P')"

    End If

    If acu = "V" Then
        '19/06/2017 kenyo NOTA DE CREDITO
        ' buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G')"
        buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E')"
  
        '19/06/2017 kenyo NOTA DE CREDITO
   
    End If

    If servicio <> "%" Then
        buf = buf & " and  servicio='" & extra_loquesea(servicio) & "'"

    End If

    If estado <> "%" Then
        buf = buf & " and estado like '" & estado & "'"

    End If

    If Combo1 = "TipoDocumento" Then
        buf = buf & " group by Tipo,(servicio) order by tipo"

    End If

    If Combo1 = "Codigo" Then
        buf = buf & " group by Codigo,(servicio) order by codigo,sum(total) "

    End If

    If Combo1 = "Hora" Then
        buf = buf & " group by left(hora,2),(servicio) order by left(hora,2),sum(total)"

    End If

    If Combo1 = "Vendedor" Then
        buf = buf & " group by Vendedor,(servicio)  order by vendedor,sum(total)"

    End If

    If Combo1 = "Zona" Then
        buf = buf & " group by Zona,(servicio) order by zona,sum(total)"

    End If

    If Combo1 = "Cajero" Then
        buf = buf & " group by Usuario,(servicio) order by usuario,sum(total)"

    End If

    If Combo1 = "Caja" Then
        buf = buf & " group by Caja,(servicio) order by caja,sum(total)"

    End If

    If Combo1 = "Mesa" Then
        buf = buf & " group by Mesa,(servicio) order by mesa,sum(total)"

    End If

    If Combo1 = "Turno" Then
        buf = buf & " group by Turno,(servicio) order by turno,sum(total)"

    End If

    If Combo1 = "Local" Then
        buf = buf & " group by Local,(servicio) order by local,sum(total)"

    End If

    If Combo1 = "Almacen" Then
        buf = buf & " group by Bodega,(servicio) order by bodega,sum(total)"

    End If

    If Combo1 = "Servicio" Then
        buf = buf & " group by Servicio,(servicio) order by Servicio,sum(total)"

    End If

    'buf = buf & " order by tot "
    'MsgBox buf
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    sql_codigo = 1

End Function

Sub cabecera_codigo()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    buf = String(164, "-")
    found = formateaa(buf, 164, 2, 0)
    found = formateaa("Descripcio ", 20, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("AutoSer", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Comanda", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Delivery", 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Total", 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    
    buf = String(164, "-")
    found = formateaa(buf, 164, 2, 0)

End Sub

Sub cuerpo_programa_codigo(mytablex As ADODB.Recordset)

    Dim Tmp   As String

    Dim sw    As Integer

    Dim buf   As String

    Dim sdx1  As Double

    Dim found As Integer

    Dim tmp1  As String

    ReDim xmeses(13) As Double
    ReDim xmeses1(13) As Double

    Dim I As Integer

    sdx1 = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0
    sw = 0

    For I = 1 To 12
        xmeses(I) = 0#
        xmeses1(I) = 0#
    Next I

    tmp1 = ""
    Do

        If mytablex.EOF Then Exit Do
   
        If Combo1 = "TipoDocumento" Then
            tmp1 = "" & mytablex.Fields("Tipo")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("Codigo")

        End If

        If Combo1 = "Hora" Then
            tmp1 = "" & mytablex.Fields("Hora")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("Vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("usuario")

        End If

        If Combo1 = "Caja" Then
            tmp1 = "" & mytablex.Fields("Caja")

        End If

        If Combo1 = "Mesa" Then
            tmp1 = "" & mytablex.Fields("Mesa")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("Turno")

        End If

        If Combo1 = "Local" Then
            tmp1 = "" & mytablex.Fields("Local")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("Bodega")

        End If

        If Combo1 = "Servicio" Then
            tmp1 = "" & mytablex.Fields("Servicio")

        End If

        If sw = 0 Then
            If Combo1 = "TipoDocumento" Then
                buf = busca_tipo("" & mytablex.Fields("Tipo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Tipo")

            End If

            If Combo1 = "Codigo" Then
                buf = busca_cliente("" & mytablex.Fields("codigo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Hora" Then
                buf = busca_cliente("" & mytablex.Fields("Hora"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Hora")

            End If

            If Combo1 = "Vendedor" Then
                buf = busca_vendedor("" & mytablex.Fields("Vendedor"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Zona")

            End If

            If Combo1 = "Cajero" Then
                buf = busca_vendedor("" & mytablex.Fields("Usuario"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Usuario")

            End If

            If Combo1 = "Caja" Then
                buf = busca_caja("" & mytablex.Fields("Caja"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("Mesa")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("mesa")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Turno")

            End If

            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Local")

            End If

            If Combo1 = "Almacen" Then
                buf = busca_bodega("" & mytablex.Fields("bodega"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Servicio" Then
                buf = busca_servicio("" & mytablex.Fields("servicio"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("servicio")

            End If

            nlineas
            sw = 1

        End If

        If Tmp <> tmp1 Then
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            sdx1 = 0

            For I = 1 To 3
                buf = Format(xmeses(I), "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                sdx1 = sdx1 + xmeses(I)
            Next I

            buf = Format(sdx1, "0.00")
            found = formateaa(buf, 10, 0, 1)

            For I = 1 To 3
                xmeses(I) = 0#
            Next I

            found = formateaa("", 1, 2, 0)
            nlineas

            If Combo1 = "TipoDocumento" Then
                buf = busca_tipo("" & mytablex.Fields("Tipo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Tipo")

            End If

            If Combo1 = "Codigo" Then
                buf = busca_cliente("" & mytablex.Fields("codigo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("Hora")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Hora")

            End If

            If Combo1 = "Vendedor" Then
                buf = busca_vendedor("" & mytablex.Fields("Vendedor"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Zona")

            End If

            If Combo1 = "Cajero" Then
                buf = busca_vendedor("" & mytablex.Fields("Usuario"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Usuario")

            End If

            If Combo1 = "Caja" Then
                buf = busca_caja("" & mytablex.Fields("Caja"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("mesa")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("mesa")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Turno")

            End If

            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Local")

            End If

            If Combo1 = "servicio" Then
                buf = busca_servicio("" & mytablex.Fields("servicio"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("servicio")

            End If

            If Combo1 = "Almacen" Then
                buf = busca_bodega("" & mytablex.Fields("Bodega"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("bodega")

            End If

            nlineas

        End If

        '-----------------aqui se debe sumar los meses-----
        If tipores = "Monto" Then
            If "" & mytablex.Fields("xmes") = "A" Then
                xmeses(1) = xmeses(I) + Val("" & mytablex.Fields("tot"))

                '15/03/2018 Suma Totales Reporte Solo Totales y por Servicio
                'xmeses1(1) = xmeses1(I) + Val("" & mytablex.Fields("tot"))
                xmeses1(1) = xmeses1(1) + Val("" & mytablex.Fields("tot"))
                '15/03/2018 Suma Totales Reporte Solo Totales y por Servicio

            End If

            If "" & mytablex.Fields("xmes") = "C" Then
                xmeses(2) = xmeses(2) + Val("" & mytablex.Fields("tot"))
                xmeses1(2) = xmeses1(2) + Val("" & mytablex.Fields("tot"))

            End If

            If "" & mytablex.Fields("xmes") = "D" Then
                xmeses(3) = xmeses(3) + Val("" & mytablex.Fields("tot"))
                xmeses1(3) = xmeses1(3) + Val("" & mytablex.Fields("tot"))

            End If

        End If

        If tipores = "Cantidad" Then
            If "" & mytablex.Fields("xmes") = "A" Then
                xmeses(1) = xmeses(1) + 1
                xmeses1(1) = xmeses1(1) + 1

            End If

            If "" & mytablex.Fields("xmes") = "C" Then
                xmeses(2) = xmeses(2) + 1
                xmeses1(2) = xmeses1(2) + 1

            End If

            If "" & mytablex.Fields("xmes") = "D" Then
                xmeses(3) = xmeses(3) + 1
                xmeses1(3) = xmeses1(3) + 1

            End If

        End If

        '--------------------------------------------------
        mytablex.MoveNext
    Loop
    sdx1 = 0

    For I = 1 To 3
        buf = Format(xmeses(I), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx1 = sdx1 + xmeses(I)
    Next I

    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 21, 0, 0)
    sdx1 = 0

    For I = 1 To 3
        buf = Format(xmeses1(I), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx1 = sdx1 + xmeses1(I)
    Next I

    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)

End Sub

Sub menu_totales()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    If Combo1 = "Fecha" Then Exit Sub
    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    contpag = 0

    found = sql_codigo(mytablex)

    If found = 0 Then
        mytablex.Close
    
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_totales
    cuerpo_totales mytablex
    '------------------------------------
    Close #1
    cerrar_archivo
    mytablex.Close
     
    'genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    found = valida_wordpad(FileName)

End Sub

Function sql_totales(mytablex As ADODB.Recordset)

    Dim buf As String

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function

    '("" & mysnap.Fields("fecha"))
    If Combo1 = "TipoDocumento" Then
        buf = "select Tipo,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Codigo" Then
        buf = "select Codigo,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Hora" Then
        buf = "select left(hora,2) as Hora,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Vendedor" Then
        buf = "select Vendedor,,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Zona" Then
        buf = "select Zona,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Cajero" Then
        buf = "select Usuario,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Caja" Then
        buf = "select Caja,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Mesa" Then
        buf = "select Mesa,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Turno" Then
        buf = "select Turno,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Local" Then
        buf = "select Local,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Almacen" Then
        buf = "select Bodega,sum(total) as Tot from " & cgusuario & " where "

    End If

    If Combo1 = "Servicio" Then
        buf = "select Servicio,sum(total) as Tot from " & cgusuario & " where "

    End If

    If tipofecha = "E" Then
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    Else
        buf = buf & "  fechae>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fechae<='" & Format(fechaf, "YYYYMMDD") & "' "

    End If

    If local1 <> "%" Then
        buf = buf & " and local='" & extra_loquesea("" & local1) & "'"

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & extra_loquesea("" & tipo) & "'"

    End If

    If serie <> "%" Then
        buf = buf & " and serie like '" & serie & "'"

    End If

    If Numero <> "%" Then
        buf = buf & " and numero like '" & Numero & "'"

    End If

    If codigo <> "%" Then
        buf = buf & " and codigo like '" & codigo & "'"

    End If

    If nombre <> "%" Then
        buf = buf & " and nombre like '" & nombre & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & caja & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & turno & "'"

    End If

    If horai <> "%" And horaf <> "%" Then
        If Val(horaf) >= Val(horai) Then
            buf = buf & " and hour(hora)>=" & Val(horai)
            buf = buf & " and hour(hora)<=" & Val(horaf)

        End If

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & cajero & "'"

    End If

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & vendedor & "'"

    End If

    If transporte <> "%" Then
        buf = buf & " and transporte like '" & transporte & "'"

    End If

    If fpago <> "%" Then
        buf = buf & " and fpago like '" & fpago & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    If acu <> "C" And acu <> "V" Then
        buf = buf & " and acu='" & acu & "'"

    End If

    If acu = "C" Then
        buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P')"

    End If

    If acu = "V" Then
        '19/06/2017 kenyo NOTA DE CREDITO
        ' buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G')"
        buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E')"
        '19/06/2017 kenyo NOTA DE CREDITO
   
    End If

    If servicio <> "%" Then
        buf = buf & " and  servicio='" & extra_loquesea(servicio) & "'"

    End If

    If estado <> "%" Then
        buf = buf & " and estado like '" & estado & "'"

    End If

    If Combo1 = "TipoDocumento" Then
        buf = buf & " group by Tipo order by sum(total)"

    End If

    If Combo1 = "Codigo" Then
        buf = buf & " group by Codigo order by sum(total) "

    End If

    If Combo1 = "Hora" Then
        buf = buf & " group by left(hora,2) order by sum(total)"

    End If

    If Combo1 = "Vendedor" Then
        buf = buf & " group by Vendedor  order by sum(total)"

    End If

    If Combo1 = "Zona" Then
        buf = buf & " group by Zona order by sum(total)"

    End If

    If Combo1 = "Cajero" Then
        buf = buf & " group by Usuario order by sum(total)"

    End If

    If Combo1 = "Caja" Then
        buf = buf & " group by Caja order by sum(total)"

    End If

    If Combo1 = "Mesa" Then
        buf = buf & " group by Mesa order by sum(total)"

    End If

    If Combo1 = "Turno" Then
        buf = buf & " group by Turno order by sum(total)"

    End If

    If Combo1 = "Local" Then
        buf = buf & " group by Local order by sum(total)"

    End If

    If Combo1 = "Almacen" Then
        buf = buf & " group by Bodega order by sum(total)"

    End If

    If Combo1 = "Servicio" Then
        buf = buf & " group by Servicio order by sum(total)"

    End If

    'buf = buf & " order by tot "
    'MsgBox buf
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    sql_totales = 1

End Function

Sub cabecera_totales()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = titulo
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fechai : " & fechai, 25, 2, 0)
    found = formateaa("Fechaf : " & fechaf, 25, 2, 0)
    buf = String(164, "-")
    found = formateaa(buf, 164, 2, 0)
    found = formateaa("Descripcio ", 20, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa("Total", 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    
    buf = String(164, "-")
    found = formateaa(buf, 164, 2, 0)

End Sub

Sub cuerpo_totales(mytablex As ADODB.Recordset)

    Dim Tmp   As String

    Dim sw    As Integer

    Dim buf   As String

    Dim sdx1  As Double

    Dim found As Integer

    Dim tmp1  As String

    ReDim xmeses(13) As Double
    ReDim xmeses1(13) As Double

    Dim I As Integer

    sdx1 = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0
    sw = 0
    tmp1 = ""
    Do

        If mytablex.EOF Then Exit Do
   
        If Combo1 = "TipoDocumento" Then
            tmp1 = "" & mytablex.Fields("Tipo")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("Codigo")

        End If

        If Combo1 = "Hora" Then
            tmp1 = "" & mytablex.Fields("Hora")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("Vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

        End If

        If Combo1 = "Cajero" Then
            tmp1 = "" & mytablex.Fields("usuario")

        End If

        If Combo1 = "Caja" Then
            tmp1 = "" & mytablex.Fields("Caja")

        End If

        If Combo1 = "Mesa" Then
            tmp1 = "" & mytablex.Fields("Mesa")

        End If

        If Combo1 = "Turno" Then
            tmp1 = "" & mytablex.Fields("Turno")

        End If

        If Combo1 = "Local" Then
            tmp1 = "" & mytablex.Fields("Local")

        End If

        If Combo1 = "Almacen" Then
            tmp1 = "" & mytablex.Fields("Bodega")

        End If

        If Combo1 = "Servicio" Then
            tmp1 = "" & mytablex.Fields("Servicio")

        End If

        If sw = 0 Then
            If Combo1 = "TipoDocumento" Then
                buf = busca_tipo("" & mytablex.Fields("Tipo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Tipo")

            End If

            If Combo1 = "Codigo" Then
                buf = busca_cliente("" & mytablex.Fields("codigo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Hora" Then
                buf = busca_cliente("" & mytablex.Fields("Hora"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Hora")

            End If

            If Combo1 = "Vendedor" Then
                buf = busca_vendedor("" & mytablex.Fields("Vendedor"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Zona")

            End If

            If Combo1 = "Cajero" Then
                buf = busca_vendedor("" & mytablex.Fields("Usuario"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Usuario")

            End If

            If Combo1 = "Caja" Then
                buf = busca_caja("" & mytablex.Fields("Caja"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("mesa")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("mesa")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Turno")

            End If

            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Local")

            End If

            If Combo1 = "Almacen" Then
                buf = busca_bodega("" & mytablex.Fields("bodega"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("bodega")

            End If

            If Combo1 = "Servicio" Then
                buf = busca_servicio("" & mytablex.Fields("servicio"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("servicio")

            End If

            nlineas
            sw = 1

        End If

        If Tmp <> tmp1 Then
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            sdx1 = 0

            For I = 1 To 1
                buf = Format(xmeses(I), "0.00")
                found = formateaa(buf, 10, 0, 1)
                found = formateaa("", 1, 0, 0)
                sdx1 = sdx1 + xmeses(I)
            Next I

            buf = Format(sdx1, "0.00")
            found = formateaa(buf, 10, 0, 1)

            For I = 1 To 1
                xmeses(I) = 0#
            Next I

            found = formateaa("", 1, 2, 0)
            nlineas

            If Combo1 = "TipoDocumento" Then
                buf = busca_tipo("" & mytablex.Fields("Tipo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Tipo")

            End If

            If Combo1 = "Codigo" Then
                buf = busca_cliente("" & mytablex.Fields("codigo"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("codigo")

            End If

            If Combo1 = "Hora" Then
                buf = "" & mytablex.Fields("Hora")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Hora")

            End If

            If Combo1 = "Vendedor" Then
                buf = busca_vendedor("" & mytablex.Fields("Vendedor"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("vendedor")

            End If

            If Combo1 = "Zona" Then
                buf = "" & mytablex.Fields("Zona")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Zona")

            End If

            If Combo1 = "Cajero" Then
                buf = busca_vendedor("" & mytablex.Fields("Usuario"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Usuario")

            End If

            If Combo1 = "Caja" Then
                buf = busca_caja("" & mytablex.Fields("Caja"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Caja")

            End If

            If Combo1 = "Mesa" Then
                buf = "" & mytablex.Fields("mesa")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("mesa")

            End If

            If Combo1 = "Turno" Then
                buf = "" & mytablex.Fields("Turno")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Turno")

            End If

            If Combo1 = "Local" Then
                buf = "" & mytablex.Fields("Local")
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("Local")

            End If

            If Combo1 = "servicio" Then
                buf = busca_servicio("" & mytablex.Fields("servicio"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("servicio")

            End If

            If Combo1 = "Almacen" Then
                buf = busca_bodega("" & mytablex.Fields("Bodega"))
                found = formateaa(buf, 20, 0, 0)
                found = formateaa("", 1, 0, 0)
                Tmp = "" & mytablex.Fields("bodega")

            End If

            nlineas

        End If

        '-----------------aqui se debe sumar los meses-----
        If tipores = "Monto" Then
            xmeses(1) = xmeses(I) + Val("" & mytablex.Fields("tot"))

            '15/03/2018 Suma Totales Reporte Solo Totales
            'xmeses1(1) = xmeses1(I) + Val("" & mytablex.Fields("tot"))
            xmeses1(1) = xmeses1(1) + Val("" & mytablex.Fields("tot"))
            '15/03/2018 Suma Totales Reporte Solo Totales

        End If

        If tipores = "Cantidad" Then
            xmeses(1) = xmeses(1) + 1
            xmeses1(1) = xmeses1(1) + 1

        End If

        '--------------------------------------------------
        mytablex.MoveNext
    Loop
    sdx1 = 0

    For I = 1 To 1
        buf = Format(xmeses(I), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx1 = sdx1 + xmeses(I)
    Next I

    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("", 21, 0, 0)
    sdx1 = 0

    For I = 1 To 1
        buf = Format(xmeses1(I), "0.00")
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx1 = sdx1 + xmeses1(I)
    Next I

    buf = Format(sdx1, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)

End Sub

Function busca_vendedorco(mytabley As ADODB.Recordset) As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from vendedor where codigo='" & "" & mytabley.Fields("vendedor") & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If Val("" & mytabley.Fields("total")) >= Val("" & mytablex.Fields("ini1")) And Val("" & mytabley.Fields("total")) <= Val("" & mytablex.Fields("fin1")) Then
            busca_vendedorco = Val("" & mytablex.Fields("por1"))
            GoTo am1

        End If

        If Val("" & mytabley.Fields("total")) >= Val("" & mytablex.Fields("ini2")) And Val("" & mytabley.Fields("total")) <= Val("" & mytablex.Fields("fin2")) Then
            busca_vendedorco = Val("" & mytablex.Fields("por2"))
            GoTo am1

        End If

        If Val("" & mytabley.Fields("total")) >= Val("" & mytablex.Fields("ini3")) And Val("" & mytabley.Fields("total")) <= Val("" & mytablex.Fields("fin3")) Then
            busca_vendedorco = Val("" & mytablex.Fields("por3"))
            GoTo am1

        End If

        If Val("" & mytabley.Fields("total")) >= Val("" & mytablex.Fields("ini4")) And Val("" & mytabley.Fields("total")) <= Val("" & mytablex.Fields("fin4")) Then
            busca_vendedorco = Val("" & mytablex.Fields("por4"))
            GoTo am1

        End If

am1:

    End If

    mytablex.Close

End Function

