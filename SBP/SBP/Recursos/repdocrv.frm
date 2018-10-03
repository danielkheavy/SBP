VERSION 5.00
Begin VB.Form repdocrv 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Documentos"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox SERVICIO 
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
      TabIndex        =   55
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Solo Sunat"
      Height          =   375
      Left            =   6360
      TabIndex        =   54
      Top             =   6000
      Value           =   1  'Checked
      Width           =   1215
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox tipoprint 
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
      TabIndex        =   50
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ComboBox consolidado 
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
      TabIndex        =   48
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox tipoimp 
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
      Top             =   4560
      Width           =   1575
   End
   Begin VB.ComboBox vendedor 
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
      TabIndex        =   45
      Top             =   3720
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
      TabIndex        =   43
      Top             =   4080
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
      TabIndex        =   41
      Top             =   3720
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
      TabIndex        =   39
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "FechaSunat"
      Height          =   375
      Left            =   6360
      TabIndex        =   38
      Top             =   5520
      Width           =   1215
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
      TabIndex        =   36
      Top             =   6360
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
      TabIndex        =   34
      Top             =   1680
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
      TabIndex        =   32
      Text            =   "%"
      Top             =   1800
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
      TabIndex        =   30
      Top             =   1320
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
      TabIndex        =   28
      Top             =   960
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
      TabIndex        =   26
      Top             =   600
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
      TabIndex        =   24
      Top             =   4920
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
      TabIndex        =   22
      Text            =   "%"
      Top             =   4560
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
      TabIndex        =   20
      Text            =   "%"
      Top             =   4080
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
      Top             =   1440
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
      Top             =   960
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
      Top             =   240
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
      Top             =   600
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
      Top             =   2400
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
      Top             =   2760
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
      Top             =   5520
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
      Top             =   5880
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
      Top             =   3360
      Width           =   1575
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
      TabIndex        =   56
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label26 
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
      Left            =   3960
      TabIndex        =   53
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label25 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   51
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label24 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Consolidado Tickets"
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
      TabIndex        =   49
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label23 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Impuesto"
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
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label22 
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
      TabIndex        =   44
      Top             =   4080
      Width           =   2175
   End
   Begin VB.Label Label20 
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
      TabIndex        =   42
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label19 
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
      TabIndex        =   40
      Top             =   3360
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
      TabIndex        =   37
      Top             =   6360
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
      TabIndex        =   35
      Top             =   1680
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
      TabIndex        =   33
      Top             =   1800
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
      TabIndex        =   31
      Top             =   1320
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
      TabIndex        =   29
      Top             =   960
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
      TabIndex        =   27
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bodega"
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
      TabIndex        =   25
      Top             =   4920
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
      TabIndex        =   23
      Top             =   4560
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
      TabIndex        =   21
      Top             =   4080
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
      TabIndex        =   19
      Top             =   3720
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
      Top             =   3360
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
      Top             =   1440
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
      Top             =   960
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
      Top             =   600
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
      Top             =   240
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
      Top             =   2400
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
      Top             =   2760
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
      Top             =   5520
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
      Top             =   5880
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
      Left            =   4200
      TabIndex        =   9
      Top             =   2880
      Width           =   255
   End
   Begin VB.Menu ejui23 
      Caption         =   "&Ejecutar"
   End
   Begin VB.Menu dlo232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "repdocrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bo_inicial As String

Dim bo_final   As String

Dim sw_ant     As Integer

Private Sub dlo232_Click()
    repdocrv.Hide
    Unload repdocrv

End Sub

Private Sub ejui23_Click()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    '14/06/2017 kenyo NOTA DE CREDITO
    'Dim bufnc As String
    '
    'bufnc = "UPDATE FACTURA SET IMPUESTO=IMPUESTO*-1,"
    'bufnc = bufnc & "subtotal=subtotal*-1,gravado=gravado*-1,tivap=tivap*-1, "
    'bufnc = bufnc & "tisc=tisc*-1,total=total*-1,tdetra=tdetra*-1, "
    'bufnc = bufnc & " percepcion=percepcion*-1,servicioco=servicioco*-1 "
    'bufnc = bufnc & "  WHERE TIPO='N' and "
    'bufnc = bufnc & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    '   bufnc = bufnc & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
    'cn.Execute (bufnc)

    '14/06/2017 kenyo NOTA DE CREDITO

    contlin = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    contpag = 0

    If MsgBox("Desea Procesar..", 1, "Aviso") <> 1 Then Exit Sub

    found = sql_documento(mytablex)

    If found = 0 Then
        'mytablex.Close
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1

    '------------------------------------
    If consolidado = "S" Then
        If tipoprint = "Excell" Then
            excel_consolidado mytablex

            Exit Sub

        End If

        cabecera7
        cuerpo_programa7 mytablex

    End If

    If consolidado <> "S" Then
        If tipoprint = "Excell" Then
            cuerpo_excell mytablex
            Exit Sub

        End If

        cabecera_documento
        cuerpo_programa_documento mytablex

    End If

    '------------------------------------
    Close #1
    cerrar_archivo
    
    mytablex.Close
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Private Sub Form_Activate()
    tipoprint.Clear
    tipoprint.AddItem "Normal"
    tipoprint.AddItem "Excell"
    tipoprint.ListIndex = 1

    Dim mytablex As New ADODB.Recordset

    tipo.Clear
    tipo.AddItem "%"

    mytablex.Open "SELECT * FROM tipo", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("grupo") = acu Then
            tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    tipo.ListIndex = 0

End Sub

Private Sub Form_Load()

    Dim mytablex As New ADODB.Recordset

    consolidado.Clear
    consolidado.AddItem "N"
    consolidado.AddItem "S"
    consolidado.ListIndex = 0

    servicio.Clear
    servicio.AddItem "%"
    'servicio.AddItem "Autoservicio"
    'servicio.AddItem "Comanda"
    'servicio.AddItem "Deliveri"

    mytablex.Open "SELECT * FROM servicio ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        servicio.AddItem "" & mytablex.Fields("servicio") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    servicio.ListIndex = 0
    mytablex.Close
    tipoimp.AddItem "%"
    tipoimp.AddItem "GRAVADO"
    tipoimp.AddItem "EXONERADO"
    tipoimp.AddItem "ISC"
    tipoimp.AddItem "IVAP"
    tipoimp.AddItem "PERCEPCION"
    tipoimp.AddItem "SERVICIO"
    tipoimp.ListIndex = 0
    Combo1.AddItem "Fecha"
    Combo1.AddItem "TipoDocumento"
    Combo1.AddItem "Codigo"
    Combo1.AddItem "Vendedor"
    Combo1.AddItem "Zona"
    Combo1.ListIndex = 0

    fechaf = Format(Now, "dd/mm/yyyy")
    fechai = "01" & "/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")

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
    '25/06/2018 Testing Almacen General
    'estado.ListIndex = 0
    estado.ListIndex = 0
    '25/06/2018 Testing Almacen General

    bodega.Clear
    bodega.AddItem "%"
    mytablex.Open "SELECT * FROM bodega", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        bodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
 
    bodega.ListIndex = 0

    vendedor.Clear
    vendedor.AddItem "%"

    cajero.Clear
    cajero.AddItem "%"
    mytablex.Open "SELECT * FROM vendedor", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    cajero.ListIndex = 0
    vendedor.ListIndex = 0

    caja.Clear
    caja.AddItem "%"
    mytablex.Open "SELECT * FROM parameca", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    caja.ListIndex = 0

    turno.Clear
    turno.AddItem "%"
    mytablex.Open "SELECT * FROM turno", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        turno.AddItem "" & mytablex.Fields("turno") & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    turno.ListIndex = 0

    local1.Clear
    local1.AddItem "%"
    mytablex.Open "SELECT * FROM tlocal", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        local1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("Nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    local1.ListIndex = 0

    If local1.ListCount = 2 Then
        local1.ListIndex = 1

    End If

    moneda.Clear
    moneda.AddItem "%"
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0

End Sub

Function sql_documento(mytablex As ADODB.Recordset)

    Dim buf As String

    On Error GoTo cmd89012_err

    If Len(fechai) <> 10 Then Exit Function
    If Len(fechaf) <> 10 Then Exit Function
    If Not IsDate(fechai) Then Exit Function
    If Not IsDate(fechaf) Then Exit Function
    If consolidado = "S" Then
        If Combo1 <> "TipoDocumento" Then
            MsgBox "Grupo debe estar en TipoDocumento", 48, "Aviso"
            Exit Function

        End If

    End If

    buf = "select * from " & cgusuario & " where "

    If Check1.Value = 0 Or Check1.Value = 2 Then
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    End If

    If Check1.Value = 1 Then
        buf = buf & "  fechasunat>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fechasunat<='" & Format(fechaf, "YYYYMMDD") & "' "

    End If

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & extra_loquesea(tipo) & "'"

    End If

    If local1 <> "%" Then
        buf = buf & " and local like '" & extra_loquesea(local1) & "'"

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

    If tipoimp = "GRAVADO" Then
        buf = buf & " and impuesto>0 "
   
    End If

    If tipoimp = "SERVICIO" Then
        buf = buf & " and servicioco>0 "
   
    End If

    If tipoimp = "EXONERADO" Then
        buf = buf & " and gravado>0 "
   
    End If

    If tipoimp = "IVAP" Then
        buf = buf & " and tivap>0 "

    End If

    If tipoimp = "ISC" Then
        buf = buf & " and tisc>0 "

    End If

    If tipoimp = "PERCEPCION" Then
        buf = buf & " and PERCEPCION>0 "

    End If

    'si es registro de compras o venta debe ser convertido
    'If moneda <> "%" Then
    'buf = buf & " and moneda like '" & moneda & "'"
    'End If
    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & extra_loquesea(vendedor) & "'"

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

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & extra_loquesea(cajero) & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & extra_loquesea(caja) & "'"

    End If

    If turno <> "%" Then
        buf = buf & " and turno like '" & extra_loquesea(turno) & "'"

    End If

    'If acu <> "%" Then
    '   buf = buf & " and acu like '" & acu & "'"
    'End If
    If acu = "C" Then
        If Check2.Value = 0 Or Check1.Value = 2 Then
            buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='P' or acu='N' or acu='O')"

        End If

        If Check2.Value = 1 Then
            buf = buf & " and (acu='J' or acu='K' or acu='L' or acu='M' or acu='N' or acu='O')"

        End If
   
    End If

    If acu = "V" Then
        If Check2.Value = 0 Or Check1.Value = 2 Then
            buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D' or acu='G' or acu='E' or acu='F')"

        End If

        If Check2.Value = 1 Then
            buf = buf & " and (acu='1' or acu='A' or acu='B' or acu='C' or acu='D'  or acu='E' or acu='F')"

        End If
   
    End If

    If estado <> "%" Then
        buf = buf & " and estado='" & estado & "'"

    End If

    If Combo1 = "TipoDocumento" Then
        buf = buf & "order by tipo,fecha"
        'mytablex.Open buf, cn, adOpenKeyset, adLockOptimistic
        'sql_documento = 1
        'Exit Function

    End If

    If Combo1 = "Codigo" Then
        buf = buf & "order by Codigo,fecha"

    End If

    If Combo1 = "Vendedor" Then
        buf = buf & "order by vendedor,fecha"

    End If

    If Combo1 = "Zona" Then
        buf = buf & "order by Zona,fecha"

    End If

    If Combo1 = "Fecha" Then
        buf = buf & "order by Fecha,Local,tipo,serie,str(numero)"

    End If

    'MsgBox buf
    mytablex.Open buf, cn, adOpenKeyset, adLockOptimistic
    sql_documento = 1
    Exit Function
cmd89012_err:
    sql_documento = 0
    Exit Function

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
    
    buf = String(183, "-")
    found = formateaa(buf, 183, 2, 0)
    found = formateaa("Fecha", 11, 0, 0)
    found = formateaa("Tipdoc", 7, 0, 0)
    found = formateaa("Serie", 7, 0, 0)
    found = formateaa("Numero", 12, 0, 0)
    found = formateaa("Codigo", 12, 0, 0)
    found = formateaa("Nombre", 31, 0, 0)
    found = formateaa("M", 2, 0, 0)
    found = formateaa("BaseImpo. ", 11, 0, 1)
    found = formateaa("Exonerado", 11, 0, 1)
    found = formateaa("I.S.C ", 11, 0, 1)
    found = formateaa("Impuesto ", 11, 0, 1)
    found = formateaa("Total ", 11, 0, 1)
    found = formateaa("Ivap ", 11, 0, 1)
    found = formateaa("Percepcion ", 11, 0, 1)
    found = formateaa("Servicio ", 11, 0, 1)
    found = formateaa("Detraccion ", 11, 2, 1)

    'found = formateaa("E", 2, 0, 0)
    'found = formateaa("TipoCamb ", 10, 0, 0)
    'found = formateaa("", 1, 2, 0)
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
        buf = "Cant "
        found = formateaa(buf, 8, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "Precio "
        found = formateaa(buf, 10, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "Total "
        found = formateaa(buf, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "E"
        found = formateaa(buf, 1, 0, 0)
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
    
    buf = String(183, "-")
    found = formateaa(buf, 183, 2, 0)

End Sub

Sub cuerpo_programa_documento(mytablex As ADODB.Recordset)

    Dim xparidad As Double

    Dim Tmp      As String

    Dim sw       As Integer

    Dim buf      As String

    Dim found    As Integer

    Dim sdx      As Double

    Dim xdx1     As Double

    Dim xdx2     As Double

    Dim xdx3     As Double

    Dim xdx4     As Double

    Dim tmp1     As String

    sdx = 0
    sw = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0

    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0

    xdx1 = 0
    xdx2 = 0
    xdx3 = 0
    xdx4 = 0
    tmp1 = ""
    Do

        If mytablex.EOF Then Exit Do

        If Combo1 = "TipoDocumento" Then
            tmp1 = "" & mytablex.Fields("tipo")

        End If

        If Combo1 = "Codigo" Then
            tmp1 = "" & mytablex.Fields("codigo")

        End If

        If Combo1 = "Vendedor" Then
            tmp1 = "" & mytablex.Fields("vendedor")

        End If

        If Combo1 = "Zona" Then
            tmp1 = "" & mytablex.Fields("Zona")

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

            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
   
            ssuma1 = 0
            ssuma2 = 0
            ssuma3 = 0
            ssuma4 = 0
            ssuma5 = 0
            ssuma6 = 0
            ssuma7 = 0
            ssuma8 = 0
            ssuma9 = 0

            xdx1 = 0
            xdx2 = 0
            xdx3 = 0
            xdx4 = 0
   
        End If

        If Tmp <> tmp1 Then
            found = formateaa("SubNeto ", 82, 0, 1)
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
   
            buf = Format(suma5, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(suma6, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = Format(suma7, "0.00")
            found = formateaa(buf, 10, 0, 1)
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
   
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0
            suma9 = 0
   
        End If

        buf = "" & mytablex.Fields("fecha")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Tipo")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("serie")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("numero")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("codigo")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("nombre")

        If "" & mytablex.Fields("estado") = "1" Then
            buf = "***ANULADO***"
            found = formateaa(buf, 30, 2, 0)
            nlineas
            GoTo amiga1

        End If

        found = formateaa(buf, 30, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("moneda")
        found = formateaa(buf, 1, 0, 0)
        found = formateaa("", 1, 0, 0)
        xparidad = 1

        If "" & mytablex.Fields("moneda") = "D" Then
            xparidad = busca_paridad("" & mytablex.Fields("fecha"))

        End If

        If xparidad <= 0 Then
            xparidad = 1

        End If

        If "" & mytablex.Fields("moneda") = "D" Then
            sdx = Val("" & mytablex.Fields("subtotal")) * Val(xparidad) - Val("" & mytablex.Fields("gravado")) * Val(xparidad)
            buf = Format(sdx, "0.00")
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("gravado") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("tisc") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("impuesto") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
     
            buf = "" & mytablex.Fields("total") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            buf = "" & mytablex.Fields("tivap") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            buf = "" & mytablex.Fields("percepcion") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            buf = "" & mytablex.Fields("servicioco") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            buf = "" & mytablex.Fields("tdetra") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
   
        End If

        If "" & mytablex.Fields("moneda") = "S" Then
            sdx = Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado"))
            buf = Format(sdx, "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("gravado")
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("tisc")
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("impuesto")
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
      
            buf = "" & mytablex.Fields("total")
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            buf = "" & mytablex.Fields("tivap")
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            buf = "" & mytablex.Fields("percepcion")
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            buf = "" & mytablex.Fields("SERVICIOCO")
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
   
            buf = "" & mytablex.Fields("tdetra")
            buf = Format(Val(buf), "0.00")
            found = formateaa(buf, 10, 0, 1)
            found = formateaa("", 1, 0, 0)
   
        End If

        found = formateaa("", 1, 2, 0)
   
        nlineas

        If "" & mytablex.Fields("moneda") = "S" And "" & mytablex.Fields("estado") = "2" Then
            suma1 = suma1 + Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado"))
            suma2 = suma2 + Val("" & mytablex.Fields("gravado"))
            suma6 = suma6 + Val("" & mytablex.Fields("tivap"))
            ssuma6 = ssuma6 + Val("" & mytablex.Fields("tivap"))
            suma3 = suma3 + Val("" & mytablex.Fields("tisc"))
      
            suma4 = suma4 + Val("" & mytablex.Fields("impuesto"))
            suma5 = suma5 + Val("" & mytablex.Fields("total"))
            suma9 = suma9 + Val("" & mytablex.Fields("tdetra"))
            ssuma9 = ssuma9 + Val("" & mytablex.Fields("tdetra"))
            ssuma1 = ssuma1 + Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado"))
            ssuma2 = ssuma2 + Val("" & mytablex.Fields("gravado"))
            ssuma3 = ssuma3 + Val("" & mytablex.Fields("tisc"))
            ssuma4 = ssuma4 + Val("" & mytablex.Fields("impuesto"))
            ssuma5 = ssuma5 + Val("" & mytablex.Fields("total"))
      
            suma7 = suma7 + Val("" & mytablex.Fields("percepcion"))
            ssuma7 = ssuma7 + Val("" & mytablex.Fields("percepcion"))
      
            suma8 = suma8 + Val("" & mytablex.Fields("servicioco"))
            ssuma8 = ssuma8 + Val("" & mytablex.Fields("servicioco"))
      
        End If

        If "" & mytablex.Fields("moneda") = "D" And "" & mytablex.Fields("estado") = "2" Then
            suma1 = suma1 + Val(Format(Val("" & mytablex.Fields("subtotal")) * Val(xparidad) - Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00"))
            suma2 = suma2 + Val(Format(Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00"))
            suma6 = suma6 + Val(Format(Val("" & mytablex.Fields("tivap")) * Val(xparidad), "0.00"))
            ssuma6 = ssuma6 + Val(Format(Val("" & mytablex.Fields("tivap")) * Val(xparidad), "0.00"))
      
            suma3 = suma3 + Val(Format(Val("" & mytablex.Fields("tisc")) * Val(xparidad), "0.00"))
            ssuma3 = ssuma3 + Val(Format(Val("" & mytablex.Fields("tisc")) * Val(xparidad), "0.00"))
      
            suma4 = suma4 + Val(Format(Val("" & mytablex.Fields("impuesto")) * Val(xparidad), "0.00"))
            suma5 = suma5 + Val(Format(Val("" & mytablex.Fields("total")) * Val(xparidad), "0.00"))
            ssuma1 = ssuma1 + Val(Format(Val("" & mytablex.Fields("subtotal")) * Val(xparidad) - Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00"))
            ssuma2 = ssuma2 + Val(Format(Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00"))
      
            ssuma4 = ssuma4 + Val(Format(Val("" & mytablex.Fields("impuesto")) * Val(xparidad), "0.00"))
            ssuma5 = ssuma5 + Val(Format(Val("" & mytablex.Fields("total")) * Val(xparidad), "0.00"))
      
            suma7 = suma7 + Val(Format(Val("" & mytablex.Fields("percepcion")) * Val(xparidad), "0.00"))
            ssuma7 = ssuma7 + Val(Format(Val("" & mytablex.Fields("percepcion")) * Val(xparidad), "0.00"))
      
            suma8 = suma8 + Val(Format(Val("" & mytablex.Fields("percepcion")) * Val(xparidad), "0.00"))
            ssuma8 = ssuma8 + Val(Format(Val("" & mytablex.Fields("percepcion")) * Val(xparidad), "0.00"))

        End If

        If vdetalle = "S" Then
            ver_detalle mytablex

        End If

amiga1:
        mytablex.MoveNext
    Loop
    found = formateaa("SubNeto ", 82, 0, 1)
    'found = formateaa("", 82, 0, 0)
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
      
    buf = Format(suma5, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
   
    buf = Format(suma6, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
   
    buf = Format(suma7, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma8, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma9, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
    nlineas
    found = formateaa("   Neto ", 82, 0, 1)
    'found = formateaa("", 82, 0, 0)
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
   
    buf = Format(ssuma5, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
   
    buf = Format(ssuma6, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
   
    buf = Format(ssuma7, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
   
    buf = Format(ssuma8, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 0, 0)
   
    buf = Format(ssuma9, "0.00")
    found = formateaa(buf, 10, 0, 1)
    found = formateaa("", 1, 2, 0)
   
End Sub

Sub ver_detalle(mytabley As ADODB.Recordset)

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    Dim found    As Integer

    mytablex.Open "SELECT * FROM " & dgusuariog & " where local='" & "" & mytabley.Fields("local") & "' and tipo='" & "" & mytabley.Fields("tipo") & "' and serie='" & "" & mytabley.Fields("serie") & "' and numero='" & "" & mytabley.Fields("numero") & "'", cn, adOpenKeyset, adLockOptimistic

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
        If "" & mytablex.Fields("codigo") = "" & mytabley.Fields("codigo") And "" & mytablex.Fields("acu") = "" & mytabley.Fields("acu") Then
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
            found = formateaa("", 1, 0, 0)
     
            buf = "" & mytablex.Fields("l1")
            found = formateaa(buf, 3, 0, 0)
            found = formateaa("", 1, 0, 0)
     
            buf = "" & mytablex.Fields("l2")
            found = formateaa(buf, 3, 0, 0)
            found = formateaa("", 1, 0, 0)
     
            buf = "" & mytablex.Fields("l3")
            found = formateaa(buf, 3, 0, 0)
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("l4")
            found = formateaa(buf, 3, 0, 0)
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

Sub nlineas()
    contlin = contlin + 1

    If contlin > Val(nrolineas) Then
        If consolidado = "S" Then
            cabecera7

        End If

        If consolidado <> "S" Then
            cabecera_documento

        End If

    End If

End Sub

Function busca_tipo(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_tipo = "" & mytablex.Fields("descripcio")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_sunat(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_sunat = Format("" & mytablex.Fields("sunat"), "00")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_vendedor(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM vendedor where codigo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_vendedor = "" & mytablex.Fields("nombre")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_zona(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM zona where zona='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_zona = "" & mytablex.Fields("descripcio")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_codigo(buf As String, sw As String) As String

    Dim mytablex As New ADODB.Recordset

    Dim buf1     As String

    buf1 = "CLIENTES"

    If sw = "C" Then
        buf1 = "clientes"

    End If

    If sw = "P" Then
        buf1 = "proveedo"

    End If

    mytablex.Open "SELECT * FROM " & buf1 & "  where codigo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

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

Function busca_paridad(buf As String) As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tcambio where fecha='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        If acu = "C" Then
            busca_paridad = Val("" & mytablex.Fields("compra"))

        End If

        If acu = "V" Then
            busca_paridad = Val("" & mytablex.Fields("venta"))

        End If

    End If

    '------------------------------------- ------------
    mytablex.Close
 
End Function

Sub cuerpo_programa7(mytablex As ADODB.Recordset)

    Dim Tmp      As String

    Dim tmpfecha As String

    Dim tmpcaja  As String

    Dim tmpx     As String

    Dim tmpx1    As String

    Dim buf      As String

    Dim buf1     As String

    Dim sw       As Integer

    Dim found    As Integer

    On Error GoTo cmd22233_err

    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0
    suma7 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0
    ssuma7 = 0
    
    sw = 0
    sw_ant = 0
    Tmp = ""
    tmpcaja = ""
    tmpfecha = ""
    tmpx = ""
    tmpx1 = ""

    Do Until mytablex.EOF

        If "" & mytablex.Fields("tipo") = "1" Or "" & mytablex.Fields("tipo") = "2" Or "" & mytablex.Fields("tipo") = "3" Or "" & mytablex.Fields("tipo") = "4" Or "" & mytablex.Fields("acu") = "E" Then  'E NOTA CREDITO
            tmpx = "" & mytablex.Fields("fecha") & "" & mytablex.Fields("tipo")

            If sw = 0 Then
                Tmp = "" & mytablex.Fields("tipo")
                tmpfecha = "" & mytablex.Fields("fecha")
                tmpcaja = "" & mytablex.Fields("caja")
                tmpx1 = "" & mytablex.Fields("fecha") & "" & mytablex.Fields("tipo")

                If Val("" & mytablex.Fields("tipo")) = 1 Or Val("" & mytablex.Fields("tipo")) = 3 Then
                    bo_inicial = "" & mytablex.Fields("numero")
                    bo_final = "" & mytablex.Fields("numero")

                End If

                sw = 1

            End If

            If tmpx <> tmpx1 Then
                subtotal_rv Tmp, tmpfecha, tmpcaja
                suma1 = 0
                suma2 = 0
                suma3 = 0
                suma4 = 0
                suma5 = 0
                suma6 = 0
                suma7 = 0
                tmpfecha = "" & mytablex.Fields("fecha")
                Tmp = "" & mytablex.Fields("tipo")
                tmpcaja = "" & mytablex.Fields("caja")
                tmpx1 = "" & mytablex.Fields("fecha") & "" & mytablex.Fields("tipo")

                If Val("" & mytablex.Fields("tipo")) = 1 Or Val("" & mytablex.Fields("tipo")) = 3 Then
                    bo_inicial = "" & mytablex.Fields("numero")
                    bo_final = "" & mytablex.Fields("numero")

                End If

            End If

            found = imprime_detalle7(Tmp, tmpfecha, tmpcaja, mytablex)

        End If

        mytablex.MoveNext
    Loop
    
    subtotal_rv Tmp, tmpfecha, tmpcaja
    subtotal_rv1 ssuma2, ssuma3, ssuma4, ssuma5, ssuma6, ssuma7
    Exit Sub
cmd22233_err:
     
    MsgBox "Error en Cuerpo Programa 5 " & error$, 24, "AVISO"
    Exit Sub

End Sub

Function imprime_detalle7(Tmp As String, _
                          tmpfecha As String, _
                          tmpcaja As String, _
                          mytablex As ADODB.Recordset)

    Dim found As Integer

    Dim buf1  As String

    Dim sdx1  As Double

    Dim sdx2  As Double

    Dim sdx3  As Double

    Dim sdx4  As Double

    Dim sdx5  As Double

    Dim sdx6  As Double

    Dim signo As Double

    Dim sw    As Integer

    On Error GoTo cmdhola2

    signo = 1

    If "" & mytablex.Fields("acu") = "E" Or "" & mytablex.Fields("acu") = "N" Then
        signo = -1

    End If

    sdx1 = Val("" & mytablex.Fields("total")) * signo
    sdx2 = Val("" & mytablex.Fields("subtotal")) * signo
    sdx3 = Val("" & mytablex.Fields("impuesto")) * signo
    sdx4 = Val("" & mytablex.Fields("descuento")) * signo
    sdx5 = Val("" & mytablex.Fields("neto")) * signo
    sdx6 = Val("" & mytablex.Fields("gravado")) * signo

    If Val("" & mytablex.Fields("estado")) = 2 Then
        ssuma2 = ssuma2 + sdx1
        ssuma3 = ssuma3 + sdx2
        ssuma4 = ssuma4 + sdx3
        ssuma5 = ssuma5 + sdx4
        ssuma6 = ssuma6 + sdx5
        ssuma7 = ssuma7 + sdx6

    End If

    If Val("" & mytablex.Fields("tipo")) = 1 Or Val("" & mytablex.Fields("tipo")) = 3 Then
        If Val("" & mytablex.Fields("estado")) = 2 Then
            bo_final = "" & mytablex.Fields("numero")

            If sw_ant = 1 Then
                bo_inicial = "" & mytablex.Fields("numero")
                sw_ant = 0

            End If

            suma1 = suma1 + 1
            suma2 = suma2 + sdx1
            suma3 = suma3 + sdx2
            suma4 = suma4 + sdx3
            suma5 = suma5 + sdx4
            suma6 = suma6 + sdx5
            suma7 = suma7 + sdx6

        End If

        If Val("" & mytablex.Fields("estado")) = 1 Then
            subtotal_rv Tmp, tmpfecha, tmpcaja
            imprime_detalle7 = 1
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            sw_ant = 1

        End If

    End If

    '-----------
    If (Val("" & mytablex.Fields("tipo")) <> 1 And Val("" & mytablex.Fields("tipo")) <> 3) Or Val("" & mytablex.Fields("estado")) = 1 Then
        buf1 = "" & mytablex.Fields("fecha") 'fecha
        found = formateaa(buf1, 10, 0, 0)
        found = formateaa("", 1, 0, 0)

        buf1 = "" & mytablex.Fields("tipo") 'tipo
        found = formateaa(buf1, 2, 0, 0)
        found = formateaa("", 1, 0, 0)

        buf1 = "" & mytablex.Fields("caja") 'caja
        found = formateaa(buf1, 2, 0, 0)
        found = formateaa("", 1, 0, 0)
    
        If ("" & mytablex.Fields("acu") = "E" Or Val("" & mytablex.Fields("tipo")) = 2 Or Val("" & mytablex.Fields("tipo")) = 4) And Val("" & mytablex.Fields("estado")) <> 1 Then
            If Val("" & mytablex.Fields("tipo")) = 4 Or "" & mytablex.Fields("acu") = "E" Then
                found = formateaa("", 24, 0, 0)

            End If

            buf1 = "" & mytablex.Fields("numero") 'numero
            found = formateaa(buf1, 11, 0, 0)
            found = formateaa("", 1, 0, 0)

            If Val("" & mytablex.Fields("tipo")) = 2 Then
                found = formateaa("", 24, 0, 0)

            End If

            buf1 = "" & mytablex.Fields("codigo")
            found = formateaa(buf1, 11, 0, 0)
            found = formateaa("", 25, 0, 0)

            buf1 = "" & mytablex.Fields("nombre")
            found = formateaa(buf1, 29, 0, 0)
            found = formateaa("", 1, 0, 0)

            buf1 = Format(sdx1, "0.00")
            found = formateaa(buf1, 9, 0, 1)
            found = formateaa("", 1, 0, 0)

            buf1 = Format(sdx2, "0.00")
            found = formateaa(buf1, 9, 0, 1)
            found = formateaa("", 1, 0, 0)
       
            buf1 = Format(sdx3, "0.00")
            found = formateaa(buf1, 9, 0, 1)
            found = formateaa("", 1, 0, 0)

            buf1 = Format(sdx6, "0.00")
            found = formateaa(buf1, 9, 0, 1)
            found = formateaa("", 1, 0, 0)

            buf1 = Format(sdx4, "0.00")
            found = formateaa(buf1, 9, 0, 1)
            found = formateaa("", 1, 0, 0)

            buf1 = Format(sdx5, "0.00")
            found = formateaa(buf1, 9, 0, 1)
            found = formateaa("", 1, 2, 0)
            'MsgBox "xx"
            nlineas

            If Val("" & mytablex.Fields("estado")) = 2 Then
                suma2 = suma2 + Val("" & mytablex.Fields("total")) * signo
                suma3 = suma3 + Val("" & mytablex.Fields("subtotal")) * signo
                suma4 = suma4 + Val("" & mytablex.Fields("impuesto")) * signo
                suma5 = suma5 + Val("" & mytablex.Fields("descuento")) * signo
                suma6 = suma6 + Val("" & mytablex.Fields("neto")) * signo
                suma7 = suma7 + Val("" & mytablex.Fields("gravado")) * signo

            End If

        End If

        If Val("" & mytablex.Fields("tipo")) = 1 And Val("" & mytablex.Fields("estado")) = 1 Then
            buf1 = "" & mytablex.Fields("numero") 'numero
            found = formateaa(buf1, 11, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("", 37, 0, 1)
            found = formateaa("ANULADO", 30, 2, 1)
            nlineas

        End If

        If (Val("" & mytablex.Fields("tipo")) = 2 Or Val("" & mytablex.Fields("tipo")) = 4) And Val("" & mytablex.Fields("estado")) = 1 Then
            buf1 = "" & mytablex.Fields("numero") 'numero
            found = formateaa(buf1, 11, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("", 37, 0, 1)
            found = formateaa("ANULADO", 30, 2, 1)
            nlineas

        End If

        If Val("" & mytablex.Fields("tipo")) = 3 And Val("" & mytablex.Fields("estado")) = 1 Then
            buf1 = "" & mytablex.Fields("numero") 'numero
            found = formateaa(buf1, 11, 0, 0)
            found = formateaa("", 1, 0, 0)
            found = formateaa("", 37, 0, 1)
            found = formateaa("ANULADO", 30, 2, 1)
            nlineas

        End If

        'nlineas
    End If

    Exit Function
cmdhola2:
    MsgBox "IMPRIME DETALLE-7 " & error$, 24, "AVISO"
    Exit Function

End Function

Sub subtotal_rv(Tmp As String, tmpfecha As String, tmpcaja As String)

    Dim buf1  As String

    Dim found As Integer

    If Val(Tmp) = 1 Or Val(Tmp) = 3 Then
        buf1 = tmpfecha 'fecha
        found = formateaa(buf1, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
                
        buf1 = Tmp
        found = formateaa(buf1, 2, 0, 0)
        found = formateaa("", 1, 0, 0)

        buf1 = tmpcaja 'caja
        found = formateaa(buf1, 2, 0, 0)
        found = formateaa("", 1, 0, 0)

    End If

    If Val(Tmp) = 1 Then
        buf1 = bo_inicial
        found = formateaa(buf1, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf1 = bo_final
        found = formateaa(buf1, 11, 0, 0)
        found = formateaa("", 49, 0, 0)
        found = formateaa("V E N T A  DEL  DIA  ", 30, 0, 0)

        buf1 = Format(suma2, "0.00")
        found = formateaa(buf1, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(suma3, "0.00")
        found = formateaa(buf1, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(suma4, "0.00")
        found = formateaa(buf1, 9, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf1 = Format(suma7, "0.00")
        found = formateaa(buf1, 9, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf1 = Format(suma5, "0.00")
        found = formateaa(buf1, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(suma6, "0.00")
        found = formateaa(buf1, 9, 0, 1)
        found = formateaa("", 1, 2, 0)
        suma2 = 0
        suma3 = 0
        suma4 = 0
        suma5 = 0
        suma6 = 0
        suma7 = 0

        nlineas

    End If

    If Val(Tmp) = 3 Then
        found = formateaa("", 48, 0, 0)
        buf1 = bo_inicial
        found = formateaa(buf1, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf1 = bo_final
        found = formateaa(buf1, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        'found = formateaa("", 49, 0, 0)
        found = formateaa("V E N T A  DEL  DIA  ", 30, 0, 0)
        buf1 = Format(suma2, "0.00")
        found = formateaa(buf1, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(suma3, "0.00")
        found = formateaa(buf1, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(suma4, "0.00")
        found = formateaa(buf1, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(suma7, "0.00")
        found = formateaa(buf1, 9, 0, 1)
        found = formateaa("", 1, 0, 0)

        buf1 = Format(suma5, "0.00")
        found = formateaa(buf1, 9, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf1 = Format(suma6, "0.00")
        found = formateaa(buf1, 9, 0, 1)
        found = formateaa("", 1, 2, 0)
        suma2 = 0
        suma3 = 0
        suma4 = 0
        suma5 = 0
        suma6 = 0
        suma7 = 0

        nlineas

    End If

End Sub

Sub cabecera7()

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
    
    buf = String(190, "-")
    found = formateaa(buf, 190, 2, 0)

    found = formateaa("", 17, 0, 0)
    found = formateaa("REGISTRADORA", 24, 0, 0)
    found = formateaa("MANUAL", 12, 0, 0)
    found = formateaa("", 12, 0, 0)
    found = formateaa("BOLETA DE VENTA", 24, 0, 0)
    found = formateaa("", 31, 0, 0)
    found = formateaa("PRECIO ", 10, 0, 1)
    found = formateaa("VALOR ", 10, 2, 1)

    found = formateaa("FECHA", 11, 0, 0)
    found = formateaa("TI", 3, 0, 0)
    found = formateaa("CA", 3, 0, 0)
    found = formateaa("INICIO", 12, 0, 0)
    found = formateaa("FINAL", 12, 0, 0)
    found = formateaa("FACTURA", 12, 0, 0)
    found = formateaa("RUC", 12, 0, 0)
    found = formateaa("INICIAL", 12, 0, 0)
    found = formateaa("FINAL", 12, 0, 0)
    found = formateaa("CLIENTE", 30, 0, 0)
    found = formateaa("VENTA ", 10, 0, 1)
    found = formateaa("VENTA ", 10, 0, 1)
    found = formateaa("IGV ", 10, 0, 1)
    found = formateaa("INAFEC ", 10, 0, 1)
    found = formateaa("DSCTO ", 10, 0, 1)
    found = formateaa("T.BRUTO ", 10, 2, 1)

    buf = String(190, "-")
    found = formateaa(buf, 190, 2, 0)

End Sub

Sub subtotal_rv1(sdx1 As Double, _
                 sdx2 As Double, _
                 sdx3 As Double, _
                 sdx4 As Double, _
                 sdx5 As Double, _
                 sdx6 As Double)

    Dim found As Integer

    Dim buf1  As String

    found = formateaa("", 119, 0, 0)
    buf1 = Format(sdx1, "0.00")
    found = formateaa(buf1, 9, 0, 1)
    found = formateaa("", 1, 0, 0)

    buf1 = Format(sdx2, "0.00")
    found = formateaa(buf1, 9, 0, 1)
    found = formateaa("", 1, 0, 0)

    buf1 = Format(sdx3, "0.00")
    found = formateaa(buf1, 9, 0, 1)
    found = formateaa("", 1, 0, 0)

    buf1 = Format(sdx6, "0.00")
    found = formateaa(buf1, 9, 0, 1)
    found = formateaa("", 1, 0, 0)

    buf1 = Format(sdx4, "0.00")
    found = formateaa(buf1, 9, 0, 1)
    found = formateaa("", 1, 0, 0)
                
    buf1 = Format(sdx5, "0.00")
    found = formateaa(buf1, 9, 0, 1)
    found = formateaa("", 1, 2, 0)
    nlineas

End Sub

'25/06/2018 Testing Almacen General
'Sub cuerpo_excell(mytablex As ADODB.Recordset)
'Dim Heading(17) As String
'    Dim found As Integer
'    Dim xparidad As Double
'    Dim v As Double
'    Dim h As Double
'    Dim sdx As Double
'    Dim buf As String
'    On Error GoTo cmd5612_err
'
'    Heading(1) = "Fecha"
'    Heading(2) = "Local"
'    Heading(3) = "Tipo"
'    Heading(4) = "Serie"
'    Heading(5) = "Numero"
'    Heading(6) = "Codigo"
'    Heading(7) = "Nombre"
'    Heading(8) = "M"
'    Heading(9) = "BaseImp"
'    Heading(10) = "Exonera"
'    Heading(11) = "Isc"
'    Heading(12) = "Impuesto"
'    Heading(13) = "Total"
'    Heading(14) = "Ivap"
'    Heading(15) = "Percepc"
'    Heading(16) = "Servicio"
'    Heading(17) = "Detraccion"
'
'    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
'    Call Formato_Excelrv(17, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
'
'
'    ''''09/10/2017 kenyo Testing Reportes
'    If acu = "C" Then
'        objExcel.ActiveSheet.Cells(1, 4) = "                                                  REGISTRO DE COMPRAS                                           "
'    End If
'
'    If acu = "V" Then
'        objExcel.ActiveSheet.Cells(1, 4) = "                                                  REGISTRO DE VENTAS                                           "
'    End If
'
'    objExcel.ActiveSheet.Cells(1, 4).Font.bold = True
'    objExcel.ActiveSheet.Cells(1, 4).Font.Size = 14
'    objExcel.ActiveSheet.Cells(1, 4).Font.color = RGB(0, 112, 184)
'    ''''09/10/2017 kenyo Testing Reportes
'
'
'   suma1 = 0
'   suma2 = 0
'   suma3 = 0
'   suma4 = 0
'   suma5 = 0
'   suma6 = 0
'   suma7 = 0
'   suma8 = 0
'   suma9 = 0
'
'   ssuma1 = 0
'   ssuma2 = 0
'   ssuma3 = 0
'   ssuma4 = 0
'   ssuma5 = 0
'   ssuma6 = 0
'   ssuma7 = 0
'   ssuma8 = 0
'   ssuma9 = 0
'
'   v = 2
'    h = 1
'    If acu = "C" Then
'    objExcel.ActiveSheet.Cells(v, h) = "" & "Fecha Inicio:" & fechai & " Fecha Final:" & fechaf
'    End If
'    If acu = "V" Then
'    objExcel.ActiveSheet.Cells(v, h) = "" & "Fecha Inicio:" & fechai & " Fecha Final:" & fechaf
'    End If
'    v = 4
'    h = 1
'
'    'objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("familia")
'    'objExcel.ActiveSheet.Cells(v, h + 1) = busca_familia("" & mytablex.Fields("familia"))
'
'    '----------------------------------------------------
'  Do
'    If mytablex.EOF Then Exit Do
'
'    objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("fecha")
'    objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("local")
'    objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("tipo")
'    objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("serie")
'    objExcel.ActiveSheet.Cells(v, h + 4) = "'" & mytablex.Fields("NUmero")
'    objExcel.ActiveSheet.Cells(v, h + 5) = "'" & mytablex.Fields("Codigo")
'    objExcel.ActiveSheet.Cells(v, h + 6) = "'" & mytablex.Fields("Nombre")
'
'
'   If "" & mytablex.Fields("estado") = "1" Then
'       objExcel.ActiveSheet.Cells(v, h + 7) = "ANULADO"
'       GoTo amiga11
'   End If
'
'   objExcel.ActiveSheet.Cells(v, h + 7) = "'" & mytablex.Fields("moneda")
'   xparidad = 1
'
'   If "" & mytablex.Fields("moneda") = "D" Then
'      xparidad = busca_paridad("" & mytablex.Fields("fecha"))
'   End If
'   If xparidad <= 0 Then
'      xparidad = 1
'   End If
'
'   If "" & mytablex.Fields("moneda") = "D" Then
'   sdx = Val("" & mytablex.Fields("subtotal")) * Val(xparidad) - Val("" & mytablex.Fields("gravado")) * Val(xparidad)
'   buf = Format(sdx, "0.00")
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 8) = "" & buf
'
'   buf = "" & mytablex.Fields("gravado") * Val(xparidad)
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 9) = "" & buf
'
'   buf = "" & mytablex.Fields("tisc") * Val(xparidad)
'   buf = Format(Val(buf), "0.00")
'
'   objExcel.ActiveSheet.Cells(v, h + 10) = "" & buf
'   found = formateaa("", 1, 0, 0)
'   buf = "" & mytablex.Fields("impuesto") * Val(xparidad)
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 11) = "" & buf
'
'
'   buf = "" & mytablex.Fields("total") * Val(xparidad)
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 12) = "" & buf
'
'
'   buf = "" & mytablex.Fields("tivap") * Val(xparidad)
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 13) = "" & buf
'
'   buf = "" & mytablex.Fields("percepcion") * Val(xparidad)
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 14) = "" & buf
'
'   buf = "" & mytablex.Fields("servicioco") * Val(xparidad)
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 15) = "" & buf
'
'   buf = "" & mytablex.Fields("tdetra") * Val(xparidad)
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 16) = "" & buf
'   End If
'
'   If "" & mytablex.Fields("moneda") = "S" Then
'   sdx = Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado"))
'   buf = Format(sdx, "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 8) = "" & buf
'
'   buf = "" & mytablex.Fields("gravado")
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 9) = "" & buf
'
'   buf = "" & mytablex.Fields("tisc")
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 10) = "" & buf
'   buf = "" & mytablex.Fields("impuesto")
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 11) = "" & buf
'
'
'   buf = "" & mytablex.Fields("total")
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 12) = "" & buf
'
'
'   buf = "" & mytablex.Fields("tivap")
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 13) = "" & buf
'
'   buf = "" & mytablex.Fields("percepcion")
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 14) = "" & buf
'
'   buf = "" & mytablex.Fields("SERVICIOCO")
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 15) = "" & buf
'
'   buf = "" & mytablex.Fields("tdetra")
'   buf = Format(Val(buf), "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 16) = "" & buf
'   End If
'
'
'
'   If "" & mytablex.Fields("moneda") = "S" And "" & mytablex.Fields("estado") = "2" Then
'      suma1 = suma1 + Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado"))
'      suma2 = suma2 + Val("" & mytablex.Fields("gravado"))
'      suma6 = suma6 + Val("" & mytablex.Fields("tivap"))
'      ssuma6 = ssuma6 + Val("" & mytablex.Fields("tivap"))
'      suma3 = suma3 + Val("" & mytablex.Fields("tisc"))
'
'      suma4 = suma4 + Val("" & mytablex.Fields("impuesto"))
'
'      suma5 = suma5 + Val("" & mytablex.Fields("total"))
'      ssuma1 = ssuma1 + Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado"))
'      ssuma2 = ssuma2 + Val("" & mytablex.Fields("gravado"))
'      ssuma3 = ssuma3 + Val("" & mytablex.Fields("tisc"))
'      ssuma4 = ssuma4 + Val("" & mytablex.Fields("impuesto"))
'      ssuma5 = ssuma5 + Val("" & mytablex.Fields("total"))
'
'      suma7 = suma7 + Val("" & mytablex.Fields("percepcion"))
'      ssuma7 = ssuma7 + Val("" & mytablex.Fields("percepcion"))
'
'       suma8 = suma8 + Val("" & mytablex.Fields("SERVICIOCO"))
'      ssuma8 = ssuma8 + Val("" & mytablex.Fields("SERVICIOCO"))
'
'      suma9 = suma9 + Val("" & mytablex.Fields("tdetra"))
'      ssuma9 = ssuma9 + Val("" & mytablex.Fields("tdetra"))
'
'   End If
'
'   If "" & mytablex.Fields("moneda") = "D" And "" & mytablex.Fields("estado") = "2" Then
'      suma1 = suma1 + Val(Format(Val("" & mytablex.Fields("subtotal")) * Val(xparidad) - Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00"))
'      suma2 = suma2 + Val(Format(Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00"))
'      suma6 = suma6 + Val(Format(Val("" & mytablex.Fields("tivap")) * Val(xparidad), "0.00"))
'      ssuma6 = ssuma6 + Val(Format(Val("" & mytablex.Fields("tivap")) * Val(xparidad), "0.00"))
'
'      suma3 = suma3 + Val(Format(Val("" & mytablex.Fields("tisc")) * Val(xparidad), "0.00"))
'      ssuma3 = ssuma3 + Val(Format(Val("" & mytablex.Fields("tisc")) * Val(xparidad), "0.00"))
'
'      suma4 = suma4 + Val(Format(Val("" & mytablex.Fields("impuesto")) * Val(xparidad), "0.00"))
'
'
'      suma5 = suma5 + Val(Format(Val("" & mytablex.Fields("total")) * Val(xparidad), "0.00"))
'      ssuma1 = ssuma1 + Val(Format(Val("" & mytablex.Fields("subtotal")) * Val(xparidad) - Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00"))
'      ssuma2 = ssuma2 + Val(Format(Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00"))
'
'      ssuma4 = ssuma4 + Val(Format(Val("" & mytablex.Fields("impuesto")) * Val(xparidad), "0.00"))
'
'
'      ssuma5 = ssuma5 + Val(Format(Val("" & mytablex.Fields("total")) * Val(xparidad), "0.00"))
'
'      suma7 = suma7 + Val(Format(Val("" & mytablex.Fields("percepcion")) * Val(xparidad), "0.00"))
'      ssuma7 = ssuma7 + Val(Format(Val("" & mytablex.Fields("percepcion")) * Val(xparidad), "0.00"))
'       suma8 = suma8 + Val(Format(Val("" & mytablex.Fields("servicioco")) * Val(xparidad), "0.00"))
'      ssuma8 = ssuma8 + Val(Format(Val("" & mytablex.Fields("servicioco")) * Val(xparidad), "0.00"))
'
'      suma9 = suma9 + Val(Format(Val("" & mytablex.Fields("tdetra")) * Val(xparidad), "0.00"))
'      ssuma9 = ssuma9 + Val(Format(Val("" & mytablex.Fields("tdetra")) * Val(xparidad), "0.00"))
'   End If
'amiga11:
'   v = v + 1
'   h = 1
'mytablex.MoveNext
'Loop
'    '-------------------------------------------------
'    h = 1
'
'    ''''09/10/2017 kenyo Testing Reportes
'    objExcel.ActiveSheet.Cells(v, h + 6) = "GRAN TOTAL"
'   Dim k As Integer
'    For k = h + 6 To 17
'      objExcel.ActiveSheet.Cells(v, k).Font.bold = True
'      objExcel.ActiveSheet.Cells(v, k).Interior.color = RGB(248, 243, 53)
'   Next
'''''09/10/2017 kenyo Testing Reportes
'   buf = Format(ssuma1, "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 8) = buf
'
'   buf = Format(ssuma2, "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 9) = buf
'
'   buf = Format(ssuma3, "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 10) = buf
'
'   buf = Format(suma4, "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 11) = buf
'
'
'   buf = Format(ssuma5, "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 12) = buf
'
'
'   buf = Format(ssuma6, "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 13) = buf
'
'
'   buf = Format(ssuma7, "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 14) = buf
'
'   buf = Format(ssuma8, "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 15) = buf
'
'   buf = Format(ssuma9, "0.00")
'   objExcel.ActiveSheet.Cells(v, h + 16) = buf
'
'   v = v + 1
'
'    '----------------------------------------------------
'    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
'    Exit Sub
'cmd5612_err:
'    MsgBox "Aviso en cuerpo excell " + error$, 48, "Aviso"
'    Exit Sub
'
'End Sub

Sub cuerpo_excell(mytablex As ADODB.Recordset)

    Dim Heading(17) As String

    Dim found       As Integer

    Dim xparidad    As Double

    Dim v           As Double

    Dim h           As Double

    Dim sdx         As Double

    Dim buf         As String

    On Error GoTo cmd5612_err
    
    Dim signo As Integer
   
    Heading(1) = "Fecha"
    Heading(2) = "Local"
    Heading(3) = "Tipo"
    Heading(4) = "Serie"
    Heading(5) = "Numero"
    Heading(6) = "Codigo"
    Heading(7) = "Nombre"
    Heading(8) = "M"
    Heading(9) = "BaseImp"
    Heading(10) = "Exonera"
    Heading(11) = "Isc"
    Heading(12) = "Impuesto"
    Heading(13) = "Total"
    Heading(14) = "Ivap"
    Heading(15) = "Percepc"
    Heading(16) = "Servicio"
    Heading(17) = "Detraccion"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excelrv(17, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    ''''09/10/2017 kenyo Testing Reportes
    If acu = "C" Then
        objExcel.ActiveSheet.Cells(1, 4) = "                                                  REGISTRO DE COMPRAS                                           "

    End If
    
    If acu = "V" Then
        objExcel.ActiveSheet.Cells(1, 4) = "                                                  REGISTRO DE VENTAS                                           "

    End If
        
    objExcel.ActiveSheet.Cells(1, 4).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 4).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 4).Font.color = RGB(0, 112, 184)
    ''''09/10/2017 kenyo Testing Reportes
    
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0
    suma7 = 0
    suma8 = 0
    suma9 = 0
   
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0
    ssuma7 = 0
    ssuma8 = 0
    ssuma9 = 0

    signo = 1
    v = 2
    h = 1

    If acu = "C" Then
        objExcel.ActiveSheet.Cells(v, h) = "" & "Fecha Inicio:" & fechai & " Fecha Final:" & fechaf

    End If

    If acu = "V" Then
        objExcel.ActiveSheet.Cells(v, h) = "" & "Fecha Inicio:" & fechai & " Fecha Final:" & fechaf

    End If

    v = 4
    h = 1
    
    Do

        If mytablex.EOF Then Exit Do
    
        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("fecha")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("local")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("tipo")
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("serie")
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & mytablex.Fields("NUmero")
        objExcel.ActiveSheet.Cells(v, h + 5) = "'" & mytablex.Fields("Codigo")
        objExcel.ActiveSheet.Cells(v, h + 6) = "'" & mytablex.Fields("Nombre")
   
        ' 01/08/2018 Testing Facturacion Electronica
        If mytablex.Fields("ACU") = "E" Then ' SI ES NOTA DE CREDITO SERA NEGATIVO
            ' signo = -1
            signo = 1

        End If

        ' 01/08/2018 Testing Facturacion Electronica
   
        If "" & mytablex.Fields("estado") = "1" Then
            objExcel.ActiveSheet.Cells(v, h + 7) = "ANULADO"
            GoTo amiga11

        End If
   
        objExcel.ActiveSheet.Cells(v, h + 7) = "'" & mytablex.Fields("moneda")
        xparidad = 1
   
        If "" & mytablex.Fields("moneda") = "D" Then
            xparidad = busca_paridad("" & mytablex.Fields("fecha"))

        End If

        If xparidad <= 0 Then
            xparidad = 1

        End If
   
        If "" & mytablex.Fields("moneda") = "D" Then
   
            sdx = Val("" & mytablex.Fields("subtotal")) * Val(xparidad) - Val("" & mytablex.Fields("gravado")) * Val(xparidad)
            buf = Format(sdx, "0.00") * signo
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 8) = "" & buf
    
            buf = "" & mytablex.Fields("gravado") * Val(xparidad) * signo
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 9) = "" & buf
     
            buf = "" & mytablex.Fields("tisc") * Val(xparidad) * signo
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 10) = "" & buf
     
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("impuesto") * Val(xparidad) * signo
            buf = Format(Val(buf), "0.00")
     
            objExcel.ActiveSheet.Cells(v, h + 11) = "" & buf
            buf = "" & mytablex.Fields("total") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 12) = "" & buf
     
            buf = "" & mytablex.Fields("tivap") * Val(xparidad) * signo
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 13) = "" & buf
     
            buf = "" & mytablex.Fields("percepcion") * Val(xparidad) * signo
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 14) = "" & buf
     
            buf = "" & mytablex.Fields("servicioco") * Val(xparidad) * signo
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 15) = "" & buf
     
            buf = "" & mytablex.Fields("tdetra") * Val(xparidad) * signo
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 16) = "" & buf

        End If
   
        If "" & mytablex.Fields("moneda") = "S" Then
   
            sdx = Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado"))
            buf = Format(sdx, "0.00") * signo
            objExcel.ActiveSheet.Cells(v, h + 8) = "" & buf
   
            buf = "" & mytablex.Fields("gravado") * signo
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 9) = "" & buf
   
            buf = "" & mytablex.Fields("tisc") * signo
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 10) = "" & buf
   
            buf = "" & mytablex.Fields("impuesto") * signo
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 11) = "" & buf
      
            buf = "" & mytablex.Fields("total") * signo
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 12) = "" & buf
   
            buf = "" & mytablex.Fields("tivap") * signo
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 13) = "" & buf
   
            buf = "" & mytablex.Fields("percepcion") * signo
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 14) = "" & buf
   
            buf = "" & mytablex.Fields("SERVICIOCO") * signo
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 15) = "" & buf
   
            buf = "" & mytablex.Fields("tdetra") * signo
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 16) = "" & buf

        End If
   
        If "" & mytablex.Fields("moneda") = "S" And "" & mytablex.Fields("estado") = "2" Then
            suma1 = suma1 + (Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado"))) * signo
            suma2 = suma2 + Val("" & mytablex.Fields("gravado") * signo)
            suma6 = suma6 + Val("" & mytablex.Fields("tivap") * signo)
            ssuma6 = ssuma6 + Val("" & mytablex.Fields("tivap") * signo)
            suma3 = suma3 + Val("" & mytablex.Fields("tisc") * signo)

            suma4 = suma4 + Val("" & mytablex.Fields("impuesto") * signo)
      
            suma5 = suma5 + Val("" & mytablex.Fields("total") * signo)
            ssuma1 = ssuma1 + (Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado"))) * signo
            ssuma2 = ssuma2 + Val("" & mytablex.Fields("gravado") * signo)
            ssuma3 = ssuma3 + Val("" & mytablex.Fields("tisc") * signo)
            ssuma4 = ssuma4 + Val("" & mytablex.Fields("impuesto") * signo)
            ssuma5 = ssuma5 + Val("" & mytablex.Fields("total") * signo)
      
            suma7 = suma7 + Val("" & mytablex.Fields("percepcion") * signo)
            ssuma7 = ssuma7 + Val("" & mytablex.Fields("percepcion") * signo)
      
            suma8 = suma8 + Val("" & mytablex.Fields("SERVICIOCO") * signo)
            ssuma8 = ssuma8 + Val("" & mytablex.Fields("SERVICIOCO") * signo)
      
            suma9 = suma9 + Val("" & mytablex.Fields("tdetra") * signo)
            ssuma9 = ssuma9 + Val("" & mytablex.Fields("tdetra") * signo)
      
        End If
   
        If "" & mytablex.Fields("moneda") = "D" And "" & mytablex.Fields("estado") = "2" Then
            suma1 = suma1 + (Val(Format(Val("" & mytablex.Fields("subtotal")) * Val(xparidad) - Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00")) * signo)
            suma2 = suma2 + Val(Format(Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00") * signo)
            suma6 = suma6 + Val(Format(Val("" & mytablex.Fields("tivap")) * Val(xparidad), "0.00") * signo)
            ssuma6 = ssuma6 + Val(Format(Val("" & mytablex.Fields("tivap")) * Val(xparidad), "0.00") * signo)
      
            suma3 = suma3 + Val(Format(Val("" & mytablex.Fields("tisc")) * Val(xparidad), "0.00") * signo)
            ssuma3 = ssuma3 + Val(Format(Val("" & mytablex.Fields("tisc")) * Val(xparidad), "0.00") * signo)
      
            suma4 = suma4 + Val(Format(Val("" & mytablex.Fields("impuesto")) * Val(xparidad), "0.00") * signo)
      
            suma5 = suma5 + Val(Format(Val("" & mytablex.Fields("total")) * Val(xparidad), "0.00") * signo)
            ssuma1 = ssuma1 + (Val(Format(Val("" & mytablex.Fields("subtotal")) * Val(xparidad) - Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00")) * signo)
            ssuma2 = ssuma2 + Val(Format(Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00") * signo)
      
            ssuma4 = ssuma4 + Val(Format(Val("" & mytablex.Fields("impuesto")) * Val(xparidad), "0.00") * signo)
      
            ssuma5 = ssuma5 + Val(Format(Val("" & mytablex.Fields("total")) * Val(xparidad), "0.00") * signo)
      
            suma7 = suma7 + Val(Format(Val("" & mytablex.Fields("percepcion")) * Val(xparidad), "0.00") * signo)
            ssuma7 = ssuma7 + Val(Format(Val("" & mytablex.Fields("percepcion")) * Val(xparidad), "0.00") * signo)
            suma8 = suma8 + Val(Format(Val("" & mytablex.Fields("servicioco")) * Val(xparidad), "0.00") * signo)
            ssuma8 = ssuma8 + Val(Format(Val("" & mytablex.Fields("servicioco")) * Val(xparidad), "0.00") * signo)
      
            suma9 = suma9 + Val(Format(Val("" & mytablex.Fields("tdetra")) * Val(xparidad), "0.00") * signo)
            ssuma9 = ssuma9 + Val(Format(Val("" & mytablex.Fields("tdetra")) * Val(xparidad), "0.00") * signo)

        End If

amiga11:
        v = v + 1
        h = 1
        mytablex.MoveNext
    Loop
    '-------------------------------------------------
    h = 1
    
    objExcel.ActiveSheet.Cells(v, h + 6) = "GRAN TOTAL"

    Dim k As Integer

    For k = h + 6 To 17
        objExcel.ActiveSheet.Cells(v, k).Font.bold = True
        objExcel.ActiveSheet.Cells(v, k).Interior.color = RGB(248, 243, 53)
    Next

    buf = Format(ssuma1, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 8) = buf
   
    buf = Format(ssuma2, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 9) = buf
   
    buf = Format(ssuma3, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 10) = buf
   
    buf = Format(suma4, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 11) = buf
  
    buf = Format(ssuma5, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 12) = buf
   
    buf = Format(ssuma6, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 13) = buf
   
    buf = Format(ssuma7, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 14) = buf
   
    buf = Format(ssuma8, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 15) = buf
   
    buf = Format(ssuma9, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 16) = buf
   
    v = v + 1

    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd5612_err:
    MsgBox "Aviso en cuerpo excell " + error$, 48, "Aviso"
    Exit Sub

End Sub

'25/06/2018 Testing Almacen General

''' 09/01/2018 Registro de Ventas para Contasis
Sub cuerpo_excellContasis(mytablex As ADODB.Recordset)

    Dim Heading(42) As String

    Dim found       As Integer

    Dim xparidad    As Double

    Dim v           As Double

    Dim h           As Double

    Dim sdx         As Double

    Dim buf         As String

    On Error GoTo cmd5612_err
   
    Heading(1) = "Fecha de Emision"
    Heading(2) = "Fecha de Vencimiento"
    
    'Documento
    Heading(3) = "Tipo"
    Heading(4) = "Serie"
    Heading(5) = "Numero"
    
    'Cliente
    Heading(6) = "Tipo"
    Heading(7) = "Numero"
    Heading(8) = "Razon Social"
    
    Heading(9) = "VALOR FACTURADO DE LA EXPORTACION"
    Heading(10) = "BASE IMPONIBLE DE LA OPERACION GRAVADA"
    Heading(11) = "IMPORTE TOTAl EXONERADA"
    Heading(12) = "IMPORTE TOTAl INAFECTA"
    Heading(13) = "ISC"
    Heading(14) = "IGV"
    Heading(15) = "OTROS TRIBUTOS "
    Heading(16) = "IMPORTE TOTAL"
    Heading(17) = "TIPO DE CAMBIO"
    
    'nota de credito
    Heading(18) = "FECHA NC"
    Heading(19) = "TIPO NC"
    Heading(20) = "SERIE NC"
    Heading(21) = "NUMERO NC"
    
    Heading(22) = "MONEDA"
    Heading(23) = "EQUIVALENTE EN DOL"
    Heading(24) = "FECHA VENC."
    Heading(25) = "CONDICION DE PAGO"
    
    Heading(26) = "CODIGO CENTRO DE COSTOS"
    Heading(27) = "CODIGO CENTRO DE COSTOS 2"
    
    Heading(28) = "CUENTA CONTABLE BASE IMPONIBLE"
    Heading(29) = "CUENTA CONTABLE OTROS TRIBUTOS Y CARGOS"
    Heading(30) = "CUENTA CONTABLE TOTAL"
    
    Heading(31) = "REGIMEN ESPECIAL"
    Heading(32) = "PORCENTAJE REGIMEN ESPECIAL"
    Heading(33) = "IMPORTE REGIMEN ESPECIAL"
    Heading(34) = "SERIE DOCUMENTO REGIMEN ESPECIAL"
    Heading(35) = "NUMERO DOCUMENTO REGIMEN ESPECIAL"
    Heading(36) = "FECHA DOCUMENTO REGIMEN ESPECIAL"
    
    Heading(37) = "CODIGO PRESUPUESTO"
    Heading(38) = "PORCENTAJE I.G.V."

    Heading(39) = "GLOSA"
    Heading(40) = "MEDIO DE PAGO"
    
    Heading(41) = "CONDICIÓN DE PERCEPCIÓN"
    Heading(42) = "IMPORTE PARA CALCULO RÉGIMEN ESPECIAL"

    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_ExcelrvContasis(42, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    If acu = "C" Then
        objExcel.ActiveSheet.Cells(1, 4) = "                                                  REGISTRO DE COMPRAS                                           "

    End If
    
    If acu = "V" Then
        objExcel.ActiveSheet.Cells(1, 4) = "                                                  REGISTRO DE VENTAS                                           "

    End If
        
    objExcel.ActiveSheet.Cells(1, 4).Font.bold = True
    objExcel.ActiveSheet.Cells(1, 4).Font.Size = 14
    objExcel.ActiveSheet.Cells(1, 4).Font.color = RGB(0, 112, 184)
    
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma6 = 0
    suma7 = 0
    suma8 = 0
    suma9 = 0
   
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0
    ssuma7 = 0
    ssuma8 = 0
    ssuma9 = 0

    v = 2
    h = 1

    If acu = "C" Then
        objExcel.ActiveSheet.Cells(v, h) = "" & "Fecha Inicio:" & fechai & " Fecha Final:" & fechaf

    End If

    If acu = "V" Then
        objExcel.ActiveSheet.Cells(v, h) = "" & "Fecha Inicio:" & fechai & " Fecha Final:" & fechaf

    End If

    v = 4
    h = 1
    'objExcel.ActiveSheet.Cells(v, h) = "" & mytablex.Fields("familia")
    'objExcel.ActiveSheet.Cells(v, h + 1) = busca_familia("" & mytablex.Fields("familia"))
    '----------------------------------------------------
    Do

        If mytablex.EOF Then Exit Do
    
        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("fecha")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("fechae")
    
        'TIPO DE DOCUMENTO
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & E_llenar_TipoDocumento(mytablex.Fields("tipo"))
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("serie")
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & E_llenar_zero(8, mytablex.Fields("numero"))
    
        objExcel.ActiveSheet.Cells(v, h + 5) = "'" & E_llenar_TipoPersona(mytablex.Fields("codigo")) ' TIPO CLINTE
        objExcel.ActiveSheet.Cells(v, h + 6) = "'" & E_llenar_Codigo(mytablex.Fields("codigo"), mytablex.Fields("estado")) 'COD CLIENTE
        objExcel.ActiveSheet.Cells(v, h + 7) = "'" & E_llenar_TipoRazonSocial(mytablex.Fields("codigo"), mytablex.Fields("nombre"), mytablex.Fields("estado"))
       
        xparidad = 1

        If "" & mytablex.Fields("moneda") = "D" Then
            xparidad = busca_paridad("" & mytablex.Fields("fecha"))

        End If

        If xparidad <= 0 Then
            xparidad = 1

        End If
   
        If "" & mytablex.Fields("moneda") = "D" Then
            sdx = Val("" & mytablex.Fields("subtotal")) * Val(xparidad) - Val("" & mytablex.Fields("gravado")) * Val(xparidad)
            buf = Format(sdx, "0.00")
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 8) = "" & buf
  
            buf = "" & mytablex.Fields("gravado") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 9) = "" & buf
   
            buf = "" & mytablex.Fields("tisc") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
   
            objExcel.ActiveSheet.Cells(v, h + 10) = "" & buf
            found = formateaa("", 1, 0, 0)
            buf = "" & mytablex.Fields("impuesto") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 11) = "" & buf
     
            buf = "" & mytablex.Fields("total") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 12) = "" & buf
   
            buf = "" & mytablex.Fields("tivap") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 13) = "" & buf
   
            buf = "" & mytablex.Fields("percepcion") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 14) = "" & buf
   
            buf = "" & mytablex.Fields("servicioco") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 15) = "" & buf
   
            buf = "" & mytablex.Fields("tdetra") * Val(xparidad)
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 16) = "" & buf

        End If
   
        If "" & mytablex.Fields("moneda") = "S" Then
   
            If "" & mytablex.Fields("estado") = "1" Then

                Dim JK As Integer

                For JK = 9 To 15
                    objExcel.ActiveSheet.Cells(v, h + JK) = "0.00"
                Next
                GoTo amiga11

            End If
   
            'VALOR FACTURADO DE LA EXPORTACION
            objExcel.ActiveSheet.Cells(v, h + 8) = ""
         
            'BASE IMPONIBLE DE LA OPERACION GRAVADA
            sdx = Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado"))
            buf = Format(sdx, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 9) = "" & buf
        
            'EXONERADA
            buf = "" & mytablex.Fields("gravado")
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 10) = "" & buf

            'INAFECTA
            objExcel.ActiveSheet.Cells(v, h + 11) = "" & "0.00"

            'ISC
            buf = "" & mytablex.Fields("tisc")
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 12) = "" & buf

            'IMPUESTO
            buf = "" & mytablex.Fields("impuesto")
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 13) = "" & buf

            'OTROS TRIBUTOS Y CARGOS QUE NO FORMAN PARTE DE LA BASE IMPONIBLE
            objExcel.ActiveSheet.Cells(v, h + 14) = "" & "0.00"

            'IMPORTE TOTAL DEL COMPROBANTE DE PAGO
            buf = "" & mytablex.Fields("total")
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 15) = "" & buf

            'TIPO DE CAMBIO
            buf = "" & mytablex.Fields("paridad")
            buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 16) = "" & buf

            'FECHA, TIPO, SERIE Y NUMERO DE COMPROBANTE
            objExcel.ActiveSheet.Cells(v, h + 17) = ""
            objExcel.ActiveSheet.Cells(v, h + 18) = ""
            objExcel.ActiveSheet.Cells(v, h + 19) = ""
            objExcel.ActiveSheet.Cells(v, h + 20) = ""
        
            'MONEDA
            objExcel.ActiveSheet.Cells(v, h + 21) = "" & mytablex.Fields("moneda")
        
            'EQUIVALENTE EN DOLARES AMERICANOS
            'buf = "" & (mytablex.Fields("total") / mytablex.Fields("paridad"))
            'buf = Format(Val(buf), "0.00")
            objExcel.ActiveSheet.Cells(v, h + 22) = ""
        
            'FECHA DE VENCIMIENTO
            objExcel.ActiveSheet.Cells(v, h + 23) = "'" & mytablex.Fields("fechae")
        
            'condicion contado / credito
            objExcel.ActiveSheet.Cells(v, h + 24) = "'" & busca_CondicionPago(mytablex.Fields("local"), mytablex.Fields("tipo"), mytablex.Fields("serie"), mytablex.Fields("numero"))

            'CODIGO CENTRO DE COSTOS
            objExcel.ActiveSheet.Cells(v, h + 25) = ""
        
            'CODIGO CENTRO DE COSTOS 2
            objExcel.ActiveSheet.Cells(v, h + 26) = ""
        
            'CUENTA CONTABLE BASE IMPONIBLE ' SubTotal
            objExcel.ActiveSheet.Cells(v, h + 27) = "'" & busca_CuentasContables(mytablex.Fields("tipo"), 1)
        
            'CUENTA CONTABLE OTROS TRIBUTOS Y CARGOS ' Impuesto
            objExcel.ActiveSheet.Cells(v, h + 28) = "'" & busca_CuentasContables(mytablex.Fields("tipo"), 2)
      
            'CUENTA CONTABLE TOTAL  ' Total
            objExcel.ActiveSheet.Cells(v, h + 29) = "'" & busca_CuentasContables(mytablex.Fields("tipo"), 3)
        
            'REGIMEN ESPECIAL
            objExcel.ActiveSheet.Cells(v, h + 30) = ""
        
            'PORCENTAJE REGIMEN ESPECIAL
            objExcel.ActiveSheet.Cells(v, h + 31) = ""
   
            'IMPORTE REGIMEN ESPECIAL
            objExcel.ActiveSheet.Cells(v, h + 32) = ""
      
            'SERIE DOCUMENTO REGIMEN ESPECIAL
            objExcel.ActiveSheet.Cells(v, h + 33) = ""
        
            'NUMERO DOCUMENTO REGIMEN ESPECIAL
            objExcel.ActiveSheet.Cells(v, h + 34) = ""
        
            'FECHA DOCUMENTO REGIMEN ESPECIAL
            objExcel.ActiveSheet.Cells(v, h + 35) = ""
    
            'CODIGO PRESUPUESTO
            objExcel.ActiveSheet.Cells(v, h + 36) = ""
      
            'PORCENTAJE I.G.V.
            objExcel.ActiveSheet.Cells(v, h + 37) = ""
                
            'GLOSA
            objExcel.ActiveSheet.Cells(v, h + 38) = "'" & mytablex.Fields("observa")
        
            'MEDIO DE PAGO
            objExcel.ActiveSheet.Cells(v, h + 39) = ""
      
            'CONDICIÓN DE PERCEPCIÓN
            objExcel.ActiveSheet.Cells(v, h + 40) = ""
        
            'IMPORTE PARA CALCULO RÉGIMEN ESPECIAL
            objExcel.ActiveSheet.Cells(v, h + 41) = ""

            '        buf = "" & mytablex.Fields("tivap")
            '        buf = Format(Val(buf), "0.00")
            '        objExcel.ActiveSheet.Cells(v, h + 13) = "" & buf
            '
            '        buf = "" & mytablex.Fields("percepcion")
            '        buf = Format(Val(buf), "0.00")
            '        objExcel.ActiveSheet.Cells(v, h + 14) = "" & buf
            '
            '        buf = "" & mytablex.Fields("SERVICIOCO")
            '        buf = Format(Val(buf), "0.00")
            '        objExcel.ActiveSheet.Cells(v, h + 15) = "" & buf
            '
            '        buf = "" & mytablex.Fields("tdetra")
            '        buf = Format(Val(buf), "0.00")
            '        objExcel.ActiveSheet.Cells(v, h + 16) = "" & buf
   
        End If
   
        If "" & mytablex.Fields("moneda") = "S" And "" & mytablex.Fields("estado") = "2" Then
            suma1 = suma1 + Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado"))
            suma2 = suma2 + Val("" & mytablex.Fields("gravado"))
            suma6 = suma6 + Val("" & mytablex.Fields("tivap"))
            ssuma6 = ssuma6 + Val("" & mytablex.Fields("tivap"))
            suma3 = suma3 + Val("" & mytablex.Fields("tisc"))

            suma4 = suma4 + Val("" & mytablex.Fields("impuesto"))
      
            suma5 = suma5 + Val("" & mytablex.Fields("total"))
            ssuma1 = ssuma1 + Val("" & mytablex.Fields("subtotal")) - Val("" & mytablex.Fields("gravado"))
            ssuma2 = ssuma2 + Val("" & mytablex.Fields("gravado"))
            ssuma3 = ssuma3 + Val("" & mytablex.Fields("tisc"))
            ssuma4 = ssuma4 + Val("" & mytablex.Fields("impuesto"))
            ssuma5 = ssuma5 + Val("" & mytablex.Fields("total"))
      
            suma7 = suma7 + Val("" & mytablex.Fields("percepcion"))
            ssuma7 = ssuma7 + Val("" & mytablex.Fields("percepcion"))
      
            suma8 = suma8 + Val("" & mytablex.Fields("SERVICIOCO"))
            ssuma8 = ssuma8 + Val("" & mytablex.Fields("SERVICIOCO"))
      
            suma9 = suma9 + Val("" & mytablex.Fields("tdetra"))
            ssuma9 = ssuma9 + Val("" & mytablex.Fields("tdetra"))
      
        End If

        If "" & mytablex.Fields("moneda") = "D" And "" & mytablex.Fields("estado") = "2" Then
            suma1 = suma1 + Val(Format(Val("" & mytablex.Fields("subtotal")) * Val(xparidad) - Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00"))
            suma2 = suma2 + Val(Format(Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00"))
            suma6 = suma6 + Val(Format(Val("" & mytablex.Fields("tivap")) * Val(xparidad), "0.00"))
            ssuma6 = ssuma6 + Val(Format(Val("" & mytablex.Fields("tivap")) * Val(xparidad), "0.00"))
      
            suma3 = suma3 + Val(Format(Val("" & mytablex.Fields("tisc")) * Val(xparidad), "0.00"))
            ssuma3 = ssuma3 + Val(Format(Val("" & mytablex.Fields("tisc")) * Val(xparidad), "0.00"))
      
            suma4 = suma4 + Val(Format(Val("" & mytablex.Fields("impuesto")) * Val(xparidad), "0.00"))
      
            suma5 = suma5 + Val(Format(Val("" & mytablex.Fields("total")) * Val(xparidad), "0.00"))
            ssuma1 = ssuma1 + Val(Format(Val("" & mytablex.Fields("subtotal")) * Val(xparidad) - Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00"))
            ssuma2 = ssuma2 + Val(Format(Val("" & mytablex.Fields("gravado")) * Val(xparidad), "0.00"))
      
            ssuma4 = ssuma4 + Val(Format(Val("" & mytablex.Fields("impuesto")) * Val(xparidad), "0.00"))
      
            ssuma5 = ssuma5 + Val(Format(Val("" & mytablex.Fields("total")) * Val(xparidad), "0.00"))
      
            suma7 = suma7 + Val(Format(Val("" & mytablex.Fields("percepcion")) * Val(xparidad), "0.00"))
            ssuma7 = ssuma7 + Val(Format(Val("" & mytablex.Fields("percepcion")) * Val(xparidad), "0.00"))
            suma8 = suma8 + Val(Format(Val("" & mytablex.Fields("servicioco")) * Val(xparidad), "0.00"))
            ssuma8 = ssuma8 + Val(Format(Val("" & mytablex.Fields("servicioco")) * Val(xparidad), "0.00"))
      
            suma9 = suma9 + Val(Format(Val("" & mytablex.Fields("tdetra")) * Val(xparidad), "0.00"))
            ssuma9 = ssuma9 + Val(Format(Val("" & mytablex.Fields("tdetra")) * Val(xparidad), "0.00"))

        End If

amiga11:
        v = v + 1
        h = 1
        mytablex.MoveNext
    Loop
    '-------------------------------------------------
    h = 1
  
    objExcel.ActiveSheet.Cells(v, h + 6) = "GRAN TOTAL"

    Dim k As Integer

    For k = h + 6 To 16
        objExcel.ActiveSheet.Cells(v, k).Font.bold = True
        objExcel.ActiveSheet.Cells(v, k).Interior.color = RGB(248, 243, 53)
    Next

    buf = Format(ssuma1, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 9) = buf
   
    buf = Format(ssuma2, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 10) = buf
   
    buf = Format(ssuma3, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 11) = buf
   
    buf = Format(suma4, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 13) = buf
  
    buf = Format(ssuma5, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 15) = buf
   
    '
    '   buf = Format(ssuma6, "0.00")
    '   objExcel.ActiveSheet.Cells(v, h + 14) = buf
    '
    '
    '   buf = Format(ssuma7, "0.00")
    '   objExcel.ActiveSheet.Cells(v, h + 15) = buf
    '
    '   buf = Format(ssuma8, "0.00")
    '   objExcel.ActiveSheet.Cells(v, h + 16) = buf
    '
    '   buf = Format(ssuma9, "0.00")
    '   objExcel.ActiveSheet.Cells(v, h + 17) = buf
   
    v = v + 1

    '----------------------------------------------------
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd5612_err:
    MsgBox "Aviso en cuerpo excell " + error$, 48, "Aviso"
    Exit Sub

End Sub

Public Function Formato_ExcelrvContasis(Num_Campos As Integer, _
                                        Nombre_Campos() As String) As Boolean

    Dim I As Integer

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Font.Size = 8
        
        ''''09/10/2017 kenyo Testing Reportes
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Interior.color = RGB(192, 192, 250)
        ''''09/10/2017 kenyo Testing Reportes
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
       
        .columns("A").ColumnWidth = 12
        .columns("B").ColumnWidth = 12
        
        .columns("C").ColumnWidth = 4
        
        .columns("D").ColumnWidth = 12
        .columns("E").ColumnWidth = 12
        .columns("F").ColumnWidth = 4
        .columns("G").ColumnWidth = 14
        .columns("H").ColumnWidth = 30
        
        .columns("I").ColumnWidth = 12
        .columns("J").ColumnWidth = 12
        .columns("K").ColumnWidth = 12
        .columns("L").ColumnWidth = 12
        .columns("M").ColumnWidth = 12
        .columns("N").ColumnWidth = 12
        .columns("O").ColumnWidth = 12
        .columns("P").ColumnWidth = 12
        .columns("Q").ColumnWidth = 12
        
        .columns("R").ColumnWidth = 15
        .columns("S").ColumnWidth = 12
        .columns("T").ColumnWidth = 12
        .columns("U").ColumnWidth = 12
        .columns("V").ColumnWidth = 12
        .columns("W").ColumnWidth = 12
        .columns("X").ColumnWidth = 12
        .columns("Y").ColumnWidth = 12
        .columns("Z").ColumnWidth = 12
   
        .columns("AA").ColumnWidth = 12
        .columns("AB").ColumnWidth = 12
        .columns("AC").ColumnWidth = 12
        .columns("AD").ColumnWidth = 12
        .columns("AE").ColumnWidth = 12
        .columns("AF").ColumnWidth = 12
        .columns("AG").ColumnWidth = 12
        .columns("AH").ColumnWidth = 12
        
        .columns("AI").ColumnWidth = 12
        .columns("AJ").ColumnWidth = 12
        .columns("AK").ColumnWidth = 12
        .columns("AL").ColumnWidth = 12
        .columns("AM").ColumnWidth = 18 ' GLOSA
        .columns("AN").ColumnWidth = 12
        .columns("AO").ColumnWidth = 12
        .columns("AP").ColumnWidth = 12
        
    End With

End Function

Function E_llenar_zero(hastaCuanto As Integer, myDato As String) As String

    Dim I   As Integer

    Dim max As Integer

    max = Len(myDato)

    For I = 1 To hastaCuanto - max
        myDato = "0" & myDato
    Next
    E_llenar_zero = myDato

End Function

Function E_llenar_TipoDocumento(myDato As String) As String
   
    '07/08/2018 Nota de credito final
    If myDato = 1 Then 'Boleta
        E_llenar_TipoDocumento = "03"
    ElseIf myDato = 2 Then 'Factura
        E_llenar_TipoDocumento = "01"
    ElseIf myDato = "71" Or myDato = "72" Then  'NC
        E_llenar_TipoDocumento = "07"
    ElseIf myDato = "81" Or myDato = "82" Then  'ND
        E_llenar_TipoDocumento = "08"
    Else ' Otros
        E_llenar_TipoDocumento = "00"

    End If
    
    '07/08/2018 Nota de credito final
    
End Function

Function E_llenar_TipoPersona(myDato As String) As String

    If Len(Trim(myDato)) = 8 Then  'Dni
        E_llenar_TipoPersona = "1"
    ElseIf Len(Trim(myDato)) = 11 Then  'RUC
        E_llenar_TipoPersona = "6"
    Else ' Otros
        E_llenar_TipoPersona = "0"

    End If

End Function

Function E_llenar_Codigo(myDato As String, myEstado As String) As String
    
    If myEstado = "1" Then
        E_llenar_Codigo = "99999999999"
    Else

        If E_llenar_TipoPersona(myDato) = 0 Then   'RUC
            E_llenar_Codigo = "00000000"
        Else
            E_llenar_Codigo = myDato

        End If

    End If
    
End Function

Function E_llenar_TipoRazonSocial(codigo As String, _
                                  nombre As String, _
                                  myEstado As String) As String
    
    If myEstado = "1" Then
        E_llenar_TipoRazonSocial = "VENTA ANULADA"
    Else

        If Len(Trim(codigo)) = 8 Then 'Dni
            E_llenar_TipoRazonSocial = nombre
        ElseIf Len(Trim(codigo)) = 11 Then  'RUC

            If Mid(codigo, 1, 2) = "10" Then
                E_llenar_TipoRazonSocial = busca_nombre_comas("" & nombre)
            Else
                E_llenar_TipoRazonSocial = nombre

            End If

        End If

    End If

End Function

Function busca_nombre_comas(buf As String) As String
    buf = Replace$(buf, " ", ",")
    busca_nombre_comas = buf

End Function

Function busca_CuentasContables(tipodoc As String, tipocuenta As String) As String

    Dim mytabley As New ADODB.Recordset

    mytabley.Open "SELECT cuenta1,cuenta2,cuenta3  FROM tipo where tipo='" & "" & tipodoc & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount > 0 Then
   
        If tipocuenta = "1" Then        'SubTotal
            busca_CuentasContables = "" & mytabley.Fields("cuenta1")
        ElseIf tipocuenta = "2" Then    'Impuesto
            busca_CuentasContables = "" & mytabley.Fields("cuenta2")
        ElseIf tipocuenta = "3" Then    'Total
            busca_CuentasContables = "" & mytabley.Fields("cuenta3")

        End If

    End If

    '------------------------------------- ------------
    mytabley.Close
 
End Function

Function busca_CondicionPago(locall As String, _
                             tipo As String, _
                             serie As String, _
                             Numero As String) As String

    Dim mytabley As New ADODB.Recordset

    mytabley.Open "SELECT top 1 fpago FROM fpagov where local='" & "" & locall & "' and tipo='" & "" & tipo & "' and serie='" & "" & serie & "' and numero='" & "" & Numero & "' order by fpago", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        If ("" & mytabley.Fields("fpago") = "3") Then
            busca_CondicionPago = "CRE"
        Else
            busca_CondicionPago = "CON"

        End If

    End If

    '------------------------------------- ------------
    mytabley.Close
 
End Function

''' 09/01/2018 Registro de Ventas para Contasis

Public Function Formato_Excelrv(Num_Campos As Integer, _
                                Nombre_Campos() As String) As Boolean

    Dim I As Integer

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 17)).Font.bold = True
        
        ''''09/10/2017 kenyo Testing Reportes
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Interior.color = RGB(192, 192, 250)
        ''''09/10/2017 kenyo Testing Reportes
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
       
        .columns("A").ColumnWidth = 10
        .columns("B").ColumnWidth = 6
        .columns("C").ColumnWidth = 4
        .columns("D").ColumnWidth = 8
        .columns("E").ColumnWidth = 9
        .columns("F").ColumnWidth = 11
        .columns("G").ColumnWidth = 30
        .columns("H").ColumnWidth = 2
        
        .columns("I").ColumnWidth = 10
        .columns("J").ColumnWidth = 10
        .columns("K").ColumnWidth = 10
        .columns("L").ColumnWidth = 10
        .columns("M").ColumnWidth = 10
        .columns("N").ColumnWidth = 10
        .columns("O").ColumnWidth = 10
        .columns("P").ColumnWidth = 10
        .columns("Q").ColumnWidth = 10
    
    End With

End Function

Sub excel_consolidado(mytablex As ADODB.Recordset)

    Dim Heading(17) As String

    Dim found       As Integer

    Dim xparidad    As Double

    Dim v           As Double

    Dim h           As Double

    Dim sdx         As Double

    Dim buf         As String

    Dim tmpx        As String

    Dim tmpx1       As String

    Dim sw          As Integer

    Dim tmpfecha    As String

    Dim tmpcaja     As String

    Dim Tmp         As String

    Dim tmpserie    As String

    On Error GoTo cmd5612_err
    
    'Heading(1) = ""
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excelrve(17, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    v = 5
    h = 1
    
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    suma5 = 0
    suma8 = 0
    suma9 = 0
   
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0
    ssuma5 = 0
    ssuma6 = 0
    ssuma7 = 0
    ssuma8 = 0
    ssuma9 = 0
    h = 1

    If acu = "C" Then
        objExcel.ActiveSheet.Cells(2, h) = "REGISTRO COMPRAS " & "Fecha Inicio:" & fechai & " Fecha Inicio:" & fechai

    End If

    If acu = "V" Then
        objExcel.ActiveSheet.Cells(2, h) = "REGISTRO VENTAS " & "Fecha Inicio:" & fechai & " Fecha Inicio:" & fechai

    End If

    v = 5
    h = 1
    sw = 0
    tmpx = ""
    tmpx1 = ""
    '----------------------------------------------------
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("tipo") = "1" Or "" & mytablex.Fields("tipo") = "2" Or "" & mytablex.Fields("tipo") = "3" Or "" & mytablex.Fields("tipo") = "4" Or "" & mytablex.Fields("acu") = "E" Then  'E NOTA CREDITO
            tmpx = "" & mytablex.Fields("fecha") & "" & mytablex.Fields("tipo") & "" & mytablex.Fields("serie")

            If sw = 0 Then
                Tmp = "" & mytablex.Fields("tipo")
                tmpfecha = "" & mytablex.Fields("fecha")
                tmpcaja = "" & mytablex.Fields("caja")
                tmpserie = "" & mytablex.Fields("serie")
                tmpx1 = "" & mytablex.Fields("fecha") & "" & mytablex.Fields("tipo") & "" & mytablex.Fields("serie")

                If Val("" & mytablex.Fields("tipo")) = 1 Or Val("" & mytablex.Fields("tipo")) = 3 Then
                    bo_inicial = "" & mytablex.Fields("numero")
                    bo_final = "" & mytablex.Fields("numero")

                End If

                sw = 1

            End If

            If tmpx <> tmpx1 Then
                subtotal_rve Tmp, tmpfecha, tmpcaja, v, h, tmpserie
                suma1 = 0
                suma2 = 0
                suma3 = 0
                suma4 = 0
                suma5 = 0
                suma6 = 0
                suma7 = 0
                suma8 = 0
                suma9 = 0
                tmpfecha = "" & mytablex.Fields("fecha")
                Tmp = "" & mytablex.Fields("tipo")
                tmpcaja = "" & mytablex.Fields("caja")
                tmpserie = "" & mytablex.Fields("serie")
                tmpx1 = "" & mytablex.Fields("fecha") & "" & mytablex.Fields("tipo") & "" & mytablex.Fields("serie")

                If Val("" & mytablex.Fields("tipo")) = 1 Or Val("" & mytablex.Fields("tipo")) = 3 Then
                    bo_inicial = "" & mytablex.Fields("numero")
                    bo_final = "" & mytablex.Fields("numero")

                End If

            End If

            h = 0
            found = imprime_detalle7e(Tmp, tmpfecha, tmpcaja, mytablex, v, h, tmpserie)

        End If

        mytablex.MoveNext
    Loop
    '----------------------------------------------------
    v = v + 1
    objExcel.ActiveSheet.Cells(v, h + 12) = Format(ssuma2, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 13) = Format(ssuma3, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 14) = Format(ssuma4, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 15) = Format(ssuma7, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 16) = Format(ssuma5, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 17) = Format(ssuma6, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 18) = Format(ssuma8, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 19) = Format(ssuma9, "0.00")
    
    '----------------------------------------------------
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd5612_err:
    MsgBox "Aviso en cuerpo excell " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function imprime_detalle7e(Tmp As String, _
                           tmpfecha As String, _
                           tmpcaja As String, _
                           mytablex As ADODB.Recordset, _
                           v As Double, _
                           h As Double, _
                           tmpserie As String)

    Dim found As Integer

    Dim buf1  As String

    Dim sdx1  As Double

    Dim sdx2  As Double

    Dim sdx3  As Double

    Dim sdx4  As Double

    Dim sdx5  As Double

    Dim sdx6  As Double

    Dim sdx8  As Double

    Dim sdx9  As Double

    Dim signo As Double

    Dim sw    As Integer

    On Error GoTo cmdhola22

    signo = 1

    If "" & mytablex.Fields("acu") = "E" Or "" & mytablex.Fields("acu") = "N" Then
        signo = -1

    End If

    sdx1 = Val("" & mytablex.Fields("total")) * signo
    sdx2 = Val("" & mytablex.Fields("subtotal")) * signo
    sdx3 = Val("" & mytablex.Fields("impuesto")) * signo
    sdx4 = Val("" & mytablex.Fields("descuento")) * signo
    sdx5 = Val("" & mytablex.Fields("neto")) * signo
    sdx6 = Val("" & mytablex.Fields("gravado")) * signo

    sdx8 = Val("" & mytablex.Fields("percepcion")) * signo
    sdx9 = Val("" & mytablex.Fields("servicioco")) * signo

    If Val("" & mytablex.Fields("estado")) = 2 Then
        ssuma2 = ssuma2 + sdx1
        ssuma3 = ssuma3 + sdx2
        ssuma4 = ssuma4 + sdx3
        ssuma5 = ssuma5 + sdx4
        ssuma6 = ssuma6 + sdx5
        ssuma7 = ssuma7 + sdx6
      
        ssuma8 = ssuma8 + sdx8
        ssuma9 = ssuma9 + sdx9

    End If

    If Val("" & mytablex.Fields("tipo")) = 1 Or Val("" & mytablex.Fields("tipo")) = 3 Then
        If Val("" & mytablex.Fields("estado")) = 2 Then
            bo_final = "" & mytablex.Fields("numero")

            If sw_ant = 1 Then
                bo_inicial = "" & mytablex.Fields("numero")
                sw_ant = 0

            End If

            suma1 = suma1 + 1
            suma2 = suma2 + sdx1
            suma3 = suma3 + sdx2
            suma4 = suma4 + sdx3
            suma5 = suma5 + sdx4
            suma6 = suma6 + sdx5
            suma7 = suma7 + sdx6
      
            suma8 = suma8 + sdx8
            suma9 = suma9 + sdx9

        End If

        If Val("" & mytablex.Fields("estado")) = 1 Then
            subtotal_rve Tmp, tmpfecha, tmpcaja, v, h, tmpserie
            imprime_detalle7e = 1
            suma1 = 0
            suma2 = 0
            suma3 = 0
            suma4 = 0
            suma5 = 0
            suma6 = 0
            suma7 = 0
            suma8 = 0
            suma9 = 0
            sw_ant = 1

        End If

    End If

    '-----------
    If (Val("" & mytablex.Fields("tipo")) <> 1 And Val("" & mytablex.Fields("tipo")) <> 3) Or Val("" & mytablex.Fields("estado")) = 1 Then

        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("fecha")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & busca_sunat("" & mytablex.Fields("tipo"))
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("caja")
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & mytablex.Fields("serie")
    
        If ("" & mytablex.Fields("acu") = "E" Or Val("" & mytablex.Fields("tipo")) = 2 Or Val("" & mytablex.Fields("tipo")) = 4) And Val("" & mytablex.Fields("estado")) <> 1 Then
            If Val("" & mytablex.Fields("tipo")) = 4 Or "" & mytablex.Fields("acu") = "E" Then

                'objExcel.ActiveSheet.Cells(v, h + 4) = "'"
            End If

            objExcel.ActiveSheet.Cells(v, h + 7) = "'" & mytablex.Fields("numero")

            If Val("" & mytablex.Fields("tipo")) = 2 Then

                'objExcel.ActiveSheet.Cells(v, h + 6) = "'"
            End If

            objExcel.ActiveSheet.Cells(v, h + 8) = "'" & mytablex.Fields("codigo")
            objExcel.ActiveSheet.Cells(v, h + 11) = "'" & mytablex.Fields("nombre")
            objExcel.ActiveSheet.Cells(v, h + 12) = Format(sdx1, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 13) = Format(sdx2, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 14) = Format(sdx3, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 15) = Format(sdx6, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 16) = Format(sdx4, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 17) = Format(sdx5, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 18) = Format(sdx8, "0.00")
            objExcel.ActiveSheet.Cells(v, h + 19) = Format(sdx9, "0.00")
            v = v + 1
       
            If Val("" & mytablex.Fields("estado")) = 2 Then
                suma2 = suma2 + Val("" & mytablex.Fields("total")) * signo
                suma3 = suma3 + Val("" & mytablex.Fields("subtotal")) * signo
                suma4 = suma4 + Val("" & mytablex.Fields("impuesto")) * signo
                suma5 = suma5 + Val("" & mytablex.Fields("descuento")) * signo
                suma6 = suma6 + Val("" & mytablex.Fields("neto")) * signo
                suma7 = suma7 + Val("" & mytablex.Fields("gravado")) * signo
                suma8 = suma8 + Val("" & mytablex.Fields("percepcion")) * signo
                suma9 = suma9 + Val("" & mytablex.Fields("servicioco")) * signo
      
            End If

        End If

        If Val("" & mytablex.Fields("tipo")) = 1 And Val("" & mytablex.Fields("estado")) = 1 Then
            objExcel.ActiveSheet.Cells(v, h + 5) = "'" & mytablex.Fields("numero")
            objExcel.ActiveSheet.Cells(v, h + 11) = "ANULADO"
            v = v + 1

        End If

        If (Val("" & mytablex.Fields("tipo")) = 2 Or Val("" & mytablex.Fields("tipo")) = 4) And Val("" & mytablex.Fields("estado")) = 1 Then
            objExcel.ActiveSheet.Cells(v, h + 5) = "'" & mytablex.Fields("numero")
            objExcel.ActiveSheet.Cells(v, h + 11) = "ANULADO"
            v = v + 1

        End If

        If Val("" & mytablex.Fields("tipo")) = 3 And Val("" & mytablex.Fields("estado")) = 1 Then
            objExcel.ActiveSheet.Cells(v, h + 5) = "'" & mytablex.Fields("numero")
            objExcel.ActiveSheet.Cells(v, h + 11) = "ANULADO"
            v = v + 1

        End If

    End If

    Exit Function
cmdhola22:
    MsgBox "IMPRIME DETALLE-7a " & error$, 24, "AVISO"
    Exit Function

End Function

Sub subtotal_rve(Tmp As String, _
                 tmpfecha As String, _
                 tmpcaja As String, _
                 v As Double, _
                 h As Double, _
                 tmpserie As String)

    Dim buf1  As String

    Dim found As Integer

    If Val(Tmp) = 1 Or Val(Tmp) = 3 Then
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & tmpfecha
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & busca_sunat(Tmp) '"'" & tmp
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & tmpcaja
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & tmpserie

    End If

    If Val(Tmp) = 1 Then
             
        objExcel.ActiveSheet.Cells(v, h + 5) = "'" & bo_inicial
        objExcel.ActiveSheet.Cells(v, h + 6) = "'" & bo_final
        objExcel.ActiveSheet.Cells(v, h + 11) = "VENTA DEL DIA"
        objExcel.ActiveSheet.Cells(v, h + 12) = Format(suma2, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 13) = Format(suma3, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 14) = Format(suma4, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 15) = Format(suma7, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 16) = Format(suma5, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 17) = Format(suma6, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 18) = Format(suma8, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 19) = Format(suma9, "0.00")
                
        suma2 = 0
        suma3 = 0
        suma4 = 0
        suma5 = 0
        suma6 = 0
        suma7 = 0
        suma8 = 0
        suma9 = 0
        v = v + 1
        nlineas

    End If

    If Val(Tmp) = 3 Then
        objExcel.ActiveSheet.Cells(v, h + 5) = "'" & bo_inicial
        objExcel.ActiveSheet.Cells(v, h + 6) = "'" & bo_final
        objExcel.ActiveSheet.Cells(v, h + 11) = "VENTA DEL DIA"
        objExcel.ActiveSheet.Cells(v, h + 12) = Format(suma2, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 13) = Format(suma3, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 14) = Format(suma4, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 15) = Format(suma7, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 16) = Format(suma5, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 17) = Format(suma6, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 18) = Format(suma8, "0.00")
        objExcel.ActiveSheet.Cells(v, h + 19) = Format(suma9, "0.00")
        v = 1
        suma2 = 0
        suma3 = 0
        suma4 = 0
        suma5 = 0
        suma6 = 0
        suma7 = 0
        suma8 = 0
        suma9 = 0

    End If

End Sub

Public Function Formato_Excelrve(Num_Campos As Integer, _
                                 Nombre_Campos() As String) As Boolean

    Dim I As Integer

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        '.Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        '.Range(.Cells(3, 1), .Cells(3, 8)).Font.Bold = True
        
        'For i = 1 To Num_Campos Step 1
        '    .Cells(3, i) = Nombre_Campos(i)
        'Next i
        'hasta aki pa colocar los titulos
        
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
        '.columns("A").ColumnWidth = 10
        '.columns("B").ColumnWidth = 6
        '.columns("C").ColumnWidth = 6
        '.columns("D").ColumnWidth = 10
        '.columns("E").ColumnWidth = 10
        '.columns("F").ColumnWidth = 10
        '.columns("G").ColumnWidth = 30
        '.columns("H").ColumnWidth = 2
        
        .Cells(4, 1) = "FECHA"
        .Cells(4, 2) = "TI"
        .Cells(4, 3) = "CA"
        .Cells(4, 4) = "SERIE"
        .Cells(4, 5) = "INICIO"
        .Cells(4, 6) = "FINAL"
        .Cells(4, 7) = "FACTURA"
        .Cells(4, 8) = dicruc
        .Cells(4, 9) = "INICIAL"
        .Cells(4, 10) = "FINAL"
        .Cells(4, 11) = "CLIENTE"
        .Cells(4, 12) = "VENTA"
        .Cells(4, 13) = "VENTA"
        .Cells(4, 14) = "dicigv"
        .Cells(4, 15) = "INAFEC"
        .Cells(4, 16) = "DSCTO"
        .Cells(4, 17) = "TBRUTO"
        .Cells(4, 18) = "PERCEPCION"
        .Cells(4, 19) = "SERVICIO"
        
        .Cells(3, 5) = "REGISTRADORA"
        .Cells(3, 7) = "MANUAL"
        .Cells(3, 9) = "BOLETA VENTA"
        .Cells(3, 12) = "PRECIO"
        .Cells(3, 13) = "VALOR"
        
        .columns("A").ColumnWidth = 15
        .columns("B").ColumnWidth = 5
        .columns("C").ColumnWidth = 5
        .columns("D").ColumnWidth = 12
        .columns("E").ColumnWidth = 12
        .columns("F").ColumnWidth = 10
        .columns("G").ColumnWidth = 10
        .columns("H").ColumnWidth = 15
        .columns("I").ColumnWidth = 10
        .columns("J").ColumnWidth = 10
        .columns("K").ColumnWidth = 25
        
        .columns("l").ColumnWidth = 10
        .columns("m").ColumnWidth = 10
   
    End With

End Function

''' 09/01/2018 Registro de Ventas para Contasis
Private Sub Label17_Click()

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

    If MsgBox("Desea Procesar..", 1, "Aviso") <> 1 Then Exit Sub

    found = sql_documento(mytablex)

    If found = 0 Then
        'mytablex.Close
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    
    If consolidado <> "S" Then
        If tipoprint = "Excell" Then
            cuerpo_excellContasis mytablex
            Exit Sub

        End If

    End If
    
    Close #1
    cerrar_archivo
    
    mytablex.Close

End Sub

''' 09/01/2018 Registro de Ventas para Contasis

