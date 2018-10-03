VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form pedprov 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Genera Orden Compra"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   14115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   14115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFF00&
      Caption         =   "Lista Precios"
      Height          =   3735
      Left            =   7200
      TabIndex        =   106
      Top             =   3720
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CommandButton Command8 
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
         Left            =   7200
         MaskColor       =   &H00E0E0E0&
         Picture         =   "pedprov.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Borrar registro"
         Top             =   360
         Width           =   735
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Height          =   3255
         Left            =   120
         OleObjectBlob   =   "pedprov.frx":1212
         TabIndex        =   108
         Top             =   360
         Width           =   6975
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consultas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   4680
      TabIndex        =   76
      Top             =   5520
      Visible         =   0   'False
      Width           =   13695
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H8000000D&
         Caption         =   "&Ejecutar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8160
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox buffer 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "pedprov.frx":2275
         Height          =   6255
         Left            =   120
         OleObjectBlob   =   "pedprov.frx":2289
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   1080
         Width           =   13455
      End
   End
   Begin VB.TextBox xproveedor 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF00&
      Caption         =   "Ingreso de Pedidos x Local"
      Height          =   5415
      Left            =   7200
      TabIndex        =   36
      Top             =   3120
      Visible         =   0   'False
      Width           =   10455
      Begin VB.TextBox precio 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1575
      End
      Begin VB.TextBox cantidad 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox xproveedorp 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox l1 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox l2 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox l3 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox l4 
         Height          =   375
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
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
         Left            =   9360
         MaskColor       =   &H00E0E0E0&
         Picture         =   "pedprov.frx":2C54
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Grabar registro"
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command5 
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
         Left            =   9360
         MaskColor       =   &H00E0E0E0&
         Picture         =   "pedprov.frx":3E66
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Borrar registro"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox observa4 
         Height          =   375
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox observa3 
         Height          =   375
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox observa2 
         Height          =   375
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox observa1 
         Height          =   375
         Left            =   3480
         MaxLength       =   20
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox lx4 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox lx3 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox lx2 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox lx1 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factor"
         Height          =   375
         Left            =   720
         TabIndex        =   73
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label factor 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   72
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unidad"
         Height          =   375
         Left            =   720
         TabIndex        =   71
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label unidad 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   70
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio"
         Height          =   375
         Left            =   720
         TabIndex        =   69
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   375
         Left            =   720
         TabIndex        =   68
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label xnombre 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   67
         Top             =   4200
         Width           =   4695
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   375
         Left            =   720
         TabIndex        =   66
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor"
         Height          =   375
         Left            =   720
         TabIndex        =   65
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         Height          =   375
         Left            =   720
         TabIndex        =   64
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label tl1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   720
         TabIndex        =   63
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label tl2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   720
         TabIndex        =   62
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label tl3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   720
         TabIndex        =   61
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label tl4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   720
         TabIndex        =   60
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label34 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   59
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label35 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   58
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label36 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   57
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
         Height          =   375
         Left            =   1920
         TabIndex        =   56
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observaciones"
         Height          =   375
         Left            =   3480
         TabIndex        =   55
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Can.Recibido"
         Height          =   375
         Left            =   6960
         TabIndex        =   54
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Generando Documentos"
      Height          =   4215
      Left            =   2640
      TabIndex        =   25
      Top             =   480
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox xobserva 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   104
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox xdias 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   2880
         MaxLength       =   3
         TabIndex        =   103
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox xVendedor 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1800
         MaxLength       =   11
         TabIndex        =   101
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox xmoneda 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   100
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox xfpago 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   97
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox xfecha 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   82
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox xnumero 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1800
         MaxLength       =   11
         TabIndex        =   30
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox xserie 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   29
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
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
         Left            =   5880
         MaskColor       =   &H00E0E0E0&
         Picture         =   "pedprov.frx":5078
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Grabar registro"
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command2 
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
         Left            =   5880
         MaskColor       =   &H00E0E0E0&
         Picture         =   "pedprov.frx":628A
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Borrar registro"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox xtipo 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label33 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observacion"
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
         Left            =   360
         TabIndex        =   105
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label32 
         BackColor       =   &H80000009&
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
         Left            =   360
         TabIndex        =   102
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label31 
         BackColor       =   &H80000009&
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
         Left            =   360
         TabIndex        =   99
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label pago 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fpago"
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
         Left            =   360
         TabIndex        =   98
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label25 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
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
         Left            =   360
         TabIndex        =   83
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000009&
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
         Left            =   360
         TabIndex        =   35
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000009&
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
         Left            =   360
         TabIndex        =   34
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000009&
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
         Left            =   360
         TabIndex        =   33
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000009&
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
         Left            =   360
         TabIndex        =   32
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label xdescripcio 
         BackColor       =   &H80000009&
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
         Left            =   1800
         TabIndex        =   31
         Top             =   720
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Cargar Pedidos"
      Height          =   3015
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   10095
      Begin VB.TextBox proveedorp 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1080
         MaxLength       =   11
         TabIndex        =   14
         Text            =   "*"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox tipo 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "*"
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox serie 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7440
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "*"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox numero 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7440
         MaxLength       =   11
         TabIndex        =   11
         Text            =   "*"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox fechai 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox fechaf 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   9
         Top             =   960
         Width           =   2055
      End
      Begin VB.ComboBox estado 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox codigo 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   7440
         MaxLength       =   11
         TabIndex        =   7
         Text            =   "*"
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox tipoclie 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
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
         Left            =   9240
         MaskColor       =   &H00E0E0E0&
         Picture         =   "pedprov.frx":749C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Borrar registro"
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command7 
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
         Left            =   9240
         MaskColor       =   &H00E0E0E0&
         Picture         =   "pedprov.frx":86AE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Grabar registro"
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor"
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
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
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
         TabIndex        =   23
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
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
         Left            =   6240
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
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
         Left            =   6240
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
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
         Left            =   2880
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaFinal"
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
         Left            =   2880
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
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
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
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
         Left            =   2880
         TabIndex        =   17
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF00&
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
         Left            =   6240
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label acu 
         BackColor       =   &H80000009&
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
         Left            =   11640
         TabIndex        =   15
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.ComboBox orden 
      BackColor       =   &H00C0FFFF&
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
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "pedprov.frx":98C0
      Height          =   6495
      Left            =   120
      OleObjectBlob   =   "pedprov.frx":98D4
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   480
      Width           =   13575
   End
   Begin VB.Label c4 
      BackColor       =   &H80000009&
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
      Left            =   120
      TabIndex        =   109
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label c3 
      BackColor       =   &H80000009&
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
      Left            =   120
      TabIndex        =   96
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label c2 
      BackColor       =   &H80000009&
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
      Left            =   120
      TabIndex        =   95
      Top             =   7440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label c1 
      BackColor       =   &H80000009&
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
      Left            =   120
      TabIndex        =   94
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label xnombre1 
      BackColor       =   &H80000009&
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
      Left            =   6600
      TabIndex        =   93
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label30 
      BackColor       =   &H80000009&
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
      Left            =   5040
      TabIndex        =   92
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Poner Ceros"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   91
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Calcular"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4200
      TabIndex        =   90
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label tximpuesto 
      BackColor       =   &H80000009&
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
      Left            =   7200
      TabIndex        =   89
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label txtotal 
      BackColor       =   &H80000009&
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
      Left            =   11880
      TabIndex        =   88
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label txdescuento 
      BackColor       =   &H80000009&
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
      Left            =   10320
      TabIndex        =   87
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label txneto 
      BackColor       =   &H80000009&
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
      Left            =   8760
      TabIndex        =   86
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label txsubtotal 
      BackColor       =   &H80000009&
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
      Left            =   5640
      TabIndex        =   85
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label27 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proveedor"
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
      TabIndex        =   75
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buscar"
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
      Height          =   375
      Left            =   3600
      TabIndex        =   74
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Orden"
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
      Left            =   240
      TabIndex        =   2
      Top             =   7800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu nu3434 
      Caption         =   "&Buscar"
   End
   Begin VB.Menu hrni343 
      Caption         =   "&Generar"
   End
   Begin VB.Menu prol343 
      Caption         =   "&Menu"
      Begin VB.Menu lo854 
         Caption         =   "&1.Actualizar Costos"
      End
      Begin VB.Menu dki4545 
         Caption         =   "&2.Cargar desde Pedidos"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu lsoere21 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "pedprov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xacu As String
Dim xproducto As String
Private Type campo_precio
    unidad As String
    factor As String
    precio As String
    costo As String
    margen As String
    stock As String
End Type
Dim campo_precios(12) As campo_precio
Private Sub codigo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
serie.SetFocus

End Sub

Private Sub Command1_Click()
Dim buf As String
Dim tmp As String
Dim sw As Integer
Dim sdx As Double
Dim mydbx As Database
Dim mytabley As Table
Dim mytablez As Table
Dim found As Integer
On Error GoTo cmd671_err

'------------------------------
If Len(xtipo) = 0 Then
   xtipo.SetFocus
   Exit Sub
End If
If Len(xserie) = 0 Then
   xserie.SetFocus
   Exit Sub
End If
If Len(xnumero) = 0 Then
   xnumero.SetFocus
   Exit Sub
End If
If Len(xfecha) <> 10 Then
   xfecha.SetFocus
   Exit Sub
End If
If Not IsDate(xfecha) Then
   xfecha.SetFocus
   Exit Sub
End If
found = busca_fpago()
If found = 0 Then
   xfpago.SetFocus
   Exit Sub
End If
If xmoneda <> "S" And xmoneda <> "D" Then
   xmoneda.SetFocus
End If

Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytabley = mydbx.OpenTable("cordenc")
mytabley.Index = "tfactura1"

Set mytablez = mydbx.OpenTable("dordenc")
mytablez.Index = "cuerpo1"

Data2.Recordset.MoveFirst
sw = 0
Do
If Data2.Recordset.EOF Then Exit Do
If Val("" & Data2.Recordset.Fields("cantidad")) > 0 Then
If sw = 0 Then
   sw = 1
   tmp = "" & Data2.Recordset.Fields("proveedorp")
   grabar_cabecera mytabley
End If
If tmp <> "" & Data2.Recordset.Fields("proveedorp") Then
   tmp = "" & Data2.Recordset.Fields("proveedorp")
   'sumar uno
   sdx = Val(xnumero) + 1
   xnumero = "" & sdx
   grabar_cabecera mytabley
End If
hacer_detalle mytablez
End If
Data2.Recordset.MoveNext
Loop
mytabley.Close
mytablez.Close
mydbx.Close
MsgBox "Proceso Terminado ", 48, "Aviso"
Command2_Click
Exit Sub
cmd671_err:
Exit Sub
End Sub
Sub hacer_detalle(mytablez As Table)
Dim i As Integer
Dim sdx As Double
mytablez.Seek "=", xtipo, xserie, xnumero, "" & Data2.Recordset.Fields("producto"), "" & Data2.Recordset.Fields("proveedorp")
If mytablez.NoMatch Then
   mytablez.AddNew
   For i = 0 To Data2.Recordset.Fields.Count - 1
       mytablez.Fields(i) = Data2.Recordset.Fields(i)
   Next i
   mytablez.Fields("tipo") = xtipo
   mytablez.Fields("serie") = xserie
   mytablez.Fields("numero") = xnumero
   mytablez.Fields("moneda") = xmoneda
   mytablez.Fields("vendedor") = xVendedor
   mytablez.Fields("acu") = xacu
   mytablez.Fields("acu1") = acu
   mytablez.Fields("fecha") = Format(Now, "dd/mm/yyyy")
   mytablez.Fields("tipoclie") = "P"
   mytablez.Fields("codigo") = "" & Data2.Recordset.Fields("proveedorp")
   calcula_igv1 mytablez
   mytablez.Update
End If
If Not mytablez.NoMatch Then
   mytablez.Edit
   mytablez.Fields("l1") = Val("" & mytablez.Fields("l1")) + Val("" & Data2.Recordset.Fields("l1"))
   mytablez.Fields("l2") = Val("" & mytablez.Fields("l2")) + Val("" & Data2.Recordset.Fields("l2"))
   mytablez.Fields("l3") = Val("" & mytablez.Fields("l3")) + Val("" & Data2.Recordset.Fields("l3"))
   mytablez.Fields("l4") = Val("" & mytablez.Fields("l4")) + Val("" & Data2.Recordset.Fields("l4"))
   mytablez.Fields("cantidad") = Val("" & mytablez.Fields("l1")) + Val("" & mytablez.Fields("l2")) + Val("" & mytablez.Fields("l3")) + Val("" & mytablez.Fields("l4"))
   mytablez.Update
End If

End Sub

Sub grabar_cabecera(mytabley As Table)
Dim found As Integer
Dim i As Integer
mytabley.Seek "=", xtipo, xserie, xnumero, "P", "" & Data2.Recordset.Fields("proveedorp")
If mytabley.NoMatch Then
   mytabley.AddNew
   pone_registro_compra mytabley
   mytabley.Update
   '------------------- grabar tipo
   found = busca_tipo("" & xtipo, 1)
End If

End Sub

Private Sub Command2_Click()
lsoere21_Click
End Sub

Private Sub Command6_Click()
lsoere21_Click
End Sub

Private Sub Command7_Click()
Dim mytablex As Snapshot
Dim mytabley As Table
Dim mydbx As Database
Dim found As Integer
Dim buf As String
Dim i As Integer
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytabley = mydbx.OpenTable("tcargar")
mytabley.Index = "tcargar"
found = sql_detalle(mydbx, mytablex)
Do
If mytablex.EOF Then Exit Do
'---------------------------------------
mytabley.Seek "=", "" & mytablex.Fields("fecha"), "" & mytablex.Fields("proveedorp"), "" & mytablex.Fields("producto")
If mytabley.NoMatch Then
   mytabley.AddNew
   For i = 0 To mytablex.Fields.Count - 1
       mytabley.Fields(i) = mytablex.Fields(i)
   Next i
   mytabley.Fields("fecha") = xfecha
   mytabley.Update
End If
If Not mytabley.NoMatch Then
   mytabley.Edit
   mytabley.Fields("l1") = Val("" & mytabley.Fields("l1")) + Val("" & mytablex.Fields("l1"))
   mytabley.Fields("l2") = Val("" & mytabley.Fields("l2")) + Val("" & mytablex.Fields("l2"))
   mytabley.Fields("l3") = Val("" & mytabley.Fields("l3")) + Val("" & mytablex.Fields("l3"))
   mytabley.Fields("l4") = Val("" & mytabley.Fields("l4")) + Val("" & mytablex.Fields("l4"))
   mytabley.Fields("cantidad") = Val("" & mytabley.Fields("cantidad")) + Val("" & mytablex.Fields("cantidad"))
   mytabley.Update
End If
'---------------------------------------
mytablex.MoveNext
Loop
mytabley.Close
mytablex.Close
mydbx.Close
Label26_Click
lsoere21_Click
End Sub

Private Sub Command8_Click()
Frame5.Visible = False
DBGrid2.SetFocus
End Sub

Private Sub Command9_Click()

End Sub

Private Sub DBGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Dim found As Integer
If ColIndex > 5 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
Case 1, 2, 4, 8, 9, 10, 11, 12, 13
     Cancel = True
     Exit Sub
Case 0
     If Len("" & DBGrid2.Columns(0)) > 0 Then  'si ya existe no cambiar
        Cancel = True
        Exit Sub
     End If
     
Case 3
     If Len("" & DBGrid2.Columns(0)) = 0 Then  '
        Cancel = True
        Exit Sub
     End If
     If Len("" & DBGrid2.Columns(17)) > 0 Then
        Cancel = True
        Exit Sub
     End If
Case 5
     If Len("" & DBGrid2.Columns(0)) = 0 Then  '
        Cancel = True
        Exit Sub
     End If
End Select

End Sub

Private Sub DBGrid2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim sdx As Double
Select Case ColIndex
     Case 3
     If Len(DBGrid2.Columns(0)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Not IsNumeric("" & DBGrid2.Columns(3)) Then
        Cancel = True
        Exit Sub
     End If
     sdx = Val("" & DBGrid2.Columns(3)) * Val("" & DBGrid2.Columns(5))
     DBGrid2.Columns(9) = Val(Format(sdx, "0.00"))
     DBGrid2.Columns(7) = Val(Format(sdx, "0.00"))
     calcula_igv
     Case 5
     If Len(DBGrid2.Columns(0)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Val("" & DBGrid2.Columns(3)) <= 0 Then
        Cancel = True
        Exit Sub
     End If
     sdx = Val("" & DBGrid2.Columns(3)) * Val("" & DBGrid2.Columns(5))
     DBGrid2.Columns(9) = Val(Format(sdx, "0.00"))
     DBGrid2.Columns(7) = Val(Format(sdx, "0.00"))
     calcula_igv
End Select

End Sub


Private Sub DBGrid3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Command8_Click
   Exit Sub
End If
If KeyCode = 13 Then
   If Len("" & DBGrid3.Columns(0)) > 0 And Val("" & DBGrid3.Columns(1)) > 0 And Len("" & DBGrid3.Columns(3)) > 0 Then
      Data2.Recordset.Edit
      Data2.Recordset.Fields("unidad") = "" & DBGrid3.Columns(0)
      Data2.Recordset.Fields("factor") = "" & DBGrid3.Columns(1)
      Data2.Recordset.Fields("precio") = "" & DBGrid3.Columns(3)
      Data2.Recordset.Update
      Command8_Click
   End If
End If
End Sub

Private Sub DBGrid3_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim dr As Integer
Dim row_num As Integer
Dim r As Integer
Dim rows_returned As Integer
If ReadPriorRows Then
        dr = -1
    Else
        dr = 1
    End If
    If IsNull(StartLocation) Then
        If ReadPriorRows Then
           row_num = RowBuf.RowCount - 1
           'row_num = 9
        Else
           row_num = 0
        End If
    Else
        row_num = CLng(StartLocation) + dr
    End If
    rows_returned = 0
    For r = 0 To RowBuf.RowCount - 1
        If row_num < 0 Or row_num > 9 Then Exit For
        RowBuf.Value(r, 0) = campo_precios(row_num).unidad
        RowBuf.Value(r, 1) = campo_precios(row_num).factor
        RowBuf.Value(r, 2) = campo_precios(row_num).precio
        RowBuf.Value(r, 3) = campo_precios(row_num).costo
        RowBuf.Value(r, 4) = campo_precios(row_num).margen
        RowBuf.Value(r, 5) = campo_precios(row_num).stock
        RowBuf.Bookmark(r) = row_num
        row_num = row_num + dr
        rows_returned = rows_returned + 1
   Next r
   RowBuf.RowCount = rows_returned

End Sub

Private Sub estado_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
proveedorp.SetFocus

End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
'fpago.SetFocus

End Sub

Private Sub fechaf_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
tipoclie.SetFocus

End Sub

Private Sub fechai_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
fechaf.SetFocus

End Sub

Private Sub gat3434_Click()


End Sub

Private Sub fpago_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then
   KeyAscii = 0
   Exit Sub
End If
found = busca_fpago("" & xfpago, 0)
Label26_Click

End Sub

Private Sub hrni343_Click()
Dim found As Integer
found = busca_proveedor(xproveedor)
If found = 0 Then Exit Sub
Label28_Click
If Val(txtotal) <= 0 Then Exit Sub


If Frame2.Visible = True Then Exit Sub
If Frame3.Visible = True Then Exit Sub
If Frame4.Visible = True Then Exit Sub
Frame2.Visible = True
xtipo = ""
xserie = ""
xnumero = ""
xfecha = Format(Now, "dd/mm/yyyy")
xtipo.SetFocus
End Sub

Private Sub Label26_Click()
If Len(xproveedor) = 0 Then
   xproveedor.SetFocus
   Exit Sub
End If
sql_cargar
End Sub

Private Sub ldoer3431_Click()

End Sub

Private Sub Label28_Click()
sumar_detalle
End Sub

Private Sub Label29_Click()
poner_ceros
End Sub

Private Sub mxiklo343_Click()

End Sub

Private Sub numero_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Command7_Click

End Sub

Private Sub proveedorp_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
fechai.SetFocus

End Sub

Private Sub serie_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
numero.SetFocus

End Sub

Private Sub tipo_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
estado.SetFocus
End Sub

Private Sub tipoclie_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
codigo.SetFocus

End Sub

Private Sub xdias_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
xmoneda.SetFocus
End Sub

Private Sub xfecha_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
xfpago.SetFocus

End Sub

Private Sub xfpago_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   KeyAscii = 0
   Exit Sub
End If
xdias.SetFocus
End Sub

Private Sub xfpago_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_fpago
End If

End Sub

Private Sub xmoneda_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
xVendedor.SetFocus
End Sub

Private Sub xnumero_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
xfecha.SetFocus

End Sub

Private Sub xobserva_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Command1.SetFocus
End Sub

Private Sub xproveedor_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then
   KeyAscii = 0
   Exit Sub
End If
Label26_Click
End Sub

Private Sub xproveedor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_xproveedor
End If

End Sub

Private Sub xserie_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
xnumero.SetFocus

End Sub

Private Sub xtipo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then
   KeyAscii = 0
   Exit Sub
End If
found = busca_tipo("" & xtipo, 0)
xserie.SetFocus

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   lsoere21_Click
   Exit Sub
End If
Command3_Click

End Sub

Private Sub Command3_Click()
Dim buf As String
If opcion1 = "3" Then
      If Len(buffer) = 0 Then
      buf = "select Descripcio,Fpago from fpago "
      Else
      buf = "select Descripcio,Fpago from fpago where  " & Combo1 & " like '" & buffer & "*'"
      End If
   End If
If opcion1 = "5" Then
      If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from vendedor "
      Else
      buf = "select Nombre,vendedor from Vendedor where  " & Combo1 & " like '" & buffer & "*'"
      End If
   End If

If opcion1 = "1" Then
      If Len(buffer) = 0 Then
      buf = "select Descripcio,Tipo from Tipo where tipodoc='R'"
      Else
      buf = "select Descripcio,Tipo from Tipo where tipodoc='R' and " & Combo1 & " like '" & buffer & "*'"
      End If
   End If
   If opcion1 = "2" Or opcion1 = "4" Then
      If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from proveedo "
      Else
      buf = "select Nombre,Codigo from proveedor where " & Combo1 & " like '" & buffer & "*'"
      End If
   End If
   
   If Combo2.ListIndex = 1 Then
      buf = buf & " order by " & Combo1
   End If
               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldir
               Data1.RecordSource = buf
               Data1.Refresh
               If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
                  Data1.Recordset.Close
                  buffer.SetFocus
                  Exit Sub
               End If
               If opcion1 = "1" Or opcion1 = "2" Or opcion1 = "4" Then
                  DBGrid1.Columns(0).Width = 4000
                  DBGrid1.Columns(1).Width = 2000
               End If
               DBGrid1.SetFocus

End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
xproveedorp.SetFocus

End Sub

Private Sub Command4_Click()
Dim sdx As Double
If Val(precio) <= 0 Then
   precio.SetFocus
   Exit Sub
End If
If Val(cantidad) <= 0 Then
   cantidad.SetFocus
   Exit Sub
End If

If Val(l1) < 0 Then
   l1.SetFocus
   Exit Sub
End If
If Val(l2) < 0 Then
   l2.SetFocus
   Exit Sub
End If
If Val(l3) < 0 Then
   l3.SetFocus
   Exit Sub
End If
If Val(l4) < 0 Then
   l4.SetFocus
   Exit Sub
End If

If Val(lx1) < 0 Then
   lx1.SetFocus
   Exit Sub
End If
If Val(lx2) < 0 Then
   lx2.SetFocus
   Exit Sub
End If
If Val(lx3) < 0 Then
   lx3.SetFocus
   Exit Sub
End If
If Val(lx4) < 0 Then
   lx4.SetFocus
   Exit Sub
End If
sdx = Val(l1) + Val(l2) + Val(l3) + Val(l4)
If sdx <> Val(cantidad) Then
   cantidad.SetFocus
   Exit Sub
End If
Data2.Recordset.Edit
Data2.Recordset.Fields("precio") = Val(precio)
Data2.Recordset.Fields("l1") = Val(l1)
Data2.Recordset.Fields("l2") = Val(l2)
Data2.Recordset.Fields("l3") = Val(l3)
Data2.Recordset.Fields("l4") = Val(l4)
'sdx = Val(l1) + Val(l2) + Val(l3) + Val(l4)
Data2.Recordset.Fields("cantidad") = Val(cantidad)

Data2.Recordset.Fields("lx1") = Val(lx1)
Data2.Recordset.Fields("lx2") = Val(lx2)
Data2.Recordset.Fields("lx3") = Val(lx3)
Data2.Recordset.Fields("lx4") = Val(lx4)

Data2.Recordset.Fields("observa1") = observa1
Data2.Recordset.Fields("observa2") = observa2
Data2.Recordset.Fields("observa3") = observa3
Data2.Recordset.Fields("observa4") = observa4
Data2.Recordset.Fields("proveedorp") = xproveedorp
Data2.Recordset.Fields("total") = Val(Format(Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & Data2.Recordset.Fields("precio")), "0.00"))
Data2.Recordset.Update
lsoere21_Click
DBGrid2.SetFocus

End Sub

Private Sub Command5_Click()
lsoere21_Click
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
If opcion1 = "1" Then
   xtipo = DBGrid1.Columns(1)
   Frame4.Visible = False
   xtipo.SetFocus
End If
If opcion1 = "3" Then
   xfpago = DBGrid1.Columns(1)
   Frame4.Visible = False
   xfpago.SetFocus
End If

If opcion1 = "2" Then
   xproveedorp = DBGrid1.Columns(1)
   Frame4.Visible = False
   xproveedorp.SetFocus
End If
If opcion1 = "5" Then
   xVendedor = DBGrid1.Columns(1)
   Frame4.Visible = False
   xVendedor.SetFocus
End If

If opcion1 = "4" Then
   xproveedor = DBGrid1.Columns(1)
   xnombre1 = DBGrid1.Columns(0)
   Frame4.Visible = False
   xproveedor.SetFocus
   xproveedor_KeyPress 13
End If

End If

End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   If Len(DBGrid2.Columns(0)) > 0 And DBGrid2.Col = 2 Then
      xproducto = "" & DBGrid2.Columns(0)
      carga_dbgrid3
   End If
End If

If KeyCode = &H72 Then  'f3
   ingreso_locales
End If

End Sub

Private Sub Form_Load()
xfecha = Format(Now, "dd/mm/yyyy")
orden.Clear
orden.AddItem "*"
orden.AddItem "Proveedorp"
orden.AddItem "Codigo"
orden.AddItem "Fecha"
orden.AddItem "Serie"
orden.AddItem "Numero"
orden.AddItem "tipoclie"
orden.ListIndex = 1


estado.Clear
estado.AddItem "*"
estado.AddItem "0"
estado.AddItem "1"
estado.ListIndex = 0

tipoclie.Clear
tipoclie.AddItem "*"
tipoclie.AddItem "P"
tipoclie.AddItem "C"
tipoclie.AddItem "I"
tipoclie.ListIndex = 0

fechai = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
fechaf = Format(Now, "dd/mm/yyyy")


End Sub
Function sql_detalle(mydbx As Database, mytablex As Snapshot)
Dim buf As String
buf = "select * from  drequisa"
buf = buf & "  where fecha>=" & "DateValue('" & fechai & "'" & ")"
buf = buf & " and  fecha<=" & "DateValue('" & fechaf & "'" & ")"
buf = buf & " and acu='" & acu & "'"
If tipo <> "*" Then
buf = buf & " and tipo like '" & tipo & "'"
End If
If serie <> "*" Then
buf = buf & " and serie like '" & serie & "'"
End If
If numero <> "*" Then
buf = buf & " and numero like '" & numero & "'"
End If
If tipoclie <> "*" Then
buf = buf & " and tipoclie like '" & tipoclie & "'"
End If
If codigo <> "*" Then
buf = buf & " and codigo like '" & codigo & "'"
End If
If proveedorp <> "*" Then
buf = buf & " and proveedorp like '" & proveedorp & "'"
End If
buf = buf & " order by proveedorp,producto"
Set mytablex = mydbx.CreateSnapshot(buf)
End Function
Sub sql_cargar()

Dim buf As String
buf = "select * from  tabprov "
buf = buf & "  where proveedorp='" & xproveedor & "'"
buf = buf & " order by descripcio"
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = buf
               Data2.Refresh
               'DBGrid2.Columns(ColIndex_List1).Button = True
               DBGrid2.SetFocus
               
               

End Sub


Private Sub l1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
l2.SetFocus
End Sub

Private Sub l2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
l3.SetFocus

End Sub

Private Sub l3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
l4.SetFocus

End Sub

Private Sub l4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
precio.SetFocus
End Sub

Private Sub lsoere21_Click()
If Frame2.Visible = True Then
   Frame2.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If
If Frame1.Visible = True Then
   Frame1.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If

If Frame3.Visible = True Then
   Frame3.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If
If Frame4.Visible = True Then
   Frame4.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If
pedprov.Hide
Unload pedprov
End Sub

Private Sub nu3434_Click()
Label26_Click
End Sub
Sub xxpone_locales()
Dim found As Integer
found = xxbusca_locales()
If found = 0 Then Exit Sub
l1 = "" & Data2.Recordset.Fields("l1")
l2 = "" & Data2.Recordset.Fields("l2")
l3 = "" & Data2.Recordset.Fields("l3")
l4 = "" & Data2.Recordset.Fields("l4")

lx1 = "" & Data2.Recordset.Fields("lx1")
lx2 = "" & Data2.Recordset.Fields("lx2")
lx3 = "" & Data2.Recordset.Fields("lx3")
lx4 = "" & Data2.Recordset.Fields("lx4")

unidad = "" & Data2.Recordset.Fields("unidad")
factor = "" & Data2.Recordset.Fields("factor")
precio = "" & Data2.Recordset.Fields("precio")
cantidad = "" & Data2.Recordset.Fields("cantidad")

observa1 = "" & Data2.Recordset.Fields("observa1")
observa2 = "" & Data2.Recordset.Fields("observa2")
observa3 = "" & Data2.Recordset.Fields("observa3")
observa4 = "" & Data2.Recordset.Fields("observa4")

xproveedorp = "" & Data2.Recordset.Fields("proveedorp")

End Sub
Function xxbusca_locales()
Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("empresa")
mytablex.Index = "codigo"
mytablex.Seek "=", menup.gempresa
If Not mytablex.NoMatch Then
   xxbusca_locales = 1
   tl1 = "" & mytablex.Fields("l1")
   tl2 = "" & mytablex.Fields("l2")
   tl3 = "" & mytablex.Fields("l3")
   tl4 = "" & mytablex.Fields("l4")
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close
End Function
Sub ingreso_locales()
xxpone_locales
Frame3.Visible = True
l1.SetFocus

End Sub

Private Sub precio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
cantidad.SetFocus

End Sub

Private Sub xproveedorp_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Command4.SetFocus

End Sub

Private Sub xproveedorp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_proveedor
End If

End Sub
Sub consulta_proveedor()
Combo1.Clear
Combo1.AddItem "Codigo"
Combo1.AddItem "Nombre"
Combo1.ListIndex = 0
Frame4.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "2"
Command3_Click


End Sub
Sub consulta_xproveedor()
Combo1.Clear
Combo1.AddItem "Codigo"
Combo1.AddItem "Nombre"
Combo1.ListIndex = 0
Frame4.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "4"
Command3_Click

End Sub

Private Sub xtipo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_tipo
End If

End Sub
Sub consulta_tipo()
Combo1.Clear
Combo1.AddItem "Tipo"
Combo1.AddItem "Descripcio"
Combo1.ListIndex = 0
Frame4.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "1"
Command3_Click
End Sub
Sub consulta_fpago()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Fpago"
Combo1.ListIndex = 0
Frame4.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "3"
Command3_Click

End Sub
Sub consulta_vendedor()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Codigo"
Combo1.ListIndex = 0
Frame4.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "5"
Command3_Click

End Sub

Function busca_numero(buf As String)
Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("cordenc")
mytablex.Index = "tfactura"
mytablex.Seek "=", xtipo, xserie, buf
If Not mytablex.NoMatch Then
   busca_numero = 1
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close
End Function
Function busca_tipo(buf As String, sw As Integer)
Dim mydbx As Database
Dim mytablex As Table
Dim sdx As Double
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("tipo")
mytablex.Index = "tipo"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   busca_tipo = 1
   xacu = "" & mytablex.Fields("tipodoc")
   If sw = 0 Then
      xserie = "" & mytablex.Fields("serie")
      xnumero = "" & mytablex.Fields("numero")
      sdx = Val(xnumero) + 1
      xnumero = "" & sdx
   End If
   If sw = 1 Then
      mytablex.Edit
      mytablex.Fields("numero") = xnumero
      mytablex.Update
   End If
End If
mytablex.Close
mydbx.Close

End Function
Sub calcula_igv()
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double
Dim found As Integer
sdx = Val("" & DBGrid2.Columns(5)) * Val("" & DBGrid2.Columns(3))
DBGrid2.Columns(7) = Val(Format(sdx, "0.00"))  'total
DBGrid2.Columns(9) = Val(Format(sdx, "0.00"))  'neto
sdx = Val("" & DBGrid2.Columns(9)) * Val("" & DBGrid2.Columns(6)) / 100
sdx2 = Val("" & DBGrid2.Columns(9)) - sdx
DBGrid2.Columns(8) = Val(Format(sdx, "0.00"))  'descuento
DBGrid2.Columns(7) = Val(Format(sdx2, "0.00"))  'total
DBGrid2.Columns(11) = 0
DBGrid2.Columns(10) = 0
If Val("" & DBGrid2.Columns(7)) > 0 And Val("" & DBGrid2.Columns(12)) > 0 Then
   sdx1 = 1 + Val("" & DBGrid2.Columns(12)) / 100
   sdx1 = Val(Format(sdx1, "0.00"))
   sdx1 = Val(DBGrid2.Columns(7)) / sdx1
   DBGrid2.Columns(11) = Val(Format(sdx1, "0.00"))  'subtotal
   sdx = Val(DBGrid2.Columns(7)) - Val(DBGrid2.Columns(11))
   DBGrid2.Columns(10) = Val(Format(sdx, "0.00"))  'total
End If
'found = pone_valoresxxx("" & DBGrid2.Columns(0), Val("" & DBGrid2.Columns(7)))
End Sub

Sub calcula_igv1(mytablex As Table)
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double
sdx = Val("" & mytablex.Fields("precio")) * Val("" & mytablex.Fields("cantidad"))
mytablex.Fields("total") = Val(Format(sdx, "0.00"))  'total
mytablex.Fields("neto") = Val(Format(sdx, "0.00"))  'neto
sdx = Val("" & mytablex.Fields("neto")) * Val("" & mytablex.Fields("deslipo")) / 100
sdx2 = Val("" & mytablex.Fields("neto")) - sdx
mytablex.Fields("descuento") = Val(Format(sdx, "0.00"))  'descuento
mytablex.Fields("total") = Val(Format(sdx2, "0.00"))  'total
mytablex.Fields("subtotal") = 0
mytablex.Fields("impuesto") = 0
If Val("" & mytablex.Fields("total")) > 0 And Val("" & mytablex.Fields("igv")) > 0 Then
   sdx1 = 1 + Val("" & mytablex.Fields("igv")) / 100
   sdx1 = Val(Format(sdx1, "0.00"))
   sdx1 = Val("" & mytablex.Fields("total")) / sdx1
   mytablex.Fields("subtotal") = Val(Format(sdx1, "0.00"))  'subtotal
   sdx = Val("" & mytablex.Fields("total")) - Val("" & mytablex.Fields("subtotal"))
   mytablex.Fields("impuesto") = Val(Format(sdx, "0.00"))  'total
End If

End Sub
Sub sumar_detalle()
On Error GoTo cmd35_err
Dim fila As Integer
Dim xtotal As Double
Dim xdescuento As Double
Dim xneto As Double
Dim ximpuesto As Double
Dim xsubtotal As Double
Dim xc1 As Double
Dim xc2 As Double
Dim xc3 As Double
Dim xc4 As Double
Dim vr
xc1 = 0
xc2 = 0
xc3 = 0
xc4 = 0
xtotal = 0
xdescuento = 0
xneto = 0
ximpuesto = 0
xsubtotal = 0
'dbrecords = Data2.Recordset.RecordCount
'For fila = 0 To DBGrid2.ApproxCount - 1
Data2.Recordset.MoveFirst
Do
If Data2.Recordset.EOF Then Exit Do
xc1 = xc1 + Val("" & Data2.Recordset.Fields("c1"))
xc2 = xc2 + Val("" & Data2.Recordset.Fields("c2"))
xc3 = xc3 + Val("" & Data2.Recordset.Fields("c3"))
xc4 = xc4 + Val("" & Data2.Recordset.Fields("c4"))

xtotal = xtotal + Val("" & Data2.Recordset.Fields("total"))
xdescuento = xdescuento + Val("" & Data2.Recordset.Fields("descuento"))
xneto = xneto + Val("" & Data2.Recordset.Fields("neto"))
ximpuesto = ximpuesto + Val("" & Data2.Recordset.Fields("impuesto"))
xsubtotal = xsubtotal + Val("" & Data2.Recordset.Fields("subtotal"))
Data2.Recordset.MoveNext
Loop
txtotal = Format(xtotal, "0.00")
txdescuento = Format(xdescuento, "0.00")
txneto = Format(xneto, "0.00")
tximpuesto = Format(ximpuesto, "0.00")
txsubtotal = Format(xsubtotal, "0.00")
c1 = Format(xc1, "0.00")
c2 = Format(xc2, "0.00")
c3 = Format(xc3, "0.00")
c4 = Format(xc4, "0.00")
Exit Sub
cmd35_err:
'MsgBox "Error " & Error$ & " " & fila, 24, "Aviso"
Exit Sub

End Sub
Sub poner_ceros()
On Error GoTo cmd789_err
Data2.Recordset.MoveFirst
Do
If Data2.Recordset.EOF Then Exit Do
Data2.Recordset.Edit
Data2.Recordset.Fields("cantidad") = 0
Data2.Recordset.Fields("total") = 0
Data2.Recordset.Update
Data2.Recordset.MoveNext
Loop
txtotal = ""
txdescuento = ""
txneto = ""
tximpuesto = ""
txsubtotal = ""
sumar_detalle
Exit Sub
cmd789_err:
Exit Sub
End Sub
Sub pone_registro_compra(mytablex As Table)

   mytablex.Fields("tipo") = xtipo
   mytablex.Fields("serie") = xserie
   mytablex.Fields("numero") = xnumero
   mytablex.Fields("acu") = xacu
   mytablex.Fields("acu1") = acu
   mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
   mytablex.Fields("fechae") = Format(Now, "dd/mm/yyyy")
   mytablex.Fields("tipoclie") = "P"
   mytablex.Fields("codigo") = "" & Data2.Recordset.Fields("proveedorp")
   mytablex.Fields("nombre") = xnombre1
   mytablex.Fields("estado") = "2"
   mytablex.Fields("partida") = ""
   mytablex.Fields("destino") = ""
   mytablex.Fields("moneda") = xmoneda
   mytablex.Fields("vendedor") = xVendedor
   mytablex.Fields("fpago") = xfpago
   mytablex.Fields("transporte") = ""
   mytablex.Fields("paridad") = 1
   mytablex.Fields("dias") = 1
   mytablex.Fields("bodega") = "01"
   mytablex.Fields("bodegaf") = ""
   mytablex.Fields("observa") = ""
   mytablex.Fields("usuario") = "" & gusuario
   mytablex.Fields("flage") = ""
   mytablex.Fields("hora") = Format(Now, "hh:MM")
   mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")

   mytablex.Fields("total") = Val("" & txtotal)
   mytablex.Fields("descuento") = Val("" & txdescuento)
   mytablex.Fields("neto") = Val("" & txneto)
   mytablex.Fields("impuesto") = Val("" & tximpuesto)
   mytablex.Fields("subtotal") = Val("" & txsubtotal)

   'mytablex.Fields("c1") = Val(c1)
   'mytablex.Fields("c2") = Val(c2)
   'mytablex.Fields("c3") = Val(c3)
   'mytablex.Fields("c4") = Val(c4)

End Sub
Function busca_fpago()
Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("fpago")
mytablex.Index = "fpago"
mytablex.Seek "=", xfpago
If Not mytablex.NoMatch Then
   busca_fpago = 1
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close

End Function
Function busca_vendedor()
Dim mydbx As Database
Dim mytablex As Table
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("fpago")
mytablex.Index = "fpago"
mytablex.Seek "=", xVendedor
If Not mytablex.NoMatch Then
   busca_vendedor = 1
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close

End Function

Private Sub xVendedor_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
xobserva.SetFocus
End Sub
Function busca_proveedor(buf As String)
Dim mydbx As Database
Dim mytablex As Table
xfpago = ""
xdias = ""
xmoneda = ""
xVendedor = ""
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbx.OpenTable("proveedo")
mytablex.Index = "codigo"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   xfpago = "" & mytablex.Fields("fpago")
   xdias = "" & mytablex.Fields("diapago")
   xmoneda = "" & mytablex.Fields("moneda")
   xVendedor = "" & mytablex.Fields("vendedor")
   busca_proveedor = 1
End If
'------------------------------------- ------------
mytablex.Close
mydbx.Close

End Function

Private Sub xVendedor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_vendedor
End If

End Sub
Sub borrar_grid3()
Dim i As Integer
Dim j As Integer

With DBGrid3
For i = 0 To .Row - 1
For j = 0 To .Col - 1
.Row = i
.Col = j
.Text = ""
Next j
Next i
End With

End Sub
Sub carga_dbgrid3()
Dim i As Integer
Dim mydbx As Database
Dim mytablex As Table
Dim mytabley As Table
Dim sw As Integer
Dim xbodega As String
Dim xsaldo As Double
Dim xbuf As String
Dim xcosto As Double
Dim xmargen As Double
For i = 0 To 9
    campo_precios(i).unidad = ""
    campo_precios(i).factor = ""
    campo_precios(i).precio = ""
    campo_precios(i).costo = ""
    campo_precios(i).margen = ""
    campo_precios(i).stock = ""
Next i
xbodega = "01"
xsaldo = 0
xcosto = 0
sw = 0
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytabley = mydbx.OpenTable("parame")
mytabley.Index = "codigo"
mytabley.Seek "=", "01"
If Not mytabley.NoMatch Then
   xbodega = "" & mytabley.Fields("bodega")
End If
mytabley.Close
Set mytabley = mydbx.OpenTable("almacen")
mytabley.Index = "almacen"
mytabley.Seek "=", xproducto, xbodega
If Not mytabley.NoMatch Then
   xsaldo = Val("" & mytabley.Fields("saldo"))
End If
mytabley.Close


Set mytablex = mydbx.OpenTable("producto")
mytablex.Index = "producto"
mytablex.Seek "=", xproducto
If Not mytablex.NoMatch Then
   xcosto = 0
   If Val("" & mytablex.Fields("factor1")) > 0 Then
   xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
   xcosto = xcosto * Val("" & mytablex.Fields("factor1"))
   End If
   '----------------
   '----------------
   campo_precios(0).unidad = "" & mytablex.Fields("unidad1")
   campo_precios(0).factor = "" & mytablex.Fields("factor1")
   campo_precios(0).precio = "" & mytablex.Fields("pventa1")
   campo_precios(0).costo = "" & xcosto
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor1")))
   campo_precios(0).stock = "" & xbuf
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa1")) - xcosto) * 100) / xcosto
   End If
   campo_precios(0).margen = "" & xmargen
   '--------
   
   '---------
   
   campo_precios(1).unidad = "" & mytablex.Fields("unidad2")
   campo_precios(1).factor = "" & mytablex.Fields("factor2")
   campo_precios(1).precio = "" & mytablex.Fields("pventa2")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor2")))
   campo_precios(1).stock = "" & xbuf
   xcosto = 0
   If Val("" & mytablex.Fields("factor2")) > 0 Then
   xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
   xcosto = xcosto * Val("" & mytablex.Fields("factor2"))
   End If
   campo_precios(1).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa2")) - xcosto) * 100) / xcosto
   End If
   campo_precios(1).margen = "" & xmargen
   
   campo_precios(2).unidad = "" & mytablex.Fields("unidad3")
   campo_precios(2).factor = "" & mytablex.Fields("factor3")
   campo_precios(2).precio = "" & mytablex.Fields("pventa3")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor3")))
   campo_precios(2).stock = "" & xbuf
   xcosto = 0
   If Val("" & mytablex.Fields("factor3")) > 0 Then
   xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
   xcosto = xcosto * Val("" & mytablex.Fields("factor3"))
   End If
   campo_precios(2).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa3")) - xcosto) * 100) / xcosto
         campo_precios(2).margen = "" & xmargen
   End If
   campo_precios(2).margen = "" & xmargen
   
   campo_precios(3).unidad = "" & mytablex.Fields("unidad4")
   campo_precios(3).factor = "" & mytablex.Fields("factor4")
   campo_precios(3).precio = "" & mytablex.Fields("pventa4")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor4")))
   campo_precios(3).stock = "" & xbuf
   xcosto = 0
   If Val("" & mytablex.Fields("factor4")) > 0 Then
   xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
   xcosto = xcosto * Val("" & mytablex.Fields("factor4"))
   End If
   campo_precios(4).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa4")) - xcosto) * 100) / xcosto
   End If
   campo_precios(3).margen = "" & xmargen
   
   campo_precios(4).unidad = "" & mytablex.Fields("unidad5")
   campo_precios(4).factor = "" & mytablex.Fields("factor5")
   campo_precios(4).precio = "" & mytablex.Fields("pventa5")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor5")))
   campo_precios(4).stock = "" & xbuf
   xcosto = 0
   If Val("" & mytablex.Fields("factor5")) > 0 Then
   xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
   xcosto = xcosto * Val("" & mytablex.Fields("factor5"))
   End If
   campo_precios(4).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa5")) - xcosto) * 100) / xcosto
         
   End If
   campo_precios(4).margen = "" & xmargen
   
   campo_precios(5).unidad = "" & mytablex.Fields("unidad6")
   campo_precios(5).factor = "" & mytablex.Fields("factor6")
   campo_precios(5).precio = "" & mytablex.Fields("pventa6")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor6")))
   campo_precios(5).stock = "" & xbuf
   xcosto = 0
   If Val("" & mytablex.Fields("factor6")) > 0 Then
   xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
   xcosto = xcosto * Val("" & mytablex.Fields("factor6"))
   End If
   campo_precios(5).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa5")) - xcosto) * 100) / xcosto
         
   End If
   campo_precios(5).margen = "" & xmargen
   
   campo_precios(6).unidad = "" & mytablex.Fields("unidad7")
   campo_precios(6).factor = "" & mytablex.Fields("factor7")
   campo_precios(6).precio = "" & mytablex.Fields("pventa7")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor7")))
   campo_precios(6).stock = "" & xbuf
   xcosto = 0
   If Val("" & mytablex.Fields("factor7")) > 0 Then
   xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
   xcosto = xcosto * Val("" & mytablex.Fields("factor7"))
   End If
   campo_precios(6).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa7")) - xcosto) * 100) / xcosto
         
   End If
   campo_precios(6).margen = "" & xmargen
   
   campo_precios(7).unidad = "" & mytablex.Fields("unidad8")
   campo_precios(7).factor = "" & mytablex.Fields("factor8")
   campo_precios(7).precio = "" & mytablex.Fields("pventa8")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor8")))
   campo_precios(7).stock = "" & xbuf
   xcosto = 0
   If Val("" & mytablex.Fields("factor8")) > 0 Then
      xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
      xcosto = xcosto * Val("" & mytablex.Fields("factor8"))
   End If
   campo_precios(7).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
      xmargen = ((Val("" & mytablex.Fields("pventa8")) - xcosto) * 100) / xcosto
   End If
   campo_precios(7).margen = "" & xmargen
   
   campo_precios(8).unidad = "" & mytablex.Fields("unidad9")
   campo_precios(8).factor = "" & mytablex.Fields("factor9")
   campo_precios(8).precio = "" & mytablex.Fields("pventa9")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor9")))
   campo_precios(8).stock = "" & xbuf
   xcosto = 0
   If Val("" & mytablex.Fields("factor9")) > 0 Then
   xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
   xcosto = xcosto * Val("" & mytablex.Fields("factor9"))
   End If
   campo_precios(8).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa9")) - xcosto) * 100) / xcosto
         
   End If
   campo_precios(8).margen = "" & xmargen
   
   campo_precios(9).unidad = "" & mytablex.Fields("unidad10")
   campo_precios(9).factor = "" & mytablex.Fields("factor10")
   campo_precios(9).precio = "" & mytablex.Fields("pventa10")
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor10")))
   campo_precios(9).stock = "" & xbuf
   xcosto = 0
   If Val("" & mytablex.Fields("factor10")) > 0 Then
   xcosto = Val("" & mytablex.Fields("costou")) / Val("" & mytablex.Fields("factor"))
   xcosto = xcosto * Val("" & mytablex.Fields("factor10"))
   End If
   campo_precios(9).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((Val("" & mytablex.Fields("pventa10")) - xcosto) * 100) / xcosto
   End If
   campo_precios(9).margen = "" & xmargen
      
   'margenes
   
   
   
   sw = 1
End If
mytablex.Close
mydbx.Close
DBGrid3.Refresh
Frame5.Visible = True
DBGrid3.SetFocus
End Sub
