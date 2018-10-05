VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ttparame 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametros Generales"
   ClientHeight    =   10410
   ClientLeft      =   150
   ClientTop       =   -60
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Parametros Generales de la Contabilidad"
      Height          =   10260
      Left            =   12720
      TabIndex        =   103
      Top             =   360
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   10080
         Picture         =   "ttparame.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Imprimir todo"
         Top             =   6960
         Width           =   1470
      End
      Begin VB.TextBox cuentacierre 
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
         Left            =   2760
         MaxLength       =   14
         TabIndex        =   106
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox cuentacapital 
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
         Left            =   2760
         MaxLength       =   14
         TabIndex        =   105
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox periodocontable 
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
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   104
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label36 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ejem. Cuenta de Cierre:540505  Cuenta Dentro del Patrimonio: 3605 o 3610"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   114
         Top             =   4680
         Width           =   10575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"ttparame.frx":08CA
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   113
         Top             =   2400
         Width           =   10575
      End
      Begin VB.Label Label33 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta Cierre Ejercicio"
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
         Left            =   120
         TabIndex        =   112
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta de Capital"
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
         Left            =   120
         TabIndex        =   111
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"ttparame.frx":0995
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4560
         TabIndex        =   110
         Top             =   1560
         Width           =   6135
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mes Contable"
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
         Left            =   120
         TabIndex        =   109
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"ttparame.frx":0A25
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4560
         TabIndex        =   108
         Top             =   480
         Width           =   6135
      End
   End
   Begin VB.Frame Frame3 
      Height          =   6735
      Left            =   12360
      TabIndex        =   52
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
      Begin VB.TextBox clientes 
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
         TabIndex        =   62
         Top             =   120
         Width           =   975
      End
      Begin VB.TextBox paricomp 
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
         TabIndex        =   61
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox parivta 
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
         TabIndex        =   60
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox proveedo 
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
         TabIndex        =   59
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox insumo 
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
         MaxLength       =   15
         TabIndex        =   58
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox plocal 
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
         MaxLength       =   1
         TabIndex        =   57
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox conteo 
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
         TabIndex        =   56
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox banco 
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
         TabIndex        =   55
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox prehora 
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
         TabIndex        =   54
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox aduana 
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
         TabIndex        =   53
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clientes"
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
         Left            =   120
         TabIndex        =   72
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Compras T/C"
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
         Left            =   120
         TabIndex        =   71
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ventas T/C"
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
         Left            =   120
         TabIndex        =   70
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Proveedor"
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
         Left            =   120
         TabIndex        =   69
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Insumo"
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
         Left            =   120
         TabIndex        =   68
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label Label27 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "pLocal"
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
         Left            =   120
         TabIndex        =   67
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label28 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Conteo"
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
         Left            =   120
         TabIndex        =   66
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label29 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Banco"
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
         Left            =   120
         TabIndex        =   65
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Prehora"
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
         Left            =   120
         TabIndex        =   64
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label Label31 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aduana"
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
         Left            =   120
         TabIndex        =   63
         Top             =   4440
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Frame2"
      Height          =   10335
      Left            =   -120
      TabIndex        =   11
      Top             =   -120
      Visible         =   0   'False
      Width           =   12495
      Begin VB.Frame Frame6 
         BackColor       =   &H00808080&
         Caption         =   "Nuevo"
         Height          =   3495
         Left            =   3240
         TabIndex        =   131
         Top             =   5640
         Width           =   9015
         Begin VB.ComboBox nuevoproducto 
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
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   154
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox cambiadescripcion 
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
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   153
            Top             =   360
            Width           =   1215
         End
         Begin VB.ComboBox colorproductofamilia 
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
            Left            =   7575
            Style           =   2  'Dropdown List
            TabIndex        =   151
            Top             =   2280
            Width           =   1215
         End
         Begin VB.ComboBox vemesa 
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
            Left            =   2300
            Style           =   2  'Dropdown List
            TabIndex        =   150
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox tiporeceta 
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
            TabIndex        =   149
            Top             =   1440
            Width           =   2520
         End
         Begin VB.ComboBox tcostoreceta 
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
            Left            =   2300
            Style           =   2  'Dropdown List
            TabIndex        =   147
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox tamanocomanda 
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
            Left            =   2300
            Style           =   2  'Dropdown List
            TabIndex        =   146
            Top             =   2760
            Width           =   1935
         End
         Begin VB.ComboBox OpcionNombre 
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
            Left            =   2300
            Style           =   2  'Dropdown List
            TabIndex        =   144
            Top             =   2400
            Width           =   1935
         End
         Begin VB.ComboBox Comanda 
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
            Left            =   2300
            Style           =   2  'Dropdown List
            TabIndex        =   143
            Top             =   2040
            Width           =   1935
         End
         Begin VB.ComboBox formatocierre 
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
            Left            =   2300
            Style           =   2  'Dropdown List
            TabIndex        =   142
            Top             =   360
            Width           =   2520
         End
         Begin VB.ComboBox opcionexportacion 
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
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   140
            Top             =   3000
            Width           =   1215
         End
         Begin VB.ComboBox EstadoSistema 
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
            Left            =   7560
            Style           =   2  'Dropdown List
            TabIndex        =   138
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label Label58 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "* Cambiar descripción Producto en ventas            * ¿Agregarlo como nuevo producto?"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   5160
            TabIndex        =   152
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label59 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CostoProductos en receta:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   120
            TabIndex        =   148
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label Label55 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tamaño comanda"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   145
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label Label66 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ventana de Ventas:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   141
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label Label65 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Factura de Exportación?"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5160
            TabIndex        =   139
            Top             =   3000
            Width           =   2370
         End
         Begin VB.Label Label64 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ESTADO SISTEMA"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5160
            TabIndex        =   137
            Top             =   2640
            Width           =   2355
         End
         Begin VB.Label Label63 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Comanda:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   136
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label Label62 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Color Producto=Color Familia?"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5160
            TabIndex        =   135
            Top             =   2280
            Width           =   2325
         End
         Begin VB.Label lblHabilitarSubfamilia 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ve OpcionMesa en Personal"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   120
            TabIndex        =   134
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label54 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Formato Cierre"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   120
            TabIndex        =   133
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label57 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tipo Receta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   380
            Left            =   120
            TabIndex        =   132
            Top             =   1440
            Width           =   2175
         End
      End
      Begin VB.TextBox cajapedido 
         Height          =   495
         Left            =   6240
         MaxLength       =   1
         TabIndex        =   129
         Top             =   9720
         Width           =   495
      End
      Begin VB.Frame Frame5 
         Height          =   4095
         Left            =   0
         TabIndex        =   115
         Top             =   0
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox saldocierre 
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
            MaxLength       =   1
            TabIndex        =   127
            Top             =   1680
            Width           =   495
         End
         Begin VB.TextBox solohuella 
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
            MaxLength       =   1
            TabIndex        =   124
            Top             =   2160
            Width           =   495
         End
         Begin VB.TextBox correocierre 
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
            MaxLength       =   1
            TabIndex        =   122
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox centralizacierre 
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
            MaxLength       =   1
            TabIndex        =   120
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox caduca 
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
            TabIndex        =   117
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox estadocaduca 
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
            MaxLength       =   1
            TabIndex        =   116
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblInventarioAl 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo al Cierre"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   126
            Top             =   1680
            Width           =   1500
         End
         Begin VB.Label Label53 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SoloHuella"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   125
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label Label52 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Correo al Cierre"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   123
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label51 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Centraliza al Cierre"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   121
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label45 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "fecha caducidad"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   119
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label46 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Activo Caducidad"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   118
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.TextBox subfamilia 
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
         Left            =   6960
         MaxLength       =   1
         TabIndex        =   101
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox mesaseccion 
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
         Left            =   6960
         MaxLength       =   8
         TabIndex        =   99
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox diasemana 
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
         Left            =   6960
         MaxLength       =   1
         TabIndex        =   97
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox selemesa 
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
         Left            =   5040
         MaxLength       =   1
         TabIndex        =   95
         Top             =   5160
         Width           =   495
      End
      Begin VB.TextBox vdolar 
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
         Left            =   8160
         MaxLength       =   1
         TabIndex        =   93
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox tipoprecio 
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
         MaxLength       =   1
         TabIndex        =   91
         Top             =   7680
         Width           =   615
      End
      Begin VB.TextBox mesavendedor 
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
         MaxLength       =   1
         TabIndex        =   89
         Top             =   9840
         Width           =   615
      End
      Begin VB.TextBox mesabierta 
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
         MaxLength       =   1
         TabIndex        =   87
         Top             =   9480
         Width           =   615
      End
      Begin VB.TextBox ingreso 
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
         Left            =   6960
         MaxLength       =   10
         TabIndex        =   85
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox centraliza 
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
         Left            =   5040
         MaxLength       =   1
         TabIndex        =   83
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox tarjetacredito 
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
         MaxLength       =   2
         TabIndex        =   81
         Top             =   9120
         Width           =   615
      End
      Begin VB.TextBox credito 
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
         MaxLength       =   2
         TabIndex        =   79
         Top             =   8760
         Width           =   615
      End
      Begin VB.TextBox dolares 
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
         MaxLength       =   2
         TabIndex        =   77
         Top             =   8400
         Width           =   615
      End
      Begin VB.TextBox efectivo 
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
         MaxLength       =   2
         TabIndex        =   75
         Top             =   8040
         Width           =   615
      End
      Begin VB.TextBox listaprecios 
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
         Left            =   6960
         MaxLength       =   6
         TabIndex        =   73
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox wordpad 
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
         Left            =   5400
         MaxLength       =   100
         TabIndex        =   50
         Top             =   9360
         Width           =   6015
      End
      Begin VB.TextBox grupoproducto 
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
         Left            =   5040
         MaxLength       =   1
         TabIndex        =   48
         Top             =   4440
         Width           =   495
      End
      Begin VB.TextBox tipo5 
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
         MaxLength       =   1
         TabIndex        =   46
         Top             =   6720
         Width           =   375
      End
      Begin VB.TextBox touch 
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
         Left            =   9120
         MaxLength       =   1
         TabIndex        =   44
         Top             =   5160
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Contable"
         Height          =   975
         Left            =   10080
         Picture         =   "ttparame.frx":0AB0
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2520
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox codigo 
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
         MaxLength       =   2
         TabIndex        =   28
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox descripcio 
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
         TabIndex        =   27
         Top             =   840
         Width           =   6015
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   10080
         Picture         =   "ttparame.frx":137A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Imprimir todo"
         Top             =   1440
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   10080
         Picture         =   "ttparame.frx":1C44
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   1470
      End
      Begin VB.TextBox ocurrencia 
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
         Left            =   8160
         MaxLength       =   10
         TabIndex        =   24
         Top             =   4800
         Width           =   1335
      End
      Begin VB.TextBox pocket 
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
         Left            =   8160
         MaxLength       =   10
         TabIndex        =   23
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox SERVIDOR 
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
         MaxLength       =   1
         TabIndex        =   22
         Top             =   5160
         Width           =   495
      End
      Begin VB.TextBox tradiario 
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
         MaxLength       =   1
         TabIndex        =   21
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox imp_und 
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
         MaxLength       =   1
         TabIndex        =   20
         Top             =   4440
         Width           =   495
      End
      Begin VB.TextBox pedauto 
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
         TabIndex        =   19
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox deliveri 
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
         TabIndex        =   18
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox cabecera2 
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
         MaxLength       =   60
         TabIndex        =   17
         Top             =   4080
         Width           =   5775
      End
      Begin VB.TextBox cabecera1 
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
         MaxLength       =   60
         TabIndex        =   16
         Top             =   3720
         Width           =   5775
      End
      Begin VB.TextBox produccion 
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
         TabIndex        =   15
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox bodega 
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
         MaxLength       =   2
         TabIndex        =   14
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox igv 
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
         TabIndex        =   13
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox producto 
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
         TabIndex        =   12
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Paranetriza Caja para  pedido de OT?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3240
         TabIndex        =   128
         Top             =   9720
         Width           =   2955
      End
      Begin VB.Label Label50 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Habilitar Subfamilia Caja"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   102
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label49 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cambiar Mesa x Seccion?"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   100
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label48 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja:SoloDiaSemana"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   98
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label47 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Definir Mesas x Mesero"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   96
         Top             =   5160
         Width           =   2175
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ve Dolar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   94
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label44 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Precio Producto  []Inlcuido [N]o incluido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   92
         Top             =   7320
         Width           =   2175
      End
      Begin VB.Label Label43 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ObligarMismoMeseroComanda"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   90
         Top             =   9840
         Width           =   2175
      End
      Begin VB.Label Label42 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mesas Abiertas - No cerrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   88
         Top             =   9480
         Width           =   2175
      End
      Begin VB.Label Label41 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Incluye IngresoEgreso  Cuadre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   86
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Centraliza"
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
         Left            =   2880
         TabIndex        =   84
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label40 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tarjeta Credito"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   82
         Top             =   9120
         Width           =   2175
      End
      Begin VB.Label Label39 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Credito"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   80
         Top             =   8760
         Width           =   2175
      End
      Begin VB.Label Dolaresh 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dolares"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   78
         Top             =   8400
         Width           =   2175
      End
      Begin VB.Label Label38 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Efectivo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   76
         Top             =   8040
         Width           =   2175
      End
      Begin VB.Label Label37 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lista Precios ..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   74
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label pathwordpad 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Editor Texto"
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
         Left            =   3240
         TabIndex        =   51
         Top             =   9360
         Width           =   2175
      End
      Begin VB.Label Label35 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grupo default P/F"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   49
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Visualiza Tipo 5"
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
         Left            =   120
         TabIndex        =   47
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label Label34 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acceso Rapido 1.RestTouch 2.MarketTouch"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   45
         Top             =   5160
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Parametro Nro"
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
         TabIndex        =   42
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
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
         TabIndex        =   41
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OcurrenciasGrabadas"
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
         Left            =   5760
         TabIndex        =   40
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CorrelativoTicketIng"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   39
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Servidor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   5160
         Width           =   2175
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Transaccion (D)iario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Veprodc.Cuadre(1=s)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDoc.Ped.Autom"
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
         Left            =   120
         TabIndex        =   35
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDoc.Deliveri"
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
         Left            =   120
         TabIndex        =   34
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cabecera_Reporte"
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
         Left            =   120
         TabIndex        =   33
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Orden Produccion"
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
         Left            =   120
         TabIndex        =   32
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bodega"
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
         Left            =   120
         TabIndex        =   31
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Igv"
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
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
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
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   12435
      TabIndex        =   2
      Top             =   0
      Width           =   12495
      Begin VB.CommandButton cmdSave 
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
         Picture         =   "ttparame.frx":250E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   735
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
         Height          =   375
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Buscar"
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
         Left            =   10560
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "ttparame.frx":3720
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Borrar registro"
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
         Left            =   2760
         MaskColor       =   &H00E0E0E0&
         Picture         =   "ttparame.frx":4932
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
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
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "ttparame.frx":5B44
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdAddEntry 
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
         Picture         =   "ttparame.frx":6D56
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7935
         Left            =   0
         TabIndex        =   1
         Top             =   720
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   13996
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   22
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "Codigo"
            Caption         =   "Codigo"
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
            DataField       =   "Periodocontable"
            Caption         =   "Periodocontable"
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
         BeginProperty Column03 
            DataField       =   "CuentaCierre"
            Caption         =   "CuentaCierre"
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
         BeginProperty Column04 
            DataField       =   "CuentaCapitalu"
            Caption         =   "CuentaCapital"
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
               ColumnWidth     =   5940.284
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label60 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(N) Normal/Sunat    (S)Simple"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   130
      Top             =   7500
      Width           =   2295
   End
   Begin VB.Menu ajdu1 
      Caption         =   "&Add"
   End
   Begin VB.Menu f8443 
      Caption         =   "&Modifica"
   End
   Begin VB.Menu bo712 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu fjh433 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "ttparame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txempre As New ADODB.Recordset

Private Sub ajdu1_Click()

    If Frame4.Visible = True Then Exit Sub
    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Nuevo"
    cmdGuardar.Enabled = True
    habilita 1
    codigo.Enabled = True
    codigo = ""
    codigo.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    If Frame4.Visible = True Then Exit Sub
    buf = txempre.Fields("codigo")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + txempre.Fields("codigo"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    txempre.Delete
    Command1_Click

    Exit Sub
cmd656_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    Command1_Click

End Sub

Private Sub cmdAddEntry_Click()
    ajdu1_Click

End Sub

Private Sub cmdCerrar_Click()
    dlo132_Click

End Sub

Private Sub cmdDelete_Click()
    bo712_Click

End Sub

Private Sub cmdExit_Click()
    dlo132_Click

End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdGuardar_Click()

    Dim found As Integer

    found = grabar()

End Sub

Private Sub cmdPrint_Click()
    djuer1_Click

End Sub

Private Sub cmdSave_Click()
    f8443_Click

End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(codigo) = 0 Then Exit Sub
    descripcio.SetFocus

End Sub

Private Sub Command1_Click()
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    opcion1 = "1"
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim cad As String

    If opcion1 = "1" Then  'bodega
        If Len(buffer) = 0 Then
            cad = "SELECT * from parame    "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT *  from parame   where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If txempre.State = 1 Then txempre.Close
        txempre.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbGrid1.DataSource = txempre
        dbGrid1.columns(0).Width = 4000
        dbGrid1.columns(1).Width = 2000

        If txempre.RecordCount > 0 Then
            dbGrid1.SetFocus

        End If

    End If

End Sub

Private Sub Command2_Click()
    Frame4.Visible = False

End Sub

Private Sub Command3_Click()
    Frame4.Visible = True

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'codigo = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'codigo.SetFocus
        'codigo_KeyPress 13
    End If

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)

    Dim buf  As String

    Dim buf2 As String

    If KeyAscii <> 13 And KeyAscii <> 27 Then
        If KeyAscii = 8 Then
            If Len(buffer) > 0 Then
                buf = Mid$(buffer, 1, Len(buffer) - 1)
                buffer = buf
                KeyAscii = 0
            Else
                KeyAscii = 0
                Exit Sub

            End If

        End If

        buf = Chr(KeyAscii)

        If Chr(KeyAscii) = "*" Then
            buf = ""
            buffer = buf

        End If

        If KeyAscii <> 13 Then
            buffer = buffer + buf

        End If

        buf = buffer
        ejecuta 0
         
    End If

End Sub

Private Sub djuer1_Click()

    If Frame4.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "parame"
    reporgen.Show 1

End Sub

Private Sub dlo132_Click()

    If Frame5.Visible = True Then
        Frame5.Visible = False
        Exit Sub

    End If

    If Frame4.Visible = True Then
        Frame4.Visible = False
        Exit Sub

    End If

    If Frame2.Visible = True Then
        habilita 0
        Frame2.Visible = False
        dbGrid1.Enabled = True
        Exit Sub

    End If

    ttparame.Hide
    Unload ttparame

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    If Frame4.Visible = True Then Exit Sub
    buf = txempre.Fields("codigo")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Modifica"
    cmdGuardar.Enabled = True
    pone_registro
    habilita 1
    codigo.Enabled = False
    descripcio.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    If Frame4.Visible = True Then Exit Sub
    buf = txempre.Fields("codigo")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Zoom"
    cmdGuardar.Enabled = False
    pone_registro
    habilita 1
    codigo.Enabled = False
    descripcio.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    Label4 = dicigv
    Command1_Click

    '''20/01/2018 kenyo Testing General Sistema
    Frame4.Top = 0: Frame4.Left = 0
    Frame3.Top = 0: Frame3.Top = 0
    '''20/01/2018 kenyo Testing General Sistema

End Sub

Private Sub Form_Load()

    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "codigo"
    Combo1.ListIndex = 0

End Sub

Sub inicializa()
    solohuella = ""
    correocierre = ""
    saldocierre = ""
    centralizacierre = ""
    subfamilia = ""
    '''26/10/2017 Listas de ciertas caja para cobrar pedido
    '''09/08/2017 kenyo. Opcion Mesa Personal
    vemesa.Clear
    vemesa.AddItem "S"
    vemesa.AddItem "N"
    vemesa.ListIndex = 0

    formatocierre.Clear
    formatocierre.AddItem "N|Normal/Sunat"
    formatocierre.AddItem "S|Simple/Control Interno"
    formatocierre.ListIndex = 0

    cajapedido = ""
    'formatocierrey = ""
    '''09/08/2017 kenyo. Opcion Mesa Personal
    '''26/10/2017 Listas de ciertas caja para cobrar pedido

    '' 11/12/2017 SubReceta
    tiporeceta.Clear
    tiporeceta.AddItem "E|ESTÁNDAR"
    tiporeceta.AddItem "P|PRODUCCIÓN"
    tiporeceta.ListIndex = 0
    '' 11/12/2017 SubReceta

    'Color por familia y producto  30/05/2018
    colorproductofamilia.Clear
    colorproductofamilia.AddItem "N"
    colorproductofamilia.AddItem "S"
    colorproductofamilia.ListIndex = 0
    'Color por familia y producto  30/05/2018

    '' 10/07/2018 Edicion Comanda
    comanda.Clear
    comanda.AddItem "CO|Codigo"
    comanda.AddItem "DL|DescrLarga"
    comanda.AddItem "DC|DescrCorta"
    comanda.ListIndex = 0

    OpcionNombre.Clear
    OpcionNombre.AddItem "CO|Codigo"
    OpcionNombre.AddItem "DL|DescrLarga"
    OpcionNombre.AddItem "DC|DescrCorta"
    OpcionNombre.ListIndex = 0

    tamanocomanda.Clear
    tamanocomanda.AddItem "12"
    tamanocomanda.AddItem "14"
    tamanocomanda.AddItem "20"
    tamanocomanda.ListIndex = 0

    '' 10/07/2018 Edicion Comanda

    ' 26/07/2018 Desactivar Facturacion Electronica
    EstadoSistema.Clear
    EstadoSistema.AddItem "FE ARIES"
    EstadoSistema.AddItem "FE BYH"
    EstadoSistema.AddItem "CONINT"
    EstadoSistema.ListIndex = 0
    ' 26/07/2018 Desactivar Facturacion Electronica

    ' 17/07/2018 Factura de Exportación
    opcionexportacion.Clear
    opcionexportacion.AddItem "N"
    opcionexportacion.AddItem "S"
    opcionexportacion.ListIndex = 0
    ' 17/07/2018 Factura de Exportación

    '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
    tcostoreceta.Clear
    tcostoreceta.AddItem "CU"
    tcostoreceta.AddItem "CP"
    tcostoreceta.ListIndex = 0
    '11/06/2018 Actualiza Precio Promedio Ponderado Masivo

    '15/08/2018 Cambiar Descripcion de producto venta de ventas
    cambiadescripcion.Clear
    cambiadescripcion.AddItem "N"
    cambiadescripcion.AddItem "S"
    cambiadescripcion.ListIndex = 0

    nuevoproducto.Clear
    nuevoproducto.AddItem "N"
    nuevoproducto.AddItem "S"
    nuevoproducto.ListIndex = 0

    '15/08/2018 Cambiar Descripcion de producto venta de ventas

    mesaseccion = ""
    diasemana = ""
    selemesa = ""
    caduca = ""
    estadocaduca = ""
    tipoprecio = ""
    mesavendedor = ""
    mesabierta = ""
    ingreso = ""
    efectivo = ""
    dolares = ""
    credito = ""
    tarjetacredito = ""

    listaprecios = ""
    wordpad = ""
    periodocontable = ""
    grupoproducto = ""
    touch = ""
    cuentacierre = ""
    cuentacapital = ""
    clientes = ""
    paricomp = ""
    parivta = ""
    proveedo = ""
    insumo = ""
    centraliza = ""
    vdolar = ""
    tipo5 = ""
    plocal = ""
    conteo = ""
    banco = ""
    prehora = ""
    aduana = ""

    ocurrencia = ""
    pocket = ""
    servidor = ""
    tradiario = ""
    imp_und = ""
    pedauto = ""
    deliveri = ""
    cabecera1 = ""
    cabecera2 = ""
    produccion = ""
    descripcio = ""
    igv = ""
    producto = ""
    bodega = ""
    'saldoini = ""
    'mesconta = ""
    'anoconta = ""

End Sub

Sub pone_registro()
    solohuella = "" & txempre.Fields("solohuella")
    correocierre = "" & txempre.Fields("correocierre")
    saldocierre = "" & txempre.Fields("saldocierre")

    centralizacierre = "" & txempre.Fields("centralizacierre")
    subfamilia = "" & txempre.Fields("subfamilia")

    '''09/08/2017 kenyo. Opcion Mesa Personal
    vemesa.ListIndex = 0

    If "" & txempre.Fields("vemesa") = "N" Then
        vemesa.ListIndex = 1

    End If

    formatocierre.ListIndex = 0

    If "" & txempre.Fields("formatocierre") = "S" Then
        formatocierre.ListIndex = 1

    End If

    cajapedido = "" & txempre.Fields("cajapedido")
    '''09/08/2017 kenyo. Opcion Mesa Personal

    '''26/10/2017 Listas de ciertas caja para cobrar pedido

    '' 11/12/2017 SubReceta
    tiporeceta.ListIndex = 0

    If "" & txempre.Fields("tiporeceta") = "P" Then
        tiporeceta.ListIndex = 1

    End If

    '' 11/12/2017 SubReceta

    'Color por familia y producto  30/05/2018
    colorproductofamilia.ListIndex = 0

    If "" & txempre.Fields("colorproductofamilia") = "S" Then
        colorproductofamilia.ListIndex = 1

    End If

    'Color por familia y producto  30/05/2018

    '' 10/07/2018 Edicion Comanda
    'Comanda.ListIndex = 0
    If "" & txempre.Fields("Comanda") = "CO" Then
        comanda.ListIndex = 0
    ElseIf "" & txempre.Fields("Comanda") = "DC" Then
        comanda.ListIndex = 2
    Else
        comanda.ListIndex = 1

    End If

    '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
    tcostoreceta.ListIndex = 0

    If "" & txempre.Fields("tcostoreceta") = "CP" Then
        tcostoreceta.ListIndex = 1

    End If

    '11/06/2018 Actualiza Precio Promedio Ponderado Masivo

    If "" & txempre.Fields("opcionnombre") = "CO" Then
        OpcionNombre.ListIndex = 0
    ElseIf "" & txempre.Fields("opcionnombre") = "DC" Then
        OpcionNombre.ListIndex = 2
    Else
        OpcionNombre.ListIndex = 1

    End If

    If "" & txempre.Fields("tamanocomanda") = "14" Then
        tamanocomanda.ListIndex = 1
    ElseIf "" & txempre.Fields("tamanocomanda") = "20" Then
        tamanocomanda.ListIndex = 2
    Else
        tamanocomanda.ListIndex = 0

    End If

    '' 10/07/2018 Edicion Comanda

    ' 26/07/2018 Desactivar Facturacion Electronica
    EstadoSistema.ListIndex = 0

    If "" & txempre.Fields("EstadoSistema") = "CONINT" Then
        EstadoSistema.ListIndex = 1

    End If

    ' 26/07/2018 Desactivar Facturacion Electronica

    ' 17/07/2018 Factura de Exportación
    opcionexportacion.ListIndex = 0

    If "" & txempre.Fields("opcionexportacion") = "S" Then
        opcionexportacion.ListIndex = 1

    End If

    ' 17/07/2018 Factura de Exportación

    '15/08/2018 Cambiar Descripcion de producto venta de ventas
    cambiadescripcion.ListIndex = 0

    If "" & txempre.Fields("cambiadescripcion") = "S" Then
        cambiadescripcion.ListIndex = 1

    End If

    nuevoproducto.ListIndex = 0

    If "" & txempre.Fields("nuevoproducto") = "S" Then
        nuevoproducto.ListIndex = 1

    End If

    '15/08/2018 Cambiar Descripcion de producto venta de ventas

    mesaseccion = "" & txempre.Fields("mesaseccion")
    diasemana = "" & txempre.Fields("diasemana")
    selemesa = "" & txempre.Fields("selemesa")
    caduca = "" & txempre.Fields("caduca")
    estadocaduca = "" & txempre.Fields("estadocaduca")

    tipoprecio = "" & txempre.Fields("tipoprecio")
    mesavendedor = "" & txempre.Fields("mesavendedor")
    mesabierta = "" & txempre.Fields("mesabierta")
    ingreso = "" & txempre.Fields("ingreso")
    efectivo = "" & txempre.Fields("efectivo")
    dolares = "" & txempre.Fields("dolares")
    credito = "" & txempre.Fields("credito")
    tarjetacredito = "" & txempre.Fields("tarjetacredito")

    listaprecios = "" & txempre.Fields("listaprecios")
    wordpad = "" & txempre.Fields("wordpad")
    periodocontable = "" & txempre.Fields("periodocontable")
    grupoproducto = "" & txempre.Fields("grupoproducto")
    touch = "" & txempre.Fields("touch")
    cuentacierre = "" & txempre.Fields("cuentacierre")
    cuentacapital = "" & txempre.Fields("cuentacapitalu")
    clientes = "" & txempre.Fields("clientes")
    paricomp = "" & txempre.Fields("paricomp")
    parivta = "" & txempre.Fields("parivta")
    proveedo = "" & txempre.Fields("proveedo")
    insumo = "" & txempre.Fields("insumo")
    centraliza = "" & txempre.Fields("centraliza")
    vdolar = "" & txempre.Fields("vdolar")
    tipo5 = "" & txempre.Fields("tipo5")
    plocal = "" & txempre.Fields("plocal")
    conteo = "" & txempre.Fields("conteo")
    banco = "" & txempre.Fields("banco")
    prehora = "" & txempre.Fields("prehora")
    aduana = "" & txempre.Fields("aduana")

    ocurrencia = "" & txempre.Fields("ocurrencia")
    pocket = "" & txempre.Fields("pocket")
    tradiario = "" & txempre.Fields("tradiario")
    imp_und = "" & txempre.Fields("imp_und")
    pedauto = "" & txempre.Fields("pedauto")
    deliveri = "" & txempre.Fields("deliveri")
    cabecera1 = "" & txempre.Fields("cabecera1")
    cabecera2 = "" & txempre.Fields("cabecera2")
    produccion = "" & txempre.Fields("produccion")
    codigo = "" & txempre.Fields("codigo")
    descripcio = "" & txempre.Fields("descripcio")
    igv = "" & txempre.Fields("igv")
    producto = "" & txempre.Fields("producto")
    bodega = "" & txempre.Fields("bodega")

    'saldoini = "" & txempre.Fields("saldoini")
    'mesconta = "" & txempre.Fields("mesconta")
    'anoconta = "" & txempre.Fields("anoconta")
End Sub

Sub grabando()
    txempre.Fields("solohuella") = Trim(solohuella)
    txempre.Fields("correocierre") = Trim(correocierre)
    txempre.Fields("saldocierre") = Trim(saldocierre)
    txempre.Fields("centralizacierre") = Trim(centralizacierre)

    txempre.Fields("subfamilia") = Trim(subfamilia)
    '''18/09/2017 KENYO Formato Simple Cierre X & Y
    '''09/08/2017 kenyo. Opcion Mesa Personal
    txempre.Fields("vemesa") = Trim(vemesa)

    txempre.Fields("formatocierre") = extra_loquesea(Trim(formatocierre))

    txempre.Fields("cajapedido") = Trim(cajapedido)
    'txempre.Fields("formatocierrey") = Trim(formatocierrey)
    '''09/08/2017 kenyo. Opcion Mesa Personal
    '''26/10/2017 Listas de ciertas caja para cobrar pedido

    '' 11/12/2017 SubReceta
    txempre.Fields("tiporeceta") = extra_loquesea(Trim(tiporeceta))
    '' 11/12/2017 SubReceta

    'Color por familia y producto  30/05/2018
    txempre.Fields("colorproductofamilia") = Trim(colorproductofamilia)
    'Color por familia y producto  30/05/2018

    '' 10/07/2018 Edicion Comanda
    txempre.Fields("comanda") = extra_loquesea(Trim(comanda))
    txempre.Fields("opcionnombre") = extra_loquesea(Trim(OpcionNombre))
    txempre.Fields("tamanocomanda") = Trim(tamanocomanda)
    '' 10/07/2018 Edicion Comanda

    ' 26/07/2018 Desactivar Facturacion Electronica
    txempre.Fields("EstadoSistema") = Trim(EstadoSistema)
    ' 26/07/2018 Desactivar Facturacion Electronica

    ' 17/07/2018 Factura de Exportación
    txempre.Fields("opcionexportacion") = Trim(opcionexportacion)
    ' 17/07/2018 Factura de Exportación

    '11/06/2018 Actualiza Precio Promedio Ponderado Masivo
    txempre.Fields("tcostoreceta") = Trim(tcostoreceta)
    '11/06/2018 Actualiza Precio Promedio Ponderado Masivo

    '15/08/2018 Cambiar Descripcion de producto venta de ventas
    txempre.Fields("cambiadescripcion") = Trim(cambiadescripcion)
    txempre.Fields("nuevoproducto") = Trim(nuevoproducto)
    '15/08/2018 Cambiar Descripcion de producto venta de ventas

    If Len(Trim(mesaseccion)) = 0 Then
        mesaseccion = "MESA"

    End If

    txempre.Fields("mesaseccion") = Trim(mesaseccion)
    txempre.Fields("selemesa") = Trim(selemesa)

    If IsDate(Trim(caduca)) Then
        txempre.Fields("caduca") = Trim(caduca)

    End If

    txempre.Fields("diasemana") = Trim(diasemana)
    txempre.Fields("estadocaduca") = Trim(estadocaduca)
    txempre.Fields("tipoprecio") = Trim(tipoprecio)
    txempre.Fields("mesavendedor") = Trim(mesavendedor)
    txempre.Fields("mesabierta") = Trim(mesabierta)
    txempre.Fields("ingreso") = Trim(ingreso)
    txempre.Fields("efectivo") = Trim(efectivo)
    txempre.Fields("dolares") = Trim(dolares)
    txempre.Fields("credito") = Trim(credito)
    txempre.Fields("tarjetacredito") = Trim(tarjetacredito)

    txempre.Fields("listaprecios") = Trim(listaprecios)
    txempre.Fields("wordpad") = Trim(wordpad)
    txempre.Fields("periodocontable") = Trim(periodocontable)
    txempre.Fields("grupoproducto") = Trim(grupoproducto)
    txempre.Fields("touch") = Trim(touch)
    txempre.Fields("cuentacierre") = Trim(cuentacierre)
    txempre.Fields("cuentacapitalu") = Trim(cuentacapital)
    txempre.Fields("codigo") = Trim(codigo)

    txempre.Fields("descripcio") = Trim(descripcio)
    txempre.Fields("igv") = Val(igv)
    txempre.Fields("bodega") = Trim(bodega)
    'txempre.Fields("saldoini") = Trim(saldoini)
    'txempre.Fields("mesconta") = Trim(mesconta)
    'txempre.Fields("anoconta") = Trim(anoconta)
    txempre.Fields("produccion") = Trim(produccion)
    txempre.Fields("cabecera1") = Trim(cabecera1)
    txempre.Fields("cabecera2") = Trim(cabecera2)
    txempre.Fields("deliveri") = Trim(deliveri)
    txempre.Fields("clientes") = Trim(clientes)
    txempre.Fields("paricomp") = Val(paricomp)
    txempre.Fields("parivta") = Val(parivta)
    txempre.Fields("proveedo") = Trim(proveedo)
    txempre.Fields("insumo") = Trim(insumo)
    txempre.Fields("pedauto") = Trim(pedauto)
    txempre.Fields("imp_und") = Trim(imp_und)
    txempre.Fields("tradiario") = Trim(tradiario)
    txempre.Fields("centraliza") = Trim(centraliza)
    txempre.Fields("vdolar") = Trim(vdolar)
    txempre.Fields("tipo5") = Trim(tipo5)
    txempre.Fields("plocal") = Trim(plocal)
    txempre.Fields("conteo") = Trim(conteo)
    txempre.Fields("banco") = Trim(banco)
    txempre.Fields("prehora") = Trim(prehora)
    txempre.Fields("pocket") = Val(pocket)
    txempre.Fields("ocurrencia") = Val(cabecera1)
    'txempre.Fields("aduana") = Trim(cabecera1)

End Sub

Private Sub grba1_Click()

End Sub

Function grabar()

    Dim found  As Integer

    Dim rbusca As New ADODB.Recordset

    found = valida()

    If found = 0 Then
        MsgBox "Campos invalidos", 48, "Aviso"
        Exit Function

    End If

    If Frame2.Caption = "Nuevo" Then
        If Len(codigo) = 0 Then
            codigo.SetFocus
            Exit Function

        End If

        rbusca.Open "select codigo from parame where codigo='" & codigo & "'", cn, adOpenStatic, adLockOptimistic

        If rbusca.RecordCount > 0 Then
            rbusca.Close
            MsgBox "Ya existe codigo ", 48, "Aviso"
            Exit Function

        End If

        txempre.AddNew
        txempre.Fields("codigo") = codigo
        grabando
        txempre.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        txempre.Fields("codigo") = codigo
        grabando
        txempre.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    'If Len(codigo) = 0 Then
    '   codigo.SetFocus
    '   Exit Function
    'End If
    If Len(periodocontable) > 0 Then
        If Len(periodocontable) <> 10 Then
            periodocontable = ""

        End If

        If Len(periodocontable) = 10 Then
            If Not IsDate(periodocontable) Then
                periodocontable = ""

            End If

        End If

        If Len(periodocontable) = 10 Then
            periodocontable = Format(periodocontable, "dd/mm/yyyy")

        End If

    End If

    If Len(descripcio) = 0 Then
        descripcio.SetFocus
        Exit Function

    End If

    valida = 1

End Function

Sub habilita(sw As Integer)

    If sw = 0 Then

        ajdu1.Enabled = True
        f8443.Enabled = True
        bo712.Enabled = True
        fjh433.Enabled = True
        djuer1.Enabled = True
        djuer1.Enabled = True
        Picture1.Enabled = True
        dbGrid1.Enabled = True

    End If

    If sw = 1 Then

        ajdu1.Enabled = False
        f8443.Enabled = False
        bo712.Enabled = False
        fjh433.Enabled = False
        djuer1.Enabled = False
        djuer1.Enabled = False
        Picture1.Enabled = False
        dbGrid1.Enabled = False
        dbGrid1.Enabled = False

    End If

End Sub

Private Sub Label38_Click()

    If Frame5.Visible = False Then
        If InputBox("Clave de Paso", "Aviso") = "KALIPO" Then
            Frame5.Visible = True

        End If

    End If

End Sub

