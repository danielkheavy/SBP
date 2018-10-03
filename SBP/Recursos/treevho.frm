VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form treevho 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Panel de Control Hotel"
   ClientHeight    =   9285
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   16950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   16950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Habitacion"
      Height          =   9255
      Left            =   13080
      TabIndex        =   84
      Top             =   3960
      Visible         =   0   'False
      Width           =   12735
      Begin VB.CommandButton Command18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NuevaEntrada"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox arribofecha 
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   112
         Top             =   4800
         Width           =   1935
      End
      Begin VB.TextBox arribofechaf 
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   111
         Top             =   5160
         Width           =   1935
      End
      Begin VB.TextBox nombre 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   80
         TabIndex        =   110
         Top             =   3720
         Width           =   5655
      End
      Begin VB.TextBox arribohora 
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   109
         Top             =   5520
         Width           =   1935
      End
      Begin VB.TextBox codigo 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   11
         TabIndex        =   108
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox noches 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   107
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox huesped 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   106
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox hnombre 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   105
         Top             =   2640
         Width           =   5655
      End
      Begin VB.TextBox personas 
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   104
         Top             =   7320
         Width           =   735
      End
      Begin VB.TextBox tipotarifa 
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   103
         Top             =   6240
         Width           =   1935
      End
      Begin VB.TextBox tipopension 
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   102
         Top             =   6600
         Width           =   1935
      End
      Begin VB.TextBox precio 
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   101
         Top             =   6960
         Width           =   1215
      End
      Begin VB.TextBox categoria 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   100
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox arribohoraf 
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   99
         Top             =   5880
         Width           =   1935
      End
      Begin VB.TextBox estado 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   7
         TabIndex        =   98
         Top             =   7680
         Width           =   1215
      End
      Begin VB.TextBox tipocodigo 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   97
         Top             =   3360
         Width           =   495
      End
      Begin VB.TextBox tipocodigoh 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
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
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   96
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nueva Reserva"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pagos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consumos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Precuenta"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Facturacion"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   6960
         Width           =   1575
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Entrada"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Incidencias"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Limpieza"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8040
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Id CheckIn"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   0
         TabIndex        =   142
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label34 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Totalhabitacion"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   141
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label33 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Abonos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   140
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Permanencia"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   139
         Top             =   4440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label31 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Consumos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   138
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   137
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label Label27 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaEntrada"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   136
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   135
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HoraLLegada(HH:MM:SS)"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   134
         Top             =   5520
         Width           =   2175
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo(QuienPaga)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   133
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label23 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NroDias"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   132
         Top             =   4440
         Width           =   2175
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombres"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   131
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo (El que se aloja)"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   130
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label38 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NroPersonas"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   129
         Top             =   7320
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaSalida"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   128
         Top             =   5160
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo de Tarifa"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   127
         Top             =   6240
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Pension"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   126
         Top             =   6600
         Width           =   2175
      End
      Begin VB.Label Label26 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   125
         Top             =   6960
         Width           =   2175
      End
      Begin VB.Label Label29 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HoraSalida(HH:MM:SS)"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   124
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Label Label30 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Categoria"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   123
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   122
         Top             =   7680
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocumento"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   121
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label25 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocumento"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   120
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label xpermanencia 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5880
         TabIndex        =   119
         Top             =   4440
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label xcheckin 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2160
         TabIndex        =   118
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label xabonos 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5880
         TabIndex        =   117
         Top             =   5520
         Width           =   1185
      End
      Begin VB.Label xtotal 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5880
         TabIndex        =   116
         Top             =   4800
         Width           =   1185
      End
      Begin VB.Label xconsumos 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5880
         TabIndex        =   115
         Top             =   5160
         Width           =   1185
      End
      Begin VB.Label xsaldos 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5880
         TabIndex        =   114
         Top             =   5880
         Width           =   1185
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Administracion"
      Height          =   8655
      Left            =   11640
      TabIndex        =   81
      Top             =   3360
      Visible         =   0   'False
      Width           =   10335
      Begin VB.CommandButton Command17 
         Caption         =   "Close"
         Height          =   495
         Left            =   9000
         TabIndex        =   82
         Top             =   240
         Width           =   1215
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   8175
         Left            =   120
         TabIndex        =   83
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   14420
         _Version        =   393217
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Dia de Trabajo"
      Height          =   4335
      Left            =   10440
      TabIndex        =   73
      Top             =   1680
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox xdiatrabajo 
         Height          =   495
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   78
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox xturno 
         Height          =   495
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   77
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Graba"
         Height          =   735
         Left            =   120
         TabIndex        =   76
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Borra"
         Height          =   735
         Left            =   1680
         TabIndex        =   75
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Close"
         Height          =   855
         Left            =   5520
         TabIndex        =   74
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         Height          =   495
         Left            =   120
         TabIndex        =   80
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
         Height          =   495
         Left            =   120
         TabIndex        =   79
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Planning"
      Height          =   495
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   0
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Entrada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   13560
      TabIndex        =   55
      Top             =   1560
      Visible         =   0   'False
      Width           =   12735
      Begin VB.CommandButton Command13 
         Caption         =   "Close"
         Height          =   615
         Left            =   5280
         TabIndex        =   62
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Grabar"
         Height          =   615
         Left            =   5280
         TabIndex        =   61
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox yprecio 
         Height          =   495
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   60
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox yhoras 
         Height          =   495
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   59
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox yhorae 
         Height          =   495
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   58
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox yfechas 
         Height          =   495
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   57
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox yfechae 
         Height          =   495
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   56
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   67
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaEntrada"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   66
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaSalida"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   65
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hora Entrada"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   64
         Top             =   1560
         Width           =   1395
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hora Salida"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   63
         Top             =   2040
         Width           =   1395
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Filtrar"
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hoy"
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   0
      Width           =   735
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
      Height          =   1215
      Index           =   47
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   6600
      Width           =   1215
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
      Height          =   1215
      Index           =   46
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   5400
      Width           =   1215
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
      Height          =   1215
      Index           =   45
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   4200
      Width           =   1215
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
      Height          =   1215
      Index           =   44
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   3000
      Width           =   1215
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
      Height          =   1215
      Index           =   43
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   1800
      Width           =   1215
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
      Height          =   1215
      Index           =   42
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   600
      Width           =   1215
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
      Height          =   1215
      Index           =   41
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6600
      Width           =   1215
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
      Height          =   1215
      Index           =   40
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   5400
      Width           =   1215
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
      Height          =   1215
      Index           =   39
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   4200
      Width           =   1215
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
      Height          =   1215
      Index           =   38
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3000
      Width           =   1215
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
      Height          =   1215
      Index           =   37
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1800
      Width           =   1215
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
      Height          =   1215
      Index           =   36
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   600
      Width           =   1215
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
      Height          =   1215
      Index           =   35
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   6600
      Width           =   1215
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
      Height          =   1215
      Index           =   34
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   5400
      Width           =   1215
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
      Height          =   1215
      Index           =   33
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4200
      Width           =   1215
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
      Height          =   1215
      Index           =   32
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3000
      Width           =   1215
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
      Height          =   1215
      Index           =   31
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1800
      Width           =   1215
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
      Height          =   1215
      Index           =   30
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   600
      Width           =   1215
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
      Height          =   1215
      Index           =   29
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6600
      Width           =   1215
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
      Height          =   1215
      Index           =   28
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   1920
      TabIndex        =   30
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Height          =   375
         Left            =   1800
         TabIndex        =   32
         Top             =   3120
         Width           =   615
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2670
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   4710
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   16777088
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StartOfWeek     =   230227969
         CurrentDate     =   41952
      End
   End
   Begin VB.TextBox fecha 
      Height          =   495
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   29
      Top             =   0
      Width           =   1575
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
      Height          =   1215
      Index           =   27
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4200
      Width           =   1215
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
      Height          =   1215
      Index           =   26
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3000
      Width           =   1215
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
      Height          =   1215
      Index           =   25
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1800
      Width           =   1215
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
      Height          =   1215
      Index           =   24
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   600
      Width           =   1215
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
      Height          =   1215
      Index           =   23
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6600
      Width           =   1215
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
      Height          =   1215
      Index           =   22
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5400
      Width           =   1215
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
      Height          =   1215
      Index           =   21
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4200
      Width           =   1215
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
      Height          =   1215
      Index           =   20
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3000
      Width           =   1215
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
      Height          =   1215
      Index           =   19
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1800
      Width           =   1215
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
      Height          =   1215
      Index           =   18
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   600
      Width           =   1215
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
      Height          =   1215
      Index           =   17
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6600
      Width           =   1215
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
      Height          =   1215
      Index           =   16
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5400
      Width           =   1215
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
      Height          =   1215
      Index           =   15
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4200
      Width           =   1215
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
      Height          =   1215
      Index           =   14
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
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
      Height          =   1215
      Index           =   13
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
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
      Height          =   1215
      Index           =   12
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   1215
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
      Height          =   1215
      Index           =   11
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6600
      Width           =   1215
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
      Height          =   1215
      Index           =   10
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   1215
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
      Height          =   1215
      Index           =   9
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
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
      Height          =   1215
      Index           =   8
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
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
      Height          =   1215
      Index           =   7
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
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
      Height          =   1215
      Index           =   6
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   1215
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
      Height          =   1215
      Index           =   5
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
      Width           =   1215
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
      Height          =   1215
      Index           =   4
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
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
      Height          =   1215
      Index           =   3
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
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
      Height          =   1215
      Index           =   2
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
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
      Height          =   1215
      Index           =   1
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
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
      Height          =   1215
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   14280
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "treevho.frx":0000
            Key             =   "picture1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "treevho.frx":059A
            Key             =   "picture2"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DiaTrabajo"
      Height          =   375
      Left            =   2040
      TabIndex        =   71
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label diatrabajo 
      BackColor       =   &H00FFFFC0&
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
      Height          =   375
      Left            =   3120
      TabIndex        =   70
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Label turno 
      BackColor       =   &H00FFFFC0&
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
      Height          =   375
      Left            =   1080
      TabIndex        =   69
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Turno:"
      Height          =   375
      Left            =   0
      TabIndex        =   68
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      Height          =   495
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   1815
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   7200
      Picture         =   "treevho.frx":0B34
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1200
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   8400
      Picture         =   "treevho.frx":2706
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1200
   End
   Begin VB.Menu dfkir44 
      Caption         =   "&Administracion"
   End
   Begin VB.Menu d89 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "treevho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim buffer(50)      As String

Dim jindx           As Integer

Dim mmesacod(15000) As String

Dim wmesacod(15000) As String

Dim wwmesacod(50)   As String

Dim mmesapag        As Integer

Dim mmesatop        As Integer

Dim msalcod(100)    As String

Dim msalpag         As Integer

Dim msaltop         As Integer

Option Explicit

Private Sub btnsalir_Click()
    d89_Click

End Sub

Private Sub Command1_Click()
    Frame1.Visible = False

End Sub

Private Sub Command10_Click()
    'If Len(Trim(turno)) = 0 Then
    '   MsgBox "Seleccione un Turno de Trabajo ", 48, "Aviso"
    '   Exit Sub
    'End If
    Frame2.Visible = False
    Command6_Click

End Sub

Private Sub Command11_Click()

    If Len(Trim(turno)) = 0 Then
        MsgBox "Seleccione un Turno de Trabajo ", 48, "Aviso"
        Exit Sub

    End If

    Frame3.Visible = True
    yfechae = Format(Now, "dd/mm/yyyy")
    yfechas = Format(Now + 1, "dd/mm/yyyy")
    yhorae = Format(Now, "HH:MM:SS")
    yhoras = "13:00:00"
    yprecio = Format(Val(precio), "0.00")
    Frame2.Visible = False
    Command6_Click

End Sub

Private Sub Command12_Click()

    Dim mytablex As New ADODB.Recordset

    If Not IsDate(yfechae) Then
        yfechae.SetFocus
        Exit Sub

    End If

    If Not IsDate(yfechas) Then
        yfechas.SetFocus
        Exit Sub

    End If

    If Len(yhorae) = 0 Then
        yhorae.SetFocus
        Exit Sub

    End If

    If Len(yhoras) = 0 Then
        yhoras.SetFocus
        Exit Sub

    End If

    If Val(yprecio) = 0 Then
        yprecio.SetFocus
        Exit Sub

    End If

    mytablex.Open "SELECT * FROM hotelcheckin where checkin=" & Val(xcheckin), cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Fields("estado") = "ENTRADA"
        mytablex.Fields("arribofecha") = Format(yfechae, "dd/mm/yyyy")
        mytablex.Fields("arribofechaf") = Format(yfechas, "dd/mm/yyyy")
        mytablex.Fields("arribohora") = yhorae
        mytablex.Fields("arribohoraf") = yhoras
        mytablex.Fields("turno") = Val(turno)
        mytablex.Update

    End If

    mytablex.Close
    Frame3.Visible = False
    Frame2.Visible = False
    Command6_Click

End Sub

Private Sub Command13_Click()
    Frame3.Visible = False

End Sub

Private Sub Command14_Click()

    Dim mytablex As New ADODB.Recordset

    If Len(Trim(turno)) = 0 Then
        MsgBox "Seleccione un Turno de Trabajo ", 48, "Aviso"
        Exit Sub

    End If

    If Val(xsaldos) > 0 Then
        MsgBox "Habitacion no cancelada,Todavia", 48, "Aviso"
        Exit Sub

    End If

    If Len(Trim(turno)) = 0 Then
        MsgBox "Seleccione Turno Trabajo", 48, "Aviso"
        Exit Sub

    End If

    If MsgBox("Desea dar Salida a la habitacion", 1, "Aviso") <> 1 Then Exit Sub
    mytablex.Open "SELECT * FROM hotelcheckin where checkin=" & Val(xcheckin), cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Fields("estado") = "CERRADO"
        mytablex.Fields("noches") = Val(xpermanencia)
        mytablex.Fields("arribofechaf") = Format(Now, "dd/mm/yyyy")
        mytablex.Fields("fechasalida") = Format(Now, "dd/mm/yyyy")
        mytablex.Fields("horasalida") = Format(Now, "HH:MM:SS")
        mytablex.Fields("arribohoraf") = Format(Now, "HH:MM:SS")
        'mytablex.Fields("turno") = Val(turno)
        mytablex.Fields("total") = Val(xpermanencia) * Val("" & mytablex.Fields("precio"))
        mytablex.Update

    End If

    mytablex.Close
    'CAMBIAR COLOR DE LA HABITACION A LIMPIEZA
    mytablex.Open "SELECT * FROM habitacion where habitacion='" & Trim(Command9.Caption) & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Fields("estado") = "S"
        mytablex.Fields("checkin") = ""
        mytablex.Update

    End If

    mytablex.Close

    Frame3.Visible = False
    Command6_Click
    busca_presentacion Command9.Caption
    visualiza_habitacion Command9.Caption
    sumar_abonos
    dias_ocupados
    sumar_consumos
    suma_saldos
    Frame2.Visible = False
    Command6_Click

End Sub

Private Sub Command15_Click()

    If Len(Trim(turno)) = 0 Then
        MsgBox "Seleccione un Turno de Trabajo ", 48, "Aviso"
        Exit Sub

    End If

    tnovedad.FLAG = Trim(Command9.Caption)
    tnovedad.Show 1
    Frame2.Visible = False
    Command6_Click

End Sub

Private Sub Command16_Click()

    If Len(Trim(turno)) = 0 Then
        MsgBox "Seleccione un Turno de Trabajo ", 48, "Aviso"
        Exit Sub

    End If

    tlimpia.FLAG = Trim(Command9.Caption)
    tlimpia.Show 1
    Frame2.Visible = False
    Command6_Click

End Sub

Private Sub Command18_Click()

    If Val(xcheckin) > 0 Then
        MsgBox "Habitacion Ocupado ", 48, "Aviso"
        Exit Sub

    End If

    If Len(Trim(turno)) = 0 Then
        MsgBox "Seleccione un Turno de Trabajo ", 48, "Aviso"
        Exit Sub

    End If

    tcheckin.vflag = "NUEVO"
    tcheckin.vestado = "ENTRADA"
    tcheckin.xhabitacion = Trim("" & Command9.Caption)
    tcheckin.Show 1
    Frame2.Visible = False
    Command6_Click

End Sub

Private Sub Command17_Click()
    Frame4.Visible = False

End Sub

Private Sub Command19_Click()

    Dim mytablex As New ADODB.Recordset

    If Not IsDate(xdiatrabajo) Then
        xdiatrabajo = ""
        xdiatrabajo.SetFocus
        Exit Sub

    End If

    If Len(xturno) = 0 Then
        xturno = ""
        xturno.SetFocus

    End If

    mytablex.Open "select * from hotelturno", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew

    End If

    mytablex.Fields("diatrabajo") = Format(xdiatrabajo, "dd/mm/yyyy")
    mytablex.Fields("hotelturno") = xturno
    mytablex.Update
    mytablex.Close
    carga_turno

End Sub

Private Sub Command2_Click()

    If Len(Trim(turno)) = 0 Then
        MsgBox "Seleccione un Turno de Trabajo ", 48, "Aviso"
        Exit Sub

    End If

    tcheckin.vflag = "NUEVO"
    tcheckin.vestado = "RESERVA"
    tcheckin.xhabitacion = Trim("" & Command9.Caption)
    tcheckin.Show 1
    Frame2.Visible = False
    Command6_Click

End Sub

Private Sub Command20_Click()
    cn.Execute ("delete from hotelturno")
    xdiatrabajo = ""
    xturno = ""
    diatrabajo = ""
    turno = ""

    carga_turno
    xdiatrabajo.SetFocus

End Sub

Private Sub Command21_Click()
    d89_Click

End Sub

Private Sub Command22_Click()
    thabitapla.Show 1

End Sub

Private Sub Command3_Click()

    Dim buf As String

    If Len(Trim(turno)) = 0 Then
        MsgBox "Seleccione un Turno de Trabajo ", 48, "Aviso"
        Exit Sub

    End If

    buf = Trim("" & xcheckin)

    If Len(buf) = 0 Then
        MsgBox "Debe existir Reserva o Entrada ", 48, "Aviso"
        Exit Sub

    End If

    'thotelan.XFLAG = "NUEVO"
    thotelan.idreserva = Trim(buf)
    thotelan.Show 1

End Sub

Private Sub Command4_Click()
    fecha = Format(Now, "dd/mm/yyyy")

End Sub

Private Sub Command5_Click()

    Dim buf As String

    If Len(Trim(turno)) = 0 Then
        MsgBox "Seleccione un Turno de Trabajo ", 48, "Aviso"
        Exit Sub

    End If

    buf = Trim("" & xcheckin)
    thotelco.idcheckin = Trim(buf)
    'thotelco.xflag = "NUEVO"
    thotelco.idhabitacion = buf
    thotelco.Show 1
    Frame2.Visible = False
    Command6_Click

End Sub

Private Sub Command6_Click()

    Dim I As Integer

    For I = 0 To 47
        groupmesa(I).BackColor = &H80FF80
    Next I

    menu_carga_mesa "TODOS"
    menu_mesa "INI"

End Sub

Private Sub Command7_Click()

    Dim buf As String

    If Len(Trim(turno)) = 0 Then
        MsgBox "Seleccione un Turno de Trabajo ", 48, "Aviso"
        Exit Sub

    End If

    buf = Trim("" & xcheckin)
    thotelpr.idcheckin = Trim(buf)
    'thotelco.xflag = "NUEVO"
    thotelpr.idhabitacion = buf
    thotelpr.Show 1
    Frame2.Visible = False
    Command6_Click

End Sub

Private Sub Command8_Click()

    If Len(Trim(turno)) = 0 Then
        MsgBox "Seleccione un Turno de Trabajo ", 48, "Aviso"
        Exit Sub

    End If

    thotelfa.idxcheckin = Trim(xcheckin)
    thotelfa.Show 1
    Frame2.Visible = False
    Command6_Click

End Sub

Private Sub d89_Click()

    If Frame5.Visible = True Then
        carga_turno
        Frame5.Visible = False
        Exit Sub

    End If

    If Frame4.Visible = True Then
        Frame4.Visible = False
        Exit Sub

    End If

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Exit Sub

    End If

    treevho.Hide
    Unload treevho

End Sub

Private Sub dfkir44_Click()
    Frame4.Visible = True

End Sub

Private Sub Form_Load()

    Dim sp       As String

    Dim spp      As String

    Dim sh       As String

    Dim shh      As String

    Dim sp1      As String

    Dim sh1      As String

    Dim sp2      As String

    Dim sh2      As String

    Dim sp3      As String

    Dim sh3      As String

    Dim sp4      As String

    Dim sh4      As String

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    carga_turno
    'MsgBox "abcd"
    fecha = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
    fecha = Format(Now, "dd/mm/yyyy")
    sp = "sp"
    spp = "spp"
    sp1 = "sp1"
    sp2 = "sp2"
    sp3 = "sp3"
    sp4 = "sp4"
    
    TreeView1.ImageList = ImageList1
    
    TreeView1.Nodes.Add , , sp, "Tablas", "picture1"
    
    TreeView1.Nodes.Add sp, tvwChild, sh, "Clientes", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Personal", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "TipoReserva", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "TipoTarifa", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "TipoPension", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "TipoHabitacion", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Caracteristicas", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Habitacion", "picture1"
    TreeView1.Nodes.Add sp, tvwChild, sh, "Fuera de Servicio", "picture1"
       
    TreeView1.Nodes.Add , , sp3, "Procesos", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Planning", "picture1"
    'TreeView1.Nodes.Add sp3, tvwChild, sh3, "Reservas", "picture1"
    ' TreeView1.Nodes.Add sp3, tvwChild, sh3, "Abonos Reservas", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Reservas", "picture1"
    ' TreeView1.Nodes.Add sp3, tvwChild, sh3, "Reservas en Grupo", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Abonos", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Hospedaje y Consumos", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "EstadoCuenta", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Facturacion", "picture1"
    'TreeView1.Nodes.Add sp3, tvwChild, sh3, "CheckOut", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Incidencias", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Limpieza", "picture1"
    ' TreeView1.Nodes.Add sp3, tvwChild, sh3, "Apertura Turno", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "Cuadre", "picture1"
    TreeView1.Nodes.Add sp3, tvwChild, sh3, "CambiaTurno", "picture1"
    'TreeView1.Nodes.Add sp3, tvwChild, sh3, "EstadoHabitacion", "picture1"
    
    '   TreeView1.Nodes.Add , , spp, "Almacen", "picture1"
    
    '   TreeView1.Nodes.Add spp, tvwChild, shh, "ParteEntrada", "picture1"
    '  TreeView1.Nodes.Add spp, tvwChild, shh, "ParteSalida", "picture1"
    
    TreeView1.Nodes.Add , , sp2, "Reportes", "picture1"
    TreeView1.Nodes.Add sp2, tvwChild, sh2, "Huespedes", "picture1"
   
    For I = 1 To 50
        buffer(I) = ""
    Next I

    TreeView1.Nodes.Add , , sp4, "ReportesUsuario", "picture1"
    
    '------------------
    jindx = 0

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select * from archivo where menu='PLANILLA' and   estado='S'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            jindx = jindx + 1
            buffer(jindx) = Trim("" & mytablex.Fields("descripcio"))
            TreeView1.Nodes.Add sp4, tvwChild, sh4, Trim("" & mytablex.Fields("descripcio")), "picture1"
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    
    For I = 1 To TreeView1.Nodes.count - 1
        'TreeView1.Nodes(i).ExpandedImage = "Open"
        TreeView1.Nodes(I).Expanded = True
    Next I

    For I = 0 To 47
        groupmesa(I).BackColor = &H80FF80
    Next I

    menu_carga_mesa "TODOS"
    menu_mesa "INI"
    'Frame4.Visible = True
     
    Exit Sub

    'cmdLlenarTree_Click

End Sub
 
Private Sub groupmesa_Click(Index As Integer)

    If Len(Trim("" & groupmesa(Index).Caption)) = 0 Then Exit Sub
    Frame2.Visible = True
    Command9.Caption = Trim("" & groupmesa(Index).Caption)
    Command9.BackColor = groupmesa(Index).BackColor
    busca_presentacion Command9.Caption
    visualiza_habitacion Command9.Caption
    sumar_abonos
    dias_ocupados
    sumar_consumos
    suma_saldos

End Sub

Private Sub Image2_Click()

    Dim I As Integer

    For I = 0 To 47
        groupmesa(I).BackColor = &H80FF80
        'mesa = ""
    Next I

    menu_mesa "SIG"

End Sub

Private Sub Image3_Click()

    Dim I As Integer

    For I = 0 To 47
        groupmesa(I).BackColor = &H80FF80
        'mesa = ""
    Next I

    menu_mesa "ANT" ', salon

End Sub

Private Sub Label1_Click()
    Frame1.Visible = True

End Sub

Private Sub Label11_Click()
    xdiatrabajo = Format(Now, "dd/mm/yyyy")

End Sub

Private Sub Label22_Click()

    Dim mytablex As New ADODB.Recordset

    Frame5.Visible = True
    xdiatrabajo = ""
    xturno = ""
    mytablex.Open "select * from hotelturno", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        xdiatrabajo = Format(Trim("" & mytablex.Fields("diatrabajo")), "dd/mm/yyyy")
        xturno = Trim("" & mytablex.Fields("hotelturno"))

    End If

    mytablex.Close

End Sub

Private Sub Label4_Click()

    If Len(Trim(turno)) = 0 Then
        MsgBox "Seleccione un Turno de Trabajo ", 48, "Aviso"
        Exit Sub

    End If

    'tcheckin.vflag = "NUEVO"
    tcheckin.xhabitacion = Trim("" & Command9.Caption)
    tcheckin.Show 1
    Frame2.Visible = False
    Command6_Click

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim I As Integer

    'If jindx > 0 Then
    'For i = 1 To jindx
    '    If Node = buffer(i) Then
    '       ejecuta_reporte buffer(i)
    '    End If
    'Next i
    'End If

    'If Node = "Fuera de Servicio" Then
    'thabitac.flag = "FS"
    'thabitac.Show 1
    'End If

    If Node = "CambiaTurno" Then
        'Frame4.Visible = True
        Exit Sub

    End If

    If Node = "Cuadre" Then
        thocpro.Show 1
        carga_turno
        Exit Sub

    End If

    If Node = "Apertura Turno" Then
        thocpro.Show 1
        carga_turno
        Exit Sub

    End If

    If Len(Trim(turno)) = 0 Then
        MsgBox "Seleccione un Turno de Trabajo ", 48, "Aviso"
        Exit Sub

    End If

    If Node = "Huespedes" Then
        trepohotel.Show 1

    End If

    If Node = "Limpieza" Then
        tlimpia.Show 1

    End If

    If Node = "ParteEntrada" Then
        thotemov.vtipo = "S"
        thotemov.Show 1

    End If

    If Node = "ParteSalida" Then
        thotemov.vtipo = "T"
        thotemov.Show 1

    End If

    If Node = "Incidencias" Then
        tnovedad.Show 1

    End If

    If Node = "TipoReserva" Then
        ttiporeserva.Show 1

    End If

    If Node = "TipoPension" Then
        ttipopension.Show 1

    End If

    If Node = "TipoTarifa" Then
        ttipotarifa.Show 1

    End If

    If Node = "Caracteristicas" Then
        tcarater.Show 1

    End If

    If Node = "Planning" Then
        thabitapla.Show 1

    End If

    If Node = "TipoHabitacion" Then
        tipohabi.Show 1

    End If

    If Node = "Clientes" Then
        tnclie.DBPROV = "clientes"
        tnclie.Show 1

    End If

    If Node = "EstadoHabitacion" Then

        'thotelet.Show 1
    End If

    If Node = "Personal" Then
        tpersona.Show 1

    End If

    If Node = "Reservas" Then
        'treserva.Show 1
        tcheckin.xhabitacion = ""
        tcheckin.Show 1

    End If

    If Node = "Reservas en Grupo" Then

        'tcheckgr.xhabitacion = ""
        'tcheckgr.Show 1
    End If

    If Node = "Abonos Reservas" Then

        'treserva.xsw = "ANTICIPO"
        'treserva.Show 1
    End If

    If Node = "CheckIn" Then
        tcheckin.xhabitacion = ""
        tcheckin.Show 1

    End If

    If Node = "Habitacion" Then
        thabitac.Show 1

    End If

    If Node = "Hospedaje y Consumos" Then
        tcheckin.xsw = "CONSUMO"
        tcheckin.Show 1

    End If

    If Node = "Apertura Turno" Then
        thocpro.Show 1

    End If

    If Node = "Abonos" Then
        tcheckin.xsw = "ANTICIPO"
        tcheckin.Show 1

    End If

    If Node = "CheckOut" Then
        tcheckin.xsw = "SALIDA"
        tcheckin.Show 1

    End If

    If Node = "EstadoCuenta" Then
        tcheckin.xsw = "PRECUENTA"
        tcheckin.Show 1

    End If

    If Node = "Facturacion" Then
        thotelfa.Show 1

    End If

End Sub

Sub menu_carga_mesa(buf As String)

    Dim mytablex As New ADODB.Recordset

    Dim I        As Integer

    For I = 0 To 47
        wwmesacod(I) = ""
    Next I

    For I = 0 To 14999
        mmesacod(I) = ""
        wmesacod(I) = ""
    Next I

    I = -1

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM habitacion order by habitacion", cn, adOpenDynamic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        I = I + 1
        mmesacod(I) = "" & mytablex.Fields("habitacion")
        wmesacod(I) = "" & mytablex.Fields("Descripcio")
  
        mytablex.MoveNext
    Loop

    mytablex.Close
    mmesatop = I
    mmesapag = 0

End Sub

Sub menu_mesa(buf As String)

    Dim I As Integer

    Dim j As Integer

    Select Case buf

        Case "INI"
            mmesapag = 0

        Case "SIG"
            mmesapag = mmesapag + 47

            If mmesapag > 102 Then
                mmesapag = 0

            End If

        Case "ANT"
            mmesapag = mmesapag - 47

            If mmesapag < 0 Then
                mmesapag = 0

            End If

    End Select

    j = -1

    For I = mmesapag To 47 + mmesapag
        j = j + 1
        groupmesa(j).Caption = wmesacod(I) 'mmesacod(i)
        verifica_habitacion j, groupmesa(j).Caption
        'verifica_mesas j, groupmesa(j).Caption
    Next I

End Sub

Sub verifica_habitacion(indx As Integer, buf1 As String)

    Dim mytablex As New ADODB.Recordset

    If Len(Trim(buf1)) = 0 Then Exit Sub
    groupmesa(indx).BackColor = &H80FF80
    verifica_colores buf1, indx
    Exit Sub

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

Sub verifica_colores(xbuf As String, indx As Integer)

    Dim sw       As Integer

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    sw = 0

    If Not IsDate(fecha) Then Exit Sub
    groupmesa(indx).BackColor = &H80FF80  'verde
    'si esta en aseo
    mytablex.Open "SELECT * FROM habitacion where habitacion='" & Trim(xbuf) & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If "" & mytablex.Fields("estado") = "S" Then 's sucio
            groupmesa(indx).BackColor = &HFF&

        End If

    End If

    If Val("" & mytablex.Fields("checkin")) > 0 Then
        groupmesa(indx).BackColor = &H80FFFF

    End If

    mytablex.Close
    Exit Sub
    buf = "SELECT dbo.hotelcheckin.ESTADO,dbo.hotelcheckin.checkin, dbo.hotelcheckin.habitacion"
    buf = buf & " FROM         dbo.hotelcheckin where"
    buf = buf & " dbo.hotelcheckin.arribofecha<='" & Format(fecha, "YYYYMMDD") & "'"
    'buf = buf & " and   dbo.hotelcheckin.arribofechaf>='" & Format(fecha, "YYYYMMDD") & "'"
    buf = buf & " and dbo.hotelcheckin.habitacion='" & Trim(xbuf) & "'"
    buf = buf & " and (dbo.hotelcheckin.estado='ENTRADA' OR dbo.hotelcheckin.estado='RESERVA')"
    mytabley.Open buf, cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        sw = 1

        'mytablex.Fields("l" & x) = "O"
        If "" & mytabley.Fields("estado") = "RESERVA" Then
            groupmesa(indx).BackColor = &HFFFF80

        End If

        If "" & mytabley.Fields("estado") = "ENTRADA" Then
            groupmesa(indx).BackColor = &H80FFFF

        End If

    End If

    mytabley.Close

End Sub

Sub busca_presentacion(buf As String)

    Dim mytablex As New ADODB.Recordset

    xcheckin = ""
    inicializa
    mytablex.Open "SELECT * FROM habitacion where habitacion='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        precio = Trim("" & mytablex.Fields("precio"))
        xcheckin = Trim("" & mytablex.Fields("checkin"))

    End If

    mytablex.Close

End Sub

Sub visualiza_habitacion(xbuf As String)

    Dim mytabley As New ADODB.Recordset

    Dim buf      As String

    If Not IsDate(fecha) Then Exit Sub
    buf = "SELECT    * "
    buf = buf & " FROM dbo.hotelcheckin where"
    buf = buf & "  dbo.hotelcheckin.checkin=" & Val(xcheckin) & ""
    mytabley.Open buf, cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        pone_registro mytabley

    End If

    mytabley.Close

End Sub

Sub sumar_abonos()

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from hotelfactura where idcheckin=" & Val(xcheckin), cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + Val("" & mytablex.Fields("total"))
        mytablex.MoveNext
    Loop
    xabonos = Format(sdx, "0.00")
    mytablex.Close

End Sub

Sub sumar_consumos()

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from hotelconsumo where idecheckin=" & Val(xcheckin), cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + Val("" & mytablex.Fields("total"))
        mytablex.MoveNext
    Loop
    xconsumos = Format(sdx, "0.00")
    mytablex.Close

End Sub

Sub dias_ocupados()

    Dim sdx  As Double

    Dim xhoy As String

    Dim dias As Integer

    sdx = Val(noches) * Val(precio)
    xtotal = Format(sdx, "0.00")

    If categoria = "HORAS" Then
        xtotal = Format(precio, "0.00")

    End If
   
    Exit Sub

    If Not IsDate(arribofecha) Then
        xtotal = "0.00"
        Exit Sub

    End If

    xhoy = Format(fecha, "dd/mm/yyyy")
    dias = DateDiff("d", arribofecha, xhoy)

    If dias = 0 Then
        dias = 1

    End If

    If categoria = "HORAS" Then
        dias = 1

    End If
   
    sdx = Val(precio) * dias
    xtotal = Format(sdx, "0.00")
    xpermanencia = "" & dias

End Sub

Sub suma_saldos()

    Dim sdx As Double

    sdx = Val(xtotal) + Val(xconsumos) - Val(xabonos)
    xsaldos = Format(sdx, "0.00")

End Sub

Sub carga_turno()

    Dim mytablex As New ADODB.Recordset

    diatrabajo = ""
    turno = ""
    mytablex.Open "select * from hotelturno", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        diatrabajo = Format(Trim("" & mytablex.Fields("diatrabajo")), "dd/mm/yyyy")
        turno = Trim("" & mytablex.Fields("hotelturno"))

    End If

    mytablex.Close

    'turno = "1"
    'diatrabajo = Format(Now, "dd/mm/yyyy")
End Sub

Sub pone_registro(mytabley As ADODB.Recordset)
    tipocodigoh = Trim("" & mytabley.Fields("tipocodigoh"))
    tipocodigo = Trim("" & mytabley.Fields("tipocodigo"))
    tipotarifa = Trim("" & mytabley.Fields("tipotarifa"))
    'tiporeserva = Trim("" & mytabley.Fields("tiporeserva"))
    tipopension = Trim("" & mytabley.Fields("tipopension"))

    personas = Trim("" & mytabley.Fields("personas"))
    'tiporeserva = Trim("" & mytabley.Fields("tiporeserva"))
    'horareserva = Trim("" & mytabley.Fields("horareserva"))
    'fechareserva = Trim("" & mytabley.Fields("fechareserva"))
    categoria = Trim("" & mytabley.Fields("categoria"))
    Label23 = categoria
    'checkin = Trim("" & mytabley.Fields("checkin"))
    arribofecha = Trim("" & mytabley.Fields("arribofecha"))
    arribofechaf = Trim("" & mytabley.Fields("arribofechaf"))
    arribohora = Trim("" & mytabley.Fields("arribohora"))
    arribohoraf = Trim("" & mytabley.Fields("arribohoraf"))
    noches = Trim("" & mytabley.Fields("noches"))

    codigo = Trim("" & mytabley.Fields("codigo"))
    nombre = Trim("" & mytabley.Fields("nombre"))
    'direccion = Trim("" & mytabley.Fields("direccion"))

    huesped = Trim("" & mytabley.Fields("huesped"))
    hnombre = Trim("" & mytabley.Fields("hnombre"))
    'hdireccion = Trim("" & mytabley.Fields("hdireccion"))

    'procedencia = Trim("" & mytabley.Fields("procedencia"))
    'operador = Trim("" & mytabley.Fields("operador"))
    'agente = Trim("" & mytabley.Fields("agente"))
    precio = Trim("" & mytabley.Fields("precio"))

    'adulto = Trim("" & mytabley.Fields("adulto"))
    'nino = Trim("" & mytabley.Fields("nino"))
    'habitacion = Trim("" & mytabley.Fields("habitacion"))
    estado = Trim("" & mytabley.Fields("estado"))
    'hotelcuadre = Trim("" & mytabley.Fields("hotelcuadre"))

End Sub

Sub inicializa()
    tipocodigo = ""
    tipocodigoh = ""

    'horas = ""
    precio = ""
    estado = ""
    Label23 = ""
    'reservadas.Clear
    tipotarifa = ""

    tipopension = ""
    personas = ""

    categoria = ""
    'idreserva = ""
    arribofecha = ""
    arribofechaf = ""
    arribohora = ""
    arribohoraf = ""  'Format(Now, "hh:mm")

    nombre = ""
    codigo = ""
    'direccion = ""
    huesped = ""
    hnombre = ""
    'hdireccion = ""
    'tipoviaje = ""
    'procedencia = ""

    'adulto = ""
    'nino = ""
    noches = ""
    'carga_precio Trim("" & habitacion)
    'tipoviaje = ""

End Sub

