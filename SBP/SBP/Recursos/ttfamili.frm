VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ttfamilia 
   BackColor       =   &H00808080&
   Caption         =   "Tabla de Familias"
   ClientHeight    =   9150
   ClientLeft      =   165
   ClientTop       =   -45
   ClientWidth     =   12630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   12630
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Frame2"
      Height          =   7815
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   12495
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
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
         Height          =   2415
         Left            =   120
         TabIndex        =   60
         Top             =   5280
         Width           =   7305
         Begin VB.CommandButton cmdCommand1 
            Caption         =   "Aceptar"
            Height          =   345
            Left            =   2280
            TabIndex        =   65
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox estaxx 
            Height          =   375
            Left            =   3120
            MaxLength       =   1
            TabIndex        =   64
            Top             =   1800
            Width           =   555
         End
         Begin VB.HScrollBar hs3 
            Height          =   375
            Left            =   120
            TabIndex        =   63
            Top             =   1320
            Width           =   3615
         End
         Begin VB.HScrollBar hs2 
            Height          =   375
            Left            =   120
            TabIndex        =   62
            Top             =   840
            Width           =   3615
         End
         Begin VB.HScrollBar HS1 
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label lblFamilia 
            BackColor       =   &H00808080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Colores Defecto>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   66
            Top             =   1920
            Width           =   1635
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000004&
            BorderStyle     =   2  'Dash
            BorderWidth     =   5
            X1              =   0
            X2              =   3720
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label elcolor 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   2055
            Left            =   3840
            TabIndex        =   59
            Top             =   240
            Width           =   3375
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   2400
         Top             =   7560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox orden 
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
         Left            =   4920
         MaxLength       =   3
         TabIndex        =   56
         Top             =   3960
         Width           =   615
      End
      Begin VB.ComboBox cboprinters 
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   4800
         Width           =   7215
      End
      Begin VB.TextBox puerto 
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
         TabIndex        =   51
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox cola 
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
         TabIndex        =   49
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox red 
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
         TabIndex        =   47
         Top             =   2040
         Width           =   3255
      End
      Begin VB.TextBox familia 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   29
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox descripcio 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   28
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox tipo 
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
         TabIndex        =   27
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox maxdscto 
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
         TabIndex        =   26
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox margen1 
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
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   25
         Top             =   7560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   9600
         Picture         =   "ttfamili.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Imprimir todo"
         Top             =   1560
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   9600
         Picture         =   "ttfamili.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   480
         Width           =   1470
      End
      Begin VB.TextBox margen2 
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
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   22
         Top             =   7920
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox margen3 
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
         Left            =   5160
         MaxLength       =   10
         TabIndex        =   21
         Top             =   8280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox margen4 
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
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   20
         Top             =   8040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox margen5 
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
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   19
         Top             =   8400
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox margen6 
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
         Left            =   9360
         MaxLength       =   10
         TabIndex        =   18
         Top             =   7560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox margen7 
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
         Left            =   9360
         MaxLength       =   10
         TabIndex        =   17
         Top             =   7920
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox margen8 
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
         Left            =   9360
         MaxLength       =   10
         TabIndex        =   16
         Top             =   8280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox margen9 
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
         Left            =   9480
         MaxLength       =   10
         TabIndex        =   15
         Top             =   8040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox margen10 
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
         Left            =   9480
         MaxLength       =   10
         TabIndex        =   14
         Top             =   8400
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox obliga 
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
         TabIndex        =   13
         Top             =   3600
         Width           =   375
      End
      Begin VB.TextBox vetouch 
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
         Left            =   2265
         MaxLength       =   1
         TabIndex        =   12
         Text            =   "S"
         Top             =   3975
         Width           =   375
      End
      Begin VB.Label fotonombre 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9600
         TabIndex        =   58
         Top             =   4800
         Width           =   2295
      End
      Begin VB.Image foto 
         BorderStyle     =   1  'Fixed Single
         Height          =   2055
         Left            =   9600
         Stretch         =   -1  'True
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label23 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Orden Touch"
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
         Left            =   2760
         TabIndex        =   57
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Copiar"
         Height          =   375
         Left            =   8760
         TabIndex        =   55
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Impresoras"
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
         TabIndex        =   53
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grupo Impresion"
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
         TabIndex        =   52
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cola"
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
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Puerto Orden Despa."
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
         TabIndex        =   48
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Familia"
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
         TabIndex        =   46
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcio"
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
         TabIndex        =   45
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label4 
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
         TabIndex        =   44
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Maxdscto"
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
         TabIndex        =   43
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Margen1"
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
         Left            =   3000
         TabIndex        =   42
         Top             =   7560
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Margen2"
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
         Left            =   3000
         TabIndex        =   41
         Top             =   7920
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Margen3"
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
         Left            =   3000
         TabIndex        =   40
         Top             =   8280
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Margen4"
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
         Left            =   3120
         TabIndex        =   39
         Top             =   8040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Margen5"
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
         Left            =   3120
         TabIndex        =   38
         Top             =   8400
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Margen6"
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
         TabIndex        =   37
         Top             =   7560
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Margen7"
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
         TabIndex        =   36
         Top             =   7920
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Margen8"
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
         TabIndex        =   35
         Top             =   8280
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Margen9"
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
         Left            =   7320
         TabIndex        =   34
         Top             =   8040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Margen10"
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
         Left            =   7320
         TabIndex        =   33
         Top             =   8400
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1.Sin Inventario"
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
         Left            =   2895
         TabIndex        =   32
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Margen Obligado"
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
         TabIndex        =   31
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Visualiza Touch"
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
         TabIndex        =   30
         Top             =   3960
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
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
         Picture         =   "ttfamili.frx":1194
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
         Picture         =   "ttfamili.frx":23A6
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
         Picture         =   "ttfamili.frx":35B8
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
         Picture         =   "ttfamili.frx":47CA
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
         Picture         =   "ttfamili.frx":59DC
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
      Top             =   720
      Width           =   12495
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7905
         Left            =   75
         TabIndex        =   1
         Top             =   270
         Width           =   12060
         _ExtentX        =   21273
         _ExtentY        =   13944
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
         ColumnCount     =   7
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
            DataField       =   "Familia"
            Caption         =   "Familia"
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
            DataField       =   "Obliga"
            Caption         =   "Obliga"
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
            DataField       =   "vetouch"
            Caption         =   "Vetouch"
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
            DataField       =   "Red"
            Caption         =   "Red"
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
         BeginProperty Column05 
            DataField       =   "Puerto"
            Caption         =   "Grupo"
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
         BeginProperty Column06 
            DataField       =   "Orden"
            Caption         =   "Orden"
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
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column06 
            EndProperty
         EndProperty
      End
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
   Begin VB.Menu subt4 
      Caption         =   "&Subfamilia"
   End
   Begin VB.Menu fdk8923 
      Caption         =   "&Copiar"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "ttfamilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txempre As New ADODB.Recordset

Private Sub ajdu1_Click()

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Nuevo"

    cmdGuardar.Enabled = True
    habilita 1
    familia.Enabled = True
    familia = ""
    carga_impresoras
    familia.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    buf = txempre.Fields("familia")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + txempre.Fields("familia"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    cn.Execute ("delete from subfamil where familia='" & txempre.Fields("familia") & "'")
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

'Color por familia y producto  30/05/2018
Private Sub cmdCommand1_Click()

    If MsgBox("Seguro de Poner Color por defecto?", 1, "Aviso") <> 1 Then Exit Sub

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM paramecacolor where  caja='01'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        HS1.Value = Val("" & mytablex.Fields("colorfamilia1"))
        hs2.Value = Val("" & mytablex.Fields("colorfamilia2"))
        hs3.Value = Val("" & mytablex.Fields("colorfamilia3"))

    End If

    mytablex.Close

End Sub

'Color por familia y producto  30/05/2018

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

Private Sub familia_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(familia) = 0 Then Exit Sub
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
            cad = "SELECT * from familia    "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT *  from familia   where  " & Combo1 & " like '" & buffer & "%'"

        End If

        cad = cad & " order by orden "

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

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'familia = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'familia.SetFocus
        'familia_KeyPress 13
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

    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "familia"
    reporgen.Show 1

End Sub

Private Sub dlo132_Click()

    If Frame2.Visible = True Then
        habilita 0
        Frame2.Visible = False
        dbGrid1.Enabled = True
        Exit Sub

    End If

    ttfamilia.Hide
    Unload ttfamilia

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    buf = txempre.Fields("familia")

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
    carga_impresoras
    familia.Enabled = False
    descripcio.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fdk8923_Click()

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    If MsgBox("Desea Actualizar familias en producto para impresion", 1, "Aviso") <> 1 Then Exit Sub
   
    mytablex.Open "select * from familia ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
   
        buf = "update producto set puertoimpresion='" & "" & mytablex.Fields("red") & "'"
        buf = buf & ",grupoimpresion='" & "" & mytablex.Fields("puerto") & "'"
        buf = buf & ",cola='" & "" & mytablex.Fields("cola") & "' where familia='" & "" & mytablex.Fields("familia") & "'"
        cn.Execute (buf)
        mytablex.MoveNext
    Loop
    mytablex.Close
    MsgBox "Proceso Realizado ", 48, "Aviso"

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    buf = txempre.Fields("familia")

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
    familia.Enabled = False
    descripcio.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    Command1_Click

End Sub

Sub carga_impresoras()

    Dim I As Integer

    On Error GoTo cmd8912_err

    For I = 0 To Printers.count - 1
        cboprinters.AddItem Printers(I).DeviceName

        ' if this is the current printer, select it
        If Printers(I).DeviceName = Printer.DeviceName Then
            ' this indirectly executes ShowPrinterInfo
            cboprinters.ListIndex = I

        End If

    Next
    Exit Sub
cmd8912_err:
    Exit Sub

End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "descripcio"
    Combo1.AddItem "familia"
    Combo1.ListIndex = 0

    'Color por familia y producto  30/05/2018
    HS1.Min = 0
    HS1.max = 255
    HS1.LargeChange = 25
    HS1.SmallChange = 5

    hs2.Min = 0
    hs2.max = 255
    hs2.LargeChange = 25
    hs2.SmallChange = 5

    hs3.Min = 0
    hs3.max = 255
    hs3.LargeChange = 25
    hs3.SmallChange = 5
    'Color por familia y producto  30/05/2018

End Sub

Sub inicializa()
    fotonombre = ""
    orden = ""
    red = ""
    Puerto = ""
    cola = ""
    descripcio = ""
    tipo = ""
    maxdscto = ""
    margen1 = ""
    margen2 = ""
    margen3 = ""
    margen4 = ""
    margen5 = ""
    margen6 = ""
    margen7 = ""
    margen8 = ""
    margen9 = ""
    margen10 = ""
    obliga = ""
    vetouch = "S"

    'Color por familia y producto  30/05/2018
    Dim I        As Integer

    Dim I1       As Integer

    Dim I2       As Integer

    Dim I3       As Integer

    Dim I4       As Integer

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM paramecacolor where  caja='01'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        HS1.Value = Val("" & mytablex.Fields("colorfamilia1"))
        hs2.Value = Val("" & mytablex.Fields("colorfamilia2"))
        hs3.Value = Val("" & mytablex.Fields("colorfamilia3"))

    End If

    mytablex.Close

    'Color por familia y producto  30/05/2018

End Sub

Sub pone_registro()
    pone_fotonombre txempre
    fotonombre = Trim("" & txempre.Fields("familia"))
    orden = Trim("" & txempre.Fields("orden"))
    red = Trim("" & txempre.Fields("red"))
    cola = Trim("" & txempre.Fields("Cola"))
    Puerto = Trim("" & txempre.Fields("puerto"))
    vetouch = Trim("" & txempre.Fields("vetouch"))
    obliga = Trim("" & txempre.Fields("obliga"))
    familia = Trim("" & txempre.Fields("familia"))
    descripcio = Trim("" & txempre.Fields("descripcio"))
    tipo = Trim("" & txempre.Fields("tipo"))
    maxdscto = Trim("" & txempre.Fields("maxdscto"))
    margen1 = Trim("" & txempre.Fields("margen1"))
    margen2 = Trim("" & txempre.Fields("margen2"))
    margen3 = Trim("" & txempre.Fields("margen3"))
    margen4 = Trim("" & txempre.Fields("margen4"))
    margen5 = Trim("" & txempre.Fields("margen5"))
    margen6 = Trim("" & txempre.Fields("margen6"))
    margen7 = Trim("" & txempre.Fields("margen7"))
    margen8 = Trim("" & txempre.Fields("margen8"))
    margen9 = Trim("" & txempre.Fields("margen9"))
    margen10 = Trim("" & txempre.Fields("margen10"))

    'Color por familia y producto  30/05/2018
    If IsNull(txempre.Fields("c")) Then

        Dim mytablex As New ADODB.Recordset

        mytablex.Open "SELECT * FROM paramecacolor where  caja='01'", cn, adOpenKeyset, adLockOptimistic
    
        If mytablex.RecordCount > 0 Then
            HS1.Value = Val("" & mytablex.Fields("colorfamilia1"))
            hs2.Value = Val("" & mytablex.Fields("colorfamilia2"))
            hs3.Value = Val("" & mytablex.Fields("colorfamilia3"))

        End If

        mytablex.Close

    Else

        HS1.Value = "" & txempre.Fields("c")
        hs2.Value = "" & txempre.Fields("d")
        hs3.Value = "" & txempre.Fields("e")

    End If
    
    'Color por familia y producto  30/05/2018

End Sub

Sub grabando()
    txempre.Fields("fotonombre") = Trim("" & familia)
    SaveBitmap txempre, Trim("" & familia)
    txempre.Fields("orden") = Val("" & orden)
    txempre.Fields("puerto") = Trim(Puerto)
    txempre.Fields("red") = Trim(red)
    txempre.Fields("cola") = Trim(cola)
    txempre.Fields("descripcio") = Trim(descripcio)
    txempre.Fields("tipo") = Trim(tipo)
    txempre.Fields("vetouch") = Trim(vetouch)
    txempre.Fields("obliga") = Trim(obliga)
    txempre.Fields("maxdscto") = Val(maxdscto)
    txempre.Fields("margen1") = Val(margen1)
    txempre.Fields("margen2") = Val(margen2)
    txempre.Fields("margen3") = Val(margen3)
    txempre.Fields("margen4") = Val(margen4)
    txempre.Fields("margen5") = Val(margen5)
    txempre.Fields("margen6") = Val(margen6)
    txempre.Fields("margen7") = Val(margen7)
    txempre.Fields("margen8") = Val(margen8)
    txempre.Fields("margen9") = Val(margen9)
    txempre.Fields("margen10") = Val(margen10)

    'Color por familia y producto  30/05/2018
    txempre.Fields("c") = "" & HS1.Value
    txempre.Fields("d") = "" & hs2.Value
    txempre.Fields("e") = "" & hs3.Value
    'Color por familia y producto  30/05/2018

End Sub

Private Sub grba1_Click()

End Sub

Private Sub descripcio_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    tipo.SetFocus

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
        If Len(familia) = 0 Then
            familia.SetFocus
            Exit Function

        End If

        rbusca.Open "select familia from familia where familia='" & familia & "'", cn, adOpenStatic, adLockOptimistic

        If rbusca.RecordCount > 0 Then
            rbusca.Close
            MsgBox "Ya existe familia ", 48, "Aviso"
            Exit Function

        End If

        txempre.AddNew
        txempre.Fields("familia") = familia
        grabando
        txempre.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        txempre.Fields("familia") = familia
        grabando
        txempre.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    'If Len(familia) = 0 Then
    '   familia.SetFocus
    '   Exit Function
    'End If
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

Private Sub foto_Click()
    CommonDialog1.DialogTitle = "Seleccione un archivo Grafico"
    CommonDialog1.InitDir = globaldir & "\grafico"
    CommonDialog1.Filter = "Archivos Grafico|*.jpg"
    CommonDialog1.ShowOpen

    'Si seleccionamos un archivo mostramos la ruta
    If CommonDialog1.FileName <> "" Then
        fotonombre = CommonDialog1.FileName
        foto = LoadPicture(fotonombre)
    Else

        'Si no mostramos un texto de advertencia de que no se seleccionó _   ninguno, ya que FileName devuelve una cadena vacía
        'Label1 = "No se seleccionó ningún archivo"
    End If

End Sub

'Color por familia y producto  30/05/2018
Private Sub HS1_Change()
    elcolor.BackColor = RGB(HS1.Value, hs2.Value, hs3.Value)

End Sub

Private Sub hs2_Change()
    elcolor.BackColor = RGB(HS1.Value, hs2.Value, hs3.Value)

End Sub

Private Sub hs3_Change()
    elcolor.BackColor = RGB(HS1.Value, hs2.Value, hs3.Value)

End Sub

'Color por familia y producto  30/05/2018

Private Sub Label22_Click()
    red = cboprinters.Text

End Sub

Private Sub maxdscto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub maxdscto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        tipo.SetFocus
        Exit Sub

    End If

End Sub

Private Sub subt4_Click()

    Dim buf As String

    On Error GoTo cmd4556_err

    buf = txempre.Fields("familia")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    tsubfami.familia = "" & txempre.Fields("familia")
    tsubfami.nfamilia = "" & txempre.Fields("descripcio")
    tsubfami.Show 1
    Exit Sub
cmd4556_err:
    MsgBox "Seleccione un dato", 48, "Aviso"
    Exit Sub

End Sub

Private Sub tipo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub tipo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        descripcio.SetFocus
        Exit Sub

    End If

End Sub

Sub pone_fotonombre(mytablex As ADODB.Recordset)

    On Error GoTo cm897888_err

    foto = LoadPicture()
    fotonombre = Trim("" & mytablex.Fields("familia"))
    viewBMP mytablex, fotonombre

    If Len(fotonombre) > 0 Then
        If existe_archivo(fotonombre) > 0 Then
            foto = LoadPicture(fotonombre)

        End If

    End If

cm897888_err:
    Exit Sub

End Sub

