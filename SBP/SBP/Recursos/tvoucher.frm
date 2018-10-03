VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form tvoucher 
   BackColor       =   &H00FFFF00&
   Caption         =   "Vouchers"
   ClientHeight    =   6825
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   11475
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
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
      Height          =   6015
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Visible         =   0   'False
      Width           =   11295
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
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
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
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
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
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "tvoucher.frx":0000
         Height          =   5055
         Left            =   120
         OleObjectBlob   =   "tvoucher.frx":0014
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   840
         Width           =   11055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000009&
      Caption         =   "Datos del Voucher"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   21
      Top             =   1800
      Visible         =   0   'False
      Width           =   11175
      Begin VB.Frame Frame3 
         BackColor       =   &H80000009&
         Caption         =   "Registrar Documentos"
         Height          =   3735
         Left            =   5880
         TabIndex        =   37
         Top             =   360
         Width           =   5055
         Begin VB.TextBox nombre1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   1800
            MaxLength       =   60
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1680
            Width           =   3135
         End
         Begin VB.CommandButton Command3 
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
            Left            =   3480
            MaskColor       =   &H00E0E0E0&
            Picture         =   "tvoucher.frx":09DF
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Borrar registro"
            Top             =   2880
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
            Left            =   4200
            MaskColor       =   &H00E0E0E0&
            Picture         =   "tvoucher.frx":1BF1
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Grabar registro"
            Top             =   2880
            Width           =   735
         End
         Begin VB.ComboBox tipoclie 
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
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   960
            Width           =   1575
         End
         Begin VB.TextBox otros1 
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
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   5880
            Width           =   1575
         End
         Begin VB.TextBox exonerado1 
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
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   5520
            Width           =   1575
         End
         Begin VB.TextBox inafecto1 
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
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   5160
            Width           =   1575
         End
         Begin VB.TextBox igv1 
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
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   4800
            Width           =   1575
         End
         Begin VB.TextBox importe1 
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
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   4440
            Width           =   1575
         End
         Begin VB.ComboBox libro1 
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
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   4080
            Width           =   1575
         End
         Begin VB.TextBox tipo1 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox numero1 
            BackColor       =   &H00FFFFFF&
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
            Left            =   2400
            MaxLength       =   11
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox fecha1 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox codigo1 
            BackColor       =   &H00FFFFFF&
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
            Left            =   1800
            MaxLength       =   11
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label23 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tipoclie"
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
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label20 
            BackColor       =   &H00FFFF00&
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
            TabIndex        =   57
            Top             =   5880
            Width           =   1695
         End
         Begin VB.Label Label19 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Exonerado"
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
            Top             =   5520
            Width           =   1695
         End
         Begin VB.Label Label18 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Inafecto"
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
            TabIndex        =   55
            Top             =   5160
            Width           =   1695
         End
         Begin VB.Label Label17 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Igv"
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
            Top             =   4800
            Width           =   1695
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Importe"
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
            Top             =   4440
            Width           =   1695
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Libro"
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
            Top             =   4080
            Width           =   1695
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Tipo/Numero"
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
            TabIndex        =   51
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label14 
            BackColor       =   &H00FFFF00&
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
            Left            =   120
            TabIndex        =   50
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label15 
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
            Left            =   120
            TabIndex        =   49
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFFF00&
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
            TabIndex        =   48
            Top             =   1680
            Width           =   1695
         End
      End
      Begin VB.TextBox glosa 
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
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2760
         Width           =   3855
      End
      Begin VB.TextBox paridad 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1575
      End
      Begin VB.ComboBox moneda 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox haber 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox debe 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox cuenta 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ingresar Datos Documentos"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Glosa"
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
         TabIndex        =   36
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF00&
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
         TabIndex        =   35
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo/Cambio"
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
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
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
         TabIndex        =   33
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Haber"
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
         TabIndex        =   32
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Debe"
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
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label ncuenta 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   30
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta"
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
         TabIndex        =   29
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox thaberd 
      BackColor       =   &H00FFFF00&
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9960
      MaxLength       =   6
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox tdebed 
      BackColor       =   &H00FFFF00&
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9960
      MaxLength       =   6
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox thabers 
      BackColor       =   &H00FFFF00&
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8640
      MaxLength       =   6
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox tdebes 
      BackColor       =   &H00FFFF00&
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
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8640
      MaxLength       =   6
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox mes 
      BackColor       =   &H00FFFF00&
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
      Left            =   10080
      MaxLength       =   6
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox fecha 
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
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox asiento 
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
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox codigo 
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
      Left            =   2520
      MaxLength       =   11
      TabIndex        =   1
      Top             =   840
      Width           =   1695
   End
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
      Left            =   4320
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tvoucher.frx":2E03
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Grabar registro"
      Top             =   0
      Visible         =   0   'False
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
      Picture         =   "tvoucher.frx":4015
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Nuevo registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdHelp 
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
      Left            =   2880
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tvoucher.frx":5227
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ayuda"
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
      Left            =   1440
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tvoucher.frx":6439
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprimir"
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
      Left            =   3600
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tvoucher.frx":764B
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdDelete 
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
      Picture         =   "tvoucher.frx":885D
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Borrar registro"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0C0C0&
      Height          =   615
      Left            =   2160
      MaskColor       =   &H00E0E0E0&
      Picture         =   "tvoucher.frx":9A6F
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Consulta"
      Top             =   0
      Width           =   735
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "tvoucher.frx":AC81
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "tvoucher.frx":AC95
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1680
      Width           =   11175
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Dolares Debe Haber"
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
      Left            =   5880
      TabIndex        =   18
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Soles Debe Haber"
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
      Left            =   5880
      TabIndex        =   16
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Periodo"
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
      Left            =   8520
      TabIndex        =   14
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
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
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Origen/Numero"
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
      Top             =   840
      Width           =   1575
   End
   Begin VB.Menu ajdu1 
      Caption         =   "&Add"
   End
   Begin VB.Menu bo712 
      Caption         =   "&Borrar"
   End
   Begin VB.Menu grba1 
      Caption         =   "&Grabar"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tvoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ajdu1_Click()
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
inicializa
codigo = ""
mes.Enabled = True
asiento.Enabled = True
codigo.Enabled = True
habilita 1
codigo.SetFocus

End Sub

Private Sub asiento_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(mes) = 0 Then Exit Sub
found = busca_asiento()
If found = 0 Then Exit Sub
codigo.SetFocus
End Sub

Private Sub asiento_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_asiento
End If


End Sub

Private Sub bo712_Click()
Dim found As Integer
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
found = busca_registro()
If found = 0 Then
   MsgBox "No existe registro", 48, "Aviso"
   Exit Sub
End If
found = borra_registro()
If found = 0 Then Exit Sub
MsgBox "Ok,Registro Borrado", 48, "Aviso"
codigo = ""
inicializa
codigo.SetFocus
End Sub

Private Sub buffer_DblClick()
Command1_Click
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   dlo132_Click
   Exit Sub
End If
Command1_Click

End Sub

Private Sub cmdAddEntry_Click()
ajdu1_Click
End Sub

Private Sub cmdDelete_Click()
bo712_Click
End Sub

Private Sub cmdExit_Click()
dlo132_Click

End Sub

Private Sub cmdPrint_Click()
djuer1_Click
End Sub

Private Sub cmdSave_Click()
'grba1_Click
End Sub

Private Sub cmdSort_Click()
Exit Sub
Frame1.Visible = True
buffer = ""
buffer.SetFocus
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   asiento.SetFocus
   Exit Sub
End If
If Len(mes) = 0 Then Exit Sub
If Len(asiento) = 0 Then
   codigo.SetFocus
   asiento.SetFocus
   Exit Sub
End If
If Len(codigo) = 0 Then Exit Sub
found = busca_registro()
If found = 0 Then
   inicializa
End If
sql_detalle
mes.Enabled = False
asiento.Enabled = False
codigo.Enabled = False
habilita 0
sumar_glosa
fecha.SetFocus
End Sub

Private Sub codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
'cmdSort_Click
End If
If KeyCode = &H26 Then
   asiento.SetFocus
   Exit Sub
End If

End Sub

Private Sub codigo1_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(codigo1) > 0 Then
   found = busca_cliente()
   If found = 0 Then
      'MsgBox "No existe Cliente", 48, "Aviso"
      'codigo1.SetFocus
      'Exit Sub
End If
End If
nombre1.SetFocus
End Sub

Private Sub codigo1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   fecha1.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_cliente
End If

End Sub

Private Sub Command1_Click()
Dim buf As String
Dim bufx As String
If tipoclie = "C" Then
   bufx = "clientes"
End If
If tipoclie = "P" Then
   bufx = "proveedo"
End If
If tipoclie = "V" Then
   bufx = "vendedor"
End If

   If opcion1 = "7" Then
      If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from " & bufx
      Else
      buf = "select Nombre,Codigo from   " & bufx & Combo1 & " like '" & buffer & "%'"
      End If
   End If

   If opcion1 = "5" Then
      If Len(buffer) = 0 Then
      buf = "select Descripcio,Tipo from cotipodo "
      Else
      buf = "select Descripcio,Tipo from cotipodo  " & Combo1 & " like '" & buffer & "%'"
      End If
   End If
   If opcion1 = "1" Then  'si es 3 debe ser subcuenta
      If Len(buffer) = 0 Then
      buf = "select Cuenta,Nombre,rnf as Tipo,Bd as T from mdh_plan where bd='3' order by cuenta "
      Else
      buf = "select Cuenta,Nombre,rnf as Tipo,Bd as T from mdh_plan where  bd='3' and " & Combo1 & " like '" & buffer & "%' order bu cuenta"
      End If
   End If
   If opcion1 = "2" Then
      If Len(buffer) = 0 Then
      buf = "select Descripcio,Origen from origen "
      Else
      buf = "select Descripcio,Origen from origen where  " & Combo1 & " like '" & buffer & "%'"
      End If
   End If
   'MsgBox buf
               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globalcont
               Data1.RecordSource = buf
               Data1.Refresh
               If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
                  Data1.Recordset.Close
                  buffer.SetFocus
                  Exit Sub
               End If
               If opcion1 = "1" Then
               dbGrid1.Columns(0).Width = 1000
               dbGrid1.Columns(1).Width = 5000
               dbGrid1.Columns(2).Width = 600
               dbGrid1.Columns(3).Width = 600
               End If
               If opcion1 = "2" Or opcion1 = "5" Then
               dbGrid1.Columns(0).Width = 4000
               dbGrid1.Columns(1).Width = 2000
               End If
               dbGrid1.SetFocus
End Sub



Private Sub Command2_Click()
grabar_glosa
End Sub

Private Sub Command3_Click()
dlo132_Click
End Sub

Private Sub cuenta_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   Frame2.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If
If Len(cuenta) = 0 Then Exit Sub
found = busca_cuenta("" & cuenta)
If found = 0 Then
   cuenta = ""
   MsgBox "Cuenta No existe", 48, "Aviso"
   Exit Sub
End If
debe.SetFocus
End Sub

Private Sub cuenta_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
consulta_cuenta
End If


End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim xtemp As Integer
Dim found As Integer
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
If opcion1 = "2" Then
   asiento = dbGrid1.Columns(1)
   Frame1.Visible = False
   asiento.SetFocus
   asiento_KeyPress 13
End If
If opcion1 = "5" Then
   tipo1 = dbGrid1.Columns(1)
   Frame1.Visible = False
   tipo1.SetFocus
   tipo1_KeyPress 13
End If
If opcion1 = "7" Then
   codigo1 = dbGrid1.Columns(1)
   Frame1.Visible = False
   codigo1.SetFocus
   codigo1_KeyPress 13
End If

If opcion1 = "1" Then
      If Len("" & dbGrid1.Columns(0)) > 0 Then
         cuenta = dbGrid1.Columns(0)
         Frame1.Visible = False
         cuenta.SetFocus
         cuenta_KeyPress 13
     End If
End If
End If
End Sub



Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim found As Integer
On Error GoTo cmd245_err
If KeyCode = &H2D Then  'inserte  uno nuevo
   inicializa_tipo1
   habilita3 0
   Check1.Value = 0
   Frame2.Visible = True
   inicializa_frame2
   found = busca_paridad()
   cuenta.SetFocus
End If
If KeyCode = &H2E Then  'borrar linea
   If MsgBox("Desea Borrar ", 1, "Aviso") <> 1 Then Exit Sub
   Data2.Recordset.Delete
End If
Exit Sub
cmd245_err:
Exit Sub

End Sub
Sub inicializa_frame2()
cuenta = ""
ncuenta = ""
debe = ""
haber = ""
moneda.Clear
moneda.AddItem "S"
moneda.AddItem "D"
moneda.ListIndex = 0
paridad = ""
glosa = ""
libro1.Clear
libro1.AddItem ""
libro1.AddItem "LibroCompras"
libro1.AddItem "LibroVentas"
libro1.AddItem "LibroRetenciones"
libro1.ListIndex = 0
End Sub

Private Sub debe_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
haber.SetFocus
End Sub

Private Sub debe_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   cuenta.SetFocus
   Exit Sub
End If

End Sub

Private Sub djuer1_Click()
If Frame2.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
reporgen.NAMETABLA = "cvoucher"
reporgen.Show 1

End Sub

Private Sub dlo132_Click()
If Frame1.Visible = True Then
   If opcion1 = "7" Then
      Frame1.Visible = False
      codigo1.SetFocus
      Exit Sub
   End If

   If opcion1 = "5" Then
      Frame1.Visible = False
      tipo1.SetFocus
      Exit Sub
   End If

   If opcion1 = "1" Then
      Frame1.Visible = False
      cuenta.SetFocus
      Exit Sub
   End If
   If opcion1 = "2" Then
      Frame1.Visible = False
      asiento.SetFocus
      Exit Sub
   End If
End If
If Frame2.Visible = True Then
   Frame2.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If


If codigo.Enabled = False Then
   codigo.Enabled = True
   asiento.Enabled = True
   habilita 1
   codigo.SetFocus
   Exit Sub
End If

tvoucher.Hide
Unload tvoucher
End Sub



Private Sub exonerado1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
otros1.SetFocus

End Sub

Private Sub exonerado1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   inafecto1.SetFocus
   Exit Sub
End If

End Sub

Private Sub fecha_KeyPress(KeyAscii As Integer)
Dim sdx As Double
If KeyAscii <> 13 Then Exit Sub
If Len(fecha) = 0 Then
   fecha = Format(Now, "dd/mm/yyyy")
End If
If Len(fecha) <> 10 Then
   fecha = ""
   Exit Sub
End If
If Not IsDate(fecha) Then Exit Sub
If Val(Mid$(mes, 1, 2)) <> Val(Mid$(fecha, 4, 2)) Then
   MsgBox "Periodo no programado ", 48, "Aviso"
   fecha = ""
   fecha.SetFocus
   Exit Sub
End If
If Val(Mid$(mes, 3, 4)) <> Val(Mid$(fecha, 7, 4)) Then
   MsgBox "Periodo no programado ", 48, "Aviso"
   fecha = ""
   fecha.SetFocus
   Exit Sub
End If
DBGrid2.SetFocus
End Sub

Private Sub fecha_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   habilita 1
   codigo.Enabled = True
   asiento.Enabled = True
   codigo.SetFocus
   Exit Sub
End If

End Sub

Private Sub fecha1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(fecha1) = 0 Then
   fecha1 = Format(Now, "dd/mm/yyyy")
End If
codigo1.SetFocus

End Sub

Private Sub fecha1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   numero1.SetFocus
   Exit Sub
End If

End Sub

Private Sub Form_Activate()
tipoclie.Clear
tipoclie.AddItem "C"
tipoclie.AddItem "P"
tipoclie.AddItem "V"
tipoclie.ListIndex = 0
mes = busca_parame()
habilita 1
End Sub

Private Sub Form_Load()
Dim found As Integer

Combo1.Clear
Combo1.AddItem "asiento"
Combo1.AddItem "numero"
Combo1.ListIndex = 0
borrar_temporal
sql_detalle
End Sub
Sub inicializa()
fecha = ""
End Sub
Function borra_registro()

Dim mytablex As Table

Set mytablex = mydbzglo.OpenTable("cvoucher")
mytablex.Index = "cvoucher"
mytablex.Seek "=", asiento, mes, codigo
If Not mytablex.NoMatch Then
   If MsgBox("Desea Borra el registro", 1, "Aviso") = "1" Then
      mytablex.Delete
      borra_registro = 1
   End If
End If
'------------------------------------- ------------
mytablex.Close
 

End Function
Function busca_registro()

Dim mytablex As Table
Dim mytableY As Table

Set mytablex = mydbzglo.OpenTable("cvoucher")
mytablex.Index = "cvoucher"
mytablex.Seek "=", asiento, mes, codigo
If Not mytablex.NoMatch Then
   pone_registro mytablex
   busca_registro = 1
End If
'------------------------------------- ------------
mytablex.Close
 
End Function
Sub pone_registro(mytablex As Table)
codigo = "" & mytablex.Fields("voucher")
asiento = "" & mytablex.Fields("origen")
mes = "" & mytablex.Fields("mes")
fecha = "" & mytablex.Fields("fecha")
End Sub
Sub grabando(mytablex As Table)
mytablex.Fields("voucher") = codigo
mytablex.Fields("mes") = mes
mytablex.Fields("origen") = asiento
mytablex.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
End Sub

Private Sub glosa_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Val(debe) = 0 And Val(haber) = 0 Then
   debe.SetFocus
   Exit Sub
End If
'If Check1.Value = 0 Then
'   grabar_glosa
'   Exit Sub
'End If
habilita3 0
tipo1.SetFocus
End Sub
Function valida1()
Dim found As Integer
If Len(cuenta) = 0 Then
   cuenta.SetFocus
   Exit Function
End If
found = busca_cuenta("" & cuenta)
If found = 0 Then
   cuenta = ""
   cuenta.SetFocus
   Exit Function
End If
If Val(debe) = 0 And Val(haber) = 0 Then
   debe.SetFocus
   Exit Function
End If
If Val(debe) > 0 And Val(haber) > 0 Then
   debe.SetFocus
   Exit Function
End If
If Check1.Value = 1 Then  'si debe grabarse detalle obligar el uso
   If Len(tipo1) = 0 Then
      tipo1.SetFocus
      Exit Function
   End If
   
   If Len(numero1) = 0 Then
      numero1.SetFocus
      Exit Function
   End If
   If Len(fecha1) = 0 Then
      fecha1.SetFocus
      Exit Function
   End If
   If Len(codigo1) = 0 Then
      codigo1.SetFocus
      Exit Function
   End If
End If
If Len(tipo1) > 0 Then
   found = busca_tipo()
   If found = 0 Then
      MsgBox "Tipo No existe", 48, "Aviso"
      tipo1.SetFocus
      Exit Function
   End If
End If
  
If Len(fecha1) > 0 Then
   If Not IsDate(fecha1) Then
      fecha1.SetFocus
      Exit Function
   End If
End If
If Len(codigo1) > 0 Then
   found = busca_cliente()
   If found = 0 Then
      'MsgBox "No existe Cliente", 48, "Aviso"
      'codigo1.SetFocus
      'Exit Function
   End If
End If

valida1 = 1
End Function

Private Sub glosa_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   paridad.SetFocus
   Exit Sub
End If

End Sub


Private Sub haber_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
moneda.SetFocus

End Sub

Private Sub haber_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   debe.SetFocus
   Exit Sub
End If

End Sub

Private Sub igv1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
inafecto1.SetFocus

End Sub

Private Sub igv1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   importe1.SetFocus
   Exit Sub
End If

End Sub

Private Sub importe1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
igv1.SetFocus

End Sub

Private Sub importe1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   libro1.SetFocus
   Exit Sub
End If

End Sub

Private Sub inafecto1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
exonerado1.SetFocus

End Sub

Private Sub inafecto1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   igv1.SetFocus
   Exit Sub
End If

End Sub

Private Sub Label1_Click()
cmdSort_Click
End Sub


Function grabar()
Dim found As Integer
Dim mytablex As Table

found = valida()
If found = 0 Then
   MsgBox "Campos invalidos", 48, "Aviso"
   Exit Function
End If

Set mytablex = mydbzglo.OpenTable("cvoucher")
mytablex.Index = "cvoucher"
mytablex.Seek "=", asiento, mes, codigo
If mytablex.NoMatch Then
   mytablex.AddNew
   grabando mytablex
   mytablex.Update
   found = graba_origen()
   grabar = 1
End If
If Not mytablex.NoMatch Then
   'If MsgBox("Desea Reescribir?", 1, "Aviso") = 1 Then
   mytablex.Edit
   grabando mytablex
   mytablex.Update
   grabar = 1
   'End If
End If
'------------------------------------- ------------
mytablex.Close
 
End Function

Function valida()
If Len(codigo) = 0 Then
   codigo.SetFocus
   Exit Function
End If
If Len(asiento) = 0 Then
   asiento.SetFocus
   Exit Function
End If
If Len(fecha) <> 10 Then
   fecha.SetFocus
   Exit Function
End If
If Not IsDate(fecha) Then
   fecha.SetFocus
   Exit Function
End If
valida = 1
End Function


Function busca_cuenta(buf As String)

Dim mytablex As Table
ncuenta = ""
habilita3 0

Set mytablex = mydbzglo.OpenTable("mdh_plan")
mytablex.Index = "mdh_plan"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   If "" & mytablex.Fields("bd") = "3" Then
      busca_cuenta = 1
      ncuenta = "" & mytablex.Fields("nombre")
      If "" & mytablex.Fields("cta") = "S" Then
         Check1.Value = 1
      End If
   End If
End If
mytablex.Close
 

End Function
Sub consulta_cuenta()
Combo1.Clear
Combo1.AddItem "Cuenta"
Combo1.AddItem "Nombre"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "1"
Command1_Click

End Sub
Sub habilita(sw As Integer)
Dim xsw As Variant
If sw = 0 Then
   xsw = True
End If
If sw = 1 Then
   xsw = False
End If
fecha.Enabled = xsw
DBGrid2.Enabled = xsw
End Sub
Sub sql_detalle()
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globalcont
               Data2.RecordSource = "select * from mdh_vou where t='" & asiento & "' and mes='" & mes & "' and vou='" & codigo & "'"
               Data2.Refresh
End Sub
Sub borrar_temporal()
'Dim mydbx As Database
'Set mydbx = OpenDatabase(globalcont, False, False, "foxpro 2.5;")
'mydbzglo.Execute "DELETE FROM tvoucher "
'
End Sub
Sub consulta_asiento()

Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Origen"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "2"
Command1_Click
End Sub
Function busca_asiento()

Dim mytablex As Table
Dim sdx As Double
Dim xmes As String

Set mytablex = mydbzglo.OpenTable("origen")
mytablex.Index = "origen"
mytablex.Seek "=", asiento
If Not mytablex.NoMatch Then
   busca_asiento = 1
   sdx = 0
   Select Case Mid$(mes, 1, 2)
          Case "01"
               sdx = Val("" & mytablex.Fields("enero")) + 1
          Case "02"
          sdx = Val("" & mytablex.Fields("febrero")) + 1
          Case "03"
          sdx = Val("" & mytablex.Fields("marzo")) + 1
          Case "04"
          sdx = Val("" & mytablex.Fields("abril")) + 1
          Case "05"
          sdx = Val("" & mytablex.Fields("mayo")) + 1
          Case "06"
          sdx = Val("" & mytablex.Fields("junio")) + 1
          Case "07"
          sdx = Val("" & mytablex.Fields("julio")) + 1
          Case "08"
          sdx = Val("" & mytablex.Fields("agosto")) + 1
          Case "09"
          sdx = Val("" & mytablex.Fields("setiembre")) + 1
          Case "10"
          sdx = Val("" & mytablex.Fields("octubre")) + 1
          Case "11"
          sdx = Val("" & mytablex.Fields("noviembre")) + 1
          Case "12"
          sdx = Val("" & mytablex.Fields("diciembre")) + 1
   End Select
   codigo = Format(sdx, "0")
End If
mytablex.Close
 
End Function
Function busca_parame() As String

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("parame")
mytablex.Index = "codigo"
mytablex.Seek "=", "01"
If Not mytablex.NoMatch Then
   busca_parame = "" & mytablex.Fields("mesconta") & "" & mytablex.Fields("anoconta")
End If
mytablex.Close
 

End Function


Private Sub libro1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
importe1.SetFocus

End Sub

Private Sub libro1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   codigo1.SetFocus
   Exit Sub
End If

End Sub


Private Sub moneda_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
paridad.SetFocus

End Sub

Private Sub moneda_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   haber.SetFocus
   Exit Sub
End If

End Sub

Private Sub nombre1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   codigo1.SetFocus
   Exit Sub
End If

End Sub

Private Sub numero1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
fecha1.SetFocus
End Sub

Private Sub numero1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   tipo1.SetFocus
   Exit Sub
End If

End Sub

Private Sub otros1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   exonerado1.SetFocus
   Exit Sub
End If

End Sub

Private Sub paridad_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(paridad) = 0 Then
   found = busca_paridad()
   If found = 0 Then
      MsgBox "No existe tipo de cambio ", 48, "Aviso"
      paridad.SetFocus
      Exit Sub
   End If
End If
glosa.SetFocus
End Sub

Private Sub paridad_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   moneda.SetFocus
   Exit Sub
End If

End Sub
Function graba_cuenta()
   Data2.Recordset.AddNew
   Data2.Recordset.Fields("t") = asiento
   Data2.Recordset.Fields("vou") = codigo
   Data2.Recordset.Fields("mes") = mes
   Data2.Recordset.Fields("cuenta") = cuenta
   Data2.Recordset.Fields("debe") = Val(debe)
   Data2.Recordset.Fields("haber") = Val(haber)
   Data2.Recordset.Fields("glosa") = glosa
   Data2.Recordset.Fields("ncuenta") = ncuenta
   Data2.Recordset.Fields("fecha") = Format(fecha, "dd/mm/yyyy")
   If Len(glosa) = 0 Then
   Data2.Recordset.Fields("glosa") = ncuenta
   End If
   Data2.Recordset.Fields("moneda") = moneda
   Data2.Recordset.Fields("tc") = Val(paridad)
   Data2.Recordset.Fields("doc") = tipo1
   Data2.Recordset.Fields("numero") = numero1
   Data2.Recordset.Fields("rut") = codigo1
   Data2.Recordset.Fields("rs") = nombre1
   Data2.Recordset.Update
   sumar_glosa
End Function

Private Sub tipo1_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(tipo1) > 0 Then
found = busca_tipo()
If found = 0 Then
   MsgBox "Tipo No existe", 48, "Aviso"
   Exit Sub
End If
End If
numero1.SetFocus
End Sub

Private Sub tipo1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   glosa.SetFocus
   Exit Sub
End If

If KeyCode = &H70 Then  'f1
   consulta_tipo1
End If

End Sub
Sub consulta_tipo1()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Tipo"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "5"
Command1_Click
End Sub
Sub grabar_glosa()
Dim found As Integer
found = grabar()
If found = 0 Then
   Exit Sub
End If
found = valida1()
If found = 0 Then Exit Sub
graba_cuenta
sql_detalle
inicializa_tipo1
dlo132_Click
End Sub
Sub habilita3(sw As Integer)
Dim xsw
If sw = 0 Then
   xsw = True
End If
If sw = 1 Then
   xsw = False
End If
tipo1.Enabled = xsw
numero1.Enabled = xsw
fecha1.Enabled = xsw
codigo1.Enabled = xsw
nombre1.Enabled = xsw
libro1.Enabled = xsw
importe1.Enabled = xsw
igv1.Enabled = xsw
inafecto1.Enabled = xsw
otros1.Enabled = xsw

End Sub
Function busca_paridad()

Dim mytablex As Table
Dim mytableY As Table

Set mytablex = mydbzglo.OpenTable("tcambio")
mytablex.Index = "tcambio"
mytablex.Seek "=", fecha
If Not mytablex.NoMatch Then
   busca_paridad = 1
   paridad = "" & mytablex.Fields("compra")
End If
'------------------------------------- ------------
mytablex.Close
 
End Function
Function busca_tipo()

Dim mytablex As Table
Dim mytableY As Table

Set mytablex = mydbzglo.OpenTable("cotipodo")
mytablex.Index = "cotipodo"
mytablex.Seek "=", tipo1
If Not mytablex.NoMatch Then
   busca_tipo = 1
End If
mytablex.Close
 
End Function
Sub consulta_cliente()
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Codigo"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "7"
Command1_Click
End Sub
Function busca_cliente()

Dim mytablex As Table
Dim mytableY As Table
Dim bufx As String
If tipoclie = "C" Then
   bufx = "clientes"
End If
If tipoclie = "P" Then
   bufx = "proveedo"
End If
If tipoclie = "V" Then
   bufx = "vendedor"
End If
nombre1 = ""

Set mytablex = mydbzglo.OpenTable(bufx)
mytablex.Index = "codigo"
mytablex.Seek "=", codigo1
If Not mytablex.NoMatch Then
   busca_cliente = 1
   nombre1 = "" & mytablex.Fields("nombre")
End If
mytablex.Close
 
End Function
Function graba_origen()

Dim mytablex As Table
Dim sdx As Double
Dim xmes As String

Set mytablex = mydbzglo.OpenTable("origen")
mytablex.Index = "origen"
mytablex.Seek "=", asiento
If Not mytablex.NoMatch Then
   mytablex.Edit
   
   Select Case Mid$(mes, 1, 2)
          Case "01"
               mytablex.Fields("enero") = codigo
          Case "02"
               mytablex.Fields("febrero") = codigo
          Case "03"
               mytablex.Fields("marzo") = codigo
          Case "04"
               mytablex.Fields("abril") = codigo
          Case "05"
               mytablex.Fields("mayo") = codigo
          Case "06"
               mytablex.Fields("junio") = codigo
          Case "07"
               mytablex.Fields("julio") = codigo
          Case "08"
               mytablex.Fields("agosto") = codigo
          Case "09"
               mytablex.Fields("setiembre") = codigo
          Case "10"
               mytablex.Fields("octubre") = codigo
          Case "11"
          mytablex.Fields("noviembre") = codigo
          Case "12"
          mytablex.Fields("diciembre") = codigo
   End Select
   mytablex.Update
End If
mytablex.Close
 
End Function
Sub inicializa_tipo1()
tipo1 = ""
numero1 = ""
fecha1 = ""
codigo1 = ""
nombre1 = ""
tipoclie.ListIndex = 0

End Sub
Sub ir_ïnicio()
On Error GoTo cmd342_err
Data2.Recordset.MoveFirst
Exit Sub
cmd342_err:
Exit Sub
End Sub
Sub sumar_glosa()
Dim sdx2 As Double
Dim sdx1 As Double
Dim sdx3 As Double
Dim sdx4 As Double
sdx1 = 0
sdx2 = 0
sdx3 = 0
sdx4 = 0
ir_ïnicio
Do
If Data2.Recordset.EOF Then Exit Do
If "" & Data2.Recordset.Fields("moneda") = "S" Then
sdx1 = sdx1 + Val("" & Data2.Recordset.Fields("debe"))
sdx2 = sdx2 + Val("" & Data2.Recordset.Fields("haber"))
End If
If "" & Data2.Recordset.Fields("moneda") = "D" Then
sdx3 = sdx1 + Val("" & Data2.Recordset.Fields("debe"))
sdx4 = sdx2 + Val("" & Data2.Recordset.Fields("haber"))
End If
Data2.Recordset.MoveNext
Loop
tdebes = Format(sdx1, "0.00")
tdebed = Format(sdx3, "0.00")
thabers = Format(sdx2, "0.00")
thaberd = Format(sdx4, "0.00")
End Sub

