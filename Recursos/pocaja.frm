VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form pocaja 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caja Pocket"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   3255
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Up"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Dwn"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   66
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   65
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Cls"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   64
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Fin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   45
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Grab"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   44
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton groupmesa 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton groupsalon 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   3
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2400
         Width           =   375
      End
      Begin VB.CommandButton groupsalon 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   2
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2400
         Width           =   375
      End
      Begin VB.CommandButton groupsalon 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   1
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton groupsalon 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2040
         Width           =   375
      End
      Begin MSDataGridLib.DataGrid dbgrid1 
         Height          =   1695
         Left            =   0
         TabIndex        =   63
         Top             =   120
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   2990
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   4
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
            DataField       =   "Cantidad"
            Caption         =   "Cant"
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
            DataField       =   "Total"
            Caption         =   "Total"
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
         BeginProperty Column03 
            DataField       =   "Precio"
            Caption         =   "Prec"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1890.142
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   390.047
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   540.284
            EndProperty
         EndProperty
      End
      Begin VB.Label xtotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   72
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "->"
         Height          =   375
         Left            =   3000
         TabIndex        =   71
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<-"
         Height          =   375
         Left            =   3000
         TabIndex        =   70
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "->"
         Height          =   375
         Left            =   0
         TabIndex        =   69
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<-"
         Height          =   375
         Left            =   0
         TabIndex        =   68
         Top             =   2400
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Height          =   3255
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton zproducto 
         BackColor       =   &H00FFFF80&
         Height          =   375
         Index           =   7
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton zproducto 
         BackColor       =   &H00FFFF80&
         Height          =   375
         Index           =   6
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   2520
         Width           =   1575
      End
      Begin VB.CommandButton zproducto 
         BackColor       =   &H00FFFF80&
         Height          =   375
         Index           =   5
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton zproducto 
         BackColor       =   &H00FFFF80&
         Height          =   375
         Index           =   4
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton zproducto 
         BackColor       =   &H00FFFF80&
         Height          =   375
         Index           =   3
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton zproducto 
         BackColor       =   &H00FFFF80&
         Height          =   375
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CommandButton zproducto 
         BackColor       =   &H00FFFF80&
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cantidad 
         BackColor       =   &H80000009&
         Height          =   375
         Index           =   11
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cantidad 
         BackColor       =   &H80000009&
         Height          =   375
         Index           =   10
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cantidad 
         BackColor       =   &H80000009&
         Height          =   375
         Index           =   9
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cantidad 
         BackColor       =   &H80000009&
         Height          =   375
         Index           =   8
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cantidad 
         BackColor       =   &H80000009&
         Height          =   375
         Index           =   7
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cantidad 
         BackColor       =   &H80000009&
         Height          =   375
         Index           =   6
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cantidad 
         BackColor       =   &H80000009&
         Height          =   375
         Index           =   5
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cantidad 
         BackColor       =   &H80000009&
         Height          =   375
         Index           =   4
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cantidad 
         BackColor       =   &H80000009&
         Height          =   375
         Index           =   3
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cantidad 
         BackColor       =   &H80000009&
         Height          =   375
         Index           =   2
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cantidad 
         BackColor       =   &H80000009&
         Height          =   375
         Index           =   1
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cantidad 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton zfamilia 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton zfamilia 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton zfamilia 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CommandButton zproducto 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton zfamilia 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton zfamilia 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton zfamilia 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton zfamilia 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton zfamilia 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton zfamilia 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pedido"
         Height          =   255
         Left            =   2280
         TabIndex        =   62
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   61
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   60
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fn"
         Height          =   255
         Left            =   1920
         TabIndex        =   57
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ant"
         Height          =   255
         Left            =   1200
         TabIndex        =   51
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sig"
         Height          =   255
         Left            =   1560
         TabIndex        =   50
         Top             =   2880
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ant"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sig"
         Height          =   255
         Left            =   480
         TabIndex        =   48
         Top             =   120
         Width           =   375
      End
      Begin VB.Label xcantidad 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   255
         Left            =   2400
         TabIndex        =   37
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Clave"
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.TextBox clave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enter"
         Height          =   615
         Index           =   0
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         Height          =   615
         Index           =   11
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1"
         Height          =   615
         Index           =   1
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "2"
         Height          =   615
         Index           =   2
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "3"
         Height          =   615
         Index           =   3
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "4"
         Height          =   615
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "5"
         Height          =   615
         Index           =   5
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "6"
         Height          =   615
         Index           =   6
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "7"
         Height          =   615
         Index           =   7
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "8"
         Height          =   615
         Index           =   8
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "9"
         Height          =   615
         Index           =   9
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CR"
         Height          =   615
         Index           =   10
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label vendedor 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2880
         Width           =   105
      End
      Begin VB.Label conectado 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2640
         Width           =   105
      End
   End
   Begin VB.Label mesa 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   81
      Top             =   3600
      Width           =   45
   End
   Begin VB.Label salon 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   73
      Top             =   3360
      Width           =   45
   End
End
Attribute VB_Name = "pocaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mytablexx As New ADODB.Recordset
Dim mfamcod(15000) As String
Dim mfampag As Integer
Dim mfamtop As Integer

Dim mprodcod(15000) As String
Dim wprodcod(15000) As String
Dim wwprodcod(30) As String
Dim mprodpag As Integer
Dim mprodtop As Integer

Dim msalcod(100) As String
Dim msalpag As Integer
Dim msaltop As Integer


Dim mmesacod(15000) As String
Dim wmesacod(15000) As String
Dim wwmesacod(30) As String
Dim mmesapag As Integer
Dim mmesatop As Integer


Private Sub Command1_Click(Index As Integer)
Dim found As Integer
If Index = 10 Then
          clave = ""
          Exit Sub
       End If
If Index = 0 Then  'enter
If cn.State = adStateOpen Then ' si esta abierta
   cn.Close ' cierro
End If
found = conectarpo()
If found = 0 Then
   MsgBox "NO hay COnexion base datos ", 48, "Aviso"
   End
   Exit Sub
End If
If Len(clave) = 0 Then
   clave.SetFocus
   Exit Sub
End If
   found = busca_clave()
   If found = 0 Then
      MsgBox "Clave no Valido ", 48, "Aviso"
      clave = ""
      clave.SetFocus
      Exit Sub
   End If
   Frame1.Visible = False
   Frame2.Visible = True
   carga_familia
   If mytablexx.State = 1 Then
      mytablexx.Close
   End If
   
   mytablexx.Open "SELECT * FROM tmpocket where vendedor='" & Trim(vendedor) & "'", cn, adOpenKeyset, adLockOptimistic
   Set dbgrid1.DataSource = mytablexx
   dbgrid1.Refresh
           Exit Sub
End If
       clave = clave & Command1(Index).Caption

End Sub

Function busca_clave()
Dim mytablex As New ADODB.Recordset
   mytablex.Open "SELECT * FROM vendedor where clave='" & Trim(clave) & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount > 0 Then  'si existe
      busca_clave = 1
      vendedor = "" & mytablex.Fields("codigo")
   End If
   mytablex.Close

End Function

Public Function conectarpo()
Dim dbuser As String
Dim dbpassword As String
Dim dbname As String
Dim dbserver As String
On Error GoTo cmd1_error

 cn.CursorLocation = adUseClient
 cn.CommandTimeout = 1024
 'cn.Open "Driver={SQL Server};Server=" & xvservidor & ";Database=calipso;Uid=" & xusuario & ";pwd=" & xclave
 'cn.Open "Driver={SQL Server};Server=" & "(local)" & ";Database=calipso;Uid=sa;pwd=123 "
 cn.Open "Driver={SQL Server};Server=(local);Database=calipso;Uid=sa;pwd=123"
 
 conectarpo = 1
 conectado = "S"
 Exit Function
cmd1_error:
 MsgBox " " & error$, 48, "Aviso"
 Exit Function
 End Function

Private Sub Command2_Click()

End Sub

Sub menu_familia(buf As String)
Dim i As Integer
Dim j As Integer
Select Case buf
       Case "INI"
            mfampag = 0
       Case "SIG"
            mfampag = mfampag + 7
            If mfampag > 102 Then
               mfampag = 0
            End If
       Case "ANT"
            mfampag = mfampag - 7
            If mfampag < 0 Then
               mfampag = 0
            End If
End Select
j = -1
For i = mfampag To 7 + mfampag
    j = j + 1
    zfamilia(j).Caption = mfamcod(i)
Next i

End Sub
Sub carga_familia()
Dim mytablex As New ADODB.Recordset
Dim i As Integer
For i = 0 To 14999
    mfamcod(i) = ""
Next i

i = -1
mytablex.Open "select * from familia where vetouch='S'", cn, adOpenStatic, adLockOptimistic

Do
If mytablex.EOF Then Exit Do
i = i + 1
mfamcod(i) = "" & mytablex.Fields("familia")
mytablex.MoveNext
Loop
mfamtop = i
mytablex.Close
mfampag = 0
menu_familia "INI"

End Sub


Private Sub Command3_Click()

End Sub

Private Sub Command10_Click()
If Len(salon) = 0 Then Exit Sub
If Len(mesa) = 0 Then Exit Sub
If mytablexx.RecordCount = 0 Then
   MsgBox "No existen datos ", 48, "Aviso"
   Exit Sub
End If

   If mytablexx.State = 1 Then
      mytablexx.Close
   End If
   cn.Execute ("update tmpocket set salon='" & Trim(salon) & "',mesa='" & mesa & "' where vendedor='" & vendedor & "'")
   
   mytablexx.Open "SELECT * FROM tmpocket where vendedor='" & Trim(vendedor) & "'", cn, adOpenKeyset, adLockOptimistic
   Set dbgrid1.DataSource = mytablexx
   dbgrid1.Refresh
   sumar
   adicionar_comanda
   inicializa_todo
   Frame4.Visible = False
   
   
End Sub
Sub adicionar_comanda()
Dim mytablex As New ADODB.Recordset
If mytablexx.RecordCount = 0 Then Exit Sub
mytablexx.MoveFirst
mytablex.Open "SELECT * FROM dcomanda where vendedor='" & Trim(vendedor) & "'", cn, adOpenKeyset, adLockOptimistic
Do
If mytablexx.EOF Then Exit Do
    mytablex.AddNew
    mytablex.Fields("salon") = Trim("" & mytablexx.Fields("salon"))
    mytablex.Fields("mesa") = Trim("" & mytablexx.Fields("mesa"))
    mytablex.Fields("producto") = Trim("" & mytablexx.Fields("producto"))
    mytablex.Fields("descripcio") = Trim("" & mytablexx.Fields("descripcio"))
    mytablex.Fields("unidad") = Trim("" & mytablexx.Fields("unidad"))
    mytablex.Fields("factor") = Val("" & mytablexx.Fields("producto"))
    mytablex.Fields("precio") = Val("" & mytablexx.Fields("precio"))
    mytablex.Fields("total") = Val("" & mytablexx.Fields("total"))
    mytablex.Fields("cantidad") = Val("" & mytablexx.Fields("cantidad"))
    mytablex.Fields("precio") = Val("" & mytablexx.Fields("precio"))
    mytablex.Fields("igv") = 18
    mytablex.Fields("local") = "01"
    mytablex.Fields("tipo") = ""
    mytablex.Fields("serie") = ""
    mytablex.Fields("numero") = ""
    mytablex.Fields("vendedor") = Trim("" & mytablexx.Fields("vendedor"))
    mytablex.Fields("moneda") = "S"
    mytablex.Fields("bodega") = "01"
    mytablex.Fields("bodegaf") = "" 'xruc '"" & mytable11.Fields("bodega")
    mytablex.Fields("acu") = ""
    mytablex.Fields("localf") = ""
    mytablex.Fields("tipoclie") = "C"
    mytablex.Fields("acu1") = ""
    mytablex.Fields("servicio") = "C"
    mytablex.Fields("flage") = ""
    mytablex.Fields("codigo") = ""
    mytablex.Fields("caja") = ""
    mytablex.Fields("turno") = ""
    mytablex.Fields("usuario") = ""
    mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("hora") = Format(Now, "hh:MM")
    mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("estado") = "2"
    mytablex.Fields("numero") = "1"
    mytablex.Update
    mytablexx.MoveNext
Loop
mytablex.Close
End Sub
Sub inicializa_todo()
   If mytablexx.State = 1 Then
      mytablexx.Close
   End If
   cn.Execute ("delete from tmpocket where vendedor='" & vendedor & "'")
   mytablexx.Open "SELECT * FROM tmpocket where vendedor='" & Trim(vendedor) & "'", cn, adOpenKeyset, adLockOptimistic
   Set dbgrid1.DataSource = mytablexx
   dbgrid1.Refresh
   sumar
   salon = ""
   mesa = ""

End Sub

Private Sub Command11_Click()
Frame4.Visible = False
Frame2.Visible = True

End Sub

Private Sub Command6_Click()
End Sub

Private Sub Command8_Click()
borrar_linea
End Sub

Private Sub groupmesa_Click(Index As Integer)
Dim i As Integer
Dim k As Integer
For i = 0 To 7
    groupmesa(i).BackColor = &HFFFFFF
Next i
If Len(groupmesa(Index).Caption) = 0 Then Exit Sub
groupmesa(Index).BackColor = &HFF&

'---------------------------------------


If groupsalon(0).BackColor = &HFFFFFF And groupsalon(1).BackColor = &HFFFFFF And groupsalon(2).BackColor = &HFFFFFF Then
   Exit Sub
End If

k = 0
If groupsalon(0).BackColor = &HFF& Then
   k = 0
End If
If groupsalon(1).BackColor = &HFF& Then
   k = 1
End If
'MsgBox "Hola"
mesa = groupmesa(Index).Caption

End Sub

Private Sub groupsalon_Click(Index As Integer)
Dim i As Integer
If Len(groupsalon(Index).Caption) = 0 Then Exit Sub
For i = 0 To 3
  groupsalon(i).BackColor = &HFFFFFF
Next i
For i = 0 To 7
  groupmesa(i).BackColor = &HFFFFFF
Next i
groupsalon(Index).BackColor = &HFF&
menu_carga_mesa groupsalon(Index).Caption
menu_mesa "INI", groupsalon(Index).Caption
salon = groupsalon(Index).Caption
mesa = ""

End Sub

Private Sub Label1_Click()
menu_familia "SIG"
End Sub

Private Sub Label10_Click()
Dim i As Integer
For i = 0 To 3
    groupsalon(i).BackColor = &HFFFFFF
    salon = ""
    mesa = ""
Next i
For i = 0 To 23
    groupmesa(i).Caption = ""
    groupmesa(i).BackColor = &HFFFFFF
Next i

menu_salon "SIG"

End Sub

Private Sub Label11_Click()
Dim i As Integer
For i = 0 To 7
    groupmesa(i).BackColor = &HFFFFFF
    mesa = ""
Next i

menu_mesa "ANT", salon

End Sub

Private Sub Label12_Click()
Dim i As Integer
For i = 0 To 7
    groupmesa(i).BackColor = &HFFFFFF
    mesa = ""
Next i

menu_mesa "SIG", salon

End Sub

Private Sub Label2_Click()
menu_familia "ANT"
End Sub

Private Sub Label3_Click()
menu_producto "SIG"
End Sub

Private Sub Label4_Click()
menu_producto "ANT"
End Sub

Private Sub Label5_Click()
clave = ""
conectado = ""
Frame1.Visible = True
Frame2.Visible = False
clave.SetFocus

End Sub

Private Sub Label6_Click()
Dim sdx As Double
sdx = Val(xcantidad) + 1
xcantidad = Format(sdx, "0")
End Sub

Private Sub Label7_Click()
Dim sdx As Double
sdx = Val(xcantidad) - 1
If sdx <= 0 Then
   sdx = 1
End If
xcantidad = Format(sdx, "0")
End Sub

Private Sub Label8_Click()
Dim i As Integer

Frame4.Visible = True
carga_salon
For i = 0 To 2
'If Trim(groupsalon(i).Caption) = Trim("" & mytable11.Fields("salon")) Then
'   groupsalon_Click i
'   Exit For
'End If
Next i

dbgrid1.SetFocus
End Sub

Private Sub Label9_Click()
Dim i As Integer
For i = 0 To 3
    groupsalon(i).BackColor = &HFFFFFF
    salon = ""
    mesa = ""
Next i
For i = 0 To 23
    groupmesa(i).Caption = ""
    groupmesa(i).BackColor = &HFFFFFF
Next i

menu_salon "ANT"

End Sub

Private Sub zfamilia_Click(Index As Integer)
menu_carga_producto zfamilia(Index).Caption
menu_producto "INI"

End Sub
Sub menu_carga_producto(buf As String)
Dim mytablex As New ADODB.Recordset

Dim i As Integer
For i = 0 To 7
   wwprodcod(i) = ""
Next i
For i = 0 To 14999
    mprodcod(i) = ""
    wprodcod(i) = ""
Next i

i = -1

If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM producto where familia='" & buf & "' order by touch ", cn, adOpenDynamic, adLockOptimistic

Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("familia") = buf Then
   i = i + 1
   mprodcod(i) = "" & mytablex.Fields("descripcio")
   wprodcod(i) = "" & mytablex.Fields("producto")
   Else: Exit Do
End If
mytablex.MoveNext
Loop

mytablex.Close
mprodtop = i
mprodpag = 0

End Sub
Sub menu_producto(buf As String)
Dim i As Integer
Dim j As Integer

Select Case buf
       Case "INI"
            mprodpag = 0
       Case "SIG"
            mprodpag = mprodpag + 7
            If mprodpag > 102 Then
               mprodpag = 0
            End If
       Case "ANT"
            mprodpag = mprodpag - 7
            If mprodpag < 0 Then
               mprodpag = 0
            End If
End Select
j = -1
For i = mprodpag To 7 + mprodpag
    j = j + 1
    zproducto(j).Caption = mprodcod(i)
    wwprodcod(j) = wprodcod(i)
Next i


End Sub

Private Sub zproducto_Click(Index As Integer)
Dim buff As String
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
           buff = "" & wwprodcod(Index)
            If mytablex.State = 1 Then mytablex.Close
            mytablex.Open "SELECT * FROM producto where producto='" & buff & "'", cn, adOpenDynamic, adLockOptimistic
            If mytablex.RecordCount = 0 Then
               mytablex.Close
               Exit Sub
            End If
            If mytabley.State = 1 Then mytabley.Close
            mytabley.Open "SELECT * FROM precios where producto='" & buff & "' and local='01'", cn, adOpenDynamic, adLockOptimistic
            If mytabley.RecordCount = 0 Then
               mytabley.Close
               mytablex.Close
               Exit Sub
            End If
            If Val(xcantidad) <= 0 Then
               xcantidad = "1"
            End If
            
            mytablexx.AddNew
            mytablexx.Fields("vendedor") = Trim(vendedor)
            mytablexx.Fields("producto") = "" & mytablex.Fields("producto")
            mytablexx.Fields("descripcio") = "" & mytablex.Fields("descripcio")
            mytablexx.Fields("unidad") = "" & mytabley.Fields("unidad1")
            mytablexx.Fields("factor") = Val("" & mytabley.Fields("factor1"))
            mytablexx.Fields("cantidad") = Val(xcantidad)
            mytablexx.Fields("precio") = Val("" & mytabley.Fields("PVENTA1"))
            mytablexx.Fields("total") = Val(xcantidad) * Val("" & mytabley.Fields("pventa1"))
            mytablexx.Update
            mytablex.Close
            mytabley.Close
            sumar
            
 
End Sub
Sub sumar()
Dim sdx As Double
If mytablexx.RecordCount = 0 Then Exit Sub
mytablexx.MoveFirst
sdx = 0
xtotal = ""
Do
If mytablexx.EOF Then Exit Do
sdx = sdx + Val("" & mytablexx.Fields("total"))
mytablexx.MoveNext
Loop
xtotal = Format(sdx, "0.00")
End Sub
Sub carga_salon()
Dim mytablex As New ADODB.Recordset
Dim i As Integer
For i = 0 To 99
    msalcod(i) = ""
Next i

i = -1
mytablex.Open "select * from salon ", cn, adOpenStatic, adLockOptimistic

Do
If mytablex.EOF Then Exit Do
i = i + 1
msalcod(i) = "" & mytablex.Fields("salon")
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

Sub verifica_mesas(indx As Integer, buf As String, buf1 As String)
Dim mytablex As New ADODB.Recordset
groupmesa(indx).BackColor = &HFFFFFF
If Len(buf1) > 0 And Len(buf) > 0 Then
   mytablex.Open "select * from dcomanda where salon='" & buf & "' and mesa='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      groupmesa(indx).BackColor = &HFF00&
   End If
   mytablex.Close
End If
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
mytablex.Open "SELECT * FROM mesa where salon='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

Do
If mytablex.EOF Then Exit Do
If "" & mytablex.Fields("salon") = buf Then
   i = i + 1
   mmesacod(i) = "" & mytablex.Fields("mesa")
   wmesacod(i) = "" & mytablex.Fields("mesa")
   Else: Exit Do
End If
mytablex.MoveNext
Loop

mytablex.Close
mmesatop = i
mmesapag = 0

End Sub
Sub menu_mesa(buf As String, buf1 As String)
Dim i As Integer
Dim j As Integer
Select Case buf
       Case "INI"
            mmesapag = 0
       Case "SIG"
            mmesapag = mmesapag + 7
            If mmesapag > 102 Then
               mmesapag = 0
            End If
       Case "ANT"
            mmesapag = mmesapag - 7
            If mmesapag < 0 Then
               mmesapag = 0
            End If
End Select
j = -1
For i = mmesapag To 7 + mmesapag
    j = j + 1
    groupmesa(j).Caption = mmesacod(i)
    verifica_mesas j, buf1, groupmesa(j).Caption
Next i

End Sub
Sub borrar_linea()
On Error GoTo cmd9000_err
mytablexx.Delete
Exit Sub
cmd9000_err:
'msgbox "Seleccione un dato ",48,"Aviso"
Exit Sub

End Sub
