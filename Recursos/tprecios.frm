VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form tprecios 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precios por Local"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   11985
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Copiar data Locales"
      Height          =   3975
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   6975
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
         Left            =   5520
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tprecios.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Borrar registro"
         Top             =   240
         Width           =   1215
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
         Left            =   5520
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tprecios.frx":1212
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Grabar registro"
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desde"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desde"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Precios Mantenimiento"
      Height          =   5535
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   8655
      Begin VB.TextBox unidad1 
         BackColor       =   &H00C0FFFF&
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
         MaxLength       =   6
         TabIndex        =   84
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox factor1 
         BackColor       =   &H00C0FFFF&
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
         Left            =   960
         MaxLength       =   4
         TabIndex        =   83
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox pventa1 
         BackColor       =   &H00C0FFFF&
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   82
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox margen1 
         BackColor       =   &H00C0FFFF&
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
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   81
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox fechai11 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   80
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox fechaf11 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   79
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox fechaid 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   78
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox fechafd 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   77
         Top             =   4440
         Width           =   1455
      End
      Begin VB.TextBox margen11 
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
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   76
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox pventa11 
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
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   75
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox maximo11 
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
         Left            =   4080
         MaxLength       =   5
         TabIndex        =   74
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox minimo11 
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
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   73
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox dscto 
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
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   72
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox margen2 
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
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   71
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox pventa2 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   70
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox factor2 
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
         Left            =   960
         MaxLength       =   4
         TabIndex        =   69
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox unidad2 
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   68
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox margen3 
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
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   67
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox pventa3 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   66
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox factor3 
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
         Left            =   960
         MaxLength       =   4
         TabIndex        =   65
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox unidad3 
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   64
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox margen4 
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
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   63
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox pventa4 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   62
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox factor4 
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
         Left            =   960
         MaxLength       =   4
         TabIndex        =   61
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox unidad4 
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   60
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox margen5 
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
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   59
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox pventa5 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   58
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox factor5 
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
         Left            =   960
         MaxLength       =   4
         TabIndex        =   57
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox unidad5 
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   56
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox margen6 
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
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   55
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox pventa6 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   54
         Top             =   3240
         Width           =   975
      End
      Begin VB.TextBox factor6 
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
         Left            =   960
         MaxLength       =   4
         TabIndex        =   53
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox unidad6 
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   52
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox margen7 
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
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   51
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox pventa7 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   50
         Top             =   3600
         Width           =   975
      End
      Begin VB.TextBox factor7 
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
         Left            =   960
         MaxLength       =   4
         TabIndex        =   49
         Top             =   3600
         Width           =   735
      End
      Begin VB.TextBox unidad7 
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   48
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox margen8 
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
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   47
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox pventa8 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   46
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox factor8 
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
         Left            =   960
         MaxLength       =   4
         TabIndex        =   45
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox unidad8 
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   44
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox margen9 
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
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   43
         Top             =   4320
         Width           =   855
      End
      Begin VB.TextBox pventa9 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   42
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox factor9 
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
         Left            =   960
         MaxLength       =   4
         TabIndex        =   41
         Top             =   4320
         Width           =   735
      End
      Begin VB.TextBox unidad9 
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   40
         Top             =   4320
         Width           =   855
      End
      Begin VB.TextBox margen10 
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
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   39
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox pventa10 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   38
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox factor10 
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
         Left            =   960
         MaxLength       =   4
         TabIndex        =   37
         Top             =   4680
         Width           =   735
      End
      Begin VB.TextBox unidad10 
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   36
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox minimo12 
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
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   35
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox maximo12 
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
         Left            =   4080
         MaxLength       =   5
         TabIndex        =   34
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox pventa12 
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
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   33
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox margen12 
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
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   32
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox minimo13 
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
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   31
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox maximo13 
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
         Left            =   4080
         MaxLength       =   5
         TabIndex        =   30
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox pventa13 
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
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   29
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox margen13 
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
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   28
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox minimo14 
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
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   27
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox maximo14 
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
         Left            =   4080
         MaxLength       =   5
         TabIndex        =   26
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox pventa14 
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
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   25
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox margen14 
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
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   24
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox minimo15 
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
         Left            =   3600
         MaxLength       =   5
         TabIndex        =   23
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox maximo15 
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
         Left            =   4080
         MaxLength       =   5
         TabIndex        =   22
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox pventa15 
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
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   21
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox margen15 
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
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   20
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox ccosto 
         BackColor       =   &H00C0FFFF&
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
         Left            =   5280
         MaxLength       =   6
         TabIndex        =   19
         Top             =   720
         Width           =   1215
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
         Height          =   1095
         Left            =   7200
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tprecios.frx":2424
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Borrar registro"
         Top             =   1920
         Width           =   975
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
         Height          =   1095
         Left            =   7200
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tprecios.frx":3636
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Grabar registro"
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mon.Venta"
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
         Left            =   120
         TabIndex        =   101
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label monedav 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   100
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Und"
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
         Left            =   120
         TabIndex        =   99
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factor"
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
         Left            =   960
         TabIndex        =   98
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pvta"
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
         Left            =   1680
         TabIndex        =   97
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label40 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%Margen"
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
         Left            =   2640
         TabIndex        =   96
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Min"
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
         Left            =   3600
         TabIndex        =   95
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label42 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Max"
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
         Left            =   4080
         TabIndex        =   94
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label43 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pvta"
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
         Left            =   4560
         TabIndex        =   93
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label44 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%Margen"
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
         Left            =   5640
         TabIndex        =   92
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label45 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
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
         Left            =   3600
         TabIndex        =   91
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label46 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaFinal"
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
         Left            =   3600
         TabIndex        =   90
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label47 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
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
         Left            =   3600
         TabIndex        =   89
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label48 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaFinal"
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
         Left            =   3600
         TabIndex        =   88
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label49 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dscto Autom."
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
         Left            =   3600
         TabIndex        =   87
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label54 
         BackColor       =   &H00FFFF00&
         Caption         =   "Oferta.precio=0 acepta"
         Height          =   255
         Left            =   3600
         TabIndex        =   86
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ccosto"
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
         Left            =   3600
         TabIndex        =   85
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label costou 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label factor 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label unidad 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label monedac 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mon.Costo"
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
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "tprecios.frx":4848
      Height          =   4575
      Left            =   120
      OleObjectBlob   =   "tprecios.frx":485C
      TabIndex        =   0
      Top             =   960
      Width           =   11775
   End
   Begin VB.Label descripcio 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   600
      Width           =   7215
   End
   Begin VB.Label producto 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Menu dkl8823 
      Caption         =   "&Menu"
      Begin VB.Menu dk7823 
         Caption         =   "&1.Crear data Local"
      End
      Begin VB.Menu d782323 
         Caption         =   "&2.Generacion Completa"
      End
      Begin VB.Menu dkwerwwe 
         Caption         =   "&Migracion"
      End
   End
   Begin VB.Menu do2323 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tprecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DBGrid1_Click()

End Sub

Private Sub cmdDelete_Click()
Frame1.Visible = False
End Sub

Private Sub cmdSave_Click()
Dim found As Integer
found = grabando()
If found = 1 Then
   do2323_Click
   Exit Sub
End If
MsgBox "Error al Grabar ", 48, "Aviso"


End Sub
Function grabando()
On Error GoTo cmd2_err
Data2.Recordset.Edit
Data2.Recordset.Fields("ccosto") = seccion
Data2.Recordset.Fields("unidad1") = unidad1
Data2.Recordset.Fields("unidad2") = unidad2
Data2.Recordset.Fields("unidad3") = unidad3
Data2.Recordset.Fields("unidad4") = unidad4
Data2.Recordset.Fields("unidad5") = unidad5
Data2.Recordset.Fields("unidad6") = unidad6
Data2.Recordset.Fields("unidad7") = unidad7
Data2.Recordset.Fields("unidad8") = unidad8
Data2.Recordset.Fields("unidad9") = unidad9
Data2.Recordset.Fields("unidad10") = unidad10

Data2.Recordset.Fields("factor1") = Val(factor1)
Data2.Recordset.Fields("factor2") = Val(factor2)
Data2.Recordset.Fields("factor3") = Val(factor3)
Data2.Recordset.Fields("factor4") = Val(factor4)
Data2.Recordset.Fields("factor5") = Val(factor5)
Data2.Recordset.Fields("factor6") = Val(factor6)
Data2.Recordset.Fields("factor7") = Val(factor7)
Data2.Recordset.Fields("factor8") = Val(factor8)
Data2.Recordset.Fields("factor9") = Val(factor9)
Data2.Recordset.Fields("factor10") = Val(factor10)

Data2.Recordset.Fields("pventa1") = Val(pventa1)
Data2.Recordset.Fields("pventa2") = Val(pventa2)
Data2.Recordset.Fields("pventa3") = Val(pventa3)
Data2.Recordset.Fields("pventa4") = Val(pventa4)
Data2.Recordset.Fields("pventa5") = Val(pventa5)
Data2.Recordset.Fields("pventa6") = Val(pventa6)
Data2.Recordset.Fields("pventa7") = Val(pventa7)
Data2.Recordset.Fields("pventa8") = Val(pventa8)
Data2.Recordset.Fields("pventa9") = Val(pventa9)
Data2.Recordset.Fields("pventa10") = Val(pventa10)

Data2.Recordset.Fields("margen1") = Val(margen1)
Data2.Recordset.Fields("margen2") = Val(margen2)
Data2.Recordset.Fields("margen3") = Val(margen3)
Data2.Recordset.Fields("margen4") = Val(margen4)
Data2.Recordset.Fields("margen5") = Val(margen5)
Data2.Recordset.Fields("margen6") = Val(margen6)
Data2.Recordset.Fields("margen7") = Val(margen7)
Data2.Recordset.Fields("margen8") = Val(margen8)
Data2.Recordset.Fields("margen9") = Val(margen9)
Data2.Recordset.Fields("margen10") = Val(margen10)

Data2.Recordset.Fields("minimo11") = Val(minimo11)
Data2.Recordset.Fields("minimo12") = Val(minimo12)
Data2.Recordset.Fields("minimo13") = Val(minimo13)
Data2.Recordset.Fields("minimo14") = Val(minimo14)
Data2.Recordset.Fields("minimo15") = Val(minimo15)

Data2.Recordset.Fields("maximo11") = Val(maximo11)
Data2.Recordset.Fields("maximo12") = Val(maximo12)
Data2.Recordset.Fields("maximo13") = Val(maximo13)
Data2.Recordset.Fields("maximo14") = Val(maximo14)
Data2.Recordset.Fields("maximo15") = Val(maximo15)

Data2.Recordset.Fields("pventa11") = Val(pventa11)
Data2.Recordset.Fields("pventa12") = Val(pventa12)
Data2.Recordset.Fields("pventa13") = Val(pventa13)
Data2.Recordset.Fields("pventa14") = Val(pventa14)
Data2.Recordset.Fields("pventa15") = Val(pventa15)

Data2.Recordset.Fields("margen11") = Val(margen11)
Data2.Recordset.Fields("margen12") = Val(margen12)
Data2.Recordset.Fields("margen13") = Val(margen13)
Data2.Recordset.Fields("margen14") = Val(margen14)
Data2.Recordset.Fields("margen15") = Val(margen15)
Data2.Recordset.Update
grabando = 1
Exit Function
cmd2_err:
Exit Function

End Function

Private Sub Command1_Click()
Dim found As Integer
found = copiar_00("" & producto, "01")
found = copiar_00("" & producto, "02")
found = copiar_00("" & producto, "03")
found = copiar_00("" & producto, "04")
do2323_Click
End Sub

Private Sub Command2_Click()
do2323_Click
End Sub

Private Sub d782323_Click()
Dim found As Integer
If MsgBox("Se reeemplazara todos los locales", 1, "Aviso") <> 1 Then Exit Sub
found = generacion_completa()
End Sub

Private Sub DBGrid2_Click()
Dim found As Integer
If Frame2.Visible = True Then Exit Sub
found = pone_registro()
If found = 1 Then
   Frame1.Visible = True
   unidad1.SetFocus
   Exit Sub
End If
End Sub

Private Sub dk7823_Click()
If Frame2.Visible = True Then Exit Sub
Frame2.Visible = True
Combo1.SetFocus

End Sub

Private Sub dkwerwwe_Click()
Dim found As Integer
If MsgBox("Desea Migrar informacion", 1, "Aviso") <> 1 Then Exit Sub
found = actualiza_datos()
End Sub

Private Sub do2323_Click()
If Frame2.Visible = True Then
   Frame2.Visible = False
   Exit Sub
   Exit Sub
End If
If Frame1.Visible = True Then
   Frame1.Visible = False
   DBGrid2.SetFocus
   Exit Sub
End If
tprecios.Hide
Unload tprecios
End Sub

Private Sub Form_Activate()
sql_detalle
End Sub

Sub sql_detalle()
Dim buf As String
On Error GoTo cmd34_err
buf = "select * from precios where producto='" & producto & "'"
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = buf
               Data2.Refresh
               DBGrid2.Refresh
Exit Sub
cmd34_err:
MsgBox "Error en select " & error$, 48, "Aviso"
Exit Sub
End Sub
Function pone_registro()
On Error GoTo cmd12_err
ccosto = "" & Data2.Recordset.Fields("ccosto")
unidad1 = "" & Data2.Recordset.Fields("unidad1")
unidad2 = "" & Data2.Recordset.Fields("unidad2")
unidad3 = "" & Data2.Recordset.Fields("unidad3")
unidad4 = "" & Data2.Recordset.Fields("unidad4")
unidad5 = "" & Data2.Recordset.Fields("unidad5")
unidad6 = "" & Data2.Recordset.Fields("unidad6")
unidad7 = "" & Data2.Recordset.Fields("unidad7")
unidad8 = "" & Data2.Recordset.Fields("unidad8")
unidad9 = "" & Data2.Recordset.Fields("unidad9")
unidad10 = "" & Data2.Recordset.Fields("unidad10")
factor1 = "" & Data2.Recordset.Fields("factor1")
factor2 = "" & Data2.Recordset.Fields("factor2")
factor3 = "" & Data2.Recordset.Fields("factor3")
factor4 = "" & Data2.Recordset.Fields("factor4")
factor5 = "" & Data2.Recordset.Fields("factor5")
factor6 = "" & Data2.Recordset.Fields("factor6")
factor7 = "" & Data2.Recordset.Fields("factor7")
factor8 = "" & Data2.Recordset.Fields("factor8")
factor9 = "" & Data2.Recordset.Fields("factor9")
factor10 = "" & Data2.Recordset.Fields("factor10")
pventa1 = "" & Data2.Recordset.Fields("pventa1")
pventa2 = "" & Data2.Recordset.Fields("pventa2")
pventa3 = "" & Data2.Recordset.Fields("pventa3")
pventa4 = "" & Data2.Recordset.Fields("pventa4")
pventa5 = "" & Data2.Recordset.Fields("pventa5")
pventa6 = "" & Data2.Recordset.Fields("pventa6")
pventa7 = "" & Data2.Recordset.Fields("pventa7")
pventa8 = "" & Data2.Recordset.Fields("pventa8")
pventa9 = "" & Data2.Recordset.Fields("pventa9")
pventa10 = "" & Data2.Recordset.Fields("pventa10")
margen1 = "" & Data2.Recordset.Fields("margen1")
margen2 = "" & Data2.Recordset.Fields("margen2")
margen3 = "" & Data2.Recordset.Fields("margen3")
margen4 = "" & Data2.Recordset.Fields("margen4")
margen5 = "" & Data2.Recordset.Fields("margen5")
margen6 = "" & Data2.Recordset.Fields("margen6")
margen7 = "" & Data2.Recordset.Fields("margen7")
margen8 = "" & Data2.Recordset.Fields("margen8")
margen9 = "" & Data2.Recordset.Fields("margen9")
margen10 = "" & Data2.Recordset.Fields("margen10")
minimo11 = "" & Data2.Recordset.Fields("minimo11")
minimo12 = "" & Data2.Recordset.Fields("minimo12")
minimo13 = "" & Data2.Recordset.Fields("minimo13")
minimo14 = "" & Data2.Recordset.Fields("minimo14")
minimo15 = "" & Data2.Recordset.Fields("minimo15")
maximo11 = "" & Data2.Recordset.Fields("maximo11")
maximo12 = "" & Data2.Recordset.Fields("maximo12")
maximo13 = "" & Data2.Recordset.Fields("maximo13")
maximo14 = "" & Data2.Recordset.Fields("maximo14")
maximo15 = "" & Data2.Recordset.Fields("maximo15")
pventa11 = "" & Data2.Recordset.Fields("pventa11")
pventa12 = "" & Data2.Recordset.Fields("pventa12")
pventa13 = "" & Data2.Recordset.Fields("pventa13")
pventa14 = "" & Data2.Recordset.Fields("pventa14")
pventa15 = "" & Data2.Recordset.Fields("pventa15")
margen11 = "" & Data2.Recordset.Fields("margen11")
margen12 = "" & Data2.Recordset.Fields("margen12")
margen13 = "" & Data2.Recordset.Fields("margen13")
margen14 = "" & Data2.Recordset.Fields("margen14")
margen15 = "" & Data2.Recordset.Fields("margen15")
fechai11 = "" & Data2.Recordset.Fields("fechai11")
fechaf11 = "" & Data2.Recordset.Fields("fechaf11")
fechaid = "" & Data2.Recordset.Fields("fechaid")
fechafd = "" & Data2.Recordset.Fields("fechafd")
dscto = "" & Data2.Recordset.Fields("dscto")
calcula_margenes
pone_registro = 1
Exit Function
cmd12_err:
Exit Function
End Function
Sub calcula_margenes()
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double
If Val(factor) <= 0 Then
   factor = "1"
End If
If Val(factor1) <= 0 Then
   factor1 = "1"
End If

       If Val(costou) > 0 And Val(pventa1) > 0 Then  'calculando margenes
          sdx = (Val(costou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa1) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen1 = Format(sdx2, "0.00")
          GoTo siguiente1
       End If
       If Val(margen1) > 0 And Val(pventa1) <= 0 And Val(costou) > 0 Then
          sdx = Val(costou) + Val(costou) * Val(margen1) / 100
          pventa1 = Format(sdx, "0.00")
          GoTo siguiente1
       End If
       If Val(costou) <= 0 And Val(pventa1) > 0 And Val(margen1) > 0 Then
          sdx = Val(pventa1) / (1 + (Val(margen1) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente1
       End If
       
siguiente1:
       If Val(costou) > 0 And Val(pventa2) > 0 Then  'calculando margenes
          sdx = (Val(costou) / Val(factor))
          sdx = sdx * Val(factor2)
          sdx1 = Val(pventa2) '/ Val(factor1)
          sdx2 = (Val(sdx1) - sdx) * 100 / sdx
          margen2 = Format(sdx2, "0.00")
          GoTo siguiente2
       End If
       If Val(margen2) > 0 And Val(pventa2) <= 0 And Val(costou) > 0 Then
          sdx = Val(costou) + Val(costou) * Val(margen2) / 100
          pventa2 = Format(sdx, "0.00")
          GoTo siguiente2
       End If
       If Val(costou) <= 0 And Val(pventa2) > 0 And Val(margen2) > 0 Then
          sdx = Val(pventa2) / (1 + (Val(margen2) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente2
       End If
siguiente2:
       If Val(costou) > 0 And Val(pventa3) > 0 Then  'calculando margenes
          sdx = (Val(costou) / Val(factor))
          sdx = sdx * Val(factor3)
          sdx1 = Val(pventa3) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen3 = Format(sdx2, "0.00")
          GoTo siguiente3
       End If
       If Val(margen3) > 0 And Val(pventa3) <= 0 And Val(costou) > 0 Then
          sdx = Val(costou) + Val(costou) * Val(margen3) / 100
          pventa3 = Format(sdx, "0.00")
          GoTo siguiente3
       End If
       If Val(costou) <= 0 And Val(pventa3) > 0 And Val(margen3) > 0 Then
          sdx = Val(pventa3) / (1 + (Val(margen3) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente2
       End If
siguiente3:
If Val(costou) > 0 And Val(pventa4) > 0 Then  'calculando margenes
          sdx = (Val(costou) / Val(factor))
          sdx = sdx * Val(factor4)
          sdx1 = Val(pventa4) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen4 = Format(sdx2, "0.00")
          GoTo siguiente4
       End If
       If Val(margen4) > 0 And Val(pventa4) <= 0 And Val(costou) > 0 Then
          sdx = Val(costou) + Val(costou) * Val(margen4) / 100
          pventa4 = Format(sdx, "0.00")
          GoTo siguiente4
       End If
       If Val(costou) <= 0 And Val(pventa4) > 0 And Val(margen4) > 0 Then
          sdx = Val(pventa4) / (1 + (Val(margen4) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente4
       End If
siguiente4:
If Val(costou) > 0 And Val(pventa5) > 0 Then  'calculando margenes
          sdx = (Val(costou) / Val(factor))
          sdx = sdx * Val(factor5)
          sdx1 = Val(pventa5) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen5 = Format(sdx2, "0.00")
          GoTo siguiente5
       End If
       If Val(margen5) > 0 And Val(pventa5) <= 0 And Val(costou) > 0 Then
          sdx = Val(costou) + Val(costou) * Val(margen5) / 100
          pventa5 = Format(sdx, "0.00")
          GoTo siguiente5
       End If
       If Val(costou) <= 0 And Val(pventa5) > 0 And Val(margen5) > 0 Then
          sdx = Val(pventa5) / (1 + (Val(margen5) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente5
       End If
siguiente5:
If Val(costou) > 0 And Val(pventa6) > 0 Then  'calculando margenes
          sdx = (Val(costou) / Val(factor))
          sdx = sdx * Val(factor6)
          sdx1 = Val(pventa6) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen6 = Format(sdx2, "0.00")
          GoTo siguiente6
       End If
       If Val(margen6) > 0 And Val(pventa6) <= 0 And Val(costou) > 0 Then
          sdx = Val(costou) + Val(costou) * Val(margen6) / 100
          pventa6 = Format(sdx, "0.00")
          GoTo siguiente6
       End If
       If Val(costou) <= 0 And Val(pventa6) > 0 And Val(margen6) > 0 Then
          sdx = Val(pventa6) / (1 + (Val(margen6) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente6
       End If
siguiente6:
If Val(costou) > 0 And Val(pventa7) > 0 Then  'calculando margenes
          sdx = (Val(costou) / Val(factor))
          sdx = sdx * Val(factor7)
          sdx1 = Val(pventa7) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen7 = Format(sdx2, "0.00")
          GoTo siguiente7
       End If
       If Val(margen7) > 0 And Val(pventa7) <= 0 And Val(costou) > 0 Then
          sdx = Val(costou) + Val(costou) * Val(margen7) / 100
          pventa7 = Format(sdx, "0.00")
          GoTo siguiente7
       End If
       If Val(costou) <= 0 And Val(pventa7) > 0 And Val(margen7) > 0 Then
          sdx = Val(pventa7) / (1 + (Val(margen7) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente7
       End If
siguiente7:
If Val(costou) > 0 And Val(pventa8) > 0 Then  'calculando margenes
          sdx = (Val(costou) / Val(factor))
          sdx = sdx * Val(factor8)
          sdx1 = Val(pventa8) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen8 = Format(sdx2, "0.00")
          GoTo siguiente8
       End If
       If Val(margen8) > 0 And Val(pventa8) <= 0 And Val(costou) > 0 Then
          sdx = Val(costou) + Val(costou) * Val(margen8) / 100
          pventa8 = Format(sdx, "0.00")
          GoTo siguiente8
       End If
       If Val(costou) <= 0 And Val(pventa8) > 0 And Val(margen8) > 0 Then
          sdx = Val(pventa8) / (1 + (Val(margen8) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente8
       End If
siguiente8:
If Val(costou) > 0 And Val(pventa9) > 0 Then  'calculando margenes
          sdx = (Val(costou) / Val(factor))
          sdx = sdx * Val(factor9)
          sdx1 = Val(pventa9) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen9 = Format(sdx2, "0.00")
          GoTo siguiente9
       End If
       If Val(margen9) > 0 And Val(pventa9) <= 0 And Val(costou) > 0 Then
          sdx = Val(costou) + Val(costou) * Val(margen9) / 100
          pventa9 = Format(sdx, "0.00")
          GoTo siguiente9
       End If
       If Val(costou) <= 0 And Val(pventa9) > 0 And Val(margen9) > 0 Then
          sdx = Val(pventa9) / (1 + (Val(margen9) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente9
       End If
siguiente9:
If Val(costou) > 0 And Val(pventa10) > 0 Then  'calculando margenes
          sdx = (Val(costou) / Val(factor))
          sdx = sdx * Val(factor10)
          sdx1 = Val(pventa10) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen10 = Format(sdx2, "0.00")
          GoTo siguiente10
       End If
       If Val(margen10) > 0 And Val(pventa10) <= 0 And Val(costou) > 0 Then
          sdx = Val(costou) + Val(costou) * Val(margen10) / 100
          pventa2 = Format(sdx, "0.00")
          GoTo siguiente10
       End If
       If Val(costou) <= 0 And Val(pventa10) > 0 And Val(margen10) > 0 Then
          sdx = Val(pventa10) / (1 + (Val(margen10) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente10
       End If
siguiente10:
If Val(costou) > 0 And Val(pventa11) > 0 Then  'calculando margenes
          sdx = (Val(costou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa11) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen11 = Format(sdx2, "0.00")
          GoTo siguiente11
       End If
       If Val(margen11) > 0 And Val(pventa11) <= 0 And Val(costou) > 0 Then
          sdx = Val(costou) + Val(costou) * Val(margen11) / 100
          pventa11 = Format(sdx, "0.00")
          GoTo siguiente11
       End If
       If Val(costou) <= 0 And Val(pventa11) > 0 And Val(margen11) > 0 Then
          sdx = Val(pventa11) / (1 + (Val(margen11) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente11
       End If
siguiente11:
If Val(costou) > 0 And Val(pventa12) > 0 Then  'calculando margenes
          sdx = (Val(costou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa12) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen12 = Format(sdx2, "0.00")
          GoTo siguiente12
       End If
       If Val(margen12) > 0 And Val(pventa12) <= 0 And Val(costou) > 0 Then
          sdx = Val(costou) + Val(costou) * Val(margen12) / 100
          pventa12 = Format(sdx, "0.00")
          GoTo siguiente12
       End If
       If Val(costou) <= 0 And Val(pventa12) > 0 And Val(margen12) > 0 Then
          sdx = Val(pventa12) / (1 + (Val(margen12) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente12
       End If
siguiente12:
If Val(costou) > 0 And Val(pventa13) > 0 Then  'calculando margenes
          sdx = (Val(costou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa13) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen13 = Format(sdx2, "0.00")
          GoTo siguiente13
       End If
       If Val(margen13) > 0 And Val(pventa13) <= 0 And Val(costou) > 0 Then
          sdx = Val(costou) + Val(costou) * Val(margen13) / 100
          pventa13 = Format(sdx, "0.00")
          GoTo siguiente13
       End If
       If Val(costou) <= 0 And Val(pventa13) > 0 And Val(margen13) > 0 Then
          sdx = Val(pventa13) / (1 + (Val(margen13) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente13
       End If
siguiente13:
If Val(costou) > 0 And Val(pventa14) > 0 Then  'calculando margenes
          sdx = (Val(costou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa14) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen14 = Format(sdx2, "0.00")
          GoTo siguiente14
       End If
       If Val(margen14) > 0 And Val(pventa14) <= 0 And Val(costou) > 0 Then
          sdx = Val(costou) + Val(costou) * Val(margen14) / 100
          pventa14 = Format(sdx, "0.00")
          GoTo siguiente14
       End If
       If Val(costou) <= 0 And Val(pventa14) > 0 And Val(margen14) > 0 Then
          sdx = Val(pventa14) / (1 + (Val(margen14) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente14
       End If
siguiente14:
If Val(costou) > 0 And Val(pventa15) > 0 Then  'calculando margenes
          sdx = (Val(costou) / Val(factor))
          sdx = sdx * Val(factor1)
          sdx1 = Val(pventa15) '/ Val(factor1)
          sdx2 = (sdx1 - sdx) * 100 / sdx
          margen15 = Format(sdx2, "0.00")
          GoTo siguiente15
       End If
       If Val(margen15) > 0 And Val(pventa15) <= 0 And Val(costou) > 0 Then
          sdx = Val(costou) + Val(costou) * Val(margen15) / 100
          pventa15 = Format(sdx, "0.00")
          GoTo siguiente10
       End If
       If Val(costou) <= 0 And Val(pventa15) > 0 And Val(margen15) > 0 Then
          sdx = Val(pventa15) / (1 + (Val(margen15) / 100))
          costou = Format(sdx, "0.0000")
          GoTo siguiente15
       End If
siguiente15:
End Sub



Private Sub Form_Load()
Combo1.Clear
Combo1.AddItem "00"
Combo1.AddItem "01"
Combo1.AddItem "02"
Combo1.AddItem "03"
Combo1.AddItem "04"
Combo1.ListIndex = 0

Combo2.Clear
Combo2.AddItem "01"
Combo2.AddItem "02"
Combo2.AddItem "03"
Combo2.AddItem "04"
Combo2.ListIndex = 0
End Sub
Function copiar_00(buf As String, xlocal As String)

Dim mytablex As Table
Dim mytabley As Table

Set mytabley = mydbxglo.OpenTable("precios")
mytabley.Index = "tprecios"
Set mytablex = mydbxglo.OpenTable("producto")
mytablex.Index = "producto"
mytablex.Seek "=", buf
If Not mytablex.NoMatch Then
   mytabley.Seek "=", xlocal, buf
   If Not mytabley.NoMatch Then
      mytabley.Edit
      grabandoy mytablex, mytabley, "" & xlocal
      mytabley.Update
   End If
   If mytabley.NoMatch Then
      mytabley.AddNew
      grabandoy mytablex, mytabley, "" & xlocal
      mytabley.Update
   End If
End If
mytablex.Close
mytabley.Close
 
End Function
Sub grabandoy(mytablex As Table, mytabley As Table, xlocal As String)
mytabley.Fields("producto") = "" & mytablex.Fields("producto")
mytabley.Fields("local") = xlocal
mytabley.Fields("ccosto") = "" & mytablex.Fields("ccosto")
mytabley.Fields("unidad1") = "" & mytablex.Fields("unidad1")
mytabley.Fields("unidad2") = "" & mytablex.Fields("unidad2")
mytabley.Fields("unidad3") = "" & mytablex.Fields("unidad3")
mytabley.Fields("unidad4") = "" & mytablex.Fields("unidad4")
mytabley.Fields("unidad5") = "" & mytablex.Fields("unidad5")
mytabley.Fields("unidad6") = "" & mytablex.Fields("unidad6")
mytabley.Fields("unidad7") = "" & mytablex.Fields("unidad7")
mytabley.Fields("unidad8") = "" & mytablex.Fields("unidad8")
mytabley.Fields("unidad9") = "" & mytablex.Fields("unidad9")
mytabley.Fields("unidad10") = "" & mytablex.Fields("unidad10")
mytabley.Fields("factor1") = Val("" & mytablex.Fields("factor1"))
mytabley.Fields("factor2") = Val("" & mytablex.Fields("factor2"))
mytabley.Fields("factor3") = Val("" & mytablex.Fields("factor3"))
mytabley.Fields("factor4") = Val("" & mytablex.Fields("factor4"))
mytabley.Fields("factor5") = Val("" & mytablex.Fields("factor5"))
mytabley.Fields("factor6") = Val("" & mytablex.Fields("factor6"))
mytabley.Fields("factor7") = Val("" & mytablex.Fields("factor7"))
mytabley.Fields("factor8") = Val("" & mytablex.Fields("factor8"))
mytabley.Fields("factor9") = Val("" & mytablex.Fields("factor9"))
mytabley.Fields("factor10") = Val("" & mytablex.Fields("factor10"))
mytabley.Fields("pventa1") = Val("" & mytablex.Fields("pventa1"))
mytabley.Fields("pventa2") = Val("" & mytablex.Fields("pventa2"))
mytabley.Fields("pventa3") = Val("" & mytablex.Fields("pventa3"))
mytabley.Fields("pventa4") = Val("" & mytablex.Fields("pventa4"))
mytabley.Fields("pventa5") = Val("" & mytablex.Fields("pventa5"))
mytabley.Fields("pventa6") = Val("" & mytablex.Fields("pventa6"))
mytabley.Fields("pventa7") = Val("" & mytablex.Fields("pventa7"))
mytabley.Fields("pventa8") = Val("" & mytablex.Fields("pventa8"))
mytabley.Fields("pventa9") = Val("" & mytablex.Fields("pventa9"))
mytabley.Fields("pventa10") = Val("" & mytablex.Fields("pventa10"))
mytabley.Fields("margen1") = Val("" & mytablex.Fields("margen1"))
mytabley.Fields("margen2") = Val("" & mytablex.Fields("margen2"))
mytabley.Fields("margen3") = Val("" & mytablex.Fields("margen3"))
mytabley.Fields("margen4") = Val("" & mytablex.Fields("margen4"))
mytabley.Fields("margen5") = Val("" & mytablex.Fields("margen5"))
mytabley.Fields("margen6") = Val("" & mytablex.Fields("margen6"))
mytabley.Fields("margen7") = Val("" & mytablex.Fields("margen7"))
mytabley.Fields("margen8") = Val("" & mytablex.Fields("margen8"))
mytabley.Fields("margen9") = Val("" & mytablex.Fields("margen9"))
mytabley.Fields("margen10") = Val("" & mytablex.Fields("margen10"))
mytabley.Fields("minimo11") = Val("" & mytablex.Fields("minimo11"))
mytabley.Fields("minimo12") = Val("" & mytablex.Fields("minimo12"))
mytabley.Fields("minimo13") = Val("" & mytablex.Fields("minimo13"))
mytabley.Fields("minimo14") = Val("" & mytablex.Fields("minimo14"))
mytabley.Fields("minimo15") = Val("" & mytablex.Fields("minimo15"))
mytabley.Fields("maximo11") = Val("" & mytablex.Fields("maximo11"))
mytabley.Fields("maximo12") = Val("" & mytablex.Fields("maximo12"))
mytabley.Fields("maximo13") = Val("" & mytablex.Fields("maximo13"))
mytabley.Fields("maximo14") = Val("" & mytablex.Fields("maximo14"))
mytabley.Fields("maximo15") = Val("" & mytablex.Fields("maximo15"))
mytabley.Fields("pventa11") = Val("" & mytablex.Fields("pventa11"))
mytabley.Fields("pventa12") = Val("" & mytablex.Fields("pventa12"))
mytabley.Fields("pventa13") = Val("" & mytablex.Fields("pventa13"))
mytabley.Fields("pventa14") = Val("" & mytablex.Fields("pventa14"))
mytabley.Fields("pventa15") = Val("" & mytablex.Fields("pventa15"))
mytabley.Fields("margen11") = Val("" & mytablex.Fields("margen11"))
mytabley.Fields("margen12") = Val("" & mytablex.Fields("margen12"))
mytabley.Fields("margen13") = Val("" & mytablex.Fields("margen13"))
mytabley.Fields("margen14") = Val("" & mytablex.Fields("margen14"))
mytabley.Fields("margen15") = Val("" & mytablex.Fields("margen15"))
'mytabley.Fields("fechai11") = "" & mytablex.Fields("fechai11")
'mytabley.Fields("fechaf11") = "" & mytablex.Fields("fechaf11")
'mytabley.Fields("fechaid") = "" & mytablex.Fields("fechaid")
'mytabley.Fields("fechafd") = "" & mytablex.Fields("fechafd")
mytabley.Fields("dscto") = Val("" & mytablex.Fields("dscto"))
End Sub
Function generacion_completa()

Dim mytablex As Table
Dim mytabley As Table
Dim i As Integer

Set mytabley = mydbxglo.OpenTable("precios")
mytabley.Index = "tprecios"
Set mytablex = mydbxglo.OpenTable("producto")
Do
  If mytablex.EOF Then Exit Do
   For i = 1 To 4
   mytabley.Seek "=", Format(i, "00"), "" & mytablex.Fields("producto")
   If Not mytabley.NoMatch Then
      mytabley.Edit
      grabandoy mytablex, mytabley, Format(i, "00")
      mytabley.Update
   End If
   If mytabley.NoMatch Then
      mytabley.AddNew
      grabandoy mytablex, mytabley, Format(i, "00")
      mytabley.Update
   End If
   Next i
   mytablex.MoveNext
Loop
mytablex.Close
mytabley.Close
 
MsgBox "Proceso terminado ", 48, "Aviso"
End Function
Function actualiza_datos()

Dim mytablex As Table
Dim mytabley As Table
Dim mytablez As Table
Dim i As Integer

Set mytablez = mydbxglo.OpenTable("producto")
mytablez.Index = "producto"
Set mytabley = mydbxglo.OpenTable("precios")
mytabley.Index = "tprecios"
Set mytablex = mydbxglo.OpenTable("fabiolo")
Do
  If mytablex.EOF Then Exit Do
   For i = 1 To 4
                mytablez.Seek "=", "" & mytablex.Fields("codigo")
                If Not mytablez.NoMatch Then
                   mytablez.Edit
                   mytablez.Fields("ccosto") = "" & mytablex.Fields("tortu")
                   mytablez.Update
                End If
   mytabley.Seek "=", Format(i, "00"), "" & mytablex.Fields("codigo")
   If Not mytabley.NoMatch Then
      mytabley.Edit
      mytabley.Fields("producto") = "" & mytablex.Fields("codigo")
      mytabley.Fields("local") = Format(i, "00")
      Select Case i
             Case 1
             mytabley.Fields("ccosto") = "" & mytablex.Fields("tortu")
             Case 2
             mytabley.Fields("ccosto") = "" & mytablex.Fields("dos_her")
             Case 3
             mytabley.Fields("ccosto") = "" & mytablex.Fields("surco")
             Case 4
             mytabley.Fields("ccosto") = "" & mytablex.Fields("molina")
      End Select
      mytabley.Update
   End If
   If mytabley.NoMatch Then
      mytabley.AddNew
      mytabley.Fields("producto") = "" & mytablex.Fields("codigo")
      mytabley.Fields("local") = Format(i, "00")
      Select Case i
             Case 1
             mytabley.Fields("ccosto") = "" & mytablex.Fields("tortu")
             Case 2
             mytabley.Fields("ccosto") = "" & mytablex.Fields("dos_her")
             Case 3
             mytabley.Fields("ccosto") = "" & mytablex.Fields("surco")
             Case 4
             mytabley.Fields("ccosto") = "" & mytablex.Fields("molina")
      End Select
      mytabley.Update
      
   End If
   Next i
   mytablex.MoveNext
Loop
mytablez.Close
mytablex.Close
mytabley.Close
 
MsgBox "Proceso terminado ", 48, "Aviso"

End Function

