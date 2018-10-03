VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tingtalla 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso de Tallas"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "STOCK ACTUAL"
      Height          =   2490
      Left            =   0
      TabIndex        =   40
      Top             =   3720
      Width           =   10575
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   1740
         Left            =   195
         TabIndex        =   41
         Top             =   360
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   3069
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   15
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
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "INGRESO DE TALLAS"
      Height          =   3600
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin VB.CommandButton Command1 
         Caption         =   "SALIR"
         Height          =   600
         Left            =   9195
         TabIndex        =   44
         Top             =   135
         Width           =   1290
      End
      Begin VB.TextBox t1 
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
         Left            =   1905
         MaxLength       =   2
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1470
         Width           =   735
      End
      Begin VB.TextBox t2 
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
         Left            =   1905
         MaxLength       =   2
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1830
         Width           =   735
      End
      Begin VB.TextBox t3 
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
         Left            =   1905
         MaxLength       =   2
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2190
         Width           =   735
      End
      Begin VB.TextBox t4 
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
         Left            =   1905
         MaxLength       =   2
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2550
         Width           =   735
      End
      Begin VB.TextBox t5 
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
         Left            =   3345
         MaxLength       =   2
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1470
         Width           =   735
      End
      Begin VB.TextBox t6 
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
         Left            =   3345
         MaxLength       =   2
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1830
         Width           =   735
      End
      Begin VB.TextBox t7 
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
         Left            =   3345
         MaxLength       =   2
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2190
         Width           =   735
      End
      Begin VB.TextBox t8 
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
         Left            =   3345
         MaxLength       =   2
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2550
         Width           =   735
      End
      Begin VB.TextBox t9 
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
         Left            =   4785
         MaxLength       =   2
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1470
         Width           =   735
      End
      Begin VB.TextBox t10 
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
         Left            =   4785
         MaxLength       =   2
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1830
         Width           =   735
      End
      Begin VB.TextBox t11 
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
         Left            =   4785
         MaxLength       =   2
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2190
         Width           =   735
      End
      Begin VB.TextBox t12 
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
         Left            =   4785
         MaxLength       =   2
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2550
         Width           =   735
      End
      Begin VB.TextBox t13 
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
         Left            =   6225
         MaxLength       =   2
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1470
         Width           =   735
      End
      Begin VB.TextBox t14 
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
         Left            =   6225
         MaxLength       =   2
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1830
         Width           =   735
      End
      Begin VB.TextBox t15 
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
         Left            =   6225
         MaxLength       =   2
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2190
         Width           =   735
      End
      Begin VB.TextBox t16 
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
         Left            =   6225
         MaxLength       =   2
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2550
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
         Height          =   1125
         Left            =   7290
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tingtalla.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Grabar registro"
         Top             =   1470
         Width           =   1275
      End
      Begin VB.Label Descripcion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   1920
         TabIndex        =   43
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Producto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   240
         Left            =   120
         TabIndex        =   42
         Top             =   345
         Width           =   945
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linea                                       Tallas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   90
         TabIndex        =   39
         Top             =   1010
         Width           =   1095
      End
      Begin VB.Label nlinea 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   470
         Left            =   1170
         TabIndex        =   38
         Top             =   1010
         Width           =   3255
      End
      Begin VB.Label nt1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1170
         TabIndex        =   37
         Top             =   1485
         Width           =   735
      End
      Begin VB.Label nt2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1185
         TabIndex        =   36
         Top             =   1830
         Width           =   735
      End
      Begin VB.Label nt3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1185
         TabIndex        =   35
         Top             =   2190
         Width           =   735
      End
      Begin VB.Label nt4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1185
         TabIndex        =   34
         Top             =   2550
         Width           =   735
      End
      Begin VB.Label nt5 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2625
         TabIndex        =   33
         Top             =   1470
         Width           =   735
      End
      Begin VB.Label nt6 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2625
         TabIndex        =   32
         Top             =   1830
         Width           =   735
      End
      Begin VB.Label nt7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2625
         TabIndex        =   31
         Top             =   2190
         Width           =   735
      End
      Begin VB.Label nt8 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2625
         TabIndex        =   30
         Top             =   2550
         Width           =   735
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1905
         TabIndex        =   29
         Top             =   1815
         Width           =   735
      End
      Begin VB.Label Label29 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   28
         Top             =   1455
         Width           =   735
      End
      Begin VB.Label Label30 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   27
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt9 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4065
         TabIndex        =   26
         Top             =   1470
         Width           =   735
      End
      Begin VB.Label nt10 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4065
         TabIndex        =   25
         Top             =   1830
         Width           =   735
      End
      Begin VB.Label nt11 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4065
         TabIndex        =   24
         Top             =   2190
         Width           =   735
      End
      Begin VB.Label nt12 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4065
         TabIndex        =   23
         Top             =   2550
         Width           =   735
      End
      Begin VB.Label nt13 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5505
         TabIndex        =   22
         Top             =   1470
         Width           =   735
      End
      Begin VB.Label nt14 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5505
         TabIndex        =   21
         Top             =   1830
         Width           =   735
      End
      Begin VB.Label nt15 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5505
         TabIndex        =   20
         Top             =   2190
         Width           =   735
      End
      Begin VB.Label nt16 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
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
         Left            =   5490
         TabIndex        =   19
         Top             =   2550
         Width           =   735
      End
      Begin VB.Label linea 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4440
         TabIndex        =   18
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.Menu fd4541 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tingtalla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    tingtalla.Hide
    Unload tingtalla

End Sub

Private Sub Command2_Click()
    tptovta.DBGrid2.columns(18) = Val(t1)
    tptovta.DBGrid2.columns(19) = Val(t2)
    tptovta.DBGrid2.columns(20) = Val(t3)
    tptovta.DBGrid2.columns(21) = Val(t4)
    tptovta.DBGrid2.columns(22) = Val(t5)
    tptovta.DBGrid2.columns(23) = Val(t6)
    tptovta.DBGrid2.columns(24) = Val(t7)
    tptovta.DBGrid2.columns(25) = Val(t8)
    tptovta.DBGrid2.columns(26) = Val(t9)
    tptovta.DBGrid2.columns(27) = Val(t10)
    tptovta.DBGrid2.columns(28) = Val(t11)
    tptovta.DBGrid2.columns(29) = Val(t12)
    tptovta.DBGrid2.columns(30) = Val(t13)
    tptovta.DBGrid2.columns(31) = Val(t14)
    tptovta.DBGrid2.columns(32) = Val(t15)
    tptovta.DBGrid2.columns(33) = Val(t16)

    sdx = Val(t1) + Val(t2) + Val(t3) + Val(t4) + Val(t5) + Val(t6) + Val(t7) + Val(t8) + Val(t9) + Val(t10) + Val(t11) + Val(t12) + Val(t13) + Val(t14) + Val(t15) + Val(t16)

    '''27/07/2017 kenyo Testing Completo al Sistema
    'tptovta.dbgrid2.columns("cantidad") = sdx
    'sdx = Val("" & tptovta.dbgrid2.columns("cantidad")) * Val("" & tptovta.dbgrid2.columns("precio"))
    'tptovta.dbgrid2.columns("total") = sdx
    
    If sdx > 0 Then
        tptovta.DBGrid2.columns("cantidad") = sdx
        sdx = Val("" & tptovta.DBGrid2.columns("cantidad")) * Val("" & tptovta.DBGrid2.columns("precio"))
        tptovta.DBGrid2.columns("total") = sdx

    End If

    '''27/07/2017 kenyo Testing Completo al Sistema

    ''''28/09/2017 kenyo Testing Zapateria
    tptovta.hknumero.Caption = ""
    tingtalla.Hide
    Unload tingtalla

    ''''28/09/2017 kenyo Testing Zapateria
End Sub

'KENYO
Sub sql_saldo_locales(buf As String)

    Dim buf1     As String

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    On Error GoTo cmd34_err

    'mytablex.Open "SELECT * from bodega WHERE 1=2", cn, adOpenKeyset, adLockOptimistic
    'Do
    'If mytablex.EOF Then Exit Do
    'If mytabley.State = 1 Then mytabley.Close
    'mytabley.Open "SELECT * from almacen where local='" & mytablex.Fields("local") & "' and producto='" & "" & Trim(buf) & "' AND bodega='" & "" & mytablex.Fields("codigo") & "'", cn, adOpenStatic, adLockOptimistic
    'If mytabley.RecordCount = 0 Then
    '   mytabley.AddNew
    '   mytabley.Fields("producto") = buf
    '   mytabley.Fields("local") = mytablex.Fields("local")
    '   mytabley.Fields("bodega") = "" & mytablex.Fields("codigo")
    '   mytabley.Fields("minimo") = Val("" & rrproducto.Fields("minimo"))
    '   mytabley.Fields("maximo") = Val("" & rrproducto.Fields("maximo"))
    '   mytabley.Fields("saldo") = 0
    '   mytabley.Update
    'End If
    'mytabley.Close
    'mytablex.MoveNext
    'Loop
    'mytablex.Close

    buf1 = "select Almacen.saldo,t1 as '" & nt1 & "',t2 as '" & nt2 & "' ,t3 as '" & nt3 & "',t4 as '" & nt4 & "',t5 as '" & nt5 & "',t6 as '" & nt6 & "',t7 as '" & nt7 & "',t8 as '" & nt8 & "',t9 as '" & nt9 & "',t10 as '" & nt10 & "',T11 as '" & nt11 & "',T12 as '" & nt12 & "',T13 as '" & nt13 & "',T14 as '" & nt14 & "',T15 as '" & nt15 & "' , T16 as '" & nt16 & "' , Bodega.nombre,almacen.bodega as Almacen,Almacen.local from almacen,bodega where  almacen.bodega=bodega.codigo and almacen.producto='" & buf & "' order by almacen.bodega"

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open buf1, cn, adOpenKeyset, adLockOptimistic
    Set DBGrid2.DataSource = mytablex
    DBGrid2.columns(0).Width = 500 'inventario actual
   
    DBGrid2.columns(1).Width = 400
    DBGrid2.columns(2).Width = 400
    DBGrid2.columns(3).Width = 400
    DBGrid2.columns(4).Width = 400
    DBGrid2.columns(5).Width = 400
    DBGrid2.columns(6).Width = 400
    DBGrid2.columns(7).Width = 400
    DBGrid2.columns(8).Width = 400
    DBGrid2.columns(9).Width = 400
    DBGrid2.columns(10).Width = 400
    DBGrid2.columns(11).Width = 400
    DBGrid2.columns(12).Width = 400
    DBGrid2.columns(13).Width = 400
    DBGrid2.columns(14).Width = 400
    DBGrid2.columns(15).Width = 400
    DBGrid2.columns(16).Width = 400
    DBGrid2.columns(17).Width = 900
    DBGrid2.columns(18).Width = 900
    DBGrid2.columns(19).Width = 500
   
    Exit Sub
cmd34_err:
    MsgBox "Aviso en sql saldo locales " + error$, 48, "Aviso"
    Exit Sub

End Sub

'FIN KENYO

'KENYO
Private Sub Command3_Click()

End Sub

'FIN KENYO

Private Sub fd4541_Click()
    tingtalla.Hide
    Unload tingtalla

End Sub

