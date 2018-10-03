VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tcomvta 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorador de Compras/Ventas"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   14955
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
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
      Height          =   9735
      Left            =   -120
      TabIndex        =   61
      Top             =   1200
      Visible         =   0   'False
      Width           =   14895
      Begin VB.TextBox buffer 
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   63
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
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   8775
         Left            =   120
         TabIndex        =   65
         Top             =   840
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   15478
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   22
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
   Begin VB.CommandButton Command8 
      Caption         =   "Sumar"
      Height          =   615
      Left            =   11760
      TabIndex        =   60
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Generacion de Documentos"
      Height          =   2895
      Left            =   2880
      TabIndex        =   48
      Top             =   1560
      Visible         =   0   'False
      Width           =   8415
      Begin VB.TextBox gacu 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   57
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Salir"
         Height          =   615
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H000000FF&
         Caption         =   "Generar"
         Height          =   615
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox gnumero 
         Height          =   375
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   53
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox gserie 
         Height          =   375
         Left            =   1560
         MaxLength       =   11
         TabIndex        =   51
         Top             =   720
         Width           =   855
      End
      Begin VB.ComboBox gtipo 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Flag"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   58
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   54
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   52
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocumento"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   50
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clave de Acceso"
      Height          =   5055
      Left            =   3000
      TabIndex        =   41
      Top             =   1560
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tcomvta.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   4080
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tcomvta.frx":07AE
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3240
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox clave 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label Label13 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su clave para realizar esta Accion"
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
         Left            =   600
         TabIndex        =   45
         Top             =   720
         Width           =   4215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Consultas-Condiciones"
      Height          =   5175
      Left            =   4320
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   7695
      Begin VB.ComboBox servicio 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox moneda 
         Height          =   360
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox nombre 
         Height          =   375
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   14
         Text            =   "%"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tcomvta.frx":0F5C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton cmdGrabar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tcomvta.frx":170A
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.ComboBox estado 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox codigo 
         Height          =   375
         Left            =   2400
         MaxLength       =   11
         TabIndex        =   8
         Text            =   "%"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Servicio"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   67
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   1275
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   14895
      TabIndex        =   6
      Top             =   0
      Width           =   14955
      Begin VB.TextBox tipoclie 
         Height          =   375
         Left            =   10560
         MaxLength       =   1
         TabIndex        =   74
         Text            =   "%"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox serie 
         Height          =   375
         Left            =   10560
         MaxLength       =   4
         TabIndex        =   71
         Text            =   "%"
         Top             =   0
         Width           =   735
      End
      Begin VB.TextBox numero 
         Height          =   375
         Left            =   10560
         MaxLength       =   11
         TabIndex        =   70
         Text            =   "%"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox acu 
         Height          =   375
         Left            =   9480
         MaxLength       =   1
         TabIndex        =   68
         Text            =   "%"
         Top             =   360
         Width           =   375
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   8400
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   0
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "No mostrar Tipo 5"
         Height          =   255
         Left            =   7920
         TabIndex        =   39
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cajero 
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox vendedor 
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox bodegaf 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox tipo 
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox caja 
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   0
         Width           =   1695
      End
      Begin VB.ComboBox bodega 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H000000FF&
         Caption         =   "&Refrescar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   12600
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   2055
      End
      Begin VB.TextBox fechaf 
         Height          =   375
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox fechai 
         Height          =   375
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   20
         Top             =   0
         Width           =   1695
      End
      Begin VB.ComboBox local1 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "clie(CPV)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9840
         TabIndex        =   75
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9840
         TabIndex        =   73
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   9840
         TabIndex        =   72
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(C)ompra (V)enta"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7920
         TabIndex        =   69
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label tinterno 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   14520
         TabIndex        =   59
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label34 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7920
         TabIndex        =   47
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajero"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5400
         TabIndex        =   36
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2760
         TabIndex        =   34
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AlmaFin"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   33
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDoc"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5400
         TabIndex        =   30
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   5400
         TabIndex        =   28
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AlmaIni"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   26
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Final"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2760
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Inicio"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2760
         TabIndex        =   21
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   855
      End
   End
   Begin MSDataGridLib.DataGrid dbgrid2 
      Height          =   6735
      Left            =   0
      TabIndex        =   40
      Top             =   1200
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   11880
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   27
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
      ColumnCount     =   32
      BeginProperty Column00 
         DataField       =   "Yausado"
         Caption         =   "A"
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
         DataField       =   "Estado"
         Caption         =   "E"
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
         DataField       =   "Local"
         Caption         =   "Local"
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
         DataField       =   "Tipo"
         Caption         =   "Tipo"
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
      BeginProperty Column04 
         DataField       =   "Serie"
         Caption         =   "Serie"
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
         DataField       =   "Numero"
         Caption         =   "Numero"
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
         DataField       =   "fecha"
         Caption         =   "Fecha"
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
      BeginProperty Column07 
         DataField       =   "Hora"
         Caption         =   "Hora"
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
      BeginProperty Column08 
         DataField       =   "Tipoclie"
         Caption         =   "T"
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
      BeginProperty Column09 
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
      BeginProperty Column10 
         DataField       =   "Nombre"
         Caption         =   "Nombre"
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
      BeginProperty Column11 
         DataField       =   "Moneda"
         Caption         =   "M"
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
      BeginProperty Column12 
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
      BeginProperty Column13 
         DataField       =   "bodega"
         Caption         =   "BodI"
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
      BeginProperty Column14 
         DataField       =   "bodegaf"
         Caption         =   "Bodf"
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
      BeginProperty Column15 
         DataField       =   "Acuenta"
         Caption         =   "Acuenta"
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
      BeginProperty Column16 
         DataField       =   "Nro_items"
         Caption         =   "Items"
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
      BeginProperty Column17 
         DataField       =   "Usuario"
         Caption         =   "Cajero"
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
      BeginProperty Column18 
         DataField       =   "Caja"
         Caption         =   "Caja"
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
      BeginProperty Column19 
         DataField       =   "Turno"
         Caption         =   "T"
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
      BeginProperty Column20 
         DataField       =   "vendedor"
         Caption         =   "Vendedor"
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
      BeginProperty Column21 
         DataField       =   "Observa"
         Caption         =   "Observa"
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
      BeginProperty Column22 
         DataField       =   "Acu"
         Caption         =   "Acu"
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
      BeginProperty Column23 
         DataField       =   "Servicio"
         Caption         =   "S"
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
      BeginProperty Column24 
         DataField       =   "Local1"
         Caption         =   "Local1"
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
      BeginProperty Column25 
         DataField       =   "tipo1"
         Caption         =   "Tipo1"
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
      BeginProperty Column26 
         DataField       =   "Serie1"
         Caption         =   "Serie1"
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
      BeginProperty Column27 
         DataField       =   "numero1"
         Caption         =   "Numero1"
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
      BeginProperty Column28 
         DataField       =   "retipo1"
         Caption         =   "retipo1"
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
      BeginProperty Column29 
         DataField       =   "renumero3"
         Caption         =   "renumero3"
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
      BeginProperty Column30 
         DataField       =   "renumero1"
         Caption         =   "renumero1"
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
      BeginProperty Column31 
         DataField       =   "renumero2"
         Caption         =   "renumero2"
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
         MarqueeStyle    =   2
         BeginProperty Column00 
            ColumnWidth     =   180.283
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   209.764
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   360
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1094.74
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   645.165
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   180.283
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   2009.764
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   225.071
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   734.74
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   390.047
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   195.024
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column21 
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   420.095
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   345.26
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   585.071
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   569.764
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column28 
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   929.764
         EndProperty
         BeginProperty Column31 
            ColumnWidth     =   1080
         EndProperty
      EndProperty
   End
   Begin VB.Label difigv 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   11760
      TabIndex        =   89
      Top             =   8880
      Width           =   1575
   End
   Begin VB.Label totalneto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      TabIndex        =   88
      Top             =   9240
      Width           =   1575
   End
   Begin VB.Label totalvt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   9960
      TabIndex        =   87
      Top             =   9240
      Width           =   1695
   End
   Begin VB.Label totalco 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4800
      TabIndex        =   86
      Top             =   9240
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VENTAS"
      Height          =   375
      Left            =   6600
      TabIndex        =   85
      Top             =   8160
      Width           =   5055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COMPRAS"
      Height          =   375
      Left            =   1440
      TabIndex        =   84
      Top             =   8160
      Width           =   5055
   End
   Begin VB.Label Label33 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "S/."
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   83
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label comtots 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4800
      TabIndex        =   82
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label comtotd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   4800
      TabIndex        =   81
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Label comimps 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3120
      TabIndex        =   80
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label comimpd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3120
      TabIndex        =   79
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Label comsubs 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1440
      TabIndex        =   78
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label comsubd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1440
      TabIndex        =   77
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Label Label24 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "US$."
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   360
      TabIndex        =   76
      Top             =   8880
      Width           =   975
   End
   Begin VB.Label zooma 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   13800
      TabIndex        =   38
      Top             =   8280
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label YacaRGA 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   13800
      TabIndex        =   31
      Top             =   8040
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label subtotald 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6600
      TabIndex        =   5
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Label subtotals 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label impuestod 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Label impuestos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Label totald 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9960
      TabIndex        =   1
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Label totals 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   9960
      TabIndex        =   0
      Top             =   8520
      Width           =   1695
   End
   Begin VB.Menu djku232 
      Caption         =   "&Nuevo"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu agt62323 
      Caption         =   "&Borrar"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu modi343 
      Caption         =   "&Desmarca"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mio8923 
      Caption         =   "&Modifica"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu anulier 
      Caption         =   "&Anular"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu dkiw232 
      Caption         =   "&Imprimir"
      Begin VB.Menu dkifor 
         Caption         =   "&1.FormatoDefinido"
      End
      Begin VB.Menu dkiewre 
         Caption         =   "&2.Reporteador"
      End
      Begin VB.Menu dl89er 
         Caption         =   "&3.Excell Impresion Total"
      End
      Begin VB.Menu dki889343 
         Caption         =   "&4.Excell Impresion solo seleccionado"
      End
      Begin VB.Menu impso02 
         Caption         =   "&5.Excell Impresion solo Documentos"
      End
   End
   Begin VB.Menu mit56232 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu fdl89234 
      Caption         =   "&Validar"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu djbu232 
      Caption         =   "&Consulta"
   End
   Begin VB.Menu fk4844 
      Caption         =   "&Generar"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu ldo33 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcomvta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rexplorap As New ADODB.Recordset

Private Sub agt62323_Click()

    Dim buf1 As String

    On Error GoTo cmd6_err

    'If Frame4.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    If "" & DBGrid2.columns(1) <> "0" Then
        MsgBox "Para Borrar el documento debe estar en estado=0", 48, "Aviso"
        Exit Sub

    End If

    If MsgBox("Desea Borrar Documento", 1, "Aviso") <> 1 Then Exit Sub
    'MsgBox cgusuario
    buf1 = " and acu='" & "" & DBGrid2.columns("acu") & "'"
    cn.Execute "DELETE FROM  " & dgusuariog & "   where  local='" & "" & DBGrid2.columns(2) & "' and tipo='" & "" & DBGrid2.columns(3) & "' and serie='" & "" & DBGrid2.columns(4) & "' and  numero='" & "" & DBGrid2.columns(5) & "'" & buf1
    cn.Execute "DELETE FROM  fpagov   where  local='" & "" & DBGrid2.columns(2) & "' and tipo='" & "" & DBGrid2.columns(3) & "' and serie='" & "" & DBGrid2.columns(4) & "' and  numero='" & "" & DBGrid2.columns(5) & "'" & buf1
    cn.Execute "DELETE FROM  " & cgusuario & "   where  local='" & "" & DBGrid2.columns(2) & "' and tipo='" & "" & DBGrid2.columns(3) & "' and serie='" & "" & DBGrid2.columns(4) & "' and  numero='" & "" & DBGrid2.columns(5) & "'" & buf1
    MsgBox "Ok,Documento Borrado", 24, "Aviso"
    sql_cabeza
    Exit Sub
cmd6_err:
    MsgBox "Aviso en Borrar " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub anulier_Click()

    Dim buf1 As String

    Dim buf  As String

    Dim Msg  As String

    On Error GoTo cmd8_err

    'If Frame4.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    Msg = "Ojo.. Esta opcion de anular permite poner el documento en modo de anulacion ,luego de realizacion" + Chr$(10) + Chr$(13)
    Msg = Msg + "No puede ya reversar.... " + Chr$(10) + Chr$(13)

    If MsgBox(Msg, 1, "Aviso") <> 1 Then Exit Sub

    If "" & DBGrid2.columns(1) = "2" Then
        MsgBox "Para anular el documento debe estar en estado=0 or estado=1", 48, "Aviso"
        Exit Sub

    End If

    If MsgBox("Desea Anular Documento,Quedara inmodificable ", 1, "Aviso") <> 1 Then Exit Sub
    buf = "1"

    If "" & DBGrid2.columns(1) = "1" Then
        buf = "0"

    End If

    'Data2.Recordset.Edit
    'Data2.Recordset.Fields("estado") = buf
    'Data2.Recordset.Update
    'MsgBox cgusuario
    buf1 = " and acu='" & "" & DBGrid2.columns("acu") & "'"
    cn.Execute ("update  " & dgusuariog & " set estado='" & buf & "'  where  local='" & "" & DBGrid2.columns("local") & "' and tipo='" & "" & DBGrid2.columns("tipo") & "' and serie='" & "" & DBGrid2.columns("serie") & "' and  numero='" & "" & DBGrid2.columns("numero") & "'" & buf1)
    cn.Execute ("update  fpagov  set estado='" & buf & "'  where  local='" & "" & DBGrid2.columns("local") & "' and tipo='" & "" & DBGrid2.columns("tipo") & "' and serie='" & "" & DBGrid2.columns("serie") & "' and  numero='" & "" & DBGrid2.columns("numero") & "'" & buf1)
    cn.Execute ("update  " & cgusuario & " set estado='" & buf & "'  where  local='" & "" & DBGrid2.columns("local") & "' and tipo='" & "" & DBGrid2.columns("tipo") & "' and serie='" & "" & DBGrid2.columns("serie") & "' and  numero='" & "" & DBGrid2.columns("numero") & "'" & buf1)
    MsgBox "Ok,Documento Anulado", 24, "Aviso"
    sql_cabeza
    Exit Sub
cmd8_err:
    MsgBox "Aviso en anular documento " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        ldo33_Click
        Exit Sub

    End If

    Command1_Click

End Sub

Private Sub clave_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(clave) = 0 Then
        clave.SetFocus

    End If

    Command4_Click

End Sub

Private Sub cmdAddEntry_Click()

End Sub

Private Sub cmdDelete_Click()

End Sub

Private Sub cmdExit_Click()

End Sub

Private Sub cmdGrabar_Click()
    sql_cabeza
    Frame3.Visible = False

End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub cmdSort_Click()

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_codigo

    End If

End Sub

Private Sub Command1_Click()
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim buf       As String

    Dim rconsulta As New ADODB.Recordset

    Dim cad       As String

    If opcion1 = "1" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from clientes "
        Else
            buf = "select Nombre,Codigo from clientes where " & Combo1 & " like '" & buffer & "%'"

        End If

    End If

    If opcion1 = "6100" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from tlocal "
        Else
            buf = "select Nombre,Codigo from tlocal where " & Combo1 & " like '" & buffer & "%'"

        End If

    End If

    If opcion1 = "2" Then
        buf = "select Producto,Descripcio,Unidad as Und,Factor as Fac,Precio,Cantidad as Cant,Total,Local,Deslipo as Dscto from  " & dgusuariog & " where local='" & "" & DBGrid2.columns("local") & "' and serie='" & "" & DBGrid2.columns("serie") & "' and numero='" & "" & DBGrid2.columns("numero") & "'"

    End If

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True And rconsulta.BOF = True Then
        rconsulta.Close
        buffer.SetFocus
        Exit Sub

    End If

    Set dbGrid1.DataSource = rconsulta

    If opcion1 = "1" Or opcion1 = "6100" Then
        dbGrid1.columns(0).Width = 4000
        dbGrid1.columns(1).Width = 2000

    End If

    If opcion1 = "2" Then
        dbGrid1.columns(0).Width = 1500
        dbGrid1.columns(1).Width = 5000
        dbGrid1.columns(2).Width = 900
        dbGrid1.columns(3).Width = 900
        dbGrid1.columns(4).Width = 900
        dbGrid1.columns(5).Width = 900
        dbGrid1.columns(6).Width = 1500
        dbGrid1.columns(7).Width = 900
        dbGrid1.columns(8).Width = 700

    End If
   
    If sw = 1 Then
        dbGrid1.SetFocus

    End If

End Sub

Private Sub Command10_Click()

End Sub

Private Sub Command2_Click()
    ldo33_Click

End Sub

Private Sub Command3_Click()
    ldo33_Click

End Sub

Private Sub Command4_Click()

    On Error GoTo cmd7_err

    Dim found As Integer

    Dim buf   As String

    If Len(clave) = 0 Then
        clave.SetFocus
        Exit Sub

    End If

    found = valida_clave("" & clave)

    If found = 0 Then
        MsgBox "Clave no valida para realizar este proceso ", 48, "Aviso"
        clave = ""
        clave.SetFocus
        Exit Sub

    End If

    'MsgBox ""
    If Frame2.Caption = "DESMARCA" Then
        If MsgBox("Desea Desmarca el Documento", 1, "Aviso") <> 1 Then Exit Sub
        If Trim("" & DBGrid2.columns("acu")) = "A" Or Trim("" & DBGrid2.columns("acu")) = "B" Or Trim("" & DBGrid2.columns("acu")) = "C" Or Trim("" & DBGrid2.columns("acu")) = "D" Or Trim("" & DBGrid2.columns("acu")) = "G" Then  'ventas
            buf = "cuentacd"
            'MsgBox ""
            found = verificar_recibo(buf, Trim(DBGrid2.columns(2)), Trim(DBGrid2.columns(3)), Trim(DBGrid2.columns(4)), Trim(DBGrid2.columns(5)))

            If found = 1 Then
                MsgBox "Ya existe recibo ", 48, "Aviso"
                Exit Sub

            End If

            'MsgBox ""
        End If

        If Trim("" & DBGrid2.columns("acu")) = "J" Or Trim("" & DBGrid2.columns("acu")) = "K" Or Trim("" & DBGrid2.columns("acu")) = "L" Or Trim("" & DBGrid2.columns("acu")) = "M" Or Trim("" & DBGrid2.columns("acu")) = "P" Then  'ventas
            buf = "cuentaPd"
            'MsgBox ""
            found = verificar_recibo(buf, Trim(DBGrid2.columns(2)), Trim(DBGrid2.columns(3)), Trim(DBGrid2.columns(4)), Trim(DBGrid2.columns(5)))

            If found = 1 Then
                MsgBox "Ya existe recibo ", 48, "Aviso"
                Exit Sub

            End If

        End If

        'MsgBox ""
        desmarca_documento

    End If

    Frame2.Visible = False
    Exit Sub
cmd7_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    Frame2.Visible = False
    Exit Sub

End Sub

Private Sub Command5_Click()
    sql_cabeza

End Sub

Private Sub Command6_Click()

    Dim sdx      As Double

    Dim buf      As String

    Dim bufca    As String

    Dim I        As Integer

    Dim bufde    As String

    Dim mytablez As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    If Command6.Caption = "Selecciona" Then
        mytablex.Open "SELECT * FROM tipo where tipo='" & extra_loquesea(gtipo) & "'", cn, adOpenKeyset, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            gserie = ""
            gnumero = ""
            gacu = ""
            mytablex.Close
            gserie.SetFocus
            Exit Sub

        End If

        gacu = "" & mytablex.Fields("tipodoc")
        gserie = "" & mytablex.Fields("serie")
        sdx = Val("" & mytablex.Fields("numero")) + 1
        gnumero = "" & sdx
        mytablex.Close
        Command6.Caption = "Generar"
        Exit Sub

    End If

    Command6.Caption = "Selecciona"

    Select Case Trim(gacu)

        Case "A", "B", "C", "D", "E", "G", "F"  'VENTAS
            bufca = "factura"
            bufde = "detalle"

        Case "H"  'COTIZACION
            bufca = "ccotizav"
            bufde = "dcotizav"

        Case "I"  'PEDIDO
            bufca = "cpedidov"
            bufde = "dpedidov"

        Case "J", "K", "L", "M", "P", "N", "O"  'COMPRAS
            bufca = "factura"
            bufde = "detalle"

        Case "Q"  'PEDIDO COMPRA
            bufca = "cpedidoc"
            bufde = "dpedidoc"

        Case "R"  'ORDEN COMPRA
            bufca = "cordenc"
            bufde = "dordenc"

        Case "S", "T"
            bufca = "factura"
            bufde = "detalle"

        Case "Z"
            bufca = "ctraslad"
            bufde = "dtraslad"

        Case Else
            Exit Sub

    End Select

    If Len(gserie) = 0 Then
        gserie.SetFocus
        Exit Sub

    End If

    'cabecera
ax:

    If mytablex.State = 1 Then mytablex.Close
    buf = "SELECT * FROM " & bufca & " where local='" & rexplorap.Fields("local") & "' and tipo='" & extra_loquesea(gtipo) & "' and serie='" & gserie & "' and numero='" & gnumero & "'"
    mytablex.Open buf, cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sdx = Val(gnumero) + 1
        gnumero = "" & sdx
        GoTo ax
        Exit Sub

    End If

    mytablez.Open "SELECT * FROM tipo where tipo='" & extra_loquesea(gtipo) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablez.RecordCount > 0 Then
        mytablez.Fields("numero") = gnumero
        mytablez.Update

    End If

    mytablex.AddNew

    For I = 0 To rexplorap.Fields.count - 1
        mytablex.Fields(I) = rexplorap.Fields(I)
    Next I

    mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("fechae") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("fechasunat") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("tipo") = extra_loquesea(gtipo)
    mytablex.Fields("serie") = gserie
    mytablex.Fields("numero") = gnumero
    mytablex.Fields("acu") = gacu
    mytablex.Fields("estado") = "0"

    mytablex.Fields("tipo1") = "" & rexplorap.Fields("tipo")
    mytablex.Fields("serie1") = rexplorap.Fields("serie")
    mytablex.Fields("numero1") = rexplorap.Fields("numero")

    mytablex.Update
    mytablex.Close

    mytabley.Open "SELECT * FROM " & dgusuariog & " where local='" & rexplorap.Fields("local") & "' and tipo='" & rexplorap.Fields("tipo") & "' and serie='" & rexplorap.Fields("serie") & "' and numero='" & rexplorap.Fields("numero") & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        mytabley.Close
        Exit Sub

    End If

    mytablex.Open "SELECT * FROM " & bufde, cn, adOpenKeyset, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew

        For I = 0 To mytabley.Fields.count - 1
            mytablex.Fields(I) = mytabley.Fields(I)
        Next I

        mytablex.Fields("tipo") = extra_loquesea(gtipo)
        mytablex.Fields("serie") = gserie
        mytablex.Fields("numero") = gnumero
        mytablex.Fields("acu") = gacu
        mytablex.Fields("estado") = "0"
        mytablex.Fields("fecha") = Format(Now, "dd/mm/yyyy")
        mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
        mytablex.Update
        mytabley.MoveNext
    Loop
    mytabley.Close
    mytablex.Close
    'rexplorap.Fields("serie1") = gserie
    'rexplorap.Fields("numero1") = gnumero
    'rexplorap.Fields("tipo1") = extra_loquesea(gtipo)
    rexplorap.Fields("yausado") = "1"
    rexplorap.Update
    MsgBox "Proceso Realizado", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command7_Click()
    Frame4.Visible = False

End Sub

Private Sub Command9_Click()
    ldo33_Click

End Sub

Private Sub Command8_Click()
    SUMAR_CABEZA

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            codigo = dbGrid1.columns(1)
            Frame1.Visible = False
            Frame1.Enabled = False
            codigo.SetFocus

        End If

        If opcion1 = "6100" Then
            mytablex.Open "SELECT * FROM userlocal where codigo='" & gusuario & "' and local='" & dbGrid1.columns(1) & "'", cn, adOpenKeyset, adLockOptimistic

            If mytablex.RecordCount = 0 Then
                mytablex.Close
                MsgBox "Usuario No autorizado,utilizar este local ", 48, "Aviso"
                Exit Sub

            End If

            mytablex.Close
   
            buf = dbGrid1.columns(1)
            Frame1.Visible = False
            Frame1.Enabled = False

            menu_nuevo buf

            'codigo.SetFocus
        End If

    End If

End Sub

Private Sub dbgrid2_BeforeColEdit(ByVal ColIndex As Integer, _
                                  ByVal KeyAscii As Integer, _
                                  Cancel As Integer)

    If ColIndex <> 13 Then
        Cancel = True
        Exit Sub

    End If

    Select Case ColIndex

        Case 13

            If Len("" & DBGrid2.columns(2)) = 0 Then
                Cancel = True
                Exit Sub

            End If

    End Select

End Sub

Private Sub dbgrid2_BeforeColUpdate(ByVal ColIndex As Integer, _
                                    OldValue As Variant, _
                                    Cancel As Integer)

    Dim found As Integer

    Select Case ColIndex

        Case 13

            If Len(DBGrid2.columns(2)) = 0 Then
                Cancel = True
                Exit Sub

            End If

            If Len("" & DBGrid2.columns(13)) = 0 Then
                DBGrid2.columns(13) = DBGrid2.columns("fecha")
                Exit Sub

            End If

            found = valida_fecha("" & DBGrid2.columns(13))

            If found = 0 Then
                Cancel = True
                Exit Sub

            End If

    End Select

End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'consultando
        consulta_detalle

    End If

End Sub

Private Sub DBGrid4_Click()

End Sub

Sub cerrar_data4()

End Sub

Private Sub djbu232_Click()

    'If Frame4.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    Frame3.Visible = True

End Sub

Private Sub djku232_Click()
    'If Frame4.Visible = True Then Exit Sub
    'If Frame2.Visible = True Then Exit Sub
    'If Frame3.Visible = True Then Exit Sub
    'consulta_local

End Sub

Sub menu_nuevo(buf As String)

    Dim found As Integer

    On Error GoTo cmd28_err

    tfactura.local1 = buf

    If acu = "Z" Then
        'tfactura.local1 = "01"
        tfactura.codigo = "01"
        tfactura.Label2.Caption = "Cod.Int"
        'tfactura.Label14.Visible = True
        tfactura.Label38.Visible = True
        'tfactura.localf.Visible = True
        'inicio 10/02/2018 pll
        'tfactura.bodegaf.Visible = True
        'fin 10/02/2018 pll
   
    End If

    If acu = "V" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Facturacion x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        'tfactura.caja = "00"
        tfactura.bandera = "Nuevo"
        tfactura.acu = "V"
        tfactura.tipoclie = tipoclie

        'tfactura.local1=local
        tfactura.Show 1

    End If

    If acu = "H" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Cotizaciones Ventas"
        cgusuario = "CCOTIZAV"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dcotizav"
        tfactura.bandera = "Nuevo"
        tfactura.acu = "H"
        tfactura.tipoclie = tipoclie
        tfactura.Show 1

    End If

    If acu = "I" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Cotizaciones Ventas"
        cgusuario = "Cpedidov"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dpedidov"
        tfactura.bandera = "Nuevo"
        tfactura.acu = "I"
        tfactura.tipoclie = tipoclie
        tfactura.Show 1

    End If

    If acu = "T" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Guia Salida"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Nuevo"
        tfactura.acu = "T"
        tfactura.tipoclie = tipoclie
        tfactura.Show 1

    End If

    If acu = "E" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Nota Credito Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Nuevo"
        tfactura.acu = "E"
        tfactura.tipoclie = tipoclie
        tfactura.Show 1

    End If

    If acu = "F" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Nota debito Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        tfactura.bandera = "Nuevo"
        dgusuariog = "DETALLE"
        tfactura.acu = "F"
        tfactura.tipoclie = tipoclie
        tfactura.Show 1

    End If

    If acu = "R" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   
            Exit Sub

        End If

        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Orden de Compra"
        cgusuario = "CORDENC"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DORDENC"
        tfactura.acu = "R"
        tfactura.bandera = "Nuevo"
        tfactura.tipoclie = tipoclie
        tfactura.Show 1

    End If

    If acu = "S" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   
            Exit Sub

        End If

        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Guia Remision Entrada"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.tipoclie = tipoclie
        tfactura.acu = "S"
        tfactura.bandera = "Nuevo"
        tfactura.Show 1

    End If

    If acu = "C" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   
            Exit Sub

        End If

        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Factura de Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.acu = "C"
        tfactura.bandera = "Nuevo"
        tfactura.tipoclie = tipoclie
        tfactura.Show 1

    End If

    If acu = "N" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Nota Credito Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.acu = "N"
        tfactura.bandera = "Nuevo"
        tfactura.tipoclie = tipoclie
        tfactura.Show 1

    End If

    If acu = "O" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Nota debito de Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.acu = "O"
        tfactura.bandera = "Nuevo"
        tfactura.tipoclie = tipoclie
        tfactura.Show 1

    End If

    If acu = "Q" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            Exit Sub

        End If

        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Pedido Almacen"
        cgusuario = "CREQUISA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DREQUISA"
        tfactura.acu = "Q"
        tfactura.bandera = "Nuevo"
        tfactura.tipoclie = tipoclie
        tfactura.Show 1

    End If

    If acu = "Z" Then
        found = copiar_temporal()

        If found = 0 Then
            MsgBox "Error al copiar archivo temporal", 24, "Aviso"
            End
            Exit Sub

        End If

        'tfactura.Label2 = "Cod.Inicio"
        tfactura.Caption = "Traslado entre almacen de un mismo establecimiento"
        cgusuario = "CTRASLAD"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DTRASLAD"
        tfactura.acu = "Z"
        tfactura.bandera = "Nuevo"
        tfactura.tipoclie = tipoclie
        tfactura.Show 1

    End If

    sql_cabeza
    MsgBox "Presione tecla para continuar..", 48, "Aviso"
    Exit Sub
cmd28_err:
    MsgBox "Nuevo:Seleccione un dato", 48, "Aviso"
    Exit Sub

End Sub

Private Sub dki889343_Click()

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If MsgBox("Desea Exportar Excell", 1, "Aviso") <> 1 Then Exit Sub
    conteo_excell_uno

End Sub

Private Sub dkiewre_Click()
    'If Frame4.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    reporgen.NAMETABLA = cgusuario
    reporgen.Show 1

End Sub

Private Sub dkifor_Click()
    'If Frame4.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    proceso_impresion1

End Sub

Private Sub dl89er_Click()

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If MsgBox("Desea Exportar Excell", 1, "Aviso") <> 1 Then Exit Sub
    menu_excell

End Sub

Private Sub fdl89234_Click()

    Dim buf As String

    On Error GoTo cmd45112_err

    buf = DBGrid2.columns("local")

    If Trim("" & DBGrid2.columns(1)) <> "2" Then
        MsgBox "Para este fin el estado debe estar en 2", 48, "Aviso"
        Exit Sub

    End If

    'Select Case acu
    '       Case "Z", "S", "T"
    '       Case Else: Exit Sub
    'End Select

    If Trim("" & DBGrid2.columns(0)) = "0" Then
        If MsgBox("Estado Actual:Pendiente " & Chr$(10) & Chr$(13) & "Cambiar a Atendido", 1, "Aviso") = 1 Then
            DBGrid2.columns(0) = "1"
            Exit Sub

        End If

    End If

    If Trim("" & DBGrid2.columns(0)) = "1" Then
        If MsgBox("Estado Actual:Atendido " & Chr$(10) & Chr$(13) & "Cambiar a Pendiente", 1, "Aviso") = 1 Then
            DBGrid2.columns(0) = "0"

        End If

    End If

    Exit Sub
cmd45112_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fk4844_Click()

    On Error GoTo cmd8912_err

    Dim mytablex As New ADODB.Recordset

    If "" & rexplorap.Fields("estado") <> "2" Then
        MsgBox "Estado debe estar en 2 para realizar esta operacion ", 48, "Aviso"
        Exit Sub

    End If

    If "" & rexplorap.Fields("yausado") = "1" Then
        MsgBox "Ya fue Utilizado ", 48, "Aviso"
        Exit Sub

    End If

    'If Len("" & rexplorap.Fields("serie1")) > 0 Then
    '   MsgBox "Documento ya generado " & "" & rexplorap.Fields("tipo1") & " " & rexplorap.Fields("serie1") & " " & rexplorap.Fields("numero1"), 48, "Aviso"
    '   Exit Sub
    'End If
    gacu = ""
    gserie = ""
    gnumero = ""
    gtipo.Clear
    gtipo.AddItem ""
    mytablex.Open "SELECT * FROM tipo  ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If acu = "T" Then  'Guia remision
            If "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "D" Or "" & mytablex.Fields("tipodoc") = "G" Or "" & mytablex.Fields("tipodoc") = "E" Or "" & mytablex.Fields("tipodoc") = "F" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        If acu = "V" Then  'Cotizacion
            If "" & mytablex.Fields("tipodoc") = "E" Or "" & mytablex.Fields("tipodoc") = "F" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        If acu = "H" Then  'Cotizacion
            If "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "D" Or "" & mytablex.Fields("tipodoc") = "G" Or "" & mytablex.Fields("tipodoc") = "E" Or "" & mytablex.Fields("tipodoc") = "F" Or "" & mytablex.Fields("tipodoc") = "I" Or "" & mytablex.Fields("tipodoc") = "T" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        If acu = "I" Then  'pedido
            If "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "D" Or "" & mytablex.Fields("tipodoc") = "G" Or "" & mytablex.Fields("tipodoc") = "E" Or "" & mytablex.Fields("tipodoc") = "F" Or "" & mytablex.Fields("tipodoc") = "T" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        If acu = "R" Then  'orden de compra
            If "" & mytablex.Fields("tipodoc") = "J" Or "" & mytablex.Fields("tipodoc") = "K" Or "" & mytablex.Fields("tipodoc") = "L" Or "" & mytablex.Fields("tipodoc") = "M" Or "" & mytablex.Fields("tipodoc") = "P" Or "" & mytablex.Fields("tipodoc") = "N" Or "" & mytablex.Fields("tipodoc") = "O" Or "" & mytablex.Fields("tipodoc") = "S" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        If acu = "S" Then  'guia de compra
            If "" & mytablex.Fields("tipodoc") = "J" Or "" & mytablex.Fields("tipodoc") = "K" Or "" & mytablex.Fields("tipodoc") = "L" Or "" & mytablex.Fields("tipodoc") = "M" Or "" & mytablex.Fields("tipodoc") = "P" Or "" & mytablex.Fields("tipodoc") = "N" Or "" & mytablex.Fields("tipodoc") = "O" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        If acu = "C" Then  'factura de compra
            If "" & mytablex.Fields("tipodoc") = "O" Or "" & mytablex.Fields("tipodoc") = "N" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        If acu = "Q" Then  'nota pedido almacen
            If "" & mytablex.Fields("tipodoc") = "Z" Then
                gtipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")

            End If

        End If

        mytablex.MoveNext
    Loop
    mytablex.Close
    gtipo.ListIndex = 0
    Frame4.Visible = True
    Command6.Caption = "Selecciona"
    gtipo.SetFocus
    Exit Sub
cmd8912_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    carga_iniciales
    Check1.Visible = False

    'MsgBox acu
    Select Case acu

        Case "V"
            Check1.Visible = True

    End Select

    If zooma = "Zomm" Then
        Frame3.Visible = False
        zooma = ""
        Exit Sub

    End If

    zooma = ""
    'If YacaRGA = "" Then
    cmdGrabar_Click
    'End If

End Sub

Sub color_cambio()

End Sub

Private Sub Form_Load()

    Dim mytablex As New ADODB.Recordset

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

    Combo2.Clear
    Combo2.AddItem "Pendiente"
    Combo2.AddItem "Atendido"
    Combo2.AddItem "%"
    Combo2.ListIndex = 0

    moneda.Clear
    moneda.AddItem "%"
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0
    estado.Clear
    estado.AddItem "%"
    estado.AddItem "2"
    estado.AddItem "1"
    estado.AddItem "0"

    fechai = Format(Now, "dd/mm/yyyy")
    fechaf = Format(Now, "dd/mm/yyyy")
    estado.ListIndex = 0

    'cmdGrabar_Click
    'MsgBox ""
End Sub

Sub carga_iniciales()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    cajero.Clear
    cajero.AddItem "%"
    vendedor.Clear
    vendedor.AddItem "%"
    caja.Clear
    caja.AddItem "%"
    tipo.Clear
    tipo.AddItem "%"
    bodega.Clear
    bodega.AddItem "%"
    local1.Clear
    local1.AddItem "%"
    bodegaf.Clear
    bodegaf.AddItem "%"

    mytablex.Open "SELECT * FROM vendedor ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        cajero.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    vendedor.ListIndex = 0
    cajero.ListIndex = 0
    mytablex.Close

    mytablex.Open "SELECT * FROM tlocal", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        local1.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    local1.ListIndex = 0

    If local1.ListCount = 2 Then
        local1.ListIndex = 1

    End If

    mytablex.Close

    'If LOCAL1 <> "%" Then
    'buf = " where local='" & extra_loquesea(LOCAL1) & "'"
    'End If
    mytablex.Open "SELECT * FROM tipo ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("grupo") = acu Then
            tipo.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")
        End If '

        mytablex.MoveNext
    Loop
    tipo.ListIndex = 0
    mytablex.Close

    mytablex.Open "SELECT * FROM bodega ", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        bodega.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        bodegaf.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")

        mytablex.MoveNext
    Loop
    bodega.ListIndex = 0
    bodegaf.ListIndex = 0

    mytablex.Close

    If local1 <> "%" Then
        buf = " where local='" & extra_loquesea(local1) & "'"

    End If

    mytablex.Open "SELECT * FROM parameca", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        If "" & mytablex.Fields("terminal") = "C" Then
            caja.AddItem "" & mytablex.Fields("caja") & "|" & mytablex.Fields("descripcio")

        End If

        mytablex.MoveNext
    Loop
    caja.ListIndex = 0
    mytablex.Close

End Sub

Private Sub impso02_Click()
    menu_excell1

End Sub

Private Sub ldo33_Click()
    'If Frame6.Visible = True Then
    '   Frame6.Visible = False
    '   dbgrid2.SetFocus
    '   Exit Sub
    'End If

    'If Frame4.Visible = True Then
    '   Frame4.Visible = False
    '   dbgrid2.SetFocus
    '   Exit Sub
    'End If
    If Frame3.Visible = True Then
        Frame3.Visible = False
        DBGrid2.SetFocus
        Exit Sub

    End If

    If Frame2.Visible = True Then
        Frame2.Visible = False
        DBGrid2.SetFocus
        Exit Sub

    End If

    If opcion1 = "1" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            codigo.SetFocus
            Exit Sub

        End If

    End If

    If opcion1 = "6100" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            Frame1.Enabled = False
            'codigo.SetFocus
            Exit Sub

        End If

    End If

    If opcion1 = "2" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            DBGrid2.SetFocus
            Exit Sub

        End If

    End If

    tcomvta.Hide
    Unload tcomvta

End Sub

Sub sql_cabeza()

    Dim buf As String

    On Error GoTo cmd921_err

    'MsgBox caja
    'MsgBox fechai
    If Len(fechai) <> 10 Then Exit Sub
    If Len(fechaf) <> 10 Then Exit Sub
    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    If acu <> "V" And acu <> "C" And acu <> "%" Then Exit Sub

    'MsgBox cgusuario
    buf = "select * from " & cgusuario & " where "
    buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
    buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "

    'buf = buf & "  fecha>='" & fechai & "'"
    'buf = buf & "  and  fecha<='" & fechaf & "'"

    If tipo <> "%" Then
        buf = buf & " and tipo like '" & extra_loquesea(tipo) & "'"

    End If

    If caja <> "%" Then
        buf = buf & " and caja like '" & extra_loquesea(caja) & "'"

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

    If moneda <> "%" Then
        buf = buf & " and moneda like '" & moneda & "'"

    End If

    If estado <> "%" Then
        buf = buf & " and estado like '" & estado & "'"

    End If

    If local1 <> "%" Then
        buf = buf & " and local like '" & extra_loquesea(local1) & "'"

    End If

    If vendedor <> "%" Then
        buf = buf & " and vendedor like '" & extra_loquesea(vendedor) & "'"

    End If

    If cajero <> "%" Then
        buf = buf & " and usuario like '" & extra_loquesea(cajero) & "'"

    End If

    If bodega <> "%" Then
        buf = buf & " and bodega like '" & extra_loquesea(bodega) & "'"

    End If

    If bodegaf <> "%" Then
        buf = buf & " and bodegaf like '" & extra_loquesea(bodegaf) & "'"

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

    'If acu <> "C" And acu <> "V" Then
    '   buf = buf & " and acu='" & acu & "'"
    'End If
    If Combo2 <> "%" Then
        If Combo2 = "Atendido" Then
            buf = buf & " and  yausado='1'"

        End If

        If Combo2 = "Pendiente" Then
            buf = buf & " and  yausado='0'"

        End If

    End If

    If acu = "V" Then
        buf = buf & " and (acu='A' OR acu='B' OR acu='C' OR acu='D' OR acu='G' OR acu='E' OR acu='F')"

        If Check1.Value = 1 Then
            buf = buf & " and tipo<>'5'"

        End If

    End If

    If acu = "C" Then
        buf = buf & " and (acu='J' OR acu='K' OR acu='L' OR acu='M' OR acu='P' OR acu='N' OR acu='O')"

        'If Check1.Value = 1 Then
        '   buf = buf & " and tipo<>'5'"
        'End If
    End If

    If acu = "%" Then
        buf = buf & " and (acu='A' OR acu='B' OR acu='C' OR acu='D' OR acu='G' OR acu='E' OR acu='F' "
        buf = buf & "  or acu='J' OR acu='K' OR acu='L' OR acu='M' OR acu='P' OR acu='N' OR acu='O')"

        If Check1.Value = 1 Then
            buf = buf & " and tipo<>'5'"

        End If

        'If Check1.Value = 1 Then
        '   buf = buf & " and tipo<>'5'"
        'End If
    End If

    If acu <> "Z" Then

        'buf = buf & " and importacio<>'S' "
    End If

    If acu = "Q" Then
        buf = buf & " and tipoclie='V'"

    End If

    If tinterno = "S" Then
        buf = buf & " and tipoclie='V'"
    Else

        'buf = buf & " and tipoclie<>'I'"
    End If

    buf = buf & " order by fecha,Hora,tipo,serie,numero"
    'MsgBox buf

    If rexplorap.State = 1 Then rexplorap.Close
    rexplorap.Open buf, cn, adOpenStatic, adLockOptimistic
   
    'MsgBox ""
    Set DBGrid2.DataSource = rexplorap
    'If rexplorap.EOF = True And rexplorap.BOF = True Then
    'rconsulta.Close
    'buffer.SetFocus
    'Exit Sub
    'End If
   
    If rexplorap.RecordCount > 0 Then
        DBGrid2.Col = 0
        DBGrid2.Row = DBGrid2.VisibleRows - 1
        DBGrid2.SetFocus

    End If

    'Data2.Connect = "foxpro 2.5;"
    'Data2.DatabaseName = globaldir
    'Data2.RecordSource = buf
    'Data2.Refresh
    'SUMAR_CABEZA rexplorap
    'ir_ultimo
               
    'MsgBox "xxx"
    'If rexplorap.RecordCount > 0 Then
    '   dbgrid2.Col = 0
    '   dbgrid2.Row = dbgrid2.VisibleRows - 1
    '   dbgrid2.SetFocus
    'End If
               
    Exit Sub
cmd921_err:
    MsgBox "aviso en sql_cabeza   " & error$, 48, "Aviso"
    Exit Sub

End Sub

Sub SUMAR_CABEZA()

    Dim xigv  As Double

    Dim sdx   As Double

    Dim sdx1  As Double

    Dim sdx2  As Double

    Dim sdx3  As Double

    Dim sdx4  As Double

    Dim sdx5  As Double

    Dim sdx6  As Double

    Dim sdx11 As Double

    Dim sdx21 As Double

    Dim sdx31 As Double

    Dim sdx41 As Double

    Dim sdx51 As Double

    Dim sdx61 As Double

    On Error GoTo cmd7812_err

    sdx1 = 0
    sdx2 = 0
    sdx3 = 0

    sdx4 = 0
    sdx5 = 0
    sdx6 = 0

    sdx11 = 0
    sdx21 = 0
    sdx31 = 0

    sdx41 = 0
    sdx51 = 0
    sdx61 = 0
    xigv = 2.58

    If rexplorap.RecordCount = 0 Then
        Exit Sub

    End If

    rexplorap.MoveFirst
    Do

        If rexplorap.EOF Or rexplorap.BOF Then Exit Do
        If "" & rexplorap.Fields("estado") = "2" Then

            'ventas
            If "" & rexplorap.Fields("acu") = "A" Or "" & rexplorap.Fields("acu") = "B" Or "" & rexplorap.Fields("acu") = "C" Or "" & rexplorap.Fields("acu") = "D" Or "" & rexplorap.Fields("acu") = "G" Then
                If "" & rexplorap.Fields("moneda") = "S" Then
                    sdx1 = sdx1 + Val("" & rexplorap.Fields("subtotal"))
                    sdx2 = sdx2 + Val("" & rexplorap.Fields("impuesto"))
                    sdx3 = sdx3 + Val("" & rexplorap.Fields("total"))

                End If

                If "" & rexplorap.Fields("moneda") = "D" Then
                    sdx4 = sdx4 + Val("" & rexplorap.Fields("subtotal"))
                    sdx5 = sdx5 + Val("" & rexplorap.Fields("impuesto"))
                    sdx6 = sdx6 + Val("" & rexplorap.Fields("total"))

                End If

            End If

            'compras
            If "" & rexplorap.Fields("acu") = "J" Or "" & rexplorap.Fields("acu") = "K" Or "" & rexplorap.Fields("acu") = "L" Or "" & rexplorap.Fields("acu") = "M" Or "" & rexplorap.Fields("acu") = "P" Then
                If "" & rexplorap.Fields("moneda") = "S" Then
                    sdx11 = sdx11 + Val("" & rexplorap.Fields("subtotal"))
                    sdx21 = sdx21 + Val("" & rexplorap.Fields("impuesto"))
                    sdx31 = sdx31 + Val("" & rexplorap.Fields("total"))

                End If

                If "" & rexplorap.Fields("moneda") = "D" Then
                    sdx41 = sdx41 + Val("" & rexplorap.Fields("subtotal"))
                    sdx51 = sdx51 + Val("" & rexplorap.Fields("impuesto"))
                    sdx61 = sdx61 + Val("" & rexplorap.Fields("total"))

                End If

            End If

        End If

        rexplorap.MoveNext
    Loop
    subtotals = Format(sdx1, "0.00")
    impuestos = Format(sdx2, "0.00")
    totals = Format(sdx3, "0.00")
    subtotald = Format(sdx4, "0.00")
    impuestod = Format(sdx5, "0.00")
    totald = Format(sdx6, "0.00")

    comsubs = Format(sdx11, "0.00")
    comimps = Format(sdx21, "0.00")
    comtots = Format(sdx31, "0.00")
    comsubd = Format(sdx41, "0.00")
    comimpd = Format(sdx51, "0.00")
    comtotd = Format(sdx61, "0.00")

    sdx = Val(comtots) + Val(comtotd) * xigv
    totalco = Format(sdx, "0.00")

    sdx = Val(totals) + Val(totald) * xigv
    totalvt = Format(sdx, "0.00")

    sdx = -Val(totalvt) + Val(totalco)
    totalneto = Format(sdx, "00,000.00")

    sdx = Val(comimps) + Val(comimpd) * xigv
    sdx = sdx - Val(impuestos) + Val(impuestod) * xigv
    difigv = Format(sdx, "00,000.00")

    Exit Sub
cmd7812_err:
    MsgBox "Error en Suma" & error$, 48, "Aviso"
    Exit Sub

End Sub

Sub ir_inicio()

End Sub

Sub consulta_codigo()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Telefono"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "1"
    Command1_Click

End Sub

Sub consulta_local()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    buffer = ""
    Frame1.Enabled = True
    buffer.SetFocus
    opcion1 = "6100"
    Command1_Click

End Sub

Sub consulta_detalle()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Telefono"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    buffer = ""
    buffer.SetFocus
    opcion1 = "2"
    Command1_Click

End Sub

Sub ir_ultimo()

    On Error GoTo cmd123_err

    'Data2.Recordset.MoveLast
    Exit Sub
cmd123_err:
    Exit Sub

End Sub

Sub proceso_impresion1()

    Dim found    As Integer

    Dim archivot As String

    Dim ttipo    As String

    Dim tserie   As String

    Dim local1   As String

    Dim tnumero  As String

    On Error GoTo cmd6_err:

    local1 = "" & DBGrid2.columns("local")
    ttipo = "" & DBGrid2.columns("tipo")
    tserie = "" & DBGrid2.columns("serie")
    tnumero = "" & DBGrid2.columns("numero")
    cerrar_archivo
    factura_formato local1, "" & ttipo, "" & tserie, "" & tnumero, "", 0
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    Exit Sub
cmd6_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$
    Exit Sub

End Sub

Sub desmarca_documento()

    Dim buf1  As String

    Dim te    As String

    Dim ts    As String

    Dim found As Integer

    On Error GoTo cmd57_err

    found = valida_flag("" & rexplorap.Fields("acu"))

    If found = 0 Then

    End If

    If found = 1 Or found = 2 Then
        If Len(Trim("" & rexplorap.Fields("tipo1"))) = 0 And Len(Trim("" & rexplorap.Fields("serie1"))) = 0 And Len(Trim("" & rexplorap.Fields("numero1"))) = 0 Then
            descarga_saldo Trim("" & rexplorap.Fields("local")), Trim("" & rexplorap.Fields("tipo")), Trim("" & rexplorap.Fields("serie")), Trim("" & rexplorap.Fields("numero")), 1, "", "" & rexplorap.Fields("tipo1")

        End If

    End If

    If found = 3 Then  'si es traslado
        If Len(Trim("" & rexplorap.Fields("tipo1"))) = 0 And Len(Trim("" & rexplorap.Fields("serie1"))) = 0 And Len(Trim("" & rexplorap.Fields("numero1"))) = 0 Then
            descarga_saldo Trim("" & rexplorap.Fields("local")), "TE", Trim("" & rexplorap.Fields("serie")), Trim("" & rexplorap.Fields("numero")), 1, "1", "" & rexplorap.Fields("local")
            descarga_saldo Trim("" & rexplorap.Fields("local")), "TS", Trim("" & rexplorap.Fields("serie")), Trim("" & rexplorap.Fields("numero")), 1, "1", "" & rexplorap.Fields("local")
            borra_detalle Trim("" & rexplorap.Fields("local")), "TE", Trim("" & rexplorap.Fields("serie")), Trim("" & rexplorap.Fields("numero"))
            borra_detalle Trim("" & rexplorap.Fields("local")), "TS", Trim("" & rexplorap.Fields("serie")), Trim("" & rexplorap.Fields("numero"))

        End If

    End If

    'MsgBox ""

    'buf1 = " and acu='" & Trim("" & rexplorap.fields("acu")) & "'"
    buf1 = "update  " & dgusuariog & " set estado='0'  where local='" & Trim("" & rexplorap.Fields("local")) & "' and  tipo='" & Trim("" & rexplorap.Fields("tipo")) & "' and serie='" & Trim("" & rexplorap.Fields("serie")) & "' and  numero='" & Trim("" & rexplorap.Fields("numero")) & "'" & " and acu='" & Trim("" & rexplorap.Fields("acu")) & "'"
    cn.Execute (buf1)

    'adicionamos la desmarcacion de las guias
    desmarca_yausado "" & rexplorap.Fields("LOCAL"), "" & rexplorap.Fields("tipo"), "" & rexplorap.Fields("SERIE"), "" & rexplorap.Fields("numero")
    'MsgBox ""
    cn.Execute ("update  fpagov  set estado='0'  where  local='" & Trim("" & rexplorap.Fields("local")) & "' and  tipo='" & Trim("" & rexplorap.Fields("tipo")) & "' and serie='" & Trim("" & rexplorap.Fields("serie")) & "' and  numero='" & Trim("" & rexplorap.Fields("numero")) & "'" & " and acu='" & Trim("" & rexplorap.Fields("acu")) & "'")
    cn.Execute ("update  recibo  set usado='N'  where  local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("retipo1") & "' and numero='" & "" & rexplorap.Fields("renumero1") & "'")
    cn.Execute ("update  recibo  set usado='N'  where  local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("retipo1") & "' and numero='" & "" & rexplorap.Fields("renumero2") & "'")
    cn.Execute ("update  recibo  set usado='N'  where  local='" & "" & rexplorap.Fields("local") & "' and tipo='" & "" & rexplorap.Fields("retipo1") & "' and numero='" & "" & rexplorap.Fields("renumero3") & "'")

    'MsgBox ""
    If acu = "Z" Then
        cn.Execute ("DELETE FROM detallE where local='" & "" & rexplorap.Fields("local") & "' and tipo='" & te & "'and serie='" & "" & rexplorap.Fields("serie") & "'  and numero='" & "" & rexplorap.Fields("numero") & "TE" & "'")
        cn.Execute ("DELETE FROM detallE where local='" & "" & rexplorap.Fields("local") & "' and tipo='" & ts & "'and serie='" & "" & rexplorap.Fields("serie") & "'  and numero='" & "" & rexplorap.Fields("numero") & "TS" & "'")

    End If
 
    If valida_flag("" & "" & rexplorap.Fields("acu")) = 1 Or valida_flag("" & "" & rexplorap.Fields("acu")) = 2 Then  'compras o ventas
        found = desgraba_cuentac()

    End If

    MsgBox "Desmarcacion Satisfactoria ", 48, "Aviso"
    sql_cabeza
    Exit Sub
cmd57_err:
    MsgBox "Aviso en desmarca documento " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function valida_flag(buf As String)

    Select Case buf

        Case "Z"
            valida_flag = 3

        Case "T", "A", "B", "C", "D", "G", "E", "F"
            valida_flag = 1

        Case "S", "J", "K", "L", "M", "P", "N", "O"
            valida_flag = 2

    End Select

End Function

Function busca_tipo1(sw As Integer) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where  tipo='" & Trim("" & DBGrid2.columns("tipo")) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        If sw = 0 Then
            busca_tipo1 = "" & mytablex.Fields("te")

        End If

        If sw = 1 Then
            busca_tipo1 = "" & mytablex.Fields("ts")

        End If

    End If

    mytablex.Close

End Function

Sub descarga_saldo(xlocal As String, _
                   xtipo As String, _
                   xserie As String, _
                   xnumero As String, _
                   sw As Integer, _
                   tipoarch As String, _
                   xtipo1 As String)

    Dim sdx       As Double

    Dim signo     As Double

    Dim sww       As Integer

    Dim mytablefa As New ADODB.Recordset

    Dim mytablex  As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    Dim buf       As String

    Dim found     As Integer

    On Error GoTo cmd19_err

    sww = 0
    'AHORA HAY QUE VALIDAR QUE no existe ya cruzado el documento----
    mytablefa.Open "SELECT * FROM " & cgusuario & " where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablefa.RecordCount > 0 Then  'si existe
        If Len(xtipo1) > 0 Then
            found = ve_descarga(xtipo1)

            If found = 1 Then
                sww = 1

            End If

        End If

    End If

    buf = dgusuariog

    If tipoarch = "1" Then
        buf = "detalle"

    End If

    mytablex.Open "SELECT * FROM " & buf & " where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then 'si existe
        Exit Sub

    End If

    'MsgBox ""
    'If permite_entrada_salida("" & mytablex.Fields("acu1")) = 1 Then 'si existe acu1 no descontar
    '   Exit Sub
    'End If
    Do

        If mytablex.EOF Then Exit Do
        '-------------------------------------------------
        signo = 1

        Select Case "" & mytablex.Fields("acu")

            Case "S", "J", "K", "L", "M", "P", "E"
                signo = 1

            Case "T", "A", "B", "C", "D", "G", "N"
                signo = -1

        End Select

        'MsgBox signo
        If "" & mytablex.Fields("acu") = "J" Or "" & mytablex.Fields("acu") = "K" Or "" & mytablex.Fields("acu") = "L" Or "" & mytablex.Fields("acu") = "M" Then 'compras varia el precios y costo

            'graba_costos mytablex
        End If
      
        '-------------------------------------------------
        'busden:
        If sww = 0 Then
            If mytabley.State = 1 Then mytabley.Close
            mytabley.Open "select * from almacen where local='" & Trim("" & mytablex.Fields("local")) & "' and producto='" & Trim("" & mytablex.Fields("producto")) & "' and bodega='" & Trim("" & mytablex.Fields("bodega")) & "'", cn, adOpenDynamic, adLockOptimistic 'adOpenKeyset, adLockOptimistic

            'MsgBox mytabley.RecordCount
            If mytabley.RecordCount = 0 Then 'si existe
                'MsgBox ""
                mytabley.AddNew
                mytabley.Fields("local") = "" & mytablex.Fields("local")
                mytabley.Fields("producto") = "" & mytablex.Fields("producto")
                mytabley.Fields("bodega") = "" & mytablex.Fields("bodega")
                sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
                'MsgBox sdx
                mytabley.Fields("saldo") = sdx
                mytabley.Update
            Else

                If sw = 0 Then
                    'mytabley.Edit
                    'MsgBox ""
                    sdx = Val("" & mytabley.Fields("saldo")) + signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
                    'MsgBox sdx
                    mytabley.Fields("saldo") = sdx
                    decarga_saldo_talla mytabley, mytablex, signo
                    mytabley.Update

                End If

                If sw = 1 Then
                    'mytabley.Edit
         
                    sdx = Val("" & mytabley.Fields("saldo")) - signo * Val("" & mytablex.Fields("cantidad")) * Val("" & mytablex.Fields("factor"))
                    mytabley.Fields("saldo") = sdx
                    decarga_saldo_talla mytabley, mytablex, signo
                    mytabley.Update

                End If

                '-------------------------------
            End If

        End If 'fin sw sw

        '-------------------------------------------------
        mytablex.MoveNext
    Loop
    Exit Sub
cmd19_err:
    MsgBox "Aviso en descarga saldo " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub borra_detalle(xlocal As String, xtipo As String, xserie As String, xnumero As String)
    cn.Execute ("delete from detalle where local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'")

End Sub

Sub desmarca_yausado(buf0 As String, buf1 As String, buf2 As String, buf3 As String)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    On Error GoTo cmd333_err

    buf = "update " & cgusuario & " set estado='0' where local='" & buf0 & "' and tipo='" & buf1 & "' and serie='" & buf2 & "' and numero='" & buf3 & "'"
    cn.Execute (buf)

    mytablex.Open "SELECT * FROM " & cgusuario & " where  local='" & buf0 & "' and tipo='" & buf1 & "' and serie='" & buf2 & "' and numero='" & buf3 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie1"), "" & mytablex.Fields("numero1"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie2"), "" & mytablex.Fields("numero2"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie3"), "" & mytablex.Fields("numero3"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie4"), "" & mytablex.Fields("numero4"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie5"), "" & mytablex.Fields("numero5"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie6"), "" & mytablex.Fields("numero6"), "0"
        descarga_el_uso "" & mytablex.Fields("local"), "" & mytablex.Fields("tipo1"), "" & mytablex.Fields("serie7"), "" & mytablex.Fields("numero7"), "0"

    End If

    '------------------------------------- ------------
    mytablex.Close
    Exit Sub
cmd333_err:
    MsgBox "Aviso en desmarca ya usado " + error$, 48, "Aviso"
    Exit Sub
 
End Sub

Sub descarga_el_uso(buf0 As String, _
                    buf1 As String, _
                    buf2 As String, _
                    buf3 As String, _
                    xsw As String)

    If Len(buf1) = 0 Then Exit Sub
    If Len(buf2) = 0 Then Exit Sub
    If Len(buf3) = 0 Then Exit Sub
    cn.Execute ("update " & cgusuario & " set yausado=" & xsw & " where local='" & buf0 & "' and tipo='" & buf1 & "' and serie='" & buf2 & "' and numero='" & buf3 & "'")

End Sub

Sub decarga_saldo_talla(mytablex As ADODB.Recordset, _
                        mytabley As ADODB.Recordset, _
                        signo As Double)

    Dim sdx As Double

    sdx = Val("" & mytablex.Fields("t1")) + signo * Val("" & mytabley.Fields("t1"))
    mytablex.Fields("t1") = sdx
    sdx = Val("" & mytablex.Fields("t2")) + signo * Val("" & mytabley.Fields("t2"))
    mytablex.Fields("t2") = sdx
    sdx = Val("" & mytablex.Fields("t3")) + signo * Val("" & mytabley.Fields("t3"))
    mytablex.Fields("t3") = sdx
    sdx = Val("" & mytablex.Fields("t4")) + signo * Val("" & mytabley.Fields("t4"))
    mytablex.Fields("t4") = sdx
    sdx = Val("" & mytablex.Fields("t5")) + signo * Val("" & mytabley.Fields("t5"))
    mytablex.Fields("t5") = sdx
    sdx = Val("" & mytablex.Fields("t6")) + signo * Val("" & mytabley.Fields("t6"))
    mytablex.Fields("t6") = sdx
    sdx = Val("" & mytablex.Fields("t7")) + signo * Val("" & mytabley.Fields("t7"))
    mytablex.Fields("t7") = sdx
    sdx = Val("" & mytablex.Fields("t8")) + signo * Val("" & mytabley.Fields("t8"))
    mytablex.Fields("t8") = sdx
    sdx = Val("" & mytablex.Fields("t9")) + signo * Val("" & mytabley.Fields("t9"))
    mytablex.Fields("t9") = sdx
    sdx = Val("" & mytablex.Fields("t10")) + signo * Val("" & mytabley.Fields("t10"))
    mytablex.Fields("t10") = sdx
    sdx = Val("" & mytablex.Fields("t11")) + signo * Val("" & mytabley.Fields("t11"))
    mytablex.Fields("t11") = sdx
    sdx = Val("" & mytablex.Fields("t12")) + signo * Val("" & mytabley.Fields("t12"))
    mytablex.Fields("t12") = sdx
    sdx = Val("" & mytablex.Fields("t13")) + signo * Val("" & mytabley.Fields("t13"))
    mytablex.Fields("t13") = sdx
    sdx = Val("" & mytablex.Fields("t14")) + signo * Val("" & mytabley.Fields("t14"))
    mytablex.Fields("t14") = sdx
    sdx = Val("" & mytablex.Fields("t15")) + signo * Val("" & mytabley.Fields("t15"))
    mytablex.Fields("t15") = sdx
    sdx = Val("" & mytablex.Fields("t16")) + signo * Val("" & mytabley.Fields("t16"))
    mytablex.Fields("t16") = sdx

End Sub

Private Sub local1_Click()
    sql_cabeza

End Sub

Private Sub local1_KeyPress(KeyAscii As Integer)
    sql_cabeza

End Sub

Private Sub mio8923_Click()

    Dim found As Integer

    On Error GoTo cmd27_err

    'If Frame4.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    If "" & DBGrid2.columns(1) <> "0" Then
        MsgBox "Estado debe estar =0", 48, "Aviso"
        Exit Sub

    End If

    If Trim("" & DBGrid2.columns(0)) = "A" Then
        MsgBox "Modo atendido,no se puede modificar ", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    found = copiar_temporal()

    If found = 0 Then
        MsgBox "Error al copiar archivo temporal", 24, "Aviso"
        Exit Sub

    End If

    If acu = "Z" Then
        'tfactura.Label14.Visible = True
        tfactura.Label38.Visible = True
        'tfactura.localf.Visible = True
        'inicio 10/02/2018 pll
        'tfactura.bodegaf.Visible = True
        'fin 10/02/2018 pll
   
        tfactura.Label2.Caption = "Cod.Int."

    End If

    tfactura.zlocal = "" & DBGrid2.columns(2)
    tfactura.ztipo = "" & DBGrid2.columns(3)
    tfactura.zserie = "" & DBGrid2.columns(4)
    tfactura.znumero = "" & DBGrid2.columns(5)

    If acu = "V" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Facturacion x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Modifica"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.acu = "V"
        tfactura.tipoclie = tipoclie
        'MsgBox "" & DBGrid1.Columns(2)
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1
        sql_cabeza

    End If

    If acu = "H" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Cotizacion x Ventas"
        cgusuario = "ccotizav"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dcotizav"
        tfactura.bandera = "Modifica"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.acu = "H"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "I" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Pedidos x Ventas"
        cgusuario = "cpedidov"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dpedidov"
        tfactura.bandera = "Modifica"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.acu = "I"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "T" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Guia Remision x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Modifica"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.acu = "T"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "E" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Nota Credito x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Modifica"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.acu = "E"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "R" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Orden Compra"
        cgusuario = "CORDENC"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DORDENC"
        tfactura.bandera = "Modifica"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.acu = "R"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "F" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Nota Debito x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"

        tfactura.acu = "F"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Modifica"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "S" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Guia Remision Entrada"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.tipoclie = tipoclie
        tfactura.acu = "S"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Modifica"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "C" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Factura de Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.acu = "C"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Modifica"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "N" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Nota Credito Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.acu = "N"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Modifica"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "O" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Nota debito de Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.acu = "O"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Modifica"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "Q" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Pedido Almacen"
        cgusuario = "CREQUISA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DREQUISA"
        tfactura.acu = "Q"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Modifica"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "Z" Then
        tfactura.Caption = "Traslado entre almacen de un mismo establecimiento"
        cgusuario = "CTRASLAD"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DTRASLAD"
        tfactura.acu = "Z"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Modifica"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    sql_cabeza

    MsgBox "Presione tecla para continuar..", 48, "Aviso"

    Exit Sub
cmd27_err:
    MsgBox "Seleccione un dato  ", 48, "Aviso"
    Exit Sub

End Sub

Sub pone_registro()

End Sub

Private Sub mit56232_Click()

    If Frame1.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    zooma = "Zomm"
    visualizar_zoom
    MsgBox "Presione tecla para continuar..", 48, "Aviso"
    'sql_cabeza
    Exit Sub

End Sub

Sub visualizar_zoom()

    Dim found As Integer

    On Error GoTo cmd278_err

    'If Frame4.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    'If "" & Data2.Recordset.Fields("estado") <> "0" Then
    '   MsgBox "Estado debe estar =0", 48, "Aviso"
    '   Exit Sub
    'End If
    found = copiar_temporal()

    If found = 0 Then
        MsgBox "Error al copiar archivo temporal", 24, "Aviso"
        Exit Sub

    End If

    tfactura.zlocal = "" & DBGrid2.columns("local")
    tfactura.ztipo = "" & DBGrid2.columns("tipo")
    tfactura.zserie = "" & DBGrid2.columns("serie")
    tfactura.znumero = "" & DBGrid2.columns("numero")

    tfactura.Label2 = "CodClie"
    tfactura.Caption = "Facturacion x Ventas"
    cgusuario = "FACTURA"
    dgusuario = "_d" & gusuario
    fgusuario = "_f" & gusuario
    dgusuariog = "DETALLE"
    tfactura.bandera = "Ver"
    tfactura.cmdAddEntry.Enabled = False
    tfactura.dnu834.Enabled = False
    tfactura.acu = "" & rexplorap.Fields("acu")
    tfactura.tipoclie = "" & rexplorap.Fields("tipoclie")
    tfactura.Show 1
    Exit Sub

    If acu = "Z" Then
        'tfactura.Label14.Visible = True
        tfactura.Label38.Visible = True
        'tfactura.localf.Visible = True
        'inicio 10/02/2018 pll
        'tfactura.bodegaf.Visible = True
        'fin 10/02/2018 pll
   
        tfactura.Label2.Caption = "Cod.Int."

    End If

    If acu = "V" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Facturacion x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Ver"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.acu = "" & rexplorap.Fields("acu")
        tfactura.tipoclie = "" & rexplorap.Fields("tipoclie")
        tfactura.Show 1
        'sql_cabeza

    End If

    If acu = "H" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Cotizacion x Ventas"
        cgusuario = "ccotizav"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dcotizav"
        tfactura.bandera = "Ver"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.acu = "H"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "I" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Pedidos x Ventas"
        cgusuario = "cpedidov"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "dpedidov"
        tfactura.bandera = "Ver"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.acu = "I"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "T" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Guia Remision x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Ver"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.acu = "T"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "E" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Nota Credito x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.bandera = "Ver"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.acu = "E"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "R" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Orden Compra"
        cgusuario = "CORDENC"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DORDENC"
        tfactura.bandera = "Ver"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.acu = "R"
        tfactura.tipoclie = tipoclie
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "F" Then
        tfactura.Label2 = "CodClie"
        tfactura.Caption = "Nota Debito x Ventas"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"

        tfactura.acu = "F"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Ver"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "S" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Guia Remision Entrada"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.tipoclie = tipoclie
        tfactura.acu = "S"
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Ver"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "C" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Factura de Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.acu = "C"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Ver"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "N" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Nota Credito Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.acu = "N"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Ver"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "O" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Nota debito de Compra"
        cgusuario = "FACTURA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DETALLE"
        tfactura.acu = "O"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Ver"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "Q" Then
        tfactura.Label2 = "CodProv"
        tfactura.Caption = "Pedido Almacen"
        cgusuario = "CREQUISA"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DREQUISA"
        tfactura.acu = "Q"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Ver"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    If acu = "Z" Then
        tfactura.Caption = "Traslado entre almacen de un mismo establecimiento"
        cgusuario = "CTRASLAD"
        dgusuario = "_d" & gusuario
        fgusuario = "_f" & gusuario
        dgusuariog = "DTRASLAD"
        tfactura.acu = "Z"
        tfactura.tipoclie = tipoclie
        tfactura.cmdAddEntry.Enabled = False
        tfactura.dnu834.Enabled = False
        tfactura.bandera = "Ver"
        'tfactura.zlocal = "" & Data2.Recordset.Fields("local")
        'tfactura.ztipo = "" & Data2.Recordset.Fields("tipo")
        'tfactura.zserie = "" & Data2.Recordset.Fields("serie")
        'tfactura.znumero = "" & Data2.Recordset.Fields("numero")
        tfactura.Show 1

    End If

    Exit Sub
cmd278_err:
    MsgBox "Aviso en visualizar zoon " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub modi343_Click()

    On Error GoTo cmd117_err

    Dim found As Integer

    Dim buf   As String

    'If Frame4.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    buf = "" & DBGrid2.columns(1) 'Data2.Recordset.Fields("estado")

    If Trim("" & DBGrid2.columns(1)) <> "2" Then
        MsgBox "Debe encontrarse en estado 2 para desmarcar ", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    If Trim("" & DBGrid2.columns(0)) = "A" Then
        MsgBox "Modo atendido,no se puede modificar ", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    Frame2.Visible = True
    Frame2.Caption = "DESMARCA"
    clave = ""
    clave.SetFocus
    Exit Sub
cmd117_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Function verificar_recibo(buf As String, _
                          xlocal As String, _
                          xtipo As String, _
                          xserie As String, _
                          xnumero As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM " & buf & " where  local1='" & xlocal & "' and tipo1='" & xtipo & "' and serie1='" & xserie & "' and numero1='" & xnumero & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        verificar_recibo = 1

    End If

    mytablex.Close

End Function

Function desgraba_cuentac()

    '---------- validando si es cuenta corriente

    If valida_flag("" & DBGrid2.columns("acu")) = 2 Then   'compras
        cn.Execute ("delete from cuentap where local='" & DBGrid2.columns("local") & "' and tipo='" & DBGrid2.columns("tipo") & "' and serie='" & DBGrid2.columns("serie") & "' and numero='" & DBGrid2.columns("numero") & "'")

    End If

    If valida_flag("" & DBGrid2.columns("acu")) = 1 Then   'ventas
        cn.Execute ("delete from cuentac where local='" & DBGrid2.columns("local") & "' and tipo='" & DBGrid2.columns("tipo") & "' and serie='" & DBGrid2.columns("serie") & "' and numero='" & DBGrid2.columns("numero") & "'")

    End If
 
End Function

Sub menu_excell1()

    If acu = "V" Or acu = "T" Or acu = "E" Or acu = "F" Or acu = "S" Or acu = "C" Or acu = "N" Or acu = "O" Then
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"

    End If

    If acu = "H" Then
        cgusuario = "CCOTIZAV"
        dgusuariog = "DCOTIZAV"

    End If

    If acu = "I" Then
        cgusuario = "CPEDIDOV"
        dgusuariog = "DPEDIDOV"

    End If

    If acu = "R" Then
        cgusuario = "CORDENC"
        dgusuariog = "DORDENC"

    End If

    If acu = "Q" Then
        cgusuario = "CREQUISA"
        dgusuariog = "DREQUISA"

    End If

    If acu = "Z" Then
        cgusuario = "CTRASLAD"
        dgusuariog = "DTRASLAD"

    End If

    excel_paso1

End Sub

Sub menu_excell()

    If acu = "V" Or acu = "T" Or acu = "E" Or acu = "F" Or acu = "S" Or acu = "C" Or acu = "N" Or acu = "O" Then
        cgusuario = "FACTURA"
        dgusuariog = "DETALLE"

    End If

    If acu = "H" Then
        cgusuario = "CCOTIZAV"
        dgusuariog = "DCOTIZAV"

    End If

    If acu = "I" Then
        cgusuario = "CPEDIDOV"
        dgusuariog = "DPEDIDOV"

    End If

    If acu = "R" Then
        cgusuario = "CORDENC"
        dgusuariog = "DORDENC"

    End If

    If acu = "Q" Then
        cgusuario = "CREQUISA"
        dgusuariog = "DREQUISA"

    End If

    If acu = "Z" Then
        cgusuario = "CTRASLAD"
        dgusuariog = "DTRASLAD"

    End If

    excel_paso

End Sub

Sub excel_paso1()

    Dim sdx As String

    On Error GoTo cmd813_err

    sdx = "" & DBGrid2.columns("numero")
    conteo_excell1
    Exit Sub
cmd813_err:
    MsgBox "Elegir un dato ", 48, "Aviso"
    Exit Sub

End Sub

Sub excel_paso()

    Dim sdx As String

    On Error GoTo cmd81_err

    sdx = "" & DBGrid2.columns("numero")
    conteo_excell
    Exit Sub
cmd81_err:
    MsgBox "Elegir un dato ", 48, "Aviso"
    Exit Sub

End Sub

Sub conteo_excell1()

    Dim v, h As Integer

    Dim R            As Long

    Dim found        As Integer

    Dim I            As Integer

    Dim sdx          As Double

    Dim sdx1         As Double

    Dim sdx2         As Double

    Dim vprecios(11) As String

    Dim Heading(12)  As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd56124_err

    Heading(1) = "Codigo"
    Heading(2) = "Nombre"
    Heading(3) = "Local"
    Heading(4) = "Tipo"
    Heading(5) = "Serie"
    Heading(6) = "Numero"
    Heading(7) = "M"
    Heading(8) = "Fecha"
    Heading(9) = "Total"
    Heading(10) = ""
    Heading(11) = ""
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excel(11, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    v = 4
    h = 0
    sdx = 0
    sdx1 = 0
    sdx2 = 0

    For R = 0 To DBGrid2.ApproxCount - 1
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & DBGrid2.columns("Codigo")
        objExcel.ActiveSheet.Cells(v, h + 2) = "" & DBGrid2.columns("nombre")
        objExcel.ActiveSheet.Cells(v, h + 3) = "" & DBGrid2.columns("Local")
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & DBGrid2.columns("tipo")
        objExcel.ActiveSheet.Cells(v, h + 5) = "" & DBGrid2.columns("serie")
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & DBGrid2.columns("numero")
        objExcel.ActiveSheet.Cells(v, h + 7) = "" & DBGrid2.columns("M")
        objExcel.ActiveSheet.Cells(v, h + 8) = "" & DBGrid2.columns("fecha")
        objExcel.ActiveSheet.Cells(v, h + 9) = "" & DBGrid2.columns("total")

        'objExcel.ActiveSheet.Cells(v, h + 10) = "" & dbgrid2.Columns("impuesto")
        'objExcel.ActiveSheet.Cells(v, h + 11) = "" & dbgrid2.Columns("subtotal")
        If "" & DBGrid2.columns("M") = "S" Then
            sdx1 = sdx1 + Val("" & DBGrid2.columns("total"))

        End If

        If "" & DBGrid2.columns("M") = "D" Then
            sdx2 = sdx2 + Val("" & DBGrid2.columns("total"))

        End If

        v = v + 1

        If DBGrid2.Row < DBGrid2.ApproxCount - 1 Then
            DBGrid2.Row = DBGrid2.Row + 1
        Else
            Exit For

        End If

    Next R

    objExcel.ActiveSheet.Cells(v, h + 1) = ""
    objExcel.ActiveSheet.Cells(v, h + 2) = ""
    objExcel.ActiveSheet.Cells(v, h + 3) = ""
    objExcel.ActiveSheet.Cells(v, h + 4) = ""
    objExcel.ActiveSheet.Cells(v, h + 5) = ""
    objExcel.ActiveSheet.Cells(v, h + 6) = dicmoneda
    objExcel.ActiveSheet.Cells(v, h + 7) = Format(sdx1, "0.00")
    objExcel.ActiveSheet.Cells(v, h + 8) = "Dolar"
    objExcel.ActiveSheet.Cells(v, h + 9) = Format(sdx2, "0.00")

    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd56124_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub conteo_excell()

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim v, h As Long

    Dim found       As Integer

    Dim I           As Integer

    Dim R           As Long

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim sdx3        As Double

    Dim vprecios(7) As String

    Dim Heading(8)  As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd1561212_err

    'Data1.Refresh
   
    Heading(1) = "Producto"
    Heading(2) = "Descripcio"
    Heading(3) = "Und"
    Heading(4) = "Factor"
    Heading(5) = "cantidad"
    Heading(6) = "Precio"
    Heading(7) = "Total"
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excel(7, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    v = 5
    h = 1

    For R = 0 To DBGrid2.ApproxCount - 1

        sdx = 0
        sdx1 = 0

        objExcel.ActiveSheet.Cells(v, h + 1) = "Tipo:" & DBGrid2.columns("tipo") & " Serie:" & DBGrid2.columns("serie") & "Numero:" & DBGrid2.columns("numero")
        v = v + 1
        objExcel.ActiveSheet.Cells(v, h + 1) = "Cliente:" & DBGrid2.columns("codigo") & " Nombre:" & DBGrid2.columns("nombre") & " Vendedor:" & DBGrid2.columns("vendedor") & " Moneda:" & DBGrid2.columns("m")
        v = v + 1

        mytablex.Open "SELECT * FROM " & dgusuariog & " where  local='" & Trim("" & DBGrid2.columns("local")) & "' and tipo='" & Trim("" & DBGrid2.columns("tipo")) & "' and serie='" & Trim("" & DBGrid2.columns("serie")) & "' and numero='" & Trim("" & DBGrid2.columns("numero")) & "'", cn, adOpenKeyset, adLockOptimistic

        If mytablex.RecordCount > 0 Then  'si existe
            Do

                If mytablex.EOF Then Exit Do
                sdx = sdx + Val("" & mytablex.Fields("cantidad"))
                sdx1 = sdx1 + Val("" & mytablex.Fields("total"))
                objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("producto")
                objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytablex.Fields("descripcio")
                objExcel.ActiveSheet.Cells(v, h + 2) = "" & mytablex.Fields("unidad")
                objExcel.ActiveSheet.Cells(v, h + 3) = Val("" & mytablex.Fields("factor"))
                objExcel.ActiveSheet.Cells(v, h + 4) = Val("" & mytablex.Fields("cantidad"))
                objExcel.ActiveSheet.Cells(v, h + 5) = Val("" & mytablex.Fields("precio"))
                objExcel.ActiveSheet.Cells(v, h + 6) = Val("" & mytablex.Fields("total"))
            
                objExcel.ActiveSheet.Cells(v, h + 6) = Val("" & mytablex.Fields("total"))
            
                If mytabley.State = 1 Then mytabley.Close
                mytabley.Open "SELECT * FROM precios where  producto='" & mytablex.Fields("producto") & "' and local='01'", cn, adOpenKeyset, adLockOptimistic
                sdx2 = 0

                If mytabley.RecordCount > 0 Then  'si existe
                    sdx2 = Val("" & mytabley.Fields("pventa1")) / Val("" & mytabley.Fields("factor1"))

                End If

                mytabley.Close
                sdx2 = sdx2 * Val("" & mytablex.Fields("factor")) * Val("" & mytablex.Fields("cantidad"))
                sdx3 = (sdx2 - Val("" & mytablex.Fields("total"))) * 100 / Val("" & mytablex.Fields("total"))
                objExcel.ActiveSheet.Cells(v, h + 7) = "'" & Format(sdx3, "0.00") & "%"
                v = v + 1
                mytablex.MoveNext
            Loop

        End If

        mytablex.Close

        objExcel.ActiveSheet.Cells(v, h) = ""
        objExcel.ActiveSheet.Cells(v, h + 1) = ""
        objExcel.ActiveSheet.Cells(v, h + 2) = ""
        objExcel.ActiveSheet.Cells(v, h + 3) = ""
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & sdx
        objExcel.ActiveSheet.Cells(v, h + 5) = ""
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & DBGrid2.columns("total")
        v = v + 1

        If DBGrid2.Row < DBGrid2.ApproxCount - 1 Then
            DBGrid2.Row = DBGrid2.Row + 1
        Else
            Exit For

        End If

    Next R

    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd1561212_err:
    MsgBox "Error en exportacion excell " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub conteo_excell_uno()

    Dim mytablex As New ADODB.Recordset

    Dim v, h As Integer

    Dim found       As Integer

    Dim I           As Integer

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim vprecios(7) As String

    Dim Heading(8)  As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd561212_err

    'Data1.Refresh
   
    Heading(1) = "Producto"
    Heading(2) = "Descripcio"
    Heading(3) = "Und"
    Heading(4) = "Factor"
    Heading(5) = "cantidad"
    Heading(6) = "Precio"
    Heading(7) = "Total"
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excel(7, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    objExcel.ActiveSheet.Cells(1, 1) = "Tipo:" & DBGrid2.columns("tipo") & " Serie:" & DBGrid2.columns("serie") & "Numero:" & DBGrid2.columns("numero")
    objExcel.ActiveSheet.Cells(2, 1) = "Cliente:" & DBGrid2.columns("codigo") & " Nombre:" & DBGrid2.columns("nombre") & " Vendedor:" & DBGrid2.columns("vendedor") & " Moneda:" & DBGrid2.columns("m")
    
    v = 5
    h = 1
    sdx = 0
    sdx1 = 0

    mytablex.Open "SELECT * FROM " & dgusuariog & " where  local='" & Trim("" & DBGrid2.columns("local")) & "' and tipo='" & Trim("" & DBGrid2.columns("tipo")) & "' and serie='" & Trim("" & DBGrid2.columns("serie")) & "' and numero='" & Trim("" & DBGrid2.columns("numero")) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        Do

            If mytablex.EOF Then Exit Do
            sdx = sdx + Val("" & mytablex.Fields("cantidad"))
            sdx1 = sdx1 + Val("" & mytablex.Fields("total"))
            objExcel.ActiveSheet.Cells(v, h) = "'" & mytablex.Fields("producto")
            objExcel.ActiveSheet.Cells(v, h + 1) = "" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "" & mytablex.Fields("unidad")
            objExcel.ActiveSheet.Cells(v, h + 3) = Val("" & mytablex.Fields("factor"))
            objExcel.ActiveSheet.Cells(v, h + 4) = Val("" & mytablex.Fields("cantidad"))
            objExcel.ActiveSheet.Cells(v, h + 5) = Val("" & mytablex.Fields("precio"))
            objExcel.ActiveSheet.Cells(v, h + 6) = Val("" & mytablex.Fields("total"))
            v = v + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    mytablex.Open "SELECT * FROM " & cgusuario & " where  local='" & Trim("" & DBGrid2.columns("local")) & "' and tipo='" & Trim("" & DBGrid2.columns("tipo")) & "' and serie='" & Trim("" & DBGrid2.columns("serie")) & "' and numero='" & Trim("" & DBGrid2.columns("numero")) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        objExcel.ActiveSheet.Cells(v, h) = ""
        objExcel.ActiveSheet.Cells(v, h + 1) = ""
        objExcel.ActiveSheet.Cells(v, h + 2) = ""
        objExcel.ActiveSheet.Cells(v, h + 3) = ""
        objExcel.ActiveSheet.Cells(v, h + 4) = "" & sdx
        objExcel.ActiveSheet.Cells(v, h + 5) = ""
        objExcel.ActiveSheet.Cells(v, h + 6) = "" & mytablex.Fields("total")
        v = v + 1

    End If

    mytablex.Close
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    Exit Sub
cmd561212_err:
    MsgBox "Error en exportacion excell" + error$, 48, "Aviso"
    Exit Sub

End Sub

'AQUI VAMOS A GENERAR LE documento automaticamente
Function valida_clave(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM vendedor where  clave='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        valida_clave = 1

    End If

    mytablex.Close

End Function

Function ve_descarga(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe

        Select Case "" & mytablex.Fields("tipodoc")

            Case "A", "B", "C", "D", "G", "E", "F", "J", "K", "L", "M", "P", "N", "O", "S", "T"
                ve_descarga = 1

        End Select

    End If

    mytablex.Close

End Function

