VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tptovtaa 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema Orion 5.0"
   ClientHeight    =   10350
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   10350
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox cadena_balanza 
      Height          =   495
      Left            =   960
      TabIndex        =   259
      Top             =   11160
      Width           =   1935
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H0080FF80&
      Caption         =   "Personal"
      Height          =   9135
      Left            =   2280
      TabIndex        =   253
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
      Begin MSDataGridLib.DataGrid table6 
         Height          =   8415
         Left            =   720
         TabIndex        =   254
         Top             =   360
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   14843
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   3
         RowHeight       =   29
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "Nombre"
            Caption         =   "Nombre"
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
         BeginProperty Column01 
            DataField       =   "Codigo"
            Caption         =   "Codigo"
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
            MarqueeStyle    =   3
            BeginProperty Column00 
               ColumnWidth     =   4710.047
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<SUBE>"
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
         Left            =   6960
         TabIndex        =   258
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<BAJA>"
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
         Left            =   6960
         TabIndex        =   257
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label53 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<SELEC>"
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
         Left            =   6960
         TabIndex        =   256
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<CLOSE>"
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
         Left            =   6960
         TabIndex        =   255
         Top             =   2880
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Lista Precios y Saldos "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   10335
      Left            =   0
      TabIndex        =   244
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton Command4 
         Caption         =   "Selecciona"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   248
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox dvendedor 
         DataField       =   "vendedor"
         DataSource      =   "Data2"
         Height          =   375
         Left            =   8520
         MaxLength       =   11
         TabIndex        =   247
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox dcvendedor 
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
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   246
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   245
         Top             =   3480
         Width           =   1455
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         Height          =   4335
         Left            =   120
         OleObjectBlob   =   "tptovtaa.frx":0000
         TabIndex        =   249
         Top             =   360
         Width           =   6735
      End
      Begin MSDataGridLib.DataGrid dbgrid7 
         Height          =   2535
         Left            =   120
         TabIndex        =   250
         Top             =   5280
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   0   'False
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
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6960
         TabIndex        =   252
         Top             =   360
         Width           =   1575
      End
      Begin VB.Image foto 
         BorderStyle     =   1  'Fixed Single
         Height          =   3015
         Left            =   6960
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   4335
      End
      Begin VB.Label descorto 
         BackColor       =   &H00C0FFC0&
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
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   251
         Top             =   4800
         Width           =   6735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   10215
      Left            =   0
      TabIndex        =   235
      Top             =   0
      Visible         =   0   'False
      Width           =   15375
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   238
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
         Left            =   5640
         TabIndex        =   237
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   236
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbgrid6 
         Height          =   8535
         Left            =   120
         TabIndex        =   239
         Top             =   840
         Visible         =   0   'False
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   15055
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   19
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
            Size            =   9.75
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
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7935
         Left            =   120
         TabIndex        =   240
         Top             =   840
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   13996
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
      Begin VB.Label label56 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   0
         TabIndex        =   243
         Top             =   9480
         Width           =   14505
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ACEPTAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   13080
         TabIndex        =   242
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CLOSE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   13080
         TabIndex        =   241
         Top             =   2040
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Delivery"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   10335
      Left            =   0
      TabIndex        =   213
      Top             =   0
      Visible         =   0   'False
      Width           =   15255
      Begin VB.ComboBox coclasifica 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   220
         Top             =   3000
         Width           =   5535
      End
      Begin VB.TextBox referencia 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2640
         MaxLength       =   60
         TabIndex        =   219
         Top             =   1920
         Width           =   8295
      End
      Begin VB.TextBox ddireccion 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2640
         MaxLength       =   60
         TabIndex        =   218
         Top             =   1440
         Width           =   8295
      End
      Begin VB.TextBox dnombre 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2640
         MaxLength       =   60
         TabIndex        =   217
         Top             =   960
         Width           =   8295
      End
      Begin VB.TextBox dcodigo 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5640
         MaxLength       =   11
         TabIndex        =   216
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox telefono 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   525
         Left            =   2640
         MaxLength       =   11
         TabIndex        =   215
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox fechanac 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   214
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label felizc 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   4440
         TabIndex        =   234
         Top             =   2400
         Width           =   6495
      End
      Begin VB.Label clasificacion 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   233
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11520
         TabIndex        =   232
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Modifica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11520
         TabIndex        =   231
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Nacimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   230
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Referencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   229
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   228
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   227
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label command11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11520
         TabIndex        =   226
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label command12 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11520
         TabIndex        =   225
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label command10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11520
         TabIndex        =   224
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   223
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefono"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   222
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clasificacion Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   221
         Top             =   2880
         Width           =   2415
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "CONGELA PEDIDOS INGRESADOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3855
      Left            =   3960
      TabIndex        =   208
      Top             =   2400
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox xcongelax 
         Height          =   615
         Left            =   240
         MaxLength       =   12
         TabIndex        =   211
         TabStop         =   0   'False
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton cmdCancelar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tptovtaa.frx":1063
         Style           =   1  'Graphical
         TabIndex        =   210
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   1575
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
         Height          =   1095
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tptovtaa.frx":1811
         Style           =   1  'Graphical
         TabIndex        =   209
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Digite un Nombre "
         Height          =   375
         Left            =   240
         TabIndex        =   212
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ingreso de Tipos de Documentos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4935
      Left            =   120
      TabIndex        =   185
      Top             =   1920
      Visible         =   0   'False
      Width           =   14295
      Begin VB.TextBox xtipo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         MaxLength       =   1
         TabIndex        =   196
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox xvendedor 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2160
         MaxLength       =   11
         PasswordChar    =   "*"
         TabIndex        =   195
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox xruc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         MaxLength       =   11
         TabIndex        =   194
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox xnombre 
         Height          =   495
         Left            =   2160
         MaxLength       =   60
         TabIndex        =   193
         Top             =   1680
         Width           =   5415
      End
      Begin VB.TextBox xdireccion 
         Height          =   495
         Left            =   2160
         MaxLength       =   60
         TabIndex        =   192
         Top             =   2160
         Width           =   5415
      End
      Begin VB.TextBox xdistrito 
         Height          =   495
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   191
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox xnumero 
         Enabled         =   0   'False
         Height          =   495
         Left            =   2160
         MaxLength       =   11
         TabIndex        =   190
         Top             =   3600
         Width           =   2175
      End
      Begin VB.TextBox xserie 
         Enabled         =   0   'False
         Height          =   495
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   189
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox sentido 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         MaxLength       =   1
         TabIndex        =   188
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   10200
         Picture         =   "tptovtaa.frx":1FBF
         Style           =   1  'Graphical
         TabIndex        =   187
         Top             =   240
         Width           =   1470
      End
      Begin VB.CommandButton Command14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   10200
         Picture         =   "tptovtaa.frx":2889
         Style           =   1  'Graphical
         TabIndex        =   186
         ToolTipText     =   "Imprimir todo"
         Top             =   1200
         Width           =   1470
      End
      Begin VB.Label nbxtipo 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   2
         Left            =   5760
         TabIndex        =   207
         Top             =   240
         Width           =   735
      End
      Begin VB.Label nbxtipo 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   1
         Left            =   5040
         TabIndex        =   206
         Top             =   240
         Width           =   735
      End
      Begin VB.Label nbxtipo 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Index           =   0
         Left            =   4320
         TabIndex        =   205
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   204
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   203
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo                                                  Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   202
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion                                     Glosa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   201
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label48 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie                                              Numero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   200
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label ntipox 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   6480
         TabIndex        =   199
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label nvendedorx 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   6480
         TabIndex        =   198
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label ordentrabajo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4440
         TabIndex        =   197
         Top             =   3600
         Width           =   105
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Entrega"
      Height          =   5055
      Left            =   120
      TabIndex        =   150
      Top             =   1800
      Visible         =   0   'False
      Width           =   13215
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Height          =   4815
         Left            =   5760
         TabIndex        =   184
         Top             =   600
         Width           =   7695
      End
      Begin VB.TextBox tcampo6 
         Height          =   375
         Left            =   7080
         MaxLength       =   5
         TabIndex        =   171
         Top             =   2760
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox tcampo5 
         Height          =   375
         Left            =   7080
         MaxLength       =   5
         TabIndex        =   170
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox tcampo4 
         Height          =   375
         Left            =   7080
         MaxLength       =   30
         TabIndex        =   169
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox tcampo3 
         Height          =   375
         Left            =   7080
         MaxLength       =   30
         TabIndex        =   168
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox tcampo2 
         Height          =   375
         Left            =   7080
         MaxLength       =   60
         TabIndex        =   167
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox tcampo1 
         Height          =   375
         Left            =   7080
         MaxLength       =   11
         TabIndex        =   166
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton kcobra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton kcobra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   164
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton kcobra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton kcobra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   162
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton kcobra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   161
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton kcobra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   5
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton kcobra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton kcobra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   7
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton kcobra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   157
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton kcobra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   9
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton kcobra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   10
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   155
         Top             =   2520
         Width           =   975
      End
      Begin VB.CommandButton kcobra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   11
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox RGPAGO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MaxLength       =   10
         TabIndex        =   153
         Top             =   360
         Width           =   4095
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1065
         Left            =   120
         Picture         =   "tptovtaa.frx":3153
         Style           =   1  'Graphical
         TabIndex        =   152
         ToolTipText     =   "Imprimir todo"
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "tptovtaa.frx":3A1D
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   3360
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.Label totpedido 
         BackColor       =   &H00C0FFC0&
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
         Left            =   7080
         TabIndex        =   183
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label acufp 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5640
         TabIndex        =   182
         Top             =   3240
         Width           =   105
      End
      Begin VB.Label descripcio6 
         BackColor       =   &H00000000&
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
         TabIndex        =   181
         Top             =   2760
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label saldoabo 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   8880
         TabIndex        =   180
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label fpmoneda 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9960
         TabIndex        =   179
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label fpago 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   9960
         TabIndex        =   178
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label descripcio5 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   177
         Top             =   2400
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label descripcio4 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   176
         Top             =   2040
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label descripcio3 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   175
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label descripcio2 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   174
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label descripcio1 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   173
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label limite_credito 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   8880
         TabIndex        =   172
         Top             =   1680
         Width           =   1695
      End
   End
   Begin VB.Frame Framefp 
      BackColor       =   &H00FFFFC0&
      Caption         =   "COBRANZAS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   10335
      Left            =   0
      TabIndex        =   128
      Top             =   0
      Visible         =   0   'False
      Width           =   15255
      Begin VB.CommandButton COMMAND6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   12120
         Picture         =   "tptovtaa.frx":41CB
         Style           =   1  'Graphical
         TabIndex        =   129
         ToolTipText     =   "Imprimir todo"
         Top             =   2880
         Width           =   1215
      End
      Begin MSDBGrid.DBGrid DBGrid9 
         Bindings        =   "tptovtaa.frx":4A95
         Height          =   4695
         Left            =   5040
         OleObjectBlob   =   "tptovtaa.frx":4AA9
         TabIndex        =   130
         Top             =   2280
         Width           =   6975
      End
      Begin MSDataGridLib.DataGrid dbgrid10 
         Height          =   6375
         Left            =   120
         TabIndex        =   131
         Top             =   2280
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   11245
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   29
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
            Size            =   12
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
            DataField       =   "Fpago"
            Caption         =   "Fpago"
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
            DataField       =   "Moneda"
            Caption         =   "Moneda"
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
            DataField       =   "Tipo"
            Caption         =   "Tipo"
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
            DataField       =   "Dias"
            Caption         =   "Dias"
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
               ColumnWidth     =   5325.166
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3915.213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   464.882
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   494.929
            EndProperty
         EndProperty
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T/C"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5160
         TabIndex        =   149
         Top             =   840
         Width           =   615
      End
      Begin VB.Label paridadfp 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5760
         TabIndex        =   148
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "COMO PAGA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   5040
         TabIndex        =   147
         Top             =   1920
         Width           =   6975
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Formas de Pago"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   120
         TabIndex        =   146
         Top             =   1920
         Width           =   4815
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1695
         Left            =   5040
         TabIndex        =   145
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FALTA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1695
         Left            =   5040
         TabIndex        =   144
         Top             =   6960
         Width           =   2055
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   7200
         TabIndex        =   143
         Top             =   240
         Width           =   495
      End
      Begin VB.Label ttxtotals 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   7680
         TabIndex        =   142
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  US$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   7200
         TabIndex        =   141
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label ttxtotald 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   855
         Left            =   7680
         TabIndex        =   140
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   855
         Left            =   7080
         TabIndex        =   139
         Top             =   6960
         Width           =   495
      End
      Begin VB.Label stxtotals 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   855
         Left            =   7560
         TabIndex        =   138
         Top             =   6960
         Width           =   4455
      End
      Begin VB.Label Label52 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  US$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   7080
         TabIndex        =   137
         Top             =   7800
         Width           =   495
      End
      Begin VB.Label stxtotald 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   855
         Left            =   7560
         TabIndex        =   136
         Top             =   7800
         Width           =   4455
      End
      Begin VB.Label Label70 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "              <Acepta>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   135
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Borra"
         Height          =   615
         Left            =   12120
         TabIndex        =   134
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label72 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "             <Baja>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         TabIndex        =   133
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label73 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                    <Sube>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1200
         TabIndex        =   132
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.TextBox ACUENTA 
      Height          =   495
      Left            =   9960
      MaxLength       =   10
      TabIndex        =   127
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Height          =   4455
      Left            =   0
      TabIndex        =   124
      Top             =   0
      Width           =   3735
      Begin VB.PictureBox IMAGE11 
         Height          =   4215
         Left            =   120
         ScaleHeight     =   4155
         ScaleWidth      =   3435
         TabIndex        =   125
         Top             =   120
         Width           =   3495
      End
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "tptovtaa.frx":597C
      Height          =   3615
      Left            =   3720
      OleObjectBlob   =   "tptovtaa.frx":5990
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   23
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Control Personal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   22
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Efectivo Boleta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   21
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   120
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cierre Caja"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   20
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Egreso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   19
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   118
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ingreso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   18
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   117
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Servicio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   17
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   116
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Combi nacion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   16
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   115
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Comen tario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   15
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "GRABA COMAN DAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   14
      Left            =   16560
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Congela"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   13
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   112
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Des congela"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   12
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Abre Gaveta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   11
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Borra Linea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Anula Venta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Copia Ventas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cuadre Parcial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Limpia Pedido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Descto Pedido Actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Delivery"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cobrar Mesa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   15720
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cuenta Separada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   17400
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Pago Cash"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton xopciones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Auto Servicio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   12720
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Label59 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NORMAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   4560
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid table2 
      Height          =   4215
      Left            =   15360
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   2280
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   24
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
         Size            =   12
         Charset         =   0
         Weight          =   700
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
         MarqueeStyle    =   4
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      BackColor       =   &H0000FF00&
      Height          =   255
      Left            =   15360
      TabIndex        =   86
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox hkproducto 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   4200
      MaxLength       =   15
      TabIndex        =   84
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   15360
      Top             =   960
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   6120
      Width           =   1935
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   6120
      Width           =   1815
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   6
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   7
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   9
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   10
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   11
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   6960
      Width           =   1815
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   12
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   13
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   14
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   15
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   7800
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   16
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   17
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   18
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   8640
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   19
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   8640
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   20
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   8640
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   21
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   8640
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   22
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   8640
      Width           =   1935
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   23
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   24
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   9480
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   25
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   9480
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   26
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   9480
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   27
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   9480
      Width           =   2055
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   28
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CommandButton zproducto 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   29
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   9480
      Width           =   1815
   End
   Begin VB.TextBox codigo 
      Height          =   195
      Left            =   4680
      TabIndex        =   38
      Top             =   12120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   17
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   16
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   15
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   14
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   13
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton zfamilia 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4920
      Width           =   1335
   End
   Begin VB.ComboBox crucefa 
      Height          =   315
      Left            =   16080
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   13560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   11400
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   16440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   11040
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   12360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   12480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data9 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   11640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data10 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   11640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   12360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   11400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   12360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   12120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   12360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   11760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   15840
      Top             =   10320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      OutBufferSize   =   1024
      RThreshold      =   13
      RTSEnable       =   -1  'True
      SThreshold      =   2
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRODUCTO"
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
      Height          =   735
      Left            =   12840
      TabIndex        =   126
      Top             =   9600
      Width           =   1095
   End
   Begin VB.Label stkminimo 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   11880
      TabIndex        =   123
      Top             =   9000
      Width           =   3375
   End
   Begin VB.Label mesero 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   97
      Top             =   0
      Width           =   1455
   End
   Begin VB.Image fotoimagen 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   11880
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Label Label69 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ACuenta"
      Enabled         =   0   'False
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
      Height          =   495
      Left            =   8880
      TabIndex        =   96
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label68 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Todos"
      Height          =   255
      Left            =   17760
      TabIndex        =   95
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label67 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "03"
      Height          =   255
      Left            =   16920
      TabIndex        =   94
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label61 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "02"
      Height          =   255
      Left            =   16080
      TabIndex        =   93
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label60 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "01"
      Height          =   255
      Left            =   15360
      TabIndex        =   92
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label totcoma 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2880
      TabIndex        =   91
      Top             =   11160
      Width           =   1575
   End
   Begin VB.Label mesa 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2280
      TabIndex        =   90
      Top             =   12120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label salon 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   89
      Top             =   12120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image10 
      BorderStyle     =   1  'Fixed Single
      Height          =   945
      Left            =   15360
      Picture         =   "tptovtaa.frx":C26F
      Stretch         =   -1  'True
      Top             =   8640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label comanda 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Height          =   195
      Left            =   15360
      TabIndex        =   88
      Top             =   1440
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   11880
      Picture         =   "tptovtaa.frx":C6C5
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   960
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   13920
      Picture         =   "tptovtaa.frx":E297
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   1080
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   15360
      Picture         =   "tptovtaa.frx":1023D
      Stretch         =   -1  'True
      Top             =   9600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   15360
      Picture         =   "tptovtaa.frx":121E3
      Stretch         =   -1  'True
      Top             =   6720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   15360
      Picture         =   "tptovtaa.frx":12540
      Stretch         =   -1  'True
      Top             =   7680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label hknumero 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   11880
      TabIndex        =   85
      Top             =   4440
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   83
      Top             =   11160
      Width           =   975
   End
   Begin VB.Label zznumero 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   14040
      TabIndex        =   82
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label znumero 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   12
      Left            =   11880
      TabIndex        =   81
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label znumero 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   11
      Left            =   12960
      TabIndex        =   80
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label znumero 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CR"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   10
      Left            =   14040
      TabIndex        =   79
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label znumero 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   9
      Left            =   11880
      TabIndex        =   78
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Label znumero 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   8
      Left            =   14040
      TabIndex        =   77
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label znumero 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   7
      Left            =   12960
      TabIndex        =   76
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label znumero 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   6
      Left            =   11880
      TabIndex        =   75
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label znumero 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   5
      Left            =   14040
      TabIndex        =   74
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label znumero 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   12960
      TabIndex        =   73
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label znumero 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   11880
      TabIndex        =   72
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label znumero 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   14040
      TabIndex        =   71
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label znumero 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   12960
      TabIndex        =   70
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label znumero 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   11880
      TabIndex        =   69
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "    FAMILIA"
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
      Height          =   495
      Left            =   6960
      TabIndex        =   19
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   8040
      Picture         =   "tptovtaa.frx":128E1
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   840
   End
   Begin VB.Image Image6 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6120
      Picture         =   "tptovtaa.frx":14887
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   840
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   11880
      Picture         =   "tptovtaa.frx":16459
      Stretch         =   -1  'True
      Top             =   720
      Width           =   840
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   11880
      Picture         =   "tptovtaa.frx":167FA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   840
   End
   Begin VB.Label tiposervicio1 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   18
      Top             =   12120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label local1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FF80&
      Height          =   195
      Left            =   15360
      TabIndex        =   17
      Top             =   1680
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label rtxtotald 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   6600
      TabIndex        =   15
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "     S/."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   9000
      TabIndex        =   14
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "      US$."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   6120
      TabIndex        =   13
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label CAMPO2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   12120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label nombre 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   12120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label campo3 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   10
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label moneda 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10560
      TabIndex        =   9
      Top             =   0
      Width           =   375
   End
   Begin VB.Label paridad 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10920
      TabIndex        =   8
      Top             =   0
      Width           =   975
   End
   Begin VB.Label fechasis 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   7
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label horasis 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   6
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label turno 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.Label cajero 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label caja 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label rtxtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   9360
      TabIndex        =   2
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label ntcant 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin VB.Menu menju232 
      Caption         =   "&Menu"
      Begin VB.Menu dju523a 
         Caption         =   "&1.Facturacion Mensual"
      End
      Begin VB.Menu dcrt6622 
         Caption         =   "&2.Carga Venta Anterior en Pedido Actual"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu dhyori83 
         Caption         =   "&9.Cargar Proformas terminales"
         Shortcut        =   {F6}
      End
      Begin VB.Menu dj78232 
         Caption         =   "&A.CargaPedidos-Ordenes Trabajo"
      End
      Begin VB.Menu dk8923 
         Caption         =   "&B.Cargar Cotizaciones "
      End
      Begin VB.Menu djk78232 
         Caption         =   "&C.Modificar Pedido Reposicion"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu d892323 
         Caption         =   "&D.Cuadre Rapido"
         Shortcut        =   ^I
      End
      Begin VB.Menu fdk9235 
         Caption         =   "&G.Anulacion Otra Fecha"
      End
      Begin VB.Menu dfk992325 
         Caption         =   "&H.Copia Otra Fecha"
      End
   End
   Begin VB.Menu cuj6721 
      Caption         =   "&Cuadres"
      Begin VB.Menu dcupar1 
         Caption         =   "&1.Parcial - Totales de Venta"
         Shortcut        =   ^T
      End
      Begin VB.Menu hundv1 
         Caption         =   "&2.Parcial - Unidades Vendidas"
         Shortcut        =   ^Q
      End
      Begin VB.Menu pado8911 
         Caption         =   "&3.Parcial - Documentos Emitidos "
      End
      Begin VB.Menu d8do82 
         Caption         =   "&4.Parcial - Productos Vs Documentos"
      End
      Begin VB.Menu forma671 
         Caption         =   "&5.Parcial - Formas de Pago"
      End
      Begin VB.Menu eju78se 
         Caption         =   "&6.Ingreso/Egreso/Seccion"
      End
   End
   Begin VB.Menu fk88332 
      Caption         =   "&Reportes"
   End
   Begin VB.Menu losao94 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tptovtaa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'CUANDO ES A CUENTA SE GENERA UN PEDIDO
'EL TIPO DOCUMENTO PEDIDO DEBE ESTAR EN TIPO DOCUMENTO
'Y GRABAR EN EL PEDIDO CUANDO QUEDA SALDO Y CUANTO FUE DADO A CUENTA


Dim cadena As String


Dim octipo As String
Dim ocserie As String
Dim ocnumero As String


Dim cuenta_separa As String
Dim mytablefpago As New ADODB.Recordset

Dim mfamcod(15000) As String
Dim wfamcod(15000) As String
Dim wwfamcod(30) As String

Dim mfampag As Integer
Dim mfamtop As Integer


Dim mcobcod(30) As String
Dim wcobcod(30) As String
Dim wwcobcod(30) As String

Dim mcobpag As Integer
Dim mcobtop As Integer



Dim trdescuento As String

Dim mprodcod(15000) As String
Dim wprodcod(15000) As String
Dim wwprodcod(30) As String
Dim mprodpag As Integer
Dim mprodtop As Integer
Dim acu As String
Dim cmytablex As New ADODB.Recordset   'comandas maneja
Dim rcconsulta As New ADODB.Recordset



Dim saldo As String
Dim pedido As String
'Dim acuenta As String
Dim tdetra As String
Dim tpeaje As String
Dim ndetraccion As String
Dim xxacu As String
Dim swprecio As Integer
Dim bk2 As Variant
Dim xproducto As String
Dim exisdev As Integer
Private Type campo_precio
    unidad As String
    factor As String
    precio As String
    costo As String
    margen As String
    stock As String
End Type
'------- globales de proformas
'Dim trdescuento As Double 'descuento global automatico
Dim nrodecimal As String
Dim tivap As Double
Dim tisc As Double
Dim txtotald As String
Dim txtotal As String
Dim cprotipo As String
Dim cproven As String
Dim cprocod As String

Dim InBuff As String
Dim xptipo As String
Dim xpserie As String
Dim xpnumero As String
Dim campo_precios(50) As campo_precio
Dim nrolineas As Integer
Dim tiposervicio As String
Dim flag_servicio As String
Dim flag_carga As String
Dim c1 As String
Dim c2 As String
Dim c3 As String
Dim c4 As String
Dim c5 As String
Dim c6 As String
Dim c7 As String
Dim c8 As String
Dim c9 As String
Dim gravado As String
Dim control_flujo As Integer
Dim protipo As String
Dim proserie As String
Dim pronumero As String
Dim tximpuesto As String
Dim xestado As String
Dim txdescuento As String
Dim txneto As String
Dim txsubtotal As String

Dim petipo As String
Dim peserie As String
Dim penumero As String
Dim flage As String
Dim dbvarios As New ADODB.Recordset
Private Sub cmdDelete_Click()

End Sub

Private Sub cmdAddEntry_Click()

End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdSort_Click()

End Sub

Private Sub cmdHelp_Click()

End Sub


Private Sub acuenta_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Not IsNumeric(acuenta) Then
   acuenta = ""
End If
found = sumar_detalle()
               DBGrid2.Col = 0
               DBGrid2.Row = DBGrid2.VisibleRows - 1
               DBGrid2.SetFocus
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   'If opcion1 = "1" Then
   '   losao94_Click
   'End If
   losao94_Click
   Exit Sub
End If
Command1_Click

End Sub

Private Sub cajero_Click()
dki3432_Click

End Sub

Private Sub cmdCancelar_Click()
Frame9.Visible = False
DBGrid2.SetFocus
End Sub

Private Sub cmdExit_Click()
End Sub

Private Sub cmdGrabar_Click()
Dim cad As String
Dim mytablex As New ADODB.Recordset
Dim mytabley As Table
Dim sdx As Double
Dim rs
Dim i As Integer
Dim xcongela As String
Dim sw As Integer
If Len(xcongelax) = 0 Then
   xcongelax.SetFocus
   Exit Sub
End If
If Frame9.Caption = "PEDIDO PARA REPONER" Then
   pedido_reposicion
   Label14_Click
   cmdCancelar_Click
   Exit Sub
End If
sdx = Val("" & mytable11.Fields("congela")) + 1
xcongela = "" & sdx
denuevo1:
If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "SELECT * FROM congelac where numero='" & xcongela & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      sdx = Val(xcongela) + 1
      xcongela = "" & sdx
      GoTo denuevo1
   End If
   mytable11.Close
   cad = "UPDATE parameca SET "
   cad = cad & "congela = '" & Trim(xcongela) & "'"
   cad = cad & " WHERE  caja='" & Trim(caja) & "'"
   cn.Execute (cad)
   mytable11.Open "SELECT * FROM parameca where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic
   cad = "INSERT INTO congelac (nombre,numero,fecha,moneda,paridad,dias,bodega,caja,turno,usuario,total) VALUES('" & Trim(xcongelax) & "','"
   cad = cad & Trim(xcongela) & "','"
   cad = cad & Format(dia, "YYYYMMDD") & "','"
   cad = cad & Trim("" & mytable11.Fields("moneda")) & "',"
   cad = cad & Val(paridad) & ","
   cad = cad & Val("1") & ",'"
   cad = cad & Trim("" & mytable11.Fields("bodega")) & "','"
   cad = cad & Trim(caja) & "','"
   cad = cad & Trim(turno) & "','"
   cad = cad & Trim(cajero) & "',"
   cad = cad & Val(txtotal) & ")"
   cn.Execute (cad)
'---ahora grabano detalle
cn.Execute ("DELETE   FROM congelad WHERE numero='" & xnumero & "'")
Data2.refresh
Do
If Data2.Recordset.EOF Then Exit Do
   cad = "INSERT INTO congelad VALUES('" & Trim("" & Data2.Recordset.Fields("tipo")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("serie")) & "','"
   cad = cad & Trim("" & xcongela) & "','"
   cad = cad & Trim("C") & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("codigo")) & "','"
   cad = cad & Trim("" & acu) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("acu1")) & "','"
   cad = cad & Format(dia, "YYYYMMDD") & "','"
   cad = cad & Trim("" & mytable11.Fields("moneda")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("producto")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("descripcio")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("unidad")) & "',"
   cad = cad & Val("" & Data2.Recordset.Fields("factor")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("cantidad")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("precio")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("igv")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("neto")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("descuento")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("subtotal")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("impuesto")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("total")) & ",'"
   cad = cad & Trim("2") & "','"
   cad = cad & Trim("" & cajero) & "','"
   cad = cad & Trim("" & Format(Now, "YYYYMMDD")) & "','"
   cad = cad & Trim("" & Format(Now, "hh:mm:ss")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("vendedor")) & "','"
   cad = cad & Trim("" & mytable11.Fields("bodega")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("bodegaf")) & "',"
   cad = cad & Val("" & Data2.Recordset.Fields("deslipo")) & ",'"
   cad = cad & Trim("" & Data2.Recordset.Fields("flage")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("linea")) & "',"
   cad = cad & Val("" & Data2.Recordset.Fields("t1")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("t2")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("t3")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("t4")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("t5")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("t6")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("t7")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("t8")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("t9")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("t10")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("t11")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("t12")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("t13")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("t14")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("t15")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("t16")) & ",'"
   cad = cad & Trim("" & Data2.Recordset.Fields("l1")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("l2")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("l3")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("l4")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("local")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("proveedorP")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("observa1")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("observa2")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("observa3")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("observa4")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("zona")) & "',"
   cad = cad & Val("" & Data2.Recordset.Fields("isc")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("tax")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("vtaneta")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("tcosto")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("ganancia")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("comision")) & ",'"
   cad = cad & Trim("" & Data2.Recordset.Fields("cajero")) & "','"
   cad = cad & Trim("" & caja) & "','"
   cad = cad & Trim("" & turno) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("servicio")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("comanda")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("mesa")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("salon")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("mesero")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("sentido")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("ccosto")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("familia")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("subfamilia")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("marca")) & "',"
   cad = cad & Val("" & Data2.Recordset.Fields("percepcion")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("tpercepcio")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("flete")) & ",'"
   cad = cad & Trim("" & Data2.Recordset.Fields("localf")) & "',"
   cad = cad & Val("" & Data2.Recordset.Fields("ivap")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("tivap")) & ",'"
   cad = cad & Trim("" & Data2.Recordset.Fields("nroprecio")) & "',"
   cad = cad & Val("" & Data2.Recordset.Fields("tisc")) & ",'"
   cad = cad & Trim("" & Data2.Recordset.Fields("placa")) & "',"
   cad = cad & Val("" & Data2.Recordset.Fields("xneto")) & ","
   cad = cad & Val("" & Data2.Recordset.Fields("tdetra")) & ",'"
   cad = cad & Trim("" & Data2.Recordset.Fields("denumero")) & "','"
   cad = cad & Trim("" & Data2.Recordset.Fields("categoria")) & "')"
   'MsgBox cad
   cn.Execute (cad)
   Data2.Recordset.MoveNext
Loop
borra_congela
cmdCancelar_Click
End Sub
Sub pedido_reposicion() 'sirva para que ingresen lo que necesitan que repongan
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
Dim sdx As Double
Dim rs
Dim i As Integer
Dim xcongela As String
Dim sw As Integer
sdx = Val("" & mytable11.Fields("congela")) + 1
xcongela = "" & sdx

denuevo13:
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM crequisa where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='01' and serie='Q' and numero='" & xcongela & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   sdx = Val(xcongela) + 1
   xcongela = "" & sdx
   GoTo denuevo13
End If
   'mytable11.Edit
   mytable11.Fields("congela") = xcongela
   mytable11.Update
   mytablex.AddNew
   mytablex.Fields("codigo") = Trim("" & mytable11.Fields("local"))
   mytablex.Fields("local") = Trim("" & mytable11.Fields("local"))
   mytablex.Fields("serie") = "01"
   mytablex.Fields("tipo") = "Q"
   mytablex.Fields("numero") = xcongela
   mytablex.Fields("fecha") = Format(dia, "dd/mm/yyyy")
   mytablex.Fields("moneda") = Trim("" & mytable11.Fields("moneda"))
   mytablex.Fields("paridad") = Val(paridad)
   mytablex.Fields("dias") = 1
   mytablex.Fields("acu") = "Q"
   mytablex.Fields("estado") = "2"
   mytablex.Fields("bodega") = Trim("" & mytable11.Fields("bodega"))
   mytablex.Fields("caja") = "" & caja
   mytablex.Fields("nombre") = "" & busca_local_pedido(Trim("" & mytable11.Fields("local")))
   mytablex.Fields("tipoclie") = "V"
   mytablex.Fields("turno") = "" & turno
   mytablex.Fields("usuario") = "" & cajero
   mytablex.Fields("hora") = Format(Now, "hh:MM")
   mytablex.Fields("total") = Val("" & txtotal)
   mytablex.Update
   mytablex.Close
'---ahora grabano detalle
ak12:
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM drequisa where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='01' and serie='Q' and numero='" & xcongela & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   mytablex.Delete
   GoTo ak12
End If
Set rs = Data2.Recordset.Clone
Do
    If rs.EOF Then Exit Do
    mytablex.AddNew
    For i = 0 To rs.Fields.count - 1
        mytablex.Fields(i) = rs.Fields(i)
    Next i
    mytablex.Fields("local") = Trim("" & mytable11.Fields("local"))
    mytablex.Fields("serie") = "01"
    mytablex.Fields("tipo") = "Q"
    mytablex.Fields("numero") = "" & xcongela
    mytablex.Fields("vendedor") = ""
    mytablex.Fields("codigo") = Trim("" & mytable11.Fields("local"))
    mytablex.Fields("local") = Trim("" & mytable11.Fields("local"))
    mytablex.Fields("moneda") = Trim("" & mytable11.Fields("moneda"))
    mytablex.Fields("bodega") = Trim("" & mytable11.Fields("bodega"))
    mytablex.Fields("bodegaf") = ""
    mytablex.Fields("acu") = "Q"
    mytablex.Fields("acu1") = ""
    mytablex.Fields("flage") = ""
    mytablex.Fields("tipoclie") = "V"
    mytablex.Fields("codigo") = ""
    mytablex.Fields("caja") = "" & caja
    mytablex.Fields("turno") = "" & turno
    mytablex.Fields("usuario") = "" & cajero
    mytablex.Fields("fecha") = Format(dia, "dd/mm/yyyy")
    mytablex.Fields("hora") = Format(Now, "hh:MM")
    mytablex.Update
    rs.MoveNext
Loop
mytablex.Close
End Sub

Private Sub coclasifica_Click()
If coclasifica <> "%" Then
   clasificacion = extra_loquesea(coclasifica)
End If
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus
   Exit Sub
End If
If Len(codigo) > 0 Then
   found = busca_codigo_descuento("" & codigo)
   If found = 1 Then
      DBGrid2.Col = 0
      DBGrid2.Row = DBGrid2.VisibleRows - 1
      DBGrid2.SetFocus
      Exit Sub
   End If
End If
'nombre.SetFocus
End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_cliente1
End If

End Sub
Sub slq_consultax()
Dim buf As String
      buf = "select Nombre,Codigo,Direccion,Distrito,Fechanac from clientes  where nombre  like '" & buffer & "%'"
               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldir
               Data1.RecordSource = buf
               Data1.refresh
               If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
                  Data1.Recordset.Close
                  buffer = ""
               End If
               dbGrid1.columns(0).Width = 4000
               dbGrid1.columns(1).Width = 2000
               dbGrid1.SetFocus
      

End Sub

Function sql_consulta(sw As Integer)
Dim buf As String
Dim queprecio As String
Dim indx As Integer
Dim dbf1 As String
Dim dbf2 As String
Dim amfecha As String
On Error GoTo cmd8912_err
'MsgBox buffer
amfecha = Format(dia, "YYYYMMDD")
indx = -1
dbf1 = ""
dbf2 = ""
If Trim("" & mytable11.Fields("t0")) = "S" Then
   If Len("" & mytable11.Fields("t1")) > 0 Then
      dbf1 = "  (caja='" & "" & mytable11.Fields("t1") & "'"
      If Len("" & mytable11.Fields("t2")) > 0 Then
         dbf1 = dbf1 & " or caja='" & "" & mytable11.Fields("t2") & "'"
      End If
      If Len("" & mytable11.Fields("t3")) > 0 Then
         dbf1 = dbf1 & " or caja='" & "" & mytable11.Fields("t3") & "'"
      End If
      If Len("" & mytable11.Fields("t4")) > 0 Then
         dbf1 = dbf1 & " or caja='" & "" & mytable11.Fields("t4") & "'"
      End If
      If Len("" & mytable11.Fields("t5")) > 0 Then
         dbf1 = dbf1 & " or caja='" & "" & mytable11.Fields("t5") & "'"
      End If
      If Len("" & mytable11.Fields("t6")) > 0 Then
         dbf1 = dbf1 & " or caja='" & "" & mytable11.Fields("t6") & "'"
      End If
      If Len("" & mytable11.Fields("t7")) > 0 Then
         dbf1 = dbf1 & " or caja='" & "" & mytable11.Fields("t7") & "'"
      End If
      If Len("" & mytable11.Fields("t8")) > 0 Then
         dbf1 = dbf1 & " or caja='" & "" & mytable11.Fields("t8") & "'"
      End If
      If Len("" & mytable11.Fields("t9")) > 0 Then
         dbf1 = dbf1 & " or caja='" & "" & mytable11.Fields("t9") & "'"
      End If
      dbf1 = dbf1 & ")"
   End If
   'Else 'si no esta programado solamente ver los congelados mios
   'dbf2 = "  (caja='" & "" & mytable11.Fields("caja") & "')"
End If
dbf2 = "  (caja='" & "" & mytable11.Fields("caja") & "')"
'If Len(dbf2) = 0 Then
'   dbf2 = dbf1
'End If
'MsgBox dbf2
queprecio = "precios.pventa1 as Precio "
'MsgBox buffer
'0 consulta delivery
If opcion1 = "0" Then  'si es delivery
If Len(buffer) = 0 Then  'AQUI DEBE APARECER
   buf = "select Telefono,Nombre,Codigo from telefono "
   Else
   buf = "select Telefono,Nombre,Codigo from telefono where "
   buf = buf & "" & Combo1 & " like '" & buffer & "%'"
   'indx = dbGrid1.Col
End If
End If
If opcion1 = "370" Then
If Len(buffer) = 0 Then
   buf = "select Nombre,Numero,Fecha,Moneda as M,Total,Hora,Usuario,Caja,Turno from crequisa where local='" & "" & mytable11.Fields("local") & "'"
   Else
   buf = "select Nombre,Numero,Fecha,Moneda as M,Total,Hora,Usuario,Caja,Turno from reponec where "
   buf = buf & " local='" & "" & mytable11.Fields("local") & "'"
   buf = buf & " and " & Combo1 & " like '" & buffer & "%'"
   'indx = dbGrid1.Col
   End If
End If

If opcion1 = "150" Then  'descongela
If Len(buffer) = 0 Then
   buf = "select Nombre,Numero,Fecha,Moneda as M,Total,Hora,Usuario,Caja,Turno from congelac "
   If Len(dbf2) > 0 Then
      buf = buf & " where "
   End If
   buf = buf & dbf2
   Else
   buf = "select Nombre,Numero,Fecha,Moneda as M,Total,Hora,Usuario,Caja,Turno from congelac where "
   buf = buf & "" & Combo1 & " like '" & buffer & "%'"
   If Len(dbf2) > 0 Then
      buf = buf & " and "
   End If
   buf = buf & dbf2
   'indx = dbGrid1.Col
   End If
   'MsgBox buf
End If
If opcion1 = "1900" Then
   If Len(buffer) = 0 Then
      buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Vendedor,Hora,Caja,Turno,Local from cproform  where local='" & "" & mytable11.Fields("local") & "'"
      If Len(dbf1) > 0 Then
         buf = buf & " and "
      End If
      buf = buf & dbf1
   Else
   buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Vendedor,Hora,Caja,Turno,Local from cproform where local='" & "" & mytable11.Fields("local") & "' and "
   buf = buf & "  " & Combo1 & " like '" & buffer & "%'"
   If Len(dbf1) > 0 Then
      buf = buf & " and "
   End If
   buf = buf & dbf1
   buf = buf & "   order by tipo,str(numero) "
   
   End If
End If


If opcion1 = "15000" Then  'carga ordenes de trabajo
   If Len(buffer) = 0 Then
   buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Vendedor,Hora,Caja,Turno,Local,acuenta from cpedidov where local='" & "" & "" & mytable11.Fields("local") & "' and "
   'buf = buf & "  fecha=" & "DateValue('" & dia & "'" & ")"
   'buf = buf & " yausado<>'1' and "
   buf = buf & "  yausado<>'1' and "
   buf = buf & "  caja='" & caja & "'"
   buf = buf & " order by fecha,HORA"
   Else
   buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Vendedor,Hora,Caja,Turno,Local,acuenta from cpedidov where local='" & "" & "" & mytable11.Fields("local") & "' and "
   'buf = buf & "  fecha=" & "DateValue('" & dia & "'" & ")"
   'buf = buf & " yausado<>'1' and "
   buf = buf & "  yausado<>'1'  and "
   buf = buf & "  caja='" & caja & "'"
   buf = buf & " and " & Combo1 & " like '" & buffer & "%'"
   buf = buf & "  order by fecha,HORA "
   'indx = dbGrid1.Col
   
   End If
End If

If opcion1 = "30000" Then  'carga cotizaciones
   If Len(buffer) = 0 Then
   buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Vendedor,Hora,Caja,Turno,Local,acuenta from ccotizav where local='" & "" & mytable11.Fields("local") & "'"
   'buf = buf & "  fecha=" & "DateValue('" & dia & "'" & ")"
   buf = buf & " and (yausado='0' or yausado=null)"
   buf = buf & " order by HORA"
   Else
   buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Vendedor,Hora,Caja,Turno,Local,acuenta from ccotizav where local='" & "" & mytable11.Fields("local") & "'"
   buf = buf & " and (yausado='0' or yausado=null)"
   buf = buf & " and " & Combo1 & " like '" & buffer & "%'"
   buf = buf & "  order by HORA "
   'indx = dbGrid1.Col
   End If
End If

'MsgBox buf

If opcion1 = "10" Or opcion1 = "13" Or opcion1 = "14" Or opcion1 = "15" Or opcion1 = "100" Or opcion1 = "1500" Then
   If Len(buffer) = 0 Then
   buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Servicio as S,Placa as Ok,Vendedor,Hora,Caja,Turno,Local from " & gocabeza & " where local='" & "" & mytable11.Fields("local") & "' and "
   buf = buf & "  fecha='" & amfecha & "'"
   buf = buf & " and usuario='" & cajero & "'"
   buf = buf & " and caja='" & caja & "'"
   buf = buf & " and turno='" & turno & "'"
   buf = buf & " order by HORA"
   Else
   buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Servicio as S,Placa as Ok,Vendedor,Hora,Caja,Turno,Local from " & gocabeza & " where local='" & "" & mytable11.Fields("local") & "' and "
   buf = buf & "  fecha='" & amfecha & "'"
   buf = buf & " and usuario='" & cajero & "'"
   buf = buf & " and caja='" & caja & "'"
   buf = buf & " and turno='" & turno & "'"
   buf = buf & " and " & Combo1 & " like '" & buffer & "%'"
   buf = buf & "  order by HORA "
   'indx = dbGrid1.Col
   End If
End If
If opcion1 = "750" Then
If Len(buffer) = 0 Then
   buf = "select FlaG_deli,tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Servico as S,Vendedor,Hora,Caja,Turno,Local from " & gocabeza & " where local='" & "" & mytable11.Fields("local") & "' and "
   buf = buf & "  fecha='" & amfecha & "'"
   buf = buf & " and servicio='D' "
   buf = buf & " and usuario='" & cajero & "'  order by tipo,str(numero)"
   Else
   buf = "select Flaf_deli as PDeli,tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Servicio as S,Vendedor,Hora,Caja,Turno,local from " & gocabeza & " where local='" & "" & mytable11.Fields("local") & "' and "
   buf = buf & "  fecha='" & amfecha & "'"
   buf = buf & " and usuario='" & cajero & "'"
   buf = buf & " and servicio='D' "
   buf = buf & " and " & Combo1 & " like '" & buffer & "%'"
   buf = buf & "  order by HORA "
   End If
End If
If opcion1 = "19000" Then
   If Len(buffer) = 0 Then
      buf = "select Descripcio,Servicio from servicio where servicio<>'D' and servicio<>'C' and servicio<>'A' "
   Else
      buf = "select Descripcio,Servicio from servicio  where servicio<>'D' and servicio<>'C' and servicio<>'A' and " & "" & Combo1 & " like '" & buffer & "%'"
   End If
End If

If opcion1 = "1" Then
   If Len(buffer) = 0 Then
      buf = "select Deliveri.telefono,Clientes.Nombre,deliveri.Direccion,deliveri.referencia,Clientes.Codigo,clientes.fechanac,clientes.clasifica from clientes inner join deliveri on deliveri.codigo=clientes.codigo "
   Else
      buf = "select Deliveri.telefono,Clientes.Nombre,clientes.Direccion,deliveri.referencia,Clientes.Codigo,clientes.fechanac,clientes.clasifica from  clientes inner join deliveri  on deliveri.codigo=clientes.codigo and " & "" & Combo1 & " like '" & buffer & "%'"
   End If
End If
If opcion1 = "2003" Then
If Len(buffer) = 0 Then
      buf = "select Local,Tipo,Serie,Numero,Fecha,Estado,Total,Acuenta,(Total-Acuenta) as saldo,codigo from cpedidov where codigo='" & tcampo1 & "' AND (Total-Acuenta)>0"
   Else
      buf = "select Local,Tipo,Serie,Numero,Fecha,Estado,Total,Acuenta,(Total-Acuenta) as saldo,codigo from cpedidov  where " & "" & Combo1 & " like '" & buffer & "%' and codigo='" & tcampo1 & "' AND (Total-Acuenta)>0"
   End If
End If
If opcion1 = "30" Or opcion1 = "99" Then
   If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo,Tipo,Codigo1,Direccion,Distrito from clientes "
   Else
      buf = "select Nombre,Codigo,Tipo,Codigo1,Direccion,Distrito from clientes  where " & "" & Combo1 & " like '" & buffer & "%'"
      'indx = dbGrid1.Col
   End If
End If

If opcion1 = "300" Then
   If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo,Tipo,Codigo1,Direccion,Distrito from bodega "
   Else
      buf = "select Nombre,Codigo,Tipo,Codigo1,Direccion,Distrito from bodega  where " & "" & Combo1 & " like '" & buffer & "%'"
      'indx = dbGrid1.Col
   End If
End If
If opcion1 = "1750" Then  'consulta de telefonos de clientes
   If Len(buffer) = 0 Then
      buf = "select Nombre,Direccion,telefono,Distrito,Fechanac,Codigo,tipo,Codigo1 from clientes "
   Else
      buf = "select Nombre,Direccion,Telefono,Distrito,Fechanac,Codigo,Tipo,Codigo1 from clientes  where " & "" & Combo1 & " like '" & buffer & "%'"
   End If
End If

If opcion1 = "300" Then
   If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo,Direccion,Distrito from bodega "
   Else
      buf = "select Nombre,Codigo,Direccion,Distrito from bodega  where " & "" & Combo1 & " like '" & buffer & "%'"
      'indx = dbGrid1.Col
   End If
End If
If opcion1 = "29" Then
   If Len(buffer) = 0 Then
   buf = "select Descripcio,Tipo from Tipo where (tipodoc='1' or tipodoc='A' or tipodoc='B' or tipodoc='C' or tipodoc='D' or tipodoc='G') order by tipo"
   Else
   buf = "select Descripcio,Tipo from Tipo  where (tipodoc='1' or tipodoc='A' or tipodoc='B' or tipodoc='C' or tipodoc='D' or tipodoc='G') and "
   buf = buf & "" & Combo1 & " like '" & buffer & "%'"
   buf = buf & "  order by tipo"
   'indx = dbGrid1.Col
   End If
End If
If opcion1 = "31" Then
   If Len(buffer) = 0 Then
   buf = "select Nombre,Codigo from Vendedor "
   Else
   buf = "select Nombre,Codigo from Vendedor  where "
   buf = buf & "" & Combo1 & " like '" & buffer & "%'"
   'indx = dbGrid1.Col
   End If
End If
If opcion1 = "200" Then
   If Len(buffer) = 0 Then
   buf = "select Descripcio,Banco from Banco "
   Else
   buf = "select Descripcio,Banco from Banco  where "
   buf = buf & "" & Combo1 & " like '" & buffer & "%'"
   'indx = dbGrid1.Col
   End If
End If
If opcion1 = "2800" Then 'consulta cuenta corriente favor cliente
   If Len(buffer) = 0 Then
   buf = "select Nombre,Codigo,Local,Tipo,Serie,Numero,Fecha,Saldo,Abono,Total from cuentac where anticipo='1' order by nombre,fecha"
   Else
   buf = "select Nombre,Codigo,Local,Tipo,Serie,Numero,Fecha,Saldo,Abono,Total from cuentac   where  anticipo='1' and "
   buf = buf & "" & Combo1 & " like '" & buffer & "%' order by nombre,fecha"
   'indx = dbGrid1.Col
   
   End If
End If
If opcion1 = "23" Then
   If Len(buffer) = 0 Then
   buf = "select Nombre,Codigo,Direccion,Distrito,Fechanac from clientes "
   Else
   buf = "select Nombre,Codigo,Codigo,Direccion,Distrito,Fechanac from clientes  where "
   buf = buf & "" & Combo1 & " like '" & buffer & "%'"
   'indx = dbGrid1.Col
   
   End If
End If

If opcion1 = "12" Then
   If Len(buffer) = 0 Then
   buf = "select Nombre,Codigo,Direccion,Distrito,Fechanac from clientes "
   Else
   buf = "select Nombre,Codigo,Direccion,Distrito,Fechanac from clientes  where "
   buf = buf & "" & Combo1 & " like '" & buffer & "%'"
   'indx = dbGrid1.Col
   End If
End If
If opcion1 = "8" Then
      If Len(buffer) = 0 Then
      'buf = "select Producto.Descripcio,Producto.producto,Producto.Marca,precios.Unidad1 as Und1,precios.Factor1 as F," & queprecio & " ,Producto.Monedav as M,Producto.Familia,Producto.barras,Producto.oferta,producto.estado from producto  left join precios on producto.producto=precios.producto  where producto.estado<>'N'   and precios.local='" & "" & mytable11.Fields("listap") & "'"
      buf = "select Producto.Descripcio,Producto.producto,Producto.Marca,precios.Unidad1 as Und1,precios.Factor1 as F," & queprecio & " ,Producto.Monedav as M,Producto.Familia,Producto.barras,Producto.Remate,producto.estado from producto  left join precios on producto.producto=precios.producto  where  precios.local='" & "" & mytable11.Fields("listap") & "'"
      Else
      'buf = "select Producto.Descripcio,Producto.producto,Producto.Marca,precios.Unidad1 as Und1,precios.Factor1 as F," & queprecio & " ,Producto.Monedav as M,Producto.Familia,Producto.barras,Producto.oferta,producto.estado from producto left join precios on producto.producto=precios.producto WHERE  producto.estado<>'N' and precios.local='" & "" & mytable11.Fields("listap") & "' and "
      buf = "select Producto.Descripcio,Producto.producto,Producto.Marca,precios.Unidad1 as Und1,precios.Factor1 as F," & queprecio & " ,Producto.Monedav as M,Producto.Familia,Producto.barras,Producto.Remate,producto.estado from producto left join precios on producto.producto=precios.producto WHERE   precios.local='" & "" & mytable11.Fields("listap") & "' and "
      buf = buf & "" & Combo1 & " like '" & buffer & "%'"
      'indx = dbGrid1.Col
      End If
End If
If opcion1 = "8" Then
   If "" & mytable11.Fields("ordenaproducto") = "S" Then
      buf = buf & " order by descripcio"
   End If
End If
'MsgBox buf

   Set rcconsulta = Nothing
   If rcconsulta.State = 1 Then
      rcconsulta.Close
      Set rcconsulta = Nothing
   End If
   'MsgBox buf & " " & opcion1
   
   
   
   rcconsulta.Open buf, cn, adOpenStatic, adLockOptimistic
   Set dbGrid1.DataSource = rcconsulta
   dbGrid1.refresh
   'MsgBox buf
   If rcconsulta.RecordCount = 0 Then
      buffer.SetFocus
      Exit Function
   End If
   'sw_consulta = 1

               
               If opcion1 = "8" Then
                  pone_precios "" & rcconsulta.Fields("producto")
               End If
               dbGrid1.columns(0).Width = 4000
               dbGrid1.columns(1).Width = 2000
               If opcion1 = "0" Then   'consulta deliveri
               dbGrid1.columns(0).Width = 1500
               dbGrid1.columns(1).Width = 4300
               dbGrid1.columns(2).Width = 1500
              End If
              If opcion1 = "750" Then
               dbGrid1.columns(0).Width = 700
               dbGrid1.columns(1).Width = 700
               dbGrid1.columns(2).Width = 700
               dbGrid1.columns(3).Width = 1300
               dbGrid1.columns(4).Width = 1500
               dbGrid1.columns(5).Width = 3000
               dbGrid1.columns(6).Width = 1500
               dbGrid1.columns(7).Width = 400
               dbGrid1.columns(8).Width = 1400
               dbGrid1.columns(9).Width = 400
               dbGrid1.columns(10).Width = 1300
               dbGrid1.columns(11).Width = 1300
               dbGrid1.columns(12).Width = 700
               dbGrid1.columns(13).Width = 700
              End If
              If opcion1 = "10" Or opcion1 = "13" Or opcion1 = "14" Or opcion1 = "15" Or opcion1 = "100" Or opcion1 = "1500" Or opcion1 = "1900" Or opcion1 = "15000" Or opcion1 = "30000" Then
               dbGrid1.columns(0).Width = 700
               dbGrid1.columns(1).Width = 700
               dbGrid1.columns(2).Width = 1300
               dbGrid1.columns(3).Width = 1500
               dbGrid1.columns(4).Width = 3000
               dbGrid1.columns(5).Width = 1500
               dbGrid1.columns(6).Width = 400
               dbGrid1.columns(7).Width = 1400
               dbGrid1.columns(8).Width = 400
               dbGrid1.columns(9).Width = 400
               dbGrid1.columns(10).Width = 400
               dbGrid1.columns(11).Width = 1300
               dbGrid1.columns(12).Width = 1300
               dbGrid1.columns(13).Width = 700
               dbGrid1.columns(14).Width = 700
               
              End If
              If opcion1 = "8" Then
               dbGrid1.columns(0).Width = 5000
               dbGrid1.columns(1).Width = 1300
               dbGrid1.columns(2).Width = 1000
               dbGrid1.columns(3).Width = 900
               dbGrid1.columns(4).Width = 500
               dbGrid1.columns(5).Width = 800
               dbGrid1.columns(6).Width = 500
               dbGrid1.columns(7).Width = 1000
               dbGrid1.columns(8).Width = 1500
               dbGrid1.columns(9).Width = 1500
               End If
               If opcion1 = "150" Then
               dbGrid1.columns(0).Width = 5000
               dbGrid1.columns(1).Width = 1300
               dbGrid1.columns(2).Width = 1500
               dbGrid1.columns(3).Width = 900
               dbGrid1.columns(4).Width = 1500
               dbGrid1.columns(5).Width = 900
               dbGrid1.columns(6).Width = 1200
               dbGrid1.columns(7).Width = 700
               End If
               If opcion1 = "2800" Then
               dbGrid1.columns(0).Width = 5000
               dbGrid1.columns(1).Width = 1300
               dbGrid1.columns(2).Width = 500
               dbGrid1.columns(3).Width = 500
               dbGrid1.columns(4).Width = 500
               dbGrid1.columns(5).Width = 1300
               dbGrid1.columns(6).Width = 1200
               dbGrid1.columns(7).Width = 900
               End If
               If sw = 1 Then
                  dbGrid1.SetFocus
               End If
               'MsgBox opcion1
              If opcion1 = "150" Or opcion1 = "10" Or opcion1 = "13" Or opcion1 = "14" Or opcion1 = "15" Or opcion1 = "100" Or opcion1 = "1500" Or opcion1 = "1900" Or opcion1 = "15000" Or opcion1 = "30000" Then
                 ir_hasta_ultimo rcconsulta
              End If
               
sql_consulta = 1
Exit Function
cmd8912_err:
MsgBox "Aviso en sql_consulta " & error$, 48, "Aviso"
buffer = ""
Exit Function
End Function
Sub ir_hasta_ultimo(rcconsulta As ADODB.Recordset)
On Error GoTo cmd789111_err
rcconsulta.MoveLast
'dbGrid1.Col = 0
'dbGrid1.Row = dbGrid1.VisibleRows - 1
'dbGrid1.SetFocus
 
Exit Sub
cmd789111_err:
MsgBox "Aviso en ir ultimo " + error$, 48, "Aviso"
Exit Sub
End Sub

Private Sub Command1_Click()
Dim found As Integer
found = sql_consulta(1)
End Sub

Private Sub Command10_Click()
If Len(telefono) > 0 Or Len(nombre) > 0 Or Len(ddireccion) > 0 Or Len(fechanac) > 0 Or Len(codigo) > 0 Then
   MsgBox "Existen Campos", 48, "Aviso"
   Exit Sub
End If
Frame2.Visible = False
'tiposervicio1 = "Autoservicio"
'flag_servicio = "A"
DBGrid2.SetFocus
End Sub

Private Sub Command11_Click()
fechanac_KeyPress 13
End Sub

Private Sub Command12_Click()
inicializa_deliveri
telefono.SetFocus
End Sub

Private Sub Command13_Click()
Dim found As Integer
Dim sw As Integer
Dim sdx As Double
If Len(pedido) = 0 Then  'si no es modificacion
found = valida_total()
If found = 0 Then
   MsgBox "Campos invalidos", 48, "Aviso"
   Exit Sub
End If
End If
   
   ndetraccion = ""
   'If Val(tdetra) > 0 Then
   '   sdx = Val("" & mytable11.Fields("detraccion")) + 1
   '   ndetraccion = "" & sdx
   'End If
   
If Len(xnombre) > 0 Then
   If local1.Visible = False Then  'si no es traslado locales
      found = graba_cliente_tipo("" & xruc) 'ojo graba con el correlativo
   End If
End If
If Len(pedido) > 0 Then
   xtipo = "P"
   xserie = "P"
   xnumero = "" & pedido
End If
cgusuario = gocabeza
dgusuariog = godetalle
If Len(flag_servicio) = 0 Then
   MsgBox "No existe servicio seleccionado ", 48, "Aviso"
   Exit Sub
End If
'If flag_servicio = "C" Or flag_servicio = "A" Or flag_servicio = "D" Then
   found = busca_numero(xtipo, xserie, xnumero) 'busca numero libre
   If found = -1 Then  'si es boleta o factura manual
      xnumero.SetFocus
      Exit Sub
   End If
   opcion1 = "0"
   If local1.Visible = True Then
      opcion1 = "9999"
   End If
   Frame7.Enabled = False
   'DBGrid2.Enabled = False
   DBGrid2.Enabled = False
   Command13.Enabled = False
   adiciona_deliveri xtipo, xserie, xnumero
   DBGrid2.Enabled = True
   Command13.Enabled = True
   Frame7.Enabled = True
   Framefp.Enabled = False
   'Command14_Click
   Command6_Click
   'MsgBox "HOLA"
   limpia_general
'End If
'Frame10.Visible = True
End Sub

Private Sub Command14_Click()
   Frame7.Visible = False
   habilita_lab7 0
   Framefp.Enabled = True
   dbgrid10.Enabled = True
   If "" & mytable11.Fields("terminal") = "T" Then
      DBGrid2.Enabled = True
      DBGrid2.Col = 0
      DBGrid2.Row = DBGrid2.VisibleRows - 1
      DBGrid2.SetFocus
      Exit Sub
   End If
   'dbgrid10.Visible = True
   dbgrid10.SetFocus
End Sub

Private Sub Command15_Click()

End Sub

Private Sub Command5_Click()
End Sub

Private Sub Command2_Click()
   Frame6.Visible = False
   habilita_lab7 0
   dbgrid10.Enabled = True
   dbgrid10.SetFocus

End Sub

Private Sub Command3_Click()
tcampo5_KeyPress 13
End Sub

Private Sub Command4_Click()
DBGrid4_KeyDown 13, 0
End Sub

Private Sub Command6_Click()
If Frame7.Visible = True Then Exit Sub
losao94_Click
End Sub

Private Sub Command7_Click()
End Sub

Private Sub Command8_Click()
 'If Frame1.Visible = True Then
 '   Frame5.Visible = False
 '   dbGrid1.SetFocus
 '   Exit Sub
 'End If
 
 DBGrid4_KeyDown 27, 0
 
   'Frame5.Visible = False
   'DBGrid2.Col = 0
   'DBGrid2.Row = dbgrid2.visiblerows - 1
   'DBGrid2.SetFocus
End Sub

Private Sub Command9_Click()
End Sub

Private Sub d7822cua_Click()
    

End Sub

Private Sub d892323_Click()
Dim found As Double
flag_clave1 = 0
        tconcla.X = "CUADRE"
        tconcla.Show 1
        If flag_clave1 <> 1 Then  'si es descongela
           Exit Sub
        End If
        found = suma_las_ventas()
        MsgBox "VENTAS ACUMULADAS ..." & Format(found, "0.00"), 48, "AVISO"
End Sub
Function suma_las_ventas() As Double
Dim mysnapx As New ADODB.Recordset
Dim buf As String
   'MsgBox gocabeza
   buf = "select sum(total) as TOT from " & gocabeza & " where "
   buf = buf & "  fecha>='" & Format(dia, "YYYYMMDD") & "'"
   buf = buf & " and fecha<='" & Format(dia, "YYYYMMDD") & "' "

   'buf = buf & " fecha>=" & "DateValue('" & dia & "'" & ")"
   'buf = buf & " and fecha<=" & "DateValue('" & dia & "'" & ")"
   buf = buf & " and estado='2' "
   buf = buf & " and usuario='" & cajero & "'"
   buf = buf & " and caja='" & caja & "'"
   buf = buf & " and turno='" & turno & "'"
   'MsgBox buf

   mysnapx.Open buf, cn, adOpenStatic, adLockOptimistic
   If mysnapx.RecordCount > 0 Then
      'Set mysnapx = mydbxglo.CreateSnapshot(buf)
      suma_las_ventas = Val("" & mysnapx.Fields("TOT"))
   End If
   mysnapx.Close

End Function

Private Sub d8do82_Click()
Dim sw As Integer
flag_clave1 = 0
tconcla.X = "CUADRE"
tconcla.Show 1
If flag_clave1 <> 1 Then  'si es descongela
   DBGrid2.SetFocus
   Exit Sub
End If

    
    opcion1 = "4"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.cajero = usuariopos
    tcuadrc1.caja = caja
    tcuadrc1.turno = turno
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.fechai = dia
    tcuadrc1.fechaf = dia
    tcuadrc1.Caption = "PRODUCTOS VS DOCUMENTOS"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1
    
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim found As Integer
Dim xbuf As String
Dim buf As String
Dim mytablex As New ADODB.Recordset
Dim xtemp As Variant
Dim anumero As String
Dim atipo As String
Dim aserie As String
Dim sdx As Double
Dim canti As String
DBGrid2.Enabled = True
If KeyCode = 27 Then
   losao94_Click
   Exit Sub
End If
'MsgBox opcion1
'If KeyCode = 0 Then Exit Sub
'MsgBox opcion1
If KeyCode = &H2E Then  'borrar linea
   If opcion1 = "1900" Then 'borrar cproform
      If MsgBox("Desea Borrar Proforma " & rcconsulta.Fields("numero"), 1, "Aviso") <> 1 Then
         dbGrid1.SetFocus
         Exit Sub
      End If
      protipo = "" & rcconsulta.Fields("tipo")
      proserie = "" & rcconsulta.Fields("serie")
      pronumero = "" & rcconsulta.Fields("numero")
      found = borrar_proformas()
      protipo = ""
      proserie = ""
      pronumero = ""
      Frame1.Visible = False
      Frame1.Enabled = False
      DBGrid2.SetFocus
      Exit Sub
   End If
   Exit Sub
End If
If KeyCode = &H71 Then  'f2  Verifica delivery
   'MsgBox opcion1
   Select Case opcion1
          Case "15", "15000"
      'si es copia se puede grabar como ok devolvio delivery
      If Trim("" & rcconsulta.Fields("ok")) = "OK" Then
         rcconsulta.Fields("OK") = ""
         rcconsulta.Update
         Exit Sub
      End If
      If Trim("" & rcconsulta.Fields("OK")) <> "OK" Then
         rcconsulta.Fields("OK") = "OK"
         rcconsulta.Update
         Exit Sub
      End If
      
  End Select
  Exit Sub
End If
If KeyCode = &H70 Then  'f1  visualizar el detalle
   If opcion1 = "8" Then  'si esta en productos
      If Len("" & rcconsulta.Fields("producto")) > 0 Then
      
      xproducto = "" & rcconsulta.Fields("producto")
      carga_dbgrid4 "" & rcconsulta.Fields("producto")
      Exit Sub
   End If
   End If
   'MsgBox opcion1
   If opcion1 = "1500" Or opcion1 = "15" Or opcion1 = "100" Or opcion1 = "1900" Then
      If Len("" & rcconsulta.Fields("tipo")) > 0 Then
         visualiza_detalle_factura "" & rcconsulta.Fields("tipo"), "" & rcconsulta.Fields("serie"), "" & rcconsulta.Fields("numero")
         Exit Sub
      End If
   End If
   Exit Sub
End If
If KeyCode = 13 Then
If opcion1 = "8" Then
   If Trim("" & rcconsulta.Fields("estado")) = "N" Then
      MsgBox "Producto No activo ", 48, "Aviso"
      dbGrid1.SetFocus
      Exit Sub
   End If
   
   
stkminimo = ""
If "" & mytable11.Fields("stkminimo") = "S" Then
            mytablex.Open "SELECT * FROM producto where producto='" & "" & rcconsulta.Fields(1) & "'", cn, adOpenDynamic, adLockOptimistic
            If mytablex.RecordCount > 0 Then
               If familia_saldo("" & mytablex.Fields("familia")) = 0 Then
                  consulta_minimo "" & mytablex.Fields("producto"), "" & mytablex.Fields("minimo")
               End If
            End If
            mytablex.Close
            
End If

   
   
         If "" & mytable11.Fields("nosaldo") = "S" Then
            found = verifica_saldo_receta("" & rcconsulta.Fields(1), Val("" & dbGrid1.columns("cantidad")))
            If found = 2 Then
               MsgBox "Se detecto un saldo receta con saldo<=0 ", 48, "Aviso"
            Exit Sub
            End If
            If familia_saldo("" & DBGrid2.columns("familia")) = 0 Then
            If consulta_saldo("" & rcconsulta.Fields(1), 1, 0) <= 0 Then
               MsgBox "No existe saldo", 48, "Aviso"
               dbGrid1.SetFocus
               Exit Sub
            End If
            End If
         End If
       
   If Val(rcconsulta.Fields("precio")) <= 0 Then
      If "" & mytable11.Fields("noprecio") = "S" Then
         MsgBox "Precio<=0", 48, "Aviso"
         dbGrid1.SetFocus
         Exit Sub
      End If
   End If
   If Len("" & DBGrid2.columns(0)) = 0 And Len("" & rcconsulta.Fields(1)) > 0 Then
   
      'found = verifica_doble("" & rcconsulta.fields(1))
      'If found = 1 Then
      '   MsgBox "Producto ya seleccionado", 48, "Aviso"
      '   DBGrid1.SetFocus
      '   Exit Sub
      'End If
      canti = ""
      If verifica_balanza("" & rcconsulta.Fields(1)) = "S" Then
ajk922:
      
      buf = puerto_balanza1()
        If Val(buf) <= 0 Then
           If MsgBox("Balanza No leido,Continua Leyendo? ", 1, "Aviso") = 1 Then
             GoTo ajk922
           End If
           losao94_Click
           Exit Sub
        End If
        canti = Format(Val(buf), "0.000")
     End If
         'AQUI TIENES QUE REVISAR JOHNNY
         '-----------------------------------
         'MsgBox ""
         DBGrid2.Col = 0
         DBGrid2.Row = DBGrid2.VisibleRows - 1
         DBGrid2.SetFocus
         'Exit Sub
         '-------------------------------------
         'xtemp = DBGrid2.Row
         'Data2.Refresh
         'found = ir_ultimo_registrox()
         'DBGrid2.Refresh
         'DBGrid2.SetFocus
         'If xtemp = -1 Then
         '   xtemp = 0
         'End If
         'opcion1 = ""
         'DBGrid2.Row = xtemp
         'DBGrid2.Col = 0
         DBGrid2.columns(0) = "" & rcconsulta.Fields(1)
         xbuf = "" & rcconsulta.Fields(1)
         'MsgBox xbuf
         found = busca_producto("" & DBGrid2.columns(0), 0, canti, 0)
         If found = 0 Then
            dbGrid1.SetFocus
            Exit Sub
         End If
         If found = 2 Then
            dbGrid1.SetFocus
            'MsgBox "hOLA"
            Exit Sub
         End If
         '------------------------.....lee la balanza
      'buf = ""
       'If "" & mytable11.Fields("actbala") = "S" Then
     'If verifica_balanza("" & rcconsulta.fields(1)) = "S" Then
'ajk92:
 '    buf = puerto_balanza1()
 '       If Len(buf) <= 0 Then
 '          If MsgBox("Balanza No leido,Continua Leyendo? ", 1, "Aviso") = 1 Then
 '             GoTo ajk92
 '             Else
 '
 '          End If
 '       End If
 '    End If
 '    End If
     
 '    If Val(buf) > 0 Then
 '       DBGrid2.Columns("cantidad") = Val(Mid$(Val(buf), 1, 5))
 '       sdx = Val("" & DBGrid2.Columns("cantidad")) * Val("" & DBGrid2.Columns("precio"))
 '       DBGrid2.Columns("total") = sdx
 '       calcula_igv 0
 '    End If
      '------------------------------------------------
      Frame1.Visible = False
      Frame1.Enabled = False
      'found = sumar_detalle()
      'aqui ponemos si tiene mas de un precio
      'msgbox xbuf
      If ver_si_puedo_dbgrid(xbuf) = 1 Then  'existe mas de un precio
         DBGrid2.Row = DBGrid2.VisibleRows - 2
         DBGrid2.Col = 3
         xproducto = xbuf
         carga_dbgrid4 xbuf
         swprecio = 1
         Exit Sub
     End If
     If Len(DBGrid2.columns(0)) > 0 And Len(DBGrid2.columns(17)) > 0 Then
         'DBGrid2.Col = 3
         'DBGrid2.SetFocus
         'ingreso_tallas "" & DBGrid2.Columns(17)
         'Exit Sub
     End If
     'verificar si tiene talla
     found = sumar_detalle()
      DBGrid2.Row = DBGrid2.VisibleRows - 1
      DBGrid2.SetFocus
      Exit Sub
   End If
   
End If
   If opcion1 = "10" Then  'modifica
      xtipo = "" & rcconsulta.Fields("tipo")
      xserie = "" & rcconsulta.Fields("serie")
      xnumero = "" & rcconsulta.Fields("numero")
      telefono = "" & rcconsulta.Fields("telefono")
      codigo = "" & rcconsulta.Fields("codigo")
      nombre = "" & rcconsulta.Fields("nombre")
      found = busca_codigod()
      modifica_detalle
      If "" & rcconsulta.Fields("servicio") = "A" Then
         tiposervicio = "Autoservicio"
      End If
      If "" & rcconsulta.Fields("servicio") = "D" Then
         tiposervicio = "DELIVERY"
      End If
      xestado = "Modifica"
      Data2.refresh
      Frame1.Visible = False
      Frame1.Enabled = False
   End If
   If opcion1 = "0" Then
   telefono = "" & rcconsulta.Fields("telefono")
   dcodigo = "" & rcconsulta.Fields("codigo")
   dnombre = "" & rcconsulta.Fields("nombre")
   found = busca_codigod()
   Frame1.Visible = False
   Frame1.Enabled = False
   dcodigo.SetFocus
   dcodigo_KeyPress 13
   End If
   If opcion1 = "1" Then
   If Len(Trim("" & rcconsulta.Fields("codigo"))) = 0 Then
      Exit Sub
   End If
   telefono = Trim("" & rcconsulta.Fields("telefono"))
   dcodigo = Trim("" & rcconsulta.Fields("codigo"))
   dnombre = Trim("" & rcconsulta.Fields("nombre"))
   ddireccion = Trim("" & rcconsulta.Fields("direccion"))
   fechanac = Trim("" & rcconsulta.Fields("fechanac"))
   referencia = Trim("" & rcconsulta.Fields("referencia"))
   clasificacion = Trim("" & rcconsulta.Fields("clasifica"))
   saludo_cumpe

   Frame1.Visible = False
   Frame1.Enabled = False
   dnombre.SetFocus
   'dcodigo_KeyPress 13
   End If
   If opcion1 = "1750" Then
   dcodigo = "" & rcconsulta.Fields("codigo")
   dnombre = "" & rcconsulta.Fields("nombre")
   ddireccion = "" & rcconsulta.Fields("direccion")
   fechanac = "" & rcconsulta.Fields("fechanac")
   telefono = "" & rcconsulta.Fields("telefono")
   saludo_cumpe

   Frame1.Visible = False
   Frame1.Enabled = False
   dcodigo.SetFocus
   dcodigo_KeyPress 13
   End If
   
   If opcion1 = "23" Then
      tcampo1 = "" & rcconsulta.Fields("codigo")
      Frame1.Visible = False
      Frame1.Enabled = False
      tcampo1.SetFocus
      tcampo1_KeyPress 13
   End If
   If opcion1 = "200" Then
      tcampo4 = "" & rcconsulta.Fields("banco")
      Frame1.Visible = False
      Frame1.Enabled = False
      tcampo4.SetFocus
   End If

   If opcion1 = "12" Then
   codigo = "" & rcconsulta.Fields("codigo")
   nombre = "" & rcconsulta.Fields("nombre")
   Frame1.Visible = False
   Frame1.Enabled = False
   codigo.SetFocus
   codigo_KeyPress 13
   End If
   If opcion1 = "2003" Then
   tcampo3 = "" & rcconsulta.Fields("tipo")
   tcampo4 = "" & rcconsulta.Fields("serie")
   tcampo5 = "" & rcconsulta.Fields("numero")
   totpedido = "" & suma_pedidos("" & rcconsulta.Fields("codigo"), "" & rcconsulta.Fields("tipo"), "" & rcconsulta.Fields("serie"), "" & rcconsulta.Fields("numero"))
   Frame1.Visible = False
   Frame1.Enabled = False
   tcampo3.SetFocus
   tcampo3_KeyPress 13
   End If
   If opcion1 = "19000" Then
   flag_servicio = Trim("" & rcconsulta.Fields("servicio"))
   
   Frame1.Visible = False
   Frame1.Enabled = False
   '--------------cobrar --------------------------------
   Label36.Caption = "Codigo"
   cobra_servicio
   'found = proceso_cobros()
   'opcion2 = 0
   'ttxtotals = Format(Val(rtxtotal), nrodecimal)
   'ttxtotald = Format(Val(rtxtotald), nrodecimal)
   'stxtotals = Format(Val(rtxtotal), nrodecimal)
   'stxtotald = Format(Val(rtxtotald), nrodecimal)
   'Framefp.Visible = True
   'Framefp.Enabled = True
'MsgBox "Hola"
   'dbgrid10.Enabled = True
   'dbgrid10.SetFocus
   'DBGrid10_KeyDown 13, 0
   palabra_bienvenida1
   Exit Sub
   End If
   
   If opcion1 = "300" Then 'bodega de traslado
   xruc = "" & rcconsulta.Fields("codigo")
   xnombre = "" & rcconsulta.Fields("nombre")
   Frame1.Visible = False
   Frame1.Enabled = False
   xruc.SetFocus
   xruc_KeyPress 13
   End If
   
   If opcion1 = "31" Then
   xvendedor = "" & rcconsulta.Fields("codigo")
   Frame1.Visible = False
   Frame1.Enabled = False
   xvendedor.SetFocus
   xvendedor_KeyPress 13
   End If

   If opcion1 = "30" Then
      If xtipo = "2" Or xtipo = "4" Then
         'If Len("" & rcconsulta.fields("ruc")) <> 11 Then
         '   MsgBox "Ruc Invalido ", 48, "Aviso"
         '   Exit Sub
         'End If
         xruc = Trim("" & rcconsulta.Fields("codigo"))
         Else
      xruc = Trim("" & rcconsulta.Fields("codigo"))
      End If
      codigo = Trim("" & rcconsulta.Fields("codigo"))
      xnombre = Trim("" & rcconsulta.Fields("nombre"))
      nombre = Trim("" & rcconsulta.Fields("nombre"))
      xdireccion = Trim("" & rcconsulta.Fields("direccion"))
   Frame1.Visible = False
   Frame1.Enabled = False
   xdireccion_KeyPress 13
   End If
   If opcion1 = "99" Then
   tcampo1 = "" & rcconsulta.Fields("codigo")
   tcampo2 = "" & rcconsulta.Fields("nombre")
   Frame1.Visible = False
   Frame1.Enabled = False
   tcampo1.SetFocus
   tcampo1_KeyPress 13
   End If
   If opcion1 = "2800" Then
   If Val("" & rcconsulta.Fields("saldo")) < Val(stxtotals) Then
                  MsgBox "Debe ingresar valor exacto", 48, "Aviso"
                  dbGrid1.SetFocus
                  Exit Sub
   End If
   tcampo1 = "" & rcconsulta.Fields("codigo")
   tcampo2 = "" & rcconsulta.Fields("nombre")
   tcampo3 = "" & rcconsulta.Fields("numero")
   tcampo4 = "" & rcconsulta.Fields("tipo")
   tcampo5 = "" & rcconsulta.Fields("serie")
   tcampo6 = "" & rcconsulta.Fields("local")
   saldoabo = "" & rcconsulta.Fields("saldo")
   Frame1.Visible = False
   Frame1.Enabled = False
   tcampo3.SetFocus
   'tcampo3_KeyPress 13
   End If

   If opcion1 = "29" Then
   xtipo = "" & rcconsulta.Fields(1)
   Frame1.Visible = False
   Frame1.Enabled = False
   xtipo.SetFocus
   xtipo_keyPress 13
   End If

   If opcion1 = "13" Then  'copia documento
   If MsgBox("Desea Sacar Copia del Documento", 1, "Aviso") <> 1 Then Exit Sub
   proceso_impresioncopia
   Frame1.Visible = False
   Frame1.Enabled = False
   DBGrid2.SetFocus
   End If
   If opcion1 = "1500" Then  'carga documento anterior
   If MsgBox("Desea Cargar Detalle Documento ", 1, "Aviso") <> 1 Then Exit Sub
      found = proceso_carga_doc_ant("" & rcconsulta.Fields("local"), "" & rcconsulta.Fields("tipo"), "" & rcconsulta.Fields("serie"), "" & rcconsulta.Fields("numero"))
   If found = 0 Then
      MsgBox "Error de carga", 48, "Aviso"
      Exit Sub
   End If
   Frame1.Visible = False
   Frame1.Enabled = False
   DBGrid2.SetFocus
   End If
   
   If opcion1 = "15000" Then  'carga pedidos de venta anteriores para cancelar
   If MsgBox("Desea Cargar Detalle Documento ", 1, "Aviso") <> 1 Then Exit Sub
      petipo = "" & rcconsulta.Fields("tipo")
      peserie = "" & rcconsulta.Fields("serie")
      penumero = "" & rcconsulta.Fields("numero")
      acuenta = "" & rcconsulta.Fields("acuenta")
      codigo = "" & rcconsulta.Fields("codigo")
      nombre = "" & rcconsulta.Fields("nombre")
      'cproven = "" & rcconsulta.Fields("vendedor")
      found = proceso_carga_Pedido("" & rcconsulta.Fields("local"), "" & rcconsulta.Fields("tipo"), "" & rcconsulta.Fields("serie"), "" & rcconsulta.Fields("numero"))
   If found = 0 Then
      MsgBox "Error de carga", 48, "Aviso"
      Exit Sub
   End If
   Frame1.Visible = False
   Frame1.Enabled = False
   DBGrid2.SetFocus
   End If
   
If opcion1 = "30000" Then  'carga cotizaciones
   If MsgBox("Desea Cargar Detalle Documento ", 1, "Aviso") <> 1 Then Exit Sub
      petipo = "" & rcconsulta.Fields("tipo")
      peserie = "" & rcconsulta.Fields("serie")
      penumero = "" & rcconsulta.Fields("numero")
      codigo = "" & rcconsulta.Fields("codigo")
      nombre = "" & rcconsulta.Fields("nombre")
      cproven = "" & rcconsulta.Fields("vendedor")
      found = proceso_carga_cotizacion("" & rcconsulta.Fields("local"), "" & rcconsulta.Fields("tipo"), "" & rcconsulta.Fields("serie"), "" & rcconsulta.Fields("numero"))
   If found = 0 Then
      MsgBox "Error de carga", 48, "Aviso"
      Exit Sub
   End If
   Frame1.Visible = False
   Frame1.Enabled = False
   DBGrid2.SetFocus
   End If
   
   
   If opcion1 = "1900" Then  'cargar proformas
   If MsgBox("Desea Cargar Proforma ", 1, "Aviso") <> 1 Then Exit Sub
      found = proceso_proforma("" & rcconsulta.Fields("local"), "" & rcconsulta.Fields("tipo"), "" & rcconsulta.Fields("serie"), "" & rcconsulta.Fields("numero"))
   If found = 0 Then
      MsgBox "Error de carga", 48, "Aviso"
      dbGrid1.SetFocus
      Exit Sub
   End If
   'sql_detalle
   'borrar_data1
   cproven = "" & rcconsulta.Fields("vendedor")
   codigo = "" & rcconsulta.Fields("codigo")
   nombre = "" & rcconsulta.Fields("nombre")
   protipo = "" & rcconsulta.Fields("tipo")
   proserie = "" & rcconsulta.Fields("serie")
   pronumero = "" & rcconsulta.Fields("numero")
   Frame1.Visible = False
   Frame1.Enabled = False
   DBGrid2.SetFocus
   End If

   If opcion1 = "14" Then  'BORRAR
   If MsgBox("Desea Borrar del Documento", 1, "Aviso") <> 1 Then Exit Sub
   PROCESO_BORRAR_DOCUMENTO "" & rcconsulta.Fields("local"), "" & rcconsulta.Fields("tipo"), "" & rcconsulta.Fields("serie"), "" & rcconsulta.Fields("numero")
   Frame1.Visible = False
   DBGrid2.SetFocus
   End If
   If opcion1 = "150" Then 'descongelar
      found = menu_descongela("" & rcconsulta.Fields(1))
      MsgBox "Presione enter para continuar...", 48, "Aviso"
      If found = 1 Then
         borrar_descongela1 "" & rcconsulta.Fields(1)
         borrar_descongela "" & rcconsulta.Fields(1)
         sql_detalle
         found = sumar_detalle()
         losao94_Click
      End If
   End If
   If opcion1 = "370" Then 'cargar reposicion para modificar
      found = menu_repone("" & rcconsulta.Fields("numero"))
      MsgBox "Presione enter para continuar...", 48, "Aviso"
      If found = 1 Then
         borrar_repone "" & rcconsulta.Fields("numero")
         borrar_reponexx
         sql_detalle
         found = sumar_detalle()
         losao94_Click
      End If
   End If
   If opcion1 = "750" Then  'deliveri no xxx
   If "" & rcconsulta.Fields("flag_deli") = "S" Then
      flag_clave1 = 0
      tconcla.X = "N"
      tconcla.Show 1
      If flag_clave1 <> 1 Then  'si es descongela
         DBGrid2.SetFocus
      Exit Sub
      End If

      'ojo esto debe estar..veificar
      'Data1.Recordset.Edit
      'Data1.Recordset.Fields("flag_deli") = ""
      'Data1.Recordset.Update
      
      
      'Frame1.Visible = False
      'DBGrid2.SetFocus
      Exit Sub
   End If
   If "" & rcconsulta.Fields("flag_deli") = "" Then
      'esto debe estar verificar
      'Data1.Recordset.Edit
      'Data1.Recordset.Fields("flag_deli") = "S"
      'Data1.Recordset.Update
      
      'Frame1.Visible = False
      'DBGrid2.SetFocus
      Exit Sub
   End If
   Exit Sub
   End If
   
   If opcion1 = "15" Then  'copia documento
   If MsgBox("Desea Sacar Copia del Documento", 1, "Aviso") <> 1 Then
      dbGrid1.SetFocus
      Exit Sub
   End If
   atipo = "" & rcconsulta.Fields("tipo")
   aserie = "" & rcconsulta.Fields("serie")
   anumero = "" & rcconsulta.Fields("numero")
   'impresion_sin_formato atipo, aserie, anumero
   proceso_impresion11 atipo, aserie, anumero, 1, "1"
   Frame1.Visible = False
   Frame1.Enabled = False
   DBGrid2.SetFocus
   Exit Sub

   End If
   If opcion1 = "100" Then  'anula documento
   If "" & rcconsulta.Fields("e") = "1" Then
      MsgBox "Documento Anulado ", 48, "Aviso"
      dbGrid1.SetFocus
      Exit Sub
   End If
   If MsgBox("Desea Anular Documento " + "" & rcconsulta.Fields("numero"), 1, "Aviso") <> 1 Then
      dbGrid1.SetFocus
      Exit Sub
   End If
      atipo = "" & rcconsulta.Fields("tipo")
      aserie = "" & rcconsulta.Fields("serie")
      anumero = "" & rcconsulta.Fields("numero")
      found = proceso_anular(atipo, aserie, anumero)
      If found = 1 Then
         proceso_impresion11 atipo, aserie, anumero, 0, ""
         If Trim("" & mytable11.Fields("hod")) = "S" And Trim(rcconsulta.Fields("tipo")) <> "C" Then 'enviar orden de despacho
            found = orden_despacho_n("" & mytable11.Fields("local"), atipo, aserie, anumero, "***ANULADO***")
         End If
         Frame1.Visible = False
         Frame1.Enabled = False
         DBGrid2.SetFocus
         Exit Sub
      End If
      dbGrid1.SetFocus
   End If
   Exit Sub
End If
'KeyCode = 0


End Sub
Function ir_ultimo_registrox()
On Error GoTo cmd7800_err
Data2.Recordset.MoveLast
ir_ultimo_registrox = 1
Exit Function
cmd7800_err:
Exit Function
End Function
Sub borrar_data1()
On Error GoTo cmd672222_err
Data1.Recordset.Delete
Exit Sub
cmd672222_err:
Exit Sub
End Sub
Function borra_data9()
On Error GoTo cmd9000_err
   Data9.Recordset.MoveLast
   Data9.Recordset.Delete
   Data9.refresh
   borra_data9 = 1
   Exit Function
cmd9000_err:
   Exit Function

End Function

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
Dim buf As String
Dim found As Integer
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
         found = sql_consulta(0)
         
End If
End Sub

Private Sub dbgrid1_KeyUp(KeyCode As Integer, Shift As Integer)
'-------------
Dim buf As String
Dim buf2 As String
Dim sw As Integer
Dim found As Integer
On Error GoTo cmd918_err
'MsgBox ""
If opcion1 = "8" Then
   'MsgBox "" & dbGrid1.Columns(1)
   pone_precios "" & rcconsulta.Fields("producto")
End If
'If KeyCode <> 13 And KeyCode <> 27 Then
'          If KeyCode = 32 Then
'             GoTo sigue9
'          End If
'          If KeyCode >= 48 And KeyCode <= 57 Then
'             GoTo sigue9
'          End If
'          If KeyCode >= 65 And KeyCode <= 90 Then
'             GoTo sigue9
'          End If
'          If KeyCode >= 97 And KeyCode <= 122 Then
'             GoTo sigue9
'          End If
'          If KeyCode = 8 Or Chr(KeyCode) = "*" Then
'             GoTo sigue9
'          End If
'          Exit Sub
'sigue9:
'          If KeyCode = 8 Then
'            If Len(buffer) > 0 Then
'               buf = Mid$(buffer, 1, Len(buffer) - 1)
'               buffer = buf
'               KeyCode = 0
'               Else
'               KeyCode = 0
'               Exit Sub
'            End If
'         End If
'         buf = Chr(KeyCode)
'         If Chr(KeyCode) = "*" Then
'            buf = ""
'            buffer = buf
'         End If
'         If KeyCode <> 13 Then
'            buffer = buffer + buf
'         End If
'
'         buf = buffer
'         found = sql_consulta(0)
'         If found = 0 Then
'            found = sql_consulta(1)
'         End If
'Exit Sub
'End If
Exit Sub

cmd918_err:
MsgBox "Aviso en dbgridKeyup " + error$, 48, "Aviso"
Exit Sub

End Sub


Private Sub DBGrid10_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ind As Integer
Dim found As Integer
On Error GoTo cmd8911_err
If KeyCode = &H2E Then  'borrar linea
   found = borra_data9()
   If found = 0 Then
      dbgrid10.Enabled = True
      dbgrid10.SetFocus
      Exit Sub
   End If
   Exit Sub
End If
If KeyCode <> 13 And KeyCode <> 27 Then Exit Sub
If KeyCode = 27 Then
   'losao94_Click
   Framefp.Visible = False
   habilita_lab7 1
   If flag_servicio = "C" Then
      inicialIzatodo
   End If
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus
   Exit Sub
End If
suma_fpagov
If Label45.Caption = "Vuelto" Or Val(stxtotals) = 0 Then
          'If MsgBox("Desea Grabar", 1, "Aviso") <> 1 Then Exit Sub
          'if len()
             xtipo = protipo
             If "" & mytable11.Fields("habilitanota") = "S" Then
                If Val(ttxtotals) <= Val("" & mytable11.Fields("siventa")) Then
                   xtipo = "5"
                End If
             End If
                  'xruc = codigo
                  'xnombre = nombre
                  xvendedor = cproven
                  Frame7.Visible = True
                  habilita_lab7 1
                  Framefp.Enabled = False
                  If flag_servicio <> "A" And flag_servicio <> "C" And flag_servicio <> "D" Then
                     xtipo = "5"
                  End If
                  xtipo.SetFocus
          Exit Sub
End If
saldoabo = ""
acufp = "" & dbgrid10.columns(3)
Frame6.Caption = "" & dbgrid10.columns(0)
fpago = "" & dbgrid10.columns(1)
fpmoneda = "" & dbgrid10.columns(2)
dbgrid10.Enabled = False
Frame3.Visible = False
RGPAGO = ""

               If fpmoneda = "S" Then
                  'RGPAGO = ttxtotals
               End If
               If fpmoneda = "D" Then
                  'RGPAGO = ttxtotald
               End If

'If "" & dbgrid10.columns(3) = "A" Or "" & dbgrid10.columns(3) = "B" Or "" & dbgrid10.columns(3) = "E" Or "" & dbgrid10.columns(3) = "U" Then  'efectivo,dolares,euros
'   macro_inserta_registro
'   DBGrid9.Row = DBGrid9.VisibleRows - 1
'   DBGrid9.Col = 2
'   DBGrid9.SetFocus
'   Exit Sub
'End If
If "" & dbgrid10.columns(3) = "A" Or "" & dbgrid10.columns(3) = "B" Or "" & dbgrid10.columns(3) = "E" Or "" & dbgrid10.columns(3) = "U" Then  'efectivo,dolares,euros
   Frame3.Visible = True
   macro_credito 5
   RGPAGO.SetFocus
   'tcampo1.SetFocus
End If


If "" & dbgrid10.columns(3) = "C" Then   'credito
   macro_credito 3
   RGPAGO.SetFocus
End If
If "" & dbgrid10.columns(3) = "D" Then   'tarejta credito
   macro_credito 4
   tcampo1.SetFocus
End If
If "" & dbgrid10.columns(3) = "F" Then   'TARJETA DEBITO
   macro_credito 5
   tcampo1.SetFocus
End If
If "" & dbgrid10.columns(3) = "G" Then   'letra
   macro_credito 0
   tcampo1.SetFocus
End If
If "" & dbgrid10.columns(3) = "H" Or "" & dbgrid10.columns(3) = "K" Then   'bancos
   macro_credito 2
   tcampo3.SetFocus
End If
If "" & dbgrid10.columns(3) = "V" Then   'vales
   macro_credito 6
   tcampo1.SetFocus
End If

If "" & dbgrid10.columns(3) = "I" Or "" & dbgrid10.columns(3) = "K" Then   'CRUCE CON ABONO EFECTIVO
   macro_credito 1
   tcampo1.Enabled = True
   tcampo1.SetFocus
End If
If "" & dbgrid10.columns(3) = "J" Then   'vales
   macro_credito 8
   tcampo1.SetFocus
End If

Exit Sub
cmd8911_err:
MsgBox error$, 48, "Aviso"
Exit Sub
End Sub

Private Sub DBGrid2_AfterColEdit(ByVal ColIndex As Integer)
Dim found As Integer
Select Case ColIndex
       Case 0
           If control_flujo = 1 Then
              found = sumar_detalle()
              DBGrid2.Col = 0
              DBGrid2.Row = DBGrid2.VisibleRows - 1
              DBGrid2.SetFocus
              control_flujo = 0
           End If
            
       Case 3
End Select
End Sub
 
Private Sub DBGrid2_AfterColUpdate(ByVal ColIndex As Integer)
Dim found As Integer
Dim sdx As Double
Select Case ColIndex
       Case 0
            'found = busca_producto("" & DBGrid2.Columns(0), 0)
            'If found = 0 Then
            '   MsgBox "No existe producto", 48, "Aviso"
            '   Exit Sub
            'End If
            'If control_flujo = 1 Then
            '   MsgBox "Hola"
            'End If
            'MsgBox "Hola"
            
            
            
            'MsgBox "" & dbgrid2.Columns(0)
            found = busca_remate("" & DBGrid2.columns(0))
            If found = 1 Then
               DBGrid2.Col = 5
               'ingreso_tallas "" & DBGrid2.Columns(17)
               Exit Sub
            End If
            '--------------------
            
            '--------------------
            
            If ver_si_puedo_dbgrid("" & DBGrid2.columns(0)) = 1 Then  'existe mas de un precio
               'MsgBox "abc"
               xproducto = "" & DBGrid2.columns(0)
               carga_dbgrid4 "" & DBGrid2.columns(0)
               swprecio = 1
               Exit Sub
            End If
            If Len(DBGrid2.columns(0)) > 0 And Len(DBGrid2.columns(17)) > 0 Then
               'DBGrid2.Col = 3
               'ingreso_tallas "" & DBGrid2.Columns(17)
               'Exit Sub
            End If
            '-------------------
            found = existe_fuel("" & DBGrid2.columns(0))
            If found = 1 Then
               'MsgBox ""
               DBGrid2.Col = 7
               DBGrid2.SetFocus
               Exit Sub
            End If
            '-------------------
            found = sumar_detalle()
            If found = 0 Then
               'If "" & mytable11.Fields("noprecio") = "S" Then
               '   MsgBox "Error en Precio<=0", 48, "Aviso"
                  DBGrid2.SetFocus
                  Exit Sub
               'End If
            End If
            If swprecio = 1 Then
               DBGrid2.Col = 2
               DBGrid2.Row = DBGrid2.VisibleRows - 2
               'DBGrid2.SetFocus
               DBGrid4.SetFocus
               Exit Sub
            End If
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
       Case 1
       Case 2
            'MsgBox "Hola"
            found = sumar_detalle()
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
       Case 3
            'sdx = Val("" & DBGrid2.Columns("cantidad")) * Val("" & DBGrid2.Columns("precio"))
            'DBGrid2.Columns(9) = Val(Format(sdx, nrodecimal))
            'DBGrid2.Columns("total") = Val(Format(sdx, nrodecimal))
            'calcula_igv
            'ir_ultimo
            'MsgBox "Hola"
            found = sumar_detalle()
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
       Case 5
            'sdx = Val("" & DBGrid2.Columns("cantidad")) * Val("" & DBGrid2.Columns("precio"))
            'DBGrid2.Columns(9) = Val(Format(sdx, nrodecimal))
            'DBGrid2.Columns("total") = Val(Format(sdx, nrodecimal))
            'calcula_igv
            found = sumar_detalle()
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
       Case 6
            'sdx = Val("" & DBGrid2.Columns("cantidad")) * Val("" & DBGrid2.Columns("precio"))
            'DBGrid2.Columns(9) = Val(Format(sdx, nrodecimal))
            'DBGrid2.Columns("total") = Val(Format(sdx, nrodecimal))
            'calcula_igv
            found = sumar_detalle()
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
       Case 7
            'If Val("" & DBGrid2.Columns("cantidad")) > 0 Then
            '   sdx = Val("" & DBGrid2.Columns("total")) / Val("" & DBGrid2.Columns("cantidad"))
            '   DBGrid2.Columns("precio") = Val(Format(sdx, nrodecimal))
            '   DBGrid2.Columns(9) = Val("" & DBGrid2.Columns("total"))
            '   calcula_igv
            found = sumar_detalle()
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
            'End If
End Select
End Sub

Private Sub DBGrid2_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Dim found As Integer
If ColIndex >= 14 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
Case 1, 2, 4, 8, 9, 10, 12, 11
     Cancel = True
     Exit Sub
Case 0
     If Len("" & DBGrid2.columns(0)) > 0 Then  'si ya existe no cambiar
        Cancel = True
        Exit Sub
     End If
     
     'If opcion5 = 20 Then
     '   MsgBox "Hola"
     '   Cancel = True
     '   Exit Sub
     'End If
     'opcion5 = 0
     
Case 2
     If Len("" & DBGrid2.columns(0)) = 0 Then  '
        Cancel = True
        Exit Sub
     End If

Case 3
     If Len("" & DBGrid2.columns(0)) = 0 Then  '
        Cancel = True
        Exit Sub
     End If
     'If Len("" & DBGrid2.Columns(17)) > 0 Then  'ojo no se puede poner si es talla
     '   Cancel = True
     '   Exit Sub
     'End If
     Case 5, 7, 13, 6
     If Len("" & DBGrid2.columns(0)) = 0 Then  '
        Cancel = True
        Exit Sub
     End If
     'MsgBox ""
     
End Select
End Sub

Private Sub DBGrid2_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim found As Integer
Dim sdx As Double
Dim xcampo As String
Dim canti As String
Dim buf1 As String
Dim buf As String
Dim bufy As String
Dim amount As String
Dim xfound As String
Dim xnbufx As Double
Select Case ColIndex
Dim mytablex As New ADODB.Recordset
Case 0
     If Len(DBGrid2.columns(0)) = 0 Then
        'aqui vamos a valida si es el fin del pedido
        Cancel = True
        Exit Sub
     End If
     If Len(DBGrid2.columns(0)) > 15 Then
        'aqui vamos a valida si es el fin del pedido
        Cancel = True
        Exit Sub
     End If
     
     
stkminimo = ""
If "" & mytable11.Fields("stkminimo") = "S" Then
            mytablex.Open "SELECT * FROM producto where producto='" & "" & DBGrid2.columns(0) & "'", cn, adOpenDynamic, adLockOptimistic
            If mytablex.RecordCount > 0 Then
               If familia_saldo("" & mytablex.Fields("familia")) = 0 Then
                  consulta_minimo "" & mytablex.Fields("producto"), "" & mytablex.Fields("minimo")
               End If
            End If
            mytablex.Close
End If
     
     
     If "" & mytable11.Fields("nosaldo") = "S" Then
            found = verifica_saldo_receta("" & DBGrid2.columns(0), 1)
            If found = 2 Then
               MsgBox "Se detecto un saldo receta con saldo<=0 ", 48, "Aviso"
            Exit Sub
            End If
            If familia_saldo("" & DBGrid2.columns("familia")) = 0 Then
            If verifica_producto("" & DBGrid2.columns(0)) = 1 Then
               If consulta_saldo("" & DBGrid2.columns(0), 1, 0) <= 0 Then
                  Cancel = True
                  DBGrid2.SetFocus
                  MsgBox "x.No existe saldo", 48, "Aviso"
                  Exit Sub
               End If
            End If
            End If
     End If
     canti = ""
     buf = UCase(DBGrid2.columns(0))  'se modifico en U. Union
     bufy = buf
     found = 0
     If "" & mytable11.Fields("flag") = "*" Then
        found = InStr(buf, "" & mytable11.Fields("flag"))
        If found > 1 Then  ' si es cantidad
                  xcampo = Mid$(buf, found + 1, Len(buf) - found)
                  canti = Mid$(buf, 1, found - 1)
                  buf1 = Val(canti)
                  buf = xcampo
                  If Len(buf) = 0 Then
                     Cancel = True
                     Exit Sub
                  End If
                  DBGrid2.columns(0) = buf
        End If
     End If
     'MsgBox buf
     'found = verifica_doble("" & DBGrid2.Columns(0))
     'If found = 1 Then
     '   Cancel = True
     '   MsgBox "Producto ya Seleccionado", 48, "Aviso"
     '   Exit Sub
     'End If
     '----validamos el saldo
         
     control_flujo = 0
     found = busca_producto(UCase("" & DBGrid2.columns(0)), 0, canti, 0)
     'found = busca_producto(buf, 0, canti)
     If found = 2 Then  'si es precio 0
        Cancel = True
        control_flujo = 1
        'MsgBox "No se pude realiza Operacion,continue..", 48, "Aviso"
        'DBGrid2.SetFocus
        Exit Sub
     End If
     If found = 0 Then
        Cancel = True
        'MsgBox "No existe producto", 48, "Aviso"
        'consulta_producto "" & DBGrid2.Columns(0)
        opcion5 = 1
        found = consulta_producto(bufy)
        If found = 1 Then
           Cancel = True
           opcion5 = 20
           MsgBox "No existe producto", 48, "Aviso"
           DBGrid2.SetFocus
           'opcion5 = 20
           'DBGrid2.Col = 0
           'DBGrid2.Row = dbgrid2.visiblerows - 1
           'DBGrid2.SetFocus
           Exit Sub
        End If
        opcion5 = 0
        Exit Sub
     End If
     buf = ""
     'If "" & mytable11.Fields("actbala") = "S" Then
     'If verifica_balanza("" & DBGrid2.Columns(0)) = "S" Then
        
'ajk9:
 '       buf = puerto_balanza1()
 '       If Val(buf) = 0 Then
 '          If MsgBox("Balanza No leido,Continua Leyendo? ", 1, "Aviso") = 1 Then
 '             GoTo ajk9
 '             '------
 '             Else
 '
 '          End If
 '       End If
 '    End If
 '    End If
 '
     'If Val(buf) > 0 Then
     '----pro favor verficia
        'DBGrid2.Columns("cantidad") = Val(Mid$(Val(buf), 1, 5))
        'sdx = Val("" & DBGrid2.Columns("cantidad")) * Val("" & DBGrid2.Columns("precio"))
        'DBGrid2.Columns("total") = sdx
        'calcula_igv 0
     '-------------------
     'End If
     swprecio = 0
     Exit Sub
Case 2
     If Len(DBGrid2.columns(0)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Mid$("" & DBGrid2.columns(2), 1, 1) = "-" And Len("" & DBGrid2.columns(2)) > 1 Then
        'grabar_foto "" & Value
        Exit Sub
     End If
     found = valida_placa("" & DBGrid2.columns(17), Mid$("" & DBGrid2.columns(2), 1, 1))
     If found = 0 Then
        MsgBox "Placa invalida ", 48, "Aviso"
        Cancel = True
        Exit Sub
     End If
Case 3
     If Len(DBGrid2.columns(0)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Not IsNumeric("" & DBGrid2.columns("cantidad")) Then
        Cancel = True
        Exit Sub
     End If
     If Val(DBGrid2.columns("cantidad")) = 0 Then
        MsgBox "Cant=0", 48, "Aviso"
        Cancel = True
        Exit Sub
     End If
     
     If Val(DBGrid2.columns("cantidad")) < 0 Then  'devolucion
        flag_clave1 = 0
        tconcla.X = "N"
        tconcla.Show 1
        If flag_clave1 <> 1 Then  'si es descongela
           Cancel = True
           Exit Sub
        End If
        'MsgBox "Cant=0", 48, "Aviso"
        'Cancel = True
        'Exit Sub
     End If
     'MsgBox Val("" & DBGrid2.Columns("cantidad"))
     
stkminimo = ""
If "" & mytable11.Fields("stkminimo") = "S" Then
            mytablex.Open "SELECT * FROM producto where producto='" & "" & DBGrid2.columns(0) & "'", cn, adOpenDynamic, adLockOptimistic
            If mytablex.RecordCount > 0 Then
               If familia_saldo("" & mytablex.Fields("familia")) = 0 Then
                  consulta_minimo "" & mytablex.Fields("producto"), "" & mytablex.Fields("minimo")
               End If
            End If
            mytablex.Close
End If


     
     
     If "" & mytable11.Fields("nosaldo") = "S" Then
     
            found = verifica_saldo_receta("" & DBGrid2.columns(0), Val(DBGrid2.columns("cantidad")) * Val(DBGrid2.columns(4)))
            If found = 2 Then
               MsgBox "Se detecto un saldo receta con saldo<=0 ", 48, "Aviso"
            Exit Sub
            End If
    
     
            If familia_saldo("" & DBGrid2.columns("familia")) = 0 Then
            If consulta_saldo("" & DBGrid2.columns(0), Val("" & DBGrid2.columns("cantidad")) * Val("" & DBGrid2.columns(4)), 1) <= 0 Then
               Cancel = True
               DBGrid2.SetFocus
               MsgBox "No existe saldo Suficiente", 48, "Aviso"
               Exit Sub
            End If
            End If
     End If
            found = busca_unidad("" & DBGrid2.columns(0))
            If found = 1 Then
               amount = Format(Val("" & DBGrid2.columns("cantidad")), nrodecimal)
               If Val(Mid$(amount, Len(amount) - 1, 2)) > 0 Then
                  MsgBox "Solo Datos Exactos", 24, "LO SENTIMOS"
                  Cancel = True
                  Exit Sub
               End If
            End If
     'VERIFICAMOS SI ES CANTIDAD para poner que precio debe tener
     xnbufx = 0
     If "" & DBGrid2.columns("nroprecio") = "1" Then  'si me encuentro en el precio 1
         If Val("" & DBGrid2.columns("cantidad")) >= 1 Then
            xnbufx = 0
            found = verifica_ofertax("" & DBGrid2.columns(0), Val("" & DBGrid2.columns("cantidad")), xnbufx)
            If found = 1 Then
               DBGrid2.columns("precio") = xnbufx
            End If
         End If
     End If
     'MsgBox "xx"
     sdx = Val("" & DBGrid2.columns("cantidad")) * Val("" & DBGrid2.columns("precio"))
     'MsgBox sdx
     DBGrid2.columns("total") = sdx
     calcula_igv 0
Case 5
     If Len(DBGrid2.columns(0)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Not IsNumeric(DBGrid2.columns("precio")) Then
        Cancel = True
        Exit Sub
     End If
     'MsgBox "hola"
     xfound = verifica_oferta("" & DBGrid2.columns(0))
     If xfound <> "S" Then   '
        If Val(DBGrid2.columns("precio")) <= 0 Then
        If "" & mytable11.Fields("noprecio") = "S" Then
           MsgBox "Precio <=0", 48, "Aviso"
           Cancel = True
           Exit Sub
        End If
        End If
        'MsgBox "hello"
        If "" & mytable11.Fields("obligaprecio") = "S" Then
           flag_clave1 = 0
           tconcla.X = "S"
           tconcla.Show 1
           If flag_clave1 = 0 Then  'si es descongela
              Cancel = True
              Exit Sub
           End If
        End If
        
        'MsgBox found
        found = valida_rango()
        If found = 0 Then
           Cancel = True
           Exit Sub
        End If
     End If
     sdx = Val("" & DBGrid2.columns("cantidad")) * Val("" & DBGrid2.columns("precio"))
     DBGrid2.columns("total") = sdx
     calcula_igv 0
Case 6
     If Len(DBGrid2.columns(0)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     
     If Not IsNumeric(DBGrid2.columns("deslipo")) Then
        Cancel = True
        Exit Sub
     End If
     sdx = Val("" & DBGrid2.columns("cantidad")) * Val("" & DBGrid2.columns("precio"))
     DBGrid2.columns("total") = sdx
     calcula_igv 0
Case 7
     If Len(DBGrid2.columns(0)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Not IsNumeric(DBGrid2.columns("total")) Then
        Cancel = True
        Exit Sub
     End If
     If Val("" & DBGrid2.columns("cantidad")) = 0 Then
        Cancel = True
        Exit Sub
     End If
     'xfound = verifica_oferta("" & DBGrid2.Columns(0))
     'if xfound<>"S" then
     If Val(DBGrid2.columns("total")) <= 0 Then
        If "" & mytable11.Fields("noprecio") = "S" Then
        MsgBox "Precio <=0", 48, "Aviso"
        Cancel = True
        Exit Sub
        End If
     End If
    found = existe_fuel("" & DBGrid2.columns(0))
    If found <> 1 Then
     flag_clave1 = 0
     tconcla.X = "S"
     tconcla.Show 1
     If flag_clave1 = 0 Then  'si es descongela
        Cancel = True
        Exit Sub
     End If
     sdx = Val("" & DBGrid2.columns("total")) / Val("" & DBGrid2.columns("cantidad"))
     DBGrid2.columns("precio") = sdx
     calcula_igv 0
    End If
    If found = 1 Then
        If Val("" & DBGrid2.columns("precio")) = 0 Then
           Cancel = True
           Exit Sub
        End If
           sdx = Val("" & DBGrid2.columns("total")) / Val("" & DBGrid2.columns("precio"))
           DBGrid2.columns("cantidad") = sdx
           calcula_igv 0
    End If
Case 13
     If Len(DBGrid2.columns(0)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     If Not IsNumeric(DBGrid2.columns("neto")) Then
        Cancel = True
        Exit Sub
     End If
     If Val("" & DBGrid2.columns("cantidad")) = 0 Then
        Cancel = True
        Exit Sub
     End If
     calcula_sinigv
     'calcula_igv 1
     

End Select
End Sub

Private Sub DBGrid2_BeforeDelete(Cancel As Integer)
   'If MsgBox("¿Realmente desea eliminar el registro ", 1, "Confirmación de eliminación") <> 1 Then
   'Cancel = True
   'Exit Sub
   'End If
End Sub

Private Sub DBGrid2_ColEdit(ByVal ColIndex As Integer)
Dim sdx As Double
Select Case ColIndex
       Case 0
       Case 3
            
End Select
End Sub


Private Sub DBGrid2_DblClick()
Dim found As Integer
Select Case DBGrid2.Col
       Case 3
       If Val("" & DBGrid2.columns("cantidad")) > 0 Then
          tkeyboar.flag = "CANTIDAD"
          tkeyboar.Show 1
       End If
       Case 5
       If Val("" & DBGrid2.columns("cantidad")) > 0 Then
          tkeyboar.flag = "PRECIO"
          tkeyboar.Show 1
       End If

End Select
found = sumar_detalle()
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim found As Integer
On Error GoTo cmd341231_err
'Exit Sub

If KeyCode = &H71 Then  'f2 totales
    If xopciones(0).Enabled = False Then Exit Sub
       xopciones_Click 0
       'Label55_Click
    Exit Sub
 End If
 If KeyCode = &H78 Then  'exonerado
    If xopciones(0).Enabled = False Then Exit Sub
       proceso_cierre_automatico
    Exit Sub
 End If

If KeyCode = 13 Then
   'If Len(DBGrid2.Columns(0)) = 0 Then
   '   DBGrid2.Col = 0
   '   DBGrid2.Row = dbgrid2.visiblerows - 1
   '   DBGrid2.SetFocus
   '   Exit Sub
   'End If
   
Select Case DBGrid2.Col
       Case 0
            If Len("" & DBGrid2.columns(0)) = 0 Then
               If xopciones(0).Enabled = False Then Exit Sub
               xopciones_Click 0
               'If Label55.Enabled = False Then
               '   Exit Sub
               'End If
               'Label55_Click
               Exit Sub
            End If
            'If Len("" & DBGrid2.Columns(0)) > 0 Then
            '   DBGrid2.Col = 2
            'End If
       Case 3
            'If Val("" & DBGrid2.Columns("precio")) = 0 Then
            '   DBGrid2.Col = 5
            '   Exit Sub
            'End If
            found = sumar_detalle()
            
            KeyCode = 0
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
       Case 5
            'If Len("" & DBGrid2.Columns(4)) > 0 Then
            '   DBGrid2.Col = 6
            'End If
            found = sumar_detalle()
            KeyCode = 0
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
       Case 7
            found = sumar_detalle()
            KeyCode = 0
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
            
End Select
'KeyCode = vbKey0
End If
Exit Sub
cmd341231_err:
Exit Sub
End Sub

Private Sub dbgrid2_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   'DBGrid1_KeyDown 0, 0
'   MsgBox "hOLA"'
'
'End If
End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim found As Integer
On Error GoTo cmd34_err
carga_grafico "" & Data2.Recordset.Fields("producto")
carga_minimo "" & Data2.Recordset.Fields("producto")

If opcion5 = 20 Then 'SI NO EXISTE PRODUCTOS
   'Data2.Refresh
   'found = sumar_detalle()
   'If Data2.Recordset.EOF Or Data2.Recordset.BOF Then
   '   Data2.Refresh
   '   'Exit Sub
   'End If
   found = ir_ultimo_registrox()
   If found = 0 Then
      opcion5 = 0
      Data2.refresh
      Exit Sub
   End If
   Data2.refresh
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus 'found = sumar_detalle()
   'DBGrid2.SetFocus
   opcion5 = 0
   Exit Sub
End If
If KeyCode = 13 Then
If Len(DBGrid2.columns(0)) = 0 Then
      DBGrid2.Col = 0
      DBGrid2.Row = DBGrid2.VisibleRows - 1
      DBGrid2.SetFocus
      Exit Sub
   End If
End If
 If KeyCode = &H72 Then  'f3
    codigo.SetFocus
    Exit Sub
 End If
 
 If KeyCode = &H70 Then  'f1  carga los demas precios
   If Len(DBGrid2.columns(0)) > 0 And DBGrid2.Col = 2 Then
      xproducto = "" & DBGrid2.columns(0)
      carga_dbgrid4 "" & DBGrid2.columns(0)
      Exit Sub
   End If
End If
If KeyCode = &H76 Then  'f7
   flag_clave1 = 0
   tconcla.X = "N"
   tconcla.Show 1
   If flag_clave1 <> 1 Then  'si es descongela
      DBGrid2.SetFocus
      Exit Sub
   End If

   xprodet.Show 1
   DBGrid2.SetFocus
End If
If KeyCode = 13 Then
   KeyCode = 0
End If
If KeyCode = &H2E Then  'borrar linea
If DBGrid2.Row = -1 Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
End If
If MsgBox("Se va a eliminar el registro : está seguro ", _
   vbExclamation + vbYesNo, "Eliminar") = vbYes Then
   Data2.Recordset.Delete
   If Data2.Recordset.EOF = True And Data2.Recordset.BOF = True Then
      Exit Sub
   End If
   found = sumar_detalle()
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus
   Exit Sub
End If
End If
If KeyCode = &H70 Then  'f1
   If Len(DBGrid2.columns(0)) = 0 Then
      found = consulta_producto("")
   End If
End If
If KeyCode = &H72 Then  'f3
   'If Len(DBGrid2.Columns(0)) > 0 And Len(DBGrid2.Columns(17)) > 0 Then
   '   ingreso_tallas "" & DBGrid2.Columns(17)
   'End If
   
End If
If KeyCode = &H77 Then  'f8 OBSERVACIONES
   If Len(DBGrid2.columns(0)) > 0 Then
      ingreso_locales
   End If
End If
If KeyCode = &H28 Then  'flecha abajo inserta una nueva
         Exit Sub
         If DBGrid2.Col = 0 Then
            ir_ultimo
            If Len(DBGrid2.columns(0)) > 0 And Len(DBGrid2.columns(1)) > 0 And Len(DBGrid2.columns(2)) > 0 And Len(DBGrid2.columns("cantidad")) > 0 And Len(DBGrid2.columns(4)) > 0 And Len(DBGrid2.columns("precio")) > 0 Then
            'Data2.Recordset.AddNew
            'Data2.Recordset.Update
            End If
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
         End If
End If
Exit Sub
cmd34_err:
Exit Sub
End Sub

Private Sub dbgrid3_Click()

End Sub

Private Sub DBGrid4_DblClick()
DBGrid4_KeyDown 13, 0
End Sub

Private Sub DBGrid4_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sdx As Double
Dim found As Integer
Dim xpreciox As Double
Dim mytablex As New ADODB.Recordset
If KeyCode = 27 Then
   If opcion3 = "1" Then
      Frame5.Visible = False
      dbGrid1.SetFocus
      Exit Sub
   End If
   If opcion1 = "8" Then
      Frame5.Visible = False
      Frame1.Enabled = True
      dbGrid1.Enabled = True
      If dbGrid1.Visible = True Then
         dbGrid1.Visible = True
         dbGrid1.Enabled = True
         dbGrid1.SetFocus
      End If
      If dbGrid1.Visible = False Then
         DBGrid2.Enabled = True
         DBGrid2.SetFocus
      End If
      Exit Sub
   End If
   Frame5.Visible = False
   'Data2.Refresh
   found = sumar_detalle()
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus
   'DBGrid2.SetFocus
   'Command8_Click
   Exit Sub
End If
If KeyCode = 13 Then
   If Len("" & DBGrid4.columns(1)) = 0 Or Len("" & DBGrid4.columns(0)) = 0 Then
      DBGrid4.SetFocus
      Exit Sub
   End If
   'MsgBox opcion1
   'MsgBox opcion1
   
   stkminimo = ""
If "" & mytable11.Fields("stkminimo") = "S" Then
            mytablex.Open "SELECT * FROM producto where producto='" & "" & DBGrid2.columns(0) & "'", cn, adOpenDynamic, adLockOptimistic
            If mytablex.RecordCount > 0 Then
               If familia_saldo("" & mytablex.Fields("familia")) = 0 Then
                  consulta_minimo "" & mytablex.Fields("producto"), "" & mytablex.Fields("minimo")
               End If
            End If
            mytablex.Close
End If

   
   
         If "" & mytable11.Fields("nosaldo") = "S" Then
         
         found = verifica_saldo_receta("" & DBGrid2.columns(0), Val(DBGrid4.columns(1)))
            If found = 2 Then
               MsgBox "Se detecto un saldo receta con saldo<=0 ", 48, "Aviso"
            Exit Sub
            End If
         
         
         If familia_saldo("" & DBGrid2.columns("familia")) = 0 Then
            If consulta_saldo("" & DBGrid2.columns(0), Val("" & DBGrid4.columns(1)), 1) <= 0 Then
               MsgBox "No existe saldo", 48, "Aviso"
               DBGrid4.SetFocus
               Exit Sub
            End If
         End If
         End If
   If Frame1.Visible = True Then
      Frame5.Visible = False
      Frame1.Enabled = True
      dbGrid1.Enabled = True
      dbGrid1.SetFocus
      Exit Sub
   End If
   If opcion3 = "1" Then
      Frame5.Visible = False
      Frame1.Enabled = True
      dbGrid1.Enabled = True
      dbGrid1.SetFocus
      Exit Sub
   End If
   'If Val("" & DBGrid4.Columns(2)) <= 0 Then
   '   MsgBox "Precio<=0", 48, "Aviso"
   '   DBGrid4.SetFocus
   '   Exit Sub
   'End If
   '---------------validar precios-----------------------------
   xpreciox = 0
   xpreciox = Val("" & DBGrid4.columns(2))
   'If opcion1 = "8" Then
   'If Len("" & DBGrid4.Columns(0)) > 0 And Val("" & DBGrid4.Columns(1)) > 0 And Len("" & DBGrid4.Columns(2)) > 0 Then
      'Data2.Recordset.Edit
      'Data2.Recordset.Fields("unidad") = "" & DBGrid4.Columns(0)
      'Data2.Recordset.Fields("factor") = "" & DBGrid4.Columns(1)
      'Data2.Recordset.Fields("precio") = "" & DBGrid4.Columns("cantidad")
      'Data2.Recordset.Update
      'MsgBox DBGrid4.Row
      DBGrid2.columns("nroprecio") = "" & (DBGrid4.Row + 1)
      DBGrid2.columns(2) = "" & DBGrid4.columns(0)
      DBGrid2.columns(4) = Val("" & DBGrid4.columns(1))
      DBGrid2.columns("precio") = xpreciox
      sdx = Val("" & DBGrid2.columns("cantidad")) * Val("" & DBGrid2.columns("precio"))
      DBGrid2.columns("total") = sdx
      calcula_igv 0
      'found = sumar_detalle()
      Frame5.Visible = False
      'antes estaba para que se vaya al final
      If Len(DBGrid2.columns(0)) > 0 And Len(DBGrid2.columns(17)) > 0 Then
         'DBGrid2.Col = 3
         'DBGrid2.SetFocus
         'ingreso_tallas "" & DBGrid2.Columns(17)
          Else
          'que vaya a cantidad
          'DBGrid2.Col = 3
          'DBGrid2.SetFocus
          sumar_reforzar
          
          'cuando necesite que vaya a la siguiente linea
          found = sumar_detalle()
          DBGrid2.Col = 0
          DBGrid2.Row = DBGrid2.VisibleRows - 1
          DBGrid2.SetFocus
      End If

      'Command8_Click
   'End If
  'End If
End If
End Sub

Private Sub DBGrid4_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim dR As Integer
Dim row_num As Integer
Dim R As Integer
Dim rows_returned As Integer
If ReadPriorRows Then
        dR = -1
    Else
        dR = 1
    End If
    If IsNull(StartLocation) Then
        If ReadPriorRows Then
           row_num = RowBuf.RowCount - 1
           'row_num = 9
        Else
           row_num = 0
        End If
    Else
        row_num = CLng(StartLocation) + dR
    End If
    rows_returned = 0
    For R = 0 To RowBuf.RowCount - 1
        If row_num < 0 Or row_num > 9 Then Exit For
        RowBuf.Value(R, 0) = campo_precios(row_num).unidad
        RowBuf.Value(R, 1) = campo_precios(row_num).factor
        RowBuf.Value(R, 2) = campo_precios(row_num).precio
        RowBuf.Value(R, 3) = campo_precios(row_num).costo
        RowBuf.Value(R, 4) = campo_precios(row_num).margen
        RowBuf.Value(R, 5) = campo_precios(row_num).stock
        RowBuf.Bookmark(R) = row_num
        row_num = row_num + dR
        rows_returned = rows_returned + 1
   Next R
   RowBuf.RowCount = rows_returned
End Sub






Private Sub DBGrid6_KeyPress(KeyAscii As Integer)
If KeyAscii <> 27 Then Exit Sub
dbgrid6.Visible = False
dbGrid1.SetFocus
End Sub

Private Sub DBGrid9_AfterColUpdate(ByVal ColIndex As Integer)
Select Case ColIndex
       Case 2
          suma_fpagov
          If Label45.Caption = "Vuelto" Or Val(stxtotals) <= 0 Then
             xtipo = protipo
             If "" & mytable11.Fields("habilitanota") = "S" Then
                If Val(ttxtotals) <= Val("" & mytable11.Fields("siventa")) Then
                   xtipo = "5"
                End If
             End If
                  xruc = codigo
                  xnombre = nombre
                  xvendedor = cproven
             Framefp.Enabled = False
             Frame7.Visible = True
             habilita_lab7 1
             Framefp.Enabled = False
             If flag_servicio <> "A" And flag_servicio <> "C" And flag_servicio <> "D" Then
                xtipo = "5"
             End If
             xtipo.SetFocus
          Exit Sub
         End If
         dbgrid10.Enabled = True
         dbgrid10.SetFocus
End Select
End Sub

Private Sub DBGrid9_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Dim found1 As Double
If ColIndex <> 2 Then
   Cancel = True
   Exit Sub
End If
Select Case ColIndex
       Case 2
            If Len("" & DBGrid9.columns(0)) = 0 Then
               Cancel = True
               Exit Sub
            End If
End Select
End Sub


Private Sub DBGrid9_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Dim found1 As Double
Select Case ColIndex
       Case 2
            If Not IsNumeric("" & DBGrid9.columns(2)) Then
               Cancel = True
               Exit Sub
            End If
            'If "" & Data9.Recordset.Fields("acu") = "H" Then 'valida el deposito bancario
            '   If Val("" & DBGrid9.Columns(2)) > Val(stxtotals) Then
            '      MsgBox "Debe ingresar valor exacto", 48, "Aviso"
            '      Cancel = True
            '      Exit Sub
            '   End If
            '   found1 = valida_deposito("" & Data9.Recordset.Fields("codigo"), "" & Data9.Recordset.Fields("orden"), 0)
            '   If found1 < Val("" & DBGrid9.Columns(2)) Then
            '      MsgBox "No existe Saldo ", 48, "Aviso"
            '      Cancel = True
            '      Exit Sub
            '   End If
            'End If
            
            If verifica_fpago("" & DBGrid9.columns("fpago")) = "V" Then
               'MsgBox "" & Data9.Recordset.Fields("orden") & "" & Data9.Recordset.Fields("observa") & "" & Data9.Recordset.Fields("dias")
               'MsgBox codigo

               'found1 = suma_pedidos("" & codigo, "" & tcampo3, "" & tcampo4, "" & tcampo5)
               found1 = suma_pedidos("" & codigo, "" & Data9.Recordset.Fields("orden"), "" & Data9.Recordset.Fields("observa"), "" & Data9.Recordset.Fields("dias"))
               'MsgBox found1
               ', "" & Data9.Recordset.Fields("orden"), "" & Data9.Recordset.Fields("observa"), "" & Data9.Recordset.Fields("dias")
               If found1 <= 0 Then
                  MsgBox "Cantidad Mayor que el saldo del pedido ", 48, "Aviso"
                  Cancel = True
                  Exit Sub
               End If
               'MsgBox found1
               If found1 > 0 Then
                  If found1 < Val("" & DBGrid9.columns(2)) Then
                     MsgBox "Cantidad Mayor que el saldo del pedido ", 48, "Aviso"
                     Cancel = True
                     Exit Sub
                  End If
               End If
            End If
            
            If "" & Data9.Recordset.Fields("acu") = "I" Or "" & Data9.Recordset.Fields("acu") = "K" Then 'valida el deposito bancario
               If Val("" & DBGrid9.columns(2)) > Val(stxtotals) Then
                  MsgBox "Debe ingresar valor exacto", 48, "Aviso"
                  Cancel = True
                  Exit Sub
               End If
               found1 = busca_credito_adelanto1("" & Data9.Recordset.Fields("codigo"), "" & Data9.Recordset.Fields("acu"))
               If found1 <= 0 Then
                  MsgBox "No existe Saldo ", 48, "Aviso"
                  Cancel = True
                  Exit Sub
               End If
               If found1 < Val("" & DBGrid9.columns(2)) Then
                  MsgBox "Saldo actual es: " & found1 & " Debe Ingresar dicha cantidad ", 48, "Aviso"
                  Cancel = True
                  Exit Sub
               End If
            End If
            

            opcion2 = 0
            '---------- validamos a donde va
            'valida_ingresado
End Select
End Sub

Private Sub DBGrid9_KeyDown(KeyCode As Integer, Shift As Integer)
Dim found1 As Double
On Error GoTo cmd7811_err
If KeyCode = 27 Then
   Data9.Recordset.Delete
   dbgrid10.Enabled = True
   dbgrid10.SetFocus
   Exit Sub
End If
'MsgBox Shift
If KeyCode = 13 Then
   Select Case DBGrid9.Col
       Case 2
            If Len("" & DBGrid9.columns(2)) > 0 Then Exit Sub
            If Val("" & DBGrid9.columns(2)) = 0 Then
                'If "" & Data9.Recordset.Fields("acu") = "H" Then 'valida el deposito bancario
                '   DBGrid9.SetFocus
                '   Exit Sub
                'End If
                If verifica_fpago("" & DBGrid9.columns("fpago")) = "V" Then
                  found1 = suma_pedidos("" & Data9.Recordset.Fields("codigo"), "" & Data9.Recordset.Fields("orden"), "" & Data9.Recordset.Fields("observa"), "" & Data9.Recordset.Fields("dias"))
               If found1 <= 0 Then
                  MsgBox "Cantidad Mayor que el saldo del pedido ", 48, "Aviso"
                  DBGrid9.SetFocus
                  Exit Sub
               End If
               If found1 > 0 Then
                  If found1 < Val(stxtotals) Then
                     MsgBox "Cantidad Mayor que el saldo del pedido ", 48, "Aviso"
                     DBGrid9.SetFocus
                     Exit Sub
                  End If
               End If
               End If
                
                
                
               If "" & Data9.Recordset.Fields("moneda") = "S" Then
                  Data9.Recordset.Edit
                  Data9.Recordset.Fields("recibe") = Val(stxtotals)
                  Data9.Recordset.Update
               End If
               If "" & Data9.Recordset.Fields("moneda") = "D" Then
                  Data9.Recordset.Edit
                  Data9.Recordset.Fields("recibe") = Val(stxtotald)
                  Data9.Recordset.Update
               End If
               opcion2 = 0
               'valida_ingresado
               
               suma_fpagov
               
               If Label45.Caption = "Vuelto" Or Val(stxtotals) <= 0 Then
                  xtipo = protipo
                  xvendedor = cproven
                  xruc = codigo
                  If "" & mytable11.Fields("habilitanota") = "S" Then
                     If Val(ttxtotals) <= Val("" & mytable11.Fields("siventa")) Then
                        xtipo = "5"
                     End If
                  End If
                  
                  xnombre = nombre
                  Frame7.Visible = True
                  habilita_lab7 1
                  Framefp.Enabled = False
                  If flag_servicio <> "A" And flag_servicio <> "C" And flag_servicio <> "D" Then
                     xtipo = "5"
                  End If
                  xtipo.SetFocus
               Exit Sub
               End If
             End If
   End Select
End If
Exit Sub
cmd7811_err:
Exit Sub
End Sub


Private Sub dcaj8923_Click()
End Sub

Private Sub dcodigo_DblClick()
tkeyboar.flag = "DCODIGO"
tkeyboar.Show 1

End Sub

Private Sub dcodigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(dcodigo) > 0 Then
   If Len(telefono) < 6 Then
      telefono.SetFocus
      Exit Sub
   End If
   found = busca_codigod()
End If
dnombre.SetFocus
End Sub

Private Sub dcodigo_KeyUp(KeyCode As Integer, Shift As Integer)
Dim found As Integer
If KeyCode = &H26 Then
   telefono.SetFocus
   Exit Sub
End If
If KeyCode = &H76 Then  'f7 creacion
   'para crear un cliente nuevo
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_cliente ""
End If
End Sub

Private Sub dcrt6622_Click()
Dim found As Integer
If dbgrid6.Visible = True Then Exit Sub
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
'-------------------------------------
   cgusuario = gocabeza
   dgusuariog = godetalle
   found = sumar_detalle()
   If Val(txtotal) > 0 Then
      MsgBox "No deben existir pedidos Pendientes", 48, "Aviso"
      DBGrid2.SetFocus
      Exit Sub
   End If
   menu_carga_doc_anterior
   Exit Sub

End Sub
Sub menu_carga_doc_anterior()
Dim found As Integer
   Combo1.Clear
   Combo1.AddItem "Nombre"
   Combo1.ListIndex = 0
sw_consulta = 0
Frame1.Visible = True
Frame1.Enabled = True
buffer = ""

opcion1 = "1500"
found = sql_consulta(1)
'dbgrid1.SetFocus
End Sub

Private Sub dcupar1_Click()
Dim sw As Integer
Dim found As Integer
flag_clave1 = 0
tconcla.X = "CUADRE"  'cuadre parcial
tconcla.Show 1
If flag_clave1 <> 1 Then  'si es descongela
   DBGrid2.SetFocus
   Exit Sub
End If

    opcion2 = "1"
    opcion1 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True

    usuariopos = gusuario
    tcuadrc1.cajero = "" & cajero
    tcuadrc1.caja = "" & caja
    tcuadrc1.turno = "" & turno
    tcuadrc1.fechai = "" & dia
    tcuadrc1.fechaf = "" & dia
    tcuadrc1.horai = "01"
    'tcuadrc1.todos = "S"
    tcuadrc1.horaf = "24"
    tcuadrc1.Caption = "CUADRE PARCIAL DEL DIA"
    tcuadrc1.Show 1
End Sub

Private Sub dcvendedor_Click()
If dcvendedor <> "%" Then
   Data2.Recordset.Edit
   dvendedor = extra_loquesea(dcvendedor)
   Data2.Recordset.Update
   
End If
End Sub

Private Sub ddireccion_DblClick()
tkeyboar.flag = "DDIRECCION"
tkeyboar.Show 1

End Sub

Private Sub ddireccion_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(ddireccion) = 0 Then
   ddireccion.SetFocus
   Exit Sub
End If
referencia.SetFocus
End Sub

Private Sub dju2323_Click()
End Sub

Private Sub ddireccion_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   dnombre.SetFocus
   Exit Sub
End If
End Sub

Private Sub dek7834_Click()
End Sub

Private Sub dfk992325_Click()
Dim found As Integer
If dbgrid6.Visible = True Then Exit Sub
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub

flag_clave1 = 0
tconcla.X = "COPIA"  '
tconcla.Show 1
If flag_clave1 <> 1 Then  'si es descongela
   DBGrid2.SetFocus
   Exit Sub
End If
    rrlocal11 = ""
    rrtipo = ""
    rrserie = ""
    rrnumero = ""
mmenua.Caption = "COPIA"
mmenua.Show 1
'MsgBox rrlocal11 & "" & rrtipo & "" & rrserie & " " & rrnumero
If Len(rrlocal11) = 0 Then Exit Sub
If Len(rrtipo) = 0 Then Exit Sub
If Len(rrnumero) = 0 Then Exit Sub
found = valida_otros()
If found = 0 Then
   MsgBox "No existe Documento ", 48, "Aviso"
   Exit Sub
End If
proceso_impresion11 rrtipo, rrserie, rrnumero, 1, "1"
DBGrid2.SetFocus
'proceso_impresioncopia1
End Sub

Private Sub dhyori83_Click()
If "" & mytable11.Fields("terminal") = "T" Then
   MsgBox "No permitido en Pedido", 48, "Aviso"
   DBGrid2.SetFocus
   Exit Sub
End If


If dbgrid6.Visible = True Then Exit Sub
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
'-------------------------------------
   cgusuario = gocabeza
   dgusuariog = godetalle
   menu_proforma
   Exit Sub

End Sub

Private Sub dj232323_Click()


End Sub

Private Sub dj78232_Click()
Dim found As Integer
If dbgrid6.Visible = True Then Exit Sub
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
'-------------------------------------

   cgusuario = gocabeza
   dgusuariog = godetalle
   found = sumar_detalle()
   If Val(txtotal) > 0 Then
      MsgBox "No deben existir pedidos Pendientes", 48, "Aviso"
      DBGrid2.SetFocus
      Exit Sub
   End If
   menu_carga_pedidos
   Exit Sub


End Sub
Sub menu_carga_pedidos()
Dim found As Integer
   Combo1.Clear
   Combo1.AddItem "Nombre"
   Combo1.ListIndex = 0
Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "15000"
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus

End Sub
Sub menu_carga_cotizacion()
Dim found As Integer
sw_consulta = 0
Combo1.Clear
   Combo1.AddItem "Nombre"
   Combo1.ListIndex = 0
Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "30000"
found = sql_consulta(1)
'dbGrid1.SetFocus

End Sub


Private Sub djk7822_Click()

End Sub

Private Sub djk78232_Click()
If dbgrid6.Visible = True Then Exit Sub
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
flag_clave1 = 0
tconcla.X = "S"
tconcla.Show 1
If flag_clave1 = 1 Then  'si es descongela
   modifica_pedido
   Exit Sub
End If
DBGrid2.SetFocus
End Sub
Sub modifica_pedido()
Dim found As Integer
If Val(txtotal) > 0 Then
   MsgBox "No deben existir Productos ", 48, "Aviso"
   DBGrid2.SetFocus
   Exit Sub
End If
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
If Frame9.Visible = True Then Exit Sub
Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "370"
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus

End Sub

Private Sub djuborra_Click()

End Sub

Private Sub dki3432_Click()
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
Frame2.Visible = True
If Len(dnombre) > 0 And Len(telefono) > 0 And Len(codigo) > 0 Then
   ddireccion.SetFocus
   Exit Sub
End If
inicializa_deliveri
telefono.SetFocus
End Sub

Private Sub dmo3434_Click()
End Sub

Private Sub dkioiumwe_Click()

End Sub

Private Sub dklio782_Click()

End Sub

Private Sub dju523a_Click()
Dim found As Integer
If dbgrid6.Visible = True Then Exit Sub
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
facmesa.Show 1
found = sumar_detalle()
               DBGrid2.Col = 0
               DBGrid2.Row = DBGrid2.VisibleRows - 1
               DBGrid2.SetFocus
End Sub

Private Sub dk8923_Click()
Dim found As Integer
If dbgrid6.Visible = True Then Exit Sub
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
'-------------------------------------
   cgusuario = gocabeza
   dgusuariog = godetalle
   found = sumar_detalle()
   If Val(txtotal) > 0 Then
      MsgBox "No deben existir pedidos Pendientes", 48, "Aviso"
      DBGrid2.SetFocus
      Exit Sub
   End If
   menu_carga_cotizacion
   Exit Sub

End Sub

Private Sub dli992323_Click()

End Sub

Private Sub dlko343_Click()
End Sub

Private Sub dlo2323_Click()
End Sub


Private Sub dlo3434_Click()

End Sub

Private Sub dloco343_Click()
End Sub

Private Sub dmo8833_Click()
End Sub

Private Sub dmoi434_Click()
End Sub



Private Sub dnombre_DblClick()
tkeyboar.flag = "DNOMBRE"
tkeyboar.Show 1
End Sub

Private Sub dnombre_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If Len(dnombre) = 0 Then
   dnombre.SetFocus
   Exit Sub
End If
ddireccion.SetFocus
End Sub

Private Sub dnombre_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   telefono.SetFocus
   Exit Sub
End If
End Sub

Private Sub dofpago_Click()

End Sub

Private Sub eju78se_Click()
Dim sw As Integer
    
flag_clave1 = 0
tconcla.X = "CUADRE"
tconcla.Show 1
If flag_clave1 <> 1 Then  'si es descongela
   DBGrid2.SetFocus
   Exit Sub
End If
    
    opcion1 = "20"
    opcion2 = "2"
    opcion3 = ""
    tcuadrc1.flagdiario = "1"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.cajero = usuariopos
    tcuadrc1.caja = caja
    tcuadrc1.turno = turno
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.fechai = dia
    tcuadrc1.fechaf = dia
    'tcuadrc1.todos = "S"
    tcuadrc1.Caption = "DOCUMENTOS EMITIDOS PERIODICO"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1
    tcuadrc1.flagdiario = ""
    
End Sub

Private Sub fdk9235_Click()
Dim found As Integer
If dbgrid6.Visible = True Then Exit Sub
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub

flag_clave1 = 0
tconcla.X = "ANULA"  '
tconcla.Show 1
If flag_clave1 <> 1 Then  'si es descongela
   DBGrid2.SetFocus
   Exit Sub
End If
    rrlocal11 = ""
    rrtipo = ""
    rrserie = ""
    rrnumero = ""
mmenua.Caption = "ANULA"
mmenua.Show 1
If Len(rrlocal11) = 0 Then Exit Sub
If Len(rrtipo) = 0 Then Exit Sub
If Len(rrnumero) = 0 Then Exit Sub
found = valida_otros()
If found = 0 Then
   MsgBox "No existe Documento ", 48, "Aviso"
   Exit Sub
End If
anularr

End Sub
Sub anularr()
Dim found As Integer
      found = proceso_anular(rrtipo, rrserie, rrnumero)
      If found = 1 Then
         proceso_impresion11 rrtipo, rrserie, rrnumero, 0, ""
      End If
      DBGrid2.SetFocus
End Sub

Private Sub fechanac_DblClick()
tkeyboar.flag = "FECHANAC"
tkeyboar.Show 1

End Sub

Private Sub fechanac_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub




found = valida()
If found = 0 Then
   Exit Sub
End If
saludo_cumpe
tiposervicio1 = "DELIVERY"
flag_servicio = "D"
'CAMPO1 = telefono
codigo = dcodigo
nombre = dnombre
salon = ""
mesa = ""
mesero = ""
cuenta_separa = ""
Frame2.Visible = False
DBGrid2.SetFocus
End Sub

Private Sub fechanac_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   referencia.SetFocus
   Exit Sub
End If

End Sub

Private Sub fk88332_Click()
flag_clave1 = 0
tconcla.X = "MINIREPORTE"
tconcla.Show 1
If flag_clave1 <> 1 Then  'si es descongela
   DBGrid2.SetFocus
   Exit Sub
End If

tresegui.Show 1
End Sub

Private Sub Form_Activate()
' CUANDO ES UNA SEPARACION SE DEBE TENER CONFIGURADO EL TIPO PEDIDO EN PARAMECA Y LA SERIE Y EL NUMERO
'
Dim found As Integer
found = leer_visorcaja("SISTEMA ORION", "VERSION 5.0")
tptovtaa.Caption = "" & mytable11.Fields("descripcio")
If flag_carga <> "S" Then
   'MsgBox ""
   found = busca_paridad()
   sql_detalle
   cajero = "" & gusuario
   flag_carga = "S"
   'pedido.SetFocus
      DBGrid2.Col = 0
      DBGrid2.Row = DBGrid2.VisibleRows - 1
      
      DBGrid2.SetFocus
End If
'uvueltos = "S/.:" & Format(Val("" & mytable11.Fields("uvueltos")), nrodecimal)
'uvueltod = "US$:" & Format(Val("" & mytable11.Fields("uvueltod")), nrodecimal)
If "" & mytable11.Fields("terminal") = "T" Then
   'MsgBox "Hola"
   'pedido.SetFocus
End If

If "" & mytable11.Fields("tpvauto") = "N" Then  'activa autoservicio
    'Label55.Enabled = False
    'Label57.Enabled = False
    'Label8.Enabled = False
    'Label63.Enabled = False
End If
If "" & mytable11.Fields("comanda") = "N" Then  'activa autoservicio
    'Label22.Enabled = False
End If
If "" & mytable11.Fields("grabacomanda") = "N" Then  'activa autoservicio
    'Label13.Enabled = False
End If
If "" & mytable11.Fields("delivery") = "N" Then  'activa autoservicio
    'Label23.Enabled = False
End If
If "" & mytable11.Fields("cuadreparcial") = "N" Then  'activa autoservicio
    'Label32.Enabled = False
End If
If "" & mytable11.Fields("copiaventas") = "N" Then  'activa autoservicio
    'Label40.Enabled = False
End If
If "" & mytable11.Fields("anulaventas") = "N" Then  'activa autoservicio
    'Label49.Enabled = False
End If
If "" & mytable11.Fields("cierrecaja") = "N" Then  'activa autoservicio
    'Label58.Enabled = False
End If
If "" & mytable11.Fields("ingresodinero") = "N" Then  'activa autoservicio
    'Label65.Enabled = False
End If
If "" & mytable11.Fields("egresodinero") = "N" Then  'activa autoservicio
    'Label66.Enabled = False
End If
If "" & mytable11.Fields("descuento") = "N" Then  'activa autoservicio
    'Label64.Enabled = False
End If
Frame4.Visible = True
'If xopciones(3).Enabled = False Then
'   Frame4.Visible = True
'End If
cargar_grafico20

End Sub
Sub carga_familia()
Dim mytablex As New ADODB.Recordset
Dim i As Integer
For i = 0 To 14999
    mfamcod(i) = ""
    wfamcod(i) = ""
Next i



i = -1
mytablex.Open "select * from familia where vetouch='S' order by orden ", cn, adOpenStatic, adLockOptimistic

Do
If mytablex.EOF Then Exit Do
i = i + 1
'mfamcod(i) = "" & mytablex.Fields("familia")
mfamcod(i) = "" & mytablex.Fields("descripcio")
wfamcod(i) = "" & mytablex.Fields("familia")


mytablex.MoveNext
Loop
mfamtop = i
mytablex.Close
mfampag = 0
menu_familia "INI"

End Sub


Sub cargar_grafico1()
On Error GoTo cmd7779_err
'Image1.Picture = LoadPicture(globalpath & "\ico\cajaper.jpg")
Exit Sub
cmd7779_err:
MsgBox "" & error$
Exit Sub
End Sub
Sub menu_familia(buf As String)
Dim i As Integer
Dim j As Integer
Select Case buf
       Case "INI"
            mfampag = 0
       Case "SIG"
            mfampag = mfampag + 17
            If mfampag > 102 Then
               mfampag = 0
            End If
       Case "ANT"
            mfampag = mfampag - 17
            If mfampag < 0 Then
               mfampag = 0
            End If
End Select
j = -1
For i = mfampag To 17 + mfampag
    j = j + 1
    zfamilia(j).Caption = mfamcod(i)
    wwfamcod(j) = wfamcod(i)
    
Next i

End Sub

Sub sql_detalle()
Dim buf As String
Dim found As Integer
On Error GoTo cmd34_err
'MsgBox dgusuario
buf = "select * from " & dgusuario
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               Data2.RecordSource = buf
               Data2.refresh
               DBGrid2.refresh
               found = sumar_detalle()
               'DBGrid2.Row = dbgrid2.visiblerows - 2
               'DBGrid2.Col = 0
               'DBGrid2.SetFocus
Exit Sub
cmd34_err:
MsgBox "Error en select " & error$, 48, "Aviso"
Exit Sub
End Sub

Private Sub Form_Load()
Dim found As Integer
Dim sdx As Double
Dim xx As String
   nrodecimal = "0.00"
   If "" & mytable11.Fields("decimal") = "3" Then
      nrodecimal = "0.000"
   End If
   moneda = "" & mytable11.Fields("moneda")
   caja = "" & mytable11.Fields("caja")
   
   DBGrid2.columns("precio").NumberFormat = nrodecimal
   DBGrid2.columns("total").NumberFormat = nrodecimal
   cargas_iniciales
   inicia_color_familia
   inicia_color_producto
   inicia_color_comandos
   carga_familia
   carga_dcvendedor
   carga_clasificacion
   'carga_cobranza
   found = busca_paridad()
   sql_detalle
   borrar_data2
   found = sumar_detalle()
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   
   
               'DBGrid2.SetFocus
   cajero = "" & gusuario
   flag_carga = "S"
   'sumar_detalle
   tiposervicio1 = "Autoservicio"
   salon = ""
   mesa = ""
   mesero = ""
   cuenta_separa = ""
   'borrar_todo
   cargar_tmcombina
   sql_detalle
   
   
   flag_servicio = "A"
   sql_detalle
   xx = busca_parame1("", 2)
   If "" & mytable11.Fields("terminal") = "T" Then
      menju232.Visible = False
      'dlo2342.Visible = False
      'dek7834.Visible = False
      'inu781.Visible = False
      'djk7822.Visible = False
      cuj6721.Visible = False
      'Frame10.Visible = True
      'Label32.Visible = True
      'pedido.Visible = True
      End If
      'Frame10.Left = 10560
      'Frame10.Height = 1445
      'Frame10.Top = 840
      'Frame10.Width = 3855
      
      'ezVidCap1.Height = 1080
      'ezVidCap1.Top = 240
      'ezVidCap1.Left = -240
      'ezVidCap1.Width = 3960
      
      
      'Frame10.Height = 2175
      'Frame10.Top = 0
      'Frame10.Left = 10680
      'Frame10.Width = 3855
      
      'ezVidCap1.Height = 1920
      'ezVidCap1.Top = 240
      'ezVidCap1.Left = 0
      'ezVidCap1.Width = 3840
      cargar_grafico1
      consulta_comanda "" & mytable11.Fields("salon")
      
      
    
   
End Sub
Sub cargas_iniciales()
'Dim mydbx As Database
'Dim mytablex As Table
'fpago.Clear
'tipodoc.Clear
'vendedor.Clear
'vendedor.AddItem "*"
'Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
'Set mytablex = mydbxglo.OpenTable("fpago")
'Do
'If mytablex.EOF Then Exit Do
'fpago.AddItem "" & mytablex.Fields("fpago") & "|" & mytablex.Fields("descripcio")
'mytablex.MoveNext
'Loop
'mytablex.Close
'Set mytablex = mydbxglo.OpenTable("tipo")
'Do
'If mytablex.EOF Then Exit Do
'If "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "D" Or "" & mytablex.Fields("tipodoc") = "G" Then
'   tipodoc.AddItem "" & mytablex.Fields("tipo") & "|" & mytablex.Fields("descripcio")
'End If
'mytablex.MoveNext
'Loop
'mytablex.Close
'Set mytablex = mydbxglo.OpenTable("vendedor")
'Do
'If mytablex.EOF Then Exit Do
'vendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
'mytablex.MoveNext
'Loop
'mytablex.Close
'
'vendedor.ListIndex = 0
'tipodoc.ListIndex = 0
'fpago.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo cmd123_err
cerrar_puerto
Data2.Recordset.Close
Exit Sub
cmd123_err:
Exit Sub
End Sub
Sub cerrar_puerto()
On Error GoTo cmd8912_err
MSComm1.PortOpen = False
Exit Sub
cmd8912_err:
Exit Sub
End Sub


Private Sub hyu545_Click()
End Sub

Private Sub forma671_Click()
Dim sw As Integer
flag_clave1 = 0
tconcla.X = "CUADRE"
tconcla.Show 1
If flag_clave1 <> 1 Then  'si es descongela
   DBGrid2.SetFocus
   Exit Sub
End If

    
    opcion1 = "6"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.cajero = usuariopos
    tcuadrc1.caja = caja
    tcuadrc1.turno = turno
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.fechai = dia
    tcuadrc1.fechaf = dia
    
    tcuadrc1.Caption = "FORMPAGO-DOCUMENTOS"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1
End Sub

Private Sub Frame10_Click()
Exit Sub
'If Frame10.Width = 3855 Then
'      Frame10.Height = 3615
'      Frame10.Top = 2400
'      Frame10.Left = 3120
'      Frame10.Width = 6855
      
      'ezVidCap1.Height = 3240
      'ezVidCap1.Left = -240
      'ezVidCap1.Top = 240
      'ezVidCap1.Width = 4080
      Exit Sub
'End If
'If Frame10.Width = 6855 Then
     
      
'      Frame10.Height = 2175
'      Frame10.Top = 0
'      Frame10.Left = 10680
'      Frame10.Width = 3855
      
      'ezVidCap1.Height = 1920
      'ezVidCap1.Top = 240
      'ezVidCap1.Left = 0
      'ezVidCap1.Width = 3840

      
'      Exit Sub
'End If

End Sub


Sub habilita_lab7(sw As Integer)
Dim xsw
If sw = 0 Then
   xsw = True
End If
If sw = 1 Then
   xsw = False
End If

Label70.Enabled = xsw
Label71.Enabled = xsw
Label72.Enabled = xsw
Label73.Enabled = xsw
End Sub

Private Sub fotoimagen_Click()
'frmain.Show 1
'toparam.caja = "" & mytable11.Fields("caja")
'toparam.Show 1
'inicia_color_comandos
End Sub

Private Sub hundv1_Click()
Dim sw As Integer
flag_clave1 = 0
tconcla.X = "CUADRE"
tconcla.Show 1
If flag_clave1 <> 1 Then  'si es descongela
   DBGrid2.SetFocus
   Exit Sub
End If
    opcion1 = "3"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.cajero = usuariopos
    tcuadrc1.caja = caja
    tcuadrc1.turno = turno
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.fechai = dia
    tcuadrc1.fechaf = dia
    tcuadrc1.Caption = "UNIDADES VENDIDAS"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1
End Sub

Private Sub hydes8912_Click()

End Sub

Private Sub Image10_Click()
consulta_comanda "" & mytable11.Fields("salon")
End Sub

Private Sub image2_Click()
'If dbvarios.BOF = False Then
   cmytablex.MovePrevious 'movemos al registro anterior
'   Exit Sub
'End If

'dbvarios.MovePrevious
If Not cmytablex.EOF Then
   cmytablex.MoveFirst
   Exit Sub
End If

End Sub

Private Sub image3_Click()
On Error GoTo cmd89012_err
TBRCOMA.salon = Trim("" & cmytablex.Fields("salon"))
TBRCOMA.mesa = Trim("" & cmytablex.Fields("mesa"))
TBRCOMA.Show 1
consulta_comanda "" & mytable11.Fields("salon")
Exit Sub
cmd89012_err:
Exit Sub
End Sub

Private Sub Image4_Click()
menu_producto "SIG"

End Sub

Private Sub Image5_Click()
menu_familia "SIG"

End Sub

Private Sub Image6_Click()
menu_familia "ANT"

End Sub

Private Sub Image7_Click()
If Data2.Recordset.EOF = False Then
   Data2.Recordset.MoveNext 'movemos al siguiente registro
   carga_grafico "" & Data2.Recordset.Fields("producto")
   carga_minimo "" & Data2.Recordset.Fields("producto")
   Exit Sub
End If
If Not Data2.Recordset.BOF Then
   Data2.Recordset.MoveLast
   carga_grafico "" & Data2.Recordset.Fields("producto")
   carga_minimo "" & Data2.Recordset.Fields("producto")
   Exit Sub
End If

End Sub

Private Sub Image8_Click()
If Data2.Recordset.BOF = False Then
   Data2.Recordset.MovePrevious 'movemos al registro anterior
   carga_grafico "" & Data2.Recordset.Fields("producto")
   carga_minimo "" & Data2.Recordset.Fields("producto")
   Exit Sub
End If

'dbvarios.MovePrevious
If Not Data2.Recordset.EOF Then
   Data2.Recordset.MoveFirst
   carga_grafico "" & Data2.Recordset.Fields("producto")
   carga_minimo "" & Data2.Recordset.Fields("producto")
   Exit Sub
End If

End Sub

Private Sub Image9_Click()
menu_producto "ANT"

End Sub

Private Sub inu781_Click()
End Sub

Private Sub labe57_Click()

End Sub

Private Sub kcobra_Click(Index As Integer)
If Index = 11 Then
          RGPAGO = ""
          Exit Sub
End If
RGPAGO = RGPAGO & kcobra(Index).Caption

End Sub

Private Sub Label11_Click()
If dbvarios.EOF = False Then
   dbvarios.MoveNext 'movemos al siguiente registro
End If
If Not dbvarios.BOF Then
   dbvarios.MoveLast
   Exit Sub
End If

End Sub

Private Sub Label12_Click()

End Sub

Private Sub Label13_Click()
End Sub
Sub sin_meseros()
Dim found As Integer
On Error GoTo cmd1289121_err
       'nmesero = usuariog
       'mesero = usuariog
       'grabar_comandax
       Exit Sub
cmd1289121_err:
       MsgBox "Seleccione un dato ", 48, "Aviso"
       Exit Sub

End Sub



Private Sub Label14_Click()

End Sub
Sub borra_congela()
If Frame2.Visible = True Then Exit Sub
'If MsgBox("Desea Borrar ??", 1, "Aviso") <> 1 Then Exit Sub
borrar_todo
sql_detalle
tiposervicio1 = "Autoservicio"
flag_servicio = "A"
salon = ""
mesa = ""
mesero = ""
cuenta_separa = ""

End Sub
Private Sub lkcop992_Click()
End Sub

Private Sub Label15_Click()
Dim found As Integer
'      If dbclie.State = 1 Then dbclie.Close
'      dbclie.Open "SELECT * FROM clientes where codigo='" & dcodigo & "'", cn, adOpenDynamic, adLockOptimistic
'      If dbclie.RecordCount = 0 Then
'          dbclie.Close
'          Exit Sub
'      End If
'tnclie.Caption = "MODIFICA"
'tnclie.profesion = Trim("" & dbclie.Fields("profesion"))
'tnclie.religion = Trim("" & dbclie.Fields("religion"))
'tnclie.nrodepe = Trim("" & dbclie.Fields("nrodepe"))
'tnclie.Trabajo = Trim("" & dbclie.Fields("trabajo"))
'tnclie.cargo = Trim("" & dbclie.Fields("cargo"))
'tnclie.hobbie = Trim("" & dbclie.Fields("hobbie"))
'tnclie.civil = Trim("" & dbclie.Fields("civil"))
'tnclie.tipovive = Trim("" & dbclie.Fields("tipovive"))


'tnclie.barras = Trim("" & dbclie.Fields("barras"))
'tnclie.ruc = Trim("" & dbclie.Fields("ruc"))
'tnclie.dni = Trim("" & dbclie.Fields("dni"))
'tnclie.especial = Trim("" & dbclie.Fields("especial"))
'tnclie.clasifica = Trim("" & dbclie.Fields("clasifica"))
'tnclie.tipoclie = Trim("" & dbclie.Fields("tipoclie"))

'tnclie.zona = Trim("" & dbclie.Fields("zona"))
'tnclie.lunes.Value = Val("" & dbclie.Fields("lunes"))
'tnclie.martes.Value = Val("" & dbclie.Fields("martes"))
'tnclie.miercoles.Value = Val("" & dbclie.Fields("miercoles"))
'tnclie.jueves.Value = Val("" & dbclie.Fields("jueves"))
'tnclie.viernes.Value = Val("" & dbclie.Fields("viernes"))
'tnclie.sabado.Value = Val("" & dbclie.Fields("sabado"))
'tnclie.domingo.Value = Val("" & dbclie.Fields("domingo"))
'tnclie.fechalta = Trim("" & dbclie.Fields("fechanac"))
'tnclie.referencias = Trim("" & dbclie.Fields("observa"))
'tnclie.referencia = Trim("" & dbclie.Fields("referencia"))
'tnclie.garantia = Trim("" & dbclie.Fields("garantia"))
'tnclie.flete = Trim("" & dbclie.Fields("flete"))
'tnclie.moneda = Trim("" & dbclie.Fields("moneda"))
'tnclie.descuento1 = Trim("" & dbclie.Fields("descuento1"))
'tnclie.credito = Trim("" & dbclie.Fields("credito"))
'tnclie.vendedor = Trim("" & dbclie.Fields("vendedor"))
'tnclie.descuento = Trim("" & dbclie.Fields("descuento"))
'tnclie.diapago = Trim("" & dbclie.Fields("diapago"))
'tnclie.fpago = Trim("" & dbclie.Fields("fpago"))
'tnclie.cuenta = Trim("" & dbclie.Fields("cuenta"))
'tnclie.codigo = Trim("" & dbclie.Fields("codigo"))
'tnclie.codigo1 = Trim("" & dbclie.Fields("extranjeria"))
'tnclie.nombre = Trim("" & dbclie.Fields("nombre"))
'tnclie.nombrec = Trim("" & dbclie.Fields("nombrec"))
'tnclie.contacto = Trim("" & dbclie.Fields("contacto"))
'tnclie.direccion = Trim("" & dbclie.Fields("direccion"))
'tnclie.dpto = Trim("" & dbclie.Fields("dpto"))
'tnclie.distrito = Trim("" & dbclie.Fields("distrito"))
'tnclie.telefono = Trim("" & dbclie.Fields("telefono"))
'tnclie.telefono1 = Trim("" & dbclie.Fields("telefono1"))
'tnclie.telefono2 = Trim("" & dbclie.Fields("telefono2"))
'tnclie.correo = Trim("" & dbclie.Fields("correo"))
'tnclie.estado = Trim("" & dbclie.Fields("estado"))
'tnclie.codigo.Enabled = False
'tnclie.Show 1
'dbclie.Close

End Sub

Private Sub Label19_Click()
End Sub

Private Sub Label2_Click()
'valida_camara
trgb.tipo = "FAMILIA"
trgb.Show 1
inicia_color_familia
End Sub

Private Sub Label21_Click()
Dim found As Integer
      amsw = 1
      If dbclie.State = 1 Then dbclie.Close
      dbclie.Open "SELECT * FROM clientes", cn, adOpenDynamic, adLockOptimistic
      tnclie.telefono = telefono
      tnclie.moneda = "S"
      tnclie.Caption = "NUEVO"
      tnclie.Show 1
      dbclie.Close
      
      

End Sub

Private Sub Label22_Click()

End Sub

Private Sub Label23_Click()

End Sub

Private Sub Label24_Click()
   If local1.Visible <> True Then  'si no es traslado
      consulta_xtipo
   End If

End Sub

Private Sub Label25_Click()

End Sub

Private Sub Label26_Click()
DBGrid1_KeyDown 13, 0
End Sub

Private Sub Label27_Click()

End Sub

Private Sub Label3_Click()
If Label3 = "+" Then
   Label3 = "-"
   table2.Height = 4215
   table2.Top = 240
   table2.Width = 3255
   Exit Sub
End If
If Label3 = "-" Then
   Label3 = "+"
   table2.Height = 10095
   table2.Top = 240
   table2.Width = 3255
   Exit Sub
End If

End Sub

Private Sub Label31_Click()
losao94_Click
End Sub

Private Sub Label32_Click()
End Sub

Private Sub Label40_Click()

End Sub

Private Sub Label49_Click()

End Sub

Private Sub Label53_Click()
Dim found As Integer
On Error GoTo cmd56123_err
tcomanda.mesero = dbvarios.Fields("codigo")
tcomanda.nmesero = dbvarios.Fields("nombre")
flag_comanda = 0
tcomanda.Show 1

If flag_comanda = "1" Then
   'MsgBox "paso"
   flag_servicio = "C"
   found = orden_despacho()
   borrar_todo
   sql_detalle
   tiposervicio1 = "Autoservicio"
   flag_servicio = "A"
   Frame8.Visible = False
End If
consulta_comanda "" & mytable11.Fields("salon")
Exit Sub
cmd56123_err:
MsgBox "Seleccione un Salon Y Mesa ", 48, "Aviso"
Exit Sub

End Sub

Private Sub Label54_Click()
Frame8.Visible = False
End Sub

Sub proceso_cobross()
Dim found As Integer
If Frame2.Visible = True Then Exit Sub
local1 = ""
local1.Visible = False
found = sumar_detalle()
If found = 0 Then
   MsgBox "debe de Existir un Precio=0", 48, "Aviso"
   DBGrid2.SetFocus
   Exit Sub
End If

If Val(txtotal) = 0 Then
   If exisdev <> -10 Then  'si existe devolucion
     If Val(ntcant) = 0 Then
        DBGrid2.SetFocus
        Exit Sub
     End If
      
   End If
End If
If mytable11.Fields("terminal") = "T" Or Val(acuenta) > 0 And Len(petipo) = 0 Then 'pedidos o a cuenta ha dado
          'MsgBox "Hola"
          'xruc = codigo
          'xnombre = nombre
          Frame7.Visible = True
          habilita_lab7 1
          Framefp.Enabled = False
          If Val(acuenta) > 0 Then
             xtipo = "" & mytable11.Fields("tipope")
          End If
          If flag_servicio <> "A" And flag_servicio <> "C" And flag_servicio <> "D" Then
             xtipo = "5"
          End If
          xtipo.SetFocus
          Exit Sub
End If

'If Val(acuenta) > 0 Then  'si existo a cuenta entonces debe ser vendido asi
'   MsgBox "Utilizar icono "
'End If
If flag_servicio = "A" Then  'venta rapida
End If
If flag_servicio = "D" Then  'validar el deliveri si ingreso datos
End If
If flag_servicio = "C" Then  'venta mesas
End If
'Frame10.Visible = False
Label36.Caption = "Codigo"
found = proceso_cobros()
opcion2 = 0
ttxtotals = Format(Val(rtxtotal), nrodecimal)
ttxtotald = Format(Val(rtxtotald), nrodecimal)
stxtotals = Format(Val(rtxtotal), nrodecimal)
stxtotald = Format(Val(rtxtotald), nrodecimal)
found = leer_visorcaja("S/." & stxtotals, "US$  " & stxtotald)

habilita_lab7 0
Framefp.Visible = True
Framefp.Enabled = True
carga_tiposdoc "%"
'MsgBox "Hola"
dbgrid10.Enabled = True
dbgrid10.SetFocus
DBGrid10_KeyDown 13, 0
palabra_bienvenida1
'Frame10.Visible = True

End Sub

Private Sub local1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
End Sub

Private Sub Label57_Click()

End Sub

Private Sub Label58_Click()

End Sub

Private Sub Label59_Click()
If Label59.Caption = "NORMAL" Then
   Label59.Caption = "CONSUMO"
   Exit Sub
End If
If Label59.Caption <> "NORMAL" Then
   Label59.Caption = "NORMAL"
   Exit Sub
End If

End Sub

Private Sub Label60_Click()
consulta_comanda "01"
End Sub

Private Sub Label61_Click()
consulta_comanda "02"
End Sub

Private Sub Label62_Click()
End Sub

Private Sub Label63_Click()
End Sub

Private Sub Label64_Click()
End Sub

Private Sub Image1_Click()
If cmytablex.EOF = False Then
   cmytablex.MoveNext 'movemos al siguiente registro
End If
If Not cmytablex.BOF Then
   cmytablex.MoveLast
   Exit Sub
End If

End Sub

Private Sub Label65_Click()

End Sub

Private Sub Label66_Click()

End Sub

Private Sub Label67_Click()
consulta_comanda "03"
End Sub

Private Sub Label68_Click()
consulta_comanda ""
End Sub

Private Sub Label69_Click()
tmesasta.Show 1
consulta_comanda "" & mytable11.Fields("salon")
End Sub

Private Sub Label7_Click()
'If dbvarios.BOF = False Then
   dbvarios.MovePrevious 'movemos al registro anterior
'   Exit Sub
'End If

'dbvarios.MovePrevious
If Not dbvarios.EOF Then
   dbvarios.MoveFirst
   Exit Sub
End If



End Sub


Private Sub Label70_Click()
On Error GoTo cmdp97811_err
DBGrid10_KeyDown 13, 0
Exit Sub

If dbgrid10.Enabled = True Then
   seleccionamos_fpago
   Exit Sub
End If
Exit Sub
'MsgBox Shift
   Select Case DBGrid9.Col
       Case 2
            If Len("" & DBGrid9.columns(2)) > 0 Then Exit Sub
            If Val("" & DBGrid9.columns(2)) = 0 Then
               If "" & Data9.Recordset.Fields("moneda") = "S" Then
                  Data9.Recordset.Edit
                  Data9.Recordset.Fields("recibe") = Val(stxtotals)
                  Data9.Recordset.Update
               End If
               If "" & Data9.Recordset.Fields("moneda") = "D" Then
                  Data9.Recordset.Edit
                  Data9.Recordset.Fields("recibe") = Val(stxtotald)
                  Data9.Recordset.Update
               End If
               opcion2 = 0
               'valida_ingresado
               suma_fpagov
               
               If Label45.Caption = "Vuelto" Or Val(stxtotals) <= 0 Then
                  xtipo = protipo
                  xvendedor = cproven
                  xruc = codigo
                  If "" & mytable11.Fields("habilitanota") = "S" Then
                     If Val(ttxtotals) <= Val("" & mytable11.Fields("siventa")) Then
                        xtipo = "5"
                     End If
                  End If
                  
                  xnombre = nombre
                  Frame7.Visible = True
                  habilita_lab7 1
                  Framefp.Enabled = False
                  If flag_servicio <> "A" And flag_servicio <> "C" And flag_servicio <> "D" Then
                     xtipo = "5"
                  End If
                  xtipo.SetFocus
               Exit Sub
               End If
             End If
   End Select

Exit Sub
cmdp97811_err:
Exit Sub

End Sub
Sub seleccionamos_fpago()
On Error GoTo cmdk8911_err
suma_fpagov
If Label45.Caption = "Vuelto" Or Val(stxtotals) = 0 Then
          'If MsgBox("Desea Grabar", 1, "Aviso") <> 1 Then Exit Sub
          'if len()
             xtipo = protipo
             If "" & mytable11.Fields("habilitanota") = "S" Then
                If Val(ttxtotals) <= Val("" & mytable11.Fields("siventa")) Then
                   xtipo = "5"
                End If
             End If
                  'xruc = codigo
                  'xnombre = nombre
                  xvendedor = cproven
                  Frame7.Visible = True
                  habilita_lab7 1
                  Framefp.Enabled = False
                  If flag_servicio <> "A" And flag_servicio <> "C" And flag_servicio <> "D" Then
                     xtipo = "5"
                  End If
                  xtipo.SetFocus
          Exit Sub
End If
saldoabo = ""
acufp = "" & dbgrid10.columns(3)
Frame6.Caption = "" & dbgrid10.columns(0)
fpago = "" & dbgrid10.columns(1)
fpmoneda = "" & dbgrid10.columns(2)
dbgrid10.Enabled = False
If "" & dbgrid10.columns(3) = "A" Or "" & dbgrid10.columns(3) = "B" Or "" & dbgrid10.columns(3) = "E" Or "" & dbgrid10.columns(3) = "U" Then  'efectivo,dolares,euros
   macro_inserta_registro
   DBGrid9.Row = DBGrid9.VisibleRows - 1
   DBGrid9.Col = 2
   DBGrid9.SetFocus
   Exit Sub
End If
If "" & dbgrid10.columns(3) = "C" Then   'credito
   macro_credito 3
   tcampo1.SetFocus
End If
If "" & dbgrid10.columns(3) = "D" Then   'tarejta credito
   macro_credito 4
   tcampo1.SetFocus
End If
If "" & dbgrid10.columns(3) = "F" Then   'TARJETA DEBITO
   macro_credito 5
   tcampo1.SetFocus
End If
If "" & dbgrid10.columns(3) = "G" Then   'letra
   macro_credito 0
   tcampo1.SetFocus
End If
If "" & dbgrid10.columns(3) = "H" Or "" & dbgrid10.columns(3) = "K" Then   'bancos
   macro_credito 2
   tcampo3.SetFocus
End If
If "" & dbgrid10.columns(3) = "V" Then   'vales
   macro_credito 6
   tcampo1.SetFocus
End If
If "" & dbgrid10.columns(3) = "J" Then   'ORDEN TRABAJO
   macro_credito 8
   tcampo1.SetFocus
End If

If "" & dbgrid10.columns(3) = "I" Or "" & dbgrid10.columns(3) = "K" Then   'CRUCE CON ABONO EFECTIVO
   macro_credito 1
   tcampo1.Enabled = True
   tcampo1.SetFocus
End If
Exit Sub
cmdk8911_err:
MsgBox error$, 48, "Aviso"
Exit Sub

End Sub

Private Sub Label71_Click()
Dim found As Integer
   found = borra_data9()
   If found = 0 Then
      dbgrid10.Enabled = True
      dbgrid10.SetFocus
      Exit Sub
   End If
   suma_fpagov
   dbgrid10.Enabled = True
   dbgrid10.SetFocus

   Exit Sub

End Sub

Private Sub Label72_Click()
On Error GoTo cmd9001_err
mytablefpago.MoveNext 'movemos al siguiente registro
If mytablefpago.EOF Or mytablefpago.BOF Then
   mytablefpago.MoveLast
   Exit Sub
End If
Exit Sub
cmd9001_err:
Exit Sub
End Sub

Private Sub Label73_Click()
On Error GoTo cmd9002_err
mytablefpago.MovePrevious 'movemos al siguiente registro
If mytablefpago.EOF Or mytablefpago.BOF Then
   mytablefpago.MoveFirst
   Exit Sub
End If
Exit Sub
cmd9002_err:
Exit Sub

End Sub

Private Sub Label74_Click()


End Sub

Private Sub Label8_Click()
trgb.tipo = "PRODUCTO"
trgb.Show 1
inicia_color_producto
End Sub

Private Sub losao94_Click()
Dim found As Integer
'If Frame3.Visible = True Then
'   Frame3.Visible = False
'   dbgrid2.Col = 0
'   dbgrid2.Row = dbgrid2.visiblerows - 1
'   dbgrid2.SetFocus
'   Exit Sub
'End If
If Frame5.Visible = True Then
   If Frame1.Visible = True Then
      Frame5.Visible = False
      dbGrid1.SetFocus
      Exit Sub
   End If
   Command8_Click
   Exit Sub
End If
If dbgrid6.Visible = True Then
   dbgrid6.Visible = False
   dbGrid1.SetFocus
   Exit Sub
End If
If Frame6.Visible = True Then
   If opcion1 = "99" Then
      If Frame1.Visible = True Then
         Frame1.Visible = False
         Frame1.Enabled = False
         tcampo1.SetFocus
         Exit Sub
      End If
   End If
   If opcion1 = "2800" Or opcion1 = "2003" Then
      If Frame1.Visible = True Then
         Frame1.Visible = False
         Frame1.Enabled = False
         tcampo3.SetFocus
         Exit Sub
      End If
   End If
   habilita_lab7 1
   Frame6.Visible = False
   'dbgrid10.SetFocus
   Exit Sub
End If


If Frame7.Visible = True Then
   
   If opcion1 = "30" Or opcion1 = "300" Then
      If Frame1.Visible = True Then
         Frame1.Visible = False
         Frame1.Enabled = False
         xruc.SetFocus
         Exit Sub
      End If
   End If
   If opcion1 = "31" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      Frame1.Enabled = False
      xvendedor.SetFocus
      Exit Sub
   End If
   End If
   
   If opcion1 = "29" Then
      If Frame1.Visible = True Then
         Frame1.Visible = False
         Frame1.Enabled = False
         xtipo.SetFocus
         Exit Sub
      End If
   End If
   If opcion1 = "8" Then
   Frame7.Visible = False
   habilita_lab7 0
   DBGrid2.Enabled = True
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus
   Exit Sub
   End If
   Frame7.Visible = False
   habilita_lab7 0
   
   If "" & mytable11.Fields("terminal") = "T" Or opcion1 = "9999" Then
   DBGrid2.Enabled = True
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus
   Exit Sub
   End If
   If Framefp.Visible = True Then
      habilita_lab7 0
      Framefp.Enabled = True
      dbgrid10.Enabled = True
      dbgrid10.SetFocus
   Exit Sub
   End If
   If opcion1 = "1000" Then
      Frame7.Visible = False
      habilita_lab7 0
      DBGrid2.Col = 0
      DBGrid2.Row = DBGrid2.VisibleRows - 1
      DBGrid2.SetFocus
      Exit Sub
   End If
   Exit Sub
End If

'If Frame10.Visible = True Then
'   If Framefp.Visible = True Then
'      Framefp.Visible = False
'      Frame10.Visible = True
'      DBGrid2.Enabled = True
'   DBGrid2.Col = 0
'   DBGrid2.Row = dbgrid2.visiblerows - 1
'   DBGrid2.SetFocus
'      Exit Sub
'   End If
'End If

If Frame6.Visible = True Then
   Frame6.Visible = False
   habilita_lab7 0
   dbgrid10.Enabled = True
   dbgrid10.SetFocus
   Exit Sub
End If
If opcion1 = "19000" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      Frame1.Enabled = False
      'xvendedor.SetFocus
      Exit Sub
   End If
End If

If opcion1 = "31" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      Frame1.Enabled = False
      xvendedor.SetFocus
      Exit Sub
   End If
End If

If opcion1 = "23" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      Frame1.Enabled = False
      tcampo1.SetFocus
      Exit Sub
   End If
End If
If opcion1 = "29" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      Frame1.Enabled = False
      xtipo.SetFocus
      Exit Sub
   End If
End If
If opcion1 = "30" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      Frame1.Enabled = False
      xruc.SetFocus
      Exit Sub
   End If
End If
If opcion1 = "8" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      Frame1.Enabled = False
      found = ir_ultimo_registrox()
      If found = 0 Then
         Data2.refresh
      End If
      DBGrid2.Enabled = True
      DBGrid2.Col = 0
      DBGrid2.Row = DBGrid2.VisibleRows - 1
      DBGrid2.SetFocus
      Exit Sub
   End If
End If

If opcion1 = "0" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      Frame1.Enabled = False
      DBGrid2.Enabled = True
      telefono.SetFocus
      Exit Sub
   End If
End If
If opcion1 = "750" Or opcion1 = "13" Or opcion1 = "10" Or opcion1 = "14" Or opcion1 = "15" Or opcion1 = "100" Or opcion1 = "150" Or opcion1 = "370" Or opcion1 = "1500" Or opcion1 = "1900" Or opcion1 = "15000" Or opcion1 = "30000" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      Frame1.Enabled = False
      DBGrid2.Enabled = True
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus
      Exit Sub
   End If
End If

If opcion1 = "1" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      Frame1.Enabled = False
      DBGrid2.Enabled = True
      If Len(telefono) < 6 Then
         telefono.SetFocus
         Exit Sub
      End If
      dcodigo.SetFocus
      Exit Sub
   End If
End If
If opcion1 = "1750" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      Frame1.Enabled = False
      DBGrid2.Enabled = True
      telefono.SetFocus
      Exit Sub
   End If
End If

If opcion1 = "12" Then
   If Frame1.Visible = True Then
      Frame1.Visible = False
      Frame1.Enabled = False
      DBGrid2.Enabled = True
      codigo.SetFocus
      Exit Sub
   End If
End If
If Frame2.Visible = True Then
   
   If Len(telefono) > 0 Or Len(ddireccion) > 0 Or Len(fechanac) > 0 Or Len(codigo) > 0 Then
      MsgBox "Existen Campos", 48, "Aviso"
      telefono.SetFocus
      Exit Sub
   End If

   Frame2.Visible = False
   DBGrid2.Enabled = True
   Command10_Click
   Exit Sub
End If
If Framefp.Visible = True Then
   habilita_lab7 1
   Framefp.Visible = False
   If flag_servicio = "C" Then
      inicialIzatodo
   End If
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus
   Exit Sub
End If
'MsgBox opcion1
If MsgBox("Desea Salir", 1, "Aviso") <> 1 Then Exit Sub
'menucaja.vendedor = ""
'menucaja.nombre = ""
'menucaja.clave = ""
'MsgBox ""
'cerrar_data2
'cerrar_archivo
tptovtaa.Hide
Unload tptovtaa
End Sub


Private Sub mesero_Click()
dj78232_Click
'tdremoto.Show 1
End Sub

Private Sub MSComm1_OnComm()
Dim i As Integer
Dim buf As String
Exit Sub
i = 0
If MSComm1.CommEvent = 2 Then 'comEvReceive Then
   'If "" & mytable11.Fields("tipo_balanza") = "1" Then
    buf = MSComm1.Input
    i = InStr(buf, Chr(13))
    If i = 0 Then
        cadena = cadena & buf
        Else
        cadena = cadena & Left(buf, i - 1)
    End If
    cadena_balanza = Mid$(cadena, Len(cadena) - 7, 6)
    
   'End If
    
End If



'Select Case MSComm1.CommEvent
'Case comEvReceive ' Received RThreshold # of chars.
'     InBuff = InBuff + MSComm1.Input
'End Select
End Sub


Private Sub nombre_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus

End Sub

Private Sub nombre_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   codigo.SetFocus
   Exit Sub
End If
End Sub


Private Sub nbxtipo_Click(Index As Integer)
xtipo = Trim(nbxtipo(Index).Caption)
If Len(xtipo) = 0 Then
   xtipo.SetFocus
   Exit Sub
End If
xtipo_keyPress 13
End Sub

Private Sub pado8911_Click()
Dim sw As Integer
flag_clave1 = 0
tconcla.X = "CUADRE"
tconcla.Show 1
If flag_clave1 <> 1 Then  'si es descongela
   DBGrid2.SetFocus
   Exit Sub
End If

    
    opcion1 = "2"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.cajero = usuariopos
    tcuadrc1.caja = caja
    tcuadrc1.turno = turno
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.fechai = dia
    tcuadrc1.fechaf = dia
    tcuadrc1.Caption = "DOCUMENTOS EMITIDOS"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1
    
End Sub

Function verifica_ticket_ingreso(buf As String)
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM ppocket where  pedido='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   verifica_ticket_ingreso = 1
End If
mytablex.Close
End Function
Function carga_ticket_ingreso()
Dim found As Integer
found = proceso_proforma("" & mytable11.Fields("local"), "P", "P", "" & pedido)
carga_ticket_ingreso = found
End Function

Private Sub Picture2_Click()

End Sub

Private Sub referencia_DblClick()
tkeyboar.flag = "DREFERENCIA"
tkeyboar.Show 1

End Sub

Private Sub referencia_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
fechanac.SetFocus

End Sub

Private Sub referencia_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   ddireccion.SetFocus
   Exit Sub
End If

End Sub

Private Sub saldo_KeyPress(KeyAscii As Integer)
End Sub

Private Sub RGPAGO_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   Command2_Click
   Exit Sub
End If

Command3_Click
End Sub

Private Sub sentido_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   sentido.SetFocus
   Exit Sub
End If
If Len(xtipo) = 0 Then
   xtipo.SetFocus
   Exit Sub
End If
If sentido <> "S" And sentido <> "B" Then
   sentido = ""
   Exit Sub
End If
If "" & mytable11.Fields("vendedor") = "S" Then
   xvendedor.SetFocus
   Exit Sub
End If
If xtipo = "7" Then
   xruc.SetFocus
   Exit Sub
End If
If "" & mytable11.Fields("cliente") <> "S" And acu <> "B" And acu <> "D" Then
   Command13_Click
   Exit Sub
End If
xruc.SetFocus
End Sub

Private Sub sentido_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   xtipo.SetFocus
End If
End Sub


Private Sub table6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Label53_Click
End If
End Sub

Private Sub tcampo1_DblClick()
If Val(RGPAGO) = 0 Then
   RGPAGO.SetFocus
   Exit Sub
End If
tkeyboar.flag = "TCAMPO1"
tkeyboar.Show 1

End Sub

Private Sub tcampo1_KeyPress(KeyAscii As Integer)
Dim found As Integer
Dim found1 As Double

If Val(RGPAGO) = 0 Then
   RGPAGO.SetFocus
   Exit Sub
End If

If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   Frame6.Visible = False
   habilita_lab7 0
   dbgrid10.Enabled = True
   dbgrid10.SetFocus
   Exit Sub
End If
If Frame6.Caption = "CREDITO" Or "" & dbgrid10.columns("tipo") = "I" Or "" & dbgrid10.columns("tipo") = "H" Or "" & dbgrid10.columns("tipo") = "K" Or "" & dbgrid10.columns("tipo") = "V" Then
   'If Len(tcampo1) = 0 Then
   '   tcampo1.SetFocus
   '   Exit Sub
   'End If
End If

found = 0
If Len(tcampo1) > 0 Then
   found = busca_codigocl("" & tcampo1, 0)
End If

If "" & dbgrid10.columns("tipo") = "C" Then  'si es credito
   If "" & mytable11.Fields("obligacredito") = "S" Then
      found = credito_habilitado("" & tcampo1)
      If found = 0 Then
         MsgBox "Credito no permitido ", 48, "Aviso"
         tcampo1 = ""
         tcampo2 = ""
         tcampo1.SetFocus
         Exit Sub
      End If
   End If
End If
If "" & dbgrid10.columns("tipo") = "C" And found = 1 Then '
   saldoabo = ""
   found = busca_credito_credito("" & dbgrid10.columns("tipo"), "" & tcampo1)  'actualiza su saldo actual
   If "" & mytable11.Fields("obligacredito") = "S" Then
   If saldo_clientes(tcampo1, Val(RGPAGO)) <= 0 Then
      MsgBox "No existe saldo", 48, "Aviso"
      tcampo1.SetFocus
      Exit Sub
   End If
   End If
     
   'If found = 0 Then
   '   MsgBox "No existe Cliente o No tiene saldo ", 48, "Aviso"
   '   tcampo1.SetFocus
   '   Exit Sub
   'End If
   tcampo5.SetFocus
   Exit Sub
End If

If ("" & dbgrid10.columns("tipo") = "I" Or "" & dbgrid10.columns("tipo") = "K" Or "" & dbgrid10.columns("tipo") = "C") And found = 1 Then '
   saldoabo = ""
   found = busca_credito_adelanto("" & dbgrid10.columns("tipo"), "" & tcampo1)
   If found = 1 And Val(saldoabo) <= 0 Then
      MsgBox "No existe saldo", 48, "Aviso"
      tcampo1.SetFocus
      Exit Sub
   End If
   If found = 0 Then
      MsgBox "No existe Cliente o No tiene saldo ", 48, "Aviso"
      tcampo1.SetFocus
      Exit Sub
   End If
   tcampo5.SetFocus
   Exit Sub
End If
tcampo2.SetFocus
End Sub

Private Sub tcampo1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_xruc1
End If
'If KeyCode = &H26 Then
'   tcampo3.SetFocus
'   Exit Sub
'End If

End Sub

Private Sub tcampo2_DblClick()
tkeyboar.flag = "TCAMPO2"
tkeyboar.Show 1

End Sub

Private Sub tcampo2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
'If Len(tcampo2) = 0 Then
'   tcampo2.SetFocus
'   Exit Sub
'End If
tcampo3.SetFocus

End Sub

Private Sub tcampo2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   tcampo1.SetFocus
   Exit Sub
End If

End Sub

Private Sub tcampo3_DblClick()
tkeyboar.flag = "TCAMPO3"
tkeyboar.Show 1

End Sub

Private Sub tcampo3_KeyPress(KeyAscii As Integer)
Dim found1 As Double
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   Frame6.Visible = False
   habilita_lab7 0
   dbgrid10.Enabled = True
   dbgrid10.SetFocus
   Exit Sub
End If
saldoabo = ""
'If "" & dbgrid10.Columns("tipo") = "V" Then
'   If Len(tcampo1) = 0 Then
'      tcampo1.SetFocus
'      Exit Sub
'   End If
'   If Len(tcampo2) = 0 Then
'      tcampo2.SetFocus
'      Exit Sub
'   End If
'   tcampo5.SetFocus
'   Exit Sub
'End If
If "" & dbgrid10.columns("tipo") = "I" Or "" & dbgrid10.columns("tipo") = "K" Then  'valida el deposito bancario
   tcampo1.SetFocus
   Exit Sub
End If
If "" & dbgrid10.columns("tipo") = "D" Or "" & dbgrid10.columns("tipo") = "F" Then 'debito o credito
   If Len(tcampo3) < 4 Then
      tcampo3.SetFocus
      Exit Sub
   End If
End If


'If "" & dbgrid10.Columns("tipo") = "H" Then 'valida el deposito bancario
'   If Len(tcampo3) = 0 Then
'      tcampo3.SetFocus
'      Exit Sub
'   End If
'   found1 = valida_deposito("" & tcampo1, "" & tcampo3, 1)
'   If found1 <= 0 Then
'      MsgBox "No existe Saldo ", 48, "Aviso"
'      tcampo1.SetFocus
'      Exit Sub
'   End If
'   saldoabo = Format(found1, nrodecimal)
'End If
tcampo4.SetFocus
End Sub

Private Sub tcampo3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   tcampo2.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
If Frame6.Caption = "CHEQUE" Then  'consulta cheques
   consulta_banco '200
End If
If acufp = "V" Then   'si es vale
   consulta_vales  '2800
End If
If acufp = "I" Or acufp = "K" Then  'si es cruce de pago adelantado cruza
   consulta_credito  '2800
End If
End If
End Sub

Private Sub tcampo4_DblClick()
tkeyboar.flag = "TCAMPO4"
tkeyboar.Show 1

End Sub

Private Sub tcampo4_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If "" & dbgrid10.columns("tipo") = "D" Or "" & dbgrid10.columns("tipo") = "F" Then 'debito o credito
'If Len(tcampo4) = 0 Then
'   tcampo4.SetFocus
'   Exit Sub
'End If
End If
If "" & dbgrid10.columns("tipo") = "V" Or "" & dbgrid10.columns("tipo") = "C" Then  'debito o credito
   tcampo5.SetFocus
   Exit Sub
End If

If Len(tcampo4) > 0 Then
found = busca_banco("" & tcampo4)
If found = 0 Then
   MsgBox "Ingrese Entidad ", 48, "Aviso"
   tcampo4 = ""
   tcampo4.SetFocus
   Exit Sub
End If
End If
tcampo5.SetFocus

End Sub

Private Sub tcampo4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   tcampo3.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_banco
End If


End Sub

Private Sub tcampo5_DblClick()
tkeyboar.flag = "TCAMPO5"
tkeyboar.Show 1

End Sub

Private Sub tcampo5_KeyPress(KeyAscii As Integer)
Dim sdx As Double
Dim found As Integer
Dim found1 As Double
On Error GoTo cmd8_err
If KeyAscii <> 13 Then Exit Sub
'If Val(tcampo5) <= 0 Then
'   tcampo5 = "1"
'End If
'If Len(tcampo1) = 0 Then
'   tcampo1.SetFocus
'   Exit Sub
'End If
'If Len(tcampo2) = 0 Then
'   tcampo2.SetFocus
'   Exit Sub
'End If

If Val(RGPAGO) = 0 Then
               If fpmoneda = "S" Then
                  RGPAGO = ttxtotals
               End If
               If fpmoneda = "D" Then
                  RGPAGO = ttxtotald
               End If
End If

saldoabo = ""
If "" & dbgrid10.columns("tipo") = "D" Or "" & dbgrid10.columns("tipo") = "F" Then 'debito o credito
If "" & dbgrid10.columns("tipo") = "D" Then 'credito
If Len(tcampo3) = 0 Then
      tcampo3.SetFocus
      Exit Sub
End If
If Len(tcampo1) = 0 Then
   tcampo1.SetFocus
   Exit Sub
End If
If Len(tcampo2) = 0 Then
   tcampo2.SetFocus
   Exit Sub
End If
End If
If "" & dbgrid10.columns("tipo") = "V" Then 'vale
     If Len(tcampo1) = 0 Then
        tcampo1.SetFocus
        Exit Sub
     End If
If Len(tcampo2) = 0 Then
   tcampo2.SetFocus
   Exit Sub
End If

If Len(tcampo3) = 0 Then
   tcampo3.SetFocus
   Exit Sub
End If
If Len(tcampo4) = 0 Then
   tcampo4.SetFocus
   Exit Sub
End If
If Len(tcampo5) = 0 Then
   tcampo5.SetFocus
   Exit Sub
End If
End If

If "" & dbgrid10.columns("tipo") = "F" Then 'debito
If Len(tcampo3) = 0 Then
      tcampo3.SetFocus
      Exit Sub
End If
End If
End If


If "" & dbgrid10.columns("tipo") = "C" Or "" & dbgrid10.columns("tipo") = "G" Then 'c,g
   If "" & mytable11.Fields("obligacredito") = "S" Then
      found = credito_habilitado("" & tcampo1)
      If found = 0 Then
         MsgBox "Credito no permitido ", 48, "Aviso"
         tcampo1 = ""
         tcampo2 = ""
         tcampo1.SetFocus
         Exit Sub
      End If
   End If

If Len(tcampo1) = 0 Then
   tcampo1.SetFocus
   Exit Sub
End If
If Len(tcampo2) = 0 Then
   tcampo2.SetFocus
   Exit Sub
End If
End If

If ("" & dbgrid10.columns("tipo") = "C") And found = 1 Then    '
   If Len(tcampo1) = 0 Then
      tcampo1.SetFocus
      Exit Sub
   End If
   saldoabo = ""
   found = busca_credito_credito("" & dbgrid10.columns("tipo"), "" & tcampo1)
   
   'If Val(limite_credito) <= Val(saldoabo) Then
   '   MsgBox "No existe saldo", 48, "Aviso"
   '   tcampo1.SetFocus
   '   Exit Sub
   'End If
   'If found = 0 Then
   '   MsgBox "No existe Cliente o No tiene saldo ", 48, "Aviso"
   '   tcampo1.SetFocus
   '   Exit Sub
   'End If
End If


  


'If "" & dbgrid10.Columns("tipo") = "H" Then 'valida el deposito bancario
'   If Len(tcampo1) = 0 Then
'      tcampo1.SetFocus
'      Exit Sub
'   End If
'   If Len(tcampo3) = 0 Then
'      tcampo3.SetFocus
'      Exit Sub
'   End If
'    found = busca_codigocl("" & tcampo1, 0)
'   If found = 0 Then
'      MsgBox "No existe codigo ", 48, "Aviso"
'      tcampo1.SetFocus
'      Exit Sub
'   End If
'   found1 = valida_deposito("" & tcampo1, "" & tcampo3, 0)
'   If found1 <= 0 Then
'      MsgBox "No existe Saldo ", 48, "Aviso"
'      tcampo1.SetFocus
'      Exit Sub
'   End If
'   saldoabo = Format(found1, nrodecimal)
'End If
If ("" & dbgrid10.columns("tipo") = "I" Or "" & dbgrid10.columns("tipo") = "K") And found = 1 Then    '
   If Len(tcampo1) = 0 Then
      tcampo1.SetFocus
      Exit Sub
   End If
   saldoabo = ""
   found = busca_credito_adelanto("" & dbgrid10.columns("tipo"), "" & tcampo1)
   If found = 1 And Val(saldoabo) <= 0 Then
      MsgBox "No existe saldo", 48, "Aviso"
      tcampo1.SetFocus
      Exit Sub
   End If
   If found = 0 Then
      MsgBox "No existe Cliente o No tiene saldo ", 48, "Aviso"
      tcampo1.SetFocus
      Exit Sub
   End If
End If

codigo = tcampo1
nombre = tcampo2

If Val(RGPAGO) = 0 Then
   RGPAGO.SetFocus
   Exit Sub
End If


Data9.Recordset.AddNew
Data9.Recordset.Fields("descripcio") = "" & dbgrid10.columns(0)
Data9.Recordset.Fields("fpago") = "" & dbgrid10.columns(1)
Data9.Recordset.Fields("moneda") = "" & dbgrid10.columns(2)
Data9.Recordset.Fields("codigo") = tcampo1
Data9.Recordset.Fields("nombre") = tcampo2
Data9.Recordset.Fields("orden") = tcampo3
Data9.Recordset.Fields("observa") = tcampo4
Data9.Recordset.Fields("dias") = tcampo5
'Data9.Recordset.Fields("recibe") = tcampo5
Data9.Recordset.Fields("recibe") = Val(RGPAGO)
Data9.Recordset.Fields("acu") = "" & dbgrid10.columns("tipo")
Data9.Recordset.Update

If Len(tcampo1) > 0 And Len(tcampo2) > 0 Then
   found = graba_cliente_credito1("" & tcampo1)
End If

suma_fpagov
Frame6.Visible = False
habilita_lab7 0
found = leer_visorcaja("S/." & stxtotals, "US$  " & stxtotald)
           DBGrid9.Row = DBGrid9.VisibleRows - 1
               DBGrid9.Col = 2
               DBGrid9.SetFocus

'-----aqui verifica si va a cobrar el otro ------
          If Label45.Caption = "Vuelto" Or Val(stxtotals) <= 0 Then
             xtipo = protipo
             If "" & mytable11.Fields("habilitanota") = "S" Then
                If Val(ttxtotals) <= Val("" & mytable11.Fields("siventa")) Then
                   xtipo = "5"
                End If
             End If
                  xruc = codigo
                  xnombre = nombre
                  xvendedor = cproven
             Framefp.Enabled = False
             Frame7.Visible = True
             habilita_lab7 1
             Framefp.Enabled = False
             If flag_servicio <> "A" And flag_servicio <> "C" And flag_servicio <> "D" Then
                xtipo = "5"
             End If
             xtipo.SetFocus
          Exit Sub
         End If
         
         dbgrid10.Enabled = True
         dbgrid10.SetFocus
Exit Sub
'-----------
Frame6.Visible = False
habilita_lab7 0
               DBGrid9.Row = DBGrid9.VisibleRows - 1
               DBGrid9.Col = 2
               DBGrid9.SetFocus
Exit Sub
cmd8_err:
Exit Sub
End Sub
Function valida_deposito(buf0 As String, buf As String, sw As Integer) As Double
Dim mytablex As New ADODB.Recordset
Dim sdx As Double
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM chequemo where  transaccio='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   valida_deposito = Val("" & mytablex.Fields("saldo"))
   If sw = 1 Then
      tcampo1 = "" & mytablex.Fields("codigo")
   End If
End If
mytablex.Close
End Function
Sub graba_deposito(mytabley As ADODB.Recordset)
Dim mytablex As New ADODB.Recordset
Dim sdx As Double
Dim nrecibe As Double
On Error GoTo cmd7812_err
'ojo nrecibe siempres es igual o menor
nrecibe = Val("" & mytabley.Fields("recibe"))
If nrecibe = 0 Then Exit Sub

If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM chequemo where  codigo='" & Trim("" & mytabley.Fields("codigo")) & "' and transaccio='" & Trim("" & mytabley.Fields("orden")) & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
         'mytablex.Edit
         mytablex.Fields("abono") = Val("" & mytablex.Fields("abono")) + nrecibe
         sdx = Val("" & mytablex.Fields("neto")) - Val("" & mytablex.Fields("abono"))
         mytablex.Fields("saldo") = sdx
         mytablex.Update
End If
mytablex.Close
Exit Sub
cmd7812_err:
MsgBox "Error en graba deposito " + error$, 48, "Aviso"
Exit Sub
End Sub
Sub desgraba_deposito(mytabley As ADODB.Recordset)
Dim mytablex As New ADODB.Recordset
Dim sdx As Double
Dim nrecibe As Double
On Error GoTo cmd17812_err
nrecibe = Val("" & mytabley.Fields("recibe"))
If nrecibe = 0 Then Exit Sub

If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM chequemo where  codigo='" & Trim("" & mytabley.Fields("codigo")) & "' and transaccio='" & Trim("" & mytabley.Fields("orden")) & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
         'mytablex.Edit
         mytablex.Fields("abono") = Val("" & mytablex.Fields("abono")) - nrecibe
         sdx = Val("" & mytablex.Fields("neto")) - Val("" & mytablex.Fields("abono"))
         mytablex.Fields("saldo") = sdx
         mytablex.Update
End If
mytablex.Close
Exit Sub
cmd17812_err:
MsgBox "Error en graba deposito " + error$, 48, "Aviso"
Exit Sub

End Sub

Private Sub tcampo5_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   tcampo4.SetFocus
   Exit Sub
End If
End Sub

Private Sub telefono_DblClick()
tkeyboar.flag = "TELEFONO"
tkeyboar.Show 1

End Sub

Private Sub telefono_KeyPress(KeyAscii As Integer)
Dim mytablex As New ADODB.Recordset
Dim found As Integer
Dim buf As String
If KeyAscii <> 13 Then Exit Sub
If Len(telefono) < 7 Then
   telefono.SetFocus
   Exit Sub
End If
If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "SELECT * FROM deliveri where telefono like '" & telefono & "%'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      dcodigo.SetFocus
      Exit Sub
   End If
   found = consulta_cliente("" & telefono)
Exit Sub


End Sub
Function busca_deliveri()
End Function
Sub nuevo_dato()
Dim found As Integer
If found = 1 Then  'no existe en la data debe buscar en la principal
   found = busca_telefono("" & telefono)
End If
If found = 0 Then   'si no existe ningun fono debe crearse
   inicializa_data_deliveri
   If MsgBox("Cliente Nuevo,Desea Crear", 1, "Aviso") <> 1 Then
      inicialIzatodo
      telefono.SetFocus
      Exit Sub
   End If
      amsw = 1
      If dbclie.State = 1 Then dbclie.Close
      dbclie.Open "SELECT * FROM clientes", cn, adOpenDynamic, adLockOptimistic
      tnclie.telefono = telefono
      tnclie.moneda = "S"
      tnclie.Caption = "NUEVO"
      tnclie.Show 1
      dbclie.Close
      If Len(dcodigo) > 0 Then
         found = busca_codigod()
      End If
      amsw = 0
      dcodigo.SetFocus
   Exit Sub
End If
'poner los datos de los pedido
'poner_valores dotipo, doserie, donumero
sql_ver_pedido
'found = cuenta_telefonos()
'If found > 1 Then
'   consulta_delivery
'   Exit Sub
'End If
'found = busca_deliveri()
fechanac.SetFocus

End Sub
Sub inicializa_data_deliveri()
   clasificacion = ""
   dcodigo = ""
   dnombre = ""
   ddireccion = ""
   referencia = ""
   fechanac = ""
   'dotipo = ""
   'doserie = ""
   'donumero = ""
   'dototal = ""
   'dofpago = ""
   'dofecha = ""

End Sub
Sub consulta_banco()
Dim found As Integer
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.ListIndex = 0

Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "200"
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus

End Sub
Sub consulta_vales()
Dim found As Integer
Combo1.Clear
Combo1.AddItem "Numero"
Combo1.ListIndex = 0

Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "2003"
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus

End Sub

Sub consulta_credito()
Dim found As Integer
   
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.ListIndex = 0
   

Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "2800"
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus

End Sub

Function consulta_cliente(buf As String)
Dim found As Integer
Dim buf1 As String
buf1 = ""
If Len(buf) > 0 Then
   buf1 = " where telefono='" & buf & "'"
End If
   Combo1.Clear
   Combo1.AddItem "deliveri.telefono"
   Combo1.AddItem "Clientes.Nombre"
   Combo1.AddItem "deliveri.Direccion"
   Combo1.ListIndex = 0
Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
If Len(buf) > 0 Then
   buffer = buf
End If
opcion1 = "1"
sw_consulta = 0
found = sql_consulta(1)
consulta_cliente = 0

End Function
Function consulta_servicios(buf As String)
Dim found As Integer
Dim buf1 As String
buf1 = ""
'If Len(buf) > 0 Then
'   buf1 = " where telefono='" & buf & "'"
'End If
   Combo1.Clear
   Combo1.AddItem "Descripcio"
   Combo1.AddItem "servicio"
   Combo1.ListIndex = 0
Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
'If Len(buf) > 0 Then
'   buffer = buf
'End If
opcion1 = "19000"
sw_consulta = 0
found = sql_consulta(1)
consulta_servicios = 0

End Function

Sub consulta_xvendedor()
Dim found As Integer
   Combo1.Clear
   Combo1.AddItem "Nombre"
   Combo1.ListIndex = 0


Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "31"
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus

End Sub

Sub consulta_xruc()
Dim vr
Dim found As Integer
   
   Combo1.Clear
   Combo1.AddItem "Nombre"
   Combo1.AddItem "Codigo"
   Combo1.ListIndex = 0

Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "30"
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus


End Sub
Sub consulta_xruc2()
Dim found As Integer
   Combo1.Clear
   Combo1.AddItem "Nombre"
   Combo1.ListIndex = 0

Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "300"
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus

End Sub
Sub consulta_xruc1()
Dim found As Integer
   Combo1.Clear
   Combo1.AddItem "Nombre"
   Combo1.ListIndex = 0

Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "99"
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus

End Sub

Sub consulta_xtipo()
Dim found As Integer
   Combo1.Clear
   Combo1.AddItem "Descripcio"
   Combo1.ListIndex = 0


Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "29"
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus

End Sub

Sub consulta_cliente1()
Dim found As Integer
Frame1.Visible = True
Frame1.Enabled = True
Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.ListIndex = 0
buffer = ""
opcion1 = "12"
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus

End Sub
Sub consulta_clientefp()
Dim found As Integer
Frame1.Visible = True
Frame1.Enabled = True
   Combo1.Clear
   Combo1.AddItem "Nombre"
   Combo1.ListIndex = 0
buffer = ""
opcion1 = "23"
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus

End Sub



Function busca_codigod()
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM clientes where  codigo='" & dcodigo & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   clasificacion = "" & mytablex.Fields("clasifica")
   dnombre = "" & mytablex.Fields("nombre")
   ddireccion = "" & mytablex.Fields("direccion")
   fechanac = "" & mytablex.Fields("fechanac")
   referencia = "" & mytablex.Fields("observa")
   saludo_cumpe
   'dotipo = "" & mytablex.Fields("dotipo")
   'doserie = "" & mytablex.Fields("doserie")
   'donumero = "" & mytablex.Fields("donumero")
   'ruc = "" & mytablex.Fields("codigo1")
End If
'------------------------------------- ------------
mytablex.Close
End Function
Function busca_cupo()
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM clientes where  codigo='" & dcodigo & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
End If
'------------------------------------- ------------
mytablex.Close

End Function
Function busca_codigocl(buf As String, sw As Integer)
Dim mytablex As New ADODB.Recordset
Dim buf1 As String
limite_credito = ""
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM clientes where  codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   limite_credito = "" & mytablex.Fields("credito")
   If sw = 0 Then
   
      codigo = "" & mytablex.Fields("codigo")
      nombre = "" & mytablex.Fields("nombre")
      tcampo2 = "" & mytablex.Fields("nombre")
   End If
   If sw = 1 Then
      xruc = "" & mytablex.Fields("codigo")
      If Len(xnombre) = 0 Then
        xnombre = "" & mytablex.Fields("nombre")
      End If
      If Len(xdireccion) = 0 Then
         xdireccion = "" & mytablex.Fields("direccion")
      End If
   End If
   If dbgrid10.columns("tipo") = "V" Then 'si en fpago es vale
      totpedido = "" '& suma_pedidos("" & Data9.Recordset.Fields("codigo"), "" & Data9.Recordset.Fields("orden"), "" & Data9.Recordset.Fields("observa"), "" & Data9.Recordset.Fields("dias"))
   End If
   busca_codigocl = 1
End If
mytablex.Close
End Function
Function busca_localx(buf As String)
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM bodega where  codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   xnombre = "" & mytablex.Fields("nombre")
   xdireccion = "" & mytablex.Fields("direccion")
   'xdistrito = "" & mytablex.Fields("distrito")
   busca_localx = 1
End If
mytablex.Close
End Function
Function busca_local_pedido(buf As String) As String
Dim mytablex As New ADODB.Recordset

If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM tlocal where  codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   busca_local_pedido = "" & mytablex.Fields("nombre")
End If
mytablex.Close
End Function
Function busca_telefono(buf As String)
Dim mytablex As New ADODB.Recordset
Dim indx As Integer
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM clientes where  codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   pone_datos_1 mytablex
   busca_telefono = 1
End If
mytablex.Close
 
End Function
Sub pone_datos_1(mytablex As ADODB.Recordset)
   clasificacion = "" & mytablex.Fields("clasifica")
   dcodigo = "" & mytablex.Fields("codigo")
   dcodigo = "" & mytablex.Fields("codigo")
   dnombre = "" & mytablex.Fields("nombre")
   ddireccion = "" & mytablex.Fields("direccion")
   fechanac = "" & mytablex.Fields("fechanac")
   referencia = "" & mytablex.Fields("observa")
   saludo_cumpe

   'dotipo = "" & mytablex.Fields("dotipo")
   'doserie = "" & mytablex.Fields("doserie")
   'donumero = "" & mytablex.Fields("donumero")
   'ruc = "" & mytablex.Fields("codigo1")

End Sub
Function contador_telefonos(buf As String)
Dim mytablex As New ADODB.Recordset
Dim indx As Integer
indx = 0
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM telefono where  telefono='" & telefono & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   Do
   If mytablex.EOF Then Exit Do
      indx = indx + 1
      buf = "" & mytablex.Fields("codigo")
      mytablex.MoveNext
   Loop
End If
mytablex.Close

contador_telefonos = indx
End Function
Function valida()
Dim sdx As Double
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
If Len(telefono) < 6 Then
   telefono.SetFocus
   Exit Function
End If
If Len(dnombre) = 0 Then
   dnombre.SetFocus
   Exit Function
End If
If Len(ddireccion) = 0 Then
   ddireccion.SetFocus
   Exit Function
End If
If Len(fechanac) > 0 Then
   If valida_fecha(fechanac) = 0 Then
      fechanac = ""
      fechanac.SetFocus
      Exit Function
   End If
End If
'crea el cliente y valida la existencia del cliente
   If Len(dcodigo) = 0 Then
      busca_correlativo 0
   End If
      If mytabley.State = 1 Then mytabley.Close
      
      mytabley.Open "SELECT * FROM clientes where codigo='" & dcodigo & "'", cn, adOpenDynamic, adLockOptimistic
      If mytabley.RecordCount > 0 Then
         mytabley.Fields("nombre") = dnombre
         mytabley.Fields("direccion") = ddireccion
         mytabley.Fields("observa") = referencia
         mytabley.Fields("telefono") = telefono
         mytabley.Fields("clasifica") = Trim(clasificacion)
         If IsDate(fechanac) Then
            mytabley.Fields("fechanac") = fechanac
         End If
         mytabley.Update
         Else
         mytabley.AddNew
         mytabley.Fields("codigo") = dcodigo
         mytabley.Fields("tipo") = "O"
         mytabley.Fields("nombre") = dnombre
         mytabley.Fields("moneda") = "" & mytable11.Fields("moneda")
         mytabley.Fields("direccion") = ddireccion
         mytabley.Fields("observa") = referencia
         mytabley.Fields("telefono") = telefono
         If IsDate(fechanac) Then
            mytabley.Fields("fechanac") = fechanac
         End If
         mytabley.Fields("clasifica") = Trim(clasificacion)
         mytabley.Update
         busca_correlativo 1
      End If
      mytabley.Close
      mytablex.Open "SELECT * FROM deliveri where codigo='" & dcodigo & "' and telefono='" & telefono & "'", cn, adOpenDynamic, adLockOptimistic
      If mytablex.RecordCount = 0 Then
         mytablex.AddNew
         mytablex.Fields("telefono") = telefono
         mytablex.Fields("codigo") = dcodigo
         mytablex.Fields("nombre") = dnombre
         mytablex.Fields("direccion") = ddireccion
         mytablex.Fields("referencia") = referencia
         mytablex.Update
         Else
         'mytablex.Fields("telefono") = telefono
         'mytablex.Fields("codigo") = dcodigo
         mytablex.Fields("nombre") = dnombre
         mytablex.Fields("direccion") = ddireccion
         mytablex.Fields("referencia") = referencia
         mytablex.Update
      End If
      mytablex.Close
      valida = 1
End Function
Function verifica_doble(buf As String)
Dim mytabley As Table

Set mytabley = mydbxglo.OpenTable(dgusuario)
mytabley.Index = "producto"
mytabley.Seek "=", buf
If Not mytabley.NoMatch Then
   verifica_doble = 1
End If
mytabley.Close
 
End Function
Function busca_equiva(buf As String) As Integer
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "SELECT * FROM productb where barras='" & buf & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      buf = "" & mytablex.Fields("producto")
      busca_equiva = 1
      mytablex.Close
      Exit Function
   End If
   mytablex.Close
   
   mytablex.Open "SELECT * FROM producto where barras='" & buf & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      buf = "" & mytablex.Fields("producto")
      busca_equiva = 1
   End If
   mytablex.Close
End Function

Function busca_producto(buf As String, sw As Integer, canti As String, xsw As Integer)
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
Dim buf1 As String
Dim i As Integer
Dim ssw As Integer
Dim found As Integer
'------------------------------------
'verificamos si es codigo barras
i = 0
      found = 0
      If mytablex.State = 1 Then mytablex.Close
      mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic
      If mytablex.RecordCount = 0 Then
         mytablex.Close
         found = busca_equiva(buf) 'busca en la table codigo barras
         If found = 0 Then
            Exit Function
         End If
         mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic
         If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Function
         End If
      End If
'MsgBox "def"
   If "" & mytablex.Fields("estado") = "N" Then  'si no esta activo
      MsgBox "Producto no activo ", 48, "Aviso"
      mytablex.Close
      Exit Function
   End If
   'MsgBox "abc"
   
      If mytabley.State = 1 Then mytabley.Close
      mytabley.Open "SELECT * FROM precios where producto='" & buf & "' and local='" & "" & mytable11.Fields("listap") & "'", cn, adOpenStatic, adLockOptimistic
      If mytabley.RecordCount = 0 Then
         MsgBox "No existe precio alguno ", 48, "Aviso"
         mytabley.Close
         mytablex.Close
         Exit Function
      End If

   'MsgBox "abc"
   If Val("" & mytabley.Fields("pventa1")) <= 0 Then
      If "" & mytable11.Fields("noprecio") = "S" Then
         MsgBox "" & mytablex.Fields("descripcio") & "  Precio <=0 No Permitido ", 48, "Aviso"
         mytablex.Close
         busca_producto = 2
         Exit Function
      End If
      If "" & mytablex.Fields("remate") <> "S" Then
         'MsgBox "" & mytablex.Fields("descripcio") & "  Precio <=0", 48, "Aviso"
         'mytablex.Close
         'busca_producto = 2
         'Exit Function
      End If
   End If
   'End If
   'canti = ""
   buf = ""
   
   '----------- verfica a forzar la balanza
   If Val(canti) <= 0 Then
   If "" & mytable11.Fields("actbala") = "S" Then
     If "" & mytablex.Fields("peso") = "S" Then
ajk91:
       'MsgBox "busca Producto"
        
        buf = puerto_balanza1()
        If Val(buf) = 0 Then
           If MsgBox("Balanza No leido,Continua Leyendo? ", 1, "Aviso") = 1 Then
              GoTo ajk91
              '------
              Else
              'MsgBox "No leido ", 48, "Aviso"
              busca_producto = 2
              mytablex.Close
             Exit Function
           End If
        End If
     End If
   End If
   canti = Format(Val(buf), "0.000")
   'canti = buf
   End If
   If Val(canti) <= 0 Then
      canti = "1"
   End If
   'MsgBox canti
   busca_producto = 1
   '---------------------------------------
   If sw = 0 Or sw = 2 Then
      graba_temporald mytablex, sw, canti, mytabley, xsw
   End If

mytablex.Close
mytabley.Close
End Function
Sub calcula_igv(sw As Integer)
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double
Dim tdscto As Double
Dim tdscto1 As Double
Dim found As Integer
Dim xtivap As Double
Dim xtisc As Double
Dim xdetra As Double
On Error GoTo cmd4567_err
'DBGrid3.Columns("subtotal") = xneto
'DBGrid3.Columns("impuesto") = xdescuento
'DBGrid3.Columns(9) = xsubtotal
'DBGrid3.Columns("isc") = ximpuesto
'DBGrid3.Columns("total") = xtotal
'MsgBox ""
'-------------------------------------------------------------
DBGrid2.columns("neto") = Val("" & DBGrid2.columns("cantidad")) * Val("" & DBGrid2.columns("precio"))
tdscto = Val("" & DBGrid2.columns("neto")) * Val("" & DBGrid2.columns("deslipo")) / 100       'calcular descuento
DBGrid2.columns("descuento") = tdscto  'total descuento
DBGrid2.columns("total") = Val("" & DBGrid2.columns("neto")) - Val("" & DBGrid2.columns("descuento")) 'cobrar
xtivap = Val("" & DBGrid2.columns("total")) * Val("" & DBGrid2.columns("ivap")) / 100
DBGrid2.columns("tivap") = xtivap
   sdx2 = 1 + Val("" & DBGrid2.columns("igv")) / 100
   sdx1 = Val(DBGrid2.columns("total")) / sdx2
   DBGrid2.columns("subtotal") = sdx1  'subtotal
   sdx = Val("" & DBGrid2.columns("total")) - Val("" & DBGrid2.columns("subtotal"))
   DBGrid2.columns("impuesto") = sdx  'impuesto
   xtisc = Val("" & DBGrid2.columns("subtotal")) * Val("" & DBGrid2.columns("isc")) / 100
   DBGrid2.columns("tisc") = xtisc
   DBGrid2.columns("tax") = 0
   If Val("" & DBGrid2.columns("igv")) = 0 Then
      DBGrid2.columns("tax") = Val("" & DBGrid2.columns("total"))
      DBGrid2.columns("impuesto") = 0
   End If
Exit Sub
cmd4567_err:
MsgBox "Error en Calcula Igv " + error$, 48, "Aviso"
Exit Sub
End Sub
Sub calcula_sinigv()
Dim sdx As Double
Dim sdx1 As Double
Dim sdx2 As Double
Dim found As Integer
'debe sumar el igv
'DBGrid2.Columns("neto") = Val("" & DBGrid2.Columns("cantidad")) * Val("" & DBGrid2.Columns("precio"))
If Val("" & DBGrid2.columns("cantidad")) > 0 And Val("" & DBGrid2.columns("neto")) > 0 Then
   sdx = Val("" & DBGrid2.columns("neto")) * Val("" & DBGrid2.columns("deslipo")) / 100 'descuento
   DBGrid2.columns("descuento") = sdx  'descuento
   DBGrid2.columns("subtotal") = Val("" & DBGrid2.columns("neto")) - Val("" & DBGrid2.columns("descuento")) 'subtotal
   sdx = Val("" & DBGrid2.columns("subtotal")) * Val("" & DBGrid2.columns("igv")) / 100
   DBGrid2.columns("impuesto") = sdx 'igv
   DBGrid2.columns("total") = Val("" & DBGrid2.columns("subtotal")) + sdx 'neto
   sdx = Val("" & DBGrid2.columns("total")) / Val(DBGrid2.columns("cantidad"))
   DBGrid2.columns("precio") = sdx
End If

End Sub
Function consulta_producto(buf As String)
Dim found As Integer
Dim xbuf As String
xbuf = ""
If Len(buf) > 0 Then
   xbuf = buf '" AND descripcio like '" & buf & "%'"
End If
Combo1.Clear
Combo1.AddItem "producto.Descripcio"
Combo1.AddItem "producto.Producto"
Combo1.AddItem "producto.Familia"
Combo1.AddItem "producto.Marca"
Combo1.ListIndex = 0
Frame1.Enabled = True
Frame1.Visible = True

buffer = xbuf
opcion1 = "8"
DBGrid2.Enabled = False
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus
End Function
Function consulta_inicial(buf As String)
Dim buf1 As String
Dim queprecio As String
   Combo1.Clear
   Combo1.AddItem "Producto.Descripcio"
   Combo1.ListIndex = 0

queprecio = "precioS.pventa1 as Precio "
If Len(buf) > 0 Then
buf1 = "select Producto.Descripcio,Producto.producto,Producto.Marca,precios.Unidad1 as Und1,precios.Factor1 as F," & queprecio & " ,Producto.Monedav as M,Producto.Familia,Producto.Subfamilia,Producto.seccion from producto  left join precios on producto.producto=precios.producto  where precios.local='" & "" & mytable11.Fields("listap") & "' and producto.descripcio like '" & buf & "%'"
End If
If Len(buf) = 0 Then
buf1 = "select Producto.Descripcio,Producto.producto,Producto.Marca,precios.Unidad1 as Und1,precios.Factor1 as F," & queprecio & " ,Producto.Monedav as M,Producto.Familia,Producto.Subfamilia,Producto.seccion from producto  left join precios on producto.producto=precios.producto  where precios.local='" & "" & mytable11.Fields("listap") & "'"
End If
   
   If rcconsulta.State = 1 Then rcconsulta.Close
   rcconsulta.Open buf1, cn, adOpenStatic, adLockOptimistic
   If rcconsulta.EOF = True And rcconsulta.BOF = True Then
      rcconsulta.Close
      buffer.SetFocus
      Exit Function
   End If
   
   Set dbGrid1.DataSource = rcconsulta
               dbGrid1.columns(0).Width = 5000
               dbGrid1.columns(1).Width = 1300
               dbGrid1.columns(2).Width = 1000
               dbGrid1.columns("cantidad").Width = 900
               dbGrid1.columns(4).Width = 500
               dbGrid1.columns("precio").Width = 800
               dbGrid1.columns("deslipo").Width = 500
               dbGrid1.columns("total").Width = 1000
               dbGrid1.columns("isc").Width = 1500
               dbGrid1.columns(9).Width = 1500
               'End If
consulta_inicial = 1
End Function
Sub cerrar_data2()
On Error GoTo cmd4_err
Data2.Recordset.Close
Exit Sub
cmd4_err:
Exit Sub
End Sub
Sub carga_dbgrid4(uproducto As String)
Dim i As Integer
Dim xfoto As String
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
Dim sw As Integer
Dim xbodega As String
Dim xsaldo As Double
Dim xbuf As String
Dim xcosto As Double
Dim xmargen As Double
Dim xcostou As Double
Dim xfactor As Double
Dim xxr As String
Dim xxi As String
Dim xpreciox As Double
Dim dmoneda As String
On Error GoTo cmd89111_err
xcostou = 0
For i = 0 To 9
    campo_precios(i).unidad = ""
    campo_precios(i).factor = ""
    campo_precios(i).precio = ""
    campo_precios(i).costo = ""
    campo_precios(i).margen = ""
    campo_precios(i).stock = ""
Next i
'MsgBox uproducto
xfactor = 1
xbodega = "" & mytable11.Fields("bodega")
xsaldo = 0
xcosto = 0
sw = 0
      If mytabley.State = 1 Then mytabley.Close
      mytabley.Open "SELECT * FROM almacen where local='" & "" & mytable11.Fields("local") & "' and producto='" & uproducto & "' and bodega='" & xbodega & "'", cn, adOpenStatic, adLockOptimistic
      If mytabley.RecordCount > 0 Then
         xsaldo = Val("" & mytabley.Fields("saldo"))
      End If
      mytabley.Close
'MsgBox "x"
'---buscamos los datos de productos
dmoneda = "S"
xfoto = ""
descorto = ""
      If mytablex.State = 1 Then mytablex.Close
      mytablex.Open "SELECT * FROM producto where  producto='" & uproducto & "'", cn, adOpenStatic, adLockOptimistic
      If mytablex.RecordCount > 0 Then
         xcostou = 0
         If "" & mytable11.Fields("vecocaja") = "S" Then
            xcostou = Val("" & mytablex.Fields("costou"))
         End If
         xfactor = Val("" & mytablex.Fields("factor"))
         descorto = "" & mytablex.Fields("presenta")
         dmoneda = "" & mytablex.Fields("monedav")
         xfoto = "" & mytablex.Fields("fotonombre")
      End If
      mytablex.Close
      carga_foto xfoto
      If Val(paridad) <= 0 Then
         paridad = "1"
      End If
      'MsgBox "abc"
      If mytablex.State = 1 Then mytablex.Close
      mytablex.Open "SELECT * FROM precios where  producto='" & uproducto & "' and local='" & "" & mytable11.Fields("listap") & "'", cn, adOpenStatic, adLockOptimistic
      If mytablex.RecordCount > 0 Then
         xcosto = 0
         xpreciox = 0
         If Val("" & mytablex.Fields("factor1")) > 0 Then
            If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
               xpreciox = Val("" & mytablex.Fields("pventa1"))
               If dmoneda = "D" Then
                  xpreciox = Val("" & mytablex.Fields("pventa1")) * Val(paridad)
               End If
            End If
            If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
               xpreciox = Val("" & mytablex.Fields("pventa1"))
               If dmoneda = "S" Then
                  xpreciox = Val("" & mytablex.Fields("pventa1")) / Val(paridad)
               End If
            End If
            'MsgBox "abc"
           '------------------------------------------------------------
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor1"))
            campo_precios(0).unidad = "" & mytablex.Fields("unidad1")
            campo_precios(0).factor = "" & Val("" & mytablex.Fields("factor1"))
            campo_precios(0).precio = "" & xpreciox
            campo_precios(0).costo = "" & xcosto
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor1")))
            campo_precios(0).stock = "" & xbuf
            xmargen = 0
            If xcosto > 0 Then
               xmargen = ((xpreciox - xcosto) * 100) / xcosto
            End If
            campo_precios(0).margen = "" & xmargen
         End If
   '---------
   'MsgBox "abc"
   xcosto = 0
   If Val("" & mytablex.Fields("factor2")) > 0 Then
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa2"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa2")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa2"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa2")) / Val(paridad)
      End If
   End If
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor2"))
   campo_precios(1).unidad = "" & mytablex.Fields("unidad2")
   campo_precios(1).factor = "" & Val("" & mytablex.Fields("factor2"))
   campo_precios(1).precio = "" & xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor2")))
   campo_precios(1).stock = "" & xbuf
   campo_precios(1).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
      xmargen = ((xpreciox - xcosto) * 100) / xcosto
   End If
   campo_precios(1).margen = "" & xmargen
   End If
   xcosto = 0
   If Val("" & mytablex.Fields("factor3")) > 0 Then
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa3"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa3")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa3"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa3")) / Val(paridad)
      End If
   End If

   'MsgBox "abc"
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor3"))
   campo_precios(2).unidad = "" & mytablex.Fields("unidad3")
   campo_precios(2).factor = "" & Val("" & mytablex.Fields("factor3"))
   campo_precios(2).precio = "" & xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor3")))
   campo_precios(2).stock = "" & xbuf
   campo_precios(2).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((xpreciox - xcosto) * 100) / xcosto
         campo_precios(2).margen = "" & xmargen
   End If
   campo_precios(2).margen = "" & xmargen
   End If
   xcosto = 0
   If Val("" & mytablex.Fields("factor4")) > 0 Then
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa4"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa4")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa4"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa4")) / Val(paridad)
      End If
   End If
'MsgBox "abc"
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor4"))
   campo_precios(3).unidad = "" & mytablex.Fields("unidad4")
   campo_precios(3).factor = "" & Val("" & mytablex.Fields("factor4"))
   campo_precios(3).precio = "" & xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor4")))
   campo_precios(3).stock = "" & xbuf
   campo_precios(3).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((xpreciox - xcosto) * 100) / xcosto
   End If
   campo_precios(3).margen = "" & xmargen
   End If
   xcosto = 0
   
   If Val("" & mytablex.Fields("factor5")) > 0 Then
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa5"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa5")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa5"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa5")) / Val(paridad)
      End If
   End If
'MsgBox "abc"
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor5"))
      campo_precios(4).unidad = "" & mytablex.Fields("unidad5")
   campo_precios(4).factor = "" & Val("" & mytablex.Fields("factor5"))
   campo_precios(4).precio = "" & xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor5")))
   campo_precios(4).stock = "" & xbuf
   campo_precios(4).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((xpreciox - xcosto) * 100) / xcosto
   End If
   campo_precios(4).margen = "" & xmargen
   End If
   xcosto = 0
   
   If Val("" & mytablex.Fields("factor6")) > 0 Then
   
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa6"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa6")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa6"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa6")) / Val(paridad)
      End If
   End If

   'MsgBox "abcD"
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor6"))
   campo_precios(5).unidad = "" & mytablex.Fields("unidad6")
   campo_precios(5).factor = "" & Val("" & mytablex.Fields("factor6"))
   campo_precios(5).precio = "" & xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor6")))
   campo_precios(5).stock = "" & xbuf
   campo_precios(5).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
      xmargen = ((xpreciox - xcosto) * 100) / xcosto
   End If
   campo_precios(5).margen = "" & xmargen
   
   'SOLO PARA MAXIMO SE PONE PRECIO=0
   'If caja <> "08" Then
   '   campo_precios("precio").precio = 0
   'End If
   End If
   'MsgBox "xx"
   xcosto = 0
   If Val("" & mytablex.Fields("factor7")) > 0 Then
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa7"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa7")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa7"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa7")) / Val(paridad)
      End If
   End If

   'MsgBox "abcde"
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor7"))
   
   campo_precios(6).unidad = "" & mytablex.Fields("unidad7")
   campo_precios(6).factor = "" & Val("" & mytablex.Fields("factor7"))
   campo_precios(6).precio = "" & xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor7")))
   campo_precios(6).stock = "" & xbuf
   campo_precios(6).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
         xmargen = ((xpreciox - xcosto) * 100) / xcosto
         
   End If
   campo_precios(6).margen = "" & xmargen
   End If
   
   xcosto = 0
   If Val("" & mytablex.Fields("factor8")) > 0 Then
   
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa8"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa8")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa8"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa8")) / Val(paridad)
      End If
   End If

      xcosto = xcostou / xfactor
      xcosto = xcosto * Val("" & mytablex.Fields("factor8"))
   
   campo_precios(7).unidad = "" & mytablex.Fields("unidad8")
   campo_precios(7).factor = "" & Val("" & mytablex.Fields("factor8"))
   campo_precios(7).precio = "" & xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor8")))
   campo_precios(7).stock = "" & xbuf
   campo_precios(7).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
      xmargen = ((xpreciox - xcosto) * 100) / xcosto
   End If
   campo_precios(7).margen = "" & xmargen
   End If
   xcosto = 0
   If Val("" & mytablex.Fields("factor9")) > 0 Then
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa9"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa9")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa9"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa9")) / Val(paridad)
      End If
   End If

   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor9"))
   campo_precios(8).unidad = "" & mytablex.Fields("unidad9")
   campo_precios(8).factor = "" & Val("" & mytablex.Fields("factor9"))
   campo_precios(8).precio = "" & xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor9")))
   campo_precios(8).stock = "" & xbuf
   campo_precios(8).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
      xmargen = ((xpreciox - xcosto) * 100) / xcosto
   End If
   campo_precios(8).margen = "" & xmargen
   End If
   xcosto = 0
   If Val("" & mytablex.Fields("factor10")) > 0 Then
      If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa10"))
      If dmoneda = "D" Then
         xpreciox = Val("" & mytablex.Fields("pventa10")) * Val(paridad)
      End If
   End If
   If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
      xpreciox = Val("" & mytablex.Fields("pventa10"))
      If dmoneda = "S" Then
         xpreciox = Val("" & mytablex.Fields("pventa10")) / Val(paridad)
      End If
   End If
   
   xcosto = xcostou / xfactor
   xcosto = xcosto * Val("" & mytablex.Fields("factor10"))
   campo_precios(9).unidad = "" & mytablex.Fields("unidad10")
   campo_precios(9).factor = "" & Val("" & mytablex.Fields("factor10"))
   campo_precios(9).precio = "" & xpreciox
   xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor10")))
   campo_precios(9).stock = "" & xbuf
   campo_precios(9).costo = "" & xcosto
   xmargen = 0
   If xcosto > 0 Then
      xmargen = ((xpreciox - xcosto) * 100) / xcosto
   End If
   campo_precios(9).margen = "" & xmargen
   End If
   'MsgBox "xx"
   sql_saldo_locales uproducto
   'margenes
   sw = 1
End If

'mytablex.Close
'mytablez.Close
DBGrid4.refresh


'----ahora deb cargar tambien la foto del producto...
Frame1.Enabled = False
Frame5.Visible = True
Frame5.Enabled = True
DBGrid4.SetFocus
Exit Sub
cmd89111_err:
MsgBox "Error en carga dbgrid4 " + error$, 48, "Aviso"
Exit Sub
End Sub
Sub ir_ultimo()
Dim found As Integer
On Error GoTo cmd50_err
found = sumar_detalle()
DBGrid2.Col = 0
DBGrid2.Row = DBGrid2.VisibleRows - 1
DBGrid2.SetFocus
Exit Sub
cmd50_err:
MsgBox "Error en Ir-ultimo " + error$, 48, "Aviso"
Data2.refresh
DBGrid2.SetFocus
Exit Sub
End Sub
Sub ir_primero()
On Error GoTo cmd51_err
Data2.Recordset.MoveFirst
Exit Sub
cmd51_err:
Exit Sub

End Sub
Function busca_linea(buf As String)
Dim mytablex As New ADODB.Recordset

If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM linea where  linea='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   busca_linea = 1
   tingtalla.nlinea = "" & mytablex.Fields("descripcio")
   tingtalla.nt1 = "" & mytablex.Fields("t1")
   tingtalla.nt2 = "" & mytablex.Fields("t2")
   tingtalla.nt3 = "" & mytablex.Fields("t3")
   tingtalla.nt4 = "" & mytablex.Fields("t4")
   tingtalla.nt5 = "" & mytablex.Fields("t5")
   tingtalla.nt6 = "" & mytablex.Fields("t6")
   tingtalla.nt7 = "" & mytablex.Fields("t7")
   tingtalla.nt8 = "" & mytablex.Fields("t8")
   tingtalla.nt9 = "" & mytablex.Fields("t9")
   tingtalla.nt10 = "" & mytablex.Fields("t10")
   tingtalla.nt11 = "" & mytablex.Fields("t11")
   tingtalla.nt12 = "" & mytablex.Fields("t12")
   tingtalla.nt13 = "" & mytablex.Fields("t13")
   tingtalla.nt14 = "" & mytablex.Fields("t14")
   tingtalla.nt15 = "" & mytablex.Fields("t15")
   tingtalla.nt16 = "" & mytablex.Fields("t16")
End If
'------------------------------------- ------------
mytablex.Close
End Function
Sub ingreso_tallas(buf As String)
Dim found As Integer
found = busca_linea(buf)
If found = 0 Then Exit Sub
pone_tallas buf
tingtalla.Show 1
'Frame3.Visible = True
't1.SetFocus
menu_fin_tallas
End Sub
Sub pone_tallas(buf As String)
tingtalla.linea = buf
tingtalla.t1 = "" & DBGrid2.columns(18)
tingtalla.t2 = "" & DBGrid2.columns(19)
tingtalla.t3 = "" & DBGrid2.columns(20)
tingtalla.t4 = "" & DBGrid2.columns(21)
tingtalla.t5 = "" & DBGrid2.columns(22)
tingtalla.t6 = "" & DBGrid2.columns(23)
tingtalla.t7 = "" & DBGrid2.columns(24)
tingtalla.t8 = "" & DBGrid2.columns(25)
tingtalla.t9 = "" & DBGrid2.columns(26)
tingtalla.t10 = "" & DBGrid2.columns(27)
tingtalla.t11 = "" & DBGrid2.columns(28)
tingtalla.t12 = "" & DBGrid2.columns(29)
tingtalla.t13 = "" & DBGrid2.columns(30)
tingtalla.t14 = "" & DBGrid2.columns(31)
tingtalla.t15 = "" & DBGrid2.columns(32)
tingtalla.t16 = "" & DBGrid2.columns(33)
End Sub
Sub ingreso_locales()
On Error GoTo cmd11200_err
xxpone_locales
txobserv.Show 1
DBGrid2.Col = 0
               DBGrid2.Row = DBGrid2.VisibleRows - 1
               DBGrid2.SetFocus
Exit Sub
cmd11200_err:
Exit Sub
'If acu = "R" Then 'si no es orden de compra
'   l1.Enabled = False
'   l2.Enabled = False
'   l3.Enabled = False
'   l4.Enabled = False
'End If
'l1.SetFocus
End Sub
Sub xxpone_locales()
Dim found As Integer
txobserv.observa1 = "" & DBGrid2.columns(39)
txobserv.observa2 = "" & DBGrid2.columns(40)
txobserv.observa3 = "" & DBGrid2.columns(41)
txobserv.observa4 = "" & DBGrid2.columns(42)
End Sub
Sub cerrar_data1()
On Error GoTo cmd17_err
Data1.Recordset.Close
Exit Sub
cmd17_err:
Exit Sub
End Sub
Sub graba_temporald(mytablex As ADODB.Recordset, sw As Integer, canti As String, mytabley As ADODB.Recordset, xsw As Integer)
Dim fechadi As String
Dim deslipox As Double
Dim found As Integer
Dim xxca As String
Dim sdx As Double
Dim dsdx As Double
Dim xpreciox As Double
Dim mytables As New ADODB.Recordset
Dim mytablez As New ADODB.Recordset
xxca = "1"
If Val(canti) > 0 Then
   xxca = "" & canti
End If
'MsgBox xxca
xpreciox = 0
deslipox = 0
dsdx = 0
If Val(paridad) <= 0 Then
   paridad = "1"
End If
If "" & mytable11.Fields("moneda") = "S" Then 'si es soles
   xpreciox = Val("" & mytabley.Fields("pventa1"))
   If "" & mytablex.Fields("monedav") = "D" Then
      xpreciox = Val("" & mytabley.Fields("pventa1")) * Val(paridad)
   End If
End If
If "" & mytable11.Fields("moneda") = "D" Then 'si es soles
   xpreciox = Val("" & mytabley.Fields("pventa1"))
   If "" & mytablex.Fields("monedav") = "S" Then
      xpreciox = Val("" & mytabley.Fields("pventa1")) / Val(paridad)
   End If
End If

'----verificamos si el cliente tiene descuento---
dsdx = 0
If Len(codigo) > 0 And "" & mytablex.Fields("remate") <> "S" Then
     If mytablez.State = 1 Then mytablez.Close
      mytablez.Open "SELECT * FROM clientes where codigo='" & codigo & "'", cn, adOpenStatic, adLockOptimistic
      If mytablez.RecordCount > 0 Then
         dsdx = xpreciox * Val("" & mytablez.Fields("descuento")) / 100
      End If
      xpreciox = xpreciox - dsdx
End If



'ver si esta en un rango de descuento------------
   If Len("" & mytabley.Fields("fechaid")) = 10 And Len("" & mytabley.Fields("fechafd")) = 10 Then
   If IsDate("" & mytabley.Fields("fechaid")) And IsDate("" & mytabley.Fields("fechafd")) And CVDate("" & mytabley.Fields("fechafd")) >= CVDate("" & mytabley.Fields("fechaid")) And Val("" & mytabley.Fields("dscto")) > 0 Then
      fechadi = Format(Now, "dd/mm/yyyy")
      If CVDate(fechadi) >= CVDate("" & mytabley.Fields("fechaid")) And CVDate(fechadi) <= CVDate("" & mytabley.Fields("fechafd")) Then
         deslipox = Val("" & mytabley.Fields("dscto"))
      End If
   End If
   End If
   'si son cantidades que sucede y esta en el rango verificar si tiene grabado precio
   'If "" & DBGrid2.Columns(2) = "" & mytabley.Fields("unidad1") Then  'si es la misma unidad
   '
   '   If Val("" & DBGrid2.Columns("cantidad")) >= a And Val("" & DBGrid2.Columns("cantidad")) <= a Then
   '   End If
   'End If
   'If "" & mytablex.Fields("excludscto") = "S" Then
   '   Data1.Recordset.Fields("deslipo") = 0
   'End If
'------------------------------------------------
'MsgBox xpreciox
DBGrid2.refresh
DBGrid2.columns("zona") = "" & mytablex.Fields("seccion")
DBGrid2.columns("nroprecio") = "1"
DBGrid2.columns(52) = Format(Now, "hh:mm:ss")
DBGrid2.columns(56) = "" & mytablex.Fields("categoria")
DBGrid2.columns(0) = "" & mytablex.Fields("producto")
DBGrid2.columns(38) = "" '& mytablex.Fields("proveedor1")
DBGrid2.columns(44) = ""
DBGrid2.columns(14) = ""
DBGrid2.columns(15) = ""
DBGrid2.columns(16) = ""  '& mytablex.Fields("vendedor")
DBGrid2.columns(1) = Mid$("" & mytablex.Fields("descripcio"), 1, 60)
'MsgBox xxca
DBGrid2.columns("cantidad") = Val(xxca) 'Val(Format(Val(xxca), "0.000"))
'DBGrid2.Columns("descuento") = Val("" & mytablex.Fields("isc"))

DBGrid2.columns(2) = "" & mytabley.Fields("unidad1")  'ojo se cambio por placa
DBGrid2.columns(4) = Val("" & mytabley.Fields("factor1"))
DBGrid2.columns("precio") = xpreciox
DBGrid2.columns("total") = xpreciox
If "" & mytable11.Fields("hdetraccio") <> "S" Then
   DBGrid2.columns(54) = 0
End If
DBGrid2.columns("subtotal") = xpreciox

'DBGrid2.Columns("neto") = Val("" & mytablex.Fields("tax"))
'DBGrid2.Columns(2) = "" & mytabley.Fields("unidad1")



'dbgrid2.columns("COMISION") = Val("" & mytabley.Fields("comision"))
DBGrid2.columns(4) = Val("" & mytabley.Fields("factor1"))
DBGrid2.columns("precio") = xpreciox
DBGrid2.columns("total") = xpreciox
DBGrid2.columns("subtotal") = xpreciox

'DBGrid2.Columns("deslipo") = 0
DBGrid2.columns("deslipo") = deslipox
DBGrid2.columns(9) = 0
DBGrid2.columns("isc") = Val("" & mytablex.Fields("isc"))
DBGrid2.columns("impuesto") = 0
DBGrid2.columns("igv") = Val("" & mytablex.Fields("igv"))
DBGrid2.columns(17) = "" & mytablex.Fields("linea")

DBGrid2.columns("descuento") = 0
DBGrid2.columns("neto") = 0

'If "" & mytablex.Fields("recetaprn") <> "S" Then
'   dbgrid2.columns("dua") = "R"
'End If

'If xsw = 1 Then   'si es el precio que eligio grifos ojos..
If xpreciox > 0 Then
   
   If "" & mytablex.Fields("fuel") = "S" And Val(xxca) > 1 And xsw <> 1 Then
      DBGrid2.columns("total") = Val(xxca)
      DBGrid2.columns("cantidad") = Val(xxca) / xpreciox
   End If
   If xsw = 1 Then
      DBGrid2.columns("total") = Val(xxca)
      DBGrid2.columns("cantidad") = Val(xxca) / xpreciox
   End If
   
End If


carga_grafico "" & mytablex.Fields("producto")
carga_minimo "" & mytablex.Fields("producto")

'End If

mytables.Open "SELECT * FROM DUENO where  local='" & "" & mytable11.Fields("local") & "' and producto='" & "" & mytablex.Fields("producto") & "' ", cn, adOpenKeyset, adLockOptimistic
If mytables.RecordCount > 0 Then  'si existe
   DBGrid2.columns(48) = Trim("" & mytables.Fields("codigo"))  'ojo si no es por local
End If
mytables.Close

'---------pone a quien pertenece --------------------
'DBGrid2.Columns(34) = "" & mytablex.Fields("c11")
'DBGrid2.Columns(35) = "" & mytablex.Fields("c12")
'DBGrid2.Columns(36) = "" & mytablex.Fields("c13")
'DBGrid2.Columns(37) = "" & mytablex.Fields("c14")
'-----------------------------
'le pone las familias+subfamil+seccion+marca
DBGrid2.columns(45) = "" & mytablex.Fields("Familia")
DBGrid2.columns(46) = "" & mytablex.Fields("subFamilia")
DBGrid2.columns(47) = "" & mytablex.Fields("marca")
DBGrid2.columns("total") = Val(DBGrid2.columns("cantidad")) * Val(DBGrid2.columns("precio"))
DBGrid2.columns("ivap") = Val("" & mytablex.Fields("ivap"))
calcula_igv 0
found = leer_visorcaja("" & DBGrid2.columns("descripcio"), "S/." & DBGrid2.columns("Total"))

End Sub
Function sumar_detalle()
On Error GoTo cmd35_err
Dim found As Integer
Dim sdx As Double
Dim fila As Integer
Dim xtotal As Double
Dim xdescuento As Double
Dim xneto As Double
Dim ximpuesto As Double
Dim xsubtotal As Double
Dim xgravado As Double
Dim xc1 As Double
Dim xc2 As Double
Dim xc3 As Double
Dim xc4 As Double
Dim xc5 As Double
Dim xc6 As Double
Dim xc7 As Double
Dim xc8 As Double
Dim xc9 As Double

Dim difre As Double
Dim sw As Integer
Dim xredo As Double
Dim sdx1 As Double
'Dim xacuenta As Double
Dim vr
Dim stx As Double
Dim xntcant As Double
Dim xfilax As Integer
Dim xivap As Double
Dim xisc As Double
Dim xdetra As Double
Dim xpeaje As Double
xpeaje = 0
xdetra = 0
xntcant = 0
xredo = 0
sdx1 = 0
xc1 = 0
xc2 = 0
xc3 = 0
xc4 = 0
xc5 = 0
xc6 = 0
xc7 = 0
xc8 = 0
xc9 = 0
xivap = 0
xisc = 0


xredo = 0
xgravado = 0
xtotal = 0
xdescuento = 0
xneto = 0
ximpuesto = 0
xsubtotal = 0
'------------------------
'dbrecords = Data2.Recordset.RecordCount
'For fila = 0 To dbgrid2.visiblerows - 1
sw = 1
exisdev = 0
found = ir_primero1()
If found = 0 Then
   GoTo avex
End If
'Data2.Refresh
'Data2.Enabled = False
Do
If Data2.Recordset.EOF Then Exit Do
'xfilax = DBGrid2.Row
'If Len("" & Data2.Recordset.Fields("placa")) = 0 Then
   'MsgBox "Ingrese una Placa Valida ", 24, "AVISO"
'   DBGrid2.Col = 2
'   DBGrid2.SetFocus
'   Exit Function
'End If
If Val("" & Data2.Recordset.Fields("cantidad")) < 0 Then
exisdev = -10
End If
Data2.Recordset.Edit
suma_linea
Data2.Recordset.Update
If Val("" & Data2.Recordset.Fields("igv")) = 0 Then
   xgravado = xgravado + Val("" & Data2.Recordset.Fields("total"))
End If
xpeaje = xpeaje + Val("" & Data2.Recordset.Fields("xneto"))
xdetra = xdetra + Val("" & Data2.Recordset.Fields("tdetra"))
xisc = xisc + Val("" & Data2.Recordset.Fields("tisc"))
xivap = xivap + Val("" & Data2.Recordset.Fields("tivap"))
xntcant = xntcant + Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & Data2.Recordset.Fields("factor")) 'suma bruto
xtotal = xtotal + Val("" & Data2.Recordset.Fields("total"))
xdescuento = xdescuento + Val("" & Data2.Recordset.Fields("descuento"))
xneto = xneto + Val("" & Data2.Recordset.Fields("neto"))
ximpuesto = ximpuesto + Val("" & Data2.Recordset.Fields("impuesto"))
xsubtotal = xsubtotal + Val("" & Data2.Recordset.Fields("subtotal"))
Data2.Recordset.MoveNext
Loop
avex:
'Data2.Enabled = True
tpeaje = Format(xpeaje, nrodecimal)
tdetra = Format(xdetra, nrodecimal)
gravado = Format(xgravado, nrodecimal)
ntcant = Format(xntcant, nrodecimal)
txtotal = Format(xtotal, nrodecimal)
txtotlare = 0

If "" & mytable11.Fields("redondeo") = "S" Then
'MsgBox "abc"
   txtotlare = Val(redondeo1(txtotal)) - Val(txtotal)
   txtotal = redondeo1(txtotal)
   'MsgBox txtotal
End If
tisc = Val(Format(xisc, nrodecimal))
tivap = Val(Format(xivap, nrodecimal))
stx = Val(txtotal) - Val(acuenta)
rtxtotal = Format(stx, nrodecimal)
'txtotal = Format(xtotal, nrodecimal)
txdescuento = Format(xdescuento, nrodecimal)
txneto = Format(xneto, nrodecimal)
tximpuesto = Format(ximpuesto, nrodecimal)
txsubtotal = Format(xsubtotal, nrodecimal)
'calculando en dolares
If Val(paridad) = 0 Then
   paridad = "1"
End If
sdx = Val(txtotal) / Val(paridad)
txtotald = Format(sdx, nrodecimal)

sdx = Val(rtxtotal) / Val(paridad)
rtxtotald = Format(sdx, nrodecimal)

c1 = Format(xc1, nrodecimal)
c2 = Format(xc2, nrodecimal)
c3 = Format(xc3, nrodecimal)
c4 = Format(xc4, nrodecimal)
c5 = Format(xc5, nrodecimal)
c6 = Format(xc6, nrodecimal)
c7 = Format(xc7, nrodecimal)
c8 = Format(xc8, nrodecimal)
c9 = Format(xc9, nrodecimal)
'ahora con el
sumar_detalle = sw
Exit Function
cmd35_err:
MsgBox "Error en sumar_detalle " & error$, 24, "Aviso"
Exit Function
End Function
Function ir_primero1()
On Error GoTo cmd771222_err
'Data2.Recordset.MoveFirst
Data2.refresh
ir_primero1 = 1
Exit Function
cmd771222_err:
'MsgBox "aviso en ir Primero " + error$, 48, "Aviso"
Exit Function
End Function

Private Sub telefono_KeyUp(KeyCode As Integer, Shift As Integer)
Dim found As Integer
If KeyCode = &H70 Then  'f1
   found = consulta_cliente("")
End If

End Sub

Private Sub Timer1_Timer()
fechasis = Format(Now, "dd/mm/yyyy")
horasis = Format(Now, "HH:MM:SS")
End Sub


Private Sub tmrcomm_Timer()

End Sub

Private Sub txtotal_Click()
Dim found As Integer
found = sumar_detalle()
End Sub
Sub borrar_todo()
On Error GoTo cmd356_err
'If MsgBox("Desea Borrar Todo", 1, "Aviso") <> 1 Then Exit Sub
ir_primero
Do
If Data2.Recordset.EOF Then Exit Do
Data2.Recordset.Delete
Data2.refresh
Loop
inicialIzatodo
DBGrid2.SetFocus
Exit Sub
cmd356_err:
MsgBox "Aviso en borrar_todo " + error$, 48, "Aviso"
Exit Sub
End Sub
Sub inicialIzatodo()
Dim found As Integer
Dim sdx As Double
found = leer_visorcaja("SISTEMA ORION", "CASH REGISTER")
sw_acura = 0
limite_credito = ""
felizc = ""
stkminimo = ""
octipo = ""
ocserie = ""
ocnumero = ""
clasificacion = ""

cerrar_puertosmscomm

fotoimagen = LoadPicture()
totpedido = ""
Label59.Caption = "NORMAL"
cuenta_separa = ""
salon = ""
mesa = ""
mesero = ""
comanda = ""
cuenta_separa = ""
consulta_comanda "" & mytable11.Fields("salon")

DBGrid2.Enabled = True
Command13.Enabled = True
ndetraccion = ""
flage = ""
sentido.Enabled = False
If "" & mytable11.Fields("sentido") = "C" Then
   sentido = ""
   sentido.Enabled = True
   Else
   sentido = "" & mytable11.Fields("sentido")
End If
hknumero = ""
tpeaje = ""
tdetra = ""
rrlocal11 = ""
rrtipo = ""
rrserie = ""
rrnumero = ""
trdescuento = ""
saldo = ""
tcampo6 = ""
crucefa.Clear
saldoabo = ""
valordescuento = 0
tipodescuento = ""
tivap = 0
tisc = 0
local1 = ""
acuenta = ""
petipo = ""
peserie = ""
penumero = ""
txtotald = nrodecimal
txtotal = nrodecimal
rtxtotald = ""
rtxtotal = ""
cprotipo = ""
cproven = ""
cprocod = ""
pedido = ""
protipo = ""
proserie = ""
pronumero = ""
local1.Visible = False
c1 = ""
c2 = ""
c3 = ""
c4 = ""
c5 = ""
c6 = ""
c7 = ""
c8 = ""
c9 = ""

tcampo1 = ""
tcampo2 = ""
tcampo3 = ""
tcampo4 = ""
tcampo5 = ""
tcampo6 = ""
xtipo = ""
xnumero = ""
xserie = ""
xvendedor = ""
xruc = ""
xnombre = ""
xdireccion = ""
xdistrito = ""
nvendedorx = ""
ntipox = ""
gravado = ""
'dotipo = ""
'   doserie = ""
'   donumero = ""
'   dototal = ""
'   dofpago = ""
'   dofecha = ""
clasificacion = ""
xestado = ""
'monto = ""
xruc = ""
dcodigo = ""
telefono = ""
dnombre = ""
ddireccion = ""
referencia = ""
fechanac = ""
xnumero = ""
codigo = ""
nombre = ""
tiposervicio1 = "Autoservicio"
flag_servicio = "A"
'tiposervicio = "Autoservicio"
borrar_campos
'sdx = Val("" & mytable11.Fields("numero")) + 1
'xnumero = "" & sdx
ntcant = ""
txtotlare = 0

txtotal = nrodecimal
txdescuento = ""
txneto = ""
tximpuesto = ""
txsubtotal = ""
txtotald = nrodecimal
'txtotals = nrodecimal
'CAMPO1 = ""
'CAMPO2 = ""
'campo3 = ""


cargar_tmcombina
sql_detalle
found = sumar_detalle()
'uvueltos = "S/.:" & Format(Val("" & mytable11.Fields("uvueltos")), nrodecimal)
'uvueltod = "US$:" & Format(Val("" & mytable11.Fields("uvueltod")), nrodecimal)

'uvueltos = "" & mytable11.Fields("uvueltos")
'uvueltod = "" & mytable11.Fields("uvueltod")
DBGrid2.Enabled = True
DBGrid2.SetFocus

End Sub
Function cargar_tmcombina()
borratmpcombina
cn.Execute ("select * into _c" & gusuario & " from combina ")
cn.Execute ("delete from  " & "_c" & gusuario)
End Function
Sub borratmpcombina()
On Error GoTo cmdn78_err
cn.Execute ("drop table _c" & gusuario)
Exit Sub
cmdn78_err:
Exit Sub
End Sub

Sub borrar_campos()
On Error GoTo cmd212_err
inicio1:
Data2.Recordset.MoveFirst
Data2.Recordset.Delete
GoTo inicio1
Exit Sub
cmd212_err:
Exit Sub

End Sub

Sub proceso_impresion11(bxtipo As String, bxserie As String, bxnumero As String, sw As Integer, ascopia As String)
Dim found As Integer
Dim archivot As String
On Error GoTo cmd6_err:
    cerrar_archivo
    If sw = 0 Then   'si es posible
       found = abre_puerto("" & mytable11.Fields("capuerto"), Val("" & mytable11.Fields("catipo")), "" & mytable11.Fields("gavetacola"))
    End If
    
    'verificamos si es puerto LPT para no hacer formato impresion
    found = control_impresion(bxtipo, bxserie, bxnumero, 10)
    If found = 10 Then
       Exit Sub
    End If
    
    factura_formatox "" & mytable11.Fields("local"), "" & bxtipo, "" & bxserie, "" & bxnumero, ascopia, sw
    
    cerrar_archivo
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = control_impresion(bxtipo, bxserie, bxnumero, sw)
                  

    
    
    'genver.File = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    Exit Sub
cmd6_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$, 48, "Aviso"
    Exit Sub
End Sub
Function formateab(buf As String, longitud As Integer, sw As Integer, sw1 As Integer) As String
Dim xbuf As String
Dim buf1 As String
Dim sdx As Integer
On Error GoTo cmd203_err
'Open filename For Append As #1
buf1 = buf
sdx = longitud - Len(buf)
If sdx > 0 Then
   If sw1 = 0 Then
      buf1 = buf & Space$(sdx)
   End If
   If sw1 = 1 Then
      buf1 = Space$(sdx) & buf
   End If
End If
formateab = Mid$(buf1, 1, longitud)
Exit Function
cmd203_err:
MsgBox "Mensaje, Error en formateab " & error$
Exit Function

End Function

Function imprime_adifac(batipo As String, baserie As String, banumero As String, sw As Integer, xxpuerto As String)
Dim mytablex As New ADODB.Recordset
Dim found As Integer
Dim buf As String
Dim X As Double
Dim sFile As String
Dim cfilename As String
On Error GoTo cmd67112_err
Dim xmcanal
Exit Function
'---------------------------------
mytablex.Open "SELECT * FROM " & godetalle & " where  local='" & "" & mytable11.Fields("local") & "' and tipo='" & batipo & "' and serie='" & baserie & "' and numero='" & banumero & "'", cn, adOpenKeyset, adLockOptimistic
    If mytablex.RecordCount = 0 Then  'si existe
       mytablex.Close
       Exit Function
    End If

xmcanal = FreeFile
X = 0
Open globaldir & "\temporal\" & gusuario & "TX" For Output As #xmcanal
   Print #xmcanal, "      DOCUMENTO (" + batipo + " " + banumero & ")"
   Print #xmcanal, "-------------------------------"
   Print #xmcanal, "NOMBREPRODUCTO       CANTIDAD "
   
      Do
      If mytablex.EOF Then Exit Do
         buf = formateab(Mid$("" & mytablex.Fields("descripcio"), 1, 25), 25, 0, 0)
         buf = buf & formateab(Mid$("" & mytablex.Fields("cantidad"), 1, 25), 7, 2, 0)
         X = X + Val("" & mytablex.Fields("cantidad"))
         Print #xmcanal, buf
         mytablex.MoveNext
      Loop
   
mytablex.Close
Print #xmcanal, "-------------------------------"
Print #xmcanal, "Unidades       :" + Format(X, "000")
Close #xmcanal
sFile = globaldir & "\temporal\" & gusuario & "tx"
If sw = 0 Then  'cola
   found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"))
End If
If sw = 1 Then  'impresion directa
   FileName = sFile
   found = star_sp342(xxpuerto, 0)
   found = corte_papel(xxpuerto, 0)
End If
Exit Function
cmd67112_err:
MsgBox "Aviso en imprime adicional " + error$, 48, "Aviso"
Close #xmcanal
Exit Function

End Function

Function orden_despacho_n(bxlocal As String, bxtipo As String, bxserie As String, bxnumero As String, buf1 As String)
Dim xdato As String
Dim buf As String
Dim bufx As String
Dim Puerto As String
Dim puertos As String
Dim puertod As String
Dim found As Integer
Dim mytablef As New ADODB.Recordset
Dim i As Integer
Dim j As Integer
Dim X As Integer
Dim xbuf0 As String
Dim xbuf1 As String
Dim xbuf2 As String
Dim sw As Integer
Dim cola As String
Dim oldprinter
'Dim mydbf As Database
On Error GoTo cmd78901_err
'impresora por default atachado
'If MsgBox("Desea Imprimir Orden Despacho ", 1, "Aviso") <> 1 Then Exit Function
List1.Clear
suma1 = 0
Puerto = ""
puertod = ""
puertos = "OD" '& mytable11.Fields("odpuerto")
Puerto = puertos
cerrar_archivo
    'MsgBox godetalle
    FileName = caja & Puerto
    found = borra_nombre(FileName)
    'ojo es la orden de despacho
    'MsgBox "...Presione enter para continuar la orden Despacho " & bxtipo & " " & bxnumero, 48, "Aviso"
    
    
    mytablef.Open "SELECT * FROM " & godetalle & " where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenKeyset, adLockOptimistic
    If mytablef.RecordCount = 0 Then  'si existe
       mytablef.Close
       Exit Function
    End If
   
'-----------ORDEN DE DESPACHO---------------------------------------------
List1.Clear
'---OJO VERIFICAR ESTO------------------
ncanal = FreeFile
Do
   If mytablef.EOF Then Exit Do
   If "" & mytablef.Fields("dua") = "R" Then GoTo ksiguiente111
   'MsgBox "" & mytablef.Fields("producto")
   If Len("" & mytablef.Fields("producto")) > 0 And (Val("" & mytablef.Fields("cantidad")) > 0 Or Val("" & mytablef.Fields("cantidad")) < 0) Then
      found = busca_familia_orden("" & mytablef.Fields("producto"), Puerto, puertod, cola)
      'MsgBox puerto & " " & puertod & " " & cola
      If found = 0 Then   'si no existe debe tomar el defaul de la impresora
          Puerto = puertos
      End If
      'MsgBox puerto
      If Len(Puerto) = 0 Then
         Puerto = "LPT"
      End If
      'MsgBox found
   '--------------------------------------
      sw = 0
      FileName = Trim(caja & Puerto)
      found = existearchivo("" & FileName)
      If found = 1 Then  'verificar si no existe en la lista
         sw = 0
         For i = 0 To List1.ListCount - 1
          j = InStr(List1.List(i), "|")
          xbuf0 = Mid$(List1.List(i), 1, j - 1)
          If xbuf0 = FileName Then
             sw = 1
          End If
         Next i
         If sw = 0 Then  'no existe en la lista
            found = borra_nombre(FileName)
            found = 0
         End If
      End If
      cerrar_archivo
      Open FileName For Append As #ncanal
      'MsgBox found
      If found = 0 Then
         List1.AddItem FileName & "|" & puertod & "|" & cola & "|" 'adiciona en la lista
         cabecera_orden_despacho "" & mytablef.Fields("vendedor"), buf1, bxnumero, "" & xnombre
      End If
      imprime_detalle_orden1 mytablef
      Close #ncanal
   End If
ksiguiente111:
   mytablef.MoveNext
Loop
cerrar_archivo


'-------------se adiciono para agilidad--------------------------------
For i = 0 To List1.ListCount - 1
   
   'j = InStr(list1.List(i), "|")
    xdato = List1.List(i)
    'MsgBox xdato
   extrae_puertos xdato, xbuf0, xbuf1, xbuf2
   'xbuf0 = Mid$(list1.List(i), 1, j - 1)
   'xbuf1 = Mid$(list1.List(i), j + 1, Len(list1.List(i)) - (j))
   'xbuf2 = Mid$(list1.List(i), j + 1, Len(list1.List(i)) - (j))
   FileName = xbuf0
   If existearchivo(xbuf0) = 1 Then
      ncanal = FreeFile
      Open FileName For Append As #ncanal
      For X = 1 To 5
          Print #ncanal, ""
      Next X
      Print #ncanal, ""
      Close #ncanal
   End If
Next i
'MsgBox List1.ListCount
For i = 0 To List1.ListCount - 1
   xdato = List1.List(i)
   'MsgBox xdato
   extrae_puertos xdato, xbuf0, xbuf1, xbuf2
   FileName = xbuf0
   'MsgBox xdato & " " & xbuf0 & " " & xbuf1 & " " & xbuf2
   If existearchivo(xbuf0) = 1 Then
      If xbuf2 = "S" Then
         'MsgBox xbuf1
         oldprinter = Printer.DeviceName
         selecciona_impresoras (Trim(xbuf1))
         found = Imprime_archivojj(xbuf0, 0, "" & mytable11.Fields("tamanorden"), "" & mytable11.Fields("nombrefont"))
         selecciona_impresoras (Trim(oldprinter))
         'found = orden_oprn(xbuf1, "" & mytable11.Fields("tipoleta"), "" & mytable11.Fields("tamano"), "" & mytable11.Fields("negrita"))
      Else
      
      ncanal = FreeFile
      Open FileName For Append As #ncanal
      For X = 1 To 2
          Print #ncanal, ""
      Next X
      Print #ncanal, ""
      Close #ncanal
      
         found = star_sp342(Trim(xbuf1), 0)
         found = corte_papel(Trim(xbuf1), 1)
         
         'found = star_sp342(xxpuerto, 0)
         'found = star_sp342(xbuf1, ticketera_cajon)
      End If
      cerrar_archivo
      found = borra_nombre(xbuf0)
   End If
Next i
cerrar_archivo

Exit Function
mytablef.Close
Exit Function
cmd78901_err:
   MsgBox "MENSAJE, ERROR EN ORDEN DESPACHO " & error$, 24, "AVISO"
   Exit Function
End Function

Function imprime_adicional(batipo As String, baserie As String, banumero As String, sw As Integer, xxpuerto As String)
Dim mytablex As New ADODB.Recordset
Dim ax1cambio As String
Dim ax1telefono As String
Dim ax1nombre As String
Dim ax1direccio As String
Dim ax1referencia As String
Dim ax1pago As String
Dim ax1total As String
Dim ax1vuelto As String
Dim found As Integer
Dim cfilename As String
Dim sFile As String
Dim i As Integer
Dim buf As String
On Error GoTo cmd6711_err
Dim xmcanal
ax1cambio = ""
ax1telefono = ""
ax1nombre = ""
ax1direccio = ""
ax1referencia = ""
ax1pago = ""
ax1total = ""
ax1vuelto = ""
ax1cambio = "2.78"

   'MsgBox codigo
   '---------------------------------
   mytablex.Open "SELECT * FROM deliveri where  codigo='" & codigo & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount = 0 Then  'si existe
      mytablex.Close
      Exit Function
   End If
      ax1telefono = "" & mytablex.Fields("telefono")
      ax1nombre = "" & mytablex.Fields("nombre")
      
      ax1direccio = "" & mytablex.Fields("direccion")
      ax1referencia = "" & mytablex.Fields("referencia")
   
   mytablex.Close

xmcanal = FreeFile
Open globaldir & "\temporal\" & gusuario & "TX" For Output As #xmcanal
Print #xmcanal, "DELIVERY"
Print #xmcanal, "Telef:" + ax1telefono
Print #xmcanal, "Clien:" + ax1nombre
Print #xmcanal, "Direc:" + ax1direccio
Print #xmcanal, "Refer:" + ax1referencia
Print #xmcanal, "T/C  :" + ax1cambio


mytablex.Open "SELECT * FROM " & gofpago & " where  local='" & "" & mytable11.Fields("local") & "' and tipo='" & batipo & "' and serie='" & baserie & "' and numero='" & banumero & "'", cn, adOpenKeyset, adLockOptimistic
   If mytablex.RecordCount = 0 Then  'si existe
      mytablex.Close
      Exit Function
   End If

      Do
      If mytablex.EOF Then Exit Do
         Print #xmcanal, "pago :" + mytablex.Fields("descripcio")
         Print #xmcanal, "Total:" + Format(Val("" & mytablex.Fields("recibe")), "0.00")
         Print #xmcanal, "Vuelt:" + Format(Val("" & mytablex.Fields("saldos")), "0.00")
      mytablex.MoveNext
      Loop
   
mytablex.Close

'------------ PRODUCTOS
 mytablex.Open "select * from " & godetalle & " where local='" & "" & mytable11.Fields("local") & "' and tipo='" & "" & batipo & "' and serie='" & "" & baserie & "' and numero='" & "" & banumero & "'", cn, adOpenStatic, adLockOptimistic
       If mytablex.RecordCount > 0 Then
          Do
          If mytablex.EOF Then Exit Do
          buf = "" & mytablex.Fields("cantidad")
          found = formateaa(buf, 6, 0, 0)
          found = formateaa("", 1, 0, 0)
          buf = "" & mytablex.Fields("descripcio")
          found = formateaa(buf, 22, 2, 0)
          
          '----------------------
           If Len("" & mytablex.Fields("observa1")) > 0 Then
         buf = "*" & mytablex.Fields("observa1")
         found = formateaa(buf, 28, 2, 0)
  
    End If
    If Len("" & mytablex.Fields("observa2")) > 0 Then
       buf = "*" & mytablex.Fields("observa2")
       found = formateaa(buf, 28, 2, 0)
    End If
    If Len("" & mytablex.Fields("observa3")) > 0 Then
       buf = "*" & mytablex.Fields("observa3")
       found = formateaa(buf, 28, 2, 0)
    End If
    
          '----------------------
          
          mytablex.MoveNext
          Loop
       End If
       mytablex.Close
       


For i = 1 To 7
   Print #xmcanal, ""
Next i
Close #xmcanal
sFile = globaldir & "\temporal\" & gusuario & "TX"
If sw = 0 Then
   found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"))
End If
If sw = 1 Then
   FileName = sFile
   found = star_sp342(xxpuerto, 0)
   found = corte_papel(xxpuerto, Val("" & mytable11.Fields("catipo")))
End If

Exit Function
cmd6711_err:
MsgBox "Aviso en imprime adicional " + error$, 48, "Aviso"
Close #xmcanal
Exit Function
End Function




Function control_impresion(bxtipo As String, bxserie As String, bxnumero As String, psw As Integer)
Dim found As Integer
Dim sFile As String
Dim mytablex As New ADODB.Recordset
Dim sw As String
Dim xcolax As String
Dim xxpuerto As String
Dim oldprinter As String
On Error GoTo cmd67111_err
sw = ""
xcolax = ""
xxpuerto = "X_"
       If mytablex.State = 1 Then mytablex.Close
       mytablex.Open "SELECT * FROM tipo where  tipo='" & bxtipo & "'", cn, adOpenDynamic, adLockOptimistic
       If mytablex.RecordCount > 0 Then 'si existe
       Select Case "" & mytablex.Fields("tipodoc")
       Case "A"
       xxpuerto = "" & mytable11.Fields("puertobm")
       sw = "" & mytable11.Fields("ibm")
       xcolax = "" & mytable11.Fields("cbm")
       Case "B"
       xxpuerto = "" & mytable11.Fields("puertofm")
       sw = "" & mytable11.Fields("ifm")
       xcolax = "" & mytable11.Fields("cfm")
       Case "C"
       xxpuerto = "" & mytable11.Fields("puertotb")
       sw = "" & mytable11.Fields("itb")
       xcolax = "" & mytable11.Fields("ctb")
       Case "D"
       xxpuerto = "" & mytable11.Fields("puertotf")
       sw = "" & mytable11.Fields("itf")
       xcolax = "" & mytable11.Fields("ctf")
       Case "G"
       xxpuerto = "" & mytable11.Fields("puertonv")
       sw = "" & mytable11.Fields("inv")
       xcolax = "" & mytable11.Fields("cnv")
       Case "H"
       xxpuerto = "" & mytable11.Fields("puertope")
       sw = "" & mytable11.Fields("ipe")
       xcolax = "" & mytable11.Fields("cpe")
       Case "I"
       xxpuerto = "" & mytable11.Fields("puertoot")
       sw = "" & mytable11.Fields("iot")
       xcolax = "" & mytable11.Fields("cot")
       Case "1"
       xxpuerto = "" & mytable11.Fields("puertoexo")
       sw = "" & mytable11.Fields("iexo")
       xcolax = "" & mytable11.Fields("cexo")
       End Select
     End If
mytablex.Close
If psw = 10 Then  'solo es para ver si es LPT
   control_impresion = 11
   If xxpuerto = "LPT" Then
    control_impresion = 10
   End If
   Exit Function
End If
 
'found = abre_puerto("" & mytable11.Fields("capuerto"), Val("" & mytable11.Fields("catipo")))
'ahora validamos los parametros de impresion
If psw = 2 Then  'si  es orden de despacho
If xcolax = "S" Then
   'oldprinter = Printer.DeviceName
   'Set Printer = Printers(xxpuerto)
   sFile = globaldir & "\temporal\" & gusuario & ".txt"
   found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"))
   '----------------------------------
   
   'Set Printer = Printers(oldprinter)
End If
If xcolax <> "S" Then
   found = star_sp342(xxpuerto, 0)
   found = corte_papel(xxpuerto, Val("" & mytable11.Fields("catipo")))
End If
control_impresion = found
Exit Function
End If

If sw = "S" Then
   If MsgBox("Desea Imprimir", 1 + 256, "Aviso") <> 1 Then
      control_impresion = 1
      Exit Function
   End If
End If

If xcolax = "S" Then
   oldprinter = Printer.DeviceName
   selecciona_impresoras (xxpuerto)
   sFile = globaldir & "\temporal\" & gusuario & ".txt"
   found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"))
                  If bxtipo = "2" Then
                     found = imprime_adifac(bxtipo, bxserie, bxnumero, 0, "")
                  End If
                  
                  If flag_servicio = "D" Then
                     found = imprime_adicional(bxtipo, bxserie, bxnumero, 0, "")
                  End If
   
selecciona_impresoras (oldprinter)
End If
If xcolax <> "S" Then
   found = star_sp342(xxpuerto, 0)
   found = corte_papel(xxpuerto, Val("" & mytable11.Fields("catipo")))
                  If bxtipo = "2" Then
                     found = imprime_adifac(bxtipo, bxserie, bxnumero, 1, xxpuerto)
                  End If
                  
                  If flag_servicio = "D" Then
                     found = imprime_adicional(bxtipo, bxserie, bxnumero, 1, xxpuerto)
                  End If
   
   
End If
control_impresion = found
Exit Function
cmd67111_err:
Exit Function
End Function
Sub proceso_impresioncopia()
Dim found As Integer
Dim archivot As String
On Error GoTo cmd7_err:
    cerrar_archivo
    factura_formatox "" & mytable11.Fields("local"), "" & rcconsulta.Fields("tipo"), "" & rcconsulta.Fields("serie"), "" & rcconsulta.Fields("numero"), "1", 0
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    Exit Sub
cmd7_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$
    Exit Sub
End Sub
Sub proceso_impresioncopia1()
Dim found As Integer
Dim archivot As String
On Error GoTo cmd17_err:
    cerrar_archivo
    factura_formatox rrlocal11, rrtipo, rrserie, rrnumero, "1", 0
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    rrlocal11 = ""
    rrtipo = ""
    rrserie = ""
    rrnumero = ""
    MsgBox "Proceso Realizado con exito", 48, "Aviso"
    Exit Sub
cmd17_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$
    Exit Sub

End Sub

Sub factura_formatox(bxlocal As String, bxtipo As String, bxserie As String, bxnumero As String, ascopia As String, psw As Integer)
Dim vacu As String
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
Dim mytablez As New ADODB.Recordset
Dim found As Integer
Dim nro_lineas As Integer
Dim contando As Integer
Dim faltante As Integer
Dim i As Integer
Dim archivo_formato As String
On Error GoTo cmd450009_err
       vacu = ""
       'MsgBox "QU"
       nro_lineas = busca_tipo_lineas(bxtipo)
       'MsgBox ""
       'If nro_lineas <= 0 Then
       '   nro_lineas = 10
       'End If
       'MsgBox ""
       contando = 0
       FileName = globaldir & "\temporal\" & gusuario & ".txt"
       found = borra_nombre("" & FileName)
       
       If psw = 2 Then 'si es de orden
          archivo_formato = "orden"
          Else
          archivo_formato = busca_archivo_formato(bxtipo)
          If Len(archivo_formato) = 0 Then
             MsgBox "No existe archivo formato ", 48, "Aviso"
             Exit Sub
          End If
       End If
       'cabeza
       'proceso_formatos(archivo_formato , mydbx , mytablex , ubicacioni , ubicacionf , basedatos , indice , tipo , numero , ascopia , contando )
       'MsgBox gocabeza
       mytablex.Open "SELECT * FROM " & gocabeza & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic
       If mytablex.RecordCount = 0 Then 'si existe
          mytablex.Close
          Exit Sub
       End If
       'MsgBox ""
       found = proceso_formatos(archivo_formato, mytablex, "{", "}", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
       vacu = "" & mytablex.Fields("acu")
       'MsgBox ""
       '
       'detalle
       flag_contando = 0
       If "" & mytablex.Fields("observa") = "CONSUMO" Then
                  Open FileName For Append As #1
                  found = formateaa("1  POR CONSUMO            " & Format(Val("" & mytablex.Fields("total")), "0.00"), 30, 2, 0)
                  'found = formateaa("1    POR CONSUMO            ", 30, 2, 0)
                  ' found = formateaa("1    COMBUSTIBLE            ", 30, 2, 0)
                  contando = contando + 1
                  flag_contando = contando + 1
                  Close #1
       End If
       If "" & mytablex.Fields("observa") <> "CONSUMO" Then
       mytabley.Open "SELECT * FROM " & godetalle & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic
       If mytabley.RecordCount > 0 Then 'si existe
          Do
          If mytabley.EOF Then Exit Do
             If "" & mytabley.Fields("dua") <> "R" Then
                flag_contando = contando + 1
                found = proceso_formatos(archivo_formato, mytabley, "/", "\", godetalle, "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
                contando = contando + 1
             End If
          mytabley.MoveNext
          Loop
        End If
        mytabley.Close
        End If
       
        
        
        '
        If nro_lineas > 0 Then
        'If bxtipo <> "1" And bxtipo <> "2" And bxtipo <> "5" And bxtipo <> "7" Then
           If contando < nro_lineas Then
              For i = contando To nro_lineas
                  Open FileName For Append As #1
                  found = formateaa("", 1, 2, 0)
                  Close #1
              Next i
           End If
        End If
       '----- SUBTOTAL
       
       
       found = proceso_formatos(archivo_formato, mytablex, "$", "?", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
       
               
       mytablez.Open "SELECT * FROM " & gofpago & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic
       If mytablez.RecordCount > 0 Then 'si existe
           Do
           If mytablez.EOF Then Exit Do
              found = proceso_formatos(archivo_formato, mytablez, "<", ">", gofpago, "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
           mytablez.MoveNext
        Loop
        End If
        
       
       found = proceso_formatos(archivo_formato, mytablex, "^", "&", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
       
       mytablex.Close
       'mytabley.Close
       mytablez.Close
        
       Exit Sub
cmd450009_err:
       MsgBox "Aviso Proceso en Compilando " & error$, 24, "Aviso"
       'mytablex.Close
       '
       Exit Sub

End Sub
Function busca_archivo_formato(bxtipo As String) As String
Dim mytablex As New ADODB.Recordset

If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM tipo where  tipo='" & bxtipo & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
       Select Case "" & mytablex.Fields("tipodoc")
       Case "Z" 'si es traslado
       busca_archivo_formato = "" & mytablex.Fields("archivo")
       Case "A"
       busca_archivo_formato = "" & mytable11.Fields("archivobm")
       Case "B"
       busca_archivo_formato = "" & mytable11.Fields("archivofm")
       Case "C"
       busca_archivo_formato = "" & mytable11.Fields("archivotb")
       Case "1"
       busca_archivo_formato = "" & mytable11.Fields("archivoexo")
       Case "D"
       busca_archivo_formato = "" & mytable11.Fields("archivotf")
       Case "G"
       busca_archivo_formato = "" & mytable11.Fields("archivonv")
       Case "H"  'cotizacion
       busca_archivo_formato = "" & mytable11.Fields("archivope")
       Case "I"  'pedido
       'busca_archivo_formato = "" & mytable11.Fields("archivoot")
       busca_archivo_formato = "" & mytable11.Fields("archivope")
       Case "1"
       busca_archivo_formato = "" & mytable11.Fields("archivonv")
       End Select
End If
mytablex.Close
 
End Function
Function busca_parame1(buf As String, sw As Integer) As String
Dim sdx As Double
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM parame where  codigo='01'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe

   If sw = 2 Then
      
   End If
   If sw = 0 Then
   sdx = Val("" & mytablex.Fields("clientes")) + 1
   busca_parame1 = "" & sdx
   End If
   If sw = 1 Then
   'mytablex.Edit
   mytablex.Fields("clientes") = buf
   mytablex.Update
   End If
End If
mytablex.Close
End Function
Sub modifica_detalle()
Dim i As Integer

Dim mytablex As New ADODB.Recordset
borrar_campos

mytablex.Open "SELECT * FROM " & dgusuariog & "   where  local='" & "" & mytable11.Fields("local") & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount = 0 Then 'si existe
   mytablex.Close
   Exit Sub
End If
   Do
   If mytablex.EOF Then Exit Do
   
      Data2.Recordset.AddNew
      For i = 0 To mytablex.Fields.count - 1
          Data2.Recordset.Fields(i) = mytablex.Fields(i)
      Next i
      Data2.Recordset.Update
   
   mytablex.MoveNext
   Loop
   mytablex.Close
    

End Sub


Sub inicializa_deliveri()
clasificacion = ""
dcodigo = ""
telefono = ""
dnombre = ""
ddireccion = ""
referencia = ""
fechanac = ""
codigo = ""
nombre = ""
felizc = ""
'dotipo = ""
'   doserie = ""
'   donumero = ""
'   dototal = ""
'   dofpago = ""
'   dofecha = ""
End Sub
Function busca_paridad()
Dim found As Integer
Dim mytablex As New ADODB.Recordset
paridad = "1"
paridadfp = "1"

   If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "SELECT * FROM parame where codigo='01' ", cn, adOpenStatic, adLockOptimistic
   If mytablex.EOF = True Or mytablex.BOF = True Then
      mytablex.Close
      Exit Function
   End If
   paridad = "" & mytablex.Fields("parivta")
   paridadfp = "" & mytablex.Fields("parivta")
   If Val(paridad) = 0 Then
      paridad = "1"
   End If
   If Val(paridadfp) = 0 Then
      paridadfp = "1"
   End If
   busca_paridad = 1
   mytablex.Close
 
End Function
Sub ir_finalx()
On Error GoTo cmd13_err
Data1.Recordset.MoveLast
Exit Sub
cmd13_err:
Exit Sub
End Sub
Sub PROCESO_BORRAR_DOCUMENTO(buf0 As String, buf As String, buf1 As String, buf2 As String)
Dim mytablex As New ADODB.Recordset
'MsgBox "dfd"
amk1:
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM cpedidov where  local='" & buf0 & "' and tipo='" & buf & "' and serie='" & buf1 & "' and numero='" & buf2 & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   mytablex.Delete
   GoTo amk1
End If

ak12:
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM dpedidov where  local='" & buf0 & "' and tipo='" & buf & "' and serie='" & buf1 & "' and numero='" & buf2 & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   mytablex.Delete
   GoTo ak12
End If
mytablex.Close
End Sub
Function busca_clientesrpt(buf As String, sw As Integer) As String

Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM clientes where  codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   busca_clientesrpt = "" & mytablex.Fields("nombre")
End If
mytablex.Close
 
End Function
Function busca_tiporpt(buf As String, sw As Integer) As String
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   busca_tiporpt = "" & mytablex.Fields("descripcio")
End If
mytablex.Close
 

End Function
Function busca_acu() As String
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "SELECT * FROM tipo where tipo='" & xtipo & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.EOF = True Or mytablex.BOF = True Then
      mytablex.Close
      Exit Function
   End If
   busca_acu = "" & mytablex.Fields("tipodoc")
   mytablex.Close

End Function
Function busca_fpagorpt(buf As String, sw As Integer) As String
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM fpago where  fpago='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   If sw = 0 Then
   busca_fpagorpt = "" & mytablex.Fields("descripcio")
   End If
   If sw = 1 Then
   busca_fpagorpt = "" & mytablex.Fields("moneda")
   End If
End If
mytablex.Close
End Function
Sub sql_ver_pedido()
'Dim buf As String
'On Error GoTo cmd37_err
'If Len(dotipo) = 0 Then Exit Sub
'If Len(doserie) = 0 Then Exit Sub
'If Len(donumero) = 0 Then Exit Sub
'buf = "select * from dpedidov where local='" & "" & mytable11.Fields("local") & "' and tipo='" & dotipo & "' and serie='" & doserie & "' and numero='" & donumero & "'"
'               Data3.Connect = "foxpro 2.5;"
'               Data3.DatabaseName = globaldir
'               Data3.RecordSource = buf
'               Data3.Refresh'
'
'Exit Sub
'cmd37_err:
'MsgBox "Error en select " & error$, 48, "Aviso"
'Exit Sub

End Sub
Sub grabar_dato_pedido(buf As String, buf1 As String, buf2 As String, buf3 As String)
On Error GoTo cmd1203_err
Dim mytablex As New ADODB.Recordset

If Len(buf) > 0 And Len(buf1) > 0 And Len(buf2) > 0 And Len(buf3) > 0 Then
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM clientes where  codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   'mytablex.Edit
   mytablex.Fields("dotipo") = buf1
   mytablex.Fields("doserie") = buf2
   mytablex.Fields("donumero") = buf3
   mytablex.Update
End If
mytablex.Close
End If
Exit Sub
cmd1203_err:
MsgBox "Aviso en grabar_dato_pedido " + error$, 48, "Aviso"
Exit Sub
End Sub
Sub sumar_al_grabar()
Dim found As Integer
On Error GoTo cmd59_err
Data2.Recordset.MoveFirst
found = sumar_detalle()
Exit Sub
cmd59_err:
Exit Sub

End Sub
Function busca_numero(bxtipo As String, bxserie As String, bxnumero As String)
Dim buf As String
Dim sdx As Double
Dim mytablex As New ADODB.Recordset
If Len(pedido) > 0 Then
   Exit Function
End If
buf = busca_tipo_acu(bxtipo)
ahj1:
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM " & gocabeza & " where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   If buf = "A" Or buf = "B" Then
      MsgBox "Numero ya Existe ", 48, "Aviso"
      busca_numero = -1
      mytablex.Close
      Exit Function
   End If
   sdx = Val(xnumero) + 1
   xnumero = "" & sdx
   bxnumero = xnumero
   GoTo ahj1
End If
mytablex.Close
End Function
Function busca_numero_pedido()
Dim sdx As Double
Dim mytablex As New ADODB.Recordset
ahj1:
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM " & gocabeza & " where  local='" & Trim("" & mytable11.Fields("local")) & " and tipo='" & xptipo & "' and serie='" & xpserie & "' and numero='" & xpnumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   sdx = Val(xpnumero) + 1
   xpnumero = "" & sdx
   GoTo ahj1
End If
mytablex.Close
 
End Function
Function proceso_cobros()
    borra_pagos
    sql_formapago
    sql_pagos
End Function
Sub sql_formapago()
Dim buf As String
buf = "select * from fpago "
If flag_servicio <> "A" And flag_servicio <> "C" And flag_servicio <> "D" Then
   buf = "select * from fpago where fpago='6'"
End If
If mytablefpago.State = 1 Then mytablefpago.Close
mytablefpago.Open buf, cn, adOpenDynamic, adLockOptimistic
Set dbgrid10.DataSource = mytablefpago
If mytablefpago.RecordCount > 0 Then
   mytablefpago.MoveFirst
End If

               'Data10.Connect = "foxpro 2.5;"
               'Data10.DatabaseName = globaldir
               'Data10.RecordSource = "select * from fpago where bco='S' or bco=NULL "
               'Data10.Refresh
               
End Sub
Sub sql_pagos()

               Data9.Connect = "foxpro 2.5;"
               Data9.DatabaseName = globaldat
               Data9.RecordSource = "select * from  " & fpusuario
               Data9.refresh

End Sub
Sub borra_pagos()
On Error GoTo cmd8912_err
    mydbxglo.Execute "DELETE FROM " & fpusuario
    Data9.refresh
    Label45.Caption = "Falta"
    stxtotals = ttxtotals
    stxtotald = ttxtotald
    Exit Sub
cmd8912_err:
   MsgBox "Error en borra_pagos " + error$, 48, "Aviso"
   Exit Sub
    
End Sub
Sub cerrar_data9()
'On Error GoTo cmd3_err
'Data9.Recordset.Close
'Exit Sub
'cmd3_err:
'Exit Sub
End Sub
Sub macro_inserta_registro()
'ultimo_fpago
Data9.Recordset.AddNew
Data9.Recordset.Fields("descripcio") = "" & dbgrid10.columns(0)
Data9.Recordset.Fields("fpago") = "" & dbgrid10.columns(1)
Data9.Recordset.Fields("moneda") = "" & dbgrid10.columns(2)
Data9.Recordset.Fields("acu") = "" & dbgrid10.columns("tipo")
Data9.Recordset.Update
'Data9.Recordset.MoveNext
Data9.refresh
End Sub
Sub ultimo_fpago()
On Error GoTo cmd780_err
Data9.Recordset.MoveLast
Exit Sub
cmd780_err:
Exit Sub

End Sub
Sub ir_ultimo_macro()
On Error GoTo cmd78_err
Data9.Recordset.MoveFirst
Exit Sub
cmd78_err:
Exit Sub
End Sub
Function macro_credito(sw As Integer)
   Frame6.Visible = True
   habilita_lab7 1
   descripcio1.Visible = True
   descripcio2.Visible = True
   descripcio3.Visible = True
   descripcio4.Visible = True
   descripcio5.Visible = True
   descripcio6.Visible = True
   tcampo1.MaxLength = 11
   tcampo2.MaxLength = 60
   tcampo3.MaxLength = 15
   tcampo4.MaxLength = 30
   tcampo5.MaxLength = 11 '3
   tcampo6.MaxLength = 2
   tcampo1 = "" & codigo
   tcampo2 = "" & nombre
   tcampo3 = ""
   tcampo4 = ""
   tcampo5 = ""
   tcampo6 = ""
   tcampo1.Visible = True
   tcampo2.Visible = True
   tcampo3.Visible = True
   tcampo4.Visible = True
   tcampo5.Visible = True
   tcampo6.Visible = True
   
   tcampo1.Enabled = True
   tcampo2.Enabled = True
   tcampo3.Enabled = True
   tcampo4.Enabled = True
   tcampo5.Enabled = True
   tcampo6.Enabled = True
   
   descripcio1 = "Codigo"
   descripcio2 = "Nombre"
   descripcio3 = "NroTarjeta"
   descripcio4 = "Observacion"
   descripcio5 = "NroDias"
   descripcio6 = ""
   If sw = 4 Then  'tarjeta credito
   descripcio1 = "Codigo"
   descripcio2 = "Nombre"
   descripcio3 = "NroTarjeta"
   descripcio4 = "Entidad"
   descripcio5 = "Propina"
   tcampo3.MaxLength = 4
   End If
   If sw = 3 Or sw = 6 Then 'credito
   descripcio1 = "Codigo"
   descripcio2 = "Nombre"
   descripcio3 = "NroAprob"
   descripcio4 = "Observacion"
   descripcio5 = "NroDias"
   End If
   If sw = 8 Then 'orden trabajo
   descripcio1 = "Codigo"
   descripcio2 = "Nombre"
   descripcio3 = "Observa"
   descripcio4 = ""
   descripcio5 = ""
   descripcio6 = ""
   'descripcio3 = "Tipo"
   'descripcio4 = "Numero"
   'descripcio5 = "NroDias"
   End If
   If sw = 5 Then  'tarjeta Debito
   descripcio1 = "Codigo"
   descripcio2 = "Nombre"
   descripcio3 = "NroTarjeta"
   descripcio4 = "Entidad"
   descripcio5 = "Propina"
   tcampo3.MaxLength = 4
   End If
   If sw = 1 Then  'SI ES PAGO ADELANTADO
   descripcio1 = "Codigo"
   descripcio2 = "Nombre"
   descripcio3 = ""
   descripcio4 = ""
   descripcio5 = ""
   descripcio6 = ""
   'tcampo1.Enabled = False
   'tcampo2.Enabled = False
   'tcampo4.Enabled = False
   'tcampo5.Enabled = False
   'tcampo6.Enabled = False
   End If
   If sw = 2 Then
     descripcio3 = "Nro.Op.Banco"
   End If
   
End Function
Sub suma_fpagov()
Dim sdxs As Double
Dim sdxd As Double
Dim sdx As Double
Dim sdx1 As Double
On Error GoTo cmd7812_err
Label45.Caption = "Falta"
sdxs = Val(ttxtotals)  'saldoa
stxtotals = Format(sdxs, nrodecimal)
'Data9.Recordset.MoveFirst
Data9.refresh
Do
If Data9.Recordset.EOF Then Exit Do
   If Len("" & Data9.Recordset.Fields("FPAGO")) > 0 Then
   Data9.Recordset.Edit
   sdx = Val("" & Data9.Recordset.Fields("recibe"))
   If "" & Data9.Recordset.Fields("moneda") = "D" Then
      sdx = sdx * Val(paridadfp) 'Val("" & Data9.Recordset.Fields("cambio"))
      sdx = Val(Format(sdx, nrodecimal))
      Data9.Recordset.Fields("cambio") = sdx
   End If
   If sdx >= sdxs Then
      sdx1 = -sdx + sdxs
      sdx1 = Val(Format(sdx1, nrodecimal))
      Data9.Recordset.Fields("total") = sdxs
      Data9.Recordset.Fields("saldos") = sdx1
      stxtotals = Format(sdx1, nrodecimal)
      sdxs = 0
      GoTo conmuta
   End If
    If sdxs > sdx Then
      sdx1 = sdxs - sdx
      sdx1 = Val(Format(sdx1, nrodecimal))
      Data9.Recordset.Fields("total") = sdx
      Data9.Recordset.Fields("saldos") = 0
      stxtotals = Format(sdx1, nrodecimal)
      sdxs = sdx1
   End If
   If "" & Data9.Recordset.Fields("acu") = "C" Then
      'codigo = tcampo1
      'nombre = tcampo2
   End If
conmuta:
   Data9.Recordset.Update
End If
Data9.Recordset.MoveNext
Loop
stxtotald = Format(0, nrodecimal)
If Val(paridadfp) > 0 Then
   sdx = Val(stxtotals) / Val(paridadfp)
   stxtotald = Format(sdx, nrodecimal)
End If
If stxtotals <= 0 Then
   Label45.Caption = "Vuelto"
   dbgrid10.Enabled = True
   dbgrid10.SetFocus
End If
Exit Sub
cmd7812_err:
MsgBox "Error en " + error$, 48, "Aviso"
Exit Sub
End Sub
Sub valida_ingresado()
Dim sdx As Double
Dim xsoles As Double
Dim xdolares As Double
Dim xfaltas As Double
Dim xfaltad As Double
Dim xvueltos As Double
Dim xvueltod As Double
Dim sdx1 As Double
Dim sdx2 As Double
Dim sdx3 As Double
xsoles = 0
xdolares = 0
xfaltas = 0
xfaltad = 0
xvueltos = 0
xvueltod = 0
sdx3 = 0
If "" & DBGrid9.columns(1) = "S" Then
   xsoles = Val("" & DBGrid9.columns(2))
   xdolares = Val(Val("" & DBGrid9.columns(2))) / Val(paridadfp)
   sdx3 = xdolares
End If
If "" & DBGrid9.columns(1) = "D" Then
   xdolares = Val("" & DBGrid9.columns(2))
   xsoles = Val("" & DBGrid9.columns(2)) * Val(paridadfp)
   sdx3 = xsoles
End If
Data9.Recordset.Edit
Data9.Recordset.Fields("cambio") = sdx3
Data9.Recordset.Fields("recibes") = xsoles
Data9.Recordset.Fields("recibed") = xdolares
'sdx1 = Val(stxtotals) - xsoles
'sdx2 = Val(stxtotald) - xdolares
'Data9.Recordset.Fields("saldos") = sdx1
'Data9.Recordset.Fields("saldod") = sdx2
Data9.Recordset.Update
'suma_fpagov
End Sub

Private Sub xcongela_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   cmdCancelar_Click
   Exit Sub
End If


End Sub

Private Sub totcoma_Click()
trgb.tipo = "PRODUCTO"
trgb.Show 1
inicia_color_producto
End Sub

Private Sub turno_Click()
tdremoto.Show 1
End Sub

Private Sub xcongelax_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   cmdCancelar_Click
   Exit Sub
End If
cmdGrabar_Click
End Sub

Private Sub xdireccion_DblClick()
tkeyboar.flag = "DIRECCION"
tkeyboar.Show 1
End Sub

Private Sub xdireccion_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
xdistrito.SetFocus
End Sub

Private Sub xdireccion_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   xnombre.SetFocus
   Exit Sub
End If
End Sub

Private Sub xdistrito_DblClick()
tkeyboar.flag = "GLOSA"
tkeyboar.Show 1
End Sub

Private Sub xdistrito_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If xserie.Enabled = False Then  'ver si es ticket
   Command13_Click
   Exit Sub
End If
xserie.SetFocus
End Sub

Private Sub xdistrito_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   xdireccion.SetFocus
   Exit Sub
End If
End Sub

Private Sub xnombre_DblClick()
tkeyboar.flag = "NOMBRE"
tkeyboar.Show 1
End Sub

Private Sub xnombre_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
xdireccion.SetFocus
End Sub

Private Sub xnombre_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   xruc.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   If local1.Visible <> True Then
      consulta_xruc
   End If
     If local1.Visible = True Then
      consulta_xruc2
   End If
End If

End Sub

Private Sub xnumero_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Command13_Click


End Sub

Private Sub xnumero_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   xserie.SetFocus
   Exit Sub
End If
End Sub


Private Sub xopciones_Click(Index As Integer)
Dim found As Integer
Dim buf As String

If Index = 23 Then 'salir
   losao94_Click

End If
If Index = 22 Then 'control personal
   tingper.Show 1

End If
If Index = 21 Then  'guarda pedido
   'proceso_cierre_automatico
   proceso_cierre_efectivo
   'proceso_cierre_pedido
   
End If
If Index = 20 Then 'cierre caja
Dim sw As Integer
flag_clave1 = 0
tconcla.X = "CIERRE"
tconcla.Show 1
If flag_clave1 = 0 Then  'si es descongela
   'Label27_Click
   Exit Sub
End If
    
    opcion1 = "5"
    opcion2 = "1"
    opcion3 = "2"
    tcuadrc1.fechai.Enabled = True
    tcuadrc1.fechaf.Enabled = True
    usuariopos = gusuario
    tcuadrc1.cajero = "%" 'usuariopos
    tcuadrc1.caja = caja
    tcuadrc1.turno = "%" 'turno
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.fechai = dia
    tcuadrc1.fechaf = dia
    tcuadrc1.Caption = "CIERRE DEL DIA"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1

End If
If Index = 19 Then 'egreso

If dbgrid6.Visible = True Then Exit Sub
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub


   gofpago = "fpagov"
   found = copiar_recibos()
   If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
   End If
   fgusuario = "_l" & gusuario
   found = copiar_tmpfpagoR()
   If found = 0 Then
      MsgBox "No se puede copiar temporal tmpfpagor", 48, "Aviso"
      Exit Sub
   End If

'found = copiar_recibos()
'If found = 0 Then
'   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
'   End
'   Exit Sub
'End If
'explreci.afecta = "P"  'proveedor
'explreci.acu = "V"

fgusuario = "_r" & gusuario
trecaja.xcuentaco = "cuentap"
trecaja.pagocash.Visible = True
trecaja.pagocash.Value = 1
trecaja.XCUENTACO1 = "cuentapd"
trecaja.tipoclie = "P"
trecaja.Combo2.Enabled = True

trecaja.Caption = "EGRESO DINERO"
trecaja.local1 = "" & mytable11.Fields("local")
trecaja.serie = "" & mytable11.Fields("seriere")
'trecaja.local1.Enabled = False
'trecaja.afecta.Enabled = True
trecaja.afecta = "P"
trecaja.acu = "V"
trecaja.cajero = cajero
trecaja.caja = caja
trecaja.turno = turno
trecaja.fecha = dia
trecaja.dia = dia
trecaja.ch89343.Visible = True
trecaja.d7823.Visible = True

trecaja.fecha.Enabled = False
trecaja.Show 1

End If
If Index = 18 Then 'ingreso

If dbgrid6.Visible = True Then Exit Sub
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
'gofpago = "fpagov"
'found = copiar_recibos()
'If found = 0 Then
'   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
'   End
'   Exit Sub
'End If

  found = copiar_recibos()
   If found = 0 Then
   MsgBox "Error al copiar archivo temporal", 24, "Aviso"
   Exit Sub
   End If
   fgusuario = "_l" & gusuario
   found = copiar_tmpfpagoR()
   If found = 0 Then
      MsgBox "No se puede copiar temporal tmpfpagor", 48, "Aviso"
      Exit Sub
   End If


fgusuario = "_r" & gusuario
trecaja.pagocash.Visible = True
trecaja.xcuentaco = "cuentac"
trecaja.XCUENTACO1 = "cuentacd"
trecaja.tipoclie = "C"
trecaja.Combo2.Enabled = True

trecaja.Caption = "INGRESO DINERO"
trecaja.afecta = "C"
trecaja.pagocash.Value = 1

trecaja.local1 = "" & mytable11.Fields("local")
trecaja.serie = "" & mytable11.Fields("serieri")
trecaja.acu = "W"
trecaja.cajero = cajero
trecaja.caja = caja
trecaja.turno = turno
trecaja.fecha = dia
trecaja.dia = dia
trecaja.fecha.Enabled = False
trecaja.ch89343.Visible = True
trecaja.d7823.Visible = True

trecaja.Show 1

End If
If Index = 17 Then 'servicios

If Frame2.Visible = True Then Exit Sub
flag_servicio = "A"
salon = ""
mesa = ""
mesero = ""
cuenta_separa = ""
local1 = ""
local1.Visible = False
found = sumar_detalle()
If found = 0 Then
   MsgBox "debe de Existir un Precio=0", 48, "Aviso"
   DBGrid2.SetFocus
   Exit Sub
End If
If Val(txtotal) = 0 Then
   If exisdev <> -10 Then  'si existe devolucion
     If Val(ntcant) = 0 Then
        DBGrid2.SetFocus
        Exit Sub
     End If
      
   End If
End If
If mytable11.Fields("terminal") = "T" Or Val(acuenta) > 0 And Len(petipo) = 0 Then 'pedidos o a cuenta ha dado
          Frame7.Visible = True
          habilita_lab7 1
          Framefp.Enabled = False
          If Val(acuenta) > 0 Then
             xtipo = "" & mytable11.Fields("tipope")
          End If
          If flag_servicio <> "A" And flag_servicio <> "C" And flag_servicio <> "D" Then
                     xtipo = "5"
                  End If
          xtipo.SetFocus
          Exit Sub
End If
consulta_servicios ""

End If
If Index = 16 Then  'combinaciones
On Error GoTo cmd_6111err
If Len(DBGrid2.columns(0)) > 0 Then
      'cargar_tmcombina
      tcombina.producto = Trim("" & DBGrid2.columns(0))
      tcombina.Show 1
   End If
Exit Sub
cmd_6111err:
Exit Sub

End If
If Index = 15 Then ' comentario
On Error GoTo cmd_611err
If Len(DBGrid2.columns(0)) > 0 Then
      ingreso_locales
   End If
Exit Sub
cmd_611err:
Exit Sub


End If
If Index = 14 Then  'graba comandas

On Error GoTo cmd56123_err
If xopciones(0).Enabled = True Then  'si es autoservicio activado esta en la caja
   If "" & mytable11.Fields("pmesero") = "N" Then
      sin_meseros
      Exit Sub
   End If
   If dbvarios.State = 1 Then dbvarios.Close
   dbvarios.Open "select Nombre,Codigo from vendedor  order by nombre ", cn, adOpenStatic, adLockOptimistic
   If dbvarios.RecordCount = 0 Then
       dbvarios.Close
       Exit Sub
   End If
   Set table6.DataSource = dbvarios
   Frame8.Visible = True
   table6.SetFocus
   Exit Sub
End If
flag_clave1 = 0
tconcla.X = "COMANDA"  '
tconcla.Show 1
If flag_clave1 <> 1 Then  'si es descongela
   Exit Sub
End If
'-------------SI ES COMANDA-----------------
flag_comanda = 0

'MsgBox flag_comanda
tcomanda.Show 1
If flag_comanda = "1" Then
   'MsgBox "paso"
   flag_servicio = "C"
   found = orden_despacho()
   borrar_todo
   sql_detalle
   tiposervicio1 = "Autoservicio"
   flag_servicio = "A"
   Frame8.Visible = False
End If
consulta_comanda "" & mytable11.Fields("salon")
Exit Sub
cmd56123_err:
'MsgBox "Seleccione un Salon Y Mesa ", 48, "Aviso"
Exit Sub

End If
If Index = 13 Then  'descongela


If Val(txtotal) = 0 Then
   MsgBox "No existen Productos Ingresados", 48, "Aviso"
   DBGrid2.SetFocus
   Exit Sub
End If
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub

If Frame9.Visible = True Then
   xcongelax.SetFocus
   Exit Sub
End If
Frame9.Visible = True
xcongelax = ""
xcongelax.SetFocus

End If
If Index = 12 Then  'congela

If Val(txtotal) > 0 Then
   MsgBox "No deben existir Productos ", 48, "Aviso"
   DBGrid2.SetFocus
   Exit Sub
End If
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
If Frame9.Visible = True Then Exit Sub
Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "150"
sw_consulta = 0
found = sql_consulta(1)
dbGrid1.Enabled = True
'dbGrid1.SetFocus

End If
If Index = 11 Then  'abre gaveta

If dbgrid6.Visible = True Then Exit Sub
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
flag_clave1 = 0
tconcla.X = "APERTURA"
tconcla.Show 1
If flag_clave1 <> 1 Then  'si es abre cajon
   DBGrid2.SetFocus
   Exit Sub
End If
If "" & mytable11.Fields("terminal") = "T" Then
   MsgBox "No permitido en Pedido", 48, "Aviso"
   DBGrid2.SetFocus
Exit Sub
End If
found = abre_puerto("" & mytable11.Fields("capuerto"), Val("" & mytable11.Fields("catipo")), "" & mytable11.Fields("gavetacola"))

End If
If Index = 10 Then  'borra linea

On Error GoTo cmd7888_err
If DBGrid2.Row = -1 Then
        MsgBox "No hay ningún registro para eliminar", vbInformation
        Exit Sub
End If
'If MsgBox("Se va a eliminar el registro : está seguro ", _
'   vbExclamation + vbYesNo, "Eliminar") = vbYes Then
   Data2.Recordset.Delete
   If Data2.Recordset.EOF = True And Data2.Recordset.BOF = True Then
      Exit Sub
   End If
   found = sumar_detalle()
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus
   Exit Sub
'End If
Exit Sub
cmd7888_err:
MsgBox "Aviso en Borra Linea " + error$, 48, "Aviso"
Exit Sub

End If
If Index = 9 Then  'anula venta
If dbgrid6.Visible = True Then Exit Sub
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub

flag_clave1 = 0
tconcla.X = "ANULA"  '
tconcla.Show 1
If flag_clave1 <> 1 Then  'si es descongela
   DBGrid2.SetFocus
   Exit Sub
End If

cgusuario = gocabeza
dgusuariog = godetalle
menu_anula1

End If
If Index = 8 Then 'copia ventas
If dbgrid6.Visible = True Then Exit Sub
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
'-------------------------------------
If "" & mytable11.Fields("clavecopia") = "S" Then
flag_clave1 = 0
tconcla.X = "COPIA"  '
tconcla.Show 1
If flag_clave1 <> 1 Then  'si es descongela
   DBGrid2.SetFocus
   Exit Sub
End If
End If


   cgusuario = gocabeza
   dgusuariog = godetalle
   menu_copia
   Exit Sub


'-------------------------------------

End If
If Index = 7 Then 'cuadre parcial
dcupar1_Click

End If

If Index = 6 Then 'limpia pedido
If Frame2.Visible = True Then Exit Sub
If MsgBox("Desea Borrar ??", 1, "Aviso") <> 1 Then Exit Sub
borrar_todo
sql_detalle
tiposervicio1 = "Autoservicio"
flag_servicio = "A"
salon = ""
mesa = ""
mesero = ""
cuenta_separa = ""


End If
If Index = 5 Then 'descuento pedido actual
If dbgrid6.Visible = True Then Exit Sub
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
'-------------------------------------
flag_clave1 = 0
tconcla.X = "DESCUENTO"
tconcla.Show 1
If flag_clave1 <> 1 Then  'si es descongela
   DBGrid2.SetFocus
   Exit Sub
End If


Trecarg.total = txtotal
Trecarg.Show 1

grabar_descto

End If
If Index = 4 Then 'delivery
If Len(telefono) = 0 Then Exit Sub
If Len(dnombre) = 0 Then Exit Sub
salon = ""
mesa = ""
mesero = ""
cuenta_separa = ""
flag_servicio = "D"
proceso_cobross

End If
If Index = 0 Then  'autoservicio
   flag_servicio = "A"
   salon = ""
   mesa = ""
   mesero = ""
   cuenta_separa = ""
   proceso_cobross
End If
If Index = 1 Then  'pago cash
   proceso_cierre_automatico
End If
If Index = 2 Then  'cuenta separada
On Error GoTo cm889222_err
If cmytablex.RecordCount = 0 Then Exit Sub
buf = "Salon " & cmytablex.Fields("salon")
buf = buf & "Mesa " & cmytablex.Fields("mesa")
If MsgBox("Desea Cobrar " & buf, 1, "Cobrar Comanda ") <> 1 Then Exit Sub
found = carga_comanda(1)
If found = 0 Then
   MsgBox "No se Puede Cobrar,No existen Comandas ", 48, "Aviso"
   Exit Sub
End If
salon = cmytablex.Fields("salon")
mesa = cmytablex.Fields("mesa")
flag_servicio = "C"
'Label55_Click
cuenta_separa = "S"
proceso_cobross
'cuenta_separa = ""
Exit Sub
cm889222_err:
'MsgBox "Seleccione Salon Y Mesa ", 48, "Aviso"
Exit Sub

End If

If Index = 3 Then 'cobrar mesa


On Error GoTo cm89222_err
If cmytablex.RecordCount = 0 Then Exit Sub
buf = "Salon " & cmytablex.Fields("salon") '& Chr$("impuesto") & Chr$("neto")
buf = buf & "Mesa " & cmytablex.Fields("mesa")
If MsgBox("Desea Cobrar " & buf, 1, "Cobrar Comanda ") <> 1 Then Exit Sub
found = carga_comanda(0)
If found = 0 Then
   MsgBox "No se Puede Cobrar,No existen Comandas ", 48, "Aviso"
   Exit Sub
End If
salon = cmytablex.Fields("salon")
mesa = cmytablex.Fields("mesa")
'cproven = cmytablex.Fields("vendedor")

flag_servicio = "C"
'Label55_Click
proceso_cobross
Exit Sub
cm89222_err:
MsgBox "Seleccione Salon Y Mesa ", 48, "Aviso"
Exit Sub

End If

End Sub

Private Sub xruc_DblClick()
tkeyboar.flag = "RUC"
tkeyboar.Show 1
End Sub

Private Sub xruc_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(pedido) > 0 Then
   xnombre.SetFocus
   Exit Sub
End If
If local1.Visible = True Then  'si es traslado
   If Len(xruc) = 0 Then
      xruc.SetFocus
      Exit Sub
   End If
   If "" & mytable11.Fields("bodega") = xruc Then
      MsgBox "Debe ser Otro Almacen ", 48, "Aviso"
      xruc.SetFocus
      Exit Sub
   End If
   found = busca_localx("" & xruc)
   If found = 0 Then
      xruc = ""
      MsgBox "No existe Local ", 48, "Aviso"
      xruc.SetFocus
      Exit Sub
   End If
   xnombre.SetFocus
   Exit Sub
End If
If local1 = "PEDIDO" Then  'pedido a almacen
   If Len(xruc) = 0 Then
      xruc.SetFocus
      Exit Sub
   End If
   xnombre.SetFocus
   Exit Sub
End If
If acu = "B" Or acu = "D" Then
   If Len(xruc) = 0 Then
      xruc.SetFocus
      Exit Sub
   End If
   If Len(xruc) <> 11 Then
      xruc.SetFocus
      Exit Sub
   End If
      found = valida_ruc("" & xruc)
   If found = 0 Then
      MsgBox "Ruc no Valido", 48, "Aviso"
      xruc = ""
      xruc.SetFocus
      Exit Sub
   End If
   'valida el ruc
End If


If Len(xruc) > 0 Then
   found = busca_codigocl("" & xruc, 1)
   If found = 0 Then
   End If
End If
If xtipo = "7" Then
   xnombre.SetFocus
   Exit Sub
End If
If "" & mytable11.Fields("cliente") = "S" Or acu = "B" Or acu = "D" Then
   xnombre.SetFocus
   Exit Sub
End If
Command13_Click
End Sub

Private Sub xruc_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = &H26 Then
   xvendedor.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   If local1.Visible <> True Then
      consulta_xruc
   End If
     If local1.Visible = True Then
      consulta_xruc2
   End If

End If
End Sub

Private Sub xserie_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
xnumero.SetFocus

End Sub

Private Sub xserie_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   xdistrito.SetFocus
   Exit Sub
End If
End Sub

Private Sub xtipo_DblClick()
If local1.Visible <> True Then  'si no es traslado
      consulta_xtipo
   End If
End Sub

Private Sub xtipo_keyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   Frame7.Visible = False
   habilita_lab7 0
   If Framefp.Visible = False Then
      DBGrid2.Col = 0
      DBGrid2.Row = DBGrid2.VisibleRows - 1
      DBGrid2.SetFocus
      Exit Sub
   End If
   habilita_lab7 0
   Framefp.Enabled = True
   dbgrid10.Enabled = True
   dbgrid10.SetFocus
   Exit Sub
End If
'aqui es donde vamos a poner si modificacion pedido
If flag_servicio <> "A" And flag_servicio <> "C" And flag_servicio <> "D" Then
   xtipo = "5"
End If
If Len(pedido) > 0 Then
   If xtipo <> "P" Then
      xtipo = "P"
      xtipo.SetFocus
      Exit Sub
   End If
   xserie = "P"
   xvendedor.SetFocus
   Exit Sub
End If

'---si es a cuenta ---
If Val(acuenta) > 0 And Len(petipo) = 0 Then
   If xtipo <> "" & mytable11.Fields("tipope") Then
      MsgBox "Tipo documento admitido,solamente,Pedidos", 48, "Aviso"
      xtipo = "" & mytable11.Fields("tipope")
      xtipo.SetFocus
      Exit Sub
   End If
End If
'ojo aqui voy a validar si es traslado de un local a otros
If local1.Visible = True Then
   If xtipo <> "Z" Then
      xtipo = "Z"
      xtipo.SetFocus
      Exit Sub
   End If
   found = busca_xtipo("" & xtipo, 0)
   If found = 0 Then
      xtipo = ""
      MsgBox "No existe Tipo Documento", 48, "Aviso"
      xtipo.SetFocus
      Exit Sub
   End If
   xvendedor.SetFocus
   Exit Sub
End If
If local1 = "PEDIDO" Then 'pedido merca almacen
   If xtipo <> "Q" Then
      xtipo = "Q"
      xtipo.SetFocus
      Exit Sub
   End If
   found = busca_xtipo("" & xtipo, 0)
   If found = 0 Then
      xtipo = ""
      MsgBox "No existe Tipo Documento", 48, "Aviso"
      xtipo.SetFocus
      Exit Sub
   End If
   xvendedor.SetFocus
   Exit Sub
End If
If Len(xtipo) = 0 Then
   xtipo = "" & mytable11.Fields("tipodefa")
   If "" & mytable11.Fields("habilitanota") = "S" Then
      If Val(ttxtotals) <= Val("" & mytable11.Fields("siventa")) Then
         xtipo = "5"
      End If
   End If
   xtipo.SetFocus
   Exit Sub
End If

found = valida_tipo_pago("" & xtipo)
If found = 0 Then
   MsgBox "No permitido ", 48, "Aviso"
   xtipo.SetFocus
   Exit Sub
End If
found = busca_xtipo("" & xtipo, 0)
If found = 0 Then
   xtipo = ""
   MsgBox "No existe Tipo Documento", 48, "Aviso"
   xtipo.SetFocus
   Exit Sub
End If
xruc = codigo
If xtipo = "1" Or xtipo = "3" Or xtipo = "5" Then
   Label36 = "Codigo"
End If
If xtipo = "2" Or xtipo = "4" Then
   Label36 = "Ruc"
   xruc = ""
End If
sentido.Enabled = False
If sentido.Enabled = True Then
   sentido.SetFocus  'se adiciono concar.....
   Exit Sub
End If
If "" & mytable11.Fields("vendedor") = "S" Then
   xvendedor.SetFocus
   Exit Sub
End If
If flag_servicio = "D" Then  'validar el deliveri si ingreso datos
   xvendedor.SetFocus
   Exit Sub
End If
'MsgBox acu
If "" & acu <> "B" And "" & acu <> "D" Then 'si es diferente de factura
   If "" & mytable11.Fields("cliente") <> "S" Then
      Command13_Click
      Exit Sub
   End If
End If
'If "" & mytable11.Fields("cliente") <> "S" Then
'   Command13_Click
'   Exit Sub
'End If

'If xtipo = "1" Or xtipo = "3" Or xtipo = "5" Or xtipo = "7" Then
'   dni.SetFocus
'   Exit Sub
'End If
xruc.SetFocus
End Sub

Private Sub xtipo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   If local1.Visible <> True Then  'si no es traslado
      consulta_xtipo
   End If
End If

End Sub

Private Sub xvendedor_DblClick()
consulta_xvendedor
End Sub

Private Sub xvendedor_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
If Len(xvendedor) = 0 Then
   'xvendedor = gusuario
End If
If local1.Visible = True Or Len(pedido) > 0 Then 'si es traslado
   If Len(xvendedor) = 0 Then
      xvendedor.SetFocus
      Exit Sub
   End If
   found = busca_xvendedor()
   If found = 0 Then
      xvendedor = ""
      MsgBox "No existe Vendedor ", 48, "Aviso"
      xvendedor.SetFocus
      Exit Sub
   End If
   xruc.SetFocus
   Exit Sub
End If
If Len(xvendedor) > 0 Then
found = busca_xvendedor()
If found = 0 Then
   xvendedor = ""
   MsgBox "No existe Vendedor ", 48, "Aviso"
   xvendedor.SetFocus
   Exit Sub
End If
End If
If flag_servicio = "D" Then
   'If Len(xvendedor) = 0 Then
   '   xvendedor.SetFocus
   '   Exit Sub
   'End If
End If
xruc.SetFocus
End Sub

Private Sub xvendedor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   If sentido.Enabled = True Then
      If sentido.Visible = True Then
         sentido.SetFocus
         Exit Sub
      End If
   End If
   xtipo.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_xvendedor
End If

End Sub
Function busca_xtipo(buf As String, sw As Integer)
Dim sdx As Double
Dim buf1 As String
Dim mytablex As New ADODB.Recordset
ntipox = ""
buf1 = buf
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM tipo where  tipo='" & buf1 & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   ntipox = "" & mytablex.Fields("descripcio")
   acu = "" & mytablex.Fields("tipodoc")
   busca_xtipo = 1
   If sw = 0 Then
      If "" & mytablex.Fields("tipodoc") = "Z" Then  'traslado
      'gocabeza = "factura"
      'godetalle = "detalle"
      'gofpago = "fpagov"
      xserie = "" & mytablex.Fields("serie")
      sdx = Val("" & mytablex.Fields("numero")) + 1
      xnumero = "" & sdx
      xserie.Enabled = True
      xnumero.Enabled = True
      End If
      If "" & mytablex.Fields("tipodoc") = "A" Then
      'gocabeza = "factura"
      'godetalle = "detalle"
      'gofpago = "fpagov"
      xserie = "" & mytable11.Fields("seriebm")
      sdx = Val("" & mytable11.Fields("numerobm")) + 1
      If Len(xnumero) = 0 Then
         xnumero = "" & sdx
      End If
      xserie.Enabled = True
      xnumero.Enabled = True
      End If
      If "" & mytablex.Fields("tipodoc") = "B" Then
       xserie = "" & mytable11.Fields("serieFM")
      sdx = Val("" & mytable11.Fields("numeroFM")) + 1
      'xnumero = "" & sdx
      If Len(xnumero) = 0 Then
         xnumero = "" & sdx
      End If
      xserie.Enabled = True
      xnumero.Enabled = True
      End If
      
      If "" & mytablex.Fields("tipodoc") = "C" Then
      xserie = "" & mytable11.Fields("serietb")
      sdx = Val("" & mytable11.Fields("numerotb")) + 1
      xnumero = "" & sdx
      xserie.Enabled = False
      xnumero.Enabled = False
      End If
      
      If "" & mytablex.Fields("tipodoc") = "1" Then
      xserie = "" & mytable11.Fields("serieexo")
      sdx = Val("" & mytable11.Fields("numeroexo")) + 1
      xnumero = "" & sdx
      xserie.Enabled = False
      xnumero.Enabled = False
      End If
      If "" & mytablex.Fields("tipodoc") = "D" Then
      'gocabeza = "factura"
      'godetalle = "detalle"
      'gofpago = "fpagov"
      xserie = "" & mytable11.Fields("serietf")
      sdx = Val("" & mytable11.Fields("numerotf")) + 1
      xnumero = "" & sdx
      xserie.Enabled = False
      xnumero.Enabled = False
      End If
      If "" & mytablex.Fields("tipodoc") = "G" Then
      'gocabeza = "factura"
      'godetalle = "detalle"
      'gofpago = "fpagov"
      xserie = "" & mytable11.Fields("serienv")
      sdx = Val("" & mytable11.Fields("numeronv")) + 1
      xnumero = "" & sdx
      xserie.Enabled = False
      xnumero.Enabled = False
      End If
      If "" & mytablex.Fields("tipodoc") = "N" Then   '
      'gocabeza = "factura"
      'godetalle = "detalle"
      'gofpago = "fpagov"
      xserie = "" & mytable11.Fields("serienc")
      sdx = Val("" & mytable11.Fields("numeronc")) + 1
      xnumero = "" & sdx
      xserie.Enabled = True
      xnumero.Enabled = True
      End If
      If "" & mytablex.Fields("tipodoc") = "F" Then
      'gocabeza = "factura"
      'godetalle = "detalle"
      'gofpago = "fpagov"
      xserie = "" & mytable11.Fields("seriend")
      sdx = Val("" & mytable11.Fields("numerond")) + 1
      xnumero = "" & sdx
      xserie.Enabled = True
      xnumero.Enabled = True
      End If
      'si es pedidos remotos
      If "" & mytablex.Fields("tipodoc") = "I" Then   'pedido a cuenta
         xserie = "" & mytable11.Fields("seriepe")
         sdx = Val("" & mytable11.Fields("numerope")) + 1
         xnumero = "" & sdx
         xserie.Enabled = True
         xnumero.Enabled = True
      End If
      If "" & mytablex.Fields("tipodoc") = "Q" Then  'pedido reposicion
      'gocabeza = "cpedidov"
      'godetalle = "dpedidov"
      'gofpago = "fpagov"
      xserie = "" & mytable11.Fields("caja")
      sdx = Val("" & mytable11.Fields("numerope")) + 1
      xnumero = "" & sdx
      xserie.Enabled = True
      xnumero.Enabled = True
      End If
   End If
End If
vuelve1:
mytablex.Close
 
End Function
Function busca_xvendedor()

Dim mytablex As New ADODB.Recordset
nvendedorx = ""
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM vendedor where  codigo='" & xvendedor & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   
   If "" & mytablex.Fields("esvendedor") = "N" Then
      MsgBox "Usuario No permitido para ser vendedor ", 48, "Aviso"
      mytablex.Close
      Exit Function
   End If
   
   nvendedorx = "" & mytablex.Fields("nombre")
   busca_xvendedor = 1
End If
mytablex.Close
 
End Function
Function busca_xtipog(buf As String)
Dim sdx As Double
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd7888_err
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then  'si existe
      busca_xtipog = 1
      If "" & mytablex.Fields("tipodoc") = "Z" Then
         'mytablex.Edit
         mytablex.Fields("numero") = xnumero
         'mytablex.Fields("uvueltos") = Val(stxtotals)
         'mytablex.Fields("uvueltod") = Val(stxtotald)
         mytablex.Update
      End If
      If "" & mytablex.Fields("tipodoc") = "1" Then  'exonerado
         'mytable11.Edit
         If Val(tdetra) > 0 Then
         mytable11.Fields("detraccion") = Val(ndetraccion)
         End If

         mytable11.Fields("numeroexo") = xnumero
         mytable11.Fields("uvueltos") = Val(stxtotals)
         mytable11.Fields("uvueltod") = Val(stxtotald)
         mytable11.Update
      End If
      If "" & mytablex.Fields("tipodoc") = "A" Then
         'mytable11.Edit
         'If Val(tdetra) > 0 Then
         'mytable11.Fields("detraccion") = Val(ndetraccion)
         'End If
         mytable11.Fields("numerobm") = xnumero
         mytable11.Fields("uvueltos") = Val(stxtotals)
         mytable11.Fields("uvueltod") = Val(stxtotald)
         mytable11.Update
      End If
      If "" & mytablex.Fields("tipodoc") = "B" Then
         'mytablex.Edit
         mytable11.Fields("numerofm") = xnumero
         'mytable11.Update
         'mytable11.Edit
         'If Val(tdetra) > 0 Then
         'mytable11.Fields("detraccion") = Val(ndetraccion)
         'End If
         mytable11.Fields("uvueltos") = Val(stxtotals)
         mytable11.Fields("uvueltod") = Val(stxtotald)
         mytable11.Update
      End If
      If "" & mytablex.Fields("tipodoc") = "C" Then
         'mytable11.Edit
         If Val(tdetra) > 0 Then
         mytable11.Fields("detraccion") = Val(ndetraccion)
         End If
         mytable11.Fields("numerotb") = xnumero
         mytable11.Fields("uvueltos") = Val(stxtotals)
         mytable11.Fields("uvueltod") = Val(stxtotald)
         mytable11.Update
      End If
      If "" & mytablex.Fields("tipodoc") = "D" Then
         'mytable11.Edit
         'If Val(tdetra) > 0 Then
         'mytable11.Fields("detraccion") = Val(ndetraccion)
         'End If
         mytable11.Fields("numerotf") = xnumero
         mytable11.Fields("uvueltos") = Val(stxtotals)
         mytable11.Fields("uvueltod") = Val(stxtotald)
         If Val(ndetraccion) > 0 Then
         mytable11.Fields("detraccion") = Val(ndetraccion)
         End If
         mytable11.Update
      End If
      If "" & mytablex.Fields("tipodoc") = "G" Then
         'mytable11.Edit
         'If Val(tdetra) > 0 Then
         'mytable11.Fields("detraccion") = Val(ndetraccion)
         'End If
         mytable11.Fields("numeronv") = xnumero
         mytable11.Fields("uvueltos") = Val(stxtotals)
         mytable11.Fields("uvueltod") = Val(stxtotald)
         If Val(ndetraccion) > 0 Then
         mytable11.Fields("detraccion") = Val(ndetraccion)
         End If
         mytable11.Update
      End If
      If "" & mytablex.Fields("tipodoc") = "N" Then   '
         'mytable11.Edit
         'If Val(tdetra) > 0 Then
         'mytable11.Fields("detraccion") = Val(ndetraccion)
         'End If
         mytable11.Fields("numeronc") = xnumero
         mytable11.Fields("uvueltos") = Val(stxtotals)
         mytable11.Fields("uvueltod") = Val(stxtotald)
         mytable11.Update
      End If
      If "" & mytablex.Fields("tipodoc") = "F" Then
         'mytable11.Edit
         'If Val(tdetra) > 0 Then
         'mytable11.Fields("detraccion") = Val(ndetraccion)
         'End If
         mytable11.Fields("numerond") = xnumero
         mytable11.Fields("uvueltos") = Val(stxtotals)
         mytable11.Fields("uvueltod") = Val(stxtotald)
         mytable11.Update
      End If
      If "" & mytablex.Fields("tipodoc") = "I" Then
         'MsgBox "x"
         'mytable11.Edit
         'If Val(tdetra) > 0 Then
         'mytable11.Fields("detraccion") = Val(ndetraccion)
         'End If
         mytable11.Fields("numerope") = xnumero
         mytable11.Fields("uvueltos") = Val(stxtotals)
         mytable11.Fields("uvueltod") = Val(stxtotald)
         mytable11.Update
      End If
      If "" & mytablex.Fields("tipodoc") = "Q" Then
         'MsgBox "x"
         'mytable11.Edit
         'If Val(tdetra) > 0 Then
         'mytable11.Fields("detraccion") = Val(ndetraccion)
         'End If
         mytable11.Fields("numerope") = xnumero
         mytable11.Fields("uvueltos") = Val(stxtotals)
         mytable11.Fields("uvueltod") = Val(stxtotald)
         mytable11.Update
      End If
End If
mytablex.Close
Exit Function
cmd7888_err:
MsgBox "Error busa_xtipog " + error$, 48, "Aviso"
Exit Function
End Function


Function valida_total()
Dim found As Integer
If Len(xtipo) = 0 Then
   xtipo.SetFocus
   Exit Function
End If
If local1.Visible = True Then  'si es traslado
   If xtipo <> "Z" Then
      xtipo = "Z"
      xtipo.SetFocus
      Exit Function
   End If
   found = busca_xtipo("" & xtipo, 0)
   If found = 0 Then
      xtipo = ""
      MsgBox "No existe Tipo Documento", 48, "Aviso"
      xtipo.SetFocus
      Exit Function
   End If
   found = busca_xvendedor()
   If found = 0 Then
      xvendedor = ""
      MsgBox "No existe Vendedor ", 48, "Aviso"
      xvendedor.SetFocus
      Exit Function
   End If
   If Len(xruc) = 0 Then
      xruc.SetFocus
      Exit Function
   End If
   If "" & mytable11.Fields("bodega") = xruc Then
      MsgBox "Debe ser Otro Local ", 48, "Aviso"
      xruc.SetFocus
      Exit Function
   End If
   found = busca_localx("" & xruc)
   If found = 0 Then
      xruc = ""
      MsgBox "No existe Local ", 48, "Aviso"
      xruc.SetFocus
      Exit Function
   End If
   valida_total = 1
   Exit Function
End If
'------------------------------------------------
If local1 = "PEDIDO" Then 'si es pedido almacen
   If xtipo <> "Q" Then
      xtipo = "Q"
      xtipo.SetFocus
      Exit Function
   End If
   found = busca_xtipo("" & xtipo, 0)
   If found = 0 Then
      xtipo = ""
      MsgBox "No existe Tipo Documento", 48, "Aviso"
      xtipo.SetFocus
      Exit Function
   End If
   found = busca_xvendedor()
   If found = 0 Then
      xvendedor = ""
      MsgBox "No existe Vendedor ", 48, "Aviso"
      xvendedor.SetFocus
      Exit Function
   End If
   'If Len(xruc) = 0 Then
   '   xruc.SetFocus
   '   Exit Function
   'End If
   valida_total = 1
   Exit Function
End If

'------------------------------------------------
found = valida_tipo_pago("" & xtipo)
If found = 0 Then
   MsgBox "No permitido ", 48, "Aviso"
   xtipo.SetFocus
   Exit Function
End If

found = busca_xtipo("" & xtipo, 0)
If found = 0 Then
   xtipo = ""
   MsgBox "No existe Tipo Documento", 48, "Aviso"
   xtipo.SetFocus
   Exit Function
End If
If Val(acuenta) > 0 And Len(petipo) = 0 Then
   If xtipo <> "" & mytable11.Fields("tipope") Then
      MsgBox "Tipo documento admitido,solamente,Pedidos", 48, "Aviso"
      xtipo = "" & mytable11.Fields("tipope")
      xtipo.SetFocus
      Exit Function
   End If
End If
If sentido.Enabled = True Then
   If sentido <> "S" And sentido <> "B" Then
      sentido = ""
      If sentido.Visible = True Then
      sentido.SetFocus
      End If
      Exit Function
   End If
End If


If Len(xvendedor) > 0 Then
   found = busca_xvendedor()
   If found = 0 Then
      xvendedor = ""
      MsgBox "No existe Vendedor ", 48, "Aviso"
      xvendedor.SetFocus
      Exit Function
   End If
End If
If "" & mytable11.Fields("obligavendedor") = "S" Then
   If Len(xvendedor) = 0 Then
      xvendedor.SetFocus
      Exit Function
   End If
   
End If
If flag_servicio = "D" Then  'validar el deliveri si ingreso datos
  'If Len(xvendedor) = 0 Then
  '    xvendedor.SetFocus
  '    Exit Function
  ' End If
End If
If xtipo = "7" Then
   If Len(xnombre) = 0 Then
      xnombre.SetFocus
      Exit Function
   End If
   If Len(xdistrito) = 0 Then
      xdistrito.SetFocus
      Exit Function
   End If
End If

If "" & mytable11.Fields("cliente") = "S" Then
   'If xtipo = "2" Or xtipo = "4" Then
   '   If Len(xruc) = 0 Then
   '      xruc.SetFocus
   '      Exit Function
   '   End If
   '   If Len(xnombre) = 0 Then
   '      xnombre.SetFocus
   '      Exit Function
   '   End If
   'End If
End If
If acu = "B" Or acu = "D" Then
   If Len(xruc) = 0 Then
      xruc.SetFocus
      Exit Function
   End If
   If Len(xruc) <> 11 Then
      xruc.SetFocus
      Exit Function
   End If
   found = valida_ruc("" & xruc)
   If found = 0 Then
      MsgBox "Ruc no Valido", 48, "Aviso"
      xruc = ""
      xruc.SetFocus
      Exit Function
   End If
   'valida el ruc
End If
If Len(xruc) > 0 Then
   found = busca_codigocl("" & xruc, 1)
   If acu = "B" Or acu = "D" Then
      'If Len(xnombre) = 0 Then
      '   xnombre.SetFocus
      '   Exit Function
      'End If
   End If
   'If found = 0 Then
   '   xruc = ""
   '   MsgBox "No existe Codigo/Ruc", 48, "Aviso"
   '   xruc.SetFocus
   '   Exit Function
   'End If
End If
valida_total = 1
End Function
Function graba_fpagov(bxtipo As String, bxserie As String, bxnumero As String)
Dim xbuf As String
Dim mytabley As New ADODB.Recordset
Dim mytablex As Table
Dim found As Integer
On Error GoTo cdm4411_err
'---------- validando si es cuenta corriente
'If mytablex.State = 1 Then mytablex.Close
'mytablex.Open "SELECT * FROM " & fpusuario, cn, adOpenDynamic, adLockOptimistic
'If mytablex.RecordCount = 0 Then  'si no existe
'   mytablex.Close
'   Exit Function
'End If
amk223:
mytabley.Open "SELECT * FROM " & gofpago & " where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytabley.RecordCount > 0 Then  'si existe
   mytabley.Delete
   GoTo amk223
End If

'mytabley.Open "SELECT * FROM " & gofpago, cn, adOpenDynamic, adLockOptimistic
'If mytabley.RecordCount = 0 Then  'si no existe
'   mytabley.Close
'   Exit Function
'End If
'mytablez.Open "SELECT * FROM cuentac ", cn, adOpenDynamic, adLockOptimistic
'If mytablez.RecordCount = 0 Then  'si no existe
'   mytablez.Close
'   Exit Function
'End If

xbuf = "antes:" & Format(Now, "hh:mm:ss")

Set mytablex = mydbxglo.OpenTable(fpusuario)
'Set mytabley = mydbxglo.OpenTable(gofpago)
'Set mytablez = mydbxglo.OpenTable("cuentac")
'mytabley.Index = "fpagov"
Do
If mytablex.EOF Then Exit Do
   If Len("" & mytablex.Fields("fpago")) > 0 Then
      mytabley.AddNew
      grabar_registro_fpagov mytablex, mytabley
      mytabley.Update
      If "" & mytabley.Fields("acufp") = "V" Then
         graba_acumulado_clientes mytabley, 1, Val("" & mytabley.Fields("recibe"))
      End If
   End If
mytablex.MoveNext
Loop


If Len(petipo) > 0 And Len(peserie) > 0 And Len(penumero) > 0 Then
   mytabley.AddNew
   found = forma_pago_adicional(mytabley)
   mytabley.Update
End If
'xbuf = xbuf & "despues:" & Format(Now, "hh:mm:ss")
'sgBox xbuf
mytablex.Close
mytabley.Close
'mytablez.Close
Exit Function
cdm4411_err:
MsgBox "Error en graba_fpagov " + error$, 48, "Aviso"
Exit Function
End Function
Sub grabar_registro_fpagov(mytablex As Table, mytabley As ADODB.Recordset)
On Error GoTo cmd2008_err
   mytabley.Fields("vendedor") = xvendedor
   mytabley.Fields("paridad") = Val("" & paridadfp)
   mytabley.Fields("codigo") = "" & mytablex.Fields("codigo")
   mytabley.Fields("nombre") = "" & mytablex.Fields("nombre")
   mytabley.Fields("tipo") = "" & xtipo
   mytabley.Fields("serie") = "" & xserie
   mytabley.Fields("numero") = "" & xnumero
   mytabley.Fields("tipoclie") = "C"
   
   
   If Len(Trim("" & mytablex.Fields("codigo"))) = 0 Then
      mytabley.Fields("codigo") = "" & xruc
   End If
   If Len(Trim("" & mytablex.Fields("nombre"))) = 0 Then
      mytabley.Fields("nombre") = "" & xnombre
   End If

   mytabley.Fields("fecha") = Format("" & dia, "dd/mm/yyyy")
   mytabley.Fields("moneda") = "" & mytablex.Fields("moneda")
   mytabley.Fields("total") = Val(ttxtotals)
   mytabley.Fields("caja") = "" & caja
   mytabley.Fields("turno") = "" & turno
   mytabley.Fields("usuario") = "" & cajero
   'mytabley.Fields("vendedor") = "" & cajero
   
   mytabley.Fields("total") = Val("" & mytablex.Fields("total"))
   mytabley.Fields("cambio") = Val("" & mytablex.Fields("cambio"))
   mytabley.Fields("recibe") = Val("" & mytablex.Fields("recibe"))
   mytabley.Fields("recibes") = Val("" & mytablex.Fields("recibes"))
   mytabley.Fields("recibed") = Val("" & mytablex.Fields("recibed"))
   mytabley.Fields("saldos") = Val("" & mytablex.Fields("saldos"))
   mytabley.Fields("saldod") = Val("" & mytablex.Fields("saldod"))
   
   'mytabley.Fields("nombre") = "" & mytablex.Fields("nombre")
   mytabley.Fields("orden") = "" & mytablex.Fields("orden")
   mytabley.Fields("observa") = "" & mytablex.Fields("observa")
   
   'MsgBox "" & mytablex.Fields("dias")
   mytabley.Fields("dias") = "" & mytablex.Fields("dias")
   mytabley.Fields("fpago") = "" & mytablex.Fields("fpago")
   
   mytabley.Fields("acufp") = busca_fpago("" & mytablex.Fields("fpago"))
   
   mytabley.Fields("descripcio") = "" & mytablex.Fields("descripcio")
   mytabley.Fields("acu") = "" & acu
   
   mytabley.Fields("local") = Trim("" & mytable11.Fields("local"))
   If "" & mytable11.Fields("terminal") = "T" Then
    'mytabley.Fields("acu") = "I"
   End If
   mytabley.Fields("servicio") = flag_servicio
   If flag_servicio = "A" Then
      mytabley.Fields("servicio") = "A"
   End If
   If flag_servicio = "D" Then
      mytabley.Fields("servicio") = "D"
   End If
   If flag_servicio = "C" Then
      mytabley.Fields("servicio") = "C"
   End If
   
   mytabley.Fields("estado") = "2"
   'If "" & mytabley.Fields("acu") = "A" Or "" & mytabley.Fields("acu") = "B" Or "" & mytabley.Fields("acu") = "C" Or "" & mytablex.Fields("acu") = "D" Or "" & mytablex.Fields("acu") = "G" Then  'credito
   If "" & mytablex.Fields("acu") = "C" Then   'credito
         graba_credito mytablex, mytabley
   End If
   If "" & mytablex.Fields("acu") = "I" Or "" & mytablex.Fields("acu") = "K" Then     'ORDEN DE TRABAJO/DEPOSITO/BANCO
      graba_credito2 mytabley, "" & mytablex.Fields("acu")
   End If
   If "" & mytablex.Fields("acu") = "K" Then   'si es deposito a banco
      'graba_deposito mytabley
   End If
   If "" & mytablex.Fields("acu") = "D" Or "" & mytablex.Fields("acu") = "F" Then   'si tarjeta credito o debito
      'graba_tarjetas mytabley
   End If
   If xxacu = "I" Then
      mytabley.Fields("acu") = xxacu
   End If
   If xtipo = "7" Then
   mytabley.Fields("total") = 0
   mytabley.Fields("cambio") = 0
   mytabley.Fields("recibe") = 0
   mytabley.Fields("recibes") = 0
   mytabley.Fields("recibed") = 0
   mytabley.Fields("saldos") = 0
   mytabley.Fields("saldod") = 0
   End If
   mytabley.Fields("flage") = "V"
   
   Exit Sub
cmd2008_err:
   MsgBox "Aviso en grabar_registro_fpagov " + error$, 48, "Aviso"
   Exit Sub
End Sub
Function busca_fpago(buf As String) As String

Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM fpago where   fpago='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   busca_fpago = "" & mytablex.Fields("tipo")
End If
'------------------------------------- ------------
mytablex.Close
 
End Function
Function grabar_telefono()
 
End Function
Function ver_si_puedo_dbgrid(buf As String)
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "SELECT * FROM precios where producto='" & buf & "' and local='" & "" & mytable11.Fields("listap") & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.EOF = True Or mytablex.BOF = True Then
      mytablex.Close
      Exit Function
   End If
   If "" & Len("" & mytablex.Fields("unidad1")) > 0 Then
     If "" & Len("" & mytablex.Fields("unidad2")) > 0 Then
        ver_si_puedo_dbgrid = 1
     End If
   End If
mytablex.Close
 
End Function
Sub menu_anula1()
Dim found As Integer
Frame1.Visible = True
Frame1.Enabled = True
   Combo1.Clear
   Combo1.AddItem "Nombre"
   Combo1.ListIndex = 0
buffer = ""
opcion1 = "100"
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus

End Sub
Sub menu_copia()

Dim found As Integer
Dim buf As String
   Combo1.Clear
   Combo1.AddItem "Nombre"
   Combo1.ListIndex = 0

Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "15"
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus
End Sub
Sub menu_proforma()
Dim found As Integer
Combo1.Clear
   Combo1.AddItem "Nombre"
   Combo1.ListIndex = 0
Frame1.Visible = True
Frame1.Enabled = True
buffer = ""
opcion1 = "1900"
sw_consulta = 0
found = sql_consulta(1)
'dbGrid1.SetFocus
End Sub


Function proceso_anular(ytipo As String, yserie As String, ynumero As String)
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd_4356
Dim found As Integer
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM " & gocabeza & " where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & ytipo & "' and serie='" & yserie & "' and numero='" & ynumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   'mytablex.Edit
   mytablex.Fields("estado") = "1"
   mytablex.Update
End If
mytablex.Close

If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM " & godetalle & " where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & ytipo & "' and serie='" & yserie & "' and numero='" & ynumero & "'", cn, adOpenStatic, adLockOptimistic  'adOpenDynamic
If mytablex.RecordCount > 0 Then 'si existe
found = descarga_saldo("" & mytable11.Fields("local"), mytablex, ytipo, yserie, ynumero, 1, 1)
'Set mytablex = mydbxglo.OpenTable("detalle")
'mytablex.Index = "tdetalle"
'mytablex.Seek "=", ytipo, yserie, ynumero
'If Not mytablex.NoMatch Then
'   mytablex.Edit
'   mytablex.Fields("estado") = "1"
'   mytablex.Update
End If
mytablex.Close
'MsgBox "123"
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM " & gofpago & " where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & ytipo & "' and serie='" & yserie & "' and numero='" & ynumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   Do
   If mytablex.EOF Then Exit Do
      'mytablex.Edit
      mytablex.Fields("estado") = "1"
      mytablex.Update
      If "" & mytablex.Fields("acufp") = "V" Then
         graba_acumulado_clientes mytablex, -1, Val("" & mytablex.Fields("recibe"))
      End If
      found = borra_credito(ytipo, yserie, ynumero)
      'If "" & mytablex.Fields("acufp") = "I" Then
      '  found = anula_tmpcta(mytablex)
      'End If
      desgraba_deposito mytablex
   mytablex.MoveNext
   Loop
End If
mytablex.Close
reversa_guia_mensual "" & mytable11.Fields("local"), ytipo, yserie, ynumero
proceso_anular = 1
Exit Function
cmd_4356:
MsgBox "Aviso en proceso anula " + error$, 48, "Aviso"
Exit Function
End Function
Function graba_cliente_credito1(buf As String)
Dim mytablex As New ADODB.Recordset
If Len(buf) = 0 Then Exit Function
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM clientes where  codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount = 0 Then 'si existe
   mytablex.AddNew
   mytablex.Fields("codigo") = "" & tcampo1
   mytablex.Fields("nombre") = "" & tcampo2
   mytablex.Update
End If
mytablex.Close
End Function
Function graba_cliente_tipo(buf As String)
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
Dim sdx As Double
Dim buf1 As String
Dim codigogen As String
On Error GoTo cmdd7812_err

'If Len(buf) = 0 Then Exit Function
'If Len(xnombre) = 0 Then Exit Function
'If Len(buf) = 0 Then Exit Function

If Len(xruc) = 0 And Len(xnombre) > 0 Then 'no no tiene codigo
   mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.Close
      Exit Function
   End If
   sdx = Val("" & mytablex.Fields("clientes")) + 1
   codigogen = "" & sdx
   mytablex.Close
sigueb1:
   mytablex.Open "select * from clientes where codigo='" & codigogen & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      mytablex.Close
      sdx = sdx + 1
      codigogen = "" & sdx
      GoTo sigueb1
   End If
      xruc = codigogen
      mytablex.AddNew
      mytablex.Fields("codigo") = "" & xruc
      mytablex.Fields("tipo") = "O"
      mytablex.Fields("nombre") = "" & xnombre
      mytablex.Fields("direccion") = "" & xdireccion
      mytablex.Update
      xruc = "" & mytablex.Fields("codigo")
      'codigo = "" & mytablex.Fields("codigo")
      'nombre = "" & mytablex.Fields("nombre")
      mytablex.Close
      Exit Function
End If
If Len(xruc) > 0 And Len(xnombre) > 0 Then
      mytablex.Open "SELECT * FROM clientes  where  codigo='" & xruc & "'", cn, adOpenDynamic, adLockOptimistic
      If mytablex.RecordCount > 0 Then
         mytablex.Fields("nombre") = Trim("" & xnombre)
         If Len("" & xdireccion) > 0 Then
            mytablex.Fields("direccion") = Trim("" & xdireccion)
         End If
         mytablex.Update
      Else
         mytablex.AddNew
         mytablex.Fields("nombre") = "" & xnombre
         mytablex.Fields("codigo") = "" & xruc

         If xtipo = "2" Or xtipo = "4" Then
               mytablex.Fields("tipo") = "J"
               Else
               mytablex.Fields("tipo") = "O"
         End If
         If Len("" & xdireccion) > 0 Then
            mytablex.Fields("direccion") = "" & xdireccion
         End If
         mytablex.Update
      End If
      mytablex.Close
End If
Exit Function
cmdd7812_err:
MsgBox "Aviso en graba cliente tipo " + error$, 48, "Aviso"
Exit Function
      
   
  
End Function

Function graba_credito(mytabley As Table, mytablez As ADODB.Recordset)
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd6712121_err
'MsgBox ""
mytablex.Open "SELECT * FROM cuentac where local='" & Trim("" & mytablez.Fields("local")) & "' and tipo='" & Trim("" & mytablez.Fields("tipo")) & "' and serie='" & Trim("" & mytablez.Fields("serie")) & "' and numero='" & Trim("" & mytablez.Fields("numero")) & "' and cuota='1'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount = 0 Then  'si no existe
   mytablex.AddNew
   mytablex.Fields("grupo") = "C"
   'MsgBox ""
   mytablex.Fields("acu") = "" & acu
   mytablex.Fields("observa") = Mid$("" & mytabley.Fields("descripcio"), 1, 30)
   mytablex.Fields("fpago") = "" & mytablez.Fields("acufp")
   mytablex.Fields("tipo") = "" & mytablez.Fields("tipo")
   mytablex.Fields("serie") = "" & mytablez.Fields("serie")
   mytablex.Fields("numero") = "" & mytablez.Fields("numero")
   mytablex.Fields("dias") = Val("" & mytabley.Fields("dias"))
   mytablex.Fields("cuota") = "1"
   mytablex.Fields("tipoclie") = "C"
   mytablex.Fields("codigo") = "" & mytablez.Fields("codigo")
   mytablex.Fields("nombre") = "" & mytablez.Fields("nombre")
   'MsgBox ""
   mytablex.Fields("fecha") = Format("" & mytablez.Fields("fecha"), "dd/mm/yyyy")
   mytablex.Fields("fechav") = Format("" & mytablez.Fields("fecha") + Val("" & mytabley.Fields("dias")), "dd/mm/yyyy")
   'MsgBox "1"
   mytablex.Fields("moneda") = "" & mytabley.Fields("moneda")
   mytablex.Fields("total") = Val("" & mytabley.Fields("recibe"))
   mytablex.Fields("abono") = 0
   mytablex.Fields("interes") = 0
   mytablex.Fields("saldo") = Val("" & mytabley.Fields("recibe"))
   'mytablex.Fields("c1") = Val("" & mytablez.Fields("c1"))
   'mytablex.Fields("c2") = Val("" & mytablez.Fields("c2"))
   'mytablex.Fields("c3") = Val("" & mytablez.Fields("c3"))
   'mytablex.Fields("c4") = Val("" & mytablez.Fields("c4"))
   'mytablex.Fields("c5") = Val("" & mytablez.Fields("c5"))
   'mytablex.Fields("c6") = Val("" & mytablez.Fields("c6"))
   'mytablex.Fields("c7") = Val("" & mytablez.Fields("c7"))
   'mytablex.Fields("c8") = Val("" & mytablez.Fields("c8"))
   'mytablex.Fields("c9") = Val("" & mytablez.Fields("c9"))
   mytablex.Fields("estado") = "0"
   mytablex.Fields("vendedor") = xvendedor
   mytablex.Fields("usuario") = cajero
   mytablex.Fields("caja") = caja
   mytablex.Fields("turno") = turno
   mytablex.Fields("zona") = ""
   mytablex.Fields("local") = "" & mytable11.Fields("local")
   mytablex.Update
End If
mytablex.Close
Exit Function
cmd6712121_err:
MsgBox "Aviso en Graba Credito " + error$, 48, "Aviso"
Exit Function
End Function
Function graba_credito2(mytablez As ADODB.Recordset, buf As String) 'adelantos
Dim mytablex As New ADODB.Recordset
Dim mytabley As New ADODB.Recordset
Dim sdx As Double
Dim sdx1 As Double
Dim buf1 As String
If buf = "I" Then
   buf1 = "A"
End If
If buf = "K" Then
   buf1 = "D"
End If

sdx = Val("" & mytablez.Fields("total"))
mytabley.Open "SELECT * FROM cuentacd ", cn, adOpenDynamic, adLockOptimistic
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM cuentac where  tipoclie='C' and codigo='" & Trim("" & mytablez.Fields("codigo")) & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount = 0 Then 'si existe
   mytablex.Close
   Exit Function
End If
   Do
   If mytablex.EOF Then Exit Do
      If "" & mytablex.Fields("grupo") = buf1 Then
         If Val("" & mytablex.Fields("saldo")) > 0 Then
         '------------------------------------------------
         If sdx > 0 Then
         If Val("" & mytablex.Fields("saldo")) > sdx Then
            'mytablex.Edit
            graba_tmpcta mytablez, mytablex, mytabley, sdx
            mytablex.Fields("abono") = Val("" & mytablex.Fields("abono")) + sdx
            mytablex.Fields("saldo") = Val("" & mytablex.Fields("total")) + Val("" & mytablex.Fields("interes")) - Val("" & mytablex.Fields("abono"))
            mytablex.Update
            Exit Do
         End If
         If Val("" & mytablex.Fields("saldo")) <= sdx Then
            'mytablex.Edit
            sdx = sdx - Val("" & mytablex.Fields("saldo"))
            graba_tmpcta mytablez, mytablex, mytabley, Val("" & mytablex.Fields("saldo"))
            mytablex.Fields("abono") = Val("" & mytablex.Fields("abono")) + Val("" & mytablex.Fields("saldo"))
            mytablex.Fields("saldo") = Val("" & mytablex.Fields("total")) + Val("" & mytablex.Fields("interes")) - Val("" & mytablex.Fields("abono"))
            mytablex.Update
         End If
         End If
         '------------------------------------------------
         End If
      End If
   mytablex.MoveNext
   Loop
End Function
Function anula_tmpcta(mytabley As ADODB.Recordset)
Dim mytablex As New ADODB.Recordset
miramos:
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM cuentacd  where  local='" & Trim("" & mytabley.Fields("local")) & "' and tipo='" & Trim("" & mytabley.Fields("tipo")) & "' and serie='" & Trim("" & mytabley.Fields("serie")) & "' and numero='" & Trim("" & mytabley.Fields("numero")) & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
      cuentac_borra mytablex
      mytablex.Delete
      GoTo miramos
End If
mytablex.Close
End Function
Sub cuentac_borra(mytablex As ADODB.Recordset)
Dim mytablez As New ADODB.Recordset
If mytablez.State = 1 Then mytablez.Close
mytablez.Open "SELECT * FROM cuentac  where  local='" & Trim("" & mytablex.Fields("local1")) & "' and tipo='" & Trim("" & mytablex.Fields("tipo1")) & "' and serie='" & Trim("" & mytablex.Fields("serie1")) & "' and numero='" & Trim("" & mytablex.Fields("numero1")) & "' and cuota='1'", cn, adOpenDynamic, adLockOptimistic
If mytablez.RecordCount > 0 Then 'si existe
   'mytablez.Edit
   mytablez.Fields("abono") = Val("" & mytablez.Fields("abono")) - Val("" & mytablex.Fields("paga"))
   mytablez.Fields("saldo") = Val("" & mytablez.Fields("total")) + Val("" & mytablez.Fields("interes")) - Val("" & mytablez.Fields("abono"))
   mytablez.Update
End If

End Sub
Sub graba_tmpcta(mytablez As ADODB.Recordset, mytablex As ADODB.Recordset, mytabley As ADODB.Recordset, sdx As Double)
On Error GoTo cmd78121_err
mytabley.AddNew

mytabley.Fields("codigo") = "" & mytablez.Fields("codigo")
mytabley.Fields("local") = "" & mytablez.Fields("local")
mytabley.Fields("local1") = "" & mytablez.Fields("local")
mytabley.Fields("tipo") = "" & mytablez.Fields("tipo")
mytabley.Fields("serie") = "" & mytablez.Fields("serie")

mytabley.Fields("numero") = "" & mytablez.Fields("numero")
mytabley.Fields("acu") = "" & mytablez.Fields("acu")
mytabley.Fields("tipo1") = "" & mytablex.Fields("tipo")
mytabley.Fields("serie1") = "" & mytablex.Fields("serie")
mytabley.Fields("numero1") = "" & mytablex.Fields("numero")
mytabley.Fields("cuota") = "" & mytablex.Fields("cuota")
mytabley.Fields("moneda") = "" & mytablex.Fields("moneda")
mytabley.Fields("total") = Val("" & mytablex.Fields("saldo"))
mytabley.Fields("paga") = sdx
mytabley.Fields("estado") = "2"

mytabley.Fields("fecha") = CVDate("" & mytablez.Fields("fecha"))

'mytabley.Fields("hora") = "" & mytablez.Fields("hora")
mytabley.Fields("usuario") = "" & mytablez.Fields("usuario")
mytabley.Fields("caja") = "" & mytablez.Fields("caja")
mytabley.Fields("turno") = "" & mytablez.Fields("turno")

mytabley.Fields("tipoclie") = "" & mytablez.Fields("tipoclie")
mytabley.Update
Exit Sub
cmd78121_err:
MsgBox "Aviso en graba_tmpctaa " + error$, 48, "Aviso"
Exit Sub
End Sub


Function borra_credito(xtipo As String, xserie As String, xnumero As String)
Dim mytablex As New ADODB.Recordset
amk2:
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM cuentac where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "' and cuota='1'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   mytablex.Delete
   GoTo amk2
End If
mytablex.Close
End Function
Function menu_repone(xcongela As String)
Dim i As Integer
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd67112_err
mytablex.Open "SELECT * FROM drequisa  where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='01' and serie='Q' and numero='" & xcongela & "'", cn, adOpenDynamic, adLockOptimistic
Do
    If mytablex.EOF Then Exit Do
       Data2.Recordset.AddNew
       For i = 0 To mytablex.Fields.count - 1
           Data2.Recordset.Fields(i) = mytablex.Fields(i)
       Next i
       Data2.Recordset.Fields("caja") = "" & caja
       Data2.Recordset.Fields("turno") = "" & turno
       Data2.Recordset.Fields("usuario") = "" & cajero
       Data2.Recordset.Fields("fecha") = Format(dia, "dd/mm/yyyy")
       Data2.Recordset.Fields("hora") = Format(Now, "hh:MM")
       Data2.Recordset.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
       Data2.Recordset.Fields("estado") = "2"
       Data2.Recordset.Update
       mytablex.MoveNext
Loop
'--------borrando
mytablex.Close
menu_repone = 1
Exit Function
cmd67112_err:
mytablex.Close
Exit Function
End Function
Function menu_descongela(xcongela As String)
Dim i As Integer
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd6711_err

If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "SELECT * FROM congelad where numero='" & xcongela & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.EOF = True Or mytablex.BOF = True Then
      mytablex.Close
      Exit Function
   End If
   
Do
    If mytablex.EOF Then Exit Do
    If "" & mytablex.Fields("numero") = xcongela Then
       Data2.Recordset.AddNew
       For i = 0 To mytablex.Fields.count - 1
           Data2.Recordset.Fields(i) = mytablex.Fields(i)
       Next i
       Data2.Recordset.Fields("caja") = "" & caja
       Data2.Recordset.Fields("turno") = "" & turno
       Data2.Recordset.Fields("usuario") = "" & cajero
       Data2.Recordset.Fields("fecha") = Format(dia, "dd/mm/yyyy")
       'Data2.Recordset.Fields("hora") = Format(Now, "hh:MM")
       Data2.Recordset.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
       Data2.Recordset.Fields("estado") = "2"
       Data2.Recordset.Update
       mytablex.MoveNext
       Else: Exit Do
    End If
Loop

'--------borrando
mytablex.Close
menu_descongela = 1
DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus 'found = sumar_detalle()

Exit Function
cmd6711_err:
Exit Function
End Function
Sub borrar_descongela(xcongela As String)
cn.Execute ("DELETE   FROM congelac WHERE numero='" & Trim(xcongela) & "'")
End Sub
Sub borrar_reponexx()
On Error GoTo cmd133_err
Data1.Recordset.Delete
Exit Sub
cmd133_err:
Exit Sub

End Sub
Sub borrar_descongela1(xcongela As String)
cn.Execute ("DELETE   FROM congelad WHERE numero='" & Trim(xcongela) & "'")
End Sub
Sub borrar_repone(xcongela As String)
cn.Execute ("DELETE   FROM drequisa WHERE local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='01' and serie='Q' and numero='" & xcongela & "'")
End Sub
Function descarga_saldo(bxlocal As String, mytablex As ADODB.Recordset, bxtipo As String, bxserie As String, bxnumero As String, sw As Integer, sw1 As Integer)
Dim mytabley As New ADODB.Recordset
Dim sdx As Double
  
  Do
  If mytablex.EOF Or mytablex.BOF Then Exit Do
     If sw1 = 1 Then
        mytablex.Fields("estado") = "1"
        mytablex.Update
     End If
     'MsgBox "" & mytablex.Fields("producto")
     
     '--------------------------
     mytabley.Open "SELECT * FROM almacen where  local='" & "" & mytablex.Fields("local") & "' and producto='" & "" & mytablex.Fields("producto") & "' and bodega='" & "" & mytablex.Fields("bodega") & "'", cn, adOpenDynamic, adLockOptimistic
     If mytabley.RecordCount > 0 Then  'si existe
        sdx = Val("" & mytabley.Fields("saldo")) + sw * Val("" & mytablex.Fields("factor")) * Val("" & mytablex.Fields("cantidad"))
        mytabley.Fields("saldo") = sdx
        mytabley.Update
        Else
        mytabley.AddNew
        mytabley.Fields("producto") = "" & mytablex.Fields("producto")
        mytabley.Fields("bodega") = "" & mytablex.Fields("bodega")
        sdx = Val("" & mytabley.Fields("saldo")) + sw * Val("" & mytablex.Fields("factor")) * Val("" & mytablex.Fields("cantidad"))
        mytabley.Fields("saldo") = sdx
        mytabley.Fields("local") = "" & mytable11.Fields("local")
        mytabley.Update
     End If
     mytabley.Close
     Set mytabley = Nothing
     
     '--------------------------
     mytablex.MoveNext
  Loop
  
  

End Function


Function proceso_carga_doc_ant(xlocal As String, xtipo As String, xserie As String, xnumero As String)
Dim i As Integer
Dim found As Integer

Dim mytablex As New ADODB.Recordset
On Error GoTo cmd67112_err

If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM " & godetalle & " where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
Do
    If mytablex.EOF Then Exit Do
       Data2.Recordset.AddNew
       For i = 0 To mytablex.Fields.count - 1
           Data2.Recordset.Fields(i) = mytablex.Fields(i)
       Next i
       Data2.Recordset.Update
       proceso_carga_doc_ant = 1
    mytablex.MoveNext
Loop
End If
mytablex.Close
 found = sumar_detalle()
               DBGrid2.Col = 0
               DBGrid2.Row = DBGrid2.VisibleRows - 1
               DBGrid2.SetFocus
Exit Function
cmd67112_err:
mytablex.Close
 
Exit Function
End Function
Function proceso_carga_Pedido(xlocal As String, xtipo As String, xserie As String, xnumero As String)
Dim i As Integer
Dim found As Integer
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd67112_err

If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM dpedidov where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
Do
    If mytablex.EOF Then Exit Do
       Data2.Recordset.AddNew
       For i = 0 To mytablex.Fields.count - 1
           Data2.Recordset.Fields(i) = mytablex.Fields(i)
       Next i
       Data2.Recordset.Update
       proceso_carga_Pedido = 1
     mytablex.MoveNext
Loop
End If
mytablex.Close
 found = sumar_detalle()
               DBGrid2.Col = 0
               DBGrid2.Row = DBGrid2.VisibleRows - 1
               DBGrid2.SetFocus
Exit Function
cmd67112_err:
mytablex.Close
Exit Function
End Function

Function proceso_carga_cotizacion(xlocal As String, xtipo As String, xserie As String, xnumero As String)
Dim i As Integer
Dim found As Integer
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd67112_err
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM dcotizav where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
Do
    If mytablex.EOF Then Exit Do
       Data2.Recordset.AddNew
       For i = 0 To mytablex.Fields.count - 1
           Data2.Recordset.Fields(i) = mytablex.Fields(i)
       Next i
       Data2.Recordset.Update
       proceso_carga_cotizacion = 1
mytablex.MoveNext
Loop
End If
mytablex.Close
 found = sumar_detalle()
               DBGrid2.Col = 0
               DBGrid2.Row = DBGrid2.VisibleRows - 1
               DBGrid2.SetFocus
Exit Function
cmd67112_err:
mytablex.Close
Exit Function
End Function


Function proceso_proforma(xlocal As String, xtipo As String, xserie As String, xnumero As String)
Dim i As Integer
Dim found As Integer
Dim mytablex As New ADODB.Recordset
Dim sw As Integer
sw = 0
On Error GoTo cmd6711212_err
'MsgBox "" & mytable11.Fields("local") & " " & xtipo & " " & xserie & " " & xnumero
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM dproform where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   Do
    If mytablex.EOF Then Exit Do
       'MsgBox ""
       Data2.Recordset.AddNew
       For i = 0 To mytablex.Fields.count - 2
           Data2.Recordset.Fields(i) = mytablex.Fields(i)
       Next i
       Data2.Recordset.Update
       sw = 1
    mytablex.MoveNext
Loop
End If
mytablex.Close
proceso_proforma = sw
found = sumar_detalle()
               DBGrid2.Col = 0
               DBGrid2.Row = DBGrid2.VisibleRows - 1
               DBGrid2.SetFocus
Exit Function
cmd6711212_err:
MsgBox "Aviso en proceso proforma " + error$, 48, "Aviso"
mytablex.Close
Exit Function
End Function

Function verifica_balanza(buf As String) As String
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.EOF = True Or mytablex.BOF = True Then
      mytablex.Close
      Exit Function
   End If
   verifica_balanza = "" & mytablex.Fields("peso")
   mytablex.Close
 

End Function
Function puerto_balanza1() As String
On Error GoTo cmd6712_err
Dim i As Long
Dim d As Integer
Dim buffers As String

If "" & mytable11.Fields("tipo_balanza") = "1" Then
   puerto_balanza1 = acura_lectura()
   Exit Function
End If
    Select Case "" & mytable11.Fields("portbala")
           Case "COM1"
           d = 1
           Case "COM2"
           d = 2
           Case "COM3"
           d = 3
           Case "COM4"
           d = 4
           Case "COM5"
           d = 5
           
End Select

buffers = ""



MSComm1.CommPort = d
MSComm1.Settings = "9600,n,8,1"
MSComm1.InputLen = 10
MSComm1.PortOpen = True
MSComm1.Output = Chr$(80)

'For i = 1 To 9000
'Next i
i = 0
Do
'DoEvents
buffers = buffers & MSComm1.Input
i = i + 1
If i > 15000 Then
   Exit Do
End If
Loop Until Len(buffers) >= 10
cerrar_balanza
puerto_balanza1 = buffers
Exit Function
cmd6712_err:
cerrar_balanza
Exit Function
End Function
Sub cerrar_balanza()
On Error GoTo cmd892_err
MSComm1.PortOpen = False
Exit Sub
cmd892_err:
Exit Sub
End Sub
Function busca_unidad(buf As String)

Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM producto where  producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe

   If "" & mytablex.Fields("vtaund") = "S" Then
      busca_unidad = 1
   End If
   
End If
'------------------------------------- ------------
mytablex.Close
 

End Function
Sub visualiza_detalle_factura(xtipo As String, xserie As String, xnumero As String)
Dim buf As String
Dim afgodetalle As String
Dim fgodetalle As String
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd344_err
afgodetalle = godetalle
fgodetalle = godetalle
dbgrid6.Visible = True
If opcion1 = "1900" Then  'proformas
   fgodetalle = "dproform"
End If
buf = "select Producto,Descripcio,Unidad,Factor,Cantidad as Cant,Precio,Total from " & fgodetalle & " where local='" & "" & mytable11.Fields("local") & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'"
mytablex.Open buf, cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   MsgBox "No existe detalle ", 48, "Aviso"
   mytablex.Close
   Exit Sub
End If
Set dbgrid6.DataSource = mytablex

               'DBGrid6.Refresh
               dbgrid6.columns(0).Width = 1200
               dbgrid6.columns(1).Width = 4500
               dbgrid6.SetFocus
Exit Sub
godetalle = afgodetalle
cmd344_err:
MsgBox "Error en select visualiza Detalle " & error$, 48, "Aviso"
Exit Sub
End Sub
Function verifica_oferta(buf As String) As String

Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT producto,remate FROM producto where  producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   verifica_oferta = "" & mytablex.Fields("remate")
End If
'------------------------------------- ------------
mytablex.Close
 
End Function
Function valida_tipo_pago(buf As String)
      Select Case buf
             Case "1"
                If "" & mytable11.Fields("ftb") <> "S" Then
                   Exit Function
                End If
             Case "2"
                If "" & mytable11.Fields("ftf") <> "S" Then
                   Exit Function
                End If
             Case "3"
                If "" & mytable11.Fields("fbm") <> "S" Then
                   Exit Function
                End If
             Case "4"
                If "" & mytable11.Fields("ffm") <> "S" Then
                   Exit Function
                End If
             Case "5"
                If "" & mytable11.Fields("fnv") <> "S" Then
                   Exit Function
                End If
             Case "7"
                If "" & mytable11.Fields("fexo") <> "S" Then
                   Exit Function
                End If
             Case "P"  'DE PEDIDOS
                If "" & mytable11.Fields("fnv") <> "S" Then
                   Exit Function
                End If
             
             Case Else
                Exit Function
     End Select
valida_tipo_pago = 1
End Function
Function redondeo1(buf3 As String) As String
Dim buf0 As String
Dim buf1 As String
Dim buf2 As String
Dim sdx As Double
Dim buf As String
buf = buf3
buf = Format(Val(buf), nrodecimal)
buf0 = Mid$(buf, 1, Len(buf) - 3)
buf1 = Mid$(buf, Len(buf) - 1, 2)
buf2 = ""
If Val(buf1) <= 0 Then
   redondeo1 = buf3
End If
If Val(Mid$(buf1, 1, 1)) = 9 And Val(Mid$(buf1, 2, 1)) >= 5 And Val(Mid$(buf1, 2, 1)) <= 9 Then
      sdx = Val(Mid$(buf0, 1, 1)) + 1
      buf2 = Format(sdx, "0")
      buf = buf2
      buf = Format(Val(buf), nrodecimal)
      redondeo1 = buf
   Exit Function
End If
If Val(Mid$(buf1, 2, 1)) >= 1 And Val(Mid$(buf1, 2, 1)) <= 4 Then
   'buf2 = Mid$(buf1, 1, 1) & "5"
   buf2 = Mid$(buf1, 1, 1) & "0"
   buf = buf0 + "." + buf2
   buf = Format(Val(buf), nrodecimal)
   redondeo1 = buf
   Exit Function
End If
If Val(Mid$(buf1, 2, 1)) >= 5 And Val(Mid$(buf1, 2, 1)) <= 9 Then
   sdx = Val(Mid$(buf1, 1, 1)) + 1
   buf2 = Format(sdx, "0")
   buf = buf0 & "." & buf2
   buf = Format(Val(buf), nrodecimal)
   redondeo1 = buf
   Exit Function
End If
redondeo1 = buf3
End Function
Function borrar_proformas()
On Error GoTo cmd89900_err
cn.Execute ("delete from cproform where local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & protipo & "' and serie='" & proserie & "' and numero='" & pronumero & "'")
cn.Execute ("delete from dproform where local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & protipo & "' and serie='" & proserie & "' and numero='" & pronumero & "'")
cn.Execute ("delete from ppocket where pedido='" & pronumero & "'")
Exit Function
cmd89900_err:
MsgBox "Aviso en borrar proformas " + error$, 48, "Aviso"
Exit Function
End Function
Function borrar_pedidos()
Dim xbuf As String
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd1289900_err
If Len(petipo) = 0 Or Len(peserie) = 0 Or Len(penumero) = 0 Then
   Exit Function
End If
xbuf = ""
mytablex.Open "SELECT * FROM tipo where tipo='" & petipo & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   mytablex.Close
   Exit Function
End If
Select Case "" & mytablex.Fields("tipodoc")
       Case "H"
           xbuf = "ccotizav"
       Case "I"
           xbuf = "cpedidov"
       Case "T"
           xbuf = "factura"
       
       
 End Select
 mytablex.Close
       
 If Len(xbuf) = 0 Then Exit Function
 
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM " & xbuf & " where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & petipo & "' and serie='" & peserie & "' and numero='" & penumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   'mytablex.Edit
   mytablex.Fields("tipo1") = petipo
   mytablex.Fields("serie1") = peserie
   mytablex.Fields("numero1") = penumero
   mytablex.Fields("yausado") = "1"
   'mytablex.Fields("acuenta") = Val("" & mytablex.Fields("total"))
   mytablex.Update
End If
mytablex.Close
Exit Function
cmd1289900_err:
MsgBox "Aviso en borrar pedidos", 48, "Aviso"
mytablex.Close
Exit Function

End Function

Function borrar_cotizacion()
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd1289900_err
If Len(petipo) = 0 Or Len(peserie) = 0 Or Len(penumero) = 0 Then
   Exit Function
End If
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM ccotizav where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & petipo & "' and serie='" & peserie & "' and numero='" & penumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   'mytablex.Edit
   mytablex.Fields("tipo1") = petipo
   mytablex.Fields("serie1") = peserie
   mytablex.Fields("numero1") = penumero
   mytablex.Fields("yausado") = "1"
   mytablex.Fields("acuenta") = Val("" & mytablex.Fields("total"))
   mytablex.Update
End If
mytablex.Close
Exit Function
cmd1289900_err:
MsgBox "Aviso en borrar Cotizacion", 48, "Aviso"
mytablex.Close
Exit Function

End Function

Sub pone_precios(buf As String)
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd89121_err
label56 = ""
'MsgBox buf
   mytablex.Open "SELECT * FROM precios where producto='" & buf & "' and local='" & Trim("" & mytable11.Fields("listap")) & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      label56 = "Pv1:" & mytablex.Fields("Unidad1") & " " & Format(Val("" & mytablex.Fields("pventa1")), nrodecimal)
      label56 = label56 + "  Pv2:" & mytablex.Fields("Unidad2") & " " & Format(Val("" & mytablex.Fields("pventa2")), nrodecimal)
      label56 = label56 + "  Pv3:" & mytablex.Fields("Unidad3") & " " & Format(Val("" & mytablex.Fields("pventa3")), nrodecimal)
      label56 = label56 + "  Pv4:" & mytablex.Fields("Unidad4") & " " & Format(Val("" & mytablex.Fields("pventa4")), nrodecimal)
      label56 = label56 + "  Pv5:" & mytablex.Fields("Unidad5") & " " & Format(Val("" & mytablex.Fields("pventa5")), nrodecimal)
      label56 = label56 + "  Pv6:" & mytablex.Fields("Unidad6") & " " & Format(Val("" & mytablex.Fields("pventa6")), nrodecimal)
      label56 = label56 + "  Pv7:" & mytablex.Fields("Unidad7") & " " & Format(Val("" & mytablex.Fields("pventa7")), nrodecimal)
      label56 = label56 + "  Pv8:" & mytablex.Fields("Unidad8") & " " & Format(Val("" & mytablex.Fields("pventa8")), nrodecimal)
      label56 = label56 + "  Pv9:" & mytablex.Fields("Unidad9") & " " & Format(Val("" & mytablex.Fields("pventa9")), nrodecimal)
      label56 = label56 + "  Pv10:" & mytablex.Fields("Unidad10") & " " & Format(Val("" & mytablex.Fields("pventa10")), nrodecimal)
   End If
   mytablex.Close
   mytablex.Open "SELECT * FROM almacen where local='" & Trim("" & mytable11.Fields("local")) & "' and producto='" & "" & buf & "' and bodega='" & Trim("" & mytable11.Fields("bodega")) & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      label56 = label56 + Chr$("impuesto") + Chr$("neto")
      label56 = label56 + " Saldo:" & Trim("" & rcconsulta.Fields("unidad")) & " " & calcula_saldo(Val("" & mytablex.Fields("saldo")), Val("" & rcconsulta.Fields("factor")))
   End If
   mytablex.Close
Exit Sub
cmd89121_err:
'MsgBox "Aviso en pone_precios " + error$, 48, "Aviso"
Exit Sub
End Sub
Function consulta_saldo(buf As String, cant As Double, sw As Integer) As Double
Dim mytablex As New ADODB.Recordset
Dim found As Integer
   'Combo1.Clear
   'Combo1.AddItem "bodega"
   'Combo1.ListIndex = 0
   
   'vemos si existen saldo en receta
   'MsgBox "cde"
'AQUI DEBE VERIFICAR SI EXISTE PRODUCTO
If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "SELECT * FROM almacen where producto='" & buf & "' and local='" & Trim("" & mytable11.Fields("local")) & "' and bodega='" & Trim("" & mytable11.Fields("bodega")) & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.EOF = True Or mytablex.BOF = True Then
      mytablex.Close
      Exit Function
   End If
   consulta_saldo = 0.1
   If sw = 0 Then
      consulta_saldo = Val("" & mytablex.Fields("saldo"))
   End If
   If sw = 1 Then
      'MsgBox cant
      If cant > Val("" & mytablex.Fields("saldo")) Then
       consulta_saldo = 0
      End If
   End If
   mytablex.Close
End Function
Sub consulta_minimo(buf As String, buf1 As String)
Dim mytablex As New ADODB.Recordset
Dim found As Integer
   stkminimo = ""
   mytablex.Open "SELECT * FROM almacen where producto='" & buf & "' and local='" & Trim("" & mytable11.Fields("local")) & "' and bodega='" & Trim("" & mytable11.Fields("bodega")) & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.Close
      Exit Sub
   End If
   If Val("" & mytablex.Fields("saldo")) <= Val(buf1) Then
      stkminimo = "Stock en el minimo :" & "" & mytablex.Fields("saldo")
   End If
   mytablex.Close
End Sub

Sub imprime_precuenta()
Dim found As Integer
Dim sFile As String
'impresora por default atachado
On Error GoTo cmd90000_err
cerrar_archivo
FileName = globaldir & "\temporal\" & gusuario & ".txt"
cerrar_archivo
found = borra_nombre("" & FileName)
Open FileName For Append As #1
    '------------------------------------
    cabecera_estado_cuenta
    cuerpo_estado_cuenta
    '------------------------------------
    cerrar_archivo
    Close #1
    If Len(Trim("" & mytable11.Fields("ecpuerto"))) = 0 Then
       MsgBox "Puerto de Precuenta no configurado", 48, "Aviso"
       Exit Sub
    End If
    If Trim("" & mytable11.Fields("eccola")) = "S" Then
       sFile = globaldir & "\temporal\" & gusuario & ".txt"
       found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"))
       Exit Sub
    End If
    If Trim("" & mytable11.Fields("eccola")) <> "S" Then
      found = star_sp342(Trim("" & mytable11.Fields("ecpuerto")), 0)
      found = corte_papel(Trim("" & mytable11.Fields("ecpuerto")), Val("" & mytable11.Fields("catipo")))
      Exit Sub
    End If
    Exit Sub
cmd90000_err:
    MsgBox "Error en imprime precuenta" + error$, 48, "Aviso"
    Exit Sub
    Exit Sub
    
    'genver.File = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1


End Sub
Sub cuerpo_estado_cuenta()
Dim buf As String
Dim found As Integer
Dim i As Integer
On Error GoTo cmd3999_err
    suma1 = 0
    Data2.refresh
    Do
      If Data2.Recordset.EOF Then Exit Do
       imprime_estado_cuenta
       Data2.Recordset.MoveNext
    Loop
       buf = "    NroUnidades "
       found = formateaa(buf, 20, 0, 0)
       found = formateaa("", 1, 0, 0)
       buf = Str(suma1)
       buf = Format(Val(buf), nrodecimal)
       found = formateaa(buf, 7, 2, 1)
    buf = "****TOTAL       "
    found = formateaa(buf, 22, 0, 0)
    found = formateaa("", 1, 0, 0)
    found = formateaa(Trim("" & mytable11.Fields("moneda")), 3, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = Format(Val(txtotal), nrodecimal)
    found = formateaa(buf, 9, 2, 1)
    For i = 1 To 11
        found = formateaa("", 1, 2, 0)
    Next i
    DBGrid2.SetFocus
    Exit Sub
cmd3999_err:
MsgBox "Error en cuerpo estado cuenta " + error$, 48, "Aviso"
Exit Sub
End Sub
Sub cabecera_estado_cuenta()
Dim found As Integer
Dim buf As String
Dim btipo As String
On Error GoTo cmd4111_err
   buf = String(45, "-")
   found = formateaa(buf, 45, 2, 0)
   buf = "       ESTADO DE CUENTA"
   found = formateaa(buf, 36, 2, 0)
   buf = "    Cajero:" & cajero & " Caja:" & caja & " Turno:" & turno
   found = formateaa(buf, 36, 2, 0)
   buf = "  Fecha:" & Format(Now, "dd/mm/yyyy") & "  Hora:" & Format(Now, "hh:mm:ss")
   found = formateaa(buf, 36, 2, 0)
   buf = servicio_generado("" & flag_servicio)
   'If flag_servicio <> "C" And flag_servicio <> "D" Then
   '   buf = servicio_generado("" & flag_servicio)
   '   found = formateaa(buf, 25, 2, 0)
   'End If
   'If FLAG_SERVICIO = "C" Then
   '   buf = "   Salon : " & salon & " Mesa:" & mesa
   '   found = formateaa(buf, 36, 2, 0)
   'End If
   'If flag_servicio = "D" Then
   found = formateaa(buf, 36, 2, 0)
      'imprime_cliente_delivery "" & codigocli
   'End If
   buf = String(45, "-")
   found = formateaa(buf, 45, 2, 0)
Exit Sub
cmd4111_err:
  MsgBox "Mensaje,Error en cabecera Pedido " & error$
  Exit Sub

End Sub
Sub imprime_estado_cuenta()
Dim buf As String
Dim found As Integer
On Error GoTo cmd45888_err
    buf = "" & Data2.Recordset.Fields("producto")
    found = formateaa(buf, 13, 0, 0)
    found = formateaa(" ", 1, 0, 0)
    buf = "" & Data2.Recordset.Fields("unidad")
    found = formateaa(buf, 3, 2, 0)

    buf = Mid$("" & Data2.Recordset.Fields("descripcio"), 1, 20)
    found = formateaa(buf, 20, 0, 0)
    found = formateaa(" ", 1, 0, 0)

    buf = "" & Data2.Recordset.Fields("cantidad")
    'buf = Format(Val(buf), nrodecimal)
    found = formateaa(buf, 7, 0, 1)
    found = formateaa(" ", 1, 0, 0)

    buf = "" & Data2.Recordset.Fields("total")
    buf = Format(Val(buf), nrodecimal)
    found = formateaa(buf, 7, 2, 1)
    suma1 = suma1 + Val("" & Data2.Recordset.Fields("cantidad"))
Exit Sub
cmd45888_err:
MsgBox "Error en imprime estado de cuenta " + error$, 48, "Aviso"
Exit Sub
End Sub
Sub sql_saldo_locales(buf As String)
Dim rxconsulta As New ADODB.Recordset
On Error GoTo cmd87678_err
'buf = "select * from almacen where producto='" & buf & "'"
buf = "select Almacen.saldo,Bodega.nombre,almacen.bodega,Almacen.local from almacen left join bodega on almacen.bodega=bodega.codigo where almacen.producto='" & buf & "' order by almacen.bodega"
 If rxconsulta.State = 1 Then rxconsulta.Close
   rxconsulta.Open buf, cn, adOpenStatic, adLockOptimistic
   'If rxconsulta.EOF = True And rxconsulta.BOF = True Then
   '   rxconsulta.Close
   '   Exit Sub
   'End If
   Set dbgrid7.DataSource = rxconsulta
   Exit Sub
cmd87678_err:
               MsgBox "Aviso en sql-saldo local " + error$, 48, "Aviso"
               Exit Sub
               

End Sub
Sub limpia_general()
Frame7.Visible = False
Framefp.Visible = False
habilita_lab7 0
'If flag_servicio = "C" Then
'If cmytablex.RecordCount > 0 Then
'   cn.Execute ("delete from dcomanda where salon='" & cmytablex.Fields("salon") & "' and mesa='" & cmytablex.Fields("mesa") & "'")
'End If
'End If
consulta_comanda "" & mytable11.Fields("salon")
borrar_todo
sql_detalle
tiposervicio1 = "Autoservicio"
salon = ""
mesa = ""
mesero = ""
cuenta_separa = ""
flag_servicio = "A"
'Frame10.Visible = True
End Sub
Sub proceso_cierre_automatico()
Dim found As Integer
Dim buf As String


If Frame2.Visible = True Then Exit Sub
local1.Visible = False
local1.Visible = False
found = sumar_detalle()
If found = 0 Then
   MsgBox "debe de Existir un Precio=0", 48, "Aviso"
   DBGrid2.SetFocus
   Exit Sub
End If

If Val(txtotal) = 0 Then
   If exisdev <> -10 Then  'si existe devolucion
      DBGrid2.SetFocus
      Exit Sub
   End If
End If
carga_tiposdoc "%"
If Trim("" & mytable11.Fields("terminal")) = "T" Or (Val(acuenta) > 0 And Len(petipo) = 0) Then 'pedidos o acuenta>0
          'MsgBox "Hola"
          xruc = codigo
          xnombre = nombre
          Frame7.Visible = True
          habilita_lab7 1
          Framefp.Enabled = False
          If Val(acuenta) > 0 Then
             xtipo = Trim("" & mytable11.Fields("tipope"))
          End If
          xtipo.SetFocus
          Exit Sub
End If
If flag_servicio = "A" Then  'venta rapida
End If
If flag_servicio = "D" Then  'validar el deliveri si ingreso datos
End If
If flag_servicio = "C" Then  'venta mesas
End If
Label36.Caption = "Codigo"
'Frame10.Visible = False
found = proceso_cobros()  'PONE EN CERO TODAS LA FORMAS DE PAGO
opcion2 = 0
'MsgBox dbgrid10.Visible
ttxtotals = Format(Val(rtxtotal), nrodecimal)
ttxtotald = Format(Val(rtxtotald), nrodecimal)
stxtotals = Format(Val(rtxtotal), nrodecimal)
stxtotald = Format(Val(rtxtotald), nrodecimal)
Framefp.Visible = True
Framefp.Enabled = True
habilita_lab7 0
'MsgBox ""
'MsgBox dbgrid10.Enabled
buf = "select * from fpago where fpago='1'"
If mytablefpago.State = 1 Then mytablefpago.Close
mytablefpago.Open buf, cn, adOpenDynamic, adLockOptimistic
Set dbgrid10.DataSource = mytablefpago
dbgrid10.refresh
   If mytablefpago.RecordCount > 0 Then
      mytablefpago.MoveFirst
      dbgrid10.Enabled = False
      dbgrid10.Visible = True
      dbgrid10.Enabled = True
      dbgrid10.SetFocus
      DBGrid10_KeyDown 13, 0
      DBGrid9.Enabled = True
      'Exit Sub
      DBGrid9.SetFocus
      DBGrid9_KeyDown 13, 0
      RGPAGO_KeyPress 13
      'RGPAGO.SetFocus
      'xtipo = "7"
      'Else
      'MsgBox "No existe exonerado ", 48, "Aviso"
   End If
   'mytablex.Close
End Sub
Sub proceso_cierre_pedido()
Dim found As Integer
Dim buf As String
If Frame2.Visible = True Then Exit Sub
local1.Visible = False
local1.Visible = False
found = sumar_detalle()
If found = 0 Then
   MsgBox "debe de Existir un Precio=0", 48, "Aviso"
   DBGrid2.SetFocus
   Exit Sub
End If

If Val(txtotal) = 0 Then
   If exisdev <> -10 Then  'si existe devolucion
      DBGrid2.SetFocus
      Exit Sub
   End If
End If
carga_tiposdoc "I"
If Trim("" & mytable11.Fields("terminal")) = "T" Or (Val(acuenta) > 0 And Len(petipo) = 0) Then 'pedidos o acuenta>0
          'MsgBox "Hola"
          xruc = codigo
          xnombre = nombre
          Frame7.Visible = True
          habilita_lab7 1
          Framefp.Enabled = False
          If Val(acuenta) > 0 Then
             xtipo = Trim("" & mytable11.Fields("tipope"))
          End If
          xtipo.SetFocus
          Exit Sub
End If
If flag_servicio = "A" Then  'venta rapida
End If
If flag_servicio = "D" Then  'validar el deliveri si ingreso datos
End If
If flag_servicio = "C" Then  'venta mesas
End If
Label36.Caption = "Codigo"
'Frame10.Visible = False
found = proceso_cobros()  'PONE EN CERO TODAS LA FORMAS DE PAGO
opcion2 = 0
'MsgBox dbgrid10.Visible
ttxtotals = Format(Val(rtxtotal), nrodecimal)
ttxtotald = Format(Val(rtxtotald), nrodecimal)
stxtotals = Format(Val(rtxtotal), nrodecimal)
stxtotald = Format(Val(rtxtotald), nrodecimal)
Framefp.Visible = True
Framefp.Enabled = True
habilita_lab7 0
'MsgBox ""
'MsgBox dbgrid10.Enabled
buf = "select * from fpago where fpago='1'"
If mytablefpago.State = 1 Then mytablefpago.Close
mytablefpago.Open buf, cn, adOpenDynamic, adLockOptimistic
Set dbgrid10.DataSource = mytablefpago
dbgrid10.refresh
   If mytablefpago.RecordCount > 0 Then
      mytablefpago.MoveFirst
      dbgrid10.Enabled = False
      dbgrid10.Visible = True
      dbgrid10.Enabled = True
      dbgrid10.SetFocus
      DBGrid10_KeyDown 13, 0
      DBGrid9.Enabled = True
      'Exit Sub
      DBGrid9.SetFocus
      DBGrid9_KeyDown 13, 0
      RGPAGO_KeyPress 13
      'RGPAGO.SetFocus
      'xtipo = "7"
      'Else
      'MsgBox "No existe exonerado ", 48, "Aviso"
   End If
   'mytablex.Close

End Sub
Sub menu_graba_fpedido()
Dim found As Integer
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM fpagov ", cn, adOpenDynamic, adLockOptimistic
graba_fpago_pedido mytablex
found = graba_credito_trabajo() 'RECIEN LO DESHABILITE
'found = pone_recibo_caja()
mytablex.Close
End Sub
Sub graba_fpago_pedido(mytabley As ADODB.Recordset)
   mytabley.AddNew
   mytabley.Fields("paridad") = Val("" & paridadfp)
   mytabley.Fields("codigo") = "" & xruc
   mytabley.Fields("nombre") = "" & xnombre
   mytabley.Fields("tipo") = xtipo
   mytabley.Fields("serie") = xserie
   mytabley.Fields("numero") = xnumero
   mytabley.Fields("tipoclie") = "C"
   mytabley.Fields("fecha") = Format("" & dia, "dd/mm/yyyy")
   mytabley.Fields("moneda") = Trim("" & mytable11.Fields("moneda"))
   mytabley.Fields("total") = Val(acuenta)
   
   mytabley.Fields("caja") = "" & caja
   mytabley.Fields("turno") = "" & turno
   mytabley.Fields("usuario") = "" & cajero
   
   mytabley.Fields("total") = Val(acuenta)
   mytabley.Fields("cambio") = 0
   mytabley.Fields("recibe") = Val(acuenta)
   mytabley.Fields("recibes") = 0
   mytabley.Fields("recibed") = 0
   mytabley.Fields("saldos") = 0
   mytabley.Fields("saldod") = 0
   mytabley.Fields("orden") = ""
   mytabley.Fields("observa") = ""
   mytabley.Fields("dias") = ""
   mytabley.Fields("fpago") = "1"
   mytabley.Fields("acufp") = "A" 'acu de recibo ingreso por
   mytabley.Fields("descripcio") = "EFECTIVO"
   mytabley.Fields("acu") = "I"
   mytabley.Fields("local") = Trim("" & mytable11.Fields("local"))
   mytabley.Fields("servicio") = flag_servicio
   If flag_servicio = "A" Then
      mytabley.Fields("servicio") = "A"
   End If
   If flag_servicio = "D" Then
      mytabley.Fields("servicio") = "D"
   End If
   If flag_servicio = "C" Then
      mytabley.Fields("servicio") = "C"
   End If
   mytabley.Fields("estado") = "2"
   mytabley.Update
   'la diferencia al credito
   mytabley.AddNew
   mytabley.Fields("paridad") = Val("" & paridadfp)
   mytabley.Fields("codigo") = "" & xruc
   mytabley.Fields("nombre") = "" & xnombre
   mytabley.Fields("tipo") = xtipo
   mytabley.Fields("serie") = xserie
   mytabley.Fields("numero") = xnumero
   mytabley.Fields("tipoclie") = "C"
   mytabley.Fields("fecha") = Format("" & dia, "dd/mm/yyyy")
   mytabley.Fields("moneda") = Trim("" & mytable11.Fields("moneda"))
   mytabley.Fields("total") = Val(acuenta)
   
   mytabley.Fields("caja") = "" & caja
   mytabley.Fields("turno") = "" & turno
   mytabley.Fields("usuario") = "" & cajero
   
   mytabley.Fields("total") = Val(txtotal) - Val(acuenta)
   mytabley.Fields("cambio") = 0
   mytabley.Fields("recibe") = Val(txtotal) - Val(acuenta)
   mytabley.Fields("recibes") = 0
   mytabley.Fields("recibed") = 0
   mytabley.Fields("saldos") = 0
   mytabley.Fields("saldod") = 0
   mytabley.Fields("orden") = ""
   mytabley.Fields("observa") = ""
   mytabley.Fields("dias") = "1"
   mytabley.Fields("fpago") = "6" 'ojo debe existir este dato de credito formpago
   mytabley.Fields("acufp") = "J" 'acu de recibo ingreso por
   mytabley.Fields("descripcio") = "ORDENTRABAJO"
   mytabley.Fields("acu") = "I"
   mytabley.Fields("local") = Trim("" & mytable11.Fields("local"))
   mytabley.Fields("servicio") = flag_servicio
   If flag_servicio = "A" Then
      mytabley.Fields("servicio") = "A"
   End If
   If flag_servicio = "D" Then
      mytabley.Fields("servicio") = "D"
   End If
   If flag_servicio = "C" Then
      mytabley.Fields("servicio") = "C"
   End If
   mytabley.Fields("estado") = "2"
   mytabley.Fields("flage") = "I"
   mytabley.Update
End Sub
Sub grabar_descto()
On Error GoTo cmd6543_err
Dim found As Integer
Dim a As Double
            Data2.refresh
            Do
               If Data2.Recordset.EOF Then Exit Do
               If (Val("" & Data2.Recordset.Fields("cantidad")) > 0 Or Val("" & Data2.Recordset.Fields("cantidad")) < 0) And Val("" & Data2.Recordset.Fields("precio")) > 0 Then
                     Data2.Recordset.Edit
                     'MsgBox tipodescuento
                     If tipodescuento = "2" Then
                        Data2.Recordset.Fields("deslipo") = 0
                     End If
                     If tipodescuento = "0" Then
                        Data2.Recordset.Fields("deslipo") = Val(valordescuento)
                     End If
                     If tipodescuento = "1" Then
                        a = (Val(valordescuento) * 100) / Val(txtotal)
                        Data2.Recordset.Fields("deslipo") = a
                     End If
                     If tipodescuento = "3" Then   '----recargos
                        Data2.Recordset.Fields("deslipo") = 0
                        Data2.Recordset.Fields("precio") = Val("" & Data2.Recordset.Fields("precio")) + Val("" & Data2.Recordset.Fields("precio")) * valordescuento / 100
                        Data2.Recordset.Fields("TOTAL") = Val("" & Data2.Recordset.Fields("precio")) * Val("" & Data2.Recordset.Fields("cantidad"))
                     End If
                     suma_linea
                     Data2.Recordset.Update
               End If
               Data2.Recordset.MoveNext
            Loop
            'sql_detalle
            found = sumar_detalle()
               DBGrid2.Col = 0
               DBGrid2.Row = DBGrid2.VisibleRows - 1
               DBGrid2.SetFocus

            Exit Sub
cmd6543_err:
MsgBox "Aviso " + error$, 48, "Aviso"
Exit Sub

End Sub
Sub suma_linea()
resuma_precios

End Sub

Function graba_credito_trabajo()
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM cuentac where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "' and cuota='1'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then  'si existe
   mytablex.Close
   Exit Function
End If
   mytablex.AddNew
   mytablex.Fields("OBSERVA") = "ADEL.ORDENTRA"
   mytablex.Fields("GRUPO") = "O"
   mytablex.Fields("fpago") = "A"
   mytablex.Fields("acu") = "I"
   mytablex.Fields("tipo") = xtipo
   mytablex.Fields("serie") = xserie
   mytablex.Fields("numero") = xnumero
   mytablex.Fields("dias") = 1
   mytablex.Fields("cuota") = "1"
   mytablex.Fields("tipoclie") = "C"
   mytablex.Fields("codigo") = xruc
   mytablex.Fields("nombre") = xnombre
   mytablex.Fields("fecha") = Format(dia, "dd/mm/yyyy")
   mytablex.Fields("fechav") = Format(dia, "dd/mm/yyyy")
   mytablex.Fields("moneda") = Trim("" & mytable11.Fields("moneda"))
   mytablex.Fields("total") = Val(txtotal)
   mytablex.Fields("abono") = Val(acuenta)
   mytablex.Fields("interes") = 0
   mytablex.Fields("saldo") = Val(txtotal) - Val(acuenta)
   mytablex.Fields("estado") = "0"
   mytablex.Fields("vendedor") = "" & xvendedor
   mytablex.Fields("zona") = ""
   mytablex.Fields("local") = Trim("" & mytable11.Fields("local"))
   mytablex.Fields("caja") = "" & caja
   mytablex.Fields("turno") = "" & turno
   mytablex.Fields("usuario") = "" & cajero
   mytablex.Update
mytablex.Close
End Function
Function descuenta_credito_pedido()
On Error GoTo cmd65u_err
Dim mytabley As New ADODB.Recordset
Dim mytablex As New ADODB.Recordset
Dim sdx As Double



   'ADICIONAR EL PAGO
   mytabley.Open "SELECT * FROM cuentacd where local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & petipo & "' and serie='" & peserie & "' and numero='" & penumero & "' and cuota='1'", cn, adOpenStatic, adLockOptimistic
   mytabley.AddNew

mytabley.Fields("codigo") = "" & xruc
mytabley.Fields("local") = Trim("" & mytable11.Fields("local"))
mytabley.Fields("local1") = Trim("" & mytable11.Fields("local"))
mytabley.Fields("tipo") = xtipo
mytabley.Fields("serie") = xserie

mytabley.Fields("numero") = xnumero
mytabley.Fields("acu") = ""
mytabley.Fields("tipo1") = petipo
mytabley.Fields("serie1") = peserie
mytabley.Fields("numero1") = penumero
mytabley.Fields("cuota") = "1" '& mytablex.Fields("cuota")
mytabley.Fields("moneda") = moneda
mytabley.Fields("total") = Val(acuenta)
mytabley.Fields("paga") = Val(acuenta)
mytabley.Fields("estado") = "2"
mytabley.Fields("fecha") = CVDate(dia)

'mytabley.Fields("hora") = "" & mytablez.Fields("hora")
mytabley.Fields("usuario") = cajero
mytabley.Fields("caja") = caja
mytabley.Fields("turno") = turno

mytabley.Fields("tipoclie") = "C"
mytabley.Update
mytabley.Close


   mytablex.Open "SELECT * FROM cuentac where local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & petipo & "' and serie='" & peserie & "' and numero='" & penumero & "' and cuota='1'", cn, adOpenStatic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.Close
      Exit Function
   End If
   sdx = Val("" & mytablex.Fields("abono")) + (Val(txtotal) - Val(acuenta))
   mytablex.Fields("abono") = sdx
   mytablex.Fields("saldo") = (Val("" & mytablex.Fields("total")) - Val("" & mytablex.Fields("abono")))
   mytablex.Update
   mytablex.Close
   Exit Function
cmd65u_err:
   MsgBox "Aviso en descuento credito pedido " + error, 48, "Aviso"
   Exit Function
   
End Function
Function pone_recibo_caja()
On Error GoTo cmd891212_err
Dim mytablex As New ADODB.Recordset

If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM recibo where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount = 0 Then 'si existe
   mytablex.AddNew
   mytablex.Fields("local") = Trim("" & mytable11.Fields("local"))
   mytablex.Fields("caja") = "" & caja
   mytablex.Fields("turno") = "" & turno
   mytablex.Fields("usuario") = "" & cajero
  
   mytablex.Fields("tipo") = xtipo
   mytablex.Fields("serie") = xserie
   mytablex.Fields("numero") = xnumero
   

mytablex.Fields("afecta") = "C"
mytablex.Fields("fecha") = Format(dia, "dd/mm/yyyy")
mytablex.Fields("hora") = Format(Now, "hh:mm")
mytablex.Fields("tipoclie") = "C"
mytablex.Fields("codigo") = xruc
mytablex.Fields("nombre") = Trim(Mid$(nombre, 1, 60))
'mytablex.Fields("observa") = observa
mytablex.Fields("vendedor") = xvendedor
mytablex.Fields("moneda") = Trim("" & mytable11.Fields("moneda"))
mytablex.Fields("paridad") = 2.8
mytablex.Fields("total") = Val(txtotal)
mytablex.Fields("estado") = "2"
mytablex.Fields("acu") = "W"
mytablex.Fields("servicio") = "W"
'mytablex.Fields("c1") = Val(c11)
'mytablex.Fields("c2") = Val(c12)
'mytablex.Fields("c3") = Val(c13)
'mytablex.Fields("c4") = Val(c14)
mytablex.Update
End If
mytablex.Close
Exit Function
cmd891212_err:
MsgBox "Aviso en Pone recibo caja " + error$, 48, "Aviso"
Exit Function

End Function
Function forma_pago_adicional(mytabley As ADODB.Recordset)  'forma pago adicional orden pedido
 mytabley.Fields("paridad") = Val("" & paridadfp)
   mytabley.Fields("codigo") = "" & xruc
   mytabley.Fields("nombre") = "" & xnombre
   mytabley.Fields("tipo") = xtipo
   mytabley.Fields("serie") = xserie
   mytabley.Fields("numero") = xnumero
   mytabley.Fields("tipoclie") = "C"
   mytabley.Fields("fecha") = Format("" & dia, "dd/mm/yyyy")
   mytabley.Fields("moneda") = Trim("" & mytable11.Fields("moneda"))
   mytabley.Fields("total") = Val(acuenta)
   
   mytabley.Fields("caja") = "" & caja
   mytabley.Fields("turno") = "" & turno
   mytabley.Fields("usuario") = "" & cajero
   mytabley.Fields("total") = Val(acuenta)
   mytabley.Fields("cambio") = 0
   mytabley.Fields("recibe") = Val(acuenta)
   mytabley.Fields("recibes") = 0
   mytabley.Fields("recibed") = 0
   mytabley.Fields("saldos") = 0
   mytabley.Fields("saldod") = 0
   mytabley.Fields("orden") = ""
   mytabley.Fields("observa") = ""
   mytabley.Fields("dias") = ""
   mytabley.Fields("fpago") = "6"
   mytabley.Fields("acufp") = "J" 'acu de recibo ingreso por
   mytabley.Fields("descripcio") = "ORDENTRABAJO"
   mytabley.Fields("acu") = acu
   mytabley.Fields("local") = Trim("" & mytable11.Fields("local"))
   mytabley.Fields("servicio") = flag_servicio
   If flag_servicio = "A" Then
      mytabley.Fields("servicio") = "A"
   End If
   If flag_servicio = "D" Then
      mytabley.Fields("servicio") = "D"
   End If
   If flag_servicio = "C" Then
      mytabley.Fields("servicio") = "C"
   End If
   mytabley.Fields("estado") = "2"
   'If xxacu = "I" Then
   '   mytabley.Fields("acu") = xxacu
   'End If

End Function
Function verifica_producto(buf As String)
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM producto where  producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   verifica_producto = 1
End If
mytablex.Close
End Function
Function busca_tipo_lineas(buf As String)
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM tipo  where  tipo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   busca_tipo_lineas = Val("" & mytablex.Fields("nrolineas"))
   'MsgBox ""
End If
mytablex.Close

End Function
Function busca_tipo_acu(buf As String)
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM tipo  where  tipo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   busca_tipo_acu = "" & mytablex.Fields("tipodoc")
End If
mytablex.Close

End Function

Sub carga_foto(buf As String)
Dim fotonombre As String
On Error GoTo cmd4432_err
foto = LoadPicture()
fotonombre = buf
If Len(fotonombre) > 0 Then
If existe_archivo(fotonombre) > 0 Then
   foto = LoadPicture(fotonombre)
End If
End If
Exit Sub
cmd4432_err:
Exit Sub

End Sub
Sub palabra_bienvenida1()
Dim buf As String
Dim sdx As Double
Dim buf1 As String
Dim buf2 As String
On Error GoTo cmd3678112_err
Exit Sub
   sdx = Val(stxtotals)
   buf = Format(sdx, nrodecimal)
   buf1 = Mid$(buf, Len(buf) - 1, 2)
   buf = Mid$(buf, 1, Len(buf) - 3)
   buf = letras(buf, 40)
   buf = LTrim$(Trim$(buf))
   buf = UCase(buf)
   buf2 = LTrim(RTrim(buf)) & " con " & LTrim(RTrim(buf1))
'MsgBox buf2
'buf = Trim(pone_letras(stxtotals, "S", 60))
'MsgBox "" & ttxtotals
'Speech.Pitch = 170 ' Set Pitch Value
'Speech.Speed = 120 ' Set Speed Value
'Speech.AudioReset
'MsgBox "Hola"
'Speech.Speak "Su cuenta es  " & buf2 & " NUEVOS SOLES"
'Speech.Sayit = "son " + "" & ttxtotals + " SOLES "
'Sleep (5000)
Exit Sub
cmd3678112_err:
MsgBox "Error en palabra " & error$, 48, "Aviso"
Exit Sub
End Sub
Sub graba_tarjetas(mytabley As ADODB.Recordset)
Dim mytablex As New ADODB.Recordset
Dim sdx As Double
On Error GoTo cmd7811_err
sdx = busca_banco_numero()
busvf:
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM chequemo  where  transaccio='" & sdx & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   sdx = sdx + 1
   GoTo busvf
End If
mytablex.AddNew
mytablex.Fields("transaccio") = "" & sdx
mytablex.Fields("moneda") = "" & mytabley.Fields("moneda")
mytablex.Fields("tipoclie") = "C"
mytablex.Fields("codigo") = "" & mytabley.Fields("codigo")
mytablex.Fields("banco") = "BCP"
mytablex.Fields("cuenta") = ""
mytablex.Fields("tipo") = "72"
mytablex.Fields("numero") = ""
mytablex.Fields("fechan") = Format("" & mytabley.Fields("fecha"), "dd/mm/yyyy")
mytablex.Fields("fechae") = Format("" & mytabley.Fields("fecha"), "dd/mm/yyyy")
mytablex.Fields("nombre") = "" & mytabley.Fields("nombre")
mytablex.Fields("conciliado") = "N"
mytablex.Fields("concepto") = "" & mytabley.Fields("descripcio")
mytablex.Fields("acu") = "X"
mytablex.Fields("comenta") = ""
mytablex.Fields("total") = Val("" & mytabley.Fields("recibe"))
mytablex.Fields("descuento") = 0
mytablex.Fields("recargo") = 0
mytablex.Fields("abono") = 0
mytablex.Fields("neto") = Val("" & mytabley.Fields("recibe"))
mytablex.Fields("saldo") = Val("" & mytabley.Fields("recibe"))
mytablex.Fields("cajero") = "" & cajero
mytablex.Fields("caja") = "" & caja
mytablex.Fields("turno") = "" & turno
mytablex.Fields("xtipo") = "" & mytabley.Fields("tipo")
mytablex.Fields("xserie") = "" & mytabley.Fields("serie")
mytablex.Fields("xnumero") = "" & mytabley.Fields("numero")
mytablex.Update
mytablex.Close
Exit Sub
cmd7811_err:
MsgBox "Aviso en graba tarjetas " + error$, 48, "Aviso"
Exit Sub
End Sub
Function busca_banco_numero() As Double
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM parame where codigo='01'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   busca_banco_numero = Val("" & mytablex.Fields("banco"))
End If
mytablex.Close
End Function
Function graba_guia_mensual()
Dim buf As String
Dim i As Integer
Dim j As Integer
Dim AA As String
Dim BB As String
Dim cC As String
Dim dd As String
On Error GoTo cmd12004992_err
'MsgBox crucefa.ListCount
For i = 0 To crucefa.ListCount - 1
   extrae_crucefa crucefa.List(i), AA, BB, cC, dd
   buf = "update cuentac set estado='1'  where  local='" & "" & AA & "' and tipo='" & "" & BB & "' and serie='" & "" & cC & "' and  numero='" & "" & dd & "'"
   mydbxglo.Execute buf
Next i
   Exit Function
cmd12004992_err:
Exit Function
MsgBox "Aviso en graba_guia Mensual" + error$, 24, "AVISO DE NO ERROR"
Resume
End Function
Sub reversa_guia_mensual(axlocal As String, axtipo As String, axserie As String, axnumero As String)
Dim buf As String
buf = "update cuentac set estado='0'  where  local='" & axlocal & "' and tipo='" & axtipo & "' and serie='" & axserie & "' and  numero='" & axnumero & "'"
cn.Execute buf
End Sub
Sub extrae_crucefa(DATO As String, ccampo1 As String, ccampo2 As String, ccampo3 As String, ccampo4 As String)
Dim i As Integer
Dim j As Integer
Dim temp As String
i = 0
temp = Trim$(DATO)
If Len(temp) = 0 Then Exit Sub
Do
   j = InStr(temp, "|")
   If j > 0 Then
      i = i + 1
      Select Case i
             Case 1: ccampo1 = Trim(Mid$(temp, 1, j - 1))
             Case 2: ccampo2 = Trim(Mid$(temp, 1, j - 1))
             Case 3: ccampo3 = Trim(Mid$(temp, 1, j - 1))
             Case 4: ccampo4 = Trim(Mid$(temp, 1, j - 1))
             'Case 5: campo5 = Mid$(temp, 1, J - 1)
      End Select
      temp = Trim$(Mid$(temp, j + 1))
     Else
     Exit Sub
   End If
Loop
   Exit Sub

End Sub
Function verifica_ofertax(buf As String, xcant As Double, buf1 As Double)
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM precios  where  producto='" & buf & "' and local='" & Trim("" & mytable11.Fields("listap")) & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   If Val("" & mytablex.Fields("minimo11")) > 0 And Val("" & mytablex.Fields("maximo11")) > 0 Then
   If xcant >= Val("" & mytablex.Fields("minimo11")) And xcant <= Val("" & mytablex.Fields("maximo11")) Then
      buf1 = Val("" & mytablex.Fields("pventa11"))
      If Val(buf1) > 0 Then
         verifica_ofertax = 1
      End If
      mytablex.Close
      Exit Function
   End If
   End If
   If Val("" & mytablex.Fields("minimo12")) > 0 And Val("" & mytablex.Fields("maximo12")) > 0 Then
   If xcant >= Val("" & mytablex.Fields("minimo12")) And xcant <= Val("" & mytablex.Fields("maximo12")) Then
      buf1 = Val("" & mytablex.Fields("pventa12"))
      If Val(buf1) > 0 Then
         verifica_ofertax = 1
      End If
      mytablex.Close
      Exit Function
   End If
   End If
   If Val("" & mytablex.Fields("minimo13")) > 0 And Val("" & mytablex.Fields("maximo13")) > 0 Then
   If xcant >= Val("" & mytablex.Fields("minimo13")) And xcant <= Val("" & mytablex.Fields("maximo13")) Then
      buf1 = Val("" & mytablex.Fields("pventa13"))
      If Val(buf1) > 0 Then
      verifica_ofertax = 1
      End If
      mytablex.Close
      Exit Function
   End If
   End If
   If Val("" & mytablex.Fields("minimo14")) > 0 And Val("" & mytablex.Fields("maximo14")) > 0 Then
   If xcant >= Val("" & mytablex.Fields("minimo14")) And xcant <= Val("" & mytablex.Fields("maximo14")) Then
      buf1 = Val("" & mytablex.Fields("pventa14"))
      If Val(buf1) > 0 Then
      verifica_ofertax = 1
      End If
      mytablex.Close
      Exit Function
   End If
   End If
End If
mytablex.Close
End Function
Function busca_credito_adelanto(buf As String, buf2 As String)
Dim mytablex As New ADODB.Recordset
Dim sdx As Double
Dim found As Integer
Dim buf1 As String
saldoabo = ""
buf1 = buf
If buf = "I" Then
   buf1 = "A"
End If
If buf = "K" Then
   buf1 = "D"
End If

sdx = 0
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM cuentac  where  tipoclie='C' and codigo='" & buf2 & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   busca_credito_adelanto = 1
   Do
   If mytablex.EOF Then Exit Do
     If Val("" & mytablex.Fields("saldo")) > 0 Then
     If "" & mytablex.Fields("grupo") = buf1 Then
        sdx = sdx + Val("" & mytablex.Fields("saldo"))
     End If
     End If
   mytablex.MoveNext
   Loop
End If
mytablex.Close
saldoabo = Format(sdx, "0.00")
'If buf2 = "C" Then
'mytablex.Open "SELECT * FROM clientes where  codigo='" & buf2 & "'", cn, adOpenDynamic, adLockOptimistic
'If mytablex.RecordCount > 0 Then 'si existe
'   mytablex.Fields("credito_usado") = sdx
'End If
'mytablex.Close
'End If
End Function
Function busca_credito_adelanto1(buf1 As String, buf As String) As Double
Dim mytablex As New ADODB.Recordset
Dim sdx As Double
Dim buf2 As String
Dim found As Integer
If buf = "I" Then
   buf2 = "A"
End If
If buf = "K" Then
   buf2 = "D"
End If
sdx = 0
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM cuentac  where  tipoclie='C' and codigo='" & buf1 & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   Do
   If mytablex.EOF Then Exit Do
   If "" & mytablex.Fields("codigo") = buf1 Then
       If Val("" & mytablex.Fields("saldo")) > 0 Then
          If "" & mytablex.Fields("grupo") = buf2 Then
          sdx = sdx + Val("" & mytablex.Fields("saldo"))
          End If
       End If
     Else: Exit Do
   End If
   mytablex.MoveNext
   Loop
End If
mytablex.Close
busca_credito_adelanto1 = Val(Format(sdx, "0.00"))
End Function
Function busca_codigo_descuento(buf As String)
Dim mytablex As New ADODB.Recordset
Dim sdx As Double
trdescuento = ""
saldo = ""
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM clientes  where  codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   nombre = "" & mytablex.Fields("nombre")
   trdescuento = Format(Val("" & mytablex.Fields("descuento")), "0.00")
   saldo = Format(Val("" & mytablex.Fields("credito")), "0.00")
   busca_codigo_descuento = 1
End If
mytablex.Close
sdx = 0
saldo = ""
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM cuentac  where  tipoclie='C' and codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   Do
   If mytablex.EOF Then Exit Do
     If Val("" & mytablex.Fields("saldo")) > 0 Then
        sdx = sdx + Val("" & mytablex.Fields("saldo"))
     End If
   mytablex.MoveNext
   Loop
End If
mytablex.Close
saldo = Format(sdx, "0.00")
End Function
Function valida_otros()
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM factura  where  local='" & rrlocal11 & "' and tipo='" & rrtipo & "' and serie='" & rrserie & "' and numero='" & rrnumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   valida_otros = 1
End If
mytablex.Close
End Function
Function valida_rango()
Dim mytablex As New ADODB.Recordset
Dim sdx As Double
If Trim("" & mytable11.Fields("pm")) <> "S" Then
   valida_rango = 1
   Exit Function
End If
'MsgBox "" & DBGrid2.Columns(51)
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM precios  where  producto='" & Trim(DBGrid2.columns(0)) & "' and local='" & Trim("" & mytable11.Fields("local")) & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   valida_rango = 1
   Select Case "" & DBGrid2.columns("nroprecio")
          Case "1"
               'MsgBox "" & dbgrid2.columns("precio")
               sdx = Val("" & mytablex.Fields("pventa1")) - Val("" & mytablex.Fields("pventa1")) * Val("" & mytablex.Fields("pm1")) / 100
               If Val("" & DBGrid2.columns("precio")) < sdx Then
                  valida_rango = 0
               End If
               
          Case "2"
               sdx = Val("" & mytablex.Fields("pventa2")) - Val("" & mytablex.Fields("pventa2")) * Val("" & mytablex.Fields("pm2")) / 100
               If Val("" & DBGrid2.columns("precio")) < sdx Then
                  valida_rango = 0
               End If
          Case "3"
               sdx = Val("" & mytablex.Fields("pventa3")) - Val("" & mytablex.Fields("pventa3")) * Val("" & mytablex.Fields("pm3")) / 100
               If Val("" & DBGrid2.columns("precio")) < sdx Then
                  valida_rango = 0
               End If
          Case "4"
               sdx = Val("" & mytablex.Fields("pventa4")) - Val("" & mytablex.Fields("pventa4")) * Val("" & mytablex.Fields("pm4")) / 100
               If Val("" & DBGrid2.columns("precio")) < sdx Then
                  valida_rango = 0
               End If
          Case "5"
               sdx = Val("" & mytablex.Fields("pventa5")) - Val("" & mytablex.Fields("pventa5")) * Val("" & mytablex.Fields("pm5")) / 100
               If Val("" & DBGrid2.columns("precio")) < sdx Then
                  valida_rango = 0
               End If
          Case "6"
               sdx = Val("" & mytablex.Fields("pventa6")) - Val("" & mytablex.Fields("pventa6")) * Val("" & mytablex.Fields("pm6")) / 100
               If Val("" & DBGrid2.columns("precio")) < sdx Then
                  valida_rango = 0
               End If
          Case "7"
               sdx = Val("" & mytablex.Fields("pventa7")) - Val("" & mytablex.Fields("pventa7")) * Val("" & mytablex.Fields("pm7")) / 100
               If Val("" & DBGrid2.columns("precio")) < sdx Then
                  valida_rango = 0
               End If
          Case "8"
               sdx = Val("" & mytablex.Fields("pventa8")) - Val("" & mytablex.Fields("pventa8")) * Val("" & mytablex.Fields("pm8")) / 100
               If Val("" & DBGrid2.columns("precio")) < sdx Then
                  valida_rango = 0
               End If
          Case "9"
               sdx = Val("" & mytablex.Fields("pventa9")) - Val("" & mytablex.Fields("pventa9")) * Val("" & mytablex.Fields("pm9")) / 100
               If Val("" & DBGrid2.columns("precio")) < sdx Then
                  valida_rango = 0
               End If
          Case "10"
               sdx = Val("" & mytablex.Fields("pventa10")) - Val("" & mytablex.Fields("pventa10")) * Val("" & mytablex.Fields("pm10")) / 100
               If Val("" & DBGrid2.columns("precio")) < sdx Then
                  valida_rango = 0
               End If
        End Select
   End If
mytablex.Close
End Function
Function valida_placa(buf As String, buf1 As String)
Dim mytablex As New ADODB.Recordset

If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM linea  where  linea='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
      If "" & mytablex.Fields("t1") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If
      If "" & mytablex.Fields("t2") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If
      If "" & mytablex.Fields("t3") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If
      If "" & mytablex.Fields("t4") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If
      If "" & mytablex.Fields("t5") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If
      If "" & mytablex.Fields("t6") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If
      If "" & mytablex.Fields("t7") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If
      If "" & mytablex.Fields("t8") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If
      If "" & mytablex.Fields("t9") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If
      If "" & mytablex.Fields("t10") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If
      If "" & mytablex.Fields("t11") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If
      If "" & mytablex.Fields("t12") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If
      If "" & mytablex.Fields("t13") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If
      If "" & mytablex.Fields("t14") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If
      If "" & mytablex.Fields("t15") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If
      If "" & mytablex.Fields("t16") = buf1 Then
     valida_placa = 1
     GoTo usalir
      End If

   End If
usalir:
   mytablex.Close
   Exit Function

End Function
Sub graba_video_concar(buf As String)

End Sub
Sub valida_camara()
'If 0 < tdeliver.ezVidCap1.NumCapDevs Then
'     tdeliver.ezVidCap1.ShowDlgVideoSource
'Else
'    MsgBox "No Video Capture Device!", vbInformation, App.Title
'End If
Exit Sub
End Sub
Sub busca_ocurrencia()
Dim X As Double
Dim buf As String
Dim mytablex As New ADODB.Recordset
Dim ufile As String
usigue:
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM parame  where  codigo='01'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
  X = Val("" & mytablex.Fields("ocurrencia")) + 1
  ufile = globaldir + "\ocurrencia\" + caja + "-" + "" & X
  If UCase(gocabeza) = "CADIARIO" Then
     ufile = globaldir + "\ocurrencia\" + caja + "-" + "" & X
  End If
  If Dir(ufile) = "" Then 'si no existe
     Else
     'mytablex.Edit
     mytablex.Fields("ocurrencia") = X
     mytablex.Update
     GoTo usigue
  End If
  buf = caja & "-" + "" & X
  graba_video_concar1 buf
  'mytablex.Edit
  mytablex.Fields("ocurrencia") = X
  mytablex.Update
End If
mytablex.Close
End Sub
Sub graba_video_concar1(buf As String)
Dim vr
On Error GoTo cm643122_err
Dim ufile As String
      'Frame10.Visible = True
      'Frame10.Height = 3615
      'Frame10.Top = 2400
      'Frame10.Left = 3120
      'Frame10.Width = 6855
      'ezVidCap1.Height = 3240
      'ezVidCap1.Left = -240
      'ezVidCap1.Top = 240
      'ezVidCap1.Width = 5000
      'ezVidCap1.Visible = False
      'ezVidCap1.Visible = True
      MsgBox "Presione enter para continuar..", 48, "Aviso"
      
ufile = globaldir & "\ocurrencia\" & buf
If UCase(gocabeza) = "CADIARIO" Then
   ufile = globaldir & "\ocurrencia\" & buf
End If
'ezVidCap1.TimeLimit = CInt("" & mytable11.Fields("segundo"))
'ezVidCap1.CaptureFile = ufile
'Call ezVidCap1.CaptureVideo
      'Frame10.Height = 2175
      'Frame10.Top = 0
      'Frame10.Left = 10680
      'Frame10.Width = 3855
      'ezVidCap1.Height = 1920
      'ezVidCap1.Top = 240
      'ezVidCap1.Left = 0
      'ezVidCap1.Width = 3840
Exit Sub
cm643122_err:
MsgBox "Aviso en Video " + error$, 48, "Aviso"
End Sub
Function crea_nuevos_clientes(buf1 As String, buf2 As String, buf3 As String, buf4 As String, buf5 As String, buf6 As String, buf7 As String)
On Error GoTo cmd45777_err
Dim mytablex As New ADODB.Recordset
Exit Function
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM codclie  where  codigo='" & buf1 & "' and producto='" & buf2 & "' and unidad='" & buf5 & "' and factor='" & buf6 & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   'mytablex.Edit
   mytablex.Fields("codigo") = "" & buf1
   mytablex.Fields("producto") = "" & buf2
   mytablex.Fields("descripcio") = "" & buf7
   mytablex.Fields("costo") = Val("" & buf3)
   mytablex.Fields("unidad") = "" & buf5
   mytablex.Fields("factor") = Val("" & buf6)
   If Len(buf4) = 10 Then
      mytablex.Fields("fecha") = Format(buf4, "dd/mm/yyyy")
   End If
   mytablex.Update
   Else
   mytablex.AddNew
   mytablex.Fields("codigo") = "" & buf1
   mytablex.Fields("producto") = "" & buf2
   mytablex.Fields("descripcio") = "" & buf7
   mytablex.Fields("costo") = Val("" & buf3)
   mytablex.Fields("unidad") = "" & buf5
   mytablex.Fields("factor") = Val("" & buf6)
   If Len(buf4) = 10 Then
      mytablex.Fields("fecha") = Format(buf4, "dd/mm/yyyy")
   End If
   mytablex.Update
End If
mytablex.Close
Exit Function
cmd45777_err:
MsgBox "Aviso en nuevo clientes" + error$, 48, "Aviso"
Exit Function

End Function
Function busca_especial(buf, buf1 As String)
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM codclie  where  producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   buf1 = "" & mytablex.Fields("precio")
   busca_especial = 1
End If
mytablex.Close
End Function
Function familia_saldo(buf As String)

Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
   mytablex.Open "SELECT * FROM FAMILIA where familia='" & buf & "'", cn, adOpenStatic, adLockOptimistic
   If mytablex.EOF = True Or mytablex.BOF = True Then
      mytablex.Close
      Exit Function
   End If
   If "" & mytablex.Fields("tipo") = "1" Then
      familia_saldo = 1
   End If
   mytablex.Close

End Function
Function adiciona_deliveri(bxtipo As String, bxserie As String, bxnumero As String)
Dim i As Integer
Dim xbuf As String
Dim found As Integer
Dim mytabler As New ADODB.Recordset
Dim mytablex As New ADODB.Recordset
Dim mytabley As String
Dim mytableb As New ADODB.Recordset
Dim antgocabeza As String
Dim antgodetalle As String
Dim rs
Dim mytablezx As New ADODB.Recordset
On Error GoTo cmd67333_err
'MsgBox gocabeza
antgocabeza = gocabeza
antgodetalle = godetalle
If local1.Visible = True Then
   gocabeza = "ctraslad"
   godetalle = "dtraslad"
End If
If local1 = "PEDIDO" Then
   gocabeza = "crequisa"
   godetalle = "drequisa"
End If

xxacu = busca_acu()
If xxacu = "I" Then 'si es pedido
   gocabeza = "cpedidov"
   godetalle = "dpedidov"
End If

'---validar si el numero ya existe----
'MsgBox globaldir & " " & gocabeza
'AQUI ABRIMOS GAVETA PARA SER MASRAPIDO
If local1.Visible = False Or local1 <> "PEDIDO" Then 'si nos traslado
   If Trim("" & mytable11.Fields("terminal")) <> "T" Then
      found = abre_puerto(Trim("" & mytable11.Fields("capuerto")), Val("" & mytable11.Fields("catipo")), "" & mytable11.Fields("gavetacola"))
   End If
End If
'MsgBox gocabeza
DBGrid2.Enabled = True
found = sumar_detalle()
DBGrid2.Enabled = False
'MsgBox gocabeza

mytablex.Open "SELECT * FROM " & gocabeza & " where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then  'si existe
   'mytablex.Edit
   grabando_cabecera mytablex, bxtipo, bxserie, bxnumero
   mytablex.Update
   Else
   mytablex.AddNew
   grabando_cabecera mytablex, bxtipo, bxserie, bxnumero
   mytablex.Update
End If
mytablex.Close
If Len(petipo) > 0 And Len(penumero) > 0 Then  'si ha sido jalado pedido o orden trabajo descontar credito
   found = descuenta_credito_pedido()
End If
'MsgBox ""
Data2.refresh
ak1:
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM " & godetalle & " where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then  'si existe
   mytablex.Delete
   GoTo ak1
End If
'aqui debe borrar el otro si es traslado
If local1.Visible = True Then
ak12:
If mytableb.State = 1 Then mytableb.Close
mytableb.Open "SELECT * FROM detalle where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='TE' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytableb.RecordCount > 0 Then  'si existe
   mytableb.Delete
   GoTo ak12
End If
ak123:
If mytableb.State = 1 Then mytableb.Close
mytableb.Open "SELECT * FROM detalle where  local='" & Trim("" & mytable11.Fields("local")) & "' and tipo='TS' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic
If mytableb.RecordCount > 0 Then  'si existe
   mytableb.Delete
   GoTo ak123
End If

End If 'fin local visible
xbuf = "cABECERA:" & Format(Now, "hh:mm:ss")

Do
    If Data2.Recordset.EOF Then Exit Do
    mytablex.AddNew
    For i = 0 To Data2.Recordset.Fields.count - 1
        mytablex.Fields(i) = Data2.Recordset.Fields(i)
    Next i
    
    If Val(tdetra) > 0 Then
       mytablex.Fields("denumero") = Format(Val(ndetraccion), "0000000000")
    End If
    mytablex.Fields("sentido") = "" & sentido
    mytablex.Fields("local") = Trim("" & mytable11.Fields("local"))
    mytablex.Fields("tipo") = "" & bxtipo
    mytablex.Fields("serie") = "" & bxserie
    mytablex.Fields("numero") = "" & bxnumero
    If Len(Trim(xvendedor)) > 0 Then
    mytablex.Fields("vendedor") = xvendedor
    End If
    mytablex.Fields("tipoclie") = "C"
    
    mytablex.Fields("moneda") = Trim("" & mytable11.Fields("moneda"))
    mytablex.Fields("bodega") = Trim("" & mytable11.Fields("bodega"))
    mytablex.Fields("bodegaf") = ""
    
    mytablex.Fields("acu") = acu
    mytablex.Fields("localf") = Trim("" & mytable11.Fields("local"))  '& codigo  'si no es traslado
    
    If local1.Visible = True Then
       mytablex.Fields("acu") = "T"
       mytablex.Fields("bodegaf") = xruc '"" & mytable11.Fields("bodega")  'ojo si no esta vacio es traslado
       mytablex.Fields("tipoclie") = "V"
    End If
    If Trim("" & mytable11.Fields("terminal")) = "T" Then
    'mytablex.Fields("acu") = "I"
    End If
    
    mytablex.Fields("acu1") = ""
    'para traslado no debe existir nada
    mytablex.Fields("servicio") = flag_servicio
    If flag_servicio = "A" Then
       mytablex.Fields("servicio") = "A"
    End If
    If flag_servicio = "C" Then
       mytablex.Fields("servicio") = "C"
       mytablex.Fields("salon") = "" & cmytablex.Fields("salon")
       mytablex.Fields("mesa") = "" & cmytablex.Fields("mesa")
       
    End If
    If flag_servicio = "D" Then
       mytablex.Fields("servicio") = "D"
    End If
    mytablex.Fields("flage") = ""
    mytablex.Fields("codigo") = "" & xruc
    mytablex.Fields("caja") = "" & caja
    mytablex.Fields("turno") = "" & turno
    mytablex.Fields("usuario") = "" & cajero
    mytablex.Fields("fecha") = Format(dia, "dd/mm/yyyy")
    mytablex.Fields("hora") = Format(Now, "hh:MM")
    mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
    mytablex.Fields("estado") = "2"
    If local1.Visible = True Then
    mytablex.Fields("codigo") = Trim("" & mytable11.Fields("local"))
    End If
    If xxacu = "I" Then
       mytablex.Fields("acu") = xxacu
    End If
    
    If Label36.Caption = "Almac.Fuente." Then
       mytablex.Fields("bodega") = xruc
       mytablex.Fields("bodegaf") = Trim("" & mytable11.Fields("bodega"))
    End If
    If xxacu = "Q" Then
       mytablex.Fields("acu") = xxacu
    End If
    If local1 = "PEDIDO" Then
       mytablex.Fields("codigo") = ""
    End If
          If bxtipo = "7" Then
         mytablex.Fields("neto") = 0
         mytablex.Fields("descuento") = 0
         mytablex.Fields("subtotal") = 0
         mytablex.Fields("impuesto") = 0
         mytablex.Fields("total") = 0
         mytablex.Fields("xneto") = 0
         mytablex.Fields("tdetra") = 0
      End If
 'ojo aqui debe estar primero creado el codigo
 
 If Len(codigo) > 0 And (bxtipo = "1" Or bxtipo = "2" Or bxtipo = "3" Or bxtipo = "4" Or bxtipo = "5" Or bxtipo = "7") Then
 found = crea_nuevos_clientes("" & codigo, mytablex.Fields("producto"), mytablex.Fields("precio"), mytablex.Fields("fecha"), mytablex.Fields("unidad"), mytablex.Fields("factor"), mytablex.Fields("descripcio"))
 End If
 
    mytablex.Fields("flage") = "V"
    mytablex.Update
    'miramos si es combo-------------------------
    'If verifica_combo("" & Data2.Recordset.Fields("producto")) = 1 Then
    'If mytablezx.State = 1 Then mytablezx.Close
    '    mytablezx.Open "SELECT * FROM _c" & gusuario & " where producto='" & Data2.Recordset.Fields("producto") & "'", cn, adOpenDynamic, adLockOptimistic
    '    If mytablezx.RecordCount > 0 Then
    '      Do
    '      If mytablezx.EOF Then Exit Do
    '      mytablex.AddNew
    '      For i = 0 To Data2.Recordset.Fields.count - 1
    '      mytablex.Fields(i) = Data2.Recordset.Fields(i)
    '      Next i
    '
    '      mytablex.Fields("sentido") = "" & sentido
    '      mytablex.Fields("local") = Trim("" & mytable11.Fields("local"))
    '      mytablex.Fields("tipo") = "" & bxtipo
    '      mytablex.Fields("serie") = "" & bxserie
    '      mytablex.Fields("numero") = "" & bxnumero
    '      mytablex.Fields("vendedor") = xvendedor
    '      mytablex.Fields("tipoclie") = "C"
    '      mytablex.Fields("moneda") = Trim("" & mytable11.Fields("moneda"))
    '      mytablex.Fields("bodega") = Trim("" & mytable11.Fields("bodega"))
    '      mytablex.Fields("bodegaf") = ""
    '      mytablex.Fields("acu") = acu
    '      mytablex.Fields("localf") = Trim("" & mytable11.Fields("local"))  '& codigo  'si no es traslado
        
    '      mytablex.Fields("flage") = ""
    '      mytablex.Fields("codigo") = "" & xruc
    '      mytablex.Fields("caja") = "" & caja
    '      mytablex.Fields("turno") = "" & turno
    '      mytablex.Fields("usuario") = "" & cajero
    '      mytablex.Fields("fecha") = Format(dia, "dd/mm/yyyy")
    '      mytablex.Fields("hora") = Format(Now, "hh:MM")
    '      mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
    '      mytablex.Fields("estado") = "2"

    
    '      mytablex.Fields("producto") = "" & mytablezx.Fields("productop")
    '      mytablex.Fields("descripcio") = "" & mytablezx.Fields("descripciop")
    '      mytablex.Fields("unidad") = "UND" '& mytablezx.Fields("unidad")
    '      mytablex.Fields("cantidad") = Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & mytablezx.Fields("cantidad")) * 1 ' Val("" & mytablezx.Fields("factor"))
    '      mytablex.Fields("dua") = "C"  'C ES COMBO flag que dice que es receta
    '      mytablex.Fields("acu") = "T"  'guia de salida
    '      mytablex.Fields("precio") = 0
    '      mytablex.Fields("total") = 0
    '      mytablex.Fields("subtotal") = 0
    '      mytablex.Fields("impuesto") = 0
    '      mytablex.Update
    '    mytablezx.MoveNext
    '    Loop
    '    End If
    '    mytablezx.Close
       
    
    
    '------ fin de combo
    'End If
    
    'GRABANDO CLIENTES
    'verificamos si tiene receta
    'GoTo pasa_receta
    '----------------------------------------
    
    
    If verifica_receta("" & Data2.Recordset.Fields("producto")) = 1 Then
       '---------------------------------------
       If mytablezx.State = 1 Then mytablezx.Close
        mytablezx.Open "SELECT * FROM receta where producto='" & Data2.Recordset.Fields("producto") & "'", cn, adOpenDynamic, adLockOptimistic
        If mytablezx.RecordCount > 0 Then
          Do
          If mytablezx.EOF Then Exit Do
          mytablex.AddNew
          For i = 0 To Data2.Recordset.Fields.count - 1
          mytablex.Fields(i) = Data2.Recordset.Fields(i)
          Next i
        
          mytablex.Fields("sentido") = "" & sentido
          mytablex.Fields("local") = Trim("" & mytable11.Fields("local"))
          mytablex.Fields("tipo") = "" & bxtipo
          mytablex.Fields("serie") = "" & bxserie
          mytablex.Fields("numero") = "" & bxnumero
          If Len(Trim(xvendedor)) > 0 Then
             mytablex.Fields("vendedor") = xvendedor
          End If
          'mytablex.Fields("vendedor") = xvendedor
          mytablex.Fields("tipoclie") = "C"
          mytablex.Fields("moneda") = Trim("" & mytable11.Fields("moneda"))
          mytablex.Fields("bodega") = Trim("" & mytable11.Fields("bodega"))
          mytablex.Fields("bodegaf") = ""
          mytablex.Fields("acu") = acu
          mytablex.Fields("localf") = Trim("" & mytable11.Fields("local"))  '& codigo  'si no es traslado
        
          mytablex.Fields("flage") = ""
          mytablex.Fields("codigo") = "" & xruc
          mytablex.Fields("caja") = "" & caja
          mytablex.Fields("turno") = "" & turno
          mytablex.Fields("usuario") = "" & cajero
          mytablex.Fields("fecha") = Format(dia, "dd/mm/yyyy")
          mytablex.Fields("hora") = Format(Now, "hh:MM")
          mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
          mytablex.Fields("estado") = "2"

    
          mytablex.Fields("producto") = "" & mytablezx.Fields("productoi")
          mytablex.Fields("descripcio") = "" & mytablezx.Fields("descripcio")
          mytablex.Fields("unidad") = "" & mytablezx.Fields("unidad")
          mytablex.Fields("cantidad") = Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & mytablezx.Fields("cantidad")) * Val("" & mytablezx.Fields("factor"))
          mytablex.Fields("dua") = "R"  'flag que dice que es receta
          mytablex.Fields("acu") = "T"  'guia de salida
          mytablex.Fields("precio") = 0
          mytablex.Fields("total") = 0
          mytablex.Fields("subtotal") = 0
          mytablex.Fields("impuesto") = 0
          
          mytablex.Update
        mytablezx.MoveNext
        Loop
        End If
        mytablezx.Close
        '---------------------------------------
    End If
pasa_receta:
    
    'MsgBox "Hola"
    If local1.Visible = True Then  'si es traslado
    mytableb.AddNew
    'MsgBox "Hola"
    For i = 0 To Data2.Recordset.Fields.count - 1
        mytableb.Fields(i) = Data2.Recordset.Fields(i)
    Next i
    'MsgBox "Hola"
    mytableb.Fields("local") = Trim("" & mytable11.Fields("local"))
    mytableb.Fields("tipo") = "TS" '& bxtipo
    mytableb.Fields("serie") = "" & bxserie
    mytableb.Fields("numero") = "" & bxnumero
    If Len(Trim(xvendedor)) > 0 Then
       mytableb.Fields("vendedor") = xvendedor
    End If
    'mytableb.Fields("vendedor") = xvendedor
    mytableb.Fields("moneda") = Trim("" & mytable11.Fields("moneda"))
    mytableb.Fields("bodega") = Trim("" & mytable11.Fields("bodega"))
    mytableb.Fields("bodegaf") = "" 'xruc '"" & mytable11.Fields("bodega")
    mytableb.Fields("acu") = "T"
    mytableb.Fields("localf") = Trim("" & mytable11.Fields("local"))
    mytableb.Fields("tipoclie") = "V"
    If Trim("" & mytable11.Fields("terminal")) = "T" Then
    'mytablex.Fields("acu") = "I"
    End If
    mytableb.Fields("acu1") = ""
    'para traslado no debe existir nada
    mytableb.Fields("servicio") = flag_servicio
    If flag_servicio = "A" Then
       mytableb.Fields("servicio") = "A"
    End If
    If flag_servicio = "C" Then
       mytableb.Fields("servicio") = "C"
    End If
    If flag_servicio = "D" Then
       mytableb.Fields("servicio") = "D"
    End If
    mytableb.Fields("flage") = ""
    mytableb.Fields("codigo") = Trim("" & mytable11.Fields("local"))
    mytableb.Fields("caja") = "" & caja
    mytableb.Fields("turno") = "" & turno
    mytableb.Fields("usuario") = "" & cajero
    mytableb.Fields("fecha") = Format(dia, "dd/mm/yyyy")
    mytableb.Fields("hora") = Format(Now, "hh:MM")
    mytableb.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
    mytableb.Fields("estado") = "2"
    If Label36.Caption = "Almac.Fuente." Then
       mytableb.Fields("bodega") = xruc
       mytableb.Fields("bodegaf") = ""
    End If
    'mytableb.Fields("local1") = "" & mytable11.Fields("local")
    '---------------ojo no debe ir en detalle
    'mytableb.Fields("tipo1") = "" & petipo
    'mytableb.Fields("serie1") = "" & peserie
    'mytableb.Fields("numero1") = "" & penumero
    '-------------------------------------
    mytableb.Update
    'MsgBox "Hola"
    'AHORA LA ENTRADA
    '-----------------------------------
    mytableb.AddNew
    For i = 0 To Data2.Recordset.Fields.count - 1
        mytableb.Fields(i) = Data2.Recordset.Fields(i)
    Next i
    mytableb.Fields("local") = Trim("" & mytable11.Fields("local"))
    mytableb.Fields("tipo") = "TE" '& bxtipo
    mytableb.Fields("serie") = "" & bxserie
    mytableb.Fields("numero") = "" & bxnumero
    If Len(Trim(xvendedor)) > 0 Then
    mytableb.Fields("vendedor") = xvendedor
    End If
    mytableb.Fields("moneda") = Trim("" & mytable11.Fields("moneda"))
    mytableb.Fields("bodega") = xruc
    mytableb.Fields("bodegaf") = "" 'xruc '"" & mytable11.Fields("bodega")
    mytableb.Fields("acu") = "S"
    mytableb.Fields("localf") = Trim("" & mytable11.Fields("local"))
    mytableb.Fields("tipoclie") = "V"
    If Trim("" & mytable11.Fields("terminal")) = "T" Then
    'mytablex.Fields("acu") = "I"
    End If
    mytableb.Fields("acu1") = ""
    'para traslado no debe existir nada
    mytableb.Fields("servicio") = flag_servicio
    If flag_servicio = "A" Then
       mytableb.Fields("servicio") = "A"
    End If
    If flag_servicio = "C" Then
       mytableb.Fields("servicio") = "C"
    End If
    If flag_servicio = "D" Then
       mytableb.Fields("servicio") = "D"
    End If
    mytableb.Fields("flage") = ""
    mytableb.Fields("codigo") = Trim("" & mytable11.Fields("local"))
    mytableb.Fields("caja") = "" & caja
    mytableb.Fields("turno") = "" & turno
    mytableb.Fields("usuario") = "" & cajero
    mytableb.Fields("fecha") = Format(dia, "dd/mm/yyyy")
    mytableb.Fields("hora") = Format(Now, "hh:MM")
    mytableb.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
    mytableb.Fields("estado") = "2"
    If Label36.Caption = "Almac.Fuente." Then
       mytableb.Fields("bodega") = Trim("" & mytable11.Fields("bodega"))
       'mytableb.Fields("bodegaf") = ""
    End If
    mytableb.Update
    '-----------------------------------
    End If
    Data2.Recordset.MoveNext
Loop
'MsgBox "Hola"
xbuf = xbuf & " Detalle:" & Format(Now, "hh:mm:ss")

'AQUI DEBE DESCARGAR EL SALDO ACTUAL
If Trim("" & mytable11.Fields("terminal")) <> "T" And local1.Visible = False And local1 <> "PEDIDO" Then
'MsgBox "Hola1"
If crucefa.ListCount = 0 Then  'si no es facturacion mensual
   mytablex.MoveFirst
   found = descarga_saldo(Trim("" & mytable11.Fields("local")), mytablex, bxtipo, bxserie, bxnumero, -1, 0)
End If
End If
If local1.Visible = True Then  'si es traslado
   'MsgBox "Hola"
   mytableb.MoveFirst
   found = descarga_saldo(Trim("" & mytable11.Fields("local")), mytableb, "TS", bxserie, bxnumero, -1, 0)
   'MsgBox "Hola1"
   mytableb.MoveFirst
   found = descarga_saldo(Trim("" & mytable11.Fields("local")), mytableb, "TE", bxserie, bxnumero, 1, 0)
End If
If local1.Visible = True Then
   mytableb.Close
End If
mytablex.Close
xbuf = xbuf & " Saldo:" & Format(Now, "hh:mm:ss")

      If Trim("" & mytable11.Fields("terminal")) <> "T" Then  'finalizar el terminal
         found = graba_guia_mensual() 'graba cuando es cruce de guias
      End If
found = busca_xtipog("" & bxtipo)  'graba el numero al actual
xbuf = xbuf & " tipo:" & Format(Now, "hh:mm:ss")
'MsgBox "XX"
'MsgBox "Pedido Grabado con la Orden Nro:" & xnumero
If Trim("" & mytable11.Fields("terminal")) <> "T" And local1.Visible = False And local1 <> "PEDIDO" Then
   'MsgBox "x"
   If local1.Visible = False Then
      If Val(acuenta) > 0 And xtipo = Trim("" & mytable11.Fields("tipope")) And Len(petipo) = 0 Then 'si es pedido a cuenta grabar
          menu_graba_fpedido
          Else
          xbuf = xbuf & " Fpago Antes:" & Format(Now, "hh:mm:ss")
          found = graba_fpagov(bxtipo, bxserie, bxnumero) 'graba fpagov
          xbuf = xbuf & " Fpago Despues:" & Format(Now, "hh:mm:ss")
      End If
   End If
End If

xbuf = xbuf & " FIN:" & Format(Now, "hh:mm:ss")
'MsgBox xbuf

If Len(pedido) = 0 Then  'si no es modificacion de pedido
   proceso_impresion11 "" & bxtipo, "" & bxserie, "" & bxnumero, 1, ""
End If
If Trim("" & mytable11.Fields("hod")) = "S" And flag_servicio <> "C" Then 'enviar orden de despacho
        'If "" & mytable11.Fields("comanda") = "S" Then
           found = orden_despacho_n("" & mytable11.Fields("local"), bxtipo, bxserie, bxnumero, "")
        'Else
        'orden_normal
        'End If
End If

If Trim("" & mytable11.Fields("video")) = "S" Then
   If bxtipo = "7" Or Len(ndetraccion) > 0 Then
      'Frame10.Enabled = False
      graba_video_concar bxserie & "-" & bxnumero
      'Frame10.Enabled = True
   End If
End If
'impresion_sin_formato xtipo, xserie, xnumero
'MsgBox "x"
found = borrar_proformas()
'MsgBox "Hola"
found = borrar_pedidos()
'MsgBox "Hola"
found = borrar_cotizacion()

found = borrar_comanda()
inicialIzatodo
'MsgBox "Hola"
gocabeza = antgocabeza
godetalle = antgodetalle
'losao94_Click
'losao94_Click

Exit Function
cmd67333_err:
gocabeza = antgocabeza
godetalle = antgodetalle
MsgBox "Error en GRABACION TOTAL " + error$, 48, "Aviso"
Exit Function
End Function
Function borrar_comanda()
Dim buf As String
Dim buf1 As String
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd8900012_err

If flag_servicio = "C" Then
buf = cmytablex.Fields("salon")
buf1 = cmytablex.Fields("mesa")
'If cmytablex.RecordCount > 0 Then
If cuenta_separa <> "S" Then
   cn.Execute ("delete from dcomanda where salon='" & buf & "' and mesa='" & buf1 & "'")
   cn.Execute ("update mesa set estado='' where salon='" & buf & "' and mesa='" & buf1 & "'")
End If
If cuenta_separa = "S" Then
   cn.Execute ("update dcomanda set cantidad=cantidad-cantdev where salon='" & buf & "' and mesa='" & buf1 & "' and cantdev>0")
   cn.Execute ("delete from dcomanda where salon='" & buf & "' and mesa='" & buf1 & "' and cantidad=0 and cantdev>0")
   cn.Execute ("update dcomanda set cantdev=0 ,total=cantidad*precio where salon='" & buf & "' and mesa='" & buf1 & "' and  cantdev>0")
   'cn.Execute ("delete from dcomanda where salon='" & buf & "' and mesa='" & buf1 & "' and flage='S'")
   'mytablex.Open "SELECT * FROM dcomanda where  salon='" & buf & "' and mesa='" & buf1 & "' and cantdev>0", cn, adOpenDynamic, adLockOptimistic
   'If mytablex.RecordCount = 0 Then  'si existe
   '   mytablex.Close
   'End If
   'Do
   'If mytablex.EOF Then Exit Do
   'mytablex.MoveNext
   'Loop
End If
End If
Exit Function
cmd8900012_err:
MsgBox "Aviso en borrar Comanda " + error$, 48, "Aviso"
Exit Function

End Function


Sub grabando_cabecera(mytablex As ADODB.Recordset, bxtipo As String, bxserie As String, bxnumero As String)
On Error GoTo cmd232_err
'MsgBox ""
If Val(tdetra) > 0 Then
   mytablex.Fields("denumero") = Format(Val(ndetraccion), "0000000000")
End If
mytablex.Fields("sentido") = sentido
mytablex.Fields("observa") = xdistrito
mytablex.Fields("tdetra") = Val(tdetra)
mytablex.Fields("xneto") = Val(tpeaje)
mytablex.Fields("tisc") = Val(tisc)
mytablex.Fields("tivap") = Val(tivap)
mytablex.Fields("tipo1") = petipo
mytablex.Fields("serie1") = peserie
mytablex.Fields("numero1") = penumero

If Len(Trim(referencia)) > 0 Then  'que es referencia
   mytablex.Fields("observa") = Mid$("" & referencia, 1, 60)
End If
'MsgBox ""
mytablex.Fields("adetotal") = 0
mytablex.Fields("acuenta") = Val(acuenta)
mytablex.Fields("retipo1") = ""
mytablex.Fields("renumero1") = ""
mytablex.Fields("renumero2") = ""
mytablex.Fields("renumero3") = ""
mytablex.Fields("retotal1") = 0
mytablex.Fields("retotal2") = 0
mytablex.Fields("retotal3") = 0
mytablex.Fields("retotal") = 0
mytablex.Fields("zona") = ""
mytablex.Fields("nombre") = xnombre
'MsgBox ""
mytablex.Fields("estado") = "2"
mytablex.Fields("tipoclie") = "C"
mytablex.Fields("tipo") = "" & bxtipo
mytablex.Fields("serie") = "" & bxserie
mytablex.Fields("numero") = bxnumero
mytablex.Fields("codigo") = xruc
mytablex.Fields("partida") = ""
mytablex.Fields("destino") = ""
mytablex.Fields("yausado") = "0"
mytablex.Fields("nro_items") = Val(ntcant)
mytablex.Fields("fecha") = Format(dia, "dd/mm/yyyy")
mytablex.Fields("fechae") = Format(Now, "dd/mm/yyyy")
mytablex.Fields("moneda") = Trim("" & mytable11.Fields("moneda"))
If Len(Trim(xvendedor)) > 0 Then
mytablex.Fields("vendedor") = xvendedor
End If
mytablex.Fields("fpago") = ""
mytablex.Fields("transporte") = ""
mytablex.Fields("paridad") = Val(paridad)
mytablex.Fields("dias") = 1
mytablex.Fields("bodega") = Trim("" & mytable11.Fields("bodega"))
mytablex.Fields("bodegaf") = ""
'mytablex.Fields("observa") = ""
mytablex.Fields("usuario") = "" & gusuario
mytablex.Fields("caja") = "" & caja
mytablex.Fields("turno") = "" & turno
mytablex.Fields("usuario") = "" & cajero
mytablex.Fields("acu") = acu
'MsgBox acu
If Trim("" & mytable11.Fields("terminal")) = "T" Then
    'mytablex.Fields("acu") = "I"
End If
mytablex.Fields("acu1") = ""
mytablex.Fields("flage") = ""
mytablex.Fields("telefono") = "" & telefono
mytablex.Fields("hora") = Format(Now, "hh:MM")
mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
mytablex.Fields("gravado") = Val("" & gravado)
mytablex.Fields("total") = Val("" & txtotal)
mytablex.Fields("redondeo") = Val(Format(txtotlare, nrodecimal))
mytablex.Fields("descuento") = Val("" & txdescuento)
mytablex.Fields("neto") = Val("" & txneto)
mytablex.Fields("impuesto") = Val("" & tximpuesto)
mytablex.Fields("subtotal") = Val("" & txsubtotal)

'mytablex.Fields("tipo1") = ""
'mytablex.Fields("serie1") = ""
mytablex.Fields("serie2") = ""
mytablex.Fields("serie3") = ""
mytablex.Fields("serie4") = ""
mytablex.Fields("serie5") = ""
mytablex.Fields("serie6") = ""
mytablex.Fields("serie7") = ""

'mytablex.Fields("numero1") = ""
mytablex.Fields("numero2") = ""
mytablex.Fields("numero3") = ""
mytablex.Fields("numero4") = ""
mytablex.Fields("numero5") = ""
mytablex.Fields("numero6") = ""
mytablex.Fields("numero7") = ""
mytablex.Fields("c1") = Val(c1)
mytablex.Fields("c2") = Val(c2)
mytablex.Fields("c3") = Val(c3)
mytablex.Fields("c4") = Val(c4)
mytablex.Fields("c5") = Val(c5)
mytablex.Fields("c6") = Val(c6)
mytablex.Fields("c7") = Val(c7)
mytablex.Fields("c8") = Val(c8)
mytablex.Fields("c9") = Val(c9)
mytablex.Fields("local") = Trim("" & mytable11.Fields("local"))
mytablex.Fields("montopagar") = 0
mytablex.Fields("ruc") = "" & xruc
mytablex.Fields("TDOCDELI") = ""
mytablex.Fields("servicio") = flag_servicio
If flag_servicio = "A" Then
   mytablex.Fields("servicio") = "A"
End If
If flag_servicio = "D" Then
mytablex.Fields("servicio") = "D"
End If
If flag_servicio = "C" Then
mytablex.Fields("servicio") = "C"
mytablex.Fields("salon") = "" & cmytablex.Fields("salon")
mytablex.Fields("mesa") = "" & cmytablex.Fields("mesa")
End If
'validamos aqui si es traslado
If local1.Visible = True Then
   mytablex.Fields("localf") = Trim("" & mytable11.Fields("local"))
   mytablex.Fields("tipoclie") = "L"
   mytablex.Fields("bodegaf") = xruc
   mytablex.Fields("codigo") = Trim("" & mytable11.Fields("local"))
End If
If xxacu = "I" Then
   mytablex.Fields("acu") = xxacu
End If
If xxacu = "Q" Then
   mytablex.Fields("acu") = xxacu
End If
If local1 = "PEDIDO" Then
    mytablex.Fields("CODIGO") = ""
    mytablex.Fields("nombre") = "PEDIDO REPOSICION"
End If
If Label36.Caption = "Almac.Fuente." Then
 mytablex.Fields("bodega") = xruc
 mytablex.Fields("bodegaf") = Trim("" & mytable11.Fields("local"))
End If
       If bxtipo = "7" Then
         mytablex.Fields("neto") = 0
         mytablex.Fields("descuento") = 0
         mytablex.Fields("subtotal") = 0
         mytablex.Fields("impuesto") = 0
         mytablex.Fields("total") = 0
         mytablex.Fields("xneto") = 0
         mytablex.Fields("tdetra") = 0
      End If
'MsgBox "x"
mytablex.Fields("flage") = "V"
'si es consumo grabar en descripcio
If Label59.Caption = "CONSUMO" Then
   mytablex.Fields("observa") = "CONSUMO"
End If
grabar_dato_pedido codigo, bxtipo, bxserie, bxnumero
Exit Sub
cmd232_err:
MsgBox "Error en grabando Cabecera " & error$, 48, "Aviso"
Exit Sub
End Sub
Sub busca_correlativo(sw As Integer)
Dim sdx As Double
Dim mytablex As New ADODB.Recordset
mytablex.Open "select * from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   mytablex.Close
   Exit Sub
End If
If sw = 0 Then
sdx = Val("" & mytablex.Fields("clientes")) + 1
dcodigo = "" & sdx
End If
If sw = 1 Then
   If IsNumeric(dcodigo) Then
   mytablex.Fields("clientes") = dcodigo
   mytablex.Update
   End If
   mytablex.Close
   Exit Sub
End If

mytablex.Close
sigueb:
mytablex.Open "select * from clientes where codigo='" & dcodigo & "'", cn, adOpenStatic, adLockOptimistic
If mytablex.RecordCount > 0 Then
   mytablex.Close
   sdx = sdx + 1
   dcodigo = "" & sdx
   GoTo sigueb
   Exit Sub
End If
mytablex.Close
End Sub
Function busca_banco(buf As String)
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM banco where banco='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   busca_banco = 1
End If
mytablex.Close

End Function
Function verifica_receta(buf As String)
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM receta where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   verifica_receta = 1
End If
mytablex.Close
End Function
Function verifica_combo(buf As String)
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM _c" & gusuario & " where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   verifica_combo = 1
End If
mytablex.Close

End Function
Function busca_remate(buf As String)
Dim mytablex As New ADODB.Recordset
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT producto,remate FROM producto where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   If "" & mytablex.Fields("remate") = "S" Then
      busca_remate = 1
   End If
End If
mytablex.Close
End Function
Sub sumar_reforzar()
Dim sdx As Double
Dim mytablex As Table
Exit Sub
Set mytablex = mydbxglo.OpenTable(dgusuario)
sdx = 0
Do
If mytablex.EOF Then Exit Do
MsgBox "" & mytablex.Fields("producto")
'If Val("" & mytablex.Fields("total")) > 0 Then
   sdx = sdx + Val("" & mytablex.Fields("total"))
'End If
mytablex.MoveNext
Loop
MsgBox sdx
rtxtotal = Format(sdx, "0.00")
mytablex.Close



End Sub






Private Sub zfamilia_Click(Index As Integer)
Dim buff As String
If Len(wwfamcod(Index)) = 0 Then
   Exit Sub
End If
   buff = "" & wwfamcod(Index)

'menu_carga_producto zfamilia(Index).Caption
menu_carga_producto buff
menu_producto "INI"

End Sub
Sub menu_carga_producto(buf As String)
Dim mytablex As New ADODB.Recordset

Dim i As Integer
For i = 0 To 29
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
            mprodpag = mprodpag + 29
            If mprodpag > 102 Then
               mprodpag = 0
            End If
       Case "ANT"
            mprodpag = mprodpag - 29
            If mprodpag < 0 Then
               mprodpag = 0
            End If
End Select
j = -1
For i = mprodpag To 29 + mprodpag
    j = j + 1
    zproducto(j).Caption = mprodcod(i)
    wwprodcod(j) = wprodcod(i)
Next i


End Sub

Private Sub znumero_Click(Index As Integer)
Dim found As Integer
       If Index = 10 Then
          hknumero = ""
          Exit Sub
       End If
        found = InStr(hknumero, "/")
        If found > 1 Then  ' si es cantidad
           Exit Sub
        End If
       
       hknumero = hknumero & znumero(Index).Caption

End Sub

Private Sub zproducto_Click(Index As Integer)
Dim mytablex As New ADODB.Recordset
Dim buff As String
Dim bufy As String
Dim found As Integer
Dim canti As String
Dim buf As String
Dim xcampo As String
Dim buf1 As String
Dim xsw As Integer
Dim abuf As String
Data2.refresh
DBGrid2.Col = 0
DBGrid2.Row = DBGrid2.VisibleRows - 1
DBGrid2.SetFocus
stkminimo = ""

xsw = 0
If Not Data2.Recordset.EOF Then
   Data2.Recordset.MoveLast
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus
End If
   'Data2.Recordset.MoveLast
   'dbgrid2.Col = 0
   'dbgrid2.Row = dbgrid2.VisibleRows - 1
   'dbgrid2.SetFocus

If Len(wwprodcod(Index)) = 0 Then
   Exit Sub
End If
   buff = "" & wwprodcod(Index)
   If "" & mytable11.Fields("nosaldo") = "S" Then
            '-------------------------
            found = verifica_saldo_receta(buff, Val(hknumero))
            If found = 2 Then
               MsgBox "Se detecto un saldo receta con saldo<=0 ", 48, "Aviso"
               Exit Sub
            End If
            '-------------------------
            mytablex.Open "SELECT * FROM producto where producto='" & buff & "'", cn, adOpenDynamic, adLockOptimistic
            If mytablex.RecordCount > 0 Then
               If familia_saldo("" & mytablex.Fields("familia")) = 0 Then
                  If consulta_saldo(buff, Val(hknumero), 0) <= 0 Then
                     MsgBox "x.No existe saldo", 48, "Aviso"
                     Exit Sub
                  End If
               End If
            End If
   End If
     
     
stkminimo = ""
If "" & mytable11.Fields("stkminimo") = "S" Then
            mytablex.Open "SELECT * FROM producto where producto='" & buff & "'", cn, adOpenDynamic, adLockOptimistic
            If mytablex.RecordCount > 0 Then
               If familia_saldo("" & mytablex.Fields("familia")) = 0 Then
                  consulta_minimo buff, "" & mytablex.Fields("minimo")
               End If
            End If
            mytablex.Close
End If
     
     
     'MsgBox "abcd"
     canti = hknumero
     buf = buff  'se modifico en U. Union
     bufy = buf
     found = 0
         
     control_flujo = 0
     '-------------------------------------------------
     '---------- verificamos si existe /----------
     'MsgBox canti
     
     found = InStr(canti, "/")
        If found > 1 Then  ' si es cantidad
                  'MsgBox found
                  xcampo = Mid$(canti, found + 1, Len(canti) - found)
                  canti = Mid$(canti, 1, found - 1)
                  If Val(canti) = 0 Then
                     Exit Sub
                  End If
                  xsw = 1
        End If
        'AQUI SOLO SE DEBE EXTRAER LA CANTIDAD
     
     'MsgBox xcampo
     'SI EXISTE / ES PORQUE CANTIDAD ES TOTAL
     found = busca_producto(buff, 0, canti, xsw)
     'found = busca_producto(buf, 0, canti)
     'MsgBox found
     If found = 2 Then  'si es precio 0
        'MsgBox "XX-XX"
        control_flujo = 1
        Exit Sub
     End If
     'MsgBox ""
     
            If ver_si_puedo_dbgrid("" & DBGrid2.columns(0)) = 1 Then  'existe mas de un precio
               'MsgBox "abc"
               xproducto = "" & DBGrid2.columns(0)
               carga_dbgrid4 "" & DBGrid2.columns(0)
               Exit Sub
            End If
            If Len(DBGrid2.columns(0)) > 0 And Len(DBGrid2.columns(17)) > 0 Then
               'DBGrid2.Col = 3
               'ingreso_tallas "" & DBGrid2.Columns(17)
               'Exit Sub
            End If
     
     found = existe_fuel("" & DBGrid2.columns(0))
     If found = 1 And Val("" & DBGrid2.columns("cantidad")) = 1 Then
        DBGrid2.Col = 7
        DBGrid2.SetFocus
        Exit Sub
     End If
     'MsgBox ""
     found = sumar_detalle()
     'Data2.Refresh
     'Data2.Recordset.MoveLast
     DBGrid2.Col = 0
     DBGrid2.Row = DBGrid2.VisibleRows - 1
     'Data2.Refresh
     hknumero = ""
     'MsgBox ""

End Sub
Function orden_despacho()
Dim xdato As String
Dim buf As String
Dim bufx As String
Dim Puerto As String
Dim puertos As String
Dim puertod As String
Dim found As Integer
Dim cola As String
Dim i As Integer
Dim j As Integer
Dim X As Integer
Dim xbuf0 As String
Dim xbuf1 As String
Dim xbuf2 As String
Dim sFile As String
Dim sw As Integer
Dim oldprinter
'Dim mydbf As Database
On Error GoTo cmd7890_err
'impresora por default atachado
'If MsgBox("Desea Imprimir Orden Despacho ", 1, "Aviso") <> 1 Then Exit Function
'sum1 = 0
Puerto = ""
puertod = ""
puertos = "oc" '& mytable11.Fields("puertocie")  'impresora x defecto
Puerto = puertos
cerrar_archivo
FileName = caja & Puerto
found = borra_nombre("" & FileName)
'-----------ORDEN DE DESPACHO---------------------------------------------
List1.Clear
'---------------------
cola = "N"
Puerto = "LPT"
Data2.refresh
ncanal = 1

Do
   If Data2.Recordset.EOF Then Exit Do
   'MsgBox "" & Data2.Recordset.Fields("producto")
   If Len("" & Data2.Recordset.Fields("producto")) > 0 And (Val("" & Data2.Recordset.Fields("cantidad")) > 0 Or Val("" & Data2.Recordset.Fields("cantidad")) < 0) Then
      found = busca_familia_orden("" & Data2.Recordset.Fields("producto"), Puerto, puertod, cola)
      If found = 0 Then   'si no existe debe tomar el defaul de la impresora
          Puerto = puertos
      End If
      If Len(Trim(Puerto)) = 0 Then
         Puerto = "LPT"
      End If
      'MsgBox found
   '--------------------------------------
      sw = 0
      FileName = Trim(caja & Puerto)
      found = existearchivo("" & FileName)
      If found = 1 Then  'verificar si no existe en la lista
         sw = 0
         For i = 0 To List1.ListCount - 1
          j = InStr(List1.List(i), "|")
          xbuf0 = Mid$(List1.List(i), 1, j - 1)
          If xbuf0 = FileName Then
             sw = 1
          End If
         Next i
         If sw = 0 Then  'no existe en la lista
            found = borra_nombre(FileName)
            found = 0
         End If
      End If
      cerrar_archivo
      Open FileName For Append As #ncanal
      If found = 0 Then
         List1.AddItem FileName & "|" & puertod & "|" & cola & "|" 'adiciona en la lista
         cabecera_orden_despacho "" & Data2.Recordset.Fields("vendedor"), "", "", ""
      End If
      imprime_detalle_orden
      Close #ncanal
   End If
Data2.Recordset.MoveNext
Loop
cerrar_archivo
buf = ""
For i = 0 To List1.ListCount - 1
    buf = buf & List1.List(i)
Next i
'MsgBox buf

'-------------se adiciono para agilidad--------------------------------
For i = 0 To List1.ListCount - 1
   xdato = List1.List(i)
   extrae_puertos xdato, xbuf0, xbuf1, xbuf2
   FileName = xbuf0
   If existearchivo(xbuf0) = 1 Then
      Open FileName For Append As #ncanal
      For X = 1 To 5
          Print #ncanal, ""
      Next X
      Print #ncanal, ""
      Close #ncanal
   End If
Next i
'MsgBox List1.ListCount
For i = 0 To List1.ListCount - 1
   xdato = List1.List(i)
   'MsgBox xdato
   extrae_puertos xdato, xbuf0, xbuf1, xbuf2
   FileName = xbuf0
   'MsgBox xdato & " " & xbuf0 & " " & xbuf1 & " " & xbuf2
   If existearchivo(xbuf0) = 1 Then
      If xbuf2 = "S" Then
         'MsgBox xbuf1
         oldprinter = Printer.DeviceName
         selecciona_impresoras (Trim(xbuf1))
         'imprime_archivotexto xbuf0
         found = Imprime_archivojj(xbuf0, 0, "" & mytable11.Fields("tamanorden"), "" & mytable11.Fields("nombrefont"))
         'found = imprime_archivoj(xbuf0, 0, "" & mytable11.Fields("tamanorden"))
         selecciona_impresoras (Trim(oldprinter))
         'found = orden_oprn(xbuf1, "" & mytable11.Fields("tipoleta"), "" & mytable11.Fields("tamano"), "" & mytable11.Fields("negrita"))
      Else
      Open FileName For Append As #ncanal
      For X = 1 To 2
          Print #ncanal, ""
      Next X
      Print #ncanal, ""
      Close #ncanal
      found = star_sp342(Trim(xbuf1), 0)
      found = corte_papel(Trim(xbuf1), 1)
      End If
      cerrar_archivo
      found = borra_nombre(xbuf0)
   End If
Next i
cerrar_archivo
orden_despachot
Exit Function
cmd7890_err:
   MsgBox "MENSAJE, ERROR EN ORDEN DESPACHO " & error$, 24, "AVISO"
   cerrar_archivo
   Exit Function
End Function
Function busca_familia_orden(buf1 As String, Puerto As String, puertod As String, cola As String)
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd90fam_err
mytablex.Open "SELECT * FROM producto where producto='" & buf1 & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then
      Puerto = "" & mytablex.Fields("grupoimpresion")
      puertod = "" & mytablex.Fields("puertoimpresion")
      cola = "" & mytablex.Fields("cola")
      busca_familia_orden = 1
End If
mytablex.Close
'orden_despacho
   Exit Function
cmd90fam_err:
   MsgBox "Aviso en Busca Familia orden " + error$, 48, "Aviso"
   Exit Function

End Function

Sub cabecera_orden_despacho(buvendedor As String, buf1 As String, buf2 As String, buf3 As String)
Dim found As Integer
Dim buf As String
Dim btipo As String
On Error GoTo cmd114111_err
   buf = String(42, "-")
   found = formateaa(buf, 45, 2, 0)
   If Len(buf2) > 0 Then
      found = formateaa("       Numero:" & buf2, 28, 2, 0)
   End If
   buf = "     ORDEN DESPACHO " & comanda
   found = formateaa(buf, 28, 2, 0)
   buf = "     Caja :" & caja & " Turno:" & turno
   found = formateaa(buf, 28, 2, 0)
   buf = "Fecha:" & Format(Now, "dd/mm/yyyy") & " " & "Hora :" & Format(Now, "hh:mm:ss")
   found = formateaa(buf, 28, 2, 0)
   If flag_servicio = "A" Then
      found = formateaa("       *** PARA LLEVAR    ***", 28, 2, 0)
      found = formateaa("Nombre:" + Mid$(buf3, 1, 20), 28, 2, 0)
      buf = "Mozo  :"
      found = formateaa(buf, 8, 0, 0)
      found = busca_vendedor_mesero(buvendedor)
   End If
   
   If flag_servicio = "A" Then
      buf = "VENTA RAPIDA"
      found = formateaa(buf, 28, 2, 0)
      
   End If
   If flag_servicio = "C" Then
      buf = "Salon : " & salon & " Mesa:" & mesa
      found = formateaa(buf, 28, 2, 0)
      buf = "Mozo  :"
      found = formateaa(buf, 8, 0, 0)
      found = busca_vendedor_mesero(mesero)
   End If
   If flag_servicio = "D" Then
      found = formateaa("       *** DOMICILIO ***", 28, 2, 0)
      found = formateaa(buf, 28, 2, 0)
      imprime_cliente_delivery "" & codigo
   End If
   If flag_servicio <> "A" And flag_servicio <> "D" And flag_servicio <> "C" Then
      buf = "OTROS SERVICIOS"
      found = formateaa(buf, 28, 2, 0)
   End If
   buf = "///" & xnombre
   found = formateaa(buf, 28, 2, 0)
      
   If buf1 = "***ANULADO***" Then
      found = formateaa("ANULADO", 25, 2, 0)
   End If
   
   buf = String(42, "-")
   found = formateaa(buf, 45, 2, 0)

   found = formateaa("CANT", 6, 0, 0)

   found = formateaa("PRODUCTO ", 21, 0, 0)
   found = formateaa(" ", 1, 2, 0)
 
   buf = String(42, "-")
   found = formateaa(buf, 45, 2, 0)

Exit Sub
cmd114111_err:
  MsgBox "Mensaje,Error en cabecera Pedido " & error$, 48, "Aviso"
  Exit Sub
End Sub
Sub imprime_cliente_delivery(buf1 As String)
Dim mytablex As New ADODB.Recordset
Dim buf As String
Dim found As Integer
   mytablex.Open "SELECT * FROM clientes where codigo='" & buf1 & "'", cn, adOpenDynamic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      'found = formateaa(" *** DOMICILIO ***", 36, 2, 0)
      buf = "Telf:" & "" & mytablex.Fields("codigo")
      found = formateaa(buf, 36, 2, 0)
      buf = "Nomb:" & "" & mytablex.Fields("nombre")
      found = formateaa(buf, 36, 2, 0)
      buf = "Dire:" & "" & mytablex.Fields("direccion")
      found = formateaa(buf, 36, 2, 0)
   End If
   mytablex.Close


End Sub

Function busca_vendedor_mesero(buvendedor As String)
Dim buf As String
Dim found As Integer
'MsgBox buvendedor
Dim mytablex As New ADODB.Recordset
mytablex.Open "SELECT * FROM vendedor where codigo='" & buvendedor & "'", cn, adOpenDynamic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      found = formateaa("", 1, 2, 0)
   End If
   If mytablex.RecordCount > 0 Then
      buf = "" & mytablex.Fields("nombre")
      found = formateaa(buf, 20, 2, 0)
      busca_vendedor_mesero = 1
   End If
   mytablex.Close
 Exit Function
End Function

Sub imprime_detalle_orden()
Dim buf As String
Dim found As Integer
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd4711_err
    '----- formato nuevo
    buf = "" & Data2.Recordset.Fields("cantidad")
    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    'MsgBox ""
       buf = "" & Mid$("" & Data2.Recordset.Fields("descripcio"), 1, 31)
       found = formateaa(buf, 31, 0, 0)
       found = formateaa(" ", 1, 2, 0)
    If Len("" & Data2.Recordset.Fields("descripcio")) > 21 Then
       buf = "      " & Mid$("" & Data2.Recordset.Fields("descripcio"), 32, 31)
       'buf = "" & Mid$("" & Data2.Recordset.Fields("descripcio"), 32, 31)
       found = formateaa(buf, 31, 0, 0)
       found = formateaa(" ", 1, 2, 0)
    End If
    'verificar si tiene receta
    If ve_imprimecombo("" & Data2.Recordset.Fields("producto")) = 1 Then
    mytablex.Open "SELECT * FROM receta where producto='" & "" & Data2.Recordset.Fields("producto") & "'", cn, adOpenDynamic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      Do
      If mytablex.EOF Then Exit Do
      '-------------------------------------------
      found = formateaa("++", 2, 0, 0)
      buf = "" & mytablex.Fields("cantidad")
    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    'MsgBox ""
       buf = "" & Mid$("" & mytablex.Fields("descripcio"), 1, 31)
       found = formateaa(buf, 31, 0, 0)
       found = formateaa(" ", 1, 2, 0)
    If Len("" & mytablex.Fields("descripcio")) > 21 Then
       'buf = "      " & Mid$("" & mytablex.Fields("descripcio"), 32, 31)
       buf = "" & Mid$("" & mytablex.Fields("descripcio"), 32, 31)
       found = formateaa(buf, 31, 0, 0)
       found = formateaa(" ", 1, 2, 0)
    End If
      '-------------------------------------------
      mytablex.MoveNext
      Loop
   End If
   mytablex.Close
   End If
    'found = formateaa("------------------------------------- ", 28, 2, 0)
    If Len("" & Data2.Recordset.Fields("observa1")) > 0 Then
         buf = "*" & Data2.Recordset.Fields("observa1")
         found = formateaa(buf, 28, 2, 0)
  
    End If
    If Len("" & Data2.Recordset.Fields("observa2")) > 0 Then
       buf = "*" & Data2.Recordset.Fields("observa2")
       found = formateaa(buf, 28, 2, 0)
    End If
    If Len("" & Data2.Recordset.Fields("observa3")) > 0 Then
       buf = "*" & Data2.Recordset.Fields("observa3")
       found = formateaa(buf, 28, 2, 0)
    End If
    If Len(Trim("" & Data2.Recordset.Fields("observa4"))) > 0 Then
       found = imprime_combina("" & Data2.Recordset.Fields("producto"))
    End If
    Exit Sub
cmd4711_err:
    MsgBox "Aviso en imprime detalle orden " + error$, 48, "Aviso"
    Exit Sub
    
End Sub

Sub extrae_puertos(temp As String, CAMPO1 As String, CAMPO2 As String, campo3 As String)
Dim i As Integer
Dim j As Integer
i = 0

    'i = InStr(cadena, Parte)
    'If i Then
    'Extrae = Left$(cadena, i - 1) & Mid$(cadena, i + Len(Parte))
    'Else
    'Extrae = cadena
    'End If



Do
   j = InStr(temp, "|")
   If j > 0 Then
      i = i + 1
      Select Case i
             Case 1: CAMPO1 = Mid$(temp, 1, j - 1)
             Case 2: CAMPO2 = Mid$(temp, 1, j - 1)
             Case 3: campo3 = Mid$(temp, 1, j - 1)
                     'MsgBox campo3
      End Select
      temp = Trim$(Mid$(temp, j + 1))
     Else
     Exit Sub
   End If
Loop

End Sub
Sub consulta_comanda(buf2 As String)
Dim found As Integer
Dim buf As String
Dim buf3 As String
On Error GoTo cmd456_err
buf3 = ""
buf2 = Trim(buf2)
If Len(buf2) > 0 Then
   buf3 = " and salon='" & Trim(buf2) & "'"
End If
    If cmytablex.State = 1 Then cmytablex.Close
    cmytablex.Open "SELECT Salon,Mesa,SUM(TOTAL) AS Total,Count(Producto) as C FROM dcomanda where len(salon)>0 and len(mesa)>0 and len(numero)>0 " & buf3 & " group by SALON,MESA ORDER BY SALON,MESA", cn, adOpenDynamic, adLockOptimistic
    Set table2.DataSource = cmytablex
    table2.refresh
    table2.columns(0).Width = 500
    table2.columns(1).Width = 500
    table2.columns(2).Width = 1200
    table2.columns(3).Width = 700
    table2.ForeColor = 1


    suma_comanda
    If cmytablex.RecordCount > 0 Then
    cmytablex.MoveFirst
    
    End If
    'table2.SetFocus
    'table2.Col = 0
    'table2.Row = table2.VisibleRows - 1
    
    'table2.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Aviso consulta comanda " & error$, 48, "Aviso"
    Exit Sub

End Sub
Sub suma_comanda()
Dim sdx As Double
totcoma = ""
If cmytablex.RecordCount = 0 Then Exit Sub
cmytablex.MoveFirst
sdx = 0

Do
If cmytablex.EOF Then Exit Do
sdx = sdx + Val("" & cmytablex.Fields("total"))
cmytablex.MoveNext
Loop
totcoma = Format(sdx, "0.00")
End Sub
Function carga_comanda(sw As Integer)
Dim i As Integer
On Error GoTo cmd890012_err
Dim mytablex As New ADODB.Recordset
If sw = 0 Then
mytablex.Open "SELECT * FROM dcomanda where salon='" & cmytablex.Fields("salon") & "' and mesa='" & cmytablex.Fields("mesa") & "'", cn, adOpenDynamic, adLockOptimistic
End If
If sw = 1 Then
mytablex.Open "SELECT * FROM dcomanda where salon='" & cmytablex.Fields("salon") & "' and mesa='" & cmytablex.Fields("mesa") & "' and cantdev>0", cn, adOpenDynamic, adLockOptimistic
End If
If mytablex.RecordCount = 0 Then
   mytablex.Close
   Exit Function
End If
'-------------------------
Data2.refresh
borrar_campos
sql_detalle
Data2.refresh
cproven = ""
Do
If mytablex.EOF Then Exit Do
If Val("" & mytablex.Fields("cantidad")) > 0 Then
    cproven = mytablex.Fields("vendedor")
    Data2.Recordset.AddNew
    For i = 0 To mytablex.Fields.count - 1
        Data2.Recordset.Fields(i) = mytablex.Fields(i)
    Next i
    If sw = 1 Then
       Data2.Recordset.Fields("cantdev") = 0
       Data2.Recordset.Fields("cantidad") = Val("" & mytablex.Fields("cantdev"))
       Data2.Recordset.Fields("total") = Val("" & mytablex.Fields("cantdev")) * Val("" & mytablex.Fields("precio"))
    End If
    Data2.Recordset.Update
End If
mytablex.MoveNext
Loop
mytablex.Close
carga_comanda = 1
Exit Function
cmd890012_err:
MsgBox "Aviso en Carga Comanda " + error$, 48, "Aviso"
Exit Function
End Function
Sub cabecera_orden_despacho1(buf2 As String, mytablex As ADODB.Recordset)
Dim found As Integer
Dim buf As String
Dim btipo As String
Dim mytable6x As New ADODB.Recordset
On Error GoTo cmd1141111_err
  'MsgBox "xx"
   buf = String(43, "-")
   found = formateaa(buf2, 34, 2, 0)
   found = formateaa(buf, 45, 2, 0)
   buf = "     ORDEN DESPACHO " & "" & mytablex.Fields("numero")
   found = formateaa(buf, 34, 2, 0)
   buf = "     Caja :" & "" & mytablex.Fields("caja") & " Turno:" & "" & mytablex.Fields("turno")
   found = formateaa(buf, 34, 2, 0)
   buf = "  Fecha:" & Format("" & mytablex.Fields("fecha"), "dd/mm/yyyy") & " " & "Hora :" & Format(Now, "hh:mm:ss")
   found = formateaa(buf, 34, 2, 0)
   If "" & mytablex.Fields("servicio") = "A" Then
      found = formateaa("       *** PARA LLEVAR    ***", 25, 2, 0)
      If Len("" & mytablex.Fields("codigo")) > 0 Then
         buf = "CODIGO :" & "" & mytablex.Fields("codigo")
         found = formateaa(buf, 34, 2, 0)
         buf = ""
         mytable6x.Open "SELECT * FROM clientes where codigo='" & mytablex.Fields("codigo") & "'", cn, adOpenDynamic, adLockOptimistic
         If mytable6x.RecordCount > 0 Then
            buf = "" & mytable6x.Fields("nombre")
         End If
         mytable6x.Close
         'buf = "Nom:" & "" & mytablex.Fields("nombreb")
         found = formateaa(buf, 36, 2, 0)
      End If
   End If
   
   If "" & mytablex.Fields("servicio") = "C" Then
      found = formateaa("   *** ATENCION MESA ***", 25, 2, 0)
      buf = "Salon : " & "" & mytablex.Fields("salon") & " Mesa:" & "" & mytablex.Fields("mesa")
      found = formateaa(buf, 34, 2, 0)
      buf = "Mesero:"
      found = formateaa(buf, 8, 0, 0)
      found = busca_vendedor_mesero("" & mytablex.Fields("vendedor"))
   End If
   If "" & mytablex.Fields("servicio") = "D" Then
      found = formateaa("       *** DOMICILIO ***", 25, 2, 0)
      found = formateaa(buf, 34, 2, 0)
      imprime_cliente_delivery "" & mytablex.Fields("codigo")
   End If
   If Val("" & mytablex.Fields("estado")) = 1 Then
      buf = "     ***  ANULADO  ***"
      found = formateaa(buf, 34, 2, 0)
   End If
   
   buf = "///" & xnombre
   found = formateaa(buf, 28, 2, 0)
   

   buf = String(43, "-")
   found = formateaa(buf, 45, 2, 0)

   found = formateaa("CANT", 7, 0, 0)
   
   found = formateaa("PRODUCTO ", 25, 0, 0)
   found = formateaa(" ", 1, 2, 0)

   buf = String(43, "-")
   found = formateaa(buf, 45, 2, 0)

Exit Sub
cmd1141111_err:
  MsgBox "Mensaje,Error en cabecera Pedido " & error$
  Exit Sub

End Sub
Private Sub dlo2342_Click()
Dim found As Integer
'aqui probamos el autoservicio
If dbgrid6.Visible = True Then Exit Sub
If Framefp.Visible = True Then Exit Sub
If Frame7.Visible = True Then Exit Sub
If Frame6.Visible = True Then Exit Sub
If Frame1.Visible = True Then Exit Sub
If Frame2.Visible = True Then Exit Sub
found = sumar_detalle()
If Val(txtotal) > 0 Then
   MsgBox "Tiene Pedido Pendiente", 48, "Aviso"
   DBGrid2.Col = 0
   DBGrid2.Row = DBGrid2.VisibleRows - 1
   DBGrid2.SetFocus
   Exit Sub
End If
borrar_todo
sql_detalle
tiposervicio1 = "Autoservicio"
salon = ""
mesa = ""
mesero = ""
cuenta_separa = ""
flag_servicio = "A"
DBGrid2.SetFocus
End Sub
Sub borrar_data2()
On Error GoTo cmd356_err
ir_primero
Do
If Data2.Recordset.EOF Then Exit Do
Data2.Recordset.Delete
Data2.refresh
Loop
Exit Sub
cmd356_err:
Exit Sub

End Sub
Function servicio_generado(buf As String) As String
Dim mytablex As New ADODB.Recordset
   mytablex.Open "SELECT * FROM servicio where servicio='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      servicio_generado = "" & mytablex.Fields("descripcio")
   End If
   mytablex.Close

End Function
Sub cobra_servicio()
Dim found As Integer
Dim buf As String
If Frame2.Visible = True Then Exit Sub
local1.Visible = False
local1.Visible = False
found = sumar_detalle()
If found = 0 Then
   MsgBox "debe de Existir un Precio=0", 48, "Aviso"
   DBGrid2.SetFocus
   Exit Sub
End If

If Val(txtotal) = 0 Then
   If exisdev <> -10 Then  'si existe devolucion
      DBGrid2.SetFocus
      Exit Sub
   End If
End If
If Trim("" & mytable11.Fields("terminal")) = "T" Or (Val(acuenta) > 0 And Len(petipo) = 0) Then 'pedidos o acuenta>0
          'MsgBox "Hola"
          xruc = codigo
          xnombre = nombre
          Frame7.Visible = True
          habilita_lab7 1
          Framefp.Enabled = False
          If Val(acuenta) > 0 Then
             xtipo = Trim("" & mytable11.Fields("tipope"))
          End If
          xtipo.SetFocus
          Exit Sub
End If
If flag_servicio = "A" Then  'venta rapida
End If
If flag_servicio = "D" Then  'validar el deliveri si ingreso datos
End If
If flag_servicio = "C" Then  'venta mesas
End If
Label36.Caption = "Codigo"
'Frame10.Visible = False
found = proceso_cobros()  'PONE EN CERO TODAS LA FORMAS DE PAGO
opcion2 = 0
'MsgBox dbgrid10.Visible
ttxtotals = Format(Val(rtxtotal), nrodecimal)
ttxtotald = Format(Val(rtxtotald), nrodecimal)
stxtotals = Format(Val(rtxtotal), nrodecimal)
stxtotald = Format(Val(rtxtotald), nrodecimal)
Framefp.Visible = True
Framefp.Enabled = True
habilita_lab7 0

'MsgBox ""
'MsgBox dbgrid10.Enabled
buf = "select * from fpago where fpago='6'"
If mytablefpago.State = 1 Then mytablefpago.Close
mytablefpago.Open buf, cn, adOpenDynamic, adLockOptimistic
Set dbgrid10.DataSource = mytablefpago
dbgrid10.refresh
   If mytablefpago.RecordCount > 0 Then
      mytablefpago.MoveFirst
      dbgrid10.Enabled = False
      dbgrid10.Visible = True
      dbgrid10.Enabled = True
      dbgrid10.SetFocus
      DBGrid10_KeyDown 13, 0
      DBGrid9.Enabled = True
      'Exit Sub
      DBGrid9.SetFocus
      DBGrid9_KeyDown 13, 0
      'xtipo = "7"
      'Else
      'MsgBox "No existe exonerado ", 48, "Aviso"
   End If
   'mytablex.Close

End Sub
Sub carga_tiposdoc(buf As String)
Dim i As Integer
Dim mytablex As New ADODB.Recordset
For i = 0 To 2
    nbxtipo(i) = ""
Next i
i = 0
If buf = "%" Then
   mytablex.Open "select * from tipo where tipodoc='1' or tipodoc='C' or tipodoc='D' or tipodoc='G' ", cn, adOpenDynamic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.Close
      Exit Sub
   End If
   Do
   If mytablex.EOF Then Exit Do
   If "" & mytablex.Fields("tipodoc") = "A" Or "" & mytablex.Fields("tipodoc") = "B" Or "" & mytablex.Fields("tipodoc") = "C" Or "" & mytablex.Fields("tipodoc") = "D" Then
      nbxtipo(i) = "" & mytablex.Fields("tipo")
      i = i + 1
   End If
   If i > 2 Then Exit Do
   mytablex.MoveNext
   Loop
   mytablex.Close
End If
If buf <> "%" Then
   mytablex.Open "select * from tipo where tipodoc='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
   If mytablex.RecordCount = 0 Then
      mytablex.Close
      Exit Sub
   End If
   Do
   If mytablex.EOF Then Exit Do
      nbxtipo(i) = "" & mytablex.Fields("tipo")
      i = i + 1
      If i > 2 Then Exit Do
   mytablex.MoveNext
   Loop
   mytablex.Close
End If

End Sub
Function imprime_combina(buf)
Dim mytablex As New ADODB.Recordset
Dim found As Integer
mytablex.Open "select * from _c" & gusuario & " where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount = 0 Then
   mytablex.Close
   Exit Function
End If

   Do
   If mytablex.EOF Then Exit Do
      '----------------------------------------------
      found = formateaa("*" & mytablex.Fields("descripciop"), 10, 0, 0)
      found = formateaa("" & mytablex.Fields("cantidad"), 3, 2, 0)
      '----------------------------------------------
   mytablex.MoveNext
   Loop
   mytablex.Close
 
End Function
Function existe_fuel(buf As String)
Dim mytablex As New ADODB.Recordset
   mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      If "" & mytablex.Fields("fuel") = "S" Then
         existe_fuel = 1
      End If
   End If
   mytablex.Close

End Function
Sub resuma_precios()
Dim xtivap As Double
Dim tdscto As Double
Dim sdx2 As Double
Dim sdx1 As Double
Dim xtisc As Double
Dim sdx As Double
Data2.Recordset.Fields("neto") = Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & Data2.Recordset.Fields("precio"))
tdscto = Val("" & Data2.Recordset.Fields("neto")) * Val("" & Data2.Recordset.Fields("deslipo")) / 100       'calcular descuento
Data2.Recordset.Fields("descuento") = tdscto  'total descuento
Data2.Recordset.Fields("total") = Val("" & Data2.Recordset.Fields("neto")) - Val("" & Data2.Recordset.Fields("descuento")) 'cobrar
xtivap = Val("" & Data2.Recordset.Fields("total")) * Val("" & Data2.Recordset.Fields("ivap")) / 100
Data2.Recordset.Fields("tivap") = xtivap
   sdx2 = 1 + Val("" & Data2.Recordset.Fields("igv")) / 100
   sdx1 = Val(Data2.Recordset.Fields("total")) / sdx2
   Data2.Recordset.Fields("subtotal") = sdx1  'subtotal
   sdx = Val("" & Data2.Recordset.Fields("total")) - Val("" & Data2.Recordset.Fields("subtotal"))
   Data2.Recordset.Fields("impuesto") = sdx  'impuesto
   xtisc = Val("" & Data2.Recordset.Fields("subtotal")) * Val("" & Data2.Recordset.Fields("isc")) / 100
   Data2.Recordset.Fields("tisc") = xtisc
   Data2.Recordset.Fields("tax") = 0
   If Val("" & Data2.Recordset.Fields("igv")) = 0 Then
      Data2.Recordset.Fields("tax") = Val("" & Data2.Recordset.Fields("total"))
      Data2.Recordset.Fields("impuesto") = 0
   End If
Exit Sub

End Sub
Function ve_imprimecombo(buf As String)
Dim mytablex As New ADODB.Recordset
mytablex.Open "SELECT recetaprn FROM producto where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   If "" & mytablex.Fields("recetaprn") = "S" Then
      ve_imprimecombo = 1
   End If
End If
mytablex.Close
End Function
Function suma_pedidos(buf As String, buf1 As String, buf2 As String, buf3 As String) As String
Dim mytablex As New ADODB.Recordset
Dim sdx As Double
sdx = 0
'MsgBox "SELECT * FROM cpedidov where  codigo='" & buf & "' and tipo='" & buf1 & "' and serie='" & buf2 & "' and numero='" & buf3 & "'"
mytablex.Open "SELECT * FROM cpedidov where  codigo='" & buf & "' and tipo='" & buf1 & "' and serie='" & buf2 & "' and numero='" & buf3 & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   sdx = Val("" & mytablex.Fields("total")) - Val("" & mytablex.Fields("acuenta"))
End If
mytablex.Close
suma_pedidos = "" & sdx
End Function
Sub graba_acumulado_clientes(mytabley As ADODB.Recordset, signo As Double, sumador As Double)
Dim mytablex As New ADODB.Recordset
Dim sdx As Double
mytablex.Open "SELECT * FROM cpedidov where  codigo='" & mytabley.Fields("codigo") & "' and tipo='" & "" & mytabley.Fields("orden") & "' and serie='" & "" & mytabley.Fields("observa") & "' and numero='" & "" & mytabley.Fields("dias") & "'", cn, adOpenKeyset, adLockOptimistic
If mytablex.RecordCount > 0 Then
      sdx = Val("" & mytablex.Fields("acuenta")) + signo * sumador
      mytablex.Fields("acuenta") = sdx
      mytablex.Update
End If
mytablex.Close
End Sub
Function verifica_fpago(buf As String) As String
Dim mytablex As New ADODB.Recordset
mytablex.Open "SELECT * FROM fpago where  fpago='" & buf & "'", cn, adOpenKeyset, adLockOptimistic
If mytablex.RecordCount > 0 Then
   verifica_fpago = "" & mytablex.Fields("tipo")
End If
mytablex.Close
End Function
Sub inicia_color_familia()
Dim i As Integer
Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer
Dim mytablex As New ADODB.Recordset
mytablex.Open "SELECT * FROM paramecacolor where  caja='" & "" & mytable11.Fields("caja") & "'", cn, adOpenKeyset, adLockOptimistic
If mytablex.RecordCount > 0 Then

I1 = Val("" & mytablex.Fields("colorfamilia1"))
I2 = Val("" & mytablex.Fields("colorfamilia2"))
I3 = Val("" & mytablex.Fields("colorfamilia3"))

'Exit Sub
For i = 0 To 17
    'zfamilia(i).BackColor = 200
    zfamilia(i).BackColor = RGB(I1, I2, I3)
Next i
End If
mytablex.Close
End Sub
Sub inicia_color_producto()
Dim i As Integer
Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer
Dim I4 As String
Dim mytablex As New ADODB.Recordset
mytablex.Open "SELECT * FROM paramecacolor where  caja='" & "" & mytable11.Fields("caja") & "'", cn, adOpenKeyset, adLockOptimistic
If mytablex.RecordCount > 0 Then

I1 = Val("" & mytablex.Fields("colorproducto1"))
I2 = Val("" & mytablex.Fields("colorproducto2"))
I3 = Val("" & mytablex.Fields("colorproducto3"))
I4 = Val("" & mytablex.Fields("size"))

If I4 < 8 Then
   I4 = 9
End If
'Exit Sub
For i = 0 To 29
    'zfamilia(i).BackColor = 200
    zproducto(i).BackColor = RGB(I1, I2, I3)
    zproducto(i).FontSize = I4
    
Next i
End If
mytablex.Close
End Sub
Sub carga_grafico(buf As String)
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd909012_err
fotoimagen = LoadPicture()
mytablex.Open "SELECT fotonombre FROM producto where  producto='" & buf & "'", cn, adOpenKeyset, adLockOptimistic
If mytablex.RecordCount > 0 Then
   If existe_archivo("" & mytablex.Fields("fotonombre")) > 0 Then
      fotoimagen = LoadPicture("" & mytablex.Fields("fotonombre"))
   End If
End If
Exit Sub
cmd909012_err:
Exit Sub
End Sub
'Sub carga_cobranza()
'Dim mytablex As New ADODB.Recordset
'Dim i As Integer
'f or i = 0 To 20
'    mcobcod(i) = ""
'    wcobcod(i) = ""
'Next i

'i = -1
'mytablex.Open "select Descripcio,Fpago,Tipo,Moneda from fpago where vecaja='S'", cn, adOpenStatic, adLockOptimistic

'Do
'If mytablex.EOF Then Exit Do
'i = i + 1
'mcobcod(i) = "" & mytablex.Fields("descripcio")
'wcobcod(i) = "" & mytablex.Fields("fpago")
'If i > 20 Then Exit Do
'mytablex.MoveNext
'Loop
'mcobtop = i
'mytablex.Close
'mcobpag = 0
'menu_cobranza "INI"

'End Sub
'Sub menu_cobranza(buf As String)
'Dim i As Integer
'Dim j As Integer
'Select Case buf
'       Case "INI"
'            mcobpag = 0
'       Case "SIG"
'            mcobpag = mcobpag + 3
'            If mcobpag > 102 Then
'               mcobpag = 0
'            End If
'       Case "ANT"
'            mcobpag = mcobpag - 3
'            If mcobpag < 0 Then
'               mcobpag = 0
'            End If
'End Select
'j = -1
'For i = mcobpag To 3 + mcobpag
'    j = j + 1
'    tmediopa(j).Caption = mcobcod(i)
'    wwcobcod(j) = wcobcod(i)
'Next i
'
'End Sub
Sub menu_fin_tallas()
Dim found As Integer
calcula_igv 0
found = sumar_detalle()
      DBGrid2.Col = 0
      DBGrid2.Row = DBGrid2.VisibleRows - 1
      DBGrid2.SetFocus
Command3_Click

End Sub
Sub inicia_color_comandos()
Dim a As String
Dim i As Integer
Dim I1 As Integer
Dim I2 As Integer
Dim I3 As Integer
Dim mytablex As New ADODB.Recordset
mytablex.Open "SELECT * FROM paramecacolor where  caja='" & "" & mytable11.Fields("caja") & "'", cn, adOpenKeyset, adLockOptimistic
If mytablex.RecordCount > 0 Then
    For i = 0 To 23
        a = "C" & i
        I1 = Val("" & mytablex.Fields(a))
        a = "d" & i
        I2 = Val("" & mytablex.Fields(a))
        a = "e" & i
        I3 = Val("" & mytablex.Fields(a))
        xopciones(i).BackColor = RGB(I1, I2, I3)
        a = "f" & i
        If Trim("" & mytablex.Fields(a)) = "N" Then
           xopciones(i).Enabled = False
        End If
    Next i
End If
mytablex.Close
End Sub
Function verifica_saldo_receta(buf1 As String, cant As Double)
Dim mytablex As New ADODB.Recordset
Dim sdx As Double
Dim sw As Integer
Dim buf As String
sw = 0
'MsgBox buf
buf = "SELECT dbo.RECETA.PRODUCTOI, dbo.ALMACEN.SALDO "
buf = buf & "FROM dbo.RECETA INNER JOIN "
buf = buf & " dbo.ALMACEN on dbo.RECETA.PRODUCTOI = dbo.ALMACEN.PRODUCTO"
buf = buf & " and dbo.receta.producto='" & buf1 & "'"
'MsgBox buf
mytablex.Open buf, cn, adOpenKeyset, adLockOptimistic
If mytablex.RecordCount = 0 Then
   mytablex.Close
   Exit Function
End If

sdx = 1
Do
If mytablex.EOF Then Exit Do
If cant >= Val("" & mytablex.Fields(1)) Then
   sw = 2
   Exit Do
End If
mytablex.MoveNext
Loop
mytablex.Close
verifica_saldo_receta = sw
End Function
Sub carga_dcvendedor()
Dim mytablex As New ADODB.Recordset
dcvendedor.Clear
dcvendedor.AddItem "%"
mytablex.Open "select * from vendedor order by nombre", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
dcvendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
mytablex.MoveNext
Loop
mytablex.Close
dcvendedor.ListIndex = 0
End Sub
Function orden_despachot()
Dim buf As String
Dim X As Integer
Dim oldprinter
Dim found As Integer
Dim mytablex As New ADODB.Recordset
On Error GoTo cmd47113_err
If Len(Trim("" & mytable11.Fields("copiaod"))) = 0 Then Exit Function
FileName = caja & "LP"
found = borra_nombre("" & FileName)
cerrar_archivo
ncanal = 1
Open FileName For Append As #ncanal
Data2.refresh
cabecera_orden_despacho "" & Data2.Recordset.Fields("vendedor"), "", "", ""
Do
If Data2.Recordset.EOF Then Exit Do
    '----- formato nuevo
    buf = "" & Data2.Recordset.Fields("cantidad")
    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    'MsgBox ""
       buf = "" & Mid$("" & Data2.Recordset.Fields("descripcio"), 1, 31)
       found = formateaa(buf, 31, 0, 0)
       found = formateaa(" ", 1, 2, 0)
    If Len("" & Data2.Recordset.Fields("descripcio")) > 21 Then
       buf = "      " & Mid$("" & Data2.Recordset.Fields("descripcio"), 32, 31)
       'buf = "" & Mid$("" & Data2.Recordset.Fields("descripcio"), 32, 31)
       found = formateaa(buf, 31, 0, 0)
       found = formateaa(" ", 1, 2, 0)
    End If
    'verificar si tiene receta
    If ve_imprimecombo("" & Data2.Recordset.Fields("producto")) = 1 Then
    mytablex.Open "SELECT * FROM receta where producto='" & "" & Data2.Recordset.Fields("producto") & "'", cn, adOpenDynamic, adLockOptimistic
   If mytablex.RecordCount > 0 Then
      Do
      If mytablex.EOF Then Exit Do
      '-------------------------------------------
      found = formateaa("++", 2, 0, 0)
      buf = "" & mytablex.Fields("cantidad")
    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    'MsgBox ""
       buf = "" & Mid$("" & mytablex.Fields("descripcio"), 1, 31)
       found = formateaa(buf, 31, 0, 0)
       found = formateaa(" ", 1, 2, 0)
    If Len("" & mytablex.Fields("descripcio")) > 21 Then
       'buf = "      " & Mid$("" & mytablex.Fields("descripcio"), 32, 31)
       buf = "" & Mid$("" & mytablex.Fields("descripcio"), 32, 31)
       found = formateaa(buf, 31, 0, 0)
       found = formateaa(" ", 1, 2, 0)
    End If
      '-------------------------------------------
      mytablex.MoveNext
      Loop
   End If
   mytablex.Close
   End If
    'found = formateaa("------------------------------------- ", 28, 2, 0)
    If Len("" & Data2.Recordset.Fields("observa1")) > 0 Then
         buf = "*" & Data2.Recordset.Fields("observa1")
         found = formateaa(buf, 28, 2, 0)
    End If
    If Len("" & Data2.Recordset.Fields("observa2")) > 0 Then
       buf = "*" & Data2.Recordset.Fields("observa2")
       found = formateaa(buf, 28, 2, 0)
    End If
    If Len("" & Data2.Recordset.Fields("observa3")) > 0 Then
       buf = "*" & Data2.Recordset.Fields("observa3")
       found = formateaa(buf, 28, 2, 0)
    End If
    If Len(Trim("" & Data2.Recordset.Fields("observa4"))) > 0 Then
       found = imprime_combina("" & Data2.Recordset.Fields("producto"))
    End If
Data2.Recordset.MoveNext
Loop
For X = 1 To 5
      Print #ncanal, ""
      Next X
      Print #ncanal, ""
Close #ncanal
cerrar_archivo
'MsgBox "xxxx"
'---impresion del archivo------------
If existearchivo(FileName) = 1 Then
         oldprinter = Printer.DeviceName
         'MsgBox "xx"
         selecciona_impresoras ("" & mytable11.Fields("copiaod"))
         'MsgBox "xxx"
         found = Imprime_archivojj(FileName, 0, "" & mytable11.Fields("tamanorden"), "" & mytable11.Fields("nombrefont"))
         'MsgBox "xxx"
         selecciona_impresoras (Trim(oldprinter))
         'MsgBox "xabc"
End If
found = borra_nombre(FileName)
Exit Function
cmd47113_err:
    MsgBox "Aviso en orden despachot " + error$, 48, "Aviso"
    Exit Function
End Function
Sub carga_minimo(buf As String)
Dim mytablex As New ADODB.Recordset
stkminimo = ""
If "" & mytable11.Fields("stkminimo") = "S" Then
            mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
            If mytablex.RecordCount > 0 Then
               If familia_saldo("" & mytablex.Fields("familia")) = 0 Then
                  consulta_minimo buf, "" & mytablex.Fields("minimo")
               End If
            End If
 mytablex.Close
End If

End Sub
Function credito_habilitado(buf As String)
Dim mytablex As New ADODB.Recordset
Dim buf1 As String
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM clientes where  codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   If "" & mytablex.Fields("estadocredito") = "S" Then
      credito_habilitado = 1
   End If
End If
mytablex.Close

End Function
Function busca_credito_credito(buf As String, buf2 As String)
Dim mytablex As New ADODB.Recordset
Dim sdx As Double
Dim found As Integer
Dim buf1 As String
sdx = 0
If mytablex.State = 1 Then mytablex.Close
mytablex.Open "SELECT * FROM cuentac  where  tipoclie='C' and codigo='" & buf2 & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   Do
   If mytablex.EOF Then Exit Do
     If Val("" & mytablex.Fields("saldo")) > 0 Then
     If "" & mytablex.Fields("grupo") = buf Then
        sdx = sdx + Val("" & mytablex.Fields("saldo"))
     End If
     End If
   mytablex.MoveNext
   Loop
End If
mytablex.Close
saldoabo = Format(sdx, "0.00")
mytablex.Open "SELECT * FROM clientes where  codigo='" & buf2 & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   mytablex.Fields("credito_usado") = sdx
   mytablex.Update
End If
mytablex.Close

End Function
Function saldo_clientes(buf As String, deltax As Double) As Double
Dim sdx As Double
Dim mytablex As New ADODB.Recordset
sdx = 0
mytablex.Open "SELECT * FROM clientes where  codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
If mytablex.RecordCount > 0 Then 'si existe
   sdx = Val("" & mytablex.Fields("credito")) - (Val("" & mytablex.Fields("credito_usado")) + deltax)
End If
saldo_clientes = sdx
End Function
Sub carga_clasificacion()
Dim mytablex As New ADODB.Recordset
coclasifica.Clear
coclasifica.AddItem "%"
mytablex.Open "select * from clasifi ", cn, adOpenStatic, adLockOptimistic
Do
If mytablex.EOF Then Exit Do
coclasifica.AddItem "" & mytablex.Fields("clasifica") & "|" & mytablex.Fields("descripcio")
mytablex.MoveNext
Loop
mytablex.Close
coclasifica.ListIndex = 0

End Sub
Sub saludo_cumpe()
Dim dd As String
Dim mm As String

Dim ddd As String
Dim mmm As String
felizc = ""

 If Not IsDate(fechanac) Then
   Exit Sub
   End If

dd = Format(Day(Now), "00")
mm = Format(Month(Now), "00")


ddd = Format(Day(fechanac), "00")
mmm = Format(Month(fechanac), "00")

  

   
   
   If dd = ddd And mm = mmm Then
      felizc = "FELIZ CUMPLEAÑOS "
   End If
   

End Sub
Sub imprime_detalle_orden1(mytablex As ADODB.Recordset)
Dim buf As String
Dim found As Integer
Dim mytabley As New ADODB.Recordset
On Error GoTo cmd90003_err
    '----- formato nuevo
    If Val("" & mytablex.Fields("estado")) = 1 Then
       buf = "-" & mytablex.Fields("cantidad")
       Else
       buf = "" & mytablex.Fields("cantidad")
    End If
    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
       buf = "" & Mid$("" & mytablex.Fields("descripcio"), 1, 31)
       found = formateaa(buf, 31, 0, 0)
       found = formateaa(" ", 1, 2, 0)
    If Len("" & mytablex.Fields("descripcio")) > 21 Then
       buf = "      " & Mid$("" & mytablex.Fields("descripcio"), 32, 31)
       'buf = "" & Mid$("" & mytablex.Fields("descripcio"), 32, 31)
       found = formateaa(buf, 31, 0, 0)
       found = formateaa(" ", 1, 2, 0)
    End If
    
    '--------------------------------------------------------
    If ve_imprimecombo("" & mytablex.Fields("producto")) = 1 Then
    mytabley.Open "SELECT * FROM receta where producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenDynamic, adLockOptimistic
   If mytabley.RecordCount > 0 Then
      Do
      If mytabley.EOF Then Exit Do
      '-------------------------------------------
      found = formateaa("++", 1, 0, 0)
      buf = "" & mytabley.Fields("cantidad")
    found = formateaa(buf, 5, 0, 0)
    found = formateaa("", 1, 0, 0)
    'MsgBox ""
       buf = "" & Mid$("" & mytabley.Fields("descripcio"), 1, 31)
       found = formateaa(buf, 31, 0, 0)
       found = formateaa(" ", 1, 2, 0)
    If Len("" & mytabley.Fields("descripcio")) > 21 Then
       buf = "      " & Mid$("" & mytabley.Fields("descripcio"), 32, 31)
       found = formateaa(buf, 31, 0, 0)
       found = formateaa(" ", 1, 2, 0)
    End If
      '-------------------------------------------
      mytabley.MoveNext
      Loop
   End If
   mytabley.Close
   End If
   
    '--------------------------------------------------------

    'found = formateaa("------------------------------------- ", 28, 2, 0)
    If Len("" & mytablex.Fields("observa1")) > 0 Then
       'buf = "*" & mytablex.fields("observa")
       'debe descriminar los productos
       'found = verifica_receta_flag("" & mytablex.fields("observa"), 0)
       'found = imprime_combina("" & mytablex.Fields("caja"), "" & mytablex.Fields("producto"))
       'If found = 0 Then
       'found = formateaa(buf, 28, 2, 0)
       'End If
    End If
    If Len("" & mytablex.Fields("observa1")) > 0 Then
       buf = "*" & mytablex.Fields("observa1")
       found = formateaa(buf, 28, 2, 0)
    End If
    If Len("" & mytablex.Fields("observa2")) > 0 Then
       buf = "*" & mytablex.Fields("observa2")
       found = formateaa(buf, 28, 2, 0)
    End If
    If Len(Trim("" & mytablex.Fields("observa4"))) > 0 Then
       found = imprime_combina("" & mytablex.Fields("producto"))
    End If
    'If "" & mytablex.Fields("isla") = "1" Then
    '   buf = "Consumo"
    '   found = formateaa(buf, 28, 2, 0)
    'End If
    suma1 = suma1 + Val("" & mytablex.Fields("cantidad"))
    Exit Sub
cmd90003_err:
    MsgBox "Aviso en imprime detalle orden 1" + error$, 48, "Aviso"
    Exit Sub
End Sub
Sub cargar_grafico20()
On Error GoTo cmd7779_err
IMAGE11.Picture = LoadPicture(globalpath & "\ico\ORION.jpg")
Exit Sub
cmd7779_err:
'MsgBox " Carga Grafico:" & error$
Exit Sub
End Sub
Sub proceso_cierre_efectivo()
Dim found As Integer
Dim buf As String


If Frame2.Visible = True Then Exit Sub
local1.Visible = False
local1.Visible = False
found = sumar_detalle()
If found = 0 Then
   MsgBox "debe de Existir un Precio=0", 48, "Aviso"
   DBGrid2.SetFocus
   Exit Sub
End If

If Val(txtotal) = 0 Then
   If exisdev <> -10 Then  'si existe devolucion
      DBGrid2.SetFocus
      Exit Sub
   End If
End If
carga_tiposdoc "%"
If Trim("" & mytable11.Fields("terminal")) = "T" Or (Val(acuenta) > 0 And Len(petipo) = 0) Then 'pedidos o acuenta>0
          'MsgBox "Hola"
          xruc = codigo
          xnombre = nombre
          Frame7.Visible = True
          habilita_lab7 1
          Framefp.Enabled = False
          If Val(acuenta) > 0 Then
             xtipo = Trim("" & mytable11.Fields("tipope"))
          End If
          xtipo.SetFocus
          Exit Sub
End If
If flag_servicio = "A" Then  'venta rapida
End If
If flag_servicio = "D" Then  'validar el deliveri si ingreso datos
End If
If flag_servicio = "C" Then  'venta mesas
End If
Label36.Caption = "Codigo"
'Frame10.Visible = False
found = proceso_cobros()  'PONE EN CERO TODAS LA FORMAS DE PAGO
opcion2 = 0
'MsgBox dbgrid10.Visible
ttxtotals = Format(Val(rtxtotal), nrodecimal)
ttxtotald = Format(Val(rtxtotald), nrodecimal)
stxtotals = Format(Val(rtxtotal), nrodecimal)
stxtotald = Format(Val(rtxtotald), nrodecimal)
Framefp.Visible = True
Framefp.Enabled = True
habilita_lab7 0
'MsgBox ""
'MsgBox dbgrid10.Enabled
buf = "select * from fpago where fpago='1'"
If mytablefpago.State = 1 Then mytablefpago.Close
mytablefpago.Open buf, cn, adOpenDynamic, adLockOptimistic
Set dbgrid10.DataSource = mytablefpago
dbgrid10.refresh
   If mytablefpago.RecordCount > 0 Then
      mytablefpago.MoveFirst
      dbgrid10.Enabled = False
      dbgrid10.Visible = True
      dbgrid10.Enabled = True
      dbgrid10.SetFocus
      DBGrid10_KeyDown 13, 0
      DBGrid9.Enabled = True
      'Exit Sub
      DBGrid9.SetFocus
      DBGrid9_KeyDown 13, 0
      RGPAGO_KeyPress 13
      xtipo = "1"
      xtipo_keyPress 13
      'RGPAGO.SetFocus
      'xtipo = "7"
      'Else
      'MsgBox "No existe exonerado ", 48, "Aviso"
   End If
   'mytablex.Close
End Sub
Function acura_lectura() As String
Dim d As Integer
Dim i As Integer
Dim buf As String
Select Case "" & mytable11.Fields("portbala")
           Case "COM1"
           d = 1
           Case "COM2"
           d = 2
           Case "COM3"
           d = 3
           Case "COM4"
           d = 4
           Case "COM5"
           d = 5
End Select
MSComm1.CommPort = d
MSComm1.Settings = "9600,n,8,1"
MSComm1.InputLen = 10
MSComm1.PortOpen = True
buf = ""
Do
DoEvents
buf = buf & MSComm1.Input
Loop Until Len(buf) >= 10
MSComm1.PortOpen = False
acura_lectura = Mid$(buf, Len(buf) - 7, 6)

End Function













