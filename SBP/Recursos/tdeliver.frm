VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tdeliver 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VisiOrion"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   17175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   17175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   17400
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   212
      Top             =   2880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   17520
      Top             =   600
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00808080&
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
      ForeColor       =   &H8000000E&
      Height          =   9375
      Left            =   0
      TabIndex        =   196
      Top             =   0
      Visible         =   0   'False
      Width           =   12135
      Begin VB.ComboBox dcvendedor 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7080
         Style           =   2  'Dropdown List
         TabIndex        =   200
         Top             =   2400
         Visible         =   0   'False
         Width           =   3975
      End
      Begin VB.TextBox ddvendedor 
         Enabled         =   0   'False
         Height          =   495
         Left            =   9000
         MaxLength       =   11
         TabIndex        =   199
         Top             =   1920
         Visible         =   0   'False
         Width           =   2055
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   198
         Top             =   360
         Width           =   3375
      End
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
         Left            =   11280
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tdeliver.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   197
         ToolTipText     =   "Borrar registro"
         Top             =   240
         Width           =   735
      End
      Begin MSDBGrid.DBGrid DBGrid4 
         CausesValidation=   0   'False
         Height          =   4815
         Left            =   120
         OleObjectBlob   =   "tdeliver.frx":1212
         TabIndex        =   201
         Top             =   960
         Width           =   6855
      End
      Begin MSDataGridLib.DataGrid dbgrid7 
         Height          =   2535
         Left            =   120
         TabIndex        =   202
         Top             =   6480
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4471
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
      Begin VB.Label seccion 
         BackColor       =   &H00C0FFFF&
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
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   9000
         TabIndex        =   210
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label35 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   7080
         TabIndex        =   209
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
         Enabled         =   0   'False
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
         Height          =   495
         Left            =   7080
         TabIndex        =   205
         Top             =   1920
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label tproducto 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   420
         Left            =   3600
         TabIndex        =   204
         Top             =   360
         Width           =   135
      End
      Begin VB.Image foto 
         BorderStyle     =   1  'Fixed Single
         Height          =   3015
         Left            =   6960
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   4095
      End
      Begin VB.Label descorto 
         BackColor       =   &H00FF0000&
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
         TabIndex        =   203
         Top             =   6000
         Width           =   6735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Consulta"
      Height          =   9615
      Left            =   0
      TabIndex        =   188
      Top             =   0
      Visible         =   0   'False
      Width           =   17175
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
         Left            =   3600
         MaxLength       =   30
         TabIndex        =   191
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
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
         Left            =   6480
         TabIndex        =   190
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   189
         TabStop         =   0   'False
         Top             =   240
         Width           =   3255
      End
      Begin MSDataGridLib.DataGrid dbgrid6 
         Height          =   7935
         Left            =   0
         TabIndex        =   192
         Top             =   840
         Visible         =   0   'False
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   13996
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
            MarqueeStyle    =   1
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7935
         Left            =   0
         TabIndex        =   193
         Top             =   840
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   13996
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   2
         RowHeight       =   24
         TabAction       =   2
         RowDividerStyle =   6
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   10320
         TabIndex        =   195
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label label56 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   0
         TabIndex        =   194
         Top             =   9000
         Width           =   15105
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
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
      Height          =   9135
      Left            =   -7080
      TabIndex        =   172
      Top             =   7560
      Visible         =   0   'False
      Width           =   15375
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
         TabIndex        =   178
         Top             =   2400
         Width           =   1695
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
         TabIndex        =   177
         Top             =   480
         Width           =   1815
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
         TabIndex        =   176
         Top             =   480
         Width           =   2055
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
         TabIndex        =   175
         Top             =   960
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
         MaxLength       =   200
         TabIndex        =   174
         Top             =   1440
         Width           =   8295
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
         TabIndex        =   173
         Top             =   1920
         Width           =   8295
      End
      Begin VB.Label clasificacion 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   2640
         TabIndex        =   187
         Top             =   2880
         Width           =   6495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"tdeliver.frx":2275
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   186
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   185
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label command10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cerrar"
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
         Left            =   11520
         TabIndex        =   184
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label command12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Limpia Campos"
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
         Left            =   11520
         TabIndex        =   183
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label command11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Grabar"
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
         Left            =   11520
         TabIndex        =   182
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Nacimiento                        ClasificacionCliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   181
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Modifica"
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
         Left            =   11520
         TabIndex        =   180
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Crear"
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
         Left            =   11520
         TabIndex        =   179
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00808080&
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
      ForeColor       =   &H00FF0000&
      Height          =   4935
      Left            =   0
      TabIndex        =   154
      Top             =   1680
      Visible         =   0   'False
      Width           =   12735
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
         MaxLength       =   2
         TabIndex        =   165
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox xvendedor 
         Height          =   495
         Left            =   2160
         MaxLength       =   11
         TabIndex        =   164
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
         TabIndex        =   163
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox xnombre 
         Height          =   495
         Left            =   2160
         MaxLength       =   60
         TabIndex        =   162
         Top             =   1680
         Width           =   5415
      End
      Begin VB.TextBox xdireccion 
         Height          =   495
         Left            =   2160
         MaxLength       =   200
         TabIndex        =   161
         Top             =   2160
         Width           =   5415
      End
      Begin VB.TextBox xdistrito 
         Height          =   495
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   160
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox xnumero 
         Enabled         =   0   'False
         Height          =   495
         Left            =   2160
         MaxLength       =   11
         TabIndex        =   159
         Top             =   3600
         Width           =   2175
      End
      Begin VB.TextBox xserie 
         Enabled         =   0   'False
         Height          =   495
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   158
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
         TabIndex        =   157
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   8040
         Picture         =   "tdeliver.frx":2322
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   240
         Width           =   1470
      End
      Begin VB.CommandButton Command14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   8040
         Picture         =   "tdeliver.frx":2BEC
         Style           =   1  'Graphical
         TabIndex        =   155
         ToolTipText     =   "Imprimir todo"
         Top             =   1320
         Width           =   1470
      End
      Begin VB.Label Label39 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre"
         Height          =   495
         Left            =   240
         TabIndex        =   214
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label36 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
         Height          =   495
         Left            =   240
         TabIndex        =   213
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Documento                           Vendedor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   171
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label38 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion                                               Glosa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   170
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label48 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie                                                                               Numero"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   169
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label ntipox 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   4320
         TabIndex        =   168
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label nvendedorx 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   4320
         TabIndex        =   167
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label ordentrabajo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4440
         TabIndex        =   166
         Top             =   3600
         Width           =   105
      End
   End
   Begin VB.Frame Framefp 
      BackColor       =   &H00808080&
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
      Height          =   9015
      Left            =   0
      TabIndex        =   118
      Top             =   0
      Visible         =   0   'False
      Width           =   15255
      Begin VB.Frame Frame6 
         BackColor       =   &H00808080&
         Caption         =   "Entrega"
         Height          =   3855
         Left            =   5880
         TabIndex        =   136
         Top             =   1920
         Visible         =   0   'False
         Width           =   6855
         Begin VB.TextBox tcampo6 
            Height          =   375
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   142
            Top             =   2640
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox tcampo5 
            Height          =   375
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   141
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox tcampo4 
            Height          =   375
            Left            =   1680
            MaxLength       =   30
            TabIndex        =   140
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox tcampo3 
            Height          =   375
            Left            =   1680
            MaxLength       =   30
            TabIndex        =   139
            Top             =   1560
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox tcampo2 
            Height          =   375
            Left            =   1680
            MaxLength       =   60
            TabIndex        =   138
            Top             =   1200
            Width           =   3495
         End
         Begin VB.TextBox tcampo1 
            Height          =   375
            Left            =   1680
            MaxLength       =   11
            TabIndex        =   137
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label acufp 
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   240
            TabIndex        =   153
            Top             =   3120
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
            Left            =   240
            TabIndex        =   152
            Top             =   2640
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label saldoabo 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   3480
            TabIndex        =   151
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label fpmoneda 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4680
            TabIndex        =   150
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label fpago 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   4680
            TabIndex        =   149
            Top             =   240
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label descripcio5 
            BackColor       =   &H00C0C0C0&
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
            Left            =   240
            TabIndex        =   148
            Top             =   2280
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label descripcio4 
            BackColor       =   &H00C0C0C0&
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
            Left            =   240
            TabIndex        =   147
            Top             =   1920
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label descripcio3 
            BackColor       =   &H00C0C0C0&
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
            Left            =   240
            TabIndex        =   146
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Label descripcio2 
            BackColor       =   &H00C0C0C0&
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
            Left            =   240
            TabIndex        =   145
            Top             =   1200
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label descripcio1 
            BackColor       =   &H00C0C0C0&
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
            Left            =   240
            TabIndex        =   144
            Top             =   840
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label totpedido 
            BackColor       =   &H00C0C0C0&
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
            Left            =   1680
            TabIndex        =   143
            Top             =   3000
            Width           =   1575
         End
      End
      Begin VB.CommandButton COMMAND6 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   120
         Picture         =   "tdeliver.frx":34B6
         Style           =   1  'Graphical
         TabIndex        =   119
         ToolTipText     =   "Imprimir todo"
         Top             =   240
         Width           =   1335
      End
      Begin MSDBGrid.DBGrid DBGrid9 
         Bindings        =   "tdeliver.frx":3D80
         Height          =   4575
         Left            =   5880
         OleObjectBlob   =   "tdeliver.frx":3D94
         TabIndex        =   120
         Top             =   2280
         Width           =   6975
      End
      Begin MSDataGridLib.DataGrid dbgrid10 
         Height          =   6375
         Left            =   120
         TabIndex        =   121
         Top             =   2280
         Width           =   5655
         _ExtentX        =   9975
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
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T/C"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6120
         TabIndex        =   135
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label paridadfp 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6720
         TabIndex        =   134
         Top             =   1200
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
         Left            =   5880
         TabIndex        =   133
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
         TabIndex        =   132
         Top             =   1920
         Width           =   5655
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
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
         Left            =   5880
         TabIndex        =   131
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
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
         Left            =   5880
         TabIndex        =   130
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
         Left            =   8040
         TabIndex        =   129
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
         Left            =   8520
         TabIndex        =   128
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
         Left            =   8040
         TabIndex        =   127
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
         Left            =   8520
         TabIndex        =   126
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
         Left            =   8040
         TabIndex        =   125
         Top             =   6840
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
         Left            =   8520
         TabIndex        =   124
         Top             =   6840
         Width           =   4335
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
         Left            =   8040
         TabIndex        =   123
         Top             =   7680
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
         Left            =   8520
         TabIndex        =   122
         Top             =   7680
         Width           =   4335
      End
   End
   Begin VB.TextBox correo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      MaxLength       =   80
      TabIndex        =   117
      Top             =   1200
      Width           =   3735
   End
   Begin VB.ComboBox clasesunat 
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
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   115
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox pedido 
      Height          =   375
      Left            =   360
      MaxLength       =   11
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   12120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00808080&
      Caption         =   "Digite un Nombre - CONGELA PEDIDOS INGRESADOS"
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
      Height          =   3855
      Left            =   5160
      TabIndex        =   92
      Top             =   5760
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox clavecongela 
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   207
         TabStop         =   0   'False
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox xcongelax 
         Height          =   615
         Left            =   240
         MaxLength       =   12
         TabIndex        =   95
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
         Picture         =   "tdeliver.frx":4C67
         Style           =   1  'Graphical
         TabIndex        =   94
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
         Picture         =   "tdeliver.frx":5415
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave"
         Height          =   495
         Left            =   240
         TabIndex        =   208
         Top             =   1440
         Width           =   2415
      End
   End
   Begin VB.ComboBox crucefa 
      Height          =   315
      Left            =   16320
      Style           =   2  'Dropdown List
      TabIndex        =   85
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox acuenta 
      BackColor       =   &H00808080&
      Height          =   375
      Left            =   8880
      MaxLength       =   10
      TabIndex        =   82
      Top             =   8760
      Width           =   6375
   End
   Begin VB.Data Data9 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   11280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   2520
      TabIndex        =   66
      Top             =   2280
      Visible         =   0   'False
      Width           =   8775
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
         Left            =   7680
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tdeliver.frx":5BC3
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Grabar registro"
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton Command9 
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
         Left            =   7680
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tdeliver.frx":6DD5
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Borrar registro"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox observa4 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1800
         Width           =   6855
      End
      Begin VB.TextBox observa3 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1440
         Width           =   6855
      End
      Begin VB.TextBox observa2 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   1080
         Width           =   6855
      End
      Begin VB.TextBox observa1 
         Height          =   375
         Left            =   120
         MaxLength       =   40
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   720
         Width           =   6855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Ingreso de Lineas"
      Enabled         =   0   'False
      ForeColor       =   &H00808080&
      Height          =   3255
      Left            =   3600
      TabIndex        =   25
      Top             =   2040
      Visible         =   0   'False
      Width           =   7335
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
         Left            =   5400
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tdeliver.frx":7FE7
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Borrar registro"
         Top             =   2400
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
         Left            =   6240
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tdeliver.frx":91F9
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Grabar registro"
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox t16 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t15 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t14 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t13 
         Height          =   375
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t12 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t11 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t10 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t9 
         Height          =   375
         Left            =   4800
         MaxLength       =   2
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t8 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t7 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t6 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t5 
         Height          =   375
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox t4 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox t3 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox t2 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox t1 
         Height          =   375
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   840
         Width           =   735
      End
      Begin VB.Label linea 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   4440
         TabIndex        =   65
         Top             =   360
         Width           =   855
      End
      Begin VB.Label nt16 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   64
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt15 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   63
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt14 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   62
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt13 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5520
         TabIndex        =   61
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt12 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   60
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt11 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   59
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt10 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   58
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt9 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4080
         TabIndex        =   57
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label30 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   56
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label29 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   55
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label28 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1920
         TabIndex        =   54
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt8 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   53
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt7 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   52
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt6 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   51
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt5 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2640
         TabIndex        =   50
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nt4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   49
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label nt3 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   48
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label nt2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   47
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label nt1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1200
         TabIndex        =   46
         Top             =   840
         Width           =   735
      End
      Begin VB.Label nlinea 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00004040&
         Height          =   375
         Left            =   1200
         TabIndex        =   45
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Linea                                                     Tallas"
         ForeColor       =   &H00004040&
         Height          =   855
         Left            =   240
         TabIndex        =   44
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   360
      Top             =   11640
   End
   Begin VB.TextBox codigo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox nombre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1320
      MaxLength       =   60
      TabIndex        =   4
      Top             =   840
      Width           =   3735
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Top             =   10920
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "tdeliver.frx":A40B
      Height          =   5295
      Left            =   120
      OleObjectBlob   =   "tdeliver.frx":A41F
      TabIndex        =   0
      Top             =   1920
      Width           =   15135
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   16320
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      OutBufferSize   =   1024
      RThreshold      =   13
      RTSEnable       =   -1  'True
      SThreshold      =   2
   End
   Begin VB.Label acurabuffer 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3600
      TabIndex        =   211
      Top             =   9120
      Width           =   105
   End
   Begin VB.Label Label3 
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
      Left            =   120
      TabIndex        =   206
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Envio Correo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   840
      TabIndex        =   116
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Label19 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo               Nombre     Correo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   114
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      TabIndex        =   113
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   112
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label10 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Percepcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   12720
      TabIndex        =   111
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Label ytotal 
      BackColor       =   &H00808080&
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
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   11520
      TabIndex        =   110
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Label txpercepcion 
      BackColor       =   &H00808080&
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
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   14040
      TabIndex        =   109
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   10680
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cobro Credito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   4440
      TabIndex        =   108
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Verifica Precio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3000
      TabIndex        =   107
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Anula Venta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2280
      TabIndex        =   106
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Des Cuento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1560
      TabIndex        =   105
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Valida Delivery"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3720
      TabIndex        =   104
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F9  Tarjeta Credito"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   5160
      TabIndex        =   103
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label ppvendedor 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "vendedor"
      DataSource      =   "Data2"
      Height          =   495
      Left            =   2160
      TabIndex        =   102
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Label nrofilas 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1320
      TabIndex        =   101
      Top             =   9120
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Guia Remision"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2280
      TabIndex        =   100
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label cmdexit 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "         Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   4440
      TabIndex        =   98
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cash Boleta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   5160
      TabIndex        =   97
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label LABEL11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Grabar Ocurren."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3000
      TabIndex        =   96
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label local1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   14880
      TabIndex        =   91
      Top             =   3480
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label tpeaje 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2040
      TabIndex        =   90
      Top             =   10920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label tdetra 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2040
      TabIndex        =   89
      Top             =   11640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Si.Cobrar.Detraccion"
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
      Left            =   2040
      TabIndex        =   88
      Top             =   11280
      Width           =   2775
   End
   Begin VB.Label trdescuento 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3120
      TabIndex        =   87
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label saldo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1320
      TabIndex        =   86
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label54 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Req- Rimient"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3000
      TabIndex        =   84
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marcar Reloj  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3720
      TabIndex        =   83
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label70 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VueltoAnterior"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   81
      Top             =   8280
      Width           =   2895
   End
   Begin VB.Label uvueltos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   11760
      TabIndex        =   80
      Top             =   8280
      Width           =   3495
   End
   Begin VB.Label uvueltod 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8880
      TabIndex        =   79
      Top             =   8280
      Width           =   2895
   End
   Begin VB.Label Label62 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Abrir Gaveta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   840
      TabIndex        =   78
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label61 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Carga Orden Trabajo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3720
      TabIndex        =   77
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label60 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Copia Docum."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   76
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label59 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pre Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1560
      TabIndex        =   75
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label57 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F4  Cobro Dolar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   4440
      TabIndex        =   74
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label Label55 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cobro Normal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   5160
      TabIndex        =   73
      Top             =   8400
      Width           =   735
   End
   Begin VB.Label rtxtotald 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   6480
      TabIndex        =   24
      Top             =   7200
      Width           =   2415
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
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
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   8880
      TabIndex        =   23
      Top             =   7200
      Width           =   375
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
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
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   6000
      TabIndex        =   22
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Carga Proforma"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2280
      TabIndex        =   21
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Des- congela"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1560
      TabIndex        =   20
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "    Congela"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   840
      TabIndex        =   19
      Top             =   7200
      Width           =   735
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Limpia Pantalla"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   7200
      Width           =   735
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
      Height          =   375
      Left            =   5160
      TabIndex        =   17
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label CAMPO1 
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
      Height          =   375
      Left            =   5160
      TabIndex        =   16
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label campo3 
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
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label moneda 
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
      Height          =   375
      Left            =   9720
      TabIndex        =   14
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label diatrabajo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Otros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   8880
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label paridad 
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
      Height          =   375
      Left            =   8880
      TabIndex        =   12
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label fechasis 
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
      Height          =   375
      Left            =   8880
      TabIndex        =   11
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label horasis 
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
      Height          =   375
      Left            =   8880
      TabIndex        =   10
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label turno 
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
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label cajero 
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
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label caja 
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
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label tiposervicio1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Campos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label acu 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Height          =   195
      Left            =   3720
      TabIndex        =   3
      Top             =   9120
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label rtxtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   54
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Left            =   9240
      TabIndex        =   2
      Top             =   7200
      Width           =   6015
   End
   Begin VB.Label ntcant 
      BackColor       =   &H00808080&
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
      Left            =   120
      TabIndex        =   1
      Top             =   9120
      Width           =   1215
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
      Begin VB.Menu dlo3434 
         Caption         =   "&3.Copia Documento Ya Ingresado"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu dlko343 
         Caption         =   "&4.Anular Documento"
         Shortcut        =   {F5}
      End
      Begin VB.Menu dloco343 
         Caption         =   "&5.Congela"
         Shortcut        =   ^A
      End
      Begin VB.Menu dlo2323 
         Caption         =   "&6.Descongela"
         Shortcut        =   ^B
      End
      Begin VB.Menu dcaj8923 
         Caption         =   "&7.Apertura Cajon Monedero"
         Shortcut        =   ^C
      End
      Begin VB.Menu dhyori83 
         Caption         =   "&9.Cargar Proformas terminales"
         Shortcut        =   {F6}
      End
      Begin VB.Menu dj78232 
         Caption         =   "&A.CargaPedidos-Ordenes Trabajo"
      End
      Begin VB.Menu dk89230 
         Caption         =   "&B.Cargar Cotizaciones "
      End
      Begin VB.Menu dk8923 
         Caption         =   "&C.Carga Guia Remision"
      End
      Begin VB.Menu djk78232 
         Caption         =   "&D.Modificar Pedido Reposicion"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu d892323 
         Caption         =   "&E.Cuadre Rapido"
         Shortcut        =   ^I
      End
      Begin VB.Menu hydes8912 
         Caption         =   "&F.Control de Descuento Recargos"
         Shortcut        =   ^D
      End
      Begin VB.Menu dli992323 
         Caption         =   "&G.Limpiar Pantalla"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu fdk9235 
         Caption         =   "&H.Anulacion Otra Fecha"
      End
      Begin VB.Menu dfk992325 
         Caption         =   "&I.Copia Otra Fecha"
      End
      Begin VB.Menu dj7743400 
         Caption         =   "&J.CuentasxCobrar"
      End
   End
   Begin VB.Menu dlo2342 
      Caption         =   "&Autoservicio"
   End
   Begin VB.Menu inu781 
      Caption         =   "&Ingreso"
   End
   Begin VB.Menu djk7822 
      Caption         =   "&Egreso"
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
      Begin VB.Menu jur9012 
         Caption         =   "&3.Parcial - Unidades Vendidas Grupos"
      End
      Begin VB.Menu pado8911 
         Caption         =   "&4.Parcial - Documentos Emitidos "
      End
      Begin VB.Menu d8do82 
         Caption         =   "&5.Parcial - Productos Vs Documentos"
      End
      Begin VB.Menu forma671 
         Caption         =   "&6.Parcial - Formas de Pago"
      End
      Begin VB.Menu eju78se 
         Caption         =   "&7.Ingreso/Egreso/Seccion"
      End
      Begin VB.Menu d7822cua 
         Caption         =   "&8.Cierre de un Solo Turno"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu losao94 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tdeliver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'ojo revisar 18500
'CUANDO ES A CUENTA SE GENERA UN PEDIDO
'EL TIPO DOCUMENTO PEDIDO DEBE ESTAR EN TIPO DOCUMENTO
'Y GRABAR EN EL PEDIDO CUANDO QUEDA SALDO Y CUANTO FUE DADO A CUENTA
'C1 TOTAL DESCUENTO REFERENCIAL DON ARMANDO
'T15 DESCTO REFERENCIA T16 EL VALOR DESCTO REFERENCIAL
Dim serviciocobro    As Double

Dim xfpagox          As New ADODB.Recordset

Dim tmconsulta       As New ADODB.Recordset

Dim flag_especial    As String

Dim ndetraccion      As String

Dim flag_percepcion  As String

Dim xxacu            As String

Dim swprecio         As Integer

Dim bk2              As Variant

Dim xproducto        As String

Dim tabla_percepcion As Double

Dim exisdev          As Integer

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

Dim tivap             As Double

Dim tisc              As Double

Dim txtotald          As String

Dim txtotal           As String

Dim cprotipo          As String

Dim cproven           As String

Dim cprocod           As String

Dim InBuff            As String

Dim xptipo            As String

Dim xpserie           As String

Dim xpnumero          As String

Dim campo_precios(12) As campo_precio

Dim nrolineas         As Integer

Dim tiposervicio      As String

Dim flag_servicio     As String

Dim flag_carga        As String

Dim c1                As String

Dim c2                As String

Dim c3                As String

Dim c4                As String

Dim c5                As String

Dim c6                As String

Dim c7                As String

Dim c8                As String

Dim c9                As String

Dim gravado           As String

Dim control_flujo     As Integer

Dim protipo           As String

Dim proserie          As String

Dim pronumero         As String

Dim tximpuesto        As String

Dim xestado           As String

Dim txdescuento       As String

Dim txneto            As String

Dim txsubtotal        As String

Dim petipo            As String

Dim peserie           As String

Dim penumero          As String

Dim flage             As String

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
        acuenta = "0"

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

Private Sub clasesunat_Click()

    Dim found As Integer

    found = sumar_detalle()

    If Len(codigo) > 0 Then
        found = busca_codigo_descuento("" & codigo)

        If found = 1 Then
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1

            If DBGrid2.Enabled = True Then
                DBGrid2.SetFocus

            End If

            Exit Sub

        End If

    End If

End Sub

Private Sub cmdCancelar_Click()
    Frame9.Visible = False
    DBGrid2.SetFocus

End Sub

Function puede_congelar(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM vendedor where  clave='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        If "" & mytablex.Fields("congela") = "S" Then
            puede_congelar = 1

        End If

    End If

    mytablex.Close

End Function

Private Sub cmdGrabar_Click()

    Dim cad      As String

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim sdx      As Double

    Dim rs

    Dim I        As Integer

    Dim xcongela As String

    Dim sw       As Integer

    If Len(xcongelax) = 0 Then
        xcongelax.SetFocus
        Exit Sub

    End If

    If "" & mytable11.Fields("clavecongela") = "S" Then
        If Len(Trim(clavecongela)) = 0 Then
            clavecongela.SetFocus
            Exit Sub

        End If

        If puede_congelar(clavecongela) <> 1 Then
            MsgBox "Clave no autorizado para congelar ", 48, "Aviso"
            clavecongela.SetFocus
            Exit Sub

        End If

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
    '-----
    mytabley.Open "SELECT * FROM congelad where  numero='" & xcongela & "'", cn, adOpenDynamic, adLockOptimistic

    Do

        If Data2.Recordset.EOF Then Exit Do

        mytabley.AddNew

        For I = 0 To Data2.Recordset.Fields.count - 5
            mytabley.Fields(I) = Data2.Recordset.Fields(I)
        Next I

        mytabley.Fields("numero") = xcongela
        mytabley.Fields("acu") = acu
        mytabley.Fields("fecha") = Format(dia, "DD/MM/YYYY")
        mytabley.Fields("moneda") = mytable11.Fields("moneda")
        mytabley.Fields("usuario") = cajero
        mytabley.Fields("caja") = caja
        mytabley.Fields("turno") = turno
        mytabley.Fields("bodega") = mytable11.Fields("bodega")
        mytabley.Update
        Data2.Recordset.MoveNext
    Loop
    borra_congela
    cmdCancelar_Click

End Sub

Sub pedido_reposicion() 'sirva para que ingresen lo que necesitan que repongan

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim sdx      As Double

    Dim rs

    Dim I        As Integer

    Dim xcongela As String

    Dim sw       As Integer

    sdx = Val("" & mytable11.Fields("congela")) + 1
    xcongela = "" & sdx

denuevo13:

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM crequisa where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='01' and serie='Q' and numero='" & xcongela & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        sdx = Val(xcongela) + 1
        xcongela = "" & sdx
        GoTo denuevo13

    End If

    'mytable11.Edit
    mytable11.Fields("congela") = xcongela
    mytable11.Update
    mytablex.AddNew
    mytablex.Fields("codigo") = Trim("" & "" & mytable11.Fields("local"))
    mytablex.Fields("local") = Trim("" & "" & mytable11.Fields("local"))
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
    mytablex.Fields("nombre") = "" & busca_local_pedido(Trim("" & "" & mytable11.Fields("local")))
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
    mytablex.Open "SELECT * FROM drequisa where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='01' and serie='Q' and numero='" & xcongela & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        mytablex.Delete
        GoTo ak12

    End If

    Set rs = Data2.Recordset.Clone
    Do

        If rs.EOF Then Exit Do
        mytablex.AddNew

        For I = 0 To rs.Fields.count - 1
            mytablex.Fields(I) = rs.Fields(I)
        Next I

        mytablex.Fields("local") = Trim("" & "" & mytable11.Fields("local"))
        mytablex.Fields("serie") = "01"
        mytablex.Fields("tipo") = "Q"
        mytablex.Fields("numero") = "" & xcongela
        mytablex.Fields("vendedor") = ""
        mytablex.Fields("codigo") = Trim("" & "" & mytable11.Fields("local"))
        mytablex.Fields("local") = Trim("" & "" & mytable11.Fields("local"))
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

Private Sub codigo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        found = sumar_detalle()
        DBGrid2.Col = 0
        DBGrid2.Row = DBGrid2.VisibleRows - 1
        DBGrid2.SetFocus
        Exit Sub

    End If

    tabla_percepcion = 0

    If Len(codigo) > 0 Then
        found = busca_codigo_descuento("" & codigo)

        If found = 1 Then
            found = sumar_detalle()
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
            Exit Sub

        End If
   
    End If

    nombre.SetFocus

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        codigo = ""
        consulta_cliente1 "" & codigo
        Exit Sub

    End If

    If KeyCode = &H76 Then
        tnclie.DBPROV = "clientes"
        tnclie.fdlo893.Visible = True
        tnclie.Show 1

    End If

End Sub

Function sql_consulta(sw As Integer)

    Dim buf       As String

    Dim queprecio As String

    Dim indx      As Integer

    Dim dbf1      As String

    Dim dbf2      As String

    Dim amfecha   As String

    'Dim tmconsulta As New ADODB.Recordset
    On Error GoTo cmd8912_err

    'End If
    'MsgBox "ABC"

    'MsgBox buffer
    If opcion1 = "12" Then
        If Len(buffer) = 0 Then
            buffer.SetFocus
            Exit Function

        End If

    End If

    'MsgBox opcion1
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

    'dbf2 = "  (caja='" & "" & mytable11.Fields("caja") & "')"
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
            buf = "select Nombre,Numero,Fecha,Moneda as M,Total,Hora,Usuario,Caja,Turno from crequisa where local='" & "" & "" & mytable11.Fields("local") & "'"
        Else
            buf = "select Nombre,Numero,Fecha,Moneda as M,Total,Hora,Usuario,Caja,Turno from reponec where "
            buf = buf & " local='" & "" & "" & mytable11.Fields("local") & "'"
            buf = buf & " and " & Combo1 & " like '" & buffer & "%'"

            'indx = dbGrid1.Col
        End If

    End If

    If opcion1 = "150" Then  'descongela
        If Len(buffer) = 0 Then
            buf = "select Nombre,Numero,Fecha,Moneda as M,Total,Hora,Usuario,Caja,Turno from congelac "

            'MsgBox dbf2
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
            buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Vendedor,Hora,Caja,Turno,Local from cproform  where local='" & "" & "" & mytable11.Fields("local") & "'"

            If Len(dbf1) > 0 Then
                buf = buf & " and "

            End If

            buf = buf & dbf1
        Else
            buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Vendedor,Hora,Caja,Turno,Local from cproform where local='" & "" & "" & mytable11.Fields("local") & "' and "
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
            buf = buf & " yausado<>'1' and "
            buf = buf & "  caja='" & caja & "'"
            buf = buf & " order by fecha,HORA"
        Else
            buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Vendedor,Hora,Caja,Turno,Local,acuenta from cpedidov where local='" & "" & "" & mytable11.Fields("local") & "' and "
            'buf = buf & "  fecha=" & "DateValue('" & dia & "'" & ")"
            buf = buf & " yausado<>'1' and "
            buf = buf & "  caja='" & caja & "'"
            buf = buf & " and " & Combo1 & " like '" & buffer & "%'"
            buf = buf & "  order by fecha,HORA "
            'indx = dbGrid1.Col
   
        End If

    End If

    If opcion1 = "18500" Then  'carga ordenes de trabajo
        If Len(buffer) = 0 Then
            buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Vendedor,Hora,Caja,Turno,Local,acuenta from factura where local='" & "" & "" & mytable11.Fields("local") & "' and acu='T'  "
            'buf = buf & "  fecha=" & "DateValue('" & dia & "'" & ")"
            '-----------------------------------
            'buf = buf & " yausado<>'1'  "
            '-----------------------------------
            'buf = buf & "  caja='" & caja & "'"
            buf = buf & " order by fecha,HORA"
        Else
            buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Vendedor,Hora,Caja,Turno,Local,acuenta from factura where local='" & "" & "" & mytable11.Fields("local") & "' and acu='T'  "
            'buf = buf & "  fecha=" & "DateValue('" & dia & "'" & ")"
            'buf = buf & "  caja='" & caja & "'"
            'buf = buf & " yausado<>'1'  "
            buf = buf & " and " & Combo1 & " like '" & buffer & "%'"
            buf = buf & "  order by fecha,HORA "
            'indx = dbGrid1.Col
   
        End If

    End If

    If opcion1 = "30000" Then  'carga cotizaciones
        If Len(buffer) = 0 Then
            buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Vendedor,Hora,Caja,Turno,Local,acuenta from ccotizav where local='" & "" & "" & mytable11.Fields("local") & "'"
            'buf = buf & "  fecha=" & "DateValue('" & dia & "'" & ")"
            'buf = buf & " and (yausado='0' or yausado=null)"
            buf = buf & " order by fecha,HORA"
        Else
            buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Vendedor,Hora,Caja,Turno,Local,acuenta from ccotizav where local='" & "" & "" & mytable11.Fields("local") & "'"
            'buf = buf & " and (yausado='0' or yausado=null)"
            buf = buf & " and " & Combo1 & " like '" & buffer & "%'"
            buf = buf & "  order by fecha,HORA "

            'indx = dbGrid1.Col
        End If

    End If

    'MsgBox buf
    If opcion1 = "15A" Then
        If Len(buffer) = 0 Then
            buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Servicio as S,Placa as Ok,Vendedor,Hora,Caja,Turno,Local from " & gocabeza & " where local='" & "" & "" & mytable11.Fields("local") & "' and "
            buf = buf & "  fecha='" & amfecha & "'"
            buf = buf & " and usuario='" & cajero & "'"
            buf = buf & " and caja='" & caja & "'"
            buf = buf & " and turno='" & turno & "'"
            buf = buf & " and servicio='D'"
            buf = buf & " order by Fecha,HORA"
        Else
            buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Servicio as S,Placa as Ok,Vendedor,Hora,Caja,Turno,Local from " & gocabeza & " where local='" & "" & "" & mytable11.Fields("local") & "' and "
            buf = buf & "  fecha='" & amfecha & "'"
            buf = buf & " and usuario='" & cajero & "'"
            buf = buf & " and caja='" & caja & "'"
            buf = buf & " and turno='" & turno & "'"
            buf = buf & " and servicio='D'"
            buf = buf & " and " & Combo1 & " like '" & buffer & "%'"
            buf = buf & "  order by Fecha, HORA "

            'indx = dbGrid1.Col
        End If

    End If

    If opcion1 = "10" Or opcion1 = "13" Or opcion1 = "14" Or opcion1 = "15" Or opcion1 = "100" Or opcion1 = "1500" Then
        If Len(buffer) = 0 Then
            buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Servicio as S,Placa as Ok,Vendedor,Hora,Caja,Turno,Local from " & gocabeza & " where local='" & "" & "" & mytable11.Fields("local") & "' and "
            buf = buf & "  fecha='" & amfecha & "'"
            buf = buf & " and usuario='" & cajero & "'"
            buf = buf & " and caja='" & caja & "'"
            buf = buf & " and turno='" & turno & "'"
            buf = buf & " order by Fecha,HORA"
        Else
            buf = "select tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Servicio as S,Placa as Ok,Vendedor,Hora,Caja,Turno,Local from " & gocabeza & " where local='" & "" & "" & mytable11.Fields("local") & "' and "
            buf = buf & "  fecha='" & amfecha & "'"
            buf = buf & " and usuario='" & cajero & "'"
            buf = buf & " and caja='" & caja & "'"
            buf = buf & " and turno='" & turno & "'"
            buf = buf & " and " & Combo1 & " like '" & buffer & "%'"
            buf = buf & "  order by Fecha, HORA "

            'indx = dbGrid1.Col
        End If

    End If

    If opcion1 = "750" Then
        If Len(buffer) = 0 Then
            buf = "select FlaG_deli,tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Vendedor,Hora,Caja,Turno,Local from " & gocabeza & " where local='" & "" & "" & mytable11.Fields("local") & "' and "
            buf = buf & "  fecha='" & amfecha & "'"
            buf = buf & " and servicio='D' "
            buf = buf & " and usuario='" & cajero & "'  order by tipo,str(numero)"
        Else
            buf = "select Flaf_deli as PDeli,tipo,serie,Numero,Fecha,Nombre,Codigo,Moneda as M,Total,Estado as E,Vendedor,Hora,Caja,Turno,local from " & gocabeza & " where local='" & "" & "" & mytable11.Fields("local") & "' and "
            buf = buf & "  fecha='" & amfecha & "'"
            buf = buf & " and usuario='" & cajero & "'"
            buf = buf & " and servicio='D' "
            buf = buf & " and " & Combo1 & " like '" & buffer & "%'"
            buf = buf & "  order by fecha,HORA "

        End If

    End If

    If opcion1 = "1" Then
        If Len(buffer) = 0 Then
            buf = "select Deliveri.telefono,Clientes.Nombre,deliveri.Direccion,deliveri.referencia,Clientes.Codigo,clientes.fechanac from clientes,deliveri where deliveri.codigo=clientes.codigo "
        Else
            buf = "select Deliveri.telefono,Clientes.Nombre,clientes.Direccion,clientes.referencia,Clientes.Codigo,clientes.fechanac from  clientes,deliveri  where deliveri.codigo=clientes.codigo and " & "" & Combo1 & " like '" & buffer & "%'"

        End If

    End If

    If opcion1 = "800000" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo,Tipo,Codigo1,Direccion,Distrito,Moneda from proveedo where codigo like '%'"
        Else
            buf = "select Nombre,Codigo,Tipo,Codigo1,Direccion,Distrito,Moneda from proveedo  where " & "" & Combo1 & " like '" & buffer & "%'"

            'indx = dbGrid1.Col
        End If

    End If

    If opcion1 = "900000" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo,Tipo,Codigo1,Direccion,Distrito,Moneda from clientes where codigo like '% ORDER BY NOMBRE'"
        Else
            buf = "select Nombre,Codigo,Tipo,Codigo1,Direccion,Distrito,Moneda from clientes  where " & "" & Combo1 & " like '" & buffer & "%' ORDER BY NOMBRE"

            'indx = dbGrid1.Col
        End If

    End If

    If opcion1 = "30" Or opcion1 = "99" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo,Tipo,Codigo1,Direccion,Distrito from clientes where codigo like '%' ORDER BY NOMBRE"
        Else
            buf = "select Nombre,Codigo,Tipo,Codigo1,Direccion,Distrito from clientes  where " & "" & Combo1 & " like '" & buffer & "%' ORDER BY NOMBRE"

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
            buf = "select Nombre,Direccion,telefono,Distrito,Fechanac,Codigo,tipo,Codigo1 from clientes ORDER BY NOMBRE "
        Else
            buf = "select Nombre,Direccion,Telefono,Distrito,Fechanac,Codigo,Tipo,Codigo1 from clientes  where " & "" & Combo1 & " like '" & buffer & "%' ORDER BY NOMBRE"

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
            buf = "select Descripcio,Tipo from Tipo where (tipodoc='A' or tipodoc='B' or tipodoc='C' or tipodoc='D' or tipodoc='G') order by tipo"
        Else
            buf = "select Descripcio,Tipo from Tipo  where (tipodoc='A' or tipodoc='B' or tipodoc='C' or tipodoc='D' or tipodoc='G') and "
            buf = buf & "" & Combo1 & " like '" & buffer & "%'"
            buf = buf & "  order by tipo"

            'indx = dbGrid1.Col
        End If

    End If

    If opcion1 = "31" Or opcion1 = "3100" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from Vendedor order by nombre "
        Else
            buf = "select Nombre,Codigo from Vendedor  where "
            buf = buf & "" & Combo1 & " like '" & buffer & "%' order by nombre"

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
            buf = "select Nombre,Codigo,Direccion,Distrito,Fechanac from clientes ORDER BY NOMBRE "
        Else
            buf = "select Nombre,Codigo,Codigo,Direccion,Distrito,Fechanac from clientes  where "
            buf = buf & "" & Combo1 & " like '" & buffer & "%' ORDER BY NOMBRE"
            'indx = dbGrid1.Col
   
        End If

    End If

    If opcion1 = "12" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo,Direccion,Distrito,Fechanac,Correo from clientes ORDER BY NOMBRE "
        Else
            buf = "select Nombre,Codigo,Direccion,Distrito,Fechanac,Correo from clientes  where "
            buf = buf & "" & Combo1 & " like '" & buffer & "%' ORDER BY NOMBRE"

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

    'MsgBox mytable11.Fields("listap")
    'MsgBox "abc"
   
    'MsgBox "ABC"
    'Set tmconsulta = Nothing
    Set tmconsulta = Nothing

    If tmconsulta.State = 1 Then
        tmconsulta.Close
        Set tmconsulta = Nothing

    End If

    'If tmconsulta.State = 1 Then tmconsulta.Close
    'dbGrid1.refresh
    'dbGrid1.columns = 0
    tmconsulta.Open buf, cn, adOpenStatic, adLockOptimistic
    'MsgBox buf
    Set dbGrid1.DataSource = tmconsulta

    If tmconsulta.RecordCount = 0 Then
        buffer.SetFocus
        Exit Function

    End If

    'sw_consulta = 1
               
    If opcion1 = "8" Then
        pone_precios "" & dbGrid1.columns(1)

    End If

    dbGrid1.columns(0).Width = 7000
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

    If opcion1 = "10" Or opcion1 = "13" Or opcion1 = "14" Or opcion1 = "15" Or opcion1 = "15A" Or opcion1 = "100" Or opcion1 = "1500" Or opcion1 = "1900" Or opcion1 = "15000" Or opcion1 = "18500" Or opcion1 = "30000" Then
        dbGrid1.columns(0).Width = 700
        dbGrid1.columns(1).Width = 700
        dbGrid1.columns(2).Width = 1300
        dbGrid1.columns(3).Width = 1500
        dbGrid1.columns(4).Width = 3000
        dbGrid1.columns(5).Width = 1500
        dbGrid1.columns(6).Width = 400
        dbGrid1.columns(7).Width = 1400
        dbGrid1.columns(8).Width = 400
        dbGrid1.columns(9).Width = 800
        dbGrid1.columns(10).Width = 800
        dbGrid1.columns(11).Width = 800
        dbGrid1.columns(12).Width = 700
        dbGrid1.columns(13).Width = 700
               
    End If

    If opcion1 = "8" Then
        dbGrid1.columns(0).Width = 8000
        dbGrid1.columns(1).Width = 1300
        dbGrid1.columns(2).Width = 1000
        dbGrid1.columns(3).Width = 900
        dbGrid1.columns(4).Width = 500
        dbGrid1.columns(5).Width = 800
        dbGrid1.columns(6).Width = 500
        dbGrid1.columns(7).Width = 1000
        dbGrid1.columns(8).Width = 1700
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
    If opcion1 = "150" Or opcion1 = "10" Or opcion1 = "13" Or opcion1 = "14" Or opcion1 = "15" Or opcion1 = "15A" Or opcion1 = "100" Or opcion1 = "1500" Or opcion1 = "1900" Or opcion1 = "15000" Or opcion1 = "18500" Or opcion1 = "30000" Then
        ir_hasta_ultimo tmconsulta

    End If
               
    sql_consulta = 1
    Exit Function
cmd8912_err:
    MsgBox "Aviso en sql_consulta " & error$, 48, "Aviso"
    buffer = ""
    Exit Function

End Function

Sub ir_hasta_ultimo(raconsulta As ADODB.Recordset)

    On Error GoTo cmd789111_err

    raconsulta.MoveLast
    'dbGrid1.Col = 0
    'dbGrid1.Row = dbGrid1.VisibleRows - 1
    'dbGrid1.SetFocus
 
    Exit Sub
cmd789111_err:
    MsgBox "Aviso en ir ultimo " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Combo2_Click()

    If Len(tproducto) > 0 Then
        DBGrid4.refresh
        carga_dbgrid4 tproducto, Combo2.Text

    End If

End Sub

Private Sub Combo3_Change()

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

    Dim sw    As Integer

    Dim sdx   As Double

    If Len(pedido) = 0 Then  'si no es modificacion
        found = valida_total()

        If found = 0 Then
            MsgBox "Campos invalidos", 48, "Aviso"
            Exit Sub

        End If

    End If

    ndetraccion = ""

    If Val(tdetra) > 0 Then
        sdx = Val("" & mytable11.Fields("detraccion")) + 1
        ndetraccion = "" & sdx

    End If
   
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

    If flag_servicio = "A" Or flag_servicio = "D" Then
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

    End If

    'Frame10.Visible = True
End Sub

Private Sub Command14_Click()

    If Framefp.Visible = False Then
        Frame7.Visible = False
        DBGrid2.Enabled = True
        DBGrid2.Col = 0
        DBGrid2.Row = DBGrid2.VisibleRows - 1
        DBGrid2.SetFocus
        Exit Sub

    End If
   
    Frame7.Visible = False
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

Private Sub Command2_Click()

    Dim sdx   As Double

    Dim found As Integer

    DBGrid2.columns("t1") = Val(t1)
    DBGrid2.columns("t2") = Val(t2)
    DBGrid2.columns("t3") = Val(t3)
    DBGrid2.columns("t4") = Val(t4)
    DBGrid2.columns("t5") = Val(t5)
    DBGrid2.columns("t6") = Val(t6)
    DBGrid2.columns("t7") = Val(t7)
    DBGrid2.columns("t8") = Val(t8)
    DBGrid2.columns("t9") = Val(t9)
    DBGrid2.columns("t10") = Val(t10)
    DBGrid2.columns("t11") = Val(t11)
    DBGrid2.columns("t12") = Val(t12)
    DBGrid2.columns("t13") = Val(t13)
    DBGrid2.columns("t14") = Val(t14)
    DBGrid2.columns("t15") = Val(t15)
    DBGrid2.columns("t16") = Val(t16)
    sdx = Val(t1) + Val(t2) + Val(t3) + Val(t4) + Val(t5) + Val(t6) + Val(t7) + Val(t8) + Val(t9) + Val(t10) + Val(t11) + Val(t12) + Val(t13) + Val(t14) + Val(t15) + Val(t16)
    DBGrid2.columns("cantidad") = sdx
    sdx = Val("" & DBGrid2.columns("cantidad")) * Val("" & DBGrid2.columns("precio"))
    DBGrid2.columns("total") = sdx
    calcula_igv 0
    DBGrid2.Enabled = True
    found = sumar_detalle()
    DBGrid2.Col = 0
    DBGrid2.Row = DBGrid2.VisibleRows - 1
    DBGrid2.SetFocus
    Command3_Click

End Sub

Private Sub Command3_Click()
    DBGrid2.Enabled = True
    Frame3.Enabled = False
    Frame3.Visible = False
    DBGrid2.SetFocus

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command6_Click()

    If Frame7.Visible = True Then Exit Sub
    losao94_Click

End Sub

Private Sub Command7_Click()

    Dim sdx As Double

    DBGrid2.columns("observa1") = "" & observa1
    DBGrid2.columns("observa2") = "" & observa2
    DBGrid2.columns("observa3") = "" & observa3
    DBGrid2.columns("observa4") = "" & observa4
    calcula_igv 0
    Command9_Click

End Sub

Private Sub Command8_Click()
    'If Frame1.Visible = True Then
    '   Frame5.Visible = False
    '   dbGrid1.SetFocus
    '   Exit Sub
    'End If
    '  Frame5.Visible = False
    '  DBGrid2.Col = 0
    '  DBGrid2.Row = DBGrid2.VisibleRows - 1
    '  DBGrid2.SetFocus

    DBGrid4_KeyDown 27, 0

End Sub

Private Sub Command9_Click()
    losao94_Click

End Sub

Private Sub correo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    DBGrid2.Col = 0
    DBGrid2.Row = DBGrid2.VisibleRows - 1
    DBGrid2.SetFocus

End Sub

Private Sub correo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        nombre.SetFocus
        Exit Sub

    End If

End Sub

Private Sub d7822cua_Click()

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
    tcuadrc1.cajero = usuariopos
    tcuadrc1.caja = caja
    tcuadrc1.turno = turno
    tcuadrc1.horai = "01"
    tcuadrc1.horaf = "24"
    tcuadrc1.fechai = dia
    tcuadrc1.fechaf = dia
    tcuadrc1.Caption = "CIERRE DEL DIA"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1

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

    Dim buf     As String

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

Private Sub dbgrid1_DblClick()

    Dim found As Integer

    If opcion1 = "3100" Then
        DBGrid2.columns("vendedor") = "" & dbGrid1.columns("codigo")
        Frame1.Visible = False
        Frame1.Enabled = False
   
        If tmconsulta.State = 1 Then tmconsulta.Close
        found = sumar_detalle()
        DBGrid2.Col = 0
        DBGrid2.Row = DBGrid2.VisibleRows - 1
        DBGrid2.SetFocus

    End If

    If opcion1 = "12" Then
        codigo = "" & dbGrid1.columns("codigo")
        nombre = "" & dbGrid1.columns("nombre")
        correo = Trim("" & dbGrid1.columns("correo"))
        Frame1.Visible = False
        Frame1.Enabled = False

        If tmconsulta.State = 1 Then tmconsulta.Close
        codigo.SetFocus
        codigo_KeyPress 13

    End If

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found   As Integer

    Dim xbuf    As String

    Dim buf     As String

    Dim xtemp   As Variant

    Dim anumero As String

    Dim atipo   As String

    Dim aserie  As String

    Dim sdx     As Double

    Dim canti   As String

    DBGrid2.Enabled = True

    If KeyCode = 27 Then
        losao94_Click
        Exit Sub

    End If

    'MsgBox opcion1
    'MsgBox opcion1
    'buf = "" & Trim("" & dbGrid1.columns("numero"))
    'MsgBox opcion1
    'If KeyCode = 0 Then Exit Sub
    If KeyCode = &H71 Then  'f1  visualizar el detalle
        If opcion1 = "15A" Then  'Ok
            If Trim("" & tmconsulta.Fields("S")) = "D" Then
                If Trim("" & tmconsulta.Fields("Ok")) = "S" Then
                    tmconsulta.Fields("Ok") = ""
                    tmconsulta.Update
                    Exit Sub

                End If

                If Trim("" & tmconsulta.Fields("Ok")) = "" Then
                    tmconsulta.Fields("Ok") = "S"
                    tmconsulta.Update
                    Exit Sub

                End If

            End If

            Exit Sub

        End If

    End If

    If KeyCode = &H2E Then  'borrar linea
        If opcion1 = "150" Then
            Exit Sub

            If MsgBox("Desea Borrar Congelado " & dbGrid1.columns("numero"), 1, "Aviso") <> 1 Then
                dbGrid1.SetFocus
                Exit Sub

            End If

            cn.Execute ("delete FROM congelad where numero='" & dbGrid1.columns("numero") & "'")
            cn.Execute ("delete FROM congelac where numero='" & dbGrid1.columns("numero") & "'")
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            Exit Sub

        End If
   
        If opcion1 = "1900" Then 'borrar cproform
            If MsgBox("Desea Borrar Proforma " & dbGrid1.columns("numero"), 1, "Aviso") <> 1 Then
                dbGrid1.SetFocus
                Exit Sub

            End If

            protipo = "" & dbGrid1.columns("tipo")
            proserie = "" & dbGrid1.columns("serie")
            pronumero = "" & dbGrid1.columns("numero")
            found = borrar_proformas()
            protipo = ""
            proserie = ""
            pronumero = ""
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            DBGrid2.SetFocus
            Exit Sub

        End If

        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1  visualizar el detalle
        If opcion1 = "8" Then  'si esta en productos
            If Len("" & dbGrid1.columns("producto")) > 0 Then
      
                xproducto = "" & dbGrid1.columns("producto")
                tproducto = ""
                carga_combo2 "" & mytable11.Fields("listap")
                carga_dbgrid4 "" & dbGrid1.columns("producto"), "" & mytable11.Fields("listap")
                Exit Sub

            End If

        End If

        'MsgBox opcion1
        If opcion1 = "1500" Or opcion1 = "15" Or opcion1 = "15A" Or opcion1 = "100" Or opcion1 = "1900" Then
            If Len("" & dbGrid1.columns("tipo")) > 0 Then
                visualiza_detalle_factura "" & dbGrid1.columns("tipo"), "" & dbGrid1.columns("serie"), "" & dbGrid1.columns("numero")
                Exit Sub

            End If

        End If

        Exit Sub

    End If

    If KeyCode = 13 Then

        'MsgBox opcion1
        If opcion1 = "8" Then
            If Trim("" & dbGrid1.columns(10)) = "N" Then
                MsgBox "Producto No activo ", 48, "Aviso"
                dbGrid1.SetFocus
                Exit Sub

            End If

            'MsgBox "abc"
   
            If "" & mytable11.Fields("nosaldo") = "S" Then
                If familia_saldo("" & DBGrid2.columns("FAMILIA")) = 0 Then
                    If consulta_saldo("" & dbGrid1.columns("producto"), 0, 0) <= 0 Then
                        MsgBox "No existe saldo", 48, "Aviso"
                        dbGrid1.SetFocus
                        Exit Sub

                    End If

                End If

            End If
   
            If Val(dbGrid1.columns(5)) <= 0 Then
                If "" & mytable11.Fields("noprecio") = "S" Then
                    MsgBox "Precio<=0", 48, "Aviso"
                    dbGrid1.SetFocus
                    Exit Sub

                End If

            End If

            If Len("" & DBGrid2.columns("producto")) = 0 And Len("" & dbGrid1.columns(1)) > 0 Then
   
                If "" & mytable11.Fields("repite") = "S" Then
                    found = verifica_doble("" & dbGrid1.columns(1))

                    If found = 1 Then
                        MsgBox "Producto ya seleccionado", 48, "Aviso"
                        dbGrid1.SetFocus
                        Exit Sub

                    End If

                End If

                canti = ""

                If verifica_balanza("" & dbGrid1.columns(1)) = "S" Then
ajk922:
                    buf = puerto_balanza1()

                    If Val(buf) <= 0 Then
                        If MsgBox("Balanza No leido,Continua Leyendo? ", 1, "Aviso") = 1 Then
                            GoTo ajk922

                        End If

                        losao94_Click
                        Exit Sub

                    End If

                    'canti = Format(Val(buf), nrodecimal)
                    canti = Format(Val(buf), "0.000")

                End If

                'AQUI TIENES QUE REVISAR JOHNNY
                DBGrid2.Col = 0
                DBGrid2.Row = DBGrid2.VisibleRows - 1
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
                DBGrid2.columns("producto") = "" & dbGrid1.columns(1)
                xbuf = "" & dbGrid1.columns(1)
                'MsgBox xbuf
                found = busca_producto("" & DBGrid2.columns("producto"), 0, canti)

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
                'If verifica_balanza("" & DBGrid1.Columns(1)) = "S" Then
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
                '       dbgrid2.columns("cantidad") = Val(Mid$(Val(buf), 1, 5))
                '       sdx = Val("" & dbgrid2.columns("cantidad")) * Val("" & dbgrid2.columns("precio"))
                '       dbgrid2.columns("total") = sdx
                '       calcula_igv 0
                '    End If
                '------------------------------------------------
                Frame1.Visible = False
                Frame1.Enabled = False

                If tmconsulta.State = 1 Then tmconsulta.Close

                'found = sumar_detalle()
                'aqui ponemos si tiene mas de un precio
                'msgbox xbuf
                If ver_si_puedo_dbgrid(xbuf) = 1 Then  'existe mas de un precio
                    DBGrid2.Row = DBGrid2.VisibleRows - 2
                    DBGrid2.Col = 3
                    tproducto = ""
                    xproducto = xbuf
                    carga_combo2 "" & mytable11.Fields("listap")
                    carga_dbgrid4 xbuf, "" & mytable11.Fields("listap")
                    swprecio = 1
                    Exit Sub

                End If

                If Len(Trim("" & DBGrid2.columns("producto"))) > 0 And Len(Trim("" & DBGrid2.columns("linea"))) > 0 Then
                    DBGrid2.Col = 3
                    DBGrid2.SetFocus
                    ingreso_tallas "" & DBGrid2.columns("linea")
                    Exit Sub

                End If

                'verificar si tiene talla
                found = sumar_detalle()
                DBGrid2.Row = DBGrid2.VisibleRows - 1
                DBGrid2.SetFocus
                Exit Sub

            End If
   
        End If

        If opcion1 = "10" Then  'modifica
            xtipo = "" & dbGrid1.columns("tipo")
            xserie = "" & dbGrid1.columns("serie")
            xnumero = "" & dbGrid1.columns("numero")
            telefono = "" & dbGrid1.columns("telefono")
            codigo = "" & dbGrid1.columns("codigo")
            nombre = "" & dbGrid1.columns("nombre")
            found = busca_codigod()
            modifica_detalle

            If "" & dbGrid1.columns("servicio") = "A" Then
                tiposervicio = "Autoservicio"

            End If

            If "" & dbGrid1.columns("servicio") = "D" Then
                tiposervicio = "DELIVERY"

            End If

            xestado = "Modifica"
            Data2.refresh
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close

        End If

        If opcion1 = "0" Then
            telefono = "" & dbGrid1.columns("telefono")
            dcodigo = "" & dbGrid1.columns("codigo")
            dnombre = "" & dbGrid1.columns("nombre")
            found = busca_codigod()
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            dcodigo.SetFocus
            dcodigo_KeyPress 13

        End If

        If opcion1 = "1" Then
            If Len(Trim("" & dbGrid1.columns("codigo"))) = 0 Then
                Exit Sub

            End If

            telefono = Trim("" & dbGrid1.columns("telefono"))
            dcodigo = Trim("" & dbGrid1.columns("codigo"))
            dnombre = Trim("" & dbGrid1.columns("nombre"))
            ddireccion = Trim("" & dbGrid1.columns("direccion"))
            fechanac = Trim("" & dbGrid1.columns("fechanac"))
            referencia = Trim("" & dbGrid1.columns("referencia"))
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            dnombre.SetFocus

            'dcodigo_KeyPress 13
        End If

        If opcion1 = "1750" Then
            dcodigo = "" & dbGrid1.columns("codigo")
            dnombre = "" & dbGrid1.columns("nombre")
            ddireccion = "" & dbGrid1.columns("direccion")
            fechanac = "" & dbGrid1.columns("fechanac")
            telefono = "" & dbGrid1.columns("telefono")
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            dcodigo.SetFocus
            dcodigo_KeyPress 13

        End If
   
        If opcion1 = "23" Then
            tcampo1 = "" & dbGrid1.columns("codigo")
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            tcampo1.SetFocus
            tcampo1_KeyPress 13

        End If

        If opcion1 = "200" Then
            tcampo4 = "" & dbGrid1.columns("banco")
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            tcampo4.SetFocus

        End If

        If opcion1 = "12" Then
            codigo = "" & dbGrid1.columns("codigo")
            nombre = "" & dbGrid1.columns("nombre")
            correo = "" & dbGrid1.columns("correo")
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            codigo.SetFocus
            codigo_KeyPress 13

        End If

        If opcion1 = "300" Then 'bodega de traslado
            xruc = Trim("" & dbGrid1.columns("codigo"))
            xnombre = "" & dbGrid1.columns("nombre")
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            xruc.SetFocus
            xruc_KeyPress 13

        End If
   
        If opcion1 = "31" Then
            xvendedor = "" & dbGrid1.columns("codigo")
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            xvendedor.SetFocus
            xvendedor_KeyPress 13

        End If

        If opcion1 = "3100" Then
            DBGrid2.columns("vendedor") = "" & dbGrid1.columns("codigo")
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            found = sumar_detalle()
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus

        End If
   
        If opcion1 = "30" Then
            If xtipo = "2" Or xtipo = "4" Then
                'If Len("" & dbGrid1.Columns("ruc")) <> 11 Then
                '   MsgBox "Ruc Invalido ", 48, "Aviso"
                '   Exit Sub
                'End If
                xruc = Trim("" & dbGrid1.columns("codigo"))
            Else
                xruc = Trim("" & dbGrid1.columns("codigo"))

            End If

            codigo = Trim("" & dbGrid1.columns("codigo"))
            xnombre = Trim("" & dbGrid1.columns("nombre"))
            nombre = Trim("" & dbGrid1.columns("nombre"))
            xdireccion = Trim("" & dbGrid1.columns("direccion"))
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            xdireccion_KeyPress 13

        End If

        If opcion1 = "99" Then
            tcampo1 = "" & dbGrid1.columns("codigo")
            tcampo2 = "" & dbGrid1.columns("nombre")
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            tcampo1.SetFocus
            tcampo1_KeyPress 13

        End If

        If opcion1 = "2800" Then
            If Val("" & dbGrid1.columns("saldo")) < Val(stxtotals) Then
                MsgBox "Debe ingresar valor exacto", 48, "Aviso"
                dbGrid1.SetFocus
                Exit Sub

            End If

            tcampo1 = "" & dbGrid1.columns("codigo")
            tcampo2 = "" & dbGrid1.columns("nombre")
            tcampo3 = "" & dbGrid1.columns("numero")
            tcampo4 = "" & dbGrid1.columns("tipo")
            tcampo5 = "" & dbGrid1.columns("serie")
            tcampo6 = "" & dbGrid1.columns("local")
            saldoabo = "" & dbGrid1.columns("saldo")
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            tcampo3.SetFocus

            'tcampo3_KeyPress 13
        End If

        If opcion1 = "29" Then
            xtipo = "" & dbGrid1.columns(1)
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            xtipo.SetFocus
            xtipo_keyPress 13

        End If

        If opcion1 = "13" Then  'copia documento
            If MsgBox("Desea Sacar Copia del Documento", 1, "Aviso") <> 1 Then Exit Sub
            proceso_impresioncopia
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            DBGrid2.SetFocus

        End If

        If opcion1 = "1500" Then  'carga documento anterior
            If MsgBox("Desea Cargar Detalle Documento ", 1, "Aviso") <> 1 Then Exit Sub
            found = proceso_carga_doc_ant("" & dbGrid1.columns("local"), "" & dbGrid1.columns("tipo"), "" & dbGrid1.columns("serie"), "" & dbGrid1.columns("numero"))

            If found = 0 Then
                MsgBox "Error de carga", 48, "Aviso"
                Exit Sub

            End If

            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            DBGrid2.SetFocus

        End If
   
        If opcion1 = "15000" Then  'carga pedidos de venta anteriores para cancelar
            If MsgBox("Desea Cargar Detalle Documento ", 1, "Aviso") <> 1 Then Exit Sub
            petipo = "" & dbGrid1.columns("tipo")
            peserie = "" & dbGrid1.columns("serie")
            penumero = "" & dbGrid1.columns("numero")
            acuenta = "" & dbGrid1.columns("acuenta")
            codigo = "" & dbGrid1.columns("codigo")
            nombre = "" & dbGrid1.columns("nombre")
            cproven = "" & dbGrid1.columns("vendedor")
            found = proceso_carga_Pedido("" & dbGrid1.columns("local"), "" & dbGrid1.columns("tipo"), "" & dbGrid1.columns("serie"), "" & dbGrid1.columns("numero"))

            If found = 0 Then
                MsgBox "Error de carga", 48, "Aviso"
                Exit Sub

            End If

            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            DBGrid2.SetFocus

        End If
   
        If opcion1 = "18500" Then  'carga pedidos de venta anteriores para cancelar
            If MsgBox("Desea Cargar Detalle Documento ", 1, "Aviso") <> 1 Then Exit Sub
            'petipo = "" & dbGrid1.columns("tipo")
            'peserie = "" & dbGrid1.columns("serie")
            'penumero = "" & dbGrid1.columns("numero")
            'acuenta = "" & dbGrid1.columns("acuenta")
            codigo = "" & dbGrid1.columns("codigo")
            nombre = "" & dbGrid1.columns("nombre")
            cproven = "" & dbGrid1.columns("vendedor")
            found = proceso_carga_guia("" & dbGrid1.columns("local"), "" & dbGrid1.columns("tipo"), "" & dbGrid1.columns("serie"), "" & dbGrid1.columns("numero"))

            If found = 0 Then
                MsgBox "Error de carga Guia", 48, "Aviso"
                Exit Sub

            End If

            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            DBGrid2.SetFocus

        End If
   
        If opcion1 = "30000" Then  'carga cotizaciones
            If MsgBox("Desea Cargar Detalle Documento ", 1, "Aviso") <> 1 Then Exit Sub
            petipo = "" & dbGrid1.columns("tipo")
            peserie = "" & dbGrid1.columns("serie")
            penumero = "" & dbGrid1.columns("numero")
            codigo = "" & dbGrid1.columns("codigo")
            nombre = "" & dbGrid1.columns("nombre")
            cproven = "" & dbGrid1.columns("vendedor")
            found = proceso_carga_cotizacion("" & dbGrid1.columns("local"), "" & dbGrid1.columns("tipo"), "" & dbGrid1.columns("serie"), "" & dbGrid1.columns("numero"))

            If found = 0 Then
                MsgBox "Error de carga", 48, "Aviso"
                Exit Sub

            End If

            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            DBGrid2.SetFocus

        End If
   
        If opcion1 = "1900" Then  'cargar proformas
            If MsgBox("Desea Cargar Proforma ", 1, "Aviso") <> 1 Then Exit Sub
            found = proceso_proforma("" & dbGrid1.columns("local"), "" & dbGrid1.columns("tipo"), "" & dbGrid1.columns("serie"), "" & dbGrid1.columns("numero"))

            If found = 0 Then
                MsgBox "Error de carga", 48, "Aviso"
                dbGrid1.SetFocus
                Exit Sub

            End If

            'sql_detalle
            'borrar_data1
            cproven = "" & dbGrid1.columns("vendedor")
            codigo = "" & dbGrid1.columns("codigo")
            nombre = "" & dbGrid1.columns("nombre")
            protipo = "" & dbGrid1.columns("tipo")
            proserie = "" & dbGrid1.columns("serie")
            pronumero = "" & dbGrid1.columns("numero")
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            DBGrid2.SetFocus

        End If

        If opcion1 = "14" Then  'BORRAR
            If MsgBox("Desea Borrar del Documento  ", 1, "Aviso") <> 1 Then Exit Sub
            PROCESO_BORRAR_DOCUMENTO "" & dbGrid1.columns("local"), "" & dbGrid1.columns("tipo"), "" & dbGrid1.columns("serie"), "" & dbGrid1.columns("numero")
            Frame1.Visible = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            DBGrid2.SetFocus

        End If

        If opcion1 = "150" Then 'descongelar
            found = menu_descongela("" & dbGrid1.columns(1))
            MsgBox "Presione enter para continuar...", 48, "Aviso"

            If found = 1 Then
                borrar_descongela1 "" & dbGrid1.columns(1)
                borrar_descongela "" & dbGrid1.columns(1)
                '''sql_detalle
                found = sumar_detalle()
                losao94_Click

            End If

        End If

        If opcion1 = "370" Then 'cargar reposicion para modificar
            found = menu_repone("" & dbGrid1.columns("numero"))
            MsgBox "Presione enter para continuar...", 48, "Aviso"

            If found = 1 Then
                borrar_repone "" & dbGrid1.columns("numero")
                borrar_reponexx
                '''sql_detalle
                found = sumar_detalle()
                losao94_Click

            End If

        End If

        If opcion1 = "750" Then  'deliveri no xxx
            If "" & dbGrid1.columns("flag_deli") = "S" Then
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

            If "" & dbGrid1.columns("flag_deli") = "" Then
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

        If opcion1 = "15A" Then
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            DBGrid2.SetFocus
            Exit Sub

        End If

        If opcion1 = "15" Then  'copia documento
            If MsgBox("Desea Sacar Copia del Documento", 1, "Aviso") <> 1 Then
                dbGrid1.SetFocus
                Exit Sub

            End If

            atipo = "" & dbGrid1.columns("tipo")
            aserie = "" & dbGrid1.columns("serie")
            anumero = "" & dbGrid1.columns("numero")
            'impresion_sin_formato atipo, aserie, anumero
            proceso_impresion11 atipo, aserie, anumero, 1, "1"
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            DBGrid2.SetFocus
            Exit Sub

        End If

        If opcion1 = "100" Then  'anula documento
            If "" & dbGrid1.columns("e") = "1" Then
                MsgBox "Documento Anulado ", 48, "Aviso"
                dbGrid1.SetFocus
                Exit Sub

            End If

            If MsgBox("Desea Anular Documento " + "" & dbGrid1.columns("numero"), 1, "Aviso") <> 1 Then
                dbGrid1.SetFocus
                Exit Sub

            End If

            atipo = "" & dbGrid1.columns("tipo")
            aserie = "" & dbGrid1.columns("serie")
            anumero = "" & dbGrid1.columns("numero")
            found = proceso_anular(atipo, aserie, anumero)

            If found = 1 Then
                If atipo = "3" Or atipo = "4" Then
                    If MsgBox("Desea Imprimir ", 1, "Aviso") = 1 Then
                        proceso_impresion11 atipo, aserie, anumero, 0, ""

                    End If

                Else
                    proceso_impresion11 atipo, aserie, anumero, 0, ""

                End If
         
                Frame1.Visible = False
                Frame1.Enabled = False

                If tmconsulta.State = 1 Then tmconsulta.Close
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

    'On Error GoTo cmd672222_err
    'Data1.Recordset.Delete
    'Exit Sub
    'cmd672222_err:
    'Exit Sub
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

    Dim buf   As String

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
    Dim buf   As String

    Dim buf2  As String

    Dim sw    As Integer

    Dim found As Integer

    On Error GoTo cmd918_err

    'MsgBox ""
    If opcion1 = "8" Then
        'MsgBox "" & dbGrid1.Columns(1)
        pone_precios "" & dbGrid1.columns(1)

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

    Dim ind   As Integer

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

        xruc = Trim(codigo)
        xnombre = nombre

        If Len(Trim(cproven)) > 0 Then
            xvendedor = cproven

        End If

        Frame7.Visible = True
        Framefp.Enabled = False
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
        dbgrid9.Row = dbgrid9.VisibleRows - 1
        dbgrid9.Col = 2
        dbgrid9.SetFocus
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

    If "" & dbgrid10.columns(3) = "H" Then     'bancos
        macro_credito 10
        tcampo1.SetFocus

    End If

    If "" & dbgrid10.columns(3) = "K" Then     'bancos
        macro_credito 2
        tcampo1.SetFocus

    End If

    If "" & dbgrid10.columns(3) = "V" Then   'vales
        macro_credito 6
        tcampo1.SetFocus

    End If

    If "" & dbgrid10.columns(3) = "O" Then     'Otros
        macro_credito 1
        tcampo1.SetFocus

    End If

    If "" & dbgrid10.columns(3) = "I" Or "" & dbgrid10.columns(3) = "K" Then   'CRUCE CON ABONO EFECTIVO
        macro_credito 1
        tcampo1.Enabled = True
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
 
Private Sub dbgrid2_AfterColUpdate(ByVal ColIndex As Integer)

    Dim found     As Integer

    Dim sdx       As Double

    Dim fueldonde As String

    Select Case ColIndex

        Case 0
            'found = busca_producto("" & dbgrid2.columns("producto"), 0)
            'If found = 0 Then
            '   MsgBox "No existe producto", 48, "Aviso"
            '   Exit Sub
            'End If
            'If control_flujo = 1 Then
            '   MsgBox "Hola"
            'End If
            'MsgBox "Hola"
            'MsgBox "" & dbgrid2.columns("producto")
            found = busca_remate("" & DBGrid2.columns("producto"))

            If found = 1 Then
                DBGrid2.Col = 5
                ingreso_tallas "" & DBGrid2.columns("linea")
                Exit Sub

            End If

            If ver_si_puedo_dbgrid("" & DBGrid2.columns("producto")) = 1 Then  'existe mas de un precio
                'MsgBox "AQUIP PASA ALGO"
                xproducto = "" & DBGrid2.columns("producto")
                tproducto = ""
                carga_combo2 "" & mytable11.Fields("listap")
                carga_dbgrid4 "" & DBGrid2.columns("producto"), "" & mytable11.Fields("listap")
                swprecio = 1
                Exit Sub

            End If

            If Len(DBGrid2.columns("producto")) > 0 And Len(DBGrid2.columns("linea")) > 0 Then
                DBGrid2.Col = 3
                ingreso_tallas "" & DBGrid2.columns("linea")
                Exit Sub

            End If

            'MsgBox "abc"
            'dbgrid2.Col = 8
            'dbgrid2.SetFocus
            'Exit Sub
            'If "" & mytable11.Fields("vdetalle") = "S" Then
            '   If Len(Trim("" & Data2.Recordset.Fields("vendedor"))) = 0 Then
            '      dbgrid2.Col = 8
            '      Exit Sub
            '   End If
            'End If
            fueldonde = existe_fuel("" & DBGrid2.columns("producto"))

            'MsgBox fueldonde
            If Val("" & DBGrid2.columns("cantidad")) = 1 Then

                Select Case fueldonde

                    Case "C"
                        DBGrid2.Col = 3
                        DBGrid2.SetFocus
                        Exit Sub

                    Case "P"
                        DBGrid2.Col = 5
                        DBGrid2.SetFocus
                        'MsgBox "abc"
                        Exit Sub

                    Case "T"
                        'MsgBox "abc"
                        DBGrid2.Col = 7
                        DBGrid2.SetFocus
                        'MsgBox "abc"
                        Exit Sub

                    Case "V"
                        'MsgBox "abc"
                        DBGrid2.Col = 8
                        DBGrid2.SetFocus
                        'MsgBox "abc"
                        Exit Sub

                End Select

            End If
            
            'MsgBox "abc"
            'dbgrid2.Col = 8
            'dbgrid2.SetFocus
            'MsgBox "abc"
            'Exit Sub
            'valida el vendedor
            'MsgBox ""
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
            'sdx = Val("" & dbgrid2.columns("cantidad")) * Val("" & dbgrid2.columns("precio"))
            'dbgrid2.columns("tax") = Val(Format(sdx, nrodecimal))
            'dbgrid2.columns("total") = Val(Format(sdx, nrodecimal))
            'calcula_igv
            'ir_ultimo
            'MsgBox "Hola"
            found = sumar_detalle()
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus

        Case 5
            'sdx = Val("" & dbgrid2.columns("cantidad")) * Val("" & dbgrid2.columns("precio"))
            'dbgrid2.columns("tax") = Val(Format(sdx, nrodecimal))
            'dbgrid2.columns("total") = Val(Format(sdx, nrodecimal))
            'calcula_igv
            found = sumar_detalle()
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus

        Case 6
            'sdx = Val("" & dbgrid2.columns("cantidad")) * Val("" & dbgrid2.columns("precio"))
            'dbgrid2.columns("tax") = Val(Format(sdx, nrodecimal))
            'dbgrid2.columns("total") = Val(Format(sdx, nrodecimal))
            'calcula_igv
            found = sumar_detalle()
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus

        Case 7
            'If Val("" & dbgrid2.columns("cantidad")) > 0 Then
            '   sdx = Val("" & dbgrid2.columns("total")) / Val("" & dbgrid2.columns("cantidad"))
            '   dbgrid2.columns("precio") = Val(Format(sdx, nrodecimal))
            '   dbgrid2.columns("tax") = Val("" & dbgrid2.columns("total"))
            '   calcula_igv
            'calcula_igv
            'MsgBox ""
            found = sumar_detalle()
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus

            'End If
    End Select

End Sub

Private Sub dbgrid2_BeforeColEdit(ByVal ColIndex As Integer, _
                                  ByVal KeyAscii As Integer, _
                                  Cancel As Integer)

    Dim found As Integer

    Select Case ColIndex

        Case 57

            If Len(DBGrid2.columns("producto")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            Exit Sub

    End Select

    If ColIndex >= 14 Then
        Cancel = True
        Exit Sub

    End If

    Select Case ColIndex

        Case 2, 4, 8, 9, 10, 12, 11
            Cancel = True
            Exit Sub

        Case 1

            If Len("" & DBGrid2.columns("producto")) = 0 Then  'si ya existe no cambiar
                Cancel = True
                Exit Sub

            End If

        Case 0

            If Len("" & DBGrid2.columns("producto")) > 0 Then  'si ya existe no cambiar
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

            If Len("" & DBGrid2.columns("producto")) = 0 Then  '
                Cancel = True
                Exit Sub

            End If

        Case 3

            If Len("" & DBGrid2.columns("producto")) = 0 Then  '
                Cancel = True
                Exit Sub

            End If

            'If Len("" & dbgrid2.columns("linea")) > 0 Then  'ojo no se puede poner si es talla
            '   Cancel = True
            '   Exit Sub
            'End If
        Case 5, 7, 13, 6

            If Len("" & DBGrid2.columns("producto")) = 0 Then  '
                Cancel = True
                Exit Sub

            End If

            'MsgBox ""
     
    End Select

End Sub

Private Sub dbgrid2_BeforeColUpdate(ByVal ColIndex As Integer, _
                                    OldValue As Variant, _
                                    Cancel As Integer)

    Dim sw        As Integer

    Dim found     As Integer

    Dim sdx       As Double

    Dim xcampo    As String

    Dim canti     As String

    Dim buf1      As String

    Dim buf       As String

    Dim fueldonde As String

    Dim bufy      As String

    Dim amount    As String

    Dim xfound    As String

    Dim xnbufx    As Double

    Select Case ColIndex

        Case 57

            If Len(DBGrid2.columns("producto")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            calcula_igv 0

        Case 1

            If Len(DBGrid2.columns("producto")) = 0 Then
                Cancel = True
                Exit Sub

            End If

        Case 0

            If Len(DBGrid2.columns("producto")) = 0 Then
                'aqui vamos a valida si es el fin del pedido
                Cancel = True
                Exit Sub

            End If

            If Len(DBGrid2.columns("producto")) > 18 Then
                'aqui vamos a valida si es el fin del pedido
                Cancel = True
                Exit Sub

            End If

            'MsgBox "abc"
     
            If "" & mytable11.Fields("nosaldo") = "S" Then

                ' MsgBox "" & dbgrid2.columns("total")
                If familia_saldo("" & DBGrid2.columns("familia")) = 0 Then
                    If verifica_producto("" & DBGrid2.columns("producto")) = 1 Then
                        If consulta_saldo("" & DBGrid2.columns("producto"), 0, 0) <= 0 Then
                            Cancel = True
                            DBGrid2.SetFocus
                            MsgBox "x.No existe saldo", 48, "Aviso"
                            '----------------
                            'found = sumar_detalle()
                            'dbgrid2.Col = 0
                            'dbgrid2.Row = dbgrid2.VisibleRows - 1
                            'dbgrid2.SetFocus
                            '----------------
                  
                            Exit Sub

                        End If

                    End If

                End If

            End If

            canti = ""
            buf = UCase(DBGrid2.columns("producto"))  'se modifico en U. Union
            bufy = buf
            found = 0
            sw = 0
            'MsgBox "" & mytable11.Fields("flag")
            found = InStr(buf, "*")

            If found > 1 Then  ' si es cantidad
                xcampo = Mid$(buf, found + 1, Len(buf) - found)
                canti = Mid$(buf, 1, found - 1)
                buf1 = Val(canti)
                buf = xcampo

                If Len(buf) = 0 Then
                    Cancel = True
                    Exit Sub

                End If

                DBGrid2.columns("producto") = buf

            End If

            found = InStr(buf, "+")

            If found > 1 Then  ' si es cantidad
                xcampo = Mid$(buf, found + 1, Len(buf) - found)
                canti = Mid$(buf, 1, found - 1)
                buf1 = Val(canti)
                buf = xcampo

                If Len(buf) = 0 Then
                    Cancel = True
                    Exit Sub

                End If

                DBGrid2.columns("producto") = buf
                sw = 1

            End If

            'MsgBox buf
            If "" & mytable11.Fields("repite") = "S" Then
                found = verifica_doble("" & DBGrid2.columns("producto"))

                If found = 1 Then
                    Cancel = True
                    MsgBox "Producto ya Seleccionado", 48, "Aviso"
                    Exit Sub

                End If

            End If

            '----validamos el saldo
            control_flujo = 0
            found = busca_producto(UCase("" & DBGrid2.columns("producto")), sw, canti)

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
                'consulta_producto "" & dbgrid2.columns("producto")
                opcion5 = 1
                found = consulta_producto(bufy)

                If found = 1 Then
                    Cancel = True
                    opcion5 = 20
                    MsgBox "No existe producto", 48, "Aviso"
                    DBGrid2.SetFocus
                    'opcion5 = 20
                    'DBGrid2.Col = 0
                    'DBGrid2.Row = DBGrid2.VisibleRows - 1
                    'DBGrid2.SetFocus
                    Exit Sub

                End If

                opcion5 = 0
                Exit Sub

            End If

            buf = ""
            'If "" & mytable11.Fields("actbala") = "S" Then
            'If verifica_balanza("" & dbgrid2.columns("producto")) = "S" Then
        
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
            'dbgrid2.columns("cantidad") = Val(Mid$(Val(buf), 1, 5))
            'sdx = Val("" & dbgrid2.columns("cantidad")) * Val("" & dbgrid2.columns("precio"))
            'dbgrid2.columns("total") = sdx
            'calcula_igv 0
            '-------------------
            'End If
            swprecio = 0
            Exit Sub

        Case 2

            If Len(DBGrid2.columns("producto")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            If Mid$("" & DBGrid2.columns("unidad"), 1, 1) = "-" And Len("" & DBGrid2.columns("unidad")) > 1 Then
                'grabar_foto "" & Value
                Exit Sub

            End If

            found = valida_placa("" & DBGrid2.columns("linea"), Mid$("" & DBGrid2.columns("unidad"), 1, 1))

            If found = 0 Then
                MsgBox "Placa invalida ", 48, "Aviso"
                Cancel = True
                Exit Sub

            End If

        Case 3

            If Len(DBGrid2.columns("producto")) = 0 Then
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

            'MsgBox Val("" & dbgrid2.columns("cantidad"))
            If "" & mytable11.Fields("nosaldo") = "S" Then
                If familia_saldo("" & DBGrid2.columns("familia")) = 0 Then
                    If consulta_saldo("" & DBGrid2.columns("producto"), Val("" & DBGrid2.columns("cantidad")) * Val("" & DBGrid2.columns("factor")), 1) <= 0 Then
                        Cancel = True
                        DBGrid2.SetFocus
                        MsgBox "No existe saldo Suficiente", 48, "Aviso"
                        '----------------------
                        Exit Sub

                    End If

                End If

            End If

            found = busca_unidad("" & DBGrid2.columns("producto"))

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
                    found = verifica_ofertax("" & DBGrid2.columns("producto"), Val("" & DBGrid2.columns("cantidad")), xnbufx)

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

            If Len(DBGrid2.columns("producto")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            If Not IsNumeric(DBGrid2.columns("precio")) Then
                Cancel = True
                Exit Sub

            End If

            'MsgBox "hola"
            xfound = verifica_oferta("" & DBGrid2.columns("producto"))

            If xfound <> "S" Then   '
                If Val(DBGrid2.columns("precio")) <= 0 Then
                    If "" & mytable11.Fields("noprecio") = "S" Then
                        MsgBox "Precio <=0", 48, "Aviso"
                        Cancel = True
                        Exit Sub

                    End If

                End If
        
                If "" & mytable11.Fields("obligaprecio") = "S" Then
                    flag_clave1 = 0
                    tconcla.X = "S"
                    tconcla.Show 1

                    If flag_clave1 = 0 Then  'si es descongela
                        Cancel = True
                        Exit Sub

                    End If

                End If

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

            If Len(DBGrid2.columns("producto")) = 0 Then
                Cancel = True
                Exit Sub

            End If

            If "" & mytable11.Fields("obligaprecio") = "S" Then
                flag_clave1 = 0
                tconcla.X = "DESCUENTO"
                tconcla.Show 1

                If flag_clave1 = 0 Then  'si es descongela
                    Cancel = True
                    Exit Sub

                End If

            End If
     
            If Not IsNumeric(DBGrid2.columns("deslipo")) Then
                Cancel = True
                Exit Sub

            End If

            sdx = Val("" & DBGrid2.columns("cantidad")) * Val("" & DBGrid2.columns("precio"))
            DBGrid2.columns("total") = sdx
            calcula_igv 0

        Case 7

            'MsgBox ""
            If Len(DBGrid2.columns("producto")) = 0 Then
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

            'xfound = verifica_oferta("" & dbgrid2.columns("producto"))
            'if xfound<>"S" then
            If Val(DBGrid2.columns("total")) <= 0 Then
                If "" & mytable11.Fields("noprecio") = "S" Then
                    MsgBox "Precio <=0", 48, "Aviso"
                    Cancel = True
                    Exit Sub

                End If

            End If

            '-----------------xxxx--------------------------------
            fueldonde = existe_fuel("" & DBGrid2.columns("producto"))

            If fueldonde <> "C" And fueldonde <> "P" And fueldonde <> "T" And fueldonde <> "V" Then
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

            If fueldonde = "C" And fueldonde = "P" And fueldonde = "T" And fueldonde = "V" Then
                If Val("" & DBGrid2.columns("precio")) = 0 Then
                    Cancel = True
                    Exit Sub

                End If

                Select Case fueldonde

                    Case "C"
                        Exit Sub

                    Case "P"
                        Exit Sub

                    Case "V"
                        Exit Sub

                    Case "T"
                        sdx = Val("" & DBGrid2.columns("total")) / Val("" & DBGrid2.columns("precio"))
                        DBGrid2.columns("cantidad") = sdx
                        calcula_igv 0
                        Exit Sub

                End Select
           
            End If

            '-------------------------------------------------
            'flag_clave1 = 0
            'tconcla.X = "S"
            'tconcla.Show 1
            'If flag_clave1 = 0 Then  'si es descongela
            '   Cancel = True
            '   Exit Sub
            'End If
     
            'sdx = Val("" & dbgrid2.columns("total")) / Val("" & dbgrid2.columns("cantidad"))
            'dbgrid2.columns("precio") = sdx
            'calcula_igv 0
        Case 13

            If Len(DBGrid2.columns("producto")) = 0 Then
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

    Select Case DBGrid2.Col

        Case 8

            If Len("" & DBGrid2.columns("producto")) > 0 Then
                consulta_xvendedor1

            End If

    End Select

End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found     As Integer

    Dim fueldonde As String

    On Error GoTo cmd341231_err

    'Exit Sub
    If KeyCode = &H71 Then  'f2 totales
        proceso_cierre_automatico "EFECTIVO"
        'Label55_Click
        Exit Sub

    End If

    If KeyCode = &H78 Then  'exonerado
        'MsgBox ""
        proceso_cierre_automatico "TARJETACREDITO"
        Exit Sub

    End If

    If KeyCode = &H73 Then  'F4
        'MsgBox ""
        proceso_cierre_automatico "DOLAR"
        Exit Sub

    End If

    If KeyCode = 13 Then
        If DBGrid2.Col = 0 Then
            fueldonde = existe_fuel("" & DBGrid2.columns("producto"))

            If Val("" & DBGrid2.columns("cantidad")) = 1 Then

                Select Case fueldonde

                    Case "C"
                        DBGrid2.Col = 3
                        DBGrid2.SetFocus
                        KeyCode = 0
                        Exit Sub

                    Case "P"
                        'MsgBox "abc"
                        DBGrid2.Col = 4
                        DBGrid2.SetFocus
                        KeyCode = 0
                        Exit Sub

                    Case "V"
                        'MsgBox "abc"
                        DBGrid2.Col = 8
                        DBGrid2.SetFocus
                        KeyCode = 0
                        Exit Sub

                    Case "T"
                        DBGrid2.Col = 7
                        DBGrid2.SetFocus
                        KeyCode = 0
                        Exit Sub

                End Select

            End If

            'MsgBox "abc"
            If Len("" & DBGrid2.columns("producto")) > 0 Then

                Select Case fueldonde

                    Case "C"
                        DBGrid2.Col = 3
                        DBGrid2.SetFocus
                        Exit Sub

                    Case "P"
                        DBGrid2.Col = 4
                        DBGrid2.SetFocus
                        Exit Sub

                    Case "T"
                        DBGrid2.Col = 7
                        DBGrid2.SetFocus
                        KeyCode = 0
                        Exit Sub

                    Case "V"
                        DBGrid2.Col = 8
                        DBGrid2.SetFocus
                        KeyCode = 0
                        Exit Sub

                End Select
      
                DBGrid2.Col = 3
                DBGrid2.SetFocus
                KeyCode = 0
                Exit Sub

            End If
   
        End If

        If DBGrid2.Col = 57 Then
            KeyCode = 0
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
            KeyCode = 0
            Exit Sub

        End If

        If DBGrid2.Col = 3 Then
            If Val(DBGrid2.columns("cantidad")) > 0 Then
                found = sumar_detalle()
                KeyCode = 0
                DBGrid2.Col = 0
                DBGrid2.Row = DBGrid2.VisibleRows - 1
                DBGrid2.SetFocus
                KeyCode = 0
                Exit Sub

            End If

        End If

        Select Case DBGrid2.Col

            Case 0

                If Len("" & DBGrid2.columns("producto")) = 0 Then
                    Label55_Click
                    Exit Sub

                End If

                'If Len("" & dbgrid2.columns("producto")) > 0 Then
                '   DBGrid2.Col = 2
                'End If
            Case 2
                'MsgBox ""
            
            Case 3
                'MsgBox ""
                'If Val("" & dbgrid2.columns("precio")) = 0 Then
                '   dbgrid2.Col = 5
                '   Exit Sub
                'End If
                found = sumar_detalle()
                KeyCode = 0
                DBGrid2.Col = 0
                DBGrid2.Row = DBGrid2.VisibleRows - 1
                DBGrid2.SetFocus

            Case 4:
                'found = sumar_detalle()
                'KeyCode = 0
                'dbgrid2.Col = 0
                'dbgrid2.Row = dbgrid2.VisibleRows - 1
                'dbgrid2.SetFocus
       
            Case 5
                'If Len("" & dbgrid2.columns("factor")) > 0 Then
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
    '   MsgBox "hOLA" '
    '
    'End If
End Sub

Private Sub DBGrid2_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    On Error GoTo cmd34_err

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
        If Len(DBGrid2.columns("producto")) = 0 Then
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
            Exit Sub

        End If

        'End If
        Select Case DBGrid2.Col

            Case 0

                'MsgBox "xxxx"
            Case 3

                If Val(DBGrid2.columns("cantidad")) > 0 Then

                    'dbgrid2.Col = 0
                    'dbgrid2.Row = dbgrid2.VisibleRows - 1
                    'dbgrid2.SetFocus
                    'Exit Sub
                End If

        End Select

    End If

    'MsgBox "AAA"
 
    If KeyCode = &H70 Then  'f1  carga los demas precios
        If Len(DBGrid2.columns("producto")) > 0 And DBGrid2.Col = 2 Then
            xproducto = "" & DBGrid2.columns("producto")
            tproducto = ""
            carga_combo2 "" & mytable11.Fields("listap")
            carga_dbgrid4 "" & DBGrid2.columns("producto"), "" & mytable11.Fields("listap")
            Exit Sub

        End If

        If Len(DBGrid2.columns("producto")) > 0 And DBGrid2.Col = 8 Then
            consulta_xvendedor1
            Exit Sub

        End If
   
    End If

    If KeyCode = &H72 Then  'f3
        codigo.SetFocus
        Exit Sub

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

        'KeyCode = 0
    End If

    If KeyCode = &H2E Then  'borrar linea
        If DBGrid2.Row = -1 Then
            MsgBox "No hay ningún registro para eliminar", vbInformation
            Exit Sub

        End If

        If "" & mytable11.Fields("limpiapantalla") = "S" Then
            flag_clave1 = 0
            tconcla.X = "CLEAR"
            tconcla.Show 1

            If flag_clave1 = 0 Then  'si es descongela
                Exit Sub

            End If

        End If

        If MsgBox("Se va a eliminar el registro : está seguro ", vbExclamation + vbYesNo, "Eliminar") = vbYes Then
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
        If Len(DBGrid2.columns("producto")) = 0 Then
            found = consulta_producto("")

        End If

    End If

    If KeyCode = &H72 Then  'f3
        If Len(DBGrid2.columns("producto")) > 0 And Len(DBGrid2.columns("linea")) > 0 Then
            ingreso_tallas "" & DBGrid2.columns("linea")

        End If
   
    End If

    If KeyCode = &H77 Then  'f8 OBSERVACIONES
        If Len(DBGrid2.columns("producto")) > 0 Then
            ingreso_locales

        End If

    End If

    If KeyCode = &H28 Then  'flecha abajo inserta una nueva
        Exit Sub

        If DBGrid2.Col = 0 Then
            ir_ultimo

            If Len(DBGrid2.columns("producto")) > 0 And Len(DBGrid2.columns("descripcio")) > 0 And Len(DBGrid2.columns("unidad")) > 0 And Len(DBGrid2.columns("cantidad")) > 0 And Len(DBGrid2.columns("factor")) > 0 And Len(DBGrid2.columns("precio")) > 0 Then

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

    Dim sdx      As Double

    Dim found    As Integer

    Dim xpreciox As Double

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
        DBGrid2.Enabled = True
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
        If "" & mytable11.Fields("nosaldo") = "S" Then
            If familia_saldo("" & DBGrid2.columns("familia")) = 0 Then
                If consulta_saldo("" & DBGrid2.columns("producto"), Val("" & DBGrid4.columns(1)), 1) <= 0 Then
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
        'Data2.Recordset.Fields("precio") = "" & DBGrid4.Columns(3)
        'Data2.Recordset.Update
        'MsgBox DBGrid4.Row
        DBGrid2.columns("nroprecio") = "" & (DBGrid4.Row + 1)
        DBGrid2.columns("unidad") = "" & DBGrid4.columns(0)
        DBGrid2.columns("factor") = Val("" & DBGrid4.columns(1))
        DBGrid2.columns("precio") = xpreciox
        sdx = Val("" & DBGrid2.columns("cantidad")) * Val("" & DBGrid2.columns("precio"))
        DBGrid2.columns("total") = sdx
        calcula_igv 0
        'found = sumar_detalle()
        Frame5.Visible = False

        'antes estaba para que se vaya al final
        If Len(DBGrid2.columns("producto")) > 0 And Len(DBGrid2.columns("linea")) > 0 Then
            DBGrid2.Col = 3
            DBGrid2.Enabled = True
            DBGrid2.SetFocus
            ingreso_tallas "" & DBGrid2.columns("linea")
        Else
            DBGrid2.Enabled = True
            DBGrid2.Col = 3
            DBGrid2.SetFocus
            sumar_reforzar

            'sumar_reforzar
            'found = sumar_detalle()
            'DBGrid2.Col = 0
            'DBGrid2.Row = DBGrid2.VisibleRows - 1
            'DBGrid2.SetFocus
        End If

        'Command8_Click
        'End If
        'End If
    End If

End Sub

Private Sub DBGrid4_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, _
                                    StartLocation As Variant, _
                                    ByVal ReadPriorRows As Boolean)

    Dim dR            As Integer

    Dim row_num       As Integer

    Dim R             As Integer

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

                xruc = Trim(codigo)
                xnombre = Trim(nombre)

                If Len(Trim(cproven)) > 0 Then
                    xvendedor = cproven

                End If

                Framefp.Enabled = False
                Frame7.Visible = True
                Framefp.Enabled = False
                xtipo.SetFocus
                Exit Sub

            End If

            dbgrid10.Enabled = True
            dbgrid10.SetFocus

    End Select

End Sub

Private Sub DBGrid9_BeforeColEdit(ByVal ColIndex As Integer, _
                                  ByVal KeyAscii As Integer, _
                                  Cancel As Integer)

    Dim found1 As Double

    If ColIndex <> 2 Then
        Cancel = True
        Exit Sub

    End If

    Select Case ColIndex

        Case 2

            If Len("" & dbgrid9.columns(0)) = 0 Then
                Cancel = True
                Exit Sub

            End If

    End Select

End Sub

Private Sub DBGrid9_BeforeColUpdate(ByVal ColIndex As Integer, _
                                    OldValue As Variant, _
                                    Cancel As Integer)

    Dim found1 As Double

    Select Case ColIndex

        Case 2

            If Not IsNumeric("" & dbgrid9.columns(2)) Then
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
            
            If "" & Data9.Recordset.Fields("acu") = "C" Then 'credito
                If Val("" & dbgrid9.columns(2)) > Val(stxtotals) Then
                    MsgBox "Debe ingresar valor exacto", 48, "Aviso"
                    Cancel = True
                    Exit Sub

                End If

                found1 = valida_deposito("" & Data9.Recordset.Fields("codigo"), "" & Data9.Recordset.Fields("orden"), 0)

                If found1 < Val("" & dbgrid9.columns(2)) Then
                    MsgBox "No existe Saldo ", 48, "Aviso"
                    Cancel = True
                    Exit Sub

                End If

            End If
            
            '-------------------
            If verifica_fpago("" & dbgrid9.columns("fpago")) = "V" Then
                found1 = suma_pedidos("" & codigo)

                If found1 <= 0 Then
                    MsgBox "Cantidad Mayor que el saldo del pedido ", 48, "Aviso"
                    Cancel = True
                    Exit Sub

                End If

                If found1 > 0 Then
                    If found1 < Val("" & dbgrid9.columns(2)) Then
                        MsgBox "Cantidad Mayor que el saldo del pedido ", 48, "Aviso"
                        Cancel = True
                        Exit Sub

                    End If

                End If

            End If
            
            If "" & Data9.Recordset.Fields("acu") = "I" Or "" & Data9.Recordset.Fields("acu") = "K" Then 'valida el deposito bancario
                If Val("" & dbgrid9.columns(2)) > Val(stxtotals) Then
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

                If found1 < Val("" & dbgrid9.columns(2)) Then
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

        Select Case dbgrid9.Col

            Case 2

                If Len("" & dbgrid9.columns(2)) > 0 Then Exit Sub
                If Val("" & dbgrid9.columns(2)) = 0 Then
                    'If "" & Data9.Recordset.Fields("acu") = "H" Then 'valida el deposito bancario
                    '   DBGrid9.SetFocus
                    '   Exit Sub
                    'End If
                
                    If verifica_fpago("" & dbgrid9.columns("fpago")) = "V" Then
                        found1 = suma_pedidos("" & Data9.Recordset.Fields("codigo"))

                        If found1 <= 0 Then
                            MsgBox "Cantidad Mayor que el saldo del pedido ", 48, "Aviso"
                            dbgrid9.SetFocus
                            Exit Sub

                        End If

                        If found1 > 0 Then
                            If found1 < Val(stxtotals) Then
                                MsgBox "Cantidad Mayor que el saldo del pedido ", 48, "Aviso"
                                dbgrid9.SetFocus
                                Exit Sub

                            End If

                        End If

                    End If
               
aml:
               
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

                        If Len(Trim(cproven)) > 0 Then
                            xvendedor = cproven

                        End If

                        xruc = Trim(codigo)

                        If "" & mytable11.Fields("habilitanota") = "S" Then
                            If Val(ttxtotals) <= Val("" & mytable11.Fields("siventa")) Then
                                xtipo = "5"

                            End If

                        End If
                  
                        xnombre = nombre
                        Frame7.Visible = True
                        Framefp.Enabled = False
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

    Dim found As Integer

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

    found = abre_puerto("" & mytable11.Fields("capuerto"), Val("" & mytable11.Fields("catipo")), "" & mytable11.Fields("gavetacola"))

End Sub

Private Sub dcodigo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(dcodigo) > 0 Then
        If Len(telefono) < 7 Then
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

    Dim sw    As Integer

    Dim found As Integer

    flag_clave1 = 0
    tconcla.X = "CUADRE"
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
    tcuadrc1.horaf = "24"
    tcuadrc1.Caption = "CUADRE PARCIAL DEL DIA"
    tcuadrc1.Show 1

End Sub

Private Sub dcvendedor_Click()

    'If dcvendedor <> "%" Then
    'End If
    If dcvendedor <> "%" Then
        ddvendedor = extra_loquesea(dcvendedor.Text)
        Data2.Recordset.Edit
        Data2.Recordset.Fields("vendedor") = ddvendedor
        Data2.Recordset.Update
        Command8_Click

    End If

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

Private Sub dj7743400_Click()
    repctaxc.xcuentaco = "cuentac"
    repctaxc.XCUENTACO1 = "cuentacd"
    repctaxc.acu = "V"
    repctaxc.Show 1

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

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    If dbgrid6.Visible = True Then Exit Sub
    If Framefp.Visible = True Then Exit Sub
    If Frame7.Visible = True Then Exit Sub
    If Frame6.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub

    If Len(Trim("" & mytable11.Fields("tipore"))) = 0 Then
        MsgBox "No se ha definido el tipo recibo en parametros caja ", 48, "Aviso"
        Exit Sub

    End If

    mytablex.Open "SELECT * FROM tipo where  tipo='" & "" & mytable11.Fields("tipore") & "' and tipodoc='V'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si existe
        MsgBox "No existe tipo recibo definido en parametros Caja ", 48, "Aviso"
        mytablex.Close
        Exit Sub

    End If

    mytablex.Close

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
    trecaja.tipo = "" & mytable11.Fields("tipore")
    trecaja.tipo.Enabled = False
    trecaja.pagocash.Visible = True
    trecaja.pagocash.Value = 1

    fgusuario = "_r" & gusuario
    trecaja.Combo2.Enabled = True
    trecaja.xcuentaco = "cuentaC"
    trecaja.XCUENTACO1 = "cuentaCd"
    trecaja.tipoclie = "C"

    trecaja.Caption = "EGRESO DINERO"
    trecaja.local1 = "" & "" & mytable11.Fields("local")
    trecaja.serie = "" & mytable11.Fields("seriere")
    'trecaja.local1.Enabled = False
    trecaja.afecta = "C"
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

    menu_carga_guia
    Exit Sub

End Sub

Private Sub dk89230_Click()

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

Sub menu_carga_guia()

    Dim found As Integer

    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    opcion1 = "18500"
    sw_consulta = 0
    found = sql_consulta(1)

End Sub

Private Sub dli992323_Click()
    Label14_Click

End Sub

Private Sub dlko343_Click()

    If dbgrid6.Visible = True Then Exit Sub
    If Framefp.Visible = True Then Exit Sub
    If Frame7.Visible = True Then Exit Sub
    If Frame6.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub

    'If "" & mytable11.Fields("limpiapantalla") = "S" Then
    flag_clave1 = 0
    tconcla.X = "ANULA"  '
    tconcla.Show 1

    If flag_clave1 <> 1 Then  'si es descongela
        DBGrid2.SetFocus
        Exit Sub

    End If

    'End If

    cgusuario = gocabeza
    dgusuariog = godetalle
    menu_anula1

End Sub

Private Sub dlo2323_Click()

    Dim found As Integer

    If Framefp.Visible = True Then Exit Sub
    If Frame7.Visible = True Then Exit Sub
    If Frame6.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    If Frame9.Visible = True Then Exit Sub
    If Val(txtotal) > 0 Then
        MsgBox "No deben existir Productos ", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    If "" & mytable11.Fields("clavecongela") = "S" Then
        flag_clave1 = 0
        tconcla.X = "CONGELA"
        tconcla.Show 1

        If flag_clave1 = 1 Then  'si es descongela
            GoTo amj
            Exit Sub

        End If

        Exit Sub

    End If

amj:
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    opcion1 = "150"
    sw_consulta = 0
    found = sql_consulta(1)
    dbGrid1.Enabled = True

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
    found = sumar_detalle()
    tiposervicio1 = "Autoservicio"
    flag_servicio = "A"
    DBGrid2.SetFocus

End Sub

Private Sub dlo3434_Click()

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

End Sub

Private Sub dloco343_Click()

    If dbgrid6.Visible = True Then Exit Sub
    Frame9.Caption = "CONGELA PEDIDOS INGRESADOS"
    Label25_Click

End Sub

Private Sub dmo8833_Click()

End Sub

Private Sub dmoi434_Click()

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

Private Sub fechanac_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    found = valida()

    If found = 0 Then
        Exit Sub

    End If

    tiposervicio1 = "DELIVERY"
    flag_servicio = "D"
    CAMPO1 = telefono
    codigo = dcodigo
    nombre = dnombre
    Frame2.Visible = False
    DBGrid2.SetFocus

End Sub

Private Sub fechanac_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        referencia.SetFocus
        Exit Sub

    End If

End Sub

Private Sub Form_Activate()
    ' CUANDO ES UNA SEPARACION SE DEBE TENER CONFIGURADO EL TIPO PEDIDO EN PARAMECA Y LA SERIE Y EL NUMERO
    '

    Dim found As Integer

    Label34 = dicmoneda
    Label13 = "Dia:" & dia
    tdeliver.Caption = nombre_sistema & " " & mytable11.Fields("descripcio") & " "

    If flag_carga <> "S" Then
        found = busca_paridad()
        sql_detalle
        found = sumar_detalle()
        cajero = "" & gusuario
        flag_carga = "S"
        'pedido.SetFocus
        DBGrid2.refresh
        DBGrid2.Col = 0
        DBGrid2.Row = DBGrid2.VisibleRows - 1
        DBGrid2.SetFocus
      
    End If

    uvueltos = dicmoneda & ":" & Format(Val("" & mytable11.Fields("uvueltos")), nrodecimal)
    uvueltod = "US$:" & Format(Val("" & mytable11.Fields("uvueltod")), nrodecimal)

    If "" & mytable11.Fields("terminal") = "T" Then

        'MsgBox "Hola"
        'pedido.SetFocus
    End If

    found = leer_visorcaja("SISTEMA CALIPSO", "VERSION 5.0")

End Sub

Sub cargar_grafico1()

    On Error GoTo cmd7779_err

    Image1.Picture = LoadPicture(globalpath & "\ico\cajaper.jpg")
    Exit Sub
cmd7779_err:
    MsgBox "" & error$
    Exit Sub

End Sub

Sub sql_detalle()

    Dim buf   As String

    Dim found As Integer

    On Error GoTo cmd34_err

    buf = "select * from " & dgusuario & " order by hora"
    Data2.Connect = "foxpro 2.5;"
    Data2.DatabaseName = globaldir
    Data2.RecordSource = buf
    Data2.refresh
    DBGrid2.refresh
    'found = sumar_detalle()
    'DBGrid2.Row = DBGrid2.VisibleRows - 2
    'DBGrid2.Col = 0
    'DBGrid2.SetFocus
    Exit Sub
cmd34_err:
    MsgBox "Error en select " & error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Load()

    Dim found As Integer

    Dim sdx   As Double

    Dim xx    As String

    Dim I     As Integer

    nrodecimal = "0.00"

    If "" & mytable11.Fields("decimal") = "3" Then
        nrodecimal = "0.000"

    End If

    moneda = "" & mytable11.Fields("moneda")
    caja = "" & mytable11.Fields("caja")
   
    DBGrid2.columns("precio").NumberFormat = nrodecimal
    DBGrid2.columns("total").NumberFormat = nrodecimal
    carga_dcvendedor
    cargas_iniciales
    found = busca_paridad()
    'sql_detalle
    cajero = "" & gusuario
    'flag_carga = "S"
    'sumar_detalle
    tiposervicio1 = "Autoservicio"
    flag_servicio = "A"
    sql_detalle
    icerrar_puertosmscomm
    found = sumar_detalle()
    xx = busca_parame1("", 2)

    If "" & mytable11.Fields("terminal") = "T" Then
        menju232.Visible = False
        dlo2342.Visible = False
        'dek7834.Visible = False
        inu781.Visible = False
        djk7822.Visible = False
        cuj6721.Visible = False
        'Frame10.Visible = True
        'Label32.Visible = True
        pedido.Visible = True

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
    carga_local
    carga_clase_sunat
   
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

    icerrar_puertosmscomm
    cerrar_puerto
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

    Dim found As Integer

    If dbgrid6.Visible = True Then Exit Sub
    If Framefp.Visible = True Then Exit Sub
    If Frame7.Visible = True Then Exit Sub
    If Frame6.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    '-------------------------------------
    flag_clave1 = 0
    tconcla.X = "DESCUENTO"
    tconcla.Show 1

    If flag_clave1 = 0 Then  'si es descongela
        'Cancel = True
        Exit Sub

    End If

    Trecarg.total = txtotal
    Trecarg.Show 1

    grabar_descto

End Sub

Private Sub Image1_Click()

    'frmain.Show 1
End Sub

Private Sub inu781_Click()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

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
    If Len(Trim("" & mytable11.Fields("tipori"))) = 0 Then
        MsgBox "No se ha definido el tipo recibo en parametros caja ", 48, "Aviso"
        Exit Sub

    End If

    mytablex.Open "SELECT * FROM tipo where  tipo='" & "" & mytable11.Fields("tipori") & "' and tipodoc='W'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si existe
        MsgBox "No existe tipo recibo definido en parametros Caja ", 48, "Aviso"
        mytablex.Close
        Exit Sub

    End If

    mytablex.Close

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

    '

    fgusuario = "_r" & gusuario
    trecaja.pagocash.Visible = True
    trecaja.pagocash.Value = 1
    trecaja.Combo2.Enabled = True

    trecaja.xcuentaco = "cuentac"
    trecaja.XCUENTACO1 = "cuentacd"
    trecaja.tipoclie = "C"
    trecaja.tipo = "" & mytable11.Fields("tipori")
    trecaja.tipo.Enabled = False
    trecaja.Caption = "INGRESO DINERO"
    trecaja.afecta = "C"
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

End Sub

Private Sub jur9012_Click()

    Dim sw As Integer

    flag_clave1 = 0
    tconcla.X = "CUADRE"
    tconcla.Show 1

    If flag_clave1 <> 1 Then  'si es descongela
        DBGrid2.SetFocus
        Exit Sub

    End If
    
    opcion1 = "10"
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
    tcuadrc1.Caption = "UNIDADES VENDIDAS GRUPOS"
    tcuadrc1.check3d1.Visible = True
    tcuadrc1.check3d2.Visible = True
    tcuadrc1.check3d3.Visible = True
    tcuadrc1.Show 1

End Sub

Private Sub Label11_Click()
    busca_ocurrencia

End Sub

Private Sub Label14_Click()

    Dim found As Integer

    If Frame2.Visible = True Then Exit Sub
    If "" & mytable11.Fields("limpiapantalla") = "S" Then
        flag_clave1 = 0
        tconcla.X = "CLEAR"
        tconcla.Show 1

        If flag_clave1 = 0 Then  'si es descongela
            Exit Sub

        End If

    End If

    If MsgBox("Desea Borrar ??", 1, "Aviso") <> 1 Then Exit Sub
    borrar_todo
    '''sql_detalle
    found = sumar_detalle()
    tiposervicio1 = "Autoservicio"
    flag_servicio = "A"

End Sub

Sub borra_congela()

    Dim found As Integer

    If Frame2.Visible = True Then Exit Sub
    'If MsgBox("Desea Borrar ??", 1, "Aviso") <> 1 Then Exit Sub
    borrar_todo
    '''sql_detalle
    found = sumar_detalle()
    tiposervicio1 = "Autoservicio"
    flag_servicio = "A"

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

Private Sub Label17_Click()
    dlko343_Click

End Sub

Private Sub Label19_Click()

    Dim buf  As String

    Dim buf1 As String

    If Len(Trim(codigo)) = 11 Then
        nombre = ""
        xdireccion = ""
        buf1 = ""
        buf = Trim(OTROPOS(Trim(codigo), buf1))

        If Len(buf) > 0 Then
            nombre = buf
            xdireccion = buf1

        End If

    End If

End Sub

Private Sub Label2_Click()
    'valida_camara
    proceso_cierre_automatico "EFECTIVO"

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

Private Sub Label23_Click()
    proceso_cierre_automatico "TARJETACREDITO"

End Sub

Private Sub Label25_Click()

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
    clavecongela = ""
    xcongelax.SetFocus

End Sub

Private Sub Label26_Click()
    Tsms.Show 1

End Sub

Private Sub Label27_Click()
    dlo2323_Click

End Sub

Private Sub Label3_Click()
    TRUCLINE.viene = "CODIGO"
    TRUCLINE.Show 1

End Sub

Private Sub Label31_Click()

    If "" & mytable11.Fields("terminal") = "T" Then
        MsgBox "1.No permitido en Pedido", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    dhyori83_Click

End Sub

Private Sub Label32_Click()

    If dbgrid6.Visible = True Then Exit Sub
    If Framefp.Visible = True Then Exit Sub
    If Frame7.Visible = True Then Exit Sub
    If Frame6.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub

    If "" & mytable11.Fields("terminal") = "T" Then
        MsgBox "No permitido en Pedido", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    '-------------------------------------
    cgusuario = gocabeza
    dgusuariog = godetalle
    menu_delivery
    Exit Sub

End Sub

Private Sub Label33_Click()

    If Len(DBGrid2.columns("producto")) = 0 Then
        Exit Sub

    End If

    DBGrid2.Col = 57
    DBGrid2.SetFocus

End Sub

Private Sub Label36_Click()

    'TRUCLINE.viene = "XRUC"
    'TRUCLINE.Show 1
    Dim buf  As String

    Dim buf1 As String

    If Len(Trim(xruc)) = 11 Then
        xnombre = ""
        xdireccion = ""
        buf1 = ""
        buf = Trim(OTROPOS(Trim(xruc), buf1))

        If Len(buf) > 0 Then
            xnombre = buf
            xdireccion = buf1

        End If

    End If

End Sub

Private Sub Label37_Click()
    proceso_cierre_automatico "CREDITO"

End Sub

Private Sub Label4_Click()
    hydes8912_Click

End Sub

Private Sub Label44_Click()
    tncr1.local1 = Trim(mytable11.Fields("local"))
    tncr1.Show 1

End Sub

Private Sub Label5_Click()

    If "" & mytable11.Fields("terminal") = "T" Then
        MsgBox "2.No permitido en Pedido", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    proceso_guiaremision

End Sub

Sub proceso_guiaremision()

    Dim found As Integer

    'If "" & mytable11.Fields("terminal") = "T" Then
    '   MsgBox "No permitido en Pedido", 48, "Aviso"
    '   DBGrid2.SetFocus
    '  Exit Sub
    'End If
    If Frame2.Visible = True Then Exit Sub
    found = sumar_detalle()

    If found = 0 Then
        MsgBox "debe de Existir un Precio=0", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    'If Val(txtotal) = 0 Then
    '   DBGrid2.SetFocus
    '   Exit Sub
    'End If
    opcion1 = "1000"
    'Label36.Caption = "QueAlmacen"
    local1 = "GUIAREMISON"
    xruc = Trim(codigo)
    xnombre = nombre
    Frame7.Visible = True
    Framefp.Enabled = False
    xtipo = Trim("" & mytable11.Fields("tipoot"))
    xtipo.SetFocus
    Exit Sub

End Sub

Private Sub Label53_Click()

    Dim found As Integer

    flag_clave1 = 0
    tconcla.X = "RELOJ"
    tconcla.Show 1

    If flag_clave1 <> 1 Then  'si es descongela
        DBGrid2.Col = 0
        DBGrid2.Row = DBGrid2.VisibleRows - 1
        DBGrid2.SetFocus
        Exit Sub

    End If

    tingper.Show 1
    DBGrid2.Col = 0
    DBGrid2.Row = DBGrid2.VisibleRows - 1
    DBGrid2.SetFocus
 
    Exit Sub

    'If "" & mytable11.Fields("terminal") = "T" Then
    '   MsgBox "No permitido en Pedido", 48, "Aviso"
    '   DBGrid2.SetFocus
    '  Exit Sub
    'End If
    'MsgBox "Hola"
    If Frame2.Visible = True Then Exit Sub
    found = sumar_detalle()

    If found = 0 Then
        MsgBox "debe de Existir un Precio=0", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    If Val(txtotal) = 0 Then
        DBGrid2.SetFocus
        Exit Sub

    End If

    opcion1 = "1000"
    Label36.Caption = "Almac.Fuente."
    local1.Visible = True
    xruc = Trim(codigo)
    xnombre = nombre
          
    Frame7.Visible = True
    Framefp.Enabled = False
    xtipo.SetFocus
    Exit Sub

End Sub

Private Sub Label54_Click()

    Dim found As Integer

    If "" & mytable11.Fields("terminal") = "T" Then
        MsgBox "3.No permitido en Pedido", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    If Frame2.Visible = True Then Exit Sub
    found = sumar_detalle()
    DBGrid2.Col = 0
    DBGrid2.Row = DBGrid2.VisibleRows - 1
    DBGrid2.SetFocus
               
    If found = 0 Then
        MsgBox "debe de Existir un Precio=0", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    If found = 2 Then
        If "" & mytable11.Fields("vdetalle") = "S" Then
            MsgBox "No existe Vendedor", 48, "Aviso"
            DBGrid2.SetFocus
            Exit Sub

        End If

    End If

    If Val(txtotal) = 0 Then
        DBGrid2.SetFocus
        Exit Sub

    End If

    opcion1 = "1000"
    'Label36.Caption = "QueAlmacen"
    local1 = "PEDIDO"
    xruc = Trim(codigo)
    xnombre = nombre
    Frame7.Visible = True
    Framefp.Enabled = False
    xtipo = "Q"
    xtipo.SetFocus
    Exit Sub

End Sub

Private Sub Label55_Click() 'aqui es donde se totaliza la venta

    Dim found As Integer

    Dim sdx   As Double

    If "" & mytable11.Fields("terminal") = "T" Then

        'MsgBox "4.No permitido en Pedido", 48, "Aviso"
        'dbgrid2.SetFocus
        'Exit Sub
    End If

    If Frame2.Visible = True Then Exit Sub
    local1 = ""
    local1.Visible = False
    found = sumar_detalle()
    DBGrid2.Col = 0
    DBGrid2.Row = DBGrid2.VisibleRows - 1
    DBGrid2.SetFocus
               
    If found = 0 Then
        MsgBox "debe de Existir un Precio=0", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    If found = 2 Then
        If "" & mytable11.Fields("vdetalle") = "S" Then
            MsgBox "No existe Vendedor", 48, "Aviso"
            DBGrid2.SetFocus
            Exit Sub

        End If

    End If

    If flag_percepcion = "S" Then
        If Len(Trim("" & codigo)) = 0 Then
            MsgBox "Existe Percepcion ,Debe ponerse Dato Cliente ", 48, "Aviso"
            codigo.SetFocus
            Exit Sub

        End If

    End If

    If Val(txtotal) = 0 Then
        If exisdev <> -10 Then  'si existe devolucion
            DBGrid2.SetFocus
            Exit Sub

        End If

    End If

    If "" & mytable11.Fields("clienteo") = "S" Then
        If Len(nombre) = 0 Then
            MsgBox "Ingrese  Cliente ", 48, "Aviso"
            codigo.SetFocus
            Exit Sub

        End If

    End If

    If mytable11.Fields("terminal") = "T" Or Val(acuenta) > 0 And Len(petipo) = 0 Then 'pedidos o a cuenta ha dado
        'MsgBox "ojo esto estaba anterior"
        xruc = codigo
        xnombre = nombre
        Frame7.Visible = True
        Framefp.Enabled = False

        If Val(acuenta) > 0 Then
            xtipo = "" & mytable11.Fields("tipope")

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
    'MsgBox ""
    '-------------------------- xxxxx ----------------
    '-
    '-
    '-
    '-------------------------------------------------
    sdx = Val(rtxtotal)
    ttxtotals = Format(sdx, nrodecimal)

    sdx = Val(rtxtotald)
    ttxtotald = Format(sdx, nrodecimal)

    sdx = Val(rtxtotal)
    stxtotals = Format(sdx, nrodecimal)

    sdx = Val(rtxtotald)
    stxtotald = Format(sdx, nrodecimal)
    Framefp.Visible = True
    Framefp.Enabled = True
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
    proceso_cierre_automatico "DOLAR"

End Sub

Private Sub Label59_Click()

    Dim found As Integer

    If "" & mytable11.Fields("terminal") = "T" Then
        MsgBox "No permitido en Pedido", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Imprimir Pre Cuenta", 1, "Aviso") <> 1 Then
        found = ir_ultimo_registrox()
        DBGrid2.Col = 0
        DBGrid2.Row = DBGrid2.VisibleRows - 1
        DBGrid2.SetFocus 'found = sumar_detalle()
   
        DBGrid2.SetFocus
        Exit Sub

    End If

    imprime_precuenta
    found = ir_ultimo_registrox()

    If found = 0 Then
        Data2.refresh
        Exit Sub

    End If

    DBGrid2.Col = 0
    DBGrid2.Row = DBGrid2.VisibleRows - 1
    DBGrid2.SetFocus

End Sub

Private Sub Label60_Click()

    If "" & mytable11.Fields("terminal") = "T" Then
        MsgBox "No permitido en Pedido", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    dlo3434_Click

End Sub

Private Sub Label61_Click()

    Dim found As Integer

    dj78232_Click
    Exit Sub
    'If "" & mytable11.Fields("terminal") = "T" Then
    '   MsgBox "No permitido en Pedido", 48, "Aviso"
    '   DBGrid2.SetFocus
    '  Exit Sub
    'End If
    Exit Sub

    If Frame2.Visible = True Then Exit Sub
    found = sumar_detalle()
    DBGrid2.Col = 0
    DBGrid2.Row = DBGrid2.VisibleRows - 1
    DBGrid2.SetFocus
               
    If found = 0 Then
        MsgBox "debe de Existir un Precio=0", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    If found = 2 Then
        If "" & mytable11.Fields("vdetalle") = "S" Then
            MsgBox "No existe Vendedor", 48, "Aviso"
            DBGrid2.SetFocus
            Exit Sub

        End If

    End If

    If Val(txtotal) = 0 Then
        DBGrid2.SetFocus
        Exit Sub

    End If

    opcion1 = "1000"
    Label36.Caption = "Almac.Dest."
    local1.Visible = True
    xruc = Trim(codigo)
    xnombre = nombre
    Frame7.Visible = True
    Framefp.Enabled = False
    xtipo.SetFocus
    Exit Sub

End Sub

Private Sub Label62_Click()

    If "" & mytable11.Fields("terminal") = "T" Then
        MsgBox "No permitido en Pedido", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    dcaj8923_Click

End Sub

Private Sub Label65_Click()

End Sub

Private Sub losao94_Click()

    Dim found As Integer

    'MsgBox opcion1
    'If tmconsulta.State = 1 Then tmconsulta.State = 0
    If Frame3.Visible = True Then
        Frame3.Visible = False
        Frame3.Enabled = False
        DBGrid2.Enabled = True
        DBGrid2.Col = 0
        DBGrid2.Row = DBGrid2.VisibleRows - 1
        DBGrid2.SetFocus
        Exit Sub

    End If

    If Frame5.Visible = True Then
        If Frame1.Visible = True Then
            Frame5.Visible = False
            dbGrid1.Enabled = True
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

                If tmconsulta.State = 1 Then tmconsulta.Close
                tcampo1.SetFocus
                Exit Sub

            End If

        End If

        If opcion1 = "200" Then
            If Frame1.Visible = True Then
                Frame1.Visible = False
                Frame1.Enabled = False

                If tmconsulta.State = 1 Then tmconsulta.Close
                tcampo4.SetFocus
                Exit Sub

            End If

        End If

        If opcion1 = "2800" Then
            If Frame1.Visible = True Then
                Frame1.Visible = False
                Frame1.Enabled = False

                If tmconsulta.State = 1 Then tmconsulta.Close
                tcampo3.SetFocus
                Exit Sub

            End If

        End If

        Frame6.Visible = False
        'dbgrid10.SetFocus
        Exit Sub

    End If

    If Frame7.Visible = True Then
        If opcion1 = "30" Or opcion1 = "300" Then
            If Frame1.Visible = True Then
                Frame1.Visible = False
                Frame1.Enabled = False

                If tmconsulta.State = 1 Then tmconsulta.Close
                xruc.SetFocus
                Exit Sub

            End If

        End If

        If opcion1 = "31" Then
            If Frame1.Visible = True Then
                Frame1.Visible = False
                Frame1.Enabled = False

                If tmconsulta.State = 1 Then tmconsulta.Close
                xvendedor.SetFocus
                Exit Sub

            End If

        End If
   
        If opcion1 = "3100" Then
            If Frame1.Visible = True Then
                Frame1.Visible = False
                Frame1.Enabled = False

                If tmconsulta.State = 1 Then tmconsulta.Close
                DBGrid2.Col = 0
                DBGrid2.Row = DBGrid2.VisibleRows - 1
                DBGrid2.SetFocus
                Exit Sub

            End If

        End If
   
        If opcion1 = "3100" Then
            If Frame1.Visible = True Then
                Frame1.Visible = False
                Frame1.Enabled = False

                If tmconsulta.State = 1 Then tmconsulta.Close
                xvendedor.SetFocus
                Exit Sub

            End If

        End If
   
        If opcion1 = "29" Then
            If Frame1.Visible = True Then
                Frame1.Visible = False
                Frame1.Enabled = False

                If tmconsulta.State = 1 Then tmconsulta.Close
                xtipo.SetFocus
                Exit Sub

            End If

        End If

        If opcion1 = "8" Then
            Frame7.Visible = False
            DBGrid2.Enabled = True
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
            Exit Sub

        End If

        Frame7.Visible = False
   
        If "" & mytable11.Fields("terminal") = "T" Or opcion1 = "9999" Then
            DBGrid2.Enabled = True
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
            Exit Sub

        End If

        If Framefp.Visible = True Then
            Framefp.Enabled = True
            dbgrid10.Enabled = True
      
            dbgrid10.SetFocus
            Exit Sub

        End If

        If opcion1 = "1000" Then
            Frame7.Visible = False
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
            Exit Sub

        End If

        Exit Sub

    End If

    'If Frame10.Visible = True Then
    If Framefp.Visible = True Then
        Framefp.Visible = False
        'Frame10.Visible = True
        DBGrid2.Enabled = True
        DBGrid2.Col = 0
        DBGrid2.Row = DBGrid2.VisibleRows - 1
        DBGrid2.SetFocus
        Exit Sub

    End If

    'End If

    If Frame4.Visible = True Then
        Frame4.Visible = False
        DBGrid2.Enabled = True
        DBGrid2.Col = 0
        DBGrid2.Row = DBGrid2.VisibleRows - 1
        DBGrid2.SetFocus
        Exit Sub

    End If

    If Frame6.Visible = True Then
        Frame6.Visible = False
        dbgrid10.Enabled = True
        dbgrid10.SetFocus
        Exit Sub

    End If

    If opcion1 = "31" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            xvendedor.SetFocus
            Exit Sub

        End If

    End If

    If opcion1 = "3100" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus

            Exit Sub

        End If

    End If

    If opcion1 = "23" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            tcampo1.SetFocus
            Exit Sub

        End If

    End If

    If opcion1 = "29" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            xtipo.SetFocus
            Exit Sub

        End If

    End If

    If opcion1 = "30" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            xruc.SetFocus
            Exit Sub

        End If

    End If

    If opcion1 = "8" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
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

            If tmconsulta.State = 1 Then tmconsulta.Close
            DBGrid2.Enabled = True
            telefono.SetFocus
            Exit Sub

        End If

    End If

    If opcion1 = "750" Or opcion1 = "13" Or opcion1 = "10" Or opcion1 = "14" Or opcion1 = "15" Or opcion1 = "15A" Or opcion1 = "100" Or opcion1 = "150" Or opcion1 = "370" Or opcion1 = "1500" Or opcion1 = "1900" Or opcion1 = "15000" Or opcion1 = "18500" Or opcion1 = "30000" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
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

            If tmconsulta.State = 1 Then tmconsulta.Close
            DBGrid2.Enabled = True

            If Len(telefono) < 7 Then
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

            If tmconsulta.State = 1 Then tmconsulta.Close
            DBGrid2.Enabled = True
            telefono.SetFocus
            Exit Sub

        End If

    End If

    If opcion1 = "12" Then
        If Frame1.Visible = True Then
            Frame1.Visible = False
            Frame1.Enabled = False

            If tmconsulta.State = 1 Then tmconsulta.Close
            DBGrid2.Enabled = True
            codigo.SetFocus
            Exit Sub

        End If

    End If

    If Frame2.Visible = True Then
   
        If Len(telefono) > 0 Or Len(nombre) > 0 Or Len(ddireccion) > 0 Or Len(fechanac) > 0 Or Len(codigo) > 0 Then
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
        Framefp.Visible = False
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
    tdeliver.Hide
    Unload tdeliver

End Sub

Private Sub MSComm1_OnComm()

    Select Case MSComm1.CommEvent

        Case comEvReceive ' Received RThreshold # of chars.
            InBuff = InBuff + MSComm1.input

    End Select

End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    correo.SetFocus

End Sub

Private Sub nombre_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        codigo.SetFocus
        Exit Sub

    End If

End Sub

Private Sub observa1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa2.SetFocus

End Sub

Private Sub observa2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa3.SetFocus

End Sub

Private Sub observa3_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    observa4.SetFocus

End Sub

Private Sub observa4_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

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

Private Sub pedido_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If Val(txtotal) > 0 Then
        MsgBox "No Deben existir Productos Ingresados", 48, "Aviso"
        pedido = ""
        pedido.SetFocus
        Exit Sub

    End If

    If KeyAscii = 27 Then
        pedido = ""
        DBGrid2.SetFocus
        Exit Sub

    End If

    If Len(pedido) = 0 Then
        pedido.SetFocus
        Exit Sub

    End If

    found = verifica_ticket_ingreso("" & pedido)

    If found = 0 Then
        MsgBox "No existe Ticket Ingreso", 48, "Aviso"
        pedido.SetFocus
        Exit Sub

    End If

    found = carga_ticket_ingreso()

    If found = 0 Then
        MsgBox "No se puede cargar ticket ingreso", 48, "Aviso"
        pedido.SetFocus
        Exit Sub

    End If

    ir_ultimo

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

    found = proceso_proforma(Trim("" & mytable11.Fields("local")), "P", "P", "" & pedido)
    carga_ticket_ingreso = found

End Function

Private Sub Picture2_Click()

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

Private Sub t1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t2.SetFocus

End Sub

Private Sub t10_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t11.SetFocus

End Sub

Private Sub t11_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t12.SetFocus

End Sub

Private Sub t12_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t13.SetFocus

End Sub

Private Sub t13_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t14.SetFocus

End Sub

Private Sub t14_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t15.SetFocus

End Sub

Private Sub t15_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t16.SetFocus

End Sub

Private Sub t16_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Command2_Click

End Sub

Private Sub t2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t3.SetFocus

End Sub

Private Sub t3_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t3.SetFocus

End Sub

Private Sub t4_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t4.SetFocus

End Sub

Private Sub t5_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t5.SetFocus

End Sub

Private Sub t6_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t6.SetFocus

End Sub

Private Sub t7_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t8.SetFocus

End Sub

Private Sub t8_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t9.SetFocus

End Sub

Private Sub t9_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    t10.SetFocus

End Sub

Private Sub tcampo1_KeyPress(KeyAscii As Integer)

    Dim found  As Integer

    Dim found1 As Double

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Frame6.Visible = False
        dbgrid10.Enabled = True
        dbgrid10.SetFocus
        Exit Sub

    End If

    If Frame6.Caption = "CREDITO" Or "" & dbgrid10.columns("tipo") = "I" Or "" & dbgrid10.columns("tipo") = "C" Or "" & dbgrid10.columns("tipo") = "H" Or "" & dbgrid10.columns("tipo") = "K" Or "" & dbgrid10.columns("tipo") = "V" Then
        If Len(Trim(tcampo1)) = 0 Then
            tcampo1.SetFocus
            Exit Sub

        End If

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

    'MsgBox "abc"
    If Len(tcampo1) > 0 Then
        found = busca_codigocl("" & tcampo1, 0)

    End If

    'MsgBox dbgrid10.columns("tipo")
    If ("" & dbgrid10.columns("tipo") = "C") Then
        If Len(Trim(tcampo1)) = 0 Then
            tcampo1.SetFocus
            Exit Sub

        End If

        tcampo2.SetFocus
        Exit Sub

    End If

    If ("" & dbgrid10.columns("tipo") = "I" Or "" & dbgrid10.columns("tipo") = "K") And found = 1 And Len(Trim("" & tcampo1)) > 0 Then '
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
        consulta_xruc1 "" & tcampo1

    End If

    'If KeyCode = &H26 Then
    '   tcampo3.SetFocus
    '   Exit Sub
    'End If

End Sub

Private Sub tcampo2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    If Len(tcampo2) = 0 Then

        'tcampo2.SetFocus
        'Exit Sub
    End If

    tcampo3.SetFocus

End Sub

Private Sub tcampo2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        tcampo1.SetFocus
        Exit Sub

    End If

End Sub

Private Sub tcampo3_KeyPress(KeyAscii As Integer)

    Dim found1 As Double

    Dim found  As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Frame6.Visible = False
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

        If acufp = "I" Or acufp = "K" Then  'si es cruce de pago adelantado cruza
            consulta_credito  '2800

        End If

    End If

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

            'MsgBox "Ingrese Entidad ", 48, "Aviso"
            'tcampo4 = ""
            'tcampo4.SetFocus
            'Exit Sub
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

Private Sub tcampo5_KeyPress(KeyAscii As Integer)

    Dim sdx    As Double

    Dim found  As Integer

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

        If "" & dbgrid10.columns("tipo") = "V" Then 'credito
            If Len(tcampo1) = 0 Then
                tcampo1.SetFocus
                Exit Sub

            End If

            If Len(tcampo2) = 0 Then
                tcampo2.SetFocus
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

        If Not IsNumeric(tcampo5) Then
            tcampo5 = ""
            tcampo5.SetFocus
            Exit Sub

        End If

        If Val(tcampo5) <= 0 Then
            tcampo5 = ""
            tcampo5.SetFocus
            Exit Sub

        End If

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
    If ("" & dbgrid10.columns("tipo") = "I" Or "" & dbgrid10.columns("tipo") = "K" Or "" & dbgrid10.columns("tipo") = "C") And found = 1 Then   '
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

    codigo = Trim(tcampo1)
    nombre = tcampo2
    Data9.Recordset.AddNew
    Data9.Recordset.Fields("descripcio") = "" & dbgrid10.columns(0)
    Data9.Recordset.Fields("fpago") = "" & dbgrid10.columns(1)
    Data9.Recordset.Fields("moneda") = "" & dbgrid10.columns(2)
    Data9.Recordset.Fields("codigo") = Trim(tcampo1)
    Data9.Recordset.Fields("nombre") = tcampo2
    Data9.Recordset.Fields("orden") = tcampo3
    Data9.Recordset.Fields("observa") = tcampo4
    Data9.Recordset.Fields("dias") = tcampo5
    Data9.Recordset.Fields("acu") = "" & dbgrid10.columns("tipo")
    Data9.Recordset.Update

    If Len(tcampo1) > 0 And Len(tcampo2) > 0 Then
        found = graba_cliente_credito1("" & tcampo1)

    End If

    Frame6.Visible = False
    dbgrid9.Row = dbgrid9.VisibleRows - 1
    dbgrid9.Col = 2
    dbgrid9.SetFocus
    Exit Sub
cmd8_err:
    Exit Sub

End Sub

Function valida_deposito(buf0 As String, buf As String, sw As Integer) As Double

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM chequemo where  transaccio='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        valida_deposito = Val("" & mytablex.Fields("saldo"))

        If sw = 1 Then
            tcampo1 = Trim("" & mytablex.Fields("codigo"))

        End If

    End If

    mytablex.Close

End Function

Sub graba_deposito(mytabley As ADODB.Recordset)

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    Dim nrecibe  As Double

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

    Dim sdx      As Double

    Dim nrecibe  As Double

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

Private Sub telefono_KeyPress(KeyAscii As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    Dim buf      As String

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

    Dim buf1  As String

    buf1 = ""

    If Len(buf) > 0 Then
        buf1 = " where telefono='" & buf & "'"

    End If

    Combo1.Clear
    Combo1.AddItem "deliveri.telefono"
    Combo1.AddItem "Clientes.Nombre"
    Combo1.AddItem "Clientes.Codigo"
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

    If Len(Trim(buf)) > 0 Then
        found = sql_consulta(1)
        Exit Function

    End If

    Set dbGrid1.DataSource = Nothing
    buffer.SetFocus
    consulta_cliente = 0

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

Sub consulta_xvendedor1()

    Dim found As Integer

    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.ListIndex = 0

    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    opcion1 = "3100"
    sw_consulta = 0
    found = sql_consulta(1)
    'dbGrid1.SetFocus

End Sub

Sub consulta_xruc(buf As String)

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

    If Len(Trim(buf)) > 0 Then
        found = sql_consulta(1)
        Exit Sub

    End If

    Set dbGrid1.DataSource = Nothing
    buffer.SetFocus
    'dbGrid1.SetFocus

End Sub

Sub consulta_almacen()

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

Sub consulta_xruc2(buf As String)

    Dim found As Integer

    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.ListIndex = 0

    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    opcion1 = "300"
    sw_consulta = 0

    If Len(Trim(buf)) > 0 Then
        found = sql_consulta(1)
        Exit Sub

    End If

    Set dbGrid1.DataSource = Nothing
    buffer.SetFocus

    'dbGrid1.SetFocus

End Sub

Sub consulta_xruc1(buf As String)

    Dim found As Integer

    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.ListIndex = 0

    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    opcion1 = "99"
    sw_consulta = 0

    If Len(Trim(buf)) > 0 Then
        found = sql_consulta(1)
        Exit Sub

    End If

    Set dbGrid1.DataSource = Nothing
    buffer.SetFocus

    'dbGrid1.SetFocus

End Sub

Sub consulta_xtipo()

    Dim found As Integer

    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.ListIndex = 0

    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = "%"
    opcion1 = "29"
    sw_consulta = 0
    found = sql_consulta(1)
    'dbGrid1.SetFocus

End Sub

Sub consulta_cliente1(buf As String)

    Dim found As Integer

    Frame1.Visible = True
    Frame1.Enabled = True
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.ListIndex = 0
    buffer = ""
    opcion1 = "12"
    sw_consulta = 0

    If Len(Trim(buf)) > 0 Then
        found = sql_consulta(1)
        Exit Sub

    End If

    Set dbGrid1.DataSource = Nothing
    buffer.SetFocus
    'dbGrid1.SetFocus

End Sub

Sub consulta_clientefp(buf As String)

    Dim found As Integer

    Frame1.Visible = True
    Frame1.Enabled = True
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.ListIndex = 0
    buffer = ""
    opcion1 = "23"
    sw_consulta = 0

    If Len(Trim(buf)) > 0 Then
        found = sql_consulta(1)
        Exit Sub

    End If

    Set dbGrid1.DataSource = Nothing
    buffer.SetFocus
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

        If Len(Trim(correo)) = 0 Then
            correo = "" & mytablex.Fields("correo")

        End If

        'dotipo = "" & mytablex.Fields("dotipo")
        'doserie = "" & mytablex.Fields("doserie")
        'donumero = "" & mytablex.Fields("donumero")
        'ruc = "" & mytablex.Fields("codigo1")
    End If

    mytablex.Close
 
End Function

Function busca_codigocl(buf As String, sw As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim buf1     As String

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM clientes where  codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        If sw = 0 Then
            codigo = Trim("" & mytablex.Fields("codigo"))
            nombre = "" & mytablex.Fields("nombre")

            If Len(Trim(correo)) = 0 Then
                correo = "" & mytablex.Fields("correo")

            End If

            'correo = Trim("" & mytablex.Fields("correo"))
            tcampo2 = "" & mytablex.Fields("nombre")

        End If

        If sw = 1 Then
            If Len(Trim(correo)) = 0 Then
                correo = "" & mytablex.Fields("correo")

            End If

            xruc = Trim("" & mytablex.Fields("codigo"))

            If Len(xnombre) = 0 Then
                xnombre = "" & mytablex.Fields("nombre")

            End If

            If Len(xdireccion) = 0 Then
                xdireccion = "" & mytablex.Fields("direccion")

            End If

        End If

        'If dbgrid10.columns("tipo") = "V" Then 'si en fpago es vale
        '   totpedido = "" & suma_pedidos("" & mytablex.Fields("codigo"))
        'End If
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

    Dim indx     As Integer

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
    'dotipo = "" & mytablex.Fields("dotipo")
    'doserie = "" & mytablex.Fields("doserie")
    'donumero = "" & mytablex.Fields("donumero")
    'ruc = "" & mytablex.Fields("codigo1")

End Sub

Function contador_telefonos(buf As String)

    Dim mytablex As New ADODB.Recordset

    Dim indx     As Integer

    indx = 0

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM telefono where  telefono='" & telefono & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        Do

            If mytablex.EOF Then Exit Do
            indx = indx + 1
            buf = Trim("" & mytablex.Fields("codigo"))
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    contador_telefonos = indx

End Function

Function valida()

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    If Len(telefono) < 7 Then
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
    If Len(Trim(dcodigo)) = 0 Then
        busca_correlativo 0

    End If

    'If mytabley.State = 1 Then mytabley.Close
    mytabley.Open "SELECT * FROM clientes where codigo='" & dcodigo & "'", cn, adOpenDynamic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        mytabley.Fields("nombre") = dnombre
        mytabley.Fields("direccion") = ddireccion
        mytabley.Fields("observa") = referencia
        mytabley.Fields("telefono") = telefono
        mytabley.Fields("correo") = Mid$(Trim(correo), 1, 60)
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

        mytabley.Fields("correo") = Mid$(Trim(correo), 1, 60)
        mytabley.Update
        busca_correlativo 1

    End If

    mytabley.Close
    mytablex.Open "SELECT * FROM deliveri where codigo='" & dcodigo & "' and telefono='" & telefono & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        mytablex.Fields("telefono") = telefono
        mytablex.Fields("codigo") = dcodigo
        mytablex.Fields("direccion") = ddireccion
        mytablex.Fields("referencia") = referencia
        mytablex.Update
    Else
        'mytablex.Fields("telefono") = telefono
        'mytablex.Fields("codigo") = dcodigo
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

    Dim buf1 As String

    Dim sdx  As Integer

    Dim I    As Integer

    buf1 = ""

    If flag_denisse = "1" Then
        sdx = 18 - Len(buf)

        For I = 1 To sdx
            buf1 = buf1 & "0"
        Next I

    End If

    buf1 = buf1 & buf

    'MsgBox buf1
    'buf1 = "SELECT * FROM productb where RIGHT(barras," & Len(Trim(buf)) & ")='" & buf & "'"
    'MsgBox buf1
    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM productb where barras='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        'MsgBox "nose"
        buf = "" & mytablex.Fields("producto")
        busca_equiva = 1
        mytablex.Close
        Exit Function

    End If

    mytablex.Close
    'MsgBox buf
    'mytablex.Open "SELECT * FROM producto where barras='" & buf & "'", cn, adOpenStatic, adLockOptimistic
    'If mytablex.RecordCount > 0 Then
    '   buf = "" & mytablex.Fields("producto")
    '   busca_equiva = 1
    'End If
   
    mytablex.Open "SELECT * FROM producto where barras='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = "" & mytablex.Fields("producto")
        busca_equiva = 1

    End If

    mytablex.Close

End Function

Function busca_producto(buf As String, sw As Integer, canti As String)

    Dim mytablex   As New ADODB.Recordset

    Dim mytabley   As New ADODB.Recordset

    Dim xbpeso     As String

    Dim buf1       As String

    Dim I          As Integer

    Dim ssw        As Integer

    Dim found      As Integer

    Dim buf11      As String

    Dim sw_balanza As Integer

    '------------------------------------
    'verificamos si es codigo barras
    xbpeso = ""
    sw_balanza = 0

    I = 0

    'MsgBox buf
    'If Mid$(buf, 1, 1) = "2" And (Mid$(buf, 2, 1) = "1" Or Mid$(buf, 2, 1) = "2" Or Mid$(buf, 2, 1) = "3" Or Mid$(buf, 2, 1) = "0") And Len(buf) = 13 Then     'balanza+ean 13
    If Mid$(buf, 1, 1) = "2" And (Mid$(buf, 2, 1) = "1" Or Mid$(buf, 2, 1) = "2" Or Mid$(buf, 2, 1) = "3" Or Mid$(buf, 2, 1) = "0") And Len(buf) = 13 Then     'balanza+ean 13
        xbpeso = Mid$(buf, 8, 2) & "." & Mid$(buf, 10, 3)
        xbpeso = Format(Val(xbpeso), "0.000")
        'buf = Mid$(buf, 5, 3)
        buf = Mid$(buf, 3, 5)
        'mytablex.Open "SELECT * FROM producto where codigobalanza='" & buf & "'", cn, adOpenStatic, adLockOptimistic
        'If mytablex.RecordCount > 0 Then
        '   buf = "" & mytablex.Fields("producto")
        'End If
        'mytablex.Close
        'MsgBox buf
     
        'MsgBox xbpeso
        'MsgBox buf
        'MsgBox buf
        'SE AGREGO PORQUE EXISTE NUEVO CAMPO BALANZA
        'buf = busca_balanza(buf)
        'If Len(buf) = 0 Then
        '   table1.Enabled = True
        '  Exit Function
        'End If
        'MsgBox buf
        sw_balanza = 1

    End If

    'MsgBox buf
    'MsgBox xbpeso
    If buf = "XXX" Then 'codigo libre
        carga_xxx
        busca_producto = 1
        Exit Function

    End If

    found = 0

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        found = busca_equiva(buf) 'busca en la table codigo barras

        If found = 0 Then
            Exit Function

        End If

        'MsgBox buf
        mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.Close
            Exit Function

        End If

    End If

    If "" & mytablex.Fields("estado") = "N" Then  'si no esta activo
        MsgBox "Producto no activo ", 48, "Aviso"
        mytablex.Close
        Exit Function

    End If

    'MsgBox "abc"
      
    If flag_especial = "S" Then
        buf11 = " select * from precio1 where producto='" & buf & "' and local='01' and codigo='" & codigo & "'"

        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open buf11, cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount > 0 Then
            GoTo amika6

        End If

        mytabley.Close

    End If

    'MsgBox "abc"
    buf11 = " select * from precios where producto='" & buf & "' and local='" & "" & mytable11.Fields("listap") & "'"

    If mytabley.State = 1 Then mytabley.Close
    mytabley.Open buf11, cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        MsgBox "No existe precio Programado ", 48, "Aviso"
        mytabley.Close
        mytablex.Close
        Exit Function

    End If

    'MsgBox "precio"
amika6:
   
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
    If sw_balanza = 0 Then  'si no es balanza de codigo barras
        If Val(canti) <= 0 Then
            If "" & mytable11.Fields("actbala") = "S" Then
                If "" & mytablex.Fields("peso") = "S" Then
ajk91:
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

    End If

    If Val(canti) <= 0 Then
        canti = "1"

    End If

vienepeso:
    busca_producto = 1

    If sw_balanza = 1 Then
        canti = Val(xbpeso)

    End If

    '---------------------------------------
    If sw = 0 Or sw = 2 Or sw = 1 Then
        graba_temporald mytablex, sw, canti, mytabley

    End If

    mytablex.Close
    mytabley.Close

End Function

Sub calcula_igv(sw As Integer)

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim tdscto      As Double

    Dim tdscto1     As Double

    Dim found       As Integer

    Dim xtivap      As Double

    Dim xtisc       As Double

    Dim xdetra      As Double

    Dim xpercepcion As Double

    Dim clase_sunat As Double

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

    xpercepcion = Val("" & DBGrid2.columns("total")) * Val("" & DBGrid2.columns("percepcion")) / 100
    DBGrid2.columns("tpercepcio") = xpercepcion
    
    Exit Sub
cmd4567_err:
    MsgBox "Error en Calcula Igv " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub calcula_sinigv()

    Dim sdx         As Double

    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim found       As Integer

    Dim xtivap      As Double

    Dim xpercepcion As Double

    'debe sumar el igv
    'dbgrid2.columns("neto") = Val("" & dbgrid2.columns("cantidad")) * Val("" & dbgrid2.columns("precio"))
    If Val("" & DBGrid2.columns("cantidad")) > 0 And Val("" & DBGrid2.columns("neto")) > 0 Then
        sdx = Val("" & DBGrid2.columns("neto")) * Val("" & DBGrid2.columns("deslipo")) / 100 'descuento
        DBGrid2.columns("descuento") = sdx  'descuento
        DBGrid2.columns("subtotal") = Val("" & DBGrid2.columns("neto")) - Val("" & DBGrid2.columns("descuento")) 'subtotal
        sdx = Val("" & DBGrid2.columns("subtotal")) * Val("" & DBGrid2.columns("igv")) / 100
        DBGrid2.columns("impuesto") = sdx 'igv
        DBGrid2.columns("total") = Val("" & DBGrid2.columns("subtotal")) + sdx 'neto
        sdx = Val("" & DBGrid2.columns("total")) / Val(DBGrid2.columns("cantidad"))
        DBGrid2.columns("precio") = sdx
        xtivap = Val("" & DBGrid2.columns("total")) * Val("" & DBGrid2.columns("ivap")) / 100
        DBGrid2.columns("tivap") = xtivap
   
        xpercepcion = Val("" & DBGrid2.columns("total")) * Val("" & DBGrid2.columns("percepcion")) / 100
        DBGrid2.columns("tpercepcio") = xpercepcion
   
    End If

End Sub

Function consulta_producto(buf As String)

    Dim found As Integer

    Dim xbuf  As String

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

    If Len(Trim(buffer)) > 0 Then
        found = sql_consulta(1)
        Exit Function

    End If

    Set dbGrid1.DataSource = Nothing
    buffer.SetFocus

End Function

Function consulta_inicial(buf As String)

    Dim buf1      As String

    Dim queprecio As String

    Combo1.Clear
    Combo1.AddItem "Producto.Descripcio"
    Combo1.ListIndex = 0

    Dim raconsulta As New ADODB.Recordset

    queprecio = "precioS.pventa1 as Precio "

    If Len(buf) > 0 Then
        buf1 = "select Producto.Descripcio,Producto.producto,Producto.Marca,precios.Unidad1 as Und1,precios.Factor1 as F," & queprecio & " ,Producto.Monedav as M,Producto.Familia,Producto.Subfamilia,Producto.seccion from producto  left join precios on producto.producto=precios.producto  where precios.local='" & "" & mytable11.Fields("listap") & "' and producto.descripcio like '" & buf & "%'"

    End If

    If Len(buf) = 0 Then
        buf1 = "select Producto.Descripcio,Producto.producto,Producto.Marca,precios.Unidad1 as Und1,precios.Factor1 as F," & queprecio & " ,Producto.Monedav as M,Producto.Familia,Producto.Subfamilia,Producto.seccion from producto  left join precios on producto.producto=precios.producto  where precios.local='" & "" & mytable11.Fields("listap") & "'"

    End If
   
    If raconsulta.State = 1 Then raconsulta.Close
    raconsulta.Open buf1, cn, adOpenStatic, adLockOptimistic

    If raconsulta.EOF = True And raconsulta.BOF = True Then
        raconsulta.Close
        buffer.SetFocus
        Exit Function

    End If
   
    Set dbGrid1.DataSource = raconsulta
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

Sub carga_combo2(buf As String)

    Dim I As Integer

    Combo2.Clear
    Combo2.AddItem buf

    For I = 1 To 9

        If buf <> Format(I, "00") Then
            Combo2.AddItem Format(I, "00")

        End If

    Next I

    Combo2.ListIndex = 0

End Sub

Sub carga_dbgrid4(uproducto As String, xlistab As String)

    Dim I               As Integer

    Dim xfoto           As String

    Dim mytablex        As New ADODB.Recordset

    Dim mytabley        As New ADODB.Recordset

    Dim sw              As Integer

    Dim xbodega         As String

    Dim xsaldo          As Double

    Dim xbuf            As String

    Dim xcosto          As Double

    Dim xmargen         As Double

    Dim xcostou         As Double

    Dim xfactor         As Double

    Dim xxr             As String

    Dim xxi             As String

    Dim zbuf            As String

    Dim xpreciox        As Double

    Dim dmoneda         As String

    Dim empaque_visible As String

    On Error GoTo cmd89111_err

    xcostou = 0

    For I = 0 To 9
        campo_precios(I).unidad = ""
        campo_precios(I).factor = ""
        campo_precios(I).precio = ""
        campo_precios(I).costo = ""
        campo_precios(I).margen = ""
        campo_precios(I).stock = ""
    Next I

    tproducto = uproducto
    'MsgBox uproducto
    xfactor = 1
    xbodega = "" & mytable11.Fields("bodega")
    xsaldo = 0
    xcosto = 0
    sw = 0

    If mytabley.State = 1 Then mytabley.Close
    mytabley.Open "SELECT * FROM almacen where local='" & Trim("" & "" & mytable11.Fields("local")) & "' and producto='" & uproducto & "' and bodega='" & xbodega & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        xsaldo = Val("" & mytabley.Fields("saldo"))

    End If

    mytabley.Close
    'MsgBox "x"
    '---buscamos los datos de productos
    dmoneda = "S"
    xfoto = ""
    descorto = ""
    empaque_visible = ""
    seccion = ""

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM producto where  producto='" & uproducto & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        seccion = "" & mytablex.Fields("seccion")
        xcostou = 0

        If "" & mytable11.Fields("vecocaja") = "S" Then
            xcostou = Val("" & mytablex.Fields("costou"))

        End If

        xfactor = Val("" & mytablex.Fields("factor"))
        descorto = "" & mytablex.Fields("presenta")
        dmoneda = "" & mytablex.Fields("monedav")
        xfoto = "" & mytablex.Fields("fotonombre")
        empaque_visible = "" & mytablex.Fields("empaque_visible")

        If Val(empaque_visible) = 0 Then
            empaque_visible = "10"

        End If

    End If

    lectura_grafico Trim("" & mytablex.Fields("producto"))
      
    mytablex.Close

    'carga_foto xfoto
    If Val(paridad) <= 0 Then
        paridad = "1"

    End If

    If mytablex.State = 1 Then mytablex.Close
      
    '-------------------------------------------

    If flag_especial = "S" Then
        zbuf = "SELECT * FROM precio1 where  producto='" & uproducto & "' and local='01' and codigo='" & codigo & "'"
        mytablex.Open zbuf, cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            GoTo amika7

        End If

        mytablex.Close

    End If

    zbuf = "SELECT * FROM precios where  producto='" & uproducto & "' and local='" & xlistab & "'"
    mytablex.Open zbuf, cn, adOpenStatic, adLockOptimistic
amika7:

    If mytablex.RecordCount > 0 Then
        xcosto = 0
        xpreciox = 0

        If Val(empaque_visible) > 0 Then
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

                '------------------------------------------------------------
                xcosto = xcostou / xfactor
                xcosto = xcosto * Val("" & mytablex.Fields("factor1"))
                campo_precios(0).unidad = "" & mytablex.Fields("unidad1")
                campo_precios(0).factor = Val("" & mytablex.Fields("factor1"))
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

        End If '

        '---------
   
        xcosto = 0

        If Val(empaque_visible) > 1 Then
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
                campo_precios(1).factor = Val("" & mytablex.Fields("factor2"))
                campo_precios(1).precio = xpreciox
                xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor2")))
                campo_precios(1).stock = "" & xbuf
                campo_precios(1).costo = "" & xcosto
                xmargen = 0

                If xcosto > 0 Then
                    xmargen = ((xpreciox - xcosto) * 100) / xcosto

                End If

                campo_precios(1).margen = "" & xmargen

            End If

        End If 'fin de nro empaques
   
        xcosto = 0

        If Val(empaque_visible) > 2 Then
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
   
                xcosto = xcostou / xfactor
                xcosto = xcosto * Val("" & mytablex.Fields("factor3"))
                campo_precios(2).unidad = "" & mytablex.Fields("unidad3")
                campo_precios(2).factor = Val("" & mytablex.Fields("factor3"))
                campo_precios(2).precio = xpreciox
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

        End If

        xcosto = 0

        If Val(empaque_visible) > 3 Then
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

                xcosto = xcostou / xfactor
                xcosto = xcosto * Val("" & mytablex.Fields("factor4"))
                campo_precios(3).unidad = "" & mytablex.Fields("unidad4")
                campo_precios(3).factor = Val("" & mytablex.Fields("factor4"))
                campo_precios(3).precio = xpreciox
                xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor4")))
                campo_precios(3).stock = "" & xbuf
                campo_precios(3).costo = "" & xcosto
                xmargen = 0

                If xcosto > 0 Then
                    xmargen = ((xpreciox - xcosto) * 100) / xcosto

                End If

                campo_precios(3).margen = "" & xmargen

            End If

        End If

        xcosto = 0

        If Val(empaque_visible) > 4 Then
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

                xcosto = xcostou / xfactor
                xcosto = xcosto * Val("" & mytablex.Fields("factor5"))
                campo_precios(4).unidad = "" & mytablex.Fields("unidad5")
                campo_precios(4).factor = Val("" & mytablex.Fields("factor5"))
                campo_precios(4).precio = xpreciox
                xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor5")))
                campo_precios(4).stock = "" & xbuf
                campo_precios(4).costo = "" & xcosto
                xmargen = 0

                If xcosto > 0 Then
                    xmargen = ((xpreciox - xcosto) * 100) / xcosto

                End If

                campo_precios(4).margen = "" & xmargen

            End If

        End If

        xcosto = 0

        If Val(empaque_visible) > 5 Then
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
   
                xcosto = xcostou / xfactor
                xcosto = xcosto * Val("" & mytablex.Fields("factor6"))
                campo_precios(5).unidad = "" & mytablex.Fields("unidad6")
                campo_precios(5).factor = Val("" & mytablex.Fields("factor6"))
                campo_precios(5).precio = xpreciox
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
                '   campo_precios(5).precio = 0
                'End If
            End If

            'MsgBox "xx"
        End If

        xcosto = 0

        If Val(empaque_visible) > 6 Then
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
   
                xcosto = xcostou / xfactor
                xcosto = xcosto * Val("" & mytablex.Fields("factor7"))
   
                campo_precios(6).unidad = "" & mytablex.Fields("unidad7")
                campo_precios(6).factor = Val("" & mytablex.Fields("factor7"))
                campo_precios(6).precio = xpreciox
                xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor7")))
                campo_precios(6).stock = "" & xbuf
                campo_precios(6).costo = "" & xcosto
                xmargen = 0

                If xcosto > 0 Then
                    xmargen = ((xpreciox - xcosto) * 100) / xcosto
         
                End If

                campo_precios(6).margen = "" & xmargen

            End If

        End If
   
        xcosto = 0

        If Val(empaque_visible) > 7 Then
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
                campo_precios(7).factor = Val("" & mytablex.Fields("factor8"))
                campo_precios(7).precio = xpreciox
                xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor8")))
                campo_precios(7).stock = "" & xbuf
                campo_precios(7).costo = "" & xcosto
                xmargen = 0

                If xcosto > 0 Then
                    xmargen = ((xpreciox - xcosto) * 100) / xcosto

                End If

                campo_precios(7).margen = "" & xmargen

            End If

        End If

        xcosto = 0

        If Val(empaque_visible) > 8 Then
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
                campo_precios(8).factor = Val("" & mytablex.Fields("factor9"))
                campo_precios(8).precio = xpreciox
                xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor9")))
                campo_precios(8).stock = "" & xbuf
                campo_precios(8).costo = "" & xcosto
                xmargen = 0

                If xcosto > 0 Then
                    xmargen = ((xpreciox - xcosto) * 100) / xcosto

                End If

                campo_precios(8).margen = "" & xmargen

            End If

        End If

        xcosto = 0

        If Val(empaque_visible) > 9 Then
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
                campo_precios(9).factor = Val("" & mytablex.Fields("factor10"))
                campo_precios(9).precio = xpreciox
                xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor10")))
                campo_precios(9).stock = "" & xbuf
                campo_precios(9).costo = "" & xcosto
                xmargen = 0

                If xcosto > 0 Then
                    xmargen = ((xpreciox - xcosto) * 100) / xcosto

                End If

                campo_precios(9).margen = "" & xmargen

            End If

        End If

        'MsgBox "xx"
        sql_saldo_locales uproducto
        'margenes
        sw = 1

    End If

    DBGrid4.refresh
    'MsgBox ""
    'mytablex.Close
    'mytablez.Close

    '----ahora deb cargar tambien la foto del producto...
    DBGrid2.Enabled = False
    Frame1.Enabled = False
    Frame5.Enabled = True
    Frame5.Visible = True
    DBGrid4.Enabled = True
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
        nlinea = "" & mytablex.Fields("descripcio")
        nt1 = "" & mytablex.Fields("t1")
        nt2 = "" & mytablex.Fields("t2")
        nt3 = "" & mytablex.Fields("t3")
        nt4 = "" & mytablex.Fields("t4")
        nt5 = "" & mytablex.Fields("t5")
        nt6 = "" & mytablex.Fields("t6")
        nt7 = "" & mytablex.Fields("t7")
        nt8 = "" & mytablex.Fields("t8")
        nt9 = "" & mytablex.Fields("t9")
        nt10 = "" & mytablex.Fields("t10")
        nt11 = "" & mytablex.Fields("t11")
        nt12 = "" & mytablex.Fields("t12")
        nt13 = "" & mytablex.Fields("t13")
        nt14 = "" & mytablex.Fields("t14")
        nt15 = "" & mytablex.Fields("t15")
        nt16 = "" & mytablex.Fields("t16")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Sub ingreso_tallas(buf As String)

    Dim found As Integer

    MsgBox linea
    found = busca_linea(buf)

    If found = 0 Then Exit Sub
    pone_tallas
    DBGrid2.Enabled = False
    Frame3.Enabled = False
    Frame3.Enabled = True
    Frame3.Visible = True
    t1.SetFocus

End Sub

Sub pone_tallas()
    t1 = "" & DBGrid2.columns("t1")
    t2 = "" & DBGrid2.columns(19)
    t3 = "" & DBGrid2.columns(20)
    t4 = "" & DBGrid2.columns(21)
    t5 = "" & DBGrid2.columns(22)
    t6 = "" & DBGrid2.columns(23)
    t7 = "" & DBGrid2.columns(24)
    t8 = "" & DBGrid2.columns(25)
    t9 = "" & DBGrid2.columns(26)
    t10 = "" & DBGrid2.columns(27)
    t11 = "" & DBGrid2.columns(28)
    t12 = "" & DBGrid2.columns(29)
    t13 = "" & DBGrid2.columns(30)
    t14 = "" & DBGrid2.columns(31)
    t15 = "" & DBGrid2.columns(32)
    t16 = "" & DBGrid2.columns(33)

End Sub

Sub ingreso_locales()
    xxpone_locales
    Frame4.Visible = True
    observa1.SetFocus

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

    observa1 = "" & DBGrid2.columns("observa1")
    observa2 = "" & DBGrid2.columns("observa2")
    observa3 = "" & DBGrid2.columns("observa3")
    observa4 = "" & DBGrid2.columns("observa4")

End Sub

Sub cerrar_data1()

    'On Error GoTo cmd17_err
    'Data1.Recordset.Close
    'Exit Sub
    'cmd17_err:
    'Exit Sub
End Sub

Sub graba_temporald(mytablex As ADODB.Recordset, _
                    sw As Integer, _
                    canti As String, _
                    mytabley As ADODB.Recordset)

    Dim fechadi  As String

    Dim deslipox As Double

    Dim found    As Integer

    Dim xxca     As String

    Dim sdx      As Double

    Dim dsdx     As Double

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
            If Val("" & mytablez.Fields("descuento")) > 0 Then
                deslipox = Val("" & mytablez.Fields("descuento"))

            End If

        End If
      
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
    'If "" & dbgrid2.columns("unidad") = "" & mytabley.Fields("unidad1") Then  'si es la misma unidad
    '
    '   If Val("" & dbgrid2.columns("cantidad")) >= a And Val("" & dbgrid2.columns("cantidad")) <= a Then
    '   End If
    'End If
    'If "" & mytablex.Fields("excludscto") = "S" Then
    '   Data1.Recordset.Fields("deslipo") = 0
    'End If
    '------------------------------------------------
    'MsgBox xpreciox
    '--------------------------------------
    'dbgrid2.Col = 0
    'dbgrid2.Row = dbgrid2.VisibleRows - 1
    'dbgrid2.SetFocus
    '--------------------------------------
    'Data2.Recordset.MoveLast
    DBGrid2.refresh
    'found = sumar_detalle()
    'dbgrid2.Col = 0
    'dbgrid2.Row = dbgrid2.Row - 1
    DBGrid2.columns("percepcion") = 0
    DBGrid2.columns("tpercepcio") = 0

    DBGrid2.columns("nroprecio") = "1"
    DBGrid2.columns("hora") = Format(Now, "hh:mm:ss")
    DBGrid2.columns("categoria") = "" & mytablex.Fields("categoria")
    DBGrid2.columns("producto") = "" & mytablex.Fields("producto")
    DBGrid2.columns("proveedorp") = "" '& mytablex.Fields("proveedor1")
    DBGrid2.columns("tipo") = ""
    DBGrid2.columns("serie") = ""
    DBGrid2.columns("numero") = ""
    DBGrid2.columns("isc") = ""  '& mytablex.Fields("vendedor")
    DBGrid2.columns("comision") = Val("" & mytablex.Fields("comision"))
    DBGrid2.columns("descripcio") = Mid$("" & mytablex.Fields("descripcio"), 1, 60)
    'MsgBox xxca
    DBGrid2.columns("cantidad") = Val(Format(Val(xxca), "0.000"))
    'dbgrid2.columns("descuento") = Val("" & mytablex.Fields("isc"))

    DBGrid2.columns("unidad") = "" & mytabley.Fields("unidad1")  'ojo se cambio por placa
    DBGrid2.columns("factor") = Val("" & mytabley.Fields("factor1"))
    DBGrid2.columns("precio") = xpreciox
    DBGrid2.columns("precio") = xpreciox

    If "" & mytable11.Fields("hdetraccio") <> "S" Then
        DBGrid2.columns("tdetra") = 0

    End If

    DBGrid2.columns("precio") = xpreciox

    If sw = 1 Then
        DBGrid2.columns("total") = Val(Format(Val(xxca), "0.000"))
        DBGrid2.columns("cantidad") = Val(xxca) / xpreciox

    End If

    'dbgrid2.columns("neto") = Val("" & mytablex.Fields("tax"))
    'dbgrid2.columns("unidad") = "" & mytabley.Fields("unidad1")
    'MsgBox Trim("" & mytabley.Fields("percepcion"))
    sdx = 0
    DBGrid2.columns("l1") = Trim("" & mytablex.Fields("percepcion"))

    If "" & DBGrid2.columns("l1") = "S" Then  'tiene percepcion
        sdx = tabla_percepcion

    End If

    DBGrid2.columns("percepcion") = sdx
    DBGrid2.columns("tpercepcio") = xpreciox * Val("" & DBGrid2.columns("percepcion")) / 100

    DBGrid2.columns("factor") = Val("" & mytabley.Fields("factor1"))
    DBGrid2.columns("precio") = xpreciox
    DBGrid2.columns("total") = xpreciox
    DBGrid2.columns("subtotal") = xpreciox

    '-----------------
    'If "" & mytabley.Fields("fuel") = "S" And Val(xxca) > 1 Then
    'dbgrid2.columns("total") = Val(xxca)
    'dbgrid2.columns("cantidad") = Val(xxca) / xpreciox
    'End If
    '-----------------

    'dbgrid2.columns("deslipo") = 0
    DBGrid2.columns("deslipo") = deslipox
    DBGrid2.columns("tax") = 0
    DBGrid2.columns("vendedor") = "" 'Val("" & mytablex.Fields("isc"))
    DBGrid2.columns("impuesto") = 0
    DBGrid2.columns("igv") = Val("" & mytablex.Fields("igv"))
    DBGrid2.columns("linea") = "" & mytablex.Fields("linea")

    DBGrid2.columns("descuento") = 0
    DBGrid2.columns("neto") = 0

    mytables.Open "SELECT * FROM DUENO where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and producto='" & "" & mytablex.Fields("producto") & "' ", cn, adOpenKeyset, adLockOptimistic

    If mytables.RecordCount > 0 Then  'si existe
        DBGrid2.columns("ccosto") = Trim("" & mytables.Fields("codigo"))  'ojo si no es por local

    End If

    mytables.Close
    DBGrid2.columns("ccosto") = "" & mytablex.Fields("seccion")

    '---------pone a quien pertenece --------------------
    'dbgrid2.columns("l1") = "" & mytablex.Fields("c11")
    'dbgrid2.columns("l2") = "" & mytablex.Fields("c12")
    'dbgrid2.columns("l3") = "" & mytablex.Fields("c13")
    DBGrid2.columns("tcosto") = Val("" & mytablex.Fields("costou"))
    '-----------------------------
    'le pone las familias+subfamil+seccion+marca
    DBGrid2.columns("familia") = "" & mytablex.Fields("Familia")
    DBGrid2.columns("subfamilia") = "" & mytablex.Fields("subFamilia")
    DBGrid2.columns("marca") = "" & mytablex.Fields("marca")
    DBGrid2.columns("total") = Val(DBGrid2.columns("cantidad")) * Val(DBGrid2.columns("precio"))
    DBGrid2.columns("ivap") = Val("" & mytablex.Fields("ivap"))
    DBGrid2.columns("isc") = Val("" & mytablex.Fields("DSCTOREF"))  'ojo el descuento referencial se pone aqui
    DBGrid2.columns("isc") = Val("" & mytablex.Fields("isc"))
    DBGrid2.columns("l1") = Trim("" & mytablex.Fields("percepcion"))
    DBGrid2.columns("serviciopo") = Val("" & mytablex.Fields("serviciomesa"))

    calcula_igv 0
    found = leer_visorcaja("" & DBGrid2.columns("descripcio"), dicmoneda & DBGrid2.columns("Total"))

End Sub

Function sumar_detalle()

    On Error GoTo cmd35_err

    Dim found      As Integer

    Dim sdx        As Double

    Dim fila       As Integer

    Dim xtotal     As Double

    Dim xdescuento As Double

    Dim xneto      As Double

    Dim ximpuesto  As Double

    Dim xsubtotal  As Double

    Dim xgravado   As Double

    Dim xc1        As Double

    Dim xc2        As Double

    Dim xc3        As Double

    Dim xc4        As Double

    Dim xc5        As Double

    Dim xc6        As Double

    Dim xc7        As Double

    Dim xc8        As Double

    Dim xc9        As Double

    Dim difre      As Double

    Dim sw         As Integer

    Dim xredo      As Double

    Dim sdx1       As Double

    'Dim xacuenta As Double
    Dim vr

    Dim stx            As Double

    Dim xntcant        As Double

    Dim xfilax         As Integer

    Dim xivap          As Double

    Dim xisc           As Double

    Dim xdetra         As Double

    Dim xpeaje         As Double

    Dim xpercepcion    As Double

    Dim xserviciocobro As Double

    Dim xtxpercepcion  As Double

    Dim tnrofilas      As Double

    xpeaje = 0
    xserviciocobro = 0
    xdetra = 0
    xntcant = 0
    xredo = 0
    tnrofilas = 0
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
    xpercepcion = 0
    xredo = 0
    xgravado = 0
    xtotal = 0
    xdescuento = 0
    xneto = 0
    ximpuesto = 0
    xsubtotal = 0
    xtxpercepcion = 0
    '------------------------
    'dbrecords = Data2.Recordset.RecordCount
    'For fila = 0 To dbgrid2.visiblerows - 1
    sw = 1
    exisdev = 0
    found = ir_primero1()

    If found = 0 Then
        GoTo avex

    End If

    xpercepcion = selecciona_percepcion("" & codigo, "" & clasesunat)
    'Data2.Refresh
    'Data2.Enabled = False
    Do

        If Data2.Recordset.EOF Then Exit Do

        If Val("" & Data2.Recordset.Fields("cantidad")) < 0 Then
            exisdev = -10

        End If

        If Len(Trim("" & Data2.Recordset.Fields("vendedor"))) = 0 Then
            If Trim("" & mytable11.Fields("vdetalle")) = "S" Then
                sw = 2

            End If

        End If

        Data2.Recordset.Edit
        resuma_precios xpercepcion
        Data2.Recordset.Update

        If Val("" & Data2.Recordset.Fields("igv")) = 0 Then
            xgravado = xgravado + Val("" & Data2.Recordset.Fields("total"))

        End If

        tnrofilas = tnrofilas + 1
        xpeaje = xpeaje + Val("" & Data2.Recordset.Fields("xneto"))
        xserviciocobro = xserviciocobro + Val("" & Data2.Recordset.Fields("servicioco"))
        xdetra = xdetra + Val("" & Data2.Recordset.Fields("tdetra"))
        xisc = xisc + Val("" & Data2.Recordset.Fields("tisc"))
        xivap = xivap + Val("" & Data2.Recordset.Fields("tivap"))
        xntcant = xntcant + Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & Data2.Recordset.Fields("factor")) 'suma bruto
        xtotal = xtotal + Val("" & Data2.Recordset.Fields("total"))
        xdescuento = xdescuento + Val("" & Data2.Recordset.Fields("descuento"))
        xneto = xneto + Val("" & Data2.Recordset.Fields("neto"))
        ximpuesto = ximpuesto + Val("" & Data2.Recordset.Fields("impuesto"))
        xsubtotal = xsubtotal + Val("" & Data2.Recordset.Fields("subtotal"))
        xtxpercepcion = xtxpercepcion + Val("" & Data2.Recordset.Fields("tpercepcio"))
        Data2.Recordset.MoveNext
    Loop
avex:
    'Data2.Enabled = True
    'MsgBox "ABC"

    txpercepcion = Format(xtxpercepcion, nrodecimal)
    serviciocobro = redondeo1("" & xserviciocobro, nrodecimal)
    tpeaje = Format(xpeaje, nrodecimal)
    tdetra = Format(xdetra, nrodecimal)
    gravado = Format(xgravado, nrodecimal)
    ntcant = Format(xntcant, nrodecimal)
    nrofilas = Format(tnrofilas, "0")
    'MsgBox xtotal
    'xtotal = xtotal + xtxpercepcion
    txtotal = Format(xtotal, nrodecimal)
    txtotlare = 0

    If "" & mytable11.Fields("redondeo") = "S" Then
        'MsgBox "abc"
        'MsgBox redondeo2(txtotal)
        txtotlare = Val(redondeo2(txtotal, "" & nrodecimal)) - Val(txtotal)
        txtotal = Val(redondeo2(txtotal, "" & nrodecimal))

        'MsgBox txtotal
    End If

    tisc = Val(Format(xisc, nrodecimal))
    tivap = Val(Format(xivap, nrodecimal))
    'MsgBox acuenta
    stx = Val(txtotal) - Val(acuenta)
    rtxtotal = Format(stx, nrodecimal)
    'MsgBox rtxtotal
    'txtotal = Format(xtotal, nrodecimal)
    txdescuento = Format(xdescuento, nrodecimal)
    txneto = Format(xneto, nrodecimal)
    'tximpuesto = Format(ximpuesto, "0.000")
    'MsgBox ximpuesto
    tximpuesto = "" & Redondear1(ximpuesto, 2)  'redondeo3("" & ximpuesto, nrodecimal)
    'MsgBox tximpuesto
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

Function sumar_detalled()

    'antiguo
    On Error GoTo cmd35_err

    Dim sdxp       As Double

    Dim found      As Integer

    Dim sdx        As Double

    Dim fila       As Integer

    Dim xtotal     As Double

    Dim xdescuento As Double

    Dim xneto      As Double

    Dim ximpuesto  As Double

    Dim xsubtotal  As Double

    Dim xgravado   As Double

    Dim xc1        As Double

    Dim xc2        As Double

    Dim xc3        As Double

    Dim xc4        As Double

    Dim xc5        As Double

    Dim xc6        As Double

    Dim xc7        As Double

    Dim xc8        As Double

    Dim xc9        As Double

    Dim difre      As Double

    Dim sw         As Integer

    Dim xredo      As Double

    Dim sdx1       As Double

    'Dim xacuenta As Double
    Dim tnrofilas  As Double

    Dim vr

    Dim stx         As Double

    Dim xntcant     As Double

    Dim xfilax      As Integer

    Dim xivap       As Double

    Dim xisc        As Double

    Dim xdetra      As Double

    Dim xpeaje      As Double

    Dim xpercepcion As Double

    Dim nro_percep  As Double

    xpercepcion = 0
    xpeaje = 0
    nro_percep = 0
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
    tnrofilas = 0

    xredo = 0
    xgravado = 0
    xtotal = 0
    xdescuento = 0
    xneto = 0
    ximpuesto = 0
    xsubtotal = 0
    flag_percepcion = ""
    '------------------------
    'dbrecords = Data2.Recordset.RecordCount
    'For fila = 0 To DBGrid2.ApproxCount - 1
    sw = 1
    exisdev = 0
    found = ir_primero1()

    If found = 0 Then
        GoTo avex

    End If

    Do

        If Data2.Recordset.EOF Then Exit Do
        Data2.Recordset.Edit
        resuma_precios 0
        Data2.Recordset.Update

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

        If Val("" & Data2.Recordset.Fields("igv")) = 0 Then
            xgravado = xgravado + Val("" & Data2.Recordset.Fields("total"))

        End If

        tnrofilas = tnrofilas + 1
        'xc1 = xc1 + Val("" & Data2.Recordset.Fields("t16")) * Val("" & Data2.Recordset.Fields("total")) / 100
        xc1 = xc1 + Val("" & Data2.Recordset.Fields("t15"))  't15 es el descuento calculado
        xpeaje = xpeaje + Val("" & Data2.Recordset.Fields("xneto"))
        xdetra = xdetra + Val("" & Data2.Recordset.Fields("tdetra"))
        xisc = xisc + Val("" & Data2.Recordset.Fields("tisc"))

        'MsgBox "abc"
        'If "" & Data2.Recordset.Fields("l1") = "S" Then 'tiene percepcion
        '   xpercepcion = xpercepcion + Val("" & Data2.Recordset.Fields("total"))
        'End If
        'MsgBox "cba"

        If "" & Data2.Recordset.Fields("l1") = "S" Then
            nro_percep = nro_percep + 1

        End If

        xivap = xivap + Val("" & Data2.Recordset.Fields("tivap"))
        xntcant = xntcant + Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & Data2.Recordset.Fields("factor")) 'suma bruto
        xtotal = xtotal + Val("" & Data2.Recordset.Fields("total"))
        xdescuento = xdescuento + Val("" & Data2.Recordset.Fields("descuento"))
        xneto = xneto + Val("" & Data2.Recordset.Fields("neto"))
        ximpuesto = ximpuesto + Val("" & Data2.Recordset.Fields("impuesto"))
        xsubtotal = xsubtotal + Val("" & Data2.Recordset.Fields("subtotal"))
        xpercepcion = xpercepcion + Val("" & Data2.Recordset.Fields("tpercepcio"))

        If Len(Trim("" & Data2.Recordset.Fields("vendedor"))) = 0 Then
            If Trim("" & mytable11.Fields("vdetalle")) = "S" Then
                sw = 2

                'MsgBox "Ingrese vendedor ", 48, "Aviso"
                'dbgrid2.Col = 8
                'Exit Function
            End If

        End If

        Data2.Recordset.MoveNext
    Loop
avex:

    'SABER LA PERCEPCION
    'MsgBox "abc"
    If nro_percep > 0 Then
        flag_percepcion = "S"

    End If

    sdxp = Val(xpercepcion)
    'MsgBox "cba"
    ytotal = Format(xtotal, nrodecimal)
    txpercepcion = Format(sdxp, nrodecimal)
    xtotal = xtotal + Val(txpercepcion)

    tpeaje = Format(xpeaje, nrodecimal)
    tdetra = Format(xdetra, nrodecimal)
    gravado = Format(xgravado, nrodecimal)
    ntcant = Format(xntcant, nrodecimal)
    nrofilas = Format(tnrofilas, "0")
    txtotal = Format(xtotal, nrodecimal)
    txtotlare = 0

    If "" & mytable11.Fields("redondeo") = "S" Then
        txtotlare = Val(redondeo1(txtotal, nrodecimal)) - Val(txtotal)
        txtotal = redondeo1(txtotal, nrodecimal)

    End If

    tisc = Val(Format(xisc, nrodecimal))
    tivap = Val(Format(xivap, nrodecimal))

    stx = Val(txtotal)

    If Val(acuenta) > 0 Then
        If Len(petipo) > 0 Then
            stx = Val(txtotal) - Val(acuenta)
        Else
            stx = Val(acuenta)

        End If

    End If

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
    sumar_detalled = sw
    Exit Function
cmd35_err:
    MsgBox "Error en sumar_detalleD " & error$, 24, "Aviso"
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
    Exit Sub

End Sub

Sub inicialIzatodo()

    Dim found As Integer

    Dim sdx   As Double

    acurabuffer = ""
    correo = ""
    found = leer_visorcaja("SISTEMA CALIPSO", "CASH REGISTER")
    clasesunat.ListIndex = 0
    serviciocobro = 0
    tabla_percepcion = 0
    flag_percepcion = ""
    totpedido = ""
    ytotal = ""
    txpercepcion = ""
    flag_especial = ""
    DBGrid2.Enabled = True
    Command13.Enabled = True
    ndetraccion = ""
    flage = ""
    tproducto = ""

    sentido.Enabled = False

    If "" & mytable11.Fields("sentido") = "C" Then
        sentido = ""
        sentido.Enabled = True
    Else
        sentido = "" & mytable11.Fields("sentido")

    End If

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
    CAMPO1 = ""
    CAMPO2 = ""
    campo3 = ""
    sql_detalle
    found = sumar_detalle()
    uvueltos = dicmoneda & ":" & Format(Val("" & mytable11.Fields("uvueltos")), nrodecimal)
    uvueltod = "US$:" & Format(Val("" & mytable11.Fields("uvueltod")), nrodecimal)

    'uvueltos = "" & mytable11.Fields("uvueltos")
    'uvueltod = "" & mytable11.Fields("uvueltod")
    DBGrid2.Enabled = True
    DBGrid2.SetFocus

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

Sub proceso_impresion11(bxtipo As String, _
                        bxserie As String, _
                        bxnumero As String, _
                        sw As Integer, _
                        ascopia As String)

    Dim found    As Integer

    Dim archivot As String

    On Error GoTo cmd6_err:

    'MsgBox ""
    cerrar_archivo

    If sw = 0 Then   'si es posible
        found = abre_puerto("" & mytable11.Fields("capuerto"), Val("" & mytable11.Fields("catipo")), "" & mytable11.Fields("gavetacola"))

    End If

    'verificamos si es puerto LPT para no hacer formato impresion
    found = control_impresion(bxtipo, 10)

    If found = 10 And sw <> 2 Then
        Exit Sub

    End If

    'MsgBox "proceso impresion"
    factura_formatox Trim("" & "" & mytable11.Fields("local")), "" & bxtipo, "" & bxserie, "" & bxnumero, ascopia, sw
    cerrar_archivo
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = control_impresion(bxtipo, sw)
    'genver.File = globaldir & "\temporal\" & gusuario & ".txt"
    'genver.Show 1
    Exit Sub
cmd6_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$, 48, "Aviso"
    Exit Sub

End Sub

Function control_impresion(bxtipo As String, psw As Integer)

    Dim copias     As Integer

    Dim found      As Integer

    Dim sFile      As String

    Dim mytablex   As New ADODB.Recordset

    Dim sw         As String

    Dim xcolax     As String

    Dim xxpuerto   As String

    Dim oldprinter As String

    Dim I          As Integer

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

            Case "I"  'pedidos
       
                xxpuerto = "" & mytable11.Fields("puertope")
                sw = "" & mytable11.Fields("ipe")
                xcolax = "" & mytable11.Fields("cpro")
       
            Case "T"
                xxpuerto = "" & mytable11.Fields("puertoot")
                sw = "" & mytable11.Fields("iot")
                xcolax = "" & mytable11.Fields("cpro")

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
   
        If "" & mytable11.Fields("odcola") = "S" Then

            'copias = nro_copias(bxtipo)
            For I = 1 To 1
                oldprinter = Printer.DeviceName
                selecciona_impresoras ("" & mytable11.Fields("odpuerto"))
                sFile = globaldir & "\temporal\" & gusuario & ".txt"
                found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("bold"), "" & mytable11.Fields("letrainterna"))
                selecciona_impresoras (oldprinter)
            Next I

        End If

        If "" & mytable11.Fields("odcola") <> "S" Then
            'MsgBox "" & mytable11.Fields("odpuerto")
      
            found = star_sp342("" & mytable11.Fields("odpuerto"), 0)
      
            found = corte_papel("" & mytable11.Fields("odpuerto"), Val("" & mytable11.Fields("catipo")))

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
        copias = nro_copias(bxtipo)

        If copias < 1 Then
            copias = 1

        End If

        For I = 1 To copias
            oldprinter = Printer.DeviceName
            selecciona_impresoras (xxpuerto)
            sFile = globaldir & "\temporal\" & gusuario & ".txt"
            found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("bold"), "" & mytable11.Fields("letrainterna"))
            selecciona_impresoras (oldprinter)
        Next I

    End If

    If xcolax <> "S" Then
        found = star_sp342(xxpuerto, 0)
        found = corte_papel(xxpuerto, Val("" & mytable11.Fields("catipo")))

    End If

    control_impresion = found
    Exit Function
cmd67111_err:
    MsgBox "Aviso en control impresion " + error$, 48, "Aviso"
    Exit Function

End Function

Sub proceso_impresioncopia()

    Dim found    As Integer

    Dim archivot As String

    On Error GoTo cmd7_err:

    cerrar_archivo
    factura_formatox Trim("" & "" & mytable11.Fields("local")), "" & dbGrid1.columns(8), "" & dbGrid1.columns(9), "" & dbGrid1.columns(0), "1", 0
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    'found = valida_wordpad(FileName)
    Exit Sub
cmd7_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$
    Exit Sub

End Sub

Sub proceso_impresioncopia1()

    Dim found    As Integer

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

Sub factura_formatox(bxlocal As String, _
                     bxtipo As String, _
                     bxserie As String, _
                     bxnumero As String, _
                     ascopia As String, _
                     psw As Integer)

    Dim vacu            As String

    Dim mytablex        As New ADODB.Recordset

    Dim mytabley        As New ADODB.Recordset

    Dim mytablez        As New ADODB.Recordset

    Dim found           As Integer

    Dim nro_lineas      As Integer

    Dim contando        As Integer

    Dim faltante        As Integer

    Dim I               As Integer

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
        'MsgBox bxtipo
        archivo_formato = busca_archivo_formato(bxtipo)

        If Len(archivo_formato) = 0 Then
            MsgBox "No existe archivo formato ", 48, "Aviso"
            'MsgBox ""
            Exit Sub

        End If

    End If

    'cabeza
    'proceso_formatos(archivo_formato , mydbx , mytablex , ubicacioni , ubicacionf , basedatos , indice , tipo , numero , ascopia , contando )
    mytablex.Open "SELECT * FROM " & gocabeza & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then 'si existe
        mytablex.Close
        Exit Sub

    End If

    'MsgBox ""
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytablex, "{", "}", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytablex, "{", "}", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    vacu = "" & mytablex.Fields("acu")
    'MsgBox ""
    '
    'detalle
    flag_contando = 0
    mytabley.Open "SELECT * FROM " & godetalle & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytabley.RecordCount > 0 Then 'si existe
        Do

            If mytabley.EOF Then Exit Do
            If "" & mytabley.Fields("dua") <> "R" Then
                flag_contando = contando + 1
                'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                'found = proceso_formatos(archivo_formato, mytabley, "/", "\", godetalle, "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
                found = proceso_formatos(archivo_formato, mytabley, "/", "\", godetalle, "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
                'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                contando = contando + 1

            End If
          
            mytabley.MoveNext
        Loop

    End If

    'mytabley.Close
    '
    If bxtipo <> "1" And bxtipo <> "2" And bxtipo <> "5" And bxtipo <> "7" Then
        If nro_lineas > 0 Then 'si numero lineas <>0 dice que debe respetar el siguiente
            If contando < nro_lineas Then

                For I = contando To nro_lineas
                    Open FileName For Append As #1
                    found = formateaa("", 1, 2, 0)
                    Close #1
                Next I

            End If

        End If

    End If

    '----- SUBTOTAL
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytablex, "$", "?", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytablex, "$", "?", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
          
    mytablez.Open "SELECT * FROM " & gofpago & "   where  local='" & bxlocal & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablez.RecordCount > 0 Then 'si existe
        Do

            If mytablez.EOF Then Exit Do
            'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
            'found = proceso_formatos(archivo_formato, mytablez, "<", ">", gofpago, "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
            found = proceso_formatos(archivo_formato, mytablez, "<", ">", gofpago, "fpagov", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
            'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
               
            mytablez.MoveNext
        Loop

    End If

    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
    'found = proceso_formatos(archivo_formato, mytablex, "^", "&", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    found = proceso_formatos(archivo_formato, mytablex, "^", "&", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
         
    mytablex.Close
    mytabley.Close
    mytablez.Close
    Exit Sub
cmd450009_err:
    MsgBox "Aviso Proceso en Compilando " & error$, 24, "Aviso"
    'mytablex.Close
    Exit Sub

End Sub

Function busca_archivo_formato(bxtipo As String) As String

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    'MsgBox bxtipo
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

            Case "H"
                busca_archivo_formato = "" & mytable11.Fields("archivope")

            Case "T"
                busca_archivo_formato = "" & mytable11.Fields("archivoot")

            Case "I"
                busca_archivo_formato = "" & mytable11.Fields("archivope")

                'MsgBox ""
        End Select

        'MsgBox ""
    End If

    mytablex.Close
 
End Function

Function busca_parame1(buf As String, sw As Integer) As String

    Dim sdx      As Double

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

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    borrar_campos

    mytablex.Open "SELECT * FROM " & dgusuariog & "   where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then 'si existe
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
   
        Data2.Recordset.AddNew

        For I = 0 To mytablex.Fields.count - 1
            Data2.Recordset.Fields(I) = mytablex.Fields(I)
        Next I

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

    'dotipo = ""
    '   doserie = ""
    '   donumero = ""
    '   dototal = ""
    '   dofpago = ""
    '   dofecha = ""
End Sub

Function busca_paridad()

    Dim found    As Integer

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

    'On Error GoTo cmd13_err
    'Data1.Recordset.MoveLast
    'Exit Sub
    'cmd13_err:
    'Exit Sub
End Sub

Sub PROCESO_BORRAR_DOCUMENTO(buf0 As String, _
                             buf As String, _
                             buf1 As String, _
                             buf2 As String)

    Dim mytablex As New ADODB.Recordset

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
    mytablex.Open "SELECT * FROM fpago where  fpago='" & buf & "' ORDER by fpago", cn, adOpenDynamic, adLockOptimistic

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
    'buf = "select * from dpedidov where local='" & "" & "" & mytable11.Fields("local") & "' and tipo='" & dotipo & "' and serie='" & doserie & "' and numero='" & donumero & "'"
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
            'mytablex.Fields("dotipo") = buf1
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

    Dim buf      As String

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    If Len(pedido) > 0 Then
        Exit Function

    End If

    buf = busca_tipo_acu(bxtipo)
ahj1:

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM " & gocabeza & " where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

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

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

ahj1:

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM " & gocabeza & " where  local='" & Trim("" & "" & mytable11.Fields("local")) & " and tipo='" & xptipo & "' and serie='" & xpserie & "' and numero='" & xpnumero & "'", cn, adOpenDynamic, adLockOptimistic

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

    'mytablex.Open "SELECT * FROM fpago where bco='S' or bco=null ", cn, adOpenDynamic, adLockOptimistic
    If xfpagox.State = 1 Then
        xfpagox.Close
        Set xfpagox = Nothing

    End If

    xfpagox.Open "SELECT * FROM fpago order by fpago  ", cn, adOpenDynamic, adLockOptimistic
    Set dbgrid10.DataSource = xfpagox
               
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
    tcampo5.MaxLength = 3
    tcampo6.MaxLength = 2
    tcampo1 = Trim("" & codigo)
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
        descripcio5 = "NroCtas"
        tcampo3.MaxLength = 4

    End If

    If sw = 3 Or sw = 6 Then 'credito
        descripcio1 = "Codigo"
        descripcio2 = "Nombre"
        descripcio3 = "NroAprob"
        descripcio4 = "Observacion"
        descripcio5 = "NroDias"

    End If

    If sw = 5 Then  'tarjeta Debito
        descripcio1 = "Codigo"
        descripcio2 = "Nombre"
        descripcio3 = "NroTarjeta"
        descripcio4 = "Entidad"
        descripcio5 = ""
        tcampo3.MaxLength = 4

    End If

    If sw = 1 Then  'SI ES PAGO ADELANTADO
        descripcio1 = "Codigo"
        descripcio2 = "Nombre"
        descripcio3 = ""
        descripcio4 = ""
        descripcio5 = ""
        descripcio6 = ""

    End If

    If sw = 6 Then  'SI ES VALES
        descripcio1 = "Codigo"
        descripcio2 = "Nombre"
        descripcio3 = ""
        descripcio4 = ""
        descripcio5 = ""
        descripcio6 = ""

    End If
   
    If sw = 10 Then
        descripcio3 = "Banco"
        descripcio4 = "NroCheque"

    End If

    If sw = 2 Then
        descripcio3 = "Nro.OperaBanco"

    End If
   
End Function

Sub suma_fpagov()

    Dim sdxs  As Double

    Dim sdxd  As Double

    Dim sdx   As Double

    Dim sdx1  As Double

    Dim found As Integer

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
                codigo = Trim(tcampo1)
                nombre = tcampo2

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

    found = leer_visorcaja(dicmoneda & stxtotals, "US$  " & stxtotald)

    Exit Sub
cmd7812_err:
    MsgBox "Error en " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub valida_ingresado()

    Dim sdx      As Double

    Dim xsoles   As Double

    Dim xdolares As Double

    Dim xfaltas  As Double

    Dim xfaltad  As Double

    Dim xvueltos As Double

    Dim xvueltod As Double

    Dim sdx1     As Double

    Dim sdx2     As Double

    Dim sdx3     As Double

    xsoles = 0
    xdolares = 0
    xfaltas = 0
    xfaltad = 0
    xvueltos = 0
    xvueltod = 0
    sdx3 = 0

    If "" & dbgrid9.columns(1) = "S" Then
        xsoles = Val("" & dbgrid9.columns(2))
        xdolares = Val(Val("" & dbgrid9.columns(2))) / Val(paridadfp)
        sdx3 = xdolares

    End If

    If "" & dbgrid9.columns(1) = "D" Then
        xdolares = Val("" & dbgrid9.columns(2))
        xsoles = Val("" & dbgrid9.columns(2)) * Val(paridadfp)
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

Private Sub Timer2_Timer()

    Dim D           As Integer

    Dim strLectura  As String ''Contiene la cadena leida hasta el momento por el puerto serie

    Dim intContador As Integer

    On Error GoTo ManejoError:
   
    Select Case "" & mytable11.Fields("portbala")

        Case "COM1"
            D = 1

        Case "COM2"
            D = 2

        Case "COM3"
            D = 3

        Case "COM4"
            D = 4

        Case "COM5"
            D = 5

        Case Else
            Exit Sub

    End Select
   
    MSComm1.CommPort = D ''COM1
    ' 9600 baudios, sin paridad, 8 bits de datos y 1 bit de parada.
    MSComm1.Settings = "9600,n,8,1" '(ver configuracion en la ayuda del Visual Studio)
    MSComm1.InputLen = 0
    MSComm1.PortOpen = True
    ' Bucle donde se leen los datos enviados constantemente por el puerto serial
    'hasta que se encuentra un retorno de carro (enter) que marca el fin de esa cadena
    intContador = 1
    Do
        DoEvents
        strLectura = strLectura & MSComm1.input

        If Right(strLectura, 1) = Chr(13) And Len(strLectura) >= 17 Then
            Exit Do

        End If

    Loop
    MSComm1.PortOpen = False
   
    If "" & mytable11.Fields("tipo_balanza") = "1" Then
        If IsNumeric(Mid(strLectura, Len(strLectura) - 12, 6)) = True And Val(Mid(strLectura, Len(strLectura) - 12, 6)) < 80000 Then
            acurabuffer.Caption = FormatNumber(Val(Mid(strLectura, Len(strLectura) - 12, 6)), 2, vbFalse, vbFalse, vbFalse)

        End If

    End If

    'tipo2-----------------
    'MsgBox Len(strLectura)
    'MsgBox Mid(strLectura, Len(strLectura) - 9, 6)
    If "" & mytable11.Fields("tipo_balanza") = "2" Then
        If IsNumeric(Mid(strLectura, Len(strLectura) - 9, 6)) = True And Val(Mid(strLectura, Len(strLectura) - 9, 6)) < 80000 Then
            acurabuffer.Caption = FormatNumber(Val(Mid(strLectura, Len(strLectura) - 9, 6)), 2, vbFalse, vbFalse, vbFalse)

        End If

    End If
   
    ''80000 es la cantidad maxima de kilogramos registrada por la balanza
    'If IsNumeric(Mid(strLectura, Len(strLectura) - 12, 6)) = True And Val(Mid(strLectura, Len(strLectura) - 12, 6)) < 80000 Then
    '    acurabuffer.Caption = FormatNumber(Val(Mid(strLectura, Len(strLectura) - 12, 6)), 2, vbFalse, vbFalse, vbFalse)
    'End If
    Exit Sub
ManejoError:

    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False

    End If

End Sub

Private Sub tiposervicio1_Click()
    dki3432_Click

End Sub

Private Sub xcongelax_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        cmdCancelar_Click
        Exit Sub

    End If

    cmdGrabar_Click

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
            consulta_xruc "" & xnombre

        End If

        If local1.Visible = True Then
            consulta_xruc2 "" & xnombre

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

        If Trim("" & mytable11.Fields("bodega")) = Trim(xruc) Then
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

        found = busca_localx("" & xruc)

        If found = 0 Then
            MsgBox "digite un Almacen a Pedir", 48, "Aviso"
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

        'If Len(xruc) <> 11 Then
        '   xruc.SetFocus
        '   Exit Sub
        'End If
        found = valida_ruc("" & xruc)

        If found = 0 Then
            MsgBox dicruc & " No Valido", 48, "Aviso"
            xruc = ""
            xruc.SetFocus
            Exit Sub

        End If

        If Trim("" & mytable11.Fields("bodega")) = Trim(xruc) Then
            MsgBox "Almacen debe ser diferente ", 48, "Aviso"
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
        If local1 = "PEDIDO" Then  'pedido a almacen
            consulta_almacen
            Exit Sub

        End If

        If local1.Visible <> True Then
            consulta_xruc "" & xruc

        End If

        If local1.Visible = True Then
            consulta_xruc2 "" & xruc

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

Private Sub xtipo_keyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Frame7.Visible = False

        If Framefp.Visible = False Then
            DBGrid2.Col = 0
            DBGrid2.Row = DBGrid2.VisibleRows - 1
            DBGrid2.SetFocus
            Exit Sub

        End If

        Framefp.Enabled = True
        dbgrid10.Enabled = True
        dbgrid10.SetFocus
        Exit Sub

    End If

    'aqui es donde vamos a poner si modificacion pedido
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

    'MsgBox "abc"
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

    If local1 = "GUIAREMISION" Then 'pedido merca almacen
        xtipo = "" & mytable11.Fields("tipoot")
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

        If Len(codigo) = 11 Then
            xtipo = "2"

        End If

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

    'MsgBox "Hola"
    found = busca_xtipo("" & xtipo, 0)

    If found = 0 Then
        xtipo = ""
        MsgBox "No existe Tipo Documento", 48, "Aviso"
        xtipo.SetFocus
        Exit Sub

    End If

    xruc = Trim(codigo)

    If xtipo = "1" Or xtipo = "3" Or xtipo = "5" Then
        Label36 = "Codigo"

    End If

    If xtipo = "2" Or xtipo = "4" Then
        Label36 = dicruc

        If Len(xruc) <> 11 Then
            xruc = ""

        End If

    End If

    sentido.Enabled = False

    If sentido.Enabled = True Then
        sentido.SetFocus  'se adiciono concar.....
        Exit Sub

    End If

    If "" & mytable11.Fields("vendedor") = "S" Then
        If xvendedor.Visible = True Then
            xvendedor.SetFocus

        End If

        Exit Sub

    End If

    If flag_servicio = "D" Then  'validar el deliveri si ingreso datos
        xvendedor.SetFocus
        Exit Sub

    End If

    If "" & mytable11.Fields("cliente") <> "S" And acu <> "B" And acu <> "D" Then
        Command13_Click
        Exit Sub

    End If

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

Private Sub xvendedor_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If local1.Visible = True Or Len(pedido) > 0 Then 'si es traslado
        If Len(xvendedor) = 0 Then
            xvendedor = "" & cajero
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

    If Len(xvendedor) = 0 Then
        xvendedor = "" & cajero
        xvendedor.SetFocus
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

    Dim sdx      As Double

    Dim buf1     As String

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
      
            If "" & mytablex.Fields("tipodoc") = "T" Then  'guia remision
                xserie = "" & mytable11.Fields("serieot")
                sdx = Val("" & mytable11.Fields("numeroot")) + 1
                xnumero = "" & sdx
                xserie.Enabled = True
                xnumero.Enabled = True

            End If
      
            If "" & mytablex.Fields("tipodoc") = "Z" Then  'traslado
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
                'sdx = Val("" & mytable11.Fields("numeroFM")) + 1
                sdx = Val("" & mytablex.Fields("numero")) + 1
      
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
    mytablex.Open "SELECT * FROM vendedor where  codigo='" & xvendedor & "' ", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        nvendedorx = "" & mytablex.Fields("nombre")
        busca_xvendedor = 1

    End If

    '------------------------------------- ------------
    mytablex.Close
 
End Function

Function busca_xtipog(buf As String)

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd7888_err

    mytable11.Close
    mytable11.Open "SELECT * FROM parameca where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

    If mytable11.RecordCount > 0 Then
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
                If Val(tdetra) > 0 Then
                    mytable11.Fields("detraccion") = Val(ndetraccion)

                End If

                mytable11.Fields("numerobm") = xnumero
                mytable11.Fields("uvueltos") = Val(stxtotals)
                mytable11.Fields("uvueltod") = Val(stxtotald)
                mytable11.Update

            End If

            If "" & mytablex.Fields("tipodoc") = "B" Then
                'mytablex.Edit
                'MsgBox "PP"
                mytable11.Fields("numeroFM") = xnumero
                'MsgBox "PPP"
                'mytable11.Update
                'mytable11.Edit
         
                If Val(tdetra) > 0 Then
                    mytable11.Fields("detraccion") = Val(ndetraccion)

                End If

                mytable11.Fields("uvueltos") = Val(stxtotals)
                mytable11.Fields("uvueltod") = Val(stxtotald)
                mytable11.Update
                mytablex.Fields("numero") = xnumero
                mytablex.Update
         
                'MsgBox "xxx"
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

            If "" & mytablex.Fields("tipodoc") = "T" Then
                'mytable11.Edit
                'If Val(tdetra) > 0 Then
                'mytable11.Fields("detraccion") = Val(ndetraccion)
                'End If
                mytable11.Fields("numeroot") = xnumero
                'mytable11.Fields("uvueltos") = Val(stxtotals)
                'mytable11.Fields("uvueltod") = Val(stxtotald)
                mytable11.Update

            End If

            If "" & mytablex.Fields("tipodoc") = "D" Then

                'mytable11.Edit
                If Val(tdetra) > 0 Then
                    mytable11.Fields("detraccion") = Val(ndetraccion)

                End If

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
                If Val(tdetra) > 0 Then
                    mytable11.Fields("detraccion") = Val(ndetraccion)

                End If

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
                If Val(tdetra) > 0 Then
                    mytable11.Fields("detraccion") = Val(ndetraccion)

                End If

                mytable11.Fields("numeronc") = xnumero
                mytable11.Fields("uvueltos") = Val(stxtotals)
                mytable11.Fields("uvueltod") = Val(stxtotald)
                mytable11.Update

            End If

            If "" & mytablex.Fields("tipodoc") = "F" Then

                'mytable11.Edit
                If Val(tdetra) > 0 Then
                    mytable11.Fields("detraccion") = Val(ndetraccion)

                End If

                mytable11.Fields("numerond") = xnumero
                mytable11.Fields("uvueltos") = Val(stxtotals)
                mytable11.Fields("uvueltod") = Val(stxtotald)
                mytable11.Update

            End If

            If "" & mytablex.Fields("tipodoc") = "I" Then

                'MsgBox "x"
                'mytable11.Edit
                If Val(tdetra) > 0 Then
                    mytable11.Fields("detraccion") = Val(ndetraccion)

                End If

                mytable11.Fields("numerope") = xnumero
                mytable11.Fields("uvueltos") = Val(stxtotals)
                mytable11.Fields("uvueltod") = Val(stxtotald)
                mytable11.Update

            End If

            If "" & mytablex.Fields("tipodoc") = "Q" Then

                'MsgBox "x"
                'mytable11.Edit
                If Val(tdetra) > 0 Then
                    mytable11.Fields("detraccion") = Val(ndetraccion)

                End If

                mytable11.Fields("numerope") = xnumero
                mytable11.Fields("uvueltos") = Val(stxtotals)
                mytable11.Fields("uvueltod") = Val(stxtotald)
                mytable11.Update

            End If

        End If

        mytablex.Close

    End If

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

        If Trim("" & mytable11.Fields("bodega")) = Trim(xruc) Then
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

    '-----------------------------------------------
    If local1 = "GUIAREMISION" Then 'pedido merca almacen
        xtipo = "" & mytable11.Fields("tipoot")
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

        valida_total = 1
        Exit Function

    End If

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

        '-----------------------------
        If Len(xruc) = 0 Then
            xruc.SetFocus
            Exit Function

        End If

        found = busca_localx("" & xruc)

        If found = 0 Then
            MsgBox "digite un Almacen a Pedir", 48, "Aviso"
            xruc.SetFocus
            Exit Function

        End If

        If Trim("" & mytable11.Fields("bodega")) = Trim(xruc) Then
            MsgBox "Almacen debe ser diferente ", 48, "Aviso"
            xruc.SetFocus
            Exit Function

        End If

        xnombre.SetFocus
        '-----------------------------
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
            'sentido.SetFocus
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

    If "" & mytable11.Fields("vendedor") = "S" Then
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

        'If Len(xruc) <> 11 Then
        '   xruc.SetFocus
        '   Exit Function
        'End If
        found = valida_ruc("" & xruc)

        If found = 0 Then
            MsgBox dicruc & " No Valido", 48, "Aviso"
            xruc = ""
            xruc.SetFocus
            Exit Function

        End If

        'valida el ruc
    End If

    If Len(xruc) > 0 Then
        found = busca_codigocl("" & xruc, 1)

        If acu = "B" Or acu = "D" Then
            If Len(xnombre) = 0 Then
                xnombre.SetFocus
                Exit Function

            End If

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

    Dim xbuf     As String

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As Table

    Dim found    As Integer

    Dim sdx      As Double

    sdx = 0

    On Error GoTo cdm4411_err

    '---------- validando si es cuenta corriente
    'If mytablex.State = 1 Then mytablex.Close
    'mytablex.Open "SELECT * FROM " & fpusuario, cn, adOpenDynamic, adLockOptimistic
    'If mytablex.RecordCount = 0 Then  'si no existe
    '   mytablex.Close
    '   Exit Function
    'End If
amk223:
    mytabley.Open "SELECT * FROM " & gofpago & " where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

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
                graba_acumulado_clientes "" & mytabley.Fields("codigo"), 1, Val("" & mytabley.Fields("recibe"))

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

    If Trim(xvendedor) > 0 Then
        mytabley.Fields("vendedor") = xvendedor

    End If

    mytabley.Fields("paridad") = Val("" & paridadfp)
    mytabley.Fields("codigo") = "" & mytablex.Fields("codigo")
    mytabley.Fields("nombre") = "" & mytablex.Fields("nombre")
    mytabley.Fields("tipo") = "" & xtipo
    mytabley.Fields("serie") = "" & xserie
    mytabley.Fields("numero") = "" & xnumero
    mytabley.Fields("tipoclie") = "C"
   
    If Len(Trim("" & mytablex.Fields("codigo"))) = 0 Then
        mytabley.Fields("codigo") = Trim("" & xruc)

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
   
    mytabley.Fields("local") = Trim("" & "" & mytable11.Fields("local"))

    If "" & mytable11.Fields("terminal") = "T" Then

        'mytabley.Fields("acu") = "I"
    End If

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

    Dim buff     As String

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    'EXISTE MAS DE UN PRECIO-----
      
    sw = 0

    If flag_especial = "S" Then
        buff = "SELECT * FROM precio1 where producto='" & buf & "' and local='01' and codigo='" & codigo & "'"
        mytablex.Open buff, cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            GoTo amika1

        End If

    End If

    buff = "SELECT * FROM precios where producto='" & buf & "' and local='" & "" & mytable11.Fields("listap") & "'"

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open buff, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Function

    End If

amika1:

    If "" & Len("" & mytablex.Fields("unidad1")) > 0 Then
        If "" & Len("" & mytablex.Fields("unidad2")) > 0 Then
            sw = 1

        End If

    End If

    mytablex.Close

    If sw = 1 Then
        buff = "SELECT * FROM producto where producto='" & buf & "'"
        mytablex.Open buff, cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            If Val("" & mytablex.Fields("empaque_visible")) = 1 Then
                sw = 0

            End If

        End If

        mytablex.Close

    End If

    ver_si_puedo_dbgrid = sw

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
    Label22 = "2.Vendido 1.Anulado"
    sw_consulta = 0
    found = sql_consulta(1)
    'dbGrid1.SetFocus

End Sub

Sub menu_copia()

    Dim found As Integer

    Dim buf   As String

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

Sub menu_delivery()

    Dim found As Integer

    Dim buf   As String

    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.ListIndex = 0
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    opcion1 = "15A"
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

    Dim sw       As Integer

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM " & gocabeza & " where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & ytipo & "' and serie='" & yserie & "' and numero='" & ynumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        'mytablex.Edit
        mytablex.Fields("estado") = "1"
        mytablex.Update

    End If

    mytablex.Close
    sw = 0

    If mytablex.State = 1 Then mytablex.Close
    'MsgBox godetalle
    mytablex.Open "SELECT * FROM " & godetalle & " where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & ytipo & "' and serie='" & yserie & "' and numero='" & ynumero & "'", cn, adOpenStatic, adLockOptimistic  'adOpenDynamic

    If mytablex.RecordCount > 0 Then 'si existe
        found = descarga_saldo("" & "" & mytable11.Fields("local"), mytablex, ytipo, yserie, ynumero, 1, 1)
        mytablex.Close
        cn.Execute ("update " & godetalle & " set estado='1'" & " where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & ytipo & "' and serie='" & yserie & "' and numero='" & ynumero & "'")
        sw = 1

    End If

    If sw = 0 Then
        mytablex.Close

    End If

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM " & gofpago & " where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & ytipo & "' and serie='" & yserie & "' and numero='" & ynumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        Do

            If mytablex.EOF Then Exit Do
            'mytablex.Edit
            mytablex.Fields("estado") = "1"
            mytablex.Update

            If "" & mytablex.Fields("acufp") = "V" Then
                graba_acumulado_clientes "" & mytablex.Fields("codigo"), -1, Val("" & mytablex.Fields("recibe"))

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

    cn.Execute ("DELETE FROM  cuentap where local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & Trim("" & ytipo) & "' and serie='" & Trim("" & yserie) & "' and numero='" & Trim("" & ynumero) & "' and cuota='1'")

    If Len(ynumero) > 0 Then
        cn.Execute ("DELETE FROM  cuentap where local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & Trim("" & ytipo) & "' and serie='" & Trim("" & yserie) & "' and numeron='" & Trim("" & ynumero) & "' and cuota='1'")

    End If

    reversa_guia_mensual Trim("" & "" & mytable11.Fields("local")), ytipo, yserie, ynumero
    proceso_anular = 1

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
        mytablex.Fields("correo") = Mid$(Trim("" & correo), 1, 60)
        mytablex.Update

    End If

    mytablex.Close

End Function

Function graba_cliente_tipo(buf As String)

    Dim mytablex  As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    Dim sdx       As Double

    Dim buf1      As String

    Dim codigogen As String

    On Error GoTo cmdd7812_err

    'MsgBox codigo

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

        xruc = Trim(codigogen)
        mytablex.AddNew
        mytablex.Fields("codigo") = "" & xruc
        mytablex.Fields("tipo") = "O"
        mytablex.Fields("nombre") = "" & xnombre
        mytablex.Fields("correo") = Mid$(Trim("" & correo), 1, 60)
        mytablex.Fields("direccion") = "" & xdireccion
        mytablex.Update
        xruc = Trim("" & mytablex.Fields("codigo"))
        'codigo = "" & mytablex.Fields("codigo")
        'nombre = "" & mytablex.Fields("nombre")
        mytablex.Close
        Exit Function

    End If

    If Len(xruc) > 0 And Len(xnombre) > 0 Then
        mytablex.Open "SELECT * FROM clientes  where  codigo='" & xruc & "'", cn, adOpenDynamic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            mytablex.Fields("nombre") = Trim("" & xnombre)

            If Len(Trim("" & correo)) > 0 Then
                mytablex.Fields("correo") = Mid$(Trim("" & correo), 1, 60)

            End If

            If Len("" & xdireccion) > 0 Then
                mytablex.Fields("direccion") = Trim("" & xdireccion)

            End If

            mytablex.Update
        Else
            mytablex.AddNew
            mytablex.Fields("nombre") = "" & xnombre
            mytablex.Fields("codigo") = "" & xruc
            mytablex.Fields("correo") = Mid$(Trim("" & correo), 1, 60)

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

        If Len(Trim(xvendedor)) > 0 Then
            mytablex.Fields("vendedor") = xvendedor

        End If

        mytablex.Fields("usuario") = cajero
        mytablex.Fields("caja") = caja
        mytablex.Fields("turno") = turno
        mytablex.Fields("zona") = ""
        mytablex.Fields("local") = "" & "" & mytable11.Fields("local")
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

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim buf1     As String

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

Sub graba_tmpcta(mytablez As ADODB.Recordset, _
                 mytablex As ADODB.Recordset, _
                 mytabley As ADODB.Recordset, _
                 sdx As Double)

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
    mytablex.Open "SELECT * FROM cuentac where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "' and cuota='1'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        mytablex.Delete
        GoTo amk2

    End If

    mytablex.Close

End Function

Function menu_repone(xcongela As String)

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd67112_err

    mytablex.Open "SELECT * FROM drequisa  where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='01' and serie='Q' and numero='" & xcongela & "'", cn, adOpenDynamic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        Data2.Recordset.AddNew

        For I = 0 To mytablex.Fields.count - 1
            Data2.Recordset.Fields(I) = mytablex.Fields(I)
        Next I

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

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd6711_err

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM congelad where numero='" & xcongela & "' order by hora", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Function

    End If
   
    Do

        If mytablex.EOF Then Exit Do
        'If "" & mytablex.Fields("numero") = xcongela Then
        Data2.Recordset.AddNew

        For I = 0 To mytablex.Fields.count - 1
            Data2.Recordset.Fields(I) = mytablex.Fields(I)
        Next I

        Data2.Recordset.Fields("caja") = "" & caja
        Data2.Recordset.Fields("turno") = "" & turno
        Data2.Recordset.Fields("usuario") = "" & cajero
        Data2.Recordset.Fields("fecha") = Format(dia, "dd/mm/yyyy")
        'Data2.Recordset.Fields("hora") = Format(Now, "hh:MM")
        Data2.Recordset.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
        Data2.Recordset.Fields("estado") = "2"
        Data2.Recordset.Update
        mytablex.MoveNext
        '   Else: Exit Do
        'End If
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
    'On Error GoTo cmd133_err
    'Data1.Recordset.Delete
    'Exit Sub
    'cmd133_err:
    'Exit Sub

End Sub

Sub borrar_descongela1(xcongela As String)
    cn.Execute ("DELETE   FROM congelad WHERE numero='" & Trim(xcongela) & "'")

End Sub

Sub borrar_repone(xcongela As String)
    cn.Execute ("DELETE   FROM drequisa WHERE local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='01' and serie='Q' and numero='" & xcongela & "'")

End Sub

Function descarga_saldo(bxlocal As String, _
                        mytablex As ADODB.Recordset, _
                        bxtipo As String, _
                        bxserie As String, _
                        bxnumero As String, _
                        sw As Integer, _
                        sw1 As Integer)

    Dim mytabley As New ADODB.Recordset

    Dim sdx      As Double

    Dim indx     As Double

    On Error GoTo cmd8900_err

    'Exit Function
    'MsgBox "" & mytablex.Fields("local")
    indx = 0
  
    If mytablex.RecordCount = 0 Then Exit Function
    mytablex.MoveFirst
  
    Do

        If mytablex.EOF Then Exit Do
        If mytabley.State = 1 Then
            mytabley.Close
            Set mytabley = Nothing

        End If

        If Len("" & mytablex.Fields("proveedorp")) > 0 Then GoTo nohacer
     
        indx = indx + 1
        mytabley.Open "SELECT * FROM almacen where  local='" & Trim("" & mytablex.Fields("local")) & "' and producto='" & Trim("" & mytablex.Fields("producto")) & "' and bodega='" & Trim("" & mytablex.Fields("bodega")) & "'", cn, adOpenDynamic, adLockOptimistic

        If mytabley.RecordCount > 0 Then  'si existe
            sdx = Val("" & mytabley.Fields("saldo")) + sw * Val("" & mytablex.Fields("factor")) * Val("" & mytablex.Fields("cantidad"))
            mytabley.Fields("saldo") = sdx
            pone_tallas_saldo mytabley, mytablex, sw
            mytabley.Update
        Else
            mytabley.AddNew
            mytabley.Fields("local") = Trim("" & "" & mytable11.Fields("local"))
            mytabley.Fields("producto") = Trim("" & mytablex.Fields("producto"))
            mytabley.Fields("bodega") = Trim("" & mytablex.Fields("bodega"))
            mytabley.Fields("unidad") = Trim("" & mytablex.Fields("unidad"))
            sdx = Val("" & mytabley.Fields("saldo")) + sw * Val("" & mytablex.Fields("factor")) * Val("" & mytablex.Fields("cantidad"))
            mytabley.Fields("saldo") = sdx
            pone_tallas_saldo mytabley, mytablex, sw
            mytabley.Update

        End If

        mytabley.Close
nohacer:
        mytablex.MoveNext
    Loop
    Exit Function
cmd8900_err:
    MsgBox "Aviso en descarga saldo " & "" & indx & " " & error$, 48, "Aviso"
    Exit Function

End Function

Function descarga_saldos(bxlocal As String, _
                         mytablex As ADODB.Recordset, _
                         bxtipo As String, _
                         bxserie As String, _
                         bxnumero As String, _
                         sw As Integer, _
                         sw1 As Integer)

    Dim mytabley As New ADODB.Recordset

    Dim sdx      As Double

    Dim indx     As Double

    On Error GoTo cmd88900_err

    If Len("" & mytablex.Fields("proveedorp")) > 0 Then Exit Function
    mytabley.Open "SELECT * FROM almacen where  local='" & Trim("" & mytablex.Fields("local")) & "' and producto='" & Trim("" & mytablex.Fields("producto")) & "' and bodega='" & Trim("" & mytablex.Fields("bodega")) & "'", cn, adOpenDynamic, adLockOptimistic

    If mytabley.RecordCount > 0 Then  'si existe
        sdx = Val("" & mytabley.Fields("saldo")) + sw * Val("" & mytablex.Fields("factor")) * Val("" & mytablex.Fields("cantidad"))
        mytabley.Fields("saldo") = sdx
        pone_tallas_saldo mytabley, mytablex, sw
        mytabley.Update
    Else
        mytabley.AddNew
        mytabley.Fields("producto") = Trim("" & mytablex.Fields("producto"))
        mytabley.Fields("local") = Trim("" & "" & mytable11.Fields("local"))
        mytabley.Fields("producto") = Trim("" & mytablex.Fields("producto"))
        mytabley.Fields("bodega") = Trim("" & mytablex.Fields("bodega"))
        mytabley.Fields("unidad") = Trim("" & mytablex.Fields("unidad"))
        sdx = Val("" & mytabley.Fields("saldo")) + sw * Val("" & mytablex.Fields("factor")) * Val("" & mytablex.Fields("cantidad"))
        mytabley.Fields("saldo") = sdx
        pone_tallas_saldo mytabley, mytablex, sw
        mytabley.Update

    End If

    mytabley.Close
    Exit Function
cmd88900_err:
    MsgBox "Aviso en descarga saldos " & "" & indx & " " & error$, 48, "Aviso"
    Exit Function

End Function

Function proceso_carga_doc_ant(xlocal As String, _
                               xtipo As String, _
                               xserie As String, _
                               xnumero As String)

    Dim I        As Integer

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd67112_err

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM " & godetalle & " where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        Do

            If mytablex.EOF Then Exit Do
            Data2.Recordset.AddNew

            For I = 0 To mytablex.Fields.count - 1
                Data2.Recordset.Fields(I) = mytablex.Fields(I)
            Next I

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

Function proceso_carga_Pedido(xlocal As String, _
                              xtipo As String, _
                              xserie As String, _
                              xnumero As String)

    Dim I        As Integer

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd67112_err

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM dpedidov where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        Do

            If mytablex.EOF Then Exit Do
            Data2.Recordset.AddNew

            For I = 0 To mytablex.Fields.count - 1
                Data2.Recordset.Fields(I) = mytablex.Fields(I)
            Next I

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

Function proceso_carga_guia(xlocal As String, _
                            xtipo As String, _
                            xserie As String, _
                            xnumero As String)

    Dim I        As Integer

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd671124_err

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM detalle where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "' and acu='T'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        Do

            If mytablex.EOF Then Exit Do
            Data2.Recordset.AddNew

            For I = 0 To mytablex.Fields.count - 2
                Data2.Recordset.Fields(I) = mytablex.Fields(I)
            Next I

            Data2.Recordset.Update
            proceso_carga_guia = 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
    found = sumar_detalle()
    DBGrid2.Col = 0
    DBGrid2.Row = DBGrid2.VisibleRows - 1
    DBGrid2.SetFocus
    Exit Function
cmd671124_err:
    mytablex.Close
    MsgBox "Aviso en carga Guia " + error$, 48, "Aviso"
    Exit Function

End Function

Function proceso_carga_cotizacion(xlocal As String, _
                                  xtipo As String, _
                                  xserie As String, _
                                  xnumero As String)

    Dim I        As Integer

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd67112_err

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM dcotizav where  local='" & xlocal & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        Do

            If mytablex.EOF Then Exit Do
            Data2.Recordset.AddNew

            For I = 0 To mytablex.Fields.count - 1
                Data2.Recordset.Fields(I) = mytablex.Fields(I)
            Next I

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

Function proceso_proforma(xlocal As String, _
                          xtipo As String, _
                          xserie As String, _
                          xnumero As String)

    Dim I        As Integer

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    sw = 0

    On Error GoTo cmd6711212_err

    'MsgBox "" & "" & mytable11.Fields("local") & " " & xtipo & " " & xserie & " " & xnumero
    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM dproform where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        Do

            If mytablex.EOF Then Exit Do
            'MsgBox ""
            Data2.Recordset.AddNew

            For I = 0 To mytablex.Fields.count - 2
                Data2.Recordset.Fields(I) = mytablex.Fields(I)
            Next I

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

    Dim I       As Long

    Dim D       As Integer

    Dim buffers As String

    Select Case "" & mytable11.Fields("portbala")

        Case "COM1"
            D = 1

        Case "COM2"
            D = 2

        Case "COM3"
            D = 3

        Case "COM4"
            D = 4

        Case "COM5"
            D = 5
           
    End Select

    If "" & mytable11.Fields("tipo_balanza") = "1" Then
        puerto_balanza1 = acura_lectura()
        Exit Function

    End If

    If "" & mytable11.Fields("tipo_balanza") = "2" Then
        puerto_balanza1 = acura_lectura()
        Exit Function

    End If

    MSComm1.CommPort = D
    MSComm1.Settings = "9600,n,8,1"
    MSComm1.InputLen = 10
    MSComm1.PortOpen = True
    MSComm1.Output = Chr$(80)
    buffers = ""
    'For i = 1 To 9000
    'Next i
    I = 0
    Do
        'DoEvents
        buffers = buffers & MSComm1.input
        I = I + 1

        If I > 15000 Then
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

    Dim buf         As String

    Dim afgodetalle As String

    Dim fgodetalle  As String

    Dim mytablex    As New ADODB.Recordset

    On Error GoTo cmd344_err

    afgodetalle = godetalle
    fgodetalle = godetalle
    dbgrid6.Visible = True

    If opcion1 = "1900" Then  'proformas
        fgodetalle = "dproform"

    End If

    buf = "select Producto,Descripcio,Unidad,Factor,Cantidad as Cant,Precio,Total from " & fgodetalle & " where local='" & "" & "" & mytable11.Fields("local") & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'"
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

    Dim mytablex As New ADODB.Recordset

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

            If "" & mytable11.Fields("fpro") <> "S" Then
                Exit Function

            End If

        Case "T"  'DE PEDIDOS

            If "" & mytable11.Fields("fnv") <> "S" Then
                Exit Function

            End If
             
        Case Else
            'MsgBox buf
            mytablex.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic
                 
            If mytablex.RecordCount > 0 Then 'si existe

                'MsgBox "" & mytablex.Fields("tipodoc")
                Select Case "" & mytablex.Fields("tipodoc")

                    Case "T"  'guia remision
                                
                        If "" & mytable11.Fields("fot") = "S" Then
                            mytablex.Close
                            GoTo cvye

                        End If

                    Case "I"  'a cuenta da algo
                                
                        If "" & mytable11.Fields("fpro") = "S" Then
                            mytablex.Close
                            GoTo cvye

                        End If

                End Select

            End If

            mytablex.Close
            Exit Function

    End Select

cvye:
    valida_tipo_pago = 1

End Function

Function borrar_proformas()

    On Error GoTo cmd89900_err

    cn.Execute ("delete from cproform where local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & protipo & "' and serie='" & proserie & "' and numero='" & pronumero & "'")
    cn.Execute ("delete from dproform where local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & protipo & "' and serie='" & proserie & "' and numero='" & pronumero & "'")
    cn.Execute ("delete from ppocket where pedido='" & pronumero & "'")
    Exit Function
cmd89900_err:
    MsgBox "Aviso en borrar proformas " + error$, 48, "Aviso"
    Exit Function

End Function

Function borrar_pedidos()

    Dim xbuf     As String

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
    mytablex.Open "SELECT * FROM ccotizav where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & petipo & "' and serie='" & peserie & "' and numero='" & penumero & "'", cn, adOpenDynamic, adLockOptimistic

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

    Dim buf1     As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd89121_err

    label56 = ""

    'MsgBox buf
    If flag_especial = "S" Then
        buf1 = "SELECT * FROM precio1 where producto='" & buf & "' and local='01' and codigo='" & codigo & "'"
        mytablex.Open buf1, cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            GoTo amika2

        End If

        mytablex.Close

    End If

    buf1 = "SELECT * FROM precios where producto='" & buf & "' and local='" & Trim("" & mytable11.Fields("listap")) & "'"
    mytablex.Open buf1, cn, adOpenStatic, adLockOptimistic
amika2:

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
    mytablex.Open "SELECT * FROM almacen where local='" & Trim("" & "" & mytable11.Fields("local")) & "' and producto='" & "" & buf & "' and bodega='" & Trim("" & mytable11.Fields("bodega")) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        label56 = label56 + Chr$(10) + Chr$(13)
        label56 = label56 + " Saldo:" & Trim("" & dbGrid1.columns(3)) & " " & calcula_saldo(Val("" & mytablex.Fields("saldo")), Val("" & dbGrid1.columns(4)))

    End If

    mytablex.Close
    Exit Sub
cmd89121_err:
    'MsgBox "Aviso en pone_precios " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function consulta_saldo(buf As String, cant As Double, sw As Integer) As Double

    Dim mytablex As New ADODB.Recordset

    Combo1.Clear
    Combo1.AddItem "bodega"
    Combo1.ListIndex = 0
    'AQUI DEBE VERIFICAR SI EXISTE PRODUCTO
    mytablex.Open "SELECT * FROM almacen where local='" & Trim("" & "" & mytable11.Fields("local")) & "' and bodega='" & Trim("" & mytable11.Fields("bodega")) & "' and producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
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
        found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("bold"), "" & mytable11.Fields("letrainterna"))
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

    Dim buf   As String

    Dim found As Integer

    Dim I     As Integer

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

    For I = 1 To 11
        found = formateaa("", 1, 2, 0)
    Next I

    DBGrid2.SetFocus
    Exit Sub
cmd3999_err:
    MsgBox "Error en cuerpo estado cuenta " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub cabecera_estado_cuenta()

    Dim found As Integer

    Dim buf   As String

    Dim btipo As String

    On Error GoTo cmd4111_err

    buf = String(36, "-")
    found = formateaa(buf, 36, 2, 0)
    buf = "       ESTADO DE CUENTA"
    found = formateaa(buf, 36, 2, 0)
    buf = "    Cajero:" & cajero & " Caja:" & caja & " Turno:" & turno
    found = formateaa(buf, 36, 2, 0)
    buf = "  Fecha:" & Format(Now, "dd/mm/yyyy") & "  Hora:" & Format(Now, "hh:mm:ss")
    found = formateaa(buf, 36, 2, 0)

    If flag_servicio = "A" Then
        found = formateaa(" *** RAPIDO    ***", 25, 2, 0)

    End If

    'If tservicio = "C" Then
    '   buf = "   Salon : " & salon & " Mesa:" & mesa
    '   found = formateaa(buf, 36, 2, 0)
    'End If
    If flag_servicio = "D" Then
        found = formateaa(" *** DOMICILIO ***", 36, 2, 0)
        found = formateaa(buf, 36, 2, 0)

        'imprime_cliente_delivery "" & codigocli
    End If

    buf = String(36, "-")
    found = formateaa(buf, 36, 2, 0)
    Exit Sub
cmd4111_err:
    MsgBox "Mensaje,Error en cabecera Pedido " & error$
    Exit Sub

End Sub

Sub imprime_estado_cuenta()

    Dim buf   As String

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

    Dim raconsulta As New ADODB.Recordset

    On Error GoTo cmd87678_err

    'buf = "select * from almacen where producto='" & buf & "'"
    buf = "select Almacen.saldo,almacen.unidad,Bodega.nombre,almacen.bodega,Almacen.local from almacen left join bodega on almacen.bodega=bodega.codigo where almacen.producto='" & buf & "' order by almacen.bodega"

    If raconsulta.State = 1 Then raconsulta.Close
    raconsulta.Open buf, cn, adOpenStatic, adLockOptimistic

    If raconsulta.RecordCount = 0 Then
        raconsulta.Close
        Exit Sub

    End If

    Set dbgrid7.DataSource = raconsulta
    Exit Sub
cmd87678_err:
    MsgBox "Aviso en sql-saldo local " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub limpia_general()

    Dim found As Integer

    Frame7.Visible = False
    Framefp.Visible = False
    borrar_todo
    sql_detalle
    found = sumar_detalle()
    tiposervicio1 = "Autoservicio"
    flag_servicio = "A"

    'Frame10.Visible = True
End Sub

Sub proceso_cierre_automatico(buf As String)

    Dim found As Integer

    Dim buf1  As String

    If Frame2.Visible = True Then Exit Sub
    local1.Visible = False
    local1.Visible = False
    found = sumar_detalle()

    If found = 2 Then
        If "" & mytable11.Fields("vdetalle") = "S" Then
            MsgBox "Debe existir Vendedor ", 48, "Aviso"
            DBGrid2.SetFocus
            Exit Sub

        End If

    End If

    If found = 0 Then
        MsgBox "debe de Existir un Precio=0", 48, "Aviso"
        DBGrid2.SetFocus
        Exit Sub

    End If

    If flag_percepcion = "S" Then
        If Len(Trim("" & codigo)) = 0 Then
            MsgBox "Existe Percepcion ,Debe ponerse Dato Cliente ", 48, "Aviso"
            codigo.SetFocus
            Exit Sub

        End If

    End If

    If Val(txtotal) = 0 Then
        If exisdev <> -10 Then  'si existe devolucion
            DBGrid2.SetFocus
            Exit Sub

        End If

    End If

    If Trim("" & mytable11.Fields("terminal")) = "T" Or (Val(acuenta) > 0 And Len(petipo) = 0) Then 'pedidos o acuenta>0
        'MsgBox "Hola"
        xruc = Trim(codigo)
        xnombre = nombre
        Frame7.Visible = True
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
    'MsgBox ""
    ttxtotals = Format(Val(rtxtotal), nrodecimal)
    ttxtotald = Format(Val(rtxtotald), nrodecimal)
    stxtotals = Format(Val(rtxtotal), nrodecimal)
    stxtotald = Format(Val(rtxtotald), nrodecimal)
    found = leer_visorcaja(dicmoneda & stxtotals, "US$  " & stxtotald)

    Framefp.Visible = True

    If xfpagox.State = 1 Then
        xfpagox.Close
        Set xfpagox = Nothing

    End If

    buf1 = "1"

    Select Case buf

        Case "EFECTIVO"
            buf1 = "1"

        Case "DOLAR"
            buf1 = "2"

        Case "TARJETACREDITO"
            buf1 = "4"

        Case "CREDITO"
            buf1 = "3"

    End Select
    
    If xfpagox.State = 1 Then
        xfpagox.Close

    End If

    xfpagox.Open "SELECT * FROM fpago where fpago='" & buf1 & "'", cn, adOpenDynamic, adLockOptimistic
    Set dbgrid10.DataSource = xfpagox
    dbgrid10.refresh

    If xfpagox.RecordCount > 0 Then
        'MsgBox ""
        dbgrid10.Enabled = True
        Framefp.Enabled = True
        'dbgrid10.Visible = True
        dbgrid10.SetFocus
        DBGrid10_KeyDown 13, 0

        If buf1 = "1" Then
            dbgrid9.SetFocus
            xtipo = "1"
            'DBGrid9_KeyDown 13, 0
            xtipo_keyPress 13
         
        End If

        'xtipo = "7"
    Else
        MsgBox "No existe exonerado ", 48, "Aviso"

    End If

End Sub

Sub menu_graba_fpedido()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM fpagov where 2=1", cn, adOpenDynamic, adLockOptimistic
    graba_fpago_pedido mytablex
    'found = graba_credito_trabajo() 'RECIEN LO DESHABILITE
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
    mytabley.Fields("local") = Trim("" & "" & mytable11.Fields("local"))

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
    Exit Sub
   
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
    mytabley.Fields("local") = Trim("" & "" & mytable11.Fields("local"))

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

    Dim sdx   As Double

    Dim a     As Double

    'MsgBox tipodescuento
    Data2.refresh
    Do

        If Data2.Recordset.EOF Then Exit Do
        If (Val("" & Data2.Recordset.Fields("cantidad")) > 0 Or Val("" & Data2.Recordset.Fields("cantidad")) < 0) And Val("" & Data2.Recordset.Fields("precio")) > 0 Then
            Data2.Recordset.Edit

            'MsgBox tipodescuento
            If tipodescuento = "2" Then
                Data2.Recordset.Fields("deslipo") = 0
                resuma_precios 0

            End If

            If tipodescuento = "0" Then
                Data2.Recordset.Fields("deslipo") = Val(valordescuento)
                        
            End If

            If tipodescuento = "1" Then
                a = (Val(valordescuento) * 100) / Val(txtotal)
                Data2.Recordset.Fields("deslipo") = a
                        
            End If

            If tipodescuento = "3" Then   '----recargos

                'Data2.Recordset.Fields("t14") = Val(valordescuento)
                'Data2.Recordset.Fields("deslipo") = 0
                If Val(valordescuento) > 0 Then
                    Data2.Recordset.Fields("precio") = Val("" & Data2.Recordset.Fields("precio")) + Val("" & Data2.Recordset.Fields("precio")) * valordescuento / 100

                End If

                If Val(valordescuento) < 0 Then
                    sdx = 1 + Abs(Val(valordescuento)) / 100
                    Data2.Recordset.Fields("precio") = Val("" & Data2.Recordset.Fields("precio")) / sdx

                End If

            End If

            resuma_precios 0
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
    resuma_precios 0

End Sub

Function graba_credito_trabajo()

    Dim mytablex As New ADODB.Recordset

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM cuentac where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "' and cuota='1'", cn, adOpenDynamic, adLockOptimistic

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

    If Len(Trim(xvendedor)) > 0 Then
        mytablex.Fields("vendedor") = "" & xvendedor

    End If

    mytablex.Fields("zona") = ""
    mytablex.Fields("local") = Trim("" & "" & mytable11.Fields("local"))
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

    Dim sdx      As Double

    'ADICIONAR EL PAGO
    mytabley.Open "SELECT * FROM cuentacd where local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & petipo & "' and serie='" & peserie & "' and numero='" & penumero & "' and cuota='1'", cn, adOpenStatic, adLockOptimistic
    mytabley.AddNew

    mytabley.Fields("codigo") = "" & xruc
    mytabley.Fields("local") = Trim("" & "" & mytable11.Fields("local"))
    mytabley.Fields("local1") = Trim("" & "" & mytable11.Fields("local"))
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

    mytablex.Open "SELECT * FROM cuentac where local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & petipo & "' and serie='" & peserie & "' and numero='" & penumero & "' and cuota='1'", cn, adOpenStatic, adLockOptimistic

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
    mytablex.Open "SELECT * FROM recibo where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & xtipo & "' and serie='" & xserie & "' and numero='" & xnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then 'si existe
        mytablex.AddNew
        mytablex.Fields("local") = Trim("" & "" & mytable11.Fields("local"))
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
        If Len(Trim(xvendedor)) > 0 Then
            mytablex.Fields("vendedor") = xvendedor

        End If

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
    mytabley.Fields("local") = Trim("" & "" & mytable11.Fields("local"))

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
    fotonombre = globalpath & "\001d\06\grafico\tmp.jpg"

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

    Dim buf  As String

    Dim sdx  As Double

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

    Dim sdx      As Double

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

    Dim I   As Integer

    Dim j   As Integer

    Dim AA  As String

    Dim BB  As String

    Dim cC  As String

    Dim dd  As String

    On Error GoTo cmd12004992_err

    'MsgBox crucefa.ListCount
    For I = 0 To crucefa.ListCount - 1
        extrae_crucefa crucefa.List(I), AA, BB, cC, dd
        'MsgBox AA & BB & cC & dd
        buf = "update cuentac set estado='1'  where  local='" & "" & AA & "' and tipo='" & "" & BB & "' and serie='" & "" & cC & "' and  numero='" & "" & dd & "'"
        mydbxglo.Execute buf
    Next I

    Exit Function
cmd12004992_err:
    Exit Function
    MsgBox "Aviso en graba_guia Mensual" + error$, 24, "AVISO DE NO ERROR"
    Resume

End Function

Sub reversa_guia_mensual(axlocal As String, _
                         axtipo As String, _
                         axserie As String, _
                         axnumero As String)

    Dim buf As String

    buf = "update cuentac set estado='0'  where  local='" & axlocal & "' and tipo='" & axtipo & "' and serie='" & axserie & "' and  numero='" & axnumero & "'"
    cn.Execute buf

End Sub

Sub extrae_crucefa(DATO As String, _
                   ccampo1 As String, _
                   ccampo2 As String, _
                   ccampo3 As String, _
                   ccampo4 As String)

    Dim I    As Integer

    Dim j    As Integer

    Dim temp As String

    I = 0
    temp = Trim$(DATO)

    If Len(temp) = 0 Then Exit Sub
    Do
        j = InStr(temp, "|")

        If j > 0 Then
            I = I + 1

            Select Case I

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

    Dim buf11    As String

    Dim mytablex As New ADODB.Recordset

    If flag_especial = "S" Then
        buf11 = "SELECT * FROM precio1  where  producto='" & buf & "' and local='01' and codigo='" & codigo & "'"
        mytablex.Open buf11, cn, adOpenDynamic, adLockOptimistic

        If mytablex.RecordCount > 0 Then 'si existe
            GoTo amika3

        End If

        mytablex.Close

    End If

    buf11 = "SELECT * FROM precios  where  producto='" & buf & "' and local='" & Trim("" & mytable11.Fields("listap")) & "'"

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open buf11, cn, adOpenDynamic, adLockOptimistic
amika3:

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

    Dim sdx      As Double

    Dim found    As Integer

    Dim buf1     As String

    saldoabo = ""

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

End Function

Function busca_credito_adelanto1(buf1 As String, buf As String) As Double

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    Dim buf2     As String

    Dim found    As Integer

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

    Dim mytabley As New ADODB.Recordset

    Dim sdx      As Double

    trdescuento = ""
    saldo = ""
    tabla_percepcion = 0

    If Len(Trim("" & codigo)) = 11 Then
        tabla_percepcion = 2
    Else
        tabla_percepcion = 0

    End If

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM clientes  where  codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        nombre = Trim("" & mytablex.Fields("nombre"))

        If Len(Trim(correo)) = 0 Then
            correo = Trim("" & mytablex.Fields("correo"))

        End If

        trdescuento = Format(Val("" & mytablex.Fields("descuento")), "0.00")
        saldo = Format(Val("" & mytablex.Fields("credito")), "0.00")
        busca_codigo_descuento = 1

        If "" & mytablex.Fields("especial") = "1" Then
            flag_especial = "S"

        End If

        'buscamos tambien si tiene percepcion
        mytabley.Open "select * from clasesunat where clasesunat='" & Trim("" & mytablex.Fields("clasesunat")) & "'", cn, adOpenDynamic, adLockOptimistic

        If mytabley.RecordCount > 0 Then
            If Val("" & mytabley.Fields("percepcion")) > 0 Then
                tabla_percepcion = Val("" & mytabley.Fields("percepcion"))

            End If

        End If

        mytabley.Close

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

    Dim buf1     As String

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    If Trim("" & mytable11.Fields("pm")) <> "S" Then
        valida_rango = 1
        Exit Function

    End If

    If flag_especial = "S" Then
        buf1 = "SELECT * FROM precio1  where  producto='" & Trim(DBGrid2.columns("producto")) & "' and local='01' and codigo='" & codigo & "'"
        mytablex.Open buf1, cn, adOpenDynamic, adLockOptimistic

        If mytablex.RecordCount > 0 Then 'si existe
            GoTo amika4

        End If

        mytablex.Close

    End If

    buf1 = "SELECT * FROM precios  where  producto='" & Trim(DBGrid2.columns("producto")) & "' and local='" & Trim("" & "" & mytable11.Fields("local")) & "'"

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open buf1, cn, adOpenDynamic, adLockOptimistic
amika4:

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

    Dim vr

    On Error GoTo cm64312_err

    Dim ufile As String

    Exit Sub
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

    ufile = globaldir & "\video\" & buf

    If UCase(gocabeza) = "CADIARIO" Then
        ufile = globaldir & "\cavideo\" & buf

    End If

    'ezVidCap1.TimeLimit = CInt("" & mytable11.Fields("segundo"))
    'ezVidCap1.CaptureFile = ufile
    'Call ezVidCap1.CaptureVideo
    ' Frame10.Height = 2175
    'Frame10.Top = 0
    'Frame10.Left = 10680
    'Frame10.Width = 3855
      
    '     ezVidCap1.Height = 1920
    '     ezVidCap1.Top = 240
    '     ezVidCap1.Left = 0
    '     ezVidCap1.Width = 3840

    'Frame10.Left = 10560
    '      Frame10.Height = 1445
    '      Frame10.Top = 840
    '      Frame10.Width = 3855
    '
    Exit Sub
cm64312_err:
    MsgBox "Aviso en Video " + error$, 48, "Aviso"
    Exit Sub

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

    Dim X        As Double

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim ufile    As String

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

Function crea_nuevos_clientes(buf1 As String, _
                              buf2 As String, _
                              buf3 As String, _
                              buf4 As String, _
                              buf5 As String, _
                              buf6 As String, _
                              buf7 As String)

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

    Dim I            As Integer

    Dim xbuf         As String

    Dim xsw          As Integer

    Dim found        As Integer

    Dim mytableyz    As Table

    Dim mytableyzx   As New ADODB.Recordset

    Dim mytablex     As New ADODB.Recordset

    Dim mytableb     As New ADODB.Recordset

    Dim antgocabeza  As String

    Dim antgodetalle As String

    Dim indx         As Integer

    Dim rs

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

    'MsgBox gocabeza
    'MsgBox xvendedor

    found = busca_xtipog("" & bxtipo)  'graba el numero al actual

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

    mytablex.Open "SELECT * FROM " & gocabeza & " where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

    'MsgBox ""
    If mytablex.RecordCount > 0 Then  'si existe
        'mytablex.Edit
        grabando_cabecera mytablex, bxtipo, bxserie, bxnumero
        mytablex.Update
    Else
        'MsgBox "DEMOS"
        mytablex.AddNew
        grabando_cabecera mytablex, bxtipo, bxserie, bxnumero
        mytablex.Update

    End If

    mytablex.Close

    'MsgBox ""
    If Len(petipo) > 0 And Len(penumero) > 0 Then  'si ha sido jalado pedido o orden trabajo descontar credito
        found = descuenta_credito_pedido()

    End If

    'MsgBox ""

    'mirar si el producto el porcentaje comision es mayor que cero
    xsw = 0
    Set mytableyz = mydbxglo.OpenTable(fpusuario)
    Do

        If mytableyz.EOF Then Exit Do
        If "" & mytableyz.Fields("acu") = "D" Then
            xsw = 1
            Exit Do

        End If

        mytableyz.MoveNext
    Loop
    mytableyz.Close
    'MsgBox "abc"
    Data2.refresh
ak1:

    If mytablex.State = 1 Then
        mytablex.Close
        Set mytablex = Nothing

    End If

    mytablex.Open "SELECT * FROM " & godetalle & " where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='" & bxtipo & "' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then  'si existe
        mytablex.Delete
        GoTo ak1

    End If

    'MsgBox ""
    'aqui debe borrar el otro si es traslado
    '---------------MsgBox local1.Visible
    If local1.Visible = True Then
ak12:

        If mytableb.State = 1 Then
            mytableb.Close
            Set mytableb = Nothing

        End If

        'borra lsi hay traslado
        mytableb.Open "SELECT * FROM detalle where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='TE' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

        If mytableb.RecordCount > 0 Then  'si existe
            mytableb.Delete
            GoTo ak12

        End If

ak123:

        If mytableb.State = 1 Then
            mytableb.Close
            Set mytableb = Nothing

        End If

        mytableb.Open "SELECT * FROM detalle where  local='" & Trim("" & "" & mytable11.Fields("local")) & "' and tipo='TS' and serie='" & bxserie & "' and numero='" & bxnumero & "'", cn, adOpenDynamic, adLockOptimistic

        If mytableb.RecordCount > 0 Then  'si existe
            mytableb.Delete
            GoTo ak123

        End If

    End If 'fin local visible

    '------------------------------------------------------
    'MsgBox ""
    indx = 0
    xbuf = "CABECERA:" & Format(Now, "hh:mm:ss")
    Set rs = Data2.Recordset.Clone
    Do

        If rs.EOF Then Exit Do
        mytablex.AddNew

        For I = 0 To rs.Fields.count - 1
            mytablex.Fields(I) = rs.Fields(I)
        Next I
    
        If xsw = 1 Then  'si alguna venta fue a credito debera buscar la comision a credito
            If mytableyzx.State = 1 Then mytableyzx.Close
            mytableyzx.Open "SELECT comisioncredito FROM producto where  producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenDynamic, adLockOptimistic

            If mytableyzx.RecordCount > 0 Then  'si existe
                mytablex.Fields("comision") = Val("" & mytableyzx.Fields("comisioncredito"))

            End If

            mytableyzx.Close

        End If
    
        If Val(tdetra) > 0 Then
            mytablex.Fields("denumero") = Format(Val(ndetraccion), "0000000000")

        End If

        mytablex.Fields("sentido") = "" & sentido
        mytablex.Fields("local") = Trim("" & "" & mytable11.Fields("local"))
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
        mytablex.Fields("localf") = Trim("" & "" & mytable11.Fields("local"))  '& codigo  'si no es traslado
    
        If local1.Visible = True Then
            mytablex.Fields("acu") = "T"
            mytablex.Fields("bodegaf") = Trim(xruc) '"" & mytable11.Fields("bodega")  'ojo si no esta vacio es traslado
            mytablex.Fields("tipoclie") = "V"

        End If

        If Trim("" & mytable11.Fields("terminal")) = "T" Then

            'mytablex.Fields("acu") = "I"
        End If
    
        mytablex.Fields("acu1") = ""

        'para traslado no debe existir nada
        If flag_servicio = "A" Then
            mytablex.Fields("servicio") = "A"

        End If

        If flag_servicio = "C" Then
            mytablex.Fields("servicio") = "C"

        End If

        If flag_servicio = "D" Then
            mytablex.Fields("servicio") = "D"

        End If

        mytablex.Fields("flage") = ""
        mytablex.Fields("codigo") = Trim("" & xruc)
        mytablex.Fields("caja") = "" & caja
        mytablex.Fields("turno") = "" & turno
        mytablex.Fields("usuario") = "" & cajero
        mytablex.Fields("fecha") = Format(dia, "dd/mm/yyyy")
        mytablex.Fields("hora") = Format(Now, "hh:MM")
        mytablex.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
        mytablex.Fields("estado") = "2"

        If local1.Visible = True Then
            mytablex.Fields("codigo") = Trim("" & "" & mytable11.Fields("local"))

        End If

        If xxacu = "I" Then
            mytablex.Fields("acu") = xxacu

        End If

        If Label36.Caption = "Almac.Fuente." Then
            mytablex.Fields("bodega") = Trim(xruc)
            mytablex.Fields("bodegaf") = Trim("" & mytable11.Fields("bodega"))

        End If

        If xxacu = "Q" Then
            mytablex.Fields("acu") = xxacu

        End If

        If local1 = "PEDIDO" Then
            mytablex.Fields("codigo") = ""

        End If

        If local1 = "GUIAREMISION" Then
            mytablex.Fields("ACU") = "T"

        End If

        If bxtipo = "7" Then
            mytablex.Fields("neto") = 0
            mytablex.Fields("descuento") = 0
            mytablex.Fields("subtotal") = 0
            mytablex.Fields("impuesto") = 0
            mytablex.Fields("total") = 0
            mytablex.Fields("xneto") = 0
            mytablex.Fields("tdetra") = 0
            mytablex.Fields("percepcion") = 0

        End If

        'ojo aqui debe estar primero creado el codigo
        'MsgBox ""
        'MsgBox acu
        If Len(codigo) > 0 And (bxtipo = "1" Or bxtipo = "2" Or bxtipo = "3" Or bxtipo = "4" Or bxtipo = "5" Or bxtipo = "7") Then
            found = crea_nuevos_clientes("" & codigo, mytablex.Fields("producto"), mytablex.Fields("precio"), mytablex.Fields("fecha"), mytablex.Fields("unidad"), mytablex.Fields("factor"), mytablex.Fields("descripcio"))

        End If

        'MsgBox Data2.Recordset.Fields("vendedor")
        'MsgBox "" & mytablex.Fields("vendedor")
        mytablex.Fields("flage") = "V"
    
        'aqui debe descargar
        If Len("" & mytablex.Fields("proveedorp")) > 0 Then
            If "" & mytablex.Fields("l3") = "P" Then
                found = graba_creditoProveedor(mytablex, indx)
                indx = indx + 1

            End If

            If "" & mytablex.Fields("l3") = "C" Then

                'found = graba_creditocliente(mytablex)
            End If

        End If
    
        If Trim("" & mytable11.Fields("terminal")) <> "T" And local1.Visible = False And local1 <> "PEDIDO" Then
            If crucefa.ListCount = 0 Then  'si no es facturacion mensual
                'MsgBox "abc"
                found = descarga_saldos(Trim("" & "" & mytable11.Fields("local")), mytablex, bxtipo, bxserie, bxnumero, -1, 0)

            End If

        End If

        mytablex.Update
        '-------------------- TERMINA LA GRABACION MYTABLEX
    
        'GRABANDO CLIENTES
        'MsgBox "Hola"
        If local1.Visible = True Then  'si es traslado
            mytableb.AddNew

            'MsgBox "Hola"
            For I = 0 To rs.Fields.count - 1
                mytableb.Fields(I) = rs.Fields(I)
            Next I

            'MsgBox "Hola"
            mytableb.Fields("local") = Trim("" & "" & mytable11.Fields("local"))
            mytableb.Fields("tipo") = "TS" '& bxtipo
            mytableb.Fields("serie") = "" & bxserie
            mytableb.Fields("numero") = "" & bxnumero

            If Len(Trim(xvendedor)) > 0 Then
                mytableb.Fields("vendedor") = xvendedor

            End If

            mytableb.Fields("moneda") = Trim("" & mytable11.Fields("moneda"))
            mytableb.Fields("bodega") = Trim("" & mytable11.Fields("bodega"))
            mytableb.Fields("bodegaf") = "" 'xruc '"" & mytable11.Fields("bodega")
            mytableb.Fields("acu") = "T"
            mytableb.Fields("localf") = Trim("" & "" & mytable11.Fields("local"))
            mytableb.Fields("tipoclie") = "V"

            If Trim("" & mytable11.Fields("terminal")) = "T" Then

                'mytablex.Fields("acu") = "I"
            End If

            mytableb.Fields("acu1") = ""

            'para traslado no debe existir nada
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
            mytableb.Fields("codigo") = Trim("" & "" & mytable11.Fields("local"))
            mytableb.Fields("caja") = "" & caja
            mytableb.Fields("turno") = "" & turno
            mytableb.Fields("usuario") = "" & cajero
            mytableb.Fields("fecha") = Format(dia, "dd/mm/yyyy")
            mytableb.Fields("hora") = Format(Now, "hh:MM")
            mytableb.Fields("fechacrea") = Format(Now, "dd/mm/yyyy")
            mytableb.Fields("estado") = "2"

            If Label36.Caption = "Almac.Fuente." Then
                mytableb.Fields("bodega") = Trim(xruc)
                mytableb.Fields("bodegaf") = ""

            End If

            'mytableb.Fields("local1") = "" & "" & mytable11.Fields("local")
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

            For I = 0 To rs.Fields.count - 1
                mytableb.Fields(I) = rs.Fields(I)
            Next I

            mytableb.Fields("local") = Trim("" & "" & mytable11.Fields("local"))
            mytableb.Fields("tipo") = "TE" '& bxtipo
            mytableb.Fields("serie") = "" & bxserie
            mytableb.Fields("numero") = "" & bxnumero

            If Len(Trim(xvendedor)) > 0 Then
                mytableb.Fields("vendedor") = xvendedor

            End If

            mytableb.Fields("moneda") = Trim("" & mytable11.Fields("moneda"))
            mytableb.Fields("bodega") = Trim(xruc)
            mytableb.Fields("bodegaf") = "" 'xruc '"" & mytable11.Fields("bodega")
            mytableb.Fields("acu") = "S"
            mytableb.Fields("localf") = Trim("" & "" & mytable11.Fields("local"))
            mytableb.Fields("tipoclie") = "V"

            If Trim("" & mytable11.Fields("terminal")) = "T" Then

                'mytablex.Fields("acu") = "I"
            End If

            mytableb.Fields("acu1") = ""

            'para traslado no debe existir nada
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
            mytableb.Fields("codigo") = Trim("" & "" & mytable11.Fields("local"))
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
    
            '-----AQUI SE TRASLADO EL DESCARGA SALDO-----
            '---------- DESCARGA SALDO
            'MsgBox "abc"
    
            If local1.Visible = True Then  'si es traslado
                found = descarga_saldos(Trim("" & "" & mytable11.Fields("local")), mytableb, "TS", bxserie, bxnumero, -1, 0)
                found = descarga_saldos(Trim("" & "" & mytable11.Fields("local")), mytableb, "TE", bxserie, bxnumero, 1, 0)

            End If
    
            mytableb.Update

            '-----------------------------------
        End If

        rs.MoveNext
    Loop
    'despues de grabar todos....
    'MsgBox "Hola"
    xbuf = xbuf & " Detalle:" & Format(Now, "hh:mm:ss")

    'AQUI DEBE DESCARGAR EL SALDO ACTUAL
    'If Trim("" & mytable11.Fields("terminal")) <> "T" And local1.Visible = False And local1 <> "PEDIDO" Then
    'If crucefa.ListCount = 0 Then  'si no es facturacion mensual
    '   found = descarga_saldo(Trim("" & "" & mytable11.Fields("local")), mytablex, bxtipo, bxserie, bxnumero, -1, 0)
    'End If
    'End If
    'If local1.Visible = True Then  'si es traslado
    '   found = descarga_saldo(Trim("" & "" & mytable11.Fields("local")), mytableb, "TS", bxserie, bxnumero, -1, 0)
    '   found = descarga_saldo(Trim("" & "" & mytable11.Fields("local")), mytableb, "TE", bxserie, bxnumero, 1, 0)
    'End If
    '-------------------------------------------------------------------------

    If local1.Visible = True Then
        mytableb.Close
        Set mytableb = Nothing

    End If

    '-----------------------------------------------------------------------

    mytablex.Close
    Set mytablex = Nothing
    'MsgBox ""
    xbuf = xbuf & " Saldo:" & Format(Now, "hh:mm:ss")

    If Trim("" & mytable11.Fields("terminal")) <> "T" Then  'finalizar el terminal
        'MsgBox "borrar"
        found = graba_guia_mensual() 'graba cuando es cruce de guias

    End If

    'found = busca_xtipog("" & bxtipo)  'graba el numero al actual
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

    If Trim("" & mytable11.Fields("hod")) = "S" Then  'enviar orden de despacho
        proceso_impresion11 "" & bxtipo, "" & bxserie, "" & bxnumero, 2, ""

    End If

    If Trim("" & mytable11.Fields("video")) = "S" Then
        If bxtipo = "7" Or Len(ndetraccion) > 0 Then

            'Frame10.Enabled = False
            'graba_video_concar bxserie & "-" & bxnumero
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
    'MsgBox correo
    envio_correos
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

Sub grabando_cabecera(mytablex As ADODB.Recordset, _
                      bxtipo As String, _
                      bxserie As String, _
                      bxnumero As String)

    On Error GoTo cmd232_err

    'MsgBox ""
    If Val(tdetra) > 0 Then
        mytablex.Fields("denumero") = Format(Val(ndetraccion), "0000000000")

    End If

    mytablex.Fields("tipoimp") = "C"
    'MsgBox ""
    mytablex.Fields("sentido") = sentido
    mytablex.Fields("observa") = xdistrito
    mytablex.Fields("tdetra") = Val(tdetra)
    mytablex.Fields("xneto") = Val(tpeaje)
    mytablex.Fields("tisc") = Val(tisc)
    mytablex.Fields("tivap") = Val(tivap)
    mytablex.Fields("tipo1") = petipo
    mytablex.Fields("serie1") = peserie
    mytablex.Fields("numero1") = penumero

    'MsgBox ""
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
    mytablex.Fields("codigo") = Trim(xruc)
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
    mytablex.Fields("total") = Val("" & txtotal)  'txtotal ya esta con percepcion
    mytablex.Fields("retotal") = Val("" & ytotal)  'total sin percepcion

    mytablex.Fields("percepcion") = Val("" & txpercepcion)

    mytablex.Fields("redondeo") = Val(Format(txtotlare, nrodecimal))
    mytablex.Fields("descuento") = Val("" & txdescuento)
    mytablex.Fields("neto") = Val("" & txneto)
    mytablex.Fields("impuesto") = Val("" & tximpuesto)
    mytablex.Fields("subtotal") = Val("" & txsubtotal)
    mytablex.Fields("SERVICIOCO") = Val(serviciocobro)
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
    mytablex.Fields("c1") = Val(c1)   'GRABA TOTAL DESCTO REFERENCIAL
    mytablex.Fields("c2") = Val(c2)
    mytablex.Fields("c3") = Val(c3)
    mytablex.Fields("c4") = Val(c4)
    mytablex.Fields("c5") = Val(c5)
    mytablex.Fields("c6") = Val(c6)
    mytablex.Fields("c7") = Val(c7)
    mytablex.Fields("c8") = Val(c8)
    mytablex.Fields("c9") = Val(c9)
    mytablex.Fields("local") = Trim("" & "" & mytable11.Fields("local"))
    mytablex.Fields("montopagar") = 0
    mytablex.Fields("ruc") = Trim("" & xruc)
    mytablex.Fields("TDOCDELI") = ""

    If flag_servicio = "A" Then
        mytablex.Fields("servicio") = "A"

    End If

    If flag_servicio = "D" Then
        mytablex.Fields("servicio") = "D"

    End If

    If flag_servicio = "C" Then
        mytablex.Fields("servicio") = "C"

    End If

    'MsgBox ""
    'validamos aqui si es traslado
    If local1.Visible = True Then
        mytablex.Fields("localf") = Trim("" & "" & mytable11.Fields("local"))
        mytablex.Fields("tipoclie") = "L"
        mytablex.Fields("bodegaf") = Trim(xruc)
        mytablex.Fields("codigo") = Trim("" & "" & mytable11.Fields("local"))

    End If

    If xxacu = "I" Then
        mytablex.Fields("acu") = xxacu

    End If

    If xxacu = "Q" Then
        mytablex.Fields("acu") = xxacu

    End If

    If local1 = "PEDIDO" Then
        mytablex.Fields("tipoclie") = "V"
        mytablex.Fields("CODIGO") = xvendedor
        mytablex.Fields("nombre") = "PEDIDO REPOSICION"
        mytablex.Fields("bodegaf") = Trim(xruc)

    End If

    If Label36.Caption = "Almac.Fuente." Then
        mytablex.Fields("bodega") = Trim(xruc)
        mytablex.Fields("bodegaf") = Trim("" & "" & mytable11.Fields("local"))

    End If

    If local1 = "GUIAREMISION" Then
        mytablex.Fields("ACU") = "T"

        'mytablex.Fields("nombre") = "PEDIDO REPOSICION"
    End If

    If bxtipo = "7" Then
        mytablex.Fields("neto") = 0
        mytablex.Fields("descuento") = 0
        mytablex.Fields("subtotal") = 0
        mytablex.Fields("impuesto") = 0
        mytablex.Fields("total") = 0
        mytablex.Fields("xneto") = 0
        mytablex.Fields("tdetra") = 0
        mytablex.Fields("percepcion") = 0

    End If

    'MsgBox "x"

    mytablex.Fields("flage") = "V"
    grabar_dato_pedido codigo, bxtipo, bxserie, bxnumero
    Exit Sub
cmd232_err:
    MsgBox "Error en grabando Cabecera " & error$, 48, "Aviso"
    Exit Sub

End Sub

Sub busca_correlativo(sw As Integer)

    Dim sdx      As Double

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

    Dim sdx      As Double

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

Function existe_fuel(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If "" & mytablex.Fields("fuel") = "S" Then
            existe_fuel = "" & mytablex.Fields("fueldonde")

        End If

    End If

    mytablex.Close

End Function

Function suma_pedidos(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    sdx = 0
    mytablex.Open "SELECT * FROM clientes where  codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        sdx = Val("" & mytablex.Fields("pedido")) - Val("" & mytablex.Fields("pedidoentregado"))

    End If

    mytablex.Close
    suma_pedidos = "" & sdx

End Function

Sub graba_acumulado_clientes(buf As String, signo As Double, sumador As Double)

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    mytablex.Open "SELECT * FROM clientes where  codigo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        sdx = Val("" & mytablex.Fields("pedidoentregado")) + signo * sumador
        mytablex.Fields("pedidoentregado") = sdx
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

Function credito_habilitado(buf As String)

    Dim mytablex As New ADODB.Recordset

    Dim buf1     As String

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "SELECT * FROM clientes where  codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        If "" & mytablex.Fields("estadocredito") = "S" Then
            credito_habilitado = 1

        End If

    End If

    mytablex.Close

End Function

Sub carga_dcvendedor()

    Dim mytablex As New ADODB.Recordset

    dcvendedor.Clear
    dcvendedor.AddItem "%"
    mytablex.Open "SELECT * FROM vendedor order by nombre", cn, adOpenDynamic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        dcvendedor.AddItem "" & mytablex.Fields("codigo") & "|" & mytablex.Fields("nombre")
        mytablex.MoveNext
    Loop
    mytablex.Close
    dcvendedor.ListIndex = 0

End Sub

Sub carga_local()
    'Dim mytablex As New ADODB.Recordset
    'selocal.Clear
    'selocal.AddItem "" & mytable11.Fields("local")
    'mytablex.Open "SELECT * FROM tlocal ", cn, adOpenDynamic, adLockOptimistic
    'Do
    'If mytablex.EOF Then Exit Do
    'If Trim("" & mytable11.Fields("local")) <> Trim("" & mytablex.Fields("codigo")) Then
    '   selocal.AddItem "" & mytablex.Fields("codigo")
    'End If
    'mytablex.MoveNext
    'Loop
    'mytablex.Close
    'selocal.ListIndex = 0

End Sub

Sub consulta_proveedor(buf As String)

    Dim found As Integer

    Frame1.Visible = True
    Frame1.Enabled = True
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.ListIndex = 0
    buffer = ""
    opcion1 = "800000"
    sw_consulta = 0

    If Len(Trim(buf)) > 0 Then
        found = sql_consulta(1)
        Exit Sub

    End If

    Set dbGrid1.DataSource = Nothing
    buffer.SetFocus

End Sub

Sub consulta_clientepase(buf As String)

    Dim found As Integer

    Frame1.Visible = True
    Frame1.Enabled = True
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.ListIndex = 0
    buffer = ""
    opcion1 = "900000"
    sw_consulta = 0

    If Len(Trim(buf)) > 0 Then
        found = sql_consulta(1)
        Exit Sub

    End If

    Set dbGrid1.DataSource = Nothing
    buffer.SetFocus

End Sub

Function graba_creditoProveedor(mytablez As ADODB.Recordset, indx As Integer)

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd6712121_err

    'MsgBox ""
    mytablex.Open "SELECT * FROM cuentap where local='" & Trim("" & mytablez.Fields("local")) & "' and tipo='" & Trim("" & mytablez.Fields("tipo")) & "' and serie='" & Trim("" & mytablez.Fields("serie")) & "' and numero='" & Trim("" & mytablez.Fields("numero")) & "-" & indx & "' and cuota='1'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si no existe
        mytablex.AddNew

        mytablex.Fields("grupo") = "C"
        'MsgBox ""
        mytablex.Fields("acu") = "" & acu
        mytablex.Fields("observa") = "" & Mid$("" & mytablez.Fields("descripcio"), 1, 30)
        mytablex.Fields("fpago") = ""
        mytablex.Fields("tipo") = "" & mytablez.Fields("tipo")
        mytablex.Fields("serie") = "" & mytablez.Fields("serie")
        mytablex.Fields("numero") = "" & mytablez.Fields("numero") & "-" & indx
        mytablex.Fields("numeron") = "" & mytablez.Fields("numero")
        mytablex.Fields("dias") = "1"
        mytablex.Fields("cuota") = "1"
        mytablex.Fields("tipoclie") = "P"
        mytablex.Fields("codigo") = "" & mytablez.Fields("proveedorp")
        mytablex.Fields("nombre") = "" & credito_proveedor("" & mytablez.Fields("proveedorp"))
        'MsgBox ""
        mytablex.Fields("fecha") = Format("" & mytablez.Fields("fecha"), "dd/mm/yyyy")
        mytablex.Fields("fechav") = Format("" & mytablez.Fields("fecha") + Val("" & mytablex.Fields("dias")), "dd/mm/yyyy")
        'MsgBox "1"
        mytablex.Fields("moneda") = "" & mytablez.Fields("l4") 'moneda del pase
        mytablex.Fields("total") = Val("" & mytablez.Fields("tcosto"))
        mytablex.Fields("abono") = 0
        mytablex.Fields("interes") = 0
        mytablex.Fields("saldo") = Val("" & mytablez.Fields("tcosto"))
        mytablex.Fields("estado") = "0"

        If Len(Trim(xvendedor)) > 0 Then
            mytablex.Fields("vendedor") = xvendedor

        End If

        mytablex.Fields("usuario") = cajero
        mytablex.Fields("caja") = caja
        mytablex.Fields("turno") = turno
        mytablex.Fields("zona") = ""
        mytablex.Fields("local") = Trim("" & mytablez.Fields("local"))
        mytablex.Update

    End If

    mytablex.Close
    Exit Function
cmd6712121_err:
    MsgBox "Aviso en Graba Credito Proveedor " + error$, 48, "Aviso"
    Exit Function

End Function

Function graba_creditocliente(mytablez As ADODB.Recordset)

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd126712121_err

    'MsgBox ""
    mytablex.Open "SELECT * FROM cuentaC where local='" & Trim("" & mytablez.Fields("local")) & "' and tipo='" & Trim("" & mytablez.Fields("tipo")) & "' and serie='" & Trim("" & mytablez.Fields("serie")) & "' and numero='" & Trim("" & mytablez.Fields("numero")) & "' and cuota='1'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si no existe
        mytablex.AddNew
        mytablex.Fields("grupo") = "C"
        'MsgBox ""
        mytablex.Fields("acu") = "" & acu
        mytablex.Fields("observa") = "" & Mid$("" & mytablez.Fields("descripcio"), 1, 30)
        mytablex.Fields("fpago") = ""
        mytablex.Fields("tipo") = "" & mytablez.Fields("tipo")
        mytablex.Fields("serie") = "" & mytablez.Fields("serie")
        mytablex.Fields("numero") = "" & mytablez.Fields("numero")
        mytablex.Fields("dias") = "1"
        mytablex.Fields("cuota") = "1"
        mytablex.Fields("tipoclie") = "C"
        mytablex.Fields("codigo") = "" & mytablez.Fields("proveedorp")
        mytablex.Fields("nombre") = "" & credito_proveedor("" & mytablez.Fields("proveedorp"))
        'MsgBox ""
        mytablex.Fields("fecha") = Format("" & mytablez.Fields("fecha"), "dd/mm/yyyy")
        mytablex.Fields("fechav") = Format("" & mytablez.Fields("fecha") + Val("" & mytablex.Fields("dias")), "dd/mm/yyyy")
        'MsgBox "1"
        mytablex.Fields("moneda") = "" & mytablez.Fields("l4") 'moneda del pase
        mytablex.Fields("total") = Val("" & mytablez.Fields("tcosto"))
        mytablex.Fields("abono") = 0
        mytablex.Fields("interes") = 0
        mytablex.Fields("saldo") = Val("" & mytablez.Fields("tcosto"))
        mytablex.Fields("estado") = "0"

        If Len(Trim(xvendedor)) > 0 Then
            mytablex.Fields("vendedor") = xvendedor

        End If

        mytablex.Fields("usuario") = cajero
        mytablex.Fields("caja") = caja
        mytablex.Fields("turno") = turno
        mytablex.Fields("zona") = ""
        mytablex.Fields("local") = Trim("" & mytablez.Fields("local"))
        mytablex.Update

    End If

    mytablex.Close
    Exit Function
cmd126712121_err:
    MsgBox "Aviso en Graba Credito Cliente " + error$, 48, "Aviso"
    Exit Function

End Function

Function credito_proveedor(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM proveedo where codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        credito_proveedor = "" & mytablex.Fields("nombre")

    End If

    mytablex.Close

End Function

Function credito_proveedors(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    credito_proveedors = "S"
    mytablex.Open "SELECT * FROM proveedo where codigo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        credito_proveedors = "" & mytablex.Fields("moneda")

        If "" & mytablex.Fields("moneda") <> "S" And "" & mytablex.Fields("moneda") <> "D" Then
            credito_proveedors = "S"

        End If

    End If

    mytablex.Close

End Function

Sub carga_xxx()
    DBGrid2.columns("nroprecio") = "1"
    DBGrid2.columns("hora") = Format(Now, "hh:mm:ss")
    DBGrid2.columns("categoria") = ""
    DBGrid2.columns("producto") = "XXX"
    DBGrid2.columns("proveedorp") = "" '& mytablex.Fields("proveedor1")
    DBGrid2.columns("tipo") = ""
    DBGrid2.columns("serie") = ""
    DBGrid2.columns("numero") = ""
    DBGrid2.columns("isc") = ""  '& mytablex.Fields("vendedor")
    DBGrid2.columns("comision") = 0
    DBGrid2.columns("descripcio") = "X"
    'MsgBox xxca
    DBGrid2.columns("cantidad") = "1"
    'dbgrid2.columns("descuento") = Val("" & mytablex.Fields("isc"))

    DBGrid2.columns("unidad") = "UND"
    DBGrid2.columns("factor") = 1
    DBGrid2.columns("precio") = 1
    DBGrid2.columns("precio") = 1
    DBGrid2.columns("precio") = 1

    'dbgrid2.columns("neto") = Val("" & mytablex.Fields("tax"))
    'dbgrid2.columns("unidad") = "" & mytabley.Fields("unidad1")
    DBGrid2.columns("precio") = 1
    DBGrid2.columns("total") = 1
    DBGrid2.columns("subtotal") = 1

    DBGrid2.columns("deslipo") = 0
    DBGrid2.columns("tax") = 0
    DBGrid2.columns("vendedor") = "" 'Val("" & mytablex.Fields("isc"))
    DBGrid2.columns("impuesto") = 0
    DBGrid2.columns("igv") = 18
    DBGrid2.columns("linea") = ""

    DBGrid2.columns("descuento") = 0
    DBGrid2.columns("neto") = 0

    DBGrid2.columns("tcosto") = 0
    DBGrid2.columns("familia") = "XXX"
    DBGrid2.columns("subfamilia") = ""
    DBGrid2.columns("marca") = ""
    DBGrid2.columns("total") = Val(DBGrid2.columns("cantidad")) * Val(DBGrid2.columns("precio"))
    DBGrid2.columns("ivap") = 0
    DBGrid2.columns("isc") = 0
    DBGrid2.columns("isc") = 0
    calcula_igv 0

End Sub

Sub pone_tallas_saldo(mytabley As ADODB.Recordset, _
                      mytablex As ADODB.Recordset, _
                      sw As Integer)
    'MsgBox "abc"
    mytabley.Fields("t1") = Val("" & mytabley.Fields("t1")) + sw * mytablex.Fields("t1")
    mytabley.Fields("t2") = Val("" & mytabley.Fields("t2")) + sw * mytablex.Fields("t2")
    mytabley.Fields("t3") = Val("" & mytabley.Fields("t3")) + sw * mytablex.Fields("t3")
    mytabley.Fields("t4") = Val("" & mytabley.Fields("t4")) + sw * mytablex.Fields("t4")
    mytabley.Fields("t5") = Val("" & mytabley.Fields("t5")) + sw * mytablex.Fields("t5")
    mytabley.Fields("t6") = Val("" & mytabley.Fields("t6")) + sw * mytablex.Fields("t6")
    mytabley.Fields("t7") = Val("" & mytabley.Fields("t7")) + sw * mytablex.Fields("t7")
    mytabley.Fields("t8") = Val("" & mytabley.Fields("t8")) + sw * mytablex.Fields("t8")
    mytabley.Fields("t9") = Val("" & mytabley.Fields("t9")) + sw * mytablex.Fields("t9")
    mytabley.Fields("t10") = Val("" & mytabley.Fields("t10")) + sw * mytablex.Fields("t10")
    mytabley.Fields("t11") = Val("" & mytabley.Fields("t11")) + sw * mytablex.Fields("t11")
    mytabley.Fields("t12") = Val("" & mytabley.Fields("t12")) + sw * mytablex.Fields("t12")
    mytabley.Fields("t13") = Val("" & mytabley.Fields("t13")) + sw * mytablex.Fields("t13")
    mytabley.Fields("t14") = Val("" & mytabley.Fields("t14")) + sw * mytablex.Fields("t14")
    mytabley.Fields("t15") = Val("" & mytabley.Fields("t15")) + sw * mytablex.Fields("t15")

End Sub

Function tiene_percepcion(buf As String) As String

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM producto where producto='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If "" & mytablex.RecordCount > 0 Then
        If "" & mytablex.Fields("percepcion") = "S" Then
            tiene_percepcion = "S"

        End If

    End If

    mytablex.Close

End Function

Function acura_lectura() As String

    Dim D   As Integer

    Dim I   As Integer

    Dim buf As String

    acura_lectura = acurabuffer
    Exit Function

    Select Case "" & mytable11.Fields("portbala")

        Case "COM1"
            D = 1

        Case "COM2"
            D = 2

        Case "COM3"
            D = 3

        Case "COM4"
            D = 4

        Case "COM5"
            D = 5

    End Select

    MSComm1.CommPort = D
    MSComm1.Settings = "9600,n,8,1"
    MSComm1.InputLen = 10
    MSComm1.PortOpen = True
    buf = ""
    Do
        DoEvents
        buf = buf & MSComm1.input
    Loop Until Len(buf) >= 10

    MSComm1.PortOpen = False
    acura_lectura = Mid$(buf, Len(buf) - 7, 6)

End Function

Sub resuma_precios(xpercepcion As Double)

    Dim xtivap      As Double

    Dim tdscto      As Double

    Dim sdx2        As Double

    Dim sdx1        As Double

    Dim xtisc       As Double

    Dim X           As Double

    Dim Y           As Double

    Dim sdx         As Double

    Dim ypercepcion As Double

    Dim xneto       As Double

    On Error GoTo cmd94534_err

    ypercepcion = 0

    Data2.Recordset.Fields("percepcion") = xpercepcion

    If busca_tipoprecio() = "N" Then
        Data2.Recordset.Fields("neto") = Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & Data2.Recordset.Fields("precio"))
        Data2.Recordset.Fields("descuento") = Val("" & Data2.Recordset.Fields("neto")) * Val("" & Data2.Recordset.Fields("deslipo")) / 100 + Val("" & Data2.Recordset.Fields("neto")) * Val("" & Data2.Recordset.Fields("destopo")) / 100     'calcular descuento
        Data2.Recordset.Fields("subtotal") = Val("" & Data2.Recordset.Fields("neto")) - Val("" & Data2.Recordset.Fields("descuento")) 'cobrar
        Data2.Recordset.Fields("impuesto") = Val("" & Data2.Recordset.Fields("subtotal")) * Val("" & Data2.Recordset.Fields("igv")) / 100  'calcular descuento
        Data2.Recordset.Fields("total") = Val("" & Data2.Recordset.Fields("subtotal")) + Val("" & Data2.Recordset.Fields("impuesto")) 'cobrar
        Data2.Recordset.Fields("tivap") = Val("" & Data2.Recordset.Fields("total")) * Val("" & Data2.Recordset.Fields("ivap")) / 100
        Data2.Recordset.Fields("tpercepcio") = 0

        If "" & Data2.Recordset.Fields("l1") = "S" Then
            Data2.Recordset.Fields("tpercepcio") = Val("" & Data2.Recordset.Fields("total")) * Val("" & Data2.Recordset.Fields("percepcion")) / 100    'calcular descuento
            Data2.Recordset.Fields("total") = Val("" & Data2.Recordset.Fields("total")) + Val("" & Data2.Recordset.Fields("tpercepcio")) 'cobrar

        End If

        Data2.Recordset.Fields("servicioco") = Val("" & Data2.Recordset.Fields("subtotal")) * Val("" & Data2.Recordset.Fields("serviciopo")) / 100     'calcular descuento
        Data2.Recordset.Fields("total") = Val("" & Data2.Recordset.Fields("total")) + Val("" & Data2.Recordset.Fields("servicioco")) 'cobrar

    Else
        Data2.Recordset.Fields("neto") = Val("" & Data2.Recordset.Fields("cantidad")) * Val("" & Data2.Recordset.Fields("precio"))
        Data2.Recordset.Fields("descuento") = Val("" & Data2.Recordset.Fields("neto")) * Val("" & Data2.Recordset.Fields("deslipo")) / 100 + Val("" & Data2.Recordset.Fields("neto")) * Val("" & Data2.Recordset.Fields("destopo")) / 100
        Data2.Recordset.Fields("total") = Val("" & Data2.Recordset.Fields("neto")) - Val("" & Data2.Recordset.Fields("descuento")) 'cobrar
        Data2.Recordset.Fields("subtotal") = Val("" & Data2.Recordset.Fields("total")) / (1 + Val("" & Data2.Recordset.Fields("igv")) / 100) 'calcular descuento
        Data2.Recordset.Fields("impuesto") = Val("" & Data2.Recordset.Fields("total")) - Val("" & Data2.Recordset.Fields("subtotal")) 'cobrar
        xtivap = Val("" & Data2.Recordset.Fields("total")) * Val("" & Data2.Recordset.Fields("ivap")) / 100
        Data2.Recordset.Fields("tivap") = xtivap
        Data2.Recordset.Fields("tpercepcio") = 0

        If "" & Data2.Recordset.Fields("l1") = "S" Then
            Data2.Recordset.Fields("tpercepcio") = Val("" & Data2.Recordset.Fields("total")) * Val("" & Data2.Recordset.Fields("percepcion")) / 100   'calcular descuento
            Data2.Recordset.Fields("total") = Val("" & Data2.Recordset.Fields("total")) + Val("" & Data2.Recordset.Fields("tpercepcio")) 'cobrar

        End If

        Data2.Recordset.Fields("servicioco") = Val("" & Data2.Recordset.Fields("subtotal")) * Val("" & Data2.Recordset.Fields("serviciopo")) / 100      'calcular descuento

    End If

    Exit Sub
cmd94534_err:
    MsgBox "Aviso en resuma_precios ", 48, "Aviso"
    Exit Sub

End Sub

Sub carga_clase_sunat()

    Dim mytablex As New ADODB.Recordset

    clasesunat.Clear
    clasesunat.AddItem ""
    mytablex.Open "SELECT * FROM clasesunat ", cn, adOpenDynamic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        clasesunat.AddItem Trim("" & mytablex.Fields("clasesunat")) & "|" & Trim("" & mytablex.Fields("descripcio"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    clasesunat.ListIndex = 0
         
End Sub

Function nro_copias(buf As String) As Integer

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM tipo where  tipo='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then 'si existe
        nro_copias = Val("" & mytablex.Fields("copias"))

        If Val("" & mytablex.Fields("copias")) <= 0 Then
            nro_copias = 1

        End If

    End If

    mytablex.Close

End Function

Sub envio_correos()

    Dim txtserver     As String

    Dim txtusername   As String

    Dim txtpassword   As String

    Dim txtport       As String

    Dim txtto         As String

    Dim chkssl        As String

    Dim txtfromname   As String

    Dim txtfromemail  As String

    Dim txtattach     As String

    Dim txtsubject    As String

    Dim txtmsg        As String

    Dim retval        As String

    Dim txtselecciona As String

    Dim txthtml       As String

    Dim mytablex      As New ADODB.Recordset

    Dim buf           As String

    On Error GoTo cmd0905677_err

    buf = Trim("" & mytable11.Fields("correo"))

    If Trim(buf) = 0 Then Exit Sub
    mytablex.Open "select * from correos where cosms='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        txtselecciona = Trim("" & mytablex.Fields("txtselecciona"))
        txtserver = Trim("" & mytablex.Fields("txtserver"))
        txthtml = Trim("" & mytablex.Fields("txthtml"))
        txtusername = Trim("" & mytablex.Fields("txtusername"))
        txtpassword = Trim("" & mytablex.Fields("txtpassword"))
        txtfromname = Trim("" & mytablex.Fields("txtfromname"))
        txtfromemail = Trim("" & mytablex.Fields("txtfromemail"))

        txtport = Trim("" & mytablex.Fields("txtport"))
        'txtto = Trim("" & mytablex.Fields("txtto"))
        chkssl = Trim("" & mytablex.Fields("chkssl"))
        'txtfromname = Trim("" & nombre) 'Trim("" & mytablex.Fields("txtfromname"))
        txtto = Trim("" & correo) 'Trim("" & mytablex.Fields("txtfromemail"))
        txtattach = Trim("" & mytablex.Fields("txtattach"))
        txtsubject = Trim("" & mytablex.Fields("txtsubject"))
        txtmsg = Trim("" & mytablex.Fields("txtmsg"))
        txtmsg = txtmsg & Chr$(10) & Chr$(13) & ""
        txtmsg = txtmsg & Format(Now, "dd/mm/yyyy") + " " + Format(Now, "hh:mm:ss")
        retval = SendMail(Trim$(txtto), Trim$(txtsubject), Trim$(txtfromname) & "<" & Trim$(txtfromemail) & ">", Trim$(txtmsg), Trim$(txtserver), CInt(Trim$(txtport)), Trim$(txtusername), Trim$(txtpassword), Trim$(txtattach), True, txtselecciona, txthtml)

        'MsgBox retval
    End If

    mytablex.Close
    Exit Sub
cmd0905677_err:
    MsgBox "No se Pudo enviar Correo... " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub icerrar_puertosmscomm()

    Dim I As Integer

    For I = 1 To 10
        icerrando_mscomm I
    Next I

End Sub

Sub icerrando_mscomm(D As Integer)

    On Error GoTo cmdini4_err

    MSComm1.CommPort = D

    If MSComm1.PortOpen = True Then
        MSComm1.PortOpen = False

    End If

    Exit Sub
cmdini4_err:
    Exit Sub

End Sub

Sub lectura_grafico(buf As String)

    On Error GoTo cmd909090_err

    Dim mytablex   As New ADODB.Recordset

    Dim fotonombre As String

    foto.Picture = LoadPicture()
    mytablex.Open "select * from producto where producto='" & buf & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        fotonombre = globalpath & "\001d\06\grafico\tmp.jpg"
        viewBMP mytablex, fotonombre

        If Len(Trim(fotonombre)) > 0 Then
            If existe_archivo(fotonombre) > 0 Then
                redimensionar_grafico fotonombre
                foto.Picture = LoadPicture(fotonombre)

            End If

        End If

    End If

    mytablex.Close
    Exit Sub
cmd909090_err:
    MsgBox "Aviso en lectura grafico " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub redimensionar_grafico(fotonombre As String)

    On Error GoTo cmd90111_err

    Picture1.AutoRedraw = True
    Picture1.PaintPicture LoadPicture(fotonombre), 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, , , , , vbSrcCopy
    SavePicture Picture1.Image, fotonombre
    Exit Sub
cmd90111_err:
    Exit Sub

End Sub

