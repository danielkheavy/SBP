VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tcaja 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametros Generales Caja"
   ClientHeight    =   9795
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleMode       =   0  'User
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox formatocomanda 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   248
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox digitos 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   247
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox coladelivery 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9600
      MaxLength       =   1
      TabIndex        =   246
      Top             =   9240
      Width           =   495
   End
   Begin VB.TextBox puertodelivery 
      BackColor       =   &H00C0FFFF&
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
      Left            =   7920
      TabIndex        =   244
      Top             =   9240
      Width           =   1695
   End
   Begin VB.TextBox multicomanda 
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
      Left            =   8280
      MaxLength       =   1
      TabIndex        =   243
      Top             =   800
      Width           =   525
   End
   Begin VB.TextBox Fnc 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12000
      MaxLength       =   1
      TabIndex        =   241
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox cnc 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12480
      MaxLength       =   1
      TabIndex        =   240
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox archivonc 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9720
      MaxLength       =   30
      TabIndex        =   239
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox inc 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11520
      MaxLength       =   1
      TabIndex        =   238
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox puertonc 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5520
      MaxLength       =   60
      TabIndex        =   237
      Top             =   5880
      Width           =   4215
   End
   Begin VB.TextBox colanc 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   236
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox numeronc 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3720
      MaxLength       =   11
      TabIndex        =   235
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox serienc 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   4
      TabIndex        =   234
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox tiponc 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   3
      TabIndex        =   233
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Terminal Touch"
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
      Left            =   -480
      TabIndex        =   218
      Top             =   9720
      Visible         =   0   'False
      Width           =   12615
      Begin VB.TextBox anulaventas 
         Height          =   375
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   225
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox cierrecaja 
         Height          =   375
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   224
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox pmesero 
         Height          =   375
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   223
         Top             =   4200
         Width           =   495
      End
      Begin VB.TextBox ingresodinero 
         Height          =   375
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   222
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox egresodinero 
         Height          =   375
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   221
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox descuento 
         Height          =   375
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   220
         Top             =   3840
         Width           =   495
      End
      Begin VB.TextBox grabacomanda 
         Height          =   375
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   219
         Top             =   4560
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Impresoras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   960
      TabIndex        =   215
      Top             =   9720
      Visible         =   0   'False
      Width           =   11415
      Begin VB.CommandButton Command3 
         Caption         =   "Copiar"
         Height          =   615
         Left            =   8640
         TabIndex        =   217
         Top             =   330
         Width           =   975
      End
      Begin VB.ComboBox cboprinters 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   216
         Top             =   315
         Width           =   8415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9615
      Left            =   10080
      TabIndex        =   210
      Top             =   10800
      Visible         =   0   'False
      Width           =   14415
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
         Left            =   8280
         TabIndex        =   213
         TabStop         =   0   'False
         Top             =   600
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
         TabIndex        =   212
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
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
         TabIndex        =   211
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6735
         Left            =   240
         TabIndex        =   214
         Top             =   1200
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   11880
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
   Begin VB.TextBox letrainterna 
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
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   208
      Text            =   "10"
      Top             =   1560
      Width           =   570
   End
   Begin VB.TextBox bold 
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
      Left            =   6045
      MaxLength       =   1
      TabIndex        =   207
      Top             =   1155
      Width           =   495
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4275
      Style           =   2  'Dropdown List
      TabIndex        =   206
      Top             =   1200
      Width           =   1800
   End
   Begin VB.TextBox clavedescongela 
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
      Left            =   12720
      MaxLength       =   1
      TabIndex        =   205
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox clavecongela 
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
      Left            =   12480
      MaxLength       =   1
      TabIndex        =   204
      Top             =   0
      Width           =   255
   End
   Begin VB.TextBox correo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   10360
      MaxLength       =   6
      TabIndex        =   202
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox gavetasw 
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
      Left            =   9960
      MaxLength       =   1
      TabIndex        =   200
      Top             =   6720
      Width           =   255
   End
   Begin VB.TextBox tipo_balanza 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   199
      Top             =   7080
      Width           =   615
   End
   Begin VB.TextBox flag1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11205
      MaxLength       =   1
      TabIndex        =   198
      Top             =   7335
      Width           =   375
   End
   Begin VB.TextBox stkminimo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   6075
      MaxLength       =   1
      TabIndex        =   197
      Top             =   720
      Width           =   450
   End
   Begin VB.TextBox clienteo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   196
      Top             =   7800
      Width           =   375
   End
   Begin VB.TextBox vdetalle 
      BackColor       =   &H00E0E0E0&
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
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   195
      Top             =   7440
      Width           =   375
   End
   Begin VB.TextBox obligaprecio 
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
      Left            =   8280
      MaxLength       =   1
      TabIndex        =   193
      Top             =   20
      Width           =   525
   End
   Begin VB.TextBox copiaod 
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
      Left            =   12480
      MaxLength       =   30
      TabIndex        =   191
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox obligacredito 
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
      Left            =   12480
      MaxLength       =   1
      TabIndex        =   190
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox obligavendedor 
      BackColor       =   &H00E0E0E0&
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
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   189
      Top             =   7440
      Width           =   375
   End
   Begin VB.TextBox gavetacola 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   186
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox limpiapantalla 
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
      Left            =   12480
      MaxLength       =   1
      TabIndex        =   185
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox tamanorden 
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
      TabIndex        =   184
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox nombrefont 
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
      Left            =   3000
      MaxLength       =   30
      TabIndex        =   182
      Top             =   1200
      Width           =   1290
   End
   Begin VB.TextBox salon 
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
      Left            =   10365
      MaxLength       =   2
      TabIndex        =   181
      Top             =   1530
      Width           =   615
   End
   Begin VB.TextBox tamanoletra 
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
      TabIndex        =   179
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox puerto 
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
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   178
      Top             =   11160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox archivoe 
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
      Left            =   120
      MaxLength       =   10
      TabIndex        =   177
      Top             =   11160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox odcola 
      BackColor       =   &H00E0E0E0&
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
      Left            =   10650
      MaxLength       =   1
      TabIndex        =   175
      Top             =   7785
      Width           =   315
   End
   Begin VB.TextBox ordenaproducto 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   1
      TabIndex        =   173
      Top             =   9240
      Width           =   375
   End
   Begin VB.TextBox habilitanota 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   172
      Top             =   8520
      Width           =   615
   End
   Begin VB.TextBox clavecopia 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   1
      TabIndex        =   171
      Top             =   8880
      Width           =   615
   End
   Begin VB.TextBox apertura 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12000
      MaxLength       =   1
      TabIndex        =   169
      Top             =   8160
      Width           =   375
   End
   Begin VB.TextBox vecocaja 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   167
      Top             =   8880
      Width           =   615
   End
   Begin VB.TextBox colacie 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9600
      MaxLength       =   1
      TabIndex        =   166
      Top             =   8520
      Width           =   495
   End
   Begin VB.TextBox deshab 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12000
      MaxLength       =   1
      TabIndex        =   165
      Top             =   8520
      Width           =   375
   End
   Begin VB.TextBox hdetraccio 
      BackColor       =   &H00FFFF80&
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
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   164
      Top             =   10800
      Width           =   375
   End
   Begin VB.TextBox archivoexo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9720
      MaxLength       =   30
      TabIndex        =   163
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox puertoexo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5520
      MaxLength       =   60
      TabIndex        =   162
      Top             =   5520
      Width           =   4215
   End
   Begin VB.TextBox colaexo 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   161
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox tipoexo 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   3
      TabIndex        =   160
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox serieexo 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   4
      TabIndex        =   159
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox numeroexo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3720
      MaxLength       =   11
      TabIndex        =   158
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox iexo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11520
      MaxLength       =   1
      TabIndex        =   157
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox fexo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12000
      MaxLength       =   1
      TabIndex        =   156
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox cexo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12480
      MaxLength       =   1
      TabIndex        =   155
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox segundo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12000
      MaxLength       =   2
      TabIndex        =   154
      Top             =   8880
      Width           =   375
   End
   Begin VB.TextBox video 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12360
      MaxLength       =   1
      TabIndex        =   153
      Top             =   8880
      Width           =   375
   End
   Begin VB.TextBox sentido 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   1
      TabIndex        =   152
      Top             =   9240
      Width           =   375
   End
   Begin VB.TextBox parqueo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12600
      MaxLength       =   1
      TabIndex        =   151
      Top             =   9240
      Width           =   375
   End
   Begin VB.TextBox pm 
      BackColor       =   &H00E0E0E0&
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
      Left            =   6075
      MaxLength       =   1
      TabIndex        =   149
      Top             =   360
      Width           =   450
   End
   Begin VB.TextBox decimal1 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   148
      Top             =   7080
      Width           =   375
   End
   Begin VB.TextBox tipocie 
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
      Left            =   11880
      MaxLength       =   1
      TabIndex        =   145
      Top             =   7320
      Width           =   375
   End
   Begin VB.TextBox puertocie 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7920
      MaxLength       =   60
      TabIndex        =   144
      Top             =   8520
      Width           =   1695
   End
   Begin VB.TextBox cierres 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   142
      Top             =   9240
      Width           =   1575
   End
   Begin VB.TextBox seriehd 
      BackColor       =   &H00E0E0E0&
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
      Left            =   13200
      MaxLength       =   60
      TabIndex        =   141
      Top             =   9240
      Width           =   255
   End
   Begin VB.TextBox t0 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12360
      MaxLength       =   1
      TabIndex        =   140
      Top             =   6360
      Width           =   615
   End
   Begin VB.TextBox t5 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12360
      MaxLength       =   2
      TabIndex        =   138
      Top             =   8160
      Width           =   615
   End
   Begin VB.TextBox t4 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12360
      MaxLength       =   2
      TabIndex        =   137
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox t3 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12360
      MaxLength       =   2
      TabIndex        =   136
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox t2 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12360
      MaxLength       =   2
      TabIndex        =   135
      Top             =   7080
      Width           =   615
   End
   Begin VB.TextBox T1 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12360
      MaxLength       =   2
      TabIndex        =   134
      Top             =   6720
      Width           =   615
   End
   Begin VB.TextBox repite 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11040
      MaxLength       =   1
      TabIndex        =   132
      Top             =   9240
      Width           =   375
   End
   Begin VB.TextBox ecpuerto 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7920
      MaxLength       =   60
      TabIndex        =   130
      Top             =   8160
      Width           =   1695
   End
   Begin VB.TextBox eccola 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9600
      MaxLength       =   1
      TabIndex        =   129
      Top             =   8160
      Width           =   495
   End
   Begin VB.TextBox hod 
      BackColor       =   &H00E0E0E0&
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
      Left            =   10350
      MaxLength       =   30
      TabIndex        =   127
      Top             =   7800
      Width           =   300
   End
   Begin VB.TextBox odpuerto 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7920
      MaxLength       =   60
      TabIndex        =   125
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox nosaldo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   124
      Top             =   8160
      Width           =   615
   End
   Begin VB.TextBox listap 
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
      Left            =   7920
      MaxLength       =   2
      TabIndex        =   123
      Top             =   7440
      Width           =   1215
   End
   Begin VB.TextBox ctb 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12480
      MaxLength       =   1
      TabIndex        =   122
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox ctf 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12480
      MaxLength       =   1
      TabIndex        =   121
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox cnv 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12480
      MaxLength       =   1
      TabIndex        =   120
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox cbm 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12480
      MaxLength       =   1
      TabIndex        =   119
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox cfm 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12480
      MaxLength       =   1
      TabIndex        =   118
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox cot 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12480
      MaxLength       =   1
      TabIndex        =   117
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox cpro 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12480
      MaxLength       =   1
      TabIndex        =   116
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox creg 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12480
      MaxLength       =   1
      TabIndex        =   115
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox crin 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12480
      MaxLength       =   1
      TabIndex        =   114
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox local1 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   3
      TabIndex        =   112
      Top             =   6720
      Width           =   615
   End
   Begin VB.TextBox redondeo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7920
      MaxLength       =   1
      TabIndex        =   110
      Top             =   7080
      Width           =   375
   End
   Begin VB.TextBox noprecio 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   109
      Top             =   8880
      Width           =   375
   End
   Begin VB.TextBox terminal 
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
      Left            =   10365
      MaxLength       =   1
      TabIndex        =   107
      Top             =   1185
      Width           =   615
   End
   Begin VB.TextBox FRIN 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12000
      MaxLength       =   1
      TabIndex        =   105
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox FREG 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12000
      MaxLength       =   1
      TabIndex        =   104
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox FPRO 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12000
      MaxLength       =   1
      TabIndex        =   103
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox FOT 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12000
      MaxLength       =   1
      TabIndex        =   102
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox FFM 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12000
      MaxLength       =   1
      TabIndex        =   101
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox FBM 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12000
      MaxLength       =   1
      TabIndex        =   100
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox FNV 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12000
      MaxLength       =   1
      TabIndex        =   99
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox FTF 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12000
      MaxLength       =   1
      TabIndex        =   98
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox FTB 
      BackColor       =   &H00E0E0E0&
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
      Left            =   12000
      MaxLength       =   1
      TabIndex        =   97
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox itb 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11520
      MaxLength       =   1
      TabIndex        =   95
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox itf 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11520
      MaxLength       =   1
      TabIndex        =   94
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox inv 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11520
      MaxLength       =   1
      TabIndex        =   93
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox ibm 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11520
      MaxLength       =   1
      TabIndex        =   92
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox ifm 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11520
      MaxLength       =   1
      TabIndex        =   91
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox iot 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11520
      MaxLength       =   1
      TabIndex        =   90
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox ipe 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11520
      MaxLength       =   1
      TabIndex        =   89
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox ire 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11520
      MaxLength       =   1
      TabIndex        =   88
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox iri 
      BackColor       =   &H00E0E0E0&
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
      Left            =   11520
      MaxLength       =   1
      TabIndex        =   87
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox siventa 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   86
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox capuerto 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7920
      MaxLength       =   30
      TabIndex        =   84
      Top             =   6360
      Width           =   2295
   End
   Begin VB.TextBox catipo 
      BackColor       =   &H00E0E0E0&
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
      Left            =   7920
      MaxLength       =   1
      TabIndex        =   83
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox serieti 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   30
      TabIndex        =   82
      Top             =   8160
      Width           =   1455
   End
   Begin VB.TextBox flag 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   80
      Top             =   7800
      Width           =   615
   End
   Begin VB.TextBox puertori 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5505
      MaxLength       =   60
      TabIndex        =   78
      Top             =   4800
      Width           =   4215
   End
   Begin VB.TextBox colari 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   77
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox tipori 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   3
      TabIndex        =   76
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox serieri 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   4
      TabIndex        =   75
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox numerori 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3720
      MaxLength       =   11
      TabIndex        =   74
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox numerore 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3720
      MaxLength       =   11
      TabIndex        =   73
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox seriere 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   4
      TabIndex        =   72
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox tipore 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   3
      TabIndex        =   71
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox colare 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   70
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox puertore 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5520
      MaxLength       =   60
      TabIndex        =   69
      Top             =   5160
      Width           =   4215
   End
   Begin VB.TextBox archivore 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9720
      MaxLength       =   30
      TabIndex        =   68
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox archivori 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9720
      MaxLength       =   30
      TabIndex        =   67
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox actbala 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   65
      Top             =   6720
      Width           =   615
   End
   Begin VB.TextBox portbala 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   64
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox congela 
      BackColor       =   &H00E0E0E0&
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
      Left            =   4680
      MaxLength       =   11
      TabIndex        =   62
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox cliente 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   1
      TabIndex        =   60
      Top             =   7800
      Width           =   375
   End
   Begin VB.TextBox vendedor 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   1
      TabIndex        =   59
      Top             =   7440
      Width           =   375
   End
   Begin VB.TextBox bodega 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   58
      Top             =   7080
      Width           =   735
   End
   Begin VB.TextBox archivotb 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9720
      MaxLength       =   30
      TabIndex        =   56
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox archivotf 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9720
      MaxLength       =   30
      TabIndex        =   55
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox archivonv 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9720
      MaxLength       =   30
      TabIndex        =   54
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox archivobm 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9720
      MaxLength       =   30
      TabIndex        =   53
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox archivofm 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9720
      MaxLength       =   30
      TabIndex        =   52
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox archivoot 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9720
      MaxLength       =   30
      TabIndex        =   51
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox archivope 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9720
      MaxLength       =   30
      TabIndex        =   50
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox tipodefa 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   3
      TabIndex        =   48
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox puertope 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5520
      MaxLength       =   60
      TabIndex        =   47
      Top             =   4440
      Width           =   4215
   End
   Begin VB.TextBox colape 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   46
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox tipope 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   3
      TabIndex        =   45
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox seriepe 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   4
      TabIndex        =   44
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox numerope 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3720
      MaxLength       =   11
      TabIndex        =   43
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox numeroot 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3720
      MaxLength       =   11
      TabIndex        =   41
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox serieot 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   4
      TabIndex        =   40
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox tipoot 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   3
      TabIndex        =   39
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox colaot 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   38
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox puertoot 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5520
      MaxLength       =   60
      TabIndex        =   37
      Top             =   4080
      Width           =   4215
   End
   Begin VB.TextBox numerofm 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3720
      MaxLength       =   11
      TabIndex        =   36
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox seriefm 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   4
      TabIndex        =   35
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox tipofm 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   3
      TabIndex        =   34
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox colafm 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   33
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox puertofm 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5520
      MaxLength       =   60
      TabIndex        =   32
      Top             =   3720
      Width           =   4215
   End
   Begin VB.TextBox numerobm 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3720
      MaxLength       =   11
      TabIndex        =   31
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox seriebm 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   4
      TabIndex        =   30
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox tipobm 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   3
      TabIndex        =   29
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox colabm 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   28
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox puertobm 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5520
      MaxLength       =   60
      TabIndex        =   27
      Top             =   3360
      Width           =   4215
   End
   Begin VB.TextBox numeronv 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3720
      MaxLength       =   11
      TabIndex        =   26
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox serienv 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   4
      TabIndex        =   25
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox tiponv 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   3
      TabIndex        =   24
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox colanv 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   23
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox puertonv 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5520
      MaxLength       =   60
      TabIndex        =   22
      Top             =   3000
      Width           =   4215
   End
   Begin VB.TextBox numerotf 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3720
      MaxLength       =   11
      TabIndex        =   21
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox serietf 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   4
      TabIndex        =   20
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox tipotf 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   3
      TabIndex        =   19
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox colatf 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   18
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox puertotf 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5520
      MaxLength       =   60
      TabIndex        =   17
      Top             =   2640
      Width           =   4215
   End
   Begin VB.TextBox puertotb 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5520
      MaxLength       =   60
      TabIndex        =   15
      Top             =   2280
      Width           =   4215
   End
   Begin VB.TextBox colatb 
      BackColor       =   &H00E0E0E0&
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
      TabIndex        =   13
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox tipotb 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   3
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox serietb 
      BackColor       =   &H00E0E0E0&
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
      MaxLength       =   4
      TabIndex        =   9
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox numerotb 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3720
      MaxLength       =   11
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox moneda 
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
      MaxLength       =   1
      TabIndex        =   3
      Top             =   720
      Width           =   255
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
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   1
      Top             =   360
      Width           =   2655
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
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox puertocua 
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
      Left            =   7920
      MaxLength       =   60
      TabIndex        =   230
      Top             =   8880
      Width           =   1695
   End
   Begin VB.TextBox colacua 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9600
      MaxLength       =   1
      TabIndex        =   231
      Top             =   8880
      Width           =   495
   End
   Begin VB.TextBox ObligaPersonas 
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
      Left            =   8280
      MaxLength       =   1
      TabIndex        =   227
      Top             =   1155
      Width           =   525
   End
   Begin VB.TextBox obligaclavemesa 
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
      Left            =   8280
      MaxLength       =   1
      TabIndex        =   229
      Top             =   1545
      Width           =   525
   End
   Begin VB.TextBox obligavdelivery 
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
      Left            =   8280
      MaxLength       =   1
      TabIndex        =   232
      Top             =   400
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Delivery"
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
      Height          =   375
      Left            =   6360
      TabIndex        =   245
      Top             =   9240
      Width           =   1620
   End
   Begin VB.Label lblNoRepitenciaSi 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "(Si,TOUCH,MARK)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   11400
      TabIndex        =   242
      Top             =   9240
      Width           =   1170
   End
   Begin VB.Label labelObliga 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ObligaClave Mesa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   6510
      TabIndex        =   228
      Top             =   1530
      Width           =   1785
   End
   Begin VB.Label label80 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Obliga Personas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6495
      TabIndex        =   226
      Top             =   1160
      Width           =   1845
   End
   Begin VB.Label Label44 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TamLetraInterna"
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
      Height          =   375
      Left            =   4320
      TabIndex        =   209
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Al Terminar venta enviar Correo - Codigo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   8800
      TabIndex        =   203
      Top             =   0
      Width           =   1620
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Habl"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9600
      TabIndex        =   201
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PideClavePrecio                             Obliga Vend Deliv                MultiComanda"
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
      Height          =   1215
      Left            =   6520
      TabIndex        =   194
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clave Congela                              Copia Orden Despacho                             Obliga Hab Credit"
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
      Height          =   1215
      Left            =   10950
      TabIndex        =   192
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Colores Comando                                      Clave Limpia Pantalla"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   10950
      TabIndex        =   188
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label28 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cola          Decimal"
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
      Height          =   735
      Left            =   8280
      TabIndex        =   187
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Font       Formato Pedido"
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
      Height          =   735
      Left            =   2160
      TabIndex        =   183
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tamaño B/F                        Tamaño Ordenes "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   120
      TabIndex        =   180
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label77 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   """C"" o Vacio = Coc                ""P"" = Ped "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   9780
      TabIndex        =   176
      Top             =   7200
      Width           =   1410
   End
   Begin VB.Label Label72 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ordena Consulta/Sentido"
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
      Left            =   3720
      TabIndex        =   174
      Top             =   9240
      Width           =   1695
   End
   Begin VB.Label Label69 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Necesita Apertura                      Deshabilitar Caja                       Segundos/Video"
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
      Height          =   1095
      Left            =   10080
      TabIndex        =   170
      Top             =   8160
      Width           =   1935
   End
   Begin VB.Label Label68 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ve Costos Caja/ Habilitar clave Copia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2880
      TabIndex        =   168
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Label Label59 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ActivarPrecioMinimo                      Stock Minimo   "
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
      Height          =   735
      Left            =   4320
      TabIndex        =   150
      Top             =   360
      Width           =   1770
   End
   Begin VB.Label Label57 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1=Torrey 2=Digi  3=CasSw 4=CasPr"
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
      Height          =   765
      Left            =   5280
      TabIndex        =   147
      Top             =   6720
      Width           =   915
   End
   Begin VB.Label Label56 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PuertoCierre                              Cuadre Parcial"
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
      Height          =   735
      Left            =   6360
      TabIndex        =   146
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Label Label55 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Permite Precio=0                            Cierres Correlativo                           "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   120
      TabIndex        =   143
      Top             =   8880
      Width           =   2055
   End
   Begin VB.Label Label53 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Habilitar Terminales       Terminales Permitidos"
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
      Height          =   735
      Left            =   10320
      TabIndex        =   139
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label51 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NoRepitencia "
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
      Height          =   375
      Left            =   10080
      TabIndex        =   133
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label Label50 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Port.EstadoCta"
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
      Height          =   375
      Left            =   6360
      TabIndex        =   131
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label48 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hab"
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
      Height          =   375
      Left            =   9840
      TabIndex        =   128
      Top             =   7800
      Width           =   495
   End
   Begin VB.Label Label47 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Puerto.Ord.desp."
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
      Height          =   375
      Left            =   6360
      TabIndex        =   126
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Label Label42 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Local Caja-> Almacen->"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   120
      TabIndex        =   113
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label41 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Redeondeo                             Lista Precios Nro"
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
      Height          =   735
      Left            =   6360
      TabIndex        =   111
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label37 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C.aja T.erminal  Salon Defecto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   8805
      TabIndex        =   108
      Top             =   1185
      Width           =   1575
   End
   Begin VB.Label Label36 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X    Cola "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12000
      TabIndex        =   106
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label34 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PI"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   11520
      TabIndex        =   96
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label32 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PuertoCajon                            0.Star 1.Epson"
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
      Height          =   735
      Left            =   6360
      TabIndex        =   85
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Tcaja.frx":0000
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
      Height          =   1215
      Left            =   3120
      TabIndex        =   81
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ReciboEgreso                       Exonerado            NotaCredito"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1215
      Left            =   120
      TabIndex        =   79
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ActivoBalanza                              TipoBalanza                         PuertoBalanza"
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
      Height          =   975
      Left            =   3120
      TabIndex        =   66
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Congela-Num."
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
      Left            =   3120
      TabIndex        =   63
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vendedor                    Cliente Obliga   SerieTicketera                 Ventas<=0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1455
      Left            =   120
      TabIndex        =   61
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   9720
      TabIndex        =   57
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TipoDocDefault"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   120
      TabIndex        =   49
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Guia Remision                     Pedidos           ReciboIngreso"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1095
      Left            =   120
      TabIndex        =   42
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Puerto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
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
      Left            =   5040
      TabIndex        =   14
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BoletaManual             FacturaManual"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tick.Boleta            TickFactura     Nota Venta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
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
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Seriex 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TipoDoc"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caja Terminal       Descripcion->           Moneda->"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   0
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
Attribute VB_Name = "tcaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim t6            As String

Dim t7            As String

Dim t8            As String

Dim t9            As String

Dim t10           As String

Dim uvueltos      As String

Dim uvueltod      As String

Dim cuadreparcial As String

Dim delivery      As String

Dim precuenta     As String

Dim copiaventas   As String

Dim detraccion    As String

Private Sub ajdu1_Click()

    If Frame1.Visible = True Then Exit Sub
    inicializa
    codigo = ""
    codigo.SetFocus

End Sub

Private Sub bo712_Click()

    Dim found As Integer

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
        Frame1.Visible = False
        codigo.SetFocus
        Exit Sub

    End If

    Command1_Click

End Sub

Private Sub buffer_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode <> 13 And KeyCode <> 27 Then
        ejecuta 0
        Exit Sub

    End If

End Sub

Private Sub cmdAddEntry_Click()

End Sub

Private Sub cmdDelete_Click()

End Sub

Private Sub cmdExit_Click()

End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub cmdSort_Click()

End Sub

Private Sub capuerto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "GAVETAPUERTO"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(codigo) = 0 Then Exit Sub
    found = busca_registro()

    If found = 0 Then
        inicializa

    End If

    descripcio.SetFocus

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame1.Visible = True
        buffer = ""
        buffer.SetFocus
        Command1_Click

    End If

End Sub

Private Sub Combo2_Click()

    If Combo2 <> "" Then
        nombrefont = Trim(Combo2)

    End If

End Sub

Private Sub Command1_Click()
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim rconsulta As New ADODB.Recordset

    Dim buf       As String

    If Len(buffer) = 0 Then
        buf = "select Descripcio,Caja,Local,bodega from parameca "
    Else
        buf = "select Descripcio,Caja ,Local,Bodega from parameca where " & Combo1 & " like '" & buffer & "%'"

    End If

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True And rconsulta.BOF = True Then
        rconsulta.Close
        'buffer.SetFocus
        Exit Sub

    End If

    Set dbGrid1.DataSource = rconsulta

    dbGrid1.columns(0).Width = 4000
    dbGrid1.columns(1).Width = 2000

    If sw = 1 Then
        dbGrid1.SetFocus

    End If

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

    If Frame3.Caption = "COPIAOD" Then
        copiaod = cboprinters

    End If

    If Frame3.Caption = "ORDEN DESPACHO" Then
        odpuerto = cboprinters

    End If

    If Frame3.Caption = "GAVETAPUERTO" Then
        capuerto = cboprinters

    End If

    If Frame3.Caption = "PUERTOC" Then
        puertocie = cboprinters

    End If

    If Frame3.Caption = "PUERTOCUA" Then
        puertocua = cboprinters

    End If

    If Frame3.Caption = "ESTADOCUENTA" Then
        ecpuerto = cboprinters

    End If

    If Frame3.Caption = "PUERTOB" Then
        puertotb = cboprinters

    End If

    If Frame3.Caption = "PUERTOF" Then
        puertotf = cboprinters

    End If

    If Frame3.Caption = "PUERTONV" Then
        puertonv = cboprinters

    End If

    If Frame3.Caption = "PUERTOBM" Then
        puertobm = cboprinters

    End If

    If Frame3.Caption = "PUERTOFM" Then
        puertofm = cboprinters

    End If

    If Frame3.Caption = "PUERTOOT" Then
        puertoot = cboprinters

    End If

    If Frame3.Caption = "PUERTOPE" Then
        puertope = cboprinters

    End If

    If Frame3.Caption = "PUERTORI" Then
        puertori = cboprinters

    End If

    If Frame3.Caption = "PUERTORE" Then
        puertore = cboprinters

    End If

    If Frame3.Caption = "PUERTOEXO" Then
        puertoexo = cboprinters

    End If

    If Frame3.Caption = "PUERTONC" Then
        puertonc = cboprinters

    End If

    ''' 02/12/2017 Modulo Delivery Proyecto Principal
    If Frame3.Caption = "DELIVERY" Then
        puertodelivery = cboprinters

    End If

    ''' 02/12/2017 Modulo Delivery Proyecto Principal

    Frame3.Visible = False

End Sub

''' 02/12/2017 Modulo Delivery Proyecto Principal
Private Sub puertodelivery_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "DELIVERY"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

''' 02/12/2017 Modulo Delivery Proyecto Principal

Private Sub Command4_Click()
    Frame3.Visible = False

End Sub

Private Sub copiaod_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "COPIAOD"
        Frame3.Visible = True
        copiaod.SetFocus
        Exit Sub

    End If

End Sub

Private Sub Label71_Click()

End Sub

Private Sub Label22_Click()

    '''26/10/2017 Listas de ciertas caja para cobrar pedido
    If Len(Trim(codigo)) = 0 Then Exit Sub
    FrmAsignaPedidoCaja.caja = Trim(codigo)
    FrmAsignaPedidoCaja.Show 1

    '''26/10/2017 Listas de ciertas caja para cobrar pedido
End Sub

Private Sub puertocua_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "PUERTOCUA"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        codigo = dbGrid1.columns(1)
        Frame1.Visible = False
        codigo.SetFocus
        codigo_KeyPress 13

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

Private Sub descripcio_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub

End Sub

Private Sub descripcio_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        codigo.SetFocus
        Exit Sub

    End If

End Sub

Private Sub djuer1_Click()

    If Frame1.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "parameca"
    reporgen.Show 1

End Sub

Private Sub dlo132_Click()

    If Frame4.Visible = True Then
        Frame4.Visible = False
        Exit Sub

    End If

    If Frame3.Visible = True Then
        Frame3.Visible = False
        Exit Sub

    End If

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    tcaja.Hide
    Unload tcaja

End Sub

Private Sub ecpuerto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "ESTADOCUENTA"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub Form_Activate()
    Me.Width = 13905: Me.Height = 10500

End Sub

Private Sub Form_Load()

    Dim I As Integer

    On Error GoTo cmd134_err

    Frame1.Top = 10: Frame1.Left = 10
    Frame3.Top = 1800: Frame3.Left = 1200
    Frame4.Top = 600: Frame4.Left = 3480

    '10/07/2018 Edicion Comanda
    formatocomanda.Clear
    formatocomanda.AddItem "D|Driver"
    formatocomanda.AddItem "G|Genérico"
    formatocomanda.ListIndex = 0
    '10/07/2018 Edicion Comanda

    Combo2.Clear
    Combo2.AddItem ""

    For I = 0 To Printer.FontCount - 1
        Combo2.AddItem Printer.Fonts(I)
    Next I

    Combo2.ListIndex = 0

    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Caja"
    Combo1.ListIndex = 0

    If Printers.count = 0 Then
        Exit Sub

    End If
    
    ' load their device names into the combo box
    For I = 0 To Printers.count - 1
        cboprinters.AddItem Printers(I).DeviceName

        ' if this is the current printer, select it
        If Printers(I).DeviceName = Printer.DeviceName Then
            ' this indirectly executes ShowPrinterInfo
            cboprinters.ListIndex = I

        End If

    Next
    Exit Sub
cmd134_err:
    Exit Sub

End Sub

Sub inicializa()
    ObligaPersonas = ""
    obligaclavemesa = ""
    letrainterna = ""
    bold = ""
    clavedescongela = ""
    clavecongela = ""
    correo = ""
    gavetasw = ""
    tipo_balanza = ""
    flag1 = ""
    stkminimo = ""
    clienteo = ""
    vdetalle = ""
    obligaprecio = ""
    copiaod = ""
    obligacredito = ""

    obligavendedor = ""
    gavetacola = ""
    'nroempaques = ""
    limpiapantalla = ""
    grabacomanda = ""
    tamanorden = ""
    nombrefont = "Courier New"
    salon = ""
    tamanoletra = ""
    descuento = ""
    odcola = ""
    ecpuerto = ""
    eccola = ""
    'autoservicio = ""
    'comanda = ""
    'delivery = ""
    pmesero = ""
    precuenta = ""
    cuadreparcial = ""
    copiaventas = ""
    anulaventas = ""
    cierrecaja = ""
    ingresodinero = ""
    egresodinero = ""
    ordenaproducto = ""
    habilitanota = ""
    clavecopia = ""
    apertura = ""
    vecocaja = ""
    colacie = ""
    colacua = ""
    obligavdelivery = ""

    ''11/07/2017 kenyo multicomandas
    multicomanda = ""
    ''11/07/2017 kenyo multicomandas

    deshab = ""
    hdetraccio = ""
    detraccion = ""
    video = ""
    segundo = ""
    sentido = ""
    parqueo = ""
    pm = ""
    'pm.Visible = True
    decimal1 = ""
    puertocie = ""
    puertocua = ""

    ''' 02/12/2017 Modulo Delivery Proyecto Principal
    '''23/10/2017 Mejora Delivery
    puertodelivery = ""
    coladelivery = ""
    '''23/10/2017 Mejora Delivery
    ''' 02/12/2017 Modulo Delivery Proyecto Principal

    'Balanza 2/3 dígitos
    digitos = ""
    'Balanza 2/3 dígitos

    tipocie = ""
    cierres = ""
    seriehd = ""
    t0 = ""
    t1 = ""
    t2 = ""
    t3 = ""
    t4 = ""
    t5 = ""

    repite = ""
    ecpuerto = ""
    eccola = ""
    odpuerto = ""
    hod = ""
    nosaldo = ""
    listap = ""
    local1 = ""

    noprecio = ""
    terminal = "C"
    redondeo = ""
    ctb = ""
    ctf = ""
    cnv = ""
    cbm = ""
    cfm = ""
    cot = ""
    cpro = ""
    crin = ""
    creg = ""
    cexo = ""
    cnc = ""
    fexo = ""
    FTB = ""
    FTF = ""
    FNV = ""
    FBM = ""
    FFM = ""
    FOT = ""
    FPRO = ""
    FRIN = ""
    FREG = ""
    Fnc = ""

    itb = ""
    itf = ""
    inv = ""
    ibm = ""
    ifm = ""
    iot = ""
    ipe = ""
    iexo = ""
    ire = ""
    iri = ""
    inc = ""

    siventa = ""
    capuerto = ""
    catipo = ""
    serieexo = ""
    serieti = ""
    FLAG = ""
    puertore = ""
    puertori = ""
    puertoexo = ""
    tipori = ""
    tipoexo = ""
    serieri = ""
    numerori = ""
    numeroexo = ""
    colari = ""
    archivori = ""
    archivoexo = ""
    tipore = ""
    tipoexo = ""
    seriere = ""
    numerore = ""
    colare = ""
    archivore = ""

    portbala = ""
    actbala = ""
    congela = ""
    cliente = ""
    vendedor = ""
    bodega = ""
    archivotb = ""
    archivotf = ""
    archivobm = ""
    archivofm = ""
    archivonv = ""
    archivoot = ""
    archivope = ""
    archivonc = ""

    descripcio = ""
    moneda = ""
    tipodefa = ""
    tipope = ""
    seriepe = ""
    numerope = ""
    colape = ""
    puertope = ""

    tipotb = ""
    serietb = ""
    numerotb = ""
    colatb = ""
    puertotb = ""
    tipotf = ""
    serietf = ""
    numerotf = ""
    colatf = ""
    puertotf = ""
    tipobm = ""
    seriebm = ""
    numerobm = ""
    colabm = ""
    puertobm = ""
    tipofm = ""
    seriefm = ""
    numerofm = ""
    colafm = ""
    puertofm = ""
    tiponv = ""
    serienv = ""
    numeronv = ""
    colanv = ""
    puertonv = ""
    tipoot = ""
    serieot = ""
    numeroot = ""
    colaot = ""
    puertoot = ""

    'nota credito
    tiponc = ""
    serienc = ""
    numeronc = ""
    colanc = ""
    puertonc = ""

    tipoexo = ""
    serieexo = ""
    numeroexo = ""
    colaexo = ""
    puertoexo = ""
    archivoexo = ""
    iexo = ""
    fexo = ""
    cexo = ""

End Sub

Function borra_registro()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from parameca where caja='" & codigo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If MsgBox("Desea Borra el registro", 1, "Aviso") = "1" Then
            mytablex.Delete
            borra_registro = 1

        End If

    End If

    mytablex.Close

End Function

Function busca_registro()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from parameca where caja='" & codigo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        pone_registro mytablex
        busca_registro = 1

    End If

    mytablex.Close
 
End Function

Sub pone_registro(mytablex As ADODB.Recordset)
    ObligaPersonas = Trim("" & mytablex.Fields("obligapersonas"))
    obligaclavemesa = Trim("" & mytablex.Fields("obligaclavemesa"))
    letrainterna = Trim("" & mytablex.Fields("letrainterna"))
    bold = Trim("" & mytablex.Fields("Bold"))
    clavedescongela = "" & mytablex.Fields("clavedescongela")
    clavecongela = "" & mytablex.Fields("clavecongela")
    correo = "" & mytablex.Fields("correo")
    gavetasw = "" & mytablex.Fields("gavetasw")
    tipo_balanza = "" & mytablex.Fields("tipo_balanza")
    flag1 = "" & mytablex.Fields("flag1")
    stkminimo = "" & mytablex.Fields("stkminimo")
    clienteo = "" & mytablex.Fields("clienteo")
    vdetalle = "" & mytablex.Fields("vdetalle")
    obligaprecio = "" & mytablex.Fields("obligaprecio")
    copiaod = "" & mytablex.Fields("copiaod")

    obligacredito = "" & mytablex.Fields("obligacredito")
    gavetacola = "" & mytablex.Fields("gavetacola")
    'nroempaques = "" & mytablex.Fields("nroempaques")
    limpiapantalla = "" & mytablex.Fields("limpiapantalla")
    grabacomanda = "" & mytablex.Fields("grabacomanda")
    tamanorden = "" & mytablex.Fields("tamanorden")
    nombrefont = Trim("" & mytablex.Fields("nombrefont"))

    If Len(Trim(nombrefont)) = 0 Then
        nombrefont = "Courier New"

    End If

    salon = "" & mytablex.Fields("salon")
    tamanoletra = "" & mytablex.Fields("tamanoletra")
    descuento = "" & mytablex.Fields("descuento")
    odcola = "" & mytablex.Fields("odcola")
    ecpuerto = "" & mytablex.Fields("ecpuerto")
    eccola = "" & mytablex.Fields("eccola")
    'autoservicio = "" & mytablex.Fields("tpvauto")
    'comanda = "" & mytablex.Fields("comanda")
    delivery = "" & mytablex.Fields("delivery")
    pmesero = "" & mytablex.Fields("pmesero")

    precuenta = "" & mytablex.Fields("precuenta")
    cuadreparcial = "" & mytablex.Fields("cuadreparcial")
    copiaventas = "" & mytablex.Fields("copiaventas")
    anulaventas = "" & mytablex.Fields("anulaventas")
    cierrecaja = "" & mytablex.Fields("cierrecaja")
    ingresodinero = "" & mytablex.Fields("ingresodinero")
    egresodinero = "" & mytablex.Fields("egresodinero")

    obligacredito = "" & mytablex.Fields("obligacredito")
    ordenaproducto = "" & mytablex.Fields("ordenaproducto")
    habilitanota = "" & mytablex.Fields("habilitanota")
    clavecopia = "" & mytablex.Fields("clavecopia")
    apertura = "" & mytablex.Fields("apertura")
    vecocaja = "" & mytablex.Fields("vecocaja")
    colacie = "" & mytablex.Fields("colacie")
    colacua = "" & mytablex.Fields("colacua")
    obligavdelivery = "" & mytablex.Fields("obligavdelivery")

    ''11/07/2017 kenyo multicomandas
    multicomanda = "" & mytablex.Fields("multicomanda")

    ''11/07/2017 kenyo multicomandas

    deshab = "" & mytablex.Fields("deshab")
    hdetraccio = "" & mytablex.Fields("hdetraccio")
    detraccion = "" & mytablex.Fields("detraccion")
    tipoexo = "" & mytablex.Fields("tipoexo")
    serieexo = "" & mytablex.Fields("serieexo")
    numeroexo = "" & mytablex.Fields("numeroexo")
    colaexo = "" & mytablex.Fields("colaexo")
    puertoexo = "" & mytablex.Fields("puertoexo")
    archivoexo = "" & mytablex.Fields("archivoexo")
    iexo = "" & mytablex.Fields("iexo")
    fexo = "" & mytablex.Fields("fexo")
    cexo = "" & mytablex.Fields("cexo")

    segundo = "" & mytablex.Fields("segundo")
    video = "" & mytablex.Fields("video")
    sentido = "" & mytablex.Fields("sentido")
    parqueo = "" & mytablex.Fields("parqueo")
    decimal1 = "" & mytablex.Fields("decimal")
    tipocie = "" & mytablex.Fields("tipocie")
    puertocie = "" & mytablex.Fields("puertocie")
    puertocua = "" & mytablex.Fields("puertocua")

    ''' 02/12/2017 Modulo Delivery Proyecto Principal
    '''23/10/2017 Mejora Delivery
    puertodelivery = "" & mytablex.Fields("puertodelivery")
    coladelivery = "" & mytablex.Fields("coladelivery")
    '''23/10/2017 Mejora Delivery
    ''' 02/12/2017 Modulo Delivery Proyecto Principal

    'Balanza 2/3 dígitos
    digitos = "" & mytablex.Fields("digitos")
    'Balanza 2/3 dígitos

    '10/07/2018 Edicion Comanda
    formatocomanda.ListIndex = 0

    If "" & mytablex.Fields("formatocomanda") = "G" Then
        formatocomanda.ListIndex = 1

    End If

    '10/07/2018 Edicion Comanda

    cierres = "" & mytablex.Fields("cierres")
    seriehd = "" & mytablex.Fields("seriehd")
    pm = "" & mytablex.Fields("pm")

    t0 = "" & mytablex.Fields("t0")
    t1 = "" & mytablex.Fields("t1")
    t2 = "" & mytablex.Fields("t2")
    t3 = "" & mytablex.Fields("t3")
    t4 = "" & mytablex.Fields("t4")
    t5 = "" & mytablex.Fields("t5")

    repite = "" & mytablex.Fields("repite")
    odpuerto = "" & mytablex.Fields("odpuerto")
    hod = "" & mytablex.Fields("hod")
    ecpuerto = "" & mytablex.Fields("ecpuerto")
    eccola = "" & mytablex.Fields("eccola")

    listap = "" & mytablex.Fields("listap")
    nosaldo = "" & mytablex.Fields("nosaldo")

    local1 = "" & mytablex.Fields("local")
    redondeo = "" & mytablex.Fields("redondeo")
    noprecio = "" & mytablex.Fields("noprecio")

    terminal = "" & mytablex.Fields("terminal")
    FTB = "" & mytablex.Fields("ftb")
    FTF = "" & mytablex.Fields("ftf")
    FNV = "" & mytablex.Fields("fnv")
    FBM = "" & mytablex.Fields("fbm")
    FFM = "" & mytablex.Fields("ffm")
    FOT = "" & mytablex.Fields("fot")
    FPRO = "" & mytablex.Fields("fpro")
    FRIN = "" & mytablex.Fields("frin")
    FREG = "" & mytablex.Fields("freg")

    Fnc = "" & mytablex.Fields("fnc")

    ctb = "" & mytablex.Fields("Ctb")
    ctf = "" & mytablex.Fields("Ctf")
    cnv = "" & mytablex.Fields("Cnv")
    cbm = "" & mytablex.Fields("Cbm")
    cfm = "" & mytablex.Fields("Cfm")
    cot = "" & mytablex.Fields("Cot")
    cpro = "" & mytablex.Fields("Cpro")
    crin = "" & mytablex.Fields("Crin")
    creg = "" & mytablex.Fields("Creg")
    cnc = "" & mytablex.Fields("cnc")

    itb = "" & mytablex.Fields("itb")
    itf = "" & mytablex.Fields("itf")
    inv = "" & mytablex.Fields("inv")
    ibm = "" & mytablex.Fields("ibm")
    ifm = "" & mytablex.Fields("ifm")
    iot = "" & mytablex.Fields("iot")
    ipe = "" & mytablex.Fields("ipe")
    ire = "" & mytablex.Fields("ire")
    iri = "" & mytablex.Fields("iri")
    inc = "" & mytablex.Fields("inc")

    siventa = "" & mytablex.Fields("siventa")
    capuerto = "" & mytablex.Fields("capuerto")
    catipo = "" & mytablex.Fields("catipo")
    serieti = "" & mytablex.Fields("serieti")
    FLAG = "" & mytablex.Fields("flag")
    puertori = "" & mytablex.Fields("puertori")
    puertore = "" & mytablex.Fields("puertore")

    tipori = "" & mytablex.Fields("tipori")
    serieri = "" & mytablex.Fields("serieri")
    numerori = "" & mytablex.Fields("numerori")
    colari = "" & mytablex.Fields("colari")
    archivori = "" & mytablex.Fields("archivori")

    tipore = "" & mytablex.Fields("tipore")
    seriere = "" & mytablex.Fields("seriere")
    numerore = "" & mytablex.Fields("numerore")
    colare = "" & mytablex.Fields("colare")
    archivore = "" & mytablex.Fields("archivore")

    actbala = "" & mytablex.Fields("actbala")
    portbala = "" & mytablex.Fields("portbala")
    congela = "" & mytablex.Fields("congela")
    cliente = "" & mytablex.Fields("cliente")
    vendedor = "" & mytablex.Fields("vendedor")
    bodega = "" & mytablex.Fields("bodega")
    archivotb = "" & mytablex.Fields("archivotb")
    archivotf = "" & mytablex.Fields("archivotf")
    archivobm = "" & mytablex.Fields("archivobm")
    archivofm = "" & mytablex.Fields("archivofm")
    archivoot = "" & mytablex.Fields("archivoot")
    archivope = "" & mytablex.Fields("archivope")
    archivonv = "" & mytablex.Fields("archivonv")
    archivonc = "" & mytablex.Fields("archivonc")

    tipodefa = "" & mytablex.Fields("tipodefa")
    tipope = "" & mytablex.Fields("tipope")
    seriepe = "" & mytablex.Fields("seriepe")
    numerope = "" & mytablex.Fields("numerope")
    colape = "" & mytablex.Fields("colape")
    puertope = "" & mytablex.Fields("puertope")

    tipotb = "" & mytablex.Fields("tipotb")
    serietb = "" & mytablex.Fields("serietb")
    numerotb = "" & mytablex.Fields("numerotb")
    colatb = "" & mytablex.Fields("colatb")
    puertotb = "" & mytablex.Fields("puertotb")

    tipotf = "" & mytablex.Fields("tipotf")
    serietf = "" & mytablex.Fields("serietf")
    numerotf = "" & mytablex.Fields("numerotf")
    colatf = "" & mytablex.Fields("colatf")
    puertotf = "" & mytablex.Fields("puertotf")

    tipobm = "" & mytablex.Fields("tipobm")
    seriebm = "" & mytablex.Fields("seriebm")
    numerobm = "" & mytablex.Fields("numerobm")
    colabm = "" & mytablex.Fields("colabm")
    puertobm = "" & mytablex.Fields("puertobm")

    tipofm = "" & mytablex.Fields("tipofm")
    seriefm = "" & mytablex.Fields("seriefm")
    numerofm = "" & mytablex.Fields("numerofm")
    colafm = "" & mytablex.Fields("colafm")
    puertofm = "" & mytablex.Fields("puertofm")

    tiponv = "" & mytablex.Fields("tiponv")
    serienv = "" & mytablex.Fields("serienv")
    numeronv = "" & mytablex.Fields("numeronv")
    colanv = "" & mytablex.Fields("colanv")
    puertonv = "" & mytablex.Fields("puertonv")

    tipoot = "" & mytablex.Fields("tipoot")
    serieot = "" & mytablex.Fields("serieot")
    numeroot = "" & mytablex.Fields("numeroot")
    colaot = "" & mytablex.Fields("colaot")
    puertoot = "" & mytablex.Fields("puertoot")

    'nc
    tiponc = "" & mytablex.Fields("tiponc")
    serienc = "" & mytablex.Fields("serienc")
    numeronc = "" & mytablex.Fields("numeronc")
    colanc = "" & mytablex.Fields("colanc")
    puertonc = "" & mytablex.Fields("puertonc")

    moneda = "" & mytablex.Fields("moneda")
    codigo = "" & mytablex.Fields("caja")
    descripcio = "" & mytablex.Fields("descripcio")
    obligavendedor = "" & mytablex.Fields("obligavendedor")

End Sub

Sub grabando(mytablex As ADODB.Recordset)
    mytablex.Fields("obligapersonas") = Trim(ObligaPersonas)
    mytablex.Fields("obligaclavemesa") = Trim(obligaclavemesa)
    mytablex.Fields("letrainterna") = Trim(letrainterna)
    mytablex.Fields("correo") = Trim(correo)
    mytablex.Fields("bold") = Trim(bold)
    mytablex.Fields("clavecongela") = Trim(clavecongela)
    mytablex.Fields("clavedescongela") = Trim(clavedescongela)
    mytablex.Fields("gavetasw") = Trim(gavetasw)
    mytablex.Fields("tipo_balanza") = Trim(tipo_balanza)
    mytablex.Fields("flag1") = Trim(flag1)
    mytablex.Fields("stkminimo") = Trim(stkminimo)
    mytablex.Fields("clienteo") = Trim(clienteo)
    mytablex.Fields("vdetalle") = Trim(vdetalle)

    mytablex.Fields("obligaprecio") = Trim(obligaprecio)
    mytablex.Fields("copiaod") = Trim(copiaod)

    mytablex.Fields("obligacredito") = Trim(obligacredito)
    mytablex.Fields("obligavendedor") = Trim(obligavendedor)
    mytablex.Fields("gavetacola") = Trim(gavetacola)

    'mytablex.Fields("nroempaques") = Trim(nroempaques)
    mytablex.Fields("limpiapantalla") = limpiapantalla
    mytablex.Fields("grabacomanda") = Trim(grabacomanda)

    mytablex.Fields("tamanorden") = Trim(tamanorden)
    mytablex.Fields("nombrefont") = Trim(nombrefont)
    mytablex.Fields("salon") = Trim(salon)
    mytablex.Fields("tamanoletra") = tamanoletra
    mytablex.Fields("descuento") = descuento
    mytablex.Fields("odcola") = odcola
    'mytablex.Fields("tpvauto") = autoservicio
    'mytablex.Fields("comanda") = comanda
    'mytablex.Fields("delivery") = delivery
    mytablex.Fields("pmesero") = pmesero

    mytablex.Fields("precuenta") = precuenta
    mytablex.Fields("cuadreparcial") = cuadreparcial
    mytablex.Fields("copiaventas") = copiaventas
    mytablex.Fields("anulaventas") = anulaventas
    mytablex.Fields("cierrecaja") = cierrecaja
    mytablex.Fields("ingresodinero") = ingresodinero
    mytablex.Fields("egresodinero") = egresodinero

    mytablex.Fields("ordenaproducto") = ordenaproducto
    mytablex.Fields("habilitanota") = habilitanota
    mytablex.Fields("clavecopia") = clavecopia
    mytablex.Fields("apertura") = apertura
    mytablex.Fields("vecocaja") = vecocaja
    mytablex.Fields("colacie") = colacie
    mytablex.Fields("colacua") = colacua
    mytablex.Fields("obligavdelivery") = obligavdelivery

    ''11/07/2017 kenyo multicomandas
    mytablex.Fields("multicomanda") = multicomanda
    ''11/07/2017 kenyo multicomandas

    mytablex.Fields("hdetraccio") = hdetraccio
    mytablex.Fields("detraccion") = Val(detraccion)
    mytablex.Fields("tipoexo") = tipoexo
    mytablex.Fields("deshab") = deshab
    mytablex.Fields("serieexo") = serieexo
    mytablex.Fields("numeroexo") = numeroexo
    mytablex.Fields("colaexo") = colaexo
    mytablex.Fields("puertoexo") = puertoexo
    mytablex.Fields("archivoexo") = archivoexo
    mytablex.Fields("iexo") = iexo
    mytablex.Fields("fexo") = fexo
    mytablex.Fields("cexo") = cexo

    mytablex.Fields("video") = video
    mytablex.Fields("segundo") = Val(segundo)
    mytablex.Fields("sentido") = sentido
    mytablex.Fields("decimal") = decimal1
    mytablex.Fields("tipocie") = tipocie
    mytablex.Fields("puertocie") = puertocie
    mytablex.Fields("puertocua") = puertocua

    ''' 02/12/2017 Modulo Delivery Proyecto Principal
    '''23/10/2017 Mejora Delivery
    mytablex.Fields("puertodelivery") = puertodelivery
    mytablex.Fields("coladelivery") = coladelivery
    '''23/10/2017 Mejora Delivery
    ''' 02/12/2017 Modulo Delivery Proyecto Principal

    'Balanza 2/3 dígitos
    mytablex.Fields("digitos") = digitos
    'Balanza 2/3 dígitos

    '10/07/2018 Edicion Comanda
    mytablex.Fields("formatocomanda") = extra_loquesea(Trim(formatocomanda))
    '10/07/2018 Edicion Comanda

    mytablex.Fields("cierres") = cierres
    mytablex.Fields("seriehd") = seriehd
    mytablex.Fields("pm") = pm
    mytablex.Fields("parqueo") = parqueo

    mytablex.Fields("t0") = t0
    mytablex.Fields("t1") = t1
    mytablex.Fields("t2") = t2
    mytablex.Fields("t3") = t3
    mytablex.Fields("t4") = t4
    mytablex.Fields("t5") = t5

    mytablex.Fields("repite") = repite
    mytablex.Fields("ecpuerto") = ecpuerto
    mytablex.Fields("eccola") = eccola

    mytablex.Fields("odpuerto") = odpuerto
    mytablex.Fields("hod") = hod
    mytablex.Fields("nosaldo") = nosaldo
    mytablex.Fields("listap") = listap
    mytablex.Fields("local") = local1
    mytablex.Fields("redondeo") = redondeo
    mytablex.Fields("terminal") = terminal

    mytablex.Fields("noprecio") = noprecio
    mytablex.Fields("ftb") = FTB
    mytablex.Fields("ftf") = FTF
    mytablex.Fields("fnv") = FNV
    mytablex.Fields("fbm") = FBM
    mytablex.Fields("ffm") = FFM
    mytablex.Fields("fot") = FOT
    mytablex.Fields("fpro") = FPRO
    mytablex.Fields("frin") = FRIN
    mytablex.Fields("freg") = FREG
    mytablex.Fields("fnc") = Fnc

    mytablex.Fields("Ctb") = ctb
    mytablex.Fields("Ctf") = ctf
    mytablex.Fields("Cnv") = cnv
    mytablex.Fields("Cbm") = cbm
    mytablex.Fields("Cfm") = cfm
    mytablex.Fields("Cot") = cot
    mytablex.Fields("Cpro") = cpro
    mytablex.Fields("Crin") = crin
    mytablex.Fields("Creg") = creg
    mytablex.Fields("Cnc") = cnc

    mytablex.Fields("itb") = itb
    mytablex.Fields("itf") = itf
    mytablex.Fields("inv") = inv
    mytablex.Fields("ibm") = ibm
    mytablex.Fields("ifm") = ifm
    mytablex.Fields("iot") = iot
    mytablex.Fields("ipe") = ipe
    mytablex.Fields("ire") = ire
    mytablex.Fields("iri") = iri
    mytablex.Fields("inc") = inc

    mytablex.Fields("capuerto") = capuerto  'gaveta
    mytablex.Fields("catipo") = catipo
    mytablex.Fields("serieti") = serieti
    mytablex.Fields("flag") = FLAG
    mytablex.Fields("siventa") = Val(siventa)
    mytablex.Fields("puertore") = puertore
    mytablex.Fields("puertori") = puertori

    mytablex.Fields("tipori") = tipori
    mytablex.Fields("serieri") = serieri
    mytablex.Fields("numerori") = numerori
    mytablex.Fields("colari") = colari
    mytablex.Fields("archivori") = archivori

    mytablex.Fields("tipore") = tipore
    mytablex.Fields("seriere") = seriere
    mytablex.Fields("numerore") = numerore
    mytablex.Fields("colare") = colare
    mytablex.Fields("archivore") = archivore

    mytablex.Fields("actbala") = actbala
    mytablex.Fields("portbala") = portbala
    mytablex.Fields("congela") = congela
    mytablex.Fields("vendedor") = vendedor
    mytablex.Fields("bodega") = bodega
    mytablex.Fields("cliente") = cliente
    mytablex.Fields("archivotb") = archivotb
    mytablex.Fields("archivotf") = archivotf
    mytablex.Fields("archivobm") = archivobm
    mytablex.Fields("archivofm") = archivofm
    mytablex.Fields("archivonv") = archivonv
    mytablex.Fields("archivoot") = archivoot
    mytablex.Fields("archivope") = archivope
    mytablex.Fields("archivonc") = archivonc

    mytablex.Fields("tipodefa") = tipodefa
    mytablex.Fields("caja") = codigo
    mytablex.Fields("moneda") = moneda

    mytablex.Fields("tipope") = tipope
    mytablex.Fields("seriepe") = seriepe
    mytablex.Fields("numerope") = numerope
    mytablex.Fields("colape") = colape
    mytablex.Fields("puertope") = puertope

    mytablex.Fields("tipotb") = tipotb
    mytablex.Fields("serietb") = serietb
    mytablex.Fields("numerotb") = numerotb
    mytablex.Fields("colatb") = colatb
    mytablex.Fields("puertotb") = puertotb

    mytablex.Fields("tipotf") = tipotf
    mytablex.Fields("serietf") = serietf
    mytablex.Fields("numerotf") = numerotf
    mytablex.Fields("colatf") = colatf
    mytablex.Fields("puertotf") = puertotf

    mytablex.Fields("tipobm") = tipobm
    mytablex.Fields("seriebm") = seriebm
    mytablex.Fields("numerobm") = numerobm
    mytablex.Fields("colabm") = colabm
    mytablex.Fields("puertobm") = puertobm

    mytablex.Fields("tipofm") = tipofm
    mytablex.Fields("seriefm") = seriefm
    mytablex.Fields("numerofm") = numerofm
    mytablex.Fields("colafm") = colafm
    mytablex.Fields("puertofm") = puertofm

    mytablex.Fields("tiponv") = tiponv
    mytablex.Fields("serienv") = serienv
    mytablex.Fields("numeronv") = numeronv
    mytablex.Fields("colanv") = colanv
    mytablex.Fields("puertonv") = puertonv
    mytablex.Fields("tipoot") = tipoot
    mytablex.Fields("serieot") = serieot
    mytablex.Fields("numeroot") = numeroot
    mytablex.Fields("colaot") = colaot
    mytablex.Fields("puertoot") = puertoot
    mytablex.Fields("descripcio") = descripcio

    mytablex.Fields("tiponc") = tiponc
    mytablex.Fields("serienc") = serienc
    mytablex.Fields("numeronc") = numeronc
    mytablex.Fields("colanc") = colanc
    mytablex.Fields("puertonc") = puertonc

    'If terminal = "C" Then
    'serietb = "B" & codigo
    'serietf = "F" & codigo
    'serienv = "N" & codigo
    'serieri = "I" & codigo
    'seriere = "E" & codigo
    'End If
    'If terminal = "T" Then
    'serietb = "T" & codigo
    'serietf = "T" & codigo
    'serienv = "T" & codigo
    'serieri = "T" & codigo
    'seriere = "T" & codigo
    'End If

End Sub

Private Sub grba1_Click()

    Dim found As Integer

    If Frame1.Visible = True Then Exit Sub

    'Testing Facturacion Electronica 14/03/2018
    If Len(serietb) = 4 Or Len(serietb) = 0 Then
        If Mid(serietb, 1, 1) = "B" Or Mid(serietb, 1, 1) = "T" Or Len(serietb) = 0 Then
        Else
            MsgBox ("Serie debe comenzar con B "), vbCritical
            serietb.SetFocus
            Exit Sub

        End If

    Else
        MsgBox ("Serie debe tener 4 dígitos"), vbCritical
        serietb.SetFocus
        Exit Sub

    End If

    If Len(serietf) = 4 Or Len(serietf) = 0 Then
        If Mid(serietf, 1, 1) = "F" Or Len(serietf) = 0 Then
        Else
            MsgBox ("Serie debe comenzar con F "), vbCritical
            serietf.SetFocus
            Exit Sub

        End If

    Else
        MsgBox ("Serie debe tener 4 dígitos"), vbCritical
        serietf.SetFocus
        Exit Sub

    End If

    'Testing Facturacion Electronica 14/03/2018

    found = grabar()

    If found = 0 Then Exit Sub
    codigo.SetFocus

End Sub

Private Sub Label1_Click()
    cmdSort_Click

End Sub

Function grabar()

    Dim found    As Integer

    Dim mytablex As New ADODB.Recordset

    found = valida()

    If found = 0 Then
        MsgBox "Campos invalidos", 48, "Aviso"
        Exit Function

    End If

    mytablex.Open "select * from parameca where caja='" & codigo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then

        mytablex.AddNew
        grabando mytablex
        mytablex.Update
        grabar = 1
   
        '22/05/2017 KENYO
        actualiza_estadotipo
    Else

        If MsgBox("Desea Reescribir?", 1, "Aviso") = 1 Then
            'mytablex.Edit
            grabando mytablex
            mytablex.Update
            grabar = 1
            '22/05/2017 KENYO
            actualiza_estadotipo

        End If

    End If

    '------------------------------------- ------------
    mytablex.Close
 
End Function

Sub actualiza_estadotipo()

    On Error GoTo cmd9093_err

    If codigo = "01" Then
        If (FTB = "S") Then cn.Execute ("update TIPO SET ESTADOT='2' WHERE tipo='1'") Else cn.Execute ("update TIPO SET ESTADOT='1' WHERE tipo='1'")
        If (FTF = "S") Then cn.Execute ("update TIPO SET ESTADOT='2' WHERE tipo='2'") Else cn.Execute ("update TIPO SET ESTADOT='1' WHERE tipo='2'")
        If (FNV = "S") Then cn.Execute ("update TIPO SET ESTADOT='2' WHERE tipo='5'") Else cn.Execute ("update TIPO SET ESTADOT='1' WHERE tipo='5'")
        If (FBM = "S") Then cn.Execute ("update TIPO SET ESTADOT='2' WHERE tipo='3'") Else cn.Execute ("update TIPO SET ESTADOT='1' WHERE tipo='3'")
        If (FFM = "S") Then cn.Execute ("update TIPO SET ESTADOT='2' WHERE tipo='4'") Else cn.Execute ("update TIPO SET ESTADOT='1' WHERE tipo='4'")

        '''27/07/2017 kenyo Testing Completo al Sistema
        If (Fnc = "S") Then cn.Execute ("update TIPO SET ESTADOT='2' WHERE tipo='N'") Else cn.Execute ("update TIPO SET ESTADOT='1' WHERE tipo='N'")
 
        '''27/07/2017 kenyo Testing Completo al Sistema

    End If

    Exit Sub
cmd9093_err:
    MsgBox "Aviso en actualiza receta " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function valida()

    If Len(codigo) = 0 Then
        codigo.SetFocus
        Exit Function

    End If

    If Len(descripcio) = 0 Then
        descripcio.SetFocus
        Exit Function

    End If

    If sentido <> "C" And sentido <> "S" And sentido <> "B" Then
        sentido = ""
        sentido.SetFocus
        Exit Function

    End If

    valida = 1

End Function

Private Sub Label29_Click()

    If Len(Trim(codigo)) = 0 Then Exit Sub
    toparam.caja = Trim(codigo)
    toparam.Show 1

End Sub

Private Sub Label3_Click()

    'Frame4.Visible = True
End Sub

Private Sub Label54_Click()
    seriehd = serie_disco_duro()

End Sub

Private Sub odpuerto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "ORDEN DESPACHO"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub puertobm_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "PUERTOBM"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub puertocie_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "PUERTOC"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub puertoexo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "PUERTOEXO"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub puertofm_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "PUERTOFM"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub puertonc_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "PUERTONC"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub puertonv_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "PUERTONV"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub puertoot_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "PUERTOOT"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub puertope_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "PUERTOPE"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub puertore_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "PUERTORE"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub puertori_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "PUERTORI"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub puertotb_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "PUERTOB"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub puertotf_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        Frame3.Caption = "PUERTOF"
        Frame3.Visible = True
        cboprinters.SetFocus
        Exit Sub

    End If

End Sub

Private Sub velocidad_Change()

End Sub

