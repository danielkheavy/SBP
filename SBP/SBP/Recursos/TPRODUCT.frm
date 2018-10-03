VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form tproduct 
   BackColor       =   &H00808080&
   Caption         =   "Tabla de productos"
   ClientHeight    =   8940
   ClientLeft      =   165
   ClientTop       =   420
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   15120
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox deliveryautom 
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
      Left            =   12900
      Style           =   2  'Dropdown List
      TabIndex        =   248
      Top             =   2300
      Width           =   615
   End
   Begin VB.ComboBox CostoReceta 
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
      Left            =   12300
      Style           =   2  'Dropdown List
      TabIndex        =   247
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame frmOtros 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   11280
      TabIndex        =   239
      Top             =   4560
      Visible         =   0   'False
      Width           =   5270
      Begin VB.TextBox talla 
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
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   243
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox proyecto 
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
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   242
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox procedencia 
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
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   241
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox sexo 
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
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   240
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblTallaSexo 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Talla                                 Sexo                       Procedencia               Proyecto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   244
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Puertos Adicionales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   11040
      TabIndex        =   223
      Top             =   4920
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton Command6 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   230
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox cboprinters3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   229
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cboprinters2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   228
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cboprinters1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   227
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox puertoimpresion3 
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
         MaxLength       =   30
         TabIndex        =   226
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox puertoimpresion2 
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
         MaxLength       =   30
         TabIndex        =   225
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox puertoimpresion1 
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
         MaxLength       =   30
         TabIndex        =   224
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   10680
      TabIndex        =   219
      Top             =   5040
      Visible         =   0   'False
      Width           =   13695
      Begin VB.ComboBox Combo3 
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
         TabIndex        =   221
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
         TabIndex        =   220
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7710
         Left            =   240
         TabIndex        =   222
         Top             =   1200
         Width           =   12120
         _ExtentX        =   21378
         _ExtentY        =   13600
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
      Begin ChamaleonButton.ChameleonBtn Acepta 
         Height          =   825
         Left            =   12360
         TabIndex        =   236
         Top             =   1320
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1455
         BTYPE           =   4
         TX              =   "ACEPTAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   255
         BCOLO           =   255
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "TPRODUCT.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn Cerrar 
         Height          =   585
         Left            =   12360
         TabIndex        =   237
         Top             =   2280
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1032
         BTYPE           =   4
         TX              =   "CERRAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   4210752
         BCOLO           =   4210752
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "TPRODUCT.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn Command5 
         Height          =   495
         Left            =   6120
         TabIndex        =   238
         Top             =   480
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "Buscar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421504
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "TPRODUCT.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.TextBox seinventaria 
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
      Left            =   14640
      MaxLength       =   1
      TabIndex        =   218
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   10920
      TabIndex        =   211
      Top             =   8640
      Visible         =   0   'False
      Width           =   10575
      Begin VB.TextBox barras2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   215
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Borra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9000
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TPRODUCT.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   214
         ToolTipText     =   "Borrar registro"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Graba"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9000
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TPRODUCT.frx":1266
         Style           =   1  'Graphical
         TabIndex        =   213
         ToolTipText     =   "Grabar registro"
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cierra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   9000
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TPRODUCT.frx":2478
         Style           =   1  'Graphical
         TabIndex        =   212
         ToolTipText     =   "Salir"
         Top             =   2520
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dbGrid2 
         Height          =   5655
         Left            =   75
         TabIndex        =   216
         Top             =   825
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9975
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
      Begin VB.Label Label59 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CodigoAdicionar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   90
         TabIndex        =   217
         Top             =   210
         Width           =   1455
      End
   End
   Begin VB.TextBox tecla 
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
      Left            =   13920
      MaxLength       =   1
      TabIndex        =   209
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox unidadp 
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
      Left            =   4680
      MaxLength       =   6
      TabIndex        =   207
      Top             =   8160
      Width           =   1335
   End
   Begin VB.TextBox factorp 
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
      Left            =   4680
      MaxLength       =   6
      TabIndex        =   206
      Top             =   8520
      Width           =   1335
   End
   Begin VB.ComboBox dia 
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
      Left            =   14280
      Style           =   2  'Dropdown List
      TabIndex        =   205
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox serviciomesa 
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
      TabIndex        =   203
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox fueldonde 
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
      Left            =   13260
      MaxLength       =   1
      TabIndex        =   202
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox comisioncredito 
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
      TabIndex        =   201
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox costoanterior1 
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
      Left            =   1260
      MaxLength       =   10
      TabIndex        =   199
      Top             =   5490
      Width           =   1335
   End
   Begin VB.TextBox costoanterior2 
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
      Left            =   1260
      MaxLength       =   10
      TabIndex        =   198
      Top             =   5850
      Width           =   1335
   End
   Begin VB.TextBox cola 
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
      Left            =   8580
      MaxLength       =   1
      TabIndex        =   196
      Top             =   8520
      Width           =   375
   End
   Begin VB.ComboBox cboprinters 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8580
      Style           =   2  'Dropdown List
      TabIndex        =   195
      Top             =   7440
      Width           =   3975
   End
   Begin VB.TextBox puertoimpresion 
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
      Left            =   8580
      MaxLength       =   30
      TabIndex        =   194
      Top             =   7800
      Width           =   3975
   End
   Begin VB.TextBox grupoimpresion 
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
      Left            =   8580
      MaxLength       =   5
      TabIndex        =   193
      Top             =   8160
      Width           =   1335
   End
   Begin VB.TextBox recetaprn 
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
      Left            =   9540
      MaxLength       =   1
      TabIndex        =   190
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox empaque_visible 
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
      Left            =   13800
      MaxLength       =   2
      TabIndex        =   189
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox platos 
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
      Left            =   1260
      MaxLength       =   5
      TabIndex        =   187
      Top             =   6210
      Width           =   495
   End
   Begin VB.TextBox fuel 
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
      Left            =   13020
      MaxLength       =   1
      TabIndex        =   186
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox touch 
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
      Left            =   10860
      MaxLength       =   6
      TabIndex        =   185
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox dsctoref 
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
      MaxLength       =   10
      TabIndex        =   183
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox margen 
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
      Left            =   10860
      MaxLength       =   6
      TabIndex        =   181
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox fechavence 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   180
      Top             =   7250
      Width           =   1335
   End
   Begin VB.TextBox cospaqu 
      Enabled         =   0   'False
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   176
      Top             =   6120
      Width           =   1335
   End
   Begin VB.TextBox cospaqp 
      Enabled         =   0   'False
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   175
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox cospaqi 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   174
      Top             =   6850
      Width           =   1335
   End
   Begin VB.ComboBox remate 
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
      Left            =   13000
      Style           =   2  'Dropdown List
      TabIndex        =   170
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox costoini 
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   167
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox minimo 
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
      MaxLength       =   10
      TabIndex        =   166
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox maximo 
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
      TabIndex        =   165
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox detraccion 
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
      TabIndex        =   163
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox pm10 
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
      Left            =   9900
      MaxLength       =   10
      TabIndex        =   161
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox pm9 
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
      Left            =   9900
      MaxLength       =   10
      TabIndex        =   160
      Top             =   6360
      Width           =   855
   End
   Begin VB.TextBox pm8 
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
      Left            =   9900
      MaxLength       =   10
      TabIndex        =   159
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox pm7 
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
      Left            =   9900
      MaxLength       =   10
      TabIndex        =   158
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox pm6 
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
      Left            =   9900
      MaxLength       =   10
      TabIndex        =   157
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox pm5 
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
      Left            =   9900
      MaxLength       =   10
      TabIndex        =   156
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox pm4 
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
      Left            =   9900
      MaxLength       =   10
      TabIndex        =   155
      Top             =   4560
      Width           =   855
   End
   Begin VB.TextBox pm3 
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
      Left            =   9900
      MaxLength       =   10
      TabIndex        =   154
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox pm2 
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
      Left            =   9900
      MaxLength       =   10
      TabIndex        =   153
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox pm1 
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
      Left            =   9900
      MaxLength       =   10
      TabIndex        =   152
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox ivap 
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
      MaxLength       =   10
      TabIndex        =   150
      Top             =   3960
      Width           =   855
   End
   Begin VB.Data Data9 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   147
      Top             =   8550
      Width           =   3375
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Lotes"
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
      Left            =   2280
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   146
      ToolTipText     =   "Ayuda"
      Top             =   9240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Series"
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
      Left            =   1560
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   145
      ToolTipText     =   "Ayuda"
      Top             =   9240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox local2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   8940
      Style           =   2  'Dropdown List
      TabIndex        =   143
      Top             =   2760
      Width           =   615
   End
   Begin VB.ComboBox monedav 
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
      Left            =   7380
      Style           =   2  'Dropdown List
      TabIndex        =   130
      Top             =   2760
      Width           =   735
   End
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
      Left            =   6300
      MaxLength       =   6
      TabIndex        =   129
      Top             =   3480
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
      Left            =   7140
      MaxLength       =   4
      TabIndex        =   128
      Top             =   3480
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
      Left            =   7860
      MaxLength       =   10
      TabIndex        =   127
      Top             =   3480
      Width           =   1215
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
      Left            =   9060
      MaxLength       =   10
      TabIndex        =   126
      Top             =   3480
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
      Left            =   12180
      MaxLength       =   10
      TabIndex        =   125
      Top             =   5280
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
      Left            =   12180
      MaxLength       =   10
      TabIndex        =   124
      Top             =   5640
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
      Left            =   12180
      MaxLength       =   10
      TabIndex        =   123
      Top             =   6120
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
      Left            =   12180
      MaxLength       =   10
      TabIndex        =   122
      Top             =   6480
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
      Left            =   12780
      MaxLength       =   10
      TabIndex        =   121
      Top             =   3480
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
      Left            =   11700
      MaxLength       =   10
      TabIndex        =   120
      Top             =   3480
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
      Left            =   11220
      MaxLength       =   5
      TabIndex        =   119
      Top             =   3480
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
      Left            =   10740
      MaxLength       =   5
      TabIndex        =   118
      Top             =   3480
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
      Left            =   12180
      MaxLength       =   10
      TabIndex        =   117
      Top             =   6840
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
      Left            =   9060
      MaxLength       =   10
      TabIndex        =   116
      Top             =   3840
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
      Left            =   7860
      MaxLength       =   10
      TabIndex        =   115
      Top             =   3840
      Width           =   1215
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
      Left            =   7140
      MaxLength       =   4
      TabIndex        =   114
      Top             =   3840
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
      Left            =   6300
      MaxLength       =   6
      TabIndex        =   113
      Top             =   3840
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
      Left            =   9060
      MaxLength       =   10
      TabIndex        =   112
      Top             =   4200
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
      Left            =   7860
      MaxLength       =   10
      TabIndex        =   111
      Top             =   4200
      Width           =   1215
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
      Left            =   7140
      MaxLength       =   4
      TabIndex        =   110
      Top             =   4200
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
      Left            =   6300
      MaxLength       =   6
      TabIndex        =   109
      Top             =   4200
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
      Left            =   9060
      MaxLength       =   10
      TabIndex        =   108
      Top             =   4560
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
      Left            =   7860
      MaxLength       =   10
      TabIndex        =   107
      Top             =   4560
      Width           =   1215
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
      Left            =   7140
      MaxLength       =   4
      TabIndex        =   106
      Top             =   4560
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
      Left            =   6300
      MaxLength       =   6
      TabIndex        =   105
      Top             =   4560
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
      Left            =   9060
      MaxLength       =   10
      TabIndex        =   104
      Top             =   4920
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
      Left            =   7860
      MaxLength       =   10
      TabIndex        =   103
      Top             =   4920
      Width           =   1215
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
      Left            =   7140
      MaxLength       =   4
      TabIndex        =   102
      Top             =   4920
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
      Left            =   6300
      MaxLength       =   6
      TabIndex        =   101
      Top             =   4920
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
      Left            =   9060
      MaxLength       =   10
      TabIndex        =   100
      Top             =   5280
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
      Left            =   7860
      MaxLength       =   10
      TabIndex        =   99
      Top             =   5280
      Width           =   1215
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
      Left            =   7140
      MaxLength       =   4
      TabIndex        =   98
      Top             =   5280
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
      Left            =   6300
      MaxLength       =   6
      TabIndex        =   97
      Top             =   5280
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
      Left            =   9060
      MaxLength       =   10
      TabIndex        =   96
      Top             =   5640
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
      Left            =   7860
      MaxLength       =   10
      TabIndex        =   95
      Top             =   5640
      Width           =   1215
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
      Left            =   7140
      MaxLength       =   4
      TabIndex        =   94
      Top             =   5640
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
      Left            =   6300
      MaxLength       =   6
      TabIndex        =   93
      Top             =   5640
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
      Left            =   9060
      MaxLength       =   10
      TabIndex        =   92
      Top             =   6000
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
      Left            =   7860
      MaxLength       =   10
      TabIndex        =   91
      Top             =   6000
      Width           =   1215
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
      Left            =   7140
      MaxLength       =   4
      TabIndex        =   90
      Top             =   6000
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
      Left            =   6300
      MaxLength       =   6
      TabIndex        =   89
      Top             =   6000
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
      Left            =   9060
      MaxLength       =   10
      TabIndex        =   88
      Top             =   6360
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
      Left            =   7860
      MaxLength       =   10
      TabIndex        =   87
      Top             =   6360
      Width           =   1215
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
      Left            =   7140
      MaxLength       =   4
      TabIndex        =   86
      Top             =   6360
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
      Left            =   6300
      MaxLength       =   6
      TabIndex        =   85
      Top             =   6360
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
      Left            =   9060
      MaxLength       =   10
      TabIndex        =   84
      Top             =   6720
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
      Left            =   7860
      MaxLength       =   10
      TabIndex        =   83
      Top             =   6720
      Width           =   1215
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
      Left            =   7140
      MaxLength       =   4
      TabIndex        =   82
      Top             =   6720
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
      Left            =   6300
      MaxLength       =   6
      TabIndex        =   81
      Top             =   6720
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
      Left            =   10740
      MaxLength       =   5
      TabIndex        =   80
      Top             =   3840
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
      Left            =   11220
      MaxLength       =   5
      TabIndex        =   79
      Top             =   3840
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
      Left            =   11700
      MaxLength       =   10
      TabIndex        =   78
      Top             =   3840
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
      Left            =   12780
      MaxLength       =   10
      TabIndex        =   77
      Top             =   3840
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
      Left            =   10740
      MaxLength       =   5
      TabIndex        =   76
      Top             =   4200
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
      Left            =   11220
      MaxLength       =   5
      TabIndex        =   75
      Top             =   4200
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
      Left            =   11700
      MaxLength       =   10
      TabIndex        =   74
      Top             =   4200
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
      Left            =   12780
      MaxLength       =   10
      TabIndex        =   73
      Top             =   4200
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
      Left            =   10740
      MaxLength       =   5
      TabIndex        =   72
      Top             =   4560
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
      Left            =   11220
      MaxLength       =   5
      TabIndex        =   71
      Top             =   4560
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
      Left            =   11700
      MaxLength       =   10
      TabIndex        =   70
      Top             =   4560
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
      Left            =   12780
      MaxLength       =   10
      TabIndex        =   69
      Top             =   4560
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
      Left            =   10740
      MaxLength       =   5
      TabIndex        =   68
      Top             =   4920
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
      Left            =   11220
      MaxLength       =   5
      TabIndex        =   67
      Top             =   4920
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
      Left            =   11700
      MaxLength       =   10
      TabIndex        =   66
      Top             =   4920
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
      Left            =   12780
      MaxLength       =   10
      TabIndex        =   65
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox flete 
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
      MaxLength       =   10
      TabIndex        =   62
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox percepcion 
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
      MaxLength       =   1
      TabIndex        =   61
      Top             =   3960
      Width           =   375
   End
   Begin VB.ComboBox estado 
      BackColor       =   &H00C0FFFF&
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
      Left            =   8940
      Style           =   2  'Dropdown List
      TabIndex        =   59
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Precios"
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
      Left            =   120
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Ayuda"
      Top             =   9240
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13980
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox xproveedor 
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
      Left            =   9540
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Top             =   1080
      Width           =   3975
   End
   Begin VB.CheckBox insumo 
      BackColor       =   &H00FFFF00&
      Caption         =   "Ingresar solo Insumos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   56
      Top             =   9360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox fabrica 
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
      Left            =   11940
      MaxLength       =   11
      TabIndex        =   55
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox codigo 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox costop 
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   52
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox costou 
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   50
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox factor 
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
      Left            =   4680
      MaxLength       =   6
      TabIndex        =   49
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox unidad 
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
      Left            =   4680
      MaxLength       =   6
      TabIndex        =   47
      Top             =   5040
      Width           =   1335
   End
   Begin VB.ComboBox monedac 
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
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   2760
      Width           =   615
   End
   Begin VB.ComboBox vecaja 
      BackColor       =   &H00C0FFFF&
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
      Left            =   11580
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   1920
      Width           =   495
   End
   Begin VB.ComboBox oferta 
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
      Left            =   11580
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   1560
      Width           =   495
   End
   Begin VB.ComboBox servicio 
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
      Left            =   10260
      Style           =   2  'Dropdown List
      TabIndex        =   39
      Top             =   1560
      Width           =   615
   End
   Begin VB.ComboBox serie 
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
      Left            =   8940
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   1560
      Width           =   615
   End
   Begin VB.ComboBox vtaund 
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
      Left            =   10260
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   1920
      Width           =   615
   End
   Begin VB.ComboBox peso 
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
      Left            =   8940
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox comision 
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
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   31
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox pesokgr 
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
      MaxLength       =   10
      TabIndex        =   29
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox isc 
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
      MaxLength       =   5
      TabIndex        =   27
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox igv 
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
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   25
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox color 
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
      Left            =   7380
      MaxLength       =   6
      TabIndex        =   22
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox lineatalla 
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
      Left            =   7380
      MaxLength       =   6
      TabIndex        =   21
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox categoria 
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
      Left            =   7380
      MaxLength       =   6
      TabIndex        =   20
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox marca 
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
      Left            =   7380
      MaxLength       =   6
      TabIndex        =   19
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox seccion 
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
      Left            =   7380
      MaxLength       =   6
      TabIndex        =   17
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox subfamilia 
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
      Left            =   4680
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox familia 
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
      Left            =   4680
      MaxLength       =   6
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox presenta 
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
      Left            =   1200
      MaxLength       =   60
      TabIndex        =   4
      Top             =   2280
      Width           =   4815
   End
   Begin VB.TextBox descorto 
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
      Left            =   1200
      MaxLength       =   22
      TabIndex        =   3
      Top             =   1920
      Width           =   4815
   End
   Begin VB.TextBox barras 
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
      Left            =   1200
      MaxLength       =   30
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox descripcio 
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
      Left            =   1215
      MaxLength       =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4815
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
      Left            =   80
      MaskColor       =   &H00E0E0E0&
      Picture         =   "TPRODUCT.frx":368A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Grabar registro"
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
      Left            =   5760
      Picture         =   "TPRODUCT.frx":489C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Nuevo registro"
      Top             =   -15
      Visible         =   0   'False
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
      Left            =   850
      MaskColor       =   &H00E0E0E0&
      Picture         =   "TPRODUCT.frx":5AAE
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      MaskColor       =   &H00E0E0E0&
      Picture         =   "TPRODUCT.frx":6CC0
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Consulta"
      Top             =   9240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox diasalerta 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   4680
      MaxLength       =   3
      TabIndex        =   231
      Top             =   7630
      Width           =   1335
   End
   Begin VB.CommandButton cmdCommand7 
      BackColor       =   &H00004040&
      Caption         =   "Más..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12690
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   232
      Top             =   7830
      Width           =   630
   End
   Begin VB.CommandButton cmdBuscar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5535
      Picture         =   "TPRODUCT.frx":7ED2
      Style           =   1  'Graphical
      TabIndex        =   233
      ToolTipText     =   "Buscar familia"
      Top             =   705
      Width           =   645
   End
   Begin ChamaleonButton.ChameleonBtn Mas 
      Height          =   465
      Left            =   7800
      TabIndex        =   245
      ToolTipText     =   "Más Opciones"
      Top             =   240
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   820
      BTYPE           =   4
      TX              =   "+"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "TPRODUCT.frx":8C34
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox obligacomentario 
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
      Left            =   12900
      Style           =   2  'Dropdown List
      TabIndex        =   246
      Top             =   2650
      Width           =   615
   End
   Begin VB.Label lblCostoUlt 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dias Alerta"
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
      TabIndex        =   235
      Top             =   7650
      Width           =   1125
   End
   Begin VB.Label lblDescrCorto 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  F1: Buscar           F7: Agregar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   234
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label34 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Actualiza Costo de Receta?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10755
      TabIndex        =   210
      Top             =   240
      Width           =   1530
   End
   Begin VB.Label Label31 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UnidadProd       Factorpro"
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
      Left            =   3600
      TabIndex        =   208
      Top             =   8160
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Servicio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   204
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CostoAnt-1    CostoAnt-2"
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
      Left            =   180
      TabIndex        =   200
      Top             =   5490
      Width           =   1095
   End
   Begin VB.Label Label46 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cola Impresion"
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
      Left            =   6300
      TabIndex        =   197
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Label Label84 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Impresora Despacho                                        Grupo orden Despacho"
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
      Left            =   6300
      TabIndex        =   192
      Top             =   7800
      Width           =   2295
   End
   Begin VB.Label Label80 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RecetaPrn           Proveedor"
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
      Left            =   8340
      TabIndex        =   191
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label87 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Receta Para         Unidades"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   180
      TabIndex        =   188
      Top             =   6210
      Width           =   1095
   End
   Begin VB.Label Label79 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Borrar Imagen"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   184
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label75 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rentabilidad                  Orden Touch"
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
      Left            =   9540
      TabIndex        =   182
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label73 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CostoUlt."
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
      TabIndex        =   179
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label72 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CostoProm."
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
      TabIndex        =   178
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label Label71 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CostoPaq            Fechavence dd/mm/yyyy             Días Alerta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3600
      TabIndex        =   177
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label Label70 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Costo x Empaque"
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
      Left            =   3600
      TabIndex        =   173
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label69 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Empaque"
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
      Left            =   3600
      TabIndex        =   172
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label68 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remate                    Fuel"
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
      Left            =   12060
      TabIndex        =   171
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label67 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Costo Unitario"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3600
      TabIndex        =   169
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label65 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CostoInicial"
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
      TabIndex        =   168
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label62 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Detrac.  Maximo"
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
      Left            =   1920
      TabIndex        =   164
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label61 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%MiniPre"
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
      Left            =   9900
      TabIndex        =   162
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label56 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "%Ivap         Minimo          %DsctRef"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   151
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label ventanas 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   149
      Top             =   7200
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label ordename 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   148
      Top             =   7080
      Width           =   105
   End
   Begin VB.Label Label33 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lista"
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
      Left            =   8340
      TabIndex        =   144
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label36 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mon.Pvta"
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
      Left            =   6300
      TabIndex        =   142
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unidad"
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
      Left            =   6300
      TabIndex        =   141
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Left            =   7140
      TabIndex        =   140
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "P.Venta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7860
      TabIndex        =   139
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label40 
      BackColor       =   &H00C0C0C0&
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
      Left            =   9060
      TabIndex        =   138
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label41 
      BackColor       =   &H00C0C0C0&
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
      Left            =   10740
      TabIndex        =   137
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label42 
      BackColor       =   &H00C0C0C0&
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
      Left            =   11220
      TabIndex        =   136
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label43 
      BackColor       =   &H00C0C0C0&
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
      Left            =   11700
      TabIndex        =   135
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label44 
      BackColor       =   &H00C0C0C0&
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
      Left            =   12780
      TabIndex        =   134
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label45 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaInicio                  FechaFinal"
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
      Left            =   10740
      TabIndex        =   133
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label48 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaInicio                  FechaFinal                   Descuento Fijo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10740
      TabIndex        =   132
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label54 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Oferta.precio=0 acepta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9405
      TabIndex        =   131
      Top             =   7230
      Width           =   1635
   End
   Begin VB.Label Label57 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Foto"
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
      TabIndex        =   64
      Top             =   6840
      Width           =   3375
   End
   Begin VB.Label Label35 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CostoProm."
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
      TabIndex        =   63
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label32 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Activo"
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
      Left            =   8340
      TabIndex        =   60
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label fotonombre 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
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
      Height          =   255
      Left            =   3630
      TabIndex        =   54
      Top             =   8805
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image foto 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   120
      Stretch         =   -1  'True
      Top             =   7230
      Width           =   3375
   End
   Begin VB.Label paridad 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
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
      Height          =   255
      Left            =   7860
      TabIndex        =   53
      Top             =   7080
      Width           =   120
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CostoUltimo"
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
      TabIndex        =   51
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unidad                  Factor"
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
      Left            =   3600
      TabIndex        =   48
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label26 
      BackColor       =   &H00C0C0C0&
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
      Left            =   3600
      TabIndex        =   46
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label24 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Oferta        VeCaja"
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
      Left            =   10860
      TabIndex        =   42
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label23 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Servicio"
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
      Left            =   9540
      TabIndex        =   40
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
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
      Left            =   8340
      TabIndex        =   38
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "VtaUnidad"
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
      Left            =   9540
      TabIndex        =   36
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Peso"
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
      Left            =   8340
      TabIndex        =   34
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Comisiones"
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
      TabIndex        =   32
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PesoKgr         Flete"
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
      Left            =   120
      TabIndex        =   30
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Isc          Percep"
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
      Left            =   1920
      TabIndex        =   28
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Igv"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1920
      TabIndex        =   26
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fabricante"
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
      Left            =   10740
      TabIndex        =   24
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LineaTalla               Color"
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
      Left            =   6300
      TabIndex        =   23
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seccion                 Marca                       Categoria"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6300
      TabIndex        =   18
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SubFamilia"
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
      TabIndex        =   16
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Familia"
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
      TabIndex        =   15
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descr.Corto      Presentac."
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
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cod.Barras"
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
      TabIndex        =   13
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion"
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
      TabIndex        =   12
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producto"
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
      TabIndex        =   11
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DelivAutomat                         Comentario"
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
      Left            =   11595
      TabIndex        =   44
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Menu grba1 
      Caption         =   "&Grabar"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tproduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ArrBarCode(43) As String

Private Type campo_precio

    unidad As String
    factor As String
    precio As String
    costo As String
    margen As String
    stock As String

End Type

Dim campo_precios(12) As campo_precio

Dim tpventa           As Double

''17/07/2017 kenyo tienda ropa opciones producto
Dim subfamiliatr      As String

Dim marcatr           As String

Dim categoriatr       As String

Dim secciontr         As String

Dim colortr           As String

Dim tallatr           As String

''17/07/2017 kenyo tienda ropa opciones producto

Private Sub ajdu1_Click()

End Sub

Sub xadicion()

    Dim found As Integer

    'If Frame4.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    inicializa
    codigo = ""
    busca_correlativo 0
    'codigo.SetFocus
    descripcio.SetFocus

End Sub

Private Sub Acepta_Click()
    BuscaCodigo
 
End Sub

Private Sub barras_Change()
    hacer_barras

End Sub

Private Sub Barras_KeyPress(KeyAscii As Integer)

    Dim buf   As String

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    buf = convierte_barras("" & Barras)

    If Len(buf) > 0 Then
        found = valida_barras("" & buf)

        If found = 1 Then
            Barras.SetFocus
            Exit Sub

        End If

    End If

    descripcio.SetFocus

End Sub

Private Sub barras_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        'codigo.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        Label2_Click

    End If

    ''16/08/2017 kenyo tienda ropa codigo de barras
    If KeyCode = &H76 Then  'F7
        If Barras = "" Then
  
            If Len(codigo) = "1" Then
                Barras = subfamilia + marca + "0000" + codigo
            ElseIf Len(codigo) = "2" Then
                Barras = subfamilia + marca + "000" + codigo
            ElseIf Len(codigo) = "3" Then
                Barras = subfamilia + marca + "00" + codigo
            ElseIf Len(codigo) = "4" Then
                Barras = subfamilia + marca + "0" + codigo
            ElseIf Len(codigo) = "5" Then
                Barras = subfamilia + marca + codigo

            End If
  
            If Len(talla) = "2" Then
                Barras = Barras + "0" + talla
            ElseIf Len(talla) = "3" Then
                Barras = Barras + talla

            End If
      
        Else
            Exit Sub

        End If

    End If

    ''16/08/2017 kenyo tienda ropa codigo de barras

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Command5_Click

End Sub

Private Sub cboprinters_Click()

    If cboprinters.Text <> "%" Then
        puertoimpresion = "" & cboprinters.Text
  
        ' 05/06/207   grupoimpresion ''
        grupoimpresion = Left(cboprinters.Text, 1)
   
        ''11/07/2017 kenyo edicion producto cola
        cola = "S"
        ''11/07/2017 kenyo edicion producto cola
   
    End If

End Sub

Private Sub cboprinters1_Click()

    If cboprinters1.Text <> "%" Then
        puertoimpresion1 = "" & cboprinters1.Text

    End If

End Sub

Private Sub cboprinters2_Click()

    If cboprinters2.Text <> "%" Then
        puertoimpresion2 = "" & cboprinters2.Text

    End If

End Sub

Private Sub cboprinters3_Click()

    If cboprinters3.Text <> "%" Then
        puertoimpresion3 = "" & cboprinters3.Text

    End If

End Sub

Private Sub CmdAceptar_Click()
    familia = dbGrid1.columns(1)
    Frame1.Visible = False
    Frame1.Enabled = False
    familia.SetFocus
    familia_KeyPress 13

End Sub

Private Sub Cerrar_Click()
    Frame1.Visible = False

End Sub

Private Sub cmdBuscar_Click()
    'If Len(familia) = 0 Then
    consulta_Familia
    ' Exit Sub
    'End If

    '' 30/11/2017 Correción  General del Sistema Parte I
    'subfamilia.SetFocus
    '' 30/11/2017 Correción  General del Sistema Parte I

End Sub

Private Sub cmdCommand7_Click()
    Frame3.Visible = True

End Sub

Private Sub Command2_Click()

    On Error GoTo cmd432_err

    cn.Execute ("delete from productb where producto='" & Trim(codigo) & "' and barras='" & DBGrid2.columns("barras") & "'")
    Label2_Click
    Exit Sub
cmd432_err:
    Exit Sub

End Sub

Private Sub diasalerta_KeyPress(KeyAscii As Integer)

    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0

    End If

End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label2_Click()

    Dim found     As Integer

    Dim rconsulta As New ADODB.Recordset

    If Len(Trim(codigo)) = 0 Then Exit Sub
    found = busca_registro()

    If found = 0 Then

        '   Barras.SetFocus
        '
        '   Exit Sub
    End If

    Frame2.Visible = True
    Frame2.Caption = "CODIGO BARRAS"
    barras2 = ""

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open "Select Barras,producto from Productb where producto='" & Trim(codigo) & "'", cn, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = rconsulta
    DBGrid2.columns(0).Width = 3500
    DBGrid2.columns(1).Width = 1500

    If rconsulta.EOF = True And rconsulta.BOF = True Then
        Exit Sub

    End If

    barras2.SetFocus

End Sub

Private Sub barras2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    Command3_Click

End Sub

Private Sub bo712_Click()

End Sub

Private Sub buffer1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 And Combo3 = "Proveedor" Then 'f1
        consulta_proveedor

    End If

End Sub

Private Sub categoria_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    lineatalla.SetFocus

End Sub

Private Sub categoria_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        marca.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_categoria

    End If

    If KeyCode = &H76 Then  'f7
        tcategor.Show 1

    End If

End Sub

Private Sub cmdAddEntry_Click()

    If ordename = "NUEVO" Then
        inicializa
        codigo = ""
        'codigo.SetFocus
        descripcio.SetFocus
        Exit Sub

    End If

    If ordename = "MODIFICA" Then
        Barras.SetFocus

    End If

End Sub

Private Sub cmdDelete_Click()

End Sub

Private Sub cmdExit_Click()
    dlo132_Click

End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdSave_Click()
    grba1_Click

End Sub

Private Sub cmdSort_Click()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Producto"
    Combo3.AddItem "Descripcio"
    Combo3.AddItem "familia"
    Combo3.AddItem "Subfamilia"
    Combo3.AddItem "Categoria"
    Combo3.AddItem "Seccion"
    Combo3.AddItem "Color"
    Combo3.AddItem "Proveedor"
    Combo3.ListIndex = 0
    opcion1 = "1"
    Frame1.Visible = True
    buffer = ""
    found = ejecuta(0)
    dbGrid1.SetFocus

End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    descripcio.SetFocus

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1

        'cmdSort_Click
    End If

End Sub

Private Sub color_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    fabrica.SetFocus

End Sub

Private Sub color_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        lineatalla.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_color

    End If

    If KeyCode = &H76 Then  'f7
        tncolor.Show 1

    End If

End Sub

Private Sub comision_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    pesokgr.SetFocus

End Sub

Private Sub comision_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        vecaja.SetFocus
        Exit Sub

    End If

    '' 29/01/2018 Comisiones por producto por trabajador. Proyectos requerimientos Spa Cañete.
    If KeyCode = &H76 Then  'f7
        FrmComisiones.producto = codigo
        FrmComisiones.descripcion = descripcio
        FrmComisiones.Show 1

    End If

    '' 29/01/2018 Comisiones por producto por trabajador. Proyectos requerimientos Spa Cañete.

End Sub

Private Sub Command10_Click()

End Sub

Private Sub Command11_Click()

End Sub

Private Sub Command12_Click()

    Dim found As Integer

    If local2.Visible <> True Then Exit Sub  'si no es precios x locales

    'If Frame4.Visible = True Then Exit Sub
    If Frame2.Visible = True Then
        borrar_barras
        'Label2_Click
        Exit Sub

    End If

    If Frame1.Visible = True Then Exit Sub
    found = busca_registro()

    If found = 0 Then
        MsgBox "No existe registro", 48, "Aviso"
        Exit Sub

    End If

    'tprecios.monedac = monedac
    'tprecios.unidad = unidad
    'tprecios.factor = factor
    'tprecios.costou = costou
    'tprecios.monedav = monedav
    'tprecios.producto = codigo
    'tprecios.descripcio = descripcio
    'tprecios.Show 1
    'found = busca_registro()
    'codigo.SetFocus
End Sub

Private Sub Command13_Click()

End Sub

Private Sub Command14_Click()

End Sub

Private Sub Command16_Click()

End Sub

Private Sub Command3_Click()

    Dim found As Integer

    Dim buf2  As String

    buf2 = ""

    If Len(barras2) = 0 Then
        barras2.SetFocus
        Exit Sub

    End If

    If Frame2.Caption = "LOTES" Then
        found = grabar_barras()
        barras2.SetFocus
        Exit Sub

    End If

    If Frame2.Caption = "NUMERO SERIES" Then
        found = grabar_barras()
        barras2.SetFocus
        Exit Sub

    End If

    found = valida_barras("" & barras2)

    If found = 1 Then
        'found = valida_barras20("" & barras2, buf2) 'si existe el codigo en la database
        'If found = 1 Then
        MsgBox "Ya existe Barras Ingresado " + buf2, 48, "Aviso"
        barras2 = ""
        barras2.SetFocus
        Exit Sub

    End If

    'buf2 = ""
    'found = valida_barras2("" & barras2, buf2) 'si existe la barra en producto
    'If found = 1 Then
    '   MsgBox "Ya existe Barras Ingresado " + buf2, 48, "Aviso"
    '   barras2 = ""
    '   barras2.SetFocus
    '   Exit Sub
    'End If
    found = grabar_barras()
    Label2_Click
    barras2.SetFocus

End Sub

Private Sub Command4_Click()
    dlo132_Click

End Sub

Private Sub Command5_Click()
    ejecuta 1

End Sub

Function ejecuta(sw As Integer)

    Dim buf       As String

    Dim indx      As Integer

    Dim rconsulta As New ADODB.Recordset

    On Error GoTo cmd34_err

    If opcion1 = "1" Then
   
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Producto,Unidad,Factor,Costou,Costop,Unidad1,Factor1,Pventa1,Proveedor1 from producto "
        Else
            buf = "select Descripcio,Producto,Unidad,Factor,Costou,Costop,Unidad1,Factor1,Pventa1,Proveedor1 from producto where "
            buf = buf & "" & Combo3 & " like '%" & buffer & "%'"

        End If

    End If

    If opcion1 = "2" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Familia from familia "
        Else
            buf = "select Descripcio,Familia from familia where " & "" & Combo3 & " like '%" & buffer & "%'"

        End If

    End If

    If opcion1 = "39" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Id from formulacion "
        Else
            buf = "select Descripcio,Id from formulacion where " & "" & Combo3 & " like '%" & buffer & "%'"

        End If

    End If

    If opcion1 = "190" Then  'local
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from tlocal "
        Else
            buf = "select Nombre,Codigo from tlocal where " & "" & Combo3 & " like '%" & buffer & "%'"
      
        End If

    End If

    If opcion1 = "3" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,SubFamilia,Familia from Subfamil where familia='" & familia & "'"
        Else
            buf = "select Descripcio,SubFamilia,Familia from Subfamil where familia='" & familia & "' and " & Combo3 & " like '%" & buffer & "%'"
      
        End If

    End If

    If opcion1 = "16" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,margen from margen "
        Else
            buf = "select Descripcio,margen from margen where " & "" & Combo3 & " like '%" & buffer & "%'"
      
        End If

    End If

    If opcion1 = "4" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Seccion from seccion "
        Else
            buf = "select Descripcio,Seccion from seccion where " & "" & Combo3 & " like '%" & buffer & "%'"
      
        End If

    End If

    If opcion1 = "5" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Marca from Marca "
        Else
            buf = "select Descripcio,Marca from Marca where " & "" & Combo3 & " like '%" & buffer & "%'"
      
        End If

    End If

    If opcion1 = "6" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from Fabrica "
        Else
            buf = "select Nombre,Codigo from Fabrica where " & "" & Combo3 & " like '%" & buffer & "%'"
      
        End If

    End If

    If opcion1 = "7" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,categoria from categori "
        Else
            buf = "select Descripcio,categoria from categori where " & "" & Combo3 & " like '%" & buffer & "%'"
      
        End If

    End If

    If opcion1 = "8" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Linea from Linea "
        Else
            buf = "select Descripcio,Linea from Linea where " & "" & Combo3 & " like '%" & buffer & "%'"
      
        End If

    End If

    ''17/07/2017 kenyo tienda ropa opciones producto
    If opcion1 = "50" Then ' talla tienda ropa
        If Len(buffer) = 0 Then
            buf = "select Descripcio,talla from talla "
        Else
            buf = "select Descripcio,talla from talla where " & "" & Combo3 & " like '%" & buffer & "%'"
      
        End If

    End If

    If opcion1 = "51" Then ' sexo
        If Len(buffer) = 0 Then
            buf = "select Descripcio,sexo from sexo "
        Else
            buf = "select Descripcio,sexo from sexo where " & "" & Combo3 & " like '%" & buffer & "%'"
      
        End If

    End If

    If opcion1 = "52" Then ' procedencia
        If Len(buffer) = 0 Then
            buf = "select Descripcio,procedencia from procedencia "
        Else
            buf = "select Descripcio,procedencia from procedencia where " & "" & Combo3 & " like '%" & buffer & "%'"
      
        End If

    End If

    If opcion1 = "53" Then ' proyecto
        If Len(buffer) = 0 Then
            buf = "select Descripcio,proyecto from proyecto "
        Else
            buf = "select Descripcio,proyecto from proyecto where " & "" & Combo3 & " like '%" & buffer & "%'"
      
        End If

    End If

    ''17/07/2017 kenyo tienda ropa opciones producto

    If opcion1 = "9" Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Color from Color "
        Else
            buf = "select Descripcio,Color from Color where " & "" & Combo3 & " like '%" & buffer & "%'"
      
        End If

    End If

    If opcion1 = "10" Or opcion1 = "101" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from proveedo "
        Else
            buf = "select Nombre,Codigo from proveedo where " & "" & Combo3 & " like '%" & buffer & "%'"
      
        End If

    End If

    If opcion1 = "11" Then
        If Len(buffer) = 0 Then
            buf = "select Nombre,Codigo from proveedo "
        Else
            buf = "select Nombre,Codigo from proveedo where " & "" & Combo3 & " like '%" & buffer & "%'"
      
        End If

    End If

    If opcion1 = "27" Or opcion1 = 28 Or opcion1 = 29 Or opcion1 = 30 Or opcion1 = 31 Then
        If Len(buffer) = 0 Then
            buf = "select Descripcio,Ccosto from ccosto "
        Else
            buf = "select Descripcio,Ccosto from ccosto where " & "" & Combo3 & " like '%" & buffer & "%'"
            indx = dbGrid1.Col

        End If

    End If

    'MsgBox buf
   
    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = rconsulta
   
    '24/06/2017 kenyo CORRECCION acepta-busqueda
    If rconsulta.RecordCount = 0 Then
        Acepta.Enabled = False
    Else
        Acepta.Enabled = True

    End If

    '24/06/2017 kenyo CORRECCION acepta-busqueda
   
    If rconsulta.EOF = True And rconsulta.BOF = True Then
        buffer = ""
        rconsulta.Close
        buffer.SetFocus
        Exit Function

    End If
   
    pone_tamano
   
    If sw = 1 Then
        dbGrid1.SetFocus

    End If

    ejecuta = 1
    Exit Function
cmd34_err:
    'MsgBox "Error en Consulta " & error$, 48, "Aviso"
    buffer = ""
    Exit Function

End Function

Private Sub Command6_Click()
    'dlo132_Click
    Frame3.Visible = False

End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command8_Click()

End Sub

Private Sub Command9_Click()
    dlo132_Click

End Sub

Private Sub costo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

    'fechauc.SetFocus
End Sub

Private Sub costo_KeyUp(KeyCode As Integer, Shift As Integer)
    'If KeyCode = &H26 Then
    '   rcodigo.SetFocus
    '   Exit Sub
    'End If

End Sub

Private Sub costop_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes
    monedav.SetFocus

End Sub

Private Sub costop_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        costou.SetFocus
        Exit Sub

    End If

End Sub

Private Sub costou_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes
    costop.SetFocus

End Sub

Private Sub costou_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        factor.SetFocus
        Exit Sub

    End If

End Sub

Private Sub dbgrid1_DblClick()
    'buffer1 = ""
    dbgrid1_KeyDown 13, 0

End Sub

Sub BuscaCodigo()

    Dim found As Integer

    ''14/06/2017 kenyo No se cuelga el Sistema al aceptar
    'If dbGrid1.Col Then
    '  Exit Sub
    'End If
    ''14/06/2017 kenyo No se cuelga el Sistema al aceptar

    If opcion1 = "1" Then
  
        codigo = Trim(dbGrid1.columns(1))
        Frame1.Visible = False
        Frame1.Enabled = False
        codigo.SetFocus
        codigo_KeyPress 13

    End If

    If opcion1 = "27" Then

        'ccosto = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'ccosto.SetFocus
        'ccosto_KeyPress 13
    End If

    If opcion1 = "190" Then

        'tlocal = DBGrid1.Columns(1)
        'found = busca_registro()
        'Frame1.Visible = False
        'tlocal.SetFocus
        'tlocal_KeyPress 13
    End If

    If opcion1 = "29" Then

        'xccosto2 = DBGrid1.Columns(1)
        'Frame1.Visible = False
        'xccosto2.SetFocus
        'xccosto2_KeyPress 13
    End If

    If opcion1 = "30" Then

        'xccosto3 = DBGrid1.Columns(1)
        'Frame1.Visible = False
        'xccosto3.SetFocus
        'xccosto3_KeyPress 13
    End If

    If opcion1 = "31" Then

        'xccosto4 = DBGrid1.Columns(1)
        'Frame1.Visible = False
        'xccosto4.SetFocus
        'xccosto4_KeyPress 13
    End If

    If opcion1 = "2" Then ' familia
        familia = dbGrid1.columns(1)
        Frame1.Visible = False
        Frame1.Enabled = False
        familia.SetFocus
        familia_KeyPress 13

    End If

    If opcion1 = "39" Then
        'formulacion = dbGrid1.columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'formulacion.SetFocus
      
    End If
      
    If opcion1 = "3" Then
        subfamilia = dbGrid1.columns(1)
        Frame1.Visible = False
        Frame1.Enabled = False
        subfamilia.SetFocus
        subfamilia_KeyPress 13

    End If

    If opcion1 = "4" Then
        seccion = dbGrid1.columns(1)
        Frame1.Visible = False
        Frame1.Enabled = False
        seccion.SetFocus
        seccion_KeyPress 13

    End If

    If opcion1 = "16" Then
        margen = Trim(dbGrid1.columns(1))
        Frame1.Visible = False
        Frame1.Enabled = False
        margen.SetFocus
        margen_KeyPress 13

    End If
   
    If opcion1 = "5" Then
        marca = dbGrid1.columns(1)
        Frame1.Visible = False
        Frame1.Enabled = False
        marca.SetFocus
        marca_KeyPress 13

    End If

    If opcion1 = "6" Then
        fabrica = dbGrid1.columns(1)
        Frame1.Visible = False
        Frame1.Enabled = False
        fabrica.SetFocus
        fabrica_KeyPress 13

    End If

    If opcion1 = "7" Then
        categoria = dbGrid1.columns(1)
        Frame1.Visible = False
        Frame1.Enabled = False
        categoria.SetFocus
        categoria_KeyPress 13

    End If

    If opcion1 = "8" Then
        lineatalla = dbGrid1.columns(1)
        Frame1.Visible = False
        Frame1.Enabled = False
        lineatalla.SetFocus
        lineatalla_KeyPress 13

    End If

    If opcion1 = "9" Then
        color = dbGrid1.columns(1)
        Frame1.Visible = False
        Frame1.Enabled = False
        color.SetFocus
        color_KeyPress 13

    End If

    If opcion1 = "10" Then
        fabrica = dbGrid1.columns(1)
        Frame1.Visible = False
        Frame1.Enabled = False
        fabrica.SetFocus
        fabrica_KeyPress 13

    End If
   
    ''17/07/2017 kenyo tienda ropa opciones producto
    If opcion1 = "50" Then
        talla = dbGrid1.columns(1)
        Frame1.Visible = False
        Frame1.Enabled = False
        frmOtros.Visible = True
        talla.SetFocus
        talla_KeyPress 13

    End If
   
    If opcion1 = "51" Then
        sexo = dbGrid1.columns(1)
        Frame1.Visible = False
        Frame1.Enabled = False
        frmOtros.Visible = True
        sexo.SetFocus
        sexo_KeyPress 13

    End If

    If opcion1 = "52" Then
        procedencia = dbGrid1.columns(1)
        Frame1.Visible = False
        Frame1.Enabled = False
        frmOtros.Visible = True
        procedencia.SetFocus
      
        procedencia_KeyPress 13

    End If

    If opcion1 = "53" Then
        proyecto = dbGrid1.columns(1)
        Frame1.Visible = False
        Frame1.Enabled = False
        frmOtros.Visible = True
        proyecto.SetFocus
        proyecto_KeyPress 13

    End If
   
    ''17/07/2017 kenyo tienda ropa opciones producto
   
    If opcion1 = "101" Then

        'pcodigo = Trim(dbGrid1.columns(1))
        'pncodigo = Trim(dbGrid1.columns(0))
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'pcodigo.SetFocus
        'pcodigo_KeyPress 13
    End If

    If opcion1 = "11" Then

        'proveedor2 = DBGrid1.Columns(1)
        'Frame1.Visible = False
        'proveedor2.SetFocus
        'proveedor2_KeyPress 13
    End If
   
End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo cmd4323_err

    If KeyCode = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If KeyCode = &H70 Then
        If opcion1 = "1" Then
            carga_dbgrid4

        End If

        Exit Sub

    End If

    If KeyCode = 13 Then
        BuscaCodigo

    End If

    Exit Sub
cmd4323_err:
    Exit Sub

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)

    Dim buf   As String

    Dim buf2  As String

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

        If Chr(KeyAscii) = "/" Then
            buf = ""
            buffer = buf

        End If

        If KeyAscii <> 13 Then
            buffer = buffer + buf

        End If

        buf = buffer
        found = ejecuta(0)

        If found = 0 Then
            ejecuta (1)

        End If

    End If

End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        dlo132_Click
        Exit Sub

    End If

End Sub

Private Sub DBGrid3_DblClick()

End Sub

Private Sub DBGrid4_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub DBGrid5_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub DBGrid6_KeyUp(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub DBGrid6_UnboundReadData(ByVal RowBuf As MSDBGrid.RowBuffer, _
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

Private Sub descorto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If Len(descorto) = 0 Then
        descorto = Mid$(descripcio, 1, 22)

    End If

    familia.SetFocus

End Sub

Private Sub descorto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        descripcio.SetFocus
        Exit Sub

    End If

End Sub

Private Sub descripcio_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If Len(descripcio) = 0 Then Exit Sub
    descorto.SetFocus

End Sub

''17/07/2017 kenyo tienda ropa opciones producto

Sub busca_datos()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM subfamil where  subfamilia='" & subfamilia & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        subfamiliatr = mytablex.Fields("descripcio")

    End If

    mytablex.Close

    mytablex.Open "SELECT * FROM marca where  marca='" & marca & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        marcatr = mytablex.Fields("descripcio")

    End If

    mytablex.Close

    mytablex.Open "SELECT * FROM CATEGORI where  categoria='" & categoria & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        categoriatr = mytablex.Fields("descripcio")

    End If

    mytablex.Close

    mytablex.Open "SELECT * FROM seccion where  seccion='" & seccion & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        secciontr = mytablex.Fields("descripcio")

    End If

    mytablex.Close

    mytablex.Open "SELECT * FROM color where  color='" & color & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        colortr = mytablex.Fields("descripcio")

    End If

    mytablex.Close

    mytablex.Open "SELECT * FROM talla where  talla='" & talla & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        tallatr = mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Sub

''17/07/2017 kenyo tienda ropa opciones producto

Private Sub descripcio_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        Barras.SetFocus
        Exit Sub

    End If

    ''17/07/2017 kenyo tienda ropa opciones producto
    If KeyCode = &H76 Then 'F7

        busca_datos

        If descripcio = "" Then
            descripcio = subfamiliatr + " " + marcatr + " " + categoriatr + " " + secciontr + " " + colortr + " " + tallatr
        
        Else
            Exit Sub

        End If

    End If

    ''17/07/2017 kenyo tienda ropa opciones producto

End Sub

Private Sub djuer1_Click()

End Sub

Private Sub dlo132_Click()

    'If Frame5.Visible = True Then
    'Frame5.Visible = False
    'dbGrid1.SetFocus
    'Exit Sub
    'End If
    If Frame1.Visible = True Then

        'If Frame10.Visible = True Then
        '   Frame10.Visible = False
        '   buffer1.SetFocus
        '   Exit Sub
        'End If
    End If

    'If Frame9.Visible = True Then
    '   Frame9.Visible = False
    '   dbgrid4.SetFocus
    '   Exit Sub
    'End If
    'If Frame4.Visible = True Then
    '   If Frame9.Visible = True Then
    '      Frame9.Visible = False
    '      dbgrid4.SetFocus
    '      Exit Sub
    '   End If
    '   Frame4.Visible = False
    '   fabrica.SetFocus
    '   Exit Sub
    'End If

    If Frame2.Visible = True Then
        Frame2.Visible = False

        If Frame2.Caption = "LOTES" Or Frame2.Caption = "NUMERO SERIES" Then
            codigo.SetFocus
            Exit Sub

        End If

        Barras.SetFocus
        Exit Sub

    End If

    If Frame1.Visible = True Then
        If opcion1 = "1" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            codigo.SetFocus
            Exit Sub

        End If

        If opcion1 = "27" Then
   
            'Frame1.Visible = False
            'Frame1.Enabled = False
            'ccosto.SetFocus
            'Exit Sub
        End If

        If opcion1 = "28" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            vecaja.SetFocus
            Exit Sub

        End If

        If opcion1 = "29" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            vecaja.SetFocus
            Exit Sub

        End If

        If opcion1 = "30" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            vecaja.SetFocus
            Exit Sub

        End If

        If opcion1 = "31" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            vecaja.SetFocus
            Exit Sub

        End If

        If opcion1 = "2" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            familia.SetFocus
            Exit Sub

        End If

        If opcion1 = "39" Then

            'Frame1.Visible = False
            'Frame1.Enabled = False
            'formulacion.SetFocus
            'Exit Sub
        End If

        If opcion1 = "3" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            subfamilia.SetFocus
            Exit Sub

        End If

        If opcion1 = "4" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            seccion.SetFocus
            Exit Sub

        End If

        If opcion1 = "16" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            margen.SetFocus
            Exit Sub

        End If

        If opcion1 = "5" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            marca.SetFocus
            Exit Sub

        End If

        If opcion1 = "6" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            fabrica.SetFocus
            Exit Sub

        End If

        If opcion1 = "7" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            categoria.SetFocus
            Exit Sub

        End If

        ''17/07/2017 kenyo tienda ropa opciones producto
        If opcion1 = "50" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            talla.SetFocus
            Exit Sub

        End If

        If opcion1 = "51" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            sexo.SetFocus
            Exit Sub

        End If

        If opcion1 = "52" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            procedencia.SetFocus
            Exit Sub

        End If

        If opcion1 = "53" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            proyecto.SetFocus
            Exit Sub

        End If

        ''17/07/2017 kenyo tienda ropa opciones producto

        If opcion1 = "8" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            lineatalla.SetFocus
            Exit Sub

        End If

        If opcion1 = "9" Then
   
            Frame1.Visible = False
            Frame1.Enabled = False
            color.SetFocus
            Exit Sub

        End If

        If opcion1 = "10" Then
            Frame1.Visible = False
            Frame1.Enabled = False
            fabrica.SetFocus
            Exit Sub

        End If

        If opcion1 = "101" Then
            Frame1.Visible = False
            Frame1.Enabled = False
   
            Exit Sub

        End If

        If opcion1 = "11" Then
            'Frame1.Visible = False
            'proveedor2.SetFocus
            Exit Sub

        End If

    End If

    tproduct.Hide
    Unload tproduct

End Sub

Private Sub fabrica_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    xproveedor.SetFocus

End Sub

Private Sub fabrica_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        color.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_fabrica

    End If

End Sub

Private Sub factor_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes
    costou.SetFocus

End Sub

Private Sub factor_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        unidad.SetFocus
        Exit Sub

    End If

End Sub

Private Sub factor1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    pventa1.SetFocus

End Sub

Private Sub factor1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        unidad1.SetFocus
        Exit Sub

    End If

End Sub

Private Sub factor2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    pventa2.SetFocus

End Sub

Private Sub factor2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        unidad2.SetFocus
        Exit Sub

    End If

End Sub

Private Sub factor3_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    pventa3.SetFocus

End Sub

Private Sub factor3_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        unidad3.SetFocus
        Exit Sub

    End If

End Sub

Private Sub factor4_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    pventa4.SetFocus

End Sub

Private Sub factor4_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        unidad4.SetFocus
        Exit Sub

    End If

End Sub

Private Sub familia_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If Len(familia) = 0 Then
        consulta_Familia
        Exit Sub

    End If

    pventa1.SetFocus

End Sub

Private Sub familia_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        presenta.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_Familia

    End If

    If KeyCode = &H76 Then  'f7
        ttfamilia.Show 1

    End If

End Sub

Private Sub flete_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    igv.SetFocus

End Sub

Private Sub flete_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        pesokgr.SetFocus
        Exit Sub

    End If

End Sub

Private Sub Form_Activate()
    Me.Height = 9720: Me.Width = 13860
    Frame1.Top = 30: Frame1.Left = 10
    Frame2.Top = 675: Frame2.Left = 1620
    Frame3.Top = 7080: Frame3.Left = 6120
    frmOtros.Top = 600: frmOtros.Left = 8340

    Dim found As Integer

    carga_impresoras
    paridad = "T/C:" & busca_cambio()

    If ventanas = "" Then
        If ordename = "NUEVO" Then
            codigo.Enabled = True
            xadicion

        End If

        If ordename = "MODIFICA" Then
            found = busca_registro()
            codigo.Enabled = False

        End If

        If ordename = "VER" Then
            found = busca_registro()
            cmdSave.Enabled = False
            grba1.Enabled = False
            'bo712.Enabled = False
            cmdAddEntry.Enabled = False

        End If

        ventanas = "S"

    End If

End Sub

Private Sub Form_Load()
    Label16 = dicigv
    inicializa_grupos
    dia.Clear
    dia.AddItem ""
    dia.AddItem "LUNES"
    dia.AddItem "MARTES"
    dia.AddItem "MIERCOLES"
    dia.AddItem "JUEVES"
    dia.AddItem "VIERNES"
    dia.AddItem "SABADO"
    dia.AddItem "DOMINGO"
    dia.ListIndex = 0

End Sub

Sub inicializa_grupos()

    Dim mytablex As New ADODB.Recordset

    serie.Clear
    serie.AddItem "N"
    serie.AddItem "S"
    serie.ListIndex = 0

    Peso.Clear
    Peso.AddItem "N"
    Peso.AddItem "S"
    Peso.ListIndex = 0

    servicio.Clear
    servicio.AddItem "N"
    servicio.AddItem "S"
    servicio.ListIndex = 0

    vtaund.Clear
    vtaund.AddItem "S"
    vtaund.AddItem "N"
    vtaund.ListIndex = 0

    oferta.Clear
    oferta.AddItem "N"
    oferta.AddItem "S"
    oferta.ListIndex = 0

    remate.Clear
    remate.AddItem "N"
    remate.AddItem "S"
    remate.ListIndex = 0

    '' kenyo 02/11/2017 Configuración Productos Comentarios Obligatorios
    obligacomentario.Clear
    obligacomentario.AddItem "N"
    obligacomentario.AddItem "S"
    obligacomentario.ListIndex = 0
    '' kenyo 02/11/2017 Configuración Productos Comentarios Obligatorios

    '13/08/2018 Integración FE - Pizzeria
    '''' 11/12/2017 SubReceta
    CostoReceta.Clear
    CostoReceta.AddItem "S"
    CostoReceta.AddItem "N"
    CostoReceta.ListIndex = 0
    '''' 11/12/2017 SubReceta
    '13/08/2018 Integración FE - Pizzeria

    '27/08/2018 Producto delivery automatico
    deliveryautom.Clear
    deliveryautom.AddItem "N"
    deliveryautom.AddItem "S"
    deliveryautom.ListIndex = 0
    '27/08/2018 Producto delivery automatico

    vecaja.Clear
    vecaja.AddItem "S"
    vecaja.AddItem "N"
    vecaja.ListIndex = 0

    estado.Clear
    estado.AddItem "S"
    estado.AddItem "N"
    estado.ListIndex = 0

    monedac.Clear
    monedac.AddItem "S"
    monedac.AddItem "D"
    monedac.ListIndex = 0

    monedav.Clear
    monedav.AddItem "S"
    monedav.AddItem "D"
    monedav.ListIndex = 0

    local2.Clear
    mytablex.Open "SELECT * FROM parame where  codigo='01'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        If Len(Trim("" & mytablex.Fields("listaprecioS"))) > 0 Then
            local2.AddItem Trim("" & mytablex.Fields("listaprecioS"))

        End If

    End If

    local2.AddItem "00"
    local2.AddItem "01"
    local2.AddItem "02"
    local2.AddItem "03"
    local2.AddItem "04"
    local2.AddItem "05"
    local2.AddItem "06"
    local2.AddItem "07"
    local2.AddItem "08"
    local2.AddItem "09"
    local2.AddItem "10"
    local2.AddItem "11"

    'KENYO COMENTRIO
    If Trim(mytablex.Fields("listaprecioS")) = "" Then
        local2.ListIndex = 1
    Else
        local2.ListIndex = 0

    End If

    mytablex.Close

End Sub

Sub inicializa()

    Dim found As Integer

    SEINVENTARIA = ""

    ''11/07/2017 kenyo edicion producto cola
    'cola = "S"
    cola = ""
    ''11/07/2017 kenyo edicion producto cola

    tecla = ""
    puertoimpresion1 = ""
    puertoimpresion2 = ""
    puertoimpresion3 = ""
    'produccion = ""
    'formulacion = ""
    serviciomesa = ""
    fueldonde = ""
    'codigobalanza = ""
    comisioncredito = ""
    costoanterior1 = ""
    costoanterior2 = ""
    puertoimpresion = ""

    ' 05/06/207   grupoimpresion ''
    'grupoimpresion = "ZZ"
    grupoimpresion = ""

    platos = ""

    recetaprn = ""
    inicializa_grupos
    fuel = ""
    'costopais = ""
    'gastoimp = ""
    'costoimp = ""
    dsctoref = ""
    margen = ""
    touch = ""
    'l1 = ""
    'l2 = ""
    'l3 = ""
    fechavence = ""
    diasalerta = ""
    detraccion = ""
    'l4 = ""
    ivap = ""
    flete = ""
    percepcion = ""
    fotonombre = ""
    'ccosto = ""
    xproveedor.Clear
    unidadp = ""
    factorp = ""

    Barras = ""
    barras2 = ""
    descripcio = ""
    descorto = ""
    presenta = ""
    familia = ""
    subfamilia = ""
    seccion = ""
    marca = ""
    categoria = ""
    lineatalla = ""
    color = ""
    fabrica = ""

    ''17/07/2017 kenyo tienda ropa opciones producto
    talla = ""
    proyecto = ""
    sexo = ""
    procedencia = ""
    ''17/07/2017 kenyo tienda ropa opciones producto

    'proveedor1 = ""
    'proveedor2 = ""
    'proveedor3 = ""
    'proveedor4 = ""

    'codprov1 = ""
    'codprov2 = ""
    'codprov3 = ""
    'codprov4 = ""

    serie.ListIndex = 0
    Peso.ListIndex = 0
    servicio.ListIndex = 0
    vtaund.ListIndex = 0
    oferta.ListIndex = 0
    vecaja.ListIndex = 0
    estado.ListIndex = 0
    igv = ""
    isc = ""
    pesokgr = ""
    comision = ""
    monedac.ListIndex = 0
    unidad = "UND"
    factor = "1"
    costop = ""
    costou = ""
    costoini = ""
    'fechavence = ""
    monedav.ListIndex = 0
    unidad1 = "UND"
    unidad2 = ""
    unidad3 = ""
    unidad4 = ""
    unidad5 = ""
    unidad6 = ""
    unidad7 = ""
    unidad8 = ""
    unidad9 = ""
    unidad10 = ""
    'saldoini = ""

    factor1 = "1"
    factor2 = ""
    factor3 = ""
    factor4 = ""
    factor5 = ""
    factor6 = ""
    factor7 = ""
    factor8 = ""
    factor9 = ""
    factor10 = ""

    pventa1 = ""
    pventa2 = ""
    pventa3 = ""
    pventa4 = ""
    pventa5 = ""
    pventa6 = ""
    pventa7 = ""
    pventa8 = ""
    pventa9 = ""
    pventa10 = ""

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
    minimo11 = ""
    minimo12 = ""
    minimo13 = ""
    minimo14 = ""
    minimo15 = ""

    maximo11 = ""
    maximo12 = ""
    maximo13 = ""
    maximo14 = ""
    maximo15 = ""

    pventa11 = ""
    pventa12 = ""
    pventa13 = ""
    pventa14 = ""
    pventa15 = ""

    margen11 = ""
    margen12 = ""
    margen13 = ""
    margen14 = ""
    margen15 = ""
    fechai11 = ""
    fechaf11 = ""
    fechaid = ""
    fechafd = ""
    dscto = ""
    minimo = ""
    maximo = ""
    empaque_visible = ""

    found = busca_parame(2)

End Sub

Function borra_registro()

    Dim mytablex As New ADODB.Recordset

    Dim Tmp      As String

    Dim sw       As Integer

    Dim found    As Integer

    sw = 0

    On Error GoTo cmd3_err

    Tmp = ""

    mytablex.Open "SELECT * FROM producto where  producto='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If MsgBox("Desea Borra el registro", 1, "Aviso") = "1" Then
            Tmp = "" & mytablex.Fields("producto")
            mytablex.Delete
            borra_registro = 1
            sw = 1

        End If

    End If

    '------------------------------------- ------------
    If sw = 1 Then
        borra_almacen_producto Tmp
   
    End If

    mytablex.Close

    Exit Function
cmd3_err:
    MsgBox "Mensaje:" + error$, 48, "Aviso"
    mytablex.Close
 
    Exit Function

End Function

Sub borra_almacen_producto(Tmp As String)

    On Error GoTo cmd34_err

    mydbxglo.Execute "DELETE FROM ALMACEN WHERE producto='" & Tmp & "'"
    mydbxglo.Execute "DELETE FROM productob WHERE producto='" & Tmp & "'"
    mydbxglo.Execute "DELETE FROM codprov WHERE producto='" & Tmp & "'"
    mydbxglo.Execute "DELETE FROM precios WHERE producto='" & Tmp & "'"
    Exit Sub
cmd34_err:

    Exit Sub

End Sub

Function busca_registro()

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    sw = 0
    mytablex.Open "SELECT * FROM producto where  producto='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        pone_registro mytablex
        busca_registro = 1
        sw = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

    If sw = 1 Then
        carga_proveedor

    End If

End Function

Sub pone_registro(mytablex As ADODB.Recordset)

    Dim found As Integer

    Dim I     As Integer

    On Error GoTo cmd89000_err

    pone_fotonombre mytablex

    'mytablex.Fields("seinventaria") = Trim(seinventaria)
    SEINVENTARIA = Trim("" & mytablex.Fields("seinventaria"))
    'MsgBox "xx"
    'l1 = "" & mytablex.Fields("l1")
    'l2 = "" & mytablex.Fields("l2")
    'l3 = "" & mytablex.Fields("l3")
    'l4 = "" & mytablex.Fields("l4")
    'MsgBox "abc"
    'foto = "" & mytablex.Fields("fueldonde")
    'produccion = "" & mytablex.Fields("produccion")
    'formulacion = "" & mytablex.Fields("formulacion")
    tecla = Trim("" & mytablex.Fields("tecla"))
    serviciomesa = "" & mytablex.Fields("serviciomesa")
    fueldonde = "" & mytablex.Fields("fueldonde")
    'MsgBox "abc"
    'codigobalanza = "" & mytablex.Fields("codigobalanza")
    'MsgBox "abc"
    comisioncredito = "" & mytablex.Fields("comisioncredito")
    'MsgBox "abc"
    costoanterior1 = "" & mytablex.Fields("costoanterior1")
    costoanterior2 = "" & mytablex.Fields("costoanterior2")

    cola = "" & mytablex.Fields("cola")
    puertoimpresion1 = "" & mytablex.Fields("puertoimpresion1")
    puertoimpresion2 = "" & mytablex.Fields("puertoimpresion2")
    puertoimpresion3 = "" & mytablex.Fields("puertoimpresion3")

    puertoimpresion = "" & mytablex.Fields("puertoimpresion")
    grupoimpresion = "" & mytablex.Fields("grupoimpresion")
    'MsgBox "abc"
    recetaprn = "" & mytablex.Fields("recetaprn")
    empaque_visible = "" & mytablex.Fields("empaque_visible")
    platos = "" & mytablex.Fields("platos")
    fuel = "" & mytablex.Fields("fuel")
    'costopais = "" & mytablex.Fields("costopais")
    'gastoimp = "" & mytablex.Fields("costogasto")
    'costoimp = "" & mytablex.Fields("costoimp")

    touch = "" & mytablex.Fields("touch")
    dsctoref = "" & mytablex.Fields("dsctoref")
    unidadp = "" & mytablex.Fields("unidadp")
    factorp = "" & mytablex.Fields("factorp")
    margen = "" & mytablex.Fields("margen")
    fechavence = "" & mytablex.Fields("fechavence")
    diasalerta = "" & mytablex.Fields("diasalerta")
    percepcion = "" & mytablex.Fields("percepcion")
    codigo = Trim("" & mytablex.Fields("producto"))
    Barras = "" & mytablex.Fields("barras")
    'barras2 = "" & mytablex.Fields("barras2")
    descripcio = "" & mytablex.Fields("descripcio")
    descorto = "" & mytablex.Fields("descorto")
    presenta = "" & mytablex.Fields("presenta")
    familia = "" & mytablex.Fields("familia")
    subfamilia = "" & mytablex.Fields("subfamilia")
    seccion = "" & mytablex.Fields("seccion")
    marca = "" & mytablex.Fields("marca")
    categoria = "" & mytablex.Fields("categoria")
    lineatalla = "" & mytablex.Fields("linea")
    color = "" & mytablex.Fields("color")

    ''17/07/2017 kenyo tienda ropa opciones producto
    talla = "" & mytablex.Fields("talla")
    proyecto = "" & mytablex.Fields("proyecto")
    sexo = "" & mytablex.Fields("sexo")
    procedencia = "" & mytablex.Fields("procedencia")
    ''17/07/2017 kenyo tienda ropa opciones producto

    flete = "" & mytablex.Fields("flete")
    fabrica = "" & mytablex.Fields("fabrica")
    detraccion = "" & mytablex.Fields("detraccion")
    'proveedor1 = "" & mytablex.Fields("proveedor1")
    'proveedor2 = "" & mytablex.Fields("proveedor2")
    'proveedor3 = "" & mytablex.Fields("proveedor3")
    'proveedor4 = "" & mytablex.Fields("proveedor4")
    'codprov1 = "" & mytablex.Fields("codprov1")
    'codprov2 = "" & mytablex.Fields("codprov2")
    'codprov3 = "" & mytablex.Fields("codprov3")
    'codprov4 = "" & mytablex.Fields("codprov4")
    remate.ListIndex = 0

    If "" & mytablex.Fields("remate") = "S" Then
        remate.ListIndex = 1

    End If

    '' kenyo 02/11/2017 Configuración Productos Comentarios Obligatorios
    obligacomentario.ListIndex = 0

    If "" & mytablex.Fields("obligacomentario") = "S" Then
        obligacomentario.ListIndex = 1

    End If

    '' kenyo 02/11/2017 Configuración Productos Comentarios Obligatorios

    '27/08/2018 Producto delivery automatico
    deliveryautom.ListIndex = 0

    If "" & mytablex.Fields("deliveryautom") = "S" Then
        deliveryautom.ListIndex = 1

    End If

    '27/08/2018 Producto delivery automatico

    '13/08/2018 Integración FE - Pizzeria
    '''' 11/12/2017 SubReceta
    CostoReceta.ListIndex = 0

    If "" & mytablex.Fields("CostoReceta") = "N" Then
        CostoReceta.ListIndex = 1

    End If

    '''' 11/12/2017 SubReceta
    '13/08/2018 Integración FE - Pizzeria

    serie.ListIndex = 0

    If "" & mytablex.Fields("serie") = "S" Then
        serie.ListIndex = 1

    End If

    Peso.ListIndex = 0

    If "" & mytablex.Fields("peso") = "S" Then
        Peso.ListIndex = 1

    End If

    servicio.ListIndex = 0

    If "" & mytablex.Fields("servicio") = "S" Then
        servicio.ListIndex = 1

    End If

    vtaund.ListIndex = 0

    If "" & mytablex.Fields("vtaund") <> "S" Then
        vtaund.ListIndex = 1

    End If

    oferta.ListIndex = 0

    If "" & mytablex.Fields("oferta") = "S" Then
        oferta.ListIndex = 1

    End If

    vecaja.ListIndex = 0

    If "" & mytablex.Fields("vecaja") <> "S" Then
        vecaja.ListIndex = 1

    End If

    estado.ListIndex = 0

    If "" & mytablex.Fields("estado") <> "S" Then
        estado.ListIndex = 1

    End If

    'If mytablex.Fields("estado") Is Null Then
    'estado.ListIndex = 0
    'End If
    igv = "" & mytablex.Fields("igv")
    isc = "" & mytablex.Fields("isc")
    ivap = "" & mytablex.Fields("ivap")
    pesokgr = "" & mytablex.Fields("pesokgr")
    comision = "" & mytablex.Fields("comision")
    monedac.ListIndex = 0

    If "" & mytablex.Fields("monedac") = "D" Then
        monedac.ListIndex = 1

    End If

    dia.ListIndex = 0

    For I = 0 To dia.ListCount - 1

        If dia.List(I) = Trim("" & mytablex.Fields("dia")) Then
            dia.ListIndex = I
            Exit For

        End If

    Next I

    unidad = "" & mytablex.Fields("unidad")
    factor = "" & mytablex.Fields("factor")
    costop = "" & mytablex.Fields("costop")
    costoini = "" & mytablex.Fields("costoini")
    costou = "" & mytablex.Fields("costou")

    If Val(factor) <= 0 Then
        factor = "1"

    End If

    cospaqu = Format(Val(costou) * Val(factor))
    cospaqp = Format(Val(costop) * Val(factor))
    cospaqi = Format(Val(costoini) * Val(factor))

    'ccosto = "" & mytablex.Fields("ccosto")
    'fechavence = "" & mytablex.Fields("fechavence")
    monedav.ListIndex = 0

    If "" & mytablex.Fields("monedav") = "D" Then
        monedav.ListIndex = 1

    End If

    carga_precios "" & Trim(codigo)

    If Len(unidad1) = 0 Then
        unidad1 = "UND"

    End If

    If Len(factor1) = 0 Then
        factor1 = "1"

    End If

    minimo = "" & mytablex.Fields("minimo")
    maximo = "" & mytablex.Fields("maximo")

    If Len(Trim(unidadp)) = 0 Then
        unidadp = unidad

    End If

    If Val(factorp) <= 0 Then
        factorp = factor

    End If

    'ccosto = "" & mytablex.Fields("ccosto")
    'For i = 1 To 15
    'calcula_margenes i, 0
    'Next i
    'found = busca_bodega("" & codigo, "" & bodega, 0)
    calcula_margenes
    Exit Sub
cmd89000_err:
    MsgBox "Mensaje Pone Registro ", 48, "Aviso"
    Exit Sub

End Sub

Sub grabando(mytablex As ADODB.Recordset)

    Dim found As Integer

    'On Error GoTo cmd7832_err
    If Len(Trim(fotonombre)) > 0 Then
        SaveBitmap mytablex, Trim("" & fotonombre)

    End If

    '' 02/12/2017 Opcion a poder borrar imagen del modulo de productos
    If fotonombre = "" Then
        SaveBitmap mytablex, Trim("" & fotonombre)

    End If

    '' 02/12/2017 Opcion a poder borrar imagen del modulo de productos

    calcula_margenes
    'mytablex.Fields("xlocal1") = "DONNA1"
    'mytablex.Fields("xlocal1") = "DONNA2"
    'mytablex.Fields("xlocal1") = "DONNA3"
    'mytablex.Fields("xlocal1") = "DIMAGA"
    'mytablex.Fields("xccosto1") = xccosto1
    'mytablex.Fields("xccosto2") = xccosto2
    'mytablex.Fields("xccosto3") = xccosto3
    'mytablex.Fields("xccosto4") = xccosto4

    'mytablex.Fields("l1") = Val(l1)
    'mytablex.Fields("l2") = Val(l2)
    'mytablex.Fields("l3") = Val(l3)
    'mytablex.Fields("l4") = Val(l4)
    If IsDate(fechavence) = True Then
        mytablex.Fields("fechavence") = fechavence
    Else
        mytablex.Fields("fechavence") = Null

        'Else
        'mytablex.Fields("fechavence") = ""
    End If

    'If fechavence = "" Then
    '  mytablex.Fields("fechavence") = Null
    'End If

    mytablex.Fields("diasalerta") = Trim(diasalerta)
    mytablex.Fields("seinventaria") = Trim(SEINVENTARIA)
    'If Trim("" & fotonombre) <> "" Then mytablex.Fields("fotonombre") = Trim("" & fotonombre)
    mytablex.Fields("tecla") = Trim("" & tecla)
    'mytablex.Fields("c11") = "0"
    'mytablex.Fields("c12") = "0"
    'mytablex.Fields("c13") = "0"
    'mytablex.Fields("c14") = "0"

    'mytablex.Fields("c21") = "0"
    'mytablex.Fields("c22") = "0"
    'mytablex.Fields("c23") = "0"
    'mytablex.Fields("c24") = "0"

    'mytablex.Fields("c31") = "0"
    'mytablex.Fields("c32") = "0"
    'mytablex.Fields("c33") = "0"
    'mytablex.Fields("c34") = "0"

    'mytablex.Fields("c41") = "0"
    'mytablex.Fields("c42") = "0"
    'mytablex.Fields("c43") = "0"
    'mytablex.Fields("c44") = "0"
    If Len(Trim(unidadp)) = 0 Then
        unidadp = unidad

    End If

    If Val(factorp) <= 0 Then
        factorp = factor

    End If

    'mytablex.Fields("produccion") = Trim(produccion)
    'mytablex.Fields("formulacion") = Trim(formulacion)
    mytablex.Fields("dia") = Trim(dia.Text)
    mytablex.Fields("fueldonde") = Trim(fueldonde)
    'mytablex.Fields("codigobalanza") = Trim(codigobalanza)
    mytablex.Fields("comisioncredito") = Val(comisioncredito)
    mytablex.Fields("costoanterior1") = Val(costoanterior1)
    mytablex.Fields("costoanterior2") = Val(costoanterior2)
    mytablex.Fields("cola") = Trim(cola)
    mytablex.Fields("puertoimpresion1") = Trim(puertoimpresion1)
    mytablex.Fields("puertoimpresion2") = Trim(puertoimpresion2)
    mytablex.Fields("puertoimpresion3") = Trim(puertoimpresion3)

    mytablex.Fields("puertoimpresion") = Trim(puertoimpresion)
    mytablex.Fields("grupoimpresion") = Trim(grupoimpresion)
    mytablex.Fields("recetaprn") = Trim(recetaprn)

    mytablex.Fields("empaque_visible") = Trim(empaque_visible)
    mytablex.Fields("platos") = Val(platos)
    mytablex.Fields("fuel") = Trim(fuel)
    'mytablex.Fields("costopais") = Val(costopais)
    'mytablex.Fields("costogasto") = Val(gastoimp)
    'mytablex.Fields("costoimp") = Val(costoimp)
    mytablex.Fields("touch") = Val(touch)
    mytablex.Fields("dsctoref") = Val(dsctoref)
    mytablex.Fields("unidadp") = Trim(unidadp)
    mytablex.Fields("factorp") = Val(factorp)
    mytablex.Fields("margen") = Trim(margen)
    mytablex.Fields("OK") = ""
    mytablex.Fields("percepcion") = Trim(percepcion)
    mytablex.Fields("producto") = Trim(codigo)
    mytablex.Fields("detraccion") = Val(detraccion)
    mytablex.Fields("ivap") = Val(ivap)
    mytablex.Fields("flete") = Val(flete)
    'mytablex.Fields("ccosto") = ccosto
    mytablex.Fields("barras") = Trim(Barras)
    mytablex.Fields("descripcio") = UCase$(Trim(descripcio))
    mytablex.Fields("descorto") = UCase(Trim(descorto))
    mytablex.Fields("presenta") = Trim(presenta)
    mytablex.Fields("familia") = Trim(familia)
    mytablex.Fields("subfamilia") = Trim(subfamilia)
    mytablex.Fields("seccion") = Trim(seccion)
    mytablex.Fields("marca") = Trim(marca)
    mytablex.Fields("remate") = Trim(remate)

    '' kenyo 02/11/2017 Configuración Productos Comentarios Obligatorios
    mytablex.Fields("obligacomentario") = Trim(obligacomentario)
    '' kenyo 02/11/2017 Configuración Productos Comentarios Obligatorios

    '27/08/2018 Producto delivery automatico
    mytablex.Fields("deliveryautom") = Trim(deliveryautom)
    '27/08/2018 Producto delivery automatico

    '13/08/2018 Integración FE - Pizzeria
    '''' 11/12/2017 SubReceta
    mytablex.Fields("CostoReceta") = Trim(CostoReceta)
    '''' 11/12/2017 SubReceta
    '13/08/2018 Integración FE - Pizzeria

    mytablex.Fields("categoria") = Trim(categoria)
    mytablex.Fields("linea") = Trim(lineatalla)

    mytablex.Fields("color") = Trim(color)

    ''17/07/2017 kenyo tienda ropa opciones producto

    mytablex.Fields("talla") = Trim(talla)
    mytablex.Fields("proyecto") = Trim(proyecto)
    mytablex.Fields("sexo") = Trim(sexo)
    mytablex.Fields("procedencia") = Trim(procedencia)
    ''17/07/2017 kenyo tienda ropa opciones producto

    mytablex.Fields("fabrica") = Trim(fabrica)
    mytablex.Fields("serviciomesa") = Val(serviciomesa)
    'mytablex.Fields("proveedor1") = proveedor1
    'mytablex.Fields("proveedor2") = proveedor2
    'mytablex.Fields("proveedor3") = proveedor3
    'mytablex.Fields("proveedor4") = proveedor4

    'mytablex.Fields("codprov1") = codprov1
    'mytablex.Fields("codprov2") = codprov2
    'mytablex.Fields("codprov3") = codprov3
    'mytablex.Fields("codprov4") = codprov4

    mytablex.Fields("serie") = serie
    mytablex.Fields("peso") = Peso
    mytablex.Fields("servicio") = servicio
    mytablex.Fields("vtaund") = vtaund
    mytablex.Fields("oferta") = oferta
    mytablex.Fields("vecaja") = vecaja
    mytablex.Fields("estado") = estado
    mytablex.Fields("igv") = Val(igv)
    mytablex.Fields("isc") = Val(isc)
    mytablex.Fields("pesokgr") = Val(pesokgr)
    mytablex.Fields("comision") = Val(comision)
    mytablex.Fields("monedac") = monedac
    mytablex.Fields("unidad") = unidad
    mytablex.Fields("factor") = Val(factor)
    mytablex.Fields("costou") = Val(costou)
    mytablex.Fields("costop") = Val(costop)
    mytablex.Fields("costoini") = Val(costoini)
    mytablex.Fields("minimo") = Val(minimo)
    mytablex.Fields("maximo") = Val(maximo)
    'If IsDate(fechavence) Then
    '   mytablex.Fields("fechavence") = fechavence
    'End If
    mytablex.Fields("monedav") = monedav

    ''''26/09/2017 kenyo Costos al por mayor Factura de Compra
    'mytablex.Fields("cospaqu") = Val(cospaqu) * Val(factor)
    mytablex.Fields("cospaqu") = Val(costou) * Val(factor)
    ''''26/09/2017 kenyo Costos al por mayor Factura de Compra

    mytablex.Fields("cospaqp") = Val(cospaqp) * Val(factor)
    mytablex.Fields("cospaqi") = Val(cospaqi) * Val(factor)

    actualiza_receta

    ''' 11/12/2017 SubReceta
    actualiza_sumacostoreceta
    ''' 11/12/2017 SubReceta

    'found = busca_bodega("" & codigo, "" & bodega, 1)
    Exit Sub
cmd7832_err:
    MsgBox "Aviso en grabando " + error$, 48, "Aviso"
    Exit Sub

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

Private Sub grba1_Click()

    Dim found As Integer

    'If Frame4.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    found = grabar()

    If found = 1 Then
        xprodet.FLAG = "1"
        dlo132_Click
        Exit Sub

    End If

    Barras.SetFocus

End Sub

Private Sub igv_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    isc.SetFocus

End Sub

Private Sub igv_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        flete.SetFocus
        Exit Sub

    End If

End Sub

Private Sub isc_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    percepcion.SetFocus

End Sub

Private Sub isc_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        igv.SetFocus
        Exit Sub

    End If

End Sub

Private Sub Label1_Click()

    If ordename = "NUEVO" Then
        busca_correlativo 0

    End If

    'cmdSort_Click
End Sub

Function grabar()

    Dim found    As Integer

    Dim sdx      As Double

    Dim mytabley As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    Dim sw       As Integer

    found = valida()

    If found = 0 Then
        MsgBox "Campos invalidos", 48, "Aviso"
        Exit Function

    End If

    sw = 0

    If ordename = "NUEVO" Then
        If Len(Trim(codigo)) = 0 Then
            ' codigo.SetFocus
            descripcio.SetFocus
            Exit Function

        End If

        If mytablex.State = 1 Then mytablex.Close
        If MsgBox("Desea Adicionar?", 1, "Aviso") = 1 Then
            mytablex.Open "SELECT * FROM producto where producto='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic

            If mytablex.RecordCount > 0 Then
                mytablex.Close
                MsgBox "Ya existe codigo Usado ", 48, "Aviso"
                Exit Function

            End If

            mytablex.AddNew
            grabando mytablex
            mytablex.Fields("tipocreacion") = "NUEVO"
            mytablex.Update
            graba_precios
            busca_correlativo 1
            grabar = 1
            mytablex.Close
            MsgBox "Proceso Grabado ", 48, "Aviso"

        End If

        Exit Function

    End If

    If ordename = "MODIFICA" Then
        If mytablex.State = 1 Then mytablex.Close

        If puede_modificar() = 0 Then
            graba_precios
            grabar = 1
            Exit Function

        End If

        mytablex.Open "SELECT * FROM producto where  producto='" & Trim(Trim(codigo)) & "'", cn, adOpenKeyset, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            If MsgBox("Desea Reescribir?", 1, "Aviso") = 1 Then
                grabando mytablex
                mytablex.Fields("tipocreacion") = "MODIFICA"
                mytablex.Update
                graba_precios
                grabar = 1

            End If

        End If

        mytablex.Close

    End If

End Function

Function valida()

    Dim found As Integer

    Dim buf   As String

    If Len(Trim(codigo)) = 0 Then
        ' codigo.SetFocus
        descripcio.SetFocus
        Exit Function

    End If

    buf = convierte_barras("" & Barras)

    If Len(buf) > 0 Then
        Barras = buf

    End If

    If Len(Barras) > 0 Then
        found = valida_barras("" & Barras)

        If found = 1 Then
            Barras.SetFocus
            Exit Function

        End If

    End If

    If Len(Trim(tecla)) > 0 Then
        found = valida_tecla("" & tecla)

        If found = 1 Then
            tecla.SetFocus
            Exit Function

        End If

    End If

    If Len(descripcio) = 0 Then
        descripcio.SetFocus
        Exit Function

    End If

    If Len(familia) = 0 Then
        familia.SetFocus
        Exit Function

    End If

    found = busca_familia()

    If found = 0 Then
        MsgBox "No existe Familia", 48, "Aviso"
        familia.SetFocus
        Exit Function

    End If

    If Len(subfamilia) > 0 Then
        found = busca_subfamilia()

        If found = 0 Then
            MsgBox "No existe SubFamilia", 48, "Aviso"
            subfamilia.SetFocus
            Exit Function

        End If

    End If

    If Len(seccion) > 0 Then
        found = busca_seccion()

        If found = 0 Then
            MsgBox "No existe Seccion", 48, "Aviso"
            seccion.SetFocus
            Exit Function

        End If

    End If

    If Len(categoria) > 0 Then
        found = busca_categoria()

        If found = 0 Then
            MsgBox "No existe Categoria", 48, "Aviso"
            categoria.SetFocus
            Exit Function

        End If

    End If

    ''17/07/2017 kenyo tienda ropa opciones producto

    If Len(talla) > 0 Then
        found = busca_talla()

        If found = 0 Then
            MsgBox "No existe talla", 48, "Aviso"
            talla.SetFocus
            Exit Function

        End If

    End If

    If Len(sexo) > 0 Then
        found = busca_sexo()

        If found = 0 Then
            MsgBox "No existe sexo", 48, "Aviso"
            sexo.SetFocus
            Exit Function

        End If

    End If

    If Len(procedencia) > 0 Then
        found = busca_procedencia()

        If found = 0 Then
            MsgBox "No existe procedencia", 48, "Aviso"
            procedencia.SetFocus
            Exit Function

        End If

    End If

    If Len(proyecto) > 0 Then
        found = busca_proyecto()

        If found = 0 Then
            MsgBox "No existe proyecto", 48, "Aviso"
            proyecto.SetFocus
            Exit Function

        End If

    End If

    ''17/07/2017 kenyo tienda ropa opciones producto

    If Len(marca) > 0 Then
        found = busca_marca()

        If found = 0 Then
            MsgBox "No existe Marca", 48, "Aviso"
            marca.SetFocus
            Exit Function

        End If

    End If

    If Len(color) > 0 Then
        found = busca_color()

        If found = 0 Then
            MsgBox "No existe Color", 48, "Aviso"
            color.SetFocus
            Exit Function

        End If

    End If

    If Len(lineatalla) > 0 Then
        found = busca_lineatalla()

        If found = 0 Then
            MsgBox "No existe Talla", 48, "Aviso"
            lineatalla.SetFocus
            Exit Function

        End If

    End If

    If Len(Trim(margen)) > 0 Then
        found = busca_margen()

        If found = 0 Then
            MsgBox "No existe margen", 48, "Aviso"
            margen.SetFocus
            Exit Function

        End If

    End If

    'If Len(ccosto) > 0 Then
    'found = busca_ccosto()
    'If found = 0 Then
    '   MsgBox "No existe Centro Costo", 48, "Aviso"
    '   ccosto.SetFocus
    '   Exit Function
    'End If
    '
    'End If

    If Len(unidad1) = 0 Then
        unidad1 = "UND"

    End If

    If Val(factor1) = 0 Then
        factor1 = "1"

    End If

    If ordename = "NUEVO" Then

    End If

    If ordename = "MODIFICA" Then

    End If

    valida = 1

End Function

Private Sub Label46_Click()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from familia where familia='" & Trim(familia) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        puertoimpresion = Mid$(Trim("" & mytablex.Fields("red")), 1, 30)
        grupoimpresion = "" & mytablex.Fields("puerto")
        cola = "" & mytablex.Fields("cola")

    End If

    mytablex.Close
 
End Sub

Private Sub label51_Click()

    'Label66.Caption = "Nuevo"
    'Label60.Visible = True
    'Label58.Visible = True
    'pcodigo.Enabled = True
    'pcodigo = ""
    'pncodigo = ""
    'pcodigop = ""
End Sub

Private Sub Label52_Click()
    'On Error GoTo cmd435_err
    'Label66.Caption = "Modifica"
    'Label60.Visible = True
    'Label58.Visible = True
    'pcodigo = Trim("" & DBGrid9.columns(0))
    'pncodigo = Trim("" & DBGrid9.columns(1))
    'pcodigop = Trim("" & DBGrid9.columns(2))
    'pcodigo.Enabled = False
    'DBGrid9.Enabled = False
    'Exit Sub
    'cmd435_err:
    'MsgBox "Seleccione un dato", 48, "Aviso"
    'Exit Sub

End Sub

Private Sub Label53_Click()
    'On Error GoTo cmd490_err
    'cn.Execute ("delete from codprov where codigo='" & Trim("" & DBGrid9.columns(0)) & "' and producto='" & Trim(codigo) & "'")
    'Exit Sub
    'cmd490_err:
    'MsgBox "Seleccione un dato", 48, "Aviso"
    'Exit Sub

End Sub

Private Sub Label79_Click()
    '' 02/12/2017 Opcion a poder borrar imagen del modulo de productos
    foto.Picture = Nothing
    fotonombre = ""
    '' 02/12/2017 Opcion a poder borrar imagen del modulo de productos

End Sub

Private Sub Label84_Click()
    Frame3.Visible = True

End Sub

Private Sub lblSeccionMarca_Click()

End Sub

Private Sub lineatalla_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    color.SetFocus

End Sub

Private Sub lineatalla_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        categoria.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_talla

    End If

    If KeyCode = &H76 Then  'f7
        tlinea.Show 1

    End If

End Sub

Private Sub local2_Click()
    carga_precios "" & Trim(codigo)

End Sub

Private Sub local2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    carga_precios "" & Trim(codigo)

End Sub

Private Sub marca_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    categoria.SetFocus

End Sub

Private Sub marca_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        seccion.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_marca

    End If

    If KeyCode = &H76 Then  'f7
        tnmarca.Show 1

    End If

End Sub

Private Sub margen_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub margen_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        'subfamilia.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_margen

    End If

    If KeyCode = &H76 Then  'f7
        tmargen.Show 1

    End If

End Sub

Private Sub margen1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes
    unidad2.SetFocus

End Sub

Private Sub margen1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        pventa1.SetFocus
        Exit Sub

    End If

End Sub

Private Sub margen10_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub margen11_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes

End Sub

Private Sub margen12_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes

End Sub

Private Sub margen13_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes

End Sub

Private Sub margen14_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes

End Sub

Private Sub margen15_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes

End Sub

Private Sub margen2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes
    unidad3.SetFocus

End Sub

Private Sub margen2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        pventa2.SetFocus
        Exit Sub

    End If

End Sub

Private Sub margen3_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes
    unidad4.SetFocus

End Sub

Private Sub margen3_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        pventa3.SetFocus
        Exit Sub

    End If

End Sub

Private Sub margen4_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes
    unidad5.SetFocus

End Sub

Private Sub margen4_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        pventa4.SetFocus
        Exit Sub

    End If

End Sub

Private Sub margen5_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub margen6_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub margen7_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub margen8_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub margen9_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub mas_Click()

    If frmOtros.Visible = True Then
        frmOtros.Visible = False
        Exit Sub

    End If
  
    If frmOtros.Visible = False Then
        frmOtros.Visible = True
        Exit Sub

    End If
    
End Sub

Private Sub monedac_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    unidad.SetFocus

End Sub

Private Sub monedac_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        percepcion.SetFocus
        Exit Sub

    End If

End Sub

Private Sub monedav_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    unidad1.SetFocus

End Sub

Private Sub monedav_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        costop.SetFocus
        Exit Sub

    End If

End Sub

Private Sub oferta_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    vecaja.SetFocus

End Sub

Private Sub oferta_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        vtaund.SetFocus
        Exit Sub

    End If

End Sub

Private Sub percepcion_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    monedac.SetFocus

End Sub

Private Sub percepcion_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        isc.SetFocus
        Exit Sub

    End If

End Sub

Private Sub peso_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    servicio.SetFocus

End Sub

Private Sub peso_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        serie.SetFocus
        Exit Sub

    End If

End Sub

Private Sub pesokgr_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    flete.SetFocus

End Sub

Private Sub pesokgr_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        comision.SetFocus
        Exit Sub

    End If

End Sub

Private Sub pproducto_KeyPress(KeyAscii As Integer)

End Sub

Private Sub presenta_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    familia.SetFocus

End Sub

Private Sub presenta_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        descorto.SetFocus
        Exit Sub

    End If

End Sub

Private Sub proveedor1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    xproveedor.SetFocus

End Sub

Private Sub proveedor1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        color.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_proveedor1

    End If

    If KeyCode = &H76 Then  'f7

        'tnprov.show 1
    End If

End Sub

Private Sub procedencia_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    proyecto.SetFocus

End Sub

Private Sub procedencia_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        proyecto.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        frmOtros.Visible = False
        consulta_procedencia

    End If

    If KeyCode = &H76 Then  'f7
        tprocedencia.Show 1

    End If

End Sub

Private Sub proyecto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    procedencia.SetFocus

End Sub

Private Sub proyecto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        procedencia.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        frmOtros.Visible = False
        consulta_proyecto

    End If

    If KeyCode = &H76 Then  'f7
        tproyecto.Show 1

    End If

End Sub

Private Sub puertoimpresion_KeyPress(KeyAscii As Integer)

    ' 05/06/207   grupoimpresion ''
    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    grupoimpresion.SetFocus
    grupoimpresion = Left(puertoimpresion.Text, 1)
   
End Sub

Private Sub pventa1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes
    margen1.SetFocus

End Sub

Private Sub pventa1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        factor1.SetFocus
        Exit Sub

    End If

End Sub

Private Sub pventa10_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub pventa11_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub pventa12_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub pventa13_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub pventa14_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub pventa15_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub pventa2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes
    margen2.SetFocus

End Sub

Private Sub pventa2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        factor2.SetFocus
        Exit Sub

    End If

End Sub

Private Sub pventa3_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes
    margen3.SetFocus

End Sub

Private Sub pventa3_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        factor3.SetFocus
        Exit Sub

    End If

End Sub

Private Sub pventa4_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    calcula_margenes
    margen4.SetFocus

End Sub

Private Sub pventa4_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        factor4.SetFocus
        Exit Sub

    End If

End Sub

Private Sub pventa5_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub pventa6_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub pventa7_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub pventa8_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub pventa9_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    calcula_margenes

End Sub

Private Sub rect398912_Click()

End Sub

Private Sub seccion_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    marca.SetFocus

End Sub

Private Sub seccion_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        subfamilia.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_seccion

    End If

    If KeyCode = &H76 Then  'f7
        tseccion.Show 1

    End If

End Sub

Private Sub serie_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    Peso.SetFocus

End Sub

Private Sub serie_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        xproveedor.SetFocus
        Exit Sub

    End If

End Sub

Private Sub servicio_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    vtaund.SetFocus

End Sub

Private Sub servicio_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        Peso.SetFocus
        Exit Sub

    End If

End Sub

Private Sub sexo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    procedencia.SetFocus

End Sub

Private Sub sexo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        procedencia.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        frmOtros.Visible = False
        consulta_sexo

    End If

    If KeyCode = &H76 Then  'f7
        tsexo.Show 1

    End If

End Sub

Private Sub subfamilia_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    seccion.SetFocus

End Sub

Private Sub subfamilia_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    If KeyCode = &H26 Then
        familia.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        consulta_subFamilia

    End If

    If KeyCode = &H76 Then  'f7
        If Len(familia) = 0 Then
            MsgBox "Debe existir Familia ", 48, "Aviso"
            Exit Sub

        End If

        found = busca_familia()

        If found = 0 Then
            MsgBox "No existe Familia", 48, "Aviso"
            familia.SetFocus
            Exit Sub

        End If

        tsubfami.familia = familia
        tsubfami.Show 1

    End If

End Sub

Private Sub tlocal_KeyPress(KeyAscii As Integer)
    KeyAscii = 0

    If KeyAscii <> 13 Then Exit Sub

    'ccosto.SetFocus
End Sub

Private Sub tlocal_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1

        'consulta_local
    End If

End Sub

Private Sub talla_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    sexo.SetFocus

End Sub

Private Sub talla_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        sexo.SetFocus
        Exit Sub

    End If

    If KeyCode = &H70 Then  'f1
        frmOtros.Visible = False
        consulta_tallar

    End If

    If KeyCode = &H76 Then  'f7
        ttalla.Show 1

    End If

End Sub

Private Sub unidad_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    factor.SetFocus

End Sub

Private Sub unidad_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        monedac.SetFocus
        Exit Sub

    End If

End Sub

Private Sub unidad1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    factor1.SetFocus

End Sub

Private Sub unidad1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        monedav.SetFocus
        Exit Sub

    End If

End Sub

Private Sub unidad2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    factor2.SetFocus

End Sub

Private Sub unidad2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        margen1.SetFocus
        Exit Sub

    End If

End Sub

Private Sub unidad3_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    factor3.SetFocus

End Sub

Private Sub unidad3_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        margen2.SetFocus
        Exit Sub

    End If

End Sub

Private Sub unidad4_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    factor4.SetFocus

End Sub

Private Sub unidad4_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        margen3.SetFocus
        Exit Sub

    End If

End Sub

Private Sub unidad5_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        margen4.SetFocus
        Exit Sub

    End If

End Sub

Private Sub vecaja_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    comision.SetFocus

End Sub

Private Sub vecaja_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        oferta.SetFocus
        Exit Sub

    End If

End Sub

Private Sub vtaund_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    oferta.SetFocus

End Sub

Private Sub vtaund_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        servicio.SetFocus
        Exit Sub

    End If

End Sub

Sub consulta_Familia()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Descripcio"
    Combo3.AddItem "familia"
    Combo3.ListIndex = 1
    opcion1 = "2"
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    found = ejecuta(1)
    dbGrid1.SetFocus

End Sub

Sub consulta_subFamilia()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Descripcio"
    Combo3.AddItem "Subfamilia"
    Combo3.ListIndex = 1

    opcion1 = "3"
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    found = ejecuta(1)
    dbGrid1.SetFocus

End Sub

Sub consulta_seccion()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "seccion"
    Combo3.AddItem "Descripcio"
    Combo3.ListIndex = 2
    opcion1 = "4"
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    found = ejecuta(1)
    dbGrid1.SetFocus

End Sub

Sub consulta_margen()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Margen"
    Combo3.AddItem "Descripcio"
    Combo3.ListIndex = 2
    opcion1 = "16"
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    found = ejecuta(1)
    dbGrid1.SetFocus

End Sub

Sub consulta_local()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Codigo"
    Combo3.AddItem "Nombre"
    Combo3.ListIndex = 2
    opcion1 = "190"
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    found = ejecuta(1)
    dbGrid1.SetFocus

End Sub

Sub consulta_marca()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Descripcio"
    Combo3.AddItem "marca"
    Combo3.ListIndex = 1
    opcion1 = "5"
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    found = ejecuta(1)
    dbGrid1.SetFocus

End Sub

Sub consulta_fabrica()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Descripcio"
    Combo3.AddItem "codigo"
    Combo3.ListIndex = 1
    opcion1 = "6"
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    found = ejecuta(1)

    dbGrid1.SetFocus

End Sub

''17/07/2017 kenyo tienda ropa opciones producto

Sub consulta_sexo()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Descripcio"
    Combo3.AddItem "sexo"
    Combo3.ListIndex = 1
    opcion1 = "51"
    Frame1.Visible = True
    Frame1.Enabled = True
    found = ejecuta(1)
    buffer = ""

    dbGrid1.SetFocus

End Sub

Sub consulta_tallar()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Descripcio"
    Combo3.AddItem "talla"
    Combo3.ListIndex = 1
    opcion1 = "50"
    Frame1.Visible = True
    Frame1.Enabled = True
    found = ejecuta(1)
    buffer = ""

    dbGrid1.SetFocus

End Sub

Sub consulta_procedencia()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Descripcio"
    Combo3.AddItem "procedencia"
    Combo3.ListIndex = 1
    opcion1 = "52"
    Frame1.Visible = True
    Frame1.Enabled = True
    found = ejecuta(1)
    buffer = ""

    dbGrid1.SetFocus

End Sub

Sub consulta_proyecto()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Descripcio"
    Combo3.AddItem "proyecto"
    Combo3.ListIndex = 1
    opcion1 = "53"
    Frame1.Visible = True
    Frame1.Enabled = True
    found = ejecuta(1)
    buffer = ""

    dbGrid1.SetFocus

End Sub

''17/07/2017 kenyo tienda ropa opciones producto

Sub consulta_categoria()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Descripcio"
    Combo3.AddItem "Categoria"
    Combo3.ListIndex = 1
    opcion1 = "7"
    Frame1.Visible = True
    Frame1.Enabled = True
    found = ejecuta(1)
    buffer = ""

    dbGrid1.SetFocus

End Sub

Sub consulta_talla()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Descripcio"
    Combo3.AddItem "Linea"
    Combo3.ListIndex = 1

    opcion1 = "8"
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    found = ejecuta(1)

    dbGrid1.SetFocus

End Sub

Sub consulta_color()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Descripcio"
    Combo3.AddItem "Color"
    Combo3.ListIndex = 0

    opcion1 = "9"
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    found = ejecuta(1)

    dbGrid1.SetFocus

End Sub

Sub consulta_formulacion()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Descripcio"
    Combo3.ListIndex = 0

    opcion1 = "39"
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    found = ejecuta(1)

    dbGrid1.SetFocus

End Sub

Sub consulta_proveedor()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "Nombre"
    Combo3.AddItem "codigo"
    Combo3.ListIndex = 1
    opcion1 = "101"
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    found = ejecuta(1)
    dbGrid1.SetFocus

End Sub

Sub consulta_proveedor1()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Nombre"
    Combo3.AddItem "codigo"
    Combo3.ListIndex = 1
    opcion1 = "10"
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    found = ejecuta(1)
    dbGrid1.SetFocus

End Sub

Sub consulta_proveedor2()

    Dim found As Integer

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "Nombre"
    Combo3.AddItem "codigo"
    Combo3.ListIndex = 1

    opcion1 = "11"
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    found = ejecuta(1)

    dbGrid1.SetFocus

End Sub

Function busca_familia()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM familia where  familia='" & familia & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_familia = 1

        If Val(margen1) = 0 Then
            margen1 = "" & mytablex.Fields("margen1")

        End If

        If Val(margen2) = 0 Then
            margen2 = "" & mytablex.Fields("margen2")

        End If

        If Val(margen3) = 0 Then
            margen3 = "" & mytablex.Fields("margen3")

        End If

        If Val(margen4) = 0 Then
            margen4 = "" & mytablex.Fields("margen4")

        End If

        If Val(margen5) = 0 Then
            margen5 = "" & mytablex.Fields("margen5")

        End If

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_subfamilia()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM subfamil where  familia='" & familia & "' and subfamilia='" & subfamilia & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_subfamilia = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_seccion()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM seccion where  seccion='" & seccion & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        busca_seccion = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_categoria()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM categori where  categoria='" & categoria & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        busca_categoria = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_marca()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM marca where  marca='" & marca & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        busca_marca = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_margen()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM margen where  margen='" & Trim(margen) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        busca_margen = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_color()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM color where  color='" & color & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        busca_color = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

''17/07/2017 kenyo tienda ropa opciones producto

Function busca_talla()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM talla where  talla='" & talla & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        busca_talla = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_sexo()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM sexo where  sexo='" & sexo & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        busca_sexo = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_procedencia()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM procedencia where  procedencia='" & procedencia & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        busca_procedencia = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_proyecto()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM proyecto where  proyecto='" & proyecto & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        busca_proyecto = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

''17/07/2017 kenyo tienda ropa opciones producto

Function busca_lineatalla()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM linea where  linea='" & lineatalla & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_lineatalla = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function busca_parame(sw As Integer)

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM parame where  codigo='01'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If "" & mytablex.Fields("plocal") = "S" Then
            Label33.Visible = True
            local2.Visible = True

        End If

        If sw = 0 Then
            If insumo.Value = 0 Then
                sdx = Val("" & mytablex.Fields("producto")) + 1
                codigo = "" & sdx

            End If

            If insumo.Value = 1 Then
                sdx = Val("" & mytablex.Fields("insumo")) + 1
                codigo = "I" & sdx

            End If

        End If

        If sw = 3 Then

            'bodega = "" & mytablex.Fields("bodega")
            'fsaldoini = "" & mytablex.Fields("saldoini")
        End If

        If sw = 2 Then
            igv = "" & mytablex.Fields("igv")

            'MsgBox "" & igv
        End If

        If sw = 1 Then
            If insumo.Value = 0 Then
                If IsNumeric(codigo) Then
                    'mytablex.Edit
                    mytablex.Fields("producto") = Trim(codigo)
                    mytablex.Update

                End If

            End If

            If insumo.Value = 1 Then
                If IsNumeric(Mid$(codigo, 2, Len(codigo))) Then
                    'mytablex.Edit
                    mytablex.Fields("insumo") = Mid$(codigo, 2, Len(codigo))
                    mytablex.Update

                End If

            End If

        End If

        busca_parame = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function valida_barras(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM producto where  barras='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If Trim(codigo) <> Trim("" & mytablex.Fields("producto")) Then
            MsgBox "Ya existe Codigo Barras en codigo:" & mytablex.Fields("producto"), 48, "Aviso"
            valida_barras = 1
            mytablex.Close
            Exit Function

        End If

    End If

    mytablex.Close
    mytablex.Open "SELECT * FROM productb where  barras='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If Trim(codigo) <> Trim("" & mytablex.Fields("producto")) Then
            MsgBox "Ya existe Codigo Barras en codigo:" & mytablex.Fields("producto"), 48, "Aviso"
            valida_barras = 1
            mytablex.Close
            Exit Function

        End If

    End If

    mytablex.Close

End Function

Sub borrar_barras()

End Sub

Function grabar_barras()

    Dim found     As Integer

    Dim rconsulta As New ADODB.Recordset

    On Error GoTo cmd3_error

    If Frame2.Caption = "LOTES" Or Frame2.Caption = "NUMERO SERIES" Then
        rconsulta.AddNew
        rconsulta.Fields("descripcio") = "" & barras2
        rconsulta.Fields("producto") = "" & Trim(codigo)
        rconsulta.Update
        grabar_barras = 1
        Exit Function

    End If

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open "Select * from Productb where producto='" & Trim(codigo) & "' and barras='" & barras2 & "'", cn, adOpenStatic, adLockOptimistic

    If rconsulta.RecordCount = 0 Then
        rconsulta.AddNew
        rconsulta.Fields("barras") = "" & barras2
        rconsulta.Fields("producto") = "" & Trim(codigo)
        rconsulta.Update
        grabar_barras = 1
        rconsulta.Close

    End If

    Exit Function
cmd3_error:
    MsgBox "Error en grabar Barras " + error$, 48, "Aviso"
    Exit Function

End Function

Function valida_barras2(buf As String, buf2 As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM productb where  barras='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        valida_barras2 = 1
        buf2 = "" & mytablex.Fields("producto")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function valida_barras20(buf As String, buf2 As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM producto where  barras='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        valida_barras20 = 1
        buf2 = "" & mytablex.Fields("producto")

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Function existe_proveedor(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM proveedo where  codigo='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        existe_proveedor = 1

    End If

    '------------------------------------- ------------
    mytablex.Close

End Function

Sub calcula_margenes()

    Dim sdx     As Double

    Dim sdx1    As Double

    Dim sdx2    As Double

    Dim acostou As String

    On Error GoTo cmd786_err

    sdx = Val(costou) + Val(flete)
    acostou = "" & sdx

    If Val(factor) <= 0 Then
        factor = "1"

    End If

    If Val(factor1) <= 0 Then
        factor1 = "1"

    End If

    If Val(costou) = 0 And Val(costop) = 0 Then
        margen1 = "0"
        margen2 = "0"
        margen3 = "0"
        margen4 = "0"
        margen5 = "0"
        margen6 = "0"
        margen7 = "0"
        margen8 = "0"
        margen9 = "0"
        margen10 = "0"
   
        margen11 = "0"
        margen12 = "0"
        margen13 = "0"
        margen14 = "0"
        margen15 = "0"
        pone_margen

        Exit Sub

    End If

    pone_margen

    If monedac = "S" Then
        If monedav = "D" Then
            sdx = Val(acostou) / busca_cambio()

            If sdx <= 0 Then
                sdx = 1

            End If

            acostou = "" & sdx

        End If

    End If

    If monedac = "D" Then
        If monedav = "S" Then
            sdx = Val(acostou) * busca_cambio()

            If sdx <= 0 Then
                sdx = 1

            End If

            acostou = "" & sdx

        End If

    End If
          
    If Val(acostou) > 0 And Val(pventa1) > 0 And Val(factor1) > 0 Then 'calculando margenes
        sdx = (Val(acostou))
        sdx = sdx * Val(factor1)
        sdx1 = Val(pventa1) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen1 = Format(sdx2, "0.00")
        GoTo siguiente1

    End If

    If Val(margen1) > 0 And Val(pventa1) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen1) / 100
        sdx = sdx * Val(factor1)
        pventa1 = Format(sdx, "0.00")
        GoTo siguiente1

    End If

    If Val(acostou) <= 0 And Val(pventa1) > 0 And Val(margen1) > 0 Then
        sdx = Val(pventa1) / (1 + (Val(margen1) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente1

    End If
       
siguiente1:

    If Val(acostou) > 0 And Val(pventa2) > 0 And Val(factor2) > 0 Then 'calculando margenes
        sdx = (Val(acostou))
        sdx = sdx * Val(factor2)
        sdx1 = Val(pventa2) '/ Val(factor2)
        sdx2 = (Val(sdx1) - sdx) * 100 / sdx
        margen2 = Format(sdx2, "0.00")
        GoTo siguiente2

    End If

    If Val(margen2) > 0 And Val(pventa2) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen2) / 100
        sdx = sdx * Val(factor2)
        pventa2 = Format(sdx, "0.00")
        GoTo siguiente2

    End If

    If Val(acostou) <= 0 And Val(pventa2) > 0 And Val(margen2) > 0 Then
        sdx = Val(pventa2) / (1 + (Val(margen2) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente2

    End If

siguiente2:

    If Val(acostou) > 0 And Val(pventa3) > 0 And Val(factor3) > 0 Then 'calculando margenes
        sdx = (Val(acostou))  '/ Val(factor))
        sdx = sdx * Val(factor3)
        sdx1 = Val(pventa3) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen3 = Format(sdx2, "0.00")
        GoTo siguiente3

    End If

    If Val(margen3) > 0 And Val(pventa3) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen3) / 100
        sdx = sdx * Val(factor3)
        pventa3 = Format(sdx, "0.00")
        GoTo siguiente3

    End If

    If Val(acostou) <= 0 And Val(pventa3) > 0 And Val(margen3) > 0 Then
        sdx = Val(pventa3) / (1 + (Val(margen3) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente3

    End If

siguiente3:

    If Val(acostou) > 0 And Val(pventa4) > 0 And Val(factor4) > 0 Then 'calculando margenes
        sdx = (Val(acostou)) '/ Val(factor))
        sdx = sdx * Val(factor4)
        sdx1 = Val(pventa4) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen4 = Format(sdx2, "0.00")
        GoTo siguiente4

    End If

    If Val(margen4) > 0 And Val(pventa4) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen4) / 100
        sdx = sdx * Val(factor4)
        pventa4 = Format(sdx, "0.00")
        GoTo siguiente4

    End If

    If Val(acostou) <= 0 And Val(pventa4) > 0 And Val(margen4) > 0 Then
        sdx = Val(pventa4) / (1 + (Val(margen4) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente4

    End If

siguiente4:

    '''15/09/2017 KENYO Correccion margen de utilidad lista de precios
    '      If Val(acostou) > 0 And Val(pventa5) > 0 And Val(factor5) > 0 Then 'calculando margenes
    '          sdx = (Val(acostou) / Val(factor))
    '          sdx = sdx * Val(factor5)
    '          sdx1 = Val(pventa5) '/ Val(factor1)
    '          sdx2 = (sdx1 - sdx) * 100 / sdx
    '          margen5 = Format(sdx2, "0.00")
    '          GoTo siguiente5
    '       End If
    '       If Val(margen5) > 0 And Val(pventa5) <= 0 And Val(acostou) > 0 Then
    '          sdx = Val(acostou) + Val(acostou) * Val(margen5) / 100
    '          sdx = sdx * Val(factor5)
    '          pventa5 = Format(sdx, "0.00")
    '          GoTo siguiente5
    '       End If
    '       If Val(acostou) <= 0 And Val(pventa5) > 0 And Val(margen5) > 0 Then
    '          sdx = Val(pventa5) / (1 + (Val(margen5) / 100))
    '          costou = Format(sdx, "0.0000")
    '          GoTo siguiente5
    '       End If
    '
    'siguiente5:
    '       If Val(acostou) > 0 And Val(pventa6) > 0 And Val(factor6) > 0 Then 'calculando margenes
    '          sdx = (Val(acostou) / Val(factor))
    '          sdx = sdx * Val(factor6)
    '          sdx1 = Val(pventa6) '/ Val(factor1)
    '          sdx2 = (sdx1 - sdx) * 100 / sdx
    '          margen6 = Format(sdx2, "0.00")
    '          GoTo siguiente6
    '       End If
    '       If Val(margen6) > 0 And Val(pventa6) <= 0 And Val(acostou) > 0 Then
    '          sdx = Val(acostou) + Val(acostou) * Val(margen6) / 100
    '          sdx = sdx * Val(factor6)
    '          pventa6 = Format(sdx, "0.00")
    '          GoTo siguiente6
    '       End If
    '       If Val(acostou) <= 0 And Val(pventa6) > 0 And Val(margen6) > 0 Then
    '          sdx = Val(pventa6) / (1 + (Val(margen6) / 100))
    '          costou = Format(sdx, "0.0000")
    '          GoTo siguiente6
    '       End If
    'siguiente6:
    'If Val(acostou) > 0 And Val(pventa7) > 0 And Val(factor7) > 0 Then 'calculando margenes
    '          sdx = (Val(acostou) / Val(factor))
    '          sdx = sdx * Val(factor7)
    '          sdx1 = Val(pventa7) '/ Val(factor1)
    '          sdx2 = (sdx1 - sdx) * 100 / sdx
    '          margen7 = Format(sdx2, "0.00")
    '          GoTo siguiente7
    '       End If
    '       If Val(margen7) > 0 And Val(pventa7) <= 0 And Val(acostou) > 0 Then
    '          sdx = Val(acostou) + Val(acostou) * Val(margen7) / 100
    '          sdx = sdx * Val(factor7)
    '          pventa7 = Format(sdx, "0.00")
    '          GoTo siguiente7
    '       End If
    '       If Val(acostou) <= 0 And Val(pventa7) > 0 And Val(margen7) > 0 Then
    '          sdx = Val(pventa7) / (1 + (Val(margen7) / 100))
    '          costou = Format(sdx, "0.0000")
    '          GoTo siguiente7
    '       End If
    'siguiente7:
    'If Val(costou) > 0 And Val(pventa8) > 0 And Val(factor8) > 0 Then 'calculando margenes
    '          sdx = (Val(acostou) / Val(factor))
    '          sdx = sdx * Val(factor8)
    '          sdx1 = Val(pventa8) '/ Val(factor1)
    '          sdx2 = (sdx1 - sdx) * 100 / sdx
    '          margen8 = Format(sdx2, "0.00")
    '          GoTo siguiente8
    '       End If
    '       If Val(margen8) > 0 And Val(pventa8) <= 0 And Val(acostou) > 0 Then
    '          sdx = Val(acostou) + Val(acostou) * Val(margen8) / 100
    '          sdx = sdx * Val(factor8)
    '          pventa8 = Format(sdx, "0.00")
    '          GoTo siguiente8
    '       End If
    '       If Val(acostou) <= 0 And Val(pventa8) > 0 And Val(margen8) > 0 Then
    '          sdx = Val(pventa8) / (1 + (Val(margen8) / 100))
    '          costou = Format(sdx, "0.0000")
    '          GoTo siguiente8
    '       End If
    'siguiente8:
    'If Val(acostou) > 0 And Val(pventa9) > 0 And Val(factor9) > 0 Then 'calculando margenes
    '          sdx = (Val(acostou) / Val(factor))
    '          sdx = sdx * Val(factor9)
    '          sdx1 = Val(pventa9) '/ Val(factor1)
    '          sdx2 = (sdx1 - sdx) * 100 / sdx
    '          margen9 = Format(sdx2, "0.00")
    '          GoTo siguiente9
    '       End If
    '       If Val(margen9) > 0 And Val(pventa9) <= 0 And Val(acostou) > 0 Then
    '          sdx = Val(acostou) + Val(acostou) * Val(margen9) / 100
    '          sdx = sdx * Val(factor9)
    '          pventa9 = Format(sdx, "0.00")
    '          GoTo siguiente9
    '       End If
    '       If Val(acostou) <= 0 And Val(pventa9) > 0 And Val(margen9) > 0 Then
    '          sdx = Val(pventa9) / (1 + (Val(margen9) / 100))
    '          costou = Format(sdx, "0.0000")
    '          GoTo siguiente9
    '       End If
    'siguiente9:
    'If Val(acostou) > 0 And Val(pventa10) > 0 And Val(factor10) > 0 Then 'calculando margenes
    '          sdx = (Val(acostou) / Val(factor))
    '          sdx = sdx * Val(factor10)
    '          sdx1 = Val(pventa10) '/ Val(factor1)
    '          sdx2 = (sdx1 - sdx) * 100 / sdx
    '          margen10 = Format(sdx2, "0.00")
    '          GoTo siguiente10
    '       End If
    '       If Val(margen10) > 0 And Val(pventa10) <= 0 And Val(acostou) > 0 Then
    '          sdx = Val(acostou) + Val(acostou) * Val(margen10) / 100
    '          sdx = sdx * Val(factor10)
    '          pventa2 = Format(sdx, "0.00")
    '          GoTo siguiente10
    '       End If
    '       If Val(acostou) <= 0 And Val(pventa10) > 0 And Val(margen10) > 0 Then
    '          sdx = Val(pventa10) / (1 + (Val(margen10) / 100))
    '          costou = Format(sdx, "0.0000")
    '          GoTo siguiente10
    '       End If
    'siguiente10:
    
    If Val(acostou) > 0 And Val(pventa5) > 0 And Val(factor5) > 0 Then 'calculando margenes
        sdx = (Val(acostou)) '/ Val(factor))
        sdx = sdx * Val(factor5)
        sdx1 = Val(pventa5) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen5 = Format(sdx2, "0.00")
        GoTo siguiente5

    End If

    If Val(margen5) > 0 And Val(pventa5) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen5) / 100
        sdx = sdx * Val(factor5)
        pventa5 = Format(sdx, "0.00")
        GoTo siguiente5

    End If

    If Val(acostou) <= 0 And Val(pventa5) > 0 And Val(margen5) > 0 Then
        sdx = Val(pventa5) / (1 + (Val(margen5) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente5

    End If

siguiente5:
    
    If Val(acostou) > 0 And Val(pventa6) > 0 And Val(factor6) > 0 Then 'calculando margenes
        sdx = (Val(acostou)) '/ Val(factor))
        sdx = sdx * Val(factor6)
        sdx1 = Val(pventa6) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen6 = Format(sdx2, "0.00")
        GoTo siguiente6

    End If

    If Val(margen6) > 0 And Val(pventa6) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen6) / 100
        sdx = sdx * Val(factor6)
        pventa6 = Format(sdx, "0.00")
        GoTo siguiente6

    End If

    If Val(acostou) <= 0 And Val(pventa6) > 0 And Val(margen6) > 0 Then
        sdx = Val(pventa6) / (1 + (Val(margen6) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente6

    End If

siguiente6:
       
    If Val(acostou) > 0 And Val(pventa7) > 0 And Val(factor7) > 0 Then 'calculando margenes
        sdx = (Val(acostou)) '/ Val(factor))
        sdx = sdx * Val(factor7)
        sdx1 = Val(pventa7) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen7 = Format(sdx2, "0.00")
        GoTo siguiente7

    End If

    If Val(margen7) > 0 And Val(pventa7) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen7) / 100
        sdx = sdx * Val(factor7)
        pventa7 = Format(sdx, "0.00")
        GoTo siguiente7

    End If

    If Val(acostou) <= 0 And Val(pventa7) > 0 And Val(margen7) > 0 Then
        sdx = Val(pventa7) / (1 + (Val(margen7) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente7

    End If
       
siguiente7:
       
    If Val(acostou) > 0 And Val(pventa8) > 0 And Val(factor8) > 0 Then 'calculando margenes
        sdx = (Val(acostou)) '/ Val(factor))
        sdx = sdx * Val(factor8)
        sdx1 = Val(pventa8) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen8 = Format(sdx2, "0.00")
        GoTo siguiente8

    End If

    If Val(margen8) > 0 And Val(pventa8) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen8) / 100
        sdx = sdx * Val(factor8)
        pventa8 = Format(sdx, "0.00")
        GoTo siguiente8

    End If

    If Val(acostou) <= 0 And Val(pventa8) > 0 And Val(margen8) > 0 Then
        sdx = Val(pventa8) / (1 + (Val(margen8) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente8

    End If

siguiente8:
       
    If Val(acostou) > 0 And Val(pventa9) > 0 And Val(factor9) > 0 Then 'calculando margenes
        sdx = (Val(acostou)) '/ Val(factor))
        sdx = sdx * Val(factor9)
        sdx1 = Val(pventa9) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen9 = Format(sdx2, "0.00")
        GoTo siguiente9

    End If

    If Val(margen9) > 0 And Val(pventa9) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen9) / 100
        sdx = sdx * Val(factor9)
        pventa9 = Format(sdx, "0.00")
        GoTo siguiente9

    End If

    If Val(acostou) <= 0 And Val(pventa9) > 0 And Val(margen9) > 0 Then
        sdx = Val(pventa9) / (1 + (Val(margen9) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente9

    End If
       
siguiente9:
  
    If Val(acostou) > 0 And Val(pventa10) > 0 And Val(factor10) > 0 Then 'calculando margenes
        sdx = (Val(acostou)) '/ Val(factor))
        sdx = sdx * Val(factor10)
        sdx1 = Val(pventa10) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen10 = Format(sdx2, "0.00")
        GoTo siguiente10

    End If

    If Val(margen10) > 0 And Val(pventa10) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen10) / 100
        sdx = sdx * Val(factor10)
        pventa10 = Format(sdx, "0.00")
        GoTo siguiente10

    End If

    If Val(acostou) <= 0 And Val(pventa10) > 0 And Val(margen10) > 0 Then
        sdx = Val(pventa10) / (1 + (Val(margen10) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente10

    End If

siguiente10:
    '''15/09/2017 KENYO Correccion margen de utilidad lista de precios

    If Val(acostou) > 0 And Val(pventa11) > 0 And Val(factor1) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor1)
        sdx1 = Val(pventa11) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen11 = Format(sdx2, "0.00")
        GoTo siguiente11

    End If

    If Val(margen11) > 0 And Val(pventa11) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen11) / 100
        sdx = sdx * Val(factor1)
        pventa11 = Format(sdx, "0.00")
        GoTo siguiente11

    End If

    If Val(acostou) <= 0 And Val(pventa11) > 0 And Val(margen11) > 0 Then
        sdx = Val(pventa11) / (1 + (Val(margen11) / 100))
        sdx = sdx * Val(factor1)
        costou = Format(sdx, "0.0000")
        GoTo siguiente11

    End If

siguiente11:

    If Val(acostou) > 0 And Val(pventa12) > 0 And Val(factor1) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor1)
        sdx1 = Val(pventa12) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen12 = Format(sdx2, "0.00")
        GoTo siguiente12

    End If

    If Val(margen12) > 0 And Val(pventa12) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen12) / 100
        sdx = sdx * Val(factor1)
        pventa12 = Format(sdx, "0.00")
        GoTo siguiente12

    End If

    If Val(acostou) <= 0 And Val(pventa12) > 0 And Val(margen12) > 0 Then
        sdx = Val(pventa12) / (1 + (Val(margen12) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente12

    End If

siguiente12:

    If Val(acostou) > 0 And Val(pventa13) > 0 And Val(factor1) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor1)
        sdx1 = Val(pventa13) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen13 = Format(sdx2, "0.00")
        GoTo siguiente13

    End If

    If Val(margen13) > 0 And Val(pventa13) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen13) / 100
        sdx = sdx * Val(factor1)
        pventa13 = Format(sdx, "0.00")
        GoTo siguiente13

    End If

    If Val(acostou) <= 0 And Val(pventa13) > 0 And Val(margen13) > 0 Then
        sdx = Val(pventa13) / (1 + (Val(margen13) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente13

    End If

siguiente13:

    If Val(acostou) > 0 And Val(pventa14) > 0 And Val(factor1) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor1)
        sdx1 = Val(pventa14) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen14 = Format(sdx2, "0.00")
        GoTo siguiente14

    End If

    If Val(margen14) > 0 And Val(pventa14) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen14) / 100
        sdx = sdx * Val(factor1)
        pventa14 = Format(sdx, "0.00")
        GoTo siguiente14

    End If

    If Val(acostou) <= 0 And Val(pventa14) > 0 And Val(margen14) > 0 Then
        sdx = Val(pventa14) / (1 + (Val(margen14) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente14

    End If

siguiente14:

    If Val(acostou) > 0 And Val(pventa15) > 0 And Val(factor1) > 0 Then 'calculando margenes
        sdx = (Val(acostou) / Val(factor))
        sdx = sdx * Val(factor1)
        sdx1 = Val(pventa15) '/ Val(factor1)
        sdx2 = (sdx1 - sdx) * 100 / sdx
        margen15 = Format(sdx2, "0.00")
        GoTo siguiente15

    End If

    If Val(margen15) > 0 And Val(pventa15) <= 0 And Val(acostou) > 0 Then
        sdx = Val(acostou) + Val(acostou) * Val(margen15) / 100
        sdx = sdx * Val(factor1)
        pventa15 = Format(sdx, "0.00")
        GoTo siguiente15

    End If

    If Val(acostou) <= 0 And Val(pventa15) > 0 And Val(margen15) > 0 Then
        sdx = Val(pventa15) / (1 + (Val(margen15) / 100))
        costou = Format(sdx, "0.0000")
        GoTo siguiente15

    End If
       
siguiente15:
    cospaqu = Format(Val(costou) * Val(factor))
    cospaqp = Format(Val(costop) * Val(factor))
    cospaqi = Format(Val(costoini) * Val(factor))

    Exit Sub
cmd786_err:
    MsgBox "Error en calcula margenes", 48, "Aviso"
    Exit Sub

End Sub

Function busca_cambio() As Double

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    sdx = 1
    mytablex.Open "SELECT * FROM parame where  codigo='01'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        sdx = Val("" & mytablex.Fields("paricomp"))

        If Val("" & mytablex.Fields("paricomp")) <= 0 Then
            sdx = 1

        End If

    End If

    busca_cambio = sdx
    mytablex.Close

End Function

Sub pone_tamano()

    If opcion1 = "1" Then
        dbGrid1.columns(0).Width = 6000
        dbGrid1.columns(1).Width = 2000
        dbGrid1.columns(2).Width = 1000
        dbGrid1.columns(3).Width = 1000
        dbGrid1.columns(4).Width = 1000
        dbGrid1.columns(5).Width = 1000
        dbGrid1.columns(6).Width = 1000
        dbGrid1.columns(7).Width = 1000
        dbGrid1.columns(8).Width = 1000
        dbGrid1.SetFocus

    End If

    If opcion1 = "2" Or opcion1 = "27" Or opcion1 = "28" Or opcion1 = "29" Or opcion1 = "30" Or opcion1 = "31" Or opcion1 = "190" Then
        dbGrid1.columns(0).Width = 6000
        dbGrid1.columns(1).Width = 2000
        dbGrid1.SetFocus

    End If

    If opcion1 = "3" Then
        dbGrid1.columns(0).Width = 6000
        dbGrid1.columns(1).Width = 2000
        dbGrid1.columns(2).Width = 2000
        dbGrid1.SetFocus

    End If

    If opcion1 = "101" Or opcion1 = "4" Or opcion1 = "16" Or opcion1 = "5" Or opcion1 = "6" Or opcion1 = "7" Or opcion1 = "8" Or opcion1 = "9" Or opcion1 = "10" Or opcion1 = "11" Then
        dbGrid1.columns(0).Width = 6000
        dbGrid1.columns(1).Width = 2000
        dbGrid1.SetFocus

    End If

End Sub

Function borra_proveedor(buf1 As String, buf2 As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM codprov where  codigo='" & buf1 & "' and producto='" & buf2 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If MsgBox("Desea Borrar", 1, "Aviso") = 1 Then
            mytablex.Delete
            borra_proveedor = 1

        End If

    End If

    mytablex.Close

End Function

Sub carga_proveedor()

    Dim mytablex As New ADODB.Recordset

    Dim indx     As Integer

    xproveedor.Clear
    indx = 0
    mytablex.Open "SELECT proveedo.codigo,proveedo.nombre,codprov.codigop  FROM codprov,proveedo where codprov.codigo=proveedo.codigo and   codprov.producto='" & Trim(codigo) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            xproveedor.AddItem "" & mytablex.Fields("codigo") & " " & mytablex.Fields("nombre") & " " & mytablex.Fields("codigop")
            indx = indx + 1
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    If indx > 0 Then
        xproveedor.ListIndex = 0

    End If

End Sub

Private Sub xproveedor_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    serie.SetFocus

End Sub

Function borrar_codprov()

    Dim indx As Integer

    cn.Execute ("delete from codprov where producto='" & Trim(codigo) & "'")
    xproveedor.Clear

End Function

Function graba_rcodigo()

End Function

Sub carga_precios(buf As String)

    Dim mytablex As New ADODB.Recordset

    inicializa_precios
    mytablex.Open "SELECT * FROM precios where  producto='" & buf & "' and local='" & local2 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        pone_xprecio mytablex
        calcula_margenes

    End If

    mytablex.Close

    'pventa1.SetFocus
End Sub

Sub graba_precios()

    Dim mytablexx As New ADODB.Recordset

    Dim mytablex  As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    'si es modifica

    mytablexx.Open "SELECT * FROM vendedor where codigo='" & gusuario & "' and cprecios='S'", cn, adOpenKeyset, adLockOptimistic

    If mytablexx.RecordCount = 0 Then
        mytablexx.Close
        MsgBox "Usuario No autorizado,Solo puede cambiar local que esta autorizado ", 48, "Aviso"
        Exit Sub

    End If

    mytablexx.Close

    mytablex.Open "SELECT * FROM precios where  producto='" & Trim(codigo) & "' and local='" & "" & local2 & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        'mytablex.Edit
        graba_xprecio mytablex
        mytablex.Update
    Else
        mytablex.AddNew
        mytablex.Fields("producto") = Trim(codigo)
        mytablex.Fields("local") = local2
        graba_xprecio mytablex
        mytablex.Update

    End If

    mytablex.Close
    Exit Sub

    'If ordename = "NUEVO" Then
    mytabley.Open "SELECT * FROM tlocal ", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        mytabley.Close
        Exit Sub

    End If

    Do

        If mytabley.EOF Then Exit Do
        If "" & mytabley.Fields("codigo") <> local2 Then
            mytablex.Open "SELECT * FROM precios where  producto='" & Trim(codigo) & "' and local='" & "" & mytabley.Fields("codigo") & "'", cn, adOpenKeyset, adLockOptimistic

            If mytablex.RecordCount = 0 Then
                mytablex.AddNew
                mytablex.Fields("producto") = Trim(codigo)
                mytablex.Fields("local") = "" & mytabley.Fields("codigo")
                graba_xprecio mytablex
                mytablex.Update

            End If

            mytablex.Close

        End If
   
        mytabley.MoveNext
    Loop
    mytabley.Close
    Exit Sub
    'End If

End Sub

Sub graba_xprecio(mytablex As ADODB.Recordset)
    mytablex.Fields("pm1") = Val(pm1)
    mytablex.Fields("pm2") = Val(pm2)
    mytablex.Fields("pm3") = Val(pm3)
    mytablex.Fields("pm4") = Val(pm4)
    mytablex.Fields("pm5") = Val(pm5)
    mytablex.Fields("pm6") = Val(pm6)
    mytablex.Fields("pm7") = Val(pm7)
    mytablex.Fields("pm8") = Val(pm8)
    mytablex.Fields("pm9") = Val(pm9)
    mytablex.Fields("pm10") = Val(pm10)

    'mytablex.Fields("ccosto") = ccosto
    'mytablex.Fields("unidad1") = ""
    'mytablex.Fields("unidad2") = ""
    'mytablex.Fields("unidad3") = ""
    'mytablex.Fields("unidad4") = ""
    'mytablex.Fields("unidad5") = ""
    'mytablex.Fields("unidad6") = ""
    'mytablex.Fields("unidad7") = ""
    'mytablex.Fields("unidad8") = ""
    'mytablex.Fields("unidad9") = ""
    'mytablex.Fields("unidad10") = ""
    'mytablex.Fields("factor1") = 0
    'mytablex.Fields("factor2") = 0
    'mytablex.Fields("factor3") = 0
    'mytablex.Fields("factor4") = 0
    'mytablex.Fields("factor5") = 0
    'mytablex.Fields("factor6") = 0
    'mytablex.Fields("factor7") = 0
    'mytablex.Fields("factor8") = 0
    'mytablex.Fields("factor9") = 0
    'mytablex.Fields("factor10") = 0
    'mytablex.Fields("pventa1") = 0
    'mytablex.Fields("pventa2") = 0
    'mytablex.Fields("pventa3") = 0
    'mytablex.Fields("pventa4") = 0
    'mytablex.Fields("pventa5") = 0
    'mytablex.Fields("pventa6") = 0
    'mytablex.Fields("pventa7") = 0
    'mytablex.Fields("pventa8") = 0
    'mytablex.Fields("pventa9") = 0
    'mytablex.Fields("pventa10") = 0
    'mytablex.Fields("margen1") = 0
    'mytablex.Fields("margen2") = 0
    'mytablex.Fields("margen3") = 0
    'mytablex.Fields("margen4") = 0
    'mytablex.Fields("margen5") = 0
    'mytablex.Fields("margen6") = 0
    'mytablex.Fields("margen7") = 0
    'mytablex.Fields("margen8") = 0
    'mytablex.Fields("margen9") = 0
    'mytablex.Fields("margen10") = 0

    'If Val(pventa1) <> tpventa Then
    If Val("" & mytablex.Fields("pventa1")) <> Val(pventa1) Then
        mytablex.Fields("fechavp") = Format(Now, "dd/mm/yyyy")

    End If

    If Val("" & mytablex.Fields("pventa2")) <> Val(pventa2) Then
        mytablex.Fields("fechavp") = Format(Now, "dd/mm/yyyy")

    End If

    If Val("" & mytablex.Fields("pventa3")) <> Val(pventa3) Then
        mytablex.Fields("fechavp") = Format(Now, "dd/mm/yyyy")

    End If

    If Val("" & mytablex.Fields("pventa4")) <> Val(pventa4) Then
        mytablex.Fields("fechavp") = Format(Now, "dd/mm/yyyy")

    End If

    If Val("" & mytablex.Fields("pventa5")) <> Val(pventa5) Then
        mytablex.Fields("fechavp") = Format(Now, "dd/mm/yyyy")

    End If

    If Val("" & mytablex.Fields("pventa6")) <> Val(pventa6) Then
        mytablex.Fields("fechavp") = Format(Now, "dd/mm/yyyy")

    End If

    If Val("" & mytablex.Fields("pventa7")) <> Val(pventa7) Then
        mytablex.Fields("fechavp") = Format(Now, "dd/mm/yyyy")

    End If

    If Val("" & mytablex.Fields("pventa8")) <> Val(pventa8) Then
        mytablex.Fields("fechavp") = Format(Now, "dd/mm/yyyy")

    End If

    If Val("" & mytablex.Fields("pventa9")) <> Val(pventa9) Then
        mytablex.Fields("fechavp") = Format(Now, "dd/mm/yyyy")

    End If

    If Val("" & mytablex.Fields("pventa10")) <> Val(pventa10) Then
        mytablex.Fields("fechavp") = Format(Now, "dd/mm/yyyy")

    End If

    'If Val(factor1) > 0 And Len(unidad1) > 0 Then
    mytablex.Fields("unidad1") = unidad1
    mytablex.Fields("factor1") = Val(factor1)
    mytablex.Fields("pventa1") = Val(pventa1)
    mytablex.Fields("margen1") = Val(margen1)
    'End If

    'If Val(factor2) > 0 And Len(unidad2) > 0 Then
    mytablex.Fields("unidad2") = unidad2
    mytablex.Fields("factor2") = Val(factor2)
    mytablex.Fields("pventa2") = Val(pventa2)
    mytablex.Fields("margen2") = Val(margen2)
    'End If

    'If Val(factor3) > 0 And Len(unidad3) > 0 Then
    mytablex.Fields("unidad3") = unidad3
    mytablex.Fields("factor3") = Val(factor3)
    mytablex.Fields("pventa3") = Val(pventa3)
    mytablex.Fields("margen3") = Val(margen3)
    'End If

    'If Val(factor4) > 0 And Len(unidad4) > 0 Then
    mytablex.Fields("unidad4") = unidad4
    mytablex.Fields("factor4") = Val(factor4)
    mytablex.Fields("pventa4") = Val(pventa4)
    mytablex.Fields("margen1") = Val(margen1)
    'End If

    'If Val(factor5) > 0 And Len(unidad5) > 0 Then
    mytablex.Fields("unidad5") = unidad5
    mytablex.Fields("factor5") = Val(factor5)
    mytablex.Fields("pventa5") = Val(pventa5)
    mytablex.Fields("margen5") = Val(margen5)
    'End If

    'If Val(factor6) > 0 And Len(unidad6) > 0 Then
    mytablex.Fields("unidad6") = unidad6
    mytablex.Fields("factor6") = Val(factor6)
    mytablex.Fields("pventa6") = Val(pventa6)
    mytablex.Fields("margen1") = Val(margen1)
    'End If

    'If Val(factor7) > 0 And Len(unidad7) > 0 Then
    mytablex.Fields("unidad7") = unidad7
    mytablex.Fields("factor7") = Val(factor7)
    mytablex.Fields("pventa7") = Val(pventa7)
    mytablex.Fields("margen7") = Val(margen7)
    'End If

    'If Val(factor8) > 0 And Len(unidad8) > 0 Then
    mytablex.Fields("unidad8") = unidad8
    mytablex.Fields("factor8") = Val(factor8)
    mytablex.Fields("pventa8") = Val(pventa8)
    mytablex.Fields("margen8") = Val(margen8)
    'End If

    'If Val(factor9) > 0 And Len(unidad9) > 0 Then
    mytablex.Fields("unidad9") = unidad9
    mytablex.Fields("factor9") = Val(factor9)
    mytablex.Fields("pventa9") = Val(pventa9)
    mytablex.Fields("margen9") = Val(margen9)
    'End If

    'If Val(factor10) > 0 And Len(unidad10) > 0 Then
    mytablex.Fields("unidad10") = unidad10
    mytablex.Fields("factor10") = Val(factor10)
    mytablex.Fields("pventa10") = Val(pventa10)
    mytablex.Fields("margen10") = Val(margen10)
    'End If

    mytablex.Fields("minimo11") = Val(minimo11)
    mytablex.Fields("minimo12") = Val(minimo12)
    mytablex.Fields("minimo13") = Val(minimo13)
    mytablex.Fields("minimo14") = Val(minimo14)
    mytablex.Fields("minimo15") = Val(minimo15)
    mytablex.Fields("maximo11") = Val(maximo11)
    mytablex.Fields("maximo12") = Val(maximo12)
    mytablex.Fields("maximo13") = Val(maximo13)
    mytablex.Fields("maximo14") = Val(maximo14)
    mytablex.Fields("maximo15") = Val(maximo15)
    mytablex.Fields("pventa11") = Val(pventa11)
    mytablex.Fields("pventa12") = Val(pventa12)
    mytablex.Fields("pventa13") = Val(pventa13)
    mytablex.Fields("pventa14") = Val(pventa14)
    mytablex.Fields("pventa15") = Val(pventa15)
    mytablex.Fields("margen11") = Val(margen11)
    mytablex.Fields("margen12") = Val(margen12)
    mytablex.Fields("margen13") = Val(margen13)
    mytablex.Fields("margen14") = Val(margen14)
    mytablex.Fields("margen15") = Val(margen15)

    If IsDate(fechai11) Then
        mytablex.Fields("fechai11") = fechai11

    End If

    If IsDate(fechaf11) Then
        mytablex.Fields("fechaf11") = fechaf11

    End If

    If IsDate(fechaid) Then
        mytablex.Fields("fechaid") = fechaid

    End If

    If IsDate(fechafd) Then
        mytablex.Fields("fechafd") = fechafd

    End If

    mytablex.Fields("dscto") = Val(dscto)

End Sub

Sub pone_xprecio(mytablex As ADODB.Recordset)
    tpventa = Val("" & mytablex.Fields("pventa1"))
    unidad1 = "" & mytablex.Fields("unidad1")
    unidad2 = "" & mytablex.Fields("unidad2")
    unidad3 = "" & mytablex.Fields("unidad3")
    unidad4 = "" & mytablex.Fields("unidad4")
    unidad5 = "" & mytablex.Fields("unidad5")
    unidad6 = "" & mytablex.Fields("unidad6")
    unidad7 = "" & mytablex.Fields("unidad7")
    unidad8 = "" & mytablex.Fields("unidad8")
    unidad9 = "" & mytablex.Fields("unidad9")
    unidad10 = "" & mytablex.Fields("unidad10")
    factor1 = "" & mytablex.Fields("factor1")
    factor2 = "" & mytablex.Fields("factor2")
    factor3 = "" & mytablex.Fields("factor3")
    factor4 = "" & mytablex.Fields("factor4")
    factor5 = "" & mytablex.Fields("factor5")
    factor6 = "" & mytablex.Fields("factor6")
    factor7 = "" & mytablex.Fields("factor7")
    factor8 = "" & mytablex.Fields("factor8")
    factor9 = "" & mytablex.Fields("factor9")
    factor10 = "" & mytablex.Fields("factor10")
    pventa1 = "" & mytablex.Fields("pventa1")
    pventa2 = "" & mytablex.Fields("pventa2")
    pventa3 = "" & mytablex.Fields("pventa3")
    pventa4 = "" & mytablex.Fields("pventa4")
    pventa5 = "" & mytablex.Fields("pventa5")
    pventa6 = "" & mytablex.Fields("pventa6")
    pventa7 = "" & mytablex.Fields("pventa7")
    pventa8 = "" & mytablex.Fields("pventa8")
    pventa9 = "" & mytablex.Fields("pventa9")
    pventa10 = "" & mytablex.Fields("pventa10")
    margen1 = "" & mytablex.Fields("margen1")
    margen2 = "" & mytablex.Fields("margen2")
    margen3 = "" & mytablex.Fields("margen3")
    margen4 = "" & mytablex.Fields("margen4")
    margen5 = "" & mytablex.Fields("margen5")
    margen6 = "" & mytablex.Fields("margen6")
    margen7 = "" & mytablex.Fields("margen7")
    margen8 = "" & mytablex.Fields("margen8")
    margen9 = "" & mytablex.Fields("margen9")
    margen10 = "" & mytablex.Fields("margen10")
    minimo11 = "" & mytablex.Fields("minimo11")
    minimo12 = "" & mytablex.Fields("minimo12")
    minimo13 = "" & mytablex.Fields("minimo13")
    minimo14 = "" & mytablex.Fields("minimo14")
    minimo15 = "" & mytablex.Fields("minimo15")
    maximo11 = "" & mytablex.Fields("maximo11")
    maximo12 = "" & mytablex.Fields("maximo12")
    maximo13 = "" & mytablex.Fields("maximo13")
    maximo14 = "" & mytablex.Fields("maximo14")
    maximo15 = "" & mytablex.Fields("maximo15")
    pventa11 = "" & mytablex.Fields("pventa11")
    pventa12 = "" & mytablex.Fields("pventa12")
    pventa13 = "" & mytablex.Fields("pventa13")
    pventa14 = "" & mytablex.Fields("pventa14")
    pventa15 = "" & mytablex.Fields("pventa15")
    margen11 = "" & mytablex.Fields("margen11")
    margen12 = "" & mytablex.Fields("margen12")
    margen13 = "" & mytablex.Fields("margen13")
    margen14 = "" & mytablex.Fields("margen14")
    margen15 = "" & mytablex.Fields("margen15")
    fechai11 = "" & mytablex.Fields("fechai11")
    fechaf11 = "" & mytablex.Fields("fechaf11")
    fechaid = "" & mytablex.Fields("fechaid")
    fechafd = "" & mytablex.Fields("fechafd")
    dscto = "" & mytablex.Fields("dscto")
    'ccosto = "" & mytablex.Fields("ccosto")

    pm1 = "" & mytablex.Fields("pm1")
    pm2 = "" & mytablex.Fields("pm2")
    pm3 = "" & mytablex.Fields("pm3")
    pm4 = "" & mytablex.Fields("pm4")
    pm5 = "" & mytablex.Fields("pm5")
    pm6 = "" & mytablex.Fields("pm6")
    pm7 = "" & mytablex.Fields("pm7")
    pm8 = "" & mytablex.Fields("pm8")
    pm9 = "" & mytablex.Fields("pm9")
    pm10 = "" & mytablex.Fields("pm10")

    If Len(unidad1) = 0 Then
        unidad1 = "UND"

    End If

    If Len(factor1) = 0 Then
        factor1 = "1"

    End If

End Sub

Sub inicializa_precios()
    cospaqu = ""
    cospaqp = ""
    cospaqi = ""
    unidad1 = "UND"
    unidad2 = ""
    unidad3 = ""
    unidad4 = ""
    unidad5 = ""
    unidad6 = ""
    unidad7 = ""
    unidad8 = ""
    unidad9 = ""
    unidad10 = ""
    'saldoini = ""

    pm1 = ""
    pm2 = ""
    pm3 = ""
    pm4 = ""
    pm5 = ""
    pm6 = ""
    pm7 = ""
    pm8 = ""
    pm9 = ""
    pm10 = ""

    factor1 = "1"
    factor2 = ""
    factor3 = ""
    factor4 = ""
    factor5 = ""
    factor6 = ""
    factor7 = ""
    factor8 = ""
    factor9 = ""
    factor10 = ""

    pventa1 = ""
    pventa2 = ""
    pventa3 = ""
    pventa4 = ""
    pventa5 = ""
    pventa6 = ""
    pventa7 = ""
    pventa8 = ""
    pventa9 = ""
    pventa10 = ""

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
    minimo11 = ""
    minimo12 = ""
    minimo13 = ""
    minimo14 = ""
    minimo15 = ""

    maximo11 = ""
    maximo12 = ""
    maximo13 = ""
    maximo14 = ""
    maximo15 = ""

    pventa11 = ""
    pventa12 = ""
    pventa13 = ""
    pventa14 = ""
    pventa15 = ""
    margen11 = ""
    margen12 = ""
    margen13 = ""
    margen14 = ""
    margen15 = ""
    fechai11 = ""
    fechaf11 = ""
    fechaid = ""
    fechafd = ""
    dscto = ""

    'ccosto = ""
End Sub

Sub hacer_barras()

    Dim X         As Integer, Y As Integer, z As Integer, pos As Integer

    Dim temp      As String

    Dim Codevalue As Integer

    Dim BarCode   As String

    Call equivalentvalue
    
    Picture1.Cls
    pos = 10
    BarCode = UCase(Barras.Text)

    For X = 1 To Len(BarCode)
        temp = Mid$(BarCode, X, 1)

        Select Case temp

            Case "0" To "9"
                Codevalue = Val(temp)

            Case "A" To "Z"
                Codevalue = Asc(temp) - 55

            Case "-"
                Codevalue = 36

            Case "."
                Codevalue = 37

            Case " "
                Codevalue = 38

            Case "$"
                Codevalue = 39

            Case "/"
                Codevalue = 40

            Case "+"
                Codevalue = 41

            Case "%"
                Codevalue = 42

            Case "*"
                Codevalue = 43

            Case Else
                Picture1.Cls
                Picture1.Print temp & " is not valid"
                Exit Sub

        End Select
    
        For Y = 1 To 9

            If Y / 2 = Int(Y / 2) Then
                pos = pos + 1 + (3 * Val(Mid$(ArrBarCode(Codevalue), Y, 1)))
            Else

                For z = 1 To 1 + (3 * Val(Mid$(ArrBarCode(Codevalue), Y, 1)))
                    Picture1.Line (pos, 1)-(pos, 50)
                    pos = pos + 1
                Next z

            End If

        Next Y

        pos = pos + 1
    Next X

    Picture1.CurrentX = Len(BarCode) * 7
    Picture1.Print BarCode

End Sub

Private Sub equivalentvalue()
    ArrBarCode(0) = "000110100"
    ArrBarCode(1) = "100100001"
    ArrBarCode(2) = "001100001"
    ArrBarCode(3) = "101100000"
    ArrBarCode(4) = "000110001"
    ArrBarCode(5) = "100110000"
    ArrBarCode(6) = "001110000"
    ArrBarCode(7) = "000100101"
    ArrBarCode(8) = "100100100"
    ArrBarCode(9) = "001100100"
    ArrBarCode(10) = "100001001"
    ArrBarCode(11) = "001001001"
    ArrBarCode(12) = "101001000"
    ArrBarCode(13) = "000011001"
    ArrBarCode(14) = "100011000"
    ArrBarCode(15) = "001011000"
    ArrBarCode(16) = "000001101"
    ArrBarCode(17) = "100001100"
    ArrBarCode(18) = "001001100"
    ArrBarCode(19) = "000011100"
    ArrBarCode(20) = "100000011"
    ArrBarCode(21) = "001000011"
    ArrBarCode(22) = "101000010"
    ArrBarCode(23) = "000010011"
    ArrBarCode(24) = "100010010"
    ArrBarCode(25) = "001010010"
    ArrBarCode(26) = "000000111"
    ArrBarCode(27) = "100000110"
    ArrBarCode(28) = "001000110"
    ArrBarCode(29) = "000010110"
    ArrBarCode(30) = "110000001"
    ArrBarCode(31) = "011000001"
    ArrBarCode(32) = "111000000"
    ArrBarCode(33) = "010010001"
    ArrBarCode(34) = "110010000"
    ArrBarCode(35) = "011010000"
    ArrBarCode(36) = "010000101"
    ArrBarCode(37) = "110000100"
    ArrBarCode(38) = "011000100"
    ArrBarCode(39) = "010101000"
    ArrBarCode(40) = "010100010"
    ArrBarCode(41) = "010001010"
    ArrBarCode(42) = "000101010"
    ArrBarCode(43) = "010010100"

End Sub

Sub carga_dbgrid4()

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As ADODB.Recordset

    Dim sw       As Integer

    Dim xbodega  As String

    Dim xsaldo   As Double

    Dim xbuf     As String

    Dim xcosto   As Double

    Dim xmargen  As Double

    Dim xcostou  As Double

    Dim xfactor  As Double

    Dim xxr      As String

    Dim xxi      As String

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

    xfactor = 1
    xbodega = "01"
    xsaldo = 0
    xcosto = 0
    sw = 0
    mytabley.Open "SELECT * FROM almacen where  local='01' and producto='" & Trim(dbGrid1.columns(1)) & "' and bodega='" & xbodega & "'", cn, adOpenKeyset, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        xsaldo = Val("" & mytabley.Fields("saldo"))

    End If

    mytabley.Close
    '---buscamos los datos de productos
    mytablex.Open "SELECT * FROM producto where  producto='" & Trim(dbGrid1.columns(1)) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        xcostou = Val("" & mytablex.Fields("costou"))
        xfactor = Val("" & mytablex.Fields("factor"))

    End If

    mytablex.Close

    mytablex.Open "SELECT * FROM precios where  producto='" & Trim(dbGrid1.columns(1)) & "' and local='01'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then

        xcosto = 0
   
        If Val("" & mytablex.Fields("factor1")) > 0 Then
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor1"))
            campo_precios(0).unidad = "" & mytablex.Fields("unidad1")
            campo_precios(0).factor = "" & mytablex.Fields("factor1")
            campo_precios(0).precio = "" & mytablex.Fields("pventa1")
            campo_precios(0).costo = "" & xcosto
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor1")))
            campo_precios(0).stock = "" & xbuf
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa1")) - xcosto) * 100) / xcosto

            End If

            campo_precios(0).margen = "" & xmargen

        End If

        '---------
        xcosto = 0

        If Val("" & mytablex.Fields("factor2")) > 0 Then
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor2"))
            campo_precios(1).unidad = "" & mytablex.Fields("unidad2")
            campo_precios(1).factor = "" & mytablex.Fields("factor2")
            campo_precios(1).precio = "" & mytablex.Fields("pventa2")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor2")))
            campo_precios(1).stock = "" & xbuf
            campo_precios(1).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa2")) - xcosto) * 100) / xcosto

            End If

            campo_precios(1).margen = "" & xmargen

        End If

        xcosto = 0

        If Val("" & mytablex.Fields("factor3")) > 0 Then
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor3"))
            campo_precios(2).unidad = "" & mytablex.Fields("unidad3")
            campo_precios(2).factor = "" & mytablex.Fields("factor3")
            campo_precios(2).precio = "" & mytablex.Fields("pventa3")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor3")))
            campo_precios(2).stock = "" & xbuf
            campo_precios(2).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa3")) - xcosto) * 100) / xcosto
                campo_precios(2).margen = "" & xmargen

            End If

            campo_precios(2).margen = "" & xmargen

        End If

        xcosto = 0

        If Val("" & mytablex.Fields("factor4")) > 0 Then
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor4"))
            campo_precios(3).unidad = "" & mytablex.Fields("unidad4")
            campo_precios(3).factor = "" & mytablex.Fields("factor4")
            campo_precios(3).precio = "" & mytablex.Fields("pventa4")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor4")))
            campo_precios(3).stock = "" & xbuf
            campo_precios(4).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa4")) - xcosto) * 100) / xcosto

            End If

            campo_precios(3).margen = "" & xmargen

        End If

        xcosto = 0

        If Val("" & mytablex.Fields("factor5")) > 0 Then
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor5"))
            campo_precios(4).unidad = "" & mytablex.Fields("unidad5")
            campo_precios(4).factor = "" & mytablex.Fields("factor5")
            campo_precios(4).precio = "" & mytablex.Fields("pventa5")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor5")))
            campo_precios(4).stock = "" & xbuf
            campo_precios(4).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa5")) - xcosto) * 100) / xcosto

            End If

            campo_precios(4).margen = "" & xmargen

        End If

        xcosto = 0

        If Val("" & mytablex.Fields("factor6")) > 0 Then
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor6"))
   
            campo_precios(5).unidad = "" & mytablex.Fields("unidad6")
            campo_precios(5).factor = "" & mytablex.Fields("factor6")
            campo_precios(5).precio = "" & mytablex.Fields("pventa6")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor6")))
            campo_precios(5).stock = "" & xbuf
            campo_precios(5).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa5")) - xcosto) * 100) / xcosto

            End If

            campo_precios(5).margen = "" & xmargen

        End If

        xcosto = 0

        If Val("" & mytablex.Fields("factor7")) > 0 Then
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor7"))
   
            campo_precios(6).unidad = "" & mytablex.Fields("unidad7")
            campo_precios(6).factor = "" & mytablex.Fields("factor7")
            campo_precios(6).precio = "" & mytablex.Fields("pventa7")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor7")))
            campo_precios(6).stock = "" & xbuf
            campo_precios(6).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa7")) - xcosto) * 100) / xcosto
         
            End If

            campo_precios(6).margen = "" & xmargen

        End If
   
        xcosto = 0

        If Val("" & mytablex.Fields("factor8")) > 0 Then
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor8"))
   
            campo_precios(7).unidad = "" & mytablex.Fields("unidad8")
            campo_precios(7).factor = "" & mytablex.Fields("factor8")
            campo_precios(7).precio = "" & mytablex.Fields("pventa8")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor8")))
            campo_precios(7).stock = "" & xbuf
            campo_precios(7).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa8")) - xcosto) * 100) / xcosto

            End If

            campo_precios(7).margen = "" & xmargen

        End If

        xcosto = 0

        If Val("" & mytablex.Fields("factor9")) > 0 Then
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor9"))
            campo_precios(8).unidad = "" & mytablex.Fields("unidad9")
            campo_precios(8).factor = "" & mytablex.Fields("factor9")
            campo_precios(8).precio = "" & mytablex.Fields("pventa9")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor9")))
            campo_precios(8).stock = "" & xbuf
            campo_precios(8).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa9")) - xcosto) * 100) / xcosto

            End If

            campo_precios(8).margen = "" & xmargen

        End If

        xcosto = 0

        If Val("" & mytablex.Fields("factor10")) > 0 Then
            xcosto = xcostou / xfactor
            xcosto = xcosto * Val("" & mytablex.Fields("factor10"))
            campo_precios(9).unidad = "" & mytablex.Fields("unidad10")
            campo_precios(9).factor = "" & mytablex.Fields("factor10")
            campo_precios(9).precio = "" & mytablex.Fields("pventa10")
            xbuf = calcula_saldo(xsaldo, Val("" & mytablex.Fields("factor10")))
            campo_precios(9).stock = "" & xbuf
            campo_precios(9).costo = "" & xcosto
            xmargen = 0

            If xcosto > 0 Then
                xmargen = ((Val("" & mytablex.Fields("pventa10")) - xcosto) * 100) / xcosto

            End If

            campo_precios(9).margen = "" & xmargen

        End If

        sql_saldo_locales Trim(codigo)
        'margenes
        sw = 1

    End If

    mytablex.Close
    'mytablez.Close
    'dbgrid6.Refresh
    'Frame5.Visible = True
    'dbgrid6.SetFocus
    Exit Sub
cmd89111_err:
    MsgBox "Error en carga dbgrid4 " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub sql_saldo_locales(buf As String)
    'buf = "select * from almacen where producto='" & buf & "'"
    buf = "select Almacen.saldo,Bodega.nombre,almacen.bodega,Almacen.local from almacen left join bodega on almacen.bodega=bodega.codigo where almacen.producto='" & "" & dbGrid1.columns(1) & "' order by val(bodega.codigo)"
    'producto  left join precios on producto.producto=precios.producto  where precios.local='" & "" & my
    Data9.Connect = "foxpro 2.5;"
    Data9.DatabaseName = globaldir
    Data9.RecordSource = buf
    Data9.refresh

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
        sdx = Val("" & mytablex.Fields("producto")) + 1
        codigo = "" & sdx

    End If

    If sw = 1 Then
        If IsNumeric(codigo) Then
            mytablex.Fields("producto") = Trim(codigo)
            mytablex.Update

        End If

        mytablex.Close
        Exit Sub

    End If

    mytablex.Close
sigueb:
    mytablex.Open "select * from producto where producto='" & Trim(codigo) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Close
        sdx = sdx + 1
        codigo = "" & sdx
        GoTo sigueb
        Exit Sub

    End If

End Sub

Sub pone_margen()

    Dim mytablex As New ADODB.Recordset

    If Len(Trim(familia)) = 0 Then Exit Sub
    mytablex.Open "select * from familia where familia='" & Trim(familia) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
   
        If "" & mytablex.Fields("obliga") = "S" Then
      
            If Val("" & mytablex.Fields("margen1")) > 0 And Val(factor1) > 0 Then
                margen1 = "" & mytablex.Fields("margen1")

            End If

            If Val("" & mytablex.Fields("margen2")) > 0 And Val(factor2) > 0 Then
                margen2 = "" & mytablex.Fields("margen2")

            End If

            If Val("" & mytablex.Fields("margen3")) > 0 And Val(factor3) > 0 Then
                margen3 = "" & mytablex.Fields("margen3")

            End If

            If Val("" & mytablex.Fields("margen4")) > 0 And Val(factor4) > 0 Then
                margen4 = "" & mytablex.Fields("margen4")

            End If

            If Val("" & mytablex.Fields("margen5")) > 0 And Val(factor5) > 0 Then
                margen5 = "" & mytablex.Fields("margen5")

            End If

            If Val("" & mytablex.Fields("margen6")) > 0 And Val(factor6) > 0 Then
                margen6 = "" & mytablex.Fields("margen6")

            End If

            If Val("" & mytablex.Fields("margen7")) > 0 And Val(factor7) > 0 Then
                margen7 = "" & mytablex.Fields("margen7")

            End If

            If Val("" & mytablex.Fields("margen8")) > 0 And Val(factor8) > 0 Then
                margen8 = "" & mytablex.Fields("margen8")

            End If

            If Val("" & mytablex.Fields("margen9")) > 0 And Val(factor9) > 0 Then
                margen9 = "" & mytablex.Fields("margen9")

            End If

            If Val("" & mytablex.Fields("margen10")) > 0 And Val(factor10) > 0 Then
                margen10 = "" & mytablex.Fields("margen10")

            End If

        End If

    End If

    mytablex.Close

End Sub

Sub carga_impresoras()

    Dim I As Integer

    On Error GoTo cmd8912_err

    cboprinters.Clear
    cboprinters.AddItem "%"
    cboprinters1.Clear
    cboprinters1.AddItem "%"
    cboprinters2.Clear
    cboprinters2.AddItem "%"
    cboprinters3.Clear
    cboprinters3.AddItem "%"
    
    For I = 0 To Printers.count - 1
        cboprinters.AddItem Printers(I).DeviceName
        cboprinters1.AddItem Printers(I).DeviceName
        cboprinters2.AddItem Printers(I).DeviceName
        cboprinters3.AddItem Printers(I).DeviceName

        ' if this is the current printer, select it
        If Printers(I).DeviceName = Printer.DeviceName Then
            ' this indirectly executes ShowPrinterInfo
            cboprinters.ListIndex = I
            cboprinters1.ListIndex = I
            cboprinters2.ListIndex = I
            cboprinters3.ListIndex = I

        End If

    Next
    cboprinters.ListIndex = 0
    cboprinters1.ListIndex = 0
    cboprinters2.ListIndex = 0
    cboprinters3.ListIndex = 0
    Exit Sub
cmd8912_err:
    Exit Sub

End Sub

Sub actualiza_receta()

    On Error GoTo cmd9093_err

    cn.Execute ("update receta set precio=" & Val(costou) & ",total=cantidad*" & Val(costou) & " where productoi='" & codigo & "'")
    Exit Sub
cmd9093_err:
    MsgBox "Aviso en actualiza receta " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function convierte_barras(buf) As String

    Dim buf1 As String

    Dim I    As Integer

    Dim sdx  As Integer

    If flag_denisse = "0" Then Exit Function
    If Len(Trim(buf)) = 0 Then Exit Function
    buf1 = ""
    sdx = 18 - Len(buf)

    For I = 1 To sdx
        buf1 = buf1 & "0"
    Next I

    buf1 = buf1 & buf
    convierte_barras = buf1

End Function

Sub pone_fotonombre(mytablex As ADODB.Recordset)
    'On Error GoTo cm897888_err
    foto = LoadPicture()
    'fotonombre = globalpath & "\001d\06\grafico\temp.jpg" 'Trim("" & mytablex.Fields("fotonombre"))
    fotonombre = App.path & "\001d\06\grafico\temp.jpg" 'Trim("" & mytablex.Fields("fotonombre"))
    viewBMP mytablex, fotonombre

    If Len(fotonombre) > 0 Then
        If existe_archivo(fotonombre) > 0 Then
            foto = LoadPicture(fotonombre)

        End If

    End If

cm897888_err:
    Exit Sub

End Sub

Function valida_tecla(buf As String)

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM producto where  tecla='" & buf & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        If Trim(codigo) <> Trim("" & mytablex.Fields("producto")) Then
            MsgBox "Ya existe Tecla:" & mytablex.Fields("producto"), 48, "Aviso"
            valida_tecla = 1
            mytablex.Close
            Exit Function

        End If

    End If

    mytablex.Close

End Function

Function puede_modificar()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "SELECT * FROM vendedor where codigo='" & gusuario & "' and modificaproducto='N'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        mytablex.Close
        Exit Function

    End If

    mytablex.Close
    puede_modificar = 1

End Function

'13/08/2018 Integración FE - Pizzeria
'Cambios Pizzeria 24/05/2018
''' 11/12/2017 SubReceta
Sub actualiza_sumacostoreceta()

    Dim mytablex   As New ADODB.Recordset

    Dim mytablexyz As New ADODB.Recordset

    Dim suma       As Double

    If mytablex.State = 1 Then mytablex.Close
 
    mytablex.Open "SELECT * FROM receta  where  linea='' and  productoI='" & codigo & "' order by str(producto) ", cn, adOpenDynamic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            
            If mytablexyz.State = 1 Then mytablexyz.Close
            mytablexyz.Open "SELECT * FROM receta  where linea='' and  producto='" & mytablex.Fields("producto") & "' ", cn, adOpenDynamic, adLockOptimistic
            suma = 0

            If mytablexyz.RecordCount > 0 Then
                Do

                    If mytablexyz.EOF Then Exit Do
                    suma = suma + Val("" & mytablexyz.Fields("cantidad")) * Val("" & mytablexyz.Fields("precio"))
                    mytablexyz.MoveNext
                Loop
                cn.Execute ("UPDATE PRODUCTO SET COSTOU= '" & suma & "'  WHERE  PRODUCTO='" & mytablex.Fields("producto") & "'")
                    
            End If

            mytablexyz.Close
            mytablex.MoveNext
        Loop
        
    End If

    mytablex.Close
 
End Sub

''' 11/12/2017 SubReceta
'Cambios Pizzeria 24/05/2018
'13/08/2018 Integración FE - Pizzeria
