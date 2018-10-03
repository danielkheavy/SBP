VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ttordent 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ordenes de Trabajo"
   ClientHeight    =   9930
   ClientLeft      =   165
   ClientTop       =   -765
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   15960
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      Height          =   9975
      Left            =   15
      TabIndex        =   58
      Top             =   15
      Visible         =   0   'False
      Width           =   15015
      Begin VB.TextBox Text1 
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
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox Combo2 
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
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar"
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   240
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid dbgrid13 
         Height          =   8895
         Left            =   120
         TabIndex        =   62
         Top             =   840
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   15690
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   29
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
               LCID            =   10250
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
               LCID            =   10250
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pedidos"
      Height          =   8655
      Left            =   0
      TabIndex        =   49
      Top             =   45
      Visible         =   0   'False
      Width           =   14535
      Begin VB.TextBox pefechai 
         Height          =   495
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   54
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox pefechaf 
         Height          =   495
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   53
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtrar"
         Height          =   495
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   705
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enviar->Orden"
         Height          =   495
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   720
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid dbgrid12 
         Height          =   6375
         Left            =   120
         TabIndex        =   51
         Top             =   1320
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   11245
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   29
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
               LCID            =   10250
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
               LCID            =   10250
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
      Begin VB.Label Label17 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
         Height          =   495
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label18 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaFinal"
         Height          =   495
         Left            =   120
         TabIndex        =   56
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label nround 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   8520
         TabIndex        =   55
         Top             =   7800
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   8775
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   15015
      Begin VB.TextBox serie 
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
         MaxLength       =   4
         TabIndex        =   34
         Top             =   1800
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox numero 
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
         MaxLength       =   11
         TabIndex        =   33
         Top             =   2160
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox subtablapro 
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
         TabIndex        =   32
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox glosa 
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
         TabIndex        =   31
         Top             =   3240
         Width           =   3255
      End
      Begin VB.TextBox turno 
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
         TabIndex        =   30
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   10560
         Picture         =   "ttordent.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   480
         Width           =   1470
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   10560
         Picture         =   "ttordent.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Imprimir todo"
         Top             =   1560
         Width           =   1470
      End
      Begin VB.TextBox fecha 
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
         TabIndex        =   27
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox ordentrabajo 
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
         TabIndex        =   26
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox fechai 
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
         TabIndex        =   25
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox fechae 
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
         TabIndex        =   24
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox moneda 
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
         TabIndex        =   23
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox observa 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   2280
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   3600
         Width           =   5895
      End
      Begin VB.TextBox local1 
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
         MaxLength       =   11
         TabIndex        =   21
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2880
         Picture         =   "ttordent.frx":1194
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   375
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         Picture         =   "ttordent.frx":149E
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Left            =   4560
         TabIndex        =   48
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
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
         Height          =   375
         Left            =   4560
         TabIndex        =   47
         Top             =   2160
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label tablapro 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TF"
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
         Left            =   4560
         TabIndex        =   46
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoFormula"
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
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observa"
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
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Turno"
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
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   42
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OrdenTrabajo"
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
         TabIndex        =   41
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
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
         TabIndex        =   40
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaEntrega"
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
         TabIndex        =   39
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   38
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   37
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pedido"
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
         Top             =   6240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local"
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
         Left            =   4560
         TabIndex        =   35
         Top             =   1440
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Finaliza"
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
      Left            =   13320
      TabIndex        =   19
      Top             =   4560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Detalle<-Pedido"
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
      Left            =   13320
      TabIndex        =   17
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Aprobar"
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
      Left            =   13320
      TabIndex        =   10
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Detalle"
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
      Left            =   13320
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   15900
      TabIndex        =   2
      Top             =   0
      Width           =   15960
      Begin VB.ComboBox mostrar 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox ordenado 
         Height          =   315
         Left            =   8760
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   0
         Width           =   2415
      End
      Begin VB.ComboBox tipoformula 
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   0
         Width           =   2415
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
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "ttordent.frx":17A8
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Filtrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11520
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
         Picture         =   "ttordent.frx":29BA
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
         Picture         =   "ttordent.frx":3BCC
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
         Picture         =   "ttordent.frx":4DDE
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
         Picture         =   "ttordent.frx":5FF0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label viene 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8760
         TabIndex        =   18
         Top             =   480
         Width           =   105
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mostrar"
         Height          =   375
         Left            =   4080
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ordenado"
         Height          =   375
         Left            =   7800
         TabIndex        =   12
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoFormula"
         Height          =   375
         Left            =   4080
         TabIndex        =   11
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   13215
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7935
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   13996
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "Ordentrabajo"
            Caption         =   "OrdenTrabajo"
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
         BeginProperty Column02 
            DataField       =   "fechai"
            Caption         =   "Fechai"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "fechae"
            Caption         =   "fechae"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Moneda"
            Caption         =   "M"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Turno"
            Caption         =   "T"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "tablapro"
            Caption         =   "Tablapro"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Subtablapro"
            Caption         =   "Subtablapro"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Serie"
            Caption         =   "Serie"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "Numero"
            Caption         =   "Numero"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "Aprobado"
            Caption         =   "Aprobado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "Estado"
            Caption         =   "Estado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "Glosa"
            Caption         =   "Glosa"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1110.047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   329.953
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   420.095
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
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
      Begin VB.Menu rep093 
         Caption         =   "&0.Reporte Orden Trabajo"
      End
      Begin VB.Menu fdj7744 
         Caption         =   "&1.Reporte Orden Trabajo Formula"
      End
      Begin VB.Menu dk9893 
         Caption         =   "&2.Generador"
      End
      Begin VB.Menu mnuArchivoArray 
         Caption         =   "Novisible"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu paol844 
      Caption         =   "&ParteProduccion"
      Visible         =   0   'False
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "ttordent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txformxu As New ADODB.Recordset

Private Sub ajdu1_Click()

    If Frame4.Visible = True Then Exit Sub
    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Nuevo"
    cmdGuardar.Enabled = True
    habilita 1
    ordentrabajo.Enabled = False
    ordentrabajo = ""
    subtablapro.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    If Frame4.Visible = True Then Exit Sub
    buf = "" & txformxu.Fields("ordentrabajo")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + "" & txformxu.Fields("ordentrabajo"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    txformxu.Delete
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

    'djuer1_Click
End Sub

Private Sub cmdSave_Click()
    f8443_Click

End Sub

Private Sub Command2_Click()

    On Error GoTo cmfd5611_err

    'tordentd.ordentrabajod = "" & txformxu.Fields("ordentrabajod")
    tordentd.idx = "" & txformxu.Fields("ordentrabajo")
    tordentd.orlocal = Trim("" & txformxu.Fields("local"))
    tordentd.orserie = Trim("" & txformxu.Fields("serie"))
    tordentd.ornumero = Trim("" & txformxu.Fields("numero"))
    tordentd.Show 1
    Exit Sub
cmfd5611_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command3_Click()

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd901222_err

    mytablex.Open "select * from ordentrabajoD where ordentrabajo=" & "" & txformxu.Fields("ordentrabajo") & "", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "debe Ingresar Producto que se van a Fabricar ", 48, "Aviso"
        mytablex.Close

    End If

    mytablex.Close

    Select Case Trim("" & txformxu.Fields("estado"))

        Case "PLANIFICACION"

            If MsgBox("Desea Aprobar ", 1, "Aviso") <> 1 Then Exit Sub
            txformxu.Fields("aprobado") = "S"
            txformxu.Fields("estado") = "PRODUCCION"
            txformxu.Update
            txformxu.Requery
            MsgBox "Proceso Realizado", 48, "Aviso"

        Case "PRODUCCION"

            If MsgBox("Desea desAprobar ", 1, "Aviso") <> 1 Then Exit Sub
            txformxu.Fields("aprobado") = ""
            txformxu.Fields("estado") = "PLANIFICACION"
            txformxu.Update
            txformxu.Requery
            MsgBox "Proceso Realizado", 48, "Aviso"

    End Select

    Exit Sub
cmd901222_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command4_Click()
    filtro

End Sub

Private Sub Command5_Click()

    On Error GoTo cmd09012_err

    If Len("" & txformxu.Fields("ordentrabajo")) = 0 Then Exit Sub
    Frame4.Visible = True
    pefechai = Format(Now, "dd/mm/yyyy")
    pefechaf = Format(Now, "dd/mm/yyyy")
    pefechaf.SetFocus
    Exit Sub
cmd09012_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command6_Click()

    Dim sw       As Integer

    Dim buf      As String

    Dim buf1     As String

    Dim mytablez As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim sdx      As Double

    If Not IsDate(pefechai) Then Exit Sub
    If Not IsDate(pefechaf) Then Exit Sub

    'condicion busqueda
    buf = "select * from dpedidov where "
    buf = buf & "  fecha>='" & Format(pefechai, "YYYYMMDD") & "'"
    buf = buf & "  and fecha<='" & Format(pefechaf, "YYYYMMDD") & "' "
    mytablez.Open buf, cn, adOpenStatic, adLockOptimistic
    sw = 0
    cn.Execute ("delete from tmpedido")
    Do

        If mytablez.EOF Then Exit Do
        'If mytablez.Fields("dflag") = "S" Then
        '----------------------
        buf = "select * from dpedidov where local='" & "" & mytablez.Fields("local") & "' and tipo='" & "" & mytablez.Fields("tipo") & "' and serie='" & "" & mytablez.Fields("serie") & "' and numero='" & "" & mytablez.Fields("numero") & "'"
        mytablex.Open buf, cn, adOpenKeyset, adLockOptimistic
        Do

            If mytablex.EOF Then Exit Do
            mytabley.Open "select * from tmpedido where producto='" & "" & mytablex.Fields("producto") & "'", cn, adOpenKeyset, adLockOptimistic

            If mytabley.RecordCount = 0 Then
                mytabley.AddNew
                mytabley.Fields("producto") = Trim("" & mytablex.Fields("producto"))
                mytabley.Fields("descripcio") = Trim("" & mytablex.Fields("descripcio"))
                mytabley.Fields("cantidad") = Val("" & mytablex.Fields("cantidad"))
                mytabley.Fields("unidad") = Trim("" & mytablex.Fields("unidad"))
                mytabley.Fields("factor") = Val("" & mytablex.Fields("factor"))
                mytabley.Update
            Else
                mytabley.Fields("cantidad") = Val("" & mytabley.Fields("cantidad")) + Val("" & mytablex.Fields("cantidad"))
                mytabley.Update

            End If

            mytabley.Close
            mytablex.MoveNext
        Loop
        mytablex.Close
        '----------------------
        'End If
        mytablez.MoveNext
    Loop
    mytablex.Open "select * from tmpedido", cn, adOpenKeyset, adLockOptimistic
    Set dbgrid12.DataSource = mytablex
    dbgrid12.columns(0).Width = 1000
    dbgrid12.columns(1).Width = 4000
    dbgrid12.columns(2).Width = 1000
    dbgrid12.columns(3).Width = 1000
    dbgrid12.columns(4).Width = 1000
         
    dbgrid12.refresh
    sdx = 0
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + Val("" & mytablex.Fields("cantidad"))
        mytablex.MoveNext
    Loop
    nround = "" & sdx

End Sub

Private Sub Command7_Click()

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    mytablex.Open "select * from ordentrabajod where ordentrabajo=" & Trim("" & txformxu.Fields("ordentrabajo")), cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        MsgBox "Ya existen datos ", 48, "Aviso"
        mytablex.Close
        Exit Sub

    End If

    mytabley.Open "select * from tmpedido", cn, adOpenKeyset, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        mytablex.AddNew
        mytablex.Fields("producto") = Trim("" & mytabley.Fields("producto"))
        mytablex.Fields("descripcio") = Trim("" & mytabley.Fields("descripcio"))
        mytablex.Fields("unidad") = Trim("" & mytabley.Fields("unidad"))
        mytablex.Fields("factor") = Val("" & mytabley.Fields("factor"))
        mytablex.Fields("cantidad") = Val("" & mytabley.Fields("cantidad"))
        mytablex.Fields("bodega") = "01"
        mytablex.Fields("formula") = ""
        mytablex.Fields("ordentrabajo") = Trim("" & txformxu.Fields("ordentrabajo"))
        mytablex.Update
        mytabley.MoveNext
    Loop
    mytabley.Close

End Sub

Private Sub Command8_Click()

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd9012226_err

    mytablex.Open "select * from ordentrabajoD where ordentrabajo=" & "" & txformxu.Fields("ordentrabajo") & "", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        MsgBox "debe Ingresar Producto que se van a Fabricar ", 48, "Aviso"
        mytablex.Close
        Exit Sub

    End If

    mytablex.Close

    Select Case Trim("" & txformxu.Fields("estado"))

        Case "PRODUCCION"

            If MsgBox("Desea Cerrar ", 1, "Aviso") <> 1 Then Exit Sub
            txformxu.Fields("estado") = "TERMINADO"
            txformxu.Update
            txformxu.Requery
            MsgBox "Proceso Realizado", 48, "Aviso"

        Case "TERMINADO"

            If MsgBox("Desea desCerrar ", 1, "Aviso") <> 1 Then Exit Sub
            txformxu.Fields("estado") = "PRODUCCION"
            txformxu.Update
            txformxu.Requery
            MsgBox "Proceso Realizado", 48, "Aviso"

    End Select

    Exit Sub
cmd9012226_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub dbgrid13_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    If KeyCode = 27 Then
        Text1.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            subtablapro = Trim("" & dbgrid13.columns(1))
            Frame3.Visible = False
            subtablapro.SetFocus

        End If

        If opcion1 = "2" Then
            turno = Trim("" & dbgrid13.columns(1))
            Frame3.Visible = False
            turno.SetFocus

        End If

        If opcion1 = "3" Then
            local1 = Trim("" & dbgrid13.columns("local"))
            serie = Trim("" & dbgrid13.columns("serie"))
            Numero = Trim("" & dbgrid13.columns("numero"))
            Frame3.Visible = False
            serie.SetFocus

        End If

    End If

End Sub

Private Sub dk9893_Click()

    If Frame4.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "ordentrabajoc"
    reporgen.Show 1

End Sub

Sub prueba_reporte()

    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\formulacionesproducto.rpt", "")
End Sub

Private Sub Command1_Click()
    'If tipoformula = "%" Then
    '   MsgBox "Seleccione Tipo Formula ", 48, "Aviso"
    '   Exit Sub
    'Exit Sub
    Frame1.Visible = True
    Frame1.Enabled = True

    'buffer = ""
    opcion1 = "1"
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim cad As String

    'If opcion1 = "1" Then  'bodega
    'If Len(buffer) = 0 Then
    cad = "SELECT * from ordentrabajoc  where tablapro='" & tablapro & "'"

    If tipoformula <> "%" Then
        cad = cad & " and subtablapro='" & extra_loquesea1(tipoformula) & "' "

    End If

    If mostrar <> "%" Then
        cad = cad & " and estado='" & mostrar & "' "

    End If

    If viene = "ParteProduccion" Then
        cad = cad & " and aprobado='S'"

    End If
   
    If ordenado <> "%" Then
        cad = cad & " order by " & ordenado

    End If
   
    'End If
    'If Len(buffer) > 0 Then
    '   cad = "SELECT *  from ordentrabajoc   where  " & Combo1 & " like '" & buffer & "%'"
    'End If
    If txformxu.State = 1 Then txformxu.Close
    txformxu.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = txformxu
    'If txformxu.RecordCount > 0 Then
    dbGrid1.SetFocus

    'End If
    'End If
End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then

        'buffer.SetFocus
        'Exit Sub
    End If

    If KeyCode = 13 Then

        'formulacion = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'formulacion.SetFocus
        'formulacion_KeyPress 13
    End If

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

    If Frame2.Visible = True Then
        habilita 0
        Frame2.Visible = False
        dbGrid1.Enabled = True
        Exit Sub

    End If

    ttordent.Hide
    Unload ttordent

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    If Frame4.Visible = True Then Exit Sub
    buf = "" & txformxu.Fields("ordentrabajo")

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
    ordentrabajo.Enabled = False
    subtablapro.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fdj7744_Click()

    If Frame4.Visible = True Then Exit Sub
    reporte_orden_excell 1

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    If Frame4.Visible = True Then Exit Sub
    buf = "" & txformxu.Fields("ordentrabajo")

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
    ordentrabajo.Enabled = False
    'MsgBox "ABC"
    subtablapro.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()

    'agregar_menus
    If viene = "ParteProduccion" Then
        paol844.Visible = True
        ajdu1.Visible = False
        f8443.Visible = False
        bo712.Visible = False
        cmdAddEntry.Enabled = False
        cmdSave.Enabled = False
        cmdDelete.Enabled = False
        Command2.Visible = False
        Command3.Visible = False
        Command5.Visible = False
        Command8.Visible = True

    End If

    Command1_Click

End Sub

Private Sub Form_Load()

    Dim mytablex As New ADODB.Recordset

    'Combo1.Clear
    'Combo1.AddItem "Subtablapro"
    'Combo1.ListIndex = 0

    ordenado.Clear
    ordenado.AddItem "%"
    ordenado.AddItem "OrdenTrabajo"
    ordenado.AddItem "Fecha"
    ordenado.AddItem "FechaE"
    ordenado.ListIndex = 0

    mostrar.Clear
    mostrar.AddItem "%"
    mostrar.AddItem "PLANIFICACION"
    mostrar.AddItem "PRODUCCION"
    mostrar.AddItem "ANULADAS"
    mostrar.AddItem "TERMINADOS"
    mostrar.AddItem "ATRASADOS"
    mostrar.ListIndex = 0

    carga_tipoformula

End Sub

Sub inicializa()
    fecha = Format(Now, "dd/mm/yyyy")
    fechai = Format(Now, "dd/mm/yyyy")
    fechae = Format(Now, "dd/mm/yyyy")
    moneda = "S"
    turno = ""
    glosa = ""
    observa = ""
    subtablapro = ""
    ordentrabajo = ""
    serie = ""
    Numero = ""
    local1 = ""

End Sub

Sub pone_registro()
    local1 = Trim("" & txformxu.Fields("local"))
    serie = Trim("" & txformxu.Fields("serie"))
    Numero = Trim("" & txformxu.Fields("numero"))
    ordentrabajo = Trim("" & txformxu.Fields("ordentrabajo"))
    fecha = Trim("" & txformxu.Fields("fecha"))
    fechai = Trim("" & txformxu.Fields("fechai"))
    fechae = Trim("" & txformxu.Fields("fechae"))
    moneda = Trim("" & txformxu.Fields("moneda"))
    turno = Trim("" & txformxu.Fields("turno"))
    glosa = Trim("" & txformxu.Fields("glosa"))
    observa = Trim("" & txformxu.Fields("observa"))
    subtablapro = Trim("" & txformxu.Fields("subtablapro"))

End Sub

Sub grabando()
    txformxu.Fields("local") = Trim(local1)
    txformxu.Fields("serie") = Trim(serie)
    txformxu.Fields("numero") = Trim(Numero)
    txformxu.Fields("fecha") = Trim(fecha)
    txformxu.Fields("fechai") = Trim(fechai)
    txformxu.Fields("fechae") = Trim(fechae)
    txformxu.Fields("moneda") = Trim(moneda)
    txformxu.Fields("turno") = Trim(turno)
    txformxu.Fields("glosa") = Trim(glosa)
    txformxu.Fields("observa") = Trim(observa)
    txformxu.Fields("subtablapro") = Trim(subtablapro)
    txformxu.Fields("tablapro") = Trim(tablapro)

End Sub

Private Sub grba1_Click()

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
        txformxu.AddNew
        grabando
        txformxu.Fields("ESTADO") = "PLANIFICACION"
        txformxu.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        grabando
        txformxu.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    If Not IsDate(fechai) Then
        fechai.SetFocus
        Exit Function

    End If

    If Not IsDate(fechae) Then
        fechae.SetFocus
        Exit Function

    End If

    If moneda <> "S" And moneda <> "D" Then
        moneda.SetFocus
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

Sub agregar_menus()

    Dim I As Integer

    For I = 1 To mnuArchivoArray.count - 1
        Unload mnuArchivoArray(I)
    Next
     
    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from archivo where menu='formulacion' and   estado='S'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        Agregarm "" & mytablex.Fields("descripcio"), mnuArchivoArray
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Sub Agregarm(TextoDeMenu As String, QueMenu As Object)

    Dim indice As Integer

    'MsgBox QueMenu.count
    indice = QueMenu.count

    Load QueMenu(indice)

    QueMenu(indice).Caption = TextoDeMenu
    QueMenu(indice).Visible = True

End Sub

Private Sub Grupo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1

        'consulta_grupo
    End If

End Sub

Private Sub Image1_Click()
    consulta_turno

End Sub

Private Sub Image4_Click()
    consulta_tipoformula

End Sub

Private Sub local1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_pedido

    End If

End Sub

Sub mnuarchivoarray_click(Index As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    buf = mnuArchivoArray(Index).Caption
    mytablex.Open "select * from archivo where menu='formulacion' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close

    End If

    'busca el reporte
    buf = mytablex.Fields("archivo")
    mytablex.Close
    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\" & buf, "")

End Sub

Sub consulta_tipoformula()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "subtablapro"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "1"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_turno()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "Turno"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "2"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_pedido()
    Combo2.Clear
    Combo2.AddItem "Codigo"
    Combo2.AddItem "Numero"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "3"
    Text1.SetFocus
    Command4_Click

End Sub

Private Sub paol844_Click()

    On Error GoTo cmd9011_err

    If Trim("" & txformxu.Fields("ESTADO")) <> "PRODUCCION" Then
        MsgBox "Debe Encontrarse en Estado Aprobado", 48, "Aviso"
        Exit Sub

    End If

    tpartepc.vieneorden = "" & txformxu.Fields("ordentrabajo")
    tpartepc.Show 1
    Exit Sub
cmd9011_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub rep093_Click()

    If Frame4.Visible = True Then Exit Sub
    reporte_orden_excell 0

    'reporte
End Sub

Private Sub serie_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_pedido

    End If

End Sub

Private Sub subtablapro_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_tipoformula

    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    Command4_Click

End Sub

Sub filtro()

    Dim mytablex As New ADODB.Recordset

    Dim cad      As String

    If opcion1 = "1" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Descripcio,subTablapro from subTablapro where tablapro='" & tablapro & "'"

        End If

        If Len(Text1) > 0 Then
            cad = "select Descripcio,Tablapro from subtablapro where tablapro='" & tablapro & "' and " & Combo2 & " like '" & Text1 & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 5000
        dbgrid13.columns(1).Width = 2000
        'dbgrid13.columns(2).Width = 1000
        'dbgrid13.columns(3).Width = 1000
           
    End If

    If opcion1 = "2" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Descripcio,Turno from turno "

        End If

        If Len(Text1) > 0 Then
            cad = "select Descripcio,Turno from turno where  " & Combo2 & " like '" & Text1 & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 5000
        dbgrid13.columns(1).Width = 2000
        'dbgrid13.columns(2).Width = 1000
        'dbgrid13.columns(3).Width = 1000
           
    End If
   
    If opcion1 = "3" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Local,Serie,Numero,Codigo,fecha,Total from cpedidov "

        End If

        If Len(Text1) > 0 Then
            cad = "select Local,Serie,Numero,Codigo,fecha,Total from cpedidov where  " & Combo2 & " like '" & Text1 & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        'MsgBox cad
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        'dbgrid13.columns(0).Width = 5000
        'dbgrid13.columns(1).Width = 2000
        'dbgrid13.columns(2).Width = 1000
        'dbgrid13.columns(3).Width = 1000
           
    End If
   
    If mytablex.RecordCount > 0 Then
        dbgrid13.SetFocus

    End If

    Exit Sub

End Sub

Private Sub turno_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_turno

    End If

End Sub

Sub carga_tipoformula()

    Dim mytablex As New ADODB.Recordset

    tipoformula.Clear
    tipoformula.AddItem "%"
    mytablex.Open "select * from subtablapro where tablapro='" & tablapro & "'", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        tipoformula.AddItem Trim("" & mytablex.Fields("descripcio")) & "|" & Trim("" & mytablex.Fields("subtablapro"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    tipoformula.ListIndex = 0

End Sub

Sub Reporte()

    Dim found As Integer

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    borrar_archivo FileName
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento1
    cuerpo_programa_documento1
    '------------------------------------
    Close #1
    cerrar_archivo
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Sub cabecera_documento1()

    Dim buf   As String

    Dim I     As Integer

    Dim found As Integer

    If contlin > 0 Then
        buf = Chr$(12)
        found = formateaa(buf, Len(buf), 0, 0)

    End If

    contpag = contpag + 1
    contlin = 0
    cabecera_tipico "", "", "" & "" & gusuario
    buf = "Reporte de Ordenes trabajo  "
    found = formateaa(buf, 90, 2, 0)
    
    found = formateaa("Nro", 8, 0, 0)
    found = formateaa("Fecha", 11, 0, 0)
    found = formateaa("FechaIn", 11, 0, 0)
    found = formateaa("FechaEn", 11, 0, 0)
    found = formateaa("Tablapr", 8, 0, 0)
    found = formateaa("Stablap ", 8, 0, 0)
    found = formateaa("Turno ", 6, 0, 0)
    
    found = formateaa("Lo ", 7, 0, 0)
    found = formateaa("Serie ", 7, 0, 0)
    found = formateaa("Numero ", 12, 2, 0)
    
    buf = "Producto"
    found = formateaa(buf, 7, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Descripcio"
    found = formateaa(buf, 59, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Unidad"
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Factor"
    found = formateaa(buf, 6, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Cantidad"
    found = formateaa(buf, 7, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Formula"
    found = formateaa(buf, 7, 0, 0)
    found = formateaa("", 1, 0, 0)
    buf = "Almacen"
    found = formateaa(buf, 7, 0, 0)
    found = formateaa("", 1, 0, 0)
      
    buf = "Stock"
    found = formateaa(buf, 7, 0, 0)
    found = formateaa("", 1, 2, 0)
        
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)

End Sub

Sub cuerpo_programa_documento1()

    Dim buf   As String

    Dim found As Integer

    On Error GoTo cmd78812_err

    Do

        If txformxu.EOF Then Exit Do
        buf = "+" & txformxu.Fields("ordentrabajo")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txformxu.Fields("fecha")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txformxu.Fields("fechai")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txformxu.Fields("fechae")
        found = formateaa(buf, 10, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txformxu.Fields("tablapro")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txformxu.Fields("subtablapro")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txformxu.Fields("turno")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txformxu.Fields("Local")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txformxu.Fields("Serie")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txformxu.Fields("numero")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas
        imprime_detalle "" & txformxu.Fields("ordentrabajo")
        txformxu.MoveNext
    Loop
    Exit Sub
cmd78812_err:
    MsgBox "Aviso en cuerpo " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > 45 Then
        cabecera_documento1

    End If

End Sub

Sub imprime_detalle(buf1 As String)

    Dim buf   As String

    Dim found As Integer

    On Error GoTo cmd903_err

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from ordentrabajod where ordentrabajo=" & buf1, cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        buf = "-" & mytablex.Fields("Producto")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("descripcio")
        found = formateaa(buf, 59, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("unidad")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("factor")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("cantidad")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Formula")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("Bodega")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas
        imprime_receta "" & mytablex.Fields("formula")
        mytablex.MoveNext
    Loop
    mytablex.Close
    Exit Sub
cmd903_err:
    MsgBox "Aviso en imprime Detalle " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub imprime_receta(buf1 As String)

    Dim buf   As String

    Dim found As Integer

    On Error GoTo cmd32_err

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from componente where id=" & Val(buf1), cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        buf = "*" & mytablex.Fields("Producto")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("descripcio")
        found = formateaa(buf, 59, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("unidad")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("factor")
        found = formateaa(buf, 6, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & mytablex.Fields("cantidad")
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = ""
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = ""
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & stock_actual("" & mytablex.Fields("Producto"))
        found = formateaa(buf, 7, 0, 0)
        found = formateaa("", 1, 2, 0)
        nlineas
        mytablex.MoveNext
    Loop
    mytablex.Close
    Exit Sub
cmd32_err:
    MsgBox "Aviso en imprime receta " + error$, 48, "Aviso"
    Exit Sub

End Sub

Function stock_actual(buf1 As String) As Double

    Dim sdx      As Double

    Dim mytablex As New ADODB.Recordset

    sdx = 0
    mytablex.Open "select * from almacen where producto='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + Val("" & mytablex.Fields("saldo"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    stock_actual = Val(Format(sdx, "0.00"))

End Function

Sub reporte_orden_excell(sw As Integer)

    Dim found       As Integer

    Dim buf         As String

    Dim mytablex    As New ADODB.Recordset

    Dim mytabley    As New ADODB.Recordset

    Dim Heading(15) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    Dim v           As Long

    Dim h           As Long

    Command1.Visible = True

    On Error GoTo cmd6561245_err

    Heading(1) = "OrdenNro"
    Heading(2) = "Fecha"
    Heading(3) = "Tipo"
    Heading(4) = "FechaEntrega"
    Heading(5) = "Estado"
    Heading(6) = "Aprobado"
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_orden(15, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook

    objExcel.ActiveSheet.Cells(1, 1) = "ORDENES DE TRABAJO  "

    v = 4
    h = 1
    
    Do

        If txformxu.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, h) = "'" & txformxu.Fields("ordentrabajo")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & txformxu.Fields("fecha")
        objExcel.ActiveSheet.Cells(v, h + 2) = busca_subgrupo()
        objExcel.ActiveSheet.Cells(v, h + 3) = "'" & txformxu.Fields("fechae")
        objExcel.ActiveSheet.Cells(v, h + 4) = "'" & txformxu.Fields("estado")
        objExcel.ActiveSheet.Cells(v, h + 5) = "'" & txformxu.Fields("aprobado")
        v = v + 1
        '--------------------------

        objExcel.ActiveSheet.Cells(v, h + 1) = "Descripcio"
        objExcel.ActiveSheet.Cells(v, h + 2) = "Unidad"
        objExcel.ActiveSheet.Cells(v, h + 3) = "factor"
        objExcel.ActiveSheet.Cells(v, h + 4) = "PorFabricar"
        objExcel.ActiveSheet.Cells(v, h + 5) = "Avance"

        v = v + 1
        mytablex.Open "select * from ordentrabajod where ordentrabajo=" & Val("" & txformxu.Fields("ordentrabajo")), cn, adOpenStatic, adLockOptimistic
        Do

            If mytablex.EOF Then Exit Do
            objExcel.ActiveSheet.Cells(v, h) = "Formula"
            objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablex.Fields("descripcio")
            objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablex.Fields("unidad")
            objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytablex.Fields("factor")
            objExcel.ActiveSheet.Cells(v, h + 4) = "" & mytablex.Fields("cantidad")
            objExcel.ActiveSheet.Cells(v, h + 5) = avance_produccion()
            v = v + 1

            '--------------------------------------------------------
            If sw = 1 Then   'insumos
                '---------------------------------------

                mytabley.Open "select * from componente where id=" & Val("" & mytablex.Fields("formula")), cn, adOpenStatic, adLockOptimistic

                If mytabley.RecordCount > 0 Then
                    objExcel.ActiveSheet.Cells(v, h) = "Insumos"
                    objExcel.ActiveSheet.Cells(v, h + 1) = "Descripcio"
                    objExcel.ActiveSheet.Cells(v, h + 2) = "Unidad"
                    objExcel.ActiveSheet.Cells(v, h + 3) = "factor"
                    objExcel.ActiveSheet.Cells(v, h + 4) = "Cantidad"
                    objExcel.ActiveSheet.Cells(v, h + 5) = "StockActual"
                    v = v + 1

                End If

                Do

                    If mytabley.EOF Then Exit Do

                    objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytabley.Fields("descripcio")
                    objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytabley.Fields("unidad")
                    objExcel.ActiveSheet.Cells(v, h + 3) = "'" & mytabley.Fields("factor")
                    objExcel.ActiveSheet.Cells(v, h + 4) = (Val("" & mytabley.Fields("cantidad")) / Val("" & mytabley.Fields("factor"))) * Val("" & mytablex.Fields("cantidad"))
                    objExcel.ActiveSheet.Cells(v, h + 5) = "" & stock_actual("" & mytabley.Fields("Producto"))
                    v = v + 1
                    mytabley.MoveNext
                Loop
                mytabley.Close

                '----------------------------------------
            End If

            '--------------------------------------------------------

            mytablex.MoveNext
        Loop
        mytablex.Close
        v = v + 1
        '--------------------------
        txformxu.MoveNext
    Loop
    Set objExcel = Nothing
    Exit Sub
cmd6561245_err:
    MsgBox "Aviso en reporte orden ", 48, "Aviso"
    Exit Sub

End Sub

Function Formato_orden(Num_Campos As Integer, Nombre_Campos() As String) As Boolean

    Dim I As Integer

    With objExcel.ActiveSheet
        'MsgBox Num_Campos
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, Num_Campos)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 8)).Font.bold = True
        
        For I = 1 To Num_Campos Step 1
            .Cells(3, I) = Nombre_Campos(I)
        Next I

        'hasta aki pa colocar los titulos
        
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
       
        .columns("A").ColumnWidth = 10
        .columns("B").ColumnWidth = 40
        .columns("C").ColumnWidth = 20
        .columns("D").ColumnWidth = 20
        .columns("E").ColumnWidth = 20
        .columns("F").ColumnWidth = 10
        .columns("G").ColumnWidth = 10
        .columns("H").ColumnWidth = 10
    
    End With

End Function

Function busca_subgrupo() As String

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    buf = "select * from subTablapro where tablapro='" & "" & txformxu.Fields("tablapro") & "' and subtablapro='" & "" & txformxu.Fields("subtablapro") & "'"
  
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_subgrupo = "" & mytablex.Fields("descripcio")

    End If

    mytablex.Close

End Function

Function avance_produccion() As Double

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    Dim sdx      As Double

    sdx = 0
    buf = "select * from parteproducciond where ordentrabajo=" & Val("" & txformxu.Fields("ordentrabajo"))
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + Val("" & mytablex.Fields("cantidad"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    avance_produccion = sdx

End Function

