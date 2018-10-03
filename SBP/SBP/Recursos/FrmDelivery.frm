VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmDelivery 
   Caption         =   "VENTAS POR DELIVERY"
   ClientHeight    =   10380
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   10380
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   14520
      Top             =   120
   End
   Begin VB.Frame frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10935
      Left            =   2280
      TabIndex        =   25
      Top             =   5040
      Width           =   15105
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn9 
         Height          =   765
         Left            =   4800
         TabIndex        =   26
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1349
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
         BCOL            =   255
         BCOLO           =   255
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDelivery.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataGridLib.DataGrid DataGrid4 
         Height          =   4335
         Left            =   9480
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   7646
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16744576
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
      Begin MSDataGridLib.DataGrid table7 
         Height          =   8055
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   14208
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         HeadLines       =   3
         RowHeight       =   26
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
               ColumnWidth     =   2880
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1094.74
            EndProperty
         EndProperty
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn6 
         Height          =   735
         Left            =   11400
         TabIndex        =   34
         Top             =   8280
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1296
         BTYPE           =   5
         TX              =   "FINALIZAR"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDelivery.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   7095
         Left            =   6000
         TabIndex        =   38
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   12515
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16744576
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total a Recibir:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   11160
         TabIndex        =   37
         Top             =   7560
         Width           =   1695
      End
      Begin VB.Label rtxentrega 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Height          =   660
         Left            =   13080
         TabIndex        =   36
         Top             =   7440
         Width           =   1890
      End
      Begin VB.Label rtxtotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Height          =   660
         Left            =   9120
         TabIndex        =   33
         Top             =   7440
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total a Entregar:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7080
         TabIndex        =   32
         Top             =   7560
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   10200
         Width           =   14505
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   15060
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Width           =   15120
      Begin ChamaleonButton.ChameleonBtn Asignar 
         Height          =   705
         Left            =   240
         TabIndex        =   22
         Top             =   120
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1244
         BTYPE           =   4
         TX              =   "ASIGNAR VENDEDOR"
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
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDelivery.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn entregar 
         Height          =   705
         Left            =   1800
         TabIndex        =   23
         Top             =   120
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1244
         BTYPE           =   4
         TX              =   "ENTREGA DE DELIVERY"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDelivery.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn7 
         Height          =   705
         Left            =   4920
         TabIndex        =   24
         Top             =   120
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1244
         BTYPE           =   4
         TX              =   "SALIR"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDelivery.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn8 
         Height          =   705
         Left            =   3360
         TabIndex        =   35
         Top             =   120
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1244
         BTYPE           =   4
         TX              =   "REPORTE"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDelivery.frx":008C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label horasistema 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   13380
         TabIndex        =   31
         Top             =   360
         Width           =   60
      End
      Begin VB.Label diasistema 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11760
         TabIndex        =   30
         Top             =   360
         Width           =   60
      End
      Begin VB.Label znumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   11040
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label zserie 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   10200
         TabIndex        =   20
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label ztipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   9480
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
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
      Height          =   10935
      Left            =   -120
      TabIndex        =   3
      Top             =   720
      Width           =   15105
      Begin VB.TextBox buffer 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   360
         Width           =   2880
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   360
         Width           =   2670
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00808080&
         Caption         =   "Personal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   8535
         Left            =   13680
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   6135
         Begin MSDataGridLib.DataGrid table6 
            Height          =   8055
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   4440
            _ExtentX        =   7832
            _ExtentY        =   14208
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16744576
            HeadLines       =   3
            RowHeight       =   26
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
                  ColumnWidth     =   2880
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1200.189
               EndProperty
            EndProperty
         End
         Begin ChamaleonButton.ChameleonBtn ChameleonBtn3 
            Height          =   1005
            Left            =   4680
            TabIndex        =   6
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   1773
            BTYPE           =   5
            TX              =   "Asignar Personal"
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
            MICON           =   "FrmDelivery.frx":00A8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ChamaleonButton.ChameleonBtn ChameleonBtn5 
            Height          =   975
            Left            =   4680
            TabIndex        =   7
            Top             =   1800
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   1720
            BTYPE           =   5
            TX              =   "Cerrar"
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
            FCOL            =   16777215
            FCOLO           =   16777215
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmDelivery.frx":00C4
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
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   8655
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   15266
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
      Begin ChamaleonButton.ChameleonBtn Label26 
         Height          =   765
         Left            =   7800
         TabIndex        =   11
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1349
         BTYPE           =   5
         TX              =   ">>>"
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
         BCOL            =   255
         BCOLO           =   255
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDelivery.frx":00E0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn Command1 
         Height          =   495
         Left            =   5880
         TabIndex        =   12
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
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
         MICON           =   "FrmDelivery.frx":00FC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn1 
         Height          =   765
         Left            =   7800
         TabIndex        =   13
         Top             =   2040
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1349
         BTYPE           =   5
         TX              =   "<<<"
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
         BCOL            =   255
         BCOLO           =   255
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDelivery.frx":0118
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn2 
         Height          =   735
         Left            =   8880
         TabIndex        =   14
         Top             =   8880
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1296
         BTYPE           =   5
         TX              =   "Buscar Personal"
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDelivery.frx":0134
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   7695
         Left            =   8880
         TabIndex        =   15
         Top             =   1080
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   13573
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16744576
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
      Begin ChamaleonButton.ChameleonBtn ChameleonBtn4 
         Height          =   735
         Left            =   12600
         TabIndex        =   16
         Top             =   8880
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1296
         BTYPE           =   5
         TX              =   "ASIGNAR / IMPRIMIR"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmDelivery.frx":0150
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label label56 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   10200
         Width           =   14505
      End
   End
   Begin VB.Label caja 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1380
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label cajero 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label turno 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1770
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "FrmDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rcconsulta        As New ADODB.Recordset

Dim rcconsultax       As New ADODB.Recordset

Dim rcconsultae       As New ADODB.Recordset

Dim rcconsultageneral As New ADODB.Recordset

Dim dbvarios          As New ADODB.Recordset

Dim dbvariose         As New ADODB.Recordset

Dim stx               As Double

Dim xtotal            As Double

Private Sub Asignar_Click()
 
    Asignar.BackColor = &H8080FF
    Asignar.BackOver = &H8080FF
 
    entregar.BackColor = &HC0C0C0
    entregar.BackOver = &HC0C0C0
 
    Frame1.Top = 1080
    Frame1.Left = 0
    Frame1.Visible = True
 
    Frame2.Visible = False
 
    buffer = ""
 
    'Ventas por delivery
    opcion1 = "15"
    sw_consulta = 0
    found = sql_consulta(1)

    opcion1 = "16"
    sw_consulta = 0
    found = sql_consultaX(1)

End Sub

Private Sub ChameleonBtn1_Click()

    Dim found As Integer

    On Error GoTo cmd8966_err

    If rcconsultax.RecordCount = 0 Then Exit Sub

    ir_hasta_primerox rcconsultax

    If "" & rcconsultax.Fields("X") = "S" Then
        rcconsultax.Fields("X") = ""
        rcconsultax.Fields("personal") = ""
        rcconsultax.Fields("HoraSal") = ""
        rcconsultax.Update
   
        opcion1 = "15"
        found = sql_consulta(1)

        opcion1 = "16"
        found = sql_consultaX(1)
   
        Exit Sub

    End If

    rcconsultax.Fields("X") = "S"
    rcconsultax.Update
   
    opcion1 = "15"
    found = sql_consulta(1)
   
    opcion1 = "16"
    found = sql_consultaX(1)

    Exit Sub
cmd8966_err:
    MsgBox "Seleccione un Dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub ChameleonBtn2_Click()
 
    If rcconsultax.RecordCount = 0 Then Exit Sub
 
    If Frame8.Visible = True Then
        Frame8.Visible = False
    Else
        Frame8.Visible = True
 
        If dbvarios.State = 1 Then dbvarios.Close
        dbvarios.Open "select Nombre,Codigo from vendedor   where  CARGO='MOTORIZADO' AND local='" & "" & mytable11.Fields("local") & "' order by nombre ", cn, adOpenStatic, adLockOptimistic

        If dbvarios.RecordCount = 0 Then
            MsgBox "No existe Vendedor asignado al local ", 48, "Aviso"
            dbvarios.Close
            Exit Sub

        End If

        Set table6.DataSource = dbvarios
        Frame8.Caption = "Personal"
        Frame8.Visible = True
        table6.SetFocus
        Exit Sub

    End If
  
End Sub

Private Sub ChameleonBtn3_Click()

    If rcconsultax.RecordCount = 0 Then Exit Sub
    rcconsultax.MoveFirst
    Do

        If rcconsultax.EOF Then Exit Do
        rcconsultax.Fields("personal") = Trim("" & dbvarios.Fields("codigo"))
        rcconsultax.Fields("horasal") = Format(Now, "HH:MM")
        rcconsultax.Update
        rcconsultax.MoveNext
    Loop

    Frame8.Visible = False
    opcion1 = "16"
    found = sql_consultaX(1)

End Sub

Function formateab(buf As String, _
                   longitud As Integer, _
                   sw As Integer, _
                   sw1 As Integer) As String

    Dim xbuf As String

    Dim buf1 As String

    Dim sdx  As Integer

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

Function imprime_adifac(batipo As String, _
                        baserie As String, _
                        banumero As String, _
                        sw As Integer, _
                        xxpuerto As String)

    Dim mytablex  As New ADODB.Recordset

    Dim found     As Integer

    Dim buf       As String

    Dim X         As Double

    Dim sFile     As String

    Dim cfilename As String

    On Error GoTo cmd67112_err

    Dim xmcanal

    Exit Function
    '---------------------------------
    mytablex.Open "SELECT * FROM factura where  local='" & "" & mytable11.Fields("local") & "' and tipo='" & batipo & "' and serie='" & baserie & "' and numero='" & banumero & "'", cn, adOpenKeyset, adLockOptimistic

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
        found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))

    End If

    If sw = 1 Then  'impresion directa
        FileName = sFile
        found = star_sp342(xxpuerto, 0)

        ' found = corte_papel(xxpuerto, 0)
    End If

    Exit Function
cmd67112_err:
    MsgBox "Aviso en imprime adicional " + error$, 48, "Aviso"
    Close #xmcanal
    Exit Function

End Function

Function imprime_adicional(batipo As String, _
                           baserie As String, _
                           banumero As String, _
                           sw As Integer, _
                           xxpuerto As String)

    Dim mytablex      As New ADODB.Recordset

    Dim ax1cambio     As String

    Dim ax1telefono   As String

    Dim ax1nombre     As String

    Dim ax1direccio   As String

    Dim ax1referencia As String

    Dim ax1pago       As String

    Dim ax1total      As String

    Dim ax1vuelto     As String

    Dim found         As Integer

    Dim cfilename     As String

    Dim sFile         As String

    Dim I             As Integer

    Dim buf           As String

    Dim ax1codigo     As String

    On Error GoTo cmd6711_err

    Dim xmcanal

    ax1codigo = ""
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
    mytablex.Open "SELECT * FROM deliveri where  codigo='" & Trim("" & rcconsultax.Fields("codigo")) & "'", cn, adOpenKeyset, adLockOptimistic

    If mytablex.RecordCount = 0 Then  'si existe
        mytablex.Close
        Exit Function

    End If

    ax1telefono = "" & mytablex.Fields("telefono")
    ax1nombre = "" & mytablex.Fields("nombre")
      
    ax1direccio = "" & mytablex.Fields("direccion")
    ax1referencia = "" & mytablex.Fields("referencia")
    ax1codigo = "" & mytablex.Fields("codigo")
    mytablex.Close

    xmcanal = FreeFile
    Open globaldir & "\temporal\" & gusuario & "TX" For Output As #xmcanal
    Print #xmcanal, "              DELIVERY              "
    Print #xmcanal, "===================================="
    Print #xmcanal, "Fecha:" + Format(Now, "dd/mm/yyyy")
    Print #xmcanal, "Hora :" + Format(Now, "HH:MM:SS")
    Print #xmcanal, "Telef:" + ax1telefono
    Print #xmcanal, "Clien:" + ax1nombre
    Print #xmcanal, "Direc:" + ax1direccio
    Print #xmcanal, "Refer:" + ax1referencia
    Print #xmcanal, "------------------------------------"
    Print #xmcanal, "T/C  :" + ax1cambio
    buf = imprime_tipodoc("" & batipo)
    Print #xmcanal, "T.Doc:" + buf
    Print #xmcanal, "Numer:" + baserie + " " + banumero
    buf = imprime_clasifica_cliente(ax1codigo)
    found = formateaa(buf, 30, 2, 0)

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
        Print #xmcanal, "------------------------------------"
        mytablex.MoveNext
    Loop
   
    mytablex.Close

    'tipo de documento
    mytablex.Open "select * from " & godetalle & " where local='" & "" & mytable11.Fields("local") & "' and tipo='" & "" & batipo & "' and serie='" & "" & baserie & "' and numero='" & "" & banumero & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        buf = "" & mytablex.Fields("tipo")

    End If

    mytablex.Close
    '------------ PRODUCTOS
    mytablex.Open "select * from " & godetalle & " where local='" & "" & mytable11.Fields("local") & "' and tipo='" & "" & batipo & "' and serie='" & "" & baserie & "' and numero='" & "" & banumero & "' ", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            If "" & mytablex.Fields("dua") <> "R" Then
                buf = "" & mytablex.Fields("cantidad")
                found = formateaa(buf, 6, 0, 0)
                found = formateaa("", 1, 0, 0)
                buf = Mid$("" & mytablex.Fields("descripcio"), 1, 22)
                found = formateaa(buf, 22, 2, 0)

                If Len(Mid$("" & mytablex.Fields("descripcio"), 23, 22)) > 0 Then
                    buf = Mid$("" & mytablex.Fields("descripcio"), 23, 22)
                    found = formateaa(buf, 22, 2, 0)

                End If

                If Len(Mid$("" & mytablex.Fields("descripcio"), 45, 22)) > 0 Then
                    buf = Mid$("" & mytablex.Fields("descripcio"), 45, 22)
                    found = formateaa(buf, 22, 2, 0)

                End If

                If Len(Mid$("" & mytablex.Fields("descripcio"), 68, 22)) > 0 Then
                    buf = Mid$("" & mytablex.Fields("descripcio"), 68, 22)
                    found = formateaa(buf, 22, 2, 0)

                End If

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
            End If

            mytablex.MoveNext
        Loop

    End If

    mytablex.Close
       
    For I = 1 To 7
        Print #xmcanal, ""
    Next I

    Close #xmcanal
    sFile = globaldir & "\temporal\" & gusuario & "TX"

    If sw = 0 Then
        found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))

    End If

    If sw = 1 Then
        FileName = sFile
        found = star_sp342(xxpuerto, 0)

        'found = corte_papel(xxpuerto, Val("" & mytable11.Fields("catipo")))
    End If

    Exit Function
cmd6711_err:
    MsgBox "Aviso en imprime adicional " + error$, 48, "Aviso"
    Close #xmcanal
    Exit Function

End Function

Function control_impresion(bxtipo As String, _
                           bxserie As String, _
                           bxnumero As String, _
                           psw As Integer)

    Dim found      As Integer

    Dim sFile      As String

    Dim mytablex   As New ADODB.Recordset

    Dim sw         As String

    Dim xcolax     As String

    Dim xxpuerto   As String

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
                sw = "" & mytable11.Fields("ibm")
                xcolax = "" & mytable11.Fields("cbm")
       
            Case "B"
                sw = "" & mytable11.Fields("ifm")
                xcolax = "" & mytable11.Fields("cfm")
       
            Case "C"
                sw = "" & mytable11.Fields("itb")
                xcolax = "" & mytable11.Fields("ctb")
       
            Case "D"
                sw = "" & mytable11.Fields("itf")
                xcolax = "" & mytable11.Fields("ctf")
       
            Case "G"
                sw = "" & mytable11.Fields("inv")
                xcolax = "" & mytable11.Fields("cnv")
       
            Case "H"
                sw = "" & mytable11.Fields("ipe")
                xcolax = "" & mytable11.Fields("cpe")
       
            Case "I"
                sw = "" & mytable11.Fields("ipe")
                xcolax = "" & mytable11.Fields("cpro")
       
            Case "1"
                sw = "" & mytable11.Fields("iexo")
                xcolax = "" & mytable11.Fields("cexo")
       
            Case "E"
                sw = "" & mytable11.Fields("iNC")
                xcolax = "" & mytable11.Fields("cNC")
        
        End Select

    End If

    mytablex.Close

    '''23/10/2017 Mejora Delivery
    Dim mytablexyz As New ADODB.Recordset

    If mytablexyz.State = 1 Then mytablexyz.Close
    mytablexyz.Open "SELECT puertodelivery,coladelivery FROM parameca where  caja='" & caja & "'", cn, adOpenDynamic, adLockOptimistic

    If mytablexyz.RecordCount > 0 Then
        xxpuerto = mytable11.Fields("puertodelivery")
        xcolax = mytable11.Fields("coladelivery")
            
    End If

    mytablexyz.Close
    '''23/10/2017 Mejora Delivery

    If xcolax = "N" Or xcolax = " " Then
        Exit Function

    End If

    If psw = 10 Then  'solo es para ver si es LPT
        control_impresion = 11

        If xxpuerto = "LPT" Then
            control_impresion = 10

        End If

        Exit Function

    End If

    If psw = 2 Then  'si  es orden de despacho
        If xcolax = "S" Then
            sFile = globaldir & "\temporal\" & gusuario & ".txt"
            found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))

        End If

        If xcolax <> "S" Then
            found = star_sp342(xxpuerto, 0)

        End If

        control_impresion = found
        Exit Function

    End If

    If flag_servicio = "D" Then
        If impresiondelivery = "N" Then
            control_impresion = 1
            Exit Function

        End If

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
        found = Imprime_archivojj(sFile, 0, "" & mytable11.Fields("tamanoletra"), "" & mytable11.Fields("nombrefont"), "" & mytable11.Fields("BOLD"), "" & mytable11.Fields("letrainterna"))

        If bxtipo = "2" Then
            found = imprime_adifac(bxtipo, bxserie, bxnumero, 0, "")

        End If
               
        selecciona_impresoras (oldprinter)

    End If

    If xcolax <> "S" Then
        found = star_sp342(xxpuerto, 0)

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
    MsgBox "Aviso en control impresion " + error$, 48, "Aviso"
    Exit Function

End Function

Sub proceso_impresion11(bxtipo As String, _
                        bxserie As String, _
                        bxnumero As String, _
                        sw As Integer, _
                        ascopia As String)

    ''IMPRESION- IMPRESORA
    Dim found    As Integer

    Dim archivot As String

    On Error GoTo cmd6_err:

    cerrar_archivo
    
    If Trim("" & mytable11.Fields("coladelivery")) <> "S" Then
        Exit Sub

    End If
    
    If sw = 0 Then
        If Trim("" & mytable11.Fields("gavetasw")) <> "N" Then
            found = abre_puerto("" & mytable11.Fields("capuerto"), Val("" & mytable11.Fields("catipo")), "" & mytable11.Fields("gavetacola"))

        End If

    End If
    
    found = control_impresion(bxtipo, bxserie, bxnumero, 10)

    If found = 10 Then
        Exit Sub

    End If
    
    factura_formatox "01", "" & bxtipo, "" & bxserie, "" & bxnumero, ascopia, sw
    cerrar_archivo
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    found = control_impresion(bxtipo, bxserie, bxnumero, sw)

    Exit Sub
cmd6_err:
    MsgBox "Mensaje, Error al iniciar Impresion " & error$, 48, "Aviso"
    Exit Sub

End Sub

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

            Case "E" 'NC
                busca_archivo_formato = "" & mytable11.Fields("archivonc")

        End Select

    End If

    mytablex.Close
 
End Function

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
    'IMPRESION - IMPRESORA
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
    'inicio 30/05/2017 pll posiblemente abajo aqui hace la estructura
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
                    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                    'found = proceso_formatos(archivo_formato, mytabley, "/", "\", godetalle, "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
                    found = proceso_formatos(archivo_formato, mytabley, "/", "\", godetalle, "TDETALLE", bxlocal, bxtipo, bxserie, bxnumero, ascopia, contando)
                    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
                   
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

            For I = contando To nro_lineas
                Open FileName For Append As #1
                found = formateaa("", 1, 2, 0)
                Close #1
            Next I

        End If

    End If

    '----- SUBTOTAL
       
    'inicio 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
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
    found = proceso_formatos(archivo_formato, mytablex, "^", "&", gocabeza, "TFACTURA", bxlocal, bxtipo, bxserie, bxnumero, ascopia, 0)
    'fin 30/05/2017 pll para la parametrizacion nombre consistencvia cliente ticketera
        
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

Private Sub ChameleonBtn4_Click()

    Dim tipo   As String

    Dim serie  As String

    Dim Numero As String

    flag_servicio = "D"

    If rcconsultax.RecordCount = 0 Then
        Exit Sub

    End If

    If rcconsultax.Fields("PERSONAL") = "" Then
        MsgBox ("Asigne Vendedor"), vbCritical
        ChameleonBtn2_Click
        Exit Sub

    End If
    
    If rcconsultax.RecordCount = 0 Then Exit Sub
    rcconsultax.MoveFirst

    For I = 0 To rcconsultax.RecordCount - 1
        proceso_impresion11 rcconsultax.Fields("tipo"), rcconsultax.Fields("serie"), rcconsultax.Fields("numero"), 1, "1"
        rcconsultax.Fields("RENUMERO1") = "S"
        rcconsultax.Update
        rcconsultax.MoveNext
    Next
 
    tipo = ""
    serie = ""
    Numero = ""

    opcion1 = "16"
    found = sql_consultaX(1)

End Sub

Private Sub ChameleonBtn5_Click()
    Frame8.Visible = False

End Sub

Private Sub ChameleonBtn6_Click()

    If rcconsultae.RecordCount = 0 Then
        Exit Sub

    End If
 
    If MsgBox("Desea FINALIZAR los pedidos de delivery", 1, "Aviso") <> 1 Then Exit Sub
 
    If rcconsultae.RecordCount = 0 Then Exit Sub
    rcconsultae.MoveFirst

    For I = 0 To rcconsultae.RecordCount - 1
        rcconsultae.Fields("RENUMERO1") = "E"
        rcconsultae.Fields("E") = "ENTREGADO"
        rcconsultae.Fields("RENUMERO2") = Format(Now, "HH:MM")
        rcconsultae.Update
        rcconsultae.MoveNext
    Next
 
    rtxtotal = Format(0, "0.00")
    rtxentrega = Format(0, "0.00")

    opcion1 = "17"
    found = sql_consultae(1)

    found2 = sql_consultageneral(1)

End Sub

Private Sub ChameleonBtn7_Click()
    FrmDelivery.Hide
    Unload FrmDelivery

End Sub

Sub BuscarDeliveryXVendedor()

    Dim found As Integer

    opcion1 = "17"
    found = sql_consultae(1)

    Dim found2 As Integer

    found = sql_consultageneral(1)

    rtxtotal = Format(0, "0.00")
    rtxentrega = Format(0, "0.00")

    For I = 0 To rcconsultageneral.RecordCount - 1
        ' rtxtotal = rtxtotal + rcconsultae.Fields("total")
        ' rtxtotal = Format(rtxtotal, "0.00")
 
        rtxentrega = rtxentrega + rcconsultageneral.Fields("recibe")
        rtxentrega = Format(rtxentrega, "0.00")
        rcconsultageneral.MoveNext
    Next

End Sub

Private Sub ChameleonBtn8_Click()
    FrmReporteDelivery.Show 1

End Sub

Private Sub ChameleonBtn9_Click()
    BuscarDeliveryXVendedor

End Sub

Private Sub Command1_Click()

    Dim found As Integer

    opcion1 = "15"
    found = sql_consulta(1)

End Sub

Private Sub Command2_Click()
    rcconsultax.MoveLast

End Sub

Private Sub entregar_Click()
    Asignar.BackColor = &HC0C0C0
    Asignar.BackOver = &HC0C0C0
 
    entregar.BackColor = &H8080FF
    entregar.BackOver = &H8080FF
  
    Frame2.Top = 1080
    Frame2.Left = 0
    Frame2.Visible = True
    Frame1.Visible = False
 
    If dbvariose.State = 1 Then dbvariose.Close
    dbvariose.Open "select Nombre,Codigo from vendedor  WHERE CARGO='MOTORIZADO' AND local='" & "" & mytable11.Fields("local") & "' order by nombre ", cn, adOpenStatic, adLockOptimistic

    If dbvariose.RecordCount = 0 Then
        MsgBox "No existe Vendedor asignado al local ", 48, "Aviso"
        dbvariose.Close
        Exit Sub

    End If

    Set table7.DataSource = dbvariose
    table7.SetFocus

    BuscarDeliveryXVendedor
    Exit Sub
 
End Sub

Private Sub fechasis_Click()

End Sub

Private Sub Form_Activate()
    Frame8.Top = 1080
    Frame8.Left = 8880

    Frame1.Top = 1080
    Frame1.Left = 0
 
    Frame1.Visible = True
    Frame2.Visible = False
 
End Sub

Private Sub Form_Load()

    Dim found As Integer

    Dim buf   As String

    flag_servicio = "D"
    cajero = tptovta.cajero
    caja = tptovta.caja
    turno = tptovta.turno

    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0

    Frame1.Visible = True
    Frame1.Enabled = True

    buffer = ""
    opcion1 = "15"
    sw_consulta = 0
    found = sql_consulta(1)

    opcion1 = "16"
    sw_consulta = 0
    found = sql_consultaX(1)

End Sub

Function sql_consultaX(sw As Integer)

    Dim buf       As String

    Dim queprecio As String

    Dim indx      As Integer

    Dim dbf1      As String

    Dim dbf2      As String

    Dim amfecha   As String

    'On Error GoTo cmd8912_err
    'MsgBox buffer
    amfecha = Format(dia, "YYYYMMDD")
    indx = -1
    dbf1 = ""
    dbf2 = ""
              
    If opcion1 = "16" Then

        buf = "select FLAG_DELI as 'X',f.tipo,f.serie,f.Numero,f.Fecha,f.Nombre as Cliente,f.Codigo,f.Moneda as M,f.Total,f.Estado as E,f.Servicio as S,f.Placa as 'E',f.Vendedor AS Personal,f.Hora,f.Caja,f.Turno,f.Local,f.Tipo1,f.Serie1,f.Numero1,horae as HoraSal,RENUMERO1  from factura f  where f.local='01' and "
        buf = buf & "  f.fecha='" & amfecha & "'"
        buf = buf & " and estado='2' AND f.FLAG_DELI ='S' AND (F.RENUMERO1 IS NULL OR F.RENUMERO1='') and f.servicio='D' and f.usuario='" & cajero & "'"
        ''''23/09/2017 kenyo Multicaja
        'buf = buf & " and f.caja='" & caja & "'"
        ''''23/09/2017 kenyo Multicaja
   
        buf = buf & " and f.turno='" & turno & "'"
        buf = buf & " order by f.HORA"

    End If
              
    Set rcconsultax = Nothing

    If rcconsultax.State = 1 Then
        rcconsultax.Close
        Set rcconsultax = Nothing

    End If
   
    'MsgBox buf & " " & opcion1
    rcconsultax.Open buf, cn, adOpenStatic, adLockOptimistic
    Set DataGrid2.DataSource = rcconsultax
    DataGrid2.refresh
    
    'sw_consulta = 1
         
    If opcion1 = "16" Then
        DataGrid2.columns(0).Width = 1
        DataGrid2.columns(1).Width = 1 'TIPO
        DataGrid2.columns(2).Width = 430 'SERIE
        DataGrid2.columns(3).Width = 800 'NUMERO
        DataGrid2.columns(4).Width = 1 ' fecha
        DataGrid2.columns(5).Width = 2000 'CLIENTE
        DataGrid2.columns(6).Width = 1100
        DataGrid2.columns(7).Width = 1 'moneda
        DataGrid2.columns(8).Width = 1 'total
        DataGrid2.columns(9).Width = 1 'estado
        DataGrid2.columns(10).Width = 1 'servicio
        DataGrid2.columns(11).Width = 1 'estado
        DataGrid2.columns(12).Width = 780 'vendedor
        DataGrid2.columns(13).Width = 1
        DataGrid2.columns(14).Width = 1
        DataGrid2.columns(15).Width = 1
        DataGrid2.columns(16).Width = 1
        DataGrid2.columns(17).Width = 1
        DataGrid2.columns(18).Width = 1
        DataGrid2.columns(19).Width = 1
        DataGrid2.columns(20).Width = 670
        DataGrid2.columns(21).Width = 1

    End If
             
    If rcconsultax.RecordCount = 0 Then
        Exit Function

    End If
              
    If opcion1 = "16" Then
        If rcconsultax.RecordCount > 0 Then
            ir_hasta_primerox rcconsultax

        End If

    End If
              
    sql_consultaX = 1
    Exit Function
    'cmd8912_err:
    'MsgBox "Aviso en sql_consulta " & error$, 48, "Aviso"
    buffer = ""
    Exit Function

End Function

Function sql_consulta(sw As Integer)

    Dim buf       As String

    Dim queprecio As String

    Dim indx      As Integer

    Dim dbf1      As String

    Dim dbf2      As String

    Dim amfecha   As String

    'On Error GoTo cmd8912_err
    'MsgBox buffer
    amfecha = Format(dia, "YYYYMMDD")
    indx = -1
    dbf1 = ""
    dbf2 = ""

    If opcion1 = "15" Then
  
        If Len(buffer) = 0 Then
            buf = "select FLAG_DELI as 'X',tipo,serie,Numero,Fecha,Nombre as Cliente,Codigo,Moneda as M,Total,Estado as E,Servicio as S,Placa as 'E',Vendedor AS PERSONAL,Hora AS 'Hora P.',Caja,Turno,Local,Tipo1,Serie1,Numero1,RENUMERO1,horae as 'horasal' from factura where local='01' and "
            buf = buf & "  fecha='" & amfecha & "'"
            buf = buf & " AND (FLAG_DELI is null or  FLAG_DELI='')  and servicio='D' and usuario='" & cajero & "'"
 
            ''''23/09/2017 kenyo Multicaja
            buf = buf & " and estado='2' "
            'buf = buf & " and estado='2' and caja='" & caja & "'"
            ''''23/09/2017 kenyo Multicaja
   
            buf = buf & " and turno='" & turno & "'"
            buf = buf & " order by HORA"
        Else
            buf = "select FLAG_DELI  as 'X', tipo,serie,Numero,Fecha,Nombre as Cliente,Codigo,Moneda as M,Total,Estado as E,Servicio as S,Placa as 'E',Vendedor AS PERSONAL,Hora as 'Hora P.',Caja,Turno,Local,Tipo1,Serie1,Numero1,RENUMERO1,horae as 'horasal' from factura where local='01' and "
            buf = buf & "  fecha='" & amfecha & "'"
            buf = buf & " AND (FLAG_DELI is null or  FLAG_DELI='')     and servicio='D' and usuario='" & cajero & "'"
   
            ''''23/09/2017 kenyo Multicaja
            buf = buf & " and estado='2' "
            'buf = buf & " and estado='2' and caja='" & caja & "'"
            ''''23/09/2017 kenyo Multicaja
   
            buf = buf & " and turno='" & turno & "'"
            buf = buf & " and " & Combo1 & " like '%" & buffer & "%'"
            buf = buf & "  order by HORA "

            'indx = dbGrid1.Col
        End If

    End If
              
    Set rcconsulta = Nothing

    If rcconsulta.State = 1 Then
        rcconsulta.Close
        Set rcconsulta = Nothing

    End If
   
    rcconsulta.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = rcconsulta
    dbGrid1.refresh
    'MsgBox buf
                 
    'sw_consulta = 1
         
    If opcion1 = "15" Then
        dbGrid1.columns(0).Width = 1
        dbGrid1.columns(1).Width = 1 'TIPO
        dbGrid1.columns(2).Width = 430 'SERIE
        dbGrid1.columns(3).Width = 800 'NUMERO
        dbGrid1.columns(4).Width = 1 ' fecha
        dbGrid1.columns(5).Width = 2950
        dbGrid1.columns(6).Width = 1200
        dbGrid1.columns(7).Width = 1 'moneda
        dbGrid1.columns(8).Width = 800 'total
        dbGrid1.columns(9).Width = 1 'estado
        dbGrid1.columns(10).Width = 1 'servicio
        dbGrid1.columns(11).Width = 250 'estado
        dbGrid1.columns(12).Width = 1 'vendedor
        dbGrid1.columns(13).Width = 800
        dbGrid1.columns(14).Width = 1
        dbGrid1.columns(15).Width = 1
        dbGrid1.columns(16).Width = 1
        dbGrid1.columns(17).Width = 1
        dbGrid1.columns(18).Width = 1
        dbGrid1.columns(19).Width = 1
        dbGrid1.columns(20).Width = 1
        dbGrid1.columns(21).Width = 1

    End If
             
    If rcconsulta.RecordCount = 0 Then Exit Function

    If opcion1 = "15" Then
        If rcconsulta.RecordCount > 0 Then
            ir_hasta_abajo rcconsulta

        End If

    End If
               
    sql_consulta = 1
    Exit Function
    buffer = ""
    Exit Function

End Function

Function sql_consultae(sw As Integer)

    Dim buf       As String

    Dim queprecio As String

    Dim indx      As Integer

    Dim dbf1      As String

    Dim dbf2      As String

    Dim amfecha   As String

    indx = -1
    dbf1 = ""
    dbf2 = ""
    amfecha = Format(dia, "YYYYMMDD")
    buf = "select f.FLAG_DELI as 'X',f.tipo,f.serie,f.Numero,f.Fecha,f.Nombre as Cliente,f.Codigo,f.Moneda as M,f.Total,f.Estado as E,f.Servicio as S,f.Placa as 'E',f.Vendedor AS Personal,f.Hora,f.Caja,f.Turno,f.Local,f.Tipo1,f.Serie1,f.Numero1,f.horae as HoraSal,f.RENUMERO1,DATEDIFF(MI,f.horae,Convert(varchar(8),GetDate(), 108)) as 'Minutos',f.Total,f.RENUMERO2   from factura f where f.local='01' and "
    buf = buf & "  f.fecha='" & amfecha & "'"
    buf = buf & " AND f.RENUMERO1 ='S' and f.servicio='D' and f.usuario='" & cajero & "'"
    ''''23/09/2017 kenyo Multicaja
    buf = buf & " and f.estado='2' "
    'buf = buf & " and estado='2' and caja='" & caja & "'"
    ''''23/09/2017 kenyo Multicaja
    buf = buf & " and f.turno='" & turno & "'"
    buf = buf & " and f.vendedor='" & Trim("" & dbvariose.Fields("codigo")) & "'"
    buf = buf & " order by f.HORA"
            
    Set rcconsultae = Nothing

    If rcconsultae.State = 1 Then
        rcconsultae.Close
        Set rcconsultae = Nothing

    End If
   
    rcconsultae.Open buf, cn, adOpenStatic, adLockOptimistic
    Set DataGrid4.DataSource = rcconsultae
    DataGrid4.refresh
             
    DataGrid4.columns(0).Width = 1
    DataGrid4.columns(1).Width = 1 'TIPO
    DataGrid4.columns(2).Width = 430 'SERIE
    DataGrid4.columns(3).Width = 800 'NUMERO
    DataGrid4.columns(4).Width = 1 ' fecha
    DataGrid4.columns(5).Width = 2000 'CLIENTE
    DataGrid4.columns(6).Width = 1100
    DataGrid4.columns(7).Width = 1 'moneda
    DataGrid4.columns(8).Width = 1 'total
    DataGrid4.columns(9).Width = 1 'estado
    DataGrid4.columns(10).Width = 1 'servicio
    DataGrid4.columns(11).Width = 1 'estado
    DataGrid4.columns(12).Width = 780 'vendedor
    DataGrid4.columns(13).Width = 1
    DataGrid4.columns(14).Width = 1
    DataGrid4.columns(15).Width = 1
    DataGrid4.columns(16).Width = 1
    DataGrid4.columns(17).Width = 1
    DataGrid4.columns(18).Width = 1
    DataGrid4.columns(19).Width = 1
    DataGrid4.columns(20).Width = 670
    DataGrid4.columns(21).Width = 1
    DataGrid4.columns(22).Width = 670
    DataGrid4.columns(23).Width = 670
    DataGrid4.columns(24).Width = 1
           
    If rcconsultae.RecordCount = 0 Then
        Exit Function

    End If
              
    If rcconsultae.RecordCount > 0 Then
        ir_hasta_primeroe rcconsultae

    End If
              
    sql_consultae = 1
    Exit Function

    buffer = ""
    Exit Function

End Function

Function sql_consultageneral(sw As Integer)

    Dim buf       As String

    Dim queprecio As String

    Dim indx      As Integer

    Dim dbf1      As String

    Dim dbf2      As String

    Dim amfecha   As String

    indx = -1
    dbf1 = ""
    dbf2 = ""
    amfecha = Format(dia, "YYYYMMDD")
    buf = "select f.FLAG_DELI as 'X',f.tipo,f.serie,f.Numero,f.Fecha,f.Nombre as Cliente,f.Codigo,f.Moneda as M,f.Total,f.Estado as E,f.Servicio as S,f.Placa as 'E',f.Vendedor AS Personal,f.Hora,f.Caja,f.Turno,f.Local,f.Tipo1,f.Serie1,f.Numero1,f.horae as HoraSal,f.RENUMERO1,DATEDIFF(MI,f.horae,Convert(varchar(8),GetDate(), 108)) as 'Minutos',f.Total,f.RENUMERO2 ,fp.descripcio as 'F.PAGO' ,fp.recibe as 'recibe'  from factura f, fpagov fp where (f.tipo=fp.tipo and f.serie=fp.serie and f.numero=fp.numero) AND f.local='01' and "
    buf = buf & "  f.fecha='" & amfecha & "'"
    buf = buf & " AND f.RENUMERO1 ='S' and f.servicio='D' and f.usuario='" & cajero & "'"
    buf = buf & " and f.estado='2' "

    buf = buf & " and f.turno='" & turno & "'"
    buf = buf & " and f.vendedor='" & Trim("" & dbvariose.Fields("codigo")) & "'"
    buf = buf & " order by f.HORA"
            
    Set rcconsultageneral = Nothing

    If rcconsultageneral.State = 1 Then
        rcconsultageneral.Close
        Set rcconsultageneral = Nothing

    End If
   
    rcconsultageneral.Open buf, cn, adOpenStatic, adLockOptimistic
    Set DataGrid1.DataSource = rcconsultageneral
    DataGrid1.refresh
             
    DataGrid1.columns(0).Width = 1
    DataGrid1.columns(1).Width = 1 'TIPO
    DataGrid1.columns(2).Width = 430 'SERIE
    DataGrid1.columns(3).Width = 800 'NUMERO
    DataGrid1.columns(4).Width = 1 ' fecha
    DataGrid1.columns(5).Width = 2000 'CLIENTE
    DataGrid1.columns(6).Width = 1100
    DataGrid1.columns(7).Width = 1 'moneda
    DataGrid1.columns(8).Width = 1 'total
    DataGrid1.columns(9).Width = 1 'estado
    DataGrid1.columns(10).Width = 1 'servicio
    DataGrid1.columns(11).Width = 1 'estado
    DataGrid1.columns(12).Width = 780 'vendedor
    DataGrid1.columns(13).Width = 1
    DataGrid1.columns(14).Width = 1
    DataGrid1.columns(15).Width = 1
    DataGrid1.columns(16).Width = 1
    DataGrid1.columns(17).Width = 1
    DataGrid1.columns(18).Width = 1
    DataGrid1.columns(19).Width = 1
    DataGrid1.columns(20).Width = 670
    DataGrid1.columns(21).Width = 1
    DataGrid1.columns(22).Width = 670
    DataGrid1.columns(23).Width = 670
    DataGrid1.columns(24).Width = 1
               
    DataGrid1.columns(25).Width = 850
    DataGrid1.columns(26).Width = 700
               
    If rcconsultageneral.RecordCount = 0 Then
        Exit Function

    End If
              
    If rcconsultageneral.RecordCount > 0 Then
        ir_hasta_primeroe rcconsultageneral

    End If
              
    sql_consultageneral = 1
    Exit Function
    buffer = ""
    Exit Function

End Function

Sub ir_hasta_primerogeneral(rcconsultageneral As ADODB.Recordset)

    On Error GoTo cmd789111_err

    rcconsultageneral.MoveFirst
 
    Exit Sub
cmd789111_err:
    MsgBox "Aviso en ir ultimo " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub ir_hasta_primeroe(rcconsultae As ADODB.Recordset)

    On Error GoTo cmd789111_err

    rcconsultae.MoveFirst
 
    Exit Sub
cmd789111_err:
    MsgBox "Aviso en ir ultimo " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub ir_hasta_primerox(rcconsultax As ADODB.Recordset)

    On Error GoTo cmd789111_err

    rcconsultax.MoveFirst

    Exit Sub
cmd789111_err:
    MsgBox "Aviso en ir ultimo " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub ir_hasta_abajo(rcconsulta As ADODB.Recordset)

    On Error GoTo cmd789111_err

    rcconsulta.MoveFirst
 
    Exit Sub
cmd789111_err:
    MsgBox "Aviso en ir ultimo " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Label26_Click()

    Dim found As Integer

    On Error GoTo cmd8966_err

    If rcconsulta.RecordCount = 0 Then Exit Sub

    If "" & rcconsulta.Fields("X") = "S" Then
        rcconsulta.Fields("X") = ""
        rcconsulta.Fields("HoraSal") = ""
        rcconsulta.Fields("personal") = ""
        rcconsulta.Update
   
        opcion1 = "15"
        found = sql_consulta(1)

        opcion1 = "16"
        found = sql_consultaX(1)
   
        Exit Sub

    End If

    rcconsulta.Fields("X") = "S"
    rcconsulta.Fields("HoraSal") = ""
    rcconsulta.Fields("personal") = ""
    rcconsulta.Update

    opcion1 = "15"
    found = sql_consulta(1)
   
    opcion1 = "16"
    found = sql_consultaX(1)

    Exit Sub
cmd8966_err:
    MsgBox "Seleccione un Dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub table7_DblClick()
    BuscarDeliveryXVendedor

End Sub

Private Sub Timer1_Timer()
    diasistema = Format(Now, "dd/mm/yyyy")
    horasistema = Format(Now, "HH:MM:SS")

End Sub

