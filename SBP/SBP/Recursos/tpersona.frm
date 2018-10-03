VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form tpersona 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tabla de Personal"
   ClientHeight    =   9030
   ClientLeft      =   165
   ClientTop       =   -1020
   ClientWidth     =   14055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   14055
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   12600
      TabIndex        =   200
      Top             =   8760
      Visible         =   0   'False
      Width           =   13935
      Begin VB.CommandButton Command3 
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
         Left            =   10800
         TabIndex        =   203
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   202
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
      End
      Begin VB.TextBox bufferx 
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
         TabIndex        =   201
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   6735
         Left            =   240
         TabIndex        =   204
         Top             =   1080
         Width           =   12135
         _ExtentX        =   21405
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
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mesas en donde Pueden Trabajar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8655
      Left            =   120
      TabIndex        =   191
      Top             =   0
      Visible         =   0   'False
      Width           =   14055
      Begin VB.ComboBox psalon 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   193
         Top             =   960
         Width           =   3615
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   11760
         Picture         =   "tpersona.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   192
         ToolTipText     =   "Imprimir todo"
         Top             =   240
         Width           =   1470
      End
      Begin MSDataGridLib.DataGrid dbgrid11 
         Height          =   6855
         Left            =   240
         TabIndex        =   194
         Top             =   1560
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   12091
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
      Begin MSDataGridLib.DataGrid dbgrid12 
         Height          =   6855
         Left            =   7680
         TabIndex        =   195
         Top             =   1560
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   12091
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
      Begin VB.Label Label98 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Salones Mesas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   199
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label97 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Salones Mesas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7680
         TabIndex        =   198
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label96 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(Seleccionar)"
         Height          =   495
         Left            =   6000
         TabIndex        =   197
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label95 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(Borrar)"
         Height          =   495
         Left            =   6000
         TabIndex        =   196
         Top             =   2160
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Locales en donde Pueden Trabajar"
      Height          =   7695
      Left            =   0
      TabIndex        =   183
      Top             =   0
      Visible         =   0   'False
      Width           =   12135
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   10320
         Picture         =   "tpersona.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   184
         ToolTipText     =   "Imprimir todo"
         Top             =   6360
         Width           =   1470
      End
      Begin MSDataGridLib.DataGrid dbgrid9 
         Height          =   3135
         Left            =   240
         TabIndex        =   185
         Top             =   1560
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
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
      Begin MSDataGridLib.DataGrid dbgrid10 
         Height          =   3015
         Left            =   7680
         TabIndex        =   186
         Top             =   1560
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   5318
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
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
      Begin VB.Label Label93 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(Todos)"
         Height          =   495
         Left            =   6000
         TabIndex        =   190
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label70 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(Borrar)"
         Height          =   495
         Left            =   6000
         TabIndex        =   189
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label69 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(Seleccionar)"
         Height          =   495
         Left            =   6000
         TabIndex        =   188
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label68 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Autorizados"
         Height          =   495
         Left            =   7680
         TabIndex        =   187
         Top             =   1080
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Parametros Adicionales"
      Height          =   8775
      Left            =   0
      TabIndex        =   106
      Top             =   120
      Visible         =   0   'False
      Width           =   14055
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   2775
         Left            =   11160
         TabIndex        =   211
         Top             =   6840
         Visible         =   0   'False
         Width           =   2655
         Begin VB.TextBox rw1 
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
            IMEMode         =   3  'DISABLE
            Left            =   0
            MaxLength       =   1
            TabIndex        =   224
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox rw2 
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
            IMEMode         =   3  'DISABLE
            Left            =   0
            MaxLength       =   1
            TabIndex        =   223
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox rw3 
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
            IMEMode         =   3  'DISABLE
            Left            =   0
            MaxLength       =   1
            TabIndex        =   222
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox rw4 
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
            IMEMode         =   3  'DISABLE
            Left            =   0
            MaxLength       =   1
            TabIndex        =   221
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox rw5 
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
            IMEMode         =   3  'DISABLE
            Left            =   0
            MaxLength       =   1
            TabIndex        =   220
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox rw6 
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
            IMEMode         =   3  'DISABLE
            Left            =   0
            MaxLength       =   1
            TabIndex        =   219
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox rw7 
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
            IMEMode         =   3  'DISABLE
            Left            =   0
            MaxLength       =   1
            TabIndex        =   218
            Top             =   2520
            Width           =   615
         End
         Begin VB.TextBox rw8 
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
            IMEMode         =   3  'DISABLE
            Left            =   0
            MaxLength       =   1
            TabIndex        =   217
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox rw9 
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
            IMEMode         =   3  'DISABLE
            Left            =   0
            MaxLength       =   1
            TabIndex        =   216
            Top             =   3240
            Width           =   615
         End
         Begin VB.TextBox rw10 
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
            IMEMode         =   3  'DISABLE
            Left            =   0
            MaxLength       =   1
            TabIndex        =   215
            Top             =   3600
            Width           =   615
         End
         Begin VB.TextBox rw11 
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
            IMEMode         =   3  'DISABLE
            Left            =   0
            MaxLength       =   1
            TabIndex        =   214
            Top             =   3960
            Width           =   615
         End
         Begin VB.TextBox rw13 
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
            IMEMode         =   3  'DISABLE
            Left            =   0
            MaxLength       =   1
            TabIndex        =   213
            Top             =   4320
            Width           =   615
         End
         Begin VB.TextBox rw14 
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
            IMEMode         =   3  'DISABLE
            Left            =   0
            MaxLength       =   1
            TabIndex        =   212
            Top             =   4680
            Width           =   615
         End
         Begin VB.Label lblRW 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "R/W"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   960
            TabIndex        =   225
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.TextBox modificaproducto 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   208
         Top             =   8160
         Width           =   615
      End
      Begin VB.TextBox modificacompra 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   206
         Top             =   7800
         Width           =   615
      End
      Begin VB.TextBox v14 
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
         IMEMode         =   3  'DISABLE
         Left            =   13200
         MaxLength       =   1
         TabIndex        =   205
         Top             =   6120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox v13 
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
         IMEMode         =   3  'DISABLE
         Left            =   13200
         MaxLength       =   1
         TabIndex        =   145
         Top             =   5760
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox conexionremota 
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
         IMEMode         =   3  'DISABLE
         Left            =   10320
         MaxLength       =   1
         TabIndex        =   144
         Top             =   7800
         Width           =   615
      End
      Begin VB.TextBox escajero 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   143
         Top             =   4560
         Width           =   615
      End
      Begin VB.TextBox esvendedor 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   142
         Top             =   4920
         Width           =   615
      End
      Begin VB.TextBox minireporte 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   141
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox tienda 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   140
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox productos 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   139
         Top             =   6720
         Width           =   615
      End
      Begin VB.TextBox terminal 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   138
         Top             =   7080
         Width           =   615
      End
      Begin VB.TextBox cierre 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   137
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox descuento 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   136
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox verificador 
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
         IMEMode         =   3  'DISABLE
         Left            =   10320
         MaxLength       =   1
         TabIndex        =   135
         Top             =   8160
         Width           =   615
      End
      Begin VB.TextBox vreloj 
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
         IMEMode         =   3  'DISABLE
         Left            =   10320
         MaxLength       =   1
         TabIndex        =   134
         Top             =   7080
         Width           =   615
      End
      Begin VB.TextBox inicia 
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
         Left            =   10320
         MaxLength       =   1
         TabIndex        =   133
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox cprecios 
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
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   132
         Top             =   5280
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   12000
         Picture         =   "tpersona.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   131
         ToolTipText     =   "Imprimir todo"
         Top             =   1200
         Width           =   1470
      End
      Begin VB.TextBox vecostoimp 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   130
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox v11 
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
         IMEMode         =   3  'DISABLE
         Left            =   13200
         MaxLength       =   1
         TabIndex        =   129
         Top             =   6480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox apertura 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   128
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox congela 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   127
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox copia 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   126
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox anula 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   125
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox cuadre 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   124
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox parame4 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   123
         Top             =   7800
         Width           =   615
      End
      Begin VB.TextBox parame2 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   122
         Top             =   7080
         Width           =   615
      End
      Begin VB.TextBox parame3 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   121
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox parame1 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   3
         TabIndex        =   120
         Top             =   6720
         Width           =   615
      End
      Begin VB.TextBox parame 
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
         IMEMode         =   3  'DISABLE
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   119
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox caja 
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
         IMEMode         =   3  'DISABLE
         Left            =   10340
         MaxLength       =   1
         TabIndex        =   105
         Top             =   165
         Width           =   495
      End
      Begin VB.TextBox veclave 
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
         IMEMode         =   3  'DISABLE
         Left            =   10320
         MaxLength       =   1
         TabIndex        =   118
         Top             =   6720
         Width           =   615
      End
      Begin VB.TextBox vevend 
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
         IMEMode         =   3  'DISABLE
         Left            =   10320
         MaxLength       =   1
         TabIndex        =   117
         Top             =   6360
         Width           =   615
      End
      Begin VB.TextBox v10 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   116
         Top             =   4200
         Width           =   615
      End
      Begin VB.TextBox v9 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   115
         Top             =   5400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox v8 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   114
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox v7 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   113
         Top             =   5040
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox v6 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   112
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox v5 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   111
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox v4 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   110
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox v3 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   109
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox v2 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   108
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox v1 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   1
         TabIndex        =   107
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox clave 
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
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   104
         Top             =   480
         Width           =   1695
      End
      Begin ChamaleonButton.ChameleonBtn BtnPermisos 
         Height          =   585
         Left            =   11040
         TabIndex        =   231
         Top             =   240
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1032
         BTYPE           =   4
         TX              =   "Generar Permisos"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         MICON           =   "tpersona.frx":1A5E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame vemesa 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   8040
         TabIndex        =   232
         Top             =   1320
         Width           =   3015
         Begin VB.TextBox mueveproducto 
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
            IMEMode         =   3  'DISABLE
            Left            =   2160
            MaxLength       =   1
            TabIndex        =   245
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox muevemesa 
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
            IMEMode         =   3  'DISABLE
            Left            =   2160
            MaxLength       =   1
            TabIndex        =   244
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox juntamesa 
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
            IMEMode         =   3  'DISABLE
            Left            =   2160
            MaxLength       =   1
            TabIndex        =   243
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox borra_comanda 
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
            IMEMode         =   3  'DISABLE
            Left            =   2160
            MaxLength       =   1
            TabIndex        =   242
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox precuenta 
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
            IMEMode         =   3  'DISABLE
            Left            =   2160
            MaxLength       =   1
            TabIndex        =   241
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox despacho 
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
            IMEMode         =   3  'DISABLE
            Left            =   2160
            MaxLength       =   1
            TabIndex        =   240
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblPrecuenta 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Precuenta"
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
            Left            =   0
            TabIndex        =   239
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label lblDespachos 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Despachos"
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
            Left            =   0
            TabIndex        =   238
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label lblBorraComandas 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Borra Comandas"
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
            Left            =   0
            TabIndex        =   237
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblJuntarMesa 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Juntar Mesa"
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
            Left            =   0
            TabIndex        =   236
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label lblMoverMesa 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Mover Mesa"
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
            Index           =   1
            Left            =   0
            TabIndex        =   235
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label lblMoverMesa 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Mover Mesa Producto"
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
            Index           =   0
            Left            =   0
            TabIndex        =   234
            Top             =   2160
            Width           =   2175
         End
         Begin VB.Label lblAccesosMenúF 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Accesos Menú Mesa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   0
            TabIndex        =   233
            Top             =   0
            Width           =   2745
         End
      End
      Begin VB.Label lblMásPermisos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Más Permisos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8160
         TabIndex        =   230
         Top             =   5940
         Width           =   2775
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   11880
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lblRestriccion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Restricción de Cajas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4320
         TabIndex        =   229
         Top             =   5940
         Width           =   2775
      End
      Begin VB.Label lblAccesosMenú 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Permiso Menú Caja"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4320
         TabIndex        =   228
         Top             =   1240
         Width           =   2775
      End
      Begin VB.Label lblAccesosDirectos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Accesos Directos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   600
         TabIndex        =   227
         Top             =   5940
         Width           =   2775
      End
      Begin VB.Label lblMenuPrincipalh 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Menu Principal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   600
         TabIndex        =   226
         Top             =   1240
         Width           =   2775
      End
      Begin VB.Label Label100 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ModificaProducto"
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
         Left            =   600
         TabIndex        =   209
         Top             =   8160
         Width           =   2175
      End
      Begin VB.Label Label64 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ModificaCompras/Guias"
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
         Left            =   600
         TabIndex        =   207
         Top             =   7800
         Width           =   2175
      End
      Begin VB.Label Label94 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acivo Fijo                                      Hotel                      HistoriaClinica"
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
         Height          =   1095
         Left            =   11160
         TabIndex        =   182
         Top             =   5760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label92 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Conexion Remota"
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
         Left            =   8160
         TabIndex        =   181
         Top             =   7800
         Width           =   2175
      End
      Begin VB.Label Label87 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Puede Hacer Caja?"
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
         TabIndex        =   180
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label Label86 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "En caja es Vendedor?"
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
         TabIndex        =   179
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label Label82 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reporte Consolidad"
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
         Left            =   600
         TabIndex        =   178
         Top             =   7440
         Width           =   2175
      End
      Begin VB.Label Label90 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tienda"
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
         Left            =   600
         TabIndex        =   177
         Top             =   6360
         Width           =   2175
      End
      Begin VB.Label Label89 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Productos"
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
         Left            =   600
         TabIndex        =   176
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label Label88 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Terminal Pedidos"
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
         Left            =   600
         TabIndex        =   175
         Top             =   7080
         Width           =   2175
      End
      Begin VB.Label Label80 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cierre Caja"
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
         TabIndex        =   174
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label78 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hace Descuentos"
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
         TabIndex        =   173
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label76 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SoloVerificador"
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
         Left            =   8160
         TabIndex        =   172
         Top             =   8160
         Width           =   2175
      End
      Begin VB.Label Label73 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Personal Reloj"
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
         Left            =   8160
         TabIndex        =   171
         Top             =   7080
         Width           =   2175
      End
      Begin VB.Label Label74 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inicializa Data"
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
         Left            =   8160
         TabIndex        =   170
         Top             =   7440
         Width           =   2175
      End
      Begin VB.Label Label71 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Permiso Cambio Precios"
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
         Left            =   4320
         TabIndex        =   169
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label Label65 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ve Costo Import."
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
         TabIndex        =   168
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label clavex 
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
         Left            =   4440
         TabIndex        =   167
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label62 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Abre Gaveta"
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
         TabIndex        =   166
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label61 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descongela"
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
         TabIndex        =   165
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label60 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saca Copia en Caja"
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
         TabIndex        =   164
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label59 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Anula en Caja"
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
         TabIndex        =   163
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label58 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuadre Parcial Caja"
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
         TabIndex        =   162
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label56 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cajas        01,02,03 Terminales T1,T2,T3"
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
         Height          =   1455
         Left            =   4320
         TabIndex        =   161
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label Label55 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Solamente estas cajas?"
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
         TabIndex        =   160
         Top             =   6360
         Width           =   2175
      End
      Begin VB.Label Label54 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1.Ger    2.Adm  3.Sup  4.Caja   5.Caja/Pedido  6.Pedidos 7.Almacen "
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
         Left            =   8160
         TabIndex        =   159
         Top             =   165
         Width           =   2175
      End
      Begin VB.Label Label53 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CambiaClave"
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
         Left            =   8160
         TabIndex        =   158
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label Label52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AccePersonal"
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
         Left            =   8160
         TabIndex        =   157
         Top             =   6360
         Width           =   2175
      End
      Begin VB.Label Label37 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reporte"
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
         Left            =   600
         TabIndex        =   156
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label36 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contable"
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
         Left            =   600
         TabIndex        =   155
         Top             =   5400
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label35 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Personal"
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
         Left            =   600
         TabIndex        =   154
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label34 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Produccion"
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
         Left            =   600
         TabIndex        =   153
         Top             =   5040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label33 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tesoreria"
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
         Left            =   600
         TabIndex        =   152
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Almacen"
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
         Left            =   600
         TabIndex        =   151
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label31 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Compras"
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
         Left            =   600
         TabIndex        =   150
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label30 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ventas"
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
         Left            =   600
         TabIndex        =   149
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label29 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tienda"
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
         Left            =   600
         TabIndex        =   148
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tablas"
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
         Left            =   600
         TabIndex        =   147
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clave Acceso"
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
         Left            =   600
         TabIndex        =   146
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   8775
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   14055
      Begin VB.TextBox clavere 
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
         MaxLength       =   10
         TabIndex        =   66
         Top             =   6840
         Width           =   375
      End
      Begin VB.TextBox local1 
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
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   65
         Top             =   6240
         Width           =   1095
      End
      Begin VB.TextBox dueno 
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
         Left            =   7440
         MaxLength       =   1
         TabIndex        =   64
         Top             =   6240
         Width           =   735
      End
      Begin VB.TextBox tipo 
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
         Left            =   2280
         MaxLength       =   1
         TabIndex        =   63
         Text            =   "N"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   12480
         Picture         =   "tpersona.frx":1A7A
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Imprimir todo"
         Top             =   1320
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   12480
         Picture         =   "tpersona.frx":2344
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   240
         Width           =   1470
      End
      Begin VB.TextBox pocket 
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
         Left            =   7440
         MaxLength       =   3
         TabIndex        =   60
         Top             =   5880
         Width           =   735
      End
      Begin VB.TextBox reloj 
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
         MaxLength       =   6
         TabIndex        =   59
         Top             =   7200
         Width           =   1335
      End
      Begin VB.TextBox seccion 
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
         Left            =   7440
         MaxLength       =   6
         TabIndex        =   58
         Top             =   5520
         Width           =   975
      End
      Begin VB.TextBox zona 
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
         TabIndex        =   57
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox profesion 
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
         MaxLength       =   15
         TabIndex        =   56
         Top             =   5880
         Width           =   3375
      End
      Begin VB.TextBox cargo 
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
         MaxLength       =   15
         TabIndex        =   55
         Top             =   5520
         Width           =   3375
      End
      Begin VB.ComboBox civil 
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
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   5160
         Width           =   975
      End
      Begin VB.ComboBox sexo 
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   5160
         Width           =   1335
      End
      Begin VB.ComboBox estado 
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   7440
         Width           =   1935
      End
      Begin VB.TextBox correo 
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
         MaxLength       =   60
         TabIndex        =   51
         Top             =   4560
         Width           =   6135
      End
      Begin VB.TextBox telefono2 
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
         Left            =   6120
         MaxLength       =   15
         TabIndex        =   50
         Top             =   4200
         Width           =   2295
      End
      Begin VB.TextBox telefono1 
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
         Left            =   4200
         MaxLength       =   15
         TabIndex        =   49
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox telefono 
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
         MaxLength       =   15
         TabIndex        =   48
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox distrito 
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
         MaxLength       =   15
         TabIndex        =   47
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox dpto 
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
         MaxLength       =   15
         TabIndex        =   46
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox direccion 
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
         MaxLength       =   60
         TabIndex        =   45
         Top             =   2640
         Width           =   6135
      End
      Begin VB.TextBox contacto 
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
         MaxLength       =   60
         TabIndex        =   44
         Top             =   2160
         Width           =   6135
      End
      Begin VB.TextBox nombrec 
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
         MaxLength       =   60
         TabIndex        =   43
         Top             =   1800
         Width           =   6135
      End
      Begin VB.TextBox nombre 
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
         Left            =   2280
         MaxLength       =   60
         TabIndex        =   42
         Top             =   1440
         Width           =   6135
      End
      Begin VB.TextBox codigo1 
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
         MaxLength       =   11
         TabIndex        =   41
         Top             =   960
         Width           =   1815
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
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   40
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox ipss 
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
         MaxLength       =   20
         TabIndex        =   39
         Top             =   6480
         Width           =   1335
      End
      Begin VB.ComboBox regimen1 
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
         Left            =   11040
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox fechaing 
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
         MaxLength       =   10
         TabIndex        =   37
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox fechacese 
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
         MaxLength       =   10
         TabIndex        =   36
         Top             =   3480
         Width           =   1335
      End
      Begin VB.ComboBox moneda 
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
         Left            =   11040
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox basico 
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
         MaxLength       =   10
         TabIndex        =   34
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox jornal 
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
         MaxLength       =   10
         TabIndex        =   33
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox mtardanza 
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
         MaxLength       =   10
         TabIndex        =   32
         Top             =   5760
         Width           =   1335
      End
      Begin VB.TextBox dtardanza 
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
         MaxLength       =   10
         TabIndex        =   31
         Top             =   6120
         Width           =   1335
      End
      Begin VB.TextBox fechavaca 
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
         MaxLength       =   10
         TabIndex        =   30
         Top             =   4200
         Width           =   1335
      End
      Begin VB.TextBox ini1 
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
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   29
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox fin1 
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
         Left            =   9480
         MaxLength       =   10
         TabIndex        =   28
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox por1 
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
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox valor1 
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
         Left            =   11400
         MaxLength       =   10
         TabIndex        =   26
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox ini2 
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
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   25
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox fin2 
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
         Left            =   9480
         MaxLength       =   10
         TabIndex        =   24
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox por2 
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
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox valor2 
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
         Left            =   11400
         MaxLength       =   10
         TabIndex        =   22
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox ini3 
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
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox fin3 
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
         Left            =   9480
         MaxLength       =   10
         TabIndex        =   20
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox por3 
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
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   19
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox valor3 
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
         Left            =   11400
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox ini4 
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
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   17
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox fin4 
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
         Left            =   9480
         MaxLength       =   10
         TabIndex        =   16
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox por4 
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
         Left            =   10440
         MaxLength       =   10
         TabIndex        =   15
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox valor4 
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
         Left            =   11400
         MaxLength       =   10
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin ChamaleonButton.ChameleonBtn Label46 
         Height          =   585
         Left            =   6720
         TabIndex        =   210
         Top             =   360
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1032
         BTYPE           =   4
         TX              =   "CLAVES DE ACCESO"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
         MICON           =   "tpersona.frx":2C0E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label45 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AccesoPersonal 1.Marca 2.Adm"
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
         Left            =   8520
         TabIndex        =   103
         Top             =   6840
         Width           =   2535
      End
      Begin VB.Label Label72 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local Perteneciente"
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
         TabIndex        =   102
         Top             =   6240
         Width           =   2175
      End
      Begin VB.Label Label49 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dueño"
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
         Left            =   5760
         TabIndex        =   101
         Top             =   6240
         Width           =   1695
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "N.atural J.Juridica O.Otros"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2760
         TabIndex        =   100
         Top             =   600
         Width           =   2430
      End
      Begin VB.Label Label47 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Persona"
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
         TabIndex        =   99
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label57 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Terminal Pocket"
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
         Left            =   5760
         TabIndex        =   98
         Top             =   5880
         Width           =   1695
      End
      Begin VB.Label Label44 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Forma de Pago (Subgrupo)"
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
         Left            =   8520
         TabIndex        =   97
         Top             =   7200
         Width           =   2535
      End
      Begin VB.Label Label43 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SeccionProdcc."
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
         Left            =   5760
         TabIndex        =   96
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Profesion"
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
         TabIndex        =   95
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cargo"
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
         Left            =   105
         TabIndex        =   94
         Top             =   5490
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado Civil"
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
         Left            =   5760
         TabIndex        =   93
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sexo"
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
         Left            =   135
         TabIndex        =   92
         Top             =   5160
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
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
         TabIndex        =   91
         Top             =   7440
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Zona"
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
         TabIndex        =   90
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Correo Electronico"
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
         TabIndex        =   89
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Telefonos"
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
         TabIndex        =   88
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Distrito"
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
         TabIndex        =   87
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Departamento"
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
         TabIndex        =   86
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Direccion"
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
         TabIndex        =   85
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Contacto"
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
         TabIndex        =   84
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombre Comercial"
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
         TabIndex        =   83
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Apellidos Nombres"
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
         TabIndex        =   82
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo Alterno"
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
         TabIndex        =   81
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo"
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
         TabIndex        =   80
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label51 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ipss"
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
         Left            =   8520
         TabIndex        =   79
         Top             =   6480
         Width           =   2535
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Datos Planilla"
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
         Left            =   8535
         TabIndex        =   78
         Top             =   2745
         Width           =   3855
      End
      Begin VB.Label regimen 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Regimen"
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
         Left            =   8520
         TabIndex        =   77
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F.Cese                                                                   F.Ingreso"
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
         Left            =   8520
         TabIndex        =   76
         Top             =   3480
         Width           =   2535
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Moneda"
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
         Left            =   8520
         TabIndex        =   75
         Top             =   4560
         Width           =   2535
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Basico                                                                          Jornal"
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
         Left            =   8520
         TabIndex        =   74
         Top             =   4920
         Width           =   2535
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tolerancia"
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
         Left            =   8520
         TabIndex        =   73
         Top             =   5760
         Width           =   2535
      End
      Begin VB.Label Label23 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DsctoTarda"
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
         Left            =   8520
         TabIndex        =   72
         Top             =   6120
         Width           =   2535
      End
      Begin VB.Label Label66 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaVacac"
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
         Left            =   8520
         TabIndex        =   71
         Top             =   4200
         Width           =   2535
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tabla Comisiones"
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
         Left            =   8520
         TabIndex        =   70
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Ini    |   Final"
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
         Left            =   8520
         TabIndex        =   69
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "%"
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
         Left            =   10440
         TabIndex        =   68
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label42 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor"
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
         Left            =   11400
         TabIndex        =   67
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "HuellaVerifica"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12600
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Huella"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12600
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   -15
      ScaleHeight     =   675
      ScaleWidth      =   12435
      TabIndex        =   2
      Top             =   15
      Width           =   12495
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
         Picture         =   "tpersona.frx":2C2A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   735
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
         Height          =   375
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   5280
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Buscar"
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
         Left            =   10560
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
         Picture         =   "tpersona.frx":3E3C
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
         Picture         =   "tpersona.frx":504E
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
         Picture         =   "tpersona.frx":6260
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
         Picture         =   "tpersona.frx":7472
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   45
      TabIndex        =   0
      Top             =   810
      Width           =   12495
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7935
         Left            =   135
         TabIndex        =   1
         Top             =   255
         Width           =   12255
         _ExtentX        =   21616
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
         ColumnCount     =   4
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "Local"
            Caption         =   "Local"
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
               ColumnWidth     =   5940.284
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
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
   End
   Begin VB.Menu ki8933 
      Caption         =   "Sal&OnMesas"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tpersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mytablel As New ADODB.Recordset

Dim txempre  As New ADODB.Recordset

Private Sub ajdu1_Click()

    If Frame6.Visible = True Then Exit Sub
    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Nuevo"
    cmdGuardar.Enabled = True
    habilita 1
    codigo.Enabled = True
    codigo = ""
    codigo.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    If Frame6.Visible = True Then Exit Sub

    buf = txempre.Fields("codigo")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + txempre.Fields("codigo"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    cn.Execute ("delete from userlocal where codigo='" & buf & "'")
    txempre.Delete
    Command1_Click

    Exit Sub
cmd656_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

'''26/07/2017 kenyo Registro Rapido de Personal
Sub TodoMenuPrincipal()
    v1 = "S"
    v2 = "S"
    v3 = "S"
    v4 = "S"
    v5 = "S"
    v6 = "S"
    v7 = "S"
    v8 = "S"
    v9 = "S"
    v10 = "S"
    v11 = "S"

End Sub

Sub NadaMenuPrincipal()
    v1 = "N"
    v2 = "N"
    v3 = "N"
    v4 = "N"
    v5 = "N"
    v6 = "N"
    v7 = "N"
    v8 = "N"
    v9 = "N"
    v10 = "N"
    v11 = "N"

End Sub

Sub TodoAccesoDirectos()
    tienda = "S"
    productos = "S"
    terminal = "S"
    minireporte = "S"
    modificacompra = "S"
    modificaproducto = "S"

End Sub

Sub NadaAccesoDirectos()
    tienda = "N"
    productos = "N"
    terminal = "N"
    minireporte = "N"
    modificacompra = "N"
    modificaproducto = "N"

End Sub

Sub TodoPermisoCaja()
    cuadre = "S"
    cierre = "S"
    anula = "S"
    copia = "S"
    congela = "S"
    apertura = "S"
    vecostoimp = "S"
    descuento = "S"
    escajero = "S"
    esvendedor = "S"
    cprecios = "S"

End Sub

Sub NadaPermisoCaja()
    cuadre = "N"
    cierre = "N"
    anula = "N"
    copia = "N"
    congela = "N"
    apertura = "N"
    vecostoimp = "N"
    descuento = "N"
    escajero = "N"
    esvendedor = "N"
    cprecios = "N"

End Sub

Sub TodoAccesosPedidos()
    despacho = "S"
    precuenta = "S"
    borra_comanda = "S"
    juntamesa = "S"
    muevemesa = "S"
    mueveproducto = "S"

End Sub

Sub NadaAccesosPedidos()
    despacho = "N"
    precuenta = "N"
    borra_comanda = "N"
    juntamesa = "N"
    muevemesa = "N"
    mueveproducto = "N"

End Sub

Sub TodoMasPermiso()
    vevend = "S"
    veclave = "S"
    vreloj = "S"
    inicia = "S"
    conexionremota = "S"

End Sub

Sub NadaMasPermiso()
    vevend = "N"
    veclave = "N"
    vreloj = "N"
    inicia = "N"
    conexionremota = "N"

End Sub

Private Sub BtnPermisos_Click()

    If Frame2.Caption = "Modifica" Then
    
        '''27/10/2017 Testing registro de permisos personal
        If caja = " " Or caja = "" Then
            MsgBox "Seleccione tipo de usuario", 48, "Aviso"
            caja.SetFocus
            Exit Sub

        End If

        '''27/10/2017 Testing registro de permisos personal
    
        If MsgBox("Desea ACTUALIZAR permisos???", 1, "Aviso") <> 1 Then Exit Sub

    End If
    
    If caja = 1 Then ' Si es gerente

        TodoMenuPrincipal
        TodoAccesoDirectos
        TodoPermisoCaja
        TodoAccesosPedidos
        TodoMasPermiso
    
        dueno = "S"
        clavere = "2"
    
    ElseIf caja = 2 Then ' Si es administrador
        TodoMenuPrincipal
        TodoAccesoDirectos
        TodoPermisoCaja
        TodoAccesosPedidos
        TodoMasPermiso
    
        inicia = "N"
        dueno = "N"
        clavere = "2"
    
    ElseIf caja = 3 Then ' Si es supervisor
        TodoMenuPrincipal
        TodoAccesoDirectos
        TodoPermisoCaja
        TodoAccesosPedidos
        TodoMasPermiso
    
        inicia = "N"
        v10 = "N"
        minireporte = "N"
        dueno = "N"
        clavere = "1"
    
    ElseIf caja = 4 Then ' Si es cajero
        NadaMenuPrincipal
        NadaAccesoDirectos
        TodoPermisoCaja
        NadaAccesosPedidos
        NadaMasPermiso
    
        tienda = "S"
        cprecios = "N"
        vreloj = "S"
        despacho = "S"
        dueno = "N"
        clavere = "1"
     
    ElseIf caja = 5 Then ' Si es cajero/pedido
        NadaMenuPrincipal
        NadaAccesoDirectos
        TodoPermisoCaja
        TodoAccesosPedidos
        NadaMasPermiso
    
        tienda = "S"
        cprecios = "N"
        vreloj = "S"
        dueno = "N"
        clavere = "1"
     
        'Color por familia y producto  30/05/2018
        esvendedor = "S"
        'Color por familia y producto  30/05/2018
    
    ElseIf caja = 6 Then ' Si es pedidos mesas
        NadaMenuPrincipal
        NadaAccesoDirectos
        NadaPermisoCaja
        NadaAccesosPedidos
        NadaMasPermiso
    
        tienda = "S"
        vreloj = "S"
        despacho = "S"
        escajero = "S"
        dueno = "N"
        clavere = "1"
            
        'Color por familia y producto  30/05/2018
        esvendedor = "S"
        'Color por familia y producto  30/05/2018
    
    ElseIf caja = 7 Then ' Si es almacenero
        NadaMenuPrincipal
        NadaAccesoDirectos
        NadaPermisoCaja
        NadaAccesosPedidos
        NadaMasPermiso

        vreloj = "S"
        v4 = "S"
        v5 = "S"
        dueno = "N"
        clavere = "1"
    
    End If

End Sub

''26/07/2017 kenyo Registro Rapido de Personal

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    Command1_Click

End Sub

Private Sub bufferx_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Frame4.Visible = False
        Frame4.Enabled = False

        If opcion1 = "22" Then
            reloj.SetFocus
            Exit Sub

        End If

        If opcion1 = "2" Then
            seccion.SetFocus
            Exit Sub

        End If
   
        Exit Sub

    End If

    Command3_Click

End Sub

Private Sub bufferx_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode <> 13 And KeyCode <> 27 Then
        ejecutax 0

    End If

End Sub

Private Sub CAJA_KeyPress(KeyAscii As Integer)

    '''26/07/2017 kenyo Registro Rapido de Personal
    If (KeyAscii >= 97) And (KeyAscii < 122) Or (KeyAscii >= 65) And (KeyAscii < 90) Then
        KeyAscii = 8

    End If

    '''26/07/2017 kenyo Registro Rapido de Personal
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
    djuer1_Click

End Sub

Private Sub cmdSave_Click()
    f8443_Click

End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(codigo) = 0 Then Exit Sub
    nombre.SetFocus

End Sub

Private Sub Command1_Click()
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    opcion1 = "1"
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim cad As String

    If opcion1 = "1" Then  'bodega
        If Len(buffer) = 0 Then
            cad = "SELECT * from vendedor    "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT *  from vendedor   where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If txempre.State = 1 Then txempre.Close
        txempre.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbGrid1.DataSource = txempre
        dbGrid1.columns(0).Width = 4000
        dbGrid1.columns(1).Width = 2000

        If txempre.RecordCount > 0 Then
            dbGrid1.SetFocus

        End If

    End If

End Sub

Private Sub Command2_Click()
    Frame3.Visible = False

End Sub

Private Sub Command3_Click()
    ejecutax 1

End Sub

Private Sub Command4_Click()
    Frame5.Visible = False

End Sub

Private Sub Command5_Click()
    Frame6.Visible = False

End Sub

Private Sub Command6_Click()

    Dim buf As String

    On Error GoTo cmd8000_err

    buf = Trim("" & txempre.Fields("CODIGO"))

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If Len(Trim(buf)) = 0 Then
        Exit Sub

    End If

    thuellad.tipo = "personal"
    thuellad.codigo = buf
    thuellad.nombre = "" & txempre.Fields("nombre")
    thuellad.Show 1
    Exit Sub
cmd8000_err:
    MsgBox "Seleccione Dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command7_Click()

    Dim buf As String

    On Error GoTo cmd88000_err

    buf = "" & txempre.Fields("CODIGO")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If Len(Trim(buf)) = 0 Then
        Exit Sub

    End If

    thuellat.tipo = "personal"
    thuellat.codigo = buf
    thuellat.nombre = "" & txempre.Fields("nombre")
    thuellat.Show 1
    Exit Sub
cmd88000_err:
    MsgBox "Seleccione Dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'codigo = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'codigo.SetFocus
        'codigo_KeyPress 13
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

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        bufferx.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            codigo = DBGrid2.columns(1)
            Frame4.Visible = False
            Frame4.Enabled = False
            codigo.SetFocus
            codigo_KeyPress 13

        End If

        If opcion1 = "2" Then
            seccion = DBGrid2.columns(1)
            Frame4.Visible = False
            Frame4.Enabled = False
            seccion.SetFocus

        End If

        If opcion1 = "4" Then

            'tipopla = DBGrid1.Columns(1)
            'Frame1.Visible = False
            'tipopla.SetFocus
        End If

        If opcion1 = "22" Then
            reloj = DBGrid2.columns(1)
            Frame4.Visible = False
            Frame4.Enabled = False
            reloj.SetFocus

        End If

    End If

End Sub

Private Sub dbgrid2_KeyPress(KeyAscii As Integer)

    Dim buf  As String

    Dim buf2 As String

    If KeyAscii <> 13 And KeyAscii <> 27 Then
        If KeyAscii = 8 Then
            If Len(bufferx) > 0 Then
                buf = Mid$(bufferx, 1, Len(bufferx) - 1)
                bufferx = buf
                KeyAscii = 0
            Else
                KeyAscii = 0
                Exit Sub

            End If

        End If

        buf = Chr(KeyAscii)

        If Chr(KeyAscii) = "*" Then
            buf = ""
            bufferx = buf

        End If

        If KeyAscii <> 13 Then
            bufferx = bufferx + buf

        End If

        buf = bufferx
        ejecutax 0
         
    End If

End Sub

Private Sub djuer1_Click()

    If Frame6.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "vendedor"
    reporgen.Show 1

End Sub

Private Sub dlo132_Click()

    If Frame6.Visible = True Then
        Frame6.Visible = False
        Exit Sub

    End If

    If Frame5.Visible = True Then
        Frame5.Visible = False
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

    tpersona.Hide
    Unload tpersona

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    If Frame6.Visible = True Then Exit Sub

    buf = txempre.Fields("codigo")

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
    codigo.Enabled = False
    nombre.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    If Frame6.Visible = True Then Exit Sub

    buf = txempre.Fields("codigo")

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
    codigo.Enabled = False
    nombre.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    Command1_Click

    'cn.Execute ("update vendedor set value local='01'")
End Sub

Sub consulta_reloj()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.ListIndex = 0

    Frame4.Visible = True
    Frame4.Enabled = True
    bufferx = ""
    opcion1 = "22"
    ejecutax 1

End Sub

Sub consulta_seccion()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.ListIndex = 0

    Frame4.Visible = True
    Frame4.Enabled = True
    buffer = ""
    opcion1 = "2"
    ejecutax 1

End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "nombre"
    Combo1.AddItem "codigo"
    Combo1.ListIndex = 0

    regimen1.Clear
    regimen1.AddItem "E"
    regimen1.AddItem "O"
    regimen1.AddItem "I"
    regimen1.ListIndex = 0

    sexo.Clear
    sexo.AddItem "M"
    sexo.AddItem "F"
    sexo.ListIndex = 0

    moneda.Clear
    moneda.AddItem "S"
    moneda.AddItem "D"
    moneda.ListIndex = 0

    civil.Clear
    civil.AddItem "S"
    civil.AddItem "C"
    civil.ListIndex = 0

    'Combo1.Clear
    'Combo1.AddItem "NOMBRE"
    'Combo1.AddItem "CODIGO"
    'Combo1.ListIndex = 0

    estado.Clear
    estado.AddItem "ACTIVO"
    estado.AddItem "NOACTIVO"
    estado.ListIndex = 0
    carga_salon
 
    '''09/08/2017 kenyo. Opcion Mesa Personal

    Dim mytable11 As New ADODB.Recordset

    mytable11.Open "select vemesa from parame where codigo='01'", cn, adOpenStatic, adLockOptimistic

    If mytable11.RecordCount > 0 Then
        If Trim("" & mytable11.Fields("vemesa")) = "S" Or Trim("" & mytable11.Fields("vemesa")) = "" Then
            vemesa.Visible = True
            ki8933.Visible = True
        Else
            vemesa.Visible = False
            ki8933.Visible = False

        End If

    End If

    mytable11.Close
    '''09/08/2017 kenyo. Opcion Mesa Personal

End Sub

Sub inicializa()
    modificaproducto = ""
    modificacompra = ""
    rw14 = ""
    v14 = ""
    conexionremota = ""
    escajero = ""
    esvendedor = ""
    mueveproducto = ""
    muevemesa = ""
    juntamesa = ""
    minireporte = ""
    terminal = ""
    tienda = ""
    productos = ""
    borra_comanda = ""
    cierre = ""
    despacho = ""
    descuento = ""
    precuenta = ""
    verificador = ""
    vreloj = ""
    inicia = ""
    local1 = "01"
    cprecios = ""
    dueno = ""
    ipss = ""
    tipo = "N"
    fechavaca = ""
    vecostoimp = ""
    apertura = ""
    anula = ""
    copia = ""
    congela = ""
    cuadre = ""
    pocket = ""
    'tipoca = ""
    'puertoca = ""
    reloj = ""
    caja = ""
    'tipopla = ""
    parame = ""
    parame1 = ""
    parame2 = ""
    parame3 = ""
    parame4 = ""

    seccion = ""
    valor1 = ""
    valor2 = ""
    valor3 = ""
    valor4 = ""
    por1 = ""
    por2 = ""
    por3 = ""
    por4 = ""
    ini1 = ""
    ini2 = ""
    ini3 = ""
    ini4 = ""
    fin1 = ""
    fin2 = ""
    fin3 = ""
    fin4 = ""

    rw1 = "W"
    rw2 = "W"
    rw3 = "W"
    rw4 = "W"
    rw5 = "W"
    rw6 = "W"
    rw7 = "W"
    rw8 = "W"
    rw9 = "W"
    rw10 = "W"
    rw11 = "W"
    rw13 = "W"

    v1 = "S"
    v2 = "S"
    v3 = "N"
    v4 = "N"
    v5 = "N"
    v6 = "N"
    v7 = "N"
    v8 = "N"
    v9 = "N"
    v10 = "S"
    v11 = "N"
    'v12 = "N"
    v13 = "N"
    vevend = ""
    veclave = ""

    clavere = ""
    clave = ""
    codigo1 = ""
    nombre = ""
    nombrec = ""
    contacto = ""
    direccion = ""
    dpto = ""
    distrito = ""
    zona = ""
    telefono = ""
    telefono1 = ""
    telefono2 = ""
    correo = ""

    regimen1.ListIndex = 0
    sexo.ListIndex = 0
    civil.ListIndex = 0
    cargo = ""
    profesion = ""
    fechaing = ""
    fechacese = ""
    moneda.ListIndex = 0
    basico = ""
    jornal = ""
    mtardanza = ""
    dtardanza = ""
    precuenta = ""

End Sub

Sub pone_registro()
    modificaproducto = "" & txempre.Fields("modificaproducto")
    modificacompra = "" & txempre.Fields("modificacompra")
    conexionremota = "" & txempre.Fields("conexionremota")

    escajero = "" & txempre.Fields("escajero")
    esvendedor = "" & txempre.Fields("esvendedor")

    juntamesa = "" & txempre.Fields("juntamesa")
    muevemesa = "" & txempre.Fields("muevemesa")
    mueveproducto = "" & txempre.Fields("mueveproducto")
    tienda = "" & txempre.Fields("tienda")
    terminal = "" & txempre.Fields("terminal")
    productos = "" & txempre.Fields("productos")
    minireporte = "" & txempre.Fields("minireporte")
    borra_comanda = "" & txempre.Fields("borra_comanda")
    cierre = "" & txempre.Fields("cierre")
    despacho = "" & txempre.Fields("despacho")
    descuento = "" & txempre.Fields("descuento")
    precuenta = "" & txempre.Fields("precuenta")
    verificador = "" & txempre.Fields("verificador")
    vreloj = "" & txempre.Fields("vreloj")
    inicia = "" & txempre.Fields("inicializa")
    local1 = "" & txempre.Fields("local")
    cprecios = "" & txempre.Fields("cprecios")
    dueno = "" & txempre.Fields("dueno")
    vecostoimp = "" & txempre.Fields("vecostoimp")
    tipo = "" & txempre.Fields("tipo")
    rw1 = "" & txempre.Fields("rw1")
    rw2 = "" & txempre.Fields("rw2")
    rw3 = "" & txempre.Fields("rw3")
    rw4 = "" & txempre.Fields("rw4")
    rw5 = "" & txempre.Fields("rw5")
    rw6 = "" & txempre.Fields("rw6")
    rw7 = "" & txempre.Fields("rw7")
    rw8 = "" & txempre.Fields("rw8")
    rw9 = "" & txempre.Fields("rw9")
    rw10 = "" & txempre.Fields("rw10")
    rw11 = "" & txempre.Fields("rw11")
    'rw12 = "" & txempre.Fields("rw12")
    rw13 = "" & txempre.Fields("rw13")
    rw14 = "" & txempre.Fields("rw14")

    v1 = "" & txempre.Fields("v1")
    v2 = "" & txempre.Fields("v2")
    v3 = "" & txempre.Fields("v3")
    v4 = "" & txempre.Fields("v4")
    v5 = "" & txempre.Fields("v5")
    v6 = "" & txempre.Fields("v6")
    v7 = "" & txempre.Fields("v7")
    v8 = "" & txempre.Fields("v8")
    v9 = "" & txempre.Fields("v9")
    v10 = "" & txempre.Fields("v10")
    v11 = "" & txempre.Fields("v11")
    'v12 = "" & txempre.Fields("v12")
    v13 = "" & txempre.Fields("v13")
    v14 = "" & txempre.Fields("v14")
    pocket = "" & txempre.Fields("pocket")
    vevend = "" & txempre.Fields("vevend")
    veclave = "" & txempre.Fields("veclave")
    clavere = "" & txempre.Fields("claverE")
    reloj = "" & txempre.Fields("reloj")
    caja = "" & txempre.Fields("caja")
    cuadre = "" & txempre.Fields("cuadre")
    anula = "" & txempre.Fields("anula")
    congela = "" & txempre.Fields("congela")
    copia = "" & txempre.Fields("copia")
    apertura = "" & txempre.Fields("apertura")

    seccion = "" & txempre.Fields("seccion")

    parame = "" & txempre.Fields("parame")
    parame1 = "" & txempre.Fields("parame1")
    parame2 = "" & txempre.Fields("parame2")
    parame3 = "" & txempre.Fields("parame3")
    parame4 = "" & txempre.Fields("parame4")

    ini1 = "" & txempre.Fields("ini1")
    ini2 = "" & txempre.Fields("ini2")
    ini3 = "" & txempre.Fields("ini3")
    ini4 = "" & txempre.Fields("ini4")
    fin1 = "" & txempre.Fields("fin1")
    fin2 = "" & txempre.Fields("fin2")
    fin3 = "" & txempre.Fields("fin3")
    fin4 = "" & txempre.Fields("fin4")
    por1 = "" & txempre.Fields("por1")
    por2 = "" & txempre.Fields("por2")
    por3 = "" & txempre.Fields("por3")
    por4 = "" & txempre.Fields("por4")
    valor1 = "" & txempre.Fields("valor1")
    valor2 = "" & txempre.Fields("valor2")
    valor3 = "" & txempre.Fields("valor3")
    valor4 = "" & txempre.Fields("valor4")

    clave = "" & txempre.Fields("clave")
    codigo = "" & txempre.Fields("codigo")
    codigo1 = "" & txempre.Fields("codigo1")
    nombre = "" & txempre.Fields("nombre")
    nombrec = "" & txempre.Fields("nombrec")
    contacto = "" & txempre.Fields("contacto")
    direccion = "" & txempre.Fields("direccion")
    dpto = "" & txempre.Fields("dpto")
    distrito = "" & txempre.Fields("distrito")
    zona = "" & txempre.Fields("zona")
    telefono = "" & txempre.Fields("telefono")
    telefono1 = "" & txempre.Fields("telefono1")
    telefono2 = "" & txempre.Fields("telefono2")
    correo = "" & txempre.Fields("correo")
    estado.ListIndex = 0

    If "" & txempre.Fields("estado") = "NOACTIVO" Then
        estado.ListIndex = 1

    End If

    regimen1.ListIndex = 0

    If "" & txempre.Fields("regimen") = "O" Then
        regimen1.ListIndex = 1

    End If

    If "" & txempre.Fields("regimen") = "X" Then
        regimen1.ListIndex = 2

    End If

    cargo = "" & txempre.Fields("cargo")
    profesion = "" & txempre.Fields("profesion")

    civil.ListIndex = 0

    If "" & txempre.Fields("civil") = "C" Then
        civil.ListIndex = 1

    End If

    sexo.ListIndex = 0

    If "" & txempre.Fields("sexo") = "F" Then
        sexo.ListIndex = 1

    End If

    fechaing = "" & txempre.Fields("fechaingr")
    fechacese = "" & txempre.Fields("fechacese")
    moneda.ListIndex = 0

    If "" & txempre.Fields("moneda") = "D" Then
        moneda.ListIndex = 1

    End If

    basico = "" & txempre.Fields("basico")
    jornal = "" & txempre.Fields("jornal")
    mtardanza = "" & txempre.Fields("mtardanza")
    dtardanza = "" & txempre.Fields("dtardanza")

End Sub

Sub grabando()

    Dim cad As String

    Dim buf As String

    If Len(rw1) = 0 Then
        rw1 = "N"

    End If

    If Len(rw13) = 0 Then
        rw13 = "N"

    End If

    If Len(rw2) = 0 Then
        rw2 = "N"

    End If

    If Len(rw3) = 0 Then
        rw3 = "N"

    End If

    If Len(rw4) = 0 Then
        rw4 = "N"

    End If

    If Len(rw5) = 0 Then
        rw5 = "N"

    End If

    If Len(rw6) = 0 Then
        rw6 = "N"

    End If

    If Len(rw7) = 0 Then
        rw7 = "N"

    End If

    If Len(rw8) = 0 Then
        rw8 = "N"

    End If

    If Len(rw9) = 0 Then
        rw9 = "N"

    End If

    If Len(rw10) = 0 Then
        rw10 = "N"

    End If

    If Len(rw11) = 0 Then
        rw11 = "N"

    End If

    'If Len(rw12) = 0 Then
    '   rw12 = "N"
    'End If
    If Len(rw13) = 0 Then
        rw13 = "N"

    End If

    If Len(rw14) = 0 Then
        rw14 = "N"

    End If

    If Len(v1) = 0 Then
        v1 = "N"

    End If

    If Len(v13) = 0 Then
        v13 = "N"

    End If

    If Len(v14) = 0 Then
        v14 = "N"

    End If

    If Len(v2) = 0 Then
        v2 = "N"

    End If

    If Len(v3) = 0 Then
        v3 = "N"

    End If

    If Len(v4) = 0 Then
        v4 = "N"

    End If

    If Len(v5) = 0 Then
        v5 = "N"

    End If

    If Len(v6) = 0 Then
        v6 = "N"

    End If

    If Len(v7) = 0 Then
        v7 = "N"

    End If

    If Len(v8) = 0 Then
        v8 = "N"

    End If

    If Len(v9) = 0 Then
        v9 = "N"

    End If

    If Len(v10) = 0 Then
        v10 = "N"

    End If

    If Len(v11) = 0 Then
        v11 = "N"

    End If

    'If Len(v12) = 0 Then
    '   v12 = "N"
    'End If
    txempre.Fields("modificaproducto") = Trim(modificaproducto)
    txempre.Fields("modificacompra") = Trim(modificacompra)
    txempre.Fields("conexionremota") = Trim(conexionremota)
    txempre.Fields("escajero") = Trim(escajero)
    txempre.Fields("esvendedor") = Trim(esvendedor)

    txempre.Fields("juntamesa") = Trim(juntamesa)
    txempre.Fields("muevemesa") = Trim(muevemesa)
    txempre.Fields("mueveproducto") = Trim(mueveproducto)

    txempre.Fields("minireporte") = Trim(minireporte)
    txempre.Fields("tienda") = Trim(tienda)
    txempre.Fields("terminal") = Trim(terminal)
    txempre.Fields("productos") = Trim(productos)
    txempre.Fields("borra_comanda") = Trim(borra_comanda)
    txempre.Fields("cierre") = Trim(cierre)
    txempre.Fields("despacho") = Trim(despacho)
    txempre.Fields("descuento") = Trim(descuento)
    txempre.Fields("precuenta") = Trim(precuenta)
    txempre.Fields("verificador") = Trim(verificador)
    txempre.Fields("vreloj") = Trim(vreloj)
    txempre.Fields("inicializa") = Trim(inicia)
    txempre.Fields("local") = Trim(local1)
    txempre.Fields("cprecios") = Trim(cprecios)
    txempre.Fields("dueno") = Trim(dueno)
    txempre.Fields("codigo") = Trim(codigo)
    txempre.Fields("tipo") = Trim(tipo)
    txempre.Fields("codigo1") = Trim(codigo1)
    txempre.Fields("nombre") = Trim(nombre)
    txempre.Fields("nombrec") = Trim(nombrec)
    txempre.Fields("contacto") = Trim(contacto)
    txempre.Fields("direccion") = Trim(direccion)
    txempre.Fields("dpto") = Trim(dpto)
    txempre.Fields("distrito") = Trim(distrito)
    txempre.Fields("zona") = Trim(zona)
    txempre.Fields("telefono") = Trim(telefono)
    txempre.Fields("telefono1") = Trim(telefono1)
    txempre.Fields("telefono2") = Trim(telefono2)
    txempre.Fields("correo") = Trim(correo)
    txempre.Fields("estado") = Trim(estado)
    txempre.Fields("regimen") = Trim(regimen)
    txempre.Fields("cargo") = Trim(cargo)
    txempre.Fields("profesion") = Trim(profesion)
    txempre.Fields("civil") = Trim(civil)
    txempre.Fields("sexo") = Trim(sexo)

    If IsDate(fechaing) Then
        txempre.Fields("fechaing") = (fechaing)

    End If

    If IsDate(fechavaca) Then
        txempre.Fields("fechavaca") = (fechavaca)

    End If

    If IsDate(fechacese) Then
        txempre.Fields("fechacese") = (fechacese)

    End If

    txempre.Fields("moneda") = Trim(moneda)
    txempre.Fields("basico") = Val(basico)
    txempre.Fields("jornal") = Val(jornal)
    txempre.Fields("ipss") = Trim(ipss)
    txempre.Fields("mtardanza") = Trim(mtardanza)
    txempre.Fields("dtardanza") = Val(dtardanza)
    txempre.Fields("clave") = Trim(clave)
    txempre.Fields("seccion") = Trim(seccion)
    txempre.Fields("ini1") = Val(ini1)
    txempre.Fields("ini2") = Val(ini2)
    txempre.Fields("ini3") = Val(ini3)
    txempre.Fields("ini4") = Val(ini4)

    txempre.Fields("fin1") = Val(fin1)
    txempre.Fields("fin2") = Val(fin2)
    txempre.Fields("fin3") = Val(fin3)
    txempre.Fields("fin4") = Val(fin4)

    txempre.Fields("por1") = Val(por1)
    txempre.Fields("por2") = Val(por2)
    txempre.Fields("por3") = Val(por3)
    txempre.Fields("por4") = Val(por4)
   
    txempre.Fields("valor1") = Val(valor1)
    txempre.Fields("valor2") = Val(valor2)
    txempre.Fields("valor3") = Val(valor3)
    txempre.Fields("valor4") = Val(valor4)
   
    txempre.Fields("reloj") = Trim(reloj)
    txempre.Fields("clavere") = Trim(clavere)
    txempre.Fields("clave") = Trim(clave)
    txempre.Fields("vevend") = Trim(vevend)
    txempre.Fields("veclave") = Trim(veclave)
    txempre.Fields("caja") = Trim(caja)
    txempre.Fields("parame") = Trim(parame)
    txempre.Fields("parame1") = Trim(parame1)
    txempre.Fields("parame2") = Trim(parame2)
    txempre.Fields("parame3") = Trim(parame3)
    txempre.Fields("parame4") = Trim(parame4)
   
    txempre.Fields("pocket") = Trim(pocket)
    txempre.Fields("cuadre") = Trim(cuadre)
    txempre.Fields("anula") = Trim(anula)
    txempre.Fields("copia") = Trim(copia)
    txempre.Fields("congela") = Trim(congela)
    txempre.Fields("apertura") = Trim(apertura)
    txempre.Fields("v1") = Trim(v1)
    txempre.Fields("v2") = Trim(v2)
    txempre.Fields("v3") = Trim(v3)
    txempre.Fields("v4") = Trim(v4)
    txempre.Fields("v5") = Trim(v5)
    txempre.Fields("v6") = Trim(v6)
    txempre.Fields("v7") = Trim(v7)
    txempre.Fields("v8") = Trim(v8)
    txempre.Fields("v9") = Trim(v9)
    txempre.Fields("v10") = Trim(v10)
    txempre.Fields("v11") = Trim(v11)
    'txempre.Fields("v12") = Trim(v12)
    txempre.Fields("v13") = Trim(v13)
    txempre.Fields("v14") = Trim(v14)

    txempre.Fields("rw1") = Trim(rw1)
    txempre.Fields("rw2") = Trim(rw2)
    txempre.Fields("rw3") = Trim(rw3)
    txempre.Fields("rw4") = Trim(rw4)
    txempre.Fields("rw5") = Trim(rw5)
    txempre.Fields("rw6") = Trim(rw6)
    txempre.Fields("rw7") = Trim(rw7)
    txempre.Fields("rw8") = Trim(rw8)
    txempre.Fields("rw9") = Trim(rw9)
    txempre.Fields("rw10") = Trim(rw10)
    txempre.Fields("rw11") = Trim(rw11)
    'txempre.Fields("rw12") = Trim(rw12)
    txempre.Fields("rw13") = Trim(rw13)
    txempre.Fields("rw14") = Trim(rw14)
    txempre.Fields("vecostoimp") = Trim(vecostoimp)
   
End Sub

Private Sub grba1_Click()

End Sub

Sub grabalocal1()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd2325_err

    mytablex.Open "select * from userlocal where codigo='" & codigo & "' and local='01'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        mytablex.Fields("codigo") = codigo
        mytablex.Fields("local") = "01"
        mytablex.Update

    End If

    mytablex.Close
    Exit Sub
cmd2325_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

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
        If Len(codigo) = 0 Then
            codigo.SetFocus
            Exit Function

        End If

        rbusca.Open "select codigo from vendedor where codigo='" & codigo & "'", cn, adOpenStatic, adLockOptimistic

        If rbusca.RecordCount > 0 Then
            rbusca.Close
            MsgBox "Ya existe codigo ", 48, "Aviso"
            Exit Function

        End If

        txempre.AddNew
        txempre.Fields("codigo") = codigo
        grabando
        grabalocal1
        txempre.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        txempre.Fields("codigo") = codigo
        grabando
        grabalocal1
        txempre.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    Dim found As Integer

    If Len(codigo) = 0 Then
        codigo.SetFocus
        Exit Function

    End If

    If tipo <> "N" And tipo <> "J" And tipo <> "I" And tipo <> "O" Then
        tipo.SetFocus
        Exit Function

    End If

    If Len(nombre) = 0 Then
        nombre.SetFocus
        Exit Function

    End If

    found = busca_local()

    If found = 0 Then
        MsgBox "No existe Local ", 48, "Aviso"
        local1.SetFocus
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
        ki8933.Enabled = True
            
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
        ki8933.Enabled = False
            
    End If
      
End Sub

Private Sub cargo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    profesion.SetFocus

End Sub

Private Sub cargo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        regimen1.SetFocus
        Exit Sub

    End If

End Sub

Private Sub civil_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    regimen1.SetFocus

End Sub

Private Sub civil_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        sexo.SetFocus
        Exit Sub

    End If

End Sub

Private Sub clave_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    clave = UCase(clave)

End Sub

Private Sub codigo1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    nombre.SetFocus

End Sub

Private Sub codigo1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If codigo.Enabled = True Then
            codigo.SetFocus

        End If

        Exit Sub

    End If

End Sub

Private Sub contacto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    direccion.SetFocus

End Sub

Private Sub contacto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        nombrec.SetFocus
        Exit Sub

    End If

End Sub

Private Sub correo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    sexo.SetFocus

End Sub

Private Sub correo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        telefono2.SetFocus
        Exit Sub

    End If

End Sub

Private Sub Dia_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 Then Exit Sub
    If Len(dia) <> 10 Then Exit Sub
    If Not IsDate(dia) Then
        dia = ""
        Exit Sub

    End If

    'limpia_sisper
    'found = busca_sisper()
    If found = 0 Then
        MsgBox "No hay ninguna Transaccion", 48, "Aviso"
        Exit Sub

    End If

    'eh1.SetFocus
End Sub

Private Sub direccion_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    dpto.SetFocus

End Sub

Private Sub direccion_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        contacto.SetFocus
        Exit Sub

    End If

End Sub

Private Sub distrito_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    zona.SetFocus

End Sub

Private Sub distrito_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        dpto.SetFocus
        Exit Sub

    End If

End Sub

Private Sub dpto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    distrito.SetFocus

End Sub

Private Sub dpto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        direccion.SetFocus
        Exit Sub

    End If

End Sub

Private Sub estado_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fechacese.SetFocus
        Exit Sub

    End If

End Sub

Private Sub fechacese_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    estado.SetFocus

End Sub

Private Sub fechacese_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        fechaing.SetFocus
        Exit Sub

    End If

End Sub

Private Sub fechaing_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    fechacese.SetFocus

End Sub

Private Sub fechaing_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        profesion.SetFocus
        Exit Sub

    End If

End Sub

Private Sub ki8933_Click()

    Dim buf As String

    On Error GoTo cmd00012_err

    Dim mytablex As New ADODB.Recordset

    buf = txempre.Fields("codigo")
     
    Frame6.Visible = True

    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open "select Salon,Mesa from Mesa where salon='" & extra_loquesea(psalon) & "'", cn, adOpenStatic, adLockOptimistic
    Set dbgrid11.DataSource = mytablex
    dbgrid11.columns(0).Width = 4000
    dbgrid11.columns(1).Width = 1000

    If mytablel.State = 1 Then mytablel.Close
    mytablel.Open "select Salon,Mesa,Codigo from usermesa where codigo='" & Trim("" & txempre.Fields("codigo")) & "'", cn, adOpenStatic, adLockOptimistic
    Set dbgrid12.DataSource = mytablel
    dbgrid12.columns(0).Width = 2000
    Exit Sub
cmd00012_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Label24_Click()
    clavex = clave

End Sub

Private Sub Label40_Click()

End Sub

Private Sub Label46_Click()

    Dim found As Integer

    'found = busca_registro()
    'If found = 0 Then
    '   MsgBox "No existe Codigo", 48, "Aviso"
    '   codigo.SetFocus
    '   Exit Sub
    'End If
    If busca_clave1("" & gusuario) <> "S" Then
        MsgBox "No tiene permiso", 48, "Aviso"
        Exit Sub

    End If

    Frame3.Visible = True
    clave.SetFocus

End Sub

Private Sub Label69_Click()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd2325_err

    buf = "" & dbgrid9.columns(1)
    mytablex.Open "select * from userlocal where codigo='" & codigo & "' and local='" & buf & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.AddNew
        mytablex.Fields("codigo") = codigo
        mytablex.Fields("local") = buf
        mytablex.Update

    End If

    mytablex.Close
    Label71_Click
    Exit Sub
cmd2325_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Label70_Click()

    On Error GoTo cmd190_err

    mytablel.Delete
    mytablel.MoveFirst
    Exit Sub
cmd190_err:
    Exit Sub

End Sub

Private Sub Label71_Click()

    Dim mytablex As New ADODB.Recordset

    If cprecios = "S" Then
        Frame5.Visible = True

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open "select Nombre,Codigo from tlocal", cn, adOpenStatic, adLockOptimistic
        Set dbgrid9.DataSource = mytablex
        dbgrid9.columns(0).Width = 4000
        dbgrid9.columns(1).Width = 1000

        If mytablel.State = 1 Then mytablel.Close
        mytablel.Open "select Local,Codigo from userlocal where codigo='" & codigo & "'", cn, adOpenStatic, adLockOptimistic
        Set dbgrid10.DataSource = mytablel
        dbgrid10.columns(0).Width = 2000
        Exit Sub

    End If

End Sub

Private Sub Label93_Click()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    On Error GoTo cmd42325_err

    mytabley.Open "select * from tlocal", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        mytabley.Close
        Exit Sub

    End If

    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from userlocal where codigo='" & codigo & "' and local='" & Trim("" & mytabley.Fields("codigo")) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablel.AddNew
            mytablel.Fields("codigo") = codigo
            mytablel.Fields("local") = Trim("" & mytabley.Fields("codigo"))
            mytablel.Update

        End If

        mytablex.Close
        mytabley.MoveNext
    Loop
    mytabley.Close
    Label71_Click
    Exit Sub
cmd42325_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Label95_Click()

    On Error GoTo cmd9190_err

    mytablel.Delete
    mytablel.MoveFirst
    Exit Sub
cmd9190_err:
    Exit Sub

End Sub

Private Sub Label96_Click()

    Dim buf      As String

    Dim buf1     As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd22325_err

    buf = "" & dbgrid11.columns(0)
    buf1 = "" & dbgrid11.columns(1)
    mytablex.Open "select * from usermesa where codigo='" & Trim("" & txempre.Fields("codigo")) & "' and salon='" & buf & "' and mesa='" & buf1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablel.AddNew
        mytablel.Fields("codigo") = Trim("" & txempre.Fields("codigo"))
        mytablel.Fields("salon") = buf
        mytablel.Fields("mesa") = buf1
        mytablel.Update

    End If

    mytablex.Close
    ki8933_Click
    Exit Sub
cmd22325_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Label99_Click()

End Sub

Private Sub nombre_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    nombrec.SetFocus

End Sub

Private Sub nombre_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        codigo1.SetFocus
        Exit Sub

    End If

End Sub

Private Sub nombrec_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    contacto.SetFocus

End Sub

Private Sub nombrec_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        nombre.SetFocus
        Exit Sub

    End If

End Sub

Private Sub profesion_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    fechaing.SetFocus

End Sub

Private Sub profesion_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        cargo.SetFocus
        Exit Sub

    End If

End Sub

Private Sub psalon_Click()

    If psalon = "%" Then Exit Sub
    ki8933_Click

End Sub

Private Sub regimen1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    cargo.SetFocus

End Sub

Private Sub regimen1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        civil.SetFocus
        Exit Sub

    End If

End Sub

Private Sub reloj_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_reloj

    End If

End Sub

Private Sub seccion_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_seccion

    End If

End Sub

Private Sub sexo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    civil.SetFocus

End Sub

Private Sub sexo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        correo.SetFocus
        Exit Sub

    End If

End Sub

Private Sub telefono_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    telefono1.SetFocus

End Sub

Private Sub telefono_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        zona.SetFocus
        Exit Sub

    End If

End Sub

Private Sub telefono1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    telefono2.SetFocus

End Sub

Private Sub telefono1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        telefono.SetFocus
        Exit Sub

    End If

End Sub

Private Sub telefono2_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    correo.SetFocus

End Sub

Private Sub telefono2_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        telefono1.SetFocus
        Exit Sub

    End If

End Sub

Private Sub zona_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    telefono.SetFocus

End Sub

Private Sub zona_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        distrito.SetFocus
        Exit Sub

    End If

End Sub

Function busca_clave1(buf As String) As String

    Dim rconsulta As New ADODB.Recordset

    Dim cad       As String

    cad = "select * from vendedor where codigo='" & buf & "'"

    If rconsulta.State = 1 Then rconsulta.Close
    rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

    If rconsulta.EOF = True And rconsulta.BOF = True Then
        Exit Function

    End If

    busca_clave1 = "" & rconsulta.Fields("veclave")

End Function

Function ejecutax(sw As Integer)

    Dim rconsulta As New ADODB.Recordset

    Dim cad       As String

    If opcion1 = "1" Then  'bodega
        If Len(bufferx) = 0 Then
            cad = "select Nombre,Codigo,Direccion,Telefono from vendedor "

        End If

        If Len(bufferx) > 0 Then
            cad = "SELECT Nombre,Codigo,Direccion,Telefono from vendedor   where  " & Combo1 & " like '" & bufferx & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            rconsulta.Close
            bufferx.SetFocus
            Exit Function

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 4000
        DBGrid2.columns(1).Width = 2000
        DBGrid2.columns(2).Width = 4000
        DBGrid2.columns(3).Width = 2000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

        Exit Function

    End If

    If opcion1 = "2" Then  'bodega
        If Len(bufferx) = 0 Then
            cad = "select Descripcio,Seccion from pseccion "

        End If

        If Len(bufferx) > 0 Then
            cad = "SELECT Descripcio,Seccion from pseccion   where  " & Combo1 & " like '" & bufferx & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            rconsulta.Close
            bufferx.SetFocus
            Exit Function

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 4000
        DBGrid2.columns(1).Width = 2000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

        Exit Function

    End If

    If opcion1 = "22" Then  'bodega
        If Len(bufferx) = 0 Then
            cad = "select Descripcio,tprohora from tprohora "

        End If

        If Len(bufferx) > 0 Then
            cad = "SELECT Descripcio,tprohora from tprohora   where  " & Combo1 & " like '" & bufferx & "%'"

        End If
   
        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            rconsulta.Close
            bufferx.SetFocus
            Exit Function

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 4000
        DBGrid2.columns(1).Width = 2000

        If sw = 1 Then
            DBGrid2.SetFocus

        End If

        Exit Function

    End If

    If opcion1 = "4" Then  'bodega
        If Len(bufferx) = 0 Then
            cad = "select Descripcio,tipopla from tipopla "

        End If

        If Len(bufferx) > 0 Then
            cad = "SELECT Descripcio,tipopla from tipopla   where  " & Combo1 & " like '" & bufferx & "%'"

        End If

        If rconsulta.State = 1 Then rconsulta.Close
        rconsulta.Open cad, cn, adOpenStatic, adLockOptimistic

        If rconsulta.EOF = True And rconsulta.BOF = True Then
            rconsulta.Close
            bufferx.SetFocus
            Exit Function

        End If

        Set DBGrid2.DataSource = rconsulta
        DBGrid2.columns(0).Width = 4000
        DBGrid2.columns(1).Width = 2000
   
        If sw = 1 Then
            DBGrid2.SetFocus

        End If

        Exit Function

    End If

End Function

Function busca_local()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select codigo from tlocal where codigo='" & local1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        busca_local = 1

    End If

    mytablex.Close

End Function

Sub carga_salon()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset
   
    mytablex.Open "select * from salon  ", cn, adOpenStatic, adLockOptimistic
    psalon.Clear
    psalon.AddItem "%"
    Do

        If mytablex.EOF Then Exit Do
        psalon.AddItem Trim("" & mytablex.Fields("Descripcio")) & "|" & Trim("" & mytablex.Fields("salon"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    psalon.ListIndex = 0

End Sub

