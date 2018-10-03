VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tmopla 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Tabla de Modelos Planilla"
   ClientHeight    =   10080
   ClientLeft      =   165
   ClientTop       =   -615
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   8655
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   13935
      Begin VB.CommandButton Command10 
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   6735
         Left            =   240
         TabIndex        =   38
         Top             =   1200
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Planilla del Personal"
      Height          =   9735
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   14535
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   12360
         Picture         =   "tmopla.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Imprimir todo"
         Top             =   2880
         Width           =   1470
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "<-Borra"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "<-Agrega"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "<-Borra"
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
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "<-Agrega"
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
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFC0&
         Caption         =   "<-Borra"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5880
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "<-Agrega"
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5520
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid dbgrid3 
         Height          =   2775
         Left            =   240
         TabIndex        =   26
         Top             =   2280
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4895
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Concepto"
            Caption         =   "Concepto"
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
            DataField       =   "Importe"
            Caption         =   "Importe"
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
         BeginProperty Column02 
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   3075.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   629.858
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dbgrid4 
         Height          =   2775
         Left            =   6240
         TabIndex        =   27
         Top             =   2280
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4895
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Concepto"
            Caption         =   "Concepto"
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
            DataField       =   "Importe"
            Caption         =   "Importe"
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
         BeginProperty Column02 
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   3075.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   629.858
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dbgrid5 
         Height          =   2655
         Left            =   240
         TabIndex        =   28
         Top             =   5520
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4683
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Concepto"
            Caption         =   "Concepto"
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
            DataField       =   "Importe"
            Caption         =   "Importe"
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
         BeginProperty Column02 
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   3075.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   629.858
            EndProperty
         EndProperty
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Modelo"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingresos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   1920
         Width           =   4935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aportes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   5160
         Width           =   4935
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descuentos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6240
         TabIndex        =   30
         Top             =   1920
         Width           =   4935
      End
      Begin VB.Label tipopla 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   29
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   12435
      TabIndex        =   7
      Top             =   0
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
         Picture         =   "tmopla.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   17
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         Picture         =   "tmopla.frx":1ADC
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "tmopla.frx":2CEE
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "tmopla.frx":3F00
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "tmopla.frx":5112
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Frame2"
      Height          =   6615
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   10575
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   8400
         Picture         =   "tmopla.frx":6324
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2160
         Width           =   1470
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   8400
         Picture         =   "tmopla.frx":6BEE
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Imprimir todo"
         Top             =   3120
         Width           =   1470
      End
      Begin VB.TextBox descripcio 
         BeginProperty Font 
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
         TabIndex        =   4
         Top             =   600
         Width           =   6015
      End
      Begin VB.TextBox modelo 
         BeginProperty Font 
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
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
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
         TabIndex        =   6
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Modelo"
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
         TabIndex        =   5
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   12495
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7935
         Left            =   120
         TabIndex        =   1
         Top             =   240
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
         ColumnCount     =   2
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
            DataField       =   "Modelo"
            Caption         =   "Modelo"
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
               ColumnWidth     =   5940.284
            EndProperty
            BeginProperty Column01 
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
   Begin VB.Menu dk833 
      Caption         =   "&VerModelo"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tmopla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txempre   As New ADODB.Recordset

Dim mytablexx As New ADODB.Recordset

Dim mytableyy As New ADODB.Recordset

Dim mytablezz As New ADODB.Recordset

Dim mytablexr As New ADODB.Recordset

Private Sub ajdu1_Click()

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Nuevo"
    cmdGuardar.Enabled = True
    habilita 1
    modelo.Enabled = True
    modelo = ""
    modelo.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub

    buf = txempre.Fields("modelo")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + txempre.Fields("modelo"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    txempre.Delete
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

Private Sub bufferx_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        Frame5.Visible = False
        Frame5.Enabled = False
        'If opcion1 = "22" Then
        '   fcodigo.SetFocus
        '   Exit Sub
        'End If
     
        Exit Sub

    End If

    Command10_Click

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

Sub ejecutax(sw As Integer)

    Dim cad As String

    On Error GoTo cmd89123_err

    'Command10.Visible = False
    If opcion1 = "1222" Then
        Command10.Visible = True

        'If Len(bufferx) > 0 Then
        '   If Len(bufferx) <> 6 Then Exit Sub
        '   If Val(Mid$(bufferx, 1, 2)) < 1 And Val(Mid$(bufferx, 1, 2)) > 12 Then
        '      Exit Sub
        '   End If
        '   If Val(Mid$(bufferx, 3, 4)) <= 0 Then
        '      Exit Sub
        '   End If
        'End If
   
    End If

    If opcion1 = "22" Or opcion1 = "23" Or opcion1 = "24" Then
        If Len(bufferx) = 0 Then
            cad = "SELECT Descripcio,Tipo from tplanico    "

        End If

        If Len(bufferx) > 0 Then
            cad = "SELECT Descripcio,Tipo from tplanico   where  " & Combo1 & " like '" & bufferx & "%'"

        End If

    End If

    If opcion1 = "233" Then
        If Len(bufferx) = 0 Then
            cad = "SELECT Descripcio,grupo from grupopla    "

        End If

        If Len(bufferx) > 0 Then
            cad = "SELECT Descripcio,grupo from grupopla   where  " & Combo1 & " like '" & bufferx & "%'"

        End If

    End If
   
    If mytablexr.State = 1 Then mytablexr.Close
    mytablexr.Open cad, cn, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = mytablexr
    DBGrid2.columns(0).Width = 4000
    DBGrid2.columns(1).Width = 2000

    If mytablexr.RecordCount > 0 Then
        DBGrid2.SetFocus

    End If

    Exit Sub
cmd89123_err:
    Exit Sub

End Sub

Private Sub Command10_Click()
    ejecutax 1

End Sub

Private Sub Command2_Click()
    dlo132_Click

End Sub

Private Sub Command3_Click()

    On Error GoTo cmd7811_err

    If MsgBox("Desea Borrar " + mytablexx.Fields("concepto"), 1, "Aviso") <> 1 Then Exit Sub
    mytablexx.Delete
    Exit Sub
cmd7811_err:
    MsgBox "Seleccione dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command4_Click()
    consulta_tipo

End Sub

Private Sub Command5_Click()

    On Error GoTo cmd78112_err

    If MsgBox("Desea Borrar " + mytableyy.Fields("concepto"), 1, "Aviso") <> 1 Then Exit Sub
    mytableyy.Delete
    Exit Sub
cmd78112_err:
    MsgBox "Seleccione dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command6_Click()
    consulta_tipo1

End Sub

Private Sub Command7_Click()

    On Error GoTo cmd17811_err

    If MsgBox("Desea Borrar " + mytablezz.Fields("concepto"), 1, "Aviso") <> 1 Then Exit Sub
    mytablezz.Delete
    Exit Sub
cmd17811_err:
    MsgBox "Seleccione dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Command8_Click()
    consulta_tipo2

End Sub

Private Sub Command9_Click()

End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    If KeyCode = 27 Then
        bufferx.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "22" Then
            found = existe_tipo1(Trim(DBGrid2.columns(1)), 0)

            If found = 1 Then
                MsgBox "Ya existe tipo ", 48, "Aviso"
                Exit Sub

            End If

            mytablexx.AddNew
            'mytablexx.Fields("seccion") = seccion
            mytablexx.Fields("tipopla") = tipopla
            mytablexx.Fields("modelo") = tipopla
            mytablexx.Fields("tipo") = Trim(DBGrid2.columns(1))
            mytablexx.Fields("concepto") = Trim(DBGrid2.columns(0))
            mytablexx.Update
            Frame5.Visible = False
            Frame5.Enabled = False

        End If

        If opcion1 = "23" Then
            found = existe_tipo1(Trim(DBGrid2.columns(1)), 1)

            If found = 1 Then
                MsgBox "Ya existe tipo ", 48, "Aviso"
                Exit Sub

            End If

            mytableyy.AddNew
            'mytableyy.Fields("seccion") = seccion
            mytableyy.Fields("tipopla") = tipopla
            mytableyy.Fields("modelo") = tipopla
            mytableyy.Fields("tipo") = Trim(DBGrid2.columns(1))
            mytableyy.Fields("concepto") = Trim(DBGrid2.columns(0))
            mytableyy.Update
            Frame5.Visible = False
            Frame5.Enabled = False

        End If

        If opcion1 = "233" Then

            'seccion = Trim(dbgrid2.Columns(1))
            'Frame5.Visible = False
            'Frame5.Enabled = False
            'If Len(seccion) > 0 Then
            '   crea_planilla
            'End If
        End If
      
        If opcion1 = "24" Then
            found = existe_tipo1(Trim(DBGrid2.columns(1)), 2)

            If found = 1 Then
                MsgBox "Ya existe tipo ", 48, "Aviso"
                Exit Sub

            End If

            mytablezz.AddNew
            'mytablezz.Fields("seccion") = seccion
            mytablezz.Fields("tipopla") = tipopla
            mytablezz.Fields("modelo") = tipopla
            mytablezz.Fields("tipo") = Trim(DBGrid2.columns(1))
            mytablezz.Fields("concepto") = Trim(DBGrid2.columns(0))
            mytablezz.Update
            Frame5.Visible = False
            Frame5.Enabled = False

        End If

    End If

End Sub

Private Sub dk833_Click()

    On Error GoTo cmd7812_err

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub

    tipopla = "" & txempre.Fields("modelo")

    Frame3.Visible = True
    'Command4.Enabled = True
    'Command3.Enabled = True
    'Command5.Enabled = True
    'Command6.Enabled = True
    'Command7.Enabled = True
    'Command8.Enabled = True

    dbgrid3.Enabled = True
    DBGrid4.Enabled = True
    dbgrid5.Enabled = True

    consulta_planilla
    Exit Sub
cmd7812_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Sub consulta_planilla()

    Dim buf As String

    If mytablexx.State = 1 Then mytablexx.Close
    mytablexx.Open "select * from remune00 where modelo='" & tipopla & "'", cn, adOpenStatic, adLockOptimistic
    Set dbgrid3.DataSource = mytablexx
   
    If mytableyy.State = 1 Then mytableyy.Close
    mytableyy.Open "select * from descue00 where modelo='" & tipopla & "'", cn, adOpenStatic, adLockOptimistic
    Set DBGrid4.DataSource = mytableyy
   
    If mytablezz.State = 1 Then mytablezz.Close
    mytablezz.Open "select * from aporta00 where modelo='" & tipopla & "'", cn, adOpenStatic, adLockOptimistic
    Set dbgrid5.DataSource = mytablezz
   
End Sub

Private Sub modelo_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(modelo) = 0 Then Exit Sub
    descripcio.SetFocus

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
            cad = "SELECT * from modelo    "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT *  from modelo   where  " & Combo1 & " like '" & buffer & "%'"

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

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'modelo = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'modelo.SetFocus
        'modelo_KeyPress 13
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

Private Sub djuer1_Click()

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "modelo"
    reporgen.Show 1

End Sub

Private Sub dlo132_Click()

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

    tmopla.Hide
    Unload tmopla

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub

    buf = txempre.Fields("modelo")

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
    modelo.Enabled = False
    descripcio.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub

    buf = txempre.Fields("modelo")

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
    modelo.Enabled = False
    descripcio.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    Command1_Click

End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "modelo"
    Combo1.ListIndex = 0

End Sub

Sub inicializa()

    descripcio = ""

End Sub

Sub pone_registro()
    modelo = Trim("" & txempre.Fields("modelo"))
    descripcio = Trim("" & txempre.Fields("descripcio"))

End Sub

Sub grabando()
    txempre.Fields("modelo") = Trim(modelo)
    txempre.Fields("descripcio") = Trim(descripcio)

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
        If Len(modelo) = 0 Then
            modelo.SetFocus
            Exit Function

        End If

        rbusca.Open "select modelo from modelo where modelo='" & modelo & "'", cn, adOpenStatic, adLockOptimistic

        If rbusca.RecordCount > 0 Then
            rbusca.Close
            MsgBox "Ya existe modelo ", 48, "Aviso"
            Exit Function

        End If

        txempre.AddNew
        txempre.Fields("modelo") = modelo
        grabando
        txempre.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        txempre.Fields("modelo") = modelo
        grabando
        txempre.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    'If Len(modelo) = 0 Then
    '   modelo.SetFocus
    '   Exit Function
    'End If
    If Len(descripcio) = 0 Then
        descripcio.SetFocus
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

Sub consulta_tipo()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.ListIndex = 0

    Frame5.Visible = True
    Frame5.Enabled = True
    bufferx = ""
    opcion1 = "22"
    ejecutax 1

End Sub

Sub consulta_tipo1()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.ListIndex = 0

    Frame5.Visible = True
    Frame5.Enabled = True
    bufferx = ""
    opcion1 = "23"
    ejecutax 1

End Sub

Sub consulta_tipo2()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.ListIndex = 0

    Frame5.Visible = True
    Frame5.Enabled = True
    bufferx = ""
    opcion1 = "24"
    ejecutax 1

End Sub

Function existe_tipo1(buf As String, sw As Integer)

    Dim mytablex As New ADODB.Recordset

    If sw = 0 Then
        mytablex.Open "select * from remune00 where modelo='" & tipopla & "'  and tipo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            existe_tipo1 = 1

        End If

        mytablex.Close

    End If

    If sw = 1 Then
        mytablex.Open "select * from descue00 where modelo='" & tipopla & "' and tipo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            existe_tipo1 = 1

        End If

        mytablex.Close

    End If

    If sw = 2 Then
        mytablex.Open "select * from aporta00 where modelo='" & tipopla & "'  and tipo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            existe_tipo1 = 1

        End If

        mytablex.Close

    End If

End Function

