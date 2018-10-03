VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tplagepe 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Planilla de Personal "
   ClientHeight    =   8715
   ClientLeft      =   165
   ClientTop       =   15
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   15345
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   13935
      Begin VB.CommandButton Command9 
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   600
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   6735
         Left            =   240
         TabIndex        =   39
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Planilla del Personal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   0
      TabIndex        =   9
      Top             =   1245
      Visible         =   0   'False
      Width           =   14535
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
         TabIndex        =   50
         Top             =   5520
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
         TabIndex        =   49
         Top             =   5880
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
         Top             =   2280
         Width           =   855
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
         TabIndex        =   45
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   12360
         Picture         =   "TPLAGEPE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Imprimir todo"
         Top             =   2880
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   12360
         Picture         =   "TPLAGEPE.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1920
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Recalculo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   14
         Top             =   6840
         Width           =   3255
      End
      Begin VB.TextBox horaextr 
         Height          =   375
         Left            =   7800
         MaxLength       =   11
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox horatraba 
         Height          =   375
         Left            =   5040
         MaxLength       =   11
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox diatraba 
         Height          =   375
         Left            =   1800
         MaxLength       =   11
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid dbgrid3 
         Height          =   2775
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         TabIndex        =   17
         Top             =   2280
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         TabIndex        =   18
         Top             =   5520
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4683
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
      Begin VB.Label basico 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   51
         Top             =   1440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TotalDscto"
         Height          =   375
         Left            =   6240
         TabIndex        =   44
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Haber Basico"
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   1440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label tipopla 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   40
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Horas Extras"
         Height          =   375
         Left            =   6240
         TabIndex        =   34
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Horas Trabajados"
         Height          =   375
         Left            =   3480
         TabIndex        =   33
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dias Trabajados"
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label totalcobrar 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7920
         TabIndex        =   31
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label totaporte 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7920
         TabIndex        =   30
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label totdscto 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7920
         TabIndex        =   29
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label totingreso 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7920
         TabIndex        =   28
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TotalCobrar"
         Height          =   375
         Left            =   6240
         TabIndex        =   27
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TotalAportes"
         Height          =   375
         Left            =   6240
         TabIndex        =   26
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Ingresos"
         Height          =   375
         Left            =   6240
         TabIndex        =   25
         Top             =   5160
         Width           =   1575
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   5160
         Width           =   4935
      End
      Begin VB.Label Label3 
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
         TabIndex        =   22
         Top             =   1920
         Width           =   4935
      End
      Begin VB.Label xcodigo 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label xnombre 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   7095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Planilla"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   12435
      TabIndex        =   2
      Top             =   0
      Width           =   12495
      Begin VB.ComboBox periodo 
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
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   480
         Width           =   2295
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   120
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
         Height          =   375
         Left            =   7560
         MaxLength       =   10
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
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
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TPLAGEPE.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   4
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
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "TPLAGEPE.frx":23A6
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Imprimir"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Periodo"
         Height          =   375
         Left            =   1440
         TabIndex        =   53
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Planilla"
         Height          =   375
         Left            =   1440
         TabIndex        =   42
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7935
         Left            =   0
         TabIndex        =   1
         Top             =   1320
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1935
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   3413
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
      EndProperty
   End
   Begin VB.Menu dlo2232 
      Caption         =   "&Planilla"
   End
   Begin VB.Menu fjh433 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
      Begin VB.Menu dkj2343 
         Caption         =   "&1.Reporte "
      End
      Begin VB.Menu dk82323 
         Caption         =   "&2.Planilla"
      End
      Begin VB.Menu dlo893 
         Caption         =   "3.Planilla Excel"
      End
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tplagepe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mytablexx As New ADODB.Recordset

Dim mytableyy As New ADODB.Recordset

Dim mytablezz As New ADODB.Recordset

Dim txempre   As New ADODB.Recordset

Private Sub ajdu1_Click()

End Sub

Private Sub bo712_Click()

End Sub

Private Sub buffer_DblClick()
    Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    Command1_Click

End Sub

Private Sub cmdAddEntry_Click()

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

    Command9_Click

End Sub

Private Sub cmdCerrar_Click()
    dlo132_Click

End Sub

Private Sub cmdDelete_Click()

End Sub

Private Sub cmdExit_Click()
    dlo132_Click

End Sub

Private Sub cmdHelp_Click()

End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub Command1_Click()
    Frame1.Visible = True
    Frame1.Enabled = True
    buffer = ""
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim cad As String

    If Len(buffer) = 0 Then
        cad = "SELECT * from vendedor    "

    End If

    If Len(buffer) > 0 Then
        cad = "SELECT *  from vendedor  where  " & Combo1 & " like '" & buffer & "%'"

    End If

    If txempre.State = 1 Then txempre.Close
    txempre.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = txempre
    dbGrid1.columns(0).Width = 4000
    dbGrid1.columns(1).Width = 2000

    If txempre.RecordCount > 0 Then
        dbGrid1.SetFocus

    End If

End Sub

Private Sub Command2_Click()
    suma_general1
    suma_general2
    suma_general3
    suma_total

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

    If tipopla = "%" Then Exit Sub
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

    If tipopla = "%" Then Exit Sub
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

    If tipopla = "%" Then Exit Sub
    consulta_tipo2

End Sub

Private Sub Command9_Click()
    ejecutax 1

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

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

Private Sub f8443_Click()

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
            mytablexx.Fields("periodo") = extra_loquesea(periodo)
            mytablexx.Fields("tipopla") = tipopla
            mytablexx.Fields("codigo") = xcodigo
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
            mytableyy.Fields("periodo") = extra_loquesea(periodo)
            mytableyy.Fields("tipopla") = tipopla
            mytableyy.Fields("codigo") = xcodigo
            mytableyy.Fields("tipo") = Trim(DBGrid2.columns(1))
            mytableyy.Fields("concepto") = Trim(DBGrid2.columns(0))
            mytableyy.Update
            Frame5.Visible = False
            Frame5.Enabled = False

        End If
      
        If opcion1 = "24" Then
            found = existe_tipo1(Trim(DBGrid2.columns(1)), 2)

            If found = 1 Then
                MsgBox "Ya existe tipo ", 48, "Aviso"
                Exit Sub

            End If

            mytablezz.AddNew
            mytablezz.Fields("periodo") = extra_loquesea(periodo)
            mytablezz.Fields("tipopla") = tipopla
            mytablezz.Fields("codigo") = xcodigo
            mytablezz.Fields("tipo") = Trim(DBGrid2.columns(1))
            mytablezz.Fields("concepto") = Trim(DBGrid2.columns(0))
            mytablezz.Update
            Frame5.Visible = False
            Frame5.Enabled = False

        End If

    End If

End Sub

Private Sub dbgrid3_BeforeColEdit(ByVal ColIndex As Integer, _
                                  ByVal KeyAscii As Integer, _
                                  Cancel As Integer)

    'MsgBox ColIndex
    If ColIndex <> 1 Then
        Cancel = True
        Exit Sub

    End If

End Sub

Private Sub dbgrid4_BeforeColEdit(ByVal ColIndex As Integer, _
                                  ByVal KeyAscii As Integer, _
                                  Cancel As Integer)

    If ColIndex <> 1 Then
        Cancel = True
        Exit Sub

    End If

End Sub

Private Sub dbgrid5_BeforeColEdit(ByVal ColIndex As Integer, _
                                  ByVal KeyAscii As Integer, _
                                  Cancel As Integer)

    If ColIndex <> 1 Then
        Cancel = True
        Exit Sub

    End If

End Sub

Sub cabecera_documento()

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
    buf = "Reporte de Planillas"
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Fecha   : " & Format(Now, "dd/mm/yyyy"), 25, 2, 0)
    found = formateaa("Periodo : " & periodo, 25, 2, 0)
    found = formateaa("Planilla: " & Combo3, 25, 2, 0)
    
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)
    found = formateaa("Codigo", 12, 0, 0)
    found = formateaa("Nombre", 31, 0, 0)
    found = formateaa("Ingreso ", 12, 0, 1)
    found = formateaa("Aporte ", 12, 0, 1)
    found = formateaa("Descuento ", 12, 0, 1)
    found = formateaa("Total ", 12, 2, 1)
    buf = String(150, "-")
    found = formateaa(buf, 150, 2, 0)

End Sub

Private Sub dk82323_Click()

    Dim found As Integer

    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

    If Combo3 = "%" Then
        MsgBox "Seleccione tipoplanilla ", 48, "Aviso"
        Exit Sub

    End If

    If periodo = "%" Then
        MsgBox "Seleccione periodo ", 48, "Aviso"
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    'cabecera_documento
    reporte_planilla
    '------------------------------------
    Close #1
    cerrar_archivo
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Private Sub dkj2343_Click()

    Dim found As Integer

    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

    If Combo3 = "%" Then
        MsgBox "Seleccione tipoplanilla ", 48, "Aviso"
        Exit Sub

    End If

    If periodo = "%" Then
        MsgBox "Seleccione periodo ", 48, "Aviso"
        Exit Sub

    End If

    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    cabecera_documento
    cuerpo_programa_documento
    '------------------------------------
    Close #1
    cerrar_archivo
     
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1

End Sub

Private Sub dlo132_Click()

    If Frame5.Visible = True Then
        bufferx_KeyPress 27
        Exit Sub

    End If

    If Frame2.Visible = True Then
        Frame2.Visible = False
        Exit Sub

    End If

    tplagepe.Hide
    Unload tplagepe

End Sub

Private Sub dlo2232_Click()

    On Error GoTo cmd7812_err

    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

    If Combo3 = "%" Then
        MsgBox "Definir tipo planilla ", 48, "Aviso"
        Exit Sub

    End If

    If periodo = "%" Then
        MsgBox "Definir Periodo planilla ", 48, "Aviso"
        Exit Sub

    End If

    tipopla = Combo3
    xcodigo = Trim("" & dbGrid1.columns(1))
    xnombre = Trim("" & dbGrid1.columns(0))
    'basico = Trim("" & dbGrid1.Columns(2))
    Frame2.Visible = True
    Command4.Enabled = True
    Command3.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True
    Command8.Enabled = True

    dbgrid3.Enabled = True
    DBGrid4.Enabled = True
    dbgrid5.Enabled = True

    consulta_planilla
    Exit Sub
cmd7812_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub dlo893_Click()

    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

    If Combo3 = "%" Then
        MsgBox "Seleccione tipoplanilla ", 48, "Aviso"
        Exit Sub

    End If

    If periodo = "%" Then
        MsgBox "Seleccione periodo ", 48, "Aviso"
        Exit Sub

    End If

    exporta_excel

End Sub

Private Sub fjh433_Click()

    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

    On Error GoTo cmd127812_err

    If Combo3 = "%" Then
        MsgBox "Definir tipo planilla ", 48, "Aviso"
        Exit Sub

    End If

    If periodo = "%" Then
        MsgBox "Definir Periodo planilla ", 48, "Aviso"
        Exit Sub

    End If

    tipopla = Combo3
    xcodigo = Trim("" & dbGrid1.columns(1))
    xnombre = Trim("" & dbGrid1.columns(0))
    'basico = Trim("" & dbGrid1.Columns(2))
    Frame2.Visible = True
    Command4.Enabled = False
    Command3.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command8.Enabled = False
    dbgrid3.Enabled = False
    DBGrid4.Enabled = False
    dbgrid5.Enabled = False

    consulta_planilla
    Exit Sub
cmd127812_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    Frame2.Top = 20: Frame2.Left = 20
    Frame5.Top = 20: Frame5.Left = 20
    Command1_Click

End Sub

Sub ejecutax(sw As Integer)

    Dim cad      As String

    Dim mytablex As New ADODB.Recordset

    If opcion1 = "22" Or opcion1 = "23" Or opcion1 = "24" Then
        If Len(bufferx) = 0 Then
            cad = "SELECT Descripcio,Tipo from tplanico    "

        End If

        If Len(bufferx) > 0 Then
            cad = "SELECT Descripcio,Tipo from tplanico   where  " & Combo1 & " like '" & bufferx & "%'"

        End If

    End If
   
    If mytablex.State = 1 Then mytablex.Close
    mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = mytablex
    DBGrid2.columns(0).Width = 4000
    DBGrid2.columns(1).Width = 2000

    If mytablex.RecordCount > 0 Then
        DBGrid2.SetFocus

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

Private Sub Form_Load()

    Dim mytablex As New ADODB.Recordset

    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.AddItem "Codigo"
    Combo1.ListIndex = 0

    Combo3.Clear
    Combo3.AddItem "%"
    Combo3.AddItem "NORMAL"
    Combo3.AddItem "EXTRA"
    Combo3.ListIndex = 0

    periodo.AddItem "%"
    mytablex.Open "select * from plaperiodo ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        periodo.AddItem Trim("" & mytablex.Fields("periodo")) & "|" & mytablex.Fields("descripcio")
        mytablex.MoveNext
    Loop
    mytablex.Close
    periodo.ListIndex = 0

End Sub

Sub inicializa()

End Sub

Private Sub grba1_Click()

End Sub

Sub consulta_planilla()

    Dim buf As String

    If mytablexx.State = 1 Then mytablexx.Close
    mytablexx.Open "select * from remune02 where tipopla='" & tipopla & "' and codigo='" & xcodigo & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic
    Set dbgrid3.DataSource = mytablexx
   
    If mytableyy.State = 1 Then mytableyy.Close
    mytableyy.Open "select * from descue02 where tipopla='" & tipopla & "' and codigo='" & xcodigo & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic
    Set DBGrid4.DataSource = mytableyy
   
    If mytablezz.State = 1 Then mytablezz.Close
    mytablezz.Open "select * from aporta02 where tipopla='" & tipopla & "' and codigo='" & xcodigo & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic
    Set dbgrid5.DataSource = mytablezz
   
    suma_general1
    suma_general2
    suma_general3
    suma_total
   
End Sub

Function existe_tipo1(buf As String, sw As Integer)

    Dim mytablex As New ADODB.Recordset

    If sw = 0 Then
        mytablex.Open "select * from remune02 where tipopla='" & tipopla & "' and codigo='" & xcodigo & "' and tipo='" & buf & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            existe_tipo1 = 1

        End If

        mytablex.Close

    End If

    If sw = 1 Then
        mytablex.Open "select * from descue02 where tipopla='" & tipopla & "' and codigo='" & xcodigo & "' and tipo='" & buf & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            existe_tipo1 = 1

        End If

        mytablex.Close

    End If

    If sw = 2 Then
        mytablex.Open "select * from aporta02 where tipopla='" & tipopla & "' and codigo='" & xcodigo & "' and tipo='" & buf & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            existe_tipo1 = 1

        End If

        mytablex.Close

    End If

End Function

Sub suma_general1()

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim v_importe As Double

    Dim mytablex  As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    'totingreso = ""
    If mytablexx.RecordCount = 0 Then Exit Sub
    mytablexx.MoveFirst
    sdx1 = 0
    Do

        If mytablexx.EOF Then Exit Do
        v_importe = 0

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open "select * from tplanico1 where tipo='" & "" & mytablexx.Fields("tipo") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            Do

                If mytablex.EOF Then Exit Do
                If mytabley.State = 1 Then mytabley.Close
                mytabley.Open "select * from remune02 where tipopla='" & "" & tipopla & "' and codigo='" & xcodigo & "' and tipo='" & mytablex.Fields("codigo") & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

                If mytabley.RecordCount > 0 Then
                    sdx = (Val("" & mytablex.Fields("porcentaje")) / 100#) * Val("" & mytabley.Fields("importe"))
                    v_importe = v_importe + sdx

                End If

                mytablex.MoveNext
            Loop

        End If

        If v_importe > 0 Then
            sdx1 = sdx1 + v_importe
            mytablex.Close
            mytablexx.Fields("importe") = v_importe
            mytablexx.Update

        End If

        mytablexx.MoveNext
    Loop

    'totingreso = "" & sdx1
End Sub

Sub suma_general2()

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim v_importe As Double

    Dim mytablex  As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    'totdscto = ""
    If mytableyy.RecordCount = 0 Then Exit Sub
    mytableyy.MoveFirst
    sdx1 = 0
    Do

        If mytableyy.EOF Then Exit Do
        v_importe = 0

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open "select * from tplanico1 where tipo='" & "" & mytableyy.Fields("tipo") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            Do

                If mytablex.EOF Then Exit Do
                If mytabley.State = 1 Then mytabley.Close
                mytabley.Open "select * from remune02 where tipopla='" & "" & tipopla & "' and codigo='" & xcodigo & "' and tipo='" & mytablex.Fields("codigo") & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

                If mytabley.RecordCount > 0 Then
                    sdx = (Val("" & mytablex.Fields("porcentaje")) / 100#) * Val("" & mytabley.Fields("importe"))
                    v_importe = v_importe + sdx

                End If

                mytablex.MoveNext
            Loop

        End If

        If v_importe > 0 Then
            sdx1 = sdx1 + v_importe
            mytablex.Close
            mytableyy.Fields("importe") = v_importe
            mytableyy.Update

        End If

        mytableyy.MoveNext
    Loop

    'totdscto = "" & sdx1
End Sub

Sub suma_general3()

    Dim sdx       As Double

    Dim sdx1      As Double

    Dim v_importe As Double

    Dim mytablex  As New ADODB.Recordset

    Dim mytabley  As New ADODB.Recordset

    'totaporte = ""
    If mytablezz.RecordCount = 0 Then Exit Sub
    mytablezz.MoveFirst
    sdx1 = 0
    Do

        If mytablezz.EOF Then Exit Do
        v_importe = 0

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open "select * from tplanico1 where tipo='" & "" & mytablezz.Fields("tipo") & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            Do

                If mytablex.EOF Then Exit Do
                If mytabley.State = 1 Then mytabley.Close
                mytabley.Open "select * from remune02 where tipopla='" & "" & tipopla & "' and codigo='" & xcodigo & "' and tipo='" & mytablex.Fields("codigo") & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

                If mytabley.RecordCount > 0 Then
                    sdx = (Val("" & mytablex.Fields("porcentaje")) / 100#) * Val("" & mytabley.Fields("importe"))
                    v_importe = v_importe + sdx

                End If

                mytablex.MoveNext
            Loop

        End If

        If v_importe > 0 Then
            sdx1 = sdx1 + v_importe
            mytablex.Close
            mytablezz.Fields("importe") = v_importe
            mytablezz.Update

        End If

        mytablezz.MoveNext
    Loop

    'totaporte = "" & sdx1
End Sub

Sub suma_total()

    Dim singreso As Double

    Dim sdscto   As Double

    Dim saporte  As Double

    singreso = 0

    If mytablexx.RecordCount > 0 Then
        mytablexx.MoveFirst
        Do

            If mytablexx.EOF Then Exit Do
            singreso = singreso + Val("" & mytablexx.Fields("importe"))
            mytablexx.MoveNext
        Loop

    End If

    sdscto = 0

    If mytableyy.RecordCount > 0 Then
        mytableyy.MoveFirst
        Do

            If mytableyy.EOF Then Exit Do
            sdscto = sdscto + Val("" & mytableyy.Fields("importe"))
            mytableyy.MoveNext
        Loop

    End If

    saporte = 0

    If mytablezz.RecordCount > 0 Then
        mytablezz.MoveFirst
        Do

            If mytablezz.EOF Then Exit Do
            saporte = saporte + Val("" & mytablezz.Fields("importe"))
            mytablezz.MoveNext
        Loop

    End If

    totingreso = "" & singreso
    totdscto = "" & sdscto
    totaporte = "" & saporte
    totalcobrar = Val(totingreso) - Val(totdscto)

End Sub

Sub cuerpo_programa_documento()

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim mytablez As New ADODB.Recordset

    Dim Tmp      As String

    Dim sw       As Integer

    Dim buf      As String

    Dim found    As Integer

    Dim sdx      As Double

    Dim aporta   As Double

    Dim remune   As Double

    Dim descue   As Double

    sdx = 0
    sw = 0
    suma1 = 0
    suma2 = 0
    suma3 = 0
    suma4 = 0
    ssuma1 = 0
    ssuma2 = 0
    ssuma3 = 0
    ssuma4 = 0

    If txempre.RecordCount = 0 Then Exit Sub
    txempre.MoveFirst
    Do

        If txempre.EOF Then Exit Do
   
        buf = "" & txempre.Fields("Codigo")
        found = formateaa(buf, 11, 0, 0)
        found = formateaa("", 1, 0, 0)
        buf = "" & txempre.Fields("nombre")
        found = formateaa(buf, 30, 0, 0)
        found = formateaa("", 1, 0, 0)
   
        aporta = 0
   
        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open "select importe from aporta02 where tipopla='" & Combo3 & "' and codigo='" & txempre.Fields("codigo") & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            Do

                If mytablex.EOF Then Exit Do
                aporta = aporta + Val("" & mytablex.Fields("importe"))
                mytablex.MoveNext
            Loop

        End If

        mytablex.Close
   
        descue = 0

        If mytabley.State = 1 Then mytabley.Close
        mytabley.Open "select importe from descue02 where tipopla='" & Combo3 & "' and codigo='" & txempre.Fields("codigo") & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

        If mytabley.RecordCount > 0 Then
            Do

                If mytabley.EOF Then Exit Do
                descue = descue + Val("" & mytabley.Fields("importe"))
                mytabley.MoveNext
            Loop

        End If

        mytabley.Close
   
        remune = 0

        If mytablez.State = 1 Then mytablez.Close
        mytablez.Open "select importe from remune02 where tipopla='" & Combo3 & "' and codigo='" & txempre.Fields("codigo") & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablez.RecordCount > 0 Then
            Do

                If mytablez.EOF Then Exit Do
                remune = remune + Val("" & mytablez.Fields("importe"))
                mytablez.MoveNext
            Loop

        End If

        mytablez.Close
   
        buf = "" & remune
        found = formateaa(buf, 11, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & aporta
        found = formateaa(buf, 11, 0, 1)
        found = formateaa("", 1, 0, 0)
        buf = "" & descue
        found = formateaa(buf, 11, 0, 1)
        found = formateaa("", 1, 0, 0)
        sdx = remune - descue
        buf = Format(sdx, "0.00")
        found = formateaa(buf, 11, 0, 1)
        found = formateaa("", 1, 2, 0)
   
        suma1 = suma1 + remune
        suma2 = suma2 + aporta
        suma3 = suma3 + descue
        suma4 = suma4 + sdx
   
        ssuma1 = ssuma1 + remune
        ssuma2 = ssuma2 + aporta
        ssuma3 = ssuma3 + descue
        ssuma4 = ssuma4 + sdx
   
        nlineas
        txempre.MoveNext
    Loop
    found = formateaa("", 43, 0, 0)
    buf = Format(suma1, "0.00")
    found = formateaa(buf, 11, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma2, "0.00")
    found = formateaa(buf, 11, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma3, "0.00")
    found = formateaa(buf, 11, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(suma4, "0.00")
    found = formateaa(buf, 11, 0, 1)
    found = formateaa("", 1, 2, 0)
    nlineas
   
    found = formateaa("Total-->   ", 43, 0, 1)
    buf = Format(ssuma1, "0.00")
    found = formateaa(buf, 11, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(ssuma2, "0.00")
    found = formateaa(buf, 11, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(ssuma3, "0.00")
    found = formateaa(buf, 11, 0, 1)
    found = formateaa("", 1, 0, 0)
    buf = Format(ssuma4, "0.00")
    found = formateaa(buf, 11, 0, 1)
    found = formateaa("", 1, 2, 0)
   
End Sub

Sub nlineas()
    contlin = contlin + 1

    If contlin > 50 Then
        cabecera_documento

    End If

End Sub

Sub reporte_planilla()

    Dim buf        As String

    Dim I          As Integer

    Dim j          As Integer

    Dim found      As Integer

    Dim mytabley   As New ADODB.Recordset

    Dim ynombre    As String

    Dim ycodigo    As String

    Dim ycargo     As String

    Dim ysspp      As String

    Dim yfechaingr As String

    Dim yfechavaca As String

    Dim yfechacese As String

    ReDim remune(20, 20) As String
    ReDim xcolum(4) As Integer

    Dim sdx       As Double

    Dim may       As Integer

    Dim contando  As Long

    Dim sdxremune As Double

    Dim sdxdescto As Double

    Dim sdxaporta As Double

    buf = "EMPRESA PRUEBA"
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("Direccion   : " & "Las mercedes xx", 40, 2, 0)
    found = formateaa("Ruc         : " & "2043333333", 40, 2, 0)
    found = formateaa("Reg.Patronal: " & "RP1212", 40, 2, 0)
    buf = String(150, "_")
    found = formateaa(buf, 90, 2, 0)
    'cabecera_tipico "", "", "" & "" & gusuario
    ynombre = ""
    ycodigo = ""
    ycargo = ""
    ysspp = ""
    yfechaingr = ""
    yfechavaca = ""
    yfechacese = ""
    
    If mytabley.State = 1 Then mytabley.Close
    mytabley.Open "select * from vendedor where codigo='" & txempre.Fields("codigo") & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        ynombre = "" & mytabley.Fields("nombre")
        ycodigo = "" & mytabley.Fields("codigo")
        ycargo = "" & mytabley.Fields("cargo")
        ysspp = "" & mytabley.Fields("ipss")
        yfechaingr = "" & mytabley.Fields("fechaingr")
        yfechavaca = "" & mytabley.Fields("fechavaca")
        yfechacese = "" & mytabley.Fields("fechacese")

    End If

    mytabley.Close
       
    found = formateaa("Nombres y Apellidos :" & ynombre, 60, 0, 0)
    found = formateaa(" ", 2, 0, 0)
    found = formateaa("Codigo:" & "" & ycodigo, 15, 2, 0)
    found = formateaa("Ocupacion           :" & ycargo, 60, 0, 0)
    found = formateaa(" ", 2, 0, 0)
    found = formateaa("Cod.SSP:" & "" & ysspp, 15, 2, 0)
    found = formateaa("Fecha Ingreso       :" & yfechaingr, 30, 0, 0)
    found = formateaa(" ", 2, 0, 0)
    found = formateaa("Fecha Vaca  :" & yfechavaca, 30, 0, 0)
    found = formateaa(" ", 2, 0, 0)
    found = formateaa("Fecha Cese  :" & yfechacese, 30, 0, 0)
    found = formateaa(" ", 2, 2, 0)
    buf = String(150, "_")
    found = formateaa(buf, 90, 2, 0)
    found = formateaa("REMUNERACIONES ", 29, 0, 0)
    found = formateaa("|", 1, 0, 0)
    found = formateaa("APORT. Y DESCTOS TRABAJ.", 29, 0, 0)
    found = formateaa("|", 1, 0, 0)
    found = formateaa("APORTACIONES PATRONALES ", 29, 0, 0)
    found = formateaa("|", 1, 2, 0)
    buf = String(150, "_")
    found = formateaa(buf, 90, 2, 0)

    'ahora hay que imprimir los 3 archivos
    'en tres columnas
    For I = 1 To 20
        For j = 1 To 20
            remune(I, j) = ""
        Next j
    Next I

    I = 0
    sdxremune = 0

    If mytabley.State = 1 Then mytabley.Close
    mytabley.Open "select * from remune02 where tipopla='" & Combo3 & "' and codigo='" & txempre.Fields("codigo") & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
   
        Do

            If mytabley.EOF Then Exit Do
            I = I + 1
            remune(I, 1) = "" & mytabley.Fields("tipo")
            remune(I, 2) = "" & mytabley.Fields("concepto")
            remune(I, 3) = "" & mytabley.Fields("importe")
            sdxremune = sdxremune + Val("" & mytabley.Fields("importe"))
            mytabley.MoveNext
        Loop

    End If

    mytabley.Close
    
    xcolum(1) = I
   
    I = 0
    sdxdescto = 0

    If mytabley.State = 1 Then mytabley.Close
    mytabley.Open "select * from descue02 where tipopla='" & Combo3 & "' and codigo='" & txempre.Fields("codigo") & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        Do

            If mytabley.EOF Then Exit Do
            I = I + 1
            remune(I, 4) = "" & mytabley.Fields("tipo")
            remune(I, 5) = "" & mytabley.Fields("concepto")
            remune(I, 6) = "" & mytabley.Fields("importe")
            sdxdescto = sdxdescto + Val("" & mytabley.Fields("importe"))
            mytabley.MoveNext
        Loop

    End If

    mytabley.Close
    
    xcolum(2) = I
   
    I = 0
    sdxaporta = 0

    If mytabley.State = 1 Then mytabley.Close
    mytabley.Open "select * from aporta02 where tipopla='" & Combo3 & "' and codigo='" & txempre.Fields("codigo") & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
   
        Do

            If mytabley.EOF Then Exit Do
            I = I + 1
            remune(I, 7) = "" & mytabley.Fields("tipo")
            remune(I, 8) = "" & mytabley.Fields("concepto")
            remune(I, 9) = "" & mytabley.Fields("importe")
            sdxaporta = sdxaporta + Val("" & mytabley.Fields("importe"))
            mytabley.MoveNext
        Loop

    End If

    mytabley.Close
    
    xcolum(3) = I
    '--------- imprimiendo
    may = xcolum(1)

    If xcolum(1) > xcolum(2) And xcolum(1) > xcolum(3) Then
        may = xcolum(1)

    End If

    If xcolum(2) > xcolum(1) And xcolum(2) > xcolum(3) Then
        may = xcolum(2)

    End If

    If xcolum(3) > xcolum(1) And xcolum(3) > xcolum(2) Then
        may = xcolum(3)

    End If

    'impresiones detalles ---------------------------------------
    contando = 0

    For I = 1 To may
        'Open Filename For Append As #2
        found = formateaa(remune(I, 1), 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa(remune(I, 2), 15, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa(Format(Val(remune(I, 3)), "0.00"), 8, 0, 1)
        found = formateaa("", 2, 0, 0)

        found = formateaa(remune(I, 4), 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa(remune(I, 5), 15, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa(Format(Val(remune(I, 6)), "0.00"), 8, 0, 1)
        found = formateaa("", 2, 0, 0)

        found = formateaa(remune(I, 7), 3, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa(remune(I, 8), 15, 0, 0)
        found = formateaa("", 1, 0, 0)
        found = formateaa(Format(Val(remune(I, 9)), "0.00"), 8, 0, 1)
        found = formateaa("", 1, 2, 0)
        'Close #2
        contando = contando + 1
    Next I
        
    buf = String(150, "_")
    found = formateaa(buf, 90, 2, 0)
             
    found = formateaa("", 4, 0, 0)
    found = formateaa("TOTAL REMUN.", 16, 0, 0)
    found = formateaa(Format("" & sdxremune, "0.00"), 8, 0, 1)
    found = formateaa("", 2, 0, 0)
            
    found = formateaa("", 4, 0, 0)
    found = formateaa("TOTAL DESCTOS.", 16, 0, 0)
    found = formateaa(Format("" & sdxdescto, "0.00"), 8, 0, 1)
    found = formateaa("", 2, 0, 0)
            
    found = formateaa("", 4, 0, 0)
    found = formateaa("TOTAL APORTE EMP.", 16, 0, 0)
    found = formateaa(Format("" & sdxaporta, "0.00"), 8, 0, 1)
    found = formateaa("", 2, 2, 0)
            
    buf = String(150, "_")
    found = formateaa(buf, 90, 2, 0)
        
    found = formateaa("", 4, 0, 0)
    found = formateaa("NETO PAGAR.S/.", 16, 0, 0)
    sdx = sdxremune - sdxdescto
    buf = Format(sdx, "0.00")
    found = formateaa(buf, 8, 0, 1)
    found = formateaa("", 2, 2, 0)
            
    found = formateaa("", 1, 2, 0)
    found = formateaa("", 1, 2, 0)
            
    found = formateaa("", 4, 0, 0)
    found = formateaa("__________________", 16, 0, 0)
    found = formateaa("", 20, 0, 0)
    found = formateaa("__________________", 16, 2, 0)
            
    found = formateaa("", 4, 0, 0)
    found = formateaa("Empleador", 16, 0, 0)
    found = formateaa("", 20, 0, 0)
    found = formateaa("Recibi Conforme", 16, 2, 0)
            
    found = formateaa("", 1, 2, 0)
    found = formateaa("", 1, 2, 0)

End Sub

Sub exporta_excel()

    Dim buf        As String

    Dim found      As Integer

    Dim mytabley   As New ADODB.Recordset

    Dim ynombre    As String

    Dim ycodigo    As String

    Dim ycargo     As String

    Dim ysspp      As String

    Dim yfechaingr As String

    Dim yfechavaca As String

    Dim yfechacese As String

    ReDim remune(20, 20) As String
    ReDim xcolum(4) As Integer

    Dim sdx         As Double

    Dim may         As Integer

    Dim contando    As Long

    Dim sdxremune   As Double

    Dim sdxdescto   As Double

    Dim sdxaporta   As Double

    Dim v           As Long

    Dim h           As Long

    Dim I           As Long

    Dim j           As Long

    Dim sdx1        As Double

    Dim sdx2        As Double

    Dim mytablex    As New ADODB.Recordset

    Dim strSQL      As String

    Dim Heading(15) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion
      
    ynombre = ""
    ycodigo = ""
    ycargo = ""
    ysspp = ""
    yfechaingr = ""
    yfechavaca = ""
    yfechacese = ""
    
    If mytabley.State = 1 Then mytabley.Close
    mytabley.Open "select * from vendedor where codigo='" & txempre.Fields("codigo") & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        mytabley.Close
        Exit Sub

    End If

    ynombre = "" & mytabley.Fields("nombre")
    ycodigo = "" & mytabley.Fields("codigo")
    ycargo = "" & mytabley.Fields("cargo")
    ysspp = "" & mytabley.Fields("ipss")
    yfechaingr = "" & mytabley.Fields("fechaingr")
    yfechavaca = "" & mytabley.Fields("fechavaca")
    yfechacese = "" & mytabley.Fields("fechacese")
    mytabley.Close
    
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel

    With objExcel.ActiveSheet
        .columns("A").ColumnWidth = 3
        .columns("B").ColumnWidth = 20
        .columns("C").ColumnWidth = 30
        .columns("D").ColumnWidth = 30
        .columns("D").ColumnWidth = 30
        
    End With
       
    objExcel.ActiveSheet.Cells(2, 2) = "EMPRESA "
    objExcel.ActiveSheet.Cells(2, 3) = "AXES PERU "
    objExcel.ActiveSheet.Cells(3, 2) = "RUC "
    objExcel.ActiveSheet.Cells(4, 2) = "REG. PATRONAL"
    
    objExcel.ActiveSheet.Cells(5, 2) = "APELLIDOS Y NOMBRES"
    objExcel.ActiveSheet.Cells(5, 3) = ynombre
    objExcel.ActiveSheet.Cells(5, 4) = "CODIGO:" & ycodigo
    objExcel.ActiveSheet.Cells(5, 5) = "CARGO:" & ycargo
    
    objExcel.ActiveSheet.Cells(6, 2) = "COD.SSP:" & ysspp
    objExcel.ActiveSheet.Cells(6, 3) = "" & ysspp
    
    objExcel.ActiveSheet.Cells(7, 2) = "FECHA INGRESO:"
    objExcel.ActiveSheet.Cells(7, 3) = yfechaingr
    objExcel.ActiveSheet.Cells(7, 4) = "PERIODO:"
    objExcel.ActiveSheet.Cells(7, 5) = "" & periodo
    
    objExcel.ActiveSheet.Cells(8, 1) = "TIP"
    objExcel.ActiveSheet.Cells(8, 2) = "CONCEPTO"
    objExcel.ActiveSheet.Cells(8, 3) = "REMUNERACIONES"
    objExcel.ActiveSheet.Cells(8, 4) = "APORT. Y DSCTOS TRABAJ."
    objExcel.ActiveSheet.Cells(8, 5) = "APORT.PATRONALES"
    
    For I = 1 To 20
        For j = 1 To 20
            remune(I, j) = ""
        Next j
    Next I

    I = 0
    sdxremune = 0

    If mytabley.State = 1 Then mytabley.Close
    mytabley.Open "select * from remune02 where tipopla='" & Combo3 & "' and codigo='" & txempre.Fields("codigo") & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
   
        Do

            If mytabley.EOF Then Exit Do
            I = I + 1
            remune(I, 1) = "" & mytabley.Fields("tipo")
            remune(I, 2) = "" & mytabley.Fields("concepto")
            remune(I, 3) = "" & mytabley.Fields("importe")
            sdxremune = sdxremune + Val("" & mytabley.Fields("importe"))
            mytabley.MoveNext
        Loop

    End If

    mytabley.Close
    
    xcolum(1) = I
   
    I = 0
    sdxdescto = 0

    If mytabley.State = 1 Then mytabley.Close
    mytabley.Open "select * from descue02 where tipopla='" & Combo3 & "' and codigo='" & txempre.Fields("codigo") & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        Do

            If mytabley.EOF Then Exit Do
            I = I + 1
            remune(I, 4) = "" & mytabley.Fields("tipo")
            remune(I, 5) = "" & mytabley.Fields("concepto")
            remune(I, 6) = "" & mytabley.Fields("importe")
            sdxdescto = sdxdescto + Val("" & mytabley.Fields("importe"))
            mytabley.MoveNext
        Loop

    End If

    mytabley.Close
    
    xcolum(2) = I
   
    I = 0
    sdxaporta = 0

    If mytabley.State = 1 Then mytabley.Close
    mytabley.Open "select * from aporta02 where tipopla='" & Combo3 & "' and codigo='" & txempre.Fields("codigo") & "' and periodo='" & extra_loquesea(periodo) & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
   
        Do

            If mytabley.EOF Then Exit Do
            I = I + 1
            remune(I, 7) = "" & mytabley.Fields("tipo")
            remune(I, 8) = "" & mytabley.Fields("concepto")
            remune(I, 9) = "" & mytabley.Fields("importe")
            sdxaporta = sdxaporta + Val("" & mytabley.Fields("importe"))
            mytabley.MoveNext
        Loop

    End If

    mytabley.Close
    v = 9
    xcolum(3) = I
    '--------- imprimiendo
    may = xcolum(1)

    If xcolum(1) > xcolum(2) And xcolum(1) > xcolum(3) Then
        may = xcolum(1)

    End If

    If xcolum(2) > xcolum(1) And xcolum(2) > xcolum(3) Then
        may = xcolum(2)

    End If

    If xcolum(3) > xcolum(1) And xcolum(3) > xcolum(2) Then
        may = xcolum(3)

    End If

    'impresiones detalles ---------------------------------------
    contando = 0

    For I = 1 To may
        'Open Filename For Append As #2
            
        objExcel.ActiveSheet.Cells(v, 1) = "'" & remune(I, 1)
        objExcel.ActiveSheet.Cells(v, 2) = "'" & remune(I, 2)
        objExcel.ActiveSheet.Cells(v, 3) = Format(Val(remune(I, 3)), "0.00")
        v = v + 1
        objExcel.ActiveSheet.Cells(v, 1) = "'" & remune(I, 4)
        objExcel.ActiveSheet.Cells(v, 2) = "'" & remune(I, 5)
        objExcel.ActiveSheet.Cells(v, 4) = Format(Val(remune(I, 6)), "0.00")
        v = v + 1

        If Val(remune(I, 9)) > 0 Then
            objExcel.ActiveSheet.Cells(v, 1) = "'" & remune(I, 7)
            objExcel.ActiveSheet.Cells(v, 2) = "'" & remune(I, 8)
            objExcel.ActiveSheet.Cells(v, 5) = Format(Val(remune(I, 9)), "0.00")
            v = v + 1

        End If

        contando = contando + 1
    Next I

    v = v + 1
    objExcel.ActiveSheet.Cells(v, 2) = "TOTALES"
    objExcel.ActiveSheet.Cells(v, 3) = Format(sdxremune, "0.00")
    objExcel.ActiveSheet.Cells(v, 4) = Format(sdxdescto, "0.00")
    objExcel.ActiveSheet.Cells(v, 5) = Format(sdxaporta, "0.00")
        
    v = v + 1
    sdx = sdxremune - sdxdescto
    objExcel.ActiveSheet.Cells(v, 2) = "NETO PAGAR"
    objExcel.ActiveSheet.Cells(v, 3) = Format(sdx, "0.00")

    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
 
End Sub

