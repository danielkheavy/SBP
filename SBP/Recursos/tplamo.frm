VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tplamo 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Generando Planilla x Empleado"
   ClientHeight    =   9420
   ClientLeft      =   165
   ClientTop       =   15
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   14610
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
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
      Height          =   8655
      Left            =   -15
      TabIndex        =   50
      Top             =   690
      Visible         =   0   'False
      Width           =   13935
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
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   600
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
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
      End
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
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H8000000D&
         Caption         =   "Ver en Excell"
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
         Left            =   9360
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   8040
         Width           =   3015
      End
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   6735
         Left            =   240
         TabIndex        =   55
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
      Height          =   4095
      Left            =   1890
      TabIndex        =   44
      Top             =   1395
      Visible         =   0   'False
      Width           =   9855
      Begin VB.OptionButton horas2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hora Entrada Salida"
         Height          =   375
         Left            =   1680
         TabIndex        =   57
         Top             =   1560
         Width           =   2295
      End
      Begin VB.OptionButton horas1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Horas Trabajadas"
         Height          =   375
         Left            =   1680
         TabIndex        =   56
         Top             =   1200
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procesar"
         Height          =   375
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   480
         Width           =   1215
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox fechaf 
         BeginProperty Font 
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
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaInicio"
         Height          =   375
         Left            =   360
         TabIndex        =   48
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaFinal"
         Height          =   375
         Left            =   360
         TabIndex        =   47
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Planilla del Personal"
      Height          =   8940
      Left            =   -15
      TabIndex        =   5
      Top             =   690
      Visible         =   0   'False
      Width           =   14535
      Begin VB.CommandButton Command13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&CargaPlantilla"
         Height          =   975
         Left            =   12360
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   960
         Width           =   1470
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Borrar"
         Height          =   975
         Left            =   12360
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1920
         Width           =   1470
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   41
         Top             =   5520
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   40
         Top             =   5880
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   39
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   38
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   37
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   36
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   12360
         Picture         =   "tplamo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprimir todo"
         Top             =   2880
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   12360
         Picture         =   "tplamo.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3840
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
         TabIndex        =   10
         Top             =   6840
         Width           =   3255
      End
      Begin VB.TextBox horaextr 
         Height          =   375
         Left            =   7800
         MaxLength       =   11
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox horatraba 
         Height          =   375
         Left            =   5040
         MaxLength       =   11
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox diatraba 
         Height          =   375
         Left            =   5040
         MaxLength       =   11
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid dbgrid3 
         Height          =   2775
         Left            =   240
         TabIndex        =   6
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   42
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
         TabIndex        =   35
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Haber Basico"
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label tipopla 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   31
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Horas Extras"
         Height          =   375
         Left            =   6240
         TabIndex        =   30
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
         TabIndex        =   29
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dias Trabajados"
         Height          =   375
         Left            =   3480
         TabIndex        =   28
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label totalcobrar 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7920
         TabIndex        =   27
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label totaporte 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7920
         TabIndex        =   26
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label totdscto 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7920
         TabIndex        =   25
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label totingreso 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   7920
         TabIndex        =   24
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TotalCobrar"
         Height          =   375
         Left            =   6240
         TabIndex        =   23
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TotalAportes"
         Height          =   375
         Left            =   6240
         TabIndex        =   22
         Top             =   5880
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Ingresos"
         Height          =   375
         Left            =   6240
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   1920
         Width           =   4935
      End
      Begin VB.Label xcodigo 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label xnombre 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   240
         Width           =   7095
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Planilla"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   12435
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   120
         Width           =   2535
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tplamo.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Planilla"
         Height          =   375
         Left            =   1440
         TabIndex        =   33
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid dbGrid1 
      Height          =   8295
      Left            =   0
      TabIndex        =   43
      Top             =   720
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   14631
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
   Begin VB.Menu dlo2232 
      Caption         =   "&Planilla"
   End
   Begin VB.Menu fjh433 
      Caption         =   "&Zoom"
   End
   Begin VB.Menu djuer1 
      Caption         =   "&Reporte"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu fkl9933 
      Caption         =   "&Reloj"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu fdl8923 
         Caption         =   "&0.Reporte Consolidado Dias"
      End
      Begin VB.Menu k8844 
         Caption         =   "&1.Visualizar Entradas Salidas "
      End
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tplamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mytablexx As New ADODB.Recordset

Dim mytableyy As New ADODB.Recordset

Dim mytablezz As New ADODB.Recordset

Dim mytablexr As New ADODB.Recordset

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
    'Frame1.Visible = True
    'Frame1.Enabled = True
    'buffer = ""
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

    cad = cad & " order by nombre"

    If txempre.State = 1 Then txempre.Close
    txempre.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = txempre
    dbGrid1.columns(0).Width = 4000
    dbGrid1.columns(1).Width = 2000

    If txempre.RecordCount > 0 Then
        dbGrid1.SetFocus

    End If

End Sub

Private Sub Command10_Click()
    imprimir_exell

End Sub

Private Sub Command11_Click()

    If Not IsDate(fechai) Then Exit Sub
    If Not IsDate(fechaf) Then Exit Sub
    If Frame1.Caption = "Consolidado" Then
        imprime_excell2

    End If

    If Frame1.Caption = "Entradas Salidas" Then
        consulta_reloj

    End If

End Sub

Private Sub Command12_Click()

    If mytablexx.State = 1 Then mytablexx.Close
    If mytableyy.State = 1 Then mytableyy.Close
    If mytablezz.State = 1 Then mytablezz.Close

    cn.Execute ("delete from remune01 where tipopla='" & tipopla & "' and codigo='" & xcodigo & "'")
    cn.Execute ("delete from descue01 where tipopla='" & tipopla & "' and codigo='" & xcodigo & "'")
    cn.Execute ("delete from aporta01 where tipopla='" & tipopla & "' and codigo='" & xcodigo & "'")
    consulta_planilla

End Sub

Private Sub Command13_Click()

    Dim mytablex As New ADODB.Recordset

    Command12_Click
    mytablex.Open "select * from remune00 where modelo='1'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            mytablexx.AddNew
            mytablexx.Fields("tipopla") = tipopla
            mytablexx.Fields("codigo") = xcodigo
            mytablexx.Fields("tipo") = Trim("" & mytablex.Fields("tipo"))
            mytablexx.Fields("concepto") = Trim("" & mytablex.Fields("concepto"))
            mytablexx.Update
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    mytablex.Open "select * from descue00 where modelo='1'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            mytableyy.AddNew
            mytableyy.Fields("tipopla") = tipopla
            mytableyy.Fields("codigo") = xcodigo
            mytableyy.Fields("tipo") = Trim("" & mytablex.Fields("tipo"))
            mytableyy.Fields("concepto") = Trim("" & mytablex.Fields("concepto"))
            mytableyy.Update
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

    mytablex.Open "select * from aporta00 where modelo='1'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        Do

            If mytablex.EOF Then Exit Do
            mytablezz.AddNew
            mytablezz.Fields("tipopla") = tipopla
            mytablezz.Fields("codigo") = xcodigo
            mytablezz.Fields("tipo") = Trim("" & mytablex.Fields("tipo"))
            mytablezz.Fields("concepto") = Trim("" & mytablex.Fields("concepto"))
            mytablezz.Update
            mytablex.MoveNext
        Loop

    End If

    mytablex.Close

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
            'mytablexx.Fields("seccion") = seccion
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
            'mytableyy.Fields("seccion") = seccion
            mytableyy.Fields("tipopla") = tipopla
            mytableyy.Fields("codigo") = xcodigo
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

Private Sub djuer1_Click()

    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

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

    If Frame1.Visible = True Then
        Frame1.Visible = False
        Exit Sub

    End If

    tplamo.Hide
    Unload tplamo

End Sub

Private Sub dlo2232_Click()

    If Frame1.Visible = True Then Exit Sub
    crea_planilla

End Sub

Private Sub fdl8923_Click()

    If Frame1.Visible = True Then Exit Sub
    horas1.Visible = True
    horas2.Visible = True

    Frame1.Visible = True
    Frame1.Caption = "Consolidado"

End Sub

Private Sub fjh433_Click()

    If Frame1.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

    On Error GoTo cmd127812_err

    If Combo3 = "%" Then
        MsgBox "Definir tipo planilla ", 48, "Aviso"
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
    Frame1.Top = 1080: Frame1.Left = 960
    Command1_Click

End Sub

Sub ejecutax(sw As Integer)

    Dim cad As String

    On Error GoTo cmd89123_err

    Command10.Visible = False

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

    If opcion1 = "1222" Then 'consulta del reloj
   
        cad = "SELECT Fecha,MIN(TimeIn) as HoraIn,MAX(TimeOut) as HoraOut from ingper  where codigo='" & dbGrid1.columns(1) & "' "
        cad = cad & " and fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        cad = cad & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
        cad = cad & " GROUP BY fecha order by fecha "
   
    End If

    'MsgBox cad
   
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

    fechai = Format(Now, "dd/mm/yyyy")
    fechaf = Format(Now, "dd/mm/yyyy")

End Sub

Sub inicializa()

End Sub

Private Sub grba1_Click()

End Sub

Sub consulta_planilla()

    Dim buf As String

    If mytablexx.State = 1 Then mytablexx.Close
    mytablexx.Open "select * from remune01 where tipopla='" & tipopla & "' and codigo='" & xcodigo & "'", cn, adOpenStatic, adLockOptimistic
    Set dbgrid3.DataSource = mytablexx
   
    If mytableyy.State = 1 Then mytableyy.Close
    mytableyy.Open "select * from descue01 where tipopla='" & tipopla & "'  and codigo='" & xcodigo & "'", cn, adOpenStatic, adLockOptimistic
    Set DBGrid4.DataSource = mytableyy
   
    If mytablezz.State = 1 Then mytablezz.Close
    mytablezz.Open "select * from aporta01 where tipopla='" & tipopla & "'  and codigo='" & xcodigo & "'", cn, adOpenStatic, adLockOptimistic
    Set dbgrid5.DataSource = mytablezz
   
    suma_general1
    suma_general2
    suma_general3
    suma_total
   
End Sub

Function existe_tipo1(buf As String, sw As Integer)

    Dim mytablex As New ADODB.Recordset

    If sw = 0 Then
        mytablex.Open "select * from remune01 where tipopla='" & tipopla & "' and  codigo='" & xcodigo & "' and tipo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            existe_tipo1 = 1

        End If

        mytablex.Close

    End If

    If sw = 1 Then
        mytablex.Open "select * from descue01 where tipopla='" & tipopla & "' and codigo='" & xcodigo & "' and tipo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            existe_tipo1 = 1

        End If

        mytablex.Close

    End If

    If sw = 2 Then
        mytablex.Open "select * from aporta01 where tipopla='" & tipopla & "' and codigo='" & xcodigo & "' and tipo='" & buf & "'", cn, adOpenStatic, adLockOptimistic

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
                mytabley.Open "select * from remune01 where tipopla='" & "" & tipopla & "' and  codigo='" & xcodigo & "' and tipo='" & mytablex.Fields("codigo") & "'", cn, adOpenStatic, adLockOptimistic

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
                mytabley.Open "select * from remune01 where tipopla='" & "" & tipopla & "' and  codigo='" & xcodigo & "' and tipo='" & mytablex.Fields("codigo") & "'", cn, adOpenStatic, adLockOptimistic

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
                mytabley.Open "select * from remune01 where tipopla='" & "" & tipopla & "' and codigo='" & xcodigo & "' and tipo='" & mytablex.Fields("codigo") & "'", cn, adOpenStatic, adLockOptimistic

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

Sub consulta_seccion()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.ListIndex = 0

    Frame5.Visible = True
    Frame5.Enabled = True
    bufferx = ""
    opcion1 = "233"
    ejecutax 1

End Sub

Sub crea_planilla()

    On Error GoTo cmd7812_err

    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub

    If Combo3 = "%" Then
        MsgBox "Definir tipo planilla ", 48, "Aviso"
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

Sub consulta_reloj()
    Combo2.Clear
    'Combo2.AddItem "MMAAAA"
    'Combo2.ListIndex = 0

    Frame5.Visible = True
    Frame5.Enabled = True
    bufferx = ""
    opcion1 = "1222"
    ejecutax 1

End Sub

Private Sub k8844_Click()

    If Frame1.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame2.Visible = True Then Exit Sub
    If Frame1.Visible = True Then Exit Sub
    horas1.Visible = False
    horas2.Visible = False
    Frame1.Visible = True
    Frame1.Caption = "Entradas Salidas"

End Sub

Sub imprimir_exell()

    Dim v          As Long

    Dim h          As Long

    Dim found      As Integer

    Dim I          As Integer

    Dim R          As Long

    Dim sdx        As Double

    Dim sdx1       As Double

    Dim sdx2       As Double

    Dim sdx3       As Double

    Dim sdx4       As Double

    Dim xingreso   As Double

    Dim xegreso    As Double
 
    Dim Heading(5) As String 'aki vamos a guardar los nombres de los campos que despues pasamos a la funcion

    On Error GoTo cmd1561212_err

    If mytablexr.RecordCount = 0 Then Exit Sub
    mytablexr.MoveFirst
    
    Heading(1) = "Fecha"
    Heading(2) = "HoraIn"
    Heading(3) = "HoraOut"
    Heading(4) = "NroHoras"
   
    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    Call Formato_Excelre(4, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
    v = 3
    h = 1

    objExcel.ActiveSheet.Cells(2, h) = "Reporte de Entradas Salidas"
    v = v + 1
    objExcel.ActiveSheet.Cells(v, h) = "Codigo:" & txempre.Fields("codigo") & " Nombre:" & txempre.Fields("nombre")
    v = v + 1

    Do

        If mytablexr.EOF Then Exit Do
        objExcel.ActiveSheet.Cells(v, h) = "'" & mytablexr.Fields("fecha")
        objExcel.ActiveSheet.Cells(v, h + 1) = "'" & mytablexr.Fields("Horain")
        objExcel.ActiveSheet.Cells(v, h + 2) = "'" & mytablexr.Fields("Horaout")
        imprime_diferencia v, h
            
        v = v + 1
            
        mytablexr.MoveNext
    Loop

    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    'MsgBox "Proceso Terminado ", 48, "Aviso"
    Exit Sub
cmd1561212_err:
    MsgBox "Error en exportacion excell " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub imprime_excell2()

    Dim t0, t1, t2

    Dim X

    Dim sdx         As Integer

    Dim I           As Long

    Dim v           As Long

    Dim h           As Long

    Dim buf         As String

    Dim mytablex    As New ADODB.Recordset

    Dim Heading(32) As String

    Dim xhora(32)   As String

    On Error GoTo cmd11561212_err
    
    If txempre.RecordCount = 0 Then Exit Sub
    txempre.MoveFirst
    
    Heading(1) = "Nombre"

    If Inicio_Excel() = 0 Then Exit Sub 'Llamamos a la funcion que abre el workbook en excel
    
    With objExcel.ActiveSheet
        .Range(.Cells(3, 1), .Cells(3, 33)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 8)).Font.bold = True
        .Range(.Cells(3, 1), .Cells(3, 8)).Font.Size = 8
        .Cells(1, 1) = ""
        .Cells(2, 1) = "'" & fechai & "  " & fechaf
        .Cells(3, 1) = "Nombre"
        .columns("A").ColumnWidth = 40

        For I = 1 To 31 Step 1
            .Cells(3, I + 1) = "'" & Format(I, "00")
        
        Next I

        sdx = 9

        If horas2.Value = True Then
            sdx = 15

        End If

        For I = 1 To 31
            .columns(I + 1).ColumnWidth = sdx
        Next I

    End With
    
    v = 4
    h = 1
    
    Do

        If txempre.EOF Then Exit Do
    
        For I = 1 To 32
            xhora(I) = ""
        Next I

        buf = "SELECT Fecha,MIN(TimeIn) as HoraIn,MAX(TimeOut) as HoraOut from ingper  where "
        buf = buf & "  fecha>='" & Format(fechai, "YYYYMMDD") & "'"
        buf = buf & " and fecha<='" & Format(fechaf, "YYYYMMDD") & "' "
        buf = buf & " and codigo='" & txempre.Fields("codigo") & "' GROUP BY fecha order by fecha "
        mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount > 0 Then
            Do

                If mytablex.EOF Then Exit Do
                X = Day("" & mytablex.Fields("fecha"))
                xhora(X) = ""

                If horas2.Value = True Then
                    xhora(X) = Format("" & mytablex.Fields("Horain"), "hh:mm:ss") & "-" & Format("" & mytablex.Fields("Horaout"), "hh:mm:ss")

                End If

                If horas1.Value = True Then
                    If IsDate("" & mytablex.Fields("Horaout")) And IsDate("" & mytablex.Fields("Horain")) Then
                        t0 = Format("" & mytablex.Fields("Horain"), "hh:mm:ss")
                        t1 = Format("" & mytablex.Fields("Horaout"), "hh:mm:ss")
                        t2 = Format(TimeValue(t1) - TimeValue(t0), "hh:mm:ss")
                        xhora(X) = Mid$("" & t2, 1, 5)

                    End If

                End If

                mytablex.MoveNext
            Loop

        End If

        mytablex.Close
        objExcel.ActiveSheet.Cells(v, h) = "" & txempre.Fields("Nombre")
        objExcel.ActiveSheet.Cells(v, h).Font.Size = 8

        For I = 1 To 31

            If Len(Trim(xhora(I))) = 0 Then
                objExcel.ActiveSheet.columns(I + 1).ColumnWidth = 5
            Else
                objExcel.ActiveSheet.columns(I + 1).ColumnWidth = 15

            End If

            objExcel.ActiveSheet.Cells(v, h + I) = "'" & xhora(I)
            objExcel.ActiveSheet.Cells(v, h + I).Font.Size = 8
        Next I

        v = v + 1
    
        txempre.MoveNext
    Loop
     
    Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
    
    Exit Sub
cmd11561212_err:
    MsgBox "Error en exportacion excell " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub imprime_diferencia(v As Long, h As Long)

    On Error GoTo cmd8912_err

    Dim t0, t1, t2 As String

    t0 = Format("" & mytablexr.Fields("Horain"), "hh:mm:ss")
    t1 = Format("" & mytablexr.Fields("Horaout"), "hh:mm:ss")
    t2 = Format(TimeValue(t1) - TimeValue(t0), "hh:mm:ss")
    objExcel.ActiveSheet.Cells(v, h + 3) = "" & t2
    objExcel.ActiveSheet.Cells(v, h + 3).Font.Size = 8
    Exit Sub
cmd8912_err:
    Exit Sub

End Sub

Public Function Formato_Excelre(Num_Campos As Integer, _
                                Nombre_Campos() As String) As Boolean

    Dim I

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
        .columns("B").ColumnWidth = 10
        .columns("C").ColumnWidth = 10
        .columns("d").ColumnWidth = 10
    
    End With

End Function

Public Function Formato_Excelree(Num_Campos As Integer, _
                                 Nombre_Campos() As String) As Boolean

    Dim I

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
        .columns("B").ColumnWidth = 10
        .columns("C").ColumnWidth = 10
        .columns("d").ColumnWidth = 10
    
    End With

End Function

