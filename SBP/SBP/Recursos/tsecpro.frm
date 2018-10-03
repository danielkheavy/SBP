VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tsecpro 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Etapas de la Produccion"
   ClientHeight    =   10080
   ClientLeft      =   165
   ClientTop       =   -90
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consultas"
      Height          =   9615
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   14775
      Begin VB.ComboBox Combo2 
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
         TabIndex        =   43
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox bproducto 
         Height          =   375
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   42
         Top             =   240
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   8655
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   15266
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   25
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
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ejecutar Busqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   46
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acepta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   12480
         TabIndex        =   45
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   12480
         TabIndex        =   44
         Top             =   2040
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Costos Directos e Indirectos"
      Height          =   10095
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   15735
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mano de Obra"
         Height          =   9015
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   15015
         Begin VB.TextBox costo 
            BeginProperty Font 
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
            TabIndex        =   64
            Top             =   2760
            Width           =   1455
         End
         Begin VB.TextBox nrohora 
            BeginProperty Font 
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
            TabIndex        =   62
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox mobserva 
            BeginProperty Font 
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
            TabIndex        =   58
            Top             =   2040
            Width           =   3855
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "&Guardar"
            Height          =   975
            Left            =   9480
            Picture         =   "tsecpro.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   360
            Width           =   1470
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Caption         =   "&Cerrar"
            Height          =   1020
            Left            =   9480
            Picture         =   "tsecpro.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Imprimir todo"
            Top             =   1320
            Width           =   1470
         End
         Begin VB.TextBox horaf 
            BeginProperty Font 
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
            MaxLength       =   8
            TabIndex        =   54
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox horai 
            BeginProperty Font 
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
            MaxLength       =   8
            TabIndex        =   52
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox operacion 
            BeginProperty Font 
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
            TabIndex        =   50
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox operario 
            BeginProperty Font 
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
            TabIndex        =   48
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label26 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Costo"
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
            TabIndex        =   65
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label Label16 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "NroHoras"
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
            TabIndex        =   63
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label Label15 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "HH:MM:SS"
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
            Left            =   3960
            TabIndex        =   61
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label Label12 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "HH:MM:SS"
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
            Left            =   3960
            TabIndex        =   60
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label25 
            BackColor       =   &H00FFFFC0&
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
            TabIndex        =   59
            Top             =   2040
            Width           =   2175
         End
         Begin VB.Label Label24 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "HraFinal"
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
            TabIndex        =   55
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label Label23 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "HoraInicio"
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
            TabIndex        =   53
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label22 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Operacion"
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
            TabIndex        =   51
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label21 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Operario"
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
            TabIndex        =   49
            Top             =   600
            Width           =   2175
         End
      End
      Begin MSDataGridLib.DataGrid dbgrid5 
         Height          =   4215
         Left            =   120
         TabIndex        =   31
         Top             =   5640
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   7435
         _Version        =   393216
         HeadLines       =   2
         RowHeight       =   27
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Operario"
            Caption         =   "Operario"
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
            DataField       =   "Operacion"
            Caption         =   "Operacion"
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
            DataField       =   "Horai"
            Caption         =   "Horai"
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
            DataField       =   "Horaf"
            Caption         =   "Horaf"
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
         BeginProperty Column04 
            DataField       =   "Observa"
            Caption         =   "Observa"
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
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3915.213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2520
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dbgrid4 
         Height          =   4215
         Left            =   120
         TabIndex        =   32
         Top             =   840
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   7435
         _Version        =   393216
         HeadLines       =   2
         RowHeight       =   27
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "Producto"
            Caption         =   "Producto"
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
         BeginProperty Column02 
            DataField       =   "Unidad"
            Caption         =   "Unidad"
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
            DataField       =   "Factor"
            Caption         =   "Factor"
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
         BeginProperty Column04 
            DataField       =   "Cantidad"
            Caption         =   "Cant"
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
         BeginProperty Column05 
            DataField       =   "Merma"
            Caption         =   "Merma"
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
         BeginProperty Column06 
            DataField       =   "Costo"
            Caption         =   "Costo"
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
         BeginProperty Column07 
            DataField       =   "Observa"
            Caption         =   "Observa"
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
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3915.213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2520
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Materiales"
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
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   11655
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mano de Obra"
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
         Left            =   120
         TabIndex        =   38
         Top             =   5160
         Width           =   11655
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   12000
         TabIndex        =   37
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Borra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   12000
         TabIndex        =   36
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   12000
         TabIndex        =   35
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Borra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   12000
         TabIndex        =   34
         Top             =   6480
         Width           =   1815
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close"
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
         Height          =   855
         Left            =   12000
         TabIndex        =   33
         Top             =   7560
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Frame2"
      Height          =   6615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   12495
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
         Height          =   375
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   28
         Top             =   3120
         Width           =   1935
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
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   26
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox estado 
         BeginProperty Font 
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
         TabIndex        =   24
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox vendedor 
         BeginProperty Font 
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
         TabIndex        =   22
         Top             =   2400
         Width           =   1935
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   20
         Top             =   2040
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
         TabIndex        =   18
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox tarjeta 
         BeginProperty Font 
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
         TabIndex        =   16
         Top             =   600
         Width           =   1935
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
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   14
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   8400
         Picture         =   "tsecpro.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir todo"
         Top             =   3120
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   8400
         Picture         =   "tsecpro.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2160
         Width           =   1470
      End
      Begin VB.Label Observa1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Observacion"
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
         TabIndex        =   29
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF00&
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
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFF00&
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
         TabIndex        =   25
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Responsable"
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
         TabIndex        =   23
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaF"
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
         TabIndex        =   21
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fechai"
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
         TabIndex        =   19
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tarjeta"
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
         TabIndex        =   17
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seccion"
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
         TabIndex        =   15
         Top             =   1320
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   15180
      TabIndex        =   2
      Top             =   0
      Width           =   15240
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
         Picture         =   "tsecpro.frx":2328
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
         Picture         =   "tsecpro.frx":353A
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
         Picture         =   "tsecpro.frx":474C
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
         Picture         =   "tsecpro.frx":595E
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
         Picture         =   "tsecpro.frx":6B70
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   14415
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7935
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14175
         _ExtentX        =   25003
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "Numero"
            Caption         =   "Numero"
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
            DataField       =   "Seccion"
            Caption         =   "Seccion"
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
            DataField       =   "Tarjeta"
            Caption         =   "Tarjeta"
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
            DataField       =   "Fechaf"
            Caption         =   "Fechai"
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
         BeginProperty Column04 
            DataField       =   "Fechaf"
            Caption         =   "Fechaf"
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
         BeginProperty Column05 
            DataField       =   "costomaterial"
            Caption         =   "CostoMaterial"
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
         BeginProperty Column06 
            DataField       =   "costomanoobra"
            Caption         =   "Costomanoobra"
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
         BeginProperty Column07 
            DataField       =   "Costo"
            Caption         =   "Costo"
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
         BeginProperty Column08 
            DataField       =   "Merma"
            Caption         =   "Merma"
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
         BeginProperty Column09 
            DataField       =   "Vendedor"
            Caption         =   "Responsable"
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
         BeginProperty Column10 
            DataField       =   "Estado"
            Caption         =   "Estado"
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
         BeginProperty Column11 
            DataField       =   "Observa"
            Caption         =   "Observa"
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
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1365.165
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   3465.071
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
      Begin VB.Menu dk9893 
         Caption         =   "&0.GENERAL"
      End
      Begin VB.Menu mnuArchivoArray 
         Caption         =   "Novisible"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu lie7744 
      Caption         =   "&Materiales"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tsecpro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txseccionp As New ADODB.Recordset

Dim txmaterial As New ADODB.Recordset

Dim txmanoobra As New ADODB.Recordset

Private Sub ajdu1_Click()

    If Frame3.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
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
    Numero.Enabled = False
    Numero = ""
    tarjeta.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    If Frame3.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub

    buf = txseccionp.Fields("seccion")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + txseccionp.Fields("seccion"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    txseccionp.Delete
    Command1_Click

    Exit Sub
cmd656_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub bproducto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub
    Label20_Click

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
    Frame5.Visible = False

End Sub

Private Sub Command3_Click()

    If Len(Trim("" & operario)) = 0 Then
        operario.SetFocus
        Exit Sub

    End If

    If Len(Trim("" & operacion)) = 0 Then
        operacion.SetFocus
        Exit Sub

    End If

    If Len(Trim("" & horai)) <> 8 Then
        horai.SetFocus
        Exit Sub

    End If

    If Len(Trim("" & horaf)) <> 8 Then
        horaf.SetFocus
        Exit Sub

    End If

    txmanoobra.AddNew
    txmanoobra.Fields("seccion") = Trim("" & txseccionp.Fields("seccion"))
    txmanoobra.Fields("tarjeta") = Trim("" & txseccionp.Fields("tarjeta"))
    txmanoobra.Fields("numero") = Val("" & txseccionp.Fields("numero"))
    txmanoobra.Fields("operario") = Trim("" & operario)
    txmanoobra.Fields("operacion") = Trim("" & operacion)
    txmanoobra.Fields("horai") = Trim("" & horai)
    txmanoobra.Fields("horaf") = Trim("" & horaf)
    txmanoobra.Fields("observa") = Trim("" & mobserva)
    txmanoobra.Fields("nrohora") = Val("" & nrohora)
    txmanoobra.Fields("costo") = Trim("" & costo)

    txmanoobra.Update
    txmanoobra.Requery
    Command2_Click

End Sub

Private Sub dk9893_Click()

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "seccionproduccion"
    reporgen.Show 1

End Sub

Sub prueba_reporte()

    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\seccionesproducto.rpt", "")
End Sub

Private Sub Label1_Click()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.ListIndex = 0

    opcion1 = "2"
    buffer = ""
    Frame4.Visible = True
    bproducto.SetFocus
    Label20_Click

End Sub

Private Sub Label10_Click()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.ListIndex = 0

    opcion1 = "1"
    buffer = ""
    Frame4.Visible = True
    bproducto.SetFocus
    Label20_Click

End Sub

Private Sub Label11_Click()

    On Error GoTo cmd899_err

    txmaterial.Delete
    txmaterial.Requery
    Exit Sub
cmd899_err:
    MsgBox "No existe Codigo ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Label13_Click()
    Frame5.Visible = True
    operario = ""
    operacion = ""
    horai = ""
    horaf = ""
    mobserva = ""
    nrohora = ""
    costo = ""
    Frame5.Caption = "ADD"
    operario.SetFocus

End Sub

Private Sub Label14_Click()

    On Error GoTo cmd8999_err

    txmanoobra.Delete
    txmanoobra.Requery
    Exit Sub
cmd8999_err:
    MsgBox "No existe Codigo ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Label17_Click()
    Frame3.Visible = False

End Sub

Private Sub Label18_Click()
    Frame4.Visible = False

End Sub

Private Sub Label19_Click()

    On Error GoTo cmd89123_err

    If opcion1 = "1" Then
        txmaterial.AddNew
        txmaterial.Fields("seccion") = Trim("" & txseccionp.Fields("seccion"))
        txmaterial.Fields("tarjeta") = Trim("" & txseccionp.Fields("tarjeta"))

        txmaterial.Fields("numero") = Val("" & txseccionp.Fields("numero"))
        'MsgBox "abc"
        txmaterial.Fields("producto") = Trim("" & DBGrid2.columns("producto"))

        txmaterial.Fields("descripcio") = Trim("" & DBGrid2.columns("descripcio"))
        txmaterial.Fields("unidad") = Trim("" & DBGrid2.columns("unidad"))
        txmaterial.Fields("factor") = Val("" & DBGrid2.columns("factor"))
        txmaterial.Fields("costo") = Val("" & DBGrid2.columns("costo"))
        txmaterial.Update
        txmaterial.Requery

        'cantidad.SetFocus
    End If

    If opcion1 = "2" Then
        seccion = "" & DBGrid2.columns("seccion")
        seccion.SetFocus

    End If

    If opcion1 = "7" Then
        operacion = "" & DBGrid2.columns("operacion")
        operacion.SetFocus

    End If

    If opcion1 = "3" Then
        tarjeta = "" & DBGrid2.columns("tarjeta")
        tarjeta.SetFocus

    End If

    If opcion1 = "4" Then
        vendedor = "" & DBGrid2.columns("codigo")
        vendedor.SetFocus

    End If

    If opcion1 = "6" Then
        operario = "" & DBGrid2.columns("codigo")
        operario.SetFocus

    End If

    Frame4.Visible = False
    Exit Sub
cmd89123_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub Label2_Click()
    Combo2.Clear
    Combo2.AddItem "Tarjeta"
    Combo2.ListIndex = 0

    opcion1 = "3"
    buffer = ""
    Frame4.Visible = True
    bproducto.SetFocus
    Label20_Click

End Sub

Private Sub Label20_Click()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    Set mytablex = Nothing
    DBGrid2.refresh

    If opcion1 = "1" Then
        buf = "select Descripcio,Producto,Unidad,Factor,Costou as Costo from producto where "

        If Combo2 = "Descripcio" Then
            buf = buf & " descripcio like '" & bproducto & "%'" '

        End If

        If Combo2 = "Producto" Then
            buf = " Producto like '" & bproducto & "%'" '

        End If

    End If

    If opcion1 = "2" Then
        buf = "select Descripcio,seccion from pseccion where "
        buf = buf & " descripcio like '" & bproducto & "%'" '

    End If

    If opcion1 = "3" Then
        buf = "select Tarjeta,Tarjeta,Producto,Descripcio,Unidad,Factor,Cantidad from tarjetaproduccion where "
        buf = buf & " tarjeta like '" & bproducto & "%' AND estado<>'1'" '

    End If

    If opcion1 = "4" Then
        buf = "select Nombre,Codigo from Vendedor where "
        buf = buf & " Nombre like '" & bproducto & "%'" '

    End If

    If opcion1 = "6" Then
        buf = "select Nombre,Codigo from Vendedor where "
        buf = buf & " Nombre like '" & bproducto & "%'" '

    End If

    If opcion1 = "7" Then
        buf = "select Descripcio,Operacion from toperaco where "
        buf = buf & " descripcio like '" & bproducto & "%'" '

    End If

    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = mytablex
    DBGrid2.columns(0).Width = 4000
    DBGrid2.columns(1).Width = 2000
   
    DBGrid2.refresh

End Sub

Private Sub Label21_Click()
    Combo2.Clear
    Combo2.AddItem "Nombre"
    Combo2.ListIndex = 0

    opcion1 = "6"
    buffer = ""
    Frame4.Visible = True
    bproducto.SetFocus
    Label20_Click

End Sub

Private Sub Label22_Click()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.ListIndex = 0

    opcion1 = "7"
    buffer = ""
    Frame4.Visible = True
    bproducto.SetFocus
    Label20_Click

End Sub

Private Sub Label6_Click()
    Combo2.Clear
    Combo2.AddItem "Nombre"
    Combo2.ListIndex = 0

    opcion1 = "4"
    buffer = ""
    Frame4.Visible = True
    bproducto.SetFocus
    Label20_Click

End Sub

Private Sub lie7744_Click()

    Dim buf As String

    On Error GoTo cmd12656_err

    If Frame2.Visible = True Then Exit Sub

    If Frame3.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub

    buf = txseccionp.Fields("seccion")

    If txmaterial.State = 1 Then txmaterial.Close
    txmaterial.Open "select * from promaterial where seccion='" & "" & txseccionp.Fields("seccion") & "' and numero=" & Val("" & txseccionp.Fields("numero")) & "", cn, adOpenStatic, adLockOptimistic
    Set DBGrid4.DataSource = txmaterial

    If txmanoobra.State = 1 Then txmanoobra.Close
    txmanoobra.Open "select * from promanoobra where seccion='" & "" & txseccionp.Fields("seccion") & "' and numero=" & Val("" & txseccionp.Fields("numero")) & "", cn, adOpenStatic, adLockOptimistic
    Set dbgrid5.DataSource = txmanoobra

    Frame3.Visible = True
    Exit Sub
cmd12656_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub seccion_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

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

    If Len(buffer) = 0 Then
        cad = "SELECT * from seccionproduccion  order by tarjeta  "

    End If

    If Len(buffer) > 0 Then
        cad = "SELECT *  from seccionproduccion   where  " & Combo1 & " like '" & buffer & "' order by tarjeta"

    End If

    If txseccionp.State = 1 Then txseccionp.Close
    txseccionp.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = txseccionp

    If txseccionp.RecordCount > 0 Then
        dbGrid1.SetFocus

    End If

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = &H71 Then
        txseccionp.Fields("estado") = "1"
        totaliza_secciones
        txseccionp.Update
        txseccionp.Requery

    End If

    If KeyCode = 13 Then

        'seccion = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'seccion.SetFocus
        'seccion_KeyPress 13
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

Private Sub dlo132_Click()

    If Frame3.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub

    If Frame2.Visible = True Then
        habilita 0
        Frame2.Visible = False
        dbGrid1.Enabled = True
        Exit Sub

    End If

    tsecpro.Hide
    Unload tsecpro

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    If Frame3.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub

    buf = txseccionp.Fields("seccion")

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
    Numero.Enabled = False
    tarjeta.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    If Frame3.Visible = True Then Exit Sub
    If Frame5.Visible = True Then Exit Sub
    If Frame4.Visible = True Then Exit Sub

    buf = txseccionp.Fields("seccion")

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
    Numero.Enabled = False
    tarjeta.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    'agregar_menus
    Command1_Click

End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "Tarjeta"
    Combo1.ListIndex = 0

End Sub

Sub inicializa()

    tarjeta = ""
    seccion = ""
    fechai = Format(Now, "dd/mm/yyyy")
    fechaf = Format(Now, "dd/mm/yyyy")
    estado = "0"
    vendedor = ""
    observa = ""

End Sub

Sub pone_registro()
    seccion = Trim("" & txseccionp.Fields("seccion"))
    tarjeta = Trim("" & txseccionp.Fields("tarjeta"))
    fechai = Trim("" & txseccionp.Fields("fechai"))
    fechaf = Trim("" & txseccionp.Fields("fechaf"))
    vendedor = Trim("" & txseccionp.Fields("vendedor"))
    estado = Trim("" & txseccionp.Fields("estado"))

End Sub

Sub grabando()
    txseccionp.Fields("seccion") = Trim(seccion)
    txseccionp.Fields("tarjeta") = Trim(tarjeta)
    txseccionp.Fields("fechai") = Trim(fechai)
    txseccionp.Fields("fechaf") = Trim(fechaf)
    txseccionp.Fields("vendedor") = Trim(vendedor)
    txseccionp.Fields("estado") = Trim(estado)

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
        'If Len(seccion) = 0 Then
        '  seccion.SetFocus
        ' Exit Function
        'End If
        'rbusca.Open "select seccion from seccionproduccion where seccion='" & seccion & "' and tarjeta='" & tarjeta & "'", cn, adOpenStatic, adLockOptimistic
        'If rbusca.RecordCount > 0 Then
        '   rbusca.Close
        '   MsgBox "Ya existe seccion ", 48, "Aviso"
        '   Exit Function
        'End If
        txseccionp.AddNew
        'txseccionp.Fields("seccion") = seccion
        grabando
        txseccionp.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        'txseccionp.Fields("seccion") = seccion
        grabando
        txseccionp.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    Dim mytablex As New ADODB.Recordset

    'If Len(seccion) = 0 Then
    '   seccion.SetFocus
    '   Exit Function
    'End If
    If Len(tarjeta) = 0 Then
        tarjeta.SetFocus
        Exit Function

    End If

    mytablex.Open "select * from tarjetaproduccion where tarjeta='" & tarjeta & "' and estado<>'1'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        tarjeta.SetFocus
        mytablex.Close
        Exit Function

    End If

    mytablex.Close

    If Len(seccion) = 0 Then
        seccion.SetFocus
        Exit Function

    End If

    mytablex.Open "select * from pseccion where seccion='" & seccion & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        seccion.SetFocus
        mytablex.Close
        Exit Function

    End If

    mytablex.Close

    If Not IsDate(fechai) Then
        fechai.SetFocus
        Exit Function

    End If

    If Not IsDate(fechaf) Then
        fechaf.SetFocus
        Exit Function

    End If

    If Len(vendedor) = 0 Then
        vendedor.SetFocus
        Exit Function

    End If

    mytablex.Open "select * from vendedor where codigo='" & vendedor & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        vendedor.SetFocus
        mytablex.Close
        Exit Function

    End If

    mytablex.Close

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

    mytablex.Open "select * from archivo where menu='seccion' and   estado='S'", cn, adOpenStatic, adLockOptimistic

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

Sub mnuarchivoarray_click(Index As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    buf = mnuArchivoArray(Index).Caption
    mytablex.Open "select * from archivo where menu='seccion' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close

    End If

    'busca el reporte
    buf = mytablex.Fields("archivo")
    mytablex.Close
    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\" & buf, "")

End Sub

Sub totaliza_secciones()

    Dim sdx      As Double

    Dim sdx1     As Double

    Dim sdx2     As Double

    Dim mytablex As New ADODB.Recordset

    sdx = 0
    sdx2 = 0
    mytablex.Open "select * from promaterial where seccion='" & "" & txseccionp.Fields("seccion") & "' and numero=" & Val("" & txseccionp.Fields("numero")) & "", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + (Val("" & mytablex.Fields("merma")) + Val("" & mytablex.Fields("cantidad"))) * Val("" & mytablex.Fields("costo"))
        sdx2 = sdx2 + Val("" & mytablex.Fields("costo")) * Val("" & mytablex.Fields("merma"))
        mytablex.MoveNext
    Loop
    mytablex.Close

    sdx1 = 0
    mytablex.Open "select * from promanoobra where seccion='" & "" & txseccionp.Fields("seccion") & "' and numero=" & Val("" & txseccionp.Fields("numero")) & "", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        sdx1 = sdx1 + Val("" & mytablex.Fields("costo")) * Val("" & mytablex.Fields("nrohora"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    txseccionp.Fields("costomaterial") = sdx
    txseccionp.Fields("costomanoobra") = sdx1
    txseccionp.Fields("costo") = sdx + sdx1
    txseccionp.Fields("merma") = sdx2
    txseccionp.Update
    txseccionp.Requery

End Sub
