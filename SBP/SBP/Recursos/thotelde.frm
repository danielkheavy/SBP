VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form thotelde 
   BackColor       =   &H00808080&
   Caption         =   "Tabla de Detalle Facturacion"
   ClientHeight    =   9930
   ClientLeft      =   165
   ClientTop       =   -765
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Consulta"
      Height          =   9975
      Left            =   15
      TabIndex        =   32
      Top             =   15
      Visible         =   0   'False
      Width           =   14895
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
         Top             =   240
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid dbgrid13 
         Height          =   8895
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   14175
         _ExtentX        =   25003
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Frame2"
      Height          =   6615
      Left            =   15
      TabIndex        =   10
      Top             =   30
      Visible         =   0   'False
      Width           =   12495
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
         TabIndex        =   41
         Top             =   3840
         Width           =   1935
      End
      Begin VB.ComboBox ntipo 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox cantidad 
         BeginProperty Font 
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
         TabIndex        =   29
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox total 
         BeginProperty Font 
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
         Top             =   3480
         Width           =   1935
      End
      Begin VB.TextBox precio 
         BeginProperty Font 
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
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox factor 
         BeginProperty Font 
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
         MaxLength       =   4
         TabIndex        =   23
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox unidad 
         BeginProperty Font 
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
         TabIndex        =   21
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox producto 
         BeginProperty Font 
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
         TabIndex        =   19
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox tipo 
         Enabled         =   0   'False
         BeginProperty Font 
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
         TabIndex        =   17
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox iddetalle 
         BeginProperty Font 
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
         Top             =   600
         Width           =   1935
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
         MaxLength       =   80
         TabIndex        =   13
         Top             =   1680
         Width           =   6975
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   10680
         Picture         =   "thotelde.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprimir todo"
         Top             =   1320
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   10680
         Picture         =   "thotelde.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label3 
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
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cantidad"
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
         TabIndex        =   30
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
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
         TabIndex        =   28
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Precio"
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
         TabIndex        =   26
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Factor"
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
         TabIndex        =   24
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unidad"
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
         TabIndex        =   22
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Producto"
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
         TabIndex        =   20
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aplicado a:"
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
         TabIndex        =   18
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "IdDetalle"
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
         TabIndex        =   16
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   15
         Top             =   1680
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   12435
      TabIndex        =   1
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
         Picture         =   "thotelde.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         Picture         =   "thotelde.frx":23A6
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Picture         =   "thotelde.frx":35B8
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
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "thotelde.frx":47CA
         Style           =   1  'Graphical
         TabIndex        =   3
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
         Picture         =   "thotelde.frx":59DC
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label idcheckin 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4440
         TabIndex        =   43
         Top             =   120
         Width           =   735
      End
      Begin VB.Label idfactura 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3600
         TabIndex        =   40
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   12495
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6855
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   12091
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
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "idDetalle"
            Caption         =   "ID"
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
            DataField       =   "Idfactura"
            Caption         =   "Idfactura"
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
            DataField       =   "fecha"
            Caption         =   "Fecha"
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
            DataField       =   "Producto"
            Caption         =   "Producto"
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
         BeginProperty Column05 
            DataField       =   "Descripcio"
            Caption         =   "Descripcio"
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
         BeginProperty Column06 
            DataField       =   "Unidad"
            Caption         =   "Und"
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
         BeginProperty Column07 
            DataField       =   "Factor"
            Caption         =   "Factor"
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
         BeginProperty Column08 
            DataField       =   "Cantidad"
            Caption         =   "Cant"
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
            DataField       =   "Precio"
            Caption         =   "Precio"
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
         BeginProperty Column10 
            DataField       =   "Total"
            Caption         =   "Total"
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
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   450.142
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   3614.74
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column10 
            EndProperty
         EndProperty
      End
      Begin VB.Label xtotal 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   10080
         TabIndex        =   38
         Top             =   7320
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         Height          =   375
         Left            =   8280
         TabIndex        =   37
         Top             =   7320
         Width           =   1815
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
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "thotelde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txbdco As New ADODB.Recordset

Private Sub ajdu1_Click()

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Nuevo"
    cmdGuardar.Enabled = True
    habilita 1
    iddetalle.Enabled = False
    iddetalle = ""
    ntipo.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    buf = "" & txbdco.Fields("iddetalle")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + "" & txbdco.Fields("iddetalle"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    txbdco.Delete
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

Private Sub Command4_Click()
    filtro

End Sub

Sub filtro()

    Dim mytablex As New ADODB.Recordset

    Dim cad      As String

    If opcion1 = "5" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select producto.Descripcio,producto.producto,precios.Unidad1,precios.Factor1,precios.pventa1 from producto inner join precios on producto.producto=precios.producto "

        End If

        If Len(Text1) > 0 Then
            cad = "select producto.Descripcio,producto.producto,precios.Unidad1,precios.Factor1,precios.pventa1 from producto inner join precios on producto.producto=precios.producto and   " & Combo2 & " like '" & Text1.Text & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 3000
        dbgrid13.columns(1).Width = 1000
              
        If mytablex.RecordCount > 0 Then
            dbgrid13.SetFocus

        End If

    End If

    If opcion1 = "1" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Descripcio,Mesa,Precio from mesa "

        End If

        If Len(Text1) > 0 Then
            cad = "select Descripcio,Mesa,Precio from mesa where " & Combo2 & " like '" & Text1.Text & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 3000
        dbgrid13.columns(1).Width = 1000

        If mytablex.RecordCount > 0 Then
            dbgrid13.SetFocus

        End If

    End If

    If opcion1 = "12" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Descripcio,Producto,Unidad,factor,Cantidad,Precio,Total,Fecha from hotelconsumo where tipo='P' and idcheckin=" & idcheckin

        End If

        If Len(Text1) > 0 Then
            cad = "select Descripcio,Producto,Unidad,Factor,Cantidad,Precio,Total,Fecha from hotelconsumo where tipo='P' and idcheckin=" & idcheckin & " and " & Combo2 & " like '" & Text1.Text & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 3000
        dbgrid13.columns(1).Width = 1000

        If mytablex.RecordCount > 0 Then
            dbgrid13.SetFocus

        End If

    End If

    If opcion1 = "13" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Descripcio,Producto,Unidad,factor,Cantidad,Precio,Total,Fecha from hotelconsumo where tipo='H' and idcheckin=" & idcheckin

        End If

        If Len(Text1) > 0 Then
            cad = "select Descripcio,Producto,Unidad,Factor,Cantidad,Precio,Total,Fecha from hotelconsumo where tipo='H' and idcheckin=" & idcheckin & " and " & Combo2 & " like '" & Text1.Text & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 3000
        dbgrid13.columns(1).Width = 1000

        If mytablex.RecordCount > 0 Then
            dbgrid13.SetFocus

        End If

    End If

    Exit Sub

End Sub

Private Sub dbgrid13_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim found    As Integer

    If KeyCode = 27 Then
        Text1.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "5" Then
            producto = Trim("" & dbgrid13.columns("producto"))
            descripcio = Trim("" & dbgrid13.columns("descripcio"))
            unidad = Trim("" & dbgrid13.columns("unidad1"))
            factor = Val("" & dbgrid13.columns("factor1"))
            cantidad = "1"
            precio = Trim("" & dbgrid13.columns("pventa1"))
            total = Trim("" & dbgrid13.columns("pventa1"))
            producto.SetFocus
            Frame3.Visible = False

        End If

        If opcion1 = "12" Then
            producto = Trim("" & dbgrid13.columns("producto"))
            descripcio = Trim("" & dbgrid13.columns("descripcio"))
            unidad = Trim("" & dbgrid13.columns("unidad"))
            factor = Val("" & dbgrid13.columns("factor"))
            cantidad = Val("" & dbgrid13.columns("cantidad"))
            precio = Trim("" & dbgrid13.columns("precio"))
            total = Trim("" & dbgrid13.columns("total"))
            fecha = Trim("" & dbgrid13.columns("Fecha"))
            producto.SetFocus
            Frame3.Visible = False

        End If

        If opcion1 = "13" Then
            producto = Trim("" & dbgrid13.columns("producto"))
            descripcio = Trim("" & dbgrid13.columns("descripcio"))
            unidad = Trim("" & dbgrid13.columns("unidad"))
            factor = Val("" & dbgrid13.columns("factor"))
            cantidad = Val("" & dbgrid13.columns("cantidad"))
            precio = Trim("" & dbgrid13.columns("precio"))
            total = Trim("" & dbgrid13.columns("total"))
            fecha = Trim("" & dbgrid13.columns("Fecha"))
            producto.SetFocus
            Frame3.Visible = False

        End If

        If opcion1 = "1" Then
            producto = Trim("" & dbgrid13.columns("mesa"))
            descripcio = Trim("" & dbgrid13.columns("descripcio"))
            unidad = "UND"
            factor = "1"
            cantidad = "1"
            precio = Trim("" & dbgrid13.columns("precio"))
            total = Trim("" & dbgrid13.columns("precio"))
            producto.SetFocus
            Frame3.Visible = False

        End If

    End If

End Sub

Sub prueba_reporte()

    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\iddetalleesproducto.rpt", "")
End Sub

Private Sub dk9893_Click()

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

End Sub

Private Sub iddetalle_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(iddetalle) = 0 Then Exit Sub
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
            cad = "SELECT * from hoteldetalle  where idfactura=  " & idfactura & ""

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT *  from iddetalle   where idfactura=" & idfactura & " and  " & Combo1 & " like '" & buffer & "%'"

        End If

        If txbdco.State = 1 Then txbdco.Close
        txbdco.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbGrid1.DataSource = txbdco

        'dbGrid1.columns(0).Width = 4000
        'dbGrid1.columns(1).Width = 2000
        If txbdco.RecordCount > 0 Then
            dbGrid1.SetFocus

        End If

        suma_total

    End If

End Sub

Private Sub Command2_Click()

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'iddetalle = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'iddetalle.SetFocus
        'iddetalle_KeyPress 13
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

    If Frame3.Visible = True Then
        Frame3.Visible = False
        ejecuta 1
        Exit Sub

    End If

    If Frame2.Visible = True Then
        habilita 0
        Frame2.Visible = False
        dbGrid1.Enabled = True
        Exit Sub

    End If

    thotelde.Hide
    Unload thotelde

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    buf = "" & txbdco.Fields("iddetalle")

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
    iddetalle.Enabled = False
    ntipo.SetFocus
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

    buf = "" & txbdco.Fields("iddetalle")

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
    iddetalle.Enabled = False
    ntipo.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    ntipo.Clear
    ntipo.AddItem "%"
    ntipo.AddItem "P"
    ntipo.AddItem "H"
    ntipo.ListIndex = 0
    'agregar_menus
    Command1_Click

End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "Descripcio"
    Combo1.AddItem "Producto"
    Combo1.ListIndex = 0

End Sub

Sub inicializa()
    tipo = ""
    producto = ""
    descripcio = ""
    unidad = ""
    factor = ""
    precio = ""
    cantidad = ""
    total = ""
    fecha = Format(Now, "dd/mm/yyyy")

End Sub

Sub pone_registro()
    iddetalle = Trim("" & txbdco.Fields("iddetalle"))
    tipo = Trim("" & txbdco.Fields("tipo"))
    producto = Trim("" & txbdco.Fields("producto"))
    descripcio = Trim("" & txbdco.Fields("descripcio"))
    unidad = Trim("" & txbdco.Fields("unidad"))
    factor = Trim("" & txbdco.Fields("factor"))
    precio = Trim("" & txbdco.Fields("precio"))
    cantidad = Trim("" & txbdco.Fields("cantidad"))
    total = Trim("" & txbdco.Fields("total"))
    fecha = Trim("" & txbdco.Fields("fecha"))

End Sub

Sub grabando()
    txbdco.Fields("idfactura") = Val(idfactura)
    txbdco.Fields("tipo") = Trim(tipo)
    txbdco.Fields("producto") = Trim(producto)
    txbdco.Fields("descripcio") = Trim(descripcio)
    txbdco.Fields("unidad") = Trim(unidad)
    txbdco.Fields("factor") = Val(factor)
    txbdco.Fields("precio") = Val(precio)
    txbdco.Fields("cantidad") = Val(cantidad)
    txbdco.Fields("total") = Val(total)
    txbdco.Fields("fecha") = Trim(fecha)

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
        txbdco.AddNew
        grabando
        txbdco.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        grabando
        txbdco.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    Dim sdx As Double

    sdx = Val(precio) * Val(cantidad)
    total = Format(sdx, "0.00")

    If Len(tipo) = 0 Then
        tipo.SetFocus
        Exit Function

    End If

    If Len(producto) = 0 Then
        producto.SetFocus
        Exit Function

    End If

    If Len(descripcio) = 0 Then
        descripcio.SetFocus
        Exit Function

    End If

    If Len(unidad) = 0 Then
        producto.SetFocus
        Exit Function

    End If

    If Len(factor) = 0 Then
        producto.SetFocus
        Exit Function

    End If

    If Val(precio) <= 0 Then
        precio.SetFocus
        Exit Function

    End If

    If Val(cantidad) = 0 Then
        cantidad.SetFocus
        Exit Function

    End If

    If Not IsDate(fecha) Then
        fecha.SetFocus
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

    mytablex.Open "select * from archivo where menu='iddetalle' and   estado='S'", cn, adOpenStatic, adLockOptimistic

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
    mytablex.Open "select * from archivo where menu='iddetalle' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close

    End If

    'busca el reporte
    buf = mytablex.Fields("archivo")
    mytablex.Close
    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\" & buf, "")

End Sub

Private Sub ntipo_Click()

    Dim found As Integer

    If ntipo <> "%" Then
        tipo = Trim("" & ntipo)
        producto = ""
        descripcio = ""
        unidad = ""
        factor = ""
        cantidad = ""
        precio = ""
        total = ""

    End If

End Sub

Sub consulta_producto()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "Producto"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "5"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_mesas()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "habitacion"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "1"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_mesasck()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "Habitacion"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "13"
    Text1.SetFocus
    Command4_Click

End Sub

Private Sub producto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        If tipo = "P" Then
            If Len(Trim(idcheckin)) = 0 Then
                consulta_producto
            Else
                consulta_productoCK

            End If

        End If

        If tipo = "H" Then
            If Trim(idcheckin) = 0 Then
                consulta_mesas
            Else
                consulta_mesasck

            End If

        End If

    End If

End Sub

Sub suma_total()

    Dim sdx As Double

    sdx = 0

    If txbdco.RecordCount > 0 Then
        Do

            If txbdco.EOF Then Exit Do
            sdx = sdx + Val("" & txbdco.Fields("total"))
            txbdco.MoveNext
        Loop

    End If

    xtotal = Format(sdx, "0.00")

End Sub

Sub consulta_productoCK()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "12"
    Text1.SetFocus
    Command4_Click

End Sub

