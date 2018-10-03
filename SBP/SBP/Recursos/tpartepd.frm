VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tpartepd 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Detalle Parte Produccion"
   ClientHeight    =   10080
   ClientLeft      =   165
   ClientTop       =   -615
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   15120
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consulta"
      Height          =   9975
      Left            =   0
      TabIndex        =   22
      Top             =   0
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
         TabIndex        =   25
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Filtrar"
         Height          =   495
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid dbgrid13 
         Height          =   8895
         Left            =   120
         TabIndex        =   26
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Frame2"
      Height          =   8175
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   14295
      Begin VB.TextBox cantidadp 
         BackColor       =   &H0080FF80&
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
         Height          =   615
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   45
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox avance 
         BackColor       =   &H0080FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   42
         Top             =   3960
         Width           =   2055
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
         MaxLength       =   6
         TabIndex        =   28
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   10560
         Picture         =   "tpartepd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   480
         Width           =   1470
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   10560
         Picture         =   "tpartepd.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Imprimir todo"
         Top             =   1560
         Width           =   1470
      End
      Begin VB.TextBox descripcio 
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
         MaxLength       =   100
         TabIndex        =   14
         Top             =   600
         Width           =   6015
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
         MaxLength       =   6
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox unidad 
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
         MaxLength       =   6
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox factor 
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
         MaxLength       =   4
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
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
         Height          =   735
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   10
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CantidadProgramado"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   46
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Avance"
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
         Left            =   2040
         TabIndex        =   43
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
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
         TabIndex        =   29
         Top             =   2040
         Width           =   2175
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
         TabIndex        =   21
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFF00&
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
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFF00&
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
         TabIndex        =   19
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
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
         TabIndex        =   18
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CantidadProducido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         TabIndex        =   17
         Top             =   3240
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   15060
      TabIndex        =   2
      Top             =   0
      Width           =   15120
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
         Left            =   4560
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tpartepd.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Refresca"
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
         Left            =   720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tpartepd.frx":23A6
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
         Left            =   2160
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tpartepd.frx":35B8
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
         Left            =   1440
         MaskColor       =   &H00E0E0E0&
         Picture         =   "tpartepd.frx":47CA
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
         Picture         =   "tpartepd.frx":59DC
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
      Width           =   12495
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   7335
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   12938
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
         ColumnCount     =   10
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
         BeginProperty Column02 
            DataField       =   "Unidad"
            Caption         =   "Und"
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
            DataField       =   "Factor"
            Caption         =   "Factor"
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
            DataField       =   "Cantidad"
            Caption         =   "Cant"
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
            DataField       =   "Precio"
            Caption         =   "Precio"
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
         BeginProperty Column07 
            DataField       =   "Bodega"
            Caption         =   "Almacen"
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
            DataField       =   "ordentrabajo"
            Caption         =   "OrdenTrabajo"
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
            DataField       =   "Tipomov"
            Caption         =   "TipoMov"
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
               ColumnWidth     =   5025.26
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
         EndProperty
      End
      Begin VB.Label fecha 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   8280
         TabIndex        =   44
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoMov"
         Height          =   495
         Left            =   5400
         TabIndex        =   41
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label tipomov 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   6600
         TabIndex        =   40
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label ordentrabajo 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   3840
         TabIndex        =   39
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OrdenTrabajo"
         Height          =   495
         Left            =   2640
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label bodega 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   1800
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Almacen"
         Height          =   495
         Left            =   960
         TabIndex        =   36
         Top             =   240
         Width           =   855
      End
      Begin VB.Label idx 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label total 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   35
      Top             =   9840
      Width           =   2535
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   34
      Top             =   9840
      Width           =   2415
   End
   Begin VB.Label ncantidad 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   33
      Top             =   9480
      Width           =   2535
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cantidad Productos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   32
      Top             =   9480
      Width           =   2415
   End
   Begin VB.Label items 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
      TabIndex        =   31
      Top             =   9120
      Width           =   2535
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero Items"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   30
      Top             =   9120
      Width           =   2415
   End
   Begin VB.Menu ajdu1 
      Caption         =   "&Add"
   End
   Begin VB.Menu f8443 
      Caption         =   "&Modifica"
      Enabled         =   0   'False
      Visible         =   0   'False
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
Attribute VB_Name = "tpartepd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txnotadi As New ADODB.Recordset

Private Sub ajdu1_Click()

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Nuevo"
    cmdGuardar.Enabled = True
    habilita 1
    producto.Enabled = True
    producto = ""
    producto.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf      As String

    Dim mytablex As New ADODB.Recordset

    On Error GoTo cmd656_err

    buf = "" & txnotadi.Fields("producto")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + "" & txnotadi.Fields("PRODUCTO"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    buf = "delete from detalle where local='01' and tipo='" & "" & tipomov & "' and serie='" & "" & tipomov & "' and numero='" & Trim("" & txnotadi.Fields("NUMERO")) & "' and producto='" & Trim("" & txnotadi.Fields("producto")) & "'"
    'MsgBox buf
    cn.Execute (buf)
    actualiza_stock "" & txnotadi.Fields("producto"), "" & txnotadi.Fields("bodega"), "" & tipomov, Val("" & txnotadi.Fields("cantidad")), -1
    borra_formula txnotadi, -1
    txnotadi.Delete
    Command1_Click
    Exit Sub
cmd656_err:
    MsgBox "Seleccione un dato " + error$, 48, "Aviso"
    Exit Sub

End Sub

Private Sub bodega_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1

        'consulta_bodega
    End If

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
    sumar

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
    sumar

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

Private Sub dbgrid13_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim found As Integer

    If KeyCode = 27 Then
        Text1.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then
        If opcion1 = "1" Then
            producto = Trim("" & dbgrid13.columns("producto"))
            descripcio = Trim("" & dbgrid13.columns("descripcio"))
            unidad = Trim("" & dbgrid13.columns("unidad"))
            factor = Trim("" & dbgrid13.columns("factor"))
            precio = Trim("" & dbgrid13.columns("costou"))
            'formula = Trim("" & dbgrid13.columns("formula"))
   
            Frame3.Visible = False
            producto.SetFocus

        End If

        If opcion1 = "2" Then
            producto = Trim("" & dbgrid13.columns("producto"))
            descripcio = Trim("" & dbgrid13.columns("descripcio"))
            unidad = Trim("" & dbgrid13.columns("unidad"))
            factor = Trim("" & dbgrid13.columns("factor"))
            cantidadp = Trim("" & dbgrid13.columns("cantidad"))
            'precio = Trim("" & dbgrid13.columns("costou"))
            'formula = Trim("" & dbgrid13.columns("formula"))
            avance = avance_produccion("" & ordentrabajo, "" & producto)
            Frame3.Visible = False
            producto.SetFocus

        End If

    End If

End Sub

Private Sub dk9893_Click()

    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "notaingresod"
    reporgen.Show 1

End Sub

Sub prueba_reporte()

    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\formulacionesproducto.rpt", "")
End Sub

Private Sub Command1_Click()
    Frame1.Visible = True
    Frame1.Enabled = True
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim cad As String

    cad = "SELECT * from parteproducciond where numero=" & idx

    If txnotadi.State = 1 Then txnotadi.Close
    txnotadi.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = txnotadi
    dbGrid1.columns(0).Width = 4000
    dbGrid1.columns(1).Width = 2000

    If txnotadi.RecordCount > 0 Then
        dbGrid1.SetFocus

    End If

    sumar

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then

        'formulacion = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'formulacion.SetFocus
        'formulacion_KeyPress 13
    End If

End Sub

Private Sub dlo132_Click()

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

    tpartepd.Hide
    Unload tpartepd

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    buf = "" & txnotadi.Fields("producto")

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
    'MsgBox "abc"
    producto.Enabled = False
    cantidad.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    buf = "" & txnotadi.Fields("producto")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    inicializa
    Frame2.Visible = True
    Frame2.Caption = "Zoom"
    cmdGuardar.Enabled = False
    pone_registro
    avance = avance_produccion("" & ordentrabajo, "" & producto)
    habilita 1
    producto.Enabled = False
    'MsgBox "ABC"
    cantidad.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    Command1_Click

End Sub

Sub inicializa()
    producto = ""
    descripcio = ""
    unidad = ""
    factor = ""
    cantidad = "1"
    cantidadp = ""
    precio = ""
    avance = ""

    'formula = ""
End Sub

Sub pone_registro()
    producto = Trim("" & txnotadi.Fields("producto"))
    descripcio = Trim("" & txnotadi.Fields("descripcio"))
    unidad = Trim("" & txnotadi.Fields("unidad"))
    factor = Trim("" & txnotadi.Fields("factor"))
    cantidad = Trim("" & txnotadi.Fields("cantidad"))
    precio = Trim("" & txnotadi.Fields("precio"))
    cantidadp = Trim("" & txnotadi.Fields("cantidadp"))

End Sub

Sub grabando()
    txnotadi.Fields("producto") = Trim(producto)
    txnotadi.Fields("descripcio") = Trim(descripcio)
    txnotadi.Fields("unidad") = Trim(unidad)
    txnotadi.Fields("factor") = Trim(factor)
    txnotadi.Fields("cantidad") = Val(cantidad)
    txnotadi.Fields("precio") = Val(precio)
    txnotadi.Fields("numero") = Val(idx)
    txnotadi.Fields("ordentrabajo") = Val(ordentrabajo)
    txnotadi.Fields("bodega") = Trim(bodega)
    txnotadi.Fields("fecha") = Trim(fecha)
    txnotadi.Fields("tipomov") = Trim(tipomov)
    txnotadi.Fields("cantidadp") = Val(cantidadp)

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
        rbusca.Open "select producto from parteproducciond where numero=" & idx & " and producto='" & producto & "'", cn, adOpenStatic, adLockOptimistic

        If rbusca.RecordCount > 0 Then
            rbusca.Close
            MsgBox "Ya existe  ", 48, "Aviso"
            Exit Function

        End If

        txnotadi.AddNew
        grabando
        graba_kardex txnotadi, Trim("" & tipomov)
        txnotadi.Update
        actualiza_stock "" & txnotadi.Fields("producto"), Trim("" & bodega), Trim(tipomov), Val("" & txnotadi.Fields("cantidad")), 1
        descarga_formula txnotadi, 1
        sumar
        'graba_items
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        grabando
        txnotadi.Update
        sumar
        'graba_items
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    If Len(producto) = 0 Then
        producto.SetFocus
        Exit Function

    End If

    If Len(descripcio) = 0 Then
        descripcio.SetFocus
        Exit Function

    End If

    If Len(unidad) = 0 Then
        unidad.SetFocus
        Exit Function

    End If

    If Val(factor) = 0 Then
        factor.SetFocus
        Exit Function

    End If

    If Val(cantidad) <= 0 Then
        cantidad.SetFocus
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

Sub consulta_producto()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "Producto"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "1"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_ordentrabajo()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.AddItem "Producto"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "2"
    Text1.SetFocus
    Command4_Click

End Sub

Private Sub producto_KeyPress(KeyAscii As Integer)

    If Len(Trim(producto)) = 0 Then
        producto.SetFocus
        Exit Sub

    End If

    avance = avance_produccion("" & ordentrabajo, "" & producto)

End Sub

Private Sub producto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        'consulta_producto
        consulta_ordentrabajo

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
            cad = "select Descripcio,Producto,Unidad,Factor,Costou from Producto "

        End If

        If Len(Text1) > 0 Then
            cad = "select Descripcio,Producto,Unidad,factor,Costou from producto where  " & Combo2 & " like '" & Text1 & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 5000
        dbgrid13.columns(1).Width = 2000
        'dbgrid13.columns(2).Width = 1000
        'dbgrid13.columns(3).Width = 1000
           
    End If

    If opcion1 = "2" Then  'Orden Trabajo
        If Len(Text1) = 0 Then
            cad = "select Descripcio,Producto,Unidad,Factor,Cantidad from ordentrabajod where ordentrabajo=" & ordentrabajo

        End If

        If Len(Text1) > 0 Then
            cad = "select Descripcio,Producto,Unidad,factor,Cantidad from ordentrabajod where ordentrabajo=" & ordentrabajo & " and  " & Combo2 & " like '" & Text1 & "%'"

        End If

        If mytablex.State = 1 Then mytablex.Close
        mytablex.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbgrid13.DataSource = mytablex
        dbgrid13.columns(0).Width = 5000
        dbgrid13.columns(1).Width = 2000
        'dbgrid13.columns(2).Width = 1000
        'dbgrid13.columns(3).Width = 1000
           
    End If
   
    If mytablex.RecordCount > 0 Then
        dbgrid13.SetFocus

    End If

    Exit Sub

End Sub

Sub sumar()

    Dim sdx  As Double

    Dim sdx1 As Double

    Dim sdx2 As Double

    sdx = 0
    sdx1 = 0
    sdx2 = 0
    txnotadi.Requery
    Do

        If txnotadi.EOF Then Exit Do
        sdx = sdx + 1
        sdx1 = sdx1 + Val("" & txnotadi.Fields("cantidad"))
        sdx2 = sdx2 + Val("" & txnotadi.Fields("cantidad")) * Val("" & txnotadi.Fields("precio"))
        txnotadi.MoveNext
    Loop
    items = Format(sdx, "0000")
    ncantidad = Format(sdx1, "0.00")
    total = Format(sdx2, "0.00")

End Sub

Sub actualiza_stock(idxproducto As String, _
                    idxbodega As String, _
                    idxtipomov As String, _
                    idxcantidad As Double, _
                    sw As Integer)

    Dim mytabley As New ADODB.Recordset

    mytabley.Open "select * from almacen where producto='" & idxproducto & "' and local='01' and bodega='" & "" & idxbodega & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount > 0 Then
        If idxtipomov = "S" Then
            mytabley.Fields("saldo") = Val("" & mytabley.Fields("saldo")) + (sw) * idxcantidad

        End If

        If idxtipomov = "T" Then
            mytabley.Fields("saldo") = Val("" & mytabley.Fields("saldo")) - (sw) * idxcantidad

        End If

        mytabley.Update
    Else
        mytabley.AddNew
        mytabley.Fields("producto") = Trim(idxproducto)
        mytabley.Fields("bodega") = Trim(idxbodega)
        mytabley.Fields("local") = "01"
        mytabley.Fields("producto") = Trim(idxproducto)
        mytabley.Update

    End If

    mytabley.Close

End Sub

Sub borra_formula(mytabley As ADODB.Recordset, sw As Integer)

    Dim mytablez As New ADODB.Recordset

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    buf = "SELECT     dbo.ordentrabajod.producto, dbo.ordentrabajod.formula, dbo.ordentrabajod.descripcio, dbo.formulacion.id, dbo.componente.producto AS compo,"
    buf = buf & " dbo.componente.descripcio AS Expr2, dbo.componente.cantidad AS cant, dbo.componente.unidad, dbo.componente.factor "
    buf = buf & " FROM  dbo.ordentrabajod INNER JOIN"
    buf = buf & " dbo.formulacion ON dbo.ordentrabajod.producto = dbo.formulacion.producto AND dbo.ordentrabajod.formula = dbo.formulacion.id INNER JOIN"
    buf = buf & " dbo.componente ON dbo.formulacion.id = dbo.componente.id"
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        'graba la receta
        actualiza_stock "" & mytablex.Fields("compo"), "" & mytabley.Fields("bodega"), "T", Val("" & mytabley.Fields("cantidad")) * Val("" & mytablex.Fields("cant")), sw
        cn.Execute ("delete from detalle where serie='" & "T" & "' AND NUMERO='" & Trim("" & mytabley.Fields("numero")) & "' and producto='" & Trim("" & mytablex.Fields("compo")) & "'")
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Sub descarga_formula(mytabley As ADODB.Recordset, sw As Integer)

    Dim mytablex As New ADODB.Recordset

    Dim buf      As String

    buf = "SELECT     dbo.ordentrabajod.producto, dbo.ordentrabajod.formula, dbo.ordentrabajod.descripcio, dbo.formulacion.id, dbo.ordentrabajodf.producto AS compo,"
    buf = buf & " dbo.ordentrabajodf.descripcio AS Expr2, dbo.ordentrabajodf.cantidad AS cant, dbo.ordentrabajodf.unidad, dbo.ordentrabajodf.factor "
    buf = buf & " FROM  dbo.ordentrabajod INNER JOIN"
    buf = buf & " dbo.formulacion ON dbo.ordentrabajod.producto = dbo.formulacion.producto AND dbo.ordentrabajod.formula = dbo.formulacion.id INNER JOIN"
    buf = buf & " dbo.ordentrabajodf ON dbo.formulacion.ordentrabajod = dbo.ordentrabajodf.ordentrabajod"
    mytablex.Open buf, cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        Exit Sub

    End If

    Do

        If mytablex.EOF Then Exit Do
        'graba la receta
        graba_receta mytablex, mytabley, "T"
        actualiza_stock "" & mytablex.Fields("compo"), "" & mytabley.Fields("bodega"), "T", Val("" & mytabley.Fields("cantidad")) * Val("" & mytablex.Fields("cant")), sw
        mytablex.MoveNext
    Loop
    mytablex.Close

End Sub

Function avance_produccion(xordentrabajo As String, xproducto As String) As Double

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    mytablex.Open "select * from parteproducciond where ordentrabajo=" & xordentrabajo & " and producto='" & xproducto & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close
        avance_produccion = 0
        Exit Function

    End If

    sdx = 0
    Do

        If mytablex.EOF Then Exit Do
        sdx = sdx + Val("" & mytablex.Fields("cantidad"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    avance_produccion = sdx

End Function

Sub graba_kardex(mytablex As ADODB.Recordset, xtipomov As String)

    Dim mytablez As New ADODB.Recordset

    'S ENTRADA
    'T SALIDA
    mytablez.Open "select * from detalle where 2=1", cn, adOpenDynamic, adLockOptimistic
    mytablez.AddNew
    mytablez.Fields("estado") = "2"
    mytablez.Fields("acu") = xtipomov
    mytablez.Fields("tipo") = xtipomov
    mytablez.Fields("serie") = xtipomov
    mytablez.Fields("numero") = Trim("" & mytablex.Fields("numero"))
    mytablez.Fields("cantidad") = Val("" & mytablex.Fields("cantidad"))
    mytablez.Fields("local") = "01"
    mytablez.Fields("tipoclie") = "V"
    mytablez.Fields("codigo") = "PP"
    mytablez.Fields("acu1") = ""
    'mytablez.Fields("fecha") = Format(Now, "dd/mm/yyyy")
    mytablez.Fields("moneda") = "S"
    mytablez.Fields("producto") = Trim("" & mytablex.Fields("PRODUCTO"))
    mytablez.Fields("descripcio") = Mid$("" & mytablex.Fields("descripcio"), 1, 60)
    mytablez.Fields("unidad") = Trim("" & mytablex.Fields("unidad"))
    mytablez.Fields("factor") = Val("" & mytablex.Fields("factor"))
    mytablez.Fields("precio") = 0
    mytablez.Fields("igv") = 18
    mytablez.Fields("neto") = 0
    mytablez.Fields("descuento") = 0
    mytablez.Fields("subtotal") = 0
    mytablez.Fields("impuesto") = 0
    mytablez.Fields("total") = 0
    mytablez.Fields("fecha") = Format("" & mytablex.Fields("fecha"), "dd/mm/yyyy")
    mytablez.Fields("fechacrea") = Format("" & mytablex.Fields("fecha"), "dd/mm/yyyy")
    mytablez.Fields("hora") = Format(Now, "hh:mm:ss")
    mytablez.Fields("vendedor") = ""
    mytablez.Fields("bodega") = Trim("" & mytablex.Fields("bodega"))
    mytablez.Fields("bodegaf") = ""
    mytablez.Fields("deslipo") = 0
    mytablez.Fields("flage") = ""
    mytablez.Fields("linea") = ""
    mytablez.Fields("t1") = 0
    mytablez.Fields("t2") = 0
    mytablez.Fields("t3") = 0
    mytablez.Fields("t4") = 0
    mytablez.Fields("t5") = 0
    mytablez.Fields("t6") = 0
    mytablez.Fields("t7") = 0
    mytablez.Fields("t8") = 0
    mytablez.Fields("t9") = 0
    mytablez.Fields("t10") = 0
    mytablez.Fields("t11") = 0
    mytablez.Fields("t12") = 0
    mytablez.Fields("t13") = 0
    mytablez.Fields("t14") = 0
    mytablez.Fields("t15") = 0
    mytablez.Fields("t16") = 0
    mytablez.Fields("l1") = ""
    mytablez.Fields("l2") = ""
    mytablez.Fields("l3") = ""
    mytablez.Fields("l4") = ""
    'mytablez.Fields("local") = ""
    mytablez.Fields("proveedorp") = ""
    mytablez.Fields("observa1") = ""
    mytablez.Fields("observa2") = ""
    mytablez.Fields("observa3") = ""
    mytablez.Fields("observa4") = ""
    mytablez.Fields("zona") = ""
    mytablez.Fields("isc") = 0
    mytablez.Fields("tax") = 0
    mytablez.Fields("vtaneta") = 0
    mytablez.Fields("tcosto") = 0
    mytablez.Fields("ganancia") = 0
    mytablez.Fields("comision") = 0
    mytablez.Fields("usuario") = ""
    mytablez.Fields("caja") = ""
    mytablez.Fields("turno") = ""
    mytablez.Fields("servicio") = ""
    mytablez.Fields("comanda") = ""
    mytablez.Fields("mesa") = ""
    mytablez.Fields("salon") = ""
    mytablez.Fields("mesero") = ""
    mytablez.Update
    mytablez.Close

End Sub

Sub graba_receta(mytablex As ADODB.Recordset, _
                 mytabley As ADODB.Recordset, _
                 xtipomov As String)

    Dim mytablez As New ADODB.Recordset

    'S ENTRADA
    'T SALIDA
    mytablez.Open "select * from detalle where 2=1", cn, adOpenDynamic, adLockOptimistic
    mytablez.AddNew
    mytablez.Fields("estado") = "2"
    mytablez.Fields("acu") = xtipomov
    mytablez.Fields("tipo") = xtipomov
    mytablez.Fields("serie") = xtipomov
    mytablez.Fields("numero") = "" & mytabley.Fields("numero")
    mytablez.Fields("cantidad") = Val("" & mytablex.Fields("cant")) * Val("" & mytabley.Fields("cantidad"))
    mytablez.Fields("local") = "01"
    mytablez.Fields("tipoclie") = "V"
    mytablez.Fields("codigo") = "PP"
    mytablez.Fields("acu1") = ""
    'mytablez.Fields("fecha") = Format(Now, "dd/mm/yyyy")
    mytablez.Fields("moneda") = "S"
    mytablez.Fields("producto") = "" & mytablex.Fields("compo")
    mytablez.Fields("descripcio") = Mid$("" & mytablex.Fields("Expr2"), 1, 60)
    mytablez.Fields("unidad") = "" & mytablex.Fields("unidad")
    mytablez.Fields("factor") = Val("" & mytablex.Fields("factor"))
    mytablez.Fields("precio") = 0
    mytablez.Fields("igv") = 18
    mytablez.Fields("neto") = 0
    mytablez.Fields("descuento") = 0
    mytablez.Fields("subtotal") = 0
    mytablez.Fields("impuesto") = 0
    mytablez.Fields("total") = 0
    mytablez.Fields("fecha") = Format("" & mytabley.Fields("fecha"), "dd/mm/yyyy")
    mytablez.Fields("fechacrea") = Format("" & mytabley.Fields("fecha"), "dd/mm/yyyy")
    mytablez.Fields("hora") = Format(Now, "hh:mm:ss")
    mytablez.Fields("vendedor") = ""
    mytablez.Fields("bodega") = Trim("" & mytabley.Fields("bodega"))
    mytablez.Fields("bodegaf") = ""
    mytablez.Fields("deslipo") = 0
    mytablez.Fields("flage") = ""
    mytablez.Fields("linea") = ""
    mytablez.Fields("t1") = 0
    mytablez.Fields("t2") = 0
    mytablez.Fields("t3") = 0
    mytablez.Fields("t4") = 0
    mytablez.Fields("t5") = 0
    mytablez.Fields("t6") = 0
    mytablez.Fields("t7") = 0
    mytablez.Fields("t8") = 0
    mytablez.Fields("t9") = 0
    mytablez.Fields("t10") = 0
    mytablez.Fields("t11") = 0
    mytablez.Fields("t12") = 0
    mytablez.Fields("t13") = 0
    mytablez.Fields("t14") = 0
    mytablez.Fields("t15") = 0
    mytablez.Fields("t16") = 0
    mytablez.Fields("l1") = ""
    mytablez.Fields("l2") = ""
    mytablez.Fields("l3") = ""
    mytablez.Fields("l4") = ""
    'mytablez.Fields("local") = ""
    mytablez.Fields("proveedorp") = ""
    mytablez.Fields("observa1") = ""
    mytablez.Fields("observa2") = ""
    mytablez.Fields("observa3") = ""
    mytablez.Fields("observa4") = ""
    mytablez.Fields("zona") = ""
    mytablez.Fields("isc") = 0
    mytablez.Fields("tax") = 0
    mytablez.Fields("vtaneta") = 0
    mytablez.Fields("tcosto") = 0
    mytablez.Fields("ganancia") = 0
    mytablez.Fields("comision") = 0
    mytablez.Fields("usuario") = ""
    mytablez.Fields("caja") = ""
    mytablez.Fields("turno") = ""
    mytablez.Fields("servicio") = ""
    mytablez.Fields("comanda") = ""
    mytablez.Fields("mesa") = ""
    mytablez.Fields("salon") = ""
    mytablez.Fields("mesero") = ""
    mytablez.Update
    mytablez.Close

End Sub
