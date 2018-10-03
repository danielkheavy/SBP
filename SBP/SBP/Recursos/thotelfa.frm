VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form thotelfa 
   BackColor       =   &H00808080&
   Caption         =   "Facturacion"
   ClientHeight    =   10080
   ClientLeft      =   165
   ClientTop       =   -615
   ClientWidth     =   17400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   17400
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Consulta"
      Height          =   9975
      Left            =   60
      TabIndex        =   71
      Top             =   840
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
         TabIndex        =   74
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
         TabIndex        =   73
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
         TabIndex        =   72
         Top             =   240
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid dbgrid13 
         Height          =   8895
         Left            =   120
         TabIndex        =   75
         Top             =   840
         Width           =   14655
         _ExtentX        =   25850
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
      BackColor       =   &H00808080&
      Caption         =   "Detalle"
      Height          =   7695
      Left            =   60
      TabIndex        =   51
      Top             =   810
      Visible         =   0   'False
      Width           =   10095
      Begin VB.ComboBox ntipo 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox bcantidad 
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
         TabIndex        =   61
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox btotal 
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
         TabIndex        =   60
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox bprecio 
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
         TabIndex        =   59
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox bfactor 
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
         TabIndex        =   58
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox bunidad 
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
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox bproducto 
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
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox btipo 
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
         TabIndex        =   55
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox bdescripcio 
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
         TabIndex        =   54
         Top             =   1560
         Width           =   6975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Grabar"
         Height          =   735
         Left            =   8520
         TabIndex        =   53
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Close"
         Height          =   735
         Left            =   7560
         TabIndex        =   52
         Top             =   240
         Width           =   975
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         Picture         =   "thotelfa.frx":0000
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label13 
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
         TabIndex        =   70
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Label Label15 
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
         TabIndex        =   69
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label16 
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
         TabIndex        =   68
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label17 
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
         TabIndex        =   67
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label18 
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
         TabIndex        =   66
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label20 
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
         TabIndex        =   65
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
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
         TabIndex        =   64
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label23 
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
         TabIndex        =   63
         Top             =   1560
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Frame2"
      Height          =   8895
      Left            =   90
      TabIndex        =   11
      Top             =   810
      Visible         =   0   'False
      Width           =   14895
      Begin VB.TextBox hotelcuadre 
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
         Left            =   7080
         MaxLength       =   7
         TabIndex        =   78
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Modifica"
         Height          =   615
         Left            =   1200
         TabIndex        =   77
         Top             =   6840
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "CarHabitac"
         Height          =   615
         Left            =   2280
         TabIndex        =   76
         Top             =   6840
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "NuevoProd"
         Height          =   615
         Left            =   120
         TabIndex        =   50
         Top             =   6840
         Width           =   1095
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
         MaxLength       =   80
         TabIndex        =   49
         Top             =   2400
         Width           =   7575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "BorraTodo"
         Height          =   615
         Left            =   5520
         TabIndex        =   48
         Top             =   6840
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "BorraLinea"
         Height          =   615
         Left            =   4440
         TabIndex        =   47
         Top             =   6840
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CarConsumo"
         Height          =   615
         Left            =   3360
         TabIndex        =   46
         Top             =   6840
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid dbgrid10 
         Height          =   3375
         Left            =   120
         TabIndex        =   45
         Top             =   2880
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   5953
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   29
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
         ColumnCount     =   9
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column02 
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "factor"
            Caption         =   "Fac"
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
            DataField       =   "cantidad"
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
         BeginProperty Column06 
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
         BeginProperty Column07 
            DataField       =   "Deslipo"
            Caption         =   "Dscto"
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
               ColumnWidth     =   420.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1244.976
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3809.764
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   959.811
            EndProperty
         EndProperty
      End
      Begin VB.TextBox habitacion 
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
         TabIndex        =   42
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox idcheckin 
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
         Left            =   7080
         MaxLength       =   11
         TabIndex        =   40
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox operador 
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
         Left            =   7080
         MaxLength       =   11
         TabIndex        =   38
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox total 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7920
         MaxLength       =   10
         TabIndex        =   36
         Top             =   6240
         Width           =   1935
      End
      Begin VB.TextBox impuesto 
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
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   34
         Top             =   6240
         Width           =   1935
      End
      Begin VB.TextBox subtotal 
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
         TabIndex        =   32
         Top             =   6240
         Width           =   1335
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
         Left            =   7080
         MaxLength       =   1
         TabIndex        =   30
         Top             =   1320
         Width           =   735
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
         TabIndex        =   28
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox codigo 
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
         TabIndex        =   25
         Top             =   1680
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
         Left            =   3000
         MaxLength       =   11
         TabIndex        =   24
         Top             =   960
         Width           =   1935
      End
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
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   21
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox tipo 
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
         MaxLength       =   3
         TabIndex        =   18
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox nombre 
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
         TabIndex        =   16
         Top             =   2040
         Width           =   7575
      End
      Begin VB.TextBox idfactura 
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
         Left            =   9000
         MaxLength       =   6
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   10320
         Picture         =   "thotelfa.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir todo"
         Top             =   1320
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   10320
         Picture         =   "thotelfa.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label14 
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
         Left            =   5640
         TabIndex        =   79
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   9000
         Picture         =   "thotelfa.frx":149E
         Stretch         =   -1  'True
         Top             =   960
         Width           =   375
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   4200
         Picture         =   "thotelfa.frx":17A8
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   375
      End
      Begin VB.Image Image6 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   3000
         Picture         =   "thotelfa.frx":1AB2
         Stretch         =   -1  'True
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Habitacion"
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
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CheckIn"
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
         Left            =   6240
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Operador"
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
         Left            =   5640
         TabIndex        =   39
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF00&
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
         Left            =   6360
         TabIndex        =   37
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Impuesto"
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
         Left            =   3000
         TabIndex        =   35
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Subtotal"
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
         TabIndex        =   33
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         Left            =   5640
         TabIndex        =   31
         Top             =   1320
         Width           =   1455
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
         TabIndex        =   29
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label7 
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
         TabIndex        =   27
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label28 
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
         TabIndex        =   26
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
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
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie-Numero"
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
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nombres"
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
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Id"
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
         Left            =   8160
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   17340
      TabIndex        =   2
      Top             =   0
      Width           =   17400
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
         Picture         =   "thotelfa.frx":1DBC
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
         Picture         =   "thotelfa.frx":2FCE
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
         Picture         =   "thotelfa.frx":41E0
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
         Picture         =   "thotelfa.frx":53F2
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
         Picture         =   "thotelfa.frx":6604
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
      End
      Begin VB.Label idxcheckin 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   44
         Top             =   120
         Width           =   105
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Consulta"
      Height          =   8295
      Left            =   45
      TabIndex        =   0
      Top             =   765
      Width           =   14895
      Begin MSDataGridLib.DataGrid dbGrid1 
         Height          =   6855
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14655
         _ExtentX        =   25850
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "Idfactura"
            Caption         =   "Id"
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
         BeginProperty Column02 
            DataField       =   "Serie"
            Caption         =   "Serie"
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
            DataField       =   "Numero"
            Caption         =   "Numero"
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
            DataField       =   "Fecha"
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
         BeginProperty Column05 
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
         BeginProperty Column06 
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
         BeginProperty Column07 
            DataField       =   "Moneda"
            Caption         =   "M"
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
         BeginProperty Column09 
            DataField       =   "Impuesto"
            Caption         =   "Impuesto"
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
            DataField       =   "Subtotal"
            Caption         =   "Subtotal"
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
         BeginProperty Column11 
            DataField       =   "Operario"
            Caption         =   "Operario"
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
         BeginProperty Column12 
            DataField       =   "idcheckin"
            Caption         =   "Idcheckin"
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
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   480.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1335.118
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   3314.835
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   315.213
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
         EndProperty
      End
      Begin VB.Label totalreserva 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   11760
         TabIndex        =   23
         Top             =   7200
         Width           =   2895
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         Height          =   495
         Left            =   10080
         TabIndex        =   22
         Top             =   7200
         Width           =   1695
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
   Begin VB.Menu Det8912 
      Caption         =   "&Detalle"
      Enabled         =   0   'False
      Visible         =   0   'False
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
Attribute VB_Name = "thotelfa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txidfacturax As New ADODB.Recordset

Dim mytablexx    As New ADODB.Recordset

Dim mytableyy    As New ADODB.Recordset

Dim dmytablex    As New ADODB.Recordset

Private Sub agente_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_agente

    End If

End Sub

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
    idfactura.Enabled = False
    idfactura = ""
    tipo.SetFocus

End Sub

Private Sub bcantidad_KeyPress(KeyAscii As Integer)

    Dim sdx As Double

    If KeyAscii <> 13 Then Exit Sub
    sdx = Val(bcantidad) * Val(bprecio)
    btotal = Format(sdx, "0.00")

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    buf = "" & txidfacturax.Fields("idfactura")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + "" & txidfacturax.Fields("idfactura"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    cn.Execute ("delete from hoteldetalle where idfactura=" & Val("" & txidfacturax.Fields("idfactura")))
    txidfacturax.Delete
    Command1_Click

    Exit Sub
cmd656_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub bprecio_KeyPress(KeyAscii As Integer)

    Dim sdx As Double

    If KeyAscii <> 13 Then Exit Sub
    sdx = Val(bcantidad) * Val(bprecio)
    btotal = Format(sdx, "0.00")

End Sub

Private Sub bproducto_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub bproducto_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        If btipo = "P" Then
            consulta_producto

        End If

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

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_codigo1

    End If

End Sub

Private Sub Command10_Click()

    If modifica_producto() = 1 Then
        Frame4.Visible = True
        Frame4.Caption = "MODIFICA"

    End If

End Sub

Private Sub Command2_Click()

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    'Dim dmytablex As New adodb.Recordset
    'cn.Execute ("delete from hotedetalletmp ")
    'dmytablex.Open "select * from hotedetalletmp where idecheckin=" & Val(idxcheckin), cn, adOpenStatic, adLockOptimistic
    mytablex.Open "select * from hotelconsumo where idecheckin=" & Val(idxcheckin), cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        dmytablex.AddNew
        dmytablex.Fields("idecheckin") = Val("" & mytablex.Fields("idecheckin"))
        'dmytablex.Fields("habitacion") = Trim("" & mytablex.Fields("habitacion"))
        dmytablex.Fields("tipo") = Trim("" & mytablex.Fields("tipo"))
        dmytablex.Fields("producto") = Trim("" & mytablex.Fields("habitacion"))
        dmytablex.Fields("descripcio") = Trim("" & mytablex.Fields("descripcio"))
        dmytablex.Fields("unidad") = Trim("" & mytablex.Fields("unidad"))
        dmytablex.Fields("factor") = Val("" & mytablex.Fields("factor"))
        dmytablex.Fields("precio") = Val("" & mytablex.Fields("precio"))
        dmytablex.Fields("cantidad") = Trim("" & mytablex.Fields("cantidad"))
        dmytablex.Fields("total") = Val("" & mytablex.Fields("total"))
        dmytablex.Fields("fecha") = Format("" & mytablex.Fields("fecha"), "dd/mm/yyyy")
        dmytablex.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    carga_detalles

End Sub

Private Sub Command3_Click()

    Dim buf As String

    On Error GoTo cmd903_err

    buf = "" & dmytablex.Fields("idfactura")
    dmytablex.Delete
    suma_detalle
    Exit Sub
cmd903_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Sub suma_detalle()

    Dim sdx As Double

    sdx = 0
    Do

        If dmytablex.EOF Then Exit Do
        sdx = sdx + Val("" & dmytablex.Fields("total"))
        dmytablex.MoveNext
    Loop
    total = Format(sdx, "0.00")

End Sub

Private Sub Command4_Click()
    filtro

End Sub

Sub filtro()

    Dim mytablex As New ADODB.Recordset

    Dim cad      As String

    If opcion1 = "1" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Descripcio,Tipo,Tipodoc from tipo where (tipodoc='A' OR tipodoc='B' or tipodoc='C' or tipodoc='D' or tipodoc='G') "

        End If

        If Len(Text1) > 0 Then
            cad = "select Descripcio,Tipo,Tipodoc from tipo where (tipodoc='A' OR tipodoc='B' or tipodoc='C' or tipodoc='D' or tipodoc='G') AND   " & Combo2 & " like '" & Text1.Text & "%'"

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

    If opcion1 = "6" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Nombre,Codigo,Direccion,Correo from clientes "

        End If

        If Len(Text1) > 0 Then
            cad = "select Nombre,Codigo,Direccion,Correo from clientes where  " & Combo2 & " like '" & Text1.Text & "%'"

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

    If opcion1 = "3" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Nombre,Codigo from Vendedor "

        End If

        If Len(Text1) > 0 Then
            cad = "select Nombre,Codigo from Vendedor where  " & Combo2 & " like '" & Text1.Text & "%'"

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

    If opcion1 = "4" Then  'producto
        If Len(Text1) = 0 Then
            cad = "select Nombre,Codigo from Vendedor "

        End If

        If Len(Text1) > 0 Then
            cad = "select Nombre,Codigo from Vendedor where  " & Combo2 & " like '" & Text1.Text & "%'"

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

    If opcion1 = "5" Then  'producto
 
        If Len(Text1) = 0 Then
            cad = "select producto.Descripcio,producto.producto,precios.Unidad1,precios.Factor1,precios.pventa1 from producto INNER join precios on producto.producto=precios.producto AND PRECIOS.LOCAL='01'"

        End If

        If Len(Text1) > 0 Then
            cad = "select producto.Descripcio,producto.producto,precios.Unidad1,precios.Factor1,precios.pventa1 from producto INNER join precios on producto.producto=precios.producto AND PRECIOS.LOCAL='01' and   " & Combo2 & " like '" & Text1.Text & "%'"

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

    If opcion1 = "8" Then  'checkin
        If Len(Text1) = 0 Then
            cad = "select Habitacion,Checkin,Codigo,Nombre,Direccion,Arribofecha,Arribofechaf,ArriboHora,ArriboHoraf  from hotelcheckin where estado='0'"

        End If

        If Len(Text1) > 0 Then
            cad = "select Habitacion,Checkin,Codigo,Nombre,Direccion,Arribofecha,Arribofechaf,ArriboHora,ArriboHoraf from hotelcheckin where estado='0' and  " & Combo2 & " like '" & Text1.Text & "%'"

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

Private Sub DataGrid2_AfterColEdit(ByVal ColIndex As Integer)

    Select Case ColIndex

        Case 2

    End Select

End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub Command5_Click()
    cn.Execute ("delete from hotedetalletmp")
    carga_detalles

End Sub

Private Sub Command6_Click()

    Dim sdx As Double

    If Trim(btipo) <> "P" And Trim(btipo) <> "H" Then
        btipo.SetFocus
        Exit Sub

    End If

    If Len(Trim(bproducto)) = 0 Then
        bproducto.SetFocus
        Exit Sub

    End If

    If Len(Trim(bdescripcio)) = 0 Then
        bdescripcio.SetFocus
        Exit Sub

    End If

    If Len(Trim(bunidad)) = 0 Then
        bproducto.SetFocus
        Exit Sub

    End If

    If Val(Trim(bfactor)) <= 0 Then
        bfactor.SetFocus
        Exit Sub

    End If

    If Val(Trim(bcantidad)) <= 0 Then
        bcantidad.SetFocus
        Exit Sub

    End If

    If Val(Trim(bprecio)) <= 0 Then
        bprecio.SetFocus
        Exit Sub

    End If

    sdx = Val(bcantidad) * Val(bprecio)
    btotal = Format(sdx, "0.00")

    If Frame4.Caption = "NUEVO" Then
        dmytablex.AddNew

    End If

    'dmytablex.Fields("habitacion") = Trim(habitacion)

    dmytablex.Fields("idecheckin") = Trim(idxcheckin)
    dmytablex.Fields("tipo") = Trim(btipo)
    dmytablex.Fields("producto") = Trim(bproducto)
    dmytablex.Fields("descripcio") = Trim(bdescripcio)
    dmytablex.Fields("unidad") = Trim(bunidad)
    dmytablex.Fields("factor") = Val(bfactor)
    dmytablex.Fields("precio") = Val(bprecio)
    dmytablex.Fields("cantidad") = Val(bcantidad)
    dmytablex.Fields("total") = Val(btotal)
    dmytablex.Fields("fecha") = Trim(fecha)
    dmytablex.Update
    carga_detalles
    Frame4.Visible = False

End Sub

Private Sub Command7_Click()
    dlo132_Click

End Sub

Private Sub Command8_Click()
    Frame4.Visible = True
    Frame4.Caption = "NUEVO"
    inicializa_detalle

    'btipo.SetFocus
End Sub

Private Sub Command9_Click()

    carga_entrefechas

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
            bproducto = Trim("" & dbgrid13.columns("producto"))
            bdescripcio = Trim("" & dbgrid13.columns("descripcio"))
            bunidad = Trim("" & dbgrid13.columns("unidad1"))
            bfactor = Val("" & dbgrid13.columns("factor1"))
            bcantidad = "1"
            bprecio = Trim("" & dbgrid13.columns("pventa1"))
            btotal = Trim("" & dbgrid13.columns("pventa1"))
            bproducto.SetFocus
            Frame3.Visible = False

        End If

        If opcion1 = "6" Then
            codigo = Trim("" & dbgrid13.columns("codigo"))
            nombre = Trim("" & dbgrid13.columns("nombre"))
            direccion = Trim("" & dbgrid13.columns("direccion"))
            'correo = Trim("" & dbgrid13.columns("correo"))
            nombre.SetFocus
            Frame3.Visible = False
   
        End If

        If opcion1 = "1" Then
            tipo = Trim("" & dbgrid13.columns("tipo"))
            busca_parameca
            Frame3.Visible = False

        End If

        If opcion1 = "8" Then
            idcheckin = Trim("" & dbgrid13.columns("checkin"))
            habitacion = Trim("" & dbgrid13.columns("habitacion"))
            codigo = Trim("" & dbgrid13.columns("Codigo"))
            nombre = Trim("" & dbgrid13.columns("Nombre"))
            direccion = Trim("" & dbgrid13.columns("Direccion"))
            Frame3.Visible = False

        End If

        If opcion1 = "3" Then
            operador = Trim("" & dbgrid13.columns("codigo"))
            Frame3.Visible = False

        End If

    End If

End Sub

Private Sub Det8912_Click()

    On Error GoTo cmd67890_err

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub
    thotelde.idcheckin = "" & txidfacturax.Fields("idcheckin")
    thotelde.idfactura = "" & txidfacturax.Fields("idfactura")
    thotelde.Show 1
    Exit Sub
cmd67890_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub dk9893_Click()

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "hotelfactura"
    reporgen.Show 1

End Sub

Sub prueba_reporte()

    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\idfacturaesproducto.rpt", "")
End Sub

Private Sub huesped_KeyUp(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub idcheckin_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_checkin

    End If

End Sub

Private Sub Image1_Click()
    consulta_codigo1

End Sub

Private Sub Image2_Click()
    consulta_vendedor

End Sub

Private Sub Image3_Click()

    If btipo = "P" Then
        consulta_producto

    End If

End Sub

Private Sub Image6_Click()
    consulta_tipo

End Sub

Private Sub Label6_Click()
    consulta_producto

End Sub

Private Sub Label7_Click()
    consulta_codigo

End Sub

Private Sub mesa_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_producto

    End If

End Sub

Private Sub nombre_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_codigo

    End If

End Sub

Private Sub idfactura_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(idfactura) = 0 Then Exit Sub

    'descripcio.SetFocus
End Sub

Private Sub Command1_Click()
    Frame1.Visible = True
    Frame1.Enabled = True
    'buffer = ""
    opcion1 = "1"
    ejecuta 1
    'carga_detalles

End Sub

Sub ejecuta(sw As Integer)

    Dim cad As String

    Dim sdx As Double

    If Len(buffer) = 0 Then
        cad = "SELECT * from hotelfactura  where idcheckin=" & Val(idxcheckin)

    End If

    If Len(buffer) > 0 Then
        cad = "SELECT *  from hotelfactura   where idecheckin=" & Val(idxcheckin) & " and " & Combo1 & " like '" & buffer & "%'"

    End If

    If txidfacturax.State = 1 Then txidfacturax.Close
    txidfacturax.Open cad, cn, adOpenStatic, adLockOptimistic
    Set dbGrid1.DataSource = txidfacturax

    'dbGrid1.columns(0).Width = 4000
    'dbGrid1.columns(1).Width = 2000
    If txidfacturax.RecordCount > 0 Then
        dbGrid1.SetFocus

    End If

    suma_total

End Sub

Sub suma_total()

    Dim sdx As Double

    sdx = 0
    Do

        If txidfacturax.EOF Then Exit Do
        sdx = sdx + Val("" & txidfacturax.Fields("total"))
        txidfacturax.MoveNext
    Loop
    totalreserva = Format(sdx, "0.00")

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'idfactura = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'idfactura.SetFocus
        'idfactura_KeyPress 13
    End If

End Sub

Private Sub dlo132_Click()

    If Frame3.Visible = True Then
        Frame3.Visible = False
        Exit Sub

    End If

    If Frame4.Visible = True Then
        Frame4.Visible = False
        Exit Sub

    End If

    If Frame2.Visible = True Then
        habilita 0
        Frame2.Visible = False
        dbGrid1.Enabled = True
        ejecuta 1
        Exit Sub

    End If

    thotelfa.Hide
    Unload thotelfa

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    If Frame2.Visible = True Then Exit Sub
    If Frame3.Visible = True Then Exit Sub

    buf = "" & txidfacturax.Fields("idfactura")

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
    idfactura.Enabled = False

    tipo.SetFocus
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

    buf = "" & txidfacturax.Fields("idfactura")

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
    idfactura.Enabled = False
    carga_detalles
    tipo.SetFocus
    Exit Sub
cmd556_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub Form_Activate()
    'agregar_menus
    Command1_Click

End Sub

Sub consulta_checkin()
    Combo2.Clear
    Combo2.AddItem "habitacion"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "8"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_tipo()
    Combo2.Clear
    Combo2.AddItem "Descripcio"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "1"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_codigo()
    Combo2.Clear
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "2"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_codigo1()
    Combo2.Clear
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "6"
    Text1.SetFocus
    Command4_Click

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

Sub consulta_vendedor()
    Combo2.Clear
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "3"
    Text1.SetFocus
    Command4_Click

End Sub

Sub consulta_agente()
    Combo2.Clear
    Combo2.AddItem "Nombre"
    Combo2.AddItem "Codigo"
    Combo2.ListIndex = 0
    Frame3.Enabled = True
    Frame3.Visible = True
    Text1 = ""
    opcion1 = "4"
    Text1.SetFocus
    Command4_Click

End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "Nombre"
    Combo1.ListIndex = 0

    ntipo.Clear
    ntipo.AddItem "%"
    ntipo.AddItem "P"
    ntipo.AddItem "H"
    ntipo.ListIndex = 1

End Sub

Sub inicializa()

    Dim mytablex As New ADODB.Recordset

    hotelcuadre = Trim("" & treevho.turno)
    fecha = Format(Now, "dd/mm/yyyy")
    'hora = Format(Now, "hh:mm")
    habitacion = ""
    idcheckin = Trim("" & idxcheckin)
    tipo = ""
    serie = ""
    Numero = ""
    nombre = ""
    codigo = ""
    direccion = ""
    operador = Trim(gusuario)
    moneda = "S"
    impuesto = ""
    subtotal = ""
    total = ""
    mytablex.Open "select * from hotelcheckin where checkin=" & Val(idxcheckin), cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        habitacion = Trim("" & mytablex.Fields("habitacion"))
        codigo = Trim("" & mytablex.Fields("codigo"))
        nombre = Trim("" & mytablex.Fields("nombre"))
        direccion = Trim("" & mytablex.Fields("direccion"))

    End If

    mytablex.Close
    cn.Execute ("delete from hotedetalletmp ")
    carga_detalles

End Sub

Sub pone_registro()
    idcheckin = Trim("" & txidfacturax.Fields("idcheckin"))
    'habitacion = Trim("" & txidfacturax.Fields("habitacion"))
    idfactura = Trim("" & txidfacturax.Fields("idfactura"))
    fecha = Trim("" & txidfacturax.Fields("fecha"))
    'hora = Trim("" & txidfacturax.Fields("hora"))
    tipo = Trim("" & txidfacturax.Fields("tipo"))
    serie = Trim("" & txidfacturax.Fields("serie"))
    Numero = Trim("" & txidfacturax.Fields("numero"))
    codigo = Trim("" & txidfacturax.Fields("codigo"))
    nombre = Trim("" & txidfacturax.Fields("nombre"))
    direccion = Trim("" & txidfacturax.Fields("direccion"))
    operador = Trim("" & txidfacturax.Fields("operador"))
    subtotal = Trim("" & txidfacturax.Fields("subtotal"))
    impuesto = Trim("" & txidfacturax.Fields("impuesto"))
    total = Trim("" & txidfacturax.Fields("total"))
    moneda = Trim("" & txidfacturax.Fields("moneda"))
    hotelcuadre = Trim("" & txidfacturax.Fields("hotelcuadre"))
    carga_tmp
    carga_detalles

End Sub

Sub grabando()

    If Val(hotelcuadre) <= 0 Then
        hotelcuadre = Trim("" & treevho.turno)

    End If

    txidfacturax.Fields("hotelcuadre") = Val(hotelcuadre)
    txidfacturax.Fields("idcheckin") = Val(idcheckin)
    'txidfacturax.Fields("habitacion") = Trim(habitacion)
    txidfacturax.Fields("fecha") = Trim(fecha)
    txidfacturax.Fields("hora") = Format(Now, "hh:mm:ss")
    txidfacturax.Fields("tipo") = Trim(tipo)
    txidfacturax.Fields("serie") = Trim(serie)
    txidfacturax.Fields("numero") = Trim(Numero)
    txidfacturax.Fields("codigo") = Trim(codigo)
    txidfacturax.Fields("nombre") = Trim(nombre)
    txidfacturax.Fields("direccion") = Trim(direccion)
    txidfacturax.Fields("operador") = Trim(operador)
    txidfacturax.Fields("subtotal") = Val(subtotal)
    txidfacturax.Fields("impuesto") = Val(impuesto)
    txidfacturax.Fields("total") = Val(total)
    txidfacturax.Fields("moneda") = Trim(moneda)

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
        'If Len(idfactura) = 0 Then
        '   idfactura.SetFocus
        '   Exit Function
        'End If
        'rbusca.Open "select idfactura from idfactura where idfactura='" & idfactura & "'", cn, adOpenStatic, adLockOptimistic
        'If rbusca.RecordCount > 0 Then
        '   rbusca.Close
        '   MsgBox "Ya existe idfactura ", 48, "Aviso"
        '   Exit Function
        'End If
        txidfacturax.AddNew
        'txidfacturax.Fields("idfactura") = idfactura
        grabando
        txidfacturax.Update
        graba_detalle
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        'txidfacturax.Fields("idfactura") = idfactura
        grabando
        txidfacturax.Update
        graba_detalle
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    If Not IsDate(fecha) Then
        fecha.SetFocus
        Exit Function

    End If

    If Len(Trim(tipo)) = 0 Then
        tipo.SetFocus
        Exit Function

    End If

    If Len(Trim(serie)) = 0 Then
        serie.SetFocus
        Exit Function

    End If

    If Len(Trim(Numero)) = 0 Then
        Numero.SetFocus
        Exit Function

    End If

    If Len(Trim(codigo)) = 0 Then
        codigo.SetFocus
        Exit Function

    End If

    If Len(Trim(nombre)) = 0 Then
        nombre.SetFocus
        Exit Function

    End If

    If Len(Trim(operador)) = 0 Then
        operador.SetFocus
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

    mytablex.Open "select * from archivo where menu='idfactura' and   estado='S'", cn, adOpenStatic, adLockOptimistic

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
    mytablex.Open "select * from archivo where menu='idfactura' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic

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
        btipo = Trim("" & ntipo)
        bproducto = ""
        bdescripcio = ""
        bunidad = ""
        bfactor = ""
        bcantidad = ""
        bprecio = ""
        btotal = ""

        If btipo = "H" Then
            found = busca_habitacion()

        End If

    End If

End Sub

Private Sub operador_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_vendedor

    End If

End Sub

Private Sub tipo_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 Then Exit Sub

End Sub

Private Sub tipo_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H70 Then  'f1
        consulta_tipo

    End If

End Sub

Sub carga_tmp()

    Dim I As Integer

    On Error GoTo cmd9012_err

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from hotedetalletmp ")
    mytabley.Open "select * from hotedetalletmp where idfactura=" & Val(idfactura), cn, adOpenStatic, adLockOptimistic
    mytablex.Open "select * from hoteldetalle where idfactura=" & Val(idfactura), cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        mytabley.AddNew
        mytabley.Fields("idecheckin") = Val("" & mytablex.Fields("idecheckin"))
        'mytabley.Fields("habitacion") = Trim("" & mytablex.Fields("habitacion"))
        mytabley.Fields("tipo") = Trim("" & mytablex.Fields("tipo"))
        mytabley.Fields("producto") = Trim("" & mytablex.Fields("producto"))
        mytabley.Fields("descripcio") = Trim("" & mytablex.Fields("descripcio"))
        mytabley.Fields("unidad") = Trim("" & mytablex.Fields("unidad"))
        mytabley.Fields("factor") = Val("" & mytablex.Fields("factor"))
        mytabley.Fields("precio") = Val("" & mytablex.Fields("precio"))
        mytabley.Fields("cantidad") = Trim("" & mytablex.Fields("cantidad"))
        mytabley.Fields("total") = Val("" & mytablex.Fields("total"))
        'mytabley.Fields("fecha") = Format("" & mytablex.Fields("fecha"), "dd/mm/yyyy")
        mytabley.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    mytabley.Close
    Exit Sub
cmd9012_err:
    MsgBox "Aviso en carga Tmp " + error$, 48, "Aviso"
    Exit Sub

End Sub

Sub carga_detalles()

    Dim sdx As Double

    If dmytablex.State = 1 Then
        dmytablex.Close

    End If

    sdx = 0
    dmytablex.Open "select * from hotedetalletmp", cn, adOpenStatic, adLockOptimistic
    Set dbgrid10.DataSource = dmytablex
    Do

        If dmytablex.EOF Then Exit Do
        sdx = sdx + Val("" & dmytablex.Fields("total"))
        dmytablex.MoveNext
    Loop
    total = Format(sdx, "0.00")

End Sub

Sub ir_inicio()

    On Error GoTo cmd6789_err

    dmytablex.MoveFirst
    Exit Sub
cmd6789_err:
    Exit Sub

End Sub

Sub graba_detalle()

    Dim I        As Integer

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from hoteldetalle where idfactura=" & Val("" & txidfacturax.Fields("idfactura")))
    mytabley.Open "select * from hoteldetalle where idfactura=" & Val("" & txidfacturax.Fields("idfactura")), cn, adOpenStatic, adLockOptimistic
    ir_inicio
    Do

        If dmytablex.EOF Then Exit Do
        mytabley.AddNew
        mytabley.Fields("idfactura") = Val("" & txidfacturax.Fields("idfactura"))
        mytabley.Fields("idecheckin") = Val("" & dmytablex.Fields("idecheckin"))
        'mytabley.Fields("habitacion") = Trim("" & dmytablex.Fields("habitacion"))
        mytabley.Fields("tipo") = Trim("" & dmytablex.Fields("tipo"))
        mytabley.Fields("producto") = Trim("" & dmytablex.Fields("producto"))
        mytabley.Fields("descripcio") = Trim("" & dmytablex.Fields("descripcio"))
        mytabley.Fields("unidad") = Trim("" & dmytablex.Fields("unidad"))
        mytabley.Fields("factor") = Val("" & dmytablex.Fields("factor"))
        mytabley.Fields("precio") = Val("" & dmytablex.Fields("precio"))
        mytabley.Fields("cantidad") = Trim("" & dmytablex.Fields("cantidad"))
        mytabley.Fields("total") = Val("" & dmytablex.Fields("total"))
        'mytabley.Fields("fecha") = Format("" & dmytablex.Fields("fecha"), "dd/mm/yyyy")
        mytabley.Update
        dmytablex.MoveNext
    Loop
    mytabley.Close
  
End Sub

Function busca_parameca()

    Dim mytablex As New ADODB.Recordset

    Dim sdx      As Double

    mytablex.Open "select * from tipo where tipo='" & tipo & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        serie = Trim("" & mytablex.Fields("serie"))
        sdx = Val("" & mytablex.Fields("numero")) + 1
        Numero = "" & sdx

    End If

    If Len(Trim(serie)) = 0 Then
        serie = "001"

    End If

    mytablex.Close

End Function

Sub inicializa_detalle()
    btipo = "P"
    bproducto = ""
    bdescripcio = ""
    bunidad = ""
    bfactor = ""
    bcantidad = ""
    bprecio = ""
    btotal = ""

End Sub

Function busca_habitacion()

    Dim mytablex As New ADODB.Recordset

    mytablex.Open "select * from Habitacion where habitacion='" & Trim(habitacion) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        bproducto = Trim("" & habitacion)
        bdescripcio = Trim("" & mytablex.Fields("descripcio"))
        bunidad = "UND"
        bfactor = "1"
        bcantidad = "1"
        bprecio = Trim("" & mytablex.Fields("precio"))
        btotal = Trim("" & mytablex.Fields("precio"))

    End If

    mytablex.Close

End Function

Sub carga_entrefechas()

    Dim dias     As Integer

    Dim xhoy     As String

    Dim I        As Integer

    Dim mytablex As New ADODB.Recordset

    xhoy = Format(Now, "dd/mm/yyyy")
    dias = -1
    mytablex.Open "select * from hotelcheckin where checkin=" & Val(idxcheckin), cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount > 0 Then
        dias = DateDiff("d", Format("" & mytablex.Fields("arribofecha"), "dd/mm/yyyy"), xhoy)

        If dias = 0 Then
            dias = 1

        End If

        'MsgBox dias
        xhoy = Format("" & mytablex.Fields("arribofecha"), "dd/mm/yyyy")

        For I = 1 To dias
            dmytablex.AddNew
            dmytablex.Fields("idecheckin") = Val("" & mytablex.Fields("checkin"))
            'dmytablex.Fields("habitacion") = Trim("" & mytablex.Fields("habitacion"))
            dmytablex.Fields("tipo") = "H"
            dmytablex.Fields("producto") = Trim("" & mytablex.Fields("habitacion"))
            dmytablex.Fields("descripcio") = "Habitacion" & " " & Format(xhoy, "dd/mm/yyyy")
            dmytablex.Fields("unidad") = "UND"
            dmytablex.Fields("factor") = 1
            dmytablex.Fields("precio") = Val("" & mytablex.Fields("precio"))
            dmytablex.Fields("cantidad") = 1
            dmytablex.Fields("total") = Val("" & mytablex.Fields("precio"))
            xhoy = DateAdd("D", 1, xhoy)
            xhoy = Format(xhoy, "dd/mm/yyyy")
            'dmytablex.Fields("fecha") = Format(xhoy, "dd/mm/yyyy")
            dmytablex.Update
        Next I

    End If

    mytablex.Close
    carga_detalles

End Sub

Function modifica_producto()

    On Error GoTo cmd673_err

    btipo = "" & dmytablex.Fields("tipo")
    bproducto = "" & dmytablex.Fields("producto")
    bdescripcio = "" & dmytablex.Fields("descripcio")
    bunidad = "" & dmytablex.Fields("unidad")
    bfactor = "" & dmytablex.Fields("factor")
    bcantidad = "" & dmytablex.Fields("cantidad")
    bprecio = "" & dmytablex.Fields("precio")
    btotal = "" & dmytablex.Fields("total")
    modifica_producto = 1
    'btipo.SetFocus
    Exit Function
cmd673_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Function

End Function

