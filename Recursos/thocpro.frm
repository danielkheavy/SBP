VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form thocpro 
   BackColor       =   &H00808080&
   Caption         =   "Cuadre de Producto x Turno"
   ClientHeight    =   10080
   ClientLeft      =   165
   ClientTop       =   -615
   ClientWidth     =   17265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   17265
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      Caption         =   "Control Inventario"
      Height          =   7335
      Left            =   30
      TabIndex        =   43
      Top             =   60
      Visible         =   0   'False
      Width           =   13935
      Begin VB.CommandButton Command3 
         Caption         =   "Limpia Pantalla"
         Height          =   495
         Left            =   1800
         TabIndex        =   47
         Top             =   6600
         Width           =   1455
      End
      Begin VB.CommandButton Command11 
         Caption         =   "CargaProducto"
         Height          =   495
         Left            =   120
         TabIndex        =   46
         Top             =   6600
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "LimpiaPantalla"
         Height          =   495
         Left            =   -1800
         TabIndex        =   45
         Top             =   6120
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dbgrid2 
         Height          =   5895
         Left            =   0
         TabIndex        =   44
         Top             =   600
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   10398
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         BeginProperty Column01 
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
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "Factor"
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
         BeginProperty Column04 
            DataField       =   "Saldoa"
            Caption         =   "SaldoA"
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
            DataField       =   "Entradas"
            Caption         =   "Entradas"
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
            DataField       =   "Salidas"
            Caption         =   "Salidas"
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
            DataField       =   "Saldo"
            Caption         =   "Saldo"
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
            DataField       =   "hotelcuadre"
            Caption         =   "hotelcuadre"
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
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3644.788
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column08 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Limpieza"
      Height          =   6495
      Left            =   15
      TabIndex        =   40
      Top             =   60
      Visible         =   0   'False
      Width           =   13935
      Begin VB.CommandButton Command6 
         Caption         =   "Refresca"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   6000
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid dbgrid4 
         Height          =   5415
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   9551
         _Version        =   393216
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
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "habitacion"
            Caption         =   "Habitacion"
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
            DataField       =   "sabana"
            Caption         =   "Sabana"
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
            DataField       =   "Toalla"
            Caption         =   "Toalla"
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
            DataField       =   "Jabon"
            Caption         =   "Jabon"
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
            DataField       =   "Ph"
            Caption         =   "Ph"
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
            DataField       =   "cubrecama"
            Caption         =   "Cubrecama"
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
            DataField       =   "Frazada"
            Caption         =   "Frazada"
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
            DataField       =   "Fecha"
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
         BeginProperty Column08 
            DataField       =   "Hora"
            Caption         =   "Hora"
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
            DataField       =   "Empleado"
            Caption         =   "Responsable"
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
            DataField       =   "Limpia"
            Caption         =   "Id"
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
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Facturacion"
      Height          =   6495
      Left            =   30
      TabIndex        =   36
      Top             =   60
      Visible         =   0   'False
      Width           =   13815
      Begin VB.CommandButton Command9 
         Caption         =   "Refresca"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   5880
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid dbgrid5 
         Height          =   5415
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   9551
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "Habitacion"
            Caption         =   "Habitacion"
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
         BeginProperty Column05 
            DataField       =   "Fecha"
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
         BeginProperty Column06 
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
         BeginProperty Column07 
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   540.284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   3960
            EndProperty
         EndProperty
      End
      Begin VB.Label totalf 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   10440
         TabIndex        =   39
         Top             =   5880
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ServicioHabitacion"
      Height          =   6255
      Left            =   0
      TabIndex        =   32
      Top             =   30
      Visible         =   0   'False
      Width           =   13935
      Begin VB.CommandButton Command7 
         Caption         =   "Refresca"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   5400
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid dbgrid3 
         Height          =   5055
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   8916
         _Version        =   393216
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
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "Habitacion"
            Caption         =   "Habitacion"
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
            DataField       =   "Arribofecha"
            Caption         =   "Fechae"
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
            DataField       =   "arribofechaf"
            Caption         =   "Fechas"
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
            DataField       =   "arribohora"
            Caption         =   "Horae"
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
            DataField       =   "arribohoraf"
            Caption         =   "Horas"
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
            DataField       =   "Huesped"
            Caption         =   "Huesped"
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
            DataField       =   "tipocodigoh"
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
         BeginProperty Column07 
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
         BeginProperty Column08 
            DataField       =   "Categoria"
            Caption         =   "Categoria"
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
            DataField       =   "Noches"
            Caption         =   "Permanencia"
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
         BeginProperty Column11 
            DataField       =   "Checkin"
            Caption         =   "CheckIn"
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
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2940.095
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   450.142
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   705.26
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
         EndProperty
      End
      Begin VB.Label xtotal 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   10440
         TabIndex        =   35
         Top             =   5400
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Frame2"
      Height          =   9015
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   13695
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Guardar"
         Height          =   1095
         Left            =   8040
         Picture         =   "thocpro.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   1470
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Cerrar"
         Height          =   1065
         Left            =   6600
         Picture         =   "thocpro.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Imprimir todo"
         Top             =   240
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
         TabIndex        =   23
         Top             =   1680
         Width           =   4095
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
         MaxLength       =   1
         TabIndex        =   22
         Top             =   600
         Width           =   975
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
         TabIndex        =   21
         Top             =   960
         Width           =   1935
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
         TabIndex        =   20
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox hotelcuadre 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox nturno 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox nvendedor 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox nestado 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2040
         Width           =   1575
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
         MaxLength       =   7
         TabIndex        =   15
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         TabIndex        =   31
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label1 
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
         TabIndex        =   30
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   29
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaTRabajo"
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
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Id"
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
         TabIndex        =   26
         Top             =   2040
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Inventario"
      Height          =   735
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ServicioHabit"
      Height          =   735
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "LimpiezaHab"
      Height          =   735
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0080FF80&
      Caption         =   "Facturacion"
      Height          =   735
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4200
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   17205
      TabIndex        =   2
      Top             =   0
      Width           =   17265
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
         Picture         =   "thocpro.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Grabar registro"
         Top             =   0
         Width           =   735
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
         Left            =   3600
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
         Picture         =   "thocpro.frx":23A6
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
         Picture         =   "thocpro.frx":35B8
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
         Picture         =   "thocpro.frx":47CA
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
         Picture         =   "thocpro.frx":59DC
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
         ColumnCount     =   6
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "Turno"
            Caption         =   "Turno"
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
            DataField       =   "Vendedor"
            Caption         =   "Responsable"
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
         BeginProperty Column04 
            DataField       =   "hotelcuadre"
            Caption         =   "HotelCuadre"
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
            DataField       =   "estado"
            Caption         =   "Estado"
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
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
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
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "thocpro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txcuaho  As New ADODB.Recordset

Dim txcuahod As New ADODB.Recordset

Dim txcuahoh As New ADODB.Recordset

Dim txcuahol As New ADODB.Recordset

Dim txcuahof As New ADODB.Recordset

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
    hotelcuadre.Enabled = False
    hotelcuadre = ""
    turno.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    buf = "" & txcuaho.Fields("hotelcuadre")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + "" & "" & txcuaho.Fields("hotelcuadre"), 1, "Aviso") <> 1 Then
        Exit Sub

    End If

    txcuaho.Delete
    Command1_Click

    Exit Sub
cmd656_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub cmdAddEntry_Click()
    ajdu1_Click

End Sub

Private Sub cmdCerrar_Click()
    'If Frame3.Visible = True Then
    '   Frame3.Visible = False
    '   Exit Sub
    'End If
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

    'If Frame3.Visible = True Then
    '   Frame3.Visible = False
    'Exit Sub
    'End If

    found = grabar()

End Sub

Private Sub cmdPrint_Click()

    'djuer1_Click
End Sub

Private Sub cmdSave_Click()
    f8443_Click

End Sub

Private Sub Command10_Click()
    Frame4.Visible = False
    Frame3.Visible = False
    Frame5.Visible = True
    filtro_facturacion

End Sub

Private Sub Command11_Click()

    Dim mytablex As New ADODB.Recordset

    'If Val(hotelcuadre) = 0 Then
    '   MsgBox "Modo Modifica ", 48, "Aviso"
    '   Exit Sub
    'End If
    If MsgBox("Desea cargar Productos ", 1, "Aviso") <> 1 Then Exit Sub
    If txcuahod.State = 1 Then
        txcuahod.Close

    End If

    cn.Execute ("delete from hotelcuadred where hotelcuadre=" & Val(hotelcuadre))
    filtro_producto
    mytablex.Open "select * from producto ", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        txcuahod.AddNew
        txcuahod.Fields("hotelcuadre") = Val(hotelcuadre)
        txcuahod.Fields("producto") = Trim("" & mytablex.Fields("producto"))
        txcuahod.Fields("descripcio") = Trim("" & mytablex.Fields("descripcio"))
        txcuahod.Fields("unidad") = Trim("" & mytablex.Fields("unidad"))
        txcuahod.Fields("factor") = Val("" & mytablex.Fields("factor"))
        txcuahod.Update
        mytablex.MoveNext
    Loop
    mytablex.Close
    MsgBox "Proceso Terminado ", 48, "Aviso"

End Sub

Private Sub Command3_Click()

    'If Val(hotelcuadre) = 0 Then
    '   MsgBox "Modo Modifica ", 48, "Aviso"
    '   Exit Sub
    'End If
    If MsgBox("Desea Borrar todos los producto ", 1, "Aviso") <> 1 Then Exit Sub
    cn.Execute ("delete from hotelcuadred where hotelcuadre=" & Val(hotelcuadre))
    filtro_producto

End Sub

Private Sub Command4_Click()
    'If Val(hotelcuadre) = 0 Then
    '   MsgBox "Modo Modifica ", 48, "Aviso"
    '   Exit Sub
    'End If
    Frame3.Visible = False
    Frame4.Visible = False
    Frame5.Visible = False
    Frame6.Visible = True
    filtro_producto
    filtro_habitacion
    filtro_limpieza

End Sub

Private Sub Command5_Click()
    'If Val(hotelcuadre) = 0 Then
    '   MsgBox "Modo Modifica ", 48, "Aviso"
    '   Exit Sub
    'End If
    Frame4.Visible = False
    Frame3.Visible = True
    Frame5.Visible = False
    filtro_habitacion

End Sub

Private Sub Command6_Click()
    filtro_limpieza

End Sub

Private Sub Command7_Click()
    filtro_habitacion

End Sub

Private Sub Command8_Click()
    Frame5.Visible = False
    Frame3.Visible = False
    Frame4.Visible = True
    filtro_limpieza

End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub Command9_Click()
    filtro_facturacion

End Sub

Private Sub dbgrid2_BeforeColEdit(ByVal ColIndex As Integer, _
                                  ByVal KeyAscii As Integer, _
                                  Cancel As Integer)

    Select Case ColIndex

        Case 4
            'If Not IsNumeric(dbgrid2.columns(4)) Then
            'Cancel = True
            'Exit Sub
            'End If
            'Case 5
            'If Not IsNumeric(dbgrid2.columns(5)) Then
            'Cancel = True
            'Exit Sub
            'End If
            'Case 6
            'If Not IsNumeric(dbgrid2.columns(6)) Then
            'Cancel = True
            'Exit Sub
            'End If
       
            'If KeyAscii = 13 Then
            'MsgBox "BEFORECOLEDIT:" & dbgrid2.columns(4)
            'If dbGrid1.columns(1) <> "HOLA" Then
            '   MsgBox "error "
            '   Cancel = True
            '   Exit Sub
            'End If
            'End If
            
    End Select

End Sub

Private Sub dbgrid2_BeforeColUpdate(ByVal ColIndex As Integer, _
                                    OldValue As Variant, _
                                    Cancel As Integer)

    Dim sdx As Double

    Select Case ColIndex

        Case 4

            If Not IsNumeric(DBGrid2.columns(4)) Then
                Cancel = True
                Exit Sub

            End If

            sdx = Val("" & DBGrid2.columns(4)) + Val("" & DBGrid2.columns(5)) - Val("" & DBGrid2.columns(6))
            DBGrid2.columns(7) = sdx

        Case 5

            If Not IsNumeric(DBGrid2.columns(5)) Then
                Cancel = True
                Exit Sub

            End If

            sdx = Val("" & DBGrid2.columns(4)) + Val("" & DBGrid2.columns(5)) - Val("" & DBGrid2.columns(6))
            DBGrid2.columns(7) = sdx

        Case 6

            If Not IsNumeric(DBGrid2.columns(6)) Then
                Cancel = True
                Exit Sub

            End If

            sdx = Val("" & DBGrid2.columns(4)) + Val("" & DBGrid2.columns(5)) - Val("" & DBGrid2.columns(6))
            DBGrid2.columns(7) = sdx

            'If KeyAscii = 13 Then
            'MsgBox "BEFORECOLEDIT:" & dbgrid2.columns(4)
            'If dbGrid1.columns(1) <> "HOLA" Then
            '   MsgBox "error "
            '   Cancel = True
            '   Exit Sub
            'End If
            'End If
        Case Else
            Cancel = True
            Exit Sub
     
    End Select

End Sub

Private Sub dbgrid4_BeforeColUpdate(ByVal ColIndex As Integer, _
                                    OldValue As Variant, _
                                    Cancel As Integer)

    Dim sdx As Double

    Select Case ColIndex

        Case 1, 2, 3, 4, 5, 6

            'If KeyAscii = 13 Then
            'MsgBox "BEFORECOLEDIT:" & dbgrid2.columns(4)
            'If dbGrid1.columns(1) <> "HOLA" Then
            '   MsgBox "error "
            '   Cancel = True
            '   Exit Sub
            'End If
            'End If
        Case Else
            Cancel = True
            Exit Sub
     
    End Select

End Sub

Private Sub dk9893_Click()

    If Frame2.Visible = True Then Exit Sub
    reporgen.NAMETABLA = "hotelcuadre"
    reporgen.Show 1

End Sub

Sub prueba_reporte()

    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\hotelcuadreesproducto.rpt", "")
End Sub

Private Sub hotelcuadre_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(hotelcuadre) = 0 Then Exit Sub
    descripcio.SetFocus

End Sub

Private Sub Command1_Click()
    Frame1.Visible = True
    Frame1.Enabled = True
    'buffer = ""
    opcion1 = "1"
    ejecuta 1

End Sub

Sub ejecuta(sw As Integer)

    Dim cad As String

    If opcion1 = "1" Then  'bodega
        If Trim(Combo1) = 0 Then
            cad = "SELECT * from hotelcuadre  "
        Else
            cad = "SELECT * from hotelcuadre  where estado='" & estado & "'"

        End If

        If txcuaho.State = 1 Then txcuaho.Close
        txcuaho.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbGrid1.DataSource = txcuaho

        'dbGrid1.columns(0).Width = 4000
        'dbGrid1.columns(1).Width = 2000
        If txcuaho.RecordCount > 0 Then
            dbGrid1.SetFocus

        End If

    End If

End Sub

Private Sub dbgrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        'buffer.SetFocus
        Exit Sub

    End If

    If KeyCode = 13 Then

        'hotelcuadre = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'hotelcuadre.SetFocus
        'hotelcuadre_KeyPress 13
    End If

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

    thocpro.Hide
    Unload thocpro

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    buf = "" & txcuaho.Fields("hotelcuadre")

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
    hotelcuadre.Enabled = False
    turno.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    buf = "" & txcuaho.Fields("hotelcuadre")

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
    hotelcuadre.Enabled = False
    turno.SetFocus
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
    Combo1.AddItem "PENDIENTE"
    Combo1.AddItem "CERRADO"
    Combo1.ListIndex = 0
    carga_turno
    carga_vendedor

    nestado.Clear
    nestado.AddItem "ABIERTO"
    nestado.AddItem "CERRADO"
    nestado.ListIndex = 0

End Sub

Sub inicializa()
    turno = ""
    vendedor = ""
    fecha = Format(Now, "dd/mm/yyyy")
    descripcio = " "
    estado = "ABIERTO"

End Sub

Sub pone_registro()
    estado = Trim("" & txcuaho.Fields("estado"))
    turno = Trim("" & txcuaho.Fields("turno"))
    fecha = Trim("" & txcuaho.Fields("fecha"))
    vendedor = Trim("" & txcuaho.Fields("vendedor"))
    hotelcuadre = Trim("" & txcuaho.Fields("hotelcuadre"))
    descripcio = Trim("" & txcuaho.Fields("descripcio"))

End Sub

Sub grabando()
    txcuaho.Fields("estado") = Trim(estado)
    txcuaho.Fields("turno") = Trim(turno)
    txcuaho.Fields("fecha") = Trim(fecha)
    txcuaho.Fields("vendedor") = Trim(vendedor)
    'txcuaho.Fields("hotelcuadre") = Trim(hotelcuadre)
    txcuaho.Fields("descripcio") = Trim(descripcio)

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
        'If Len(hotelcuadre) = 0 Then
        'hotelcuadre.SetFocus
        'Exit Function
        'End If
        txcuaho.AddNew
        grabando
        txcuaho.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        grabando
        txcuaho.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    If Len(Trim(turno)) = 0 Then
        turno.SetFocus
        Exit Function

    End If

    If Not IsDate(fecha) Then
        fecha.SetFocus
        Exit Function

    End If

    If Len(Trim(vendedor)) = 0 Then
        vendedor.SetFocus
        Exit Function

    End If

    'If Len(descripcio) = 0 Then
    '   descripcio.SetFocus
    '   Exit Function
    'End If
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

    mytablex.Open "select * from archivo where menu='hotelcuadre' and   estado='S'", cn, adOpenStatic, adLockOptimistic

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
    mytablex.Open "select * from archivo where menu='hotelcuadre' and descripcio='" & Trim(buf) & "'", cn, adOpenStatic, adLockOptimistic

    If mytablex.RecordCount = 0 Then
        mytablex.Close

    End If

    'busca el reporte
    buf = mytablex.Fields("archivo")
    mytablex.Close
    'Call Reporter(CrystalReport1, globalpath & "\001d\06\reporte\" & buf, "")

End Sub

Sub carga_turno()

    Dim mytablex As New ADODB.Recordset

    nturno.Clear
    nturno.AddItem ""
    mytablex.Open "select * from turno order by turno", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        nturno.AddItem Trim("" & mytablex.Fields("turno"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    nturno.ListIndex = 0

End Sub

Sub carga_vendedor()

    Dim mytablex As New ADODB.Recordset

    nvendedor.Clear
    nvendedor.AddItem ""
    mytablex.Open "select * from vendedor order by nombre", cn, adOpenStatic, adLockOptimistic
    Do

        If mytablex.EOF Then Exit Do
        nvendedor.AddItem Trim("" & mytablex.Fields("nombre")) & "|" & Trim("" & mytablex.Fields("codigo"))
        mytablex.MoveNext
    Loop
    mytablex.Close
    nvendedor.ListIndex = 0

End Sub

Private Sub nestado_Click()
    estado = Trim(nestado)

End Sub

Private Sub nturno_Click()
    turno = Trim(nturno)

End Sub

Private Sub nvendedor_Click()
    vendedor = extra_loquesea1(nvendedor)

End Sub

Sub filtro_producto()

    If txcuahod.State = 1 Then
        txcuahod.Close

    End If

    txcuahod.Open "select * from hotelcuadred where hotelcuadre=" & Val("" & txcuaho.Fields("hotelcuadre")), cn, adOpenStatic, adLockOptimistic
    Set DBGrid2.DataSource = txcuahod

End Sub

Sub filtro_habitacion()

    Dim buf As String

    Dim sdx As Double

    If txcuahoh.State = 1 Then
        txcuahoh.Close

    End If

    sdx = 0
    buf = "select * from hotelcheckin where "
    buf = buf & "  hotelcuadre=" & Val("" & txcuaho.Fields("hotelcuadre")) & ""
    txcuahoh.Open buf, cn, adOpenStatic, adLockOptimistic
    Set dbgrid3.DataSource = txcuahoh
    Do

        If txcuahoh.EOF Then Exit Do
        sdx = sdx + Val("" & txcuahoh.Fields("total"))
        txcuahoh.MoveNext
    Loop
    xtotal = Format(sdx, "0.00")

End Sub

Sub filtro_limpieza()

    'Exit Sub
    If txcuahol.State = 1 Then
        txcuahol.Close

    End If

    txcuahol.Open "select * from hotellimpia where hotelcuadre=" & Val("" & txcuaho.Fields("hotelcuadre")), cn, adOpenStatic, adLockOptimistic
    Set DBGrid4.DataSource = txcuahol

End Sub

Sub filtro_facturacion()

    'Exit Sub
    Dim sdx As Double

    sdx = 0

    If txcuahof.State = 1 Then
        txcuahof.Close

    End If

    txcuahof.Open "select hotelcheckin.habitacion,hotelfactura.tipo,hotelfactura.serie,hotelfactura.numero,hotelfactura.total,hotelfactura.codigo,hotelfactura.nombre,hotelfactura.fecha from hotelfactura inner join hotelcheckin on hotelfactura.idcheckin=hotelcheckin.checkin and hotelfactura.hotelcuadre=" & Val("" & txcuaho.Fields("hotelcuadre")), cn, adOpenStatic, adLockOptimistic
    Set dbgrid5.DataSource = txcuahof
    Do

        If txcuahof.EOF Then Exit Do
        sdx = sdx + Val("" & txcuahof.Fields("total"))
        txcuahof.MoveNext
    Loop
   
    totalf = Format(sdx, "0.00")

End Sub
