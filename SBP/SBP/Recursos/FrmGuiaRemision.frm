VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmGuiaRemision 
   Caption         =   "Documento Guia de Remision"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   660
   ClientWidth     =   12045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmGuiaRemision.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7560
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6840
      Left            =   45
      TabIndex        =   0
      Top             =   240
      Width           =   11910
      Begin VB.CommandButton CmdCancelarGuia 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   10710
         Picture         =   "FrmGuiaRemision.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Cancelar:"
         Top             =   270
         Width           =   810
      End
      Begin VB.CommandButton CmdGuarGuia 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   9915
         Picture         =   "FrmGuiaRemision.frx":0B82
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Guardar Guia"
         Top             =   270
         Width           =   810
      End
      Begin VB.TextBox TxtNotas 
         Appearance      =   0  'Flat
         Height          =   765
         Left            =   315
         TabIndex        =   21
         Top             =   5940
         Width           =   11055
      End
      Begin VB.TextBox TxtNumRef 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2895
         TabIndex        =   3
         Top             =   645
         Width           =   2655
      End
      Begin VB.TextBox TxtSerie 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         TabIndex        =   2
         Top             =   645
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   360
         Left            =   5655
         TabIndex        =   4
         Top             =   645
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   278790145
         CurrentDate     =   43368
         MaxDate         =   47848
         MinDate         =   43101
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3885
         Left            =   180
         TabIndex        =   8
         Top             =   1470
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   6853
         _Version        =   393216
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Datos de Envio"
         TabPicture(0)   =   "FrmGuiaRemision.frx":1328
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "LblTranspor"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "LblModalidadTras"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "LblDescrTras"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "LblMotivoTras"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "DTPicker2"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "TxtTranspor"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "CboModalidadTras"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "ChkTransPro"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "TxtDescTras"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "CboMotivoTras"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "CmdBuscarTrans"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).ControlCount=   12
         TabCaption(1)   =   "Destinatario"
         TabPicture(1)   =   "FrmGuiaRemision.frx":1344
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TxtDetinatario"
         Tab(1).Control(1)=   "LblDestinatario"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Proveedor"
         TabPicture(2)   =   "FrmGuiaRemision.frx":1360
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         TabCaption(3)   =   "Punto Partida"
         TabPicture(3)   =   "FrmGuiaRemision.frx":137C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "CmdBuscaProvi"
         Tab(3).Control(1)=   "CmdBuscaDepart"
         Tab(3).Control(2)=   "TxtDireccion"
         Tab(3).Control(3)=   "TxtUbigeoP"
         Tab(3).Control(4)=   "TxtDistritoP"
         Tab(3).Control(5)=   "TxtProvinciaP"
         Tab(3).Control(6)=   "TxtDepartamentoP"
         Tab(3).Control(7)=   "LblDireccion(2)"
         Tab(3).Control(8)=   "LblDistrito(0)"
         Tab(3).Control(9)=   "LblProvincia(0)"
         Tab(3).Control(10)=   "LblDeparta(0)"
         Tab(3).Control(11)=   "LblUbigeo(0)"
         Tab(3).ControlCount=   12
         TabCaption(4)   =   "Punto Llegada"
         TabPicture(4)   =   "FrmGuiaRemision.frx":1398
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Text5"
         Tab(4).Control(1)=   "Text4"
         Tab(4).Control(2)=   "Text3"
         Tab(4).Control(3)=   "Text2"
         Tab(4).Control(4)=   "Text1"
         Tab(4).Control(5)=   "LblDireccion(0)"
         Tab(4).Control(6)=   "LblDistrito(1)"
         Tab(4).Control(7)=   "LblProvincia(1)"
         Tab(4).Control(8)=   "LblDeparta(1)"
         Tab(4).Control(9)=   "LblUbigeo(1)"
         Tab(4).ControlCount=   10
         Begin VB.CommandButton CmdBuscaProvi 
            Height          =   510
            Left            =   -67245
            Picture         =   "FrmGuiaRemision.frx":13B4
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Buscar Transportista:"
            Top             =   2025
            Width           =   510
         End
         Begin VB.CommandButton CmdBuscaDepart 
            Height          =   510
            Left            =   -67260
            Picture         =   "FrmGuiaRemision.frx":1573
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Buscar Transportista:"
            Top             =   1485
            Width           =   510
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71715
            TabIndex        =   45
            Top             =   3180
            Width           =   4380
         End
         Begin VB.CommandButton CmdBuscarTrans 
            Height          =   510
            Left            =   6570
            Picture         =   "FrmGuiaRemision.frx":1732
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Buscar Transportista:"
            Top             =   2730
            Width           =   510
         End
         Begin VB.TextBox TxtDireccion 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71715
            TabIndex        =   42
            Top             =   3195
            Width           =   4380
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71715
            TabIndex        =   35
            Top             =   990
            Width           =   4380
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71715
            TabIndex        =   34
            Top             =   2085
            Width           =   4380
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71715
            TabIndex        =   33
            Top             =   1545
            Width           =   4380
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71715
            TabIndex        =   32
            Top             =   2640
            Width           =   4380
         End
         Begin VB.TextBox TxtUbigeoP 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71715
            TabIndex        =   27
            Top             =   990
            Width           =   4380
         End
         Begin VB.TextBox TxtDistritoP 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71715
            TabIndex        =   26
            Top             =   2085
            Width           =   4380
         End
         Begin VB.TextBox TxtProvinciaP 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71715
            TabIndex        =   25
            Top             =   1545
            Width           =   4380
         End
         Begin VB.TextBox TxtDepartamentoP 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71715
            TabIndex        =   24
            Top             =   2640
            Width           =   4380
         End
         Begin VB.TextBox TxtDetinatario 
            Appearance      =   0  'Flat
            Height          =   555
            Left            =   -74655
            TabIndex        =   22
            Top             =   1290
            Width           =   8505
         End
         Begin VB.ComboBox CboMotivoTras 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            ItemData        =   "FrmGuiaRemision.frx":18F1
            Left            =   195
            List            =   "FrmGuiaRemision.frx":1910
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1170
            Width           =   3285
         End
         Begin VB.TextBox TxtDescTras 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3585
            TabIndex        =   12
            Top             =   1170
            Width           =   4890
         End
         Begin VB.CheckBox ChkTransPro 
            Appearance      =   0  'Flat
            Caption         =   "Transbordo Programado"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   630
            Left            =   8730
            TabIndex        =   11
            Top             =   1050
            Width           =   1935
         End
         Begin VB.ComboBox CboModalidadTras 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   210
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   2055
            Width           =   3285
         End
         Begin VB.TextBox TxtTranspor 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   210
            TabIndex        =   9
            Top             =   2805
            Width           =   6285
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   360
            Left            =   3585
            TabIndex        =   14
            Top             =   2070
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   254738433
            CurrentDate     =   43368
            MaxDate         =   47848
            MinDate         =   43101
         End
         Begin VB.Label LblDireccion 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   -73170
            TabIndex        =   46
            Top             =   3240
            Width           =   1350
         End
         Begin VB.Label LblDireccion 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   -73170
            TabIndex        =   43
            Top             =   3255
            Width           =   1350
         End
         Begin VB.Label LblDistrito 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Distrito:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   -73050
            TabIndex        =   39
            Top             =   2685
            Width           =   1215
         End
         Begin VB.Label LblProvincia 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Provincia:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   -73185
            TabIndex        =   38
            Top             =   2130
            Width           =   1350
         End
         Begin VB.Label LblDeparta 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   -73590
            TabIndex        =   37
            Top             =   1590
            Width           =   1755
         End
         Begin VB.Label LblUbigeo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ubigeo:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   -72780
            TabIndex        =   36
            Top             =   1005
            Width           =   945
         End
         Begin VB.Label LblDistrito 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Distrito:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   -73035
            TabIndex        =   31
            Top             =   2685
            Width           =   1215
         End
         Begin VB.Label LblProvincia 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Provincia:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   -73170
            TabIndex        =   30
            Top             =   2130
            Width           =   1350
         End
         Begin VB.Label LblDeparta 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   -73560
            TabIndex        =   29
            Top             =   1590
            Width           =   1740
         End
         Begin VB.Label LblUbigeo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ubigeo:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   -72765
            TabIndex        =   28
            Top             =   1005
            Width           =   945
         End
         Begin VB.Label LblDestinatario 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destinatario:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -74655
            TabIndex        =   23
            Top             =   990
            Width           =   2700
         End
         Begin VB.Label LblMotivoTras 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Motivo Del Traslado:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   210
            TabIndex        =   19
            Top             =   885
            Width           =   2700
         End
         Begin VB.Label LblDescrTras 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción Traslado:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3585
            TabIndex        =   18
            Top             =   900
            Width           =   2835
         End
         Begin VB.Label LblModalidadTras 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Modalidad de Traslado:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   255
            TabIndex        =   17
            Top             =   1785
            Width           =   2970
         End
         Begin VB.Label LblTranspor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transportista:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   210
            TabIndex        =   16
            Top             =   2550
            Width           =   1890
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Inicio Traslado:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   3570
            TabIndex        =   15
            Top             =   1815
            Width           =   2970
         End
      End
      Begin VB.Label LblNotas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notas:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   315
         TabIndex        =   20
         Top             =   5640
         Width           =   810
      End
      Begin VB.Label LblFechaEmi 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Emisión:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   7
         Top             =   390
         Width           =   2295
      End
      Begin VB.Label LblSerie 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serie:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   135
         TabIndex        =   6
         Top             =   390
         Width           =   810
      End
      Begin VB.Label LblNumero 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2895
         TabIndex        =   5
         Top             =   375
         Width           =   945
      End
      Begin VB.Label LblFondo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Datos Guia de Remisión"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Top             =   0
         Width           =   3000
      End
   End
   Begin VB.Menu CmdSalir 
      Caption         =   "Salir"
      Index           =   0
   End
End
Attribute VB_Name = "FrmGuiaRemision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CboMotivoTras_LostFocus()
    TxtDescTras.Text = Mid(CboMotivoTras.Text, 4, 41)
End Sub

Private Sub CmdSalir_Click(Index As Integer)
    Unload Me

End Sub

Private Sub Form_Load()
'CboMotivoTras.Clear
CboMotivoTras.ListIndex = 0
End Sub

