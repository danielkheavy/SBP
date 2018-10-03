VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form tcajado 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tabla de Cajas"
   ClientHeight    =   9120
   ClientLeft      =   150
   ClientTop       =   120
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Frame2"
      Height          =   8055
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   12135
      Begin VB.TextBox cierres 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   105
         Top             =   7440
         Width           =   1575
      End
      Begin VB.TextBox moneda 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   62
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox tipodefa 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   61
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox bodega 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   2
         TabIndex        =   60
         Top             =   4920
         Width           =   855
      End
      Begin VB.TextBox vendedor 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   59
         Top             =   5280
         Width           =   375
      End
      Begin VB.TextBox cliente 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   58
         Top             =   5640
         Width           =   375
      End
      Begin VB.TextBox congela 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   11
         TabIndex        =   57
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox portbala 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   10
         TabIndex        =   56
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox actbala 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   55
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox flag 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   54
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox serieti 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   53
         Top             =   6000
         Width           =   3015
      End
      Begin VB.TextBox catipo 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   52
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox capuerto 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   30
         TabIndex        =   51
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox siventa 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   50
         Top             =   6360
         Width           =   1215
      End
      Begin VB.TextBox terminal 
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
         TabIndex        =   49
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox noprecio 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   48
         Top             =   6720
         Width           =   375
      End
      Begin VB.TextBox redondeo 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   47
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox local1 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   46
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox listap 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   2
         TabIndex        =   45
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox nosaldo 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   44
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox odpuerto 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   30
         TabIndex        =   43
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox hod 
         BackColor       =   &H00C0C0C0&
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
         Left            =   9600
         MaxLength       =   30
         TabIndex        =   42
         Top             =   4200
         Width           =   255
      End
      Begin VB.TextBox eccola 
         BackColor       =   &H00C0C0C0&
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
         Left            =   9600
         MaxLength       =   30
         TabIndex        =   41
         Top             =   4560
         Width           =   255
      End
      Begin VB.TextBox ecpuerto 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   30
         TabIndex        =   40
         Top             =   4560
         Width           =   1215
      End
      Begin VB.TextBox repite 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   39
         Top             =   4920
         Width           =   375
      End
      Begin VB.TextBox T1 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   2
         TabIndex        =   38
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox t2 
         BackColor       =   &H00C0C0C0&
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
         Left            =   8040
         MaxLength       =   2
         TabIndex        =   37
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox t3 
         BackColor       =   &H00C0C0C0&
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
         Left            =   8640
         MaxLength       =   2
         TabIndex        =   36
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox t4 
         BackColor       =   &H00C0C0C0&
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
         Left            =   9240
         MaxLength       =   2
         TabIndex        =   35
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox t5 
         BackColor       =   &H00C0C0C0&
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
         Left            =   9840
         MaxLength       =   2
         TabIndex        =   34
         Top             =   7440
         Width           =   615
      End
      Begin VB.TextBox t0 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   33
         Top             =   7080
         Width           =   615
      End
      Begin VB.TextBox seriehd 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   60
         TabIndex        =   32
         Top             =   5280
         Width           =   2415
      End
      Begin VB.TextBox puertocie 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   60
         TabIndex        =   31
         Top             =   5640
         Width           =   1695
      End
      Begin VB.TextBox tipocie 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   30
         Top             =   6000
         Width           =   375
      End
      Begin VB.TextBox decimal1 
         BackColor       =   &H00C0C0C0&
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
         Left            =   8640
         MaxLength       =   1
         TabIndex        =   29
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox pm 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   28
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox parqueo 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   27
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox sentido 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   26
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox video 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   25
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox segundo 
         BackColor       =   &H00C0C0C0&
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
         MaxLength       =   2
         TabIndex        =   24
         Top             =   3840
         Width           =   375
      End
      Begin VB.TextBox detraccion 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   23
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox hdetraccio 
         BackColor       =   &H00C0C0C0&
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
         Top             =   2400
         Width           =   375
      End
      Begin VB.TextBox deshab 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   21
         Top             =   6720
         Width           =   375
      End
      Begin VB.TextBox colacie 
         BackColor       =   &H00C0C0C0&
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
         Left            =   9360
         MaxLength       =   1
         TabIndex        =   20
         Top             =   5640
         Width           =   495
      End
      Begin VB.TextBox vecocaja 
         BackColor       =   &H00C0C0C0&
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
         Top             =   7080
         Width           =   615
      End
      Begin VB.TextBox apertura 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   18
         Top             =   6360
         Width           =   375
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
         Height          =   375
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   15
         Top             =   240
         Width           =   855
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
         TabIndex        =   14
         Top             =   600
         Width           =   7695
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   10560
         Picture         =   "tcajado.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Imprimir todo"
         Top             =   1320
         Width           =   1470
      End
      Begin VB.CommandButton cmdGuardar 
         BackColor       =   &H00808080&
         Caption         =   "&Guardar"
         Height          =   975
         Left            =   10560
         Picture         =   "tcajado.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label55 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cierres Correlativo"
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
         Left            =   120
         TabIndex        =   106
         Top             =   7440
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
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
         Left            =   120
         TabIndex        =   104
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TipoDocDefault"
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
         TabIndex        =   103
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Almacen Trab."
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
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vendedor"
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
         TabIndex        =   101
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cliente Obliga.."
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
         TabIndex        =   100
         Top             =   5640
         Width           =   2175
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Congela-Num."
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
         Left            =   5880
         TabIndex        =   99
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PuertoBalanza"
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
         Left            =   5880
         TabIndex        =   98
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ActivoBalanza"
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
         Left            =   5880
         TabIndex        =   97
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Flag-Cantidad"
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
         Left            =   5880
         TabIndex        =   96
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label30 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie Ticketera"
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
         Top             =   6000
         Width           =   2175
      End
      Begin VB.Label Label31 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.Star 1.Epson"
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
         Left            =   5880
         TabIndex        =   94
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label32 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PuertoCajon"
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
         Left            =   5880
         TabIndex        =   93
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label33 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ventas <="
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
         TabIndex        =   92
         Top             =   6360
         Width           =   2175
      End
      Begin VB.Label Label37 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C.aja T.erminal"
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
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label39 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hacer Nota Vta"
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
         Left            =   3480
         TabIndex        =   90
         Top             =   6360
         Width           =   1815
      End
      Begin VB.Label Label40 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Permite Precio=0"
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
         Left            =   120
         TabIndex        =   89
         Top             =   6720
         Width           =   2175
      End
      Begin VB.Label Label41 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Redeondeo"
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
         Left            =   5880
         TabIndex        =   88
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label42 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Local Caja"
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
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label Label44 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ListaPreciosNro"
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
         Left            =   5880
         TabIndex        =   86
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label45 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "01 o 02"
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
         Left            =   8640
         TabIndex        =   85
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label46 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NoVtaSdo<=0"
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
         Left            =   5880
         TabIndex        =   84
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label47 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Puerto.Ord.desp."
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
         Left            =   5880
         TabIndex        =   83
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label48 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Habilitado"
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
         Left            =   8640
         TabIndex        =   82
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label49 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cola"
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
         Left            =   8640
         TabIndex        =   81
         Top             =   4560
         Width           =   975
      End
      Begin VB.Label Label50 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Puerto.EstadoCta"
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
         Left            =   5880
         TabIndex        =   80
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label51 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NoRepitencia"
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
         Left            =   5880
         TabIndex        =   79
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label Label52 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Terminales Permitidos"
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
         Left            =   5880
         TabIndex        =   78
         Top             =   7440
         Width           =   1575
      End
      Begin VB.Label Label53 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Habilitar Terminales"
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
         Left            =   5880
         TabIndex        =   77
         Top             =   7080
         Width           =   1575
      End
      Begin VB.Label Label54 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie DiscoDuro"
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
         Left            =   5880
         TabIndex        =   76
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label Label56 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PuertoCierre"
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
         Left            =   5880
         TabIndex        =   75
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label Label57 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.Star 1.Epson"
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
         Left            =   5880
         TabIndex        =   74
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label Label58 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Decimal"
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
         Left            =   7800
         TabIndex        =   73
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label59 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Activar precio MInimo"
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
         Left            =   120
         TabIndex        =   72
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label60 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Solo Parking"
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
         Left            =   120
         TabIndex        =   71
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label61 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sentido"
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
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label62 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Video"
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
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label63 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Segundos"
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
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label65 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nro.Detraccion"
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
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label66 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "HabilitaDet"
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
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label67 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Deshabilitar Caja"
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
         Left            =   5880
         TabIndex        =   65
         Top             =   6720
         Width           =   1575
      End
      Begin VB.Label Label68 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ve Costos en Caja"
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
         Left            =   120
         TabIndex        =   64
         Top             =   7080
         Width           =   2175
      End
      Begin VB.Label Label69 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Necesita Apertura"
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
         Left            =   5880
         TabIndex        =   63
         Top             =   6360
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Caja"
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
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcio"
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
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   12435
      TabIndex        =   2
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
         Picture         =   "tcajado.frx":1194
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
         Picture         =   "tcajado.frx":23A6
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
         Picture         =   "tcajado.frx":35B8
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
         Picture         =   "tcajado.frx":47CA
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
         Picture         =   "tcajado.frx":59DC
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Nuevo registro"
         Top             =   0
         Width           =   735
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
            DataField       =   "Caja"
            Caption         =   "Caja"
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
   Begin VB.Menu dlo333 
      Caption         =   "&Seriales"
   End
   Begin VB.Menu dlo132 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcajado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txempre As New ADODB.Recordset

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
    caja.Enabled = True
    caja = ""
    caja.SetFocus

End Sub

Private Sub bo712_Click()

    Dim buf As String

    On Error GoTo cmd656_err

    buf = txempre.Fields("caja")

    If Frame2.Visible = True Then
        dbGrid1.SetFocus
        Exit Sub

    End If

    If MsgBox("Desea Borra " + txempre.Fields("caja"), 1, "Aviso") <> 1 Then
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

Private Sub CAJA_KeyPress(KeyAscii As Integer)

    Dim found As Integer

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    If KeyAscii = 27 Then
        dlo132_Click
        Exit Sub

    End If

    If Len(caja) = 0 Then Exit Sub

End Sub

Private Sub caja1_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
    descripcio.SetFocus

End Sub

Private Sub caja1_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = &H26 Then
        If caja.Enabled = True Then
            caja.SetFocus

        End If

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

    If opcion1 = "1" Then  'bodega
        If Len(buffer) = 0 Then
            cad = "SELECT * from parameca    "

        End If

        If Len(buffer) > 0 Then
            cad = "SELECT *  from parameca   where  " & Combo1 & " like '" & buffer & "%'"

        End If

        If txempre.State = 1 Then txempre.Close
        txempre.Open cad, cn, adOpenStatic, adLockOptimistic
        Set dbGrid1.DataSource = txempre
        dbGrid1.columns(0).Width = 4000
        dbGrid1.columns(1).Width = 2000

        'dbGrid1.Columns(2).Width = 2000
        If txempre.RecordCount > 0 Then
            dbGrid1.SetFocus

        End If

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

        'caja = dbGrid1.Columns(1)
        'Frame1.Visible = False
        'Frame1.Enabled = False
        'caja.SetFocus
        'caja_KeyPress 13
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
    reporgen.NAMETABLA = "parameca"
    reporgen.Show 1

End Sub

Private Sub dlo132_Click()

    If Frame2.Visible = True Then
        habilita 0
        Frame2.Visible = False
        dbGrid1.Enabled = True
        Exit Sub

    End If

    tcajado.Hide
    Unload tcajado

End Sub

Private Sub dlo333_Click()

    On Error GoTo cmd6733_err

    'tcajaser.caja = txempre.Fields("caja")
    'tcajaser.Show 1
    Exit Sub
cmd6733_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub f8443_Click()

    Dim buf As String

    On Error GoTo cmd456_err

    buf = txempre.Fields("caja")

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
    caja.Enabled = False
    descripcio.SetFocus
    Exit Sub
cmd456_err:
    MsgBox "Seleccione un dato ", 48, "Aviso"
    Exit Sub

End Sub

Private Sub fjh433_Click()

    Dim buf As String

    On Error GoTo cmd556_err

    buf = txempre.Fields("caja")

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
    caja.Enabled = False
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
    Combo1.AddItem "descripcio"
    Combo1.AddItem "caja"
    Combo1.ListIndex = 0

End Sub

Sub inicializa()
    apertura = ""
    vecocaja = ""
    colacie = ""
    deshab = ""
    hdetraccio = ""
    detraccion = ""
    video = ""
    segundo = ""
    sentido = ""
    parqueo = ""
    pm = ""
    decimal1 = ""
    puertocie = ""
    tipocie = ""
    cierres = ""
    seriehd = ""
    t0 = ""
    t1 = ""
    t2 = ""
    t3 = ""
    t4 = ""
    t5 = ""

    repite = ""
    ecpuerto = ""
    eccola = ""
    odpuerto = ""
    hod = ""
    nosaldo = ""
    listap = ""
    local1 = ""

    noprecio = ""
    terminal = "C"
    redondeo = ""

    siventa = ""
    capuerto = ""
    catipo = ""

    FLAG = ""

    portbala = ""
    actbala = ""
    congela = ""
    cliente = ""
    vendedor = ""
    bodega = ""

    descripcio = ""
    moneda = ""
    tipodefa = ""

End Sub

Sub pone_registro()
    apertura = "" & txempre.Fields("apertura")
    vecocaja = "" & txempre.Fields("vecocaja")
    colacie = "" & txempre.Fields("colacie")
    deshab = "" & txempre.Fields("deshab")
    hdetraccio = "" & txempre.Fields("hdetraccio")
    detraccion = "" & txempre.Fields("detraccion")

    segundo = "" & txempre.Fields("segundo")
    video = "" & txempre.Fields("video")
    sentido = "" & txempre.Fields("sentido")
    parqueo = "" & txempre.Fields("parqueo")
    decimal1 = "" & txempre.Fields("decimal")
    tipocie = "" & txempre.Fields("tipocie")
    puertocie = "" & txempre.Fields("puertocie")
    cierres = "" & txempre.Fields("cierres")
    seriehd = "" & txempre.Fields("seriehd")
    pm = "" & txempre.Fields("pm")

    t0 = "" & txempre.Fields("t0")
    t1 = "" & txempre.Fields("t1")
    t2 = "" & txempre.Fields("t2")
    t3 = "" & txempre.Fields("t3")
    t4 = "" & txempre.Fields("t4")
    t5 = "" & txempre.Fields("t5")

    repite = "" & txempre.Fields("repite")
    odpuerto = "" & txempre.Fields("odpuerto")
    hod = "" & txempre.Fields("hod")
    ecpuerto = "" & txempre.Fields("ecpuerto")
    eccola = "" & txempre.Fields("eccola")

    listap = "" & txempre.Fields("listap")
    nosaldo = "" & txempre.Fields("nosaldo")

    local1 = "" & txempre.Fields("local")
    redondeo = "" & txempre.Fields("redondeo")
    noprecio = "" & txempre.Fields("noprecio")
    terminal = "" & txempre.Fields("terminal")

    siventa = "" & txempre.Fields("siventa")
    capuerto = "" & txempre.Fields("capuerto")
    catipo = "" & txempre.Fields("catipo")
    serieti = "" & txempre.Fields("serieti")
    FLAG = "" & txempre.Fields("flag")

    actbala = "" & txempre.Fields("actbala")
    portbala = "" & txempre.Fields("portbala")
    congela = "" & txempre.Fields("congela")
    cliente = "" & txempre.Fields("cliente")
    vendedor = "" & txempre.Fields("vendedor")
    bodega = "" & txempre.Fields("bodega")

    tipodefa = "" & txempre.Fields("tipodefa")

    moneda = "" & txempre.Fields("moneda")
    caja = "" & txempre.Fields("caja")
    descripcio = "" & txempre.Fields("descripcio")

End Sub

Sub grabando()

    txempre.Fields("apertura") = apertura
    txempre.Fields("vecocaja") = vecocaja
    txempre.Fields("colacie") = colacie
    txempre.Fields("hdetraccio") = hdetraccio
    txempre.Fields("detraccion") = Val(detraccion)
    txempre.Fields("deshab") = deshab

    txempre.Fields("video") = video
    txempre.Fields("segundo") = Val(segundo)
    txempre.Fields("sentido") = sentido
    txempre.Fields("decimal") = decimal1
    txempre.Fields("tipocie") = tipocie
    txempre.Fields("puertocie") = puertocie
    txempre.Fields("cierres") = cierres
    txempre.Fields("seriehd") = seriehd
    txempre.Fields("pm") = pm
    txempre.Fields("parqueo") = parqueo

    txempre.Fields("t0") = t0
    txempre.Fields("t1") = t1
    txempre.Fields("t2") = t2
    txempre.Fields("t3") = t3
    txempre.Fields("t4") = t4
    txempre.Fields("t5") = t5

    txempre.Fields("repite") = repite
    txempre.Fields("ecpuerto") = ecpuerto
    txempre.Fields("eccola") = eccola

    txempre.Fields("odpuerto") = odpuerto
    txempre.Fields("hod") = hod
    txempre.Fields("nosaldo") = nosaldo
    txempre.Fields("listap") = listap
    txempre.Fields("local") = local1
    txempre.Fields("redondeo") = redondeo
    txempre.Fields("terminal") = terminal
    txempre.Fields("noprecio") = noprecio

    txempre.Fields("capuerto") = capuerto  'gaveta
    txempre.Fields("catipo") = catipo
    txempre.Fields("serieti") = serieti
    txempre.Fields("flag") = FLAG
    txempre.Fields("siventa") = Val(siventa)

    txempre.Fields("actbala") = actbala
    txempre.Fields("portbala") = portbala
    txempre.Fields("congela") = congela
    txempre.Fields("vendedor") = vendedor
    txempre.Fields("bodega") = bodega
    txempre.Fields("cliente") = cliente

    txempre.Fields("tipodefa") = tipodefa
    txempre.Fields("caja") = caja
    txempre.Fields("moneda") = moneda

    txempre.Fields("descripcio") = descripcio

End Sub

Private Sub grba1_Click()

End Sub

Private Sub descripcio_KeyPress(KeyAscii As Integer)

    If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub

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
        If Len(caja) = 0 Then
            caja.SetFocus
            Exit Function

        End If

        rbusca.Open "select caja from parameca where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

        If rbusca.RecordCount > 0 Then
            rbusca.Close
            MsgBox "Ya existe caja ", 48, "Aviso"
            Exit Function

        End If

        txempre.AddNew
        txempre.Fields("caja") = caja
        grabando
        txempre.Update
        dlo132_Click
        Exit Function

    End If

    If Frame2.Caption = "Modifica" Then
        txempre.Fields("caja") = caja
        grabando
        txempre.Update
        dlo132_Click
        Exit Function

    End If

End Function

Function valida()

    'If Len(caja) = 0 Then
    '   caja.SetFocus
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
        dlo333.Enabled = True
            
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
        dlo333.Enabled = False
            
    End If
      
End Sub

