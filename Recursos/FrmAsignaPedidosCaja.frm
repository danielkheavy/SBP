VERSION 5.00
Object = "{5C6863A4-877B-4EF1-9BD4-A17AD61FBEDB}#1.0#0"; "ChamaleonButton.ocx"
Begin VB.Form FrmAsignaPedidoCaja 
   Caption         =   "Form2"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4020
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Height          =   7695
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   12135
      Begin VB.TextBox t17 
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
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox t18 
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
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox t19 
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
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox t20 
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
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton Command2 
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
         Height          =   1125
         Left            =   8400
         MaskColor       =   &H00E0E0E0&
         Picture         =   "FrmAsignaPedidosCaja.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Grabar registro"
         Top             =   1440
         Width           =   1275
      End
      Begin VB.TextBox t16 
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
         Left            =   5160
         MaxLength       =   2
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox t15 
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
         Left            =   5160
         MaxLength       =   2
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox t14 
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
         Left            =   5160
         MaxLength       =   2
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox t13 
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
         Left            =   5160
         MaxLength       =   2
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox t12 
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
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox t11 
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
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox t10 
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
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox t9 
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
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox t8 
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
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox t7 
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
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox t6 
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
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox t5 
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
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox t4 
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
         Left            =   960
         MaxLength       =   2
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox t3 
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
         Left            =   960
         MaxLength       =   2
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox t2 
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
         Left            =   960
         MaxLength       =   2
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox t1 
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
         Left            =   960
         MaxLength       =   2
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox caja 
         Enabled         =   0   'False
         Height          =   495
         Left            =   1800
         TabIndex        =   2
         Text            =   "caja"
         Top             =   240
         Width           =   855
      End
      Begin ChamaleonButton.ChameleonBtn ChaCERRAR 
         Height          =   585
         Left            =   8520
         TabIndex        =   5
         Top             =   2760
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1032
         BTYPE           =   4
         TX              =   "CERRAR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   4210752
         BCOLO           =   4210752
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmAsignaPedidosCaja.frx":1212
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label71"
         Height          =   195
         Left            =   8880
         TabIndex        =   4
         Top             =   7680
         Width           =   555
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   11640
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CAJA:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label68 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Autorizados"
         Height          =   495
         Left            =   960
         TabIndex        =   1
         Top             =   1080
         Width           =   4935
      End
   End
End
Attribute VB_Name = "FrmAsignaPedidoCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''26/10/2017 Listas de ciertas caja para cobrar pedido

Option Explicit

Dim mytablel As New ADODB.Recordset

Dim txempre  As New ADODB.Recordset

Private Sub ChaCERRAR_Click()
    FrmAsignaPedidoCaja.Hide
    Unload FrmAsignaPedidoCaja

End Sub

Private Sub Command2_Click()

    Dim mytabley As New ADODB.Recordset

    If mytabley.State = 1 Then mytabley.Close

    mytabley.Open "SELECT * from usercajapedido where caja='" & caja & "' ", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        mytabley.AddNew
        mytabley.Fields("caja") = caja
       
        mytabley.Fields("p1") = t1
        mytabley.Fields("p2") = t2
        mytabley.Fields("p3") = t3
        mytabley.Fields("p4") = t4
    
        mytabley.Fields("p5") = t5
        mytabley.Fields("p6") = t6
        mytabley.Fields("p7") = t7
        mytabley.Fields("p8") = t8
       
        mytabley.Fields("p9") = t9
        mytabley.Fields("p10") = t10
        mytabley.Fields("p11") = t11
        mytabley.Fields("p12") = t12
       
        mytabley.Fields("p13") = t13
        mytabley.Fields("p14") = t14
        mytabley.Fields("p15") = t15
        mytabley.Fields("p16") = t16
       
        mytabley.Fields("p17") = t17
        mytabley.Fields("p18") = t18
        mytabley.Fields("p19") = t19
        mytabley.Fields("p20") = t20
       
        mytabley.Update
    Else
       
        mytabley.Fields("p1") = t1
        mytabley.Fields("p2") = t2
        mytabley.Fields("p3") = t3
        mytabley.Fields("p4") = t4
    
        mytabley.Fields("p5") = t5
        mytabley.Fields("p6") = t6
        mytabley.Fields("p7") = t7
        mytabley.Fields("p8") = t8
       
        mytabley.Fields("p9") = t9
        mytabley.Fields("p10") = t10
        mytabley.Fields("p11") = t11
        mytabley.Fields("p12") = t12
       
        mytabley.Fields("p13") = t13
        mytabley.Fields("p14") = t14
        mytabley.Fields("p15") = t15
        mytabley.Fields("p16") = t16
              
        mytabley.Fields("p17") = t17
        mytabley.Fields("p18") = t18
        mytabley.Fields("p19") = t19
        mytabley.Fields("p20") = t20
       
        mytabley.Update

    End If

    MsgBox "Proceso Realizado con exito", 48, "Aviso"
    
    mytabley.Close

End Sub

Private Sub Form_Activate()
    Label71_Click

End Sub

Private Sub Label71_Click()

    Dim mytablexyz As New ADODB.Recordset
 
    mytablexyz.Open "select *from usercajapedido where caja='" & caja & "'", cn, adOpenStatic, adLockOptimistic

    If mytablexyz.RecordCount > 0 Then
        t1 = mytablexyz.Fields("p1")
        t2 = mytablexyz.Fields("p2")
        t3 = mytablexyz.Fields("p3")
        t4 = mytablexyz.Fields("p4")
      
        t5 = mytablexyz.Fields("p5")
        t6 = mytablexyz.Fields("p6")
        t7 = mytablexyz.Fields("p7")
        t8 = mytablexyz.Fields("p8")
      
        t9 = mytablexyz.Fields("p9")
        t10 = mytablexyz.Fields("p10")
        t11 = mytablexyz.Fields("p11")
        t12 = mytablexyz.Fields("p12")
      
        t13 = mytablexyz.Fields("p13")
        t14 = mytablexyz.Fields("p14")
        t15 = mytablexyz.Fields("p15")
        t16 = mytablexyz.Fields("p16")
      
        t17 = mytablexyz.Fields("p17")
        t18 = mytablexyz.Fields("p18")
        t19 = mytablexyz.Fields("p19")
        t20 = mytablexyz.Fields("p20")
      
    End If

    mytablexyz.Close

End Sub

'''26/10/2017 Listas de ciertas caja para cobrar pedido

