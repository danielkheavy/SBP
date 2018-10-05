VERSION 5.00
Begin VB.Form FrmLimpiaTablas 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Borra registro de Tablas"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8850
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
   ScaleHeight     =   5475
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkClientes 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Borra Todas los CLIENTES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   4800
      TabIndex        =   9
      Top             =   3240
      Width           =   3855
   End
   Begin VB.CheckBox chkProveedores 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Borra Todas los PROVEEDORES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   4800
      TabIndex        =   8
      Top             =   2520
      Width           =   3855
   End
   Begin VB.CheckBox chkRecetas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Borra Todas las RECETAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   720
      TabIndex        =   7
      Top             =   2520
      Width           =   3255
   End
   Begin VB.CheckBox chkalmacen 
      BackColor       =   &H00C0C0C0&
      Caption         =   "BORRA ALMACEN"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   4800
      TabIndex        =   6
      Top             =   480
      Width           =   3855
   End
   Begin VB.CheckBox chkmarcas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Borra Todas las MARCAS"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   4800
      TabIndex        =   5
      Top             =   1800
      Width           =   3855
   End
   Begin VB.CheckBox chksubfamilias 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Borra Todas las SUBFAMILIAS"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   4800
      TabIndex        =   4
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CheckBox chkprecios 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Borra Todas los PRECIOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   3255
   End
   Begin VB.CheckBox chkproductos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Borra Todas los PRODUCTOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CheckBox chkfamilias 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Borra Todas las FAMILIAS"
      BeginProperty DataFormat 
         Type            =   4
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "EJECUTAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3600
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   0
      X2              =   8880
      Y1              =   5280
      Y2              =   5280
   End
End
Attribute VB_Name = "FrmLimpiaTablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'25/06/2018 Mejora LIMPIA REGISTRO DE TABLAS
Private Sub cmdCommand1_Click()

    Dim buf As String
  
    If chkClientes.Value = 0 And chkRecetas.Value = 0 And chkProveedores.Value = 0 And chkfamilias.Value = 0 And chkproductos.Value = 0 And chkprecios.Value = 0 And chksubfamilias.Value = 0 And chkalmacen.Value = 0 And chkmarcas.Value = 0 Then
        MsgBox ("Seleccione una opción"), vbCritical
        Exit Sub

    End If
  
    '
    If MsgBox("Se va a ELIMINAR los registros de las TABLAS SELECCCIONADA : está seguro???", vbExclamation + vbYesNo, "Eliminar") = vbYes Then
       
        'RECIBO
        If chkfamilias.Value = 1 Then
            buf = buf & "delete from familia "

        End If
    
        If chkprecios.Value = 1 Then
            buf = buf & "delete from precios "

        End If
    
        If chkproductos.Value = 1 Then
            buf = buf & "delete from producto delete from codclie delete from codprov delete from  productb"

        End If
     
        If chkmarcas.Value = 1 Then
            buf = buf & "delete from marca "

        End If
     
        If chksubfamilias.Value = 1 Then
            buf = buf & "delete from SUBFAMIL "

        End If
     
        If chkalmacen.Value = 1 Then
            buf = buf & "delete from almacen "

        End If
     
        If chkRecetas.Value = 1 Then
            buf = buf & "delete from recetas "

        End If
      
        If chkProveedores.Value = 1 Then
            buf = buf & "delete from proveedo delete from codprov "

        End If
              
        If chkClientes.Value = 1 Then
            buf = buf & "delete from clientes delete from codclie  "

        End If
     
        'EJECUTAMOS CONSULTA
        cn.Execute (buf)
        MsgBox ("Proceso Correcto")
     
    End If

End Sub

'25/06/2018 Mejora LIMPIA REGISTRO DE TABLAS
