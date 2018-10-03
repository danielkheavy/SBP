VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frm_ESunat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enviar Datos en Sunat"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_salir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   8520
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmd_EDatos 
      Caption         =   "Envia Datos"
      Height          =   360
      Left            =   8400
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lbl_datos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "frm_ESunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_EDatos_Click()

    Dim salida           As Boolean

    Dim llego            As Boolean

    Dim my_cantidad_file As Integer

    Dim con_internet     As Boolean

    cmd_EDatos.Enabled = False

    Call read_caja(my_caja)
    Call Datos_Empresa(my_struc_datos_empresa(), my_caja, salida, 0)

    If my_struc_datos_empresa(0).esunat = "M" Then
        Call control_llegada_file(my_cantidad_file, my_caja)

        If my_cantidad_file = 0 Then
            MsgBox "No existe Datos para enviar a Sunat ", 48, "Aviso"
        Else
            ProgressBar1.Visible = True
            Call Enviar_Sunat(salida, my_cantidad_file, my_caja)
            Call bck_en_d_firmado_envia_POR_ENVIAR
            frm_ESunat.ProgressBar1.Value = (100 / 100 * 100)
            frm_ESunat.lbl_datos = "Elaborando Envio Sunat " & my_cantidad_file & "   Archivos al.." & ((100 / 100) * 100) & "%"
            MsgBox "Proceso Realizado ", 48, "Aviso"

        End If

    End If

End Sub

Private Sub cmd_salir_Click()
    Unload Me

End Sub

