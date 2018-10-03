VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ingclie 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Ingreso Clientes"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10530
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   10530
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label pedido 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5115
      Left            =   120
      Picture         =   "ingclie.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6690
   End
   Begin VB.Menu dd12 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "ingclie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dd12_Click()
    ingclie.Hide
    Unload ingclie

End Sub

Private Sub Form_Load()

    Dim sdx As Double

    sdx = busca_numero()
    pedido = "" & sdx

End Sub

Private Sub Image1_Click()

    Dim found As Integer

    found = prepara_achivo()

End Sub

Function busca_numero() As Double

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("parame")
    mytablex.Index = "codigo"
    mytablex.Seek "=", "01"

    If Not mytablex.NoMatch Then
        busca_numero = Val("" & mytablex.Fields("pocket")) + 1

    End If

    mytablex.Close

End Function

Function graba_numero()

    Dim mytablex As Table

    Set mytablex = mydbxglo.OpenTable("parame")
    mytablex.Index = "codigo"
    mytablex.Seek "=", "01"

    If Not mytablex.NoMatch Then
        mytablex.Edit
        mytablex.Fields("pocket") = Val(pedido)
        mytablex.Update
        graba_numero = 1

    End If

    mytablex.Close

End Function

Function valida_numero()

    Dim mytablex As Table

    Dim sdx      As Double

    Set mytablex = mydbxglo.OpenTable("Ppocket")
    mytablex.Index = "ppocket"
contix:
    mytablex.Seek "=", pedido

    If Not mytablex.NoMatch Then
        sdx = Val(pedido) + 1
        pedido = "" & sdx
        GoTo contix

    End If

    mytablex.Close

End Function

Function prepara_achivo()

    Dim found As Integer

    Dim I     As Integer

    Dim sdx   As Double

    found = valida_numero()
    found = valida_cliente()

    If found = 0 Then
        dd12_Click
        Exit Function

    End If

    found = graba_numero()
    FileName = globaldir & "\temporal\" & gusuario & ".txt"
    cerrar_archivo
    found = borra_nombre("" & FileName)
    Open FileName For Append As #1
    '------------------------------------
    found = formateaa("TICKET NRO " + pedido, 36, 2, 0)
    found = formateaa("URANIO", 36, 2, 0)
    found = formateaa("" & Format(Now, "dd/mm/yyyy") & "   " & "" & Format(Now, "hh:mm:ss"), 36, 2, 0)
    'found = formateaa("PRESENTAR ESTE TICKET", 36, 2, 0)
    'found = formateaa("AL VENDEDOR", 36, 2, 0)
    'For i = 1 To 2
    'found = formateaa(".", 36, 2, 0)
    'Next i
    '------------------------------------
    Close #1
    cerrar_archivo
    tipoletra = 13
    genver.file = globaldir & "\temporal\" & gusuario & ".txt"
    genver.Show 1
    Exit Function

End Function

Function valida_cliente()

    Dim mytablex As Table

    Dim found    As Integer

    On Error GoTo cmd1_err

    Set mytablex = mydbxglo.OpenTable("ppocket")
    mytablex.AddNew
    mytablex.Fields("pedido") = "" & pedido
    mytablex.Update
    mytablex.Close
    valida_cliente = 1
    Exit Function
cmd1_err:
    MsgBox "Error ,llame a servicio" + error$, 48, "Aviso"
    Exit Function

End Function

Function ir_ultimo(mytablex As Table)

    On Error GoTo cmd3_err

    mytablex.MoveLast
    ir_ultimo = 1
    Exit Function
cmd3_err:
    Exit Function

End Function

Private Sub ks2232_Click()

End Sub
