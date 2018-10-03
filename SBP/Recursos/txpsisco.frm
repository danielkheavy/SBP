VERSION 5.00
Begin VB.Form txpsisco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Contabilidad Siscont"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   14445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   14445
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Filtrar 
      Caption         =   "Filtrar"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TipoDocumento"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu lfoo4545 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "txpsisco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Filtrar_Click()

    If Combo1 = "TipoDocumento" Then
        copia_tipo_documento

    End If

    If Combo1 = "Plan de Cuentas" Then
        siscont_cuenta

    End If

    If Combo1 = "Origenes" Then
        copia_origenes

    End If

End Sub

Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem ""
    Combo1.AddItem "TipoDocumento"
    Combo1.AddItem "Plan de Cuentas"
    Combo1.AddItem "Origenes"
    Combo1.AddItem ""
    Combo1.ListIndex = 0

    'siscont1415
    'tipo documento
    'mdh_td
    'tipo de cmabio
    'mdh_tc

End Sub

Private Sub lfoo4545_Click()
    txpsisco.Hide
    Unload txpsisco

End Sub

Sub copia_origenes()

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    cn.Execute ("delete from origen")
    Set mydby = OpenDatabase(globaldir & "\siscont", False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("mdh_to")
    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from origen where origen='" & Trim("" & mytabley.Fields("tov")) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            mytablex.Fields("origen") = Trim("" & mytabley.Fields("tov"))
            mytablex.Fields("descripcio") = Trim("" & mytabley.Fields("origen"))
            mytablex.Update

        End If

        mytablex.Close
        mytabley.MoveNext
    Loop
    mytabley.Close
    MsgBox "Procesado Origenes ", 48, "Aviso"
    Exit Sub

End Sub

Sub copia_tipo_documento()

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As Table

    cn.Execute ("delete from docta")
    Set mydby = OpenDatabase(globaldir & "\siscont", False, False, "foxpro 2.5;")
    Set mytabley = mydby.OpenTable("mdh_td")

    Do

        If mytabley.EOF Then Exit Do
        mytablex.Open "select * from docta where docta='" & Trim("" & mytabley.Fields("td")) & "'", cn, adOpenStatic, adLockOptimistic

        If mytablex.RecordCount = 0 Then
            mytablex.AddNew
            mytablex.Fields("docta") = Trim("" & mytabley.Fields("td"))
            mytablex.Fields("descripcio") = Trim("" & mytabley.Fields("doc"))
            mytablex.Update

        End If

        mytablex.Close
        mytabley.MoveNext
    Loop
    mytabley.Close
    MsgBox "Procesado Tipo Documento ", 48, "Aviso"
    Exit Sub

End Sub

Sub siscont_cuenta()

    Dim I As Integer

    Dim vr

    Dim dd       As String

    Dim sdx      As Double

    Dim mydby    As Database

    Dim mytablex As Table

    Dim mytabley As New ADODB.Recordset

    cn.Execute ("delete from cuentas")
    sdx = 0
    Set mydby = OpenDatabase(globaldir & "\siscont", False, False, "foxpro 2.5;")
    Set mytablex = mydby.OpenTable("mdh_plan")

    mytabley.Open "select *  from cuentas  ", cn, adOpenStatic, adLockOptimistic
    MsgBox "Empezar..enter" + "" & mytablex.RecordCount
    Do

        If mytablex.EOF Then Exit Do
        '------------------------
        mytabley.AddNew
        mytabley.Fields("cuenta") = Trim("" & mytablex.Fields("cuenta"))
        mytabley.Fields("descripcio") = Trim("" & mytablex.Fields("nombre"))
   
        If Trim("" & mytablex.Fields("bd")) = "1" Then
            mytabley.Fields("nivelcuenta") = "B"  'BALANCE

        End If

        If Trim("" & mytablex.Fields("bd")) = "2" Then
            mytabley.Fields("nivelcuenta") = "S"  'SUBCUENTA

        End If

        If Trim("" & mytablex.Fields("bd")) = "3" Then
            mytabley.Fields("nivelcuenta") = "R"  'REGISTRO

        End If

        mytabley.Fields("tipocuenta") = Trim("" & mytablex.Fields("rnf"))
        mytabley.Fields("tipoanalisis") = Trim("" & mytablex.Fields("cta"))
        mytabley.Fields("moneda") = Trim("" & mytablex.Fields("sd"))
        mytabley.Update
        sdx = sdx + 1
        dd = "" & sdx
        vr = DoEvents()
        '------------------------
        mytablex.MoveNext
    Loop
    MsgBox "acabe,pasando la cuenta"

End Sub

