VERSION 5.00
Begin VB.Form tcolista 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copiar Listas"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.ComboBox lista2 
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.ComboBox lista1 
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
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label nregistro 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lista Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lista Inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Menu fo99333 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tcolista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    Dim mytablex As New ADODB.Recordset

    Dim mytabley As New ADODB.Recordset

    Dim I        As Integer

    Dim sdx      As Double

    Dim vr

    If lista1 = "%" Or lista2 = "%" Then
        MsgBox "Seleccione Listas ", 48, "Aviso"
        Exit Sub

    End If

    If lista1 = lista2 Then
        MsgBox "Listas diferentes ", 48, "Aviso"
        Exit Sub

    End If

    mytabley.Open "select * from precios where local='" & lista1 & "'", cn, adOpenStatic, adLockOptimistic

    If mytabley.RecordCount = 0 Then
        MsgBox "No existen Precios ", 48, "Aviso"
        mytabley.Close
        Exit Sub

    End If

    If MsgBox("Desea Copiar Lista", 1, "Aviso") <> 1 Then Exit Sub
    sdx = 0

    cn.Execute ("delete from precios where local='" & lista2 & "'")
    mytablex.Open "select * from precios where local='" & lista2 & "'", cn, adOpenStatic, adLockOptimistic
    Do

        If mytabley.EOF Then Exit Do
        sdx = sdx + 1
        nregistro = Format(sdx, "000000")
        vr = DoEvents()
        mytablex.AddNew

        For I = 0 To mytabley.Fields.count - 1
            mytablex.Fields(I) = mytabley.Fields(I)
        Next I

        mytablex.Fields("local") = Trim(lista2)
        mytablex.Update

        mytabley.MoveNext
    Loop

End Sub

Private Sub fo99333_Click()
    tcolista.Hide
    Unload tcolista

End Sub

Private Sub Form_Load()
    lista1.Clear
    lista1.AddItem "%"

    For I = 0 To 11
        lista1.AddItem Format(I, "00")
    Next I

    lista1.ListIndex = 0

    lista2.Clear
    lista2.AddItem "%"

    For I = 0 To 11
        lista2.AddItem Format(I, "00")
    Next I

    lista2.ListIndex = 0

End Sub
