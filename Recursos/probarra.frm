VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form probarra 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Browse productos"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   14820
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Consultas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   14415
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
         Height          =   495
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "&Ejecutar"
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
         Left            =   8160
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "probarra.frx":0000
         Height          =   7455
         Left            =   120
         OleObjectBlob   =   "probarra.frx":0014
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   840
         Width           =   12255
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "probarra.frx":09DF
      Height          =   7815
      Left            =   120
      OleObjectBlob   =   "probarra.frx":09F3
      TabIndex        =   10
      Top             =   960
      Width           =   14415
   End
   Begin VB.CommandButton xlocaloe 
      Caption         =   "&Localiza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12960
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox codigo 
      Height          =   375
      Left            =   10440
      MaxLength       =   10
      TabIndex        =   8
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
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
      Left            =   10440
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Buscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox loquebusca 
      Height          =   375
      Left            =   4800
      MaxLength       =   10
      TabIndex        =   4
      Text            =   "*"
      Top             =   480
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.ComboBox Combo4 
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
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Condicion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Condicion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   11
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buscar en la Seleccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ordenado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Seleccionar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu dl8923 
      Caption         =   "&Menu"
      Begin VB.Menu dj7823 
         Caption         =   "&1.Actualizar Productos desde la Central"
      End
   End
   Begin VB.Menu ldo3434 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "probarra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub buffer_DblClick()
Command1_Click

End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   ldo3434_Click
   Exit Sub
End If
Command1_Click

End Sub

Private Sub Buscar_Click()
Dim buf As String
On Error GoTo cmd6711_err
               
               Data2.Connect = "foxpro 2.5;"
               Data2.DatabaseName = globaldir
               buf = "select * from producto "
               buf = buf & " where " & Combo4 & " LIKE '" & loquebusca & "'"
               If Combo2 <> "TODOS" Then
                  buf = buf & " order by " & Combo2
               End If
               Data2.RecordSource = buf
               Data2.refresh
               dbgrid2.SetFocus
               Exit Sub
cmd6711_err:
MsgBox "Formato con Error ", 24, "Aviso"
Exit Sub

End Sub

Private Sub Command1_Click()
Dim buf As String
If opcion1 = "1" Then
   If Len(buffer) = 0 Then
   buf = "select Descripcio,Familia from Familia "
   Else
   buf = "select Descripcio,Familia from familia where " & Combo1 & " like '" & buffer & "%'"
   End If
End If
If opcion1 = "2" Then
   If Len(buffer) = 0 Then
   buf = "select Descripcio,Subfamilia from subFamil where familia='" & dbgrid2.columns(3) & "'"
   Else
   buf = "Descripcio,Subfamilia from subFamil where familia='" & dbgrid2.columns(3) & "' and " & Combo1 & " like '" & buffer & "%'"
   End If
End If
If opcion1 = "3" Then
   If Len(buffer) = 0 Then
   buf = "select Descripcio,Marca from Marca "
   Else
   buf = "select Descripcio,Marca from Marca where " & Combo1 & " like '" & buffer & "%'"
   End If
End If
If opcion1 = "4" Then
   If Len(buffer) = 0 Then
   buf = "select Descripcio,Seccion from Seccion "
   Else
   buf = "select Descripcio,Seccion from Seccion where " & Combo1 & " like '" & buffer & "%'"
   End If
End If
If opcion1 = "5" Then
   If Len(buffer) = 0 Then
   buf = "select Descripcio,Categoria from categori "
   Else
   buf = "select Descripcio,categoria from categori where " & Combo1 & " like '" & buffer & "%'"
   End If
End If


               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldir
               Data1.RecordSource = buf
               Data1.refresh
               If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
                  Data1.Recordset.Close
                  buffer.SetFocus
                  Exit Sub
               End If
               If opcion1 = "1" Or opcion1 = "2" Or opcion1 = "3" Or opcion1 = "4" Then
                  dbgrid1.columns(0).Width = 4000
                  dbgrid1.columns(1).Width = 2000
               End If
               dbgrid1.SetFocus

End Sub



Private Sub dbgrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
Select Case ColIndex
Case 0, 1, 3, 4, 5, 6, 7, 8, 9, 11, 12, 14, 15, 17, 18, 20, 21, 23, 24
     Cancel = True
     Exit Sub
Case 10
     If Val("" & dbgrid1.columns(9)) = 0 Then
        Cancel = True
        Exit Sub
     End If
Case 13
     If Val("" & dbgrid1.columns(13)) = 0 Then
        Cancel = True
        Exit Sub
     End If
Case 16
     If Val("" & dbgrid1.columns(16)) = 0 Then
        Cancel = True
        Exit Sub
     End If
Case 19
     If Val("" & dbgrid1.columns(19)) = 0 Then
        Cancel = True
        Exit Sub
     End If
Case 22
     If Val("" & dbgrid1.columns(22)) = 0 Then
        Cancel = True
        Exit Sub
     End If
Case 25
     If Val("" & dbgrid1.columns(25)) = 0 Then
        Cancel = True
        Exit Sub
     End If
     

End Select
     

End Sub
Sub consulta_categoria()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Categoria"
Combo1.ListIndex = 0
opcion1 = "5"
Frame1.Visible = True
buffer = ""
buffer.SetFocus
Command1_Click

End Sub

Sub consulta_seccion()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Seccion"
Combo1.ListIndex = 0
opcion1 = "4"
Frame1.Visible = True
buffer = ""
buffer.SetFocus
Command1_Click

End Sub

Sub consulta_marca()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Marca"
Combo1.ListIndex = 0
opcion1 = "3"
Frame1.Visible = True
buffer = ""
buffer.SetFocus
Command1_Click

End Sub

Sub consulta_Familia()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Familia"
Combo1.ListIndex = 0
opcion1 = "1"
Frame1.Visible = True
buffer = ""
buffer.SetFocus
Command1_Click

End Sub
Sub consulta_subFamilia()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "SubFamilia"
Combo1.ListIndex = 0
opcion1 = "2"
Frame1.Visible = True
buffer = ""
buffer.SetFocus
Command1_Click

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
If opcion1 = "1" Then 'familia
   dbgrid2.columns(3) = dbgrid1.columns(1)
   Frame1.Visible = False
   dbgrid2.SetFocus
End If
If opcion1 = "2" Then 'subfamilia
   dbgrid2.columns(4) = dbgrid1.columns(1)
   Frame1.Visible = False
   dbgrid2.SetFocus
End If
If opcion1 = "3" Then 'marca
   dbgrid2.columns(5) = dbgrid1.columns(1)
   Frame1.Visible = False
   dbgrid2.SetFocus
End If
If opcion1 = "4" Then 'seccion
   dbgrid2.columns(6) = dbgrid1.columns(1)
   Frame1.Visible = False
   dbgrid2.SetFocus
End If
If opcion1 = "5" Then 'categoria
   dbgrid2.columns(7) = dbgrid1.columns(1)
   Frame1.Visible = False
   dbgrid2.SetFocus
End If

End If

End Sub

Private Sub dbgrid2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 And dbgrid2.Col = 3 Then 'f1
   consulta_Familia
End If
If KeyCode = &H70 And dbgrid2.Col = 4 Then 'f1
   consulta_subFamilia
End If
If KeyCode = &H70 And dbgrid2.Col = 5 Then 'f1
   consulta_marca
End If
If KeyCode = &H70 And dbgrid2.Col = 6 Then 'f1
   consulta_seccion
End If
If KeyCode = &H70 And dbgrid2.Col = 7 Then 'f1
   consulta_categoria
End If

End Sub

Private Sub Form_Load()

    Combo2.Clear
    Combo2.AddItem "TODOS"
    Combo2.AddItem "val(PRODUCTO)"
    Combo2.AddItem "BARRAS"
    Combo2.AddItem "DESCRIPCIO"
    Combo2.AddItem "FAMILIA"
    Combo2.AddItem "SUBFAMILIA"
    'Combo1.AddItem "PROVEEDOR"
    Combo2.AddItem "MARCA"
    Combo2.AddItem "SECCION"
    Combo2.AddItem "UNIDAD1"
    Combo2.ListIndex = 0


    Combo3.Clear
    Combo3.AddItem "PRODUCTO"
    Combo3.AddItem "DESCRIPCIO"
    Combo3.AddItem "BARRAS"
    Combo3.AddItem "FAMILIA"
    Combo3.AddItem "SUBFAMILIA"
    'Combo3.AddItem "PROVEEDOR"
    Combo3.AddItem "MARCA"
    Combo3.AddItem "SECCION"
    Combo3.ListIndex = 0

    Combo4.Clear
    Combo4.AddItem "PRODUCTO"
    Combo4.AddItem "DESCRIPCIO"
    Combo4.AddItem "BARRAS"
    Combo4.AddItem "FAMILIA"
    Combo4.AddItem "SUBFAMILIA"
    'Combo4.AddItem "PROVEEDOR"
    Combo4.AddItem "MARCA"
    Combo4.AddItem "SECCION"
    Combo4.ListIndex = 1

End Sub

Private Sub ldo3434_Click()

If Frame1.Visible = True Then
   If opcion1 = "1" Or opcion1 = "2" Or opcion1 = "3" Or opcion1 = "4" Or opcion1 = "5" Then
      Frame1.Visible = False
      dbgrid2.SetFocus
      Exit Sub
   End If
End If
probarra.Hide
Unload probarra
End Sub

Private Sub xlocaloe_Click()
Dim MyCriteria As String
If Len(codigo) = 0 Then Exit Sub
MyCriteria = Combo3 & " LIKE '" & codigo & "%'"
Data2.Recordset.FindFirst MyCriteria
If Not Data2.Recordset.NoMatch Then
   Data2.Recordset.FindFirst MyCriteria
End If
Exit Sub

End Sub
