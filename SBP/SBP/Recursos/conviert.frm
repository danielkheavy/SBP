VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form conviert 
   Caption         =   "Asistente Conversion Documentos"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   11445
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11445
   StartUpPosition =   1  'CenterOwner
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
      Height          =   6015
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   10335
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
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   240
         Width           =   1815
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
         TabIndex        =   48
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
         Left            =   8520
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5160
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   240
         Width           =   2055
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "conviert.frx":0000
         Height          =   5055
         Left            =   120
         OleObjectBlob   =   "conviert.frx":0014
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   840
         Width           =   9975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Convertir a :"
      Height          =   3255
      Left            =   5040
      TabIndex        =   32
      Top             =   1200
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox fpago 
         Height          =   375
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   39
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox xfecha 
         Height          =   375
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   38
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox xnumero 
         Height          =   375
         Left            =   1680
         MaxLength       =   11
         TabIndex        =   37
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox xserie 
         Height          =   375
         Left            =   240
         MaxLength       =   3
         TabIndex        =   36
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox xtipo 
         Height          =   375
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   35
         Top             =   960
         Width           =   1215
      End
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
         Left            =   4080
         MaskColor       =   &H00E0E0E0&
         Picture         =   "conviert.frx":09DF
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Grabar registro"
         Top             =   480
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
         Height          =   615
         Left            =   4080
         MaskColor       =   &H00E0E0E0&
         Picture         =   "conviert.frx":1BF1
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Borrar registro"
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label fpago1 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FormaPago"
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FechaProceso"
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo"
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numero"
         Height          =   375
         Left            =   1680
         TabIndex        =   41
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Serie"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   2160
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Procesar"
      Height          =   495
      Left            =   3360
      TabIndex        =   31
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdDelete 
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
      Left            =   2520
      MaskColor       =   &H00E0E0E0&
      Picture         =   "conviert.frx":2E03
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Borrar registro"
      Top             =   5400
      Width           =   735
   End
   Begin VB.ComboBox bodega 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox numero7 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   25
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox serie7 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   24
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox numero6 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   23
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox serie6 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   22
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox numero5 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   21
      Top             =   4200
      Width           =   1695
   End
   Begin VB.TextBox serie5 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   20
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox numero4 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   19
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox serie4 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   18
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox numero3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   17
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox serie3 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   16
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox numero2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   15
      Top             =   3120
      Width           =   1695
   End
   Begin VB.TextBox serie2 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   14
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox fechaf 
      Height          =   375
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   13
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox fechai 
      Height          =   375
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   11
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox numero1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   9
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox serie1 
      Height          =   375
      Left            =   120
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox tipo 
      Height          =   375
      Left            =   1560
      MaxLength       =   2
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox moneda 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox codigo 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bodega"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label acu 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label tipoclie 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaFinal"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FechaInicio"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Numero"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Serie"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Moneda"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigo"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Menu ldsosalo 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "conviert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub bodega_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
serie1.SetFocus
End Sub

Private Sub buffer_DblClick()
Command1_Click
End Sub

Private Sub buffer_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 27 Then Exit Sub
If KeyAscii = 27 Then
   ldsosalo_Click
   Exit Sub
End If
Command1_Click

End Sub

Private Sub cmdDelete_Click()
serie1 = ""
serie2 = ""
serie3 = ""
serie4 = ""
serie5 = ""
serie6 = ""
serie7 = ""
numero1 = ""
numero2 = ""
numero3 = ""
numero4 = ""
numero5 = ""
numero6 = ""
numero7 = ""
End Sub

Private Sub codigo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 27 And KeyAscii <> 13 Then Exit Sub
found = busca_codigo()
If found = 0 Then
   codigo = ""
   codigo.SetFocus
   Exit Sub
End If
moneda.SetFocus

End Sub

Private Sub Codigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   tipo.SetFocus
   Exit Sub
End If
If KeyCode = &H70 Then  'f1
   consulta_codigo
End If

End Sub


Private Sub Command1_Click()
Dim buf As String
Dim buf1 As String
Dim buf2 As String
Dim buf3 As String
Dim buf4 As String
buf2 = ""
If tipoclie = "P" Then
   buf2 = "PROVEEDO"
End If
If tipoclie = "C" Then
   buf2 = "CLIENTEs"
End If
If tipoclie = "V" Then
   buf2 = "VENDEDOR"
End If
If opcion1 = "3" Or opcion1 = "4" Or opcion1 = "6" Or opcion1 = "7" Or opcion1 = "8" Or opcion1 = "9" Or opcion1 = "10" Then
   buf3 = " tipo='" & tipo & "'"
   buf3 = buf3 & " and codigo='" & codigo & "'"
   buf3 = buf3 & " and moneda='" & moneda & "'"
End If

   buf1 = ""
   If opcion1 = "1" Then
      If Len(buffer) = 0 Then
      buf = "select Descripcio,Tipo from Tipo where grupo<>'" & acu & "'"
      Else
      buf = "select Descripcio,Tipo from tipo where grupo<>'" & acu & "' and " & Combo1 & " like '" & buffer & "%'"
      End If
   End If
If opcion1 = "2" Then
      If Len(buffer) = 0 Then
      buf = "select Nombre,Codigo from  " & buf2
      Else
      buf = "select Nombre,Codigo from " & buf2 & " where " & Combo1 & " like '" & buffer & "%'"
      End If
   End If
If opcion1 = "3" Or opcion1 = "4" Or opcion1 = "6" Or opcion1 = "7" Or opcion1 = "8" Or opcion1 = "9" Or opcion1 = "10" Then
      If Len(buffer) = 0 Then
      buf = "select Tipo,Serie,Numero,Fecha,Codigo,Total,Estado from  " & xarchivo & " where " & buf3
      Else
      buf = "select Tipo,Serie,Numero,Fecha,Codigo,Total,Estado from " & xarchivo & " where " & buf3 & " and " & Combo1 & " like '" & buffer & "%'"
      End If
   End If
If opcion1 = "5" Then
      If Len(buffer) = 0 Then
      buf = "select Descripcio,Fpago from Fpago "
      Else
      buf = "select Descripcio,Fpago from Fpago where " & Combo1 & " like '" & buffer & "%'"
      End If
   End If
   
If Combo2.ListIndex = 1 Then
   buf = buf & " order by " & Combo1
End If

               Data1.Connect = "foxpro 2.5;"
               Data1.DatabaseName = globaldir
               If opcion1 = "3" Or opcion1 = "4" Or opcion1 = "6" Or opcion1 = "7" Or opcion1 = "8" Or opcion1 = "9" Or opcion1 = "10" Then
               Data1.DatabaseName = globaldat
               MsgBox buf
               End If
               Data1.RecordSource = buf
               Data1.Refresh
               If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
                  Data1.Recordset.Close
                  buffer.SetFocus
                  Exit Sub
               End If
               If opcion1 = "3" Or opcion1 = "4" Or opcion1 = "6" Or opcion1 = "7" Or opcion1 = "8" Or opcion1 = "9" Or opcion1 = "10" Then
               DBGrid1.columns(0).Width = 700
               DBGrid1.columns(1).Width = 700
               DBGrid1.columns(2).Width = 1500
               DBGrid1.columns(3).Width = 2000
               DBGrid1.columns(4).Width = 1000
               DBGrid1.columns(5).Width = 1000
               DBGrid1.columns(6).Width = 700
               End If
               If opcion1 = "1" Or opcion1 = "2" Or opcion1 = "5" Then
               DBGrid1.columns(0).Width = 4000
               DBGrid1.columns(1).Width = 2000
               End If
               DBGrid1.SetFocus

End Sub

Private Sub Command2_Click()
habilita 0
Frame2.Visible = False
tipo.SetFocus

End Sub

Private Sub Command3_Click()
habilita 1
Frame2.Visible = True
xtipo.SetFocus
End Sub
Sub habilita(sw As Integer)
Dim xsw
If sw = 0 Then
   xsw = True
End If
If sw = 1 Then
   xsw = False
End If
tipo.Enabled = xsw
codigo.Enabled = xsw
moneda.Enabled = xsw
fechai.Enabled = xsw
fechaf.Enabled = xsw
bodega.Enabled = xsw
serie1.Enabled = xsw
serie2.Enabled = xsw
serie3.Enabled = xsw
serie4.Enabled = xsw
serie5.Enabled = xsw
serie6.Enabled = xsw
serie7.Enabled = xsw
Command3.Enabled = xsw

End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim found As Integer
Dim buf As String
Dim xtemp As Variant
If KeyCode = 27 Then
   buffer.SetFocus
   Exit Sub
End If
If KeyCode = 13 Then
If opcion1 = "1" Then
   tipo = DBGrid1.columns(1)
   Frame1.Visible = False
   tipo.SetFocus
   tipo_KeyPress 13
End If
If opcion1 = "2" Then
   codigo = DBGrid1.columns(1)
   Frame1.Visible = False
   codigo.SetFocus
   codigo_KeyPress 13
End If
If opcion1 = "3" Then
   serie1 = DBGrid1.columns(1)
   numero1 = DBGrid1.columns(2)
   Frame1.Visible = False
   serie2.SetFocus
End If
If opcion1 = "4" Then
   serie2 = DBGrid1.columns(1)
   numero2 = DBGrid1.columns(2)
   Frame1.Visible = False
   serie3.SetFocus
End If
If opcion1 = "6" Then
   serie3 = DBGrid1.columns(1)
   numero3 = DBGrid1.columns(2)
   Frame1.Visible = False
   serie4.SetFocus
End If
If opcion1 = "7" Then
   serie4 = DBGrid1.columns(1)
   numero4 = DBGrid1.columns(2)
   Frame1.Visible = False
   serie5.SetFocus
End If
If opcion1 = "8" Then
   serie5 = DBGrid1.columns(1)
   numero5 = DBGrid1.columns(2)
   Frame1.Visible = False
   serie6.SetFocus
End If
If opcion1 = "9" Then
   serie6 = DBGrid1.columns(1)
   numero6 = DBGrid1.columns(2)
   Frame1.Visible = False
   serie7.SetFocus
End If
If opcion1 = "10" Then
   serie7 = DBGrid1.columns(1)
   numero7 = DBGrid1.columns(2)
   Frame1.Visible = False
End If

If opcion1 = "5" Then
   fpago = DBGrid1.columns(1)
   Frame1.Visible = False
   fpago.SetFocus
   fpago_KeyPress 13
End If
End If

End Sub

Private Sub fechaf_KeyPress(KeyAscii As Integer)
If KeyAscii <> 27 And KeyAscii <> 13 Then Exit Sub
If Len(fechaf) = 0 Then
   fechaf = Format(Now, "dd/mm/yyyy")
End If
If Not IsDate(fechaf) Then Exit Sub
fechaf = Format(fechaf, "dd/mm/yyyy")
bodega.SetFocus

End Sub

Private Sub fechai_KeyPress(KeyAscii As Integer)
If KeyAscii <> 27 And KeyAscii <> 13 Then Exit Sub
If Len(fechai) = 0 Then
   fechai = Format(Now, "dd/mm/yyyy")
End If
If Not IsDate(fechai) Then Exit Sub
fechai = Format(fechai, "dd/mm/yyyy")
fechaf.SetFocus
End Sub

Private Sub fechai_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   moneda.SetFocus
   Exit Sub
End If

End Sub

Private Sub Form_Load()
fechai = Format(Now, "dd/mm/yyyy")
fechaf = Format(Now, "dd/mm/yyyy")
xfecha = Format(Now, "dd/mm/yyyy")
moneda.Clear
moneda.AddItem "S"
moneda.AddItem "D"
moneda.ListIndex = 0
carga_bodega
End Sub

Private Sub fpago_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 13 Then Exit Sub
found = busca_fpago()
If found = 0 Then
   fpago = ""
   Exit Sub
End If


End Sub

Private Sub fpago_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_fpago
End If

End Sub

Private Sub ldsosalo_Click()
If Frame1.Visible = True Then
   Frame1.Visible = False
   If opcion1 = "1" Then
      tipo.SetFocus
   End If
   Exit Sub
End If
conviert.Hide
Unload conviert
End Sub
Function busca_tipo(sw As Integer)

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("tipo")
mytablex.Index = "tipo"
mytablex.Seek "=", tipo
If Not mytablex.NoMatch Then
   busca_tipo = 1
   Select Case "" & mytablex.Fields("tipodoc")
          Case "A", "B", "C", "D", "G", "E", "F"  'VENTAS
               xarchivo = "FACTURA"
               xarchivo1 = "DETALLE"
          Case "J", "K", "L", "M", "P", "N", "O"  'COMPRAS
               xarchivo = "FACTURA"
               xarchivo1 = "DETALLE"
          Case "H"  'COTIZACION VENTAS
               xarchivo = "CCOTIZAV"
               xarchivo1 = "DCOTIZAV"
          Case "I"  'PEDIDO VENTAS
               xarchivo = "CPEDIDOV"
               xarchivo1 = "DPEDIDOV"
          Case "Q"  'REQUISICION COMPRAS
               xarchivo = "CREQUISA"
               xarchivo1 = "DREQUISA"
          Case "R"  'ORDEN COMPRA
               xarchivo = "CORDENC"
               xarchivo1 = "DORDENC"
          Case "T", "S" 'GUIA REMISION
               xarchivo = "FACTURA"
               xarchivo1 = "DETALLE"

               
   End Select
End If
'------------------------------------- ------------
mytablex.Close
 
End Function
Function busca_codigo()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("clientes")
mytablex.Index = "codigo"
mytablex.Seek "=", codigo
If Not mytablex.NoMatch Then
   busca_codigo = 1
End If
'------------------------------------- ------------
mytablex.Close
 

End Function

Private Sub moneda_KeyPress(KeyAscii As Integer)
fechai.SetFocus
End Sub

Private Sub moneda_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H26 Then
   codigo.SetFocus
   Exit Sub
End If

End Sub

Private Sub serie1_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 27 And KeyAscii <> 13 Then Exit Sub
'found = busca_codigo()
'If found = 0 Then
'   codigo = ""
'   codigo.SetFocus
'   Exit Sub
'End If
'moneda.SetFocus

End Sub

Private Sub serie1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_documentos
End If

End Sub

Private Sub serie2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_documentos1
End If

End Sub

Private Sub serie3_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_documentos2
End If

End Sub

Private Sub serie4_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_documentos3
End If

End Sub

Private Sub serie5_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_documentos4
End If

End Sub

Private Sub serie6_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_documentos5
End If

End Sub

Private Sub serie7_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_documentos6
End If

End Sub

Private Sub tipo_KeyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 27 And KeyAscii <> 13 Then Exit Sub
If Len(tipo) = 0 Then
   tipo.SetFocus
   Exit Sub
End If
found = busca_tipo(0)
If found = 0 Then
   tipo = ""
   tipo.SetFocus
   Exit Sub
End If
codigo.SetFocus
End Sub
Sub valida_xx()
Dim found As Integer
If Len(tipo) = 0 Then
   tipo.SetFocus
   Exit Sub
End If
found = busca_tipo(0)
If found = 0 Then
   tipo = ""
   tipo.SetFocus
   Exit Sub
End If
found = busca_codigo()
If found = 0 Then
   codigo = ""
   codigo.SetFocus
   Exit Sub
End If
If Not IsDate(fechai) Then
   fechai = ""
   fechai.SetFocus
   Exit Sub
End If
If Not IsDate(fechaf) Then
   fechaf = ""
   fechaf.SetFocus
   Exit Sub
End If





End Sub

Private Sub tipo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H70 Then  'f1
   consulta_tipo
End If

End Sub
Sub consulta_tipo()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Tipo"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "1"

End Sub
Sub carga_bodega()

Dim mytablex As Table
bodega.Clear

Set mytablex = mydbxglo.OpenTable("bodega")
mytablex.Index = "codigo"
Do
If mytablex.EOF Then Exit Do
bodega.AddItem "" & mytablex.Fields("codigo")
mytablex.MoveNext
Loop
'------------------------------------- ------------
mytablex.Close
 
bodega.ListIndex = 0

End Sub
Sub consulta_codigo()

Combo1.Clear
Combo1.AddItem "Nombre"
Combo1.AddItem "Codigo"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "2"
End Sub
Sub consulta_fpago()
Combo1.Clear
Combo1.AddItem "Descripcio"
Combo1.AddItem "Fpago"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "5"

End Sub
Function busca_fpago()

Dim mytablex As Table

Set mytablex = mydbxglo.OpenTable("fpago")
mytablex.Index = "fpago"
mytablex.Seek "=", fpago
If Not mytablex.NoMatch Then
   busca_fpago = 1
End If
mytablex.Close
 
End Function
Sub consulta_documentos()
Combo1.Clear
Combo1.AddItem "Tipo"
Combo1.AddItem "Serie"
Combo1.AddItem "Numero"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "3"

End Sub
Sub consulta_documentos1()
Combo1.Clear
Combo1.AddItem "Tipo"
Combo1.AddItem "Serie"
Combo1.AddItem "Numero"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "4"

End Sub
Sub consulta_documentos2()
Combo1.Clear
Combo1.AddItem "Tipo"
Combo1.AddItem "Serie"
Combo1.AddItem "Numero"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "6"

End Sub
Sub consulta_documentos3()
Combo1.Clear
Combo1.AddItem "Tipo"
Combo1.AddItem "Serie"
Combo1.AddItem "Numero"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "7"

End Sub
Sub consulta_documentos4()
Combo1.Clear
Combo1.AddItem "Tipo"
Combo1.AddItem "Serie"
Combo1.AddItem "Numero"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "8"

End Sub
Sub consulta_documentos5()
Combo1.Clear
Combo1.AddItem "Tipo"
Combo1.AddItem "Serie"
Combo1.AddItem "Numero"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "9"

End Sub
Sub consulta_documentos6()
Combo1.Clear
Combo1.AddItem "Tipo"
Combo1.AddItem "Serie"
Combo1.AddItem "Numero"
Combo1.ListIndex = 0
Frame1.Visible = True
buffer = ""
buffer.SetFocus
opcion1 = "10"

End Sub

Private Sub xfecha_KeyPress(KeyAscii As Integer)
If KeyAscii <> 27 And KeyAscii <> 13 Then Exit Sub
If Len(xfecha) = 0 Then
   xfecha = Format(Now, "dd/mm/yyyy")
End If
If Not IsDate(xfecha) Then Exit Sub
xfecha = Format(xfecha, "dd/mm/yyyy")
fpago.SetFocus

End Sub

Private Sub xtipo_keyPress(KeyAscii As Integer)
Dim found As Integer
If KeyAscii <> 27 And KeyAscii <> 13 Then Exit Sub
If Len(xtipo) = 0 Then
   xtipo.SetFocus
   Exit Sub
End If
found = busca_tipo(1)
If found = 0 Then
   xtipo = ""
   xtipo.SetFocus
   Exit Sub
End If
xfecha.SetFocus

End Sub
