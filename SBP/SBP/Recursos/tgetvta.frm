VERSION 5.00
Begin VB.Form tgetvta 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recoger Datos de La Web"
   ClientHeight    =   3255
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   9180
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
      Left            =   2040
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox fechai 
      Height          =   495
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      Height          =   615
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo Operacion"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicio"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado del proceso"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label registro 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Menu lo89232 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "tgetvta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub recoge_todo()
recoge_cabeza
recoge_detalle
recoge_fpagov
End Sub
Private Sub Command1_Click()
Dim found As Integer
Dim xbuf As String
Dim xbuf1 As String
Dim xbuf2 As String
Dim xbuf3 As String
Dim xbuf4 As String
Dim xbuf5 As String
Dim xbuf6 As String
Dim xbuf7 As String
Dim xbuf8 As String
Dim xbuf9 As String
Dim xbuf10 As String
Dim xbuf11 As String



On Error GoTo cmd34_err
If Len(fechai) <> 10 Then Exit Sub
If Not IsDate(fechai) Then Exit Sub
If copia_tmpweb() <> 1 Then
   MsgBox "Error al copiar temporal Web ", 48, "Aviso"
   Exit Sub
End If
If Combo1 = "Producto" Then
   xbuf = "01PRODUC"
   xbuf1 = "01FAMIL"
   xbuf2 = "01SUBFAM"
   xbuf3 = "01seccion"
   xbuf4 = "01marca"
   xbuf5 = "01catego"
   xbuf6 = "01provee"
   xbuf7 = "01cliente"
   xbuf8 = "01linea"
   xbuf9 = "01color"
   xbuf10 = "01equiva"
   xbuf11 = "01codprov"
   Call ExecuteCommand(globalweb & "\r\bajarp.bat scanpos " & xbuf & " " & xbuf1 & " " & xbuf2 & " " & xbuf3 & " " & xbuf4 & " " & xbuf5 & " " & xbuf6 & " " & xbuf7 & " " & xbuf8 & " " & xbuf9 & " " & xbuf10 & " " & xbuf11)
   actualizar_productos
   MsgBox "proceso Terminado", 48, "Aviso"
   Exit Sub
End If
If Combo1 = "Compras/Ventas" Then
   xbuf = "01C" + Format(fechai, "ddmmyyyy")
   xbuf1 = "01D" + Format(fechai, "ddmmyyyy")
   xbuf2 = "01F" + Format(fechai, "ddmmyyyy")
   found = borra_nombre(globalweb & "\r\" + "01C" + Format(fechai, "ddmmyyyy"))
   found = borra_nombre(globalweb & "\r\" + "01D" + Format(fechai, "ddmmyyyy"))
   found = borra_nombre(globalweb & "\r\" + "01F" + Format(fechai, "ddmmyyyy"))
   Call ExecuteCommand(globalweb & "\r\bajar.bat scanpos " & xbuf & " " & xbuf1 & " " & xbuf2)
   recoge_todo
End If
Exit Sub
cmd34_err:
MsgBox "Error en " + error$, 48, "Aviso"
Exit Sub
End Sub
Sub recoge_cabeza()
Dim strF() As String
Dim strl As String
Dim xbuf As String
Dim p As Long
Dim vr
Dim mydbx As Database
Dim mytablex As Table
On Error GoTo cmd89_err
Set mydbx = OpenDatabase(globalweb & "\R\", False, False, "foxpro 2.5;")
Set mytablex = mydbxglo.OpenTable("_c" & gusuario)
mytablex.Index = "tfactura"
xbuf = globalweb & "\r\" + "01C" + Format(fechai, "ddmmyyyy")
Open xbuf For Binary As #1
p = 0
While Not EOF(1)
Line Input #1, strl
strF = Split(strl, "|")
p = p + 1
mytablex.Seek "=", strF(0), strF(1), strF(2)
If mytablex.NoMatch Then
   mytablex.AddNew
      'MsgBox strF(mytablex.Fields.Count - 1)
      'MsgBox "" & mydbxglo.TableDefs("factura").Fields(mytablex.Fields.Count - 1).Type
      For i = 0 To mytablex.Fields.Count - 1
       Select Case "" & mydbxglo.TableDefs("factura").Fields(i).Type
       Case "7"
            mytablex.Fields(i) = Val("" & strF(i))
       Case "8"
            If Len(strF(i)) = 10 Then
               mytablex.Fields(i) = Format(strF(i), "dd/mm/yyyy")
            End If
       Case "10"
            mytablex.Fields(i) = "" & strF(i)
       End Select
      Next i
   mytablex.Update
End If
vr = DoEvents()
registro = "" & p
Wend
Close #1
mytablex.Close
 
MsgBox "Proceso Terminado ", 48, "Aviso"
Exit Sub
cmd89_err:
MsgBox "Proceso Terminado ", 48, "Aviso"
mytablex.Close
 
Exit Sub

End Sub
Sub recoge_detalle()
Dim strF() As String
Dim strl As String
Dim xbuf As String
Dim p As Long
Dim vr
Dim mydbx As Database
Dim mytablex As Table
On Error GoTo cmd90_err
Set mydbx = OpenDatabase(globalweb & "\R\", False, False, "foxpro 2.5;")
Set mytablex = mydbxglo.OpenTable("_d" & gusuario)
'mytablex.Index = "tfactura"
xbuf = globalweb & "\r\" + "01D" + Format(fechai, "ddmmyyyy")
Open xbuf For Binary As #1
p = 0
While Not EOF(1)
Line Input #1, strl
strF = Split(strl, "|")
p = p + 1
'mytablex.Seek "=", strF(0), strF(1), strF(2)
'If mytablex.NoMatch Then
   mytablex.AddNew
      'MsgBox strF(mytablex.Fields.Count - 1)
      'MsgBox "" & mydbxglo.TableDefs("factura").Fields(mytablex.Fields.Count - 1).Type
      For i = 0 To mytablex.Fields.Count - 1
       Select Case "" & mydbxglo.TableDefs("detalle").Fields(i).Type
       Case "7"
            mytablex.Fields(i) = Val("" & strF(i))
       Case "8"
            If Len(strF(i)) = 10 Then
               mytablex.Fields(i) = Format(strF(i), "dd/mm/yyyy")
            End If
       Case "10"
            mytablex.Fields(i) = "" & strF(i)
       End Select
      Next i
   mytablex.Update
'End If
vr = DoEvents()
registro = "" & p
Wend
Close #1
mytablex.Close
 
'MsgBox "Proceso Terminado ", 48, "Aviso"
Exit Sub
cmd90_err:
'MsgBox "Proceso Terminado ", 48, "Aviso"
mytablex.Close
 
Exit Sub
End Sub
Sub recoge_fpagov()
Dim strF() As String
Dim strl As String
Dim xbuf As String
Dim p As Long
Dim vr
Dim mydbx As Database
Dim mytablex As Table
On Error GoTo cmd91_err
Set mydbx = OpenDatabase(globalweb & "\R\", False, False, "foxpro 2.5;")
Set mytablex = mydbxglo.OpenTable("_f" & gusuario)
'mytablex.Index = "tfactura"
xbuf = globalweb & "\r\" + "01F" + Format(fechai, "ddmmyyyy")
Open xbuf For Binary As #1
p = 0
While Not EOF(1)
Line Input #1, strl
strF = Split(strl, "|")
p = p + 1
'mytablex.Seek "=", strF(0), strF(1), strF(2)
'If mytablex.NoMatch Then
   mytablex.AddNew
      'MsgBox strF(mytablex.Fields.Count - 1)
      'MsgBox "" & mydbxglo.TableDefs("factura").Fields(mytablex.Fields.Count - 1).Type
      For i = 0 To mytablex.Fields.Count - 1
       Select Case "" & mydbxglo.TableDefs("fpagov").Fields(i).Type
       Case "7"
            mytablex.Fields(i) = Val("" & strF(i))
       Case "8"
            If Len(strF(i)) = 10 Then
               mytablex.Fields(i) = Format(strF(i), "dd/mm/yyyy")
            End If
       Case "10"
            mytablex.Fields(i) = "" & strF(i)
       End Select
      Next i
   mytablex.Update
'End If
vr = DoEvents()
registro = "" & p
Wend
Close #1
mytablex.Close
 
'MsgBox "Proceso Terminado ", 48, "Aviso"
Exit Sub
cmd91_err:
'MsgBox "Proceso Terminado ", 48, "Aviso"
mytablex.Close
 
Exit Sub

End Sub

Private Sub Form_Load()
fechai = Format(Now, "dd/mm/yyyy")
Combo1.Clear
Combo1.AddItem "*"
'Combo1.AddItem "Producto"
Combo1.AddItem "Compras/Ventas"
Combo1.ListIndex = 0
End Sub

Private Sub lo89232_Click()
tgetvta.Hide
Unload tgetvta
End Sub
Sub actualizar_productos()
Dim strF() As String
Dim strl As String
Dim xbuf As String
Dim p As Long
Dim vr
Dim mydbx As Database
Dim mytablex As Table
On Error GoTo cmd891_err
Set mydbx = OpenDatabase(globaldir, False, False, "foxpro 2.5;")
Set mytablex = mydbxglo.OpenTable("producto")
mytablex.Index = "producto"
xbuf = globalweb & "\r\01PRODUC"
Open xbuf For Binary As #1
p = 0
While Not EOF(1)
Line Input #1, strl
strF = Split(strl, "|")
p = p + 1
mytablex.Seek "=", strF(0)
If mytablex.NoMatch Then
   mytablex.AddNew
      'MsgBox strF(mytablex.Fields.Count - 1)
      'MsgBox "" & mydbxglo.TableDefs("factura").Fields(mytablex.Fields.Count - 1).Type
      For i = 0 To mytablex.Fields.Count - 1
       Select Case "" & mydbxglo.TableDefs("producto").Fields(i).Type
       Case "7"
            mytablex.Fields(i) = Val("" & strF(i))
       Case "8"
            If Len(strF(i)) = 10 Then
               mytablex.Fields(i) = Format(strF(i), "dd/mm/yyyy")
            End If
       Case "10"
            mytablex.Fields(i) = "" & strF(i)
       End Select
      Next i
   mytablex.Update
End If
vr = DoEvents()
registro = "" & p
Wend
Close #1
mytablex.Close
 
MsgBox "Proceso Terminado ", 48, "Aviso"
Exit Sub
cmd891_err:
MsgBox "Proceso Terminado ", 48, "Aviso"
mytablex.Close
 
Exit Sub

End Sub
