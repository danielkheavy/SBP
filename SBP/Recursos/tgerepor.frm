VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form tgerepor 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generando Reportes"
   ClientHeight    =   8940
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   12945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   13080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox busqueda 
      Height          =   375
      Left            =   4920
      TabIndex        =   29
      Top             =   8400
      Width           =   2535
   End
   Begin VB.TextBox lpag 
      Height          =   375
      Left            =   8760
      TabIndex        =   24
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox letrata 
      Height          =   375
      Left            =   6240
      TabIndex        =   22
      Top             =   7680
      Width           =   1215
   End
   Begin VB.ComboBox Combo6 
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
      Left            =   8880
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   7080
      Width           =   3975
   End
   Begin VB.ComboBox Combo5 
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
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   7080
      Width           =   3975
   End
   Begin VB.ComboBox ordenado 
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
      Left            =   8880
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4680
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Height          =   1095
      Left            =   8880
      TabIndex        =   15
      Top             =   5040
      Width           =   3975
   End
   Begin VB.TextBox campo 
      Height          =   495
      Left            =   4920
      TabIndex        =   14
      Top             =   5400
      Width           =   3975
   End
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
      Left            =   4920
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   4680
      Width           =   3975
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
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   6720
      Width           =   3975
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
      Left            =   4920
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   5040
      Width           =   3975
   End
   Begin VB.TextBox titulo 
      Height          =   495
      Left            =   4920
      MaxLength       =   50
      TabIndex        =   8
      Top             =   3720
      Width           =   7935
   End
   Begin MSDBGrid.DBGrid table1 
      Bindings        =   "tgerepor.frx":0000
      Height          =   3135
      Left            =   4920
      OleObjectBlob   =   "tgerepor.frx":0014
      TabIndex        =   6
      Top             =   120
      Width           =   7935
   End
   Begin VB.ListBox List2 
      Height          =   2985
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Width           =   4695
   End
   Begin VB.ComboBox namejoin 
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
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   5040
      Width           =   4695
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4695
   End
   Begin VB.ComboBox nombrebased 
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
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label nro 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   8760
      TabIndex        =   28
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nro"
      Height          =   375
      Left            =   7440
      TabIndex        =   27
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label registro 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6240
      TabIndex        =   26
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Registro"
      Height          =   375
      Left            =   4920
      TabIndex        =   25
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lineas x Pagina"
      Height          =   375
      Left            =   7440
      TabIndex        =   23
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tamaño Letra"
      Height          =   375
      Left            =   4920
      TabIndex        =   21
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mas Opciones"
      Height          =   375
      Left            =   8880
      TabIndex        =   20
      Top             =   6360
      Width           =   3975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Union"
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   6360
      Width           =   3975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ordenado Por"
      Height          =   375
      Left            =   8880
      TabIndex        =   10
      Top             =   4320
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Condicion de Busqueda"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   4320
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Titulo del Informe"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Campos"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre de la Tabla"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Menu dk811 
      Caption         =   "&Ejecuta"
   End
End
Attribute VB_Name = "tgerepor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim jswa As Integer

Private Sub dk811_Click()
Dim campox As String
campo = Trim(campo)
Text1 = Trim(Text1)
If Not IsNumeric(lpag) Then
   lpag = "45"
End If
If Val(lpag) < 1 Then
   lpag = "45"
End If
If Val(nro) < 1 Then
   MsgBox "No existe campos seleccionados", 24, "Aviso"
   Exit Sub
End If
letrata = "80"
If Val(nro) <= 80 Then
   letrata = "10"
   'lpag = "45"
End If
If Val(nro) > 80 And Val(nro) < 120 Then
   letrata = "8"
   'lpag = "45"
End If
If Val(nro) > 120 Then
   letrata = "7"
   'lpag = "45"
End If
If Len(Combo2) > 0 And (Len(Combo1) = 0 Or Len(campo) = 0) Then
   MsgBox "Falta LLenar un campo", 24, "Aviso"
   Exit Sub
   Exit Sub
End If
If Len(Combo1) > 0 And (Len(Combo2) = 0 Or Len(campo) = 0) Then
   MsgBox "Falta LLenar un campo", 24, "Aviso"
   Exit Sub
   Exit Sub
End If
If Len(campo) > 0 And (Len(Combo2) = 0 Or Len(Combo1) = 0) Then
   MsgBox "Falta LLenar un campo", 24, "Aviso"
   Exit Sub
   Exit Sub
End If
'combo3.ListIndex = "" & combo1.ListIndex
campox = campo
If Combo3.Text = "10" Then
   campox = "'" & campo & "'"
End If
busqueda = Combo1 & " " & Combo2 & " " & campox
If Len(Combo1) = 0 And Len(Combo2) = 0 And Len(campox) = 0 Then
   busqueda = ""
End If
DJ74.Enabled = False
proceso_impresion
DJ74.Enabled = True

End Sub

Private Sub Form_Activate()
Dim found As Integer
    table1.ColumnSumEnable(1) = True
    If Len(generador) = 0 Then
       generador = globaldir
    End If
    If Len(nombrebased) = 0 Then Exit Sub
    found = CARGA_BASE()
    refresca
    jswa = 1
End Sub

Private Sub Form_Load()
Dim found As Integer
Dim i As Integer
    Combo6.Clear
    Combo6.AddItem "CON CABECERAS"
    Combo6.AddItem "SIN CABECERAS"
End Sub
Function CARGA_BASE()
On Error GoTo cmd12_err
Dim i As Integer
   List1.Clear
   List2.Clear
   'nombrebased.Clear
   Set mydb3 = OpenDatabase(generador, False, False, "foxpro 2.5;")
   For i = 0 To mydb3.TableDefs.Count - 1
       nombrebased.AddItem mydb3.TableDefs(i)
   Next i
   For i = 0 To mydb3.TableDefs(nombrebased).Fields.Count - 1
       List1.AddItem mydb3.TableDefs(nombrebased).Fields(i).name
   Next i
    Combo1.Clear
    Combo3.Clear
    Combo1.AddItem ""
    Combo3.AddItem ""
   For i = 0 To mydb3.TableDefs(nombrebased).Fields.Count - 1
       Combo1.AddItem nombrebased & "." & mydb3.TableDefs(nombrebased).Fields(i).name
       Combo3.AddItem nombrebased & "." & mydb3.TableDefs(nombrebased).Fields(i).Type
   Next i
   Combo2.Clear
   Combo2.AddItem ""
   Combo2.AddItem "="
   Combo2.AddItem ">"
   Combo2.AddItem "<"
   Combo2.AddItem ">="
   Combo2.AddItem "<="
   Combo2.AddItem "LIKE"
   Combo2.AddItem "<>"

   ordenado.Clear
   ordenado.AddItem ""
   For i = 0 To mydb3.TableDefs(nombrebased).Fields.Count - 1
       ordenado.AddItem nombrebased & "." & mydb3.TableDefs(nombrebased).Fields(i).name
   Next i
   Combo4.Clear
   Combo5.Clear
   Combo4.AddItem ""
   Combo5.AddItem ""
   For i = 0 To mydb3.TableDefs(nombrebased).Fields.Count - 1
       Combo4.AddItem mydb3.TableDefs(nombrebased).Fields(i).name
   Next i
   If Len(namejoin) = 0 Then Exit Function
   For i = 0 To mydb3.TableDefs(namejoin).Fields.Count - 1
       List2.AddItem mydb3.TableDefs(namejoin).Fields(i).name
   Next i
   For i = 0 To mydb3.TableDefs(namejoin).Fields.Count - 1
       Combo5.AddItem mydb3.TableDefs(namejoin).Fields(i).name
       Combo1.AddItem namejoin & "." & mydb3.TableDefs(namejoin).Fields(i).name
   Next i
mydb3.Close
   Exit Function
cmd12_err:
  mydb3.Close
  Exit Function

End Function
Sub borrar_data1()
On Error GoTo cmd34_err
    Data1.Database.Execute "delete from reporte where  usuario='" & usuariopos & "'"
    Data1.Refresh
    Exit Sub
cmd34_err:
    Exit Sub
End Sub
Sub cabecera()
Dim found As Integer
On Error GoTo cmd345_err
Dim buf As String
Dim i As Integer
Dim j As Integer
Dim letternum
Dim mydbx As Database
Dim mysnapx As Snapshot
    If contlin > 0 Then
       buf = Chr$(12)
       found = formateaa(buf, Len(buf), 2, 0)
    End If
    j = Val(nro)
    If j < 80 Then
       j = 80
    End If
    contpag = contpag + 1
    contlin = 0
    found = formateaa("" & " Fecha Emision " & Format(Now, "dd/mm/yyyy") & " Pagina " & Str(contpag), 80, 2, 0)
    buf = titulo
    i = (j - Len("" & buf)) / 2
    found = formateaa(" ", i, 0, 0)
    found = formateaa("" & buf, Len("" & buf), 2, 0)

    buf = String(j + 10, "=")
    found = formateaa(buf, j + 10, 2, 0)
    '-----------------
    Set mydbx = OpenDatabase(generador, False, False, "foxpro 2.5;")
    Set mysnapx = mydbx.CreateSnapshot("select * from reporte where  usuario='" & usuariopos & "'")
    Do
      If mysnapx.EOF Then Exit Do
      'If "" & mysnapx.Fields("local") <> local_1 Then
      '   Exit Do
      'End If
      i = Val("" & mysnapx.Fields("longitud"))
      buf = "" & mysnapx.Fields("nombre")
      letternum = InStr(buf, ".")
      buf = Mid$(buf, letternum + 1, Len(buf) - letternum)
      '------------------
      Select Case Val("" & mysnapx.Fields("tipo"))
       Case 3, 4, 7 'integer
            found = formateaa(buf, i, 0, 1)
       Case Else
            found = formateaa(buf, i, 0, 0)
      End Select
      found = formateaa("", 1, 0, 0)
      mysnapx.MoveNext
    Loop
    mysnapx.Close
    mydbx.Close
    found = formateaa("", 1, 2, 0)
    '-----------------
    buf = String(j + 10, "=")
    found = formateaa(buf, j + 10, 2, 0)
    Exit Sub
cmd345_err:
    MsgBox "Error en Cabecera" & error$, 24, "Aviso"
    Exit Sub

End Sub
Sub cuerpo_programa()
Dim i As Integer
Dim buf As String
Dim found As Integer
Dim atx As Double
Dim vr
Dim mysnap1 As Snapshot
'--------------------------
    atx = 0
    registro.Visible = True
    Command1.Visible = True
    Do
       If mysnap.EOF Then Exit Do
          '------------------------------
             If Command1.Visible = False Then
                GoTo salir
             End If
             Set mydb1 = OpenDatabase(generador, False, False, "foxpro 2.5;")
             Set mysnap1 = mydb1.CreateSnapshot("select * from reporte where  usuario='" & usuariopos & "'")
             Do
             If mysnap1.EOF Then
                Exit Do
             End If
             'If "" & mysnap1.Fields("local") <> local_1 Then
             '   Exit Do
             'End If
             '---
             buf = "" & mysnap.Fields("" & mysnap1.Fields("nombre"))
             i = Val("" & mysnap1.Fields("longitud"))
             If i <= 0 Then
                i = 1
             End If
             Select Case Val("" & mysnap1.Fields("tipo"))
             Case 7
                  Select Case Val("" & mysnap1.Fields("estados"))
                         Case 0
                         buf = Format(Val(buf), "0")
                         Case 1
                         buf = Format(Val(buf), "0.0")
                         Case 2
                         buf = Format(Val(buf), "0.00")
                         Case 3
                         buf = Format(Val(buf), "0.000")
                  End Select
                  found = formateaa(buf, i, 0, 1)
             Case 3, 4 'integer
                  found = formateaa(buf, i, 0, 1)
             Case Else
                  found = formateaa(buf, i, 0, 0)
             End Select
             found = formateaa("", 1, 0, 0)
             '---
             mysnap1.MoveNext
             Loop
             mysnap1.Close
             mydb1.Close
             found = formateaa("", i, 2, 0)
             nlineas
          '------------------------------
          vr = DoEvents()
          atx = atx + 1
          registro = Format(atx, "0")
          mysnap.MoveNext
    Loop
salir:
    registro.Visible = False
    Command1.Visible = False
    Exit Sub
cmd4566_err:
    registro.Visible = False
    Command1.Visible = False
    MsgBox " ..Error en Cuerpo Programa " & error$, 24, "Aviso"
    Exit Sub
End Sub
Sub nlineas()
contlin = contlin + 1
If contlin > Val(lpag) Then
   If Not Combo6 = "SIN CABECERAS" Then
      cabecera
   End If
End If
End Sub
Sub proceso_impresion()
Dim found As Integer
Dim buf As String
Dim mytablex As Table
Dim mydbx As Database
On Error GoTo cmd7806_err
    Screen.MousePointer = 11
    Set mydbx = OpenDatabase(generador, False, False, "foxpro 2.5;")
    found = sqllo(mydbx)
    If found = 0 Then
       Screen.MousePointer = 1
       mydbx.Close
       cerrar_archivo
       Exit Sub
    End If
    contlin = 0
    contpag = 0
    tnlineas = 49
    Filename = usuariopos
    cerrar_archivo
    found = borra_archivox(Filename)
    ncanal = 1
    Open Filename For Append As #ncanal
    If Not Combo6 = "SIN CABECERAS" Then
    cabecera
    End If
    cuerpo_programa
    contlin = 0
    contpag = 0
    Close #ncanal
    mysnap.Close
    mydbx.Close
    cerrar_archivo
    Screen.MousePointer = 1
    found = ejecuta_shell(Val(letrata))
    Exit Sub
cmd7806_err:
    MsgBox "Error " & error$, 24, "AVISO"
    mysnap.Close
    mydbx.Close
    Exit Sub
End Sub

Sub refresca()
    Data1.Connect = "FOXPRO 2.5;"
    Data1.DatabaseName = generador
    Data1.RecordSource = "select * from reporte where  usuario='" & usuariopos & "'"
    Data1.Refresh
End Sub
Function sqllo(mydby As Database)
Dim mydbx As Database
Dim mytablex As Snapshot
Dim buf As String
Dim buf1 As String
Dim BUFX As String
Dim sw As Integer
Dim sw1 As Integer
    On Error GoTo cmd21_err
    Screen.MousePointer = 11
    sw1 = 0
    '----------------
    Set mydbx = OpenDatabase(generador, False, False, "foxpro 2.5;")
    Set mytablex = mydbx.CreateSnapshot("select * from reporte where  usuario='" & usuariopos & "'")
        BUFX = ""
        sw = 0
        Do
        If mytablex.EOF Then
           Exit Do
           Else
           If sw = 1 Then
              BUFX = BUFX & ","
           End If
        End If
        If Len("" & mytablex.Fields("nombre")) > 0 Then
           'bufx = bufx & " " & mytablex.Fields("estados") & " " & mytablex.Fields("nombre") & "  "
           BUFX = BUFX & " " & mytablex.Fields("nombre") & "  "
           If Len(namejoin) > 0 Then
           If Mid$("" & mytablex.Fields("nombre"), 1, Len(namejoin)) = namejoin Then
              sw1 = 1
           End If
           End If
           sw = 1
        End If
        mytablex.MoveNext
        Loop
        mytablex.Close
        mydbx.Close
        If sw = 0 Then
           MsgBox "Campos No seleccionados", 24, "Aviso"
           Exit Function
        End If
    '----------------
    'buf = "select distinct direccion all from " & nombrebased
    'MsgBox BUFX
    buf = "select " & BUFX & " from " & nombrebased
    If sw1 = 1 Then   'si la busqueda tiene join
       '-------------------------------------------------
       buf = buf & " inner join " & namejoin & " on " & nombrebased & "." & Combo4 & "=" & namejoin & "." & Combo5 & " "
       If Len(busqueda) > 0 Then
          buf = buf & " and " & busqueda
          If Len(Text1) > 0 Then
             buf = buf & " and " & Text1 & " "
          End If
       End If
       If Len(busqueda) = 0 And Len(Text1) > 0 Then
          buf = buf & " and " & Text1
       End If
       'buf = buf & " and " & nombrebased & ".local='" & local_1 & "'"
       If Len(ordenado) > 0 Then
          buf = buf & " order by " & ordenado
       End If
       '-------------------------------------------------
    End If
    If sw1 = 0 Then
    If Len(busqueda) > 0 Then
       buf = buf & " where " & busqueda
       If Len(Text1) > 0 Then
          buf = buf & " and " & Text1 & " "
       End If
    End If
    If Len(busqueda) = 0 And Len(Text1) > 0 Then
       buf = buf & " where " & Text1
    End If
    If Len(ordenado) > 0 Then
       buf = buf & " order by " & ordenado
    End If
    End If
    Set mysnap = mydby.CreateSnapshot(buf)
    sqllo = 1
    Screen.MousePointer = 1
    Exit Function
cmd21_err:
 MsgBox "FORMATO NO VALIDO " & error$, 24, "AVISO"
 Screen.MousePointer = 1
 Exit Function

End Function

Private Sub List1_DblClick()
   Dim mytablex As Table
   Dim mydbx As Database
   Set mydb3 = OpenDatabase(generador, False, False, "foxpro 2.5;")
   Set mydbx = OpenDatabase(generador, False, False, "foxpro 2.5;")
   Set mytablex = mydbx.OpenTable("reporte")
   mytablex.AddNew
   mytablex.Fields("usuario") = usuariopos
   mytablex.Fields("nombre") = nombrebased & "." & List1.Text
   mytablex.Fields("orden") = Trim(List1.ListIndex)
   mytablex.Fields("tipo") = Val("" & mydb3.TableDefs(nombrebased).Fields(List1.ListIndex).Type)
   Select Case Val("" & mydb3.TableDefs(nombrebased).Fields(List1.ListIndex).Type)
      Case 7
           mytablex.Fields("estados") = "2"
   End Select
   mytablex.Fields("longitud") = Val("" & mydb3.TableDefs(nombrebased).Fields(List1.ListIndex).Size)
   mytablex.Fields("local") = local_1
   mytablex.Update
   mytablex.Close
   mydbx.Close
   mydb3.Close
   'Set mytable = Nothing
   'Set mydb = Nothing
   'Set mydb3 = Nothing
   refresca
'MsgBox list1.Text 'list1.ListIndex

End Sub

Private Sub List2_DblClick()
If Len(namejoin) = 0 Then Exit Sub
If Val(nro) < 1 Then
   MsgBox "No existe campos seleccionados", 24, "Aviso"
   Exit Sub
End If
   Set mydb3 = OpenDatabase(generador, False, False, "foxpro 2.5;")
   Set mydb = OpenDatabase(generador, False, False, "foxpro 2.5;")
   Set mytable = mydb.OpenTable("reporte")
   mytable.AddNew
   mytable.Fields("usuario") = usuariopos
   mytable.Fields("nombre") = namejoin & "." & List2.Text
   mytable.Fields("orden") = Trim(List2.ListIndex)
   mytable.Fields("tipo") = Val("" & mydb3.TableDefs(namejoin).Fields(List2.ListIndex).Type)
   Select Case Val("" & mydb3.TableDefs(namejoin).Fields(List2.ListIndex).Type)
      Case 7
           mytable.Fields("estados") = "2"
   End Select
   mytable.Fields("longitud") = Val("" & mydb3.TableDefs(namejoin).Fields(List2.ListIndex).Size)
   mytable.Fields("local") = local_1
   mytable.Update
   mytable.Close
   mydb.Close
   mydb3.Close
   'Set mytable = Nothing
   'Set mydb = Nothing
   'Set mydb3 = Nothing
   refresca
'MsgBox list1.Text 'list1.ListIndex
End Sub

Private Sub nombrebased_Click()
Dim found As Integer
If Len(nombrebased) = 0 Then Exit Sub
asw = "0"
asw = "1"
List1.Clear
found = CARGA_BASE()
End Sub

Private Sub table1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = &H2E And table1.EditActive = False And Len(table1.ColumnText(2)) > 0 Then 'delete
   Data1.Recordset.Delete
   Data1.Refresh
   Exit Sub
End If
End Sub
