VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Otros datos"
      Height          =   2415
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   2640
      Width           =   5415
      Begin VB.TextBox txt_Field 
         Height          =   375
         Index           =   6
         Left            =   1560
         TabIndex        =   31
         Top             =   1800
         Width           =   3615
      End
      Begin VB.TextBox txt_Field 
         Height          =   375
         Index           =   5
         Left            =   1560
         TabIndex        =   26
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txt_Field 
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   25
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txt_Field 
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   24
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Obra social"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   32
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   30
         Top             =   480
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Dni"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Direccion"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   28
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Email"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   27
         Top             =   1320
         Width           =   375
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   7320
      TabIndex        =   22
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Imagen"
      Height          =   2535
      Left            =   5640
      TabIndex        =   20
      Top             =   120
      Width           =   3135
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   240
         ScaleHeight     =   1755
         ScaleWidth      =   2475
         TabIndex        =   21
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos principales"
      Height          =   2415
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5415
      Begin VB.TextBox txt_Field 
         Height          =   375
         Index           =   0
         Left            =   1560
         TabIndex        =   14
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txt_Field 
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   13
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox txt_Field 
         Height          =   375
         Index           =   2
         Left            =   1560
         TabIndex        =   12
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Id"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Telefono"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   1800
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Apellido"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   555
      End
      Begin VB.Label lblID 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   15
         Top             =   480
         Width           =   45
      End
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   ">>"
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   10
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   ">"
      Height          =   375
      Index           =   2
      Left            =   1680
      TabIndex        =   9
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "<"
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   8
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdNav 
      Caption         =   "<<"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eliminar imagen"
      Height          =   375
      Index           =   6
      Left            =   7320
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Seleccionar imagen"
      Height          =   375
      Index           =   5
      Left            =   5760
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   4
      Left            =   7320
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Editar"
      Height          =   375
      Index           =   3
      Left            =   7320
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar"
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   2
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eliminar"
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   1
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nuevo"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuVerTodo 
         Caption         =   "Ver todos los registros"
      End
      Begin VB.Menu mnuLinea 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As Connection
Dim rst As Recordset

' Primer registro, siguiente, etc...
Private Sub cmdNav_Click(Index As Integer)

    ' Si hay registro activo sale
    If rst.BOF And rst.EOF Then Exit Sub

    Select Case Index
        Case 0
            rst.MoveFirst
        Case 1
            rst.MovePrevious
            If rst.BOF Then rst.MoveFirst
        Case 2
            rst.MoveNext
            If rst.EOF Then rst.MoveLast
        Case 3
            rst.MoveLast
    End Select

    ' Carga la imagen en el Picture
    Call Mostrar_Imagen

End Sub

Private Sub Command1_Click(Index As Integer)

    Select Case Index
 
        'Agrega un nuevo registro
        Case 0
            rst.AddNew
            Picture1.Cls
            'Elimina el registro activo
            
            CmdNuevo
            
        Case 1
            If rst.EOF Or rst.BOF Then Exit Sub
            If MsgBox("Eliminar Registro", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
            Picture1.Cls
    
            'Elimina el archivo de la carpeta de imagenes
            If rst(Field_Img) <> "" Then
                Call Kill(Carpeta_IMG & rst(Field_Img))
            End If
        
            rst.Delete
            
            If rst.RecordCount > 0 Then
               cmdNormal
            Else
               cmdSinRegistros
            End If
            
            If rst.EOF Or rst.BOF Then
                Exit Sub
            End If
            rst.MoveNext
            
            If rst.EOF Then
               On Error Resume Next
               rst.MoveLast
            End If
            'Carga la imagen del registro activo
            Mostrar_Imagen
            Exit Sub
             
        ' Botón Actualizar los cambios en la base de datos
        Case 2
            If Not rst.EOF And Not rst.BOF Then
                rst.Update
                Guardar_Imagen
                cmdNormal
            End If

        ' Cancela la atualización o edición del registro que se editando o añadiendo
        Case 3
            cmdEditar
            Setear_TextBox
            Exit Sub
  
        'Botón Editar el registro activo
        Case 4
            If rst.EOF And rst.BOF Then Exit Sub
            rst.CancelUpdate
  
            If Not rst.BOF And Not rst.EOF Then
                If rst(Field_Img) <> "" Then
                    Call Dibujar_Imagen(Picture1, Carpeta_IMG & rst(Field_Img))
                End If
                
            End If
            
            If rst.RecordCount > 0 Then
                cmdNormal
            Else
                cmdSinRegistros
            End If
        'Carga una imagen en el control Picture1
        Case 5
  
            With CommonDialog1
                .DialogTitle = " Seleccionar imagen"
                .Filter = "BMP|*.bmp|JPEG|*.jpeg|GIF|*.gif|JPG|*.jpg|Todos|*.*"
     
                .ShowOpen
     
                If .FileName = "" Then
                    Exit Sub
                Else
         
                    ' Graba el nombre en el campo, el id de imagen _
                    que es el mismo que el campo Id
         
                    rst(Field_Img) = rst!id '
         
        
                    ' se dibuja la imagen en el Picture
                    Call Dibujar_Imagen(Picture1, .FileName)
         
                End If
            End With
            
            Exit Sub

        Case 6

            ' Limpia la imagen del Picture y Elimina el id de _
            imagen del registro actual de la base
            
            If MsgBox("Desea eliminar la imagen ?", vbYesNo + vbQuestion) = vbYes Then
               Picture1.Cls
               rst(Field_Img) = ""
               Exit Sub
            End If

    End Select

    
    Setear_TextBox

    ' Muestra la imagen
    Mostrar_Imagen

End Sub

Sub Guardar_Imagen()


    ' Si el campo Id_Imagen no está vacio ...
    If rst(Field_Img) <> "" And CommonDialog1.FileName <> "" Then
        ' Copia el archivo a la carpeta de imagen
        Call FileCopy(CommonDialog1.FileName, _
                      Carpeta_IMG & "\" & rst!id)

        '... si no, si el archivo está en lacarpeta lo  elimina

    ElseIf Dir(Carpeta_IMG & "\" & rst!id) <> "" And rst(Field_Img) = "" Then
       Call Kill(Carpeta_IMG & rst!id)

    End If
End Sub


Private Sub Mostrar_Imagen()

    With rst
        ' Si no hay ningún registro activo sale
        If .EOF Or .BOF Then
            Exit Sub
        End If
        
        ' Si el registro no tiene una imagen asociada Limpia el Picture
        If .Fields(Field_Img) = "" Or .Fields(Field_Img) = 0 Then
           Picture1.Cls
        Else
           ' Lee el archivo de imagen y lo dibuja en el Picture
            Call Dibujar_Imagen(Picture1, Carpeta_IMG & .Fields(Field_Img))
        End If

        'Me.Caption = "Registro N°: " & CStr(.AbsolutePosition)

    End With

End Sub

Private Sub Setear_TextBox()
    'Bloquea y desbloquea los textbox
    Dim T As TextBox
    For Each T In Me.txt_Field
        T.Locked = Not T.Locked
    Next
End Sub

' Habilita y deshabilita los CommandButton

Private Sub Setear_botones()

    Dim i As Integer

    For i = 0 To Command1.Count - 1
        Command1(i).Enabled = Not Command1(i).Enabled
    Next

    For i = 0 To cmdNav.Count - 1
        cmdNav(i).Enabled = Not cmdNav(i).Enabled
    Next

End Sub


Private Sub Imprimir()
    
Dim rsFicha As ADODB.Recordset
    
    Set rsFicha = New Recordset

    rsFicha.Open "Select * FROM clientes Where Id=" & lblID.Caption, cn, adOpenStatic, adLockReadOnly
    
    If rsFicha.RecordCount > 0 Then
        
       Set DataReport1.DataSource = rsFicha
        
       With DataReport1
            If rsFicha!id_Imagen <> "" Then
            
                .Sections.Item("Sección1").Controls("lblSinFoto").Visible = False
                Set .Sections.Item("Sección1").Controls("rptImagen").Picture = Picture1.Image
            Else
                .Sections.Item("Sección1").Controls("lblSinFoto").Visible = True
            End If
            DataReport1.Show
        End With
    Else
       MsgBox "No hay registro para imprimir ", vbInformation
    End If
    
End Sub

Private Sub Command2_Click()
    Call Imprimir
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
        Set rst = Nothing
    End If
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
End Sub

Private Sub Form_Load()

    Dim Pathbd As String, cadena As String
    Dim T As TextBox
    
    Set cn = New Connection

    Pathbd = App.Path & "\db1.mdb"

    cadena = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Pathbd & _
                                     ";Persist Security Info=False"

    
    cn.Open cadena

    Set rst = New Recordset

    rst.Open "Select * FROM clientes Order by Apellido", cn, adOpenStatic, adLockOptimistic

    ' Nombre del campo  que tiene el ID de imagen
    Field_Img = "ID_Imagen"
    ' Path de la carpeta donde están las imagenes
    Carpeta_IMG = App.Path & "\img\"

    ' Si no existe la carpeta para guardar las imagen la crea
    If Dir(App.Path & "\img", vbDirectory) = "" Then
        MkDir App.Path & "\img"
    End If
    
    If rst.RecordCount > 0 Then
        Call cmdNormal
    Else
        Call cmdSinRegistros
    End If
    
    Set txt_Field(0).DataSource = rst
    Set txt_Field(1).DataSource = rst
    Set txt_Field(2).DataSource = rst
    Set txt_Field(3).DataSource = rst
    Set txt_Field(4).DataSource = rst
    Set txt_Field(5).DataSource = rst
    Set txt_Field(6).DataSource = rst
    
    
    txt_Field(0).DataField = "Nombre"
    txt_Field(1).DataField = "Apellido"
    txt_Field(2).DataField = "Telefono"

    txt_Field(3).DataField = "Dni"
    txt_Field(4).DataField = "Direccion"
    txt_Field(5).DataField = "Email"
    txt_Field(6).DataField = "Obra social"


    'Opcional: esto visualiza el Id del registro en un label
    Set lblID.DataSource = rst
    lblID.DataField = "Id"

    Call Setear_TextBox

    ' carga la imagen en el registro si es que tiene
    Call Mostrar_Imagen

End Sub


Sub cmdNormal()

    DeshabilitarTodosCmd

    Command1(0).Enabled = True
    Command1(1).Enabled = True
    Command1(3).Enabled = True
    
End Sub

Sub cmdSinRegistros()

    DeshabilitarTodosCmd
    Command1(0).Enabled = True

End Sub

Sub cmdEditar()
        
    DeshabilitarTodosCmd
    Command1(2).Enabled = True
    Command1(4).Enabled = True
    Command1(5).Enabled = True
    Command1(6).Enabled = True
    
End Sub

Sub CmdNuevo()
    DeshabilitarTodosCmd
    Command1(2).Enabled = True
    Command1(4).Enabled = True
    
    Command1(5).Enabled = True
    Command1(6).Enabled = True
End Sub

Sub DeshabilitarTodosCmd()
    Command1(0).Enabled = False
    Command1(1).Enabled = False
    Command1(2).Enabled = False
    Command1(3).Enabled = False
    Command1(4).Enabled = False
    Command1(5).Enabled = False
    Command1(6).Enabled = False
    
End Sub

Private Sub mnuImprimir_Click()
    Call Imprimir
End Sub

Private Sub mnuVerTodo_Click()
    With Form2
         Set .MSHFlexGrid1.DataSource = rst
        .Show vbModal
    End With
End Sub
