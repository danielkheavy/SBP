VERSION 5.00
Begin VB.Form expword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar a Word"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "expword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MSWord As New Word.Application
Dim Documento As Object

Private Sub Command1_Click()
'declaramos los objetos
  

  
                'Establecemos la ruta de nuestro archivo
                ruta = App.Path & "\orden.doc"
  
                'Seteamos el archivo al objeto documento
                Set Documento = MSWord.Documents.Open(ruta)
  
                'opcionalmente podemos guardar el archivo
                'en mi caso lo guardo con una extensión diferente (cab|tmp|pot|etc)
                MSWord.Selection.Document.SaveAs (App.Path & "\printme.cab")
  
                'Establecemos la fuentre que utilizaremos
                MSWord.Selection.Font.Name = "Arial"
  
                'Configuramos la alineacion de nuestro parrafo
                MSWord.Selection.Paragraphs.Alignment = wdAlignParagraphCenter
  
                'Activamos la fuente en Negrita
                MSWord.Selection.Font.Bold = True
  
                'Y el tamaño a 16 puntos
                MSWord.Selection.Font.Size = 16
  
                'con esta opcion podemos comenzar a escribir dentro de nuestro docuemnto
                MSWord.Selection.TypeText "Aqui podemos escribir el texto en el documento" & vbCrLf
  
                'Declaramos una tabla de 1 fila por 3 columnas
                MSWord.Selection.Tables.Add MSWord.Selection.Range, 1, 3
  
                'Seleccionamos la celda 1,2
                MSWord.Selection.Tables(1).Cell(1, 2).Select
  
                'establecemos el ancho de la celda
               MSWord.Selection.Tables(1).Cell(1, 2).Width = 70
  
                'configuramos los bordes
                MSWord.Selection.Tables(1).Cell(1, 2).Borders(wdBorderTop).Visible = True
                MSWord.Selection.Tables(1).Cell(1, 2).Borders(wdBorderLeft).Visible = True
                MSWord.Selection.Tables(1).Cell(1, 2).Borders(wdBorderBottom).Visible = True
                MSWord.Selection.Tables(1).Cell(1, 2).Borders(wdBorderRight).Visible = True
  
                'Y la alineación del texto dentro de la celda
                MSWord.Selection.Paragraphs.Alignment = wdAlignParagraphLeft
  
                'Seguido escribimos texto en dicha celda
                MSWord.Selection.TypeText "Nombre"
  
                'seleccionamos la celda 1,3
                MSWord.Selection.Tables(1).Cell(1, 3).Select
  
                'Establcemos el color de fondo de la celda (Trama)
                MSWord.Selection.Cells.Shading.BackgroundPatternColor = wdColorGray20
  
                'Escribimos en dicha celda
                MSWord.Selection.TypeText "nombre2"
  
                'esta opcion nos permite salir de la edición de la tabla, o bajar una fila
                MSWord.Selection.MoveDown
  
                'por ultimo mostramos el documento de word
                MSWord.Visible = True
  
                'vaciamos los objetos de la  memoria
                Set Documento = Nothing
                Set MSWord = Nothing
  


End Sub

Private Sub Form_Load()

End Sub
