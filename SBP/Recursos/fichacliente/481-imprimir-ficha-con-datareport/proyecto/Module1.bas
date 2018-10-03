Attribute VB_Name = "Module1"
 Option Explicit

Private Pic As IPictureDisp

Public Carpeta_IMG As String
Public Field_Img As String


Private Declare Function SetErrorMode _
    Lib "kernel32" ( _
    ByVal wMode As Long) As Long

Private Declare Sub InitCommonControls Lib "Comctl32" ()

Sub Main()

    Call InitCommonControls

    Call SetErrorMode(2)

    Form1.Show

End Sub

Sub Dibujar_Imagen(Objeto As Object, Path_Imagen As String)

On Error GoTo ErrSub

Dim Pos_x As Single
Dim Pos_y As Single
Dim Ancho_IMG As Single
Dim Alto_IMG As Single
Dim Ancho_Obj As Single
Dim Alto_Obj As Single
Dim Old_Scale As Single


    Set Pic = LoadPicture(Path_Imagen)

    With Objeto
    
    .AutoRedraw = True
    .Cls
    
    Old_Scale = .ScaleMode
    
    .ScaleMode = vbPixels
    Ancho_IMG = .ScaleX(Pic.Width, vbHimetric, vbPixels)
    Alto_IMG = .ScaleY(Pic.Height, vbHimetric, vbPixels)
    
    Ancho_Obj = .ScaleWidth
    Alto_Obj = .ScaleHeight
    
    If Ancho_IMG > Ancho_Obj Then
        Alto_IMG = Alto_IMG * Ancho_Obj / Ancho_IMG
        Ancho_IMG = Ancho_Obj
    End If
    If Alto_IMG > Alto_Obj Then
        Ancho_IMG = Ancho_IMG * Alto_Obj / Alto_IMG
        Alto_IMG = Alto_Obj
    End If
    Pos_x = (Ancho_Obj - Ancho_IMG) / 2
    Pos_y = (Alto_Obj - Alto_IMG) / 2
    
    End With
    

    Objeto.PaintPicture Pic, Pos_x, Pos_y, Ancho_IMG, Alto_IMG
    
    Objeto.ScaleMode = Old_Scale
    
    Exit Sub
    
'Error
ErrSub:
    
    If Err.Number = 76 Then
       Objeto.Cls
       Exit Sub
    Else
       MsgBox Err.Description, vbCritical
    End If
End Sub





