Attribute VB_Name = "Module16"

'---------------------------------------------------------------------
'Copia un archivo.
'Sintaxis
'FileCopy fuente, destino [, sobre escritura]
'Argumentos:
'SourceFile: nombre completo de un archivo a copiarse
'DestinationPath: nombre de la ruta de destino
'OverWrite: Opción de especificar sobre escritura.
'Por Harvey Triana, Petrosoft Co., 1996
'---------------------------------------------------------------------
Sub PetrosoftCopyFile(SourceFile As String, _
                      DestinationPath As String, _
                      Optional OverWrite As Variant)

    Const INLINE = 2 ^ 10
    
    Dim Tem    As String

    Dim I      As Integer

    Dim RCnl   As Integer

    Dim WCnl   As Integer

    Dim Bytes  As Long

    Dim Groups As Long

    Dim SBytes As Long
    
    'Este bloque verifica la sobre escritura si al archivo exise
    Tem = DestinationPath + FileNameFromPath(SourceFile)
    
    If Len(Dir(Tem)) Then
        If IsMissing(OverWrite) Then
            Kill Tem
        Else

            If OverWrite Then
                Kill Tem
            Else
                Exit Sub

            End If

        End If

    End If
       
    RCnl = FreeFile
    Open SourceFile For Binary Access Read As #RCnl
    WCnl = FreeFile
    Open Tem For Binary Access Write As #WCnl
    
    'Copia por grupos de bytes
    Bytes = LOF(RCnl)
    Groups = Int((Bytes / INLINE))
    SBytes = (Bytes - Groups * INLINE)

    If Groups > 0 Then

        For I = 1 To Groups
            Tem = input$(INLINE, #RCnl)
            Put #WCnl, , Tem
        Next

    End If

    If SBytes > 0 Then
        Tem = input$(SBytes, #RCnl)
        Put #WCnl, , Tem

    End If

    Close RCnl, WCnl

End Sub

Public Function FileNameFromPath(Tem As String) As String
    
    Dim X As String, I
    
    If InStr(Tem, "\") Then
        I = Len(Tem)
        Do
            X = Mid$(Tem, I, 1)
            I = I - 1
        Loop Until X = "\" Or I = 0

        FileNameFromPath = Mid$(Tem, I + 2)
    Else
        FileNameFromPath = Tem

    End If

End Function
