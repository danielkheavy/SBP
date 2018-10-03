Attribute VB_Name = "Module11"
Option Explicit

Dim I As Integer

Dim j As Integer

Private Declare Sub ExitProcess Lib "kernel32.dll" (ByVal uExitCode As Long)

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Public Function FileCheck(path$) As Boolean
    'USAGE: If FileCheck("C:\windows\kewl.exe") then msgbox "it was found"
    FileCheck = True 'Assume Success

    On Error Resume Next

    Dim Disregard As Long

    Disregard = FileLen(path)

    If Err <> 0 Then FileCheck = False

End Function

Function ResizeElements()
    
    On Error Resume Next
    
    tcxbarra.WorkBar.Width = (tcxbarra.ScaleWidth - tcxbarra.LayerBar.ScaleWidth)
    tcxbarra.label.Left = (tcxbarra.WorkBar.ScaleWidth - tcxbarra.label.Width) / 2
    tcxbarra.label.Top = (tcxbarra.WorkBar.ScaleHeight - tcxbarra.label.Height) / 2
    tcxbarra.LabelShadow.Left = ((tcxbarra.WorkBar.ScaleWidth - tcxbarra.LabelShadow.Width) / 2) + 0.05
    tcxbarra.LabelShadow.Top = ((tcxbarra.WorkBar.ScaleHeight - tcxbarra.LabelShadow.Height) / 2) + 0.06
    tcxbarra.Frame3.Height = (tcxbarra.ScaleHeight - tcxbarra.StatusBar1.Height) - 200
    tcxbarra.ListView1.Height = (tcxbarra.ScaleHeight - tcxbarra.StatusBar1.Height) - 700
    
    'resize scrollbars
    tcxbarra.VScroll1.Height = tcxbarra.WorkBar.ScaleHeight - 0.44
    tcxbarra.VScroll1.Left = tcxbarra.WorkBar.ScaleWidth - tcxbarra.VScroll1.Width
    tcxbarra.VScroll1.Top = 0
    tcxbarra.HScroll1.Width = tcxbarra.WorkBar.ScaleWidth - 0.44
    tcxbarra.HScroll1.Left = 0
    tcxbarra.Picture1.Left = tcxbarra.WorkBar.ScaleWidth - tcxbarra.Picture1.Width
    tcxbarra.Picture1.Top = tcxbarra.WorkBar.ScaleHeight - tcxbarra.Picture1.Height
    tcxbarra.HScroll1.Top = tcxbarra.WorkBar.ScaleHeight - tcxbarra.HScroll1.Height
    
    tcxbarra.BarTop.Width = tcxbarra.label.Width
    tcxbarra.BarLeft.Height = tcxbarra.label.Height
    tcxbarra.BarEmpty.Width = tcxbarra.BarLeft.Width
    tcxbarra.BarEmpty.Height = tcxbarra.BarTop.Height
    
    'always place on top
    tcxbarra.HorizontalCenterLine.ZOrder (0)
    tcxbarra.VerticalCenterLine.ZOrder (0)
    
    'left and top bar
    tcxbarra.BarLeft.Left = ((tcxbarra.WorkBar.ScaleWidth - tcxbarra.label.Width) / 2) - tcxbarra.BarLeft.Width
    tcxbarra.BarTop.Left = ((tcxbarra.WorkBar.ScaleWidth - tcxbarra.label.Width) / 2)
    tcxbarra.BarEmpty.Left = ((tcxbarra.WorkBar.ScaleWidth - tcxbarra.label.Width) / 2) - tcxbarra.BarEmpty.Width
    tcxbarra.BarTop.Top = ((tcxbarra.WorkBar.ScaleHeight - tcxbarra.label.Height) / 2) - tcxbarra.BarTop.Height
    tcxbarra.BarLeft.Top = ((tcxbarra.WorkBar.ScaleHeight - tcxbarra.label.Height) / 2)
    tcxbarra.BarEmpty.Top = ((tcxbarra.WorkBar.ScaleHeight - tcxbarra.label.Height) / 2) - tcxbarra.BarEmpty.Height

    'center lines
    tcxbarra.HorizontalCenterLine.X1 = 0
    tcxbarra.HorizontalCenterLine.X2 = tcxbarra.label.Width
    tcxbarra.HorizontalCenterLine.Y1 = (tcxbarra.label.Height / 2)
    tcxbarra.HorizontalCenterLine.Y2 = (tcxbarra.label.Height / 2)
    tcxbarra.VerticalCenterLine.X1 = (tcxbarra.label.Width / 2)
    tcxbarra.VerticalCenterLine.X2 = (tcxbarra.label.Width / 2)
    tcxbarra.VerticalCenterLine.Y1 = 0
    tcxbarra.VerticalCenterLine.Y2 = tcxbarra.label.Height
    
    'bar numbers and lines
    'BarTop
    For I = 0 To Round(tcxbarra.label.Width, 0)
        tcxbarra.BarNumberWidth.Caption = I
        'set number position
        tcxbarra.BarTop.CurrentX = (I + 0.5) - (tcxbarra.BarNumberWidth.Width / 2) 'left
        tcxbarra.BarTop.CurrentY = 0 'top
        'ignore if is 0 or maximal
        'print number onto it
        tcxbarra.BarTop.Print I
        'draw line
        tcxbarra.BarTop.Line (I, 0.2)-(I, tcxbarra.BarTop.Height)

        'draw small lines
        For j = 1 To 9
            tcxbarra.BarTop.Line (I + (j / 10), 0.3)-(I + (j / 10), tcxbarra.BarTop.Height)
        Next j
    Next I

    'BarLeft
    For I = 0 To tcxbarra.label.Height
        tcxbarra.BarNumberWidth.Caption = I
        'set number position
        tcxbarra.BarLeft.CurrentX = 0 'left
        tcxbarra.BarLeft.CurrentY = (I + 0.5) - (tcxbarra.BarNumberWidth.Height / 2) 'top
        'print number onto it
        tcxbarra.BarLeft.Print I
        'draw line
        tcxbarra.BarLeft.Line (0.2, I)-(tcxbarra.BarLeft.Width, I)

        'draw small lines
        For j = 1 To 9
            tcxbarra.BarLeft.Line (0.3, I + (j / 10))-(tcxbarra.BarLeft.Width, I + (j / 10))
        Next j
    Next I
    
    'check scrollbars status
    If tcxbarra.label.ScaleWidth > tcxbarra.WorkBar.ScaleWidth Then
        tcxbarra.HScroll1.Enabled = True
    Else
        tcxbarra.HScroll1.Enabled = False

    End If
    
    If tcxbarra.label.ScaleHeight > tcxbarra.WorkBar.ScaleHeight Then
        tcxbarra.VScroll1.Enabled = True
    Else
        tcxbarra.VScroll1.Enabled = False

    End If

End Function

Function UnloadGuidelines()

    'Unload all Guidelines
    For I = 1 To tcxbarra.HorizontalGuideline.UBound
        Unload tcxbarra.HorizontalGuideline(I)
    Next I

    For I = 1 To tcxbarra.VerticalGuideline.UBound
        Unload tcxbarra.VerticalGuideline(I)
    Next I

End Function

Function InsertGuideline(IsType As Integer, X1, X2, Y1, Y2)

    Dim I As Integer

    'Horizontal
    If IsType = 1 Then
        Load tcxbarra.HorizontalGuideline(tcxbarra.HorizontalGuideline.count)

        With tcxbarra.HorizontalGuideline(tcxbarra.HorizontalGuideline.UBound)
            .X1 = X1
            .X2 = X2
            .Y1 = Y1
            .Y2 = Y2
            .Visible = True

        End With

        'Vertical
    ElseIf IsType = 2 Then
        Load tcxbarra.VerticalGuideline(tcxbarra.VerticalGuideline.count)

        With tcxbarra.VerticalGuideline(tcxbarra.VerticalGuideline.UBound)
            .X1 = X1
            .X2 = X2
            .Y1 = Y1
            .Y2 = Y2
            .Visible = True

        End With

    End If

End Function

Private Function AppPath(ByVal zPath As String) As String

    If Right$(zPath, 1) = "\" Then AppPath = zPath Else AppPath = zPath & "\"

End Function

Private Function FileExist(ByVal strPath As String) As Boolean
    On Local Error GoTo ErrFile
    Open strPath For Input Access Read As #1
    Close #1
    FileExist = True
    Exit Function
ErrFile:
    FileExist = False

End Function

Private Sub MakeManifest()

    Dim file$, file2$, qwe As String

    Exit Sub
    file$ = AppPath(App.path) & App.EXEName & ".exe.MANIFEST"

    If Not FileExist(file$) Then
        qwe = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbcrlf & "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbcrlf & "<assemblyIdentity type=""win32"" processorArchitecture=""*"" version=""6.0.0.0"" name=""name""/>" & vbcrlf & "<description>Enter your Description Here</description>" & vbcrlf & "<dependency>" & vbcrlf & "   <dependentAssembly>" & vbcrlf & "      <assemblyIdentity" & vbcrlf & "           type=""win32""" & vbcrlf & "           name=""Microsoft.Windows.Common-Controls"" version=""6.0.0.0""" & vbcrlf & "           language=""*""" & vbcrlf & "           processorArchitecture=""*""" & vbcrlf & "         publicKeyToken=""6595b64144ccf1df""" & vbcrlf & "      />" & vbcrlf & "   </dependentAssembly>" & vbcrlf & "</dependency>" & vbcrlf & "</assembly>" & vbcrlf
        Open file$ For Binary Access Write Lock Write As #1 Len = 1
        Put #1, , qwe
        Close #1
        SetAttr file$, vbReadOnly Or vbHidden ' Or vbSystem
        file2$ = AppPath(App.path) & App.EXEName & ".exe"
        Shell file2$, vbNormalFocus
        ExitProcess 1

    End If

End Sub

Public Sub InitControlsXP()
    MakeManifest
    InitCommonControls

End Sub

