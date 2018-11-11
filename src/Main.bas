Attribute VB_Name = "Main"
'Written in 2015-2018 by Eduard E. Tikhenko <aquaried@gmail.com>
'
'To the extent possible under law, the author(s) have dedicated all copyright
'and related and neighboring rights to this software to the public domain
'worldwide. This software is distributed without any warranty.
'You should have received a copy of the CC0 Public Domain Dedication along
'with this software.
'If not, see <http://creativecommons.org/publicdomain/zero/1.0/>

Option Explicit

Private Const appName As String = "AddRows"
Private Const section As String = "Main"
Private Const configFileName As String = "Rows.conf"
Const configNameSections As String = "[Sections]"
Const configNameItems As String = "[Items]"
Public configFullFileName As String

Dim swApp As Object
Dim swDoc As ModelDoc2
Public swTable As TableAnnotation

Enum profileFormatTitles
    idUnderline
    idUpper
End Enum

Enum profileFormatStrings
    idLeft
    idCenter
End Enum

Sub Main()
    Dim swProp As CustomPropertyManager
    Dim titles As Collection
    Dim items As Collection
    
    Set swApp = Application.SldWorks
    configFullFileName = swApp.GetCurrentMacroPathFolder + "\" + configFileName
    
    If swApp.GetDocumentCount > 0 Then
        Set swDoc = swApp.ActiveDoc
        If swDoc.GetType = swDocDRAWING Then
            If swDoc.GetPathName <> "" Then
                Set swProp = swDoc.Extension.CustomPropertyManager("")
                Set swTable = GetBOM(swDoc)
                If Not swTable Is Nothing Then
                    
                    Set titles = Nothing
                    Set items = Nothing
                    If Not GetRowsFromFile(titles, items) Then
                        MsgBox "Файл не найден!" & vbNewLine & configFullFileName, vbCritical
                        Exit Sub
                    End If
                    
                    MyWindow.lstTitles.Clear
                    MyWindow.lstItems.Clear
                    Dim i
                    For Each i In titles
                        MyWindow.lstTitles.AddItem i
                    Next
                    For Each i In items
                        MyWindow.lstItems.AddItem i
                    Next
                
                    MyWindow.Show
                Else
                    MsgBox "Не найдена спецификация.", vbCritical
                End If
            Else
                MsgBox "Безымянный чертёж.", vbCritical
            End If
        Else
            MsgBox "Открытый документ не является чертежом.", vbCritical
        End If
    Else
        MsgBox "Ничего не открыто.", vbCritical
    End If

End Sub

Function GetRowsFromFile(ByRef titles As Collection, ByRef items As Collection) As Boolean
    Dim objStream As Stream
        
    Set objStream = New Stream
    objStream.Charset = "utf-8"
    objStream.Open
    GetRowsFromFile = False
    
    On Error GoTo CreateConfig
    objStream.LoadFromFile configFullFileName
    GoTo SuccessRead

ReadConfigAgain:
    On Error GoTo ExitFunction
    objStream.LoadFromFile configFullFileName
    GoTo SuccessRead
   
SuccessRead:
    ReadRowsFromFile titles, items, objStream
    GetRowsFromFile = True
ExitFunction:
    objStream.Close
    Set objStream = Nothing
    Exit Function
    
CreateConfig:
    CreateDefaultConfigFile objStream
    GoTo ReadConfigAgain
End Function

Sub CreateDefaultConfigFile(objStream As Stream)
    'TODO: check if cannot to create file
    objStream.WriteText _
        configNameSections & vbNewLine & _
        "Сборочные единицы" & vbNewLine & _
        "Детали" & vbNewLine & _
        "Стандартные изделия" & vbNewLine & _
        "Покупные изделия" & vbNewLine & _
        "Прочее" & vbNewLine & _
        "Материалы" & vbNewLine & _
        vbNewLine & _
        configNameItems & vbNewLine & _
        "Грунт" & vbNewLine & _
        "Эмаль"
    objStream.SaveToFile configFullFileName
End Sub

Sub ReadRowsFromFile(ByRef titles As Collection, ByRef items As Collection, objStream As Stream)
    Const RowIsSection As Integer = 1
    Const RowIsItem As Integer = 2
    Dim rowIs As Integer
    
    Dim strData
    
    Set titles = New Collection
    Set items = New Collection
    rowIs = 0
    Do Until objStream.EOS
        strData = Trim(objStream.ReadText(adReadLine))
        
        If Len(strData) < 1 Then
            GoTo NextDo
        End If
        
        Select Case strData
            Case configNameSections
                rowIs = RowIsSection
            Case configNameItems
                rowIs = RowIsItem
            Case Else
                If rowIs = RowIsSection Then
                    titles.Add strData
                ElseIf rowIs = RowIsItem Then
                    items.Add strData
                End If
        End Select
NextDo:
    Loop
End Sub

Function EditConfigFile() 'mask for button
    Shell "notepad " & configFullFileName, vbNormalFocus
End Function

Private Function GetBOM(ByRef swDoc As ModelDoc2) As TableAnnotation

    Dim swFeat As Feature, swBomFeat As BomFeature
    
    Set swFeat = swDoc.FirstFeature
    Set GetBOM = Nothing
    Do While Not swFeat Is Nothing
        If "BomFeat" = swFeat.GetTypeName Then
            Set swBomFeat = swFeat.GetSpecificFeature2
            Set GetBOM = swBomFeat.GetTableAnnotations(0) 'берется первая спецификация
            Exit Do
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    
End Function

Public Function ExitApp() As Boolean

    Unload MyWindow
    ExitApp = True

End Function

Public Function Run()  'mask for button
    Dim idFormat As profileFormatTitles
    Dim appendEmptyLines As Boolean
    Dim i As Integer
    Dim addedAnything As Boolean
    Dim idFormatStrings As profileFormatStrings
    
    idFormat = IIf(MyWindow.underlineBox.value, idUnderline, idUpper)
    appendEmptyLines = MyWindow.empBox.value
    idFormatStrings = IIf(MyWindow.boxLeft.value, idLeft, idCenter)
    
    addedAnything = False
    
    With MyWindow.lstItems
    For i = .ListCount - 1 To 0 Step -1
        addedAnything = addedAnything Or AddRowIf(.Selected(i), .list(i), idFormat, appendEmptyLines, idFormatStrings, False)
    Next
    End With
    
    With MyWindow.lstTitles
    For i = .ListCount - 1 To 0 Step -1
        addedAnything = addedAnything Or AddRowIf(.Selected(i), .list(i), idFormat, appendEmptyLines, idFormatStrings, True)
    Next
    End With
    
    If addedAnything Then
        swDoc.SetSaveFlag
    End If
End Function

Function AddRowIf(condition As Boolean, name As String, idFormat As profileFormatTitles, _
        appendEmptyLines As Boolean, idFormatStrings As profileFormatStrings, _
        isTitle As Boolean) As Boolean
    
    If condition Then
        AddRowIf = AddRowToBOM(name, idFormat, appendEmptyLines, idFormatStrings, isTitle)
    Else
        AddRowIf = False
    End If
End Function

Public Sub SaveSettingFormatTitles(idFormat As profileFormatTitles)
    SaveSetting appName, section, "Format", str(idFormat)
End Sub

Public Function GetSettingFormatTitles() As profileFormatTitles
    GetSettingFormatTitles = Int(GetSetting(appName, section, "Format", "0"))
End Function

Public Sub SaveSettingEmptyStrings(appendEmptyLines As Boolean)
    SaveSetting appName, section, "AppendEmptyLines", str(Int(appendEmptyLines))
End Sub

Public Function GetSettingEmptyStrings() As Boolean
    GetSettingEmptyStrings = CBool(Int(GetSetting(appName, section, "AppendEmptyLines", "0")))
End Function

Public Sub SaveSettingFormatStrings(idFormat As profileFormatStrings)
    SaveSetting appName, section, "FormatStrings", str(idFormat)
End Sub

Public Function GetSettingFormatStrings() As profileFormatStrings
    GetSettingFormatStrings = Int(GetSetting(appName, section, "FormatStrings", "0"))
End Function
