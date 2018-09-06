Sub OpenWord()
    Dim Namedoc As String
    Namedoc = "\\Sbs01\кб\Заказной участок\ДОК-ТЫ\ОФОРМЛЕНИЕ\Расходная накладная.doc"
    If Dir(Namedoc) = "" Then
        MsgBox "Накладная не найдена !", vbExclamation
        Exit Sub
        End If

    If IsOpen(Namedoc) Then
        MsgBox "Накладную кто-то редактирует !", vbExclamation
        Exit Sub
        End If

    Dim objWrdApp As Object, objWrdDoc As Object
    'создаем новое приложение Word
    Set objWrdApp = CreateObject("Word.Application")
    'открываем документ Word - документ "Doc1.doc" должен существовать
    Set objWrdDoc = objWrdApp.Documents.Open(Namedoc)
    
   Dim MyRange
   Set MyRange = objWrdDoc.Range
   With MyRange.Find
        .ClearFormatting
        .Text = "№ * "
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute
        If .Found Then
            objWrdDoc.Range(MyRange.Start + 2, MyRange.End - 1).Copy
        Else
            MsgBox "Номер накладной не найден !", vbExclamation
            objWrdDoc.Close True
            objWrdApp.Quit
            Set objWrdDoc = Nothing: Set objWrdApp = Nothing
            Exit Sub
        End If
    End With
    
    Dim Number As New DataObject
    Number.GetFromClipboard
    
    With MyRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "№ * "
        On Error Resume Next
        .Replacement.Text = "№ " & CStr(CInt(Number.GetText) + 1) & " "
        If Err Then
            MsgBox "Не правильно записан номер накладной !", vbExclamation
            objWrdDoc.Close True
            objWrdApp.Quit
            Set objWrdDoc = Nothing: Set objWrdApp = Nothing
            Exit Sub
            End If
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=2 'wdReplaceAll
        If .Found Then
            Sheets("Расх. накладная").Cells(2, 4).Value = CInt(Number.GetText) + 1 ' здесь вписать свою строку
            MsgBox "Операция прошла успешно", vbExclamation
        Else
            MsgBox "Не удалось записать Расходную накладную", vbExclamation
            objWrdDoc.Close True
            objWrdApp.Quit
            Set objWrdDoc = Nothing: Set objWrdApp = Nothing
            Exit Sub
        End If
    End With
    
    'закрываем документ Word с сохранением
    objWrdDoc.Close True
    'закрываем приложение Word
    objWrdApp.Quit
    'очищаем переменные Word - обязательно!
    Set objWrdDoc = Nothing: Set objWrdApp = Nothing
End Sub

Function IsOpen(File$) As Boolean
 Dim FN%
 FN = FreeFile
 On Error Resume Next
 Open File For Random Access Read Write Lock Read Write As #FN
 Close #FN
 IsOpen = Err
End Function
  