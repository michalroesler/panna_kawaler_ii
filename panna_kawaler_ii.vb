Sub panna_kawaler_ii()
    
    Dim wordApp As Word.Application
    Dim wordDoc As Word.Document
    Dim excelApp As Excel.Application
    Dim ws As Worksheet
    Dim mySheet As Worksheet
    Dim rng As Word.Range
    
    Dim frstWrd As String
    Dim lastWrd As String
    
    Dim startPos As Long
    Dim endPos As Long
    Dim Length As Long
    Dim textToFind1 As String
    Dim textToFind2 As String
    Dim textToFind3 As String
    Dim textToFind4 As String
    Dim textToFind5 As String
    Dim textToFind6 As String
    Dim arr() As String
    Dim arrr() As String
    
    Dim i As Long
    Dim x As Long
    Dim wrdCount1 As Integer
    Dim wrdCount2 As Integer
    
    'Assigning object variables and values
    Set wordApp = GetObject(, "Word.Application")       'At its simplest, CreateObject creates an instance of an object,
    Set excelApp = GetObject(, "Excel.Application")     'whereas GetObject gets an existing instance of an object.
    Set wordDoc = wordApp.ActiveDocument
    Set mySheet = Application.ActiveWorkbook.ActiveSheet
    Set ws = Application.ActiveWorkbook.ActiveSheet
    Set rng = wordApp.ActiveDocument.Content
    
    textToFind1 = "§ 5. (zobowiązania stron)"     '"KRS 0000609737, REGON 364061169, NIP 951-24-09-783,"   or   "REGON 364061169, NIP 951-24-09-783,"
    textToFind2 = "§ 5. (Zobowiązania stron)"
    textToFind3 = "2. Strony postanawiają, że w umowie ustanowienia odrębnej"                         'w umowach deweloperskich FW2 było "- ad."    or   "- ad."
    textToFind4 = "Tożsamość stawających"
    textToFind5 = "cenę brutto w kwocie "
    textToFind6 = "za podaną cenę zobowiązują się kupić na zasadach wspólności ustawowej majątkowej małżeńskiej."
    x = 12
    
    'Checking multiple conditions
    If Application.WorksheetFunction.CountA(mySheet.Range("A12:D15")) = 2 And _
       mySheet.Range("I12") = "k" And _
       mySheet.Range("E12") = "" Then
        'InStr function returns a Variant (Long) specifying the position of the first occurrence of one string within another.
        startPos = InStr(1, rng, textToFind1) - 1           'here we get 40296, we're looking 4 "textToFind1"
        If startPos < 500 Then startPos = InStr(1, rng, textToFind2) - 1    'here we get 60796, we're looking 4 "textToFind2"
        endPos = InStr(1, rng, textToFind3) - 1      'here we get 42460, we're looking 4 "textToFind3"
        rng.SetRange Start:=startPos, End:=endPos
        Debug.Print "Paragraphs.Count = " & rng.Paragraphs.Count
        rng.MoveEnd wdParagraph, -1
        Debug.Print rng.Text
        arr = Split(Range("A12").Value, " ", -1)     '3rd argument is limit and it's optional. Number of substrings to be returned; -1 indicates that all substrings are returned.
        wrdCount1 = UBound(arr()) + 1
        Debug.Print wrdCount1
            For i = 0 To wrdCount1 - 1
              Debug.Print arr(i)
            Next i
        Debug.Print arr(0) & " " & arr(wrdCount1 - 1) & " oświadcza ponadto, że jest panną."
        With rng.Find
            .Text = arr(0) & " " & arr(wrdCount1 - 1) & " oświadcza ponadto, że jest panną."
            .MatchWildcards = False
            .MatchCase = False
            .Forward = True
            .Execute
               If .Found = True Then
                  mySheet.Cells(x, 5) = mySheet.Range("AE12")
               Else
                  rng.SetRange Start:=startPos, End:=endPos
                  With rng.Find
                    .Text = arr(0) & " " & arr(wrdCount1 - 1) & " oświadcza ponadto, że jest rozwiedziona."
                    .MatchWildcards = False
                    .MatchCase = False
                    .Forward = True
                    .Execute
                        If .Found = True Then
                           mySheet.Cells(x, 5) = mySheet.Range("AE14")
                        Else
                           rng.SetRange Start:=startPos, End:=endPos
                           With rng.Find
                              .Text = arr(0) & " " & arr(wrdCount1 - 1) & " oświadcza ponadto, że pozostaje w związku małżeńskim"
                              .MatchWildcards = False
                              .MatchCase = False
                              .Forward = True
                              .Execute
                                 If .Found = True Then
                                     mySheet.Cells(x, 5) = mySheet.Range("AD11")
                                     rng.SetRange Start:=startPos, End:=endPos
                                         With rng.Find
                                             .Text = "w stanie wolnym od w/w obciążeń, za podaną cenę do majątku osobistego zobowiązuje się kupić."
                                             .MatchWildcards = False
                                             .MatchCase = False
                                             .Forward = True
                                             .Execute
                                                 If .Found = True Then mySheet.Cells(x, 6) = mySheet.Range("AE14")
                                         End With
                                 Else
                                     mySheet.Cells(x, 5) = "????"
                                 End If
                            End With
                        End If
                  End With
               End If
         End With
         
   ElseIf Application.WorksheetFunction.CountA(mySheet.Range("A12:D15")) = 2 And _
          mySheet.Range("I12") = "m" And _
          mySheet.Range("E12") = "" Then
        'InStr function returns a Variant (Long) specifying the position of the first occurrence of one string within another.
        startPos = InStr(1, rng, textToFind1) - 1    'here we get 40296, we're looking 4 "textToFind1"
        If startPos < 500 Then startPos = InStr(1, rng, textToFind2) - 1    'here we get 66796, we're looking 4 "textToFind2"
        endPos = InStr(1, rng, textToFind3) - 1      'here we get 42460, we're looking 4 "textToFind3"
        rng.SetRange Start:=startPos, End:=endPos
        Debug.Print "Paragraphs.Count = " & rng.Paragraphs.Count
        rng.MoveEnd wdParagraph, -1
        Debug.Print rng.Text
        arr = Split(Range("A12").Value, " ", -1)
        wrdCount1 = UBound(arr()) + 1
        Debug.Print wrdCount1
        
        For i = 0 To wrdCount1 - 1
          Debug.Print arr(i)
        Next i
        
        Debug.Print arr(0) & " " & arr(wrdCount1 - 1) & " oświadcza ponadto, że jest kawalerem."
        
        With rng.Find
            .Text = arr(0) & " " & arr(wrdCount1 - 1) & " oświadcza ponadto, że jest kawalerem."
            .MatchWildcards = False
            .MatchCase = False
            .Forward = True
            .Execute
               If .Found = True Then
                  mySheet.Cells(x, 5) = mySheet.Range("AE13")
               Else
                  rng.SetRange Start:=startPos, End:=endPos
                  With rng.Find
                    .Text = arr(0) & " " & arr(wrdCount1 - 1) & " oświadcza ponadto, że jest rozwiedziony."
                    .MatchWildcards = False
                    .MatchCase = False
                    .Forward = True
                    .Execute
                        If .Found = True Then
                           mySheet.Cells(x, 5) = mySheet.Range("AE14")
                        Else
                           mySheet.Cells(x, 5) = "????"
                        End If
                  End With
               End If
        End With
        
    ElseIf Application.WorksheetFunction.CountA(mySheet.Range("A12:D15")) = 4 And _
           Application.WorksheetFunction.CountIf(mySheet.Range("I12:I15"), "k") = 1 And _
           Application.WorksheetFunction.CountIf(mySheet.Range("I12:I15"), "m") = 1 Then
        'InStr function returns a Variant (Long) specifying the position of the first occurrence of one string within another.
        startPos = InStr(1, rng, textToFind1) - 1    'here we get 40296, we're looking 4 "textToFind1"
        If startPos < 500 Then startPos = InStr(1, rng, textToFind2) - 1    'here we get 66796, we're looking 4 "textToFind2"
        endPos = InStr(1, rng, textToFind3) - 1      'here we get 42460, we're looking 4 "textToFind3"
        rng.SetRange Start:=startPos, End:=endPos
        Debug.Print "Paragraphs.Count = " & rng.Paragraphs.Count
        rng.MoveEnd wdParagraph, -1
        Debug.Print rng.Text
        arr = Split(Range("A12").Value, " ", -1)    '3rd argument is limit and it's optional. Number of substrings to be returned; -1 indicates that all substrings are returned.
        arrr = Split(Range("A13").Value, " ", -1)   '3rd argument is limit and it's optional. Number of substrings to be returned; -1 indicates that all substrings are returned.
        wrdCount1 = UBound(arr()) + 1
        wrdCount2 = UBound(arrr()) + 1
        
        Debug.Print arr(0) & " " & arr(wrdCount1 - 1) & " oraz " & arrr(0) & " " & arrr(wrdCount2 - 1)
        Debug.Print "First customer's name consists of " & wrdCount1 & " parts."
        Debug.Print "Second customer's name consists of " & wrdCount2 & " parts."
        Debug.Print wrdCount1 & "   " & wrdCount2
        
        Debug.Print "Second customer's name consists of " & wrdCount2 & " parts."
        Debug.Print ", a małżonkowie " & arr(0) & " " & arr(UBound(arr())) & " i " & arrr(0) & " " & arrr(UBound(arrr())) & " oświadczają, że tenże samodzielny lokal niemieszkalny"
        Debug.Print ", a " & arr(0) & " i " & arrr(0) & " małżonkowie " & "<[A-z]{3;25}>" & " oświadczają, że tenże samodzielny lokal mieszkalny"
        
        Debug.Print ", a " & arr(0) & " i " & arrr(0) & " małżonkowie " & "<[A-z]{3;25}>"
        Debug.Print ", a " & arr(0) & " i " & arrr(0) & " małżonkowie " & arr(UBound(arr()))
        Debug.Print ", a " & arr(0) & " i " & arrr(0) & " małżonkowie " & arrr(UBound(arrr()))
        
        With rng.Find
            .Text = ", a " & arr(0) & " i " & arrr(0) & " małżonkowie"
            .MatchWildcards = False
            .MatchCase = False
            .Forward = True
            .Execute
               If .Found = True Then        'If the Find object is accessed from a Range object, the selection is not changed but the Range is redefined when the find criteria is found.
                  mySheet.Range("E12:E13") = mySheet.Range("AE11")
               Else
                  rng.SetRange Start:=startPos, End:=endPos
                  With rng.Find
                     .Text = " a małżonkowie " & arr(0) & " " & arr(UBound(arr())) & " i " & arrr(0) & " " & arrr(UBound(arrr())) & " oświadczają, że tenże samodzielny lokal niemieszkalny"
                     .MatchWildcards = False
                     .MatchCase = False
                     .Forward = True
                     .Execute
                     If .Found = True Then
                        mySheet.Range("E12:E13") = mySheet.Range("AE11")
                     Else
                        rng.SetRange Start:=startPos, End:=endPos
                        With rng.Find
                          .Text = "za podaną cenę zobowiązują się kupić na zasadach wspólności ustawowej majątkowej małżeńskiej."
                          .MatchWildcards = False
                          .MatchCase = False
                          .Forward = True
                          .Execute
                            If .Found = True Then
                               mySheet.Range("F12:F13") = mySheet.Range("AE11")
                            Else
                               rng.SetRange Start:=startPos, End:=endPos
                               With rng.Find
                                 .Text = "w stanie wolnym od w/w obciążeń, za podaną cenę kupią na zasadach wspólności ustawowej majątkowej małżeńskiej."
                                 .MatchWildcards = False
                                 .MatchCase = False
                                 .Forward = True
                                 .Execute
                                   If .Found = True Then
                                      mySheet.Range("F12:F13") = mySheet.Range("AE11")
                                   End If
                                End With
                             End If
                         End With
                     End If
                  End With
              End If
          End With
    End If
    
    'Find and extract contracting parties marital status.
    
    If Application.WorksheetFunction.CountA(mySheet.Range("A12:D15")) = 2 And _
         mySheet.Range("E12") = "????" And _
         mySheet.Range("I12") = "k" Then
            rng.SetRange Start:=startPos, End:=endPos
            Debug.Print "Paragraphs.Count = " & rng.Paragraphs.Count
            With rng.Find
               .Text = arr(0) & " " & arr(wrdCount1 - 1) & " oświadcza ponadto, że jest wdową."
               .Format = False          'True if formatting is included in the find operation. Read/write Boolean.
               .MatchWildcards = False
               .MatchCase = False
               .Forward = True
               .Execute
                If .Found = True Then
                    mySheet.Cells(x, 5) = mySheet.Range("AD17")   'how to insert more actions after IF .Found here;
                Else
                    rng.SetRange Start:=startPos, End:=endPos
                    With rng.Find
                       .Text = "a przedmiotowego nabycia dokona do majątku osobistego za pieniądze pochodzące z jej majątku osobistego,"
                       .Format = False          'True if formatting is included in the find operation. Read/write Boolean.
                       .MatchWildcards = False
                       .MatchCase = False
                       .Forward = True
                       .Execute
                    If .Found = True Then mySheet.Cells(x, 6) = mySheet.Range("AE14")  'how to insert more actions after IF .Found here;
                    End With
                End If
            End With
    ElseIf Application.WorksheetFunction.CountA(mySheet.Range("A12:D15")) = 2 And _
           mySheet.Range("E12") = "????" And _
           mySheet.Range("I12") = "m" Then
               rng.SetRange Start:=startPos, End:=endPos
               Debug.Print "Paragraphs.Count = " & rng.Paragraphs.Count
               With rng.Find
                  .Text = arr(0) & " " & arr(wrdCount1 - 1) & " oświadcza ponadto, że jest wdowcem."
                  .Format = False          'True if formatting is included in the find operation. Read/write Boolean.
                  .MatchWildcards = False
                  .MatchCase = False
                  .Forward = True
                  .Execute
                   If .Found = True Then
                       mySheet.Cells(x, 5) = mySheet.Range("AD18")
                   Else
                       rng.SetRange Start:=startPos, End:=endPos
                       With rng.Find
                           .Text = "a przedmiotowego nabycia dokona do majątku osobistego za pieniądze pochodzące z jego majątku osobistego,"
                           .Format = False          'True if formatting is included in the find operation. Read/write Boolean.
                           .MatchWildcards = False
                           .MatchCase = False
                           .Forward = True
                           .Execute
                       If .Found = True Then mySheet.Cells(x, 6) = mySheet.Range("AE14")  'how to insert more actions after IF .Found here;
                       End With
                   End If
               End With
    ElseIf Application.WorksheetFunction.CountA(mySheet.Range("A12:D15")) = 4 And _
           mySheet.Range("E12") = "małżonko" And _
           Application.WorksheetFunction.CountIf(mySheet.Range("I12:I15"), "k") = 1 And _
           Application.WorksheetFunction.CountIf(mySheet.Range("I12:I15"), "m") = 1 Then
          
           Debug.Print rng.Text
           Set rng = wordApp.ActiveDocument.Content
           Debug.Print rng.Text
           'InStr function returns a Variant (Long) specifying the position of the first occurrence of one string within another.
           startPos = InStr(1, rng, textToFind1) - 1      'here we get 40296, we're looking 4 "textToFind1"
           If startPos < 500 Then startPos = InStr(1, rng, textToFind2) - 1    'here we get 66796, we're looking 4 "textToFind2"
           endPos = InStr(1, rng, textToFind3) - 1      'here we get 42460, we're looking 4 "textToFind3"
           rng.SetRange Start:=startPos, End:=endPos
           Debug.Print "Paragraphs.Count = " & rng.Paragraphs.Count
           rng.MoveEnd wdParagraph, -1
           Debug.Print rng.Text
           arr = Split(Range("A12").Value, " ", -1)    '3rd argument is limit and it's optional. Number of substrings to be returned; -1 indicates that all substrings are returned.
           arrr = Split(Range("A13").Value, " ", -1)   '3rd argument is limit and it's optional. Number of substrings to be returned; -1 indicates that all substrings are returned.
           wrdCount1 = UBound(arr()) + 1
           wrdCount2 = UBound(arrr()) + 1
               With rng.Find        'If the Find object is accessed from a Range object, the selection is not changed but the Range is redefined when the find criteria is found.
                    .Text = textToFind6
                    .Format = False          'True if formatting is included in the find operation. Read/write Boolean.
                    .MatchWildcards = False
                    .MatchCase = False
                    .Forward = True
                    .Execute
                      If .Found = True Then mySheet.Range("F12:F13") = mySheet.Range("AF11")
               End With

    End If
End Sub
