# Description
A VBA macro for MS Excel that counts patterns of correct response (C) and errors (E) across multiple trials. It was initially developed as a tool to help with data analysis of multi-trial memory experiments in which one or more subjects provide boolean-type responses (yes/now, correct/error) for multiple items across multiple tests. In such designs, each item will generate a pattern of C-E responses across tests. This VBA macro counts all such patterns across subjects and items. For more information, see related articles.


# Authorship
Jeferson Mayer and Carlos Gomes


# Related articles
[Gomes, C. F. A., Brainerd, C. J., Nakamura, K., & Reyna, V. F. (2014). Markovian interpretations of dual retrieval processes. Journal of Mathematical Psychology, 59, 50-64.](https://www.sciencedirect.com/science/article/abs/pii/S0022249613000618)

[Gomes, C. F., Brainerd, C. J., & Stein, L. M. (2013). Effects of emotional valence and arousal on recollective and nonrecollective recall. Journal of Experimental Psychology: Learning, Memory, and Cognition, 39, 663-677.](https://psycnet.apa.org/record/2012-13130-001)


# Usage
The VBA macro is part of the data-count.xls MS Excel file.  You need to enable the use of macros (they're disabled by default, for security reasons). The file has two spreadsheets: DATA, results.  The DATA spreadsheet contains the data from your multi-trial experiment, which should have the following structure:

1. The first row always contains NAMES OF VARIABLES.
2. The DATA RANGE begins at the SECOND row.
3. The DATA RANGE is CONTIGUOUS, that is, there are neither empty rows nor empty columns before or within the data range. 
4. The first column always contains an ID NUMBER for one or more subjects.
5. Names for each STIMULI used in the experiment must have the following format: [one or more uppercase letters][lowercase letter 't'][one or more digits]. For example: ARANHt1, meaning the word ARANHA (spider, in portuguese) on trial 1; then ARANHt2 for trial 2; ARANHt3; ARANHt4; etc. It is fundamental to follow this pattern because the program will search for it. 
6. Only NAMES OF STIMULI can have the format specified on the 5th item.
7. The DATA spreadsheet must be called 'DATA'. 

The results spreadsheet contains instructions and the output of the macro. That is, after adding your data to the DATA spreadsheet, manually run the macro or press the 'Compute Frequencies' button to obtain count data patterns across trials.  The data count will show up in a pop-up window and in the results spreadsheet.


# VB code
``` dif
Sub GetDataRange(Sheet As Worksheet, ByRef MinRow As Integer, ByRef MaxRow As Integer, ByRef MinCol As Integer, ByRef MaxCol As Integer)
    On Error Resume Next
    ''' find mininum column (MinCol)
    Col = 1
    Found = False
    Content = Sheet.Cells(1, Col)
    While Not Found And Not IsEmpty(Content)
        Dim Exp As New RegExp
        With Exp
            .Pattern = "([A-Z]+)|t([0-9]+)"
            .IgnoreCase = False
            .Global = True
        End With
        Set Tokens = Exp.Execute(Content)
        If Tokens.Count = 2 Then
            ''' column header matches the regular expression: MinCol found
            Found = True
            MinCol = Col
        End If
        Col = Col + 1
        Content = Sheet.Cells(1, Col)
    Wend
    ''' minimum row is assumed to be always 2
    MinRow = 2
    ''' find maximum row (MaxRow)
    Row = 2
    Found = False
    Content = Sheet.Cells(Row, 1)
    While Not Found
        Exp = RegExp
        Exp.Pattern = "[0-9]+"
        Exp.Global = True
        Set Tokens = Exp.Execute(Content)
        If Tokens.Count <> 1 Then
            Found = True
            MaxRow = Row - 1
        End If
        Row = Row + 1
        Content = Sheet.Cells(Row, 1)
    Wend
    MaxCol = Sheet.Cells.Find(What:="*", SearchDirection:=xlPrevious, SearchOrder:=xlByColumns).Column
End Sub
''' This function will be used to count C-E patterns.
Function GetPatternAsString(ByVal Pattern As Integer, ByVal NumDigits As Integer) As String
    Dim Str As String
    While NumDigits > 0
        If Pattern Mod 2 = 0 Then
            Str = "E" + Str
        Else
            Str = "C" + Str
        End If
        Pattern = Pattern \ 2
        NumDigits = NumDigits - 1
    Wend
    GetPatternAsString = Str
End Function

Sub FreqMacro()
    Dim Word As String
    Dim Trial As Integer
    Dim Limit As Range
    Dim Pattern As Integer
    Dim Patterns() As Integer
    Dim WordMap As New Dictionary

    Dim MinRow As Integer
    Dim MaxRow As Integer
    Dim MinCol As Integer
    Dim MaxCol As Integer

    Call GetDataRange(Sheets("DATA"), MinRow, MaxRow, MinCol, MaxCol)

    TrialCount = 0

    For Row = MinRow To MaxRow
        For Col = MinCol To MaxCol
            Header = Sheets("DATA").Cells(1, Col)
            Content = Sheets("DATA").Cells(Row, Col)
            Dim Exp As New RegExp
            With Exp
                .Pattern = "([A-Z]+)|t([0-9]+)"
                .IgnoreCase = False
                .Global = True
            End With

            Set Tokens = Exp.Execute(Header)
            Word = Tokens(0)
            Trial = Tokens(1).SubMatches(1)

            If Trial > TrialCount Then
                TrialCount = Trial
            End If

            If WordMap.Exists(Word) Then
                Pattern = WordMap(Word)
                Pattern = (Pattern * 2) + Content
                WordMap.Remove Word
                WordMap.Add Word, Pattern
            Else
                WordMap.Add Word, Content
            End If

            If Row = MinRow And Col = MaxCol Then
                MaxPattern = (2 ^ TrialCount) - 1
                ReDim Patterns(0 To MaxPattern)
            End If
        Next Col
        For Each Entry In WordMap
            Pattern = WordMap(Entry)
            Patterns(Pattern) = Patterns(Pattern) + 1
        Next Entry
        WordMap.RemoveAll
    Next Row

    For i = MaxPattern To 0 Step -1
        msg = msg & GetPatternAsString(i, TrialCount) & ": " & Patterns(i) & vbCrLf
        Worksheets("results").Cells(i + 3, 2).Value = GetPatternAsString(i, TrialCount) & ": " & Patterns(i)
    Next i

    MsgBox msg

End Sub
```
