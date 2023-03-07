Sub SelectForm()
    Dim activeDoc As String
    Dim isHvacCondensateDoc As Boolean
    
    activeDoc = ActiveDocument.Name
    
    isHvacCondensateDoc = CheckActiveDoc(activeDoc)
    
    If isHvacCondensateDoc Then
        CopperAndPipeForm
    Else
        NormalForm
    End If

    PickSectionForm.Show vbModeless
End Sub
Function CheckActiveDoc(activeDoc As String) As Boolean
    CheckActiveDoc = InStr(1, activeDoc, "23 21 14 - HVAC Condensate Piping", vbTextCompare)
End Function
Function CopperAndPipeForm()
    Dim txt As String
    Dim para As Word.paragraph
    Dim bool As Boolean
    Dim x As Integer
    
    txt = "ACTION SUBMITTALS"
    
    Set para = FindParagraphs(txt)
End Function
Function FindParagraphs(txt As String) As Word.paragraph
    Dim foundBool As Boolean
    Dim topMargin As Integer
    Dim cmtStyleCount As Integer
    Dim pr1Passed As Boolean
    pr1Passed = False
    topMargin = 30
    foundBool = False
    For Each p In ActiveDocument.Paragraphs
        paraFound = InStr(1, p.Range.text, txt, vbTextCompare)
        If paraFound Then
            foundBool = True
        End If
        If foundBool Then
            If p.Style = "PR1" Then
                pr1Passed = True
            End If
            If pr1Passed Then
                If p.Style = "CMT" And p.Previous.Style = "PR2" Then
                    Exit For
                End If
                If p.Style = "PR2" Or p.Style = "CMT" Then
                    AddCheckBox p.Range.text, topMargin
                    topMargin = topMargin + 25
                End If
                If p.Style = "ART" And p.Previous.Style = "PR2" Then
                    Exit For
                End If
                If p.Style = "PR1" And p.Previous.Style = "PR2" Then
                    Exit For
                End If
            End If
        End If
    Next
End Function
Function AddCheckBox(boxName As String, topMargin As Integer)
    Dim checkBox As Object
    Set checkBox = PickSectionForm.Controls.Add("Forms.Checkbox.1", "Paragraphs", True)
    With checkBox
        .caption = boxName
        .Left = 10
        .Width = 400
        .Top = topMargin
    End With
End Function
Function NormalForm()
    Dim txt As String
    Dim para As Word.paragraph
    Dim bool As Boolean
    Dim x As Integer
    
    txt = "SUMMARY"
    
    Set para = FindParagraphs(txt)
End Function
Sub SectionDeletion(txt As String)

' SectionDeletion Macro

   Debug.Print txt
   Do While True
        With Selection.Find
            .text = txt
            .Wrap = wdFindContinue
            .format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute
        Debug.Print Selection
        If Selection.Find.Found Then
            If Selection.Paragraphs(1).OutlineLevel = "2" Then
                Do While True
                    Selection.Paragraphs(1).Range.Delete
                    If Selection.Paragraphs(1).Next.OutlineLevel = "2" Then
                        Exit Do
                    End If
                Loop
            End If
        Else
            Exit Do
        End If
        Selection.Paragraphs(1).Range.Delete
    Loop
End Sub
Function GatherKeywords(caption As String)
    Debug.Print "Gathering keywords for " & caption
    If InStr(1, caption, "Copper Tube.", vbTextCompare) Then
        Dim copperTubeKeys As Variant
        Dim copperTubePhrase As String
        copperTubeKeys = Array("copper tub", "copper-", "rod size", "dielectric", "lead-free alloy", "copper alloy")
        copperTubePhrase = "Condensate-Drain Piping:  Type DWV, drawn-temper copper tubing, wrought-copper fittings, and soldered joints or"
        DeletePhrase (copperTubePhrase)
        FindKeywordsInDoc (copperTubeKeys)
        ReplaceItems ("^p^p")
    End If
    If InStr(1, caption, "Plastic pipe and fittings with solvent cement.", vbTextCompare) Then
        Dim plasticPipeKeys As Variant
        Dim plasticPipePhrase As String
        plasticPipeKeys = Array("pvc", "solvent cement", "plastic piping", "primer", "pipe-flange", "plastic pipe and fittings", "scratching")
        plasticPipePhrase = "or Schedule 40 PVC plastic pipe and fittings and solvent-welded joints."
        DeletePhrase (plasticPipePhrase)
        FindKeywordsInDoc (plasticPipeKeys)
        ReplaceItems ("^p^p")
    End If
End Function
Function DeletePhrase(phrase As String)
    With Selection.Find
        .text = phrase
        .Wrap = wdFindContinue
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
     End With
     Selection.Find.Execute
     If Selection.Find.Found Then
        Debug.Print Selection.Range
        Selection.Range.Delete
     End If
End Function
Function FindKeywordsInDoc(keywordArray As Variant)
    For Each keyword In keywordArray
        SelectExecute (keyword)
    Next
End Function
Function ReplaceItems(item As String)
    With Selection.Find
        .text = item
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
End Function
Function SelectExecute(keyword As String)
    Do While True
        With Selection.Find
            .text = keyword
            .Wrap = wdFindContinue
            .format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
         End With
         Selection.Find.Execute
         If Selection.Find.Found Then
            Debug.Print Selection.Paragraphs(1).Range.text
            Selection.Paragraphs(1).Range.Delete
         Else
            Exit Do
         End If
    Loop
End Function
'Function SelectDelete(paragraph As Word.paragraph, keyword As String)
'Debug.Print Selection.Paragraphs(1).Range
'If InStr(1, "Condensate-Drain Piping:  Type DWV, drawn-temper copper tubing, wrought-copper fittings, and soldered joints or", keyword, vbTextCompare) Or InStr(1, "or Schedule 40 PVC plastic pipe and fittings and solvent-welded joints.", keyword, vbTextCompare) Then
'    Selection.Range.Delete
'    Debug.Print keyword
'    Debug.Print True
'Else
'    Selection.Paragraphs(1).Range.Delete
'End If
''    If InStr(1, "or Schedule 40 PVC plastic pipe and fittings and solvent-welded joints.", keyword, vbTextCompare) Then
''        Selection.Range.Delete
''    Else
''        Selection.Paragraphs(1).Range.Delete
''    End If
'End Function
