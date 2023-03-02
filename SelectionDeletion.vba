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
    'cmtStyleCount = 0
    For Each p In ActiveDocument.Paragraphs
        paraFound = InStr(1, p.Range.Text, txt, vbTextCompare)
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
                    AddCheckBox p.Range.Text, topMargin
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
'        If foundBool Then
'            If p.Previous.Style = "PR2" And p.Style = "CMT" Then
'                Exit For
'            End If
'            If p.Next.Style = "PR2" Or p.Next.Style = "CMT" Then
'                If cmtStyleCount < 2 And p.Next.Style = "CMT" Then
'                    AddCheckBox p.Next.Range.Text, topMargin
'                    topMargin = topMargin + 25
'                    cmtStyleCount = cmtStyleCount + 1
'                ElseIf cmtStyleCount < 2 Then
'                    AddCheckBox p.Next.Range.Text, topMargin
'                    topMargin = topMargin + 25
'                End If
'            Else
'                Exit For
'            End If
'        End If
'        If paraFound Then
'            foundBool = True
'        End If
    Next
End Function
Function AddCheckBox(boxName As String, topMargin As Integer)
    Dim checkBox As Object
    Set checkBox = PickSectionForm.Controls.Add("Forms.Checkbox.1", "Paragraphs", True)
    With checkBox
        .Caption = boxName
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
   
   Do While True
        With Selection.Find
             .Text = txt
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
