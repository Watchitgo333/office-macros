
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
