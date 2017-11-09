Attribute VB_Name = "Module1"
Public Function fn_YesCount(pRange As Range) As String

    Dim i As Integer
    Dim s_yes, s_all As Integer
    
    i = 1
    s_yes = 0
    s_all = 0
    
    Do While i <= pRange.Rows.Count
     
        If pRange.Cells(i, 1) = "да" Then
            s_yes = s_yes + 1
        End If
        If Trim(pRange.Cells(i, 1)) <> "" Then
            s_all = s_all + 1
        End If
        i = i + 1
    Loop
    fn_YesCount = str(s_yes) + " из " + str(s_all)
End Function

