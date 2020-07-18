Function RegEx(strInput As String, strPattern As String, _
    Optional GlobalSearch As Boolean, Optional MultiLine As Boolean, _
    Optional IgnoreCase As Boolean) As MatchCollection
    
    Dim mcolResults As MatchCollection
    Dim objRegEx As New RegExp
    
    If strPattern <> vbNullString Then
        
        With objRegEx
            .Global = GlobalSearch
            .MultiLine = MultiLine
            .IgnoreCase = IgnoreCase
            .Pattern = strPattern
        End With
    
        If objRegEx.Test(strInput) Then
            Set mcolResults = objRegEx.Execute(strInput)
            Set RegEx = mcolResults
        End If
    End If
End Function
