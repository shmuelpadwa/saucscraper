Sub Scrape()

    Dim Driver As New Selenium.ChromeDriver
    Dim count As Long
    Dim count2 As Long
    count = 1
    count2 = 1
    'Both counts will be used for essentially the same purpose.
    'But count1 will be incremented in a while loop
    'count2 in a for loop

    Dim s As String 's will have the entire page of html

    Dim phrase As String
    phrase = "Dist:"
    Dim occurrences As Integer
    occurrences = 0
    Dim intCursor As Integer
    intCursor = 0

    'phrase, occurrences, and intCursor are used in the function counting the number of dists.
    'I just realized I actually did that twice, could probably cut the runtime in half by changing that
    
    Sheets("Protein1").Activate


    Set Driver = CreateObject("Selenium.ChromeDriver")
    
    Driver.Get "http://iterate.sourceforge.net/sauc-1.1.1/"
    
    'A
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/input").Click
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/input").Clear
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/input").SendKeys "42.018"
    
    'B
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[3]/td[2]/input").Click
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[3]/td[2]/input").Clear
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[3]/td[2]/input").SendKeys "81.033"
    
    'C
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]/input").Click
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]/input").Clear
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]/input").SendKeys "110.507"
    
    'Alpha
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[2]/td[4]/input").Click
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[2]/td[4]/input").Clear
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[2]/td[4]/input").SendKeys "90"
    
    'Beta
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[3]/td[4]/input").Click
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[3]/td[4]/input").Clear
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[3]/td[4]/input").SendKeys "90"
    
    'Gamma
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[4]/td[4]/input").Click
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[4]/td[4]/input").Clear
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[4]/td[4]/input").SendKeys "90"
    
    
    Driver.FindElementByXPath("/html/body/font/center[3]/p/input[1]").Click
    Driver.Wait 500
    
    s = Driver.FindElementByTag("body").Text
    'Debug.Print s
    Do Until intCursor >= Len(s)

        Dim strCheckThisString As String
        strCheckThisString = Mid(LCase(s), intCursor + 1, (Len(s) - intCursor))

        Dim intPlaceOfPhrase As Integer
        intPlaceOfPhrase = InStr(strCheckThisString, phrase)
        If intPlaceOfPhrase > 0 Then

            occurrences = occurrences + 1
            intCursor = intCursor + (intPlaceOfPhrase + Len(phrase) - 1)

        Else

            intCursor = Len(s)

        End If
        
    Loop
        
    While (count < occurrences)
        Range("B" & count) = Driver.FindElementByXPath("/html/body/font/pre/font/b[" + Str(count) + "]/a").Text
        
        count = count + 1
    

    Wend
    
    
    Dim r As Match
    Dim mcolResults As MatchCollection
    Dim regexOne As String
    regexOne = "Dist:\s([^\s]+)"
    Set mcolResults = RegEx(s, regexOne, True, , True)
    If Not mcolResults Is Nothing Then
        For Each r In mcolResults
            Range("C" & count2) = r
            count2 = count2 + 1
        Next r
    End If
    
    
    Driver.Quit
        
End Sub