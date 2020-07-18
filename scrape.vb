Sub Scrape()

    Dim Driver As New Selenium.ChromeDriver
    Dim count0 As Long
    Dim count1 As Long
    count0 = 1
    count1 = 1
    'Both counts will be used for essentially the same purpose.

    Dim s As String 's will have the entire page of text

    Dim phrase As String
    phrase = "Dist:"
    Dim occurrences As Integer
    occurrences = 0
    Dim intCursor As Integer
    intCursor = 0

    'phrase, occurrences, and intCursor are used in the function counting the number of dists.
    'I just realized I actually did that twice, could probably cut the runtime in half by changing that
    
    Dim alength, blength, clength, alphaangle, betaangle, gammaangle, myMetric, metricxpath, spheresort, spheresortxpath As String
    alength = "49.23"
    blength = "67.39"
    clength = "65.99"
    alphaangle = "90"
    betaangle = "101.8"
    gammaangle = "90"
    myMetric = "NCDist" 'In quotes, put S6, L1, L2, NCDist, V7, or D7
    spheresort = "d" 'In quotes, put f or family or d or distance
    Sheets("Protein1").Activate 'In quotes, put the name of the specific sheet


    Set Driver = CreateObject("Selenium.ChromeDriver")
    
    Driver.Get "http://iterate.sourceforge.net/sauc-1.1.1/"
    
    'Metric Selector, defaults to S6
    If StrComp(myMetric, "L1") = 0 Then
        metricxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[1]/table/tbody/tr[4]/td/input[1]"
        'Debug.Print metricxpath
    ElseIf StrComp(myMetric, "L2") = 0 Then
        metricxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[1]/table/tbody/tr[4]/td/input[2]"
        'Debug.Print metricxpath
    ElseIf StrComp(myMetric, "NCDist") = 0 Then
        metricxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[1]/table/tbody/tr[4]/td/input[3]"
        'Debug.Print metricxpath
    ElseIf StrComp(myMetric, "V7") = 0 Then
        metricxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[1]/table/tbody/tr[4]/td/input[4]"
        'Debug.Print metricxpath
    ElseIf StrComp(myMetric, "D7") = 0 Then
        metricxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[1]/table/tbody/tr[4]/td/input[5]"
        'Debug.Print metricxpath
    Else
        metricxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[1]/table/tbody/tr[4]/td/input[6]"
        'Debug.Print metricxpath
    End If
    Driver.Wait 500
    Driver.FindElementByXPath(metricxpath).Click
    
    'Sphere Sort Chooser, defaults to family
    If StrComp(spheresort, "d") = 0 Or StrComp(spheresort, "distance") = 0 Then
        spheresortxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[10]/td/input[2]"
    Else
        spheresortxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[10]/td/input[1]"
    End If
    Driver.Wait 500
    Driver.FindElementByXPath(spheresortxpath).Click
    
    'A
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/input").Click
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/input").Clear
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/input").SendKeys alength
    
    'B
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[3]/td[2]/input").Click
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[3]/td[2]/input").Clear
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[3]/td[2]/input").SendKeys blength
    
    'C
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]/input").Click
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]/input").Clear
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[4]/td[2]/input").SendKeys clength
    
    'Alpha
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[2]/td[4]/input").Click
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[2]/td[4]/input").Clear
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[2]/td[4]/input").SendKeys alphaangle
    
    'Beta
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[3]/td[4]/input").Click
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[3]/td[4]/input").Clear
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[3]/td[4]/input").SendKeys betaangle
    
    'Gamma
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[4]/td[4]/input").Click
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[4]/td[4]/input").Clear
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[4]/td[4]/input").SendKeys gammaangle
    
    
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
    
    Dim r0 As Match
    Dim mcolResults0 As MatchCollection
    Dim regexZero As String
    regexZero = "\s\w{4}\sDist:"
    Set mcolResults0 = RegEx(s, regexZero, True, , True)
    If Not mcolResults0 Is Nothing Then
        For Each r0 In mcolResults0
            Dim s0 As String
            s0 = Replace(r0, " Dist:", "", 1, 1)
            Range("B" & count0) = s0
            count0 = count0 + 1
        Next r0
    End If
    
    Dim r1 As Match
    Dim mcolResults As MatchCollection
    Dim regexOne As String
    regexOne = "Dist:\s([^\s]+)"
    Set mcolResults = RegEx(s, regexOne, True, , True)
    If Not mcolResults Is Nothing Then
        For Each r1 In mcolResults
            Dim s1 As String
            s1 = Replace(r1, "Dist: ", "", 1, 1) 'For some reason this turns "-0" into "0". Not really an issue though
            Range("C" & count1) = s1
            count1 = count1 + 1
        Next r1
    End If
    
    
    Driver.Quit
End Sub
