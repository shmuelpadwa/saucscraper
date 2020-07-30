Sub Scrape()

    Dim Driver As New Selenium.ChromeDriver
    
    
    Dim FirstDriver As New Selenium.ChromeDriver
    Set FirstDriver = CreateObject("Selenium.ChromeDriver")
    Dim protein, url As String
    
    'User-editable stuff begins here
    Dim spacegroupfinder As Boolean
    Dim myMetric, metricxpath, spheresort, spheresortxpath, centering, centeringxpath, proteincolumn, distcolumn, spacegroupcolumn, pora, poraxpath, maxradius, hitlimit As String
    protein = "6QNJ" 'MOST IMPORTANT LINE! EDIT THIS!
    myMetric = "S6" 'In quotes, put S6, L1, L2, NCDist, V7, or D7
    spheresort = "f" 'In quotes, put f or family or d or distance
    pora = "a" 'Percent or Angstroms. Put p or a or P or A
    maxradius = "5" 'Max radius of the sphere, whether percent or Angstroms
    hitlimit = "15" 'Maximum number of results
    spacegroupfinder = True ' Set to true if you also want to find spacegroups for nearby proteins. Be aware of column usage
    proteincolumn = "H" 'Put column of excel sheet you want to have protein names. Shift over by two to run a different metric on the same numbers.
    distcolumn = "I" 'Put column of excel sheet you want to have distances. Shift over by two to run a different metric on the same numbers.
    spacegroupcolumn = "J"
    'end of user-editable stuff
    
    url = "https://www.rcsb.org/structure/" + protein
    FirstDriver.Get url
    
    Dim alength, blength, clength, alphaangle, betaangle, gammaangle, originspacegroup, originspacegroupsend As String
    alength = FirstDriver.FindElementByXPath("//*[@id=""unitCellTable""]/tbody/tr[1]/td[1]").TextAsNumber
    blength = FirstDriver.FindElementByXPath("//*[@id=""unitCellTable""]/tbody/tr[2]/td[1]").TextAsNumber
    clength = FirstDriver.FindElementByXPath("//*[@id=""unitCellTable""]/tbody/tr[3]/td[1]").TextAsNumber
    alphaangle = FirstDriver.FindElementByXPath("//*[@id=""unitCellTable""]/tbody/tr[1]/td[2]").TextAsNumber
    betaangle = FirstDriver.FindElementByXPath("//*[@id=""unitCellTable""]/tbody/tr[2]/td[2]").TextAsNumber
    gammaangle = FirstDriver.FindElementByXPath("//*[@id=""unitCellTable""]/tbody/tr[3]/td[2]").TextAsNumber
    originspacegroup = FirstDriver.FindElementById("exp_undefined_xray_spaceGroup").Text 'get spacegroup
    originspacegroupsend = Replace(originspacegroup, "Space Group: ", "", 1, 1)
    centering = Left(originspacegroupsend, 1)
    'centering = "P" 'In quotes, put lattice centering as P, A, B, C, F, I, R, H, V
    Debug.Print (centering)
    
    
    Sheets(protein).Activate
    Range("A" & 1) = "A length"
    Range("B" & 1) = "B length"
    Range("C" & 1) = "C length"
    Range("D" & 1) = "Alpha angle"
    Range("E" & 1) = "Beta angle"
    Range("F" & 1) = "Gamma angle"
    Range("G" & 1) = "Space Group"
    Range("A" & 2) = alength
    Range("B" & 2) = blength
    Range("C" & 2) = clength
    Range("D" & 2) = alphaangle
    Range("E" & 2) = betaangle
    Range("F" & 2) = gammaangle
    Range("G" & 2) = originspacegroupsend
    
    FirstDriver.Quit
    
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
    
    'Lattice Centering Chooser, defaults to P
    If StrComp(centering, "A") = 0 Or StrComp(centering, "a") = 0 Then
        centeringxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[1]/table/tbody/tr[2]/td/select/option[2]"
    ElseIf StrComp(centering, "B") = 0 Or StrComp(centering, "b") = 0 Then
        centeringxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[1]/table/tbody/tr[2]/td/select/option[3]"
    ElseIf StrComp(centering, "C") = 0 Or StrComp(centering, "c") = 0 Then
        centeringxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[1]/table/tbody/tr[2]/td/select/option[4]"
    ElseIf StrComp(centering, "F") = 0 Or StrComp(centering, "f") = 0 Then
        centeringxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[1]/table/tbody/tr[2]/td/select/option[5]"
    ElseIf StrComp(centering, "I") = 0 Or StrComp(centering, "i") = 0 Then
        centeringxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[1]/table/tbody/tr[2]/td/select/option[6]"
    ElseIf StrComp(centering, "R") = 0 Or StrComp(centering, "r") = 0 Then
        centeringxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[1]/table/tbody/tr[2]/td/select/option[7]"
    ElseIf StrComp(centering, "H") = 0 Or StrComp(centering, "h") = 0 Then
        centeringxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[1]/table/tbody/tr[2]/td/select/option[8]"
    ElseIf StrComp(centering, "V") = 0 Or StrComp(centering, "v") = 0 Then
        centeringxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[1]/table/tbody/tr[2]/td/select/option[9]"
    Else
        centeringxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[1]/table/tbody/tr[2]/td/select/option[1]"
    End If
    
    Driver.Wait 500
    Driver.FindElementByXPath(centeringxpath).Click
    
    'Percent or Angstroms, defaults to percent
    If StrComp(pora, "a") = 0 Or StrComp(pora, "A") = 0 Or StrComp(pora, "angstroms") = 0 Or StrComp(pora, "Angstroms") = 0 Then 'I actually can't remember if this is case sensitive
        poraxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[9]/td/input[3]"
    Else
        poraxpath = "/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[9]/td/input[2]"
    End If
    
    Driver.Wait 500
    Driver.FindElementByXPath(poraxpath).Click
    
    'Maximum Radius
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[9]/td/input[1]").Click
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[9]/td/input[1]").Clear
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[9]/td/input[1]").SendKeys maxradius
    
    'Hit limit
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[5]/td/input").Click
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[5]/td/input").Clear
    Driver.FindElementByXPath("/html/body/font/center[3]/p/table/tbody/tr/td[2]/table/tbody/tr[5]/td/input").SendKeys hitlimit
    
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
            Dim s0, t0 As String
            s0 = Replace(r0, " Dist:", "", 1, 1)
            t0 = Replace(s0, " ", "", 1, 1)
            Range(proteincolumn & (count0 + 1)) = t0 'this leaves a space for the top of the column
            'Spacegroups of friends
            If spacegroupfinder = True Then
            
                Dim PDBDriver As New Selenium.ChromeDriver 'create a new chromedriver to avoid annoying system failures
                Set PDBDriver = CreateObject("Selenium.ChromeDriver")
                Dim longstring, spacegroup, spacegroupsend As String
                longstring = "https://www.rcsb.org/structure/" + t0 'go to each individual molecule
                PDBDriver.Get longstring
                spacegroup = PDBDriver.FindElementById("exp_undefined_xray_spaceGroup").Text 'get spacegroup
                spacegroupsend = Replace(spacegroup, "Space Group: ", "", 1, 1)
                Range(spacegroupcolumn & (count0 + 1)) = spacegroupsend 'put in column
                PDBDriver.Quit ' very important to quit the driver. for memory or something
                'Actually maybe it isn't, but better safe than sorry
            End If
            count0 = count0 + 1
        Next r0
    End If
    
    Range(proteincolumn & 1) = myMetric + " Closest proteins"
    Range(distcolumn & 1) = myMetric + " Distances"
    Range(spacegroupcolumn & 1) = "Spacegroups"
    
    Dim r1 As Match
    Dim mcolResults As MatchCollection
    Dim regexOne As String
    regexOne = "Dist:\s([^\s]+)"
    Set mcolResults = RegEx(s, regexOne, True, , True)
    If Not mcolResults Is Nothing Then
        For Each r1 In mcolResults
            Dim s1 As String
            s1 = Replace(r1, "Dist: ", "", 1, 1) 'For some reason this turns "-0" into "0". Not really an issue though
            Range(distcolumn & (count1 + 1)) = s1
            count1 = count1 + 1
        Next r1
    End If
    
    
    
    Driver.Quit
End Sub
