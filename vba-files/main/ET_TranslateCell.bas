Attribute VB_Name = "ET_TranslateCell"
'namespace=vba-files\main\

'1 : 그자리에   customButton11
'2 : 오른쪽 셀  customButton12
'3 : 아래 셀    customButton13

Public blGo As Boolean

Public Sub ET_Exe(control As IRibbonControl)

    On Error GoTo CaseErr '230224

    'Range 아니면 종료
    If (TypeName(Selection) <> "Range") Then
        MsgBox ("Range 개체가 아닙니다. 종료합니다.")
        Exit Sub
    End If

    Call Test

    If blGo = False Then
        Exit Sub
    End If

    ' '호출한 콘트롤 ID = ctID
    ' Dim ctID As String
    ' ctID = control.ID

    Application.ScreenUpdating = False

    '번역대상 => 셀 넣을 값으로 재활용
    Dim str As String
    '번역결과
    Dim tmpStr As String
    '타겟셀 (넣을 셀)
    Dim tgtRange As Range

    '순환개시
    For Each rng In Selection

        '1. 번역값 생성
        str = rng.value
        str = repAP(str) '추가. 230222
        tmpStr = GTdo(str)
        
        If (ET_SetInfo.btnDirIndex = 0) Then
            str = str + vbCrLf + tmpStr
        Else
            str = tmpStr
        End If
        
        
        '2. 결과 투입
        If (ET_SetInfo.btnDirIndex = 0) Then
            rng.value = str
            'rng = rng + tmpStr
        ElseIf (ET_SetInfo.btnDirIndex = 1) Then
            rng.Offset(0, 1).value = str
        ElseIf (ET_SetInfo.btnDirIndex = 2) Then
            rng.Offset(1, 0).value = str
        End If

    Next

CaseErr:     '230224
    If Err.Number <> 0 Then
        MsgBox "Error, 번역실패"
    End If

    Application.ScreenUpdating = True

    MsgBox "Done"

End Sub

''''''''''''''''''''''
' 이하는 공용함수부
''''''''''''''''''''''

'셀 갯수 테스트

Public Sub Test()

    blGo = True

    Dim countCells As Long
    countCells = Selection.Cells.Count

    '셀이 100개 초과면
    If (countCells > 100) Then
        If MsgBox("선택한 셀이 100개가 넘습니다. 만약 행렬전체선택을 하면 VBA가 파괴될수도 있습니다. 진행하시겠습니까?", vbYesNo) = vbNo Then
        
            blGo = False
        
        End If
    End If

End Sub

Public Function GTinput(str As String, tgt As Range)

    'GTinput은 str로 받은 값을 tgt에 집어 넣음
    tgt.value = str

End Function

Public Function GTdo(src As String)

    'GTdo로 던지면 영한번역을 반환함
    If (ET_SetInfo.btnLangIndex = 0) Then
        GTdo = GoogleTranslate(src, "en", "ko")
    ElseIf (ET_SetInfo.btnLangIndex = 1) Then
        GTdo = GoogleTranslate(src, "ko", "en")
    End If

End Function


Public Function GoogleTranslate(strInput As String, strFromSourceLanguage As String, strToTargetLanguage As String) As String
    Dim strURL As String
    Dim objHTTP As Object
    Dim objHTML As Object
    Dim objDivs As Object, objDiv As Object
    Dim strTranslated As String

    ' send query to web page
    strURL = "https://translate.google.com/m?hl=" & strFromSourceLanguage & _
        "&sl=" & strFromSourceLanguage & _
        "&tl=" & strToTargetLanguage & _
        "&ie=UTF-8&prev=_m&q=" & WorksheetFunction.EncodeURL(strInput)

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP") 'late binding
    objHTTP.Open "GET", strURL, False
    'objHTTP.setRequestHeader "Accept-Encoding", "gzip;q=1.0", "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ""

    'MsgBox objHTTP.getResponseHeader("Content-Encoding")

    Dim httpresponse() As Byte
    httpresponse = objHTTP.responseBody
    'Mod_Inflate64.Inflate httpresponse
    'MsgBox StrConv(httpresponse, vbUnicode)
    'Debug.Print StrConv(httpresponse, vbUnicode)


    ' create an html document
    Set objHTML = CreateObject("htmlfile")
    With objHTML
        .Open
        .Write objHTTP.responseText
        '#YGE 221001 for trialrun --  Z = objHTTP.responseText
        '#YGE 221001 for trialrun --  Y = InStr(Z, "result-container")
        '#YGE 221001 for trialrun --  If Y > 0 Then
        '#YGE 221001 for trialrun --      Debug.Print Z
        '#YGE 221001 for trialrun --  End If
        
        .Close
    End With
    
    '#YGE 221001 for trialrun --  Range("H1") = objHTTP.responsetext
    
    '#YGE 221001 for trialrun --  Set objDivsBody = objHTML.getElementsByTagName("body")
    Set objDivs = objHTML.getElementsByTagName("div")
    '#YGE 221001 for trialrun --  Set objDivs2 = objDivsBody(0).getElementsByTagName("div")
    '#YGE 221001 for trialrun --  Set objSpans = objHTML.getElementsByTagName("span")
    '#YGE 221001 for trialrun --  Set objSpans2 = objDivsBody(0).getElementsByTagName("span")
    
    
    '#YGE 221001 for trialrun --  Set objDivs2 = objHTML.getElementsByClassName("Q4iAWc")
    '#YGE 221001 for trialrun --  Set objDivs2 = objHTML.getElementsByClassName("JLqJ4b ChMk0b")
    
    '#YGE 221001 for trialrun --  For Each objDiv In objDivsBody
        '#YGE 221001 for trialrun --  Z = objDiv.className
        '#YGE 221001 for trialrun --  Debug.Print Z
    '#YGE 221001 for trialrun --  Next objDiv
    
    For Each objDiv In objDivs
        '#YGE 221001 for trialrun --  Z = objDiv.className
        '#YGE 221001 for trialrun --  Debug.Print Z
        
        GoogleTranslate = GoogleTranslateRecursion(objDiv.ChildNodes)
        '#YGE 221001 for trialrun --  Debug.Print GoogleTranslate
        If GoogleTranslate <> "" Then
            Exit For
        End If
        If objDiv.className = "result-container" Then
            strTranslated = objDiv.innerText
            GoogleTranslate = strTranslated
            Exit For
        End If
        
    Next objDiv


    Set objHTML = Nothing
    Set objHTTP = Nothing

End Function


Public Function GoogleTranslateRecursion(pobjDivs As Object) As String
    Dim objDivs As Object, objDiv As Object
    Dim strTranslated As String
    GoogleTranslateRecursion = ""
    Set objDivs = pobjDivs
    For Each objDiv In objDivs
        If objDiv.nodeName = "DIV" Then
            '#YGE 221001 for trialrun --  Z = objDiv.className
            '#YGE 221001 for trialrun --  Debug.Print Z
            strTranslated = GoogleTranslateRecursion(objDiv.ChildNodes)
            '#YGE 221001 for trialrun --  Debug.Print strTranslated
            If strTranslated <> "" Then
                GoogleTranslateRecursion = strTranslated
                Exit For
            End If
            
            If objDiv.className = "result-container" Then
                strTranslated = objDiv.innerText
                GoogleTranslateRecursion = strTranslated
                Exit For
            End If
        End If
        
    Next objDiv
End Function

Public Function repAP(str As String) As String

    '&는 And로
    '%는 Percent로

    Dim toStr As String
    toStr = Replace(str, "&", "and")
    toStr = Replace(toStr, "%", "percent")
    repAP = toStr

End Function
