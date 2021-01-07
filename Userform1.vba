Dim result As Double
Dim result_Array() As Double
Dim RArray() As Double '출제된 문제 배열
Dim HArray() As String
Dim EArray() As String
Dim SArray() As String
Dim qArray() As Variant
Dim q2Array_S() As Double
Dim q2Array_M() As Double
Dim Arrsize_E As Integer
Dim flag_c As Double
Dim flag_c2_S As Double
Dim flag_c2_M As Double
Public count As Integer
Public max_Count As Integer
Public Time_Next As Integer
Public Time_check As Boolean

Dim count_num As Integer
Dim Time_Value As Integer
Dim Demo_check As Integer
'Option Explicit
Dim blnOk As Boolean

Dim s1

'시간계산 프로시저
Sub TimeDemo()
    On Error Resume Next
    
    
    '변수가 true --- 프로시저종료
    '변수가 false --- 프로시저 진행
    
    If blnOk = True Then Exit Sub
    
    '변수에 시간을 저장한다
    
    s1 = Timer
    
    
    '타이머가 현재시간보다 항상 1 작으면
    '무한 순환한다
    
    Do While Timer < s1 + 1
      DoEvents
    Loop
    
    '셀에 시간을 표시한다
    Time_Value = Time_Value - 1
    Timer_text.Caption = timer_Change(Time_Value)
    
    If Time_Value <= 10 Then
        Timer_text.ForeColor = &HFF&
    End If
    
    
    If Time_Value <= 0 Then
        If Time_check = False Then
            TimeOut
        Else
            CommandButton1_Click
        End If
    End If
    
    
    '내가 나를 호출한다(재귀 프로시저)
    '언제까지?
    '사용자가 중지시킬때까지.. 무한으로 순환한다
    TimeDemo
    

End Sub


'시트에 버튼을 만들고 아래의 프러시져와 연결
'타이머를 작동하기도하고
'타이머를 중지하기도한다

Sub DoStop()
    
    '처음에는 false가되고
    '다시클릭하면 true가된다
    blnOk = blnOk
    'blnOk = Not blnOk
    
    
    '타이머를 호출한다
    Time_Value = Time_Next
    Timer_text.ForeColor = &H80000012
    Timer_text.Caption = timer_Change(Time_Value)
    
    TimeDemo
    
    
End Sub

Function timer_Change(ByVal time_V As Integer) As String
    Dim min As Integer, sec As Integer: sec = time_V
    Dim col As String: col = ":"
    min = 0
    
    Do While sec >= 60
        min = min + 1
        sec = sec - 60
    Loop
    
    If sec < 10 Then
        col = ":0"
    End If
    
    timer_Change = min & col & sec
End Function

Private Sub Label1_Click()

End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label12_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub question_Click()

End Sub

Private Sub Timer_text_Click()

End Sub

Private Sub UserForm_Initialize()
    Read
    count = 1
End Sub
Private Sub UserForm_Layout()
    question.Caption = "1번째 문제" + HArray(count - 1, 0)
    question_type.Caption = " " + HArray(count - 1, 1)
    Timer_text.Caption = timer_Change(Time_Next)
End Sub
Private Sub TimeOut()
        blnOk = True
        UserForm2.max_Count = max_Count
        UserForm2.question = HArray(count - 1, 0)
        UserForm2.question_type = HArray(count - 1, 1)
        Call UserForm2.data_Input(RArray, result_Array)
        Unload Me
        UserForm2.Show
End Sub
Private Sub CommandButton1_Click()
    Application.EnableEvents = False
    'MsgBox "정답은 : " + Str(result) + Chr(13) & Chr(10) + "입력한 답 : " + TextBox1.Text
    If TextBox1.Text = "" Then
        result_Array(count_num - 2, 0) = 0
    Else
        result_Array(count_num - 2, 0) = TextBox1.Text
    End If
    TextBox1.Text = ""
    If (max_Count >= count_num) Then
        Question_test
        If Not Time_check = False Then
        '타이머 시작함수
        Time_Value = Time_Next
        Timer_text.ForeColor = &H80000012
        Timer_text.Caption = timer_Change(Time_Value)
        s1 = Timer
        End If
        
    Else '모든문제를 다풀면 결과창을 띄워줌
        blnOk = True
        UserForm2.max_Count = max_Count
        UserForm2.question = HArray(count - 1, 0)
        UserForm2.question_type = HArray(count - 1, 1)
        UserForm2.Timer_end = Time_Value
        UserForm2.Timer_check = Time_check
        Call UserForm2.data_Input(RArray, result_Array)
        Unload Me
        UserForm2.Show
    End If
    Application.EnableEvents = True
End Sub


Private Sub Start_btn_Click()
    Start_btn.Enabled = False
    Start_btn.Visible = False
    Setup
    Question_test
    blnOk = False
    Call DoStop
End Sub

Private Sub Setup()
    Timer_text.Enabled = True
    question.Enabled = True
    question_type.Enabled = True
    TextBox1.Locked = False
    TextBox1.Enabled = True
    CommandButton1.Enabled = True
    ReDim RArray(0 To max_Count - 1, 0 To 9)
    ReDim result_Array(0 To max_Count - 1, 0 To 1)
    count_num = 1
End Sub
Public Sub Read()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Sheets("유형입력").Select
    Dim empty_Check As Boolean
    Dim max_Count As Integer
    Dim i As Integer
    
    i = 2
    empty_Check = False
    max_Count = 0
    
    Do While Not (empty_Check)
        If (Cells(i, 1) = "") Then
            max_Count = i - 2 '현재개수
            empty_Check = True
        End If
        i = i + 1
    Loop
    
    i = 2
    empty_Check = False
    
    ReDim HArray(0 To max_Count, 0 To 1)
    Do While Not (empty_Check)
        If (Cells(i, 1) = "") Then
            max_Count = i - 2 '현재개수
            empty_Check = True
        Else
            HArray(i - 2, 0) = Cells(i, 2)
            HArray(i - 2, 1) = Cells(i, 3)
        End If
        i = i + 1
    Loop
    Sheets("시작화면").Select
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then '엔터키를 눌렀을때 동작
        CommandButton1_Click
        KeyCode = 0
    End If
End Sub
Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    TextBox1.SetFocus
End Sub

Private Sub Timer_Click()

End Sub

Public Sub Question_test()
    Dim term_d(1 To 10) As Double
    Dim digit As Integer: digit = 0
    Dim term As Integer: term = 0
    Dim term_Max As Integer: term_Max = 0
    Dim sign_Arr() As String
    Dim sign_Size As Integer: sign_Size = 0
    Dim sign_Double As Boolean: sign_Double = False
    Dim sign_Max As Integer: sign_Max = 0
    Dim term_Flag As Integer: term_Flag = 1
    Dim sign_Sum() As Integer
    
    
    result = 0
    
    Call question_Logic(count - 1)
    
    question.Caption = Str(count_num) + "번째 문제" + HArray(count - 1, 0)
    
    For j = 0 To Arrsize_E - 1 '행 개수 세기
        term_Max = term_Max + Right(EArray(j), 1)
    Next
    
    sign_Arr = Split(HArray(count - 1, 1), ",")
    sign_Size = UBound(sign_Arr) - LBound(sign_Arr) + 1
    
    If sign_Size > 1 Then
        sign_Double = True
    Else
        sign_Double = False
    End If
    
    If sign_Double = True Then
        ReDim sign_Sum(0 To term_Max - 1)
        If term_Max > 4 And term_Max <= 6 Then
            sign_Max = 2
            For c = 0 To term_Max - 1
            sign_Sum(c) = 1
                If c = 1 Or c = 3 Then
                    sign_Sum(c) = -1
                End If
            Next
        ElseIf term_Max > 6 And term_Max <= 8 Then
            sign_Max = 2
            rand1 = (Int(Rnd * 2) + 2)
            rand2 = (Int(Rnd * (term_Max - rand1 - 3)) + rand1 + 3)
            For c = 0 To term_Max - 1
            sign_Sum(c) = 1
                If c = rand1 Or c = rand2 Then
                    sign_Sum(c) = -1
                End If
            Next
        ElseIf term_Max > 8 And term_Max <= 10 Then
            sign_Max = 3
            rand1 = (Int(Rnd * 2) + 5)
            rand2 = (Int(Rnd * (term_Max - rand1 - 3)) + rand1 + 3)
            For c = 0 To term_Max - 1
            sign_Sum(c) = 1
                If c = 2 Or c = rand1 Or c = rand2 Then
                    sign_Sum(c) = -1
                End If
            Next
        End If
        
        For Z = 0 To Arrsize_E - 1
            digit = LeftB(EArray(Z), 1) '여기가 문제 한 글자밖에 못 읽어오기 때문에
            term = Right(EArray(Z), 1) '문자열 자르기를 이용해 한번 해보자
            
            For a = term_Flag To term + term_Flag - 1
                Userform1.Controls("Label" & a).Visible = True
                rand_Area = (10 ^ (digit - 1))
                If sign_Sum(a - 1) = -1 Then
                    If result <= (10 * rand_Area) - 1 Then
                        term_d(a) = (Int(Rnd * (result - rand_Area)) + 1 * rand_Area) * sign_Sum(a - 1)
                    Else
                        term_d(a) = (Int(Rnd * 9 * rand_Area) + 1 * rand_Area) * sign_Sum(a - 1)
                    End If
                Else
                    term_d(a) = (Int(Rnd * 9 * rand_Area) + 1 * rand_Area) * sign_Sum(a - 1)
                End If
                
                Userform1.Controls("Label" & a).Caption = term_d(a)
                Call Comma(term_d(a), a, 1)
                result = result + term_d(a)
                RArray(count_num - 1, a - 1) = term_d(a)
            Next
            term_Flag = term_Flag + term
        Next
        '여기서 행개수 만큼 While문이나 for을 실행
        'sign은 맥스행 개수에따라 가감산 위치가 정해지고 5 = 2,4 or 2,5 7 = 3,5 or 3,6 or 4,7
        '가감산은 현재 값이 -가 되지 않도록 현재 까지 저장되어있는 값에서 - 10까지의 범위로만 랜덤 생성
        '그리고 만들어지는 행마다 따로 저장하는 전역변수가 필요함
        '한번 고민해보자.
        
    ElseIf sign_Arr(0) = "가산" Then
        digit = LeftB(EArray(0), 1)
        term = Right(EArray(0), 1)
        
        For i = 1 To term
            Userform1.Controls("Label" & i).Visible = True
            term_d(i) = Int(Rnd * 9 * (10 ^ (digit - 1))) + 1 * (10 ^ (digit - 1))
            Userform1.Controls("Label" & i).Caption = term_d(i)
            Call Comma(term_d(i), i, 1)
            result = result + term_d(i)
            RArray(count_num - 1, i - 1) = term_d(i)
        Next
        
    ElseIf sign_Arr(0) = "곱셈" Then
        Label12.Caption = "×"
        Label12.Visible = True
        ReDim digit2(0 To 1) As Integer
        
        digit2(0) = LeftB(EArray(0), 1)
        digit2(1) = Right(EArray(0), 1)
        
        For i = 1 To 2
            Userform1.Controls("Label" & (1 + (12 * (i - 1)))).Visible = True
            term_d(i) = Int(Rnd * 9 * (10 ^ (digit2(i - 1) - 1))) + 1 * (10 ^ (digit2(i - 1) - 1))
            Userform1.Controls("Label" & (1 + (12 * (i - 1)))).Caption = term_d(i)
            Call Comma(term_d(i), i, 2)
            RArray(count_num - 1, i - 1) = term_d(i)
        Next
        result = term_d(1) * term_d(2)
        term_Max = 2
    ElseIf sign_Arr(0) = "나눗셈" Then
        Label12.Caption = "÷"
        Label12.Visible = True
        ReDim digit2(0 To 1) As Integer
        ReDim LRdigit(0 To 1) As Double
        digit2(0) = LeftB(EArray(0), 1)
        digit2(1) = Right(EArray(0), 1)
        
        'if qArray가 초기화 되지 않았을 경우
        If Not isInitialised(qArray) Then
            Call nPrime(digit2(0), digit2(1), 0)
        End If
        Call rnd_Prime(digit2(1), LRdigit)
        
        For i = 1 To 2
            Userform1.Controls("Label" & (1 + (12 * (i - 1)))).Visible = True
            term_d(i) = LRdigit(i - 1)
            Userform1.Controls("Label" & (1 + (12 * (i - 1)))).Caption = term_d(i)
            Call Comma(term_d(i), i, 2)
            RArray(count_num - 1, i - 1) = term_d(i)
        Next
        result = term_d(1) / term_d(2)
        term_Max = 2
    End If
    
    If Not term_Max Then
        For i = term_Max + 1 To 10
        Userform1.Controls("Label" & i).Caption = ""
        RArray(count_num - 1, i - 1) = 0
        Next
    End If
    result_Array(count_num - 1, 1) = result
    count_num = count_num + 1
End Sub
Public Sub question_Logic(ByVal num As Integer)
    Dim TArray() As String

    Dim deLimiter As String
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    deLimiter = ","
    TArray = Split(HArray(num, 0), deLimiter)
    SArray = Split(HArray(num, 1), deLimiter)
    
    Arrsize_T = UBound(TArray()) - LBound(TArray()) + 1
    Arrsize_S = UBound(SArray()) - LBound(SArray()) + 1
    Arrsize_E = Arrsize_T
    
    ReDim EArray(0 To Arrsize_T - 1)
    
    For i = 0 To (Arrsize_T - 1)
        EArray(i) = ExtractNumber(TArray(i))
    Next
    
    'Erace (TArray)
    'Erace (SArray)
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
Private Sub Comma(ByVal arr_term As Double, ByVal q As Integer, ByVal Check As Integer)
    i = 1
    If Check = 1 Then
        str_Cap = Userform1.Controls("Label" & q).Caption
    ElseIf Check = 2 Then
        str_Cap = Userform1.Controls("Label" & (1 + (12 * (q - 1)))).Caption
    End If
    cap_Len = Len(str_Cap)
    fir_Len = Len(str_Cap)
    If arr_term < 0 Then
        cap_Len = cap_Len - 1
    End If
    
    If Check = 1 Then
        Do While cap_Len > 3
            str_Cap = Left(str_Cap, fir_Len - (3 * i)) & "," & Right(str_Cap, (3 * i) + (i - 1))
            cap_Len = cap_Len - 3
            i = i + 1
        Loop
        Userform1.Controls("Label" & q).Caption = str_Cap
    ElseIf Check = 2 Then
        Do While cap_Len > 3
            str_Cap = Left(str_Cap, fir_Len - (3 * i)) & "," & Right(str_Cap, (3 * i) + (i - 1))
            cap_Len = cap_Len - 3
            i = i + 1
        Loop
        Userform1.Controls("Label" & (1 + (12 * (q - 1)))).Caption = str_Cap
    End If
    
End Sub

Function ExtractNumber(Val, Optional iStart, Optional iEnd) As String
 
Dim i As Long
Dim Str As String
Dim match As Variant
 
If IsMissing(iStart) Then iStart = 0
If IsMissing(iEnd) Then iEnd = 32767
If iEnd <= 0 Or Not IsNumeric(iEnd) Then iEnd = 32767
If iStart <= 0 Or Not IsNumeric(iStart) Then iStart = 1
 
If iStart > iEnd Then ExtractNumber = CVErr(errvalue)
If IsObject(Val) Then Str = Val.Value Else Str = Val
 
Str = Mid(Str, iStart, iEnd - iStart + 1)
 
With CreateObject("VBScript.RegExp")
    .Pattern = "\d+"
    .Global = True
    Set match = .Execute(Str)
 
        If match.count > 0 Then
            ExtractNumber = ""
            For i = 0 To match.count - 1
               ExtractNumber = ExtractNumber & match(i)
            Next i
        End If
End With
 
End Function

Sub nPrime(ByVal term As Integer, ByVal term2 As Integer, ByVal isCheck As Integer)

    flag_c = 0
    
    ReDim qArray(0 To (10 ^ (term)) - (10 ^ (term - 1)))
    For j = 10 ^ (term - 1) To 10 ^ (term) - 1
        For h = 2 To Sqr(j)
            If j Mod h = 0 Then
                qArray(flag_c) = j
                flag_c = flag_c + 1
                Exit For
            End If
        Next
    Next
        
    'ReDim qArray(flag_c + 1)
    'nPrime = qArray(Int(Rnd * flag_c))
End Sub

Sub rnd_Prime(ByVal term As Integer, ByRef LRdigit1 As Variant)
    Dim prime_check As Boolean: prime_check = False
    Dim mid_Digit As Double
    Dim nest_Check As Boolean: nest_Check = False
    
    Do While Not prime_check
        flag_c2_S = 0
        flag_c2_M = 0
        nest_Check = False
        
        Do While Not nest_Check
        mid_Digit = qArray(Int(Rnd * (flag_c)))
        For i = 0 To count_num - 2
            If RArray(i, 0) = mid_Digit Then
                nest_Check = True
                Exit For
            End If
        Next
        
        If nest_Check = True Then
            nest_Check = False
        Else
            nest_Check = True
        End If
        
        Loop
        
        ReDim q2Array_S(0 To (10 ^ (term)) - (10 ^ (term - 1)))
        ReDim q2Array_M(0 To (10 ^ (term)) - (10 ^ (term - 1)))
        Rnd_C = Int(Rnd * 2)
        
        For j = 10 ^ (term - 1) To 10 ^ (term) - 1
            If mid_Digit Mod j = 0 Then
                If Not j = 10 ^ (term - 1) Then
                    If j >= (10 ^ (term) - 1) / 2 And Rnd_C = 1 Then
                        q2Array_M(flag_c2_M) = j
                        flag_c2_M = flag_c2_M + 1
                    End If
                    If j < (10 ^ (term) - 1) / 2 And Rnd_C = 0 Then
                        q2Array_S(flag_c2_S) = j
                        flag_c2_S = flag_c2_S + 1
                    End If
                End If
            End If
        Next
        
        If flag_c2_S > 0 Or flag_c2_M > 0 Then
            If Rnd_C = 0 Then
                LRdigit1(1) = q2Array_S(Int(Rnd * (flag_c2_S)))
            Else
                LRdigit1(1) = q2Array_M(Int(Rnd * (flag_c2_M)))
            End If
            LRdigit1(0) = mid_Digit
            
            prime_check = True
        End If
    Loop
End Sub
Function isInitialised(ByRef a() As Variant) As Boolean
isInitialised = False
On Error Resume Next
isInitialised = IsNumeric(UBound(a))
End Function

