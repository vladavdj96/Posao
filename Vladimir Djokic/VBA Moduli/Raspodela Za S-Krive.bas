Attribute VB_Name = "Module1"
Function RASPODELA(num1, num2, range_1 As range)
    'Raspodela za Forecast
    
    subs = num2 - num1
    n = range_1.Count
    
    Dim range_2 As range
    
    Delta = subs / n
    
    sum_1 = 0
    For i = 1 To n
        sum_1 = sum_1 + range_1(i)
    Next i
    
    sum_2 = sum_1 + subs
    
    RASPODELA = sum_2 / sum_1
End Function
