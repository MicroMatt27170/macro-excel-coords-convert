Function ConvertirLatitudSexacimalADecimal(lat)
    Dim parts() As String
    
    parts() = Split(lat, "°")
    Dim num As Double, isNum As Boolean
    num = 0
    ConvertirLatitudSexacimalADecimal = Null
    If UBound(parts) > 0 Then
        isNum = IsNumeric(parts(0))
        If isNum Then
            num = CDbl(parts(0))
        End If
        ConvertirLatitudSexacimalADecimal = num
        parts() = Split(parts(1), "'")
        
        If UBound(parts) > 0 Then
            isNum = IsNumeric(parts(0))
            If isNum Then
                num = num + CDbl(parts(0)) / 60
            End If
            ConvertirLatitudSexacimalADecimal = num
            parts() = Split(parts(1), Chr(34))
            
            If UBound(parts) > 0 Then
                isNum = IsNumeric(parts(0))
                If isNum Then
                    num = num + CDbl(parts(0)) / 3600
                End If
                If parts(1) = "S" Or parts(1) = "s" Or parts(1) = "W" Or parts(1) = "w" Then
                    num = num * -1
                End If
                ConvertirLatitudSexacimalADecimal = num
            End If
        End If
        
    End If
    
End Function
