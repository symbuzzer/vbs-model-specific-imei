Option Explicit

Do While True
    Dim num, i, digit, even_digits_sum, odd_digits_sum, nearest_multiple, new_num
    Dim version
    Dim title 

    title = "Model-specific IMEI generator"
    version = "v1.0.2"

    Randomize

    Do While True ' 15 basamaklı numara üretene kadar döngü oluştur
        num = "86710505" & Right(CStr(Int(Rnd * 1000000)), 6)

        even_digits_sum = 0
        odd_digits_sum = 0

        For i = 1 To 14
            If Len(num) < i Then Exit For ' check if num has enough characters
            digit = Mid(num, i, 1)
            If i Mod 2 = 0 Then
                even_digits_sum = even_digits_sum + (digit * 2) Mod 10 + Int((digit * 2) / 10)
            Else
                odd_digits_sum = odd_digits_sum + digit
            End If
        Next

        nearest_multiple = Int((odd_digits_sum + even_digits_sum + 9) / 10) * 10
        new_num = num & (nearest_multiple - (odd_digits_sum + even_digits_sum))

        If Len(new_num) = 15 Then 
            Exit Do ' exit loop if new_num is 15 digits long
        Else 
            new_num = "" ' set new_num to empty string to trigger the loop again
        End If
    Loop
    
    new_num = InputBox("Use the OK option to cancel and regenerate this IMEI number, or use the CANCEL option to exit."& vbNewLine & " " & vbNewLine & "Ali BEYAZ (avalibeyaz.com/github)", title & " - " & version, new_num)
    
    If new_num = "" Then Exit Do
Loop