Option Explicit

Dim num, i, digit, even_digits_sum, odd_digits_sum, nearest_multiple, new_num

Randomize ' seed the random number generator

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

new_num = InputBox("imei:", "imei", new_num)