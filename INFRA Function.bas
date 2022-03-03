Attribute VB_Name = "Module1"
Function NORMALIZATION(ar As Range)

'arr = Range("C4:E6")
arr = ar.Columns 'convert range into array

Dim sum As Variant 'array to hold summation
ReDim sum(1 To UBound(arr, 1))

Dim nur As Variant ' array to hold normalization
ReDim nur(1 To UBound(arr, 1) + 3, 1 To UBound(arr, 1))
'calculate sumation
For i = 1 To UBound(arr, 1)
   cc = Application.WorksheetFunction.Index(arr, 0, i) 'get each column in array
   sum(i) = WorksheetFunction.sum(cc) ' sum each column in the array
Next
'calculate normalization
For i = 1 To UBound(arr, 1)
    cc = Application.WorksheetFunction.Index(arr, 0, i) 'get each column in array
    gg = WorksheetFunction.Transpose(cc)
    For j = 1 To UBound(arr, 1)
        nur(i, j) = gg(j) / sum(i) 'divide each column by its sum Value
    Next
Next
'Calculate W
For i = UBound(nur, 1) - 2 To UBound(nur, 1) - 2
    For j = 1 To UBound(arr, 1)
        w = Application.WorksheetFunction.Index(nur, 0, j)
        nur(i, j) = Application.WorksheetFunction.Average(w)
    Next
Next
'calculate W'
w = Application.WorksheetFunction.Index(nur, UBound(nur, 1) - 2, 0)
ww = Application.WorksheetFunction.Transpose(Application.WorksheetFunction.MMult(arr, Application.WorksheetFunction.Transpose(w)))
For i = UBound(nur, 1) - 1 To UBound(nur, 1) - 1
    For j = 1 To UBound(arr, 1)
        nur(i, j) = ww(j)
     Next
Next
'calculate w''
For i = UBound(nur, 1) To UBound(nur, 1)
    For j = 1 To UBound(arr, 1)
        nur(i, j) = ww(j) / w(j)
     Next
Next

NORMALIZATION = Application.WorksheetFunction.Transpose(nur)
End Function
Function CIRICR(ar As Range)
'arr = Range("N15:N18")
Dim hh As Variant 'array to hold summation
ReDim hh(1 To 4)
arr = ar.Columns
L = UBound(arr, 1)
Lmdmax = Application.WorksheetFunction.Average(arr)
CI = (Lmdmax - L) / (L - 1)

Dim RI As Variant
If L = 1 Then RI = 0
If L = 2 Then RI = 0
If L = 3 Then RI = 0.58
If L = 4 Then RI = 0.9
If L = 5 Then RI = 1.12
If L = 6 Then RI = 1.24
If L = 7 Then RI = 1.32
If L = 8 Then RI = 1.41
If L = 9 Then RI = 1.45
If L = 10 Then RI = 1.49

CR = CI / RI
If CR < 0.1 Then gg = "consistency is acceptable"
If CR > 0.1 Then gg = "consistency is not acceptable"
hh(1) = CI
hh(2) = RI
hh(3) = CR
hh(4) = gg
CIRICR = Application.WorksheetFunction.Transpose(hh)
End Function

Function NthRoot(ar As Range)
'arr = Range("C4:E6")
arr = ar.Columns
Dim nth As Variant ' array to hold normalization
ReDim nth(1 To 5, 1 To UBound(arr, 1))
'get multiblication of the row
For i = 1 To UBound(arr, 1)
    g = 1
    For j = 1 To UBound(arr, 1)
        g = g * arr(i, j)
    Next
    nth(1, i) = g
Next
'get nthroot of each row
For i = 1 To UBound(arr, 1)
    nth(2, i) = nth(1, i) ^ (1 / UBound(arr, 1))
Next
'sum of all rows and get w
sum = Application.WorksheetFunction.sum(Application.WorksheetFunction.Index(nth, 2, 0))
For i = 1 To UBound(arr, 1)
    nth(3, i) = nth(2, i) / sum
Next
'get w'
w = Application.WorksheetFunction.Index(nth, 3, 0)
ww = Application.WorksheetFunction.Transpose(Application.WorksheetFunction.MMult(arr, Application.WorksheetFunction.Transpose(w)))
For i = 1 To UBound(arr, 1)
    nth(4, i) = ww(i)
Next
'get w''
For i = 1 To UBound(arr, 1)
    nth(5, i) = ww(i) / w(i)
Next
NthRoot = Application.WorksheetFunction.Transpose(nth)
End Function
Function SumOfYearsDigits(Cost As Range, Salvage As Range, lifetime As Range, endyear As Range)

p = Cost
F = Salvage
N = lifetime
TD = p - F
nn = endyear
Dim arr As Variant
ReDim arr(1 To nn + 2, 1 To 4)
arr(1, 1) = "Years"
arr(1, 2) = "Remaining life / sum of-years"
arr(1, 3) = "Annual depreciation"
arr(1, 4) = "Book value"

arr(2, 1) = 0
arr(2, 2) = 0
arr(2, 3) = 0
arr(2, 4) = p

BV = p
For i = 1 To nn
    RS = (N - i + 1) / ((N * (N + 1)) / 2)
    DN = TD * RS
    BV = BV - DN
    arr(2 + i, 1) = i
    arr(2 + i, 2) = RS
    arr(2 + i, 3) = DN
    arr(2 + i, 4) = BV
Next

SumOfYearsDigits = arr

End Function
Function StraightLine(Cost As Range, Salvage As Range, lifetime As Range, endyear As Range)

p = Cost
F = Salvage
N = lifetime
TD = p - F
nn = endyear
Dim arr As Variant
ReDim arr(1 To nn + 2, 1 To 3)
arr(1, 1) = "Years"
arr(1, 2) = "Annual depreciation"
arr(1, 3) = "Book value"

arr(2, 1) = 0
arr(2, 2) = 0
arr(2, 3) = p

BV = p
For i = 1 To nn
    DN = TD / N
    BV = BV - DN
    arr(2 + i, 1) = i
    arr(2 + i, 2) = DN
    arr(2 + i, 3) = BV
Next

StraightLine = arr

End Function
Function SinkingFund(Cost As Range, Salvage As Range, lifetime As Range, endyear As Range, interest_rate As Range)

p = Cost
F = Salvage
N = lifetime
TD = p - F
nn = endyear
i = interest_rate
Dim arr As Variant
ReDim arr(1 To nn + 2, 1 To 3)
arr(1, 1) = "Years"
arr(1, 2) = "Annual depreciation"
arr(1, 3) = "Book value"

arr(2, 1) = 0
arr(2, 2) = 0
arr(2, 3) = p
c = TD * (i / (((1 + i) ^ N) - 1))
BV = p
For j = 1 To nn
    DN = c * ((1 + i) ^ (j - 1))
    BV = BV - DN
    arr(2 + j, 1) = j
    arr(2 + j, 2) = DN
    arr(2 + j, 3) = BV
Next

SinkingFund = arr

End Function
Function BookValue(AcquisitionCost As Range, Usefullife As Range, endyear As Range)
p = AcquisitionCost
N = Usefullife
nn = endyear
Value = p * ((N - nn) / N)
BookValue = Value
End Function
Function EquivalentPresentWorth(AcquisitionCost As Range, Usefullife As Range, endyear As Range, interest_rate As Range)
p = AcquisitionCost
N = Usefullife
nn = endyear
i = interest_rate
p = AcquisitionCost
Value = p * ((N - nn) / N) * (1 / (1 + i) ^ nn)
EquivalentPresentWorth = Value
End Function
Function ConvertMatrix(Past As Range, Future As Range)
p = Application.WorksheetFunction.Transpose(Past)
F = Application.WorksheetFunction.Transpose(Future)
Dim ar As Variant
ReDim ar(1 To UBound(p, 1))
Dim arr As Variant
ReDim arr(1 To UBound(p, 1))

For i = 1 To UBound(p, 1)
    ar(i) = F(i) / p(i)
    arr(i) = 1 - ar(i)
Next
Dim arrr As Variant
ReDim arrr(1 To UBound(p, 1), 1 To UBound(p, 1))
For i = 1 To UBound(p, 1)
    For j = 1 To UBound(p, 1)
        If j = (UBound(p, 1) - (UBound(p, 1) - i)) Then arrr(i, j) = ar(i) Else If j = (i - 1) Then arrr(i, j) = arr(i - 1) Else: arrr(i, j) = 0
    Next
Next
ConvertMatrix = Application.WorksheetFunction.Transpose(arrr)
End Function
Function Prediction(matrix As Range, curant As Range, year As Range)
m = matrix
c = curant
y = year
Dim p As Variant
ReDim p(1 To y, 1 To UBound(c, 1) + 1)
gg = Application.WorksheetFunction.Transpose(Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(m), c))
mm = gg
For r = 1 To y
    p(r, 1) = r
    For cc = 2 To UBound(c, 1) + 1
        p(r, cc) = mm(cc - 1)
    Next
    mm = Application.WorksheetFunction.Transpose(Application.WorksheetFunction.MMult(Application.WorksheetFunction.Transpose(m), Application.WorksheetFunction.Transpose(mm)))
Next
Prediction = Application.WorksheetFunction.Transpose(p)
End Function
Function TotalCondition(condition As Range, rate As Range)
cc = Application.WorksheetFunction.Transpose(rate)
hh = Application.WorksheetFunction.Transpose(condition)
ss = Application.WorksheetFunction.sum(hh)
Dim arr As Variant
ReDim arr(1 To UBound(hh, 1))
For i = 1 To UBound(hh, 1)
    arr(i) = hh(i) / ss
Next
TotalCondition = Application.WorksheetFunction.SumProduct(cc, arr)
End Function

