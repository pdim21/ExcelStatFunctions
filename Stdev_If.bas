Attribute VB_Name = "Module1"
Function NewStDev(Criteria As Range, Values As Range, EqualTo As Double) As Double

Dim Arr1()
Dim Arr2()

Dim CurTotal As Double
Arr1 = Values
Arr2 = Criteria


For x = LBound(Arr1) To UBound(Arr1)
Debug.Print Arr1(x, 1)
If Arr2(x, 1) = EqualTo Then
    CurTotal = CurTotal + Arr1(x, 1)
    y = y + 1
End If
Next

TheAverage = CurTotal / y

For x = LBound(Arr1) To UBound(Arr1)
Debug.Print Arr1(x, 1)
If Arr2(x, 1) = EqualTo Then
    dev = dev + (Arr1(x, 1) - TheAverage) ^ 2
End If
Next
NewStDev = (dev / (y - 1)) ^ 0.5

End Function
