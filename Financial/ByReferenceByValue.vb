Option Explicit

Public Sub TestMe()

    Dim var1, var2
    Dim var3, var4
    Dim var5, var6
    
    var1 = Array(1, 1)
    var2 = Array(2, 1)
    var3 = Array(3, 1)
    var4 = Array(4, 1)
    var5 = Array(5, 1)
    var6 = Array(6, 1)
    
    increment1 (var1)
    increment2 (var2)
    increment1 var3
    increment2 var4
    var5 = increment1(var5)
    var6 = increment2(var6)
    
    Debug.Print var1(0)
    Debug.Print var2(0)
    Debug.Print var3(0)
    Debug.Print var4(0)
    Debug.Print var5(0)
    Debug.Print var6(0)
    
End Sub

Public Function increment1(ByVal testValue As Variant) As Variant
    testValue(0) = testValue(0) + 100
    increment1 = testValue
End Function

Public Function increment2(ByRef testValue As Variant) As Variant
    testValue(0) = testValue(0) + 100
    increment2 = testValue
End Function

'Immediate Window
' 1 
' 2 
' 3 
' 104 
' 105 
' 106 
