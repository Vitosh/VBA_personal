Option Explicit

Public Sub TestMe()

    Dim var1, var2, var3, var4, var5, var6
    
    var1 = Array(0, 1)
    var2 = Array(10, 11)
    var3 = Array(100, 101)
    var4 = Array(1000, 1001)
    var5 = Array(10000, 10001)
    var6 = Array(100000, 100001)
    
    'ByReference
    var1 = increment01(var1)
    var1 = increment01(var1)
    
    'ByReference
    var2 = increment02(var2)
    var2 = increment02(var2)
    
    'ByReference
    increment01 var3
    increment01 var3
    
    'ByValue
    increment02 var4
    increment02 var4
    
    'ByReference
    increment03 var5
    increment03 var5
    
    'ByValue
    increment04 var6
    increment04 var6
    
    Debug.Print var1(0)
    Debug.Print var2(0)
    Debug.Print var3(0)
    Debug.Print var4(0)
    Debug.Print var5(0)
    Debug.Print var6(0)

End Sub

Public Function increment01(ByRef testValue As Variant) As Variant()

    testValue(0) = testValue(0) + 1
    increment01 = testValue
    
End Function

Public Function increment02(ByVal testValue As Variant) As Variant

    testValue(0) = testValue(0) + 1
    increment02 = testValue
    
End Function

Public Function increment03(ByRef testValue As Variant) As Variant

    testValue(0) = testValue(0) + 1
    increment03 = testValue
    
End Function

Public Function increment04(ByVal testValue As Variant) As Variant()

    testValue(0) = testValue(0) + 1
    increment04 = testValue
    
End Function

 '2 
 '12 
 '102 
 '1000 
 '10002 
 '100000 


