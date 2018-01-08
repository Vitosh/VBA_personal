Option Explicit

Public Sub TestMe()

    Dim var1, var2, var3, var4, var5, var6, var7, var8
    
    var1 = Array(0, 1)
    var2 = Array(10, 11)
    var3 = Array(100, 101)
    var4 = Array(1000, 1001)
    var5 = Array(10000, 10001)
    var6 = Array(100000, 100001)
    var7 = Array(1000000, 10000001)
    var8 = Array(10000000, 100000001)
    
    increment01 (var1)
    increment01 (var1)
    
    increment02 (var2)
    increment02 (var2)
    
    increment03 (var3)
    increment03 (var3)
    
    increment04 (var4)
    increment04 (var4)
    
    'ByReference!
    increment01 var5
    increment01 var5

    increment02 var6
    increment02 var6
    
    'ByReference!
    increment03 var7
    increment03 var7
    
    increment04 var8
    increment04 var8
    
    Debug.Print var1(0)
    Debug.Print var2(0)
    Debug.Print var3(0)
    Debug.Print var4(0)
    Debug.Print var5(0)
    Debug.Print var6(0)
    Debug.Print var7(0)
    Debug.Print var8(0)

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

'Result in the immediate window:
 0 
 10 
 100 
 1000 
 10002 
 100000 
 1000002 
 10000000 
