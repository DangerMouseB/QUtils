Attribute VB_Name = "nTesting"
'*************************************************************************************************************************************************************************************************************************************************
'            COPYRIGHT NOTICE
'
' Copyright (C) David Briant 2009 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************

Option Explicit

Sub test1()
    Dim a() As Long, ASA As SAFEARRAY
    ReDim a(1 To 1, 1 To 2, 1 To 3)
    getSafeArrayDetails a, ASA
    Stop
End Sub

Sub test2()
    Dim a(1) As Byte, b(1) As Integer, c(1) As Long, d(1) As Long, e(1) As Single, f(1) As Double, g(1) As Variant, h   As Variant, i(1) As Boolean, j(1) As String
    Dim retVal As HRESULT, vType As Long, isBSTRArray As Boolean, isVariantArray As Boolean, SA As SAFEARRAY, vType2 As Long
    ReDim h(1 To 1)
    Erase a
    If getArrayFeatures(a, vType, isBSTRArray, isVariantArray, vType2).HRESULT <> S_OK Then Stop
    Debug.Print Hex$(vType), isBSTRArray, isVariantArray, Hex$(vType2), Hex$(varType(a))
    If getArrayFeatures(b, vType, isBSTRArray, isVariantArray, vType2).HRESULT <> S_OK Then Stop
    Debug.Print Hex$(vType), isBSTRArray, isVariantArray, Hex$(vType2), Hex$(varType(b))
    If getArrayFeatures(c, vType, isBSTRArray, isVariantArray, vType2).HRESULT <> S_OK Then Stop
    Debug.Print Hex$(vType), isBSTRArray, isVariantArray, Hex$(vType2), Hex$(varType(c))
    If getArrayFeatures(d, vType, isBSTRArray, isVariantArray, vType2).HRESULT <> S_OK Then Stop
    Debug.Print Hex$(vType), isBSTRArray, isVariantArray, Hex$(vType2), Hex$(varType(d))
    If getArrayFeatures(e, vType, isBSTRArray, isVariantArray, vType2).HRESULT <> S_OK Then Stop
    Debug.Print Hex$(vType), isBSTRArray, isVariantArray, Hex$(vType2), Hex$(varType(e))
    If getArrayFeatures(f, vType, isBSTRArray, isVariantArray, vType2).HRESULT <> S_OK Then Stop
    Debug.Print Hex$(vType), isBSTRArray, isVariantArray, Hex$(vType2), Hex$(varType(f))
    If getArrayFeatures(i, vType, isBSTRArray, isVariantArray, vType2).HRESULT <> S_OK Then Stop
    Debug.Print Hex$(vType), isBSTRArray, isVariantArray, Hex$(vType2), Hex$(varType(i))
    If getArrayFeatures(j, vType, isBSTRArray, isVariantArray, vType2).HRESULT <> S_OK Then Stop
    Debug.Print Hex$(vType), isBSTRArray, isVariantArray, Hex$(vType2), Hex$(varType(j))
    
    If getArrayFeatures(g, vType, isBSTRArray, isVariantArray, vType2).HRESULT <> S_OK Then Stop
    Debug.Print Hex$(vType), isBSTRArray, isVariantArray, Hex$(vType2), Hex$(varType(g))
    If getArrayFeatures(h, vType, isBSTRArray, isVariantArray, vType2).HRESULT <> S_OK Then Stop
    Debug.Print Hex$(vType), isBSTRArray, isVariantArray, Hex$(vType2), Hex$(varType(h))
    h(1) = a
    If getArrayFeatures(h, vType, isBSTRArray, isVariantArray, vType2).HRESULT <> S_OK Then Stop
    Debug.Print Hex$(vType), isBSTRArray, isVariantArray, Hex$(vType2), Hex$(varType(h))
    
End Sub

Function getArrayFeatures(anArray As Variant, oVType As Long, oIsBSTRArray As Boolean, oIsVariantArray As Boolean, oVType2 As Long) As HRESULT
    Dim ptr As Long, fFeatures As Integer
    ptr = getSafeArrayPointer(anArray)
    If ptr = 0 Then getArrayFeatures = CHRESULT(E_INVALIDARG): Exit Function
    apiCopyMemory fFeatures, ByVal ptr + 2, 2
    If fFeatures And FADF_HAVEVARTYPE Then apiCopyMemory oVType, ByVal ptr - 4, 4
    oIsBSTRArray = (fFeatures And FADF_BSTR) > 0
    oIsVariantArray = (fFeatures And FADF_VARIANT) > 0
    apiSafeArrayGetVartype ptr, oVType2
End Function


Sub test5(prices2D() As Double)
    Dim dates1DMap() As Double, closes1DMap() As Double, dates1DMapSA As SAFEARRAY, retVal As HRESULT
    
    If uDBCreateDoubleArrayMap(dates1DMap, VarPtr(prices2D(1, 1)), 1, 1, 5, 0, 0, 0, 0).HRESULT <> S_OK Then Stop
    If uDBCreateDoubleArrayMap(closes1DMap, VarPtr(prices2D(1, 5)), 1, 1, 5, 0, 0, 0, 0).HRESULT <> S_OK Then Stop
    
    uDBReleaseArrayMap closes1DMap
    closes1DMap = dates1DMap
    uDBReleaseArrayMap closes1DMap
    
    uDBGetSafeArrayDetails dates1DMap, dates1DMapSA

    ' demonstrate can't destroy the array - at which point I might crash
    retVal = apiSafeArrayDestroy(getSafeArrayPointer(dates1DMap))
    Select Case retVal.HRESULT
        Case E_INVALIDARG
            ' bad ptr so should be ok
        Case S_OK
            Debug.Print "we're toast"
            Stop
        Case DISP_E_ARRAYISLOCKED
            Debug.Print "DISP_E_ARRAYISLOCKED"
            ' we are ok
    End Select
     
End Sub

Sub test6()
    Dim i As Long, j As Long, blar As Double, prices2D() As Double
    blar = 1
    ReDim prices2D(1 To 5, 1 To 6)
    For i = 1 To 5
        For j = 1 To 6
            prices2D(i, j) = blar
            blar = blar + 1
        Next
    Next
    test5 prices2D
End Sub


Sub test8()
    Dim a() As Long, b() As Long, pASA As Long, pBSA As Long, temp As Long, i As Long, ASA As SAFEARRAY
    ReDim a(1 To 5)
    ReDim b(1 To 1)
    
    For i = 1 To 5
        a(i) = i
    Next
    
    pASA = getSafeArrayPointer(a)
    pBSA = getSafeArrayPointer(b)
    Stop
    
    apiCopyMemory temp, ByVal pBSA + 12, 4
    apiCopyMemory ByVal pBSA + 12, ByVal pASA + 12, 4
    apiCopyMemory ByVal pASA + 12, temp, 4
    
    Stop

    ' B now points to A's data so now can safely Redim A?
    Erase a
    ReDim a(1 To 1, 1 To 1)
    pASA = getSafeArrayPointer(a)
    getSafeArrayDetails a, ASA
    
    apiCopyMemory temp, ByVal pBSA + 12, 4
    apiCopyMemory ByVal pBSA + 12, ByVal pASA + 12, 4
    apiCopyMemory ByVal pASA + 12, temp, 4
    Stop

    ' make A (1 to 1, 1 to 5)
    
    apiCopyMemory ByVal pASA + 16, 5, 4
    apiCopyMemory ByVal pASA + 24, 1, 4
    Stop

    ' make A (1 to 5, 1 to 1)
    apiCopyMemory ByVal pASA + 16, 1, 4
    apiCopyMemory ByVal pASA + 24, 5, 4
    Stop
    
    Erase b
End Sub

Sub test9()
    Dim a() As Long, i As Long
    ReDim a(1 To 5) As Long
    For i = 1 To 5
        a(i) = i
    Next
    Stop
    redimPreserve a, 1, 11, 15, 11, 15, 0, 0
End Sub

Sub test10()
    Dim a() As String, varType As Long, ptr As Long, retVal As Long
    ReDim a(1 To 2, 1 To 1)
    a(1, 1) = "hello"
    a(2, 1) = "there"
    ptr = getSafeArrayPointer(a)
    retVal = apiSafeArrayGetVartype(ptr, varType).HRESULT
    redimPreserve a, 1, 1, 2, 0, 0, 0, 0
End Sub
