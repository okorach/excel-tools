Attribute VB_Name = "ArraysMgr"
Public Function ArraySum(oArray As Variant, Optional lb As Long = -1, Optional ub As Long = -1) As Double
    ' Calculate the sum of an (1D) array
    If lb = -1 Then
        lb = LBound(oArray)
    End If
    If ub = -1 Then
        ub = UBound(oArray)
    End If
    Dim sum As Double
    ArraySum = 0
    For i = lb To ub
        ArraySum = ArraySum + oArray(i)
    Next i
End Function

Public Function ArrayAverage(oArray As Variant, Optional lb As Long = -1, Optional ub As Long = -1) As Double
    ' Calculate the average of an (1D) array
    If lb = -1 Then
        lb = LBound(oArray)
    End If
    If ub = -1 Then
        ub = UBound(oArray)
    End If
    ArrayAverage = ArraySum(oArray, lb, ub) / (ub - lb + 1)
End Function

Public Function GetArrayColumn(oArray As Variant, colNbr As Long, Optional twoD As Boolean = True) As Variant
    Dim nbrRows As Long, i As Long
    nbrRows = UBound(oArray, 1)
    ReDim arr(1 To nbrRows) As Variant
    For i = 1 To nbrRows
        arr(i) = oArray(i, colNbr)
    Next i
    If (twoD) Then
        GetArrayColumn = OneDtoTwoD(arr)
    Else
        GetArrayColumn = arr
    End If
End Function

'------------------------------------------------------------------------------
' Returns nbr of dimensions of array
'------------------------------------------------------------------------------
Public Function ArrayNbrDimensions(oArray As Variant) As Long
    Dim i As Long, ub As Long
    On Error Resume Next
    i = 0
    Do
        i = i + 1
        ub = UBound(oArray, i)
    Loop Until Err.Number <> 0
    ArrayNbrDimensions = i - 1
End Function

'------------------------------------------------------------------------------
' Converts 1D to 2D arrays and vice versa
'------------------------------------------------------------------------------
Public Function OneDtoTwoD(arr As Variant) As Variant
    Dim lb As Long, ub As Long, i As Long
    lb = LBound(arr)
    ub = UBound(arr)
    Dim arr2d() As Variant
    ReDim arr2d(lb To ub, 1 To 1)
    For i = lb To ub
        arr2d(i, 1) = arr(i)
    Next i
    OneDtoTwoD = arr2d
End Function
'------------------------------------------------------------------------------
Public Function TwoDtoOneD(arr2d As Variant) As Variant
    Dim lb As Long, ub As Long, i As Long
    lb = LBound(arr2d, 1)
    ub = UBound(arr2d, 1)
    Dim arr1d() As Variant
    ReDim arr1d(lb To ub)
    For i = lb To ub
        arr1d(i) = arr2d(i, 1)
    Next i
    TwoDtoOneD = arr1d
End Function
'------------------------------------------------------------------------------
Public Function Create1DArray(arraySize As Long, elementValue As Variant) As Variant
    ReDim tmpArr(1 To arraySize) As Variant
    For i = 1 To arraySize
        tmpArr(i) = elementValue
    Next i
    Create1DArray = tmpArr
End Function
'------------------------------------------------------------------------------
Public Function Create2DArray(arraySize As Long, elementValue As Variant) As Variant
    ReDim tmpArr(1 To arraySize, 1 To 1) As Variant
    For i = 1 To arraySize
        tmpArr(i, 1) = elementValue
    Next i
    Create2DArray = tmpArr
End Function

