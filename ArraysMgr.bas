Attribute VB_Name = "ArraysMgr"
Public Function ArraySum(arr As Variant, Optional lb As Long = -1, Optional ub As Long = -1) As Double
    ' Calculate the sum of an (1D) array
    If lb = -1 Then
        lb = LBound(arr)
    End If
    If ub = -1 Then
        ub = UBound(arr)
    End If
    Dim sum As Double
    sum = 0
    For i = lb To ub
        sum = sum + arr(i)
    Next i
    ArraySum = sum
End Function

Public Function ArrayAverage(arr As Variant, Optional lb As Long = -1, Optional ub As Long = -1) As Double
    ' Calculate the average of an (1D) array
    If lb = -1 Then
        lb = LBound(arr)
    End If
    If ub = -1 Then
        ub = UBound(arr)
    End If
    ArrayAverage = ArraySum(arr, lb, ub) / (ub - lb + 1)
End Function

Public Function GetArrayColumn(matrix As Variant, colNbr As Long, Optional twoD As Boolean = True) As Variant
    Dim nbrRows As Long
    nbrRows = UBound(matrix, 1)
    Dim arr(1 To nbrRows) As Variant
    n = UBound(arr)
    For i = 1 To n
        arr(i) = matrix(i, colNbr)
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
Public Function ArrayNbrDimensions(arr As Variant) As Long
    Dim i As Long
    Dim ub As Long
    On Error Resume Next
    i = 0
    Do
        i = i + 1
        ub = UBound(arr, i)
    Loop Until Err.Number <> 0
    ArrayNbrDimensions = i - 1
End Function

'------------------------------------------------------------------------------
' Converts 1D to 2D arrays and vice versa
'------------------------------------------------------------------------------
Public Function OneDtoTwoD(arr As Variant) As Variant
    Dim lb As Long
    Dim ub As Long
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
    Dim lb As Long
    Dim ub As Long
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
    Dim arr1d() As Variant
    ReDim arr1d(1 To arraySize)
    For i = 1 To arraySize
        arr1d(i) = elementValue
    Next i
    Create1DArray = arr1d
End Function

'------------------------------------------------------------------------------
Public Function Create2DArray(arraySize As Long, elementValue As Variant) As Variant
    Create2DArray = OneDtoTwoD(Create1DArray(arraySize, elementValue))
End Function

