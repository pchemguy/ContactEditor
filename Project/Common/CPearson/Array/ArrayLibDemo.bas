Attribute VB_Name = "ArrayLibDemo"
'@Folder "Common.CPearson.Array"
'@IgnoreModule
'UseMeaningfulName, ImplicitActiveSheetReference
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modDemo
' By Chip Pearson, www.cpearson.com, chip@cpearson.com
' This module contains test procedures for the functions defined in modArraySupport.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Sub DemoCompareArrays()
    Dim Arr1(1 To 3) As String
    Dim Arr2(1 To 3) As String
    Dim ResArr() As Long
    Dim B As Boolean
    Dim N As Long

    Arr1(1) = "2"
    Arr1(2) = "a"
    Arr1(3) = vbNullString

    Arr2(1) = "4"
    Arr2(2) = "c"
    Arr2(3) = "x"

    B = ArrayLib.CompareArrays(Array1:=Arr1, Array2:=Arr2, ResultArray:=ResArr, CompareMode:=vbTextCompare)
    If B = True Then
        For N = LBound(ResArr) To UBound(ResArr)
            Debug.Print CStr(N), Arr1(N), Arr2(N), ResArr(N)
        Next N
    Else
        Debug.Print "CompareArrays returned false."
    End If
End Sub


Public Sub DemoConcatenateArrays()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoConcatenateArrays
    ' This demonstrates the ConcatenateArrays function.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim ResultArray() As Long                    ' MUST be dynamic
    Dim ArrayToAppend(1 To 3) As Long
    Dim N As Long
    Dim B As Boolean
    ReDim ResultArray(1 To 3)
    ResultArray(1) = 8
    ResultArray(2) = 9
    ResultArray(3) = 10

    ArrayToAppend(1) = 111
    ArrayToAppend(2) = 112
    ArrayToAppend(3) = 113

    B = ArrayLib.ConcatenateArrays(ResultArray:=ResultArray, ArrayToAppend:=ArrayToAppend)
    If B = True Then
        If ArrayLib.IsArrayAllocated(Arr:=ResultArray) = True Then
            For N = LBound(ResultArray) To UBound(ResultArray)
                If IsObject(ResultArray(N)) = True Then
                    Debug.Print CStr(N), "is object", TypeName(ResultArray(N))
                Else
                    Debug.Print CStr(N), ResultArray(N)
                End If
            Next N
        Else
            Debug.Print "Result Array Is Not Allocated."
        End If
    Else
        Debug.Print "ConcatenateArrays returned False"
    End If
End Sub


Public Sub DemoCopyArray()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoCopyArray
    ' This demonstrates the CopyArray function.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim Src(1 To 2) As Long
    Dim Dest(1 To 2) As Integer
    Dim Ndx As Long
    Dim B As Boolean

    Src(1) = 1234
    Src(2) = Rows.Count * 10

    B = ArrayLib.CopyArray(DestinationArray:=Dest, SourceArray:=Src, NoCompatabilityCheck:=True)
    If B = True Then
        If ArrayLib.IsArrayAllocated(Arr:=Dest) = True Then
            For Ndx = LBound(Dest) To UBound(Dest)
                If IsObject(Dest(Ndx)) = True Then
                    Debug.Print CStr(Ndx), "is object"
                Else
                    Debug.Print CStr(Ndx), Dest(Ndx)
                End If
            Next Ndx
        Else
            Debug.Print "Dest is not allocated."
        End If
    Else
        Debug.Print "CopyArray returneed False"
    End If
End Sub


Public Sub DemoCopyArraySubSetToArray()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoCopyArraySubSetToArray
    ' This procedure demonstrates the CopyArraySubSetToArray fuction.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim InputArray(1 To 10) As Long
    Dim ResultArray() As Long

    Dim StartNdx As Long
    Dim EndNdx As Long
    Dim DestNdx As Long
    Dim B As Boolean
    Dim N As Long

    For N = LBound(InputArray) To UBound(InputArray)
        InputArray(N) = N * 10
    Next N

    ReDim ResultArray(1 To 10)
    For N = LBound(ResultArray) To UBound(ResultArray)
        ResultArray(N) = -N
    Next N

    StartNdx = 1
    EndNdx = 5
    DestNdx = 3
    B = ArrayLib.CopyArraySubSetToArray(InputArray:=InputArray, ResultArray:=ResultArray, _
                                        FirstElementToCopy:=StartNdx, LastElementToCopy:=EndNdx, DestinationElement:=DestNdx)
    
    If B = True Then
        If ArrayLib.IsArrayAllocated(Arr:=ResultArray) = True Then
            For N = LBound(ResultArray) To UBound(ResultArray)
                If IsObject(ResultArray(N)) = True Then
                    Debug.Print CStr(N), "is object"
                Else
                    Debug.Print CStr(N), ResultArray(N)
                End If
            Next N
        Else
            Debug.Print "ResultArray is not allocated"
        End If
    Else
        Debug.Print "CopyArraySubSetToArray returned False"
    End If
End Sub


Public Sub DemoCopyNonNothingObjectsToArray()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoCopyNonNothingObjectsToArray
    ' This demonstrates the CopyNonNothingObjectsToArray procedure.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim SourceArray(1 To 5) As Object
    Dim ResultArray() As Object
    Dim B As Boolean
    Dim N As Long

    Set SourceArray(1) = Range("a1")
    Set SourceArray(2) = Range("A2")
    Set SourceArray(3) = Nothing
    Set SourceArray(4) = Nothing
    Set SourceArray(5) = Range("A5")
    B = ArrayLib.CopyNonNothingObjectsToArray(SourceArray:=SourceArray, ResultArray:=ResultArray, NoAlerts:=False)
    If B = True Then
        For N = LBound(ResultArray) To UBound(ResultArray)
            Debug.Print CStr(N), ResultArray(N).Address
        Next N
    Else
        Debug.Print "CopyNonNothingObjectsToArray returned False"
    End If
End Sub


Public Sub DemoDataTypeOfArray()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoDataTypeOfArray
    ' This demonstrates the DataTypeOfArray function.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim A(1 To 4) As String
    Dim T As VbVarType

    T = ArrayLib.DataTypeOfArray(A)
    Debug.Print T
End Sub


Public Sub DemoDeleteArrayElement()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoDeleteArrayElement
    ' This demonstrates the DeleteArrayElement procedure
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim Stat(1 To 3) As Long
    Dim Dyn() As Variant
    Dim N As Long
    Dim B As Boolean

    ReDim Dyn(1 To 3)

    Stat(1) = 1
    Stat(2) = 2
    Stat(3) = 3
    Dyn(1) = "abc"
    Dyn(2) = 1234
    Dyn(3) = "ABC"

    B = ArrayLib.DeleteArrayElement(InputArray:=Stat, ElementNumber:=1, ResizeDynamic:=False)
    If B = True Then
        For N = LBound(Stat) To UBound(Stat)
            Debug.Print CStr(N), Stat(N)
        Next N
    Else
        Debug.Print "DeleteArrayElement returned false"
    End If

    B = ArrayLib.DeleteArrayElement(InputArray:=Dyn, ElementNumber:=2, ResizeDynamic:=False)
    If B = True Then
        For N = LBound(Dyn) To UBound(Dyn)
            Debug.Print CStr(N), Dyn(N)
        Next N
    Else
        Debug.Print "DeleteArrayElement returned false"
    End If
End Sub


Public Sub DemoFirstNonEmptyStringIndexInArray()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoFirstNonEmptyStringIndexInArray
    ' This demonstrates the FirstNonEmptyStringIndexInArray procedure
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim A(1 To 4) As String
    Dim R As Long
    A(1) = vbNullString
    A(2) = vbNullString
    A(3) = "A"
    A(4) = "B"
    R = ArrayLib.FirstNonEmptyStringIndexInArray(InputArray:=A)
    Debug.Print "FirstNonEmptyStringIndexInArray", CStr(R)
End Sub


Public Sub DemoInsertElementIntoArray()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoInsertElementIntoArray
    ' This demonstartes the InsertElementIntoArray function.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim Arr() As Long
    Dim N As Long
    Dim B As Boolean

    ReDim Arr(1 To 10)
    For N = LBound(Arr) To UBound(Arr)
        Arr(N) = N * 10
    Next N

    B = ArrayLib.InsertElementIntoArray(InputArray:=Arr, Index:=5, Value:=12345)
    If B = True Then
        For N = LBound(Arr) To UBound(Arr)
            Debug.Print CStr(N), Arr(N)
        Next N
    Else
        Debug.Print "InsertElementIntoArray returned false."
    End If
End Sub


Public Sub DemoIsArrayAllDefault()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoIsArrayAllDefault
    ' This demonstartes the IsArrayAllDefault function.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim L(1 To 4) As Long
    Dim Obj(1 To 4) As Object
    Dim B As Boolean

    B = ArrayLib.IsArrayAllDefault(L)
    Debug.Print "IsArrayAllDefault L", B

    B = ArrayLib.IsArrayAllDefault(Obj)
    Debug.Print "IsArrayAllDefault Obj", B

    Set Obj(1) = Range("A1")
    B = ArrayLib.IsArrayAllDefault(Obj)
    Debug.Print "IsArrayAllDefault Obj", B
End Sub


Public Sub DemoIsArrayAllNumeric()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoIsArrayAllNumeric
    ' This demonstrates the IsArrayAllNumeric function.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim V(1 To 3) As Variant
    Dim B As Boolean
    V(1) = "abc"
    V(2) = 2
    V(3) = Empty
    B = ArrayLib.IsArrayAllNumeric(Arr:=V, AllowNumericStrings:=True)
    Debug.Print "IsArrayAllNumeric:", B
End Sub


Public Sub DemoIsArrayAllocated()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoIsArrayAllocated
    ' This demonstrates the IsArrayAllocated function.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim B As Boolean
    Dim AllocArray(1 To 3) As Variant
    Dim UnAllocArray() As Variant
    B = ArrayLib.IsArrayAllocated(Arr:=AllocArray)
    Debug.Print "IsArrayAllocated AllocArray:", B
    B = ArrayLib.IsArrayAllocated(Arr:=UnAllocArray)
    Debug.Print "IsArrayAllocated UnAllocArray:", B
End Sub


Public Sub DemoIsArrayDynamic()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoIsArrayDynamic
    ' This demonstrates the IsArrayDynamic function.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim B As Boolean
    Dim StaticArray(1 To 3) As Long
    Dim DynArray() As Long
    ReDim DynArray(1 To 3)
    B = ArrayLib.IsArrayDynamic(Arr:=StaticArray)
    Debug.Print "IsArrayDynamic StaticArray:", B

    B = ArrayLib.IsArrayDynamic(Arr:=DynArray)
    Debug.Print "IsArrayDynamic DynArray:", B
End Sub


Public Sub DemoIsArrayEmpty()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoIsArrayEmpty
    ' This demonstartes the IsArrayEmpty funtcion.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim EmptyArray() As Long
    Dim NonEmptyArray() As Long
    ReDim NonEmptyArray(1 To 3)
    Dim B As Boolean

    B = ArrayLib.IsArrayEmpty(EmptyArray)
    Debug.Print "IsArrayEmpty: EmptyArray:", B

    B = ArrayLib.IsArrayEmpty(NonEmptyArray)
    Debug.Print "IsArrayEmpty: NonEmptyArray:", B
End Sub


Public Sub DemoIsArrayObjects()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoIsArrayObjects
    ' This demonstrates the IsArrayObjects function.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim V(1 To 3) As Variant
    Dim B As Boolean
    V(1) = Range("A1")
    Set V(2) = Nothing
    Set V(3) = Range("A3")

    B = ArrayLib.IsArrayObjects(InputArray:=V, AllowNothing:=True)
    Debug.Print "IsArrayObjects With AllowNothing = True:", B

    B = ArrayLib.IsArrayObjects(InputArray:=V, AllowNothing:=False)
    Debug.Print "IsArrayObjects With AllowNothing = False:", B
End Sub


Public Sub DemoIsNumericDataType()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoIsNumericDataType
    ' This demonstrates the IsNumericDataType function.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim V As Variant
    Dim VEmpty As Variant
    Dim S As String
    Dim B As Boolean
    V = 123
    S = "123"

    B = ArrayLib.IsNumericDataType(V)
    Debug.Print "IsNumericDataType:", B

    B = ArrayLib.IsNumericDataType(S)
    Debug.Print "IsNumericDataType:", B

    B = ArrayLib.IsNumericDataType(VEmpty)
    Debug.Print "IsNumericDataType:", B

    V = Array(1, 2, 3)
    B = ArrayLib.IsNumericDataType(V)
    Debug.Print "IsNumericDataType:", B

    V = Array("a", "b", "c")
    B = ArrayLib.IsNumericDataType(V)
    Debug.Print "IsNumericDataType:", B
End Sub


Public Sub DemoIsVariantArrayConsistent()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoIsVariantArrayConsistent
    ' This demonstrates the IsVariantArrayConsistent function.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim B As Boolean
    Dim V(1 To 3) As Variant
    Set V(1) = Range("A1")
    Set V(2) = Nothing
    Set V(3) = Range("A3")


    B = ArrayLib.IsVariantArrayConsistent(V)
    Debug.Print "IsVariantArrayConsistent:", B
End Sub


Public Sub DemoIsVariantArrayNumeric()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoIsVariantArrayNumeric
    ' This demonstrates the IsVariantArrayNumeric function.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim B As Boolean
    Dim V(1 To 3) As Variant
    V(1) = 123
    Set V(2) = Range("A1")
    V(3) = 789
    B = ArrayLib.IsVariantArrayNumeric(V)
    Debug.Print "IsVariantArrayNumeric", B

End Sub


Public Sub DemoMoveEmptyStringsToEndOfArray()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoMoveEmptyStringsToEndOfArray
    ' This demonstrates the MoveEmptyStringsToEndOfArray function.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim B As Boolean
    Dim N As Long
    Dim S(1 To 5) As String
    S(1) = vbNullString
    S(2) = vbNullString
    S(3) = "C"
    S(4) = "D"
    S(5) = "E"
    B = ArrayLib.MoveEmptyStringsToEndOfArray(InputArray:=S)
    If B = True Then
        For N = LBound(S) To UBound(S)
            If S(N) = vbNullString Then
                Debug.Print CStr(N), "is vbNullString"
            Else
                Debug.Print CStr(N), S(N)
            End If
        Next N
    Else
        Debug.Print "MoveEmptyStringsToEndOfArray returned False"
    End If
End Sub


Public Sub DemoNumberOfArrayDimensions()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoNumberOfArrayDimensions
    ' This demonstrates the NumberOfArrayDimensions function.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim EmptyArray() As Long
    Dim OneArray(1 To 3) As Long
    Dim ThreeArray(1 To 3, 1 To 2, 1 To 1) As Variant
    Dim N As Long

    N = ArrayLib.NumberOfArrayDimensions(Arr:=EmptyArray)
    Debug.Print "NumberOfArrayDimensions EmptyArray", N

    N = ArrayLib.NumberOfArrayDimensions(Arr:=OneArray)
    Debug.Print "NumberOfArrayDimensions OneArray", N

    N = ArrayLib.NumberOfArrayDimensions(Arr:=ThreeArray)
    Debug.Print "NumberOfArrayDimensions ThreeArray", N
End Sub


Public Sub DemoNumElements()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoNumElements
    ' This demonstrates the NumElements function.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim N As Long
    Dim EmptyArray() As Long
    Dim OneArray(1 To 3) As Long
    Dim ThreeArray(1 To 3, 1 To 2, 1 To 1) As Variant

    N = ArrayLib.NumElements(EmptyArray, 1)
    Debug.Print "NumElements  EmptyArray", N

    N = ArrayLib.NumElements(OneArray, 1)
    Debug.Print "NumElements OneArray", N

    N = ArrayLib.NumElements(ThreeArray, 3)
    Debug.Print "NumElements ThreeArray", N
End Sub


Public Sub DemoResetVariantArrayToDefaults()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoResetVariantArrayToDefaults
    ' This demonstrates the ResetVariantArrayToDefaults procedure.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim V(1 To 5) As Variant
    Dim B As Boolean
    Dim N As Long

    V(1) = CInt(123)
    V(2) = "abcd"
    Set V(3) = Range("A1")
    V(4) = CDec(123)
    V(5) = Null
    B = ArrayLib.ResetVariantArrayToDefaults(V)
    If B = True Then
        For N = LBound(V) To UBound(V)
            If IsObject(V(N)) = True Then
                If V(N) Is Nothing Then
                    Debug.Print CStr(N), "Is Nothing"
                Else
                    Debug.Print CStr(N), "Is Object"
                End If
            Else
                Debug.Print CStr(N), TypeName(V(N)), V(N)
            End If
        Next N
    Else
        Debug.Print "ResetVariantArrayToDefaults  returned false"
    End If
End Sub


Public Sub DemoReverseArrayInPlace()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoReverseArrayInPlace
    ' This demonstrates the ReverseArrayInPlace function.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim V(1 To 5) As Long
    Dim N As Long
    Dim B As Boolean
    V(1) = 1
    V(2) = 2
    V(3) = 3
    V(4) = 4
    V(5) = 5
    B = ArrayLib.ReverseArrayInPlace(InputArray:=V)
    If B = True Then
        Debug.Print "REVERSED ARRAY --------------------------------------"
        For N = LBound(V) To UBound(V)
            Debug.Print V(N)
        Next N
    End If
End Sub


Public Sub DemoReverseArrayOfObjectsInPlace()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoReverseArrayOfObjectsInPlace
    ' This demonstrates the ReverseArrayOfObjectsInPlace function.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim B As Boolean
    Dim N As Long
    Dim V(1 To 5) As Object
    Set V(1) = Range("A1")
    Set V(2) = Nothing
    Set V(3) = Range("A3")
    Set V(4) = Range("A4")
    Set V(5) = Range("A5")
    B = ArrayLib.ReverseArrayOfObjectsInPlace(InputArray:=V)
    If B = True Then
        Debug.Print "REVERSED ARRAY --------------------------------------"
        For N = LBound(V) To UBound(V)
            If V(N) Is Nothing Then
                Debug.Print CStr(N), "Is Nothing"
            Else
                Debug.Print CStr(N), V(N).Address
            End If
        Next N
    End If
End Sub


Public Sub DemoSetObjectArrrayToNothing()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoSetObjectArrrayToNothing
    ' This demonstrates the SetObjectArrrayToNothing function.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim StaticArray(1 To 2) As Range
    Dim DynamicArray(1 To 2) As Range
    Dim B As Boolean
    Dim N As Long

    Set StaticArray(1) = Range("A1")
    Set StaticArray(2) = Nothing
    Set DynamicArray(1) = Range("A1")
    Set DynamicArray(2) = Range("A2")
    B = ArrayLib.SetObjectArrayToNothing(StaticArray)
    If B = True Then
        For N = LBound(StaticArray) To UBound(StaticArray)
            If StaticArray(N) Is Nothing Then
                Debug.Print CStr(N), "is nothing "
            End If
        Next N
    End If
    
    B = ArrayLib.SetObjectArrayToNothing(DynamicArray)
    If B = True Then
        For N = LBound(DynamicArray) To UBound(DynamicArray)
            If DynamicArray(N) Is Nothing Then
                Debug.Print CStr(N), "is nothing "
            End If
        Next N
    End If
End Sub


Public Sub DemoVectorsToArray()
    '''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoVectorsToArray
    ' This demonstartes the VectorsToArray function.
    '''''''''''''''''''''''''''''''''''''''''''''''''
    Dim A() As Variant
    Dim B As Boolean
    Dim R As Long
    Dim C As Long
    Dim S As String

    Dim AA() As Variant
    Dim bb() As Variant
    Dim CC() As String

    ReDim AA(0 To 2)
    ReDim bb(1 To 5)
    ReDim CC(2 To 5)
    
    AA(0) = 16
    AA(1) = 2
    AA(2) = 3
    'AA(3) = 3
    bb(1) = 11
    bb(2) = 22
    bb(3) = 33
    bb(4) = 44
    bb(5) = 55
    CC(2) = "A"
    CC(3) = "B"
    CC(4) = "C"
    CC(5) = "D"

    B = ArrayLib.VectorsToArray(A, AA, bb, CC)
    If B = True Then
        For R = LBound(A, 1) To UBound(A, 1)
            S = vbNullString
            For C = LBound(A, 2) To UBound(A, 2)
                S = S & A(R, C) & " "
            Next C
            Debug.Print S
        Next R
    Else
        Debug.Print "VectorsToArray Failed"
    End If
End Sub


Public Sub DemoTransposeArray()
    Dim A() As Long
    Dim B As Variant

    Dim RowNdx As Long
    Dim ColNdx As Long
    Dim S As String

    ReDim A(1 To 3, 2 To 5)
    A(1, 2) = 1
    A(1, 3) = 2
    A(1, 4) = 3
    A(1, 5) = 33
    A(2, 2) = 4
    A(2, 3) = 5
    A(2, 4) = 6
    A(2, 5) = 66
    A(3, 2) = 7
    A(3, 3) = 8
    A(3, 4) = 9
    A(3, 5) = 100
    Debug.Print "LBound1: " & CStr(LBound(A, 1)) & " Ubound1: " & CStr(UBound(A, 1)), _
        "LBound2: " & CStr(LBound(A, 2)) & " UBound2: " & CStr(UBound(A, 2))

    For RowNdx = LBound(A, 1) To UBound(A, 1)
        S = vbNullString
        For ColNdx = LBound(A, 2) To UBound(A, 2)
            S = S & A(RowNdx, ColNdx) & " "
        Next ColNdx
        Debug.Print S
    Next RowNdx
    Debug.Print "Transposed Array:"
    B = ArrayLib.TransposeArray(InputArr:=A)
    If Not IsEmpty(B) Then
        Debug.Print "LBound1: " & CStr(LBound(B, 1)) & " Ubound1: " & CStr(UBound(B, 1)), _
            "LBound2: " & CStr(LBound(B, 2)) & " UBound2: " & CStr(UBound(B, 2))
        S = vbNullString
        For RowNdx = LBound(B, 1) To UBound(B, 1)
            S = vbNullString
            For ColNdx = LBound(B, 2) To UBound(B, 2)
                S = S & B(RowNdx, ColNdx) & " "
            Next ColNdx
            Debug.Print S
        Next RowNdx
    Else
        Debug.Print "Error In Transpose Array"
    End If
End Sub


Public Sub TestChangeBoundsOfArray()
    Dim NewLB As Long
    Dim NewUB As Long
    Dim B As Boolean
    Dim N As Long
    Dim M As Long
    'Dim Arr() As Range
    'Dim Arr() As Long
    'Dim Arr() As Variant
    Dim Arr() As ArrayLibDemoClass

    ReDim Arr(5 To 7)
    'Set Arr(5) = Range("A1")
    'Set Arr(6) = Range("A2")
    'Set Arr(7) = Range("A3")
    'Arr(5) = 11
    'Arr(6) = 22
    'Arr(7) = 33
    'Arr(5) = Array(1, 2, 3)
    'Arr(6) = Array(4, 5, 6)
    'Arr(7) = Array(7, 8, 9)

    Set Arr(5) = New ArrayLibDemoClass
    Set Arr(6) = New ArrayLibDemoClass
    Set Arr(7) = New ArrayLibDemoClass
    Arr(5).Name = "Name 1"
    Arr(5).Value = 1
    Arr(6).Name = "Name 2"
    Arr(6).Value = 3
    Arr(7).Name = "Name 3"
    Arr(7).Value = 3

    NewLB = 20
    NewUB = 25
    B = ArrayLib.ChangeBoundsOfArray(InputArr:=Arr, NewLowerBound:=NewLB, NewUpperBound:=NewUB)
    Debug.Print "New LBound: " & CStr(LBound(Arr)), "New UBound: " & CStr(UBound(Arr))
    For N = LBound(Arr) To UBound(Arr)
        If IsObject(Arr(N)) = True Then
            'Debug.Print "Object: " & TypeName(Arr(N))
            If Arr(N) Is Nothing Then
                Debug.Print "Object Is Nothing"
            Else
                '            Debug.Print "Object: " & Arr(N).Name, Arr(N).Value
                Debug.Print "Object: " & TypeName(Arr(N))
            End If
        Else
            '        If IsArray(Arr(N)) = True Then
            '            For M = LBound(Arr(N)) To UBound(Arr(N))
            '                Debug.Print Arr(N)(M)
            '            Next M
            '        Else
            If IsEmpty(Arr(N)) = True Then
                Debug.Print "Empty"
            ElseIf Arr(N) = vbNullString Then
                Debug.Print "vbNullString"
            Else
                Debug.Print Arr(N)
            End If
            '        End If
        End If
    Next N
End Sub


Public Sub TestIsArraySorted()
    Dim S(1 To 3) As String
    Dim L(1 To 3) As Long
    Dim R As Variant
    Dim Desc As Boolean

    Desc = True
    S(1) = "B"
    S(2) = "B"
    S(3) = "A"

    L(1) = 1
    L(2) = 2
    L(3) = 3

    R = ArrayLib.IsArraySorted(TestArray:=S, Descending:=Desc)
    If IsNull(R) = True Then
        Debug.Print "Error From IsArraySorted"
    Else
        If R = True Then
            Debug.Print "Array Is Sorted"
        Else
            Debug.Print "Array is Unsorted"
        End If
    End If
End Sub


Public Sub TestCombineTwoDArrays()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' TestCombineTwoDArrays
    ' This illustrates the CombineTwoDArrays function.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim X As Long
    Dim Y As Long
    Dim N As Long
    Dim S As String
    Dim V As Variant
    Dim E As Variant
    
    Dim A() As String
    Dim B() As String
    Dim C() As String
    Dim D() As String
    
    '''''''''''''''''''''''''''''''''''
    ' Ensure it works on 1-Based arrays
    '''''''''''''''''''''''''''''''''''
    ReDim A(1 To 2, 1 To 2)
    ReDim B(1 To 2, 1 To 2)
    A(1, 1) = "a"
    A(1, 2) = "b"
    A(2, 1) = "c"
    A(2, 2) = "d"
    B(1, 1) = "e"
    B(1, 2) = "f"
    B(2, 1) = "g"
    B(2, 2) = "h"

    Debug.Print "--- 1 BASED ARRAY -----------------------"
    V = ArrayLib.CombineTwoDArrays(A, B)
    DebugPrint2DArray V

    '''''''''''''''''''''''''''''''''''
    ' Ensure it works on 0-Based arrays
    '''''''''''''''''''''''''''''''''''
    ReDim A(0 To 1, 0 To 1)
    ReDim B(0 To 1, 0 To 1)
    A(0, 0) = "a"
    A(0, 1) = "b"
    A(1, 0) = "c"
    A(1, 1) = "d"

    B(0, 0) = "e"
    B(0, 1) = "f"
    B(1, 0) = "g"
    B(1, 1) = "h"

    Debug.Print "--- 0 BASED ARRAY -----------------------"
    V = ArrayLib.CombineTwoDArrays(A, B)
    DebugPrint2DArray V
    
    '''''''''''''''''''''''''''''''''''''''''''
    ' Ensure it works on Positive-Based arrays
    '''''''''''''''''''''''''''''''''''''''''''
    ReDim A(5 To 6, 5 To 6)
    ReDim B(5 To 6, 5 To 6)
    A(5, 5) = "a"
    A(5, 6) = "b"
    A(6, 5) = "c"
    A(6, 6) = "d"

    B(5, 5) = "e"
    B(5, 6) = "f"
    B(6, 5) = "g"
    B(6, 6) = "h"
    
    Debug.Print "--- POSITIVE BASED ARRAY -----------------------"
    V = ArrayLib.CombineTwoDArrays(A, B)
    DebugPrint2DArray V

    '''''''''''''''''''''''''''''''''''''''''''
    ' Ensure it works on Negative-Based arrays
    '''''''''''''''''''''''''''''''''''''''''''
    ReDim A(-6 To -5, -6 To -5)
    ReDim B(-6 To -5, -6 To -5)
    A(-6, -6) = "a"
    A(-6, -5) = "b"
    A(-5, -6) = "c"
    A(-5, -5) = "d"

    B(-6, -6) = "e"
    B(-6, -5) = "f"
    B(-5, -6) = "g"
    B(-5, -5) = "h"
    
    Debug.Print "--- NEGATIVE BASED ARRAY -----------------------"
    V = ArrayLib.CombineTwoDArrays(A, B)
    DebugPrint2DArray V

    ''''''''''''''''''''''''''''''''''''''''''
    ' Ensure Nesting Works
    ''''''''''''''''''''''''''''''''''''''''''
    ReDim A(1 To 2, 1 To 2)
    ReDim B(1 To 2, 1 To 2)
    ReDim C(1 To 2, 1 To 2)
    ReDim D(1 To 2, 1 To 2)
    
    A(1, 1) = "a"
    A(1, 2) = "b"
    A(2, 1) = "c"
    A(2, 2) = "d"
    
    B(1, 1) = "e"
    B(1, 2) = "f"
    B(2, 1) = "g"
    B(2, 2) = "h"

    C(1, 1) = "i"
    C(1, 2) = "j"
    C(2, 1) = "k"
    C(2, 2) = "l"
    
    D(1, 1) = "m"
    D(1, 2) = "n"
    D(2, 1) = "o"
    D(2, 2) = "p"

    Debug.Print "--- NESTED CALLS -----------------------"
    V = ArrayLib.CombineTwoDArrays(ArrayLib.CombineTwoDArrays(ArrayLib.CombineTwoDArrays(A, B), C), D)
    DebugPrint2DArray V
End Sub


Private Sub DebugPrint2DArray(ByRef Arr As Variant)
    Dim Y As Long
    Dim X As Long
    Dim S As String

    For Y = LBound(Arr, 1) To UBound(Arr, 1)
        S = vbNullString
        For X = LBound(Arr, 2) To UBound(Arr, 2)
            S = S & Arr(Y, X) & " "
        Next X
        Debug.Print S
    Next Y
End Sub


Public Sub DemoExpandArray()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoExpandArray
    ' This demonstrates the ExpandArray function.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim A As Variant
    Dim B As Variant
    Dim RowNdx As Long
    Dim ColNdx As Long
    Dim S As String

    'ReDim A(-5 To -3, 0 To 3)
    'A(-5, 0) = "a"
    'A(-5, 1) = "b"
    'A(-5, 2) = "c"
    'A(-5, 3) = "d"
    'A(-4, 0) = "e"
    'A(-4, 1) = "f"
    'A(-4, 2) = "g"
    'A(-4, 3) = "h"
    'A(-3, 0) = "i"
    'A(-3, 1) = "j"
    'A(-3, 2) = "k"
    'A(-3, 3) = "l"
    '

    ReDim A(1 To 2, 1 To 4)
    A(1, 1) = "a"
    A(1, 2) = "b"
    A(1, 3) = "c"
    A(1, 4) = "d"
    A(2, 1) = "e"
    A(2, 2) = "f"
    A(2, 3) = "g"
    A(2, 4) = "h"

    Dim C As Variant

    Debug.Print "BEFORE:================================="
    For RowNdx = LBound(A, 1) To UBound(A, 1)
        S = vbNullString
        For ColNdx = LBound(A, 2) To UBound(A, 2)
            S = S & A(RowNdx, ColNdx) & " "
        Next ColNdx
        Debug.Print S
    Next RowNdx

    S = vbNullString
    B = ArrayLib.ExpandArray(Arr:=A, WhichDim:=1, AdditionalElements:=3, FillValue:="x")

    'C = ExpandArray(ExpandArray(Arr:=A, WhichDim:=1, AdditionalElements:=3, FillValue:="F"), _
    WhichDim:=2, AdditionalElements:=4, FillValue:="S")

    Debug.Print "AFTER:================================="
    For RowNdx = LBound(B, 1) To UBound(B, 1)
        S = vbNullString
        For ColNdx = LBound(B, 2) To UBound(B, 2)
            S = S & B(RowNdx, ColNdx) & " "
        Next ColNdx
        Debug.Print S
    Next RowNdx

    'Debug.Print "AFTER:================================="
    'For RowNdx = LBound(C, 1) To UBound(C, 1)
    '    S = vbNullString
    '    For ColNdx = LBound(C, 2) To UBound(C, 2)
    '        S = S & C(RowNdx, ColNdx) & " "
    '    Next ColNdx
    '    Debug.Print S
    'Next RowNdx
    '
End Sub


Public Sub DemoSwapArrayRowsAndColumns()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoSwapArrayRowsAndColumns
    ' This demonstrates the SwapArrayRows and SwapArrayColumns procedures.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim R As Long
    Dim C As Long
    Dim S As String
    Dim A(1 To 3, 1 To 2) As Variant
    Dim B() As Variant
    A(1, 1) = "a"
    A(1, 2) = "b"
    A(2, 1) = "c"
    A(2, 2) = "d"
    A(3, 1) = "e"
    A(3, 2) = "f"

    Debug.Print "BEFORE============================"
    For R = LBound(A, 1) To UBound(A, 1)
        S = vbNullString
        For C = LBound(A, 2) To UBound(A, 2)
            S = S & A(R, C) & " "
        Next C
        Debug.Print S
    Next R

    'B = SwapArrayRows(Arr:=A, Row1:=2, Row2:=3)
    B = ArrayLib.SwapArrayColumns(Arr:=A, Col1:=1, Col2:=2)

    Debug.Print "AFTER============================"
    For R = LBound(B, 1) To UBound(B, 1)
        S = vbNullString
        For C = LBound(B, 2) To UBound(B, 2)
            S = S & B(R, C) & " "
        Next C
        Debug.Print S
    Next R
End Sub


Public Sub DemoGetColumn()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoGetColumn
    ' This demonstrates the GetColumn function.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim InputArr(1 To 2, 1 To 3) As Variant
    Dim Result() As Variant
    Dim N As Long
    InputArr(1, 1) = 1
    InputArr(1, 2) = 2
    InputArr(1, 3) = 3
    InputArr(2, 1) = 4
    InputArr(2, 2) = 5
    InputArr(2, 3) = 6
    Result = ArrayLib.GetColumn(InputArr, 3)
    If Not IsEmpty(Result) Then
        For N = LBound(Result) To UBound(Result)
            Debug.Print Result(N)
        Next N
    Else
        Debug.Print "Error from GetColumn"
    End If
End Sub


Public Sub DemoGetRow()
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' DemoGetRow
    ' This demonstrates the GetRow function.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim InputArr(1 To 2, 1 To 3) As Variant
    Dim Result As Variant
    Dim N As Long
    InputArr(1, 1) = 1
    InputArr(1, 2) = 2
    InputArr(1, 3) = 3
    InputArr(2, 1) = 4
    InputArr(2, 2) = 5
    InputArr(2, 3) = 6
    Result = ArrayLib.GetRow(InputArr, 2)
    If Not IsEmpty(Result) Then
        For N = LBound(Result) To UBound(Result)
            Debug.Print Result(N)
        Next N
    Else
        Debug.Print "Error from GetColumn"
    End If
End Sub
