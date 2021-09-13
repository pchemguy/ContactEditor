Attribute VB_Name = "CommonRoutines"
'@Folder "Common.Shared"
Option Explicit

'@IgnoreModule MoveFieldCloserToUsage
Private lastID As Double


'@EntryPoint
Public Function GetTimeStampMs() As String
    '''' On Windows, the Timer resolution is subsecond, the fractional part (the four characters at the end
    '''' given the format) is concatenated with DateTime. It appears that the Windows' high precision time
    '''' source available via API yields garbage for the fractional part.
    GetTimeStampMs = Format$(Now, "yyyy-MM-dd HH:mm:ss") & Right$(Format$(Timer, "#0.000"), 4)
End Function


'''' The number of seconds since the Epoch is multiplied by 10^4 to bring the first
'''' four fractional places in Timer value into the whole part before trancation.
'''' Long on a 32bit machine does not provide sufficient number of digits,
'''' so returning double. Alternatively, a Currency type could be used.
'@EntryPoint
Public Function GenerateSerialID() As Double
    Dim newID As Double
    Dim secTillLastMidnight As Double
    secTillLastMidnight = CDbl(DateDiff("s", DateSerial(1970, 1, 1), Date))
    newID = Fix((secTillLastMidnight + Timer) * 10 ^ 4)
    If newID > lastID Then
        lastID = newID
    Else
        lastID = lastID + 1
    End If
    GenerateSerialID = lastID
    'GetSerialID = Fix((CDbl(Date) * 100000# + CDbl(Timer) / 8.64))
End Function


'''' When sub/function captures a list of arguments in a ParamArray and passes it
'''' to the next routine expecting a list of arguments, the second routine receives
'''' a 2D array instead of 1D with the outer dimension having a single element.
'''' This function check the arguments and unfolds the outer dimesion as necessary.
'''' Any function accepting a ParamArray argument should be able to use it.
''''
'''' Unfold if the following conditions are satisfied:
''''     - ParamArrayArg is a 1D array
''''     - UBound(ParamArrayArg, 1) = LBound(ParamArrayArg, 1) = 0
''''     - ParamArrayArg(0) is a 1D 0-based array
''''
'''' Return
''''     - ParamArrayArg(0), if unfolding is necessary
''''     - ParamArrayArg, if ParamArrayArg is array, but not all conditions are satisfied
'''' Raise an error if is not an array
'@Description "Unfolds a ParamArray argument when passed from another ParamArray."
Public Function UnfoldParamArray(ByVal ParamArrayArg As Variant) As Variant
Attribute UnfoldParamArray.VB_Description = "Unfolds a ParamArray argument when passed from another ParamArray."
    Guard.NotArray ParamArrayArg
    Dim DoUnfold As Boolean
    DoUnfold = (ArrayLib.NumberOfArrayDimensions(ParamArrayArg) = 1) And (LBound(ParamArrayArg) = 0) And (UBound(ParamArrayArg) = 0)
    If DoUnfold Then DoUnfold = IsArray(ParamArrayArg(0))
    If DoUnfold Then DoUnfold = ((ArrayLib.NumberOfArrayDimensions(ParamArrayArg(0)) = 1) And (LBound(ParamArrayArg(0), 1) = 0))
    If DoUnfold Then
        UnfoldParamArray = ParamArrayArg(0)
    Else
        UnfoldParamArray = ParamArrayArg
    End If
End Function


'@EntryPoint
Public Function GetVarType(ByRef Variable As Variant) As String
    Dim NDim As String
    NDim = IIf(IsArray(Variable), "/Array", vbNullString)
    
    Dim TypeOfVar As VBA.VbVarType
    TypeOfVar = VarType(Variable) And Not vbArray

    Dim ScalarType As String
    Select Case TypeOfVar
        Case vbEmpty
            ScalarType = "vbEmpty"
        Case vbNull
            ScalarType = "vbNull"
        Case vbInteger
            ScalarType = "vbInteger"
        Case vbLong
            ScalarType = "vbLong"
        Case vbSingle
            ScalarType = "vbSingle"
        Case vbDouble
            ScalarType = "vbDouble"
        Case vbCurrency
            ScalarType = "vbCurrency"
        Case vbDate
            ScalarType = "vbDate"
        Case vbString
            ScalarType = "vbString"
        Case vbObject
            ScalarType = "vbObject"
        Case vbError
            ScalarType = "vbError"
        Case vbBoolean
            ScalarType = "vbBoolean"
        Case vbVariant
            ScalarType = "vbVariant"
        Case vbDataObject
            ScalarType = "vbDataObject"
        Case vbDecimal
            ScalarType = "vbDecimal"
        Case vbByte
            ScalarType = "vbByte"
        Case vbUserDefinedType
            ScalarType = "vbUserDefinedType"
        Case Else
            ScalarType = "vbUnknown"
    End Select
    GetVarType = ScalarType & NDim
End Function


'''' Resolves file pathname
''''
'''' This helper routines attempts to interpret provided pathname as
'''' a reference to an existing file:
'''' 1) check if provided reference is a valid absolute file pathname, if not,
'''' 2) construct an array of possible file locations:
''''      - ThisWorkbook.Path & Application.PathSeparator
''''      - Environ("APPDATA") & Application.PathSeparator &
''''          & ThisWorkbook.VBProject.Name & Application.PathSeparator
''''    construct an array of possible file names:
''''      - FilePathName
''''          skip if len=0, or prefix is not relative
''''      - ThisWorkbook.VBProject.Name & Ext (Ext comes from the second argument
'''' 3) loop through all possible path/filename combinations until a valid
''''    pathname is found or all options are exhausted
''''
'''' Args:
''''   FilePathName (string):
''''     File pathname
''''
''''   DefaultExts (string or string/array):
''''     1D array of default extensions or a single default extension
''''
'''' Returns:
''''   String:
''''     Resolved valid absolute pathname pointing to an existing file.
''''
'''' Throws:
''''   Err.FileNotFoundErr:
''''     If provided pathname cannot be resolved to a valid file pathname.
''''
'''' Examples:
''''   >>> ?VerifyOrGetDefaultPath(Environ$("ComSpec"), "")
''''   "C:\Windows\system32\cmd.exe"
''''
'@Description "Resolves file pathname"
Public Function VerifyOrGetDefaultPath(ByVal FilePathName As String, ByVal DefaultExts As Variant) As String
Attribute VerifyOrGetDefaultPath.VB_Description = "Resolves file pathname"
    Dim PATHuSEP As String
    PATHuSEP = Application.PathSeparator
    Dim PROJuNAME As String
    PROJuNAME = ThisWorkbook.VBProject.Name
    
    Dim FileExist As Variant
    Dim PathNameCandidate As String
        
    '''' === (1) === Check if FilePathName is a valid path to an existing file.
    If Len(FilePathName) > 0 Then
        '''' If matched, Dir returns Len(String) > 0;
        '''' otherwise, returns vbNullString or raises an error
        PathNameCandidate = FilePathName
        On Error Resume Next
        FileExist = FileLen(PathNameCandidate)
        On Error GoTo 0
        If FileExist > 0 Then
            VerifyOrGetDefaultPath = PathNameCandidate
            Exit Function
        End If
    End If
    
    '''' === (2a) === Array of prefixes
    Dim Prefixes As Variant
    Prefixes = Array( _
        ThisWorkbook.Path & PATHuSEP, _
        ThisWorkbook.Path & PATHuSEP & "Library" & PATHuSEP & PROJuNAME & PATHuSEP, _
        Environ$("APPDATA") & PATHuSEP & PROJuNAME & PATHuSEP _
    )
    
    '''' === (2b) === Array of filenames
    Dim NameCount As Long
    NameCount = 0
    
    Dim UseFilePathName As Boolean
    UseFilePathName = Len(FilePathName) > 1 And _
                      Mid$(FilePathName, 1, 1) <> "\" And _
                      Mid$(FilePathName, 2, 1) <> ":"
    If UseFilePathName Then
        NameCount = NameCount + 1
    End If
    If VarType(DefaultExts) = vbString Then
        If Len(DefaultExts) > 0 Then NameCount = NameCount + 1
    ElseIf VarType(DefaultExts) >= vbArray Then
        NameCount = NameCount + UBound(DefaultExts, 1) - LBound(DefaultExts, 1) + 1
        Debug.Assert VarType(DefaultExts(0)) = vbString
    End If
    If NameCount = 0 Then
        VBA.Err.Raise _
            Number:=ErrNo.FileNotFoundErr, _
            Source:="CommonRoutines", _
            Description:="File <" & FilePathName & "> not found!"
    End If
    
    Dim FileNames() As String
    ReDim FileNames(0 To NameCount - 1)
    Dim ExtIndex As Long
    Dim FileNameIndex As Long
    FileNameIndex = 0
    If UseFilePathName Then
        FileNames(FileNameIndex) = FilePathName
        FileNameIndex = FileNameIndex + 1
    End If
    If VarType(DefaultExts) = vbString Then
        If Len(DefaultExts) > 0 Then
            FileNames(FileNameIndex) = PROJuNAME & "." & DefaultExts
        End If
    ElseIf VarType(DefaultExts) >= vbArray Then
        For ExtIndex = LBound(DefaultExts, 1) To UBound(DefaultExts, 1)
            FileNames(FileNameIndex) = PROJuNAME & "." & DefaultExts(ExtIndex)
            FileNameIndex = FileNameIndex + 1
        Next ExtIndex
    End If
    
    '''' === (3) === Loop through pathnames
    Dim PrefixIndex As Long
    
    On Error Resume Next
    For PrefixIndex = 0 To UBound(Prefixes)
        For FileNameIndex = 0 To UBound(FileNames)
            PathNameCandidate = Prefixes(PrefixIndex) & FileNames(FileNameIndex)
            FileExist = FileLen(PathNameCandidate)
            Err.Clear
            If FileExist > 0 Then
                VerifyOrGetDefaultPath = Replace$(PathNameCandidate, _
                                                  PATHuSEP & PATHuSEP, PATHuSEP)
                Exit Function
            End If
        Next FileNameIndex
    Next PrefixIndex
    On Error GoTo 0
    
    VBA.Err.Raise _
        Number:=ErrNo.FileNotFoundErr, _
        Source:="CommonRoutines", _
        Description:="File <" & FilePathName & "> not found!"
End Function


'''' Tests if argument is falsy
''''
'''' Falsy values:
''''   Numeric: 0
''''   String:  vbNullString
''''   Variant: Empty
''''   Object:  Nothing
''''   Boolean: False
''''   Null:    Null
''''
'''' Args:
''''   arg:
''''     Value to be tested for falsiness
''''
'''' Returns:
''''   True, if "arg" is Falsy
''''   Flase, if "arg" is Truthy (not Falsy)
''''
'''' Examples:
''''   >>> ?IsFalsy(0.0#)
''''   True
''''
''''   >>> ?IsFalsy(0.1)
''''   False
''''
''''   >>> ?IsFalsy(Null)
''''   True
''''
''''   >>> ?IsFalsy(Empty)
''''   True
''''
''''   >>> ?IsFalsy(False)
''''   True
''''
''''   >>> ?IsFalsy(Nothing)
''''   True
''''
''''   >>> ?IsFalsy("")
''''   True
''''
'@Description("Tests if argument is falsy: 0, False, vbNullString, Empty, Null, Nothing")
Public Function IsFalsy(ByVal arg As Variant) As Boolean
Attribute IsFalsy.VB_Description = "Tests if argument is falsy: 0, False, vbNullString, Empty, Null, Nothing"
    Select Case VarType(arg)
        Case vbEmpty, vbNull
            IsFalsy = True
        Case vbInteger, vbLong, vbSingle, vbDouble
            IsFalsy = Not CBool(arg)
        Case vbString
            IsFalsy = (arg = vbNullString)
        Case vbObject
            IsFalsy = (arg Is Nothing)
        Case vbBoolean
            IsFalsy = Not arg
        Case Else
            IsFalsy = False
    End Select
End Function
