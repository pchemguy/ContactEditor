Attribute VB_Name = "CommonRoutines"
'@Folder "Common.Shared"
Option Explicit

'@IgnoreModule MoveFieldCloserToUsage
Private lastID As Double


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


'Public Sub GetSerialIDTest()
'    Debug.Print GenerateSerialID
'    Debug.Print GenerateSerialID
'    Debug.Print GenerateSerialID
'    Debug.Print GenerateSerialID
'    Debug.Print GenerateSerialID
'End Sub

