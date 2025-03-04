Attribute VB_Name = "modLogfile"
Option Explicit

'
'   modLogfile.bas
'
'   module for writing to a logfile
'
'  The routine
'       LogErrorGer(strMsg)
'
'  will write to a logfile named %TEMP%\logSp3dGer.log
'
'  It will always append to the file.
'  If The file cannot be created, nothing will happen.
'

#Const GERLOG = 0
' to enable logging, set
'#Const GERLOG = 1

Public glngErrorSeverity As Long
Private Declare Function GetTempPath Lib "kernel32" _
   Alias "GetTempPathA" _
   (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long

Private Function GetTempDirectory() As String

    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")

    Dim strTempDir As String, lRet As Long
    Dim cTemp As String
    cTemp = "c:\temp"
    strTempDir = VBA.Space(260)
    lRet = GetTempPath(260, strTempDir)
    strTempDir = Left(strTempDir, lRet)
    
    If Not oFSO.FolderExists(strTempDir) Then
      If Not oFSO.FolderExists(cTemp) Then oFSO.CreateFolder cTemp
      strTempDir = cTemp
    Else
      strTempDir = Left(strTempDir, lRet - 1)
    End If
    
    GetTempDirectory = strTempDir
    
End Function

Public Sub LogErrorGer(strMessage As String, _
            Optional lngSeverity As Long = 1)

#If GERLOG = 1 Then
    
    Dim lngFn As Long
    Dim strPath As String
    Dim strHead As String
    
    On Error GoTo Handler
    
    If lngSeverity < glngErrorSeverity Then
        Exit Sub
    End If
    
    strPath = GetTempDirectory
    If Len(strPath) > 0 Then
        strPath = strPath & "\" & "logSp3dGer.log"
        
        lngFn = FreeFile
        
        On Error Resume Next
        Open strPath For Append As #lngFn
        If Err Then
            Exit Sub
        End If
        strHead = Format(Date, "dd.mm.yyyy ") & Format(Time, "hh:mm:ss ")
        Print #lngFn, strHead & strMessage
        Close #lngFn
    End If
    Exit Sub
Handler:

    
#End If


End Sub




' Debug-functions:
Public Function strVector(p As DVector) As String
strVector = "( " & f(p.X) & ", " & f(p.Y) & ", " & f(p.Z) & " )"
End Function
Public Function strPosition(p As AutoMath.DPosition) As String
strPosition = "( " & f(p.X) & ", " & f(p.Y) & ", " & f(p.Z) & " )"
End Function

Public Function strBx(s As String, Bx() As DVector) As String
Dim i As Long
For i = LBound(Bx) To UBound(Bx)
    strBx = strBx & vbCrLf & s & " " & i & "= " & strVector(Bx(i))
Next i
End Function

Private Function f(d As Double) As String
f = Format(d, "###0.000")
End Function




