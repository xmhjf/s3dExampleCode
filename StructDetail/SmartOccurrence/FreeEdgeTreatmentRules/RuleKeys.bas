Attribute VB_Name = "RuleKeys"
Option Explicit

Public Const INPUT_TREATMENT_EDGE = "TreatmentEdge"  'Bounded object for TeeWeld
Public Const PI As Double = 3.141592
Public Const TOL As Double = 0.0001

'constants for Tee Weld Category
Public Const gsNormal = "Normal"
Public Const gsDeep = "Deep"
Public Const gsFull = "Full"
Public Const gsChain = "Chain"
Public Const gsZigZag = "ZigZag"

'constants for Class Society question
Public Const gsLloyds = "Lloyds"
Public Const gsABS = "ABS"
Public Const gsDNV = "DNV"

'constants for Butt Weld Category
Public Const gsOneSided = "OneSided"
Public Const gsTwoSided = "TwoSided"

'constants for Workcenter
Public Const gsMachine1 = "Machine1"
Public Const gsMachine2 = "Machine2"
Public Const gsMachine3 = "Machine3"

'constants for Upside
Public Const gsReference = "Reference"
Public Const gsNonReference = "NonReference"

'constants for bevel method
Public Const gsConstant = "Constant"
Public Const gsVarying = "Varying"



Public Sub ReportError(Optional ByVal sFunctionName As String, Optional ByVal sErrorName As String)
    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description _
                & " - " & "::" & sFunctionName & " - " & sErrorName
End Sub

'********************************************************************
' ' Routine: LogError
'
' Description:  default Error logger
'********************************************************************
Public Function LogError(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "") As IJError
     
    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors
     
    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description
     
     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")
       
    ' add the error to the service : the error is also logged to the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
    Set LogError = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      , _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
End Function


