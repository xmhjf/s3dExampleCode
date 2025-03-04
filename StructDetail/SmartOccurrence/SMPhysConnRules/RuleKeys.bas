Attribute VB_Name = "RuleKeys"
Option Explicit

Public Const INPUT_CONN_OBJECT1 = "Object1"  'Bounded object for TeeWeld
Public Const INPUT_CONN_OBJECT2 = "Object2"
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

Public Const LAP_WELD1 As String = "LapWeld1"
Public Const LAP_WELD2 As String = "LapWeld2"

Public Const CHAIN_WELD As String = "ChainWeld"
Public Const FILLET_WELD1 As String = "FilletWeld1"
Public Const FILLET_WELD2 As String = "FilletWeld2"
Public Const STAGGERED_WELD As String = "StaggeredWeld"
Public Const TEE_WELD_CHILL As String = "TeeWeldChill"
Public Const TEE_WELD_K As String = "TeeWeldK"
Public Const TEE_WELD_V As String = "TeeWeldV"
Public Const TEE_WELD_X As String = "TeeWeldX"
Public Const TEE_WELD_Y As String = "TeeWeldY"
Public Const ZIG_ZAG_WELD As String = "ZigZagWeld"

Public Const BUTT_WELD_I As String = "ButtWeldI"
Public Const BUTT_WELD_IV As String = "ButtWeldIV"
Public Const BUTT_WELD_IX As String = "ButtWeldIX"
Public Const BUTT_WELD_K As String = "ButtWeldK"
Public Const BUTT_WELD_V As String = "ButtWeldV"
Public Const BUTT_WELD_X As String = "ButtWeldX"
Public Const BUTT_WELD_Y As String = "ButtWeldY"

Public Const TEE_WELD_SQUARE As String = "TeeWeldSquare"

Public Type PARAMETER_INFO
   sName As String
   vValue As Variant
   bOverridden As Boolean
   eArgumentType As IMSArgumentType
End Type

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


