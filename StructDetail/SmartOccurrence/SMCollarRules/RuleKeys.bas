Attribute VB_Name = "RuleKeys"
Option Explicit

Public Const INPUT_SLOT As String = "Slot"
Public Const INPUT_PENETRATING As String = "Penetrating"
Public Const INPUT_BOUNDINGPLATE As String = "BoundingPlate"

'Global String Constants for Questions and Answers
Public Const gsStressLevel = "StressLevel"
Public Const gsCollarCreationOrder = "CollarCreationOrder"
Public Const gsCollarSideOfPart = "CollarSideOfPlate"

'Global string constants for Collar SideOfPlate
Public Const gsFlip = "Flip"
Public Const gsNoFlip = "NoFlip"
Public Const gsCentered = "Centered"

'Global string constants for Plate Stress Level
Public Const gsHigh = "High"
Public Const gsMedium = "Medium"
Public Const gsLow = "Low"

'Global string constants for AddCornerSnipe(s) question
Public Const gsAddCornerSnipe = "AddCornerSnipe"
Public Const gsAddCornerSnipes = "AddCornerSnipes"

'Global string constants for answers to AddCornerSnipes
Public Const gsWebLeft = "Web_Left"
Public Const gsWebRight = "Web_Right"
Public Const gsBoth = "Both"
Public Const gsNone = "None"

'Global string constants for AddDrainHole question
Public Const gsAddDrainHole = "AddDrainHole"

Public Const IID_IJPlate = "{53CF4EA0-91BF-11D1-BE56-080036B3A103}"
Public Const IID_IJStructureMaterial = "{E790A7C0-2DBA-11D2-96DC-0060974FF15B}"
Public Const IID_IJCollarPart = "{138C021D-7089-11D5-B0D9-006008676515}"

Public Const CMLIBRARY_COLLARRULES As String = CUSTOMERID + "CollarRules.CollarDefCM"

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


