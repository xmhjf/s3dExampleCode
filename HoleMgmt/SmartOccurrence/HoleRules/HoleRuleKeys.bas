Attribute VB_Name = "HoleRuleKeys"
'******************************************************************************
' Copyright (C) 1998-2002 Intergraph Corporation. All Rights Reserved.
'
' Project: S:\HoleMgmt\Data\SmartOccurrence\HoleRules
'
' File: HoleRuleKeys.bas
'
' Author: Hole Mgmt Team
'
' Abstract: keys for the project
'******************************************************************************

Option Explicit

'Inputs for the Smart Class & Items
Public Const INPUT_HOLETRACE = "HoleTraceAE"

'Library for Definition Custom Methods
Public Const LIBRARY_SOURCE_ID = "HoleRules.HoleDefCM"

Public Const IID_IJPlate = "{53CF4EA0-91BF-11D1-BE56-080036B3A103}"
Public Const IID_IJStructureMaterial = "{E790A7C0-2DBA-11D2-96DC-0060974FF15B}"
Public Const IID_IJPartOcc = "{1146CF94-6B33-11D1-A300-080036409103}"
Public Const IID_IJPlateEROffsetContour_AE = "{73BA14FB-FE7E-42EB-9AB6-F3BB7B2981BB}"

'Inputs for the definition custom assy
Public Const HOLE_FEATURE = 1
Public Const EDGE_REINFORCEMENT = 2
Public Const PLATE_FITTING = 3
Public Const PROFILE_FITTING = 4
Public Const CATALOG_FITTING = 5
Public Const PHYSICAL_CONNECTION = 6

Public Const TKStatusBar = "StatusBar"

'******************************************************************************
' Routine: ReportError
'
' Abstract: Display message box with error message
'******************************************************************************
Public Sub ReportError(Optional ByVal sFunctionName As String, Optional ByVal sErrorName As String)
    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description & _
                        " - " & "::" & sFunctionName & " - " & sErrorName
End Sub

'******************************************************************************
' Routine: UpdateStatusBar
'
' Abstract: Display error message on the status bar
'******************************************************************************
Public Sub UpdateStatusBar(oStatus As String)
'    Dim oTrader As IMSTrader.Trader
'    Dim oStatusBar As Object
    
'    Set oTrader = New Trader
'    Set oStatusBar = oTrader.Service(TKStatusBar, "")
    
'    oStatusBar.Panels(1).Text = oStatus
    
'    Set oTrader = Nothing
'    Set oStatusBar = Nothing

    'this is placed into the middle tier - there should be NO reference to a client tier object
'''PML the middle tier should not be popping up any message boxes either
'''    MsgBox oStatus
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

