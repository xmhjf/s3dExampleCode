Attribute VB_Name = "PlnCheckMfgFunctions"
'********************************************************************
' Copyright (C) 1998-2004 Intergraph Corporation.  All Rights Reserved.
'
' File: PlnCheckMfgFunctions.cls
'
' Author: Dick Swager
'
' Abstract: common functions for check manufacturability rules for planning joints
'
'History
'   -               -               Creation
' Kishore       4th Jul 07          TR-CP·116757  Running Check Manufacturability vs. the ship or an assembly - results differ.
' Chaitanya    18th Jul 2010        CR-129897  [ removed client references to support Batch scheduling]
'********************************************************************

Option Explicit

Private Const Module = "PlnCheckMfcty.PlnCheckMfgFunctions"

Public Enum ESeverity
    siError = 101
    siWarining = 102
End Enum

Public Sub TrackText(strText As String, Optional nIndent As Integer)
    On Error GoTo SkipTrack
    
    Dim oTracker As GSTracker.Tracker
    Set oTracker = New GSTracker.Tracker
    On Error GoTo KillTrack
    If nIndent < 0 Then oTracker.Outdent
    oTracker.UseFontDefaults
    oTracker.ForegroundColor = RGB(0, 0, 255)
    oTracker.WriteLn strText
    If nIndent > 0 Then oTracker.Indent

KillTrack:
    Set oTracker = Nothing

SkipTrack:

End Sub


'******************************************************************************
' Routine: GatherPlanningJoints
'
' Abstract: retrives the planning joints from elements
'
' Notes:
'  the check manufacturability object has a collection of monikers. we need
'  to filter out the ones we are interested in and create a collection of the
'  actual objects from the moniker.
'
' Arguments:
'   oMonikers           a collection of monikers (input)
'   oCollection         a collection of objects (output)
'******************************************************************************
Public Sub GatherPlanningJoints(ByVal oMonikers As IJElements, oCollection As IJElements)
    Const Method As String = "GatherPlanningJoints"
    On Error GoTo ErrorHandler

    'get the active connection
'    Dim oTrader As Trader
'    Set oTrader = New Trader
'    Dim oWorkingset As IJDWorkingSet
'    Set oWorkingset = oTrader.Service(TKWorkingSet, "")
'    Dim oConnection As IJDConnection
'    Set oConnection = oWorkingset.ActiveConnection
    Dim oChildren           As IJElements
    Dim oPlnIntHelper       As IJDPlnIntHelper
    
    Dim oPOM                As IJDPOM
    Set oPOM = GetPOM("Model")
    
    'set up object info for support info checks
'    Dim oObjectInfo As WorkingSetLibrary.IJDObjectInfo ' to check interface support
'    Set oObjectInfo = oConnection ' working only on active connection

    'get the HoleTrace from moniker
    Dim oEntity As Object
    Dim vMoniker As Variant
    
    Set oPlnIntHelper = New CPlnIntHelper
    Set oCollection = New JObjectCollection
    
    For Each vMoniker In oMonikers
        On Error Resume Next
        If oPOM.SupportsInterface(vMoniker, "IJPlnJoint") Then
            If Err.Number = 0 Then
                Set oEntity = oPOM.GetObject(vMoniker)
                If TypeOf oEntity Is IJPlnJoint Then
                    oCollection.Add oEntity
                End If
            End If
            
        ElseIf oPOM.SupportsInterface(vMoniker, "IJAssembly") Then
            If Err.Number = 0 Then
                Set oEntity = oPOM.GetObject(vMoniker)
                Set oChildren = oPlnIntHelper.GetStoredProcAssemblyChildren(oEntity, _
                                "IJPlnJoint", False, Nothing, True)
                                
                If oChildren.Count > 0 Then oCollection.AddElements oChildren
                Set oChildren = Nothing
                Set oEntity = Nothing
            End If
        End If
        Err.Clear
        On Error GoTo ErrorHandler
    Next vMoniker
    
Cleanup:
'    Set oTrader = Nothing
'    Set oWorkingset = Nothing
'    Set oConnection = Nothing
'    Set oObjectInfo = Nothing
    Set oEntity = Nothing
    Set vMoniker = Nothing
    Set oPOM = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description
    GoTo Cleanup
    
End Sub ' GatherPlanningJoints

'******************************************************************************
' Routine: GatherPlanningJointSmartOccs
'
' Abstract: retrives the planning joint smart occurrences from elements
'
' Notes:
'  the check manufacturability object has a collection of monikers. we need
'  to filter out the ones we are interested in and create a collection of the
'  actual objects from the moniker.
'
' Arguments:
'   oMonikers           a collection of monikers (input)
'   oCollection         a collection of objects (output)
'******************************************************************************
Public Sub GatherPlanningJointSmartOccs _
        (ByVal oMonikers As IJElements, oCollection As IJElements)
        
    Const Method As String = "GatherPlanningJoints"
    On Error GoTo ErrorHandler

    'get the active connection
'    Dim oTrader As Trader
'    Set oTrader = New Trader
'    Dim oWorkingset As IJDWorkingSet
'    Set oWorkingset = oTrader.Service(TKWorkingSet, "")
'    Dim oConnection As IJDConnection
'    Set oConnection = oWorkingset.ActiveConnection
'
'    'set up object info for support info checks
'    Dim oObjectInfo As WorkingSetLibrary.IJDObjectInfo ' to check interface support
'    Set oObjectInfo = oConnection ' working only on active connection

    'get the HoleTrace from moniker
    Dim oEntity As Object
    Dim vMoniker As Variant
    
    Dim oPOM As IJDPOM
    Set oPOM = GetPOM("Model")
    
    Set oCollection = New JObjectCollection
    For Each vMoniker In oMonikers
        If oPOM.SupportsInterface(vMoniker, "IJPlnJoint") Then
            On Error Resume Next
            Set oEntity = oPOM.GetObject(vMoniker)
            If TypeOf oEntity Is IJPlnJoint Then
                
                ' Get the Smart Occurrence for the Planning Joint.
                Dim oPlnJoint As IJPlnJoint
                Set oPlnJoint = oEntity
                
                Dim oSmartOcc As IJPlnJointSO
                Set oSmartOcc = oPlnJoint.GetSmartOccurrence
                
                ' Add this instance to the collection if it is not
                ' already there.
                Dim bAddOccurrence As Boolean
                bAddOccurrence = True
                Dim oSO As IJPlnJointSO
                For Each oSO In oCollection
                    If oSO Is oSmartOcc Then
                        bAddOccurrence = False
                        Exit For
                    End If
                Next oSO
                
                If bAddOccurrence Then
                    oCollection.Add oSmartOcc
                End If
                
            End If
            On Error GoTo ErrorHandler
        
        ElseIf oPOM.SupportsInterface(vMoniker, "IJAssembly") Then
            On Error Resume Next
            Set oEntity = oPOM.GetObject(vMoniker)
            If TypeOf oEntity Is IJAssembly Then
                
                ' Get the Smart Occurrence for the Planning Joint.
                Dim oAssembly As IJAssembly
                Set oAssembly = oEntity
                
                ' create the helper
                Dim oPlnHelper As GSCADPlnIntHelper.IJDPlnIntHelper
                Set oPlnHelper = New GSCADPlnIntHelper.CPlnIntHelper
                
                ' get all plate and profile children in the assembly
                Dim oChildren As IJDObjectCollection
                Dim oPlateChildren As IJDObjectCollection
                Dim oProfileChildren As IJDObjectCollection
                Set oPlateChildren = oPlnHelper.GetAssemblyChildren(oAssembly, "IJPlatePart", True)
                Set oProfileChildren = oPlnHelper.GetAssemblyChildren(oAssembly, "IJProfilePart", True)
                Set oChildren = oPlateChildren
                oChildren.SetAdd oProfileChildren
                
                ' mark all planning joints dirty to force recalculation of folder
                Dim oChild As Object
                For Each oChild In oChildren
                
                    ' retrieve all of the planning joints from the object
                    Dim oPlnJoints As IJDObjectCollection
                    Set oPlnJoints = oPlnHelper.GetPlnJoints(oChild)
                    
                    For Each oPlnJoint In oPlnJoints
                        Set oSmartOcc = oPlnJoint.GetSmartOccurrence
                        
                        ' Add this instance to the collection if it is not
                        ' already there.
                        bAddOccurrence = True
                        For Each oSO In oCollection
                            If oSO Is oSmartOcc Then
                                bAddOccurrence = False
                                Exit For
                            End If
                        Next oSO
                        
                        If bAddOccurrence Then
                            oCollection.Add oSmartOcc
                        End If
                        
                    Next oPlnJoint
                    
                Next oChild
                
            End If
            On Error GoTo ErrorHandler
        End If
    Next vMoniker
    
    GoTo Cleanup

ErrorHandler:
    MsgBox Err.Description
    
Cleanup:
'    Set oTrader = Nothing
'    Set oWorkingset = Nothing
'    Set oConnection = Nothing
'    Set oObjectInfo = No
    Set oEntity = Nothing
    Set vMoniker = Nothing
    Set oPlnJoint = Nothing
    Set oSmartOcc = Nothing
    Set oPOM = Nothing
    
End Sub ' GatherPlanningJointSmartOccs


Private Function GetPOM(strDbType As String) As IJDPOM
Const Method = "GetPOM"
On Error GoTo ErrHandler
    
    Dim oContext            As IJContext
    Dim oAccessMiddle       As IJDAccessMiddle
    
    Set oContext = GetJContext()
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    Set GetPOM = oAccessMiddle.GetResourceManagerFromType(strDbType)
    
    Set oContext = Nothing
    Set oAccessMiddle = Nothing
    
Exit Function
ErrHandler:
    MsgBox Err.Description
End Function

 
