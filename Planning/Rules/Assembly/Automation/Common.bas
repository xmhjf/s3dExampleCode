Attribute VB_Name = "Common"
'********************************************************************
' Copyright (C) 1998-2002 Intergraph Corporation.  All Rights Reserved.
'
' File: AutoAssemblyAsst.cls
'
' Author: Kamrooz k. Hormoozi
'
' Abstract: Template command assistant
'
' Description:
'       implements common function for commen use.
'
' History
'   May 1 2002     oss\kaho    Created.
'********************************************************************
Option Explicit

Private Const MODULE = "Common."
Public Const IID_IJPort As String = "{5CF7C404-546D-11D2-B328-080036024603}"
Public Const TKWorkingSet = "WorkingSet"
Public Const TKApplicationContext = "ApplicationContext"


' ***************************************************************************
'
' Function
'   GetAssemblyChildrenImpl()
'
' Abstract
'   Implementation of recursive aspects of the GetAssemblyChildren() method.
'
' ***************************************************************************

Public Sub GetAssemblyChildrenImpl( _
        ByVal objAssembly As IJAssembly, _
        ByVal objElements As IJElements, Unprocessed As Boolean)
    Const METHOD As String = "GetAssemblyChildrenImpl"
    On Error GoTo ErrorHelper
          
  
    Dim oPlnIntHelper          As IJDPlnIntHelper
    Dim objTargetObjectCol     As IJElements
    Dim index                  As Long
    Dim objChild               As Object
    
    Set oPlnIntHelper = New CPlnIntHelper
    Set objTargetObjectCol = oPlnIntHelper.GetStoredProcAssemblyChildren(objAssembly, vbNullString, False)
          
    If TypeOf objAssembly Is IJPlnUnprocessedParts And Unprocessed Then
           For index = 1 To objTargetObjectCol.Count
               Set objChild = objTargetObjectCol.Item(index)
               If TypeOf objChild Is IJAssemblyChild Then
                   objElements.Add objChild
               End If
           Next index
           Exit Sub
    End If
    
    For index = 1 To objTargetObjectCol.Count
        Set objChild = objTargetObjectCol.Item(index)
            If TypeOf objChild Is IJPlatePart Or TypeOf objChild Is IJProfile Then
                 If TypeOf objChild Is IJAssemblyChild Then
                    objElements.Add objChild
                End If
            End If
    Next index
    

    Set oPlnIntHelper = Nothing
   
    Exit Sub
ErrorHelper:
    MsgBox Err.Description
End Sub

' ***************************************************************************
'
' Function
'   AddElements()
'
' Abstract
'   Adds the elements from source collection to the destination collection
'   if they're not already there.
'
' ***************************************************************************

Public Sub AddElements(colDst As IJElements, colSrc As IJElements)
Const METHOD As String = "AddElements"
On Error GoTo ErrorHelper
        
    Dim flgWaitForUpdate As Boolean
    
    ' Disable update notifications
    flgWaitForUpdate = colDst.WaitForUpdate
    colDst.WaitForUpdate = True
    
    ' Add elements from source to destination collection
    Dim oItem As Object
    
    For Each oItem In colSrc
    
        If Not colDst.Contains(oItem) Then _
            colDst.Add oItem
    
    Next oItem
    
    ' Restore update notifications
    colDst.WaitForUpdate = flgWaitForUpdate
    
    Exit Sub
ErrorHelper:
    MsgBox Err.Description
End Sub

'**************************************************************************
' Method
'   GetValidProfilesOnPlate
'
' Description:
'   When passed a Plate returns the collection of profiles on this Plate. Profiles are
'   validated that the belong to assemblies from where they can be moved. Profiles in
'   assemblies from where they must not be moved are not in the returned collection. See
'   method validateprofiles for details.
'
'   Uses 3 subs to get profiles from RootPlateSystems, LeafPlateSystems, and detailed
'   Plates.
'
' Arguments:
'   ByVal oPlate As object,             the plate selected by user
'   ByVal oPlateParent As IJAssembly,   the parent of the selected plate
'
' Return Values:
'   A collection of valid profiles. Collection is empty if no valid
'
'**************************************************************************
Public Function GetValidProfilesOnPlate(ByVal oPlate As Object, ByVal oPlateparent As IJAssembly) As Collection
    Const METHOD = "GetValidProfilesOnPlate"
    On Error GoTo ErrorHandler

    Dim oTempProfiles As Collection
    Dim oProfiles As Collection
    Dim oValidProfiles As Collection
    Dim oProfile As IJProfile
    
    'create collection to hold conten from returned collections
    Set oProfiles = New Collection
                        
    'get profiles from leafplatesystem
    If TypeOf oPlate Is IJPlatePart Then
        Set oTempProfiles = FindProfilesFromLeafPlate(oPlate)
        For Each oProfile In oTempProfiles
            oProfiles.Add oProfile
        Next
    End If
    Set oTempProfiles = Nothing
                                
    'get profiles from system
    Set oTempProfiles = FindSystemProfilesOnPlate(oPlate)
    For Each oProfile In oTempProfiles
        oProfiles.Add oProfile
    Next
    Set oTempProfiles = Nothing
    
    'get profiles from detailed part
    Set oTempProfiles = FindDetailedProfilesOnPlate(oPlate)
    For Each oProfile In oTempProfiles
        oProfiles.Add oProfile
    Next
    Set oTempProfiles = Nothing
    
    ' Get MoldedForm Beam System Parts from the PlatePart
    If TypeOf oPlate Is IJPlatePart Then
        'get beams from plate part
        Set oTempProfiles = FindBeamsOnPlate(oPlate)
        For Each oProfile In oTempProfiles
            oProfiles.Add oProfile
        Next
    End If
       Set oTempProfiles = Nothing
    
    'validate that profiles can be moved from the assembly where they are now
    Set oValidProfiles = ValidateProfiles(oPlateparent, oProfiles)
    
    Set GetValidProfilesOnPlate = oValidProfiles

Cleanup:
    Set oTempProfiles = Nothing
    Set oProfiles = Nothing
    Set oValidProfiles = Nothing
    Set oProfile = Nothing
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
    GoTo Cleanup
    
End Function

'**************************************************************************
' Method
'   FindSystemProfilesOnPlate
'
' Description:
'   From the passed in plate gets the profiles on the plate and return in a collection.
'   Only profiles implementing IJAssemblyChild are valid
'
' Arguments:
'   [in] ByVal oPlate As IJPlate, the plate to get profiles from
'
' Return Values:
'   IJDObjectcollection, with valid profiles. If no valid profiles empty collection
'                        is returned.
'**************************************************************************
Public Function FindSystemProfilesOnPlate(ByVal oPlate As IMSPlateEntity.IJPlate) As Collection
    Const METHOD As String = "FindSystemProfilesOnPlate"
    On Error GoTo ErrorHandler
    
    Dim oStructUtility As IJStructEntityUtils
    Dim oSystemChild As IJSystemChild
    Dim oProfile As IJProfile
    Dim oProfileCol As Collection
    Dim oSystem As IJSystem
    Dim oCol As IJDTargetObjectCol
    Dim oOutputColl As New Collection ' IJDObjectCollection
    Dim OObj As Object
    Dim oPlateFromChild As IMSPlateEntity.IJPlate
    
    Set oSystemChild = oPlate
    Set oPlateFromChild = oSystemChild.GetParent
    Set oSystemChild = Nothing
                    
    ' get all the ProfileSystems on plate
    Set oStructUtility = New StructEntityUtils
    Set oProfileCol = oStructUtility.GetProfilesOnPlate(oPlateFromChild)
'ERROR: At the moment this method does not return profile related to splitted
'       (Design split in molded forms) plates. TR 34023 filed against this.

    'check about the collection is empty
    If Not oProfileCol Is Nothing Then
        'find all profile in the collction
        For Each oProfile In oProfileCol
            Set oSystem = oProfile
            Set oCol = oSystem.GetChildren
            Dim i As Integer
            
            'add profiles to the output collection if they are IJAssemblyChildren
            For i = 1 To oCol.Count
                If TypeOf oCol.Item(i) Is IJAssemblyChild Then
                    'add to collection, first convert to object to be able to extract as object
                    Set OObj = oCol.Item(i)
                    oOutputColl.Add AsIJAssemblyChild(OObj)
                    Set OObj = Nothing
                End If
            Next i
        Next
    End If 'oProfileCol Is Nothing
   
    Set FindSystemProfilesOnPlate = oOutputColl
    'MsgBox "System: " & FindSystemProfilesOnPlate.Count
        
Cleanup:
    Set oStructUtility = Nothing
    Set oSystemChild = Nothing
    Set oProfile = Nothing
    Set oProfileCol = Nothing
    Set oSystem = Nothing
    Set oCol = Nothing
    Set oOutputColl = Nothing
    Set OObj = Nothing
    Set oPlateFromChild = Nothing
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
    GoTo Cleanup
    
End Function



'**************************************************************************
' Method
'   FindProfilesFromLeafPlate
'
' Description:
'   From the passed in LeafPlateSystem gets the profiles which has a StructConnection
'   to the LeafPlate and return in a collection.
'   Only profiles implementing IJAssemblyChild are valid

'   Overview
'   1. Validate it is a LeafPlateSystem and get RootPlateSystem
'   2. From RootPlateSystem get all children and among those find all IJProfileSystems.
'   3. From each profilesystem find, children of type IJStructConnection
'   4. From each IJStructConnection get children of type IJStructConnection (leaf)
'   5. From each of the leaf IJStructConnection get plate and profile
'   6. Compare plate to leafplate found in step 1. If same add profile to collection
'
'
' Arguments:
'   [in] ByVal oPlate As IJPlate, the plate to get profiles from
'
' Return Values:
'   IJDObjectcollection, with valid profiles. If no valid profiles empty collection
'                        is returned.
'**************************************************************************
Public Function FindProfilesFromLeafPlate(ByVal oPlatePart As IMSPlateEntity.IJPlatePart) As Collection
    Const METHOD As String = "FindProfilesFromLeafPlate"
    On Error GoTo ErrorHandler
    
    Dim oRootPlateSyst As IJPlateSystem
    Dim oLeafPlateSyst As IJPlateSystem
    Dim oSystem As IJSystem
    Dim oSysChildren As IJDTargetObjectCol
    Dim oOutputColl As Collection
    Dim OObj As Object
    Dim oProfSystCol As Collection
    Dim oStructConCol As Collection
    Dim oStiffSyst As IJStiffenerSystem
    Dim oStructConChildColl As Collection
    Dim oStructCon As IJStructConnection
    Dim oAppCon As IJAppConnection
    Dim oColPorts As IJElements
    Dim oPort As IJPort
    Dim oPart As Object
    Dim oChild As Object
    Dim oProfileObj As IJProfile
    Dim oPlateObj As IJPlateSystem
    Dim oSysChild As IJSystemChild
    
'    Dim oPlateFromChild As IMSPlateEntity.IJPlate

    Dim lCount As Long
    Set oOutputColl = New Collection
    
    '** 1. Validate that this is a LeafPlateSystem by checking that 2 levels above the
    'PlatePart are PlateSystems. Get the RootPlateSystem
    'Set oSystemChild = oPlatePart
    Set oSysChild = oPlatePart
    If TypeOf oSysChild.GetParent Is IJPlateSystem Then
        Set oLeafPlateSyst = oSysChild.GetParent
    End If
    
    'fix for TR40668 - need to check only if leaf plate system is not nothing
    If Not oLeafPlateSyst Is Nothing Then
        Set oSysChild = oLeafPlateSyst
        If TypeOf oSysChild.GetParent Is IJPlateSystem Then
            Set oRootPlateSyst = oSysChild.GetParent
        End If
    End If
   
    If (oLeafPlateSyst Is Nothing) Or (oRootPlateSyst Is Nothing) Then
        'this is not a leafplatesystem, exit with empty collection
        Set FindProfilesFromLeafPlate = oOutputColl
        GoTo Cleanup
    End If
    
    
    '** 2. get children of RootPlateSystem and sort out the IJProfileSystems
    Set oProfSystCol = New Collection
    Set oSystem = oRootPlateSyst
    Set oSysChildren = oSystem.GetChildren
    For lCount = 1 To oSysChildren.Count
        Set OObj = oSysChildren.Item(lCount)
        If TypeOf OObj Is IJStiffenerSystem Then
            Call oProfSystCol.Add(OObj)
        End If
        Set OObj = Nothing
    Next 'lCount
    Set oSystem = Nothing
    Set oSysChildren = Nothing
    
    
    '** 3. Get systemchildren of the profilesystem and detect IJStructConnections
    Set oStructConCol = New Collection
    For Each oStiffSyst In oProfSystCol
        Set oSystem = oStiffSyst
        Set oSysChildren = oSystem.GetChildren
        For lCount = 1 To oSysChildren.Count
            Set OObj = oSysChildren.Item(lCount)
            If TypeOf OObj Is IJStructConnection Then
                Call oStructConCol.Add(OObj)
            End If
            Set OObj = Nothing
        Next
        Set oSystem = Nothing
        Set oSysChildren = Nothing
    Next
    
    
    '** 4. From each IJStructConnection get children of type IJStructConnection
    Set oStructConChildColl = New Collection
    For Each oStructCon In oStructConCol
        Set oSystem = oStructCon
        Set oSysChildren = oSystem.GetChildren
        For lCount = 1 To oSysChildren.Count
            Set OObj = oSysChildren.Item(lCount)
            If TypeOf OObj Is IJStructConnection Then
                 oStructConChildColl.Add OObj
            End If
            Set OObj = Nothing
        Next 'lCount
        Set oSystem = Nothing
        Set oSysChildren = Nothing
    Next
    
    '** 5. Now we have a collection of leaf StructConnections. Each of them are connected
    'to a profile and a plate. Loop though and if the plate they are connected to is the
    'passed in platesystem, then add the profile to the outputcoll
    For Each OObj In oStructConChildColl
        Set oAppCon = OObj
        oAppCon.enumPorts oColPorts
        'get the profile and the plate from the connection
        For lCount = 1 To oColPorts.Count
            Set oPort = oColPorts.Item(lCount)  'get port
            Set oPart = oPort.Connectable       'get object connected to port
            If TypeOf oPart Is IJProfile Then
                Set oProfileObj = oPart
            ElseIf TypeOf oPart Is IJPlateSystem Then
                Set oPlateObj = oPart
            End If
        Next
        
        '**6. check that there is a profile and a plate and that the plate is the
        'same as the leafplatesystem passed in
        If (Not oProfileObj Is Nothing) And (Not oPlateObj Is Nothing) Then
            If oPlateObj Is oLeafPlateSyst Then
                'need to get the stiffenerpart which is a child of the profile we hold now
                Set oSystem = oProfileObj
                Set oSysChildren = oSystem.GetChildren
                '*** Fix for TR35719 - Venu*****
                For lCount = 1 To oSysChildren.Count
                    'need to get object into collection
                    Set oChild = oSysChildren.Item(lCount)
                    If TypeOf oChild Is IJAssemblyChild Then
                        oOutputColl.Add oChild
                    End If
                    Set oChild = Nothing
                Next
            End If
        End If
        
        Set oSystem = Nothing
        Set oSysChildren = Nothing
        Set oProfileObj = Nothing
        Set oPlateObj = Nothing
    Next
   
    Set FindProfilesFromLeafPlate = oOutputColl
        
Cleanup:
    Set oRootPlateSyst = Nothing
    Set oLeafPlateSyst = Nothing
    Set oSystem = Nothing
    Set oSysChildren = Nothing
    Set oOutputColl = Nothing
    Set OObj = Nothing
    Set oProfSystCol = Nothing
    Set oStructConCol = Nothing
    Set oStiffSyst = Nothing
    Set oStructConChildColl = Nothing
    Set oStructCon = Nothing
    Set oAppCon = Nothing
    Set oColPorts = Nothing
    Set oPort = Nothing
    Set oPart = Nothing
    Set oChild = Nothing
    Set oProfileObj = Nothing
    Set oPlateObj = Nothing
    Set oSysChild = Nothing
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
    GoTo Cleanup
    
End Function







'**************************************************************************
' Method
'   FindDetailedProfilesOnPlate
'
' Description:
'   From the passed in plate gets the profiles on the plate and return in a collection.
'   Only profiles implementing IJAssemblyChild are valid
'
' Arguments:
'   [in] ByVal oPlate As IJPlate, the plate to get profiles from
'
' Return Values:
'   IJDObjectcollection, with valid profiles. If no valid profiles empty collection
'                        is returned.
'**************************************************************************
Public Function FindDetailedProfilesOnPlate(ByVal oPlate As IMSPlateEntity.IJPlate) As Collection
    Const METHOD As String = "FindDetailedProfilesOnPlate"
    On Error GoTo ErrorHandler
    
    Dim oPartSupport As IJPartSupport
    Dim oPlatePartSupport As IJPlatePartSupport
    Dim oPlateDisp As Object
    Dim eConType As ConnectionType
    Dim oConObjects As Collection
    Dim oConns As Collection
    Dim oThisPartPort As Collection
    Dim oOtherPartPort As Collection
    Dim OObj As Object
    Dim oOutputColl As New Collection ' IJDObjectCollection
    Dim oSDProfileWrapper As Object
    Dim oStiffenedPlate As IJPlate
    Dim bIsSystem As Boolean
    
    Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
    'create plate helper object
    Set oPlatePartSupport = New PlatePartSupport
    Set oPartSupport = oPlatePartSupport
    
    
    'get IDispatch from plate object
    Set oPlateDisp = oPlate
    
    'connect helper with this part
    Set oPartSupport.Part = oPlateDisp
    
    'get objects physical connected with this plate
    eConType = ConnectionPhysical
    oPartSupport.GetConnectedObjects eConType, oConObjects, oConns, oThisPartPort, oOtherPartPort
    '   TR51692 : Added the additional check to verify that the connected profiles
    '             stiffen the selected plate.
    If Not oConObjects Is Nothing Then
        For Each OObj In oConObjects
            If (TypeOf OObj Is IJProfile) And (TypeOf OObj Is IJAssemblyChild) Then
                If TypeOf OObj Is IJStiffenerPart Then   'for stiffener profile
                    Set oSDProfileWrapper.object = OObj
                    oSDProfileWrapper.GetStiffenedPlate oStiffenedPlate, bIsSystem
                    'add to collection, first convert to object to be able to extract as object
                    'Set oObj = oCol.Item(i)
                    If oStiffenedPlate Is oPlate Then
                        oOutputColl.Add AsIJAssemblyChild(OObj)
                    End If
                Else   'for beam profile
                    oOutputColl.Add AsIJAssemblyChild(OObj)
                End If
                Set OObj = Nothing
            End If
        Next
    End If

    Set FindDetailedProfilesOnPlate = oOutputColl
    'MsgBox "Detail: " & FindDetailedProfilesOnPlate.Count
    
Cleanup:
    Set oPartSupport = Nothing
    Set oPlatePartSupport = Nothing
    Set oPlateDisp = Nothing
    Set oConObjects = Nothing
    Set oConns = Nothing
    Set oThisPartPort = Nothing
    Set oOtherPartPort = Nothing
    Set OObj = Nothing
    Set oOutputColl = Nothing
    Set oSDProfileWrapper = Nothing
    Set oStiffenedPlate = Nothing
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
    GoTo Cleanup
    
End Function


' ***************************************************************************
'
' Function
'   FindBeamsOnPlate(oPlateSys As IMSPlateEntity.IJPlatePart)As Collection
'
' Abstract
'   [in] the plate part as input
'   [out] Collection of BeamParts
' This method retrieves all the BeamParts (Molded form Beam Systems) associated with the plate
' The BeamSystem and the PlateSystem have a Logical Connection. We make use of this to
' get to the BeamSystem (and then to Beam Part).
' NOTE: StandAlone Beam Parts are identified in the FindDetailedProfilesOnPlate method.
' ***************************************************************************
Private Function FindBeamsOnPlate(ByVal oPlatePart As IMSPlateEntity.IJPlatePart) As Collection

    Const METHOD As String = "FindBeamsOnPlate"
    On Error GoTo ErrorHandler
    
    Dim oAppCon As IJAppConnection
    Dim oObjCon As Object
    Dim oAppConType As IJAppConnectionType
    Dim oColPorts As IJElements
    Dim oPort As IJPort
    Dim oPart As Object
    Dim oPlateObj As IJPlateSystem
    
    Dim oConnectable As IJConnectable
    Dim oStructConnectable As IJStructConnectable
    Dim oAppConnCol As IJDObjectCollection
    Dim oTargetObjCol As IJDTargetObjectCol
    Dim oAssocReln As IJDAssocRelation
    Dim oElemPorts As IJElements
    Dim oObjPort As IJPort
    Dim oBeamSystemsCol As IJDObjectCollection
    Set oBeamSystemsCol = New JObjectCollection
    Dim oBeamPartsCol As Collection
    Set oBeamPartsCol = New Collection
    Set oAppConnCol = New JObjectCollection
    
    Dim oRootPlateSyst As IJPlateSystem
    Dim oLeafPlateSyst As IJPlateSystem
    Dim oSysChild As IJSystemChild
    
    ''''''**** Navigate to the Root Level PlateSystem from PlatePart'''''''''''
    '** Validate that this is a LeafPlateSystem by checking that 2 levels above the
    'PlatePart are PlateSystems. Get the RootPlateSystem
    Set oSysChild = oPlatePart
    If TypeOf oSysChild.GetParent Is IJPlateSystem Then
        Set oLeafPlateSyst = oSysChild.GetParent
    End If
        
    If Not oLeafPlateSyst Is Nothing Then
        Set oSysChild = oLeafPlateSyst
        If TypeOf oSysChild.GetParent Is IJPlateSystem Then
            Set oRootPlateSyst = oSysChild.GetParent
        End If
    End If
   
    If (oLeafPlateSyst Is Nothing) Or (oRootPlateSyst Is Nothing) Then
        'this is not a leafplatesystem, exit with empty collection
        Set FindBeamsOnPlate = oBeamPartsCol
        GoTo Cleanup
    End If
    
    '''''1.Get the Ports from the Plate System
    Set oStructConnectable = oRootPlateSyst
    If Not oStructConnectable Is Nothing Then
        Set oColPorts = oStructConnectable.enumAllPorts
    End If
    Set oStructConnectable = Nothing
    
    If oColPorts.Count > 0 Then
        For Each oPort In oColPorts
            '''''2.Get the AppConns from each Port
            Set oAssocReln = oPort
            Set oTargetObjCol = oAssocReln.CollectionRelations(IID_IJPort, "ConnHasPorts_DEST")
            If oTargetObjCol.Count = 1 Then
                Set oAppCon = oTargetObjCol.Item(1)
                oAppConnCol.Add oAppCon
            End If
            Set oAssocReln = Nothing
            Set oTargetObjCol = Nothing
            Set oAppCon = Nothing
        Set oPort = Nothing
        Next
    End If
    
    
    If oAppConnCol.Count > 0 Then
    '''''''3. For each App Connection get the Ports
        For Each oAppCon In oAppConnCol
            oAppCon.enumPorts oElemPorts
            If oElemPorts.Count > 0 Then
            ''''''''4. For each Port get the connectable
                For Each oObjPort In oElemPorts
                    Set oConnectable = oObjPort.Connectable
                    ''''5. QI for IJBeamSystem and add if supported
                    If TypeOf oConnectable Is IJBeamSystem Then
                        oBeamSystemsCol.Add oConnectable
                    End If
                    Set oConnectable = Nothing
                    Set oObjPort = Nothing
                Next
            End If
            Set oElemPorts = Nothing
            Set oObjPort = Nothing
            Set oAppCon = Nothing
        Next
    End If
    
    Dim OObj As Object
    Dim oChild As Object
    Dim oChildCol As IJDObjectCollection
    Dim oBeamObj As IJBeamPart
    
    If Not oBeamSystemsCol Is Nothing And oBeamSystemsCol.Count > 0 Then
        For Each OObj In oBeamSystemsCol
            Set oBeamObj = GetBeamPartFromBeamSys(OObj)
            If Not oBeamObj Is Nothing Then
                oBeamPartsCol.Add oBeamObj
            End If
            Set oBeamObj = Nothing
            Set OObj = Nothing
        Next
    End If
    Set FindBeamsOnPlate = oBeamPartsCol

Cleanup:
    Set oAppCon = Nothing
    Set oConnectable = Nothing
    Set oStructConnectable = Nothing
    If oAppConnCol.Count > 0 Then
        oAppConnCol.Clear
    End If
    Set oAppConnCol = Nothing
    Set oRootPlateSyst = Nothing
    Set oLeafPlateSyst = Nothing
    Set oSysChild = Nothing
    
Exit Function
ErrorHandler:
    MsgBox Err.Description
    GoTo Cleanup
    
End Function

'**************************************************************************
' Method
'   ValidateProfiles
'
' Description:
'   A profile must satisfy the following condition to be valid for new assembly for
'   adding by the "include profiles" option
'   - profile parent must be FailedParts, UnAssignedParts or none
'   - or profile must be in same assembly as selected plate
'
' Arguments:
'   ByVal oPlateParent As IJAssembly,   the parent of the selected plate
'   ByVal oProfColl As Collection,      the collection of profiles to check
'
' Return Values:
'   A collection of valid profiles. Collection is empty if no valid
'
'**************************************************************************
Public Function ValidateProfiles(ByVal oPlateparent As IJAssembly, ByVal oProfColl As Collection) As Collection
    Const METHOD = "IsInRealAssembly"
    On Error GoTo ErrorHandler

    Dim oProfAssyChild As IJAssemblyChild
    Dim oAssy As IJAssembly
    Dim bProfileValid As Boolean
    Dim oOutputColl As Collection
    
    'create a collection to return
    Set oOutputColl = New Collection
    
    For Each oProfAssyChild In oProfColl
        bProfileValid = True
        'get parent from passed in
        Set oAssy = oProfAssyChild.Parent
        
        'did we get a valid parent (if part is below ship oAssy could be nothing)
        If (Not oAssy Is Nothing) And (Not oPlateparent Is Nothing) Then
            'check if parent is unassigned or failedparts
            If (Not TypeOf oAssy Is IJPlnUnprocessedParts) And (Not TypeOf oAssy Is IJPlnFailedParts) Then
                'check if profile parent assy is same as plate parent assy
                If Not oPlateparent Is oAssy Then
                    'parent is a normal assembly and not the same as plate, set profile false
                    bProfileValid = False
                End If
            End If
        End If
                    
        If bProfileValid = True Then oOutputColl.Add oProfAssyChild
        Set oProfAssyChild = Nothing
        Set oAssy = Nothing
    Next
    
    Set ValidateProfiles = oOutputColl
    
Cleanup:
    Set oProfAssyChild = Nothing
    Set oPlateparent = Nothing
    Set oAssy = Nothing
    Set oOutputColl = Nothing
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
    GoTo Cleanup
    
End Function
 Public Function AsIJAssemblyChild(ByVal Entity As Object) As IJAssemblyChild
    Set AsIJAssemblyChild = Entity
End Function

' ***************************************************************************
'
' Function
'   GetBeamPartFromBeamSys(ByVal oTopLevelBeamSys As IJBeamSystem) As IJBeamPart
'
' Abstract
' Gets the BeamPart from the BeamSystem
' As the immediate thickening is ON, we get to the Middle Level BeamSystem from the Top
' Level BeamSYstem and from the middle level BeamSystem to the Part.
' ***************************************************************************
Private Function GetBeamPartFromBeamSys(ByVal oTopLevelBeamSys As IJBeamSystem) As IJBeamPart
    Dim oLoop As Object
    Dim oDesParent As IJDesignParent
    Dim oTemp As Object
    Dim oChildcoll As IJDObjectCollection
    Set oChildcoll = New JObjectCollection
    Dim oMiddleLevelBeamSys As IJBeamSystem
    Dim oMiddleLevelBeamSysCol As IJDObjectCollection
        
        If TypeOf oTopLevelBeamSys Is IJBeamSystem Then
            Set oDesParent = oTopLevelBeamSys
            oDesParent.GetChildren oChildcoll
            If oChildcoll.Count <> 0 Then
                For Each oTemp In oChildcoll
                    If TypeOf oTemp Is IJBeamSystem Then
                        Set oMiddleLevelBeamSys = oTemp
                    End If
                    Set oTemp = Nothing
                Next
            End If
            Set oDesParent = Nothing
            Set oChildcoll = Nothing
        End If

        Set oTemp = Nothing
        Set oMiddleLevelBeamSysCol = New JObjectCollection
        Set oDesParent = Nothing
        Set oDesParent = oMiddleLevelBeamSys
        oDesParent.GetChildren oMiddleLevelBeamSysCol
        For Each oTemp In oMiddleLevelBeamSysCol
            If TypeOf oTemp Is IJBeamPart Then
                Set GetBeamPartFromBeamSys = oTemp
                GoTo Cleanup
            End If
        Next
Cleanup:
    Set oTemp = Nothing
    Set oMiddleLevelBeamSys = Nothing
    Set oMiddleLevelBeamSysCol = Nothing
    Set oDesParent = Nothing
    Set oChildcoll = Nothing
    Exit Function
End Function

