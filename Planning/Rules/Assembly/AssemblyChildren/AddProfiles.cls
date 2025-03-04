VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AddProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULE = "PlnAssyChildrenRules.AddProfiles"
Private Const IID_IJPort As String = "{5CF7C404-546D-11D2-B328-080036024603}"

'For adding all the profiles on Plates, Set this boolean to TRUE
Private Const AddAllParts As Boolean = False

Implements IJAssemblyChildrenRules


Private Function IJAssemblyChildrenRules_GetSecondaryObjects(ByVal pChildren As IJElements) As IJElements
Const METHOD As String = "IJAssemblyChildrenRules_GetSecondaryObjects"
On Error GoTo ErrorHelper

    Dim oObject             As Object
    Dim oProfiles           As IJElements
    Dim oAsChild            As IJAssemblyChild
    Dim oAssembly           As IJAssembly
    Dim oAllProfColl        As Collection
    Dim oSDOProfilePart     As StructDetailObjects.ProfilePart
    Dim dAcceptableratio    As Double
    Dim dProfLength         As Double
    Dim dCommonLength       As Double
    Dim i                   As Long
    Dim lCount              As Long
    Dim oProfilePart        As IJProfilePart
    Dim oResult             As Object
        
    Set oProfiles = New JObjectCollection
    
    For Each oObject In pChildren
    
        If (TypeOf oObject Is IJPlatePart) And (Not TypeOf oObject Is IJCollarPart) Then
            
            Set oAsChild = oObject
            Set oAssembly = oAsChild.Parent
            
            Set oAllProfColl = GetValidProfilesOnPlate(oAsChild, oAssembly)
            
            If Not oAllProfColl Is Nothing Then
                lCount = oAllProfColl.Count
            End If
            
            If AddAllParts = False Then
                If lCount > 0 Then
                
                    Set oSDOProfilePart = New StructDetailObjects.ProfilePart
                    
                    For i = 1 To lCount
                        Set oSDOProfilePart.object = oAllProfColl.Item(i)
                        Set oProfilePart = oAllProfColl.Item(i)
                        
                        dCommonLength = GetProfileLengthOnThePlate(oProfilePart, oAsChild)
                        dProfLength = oSDOProfilePart.MoldedLength
    
                        If (dCommonLength / dProfLength) > 0.5 Then
                            oProfiles.Add oAllProfColl.Item(i)
                        End If
                    Next
                End If
            Else
                For Each oResult In oAllProfColl
                    oProfiles.Add oResult
                Next
            End If
        End If
    Next
    
    Set IJAssemblyChildrenRules_GetSecondaryObjects = oProfiles
    
    Set oObject = Nothing
    Set oProfiles = Nothing
    Set oAsChild = Nothing
    Set oAssembly = Nothing
    Set oAllProfColl = Nothing
    Set oSDOProfilePart = Nothing
                
Exit Function
ErrorHelper:
    MsgBox Err.Description
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
    Dim oOutputColl As New Collection
    Dim oObj As Object
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
            Dim oAsChild As IJAssemblyChild
            
            'add profiles to the output collection if they are IJAssemblyChildren
            For i = 1 To oCol.Count
                If TypeOf oCol.Item(i) Is GSCADAsmHlpers.IJAssemblyChild Then
        
                    'add to collection, first convert to object to be able to extract as object
                    Set oObj = oCol.Item(i)
                    Set oAsChild = oObj
                    
                    oOutputColl.Add oAsChild
                    Set oObj = Nothing
                End If
            Next i
            Set oProfile = Nothing
        Next
    End If

    Set FindSystemProfilesOnPlate = oOutputColl
        
CleanUp:
    Set oStructUtility = Nothing
    Set oSystemChild = Nothing
    Set oProfile = Nothing
    Set oProfileCol = Nothing
    Set oSystem = Nothing
    Set oCol = Nothing
    Set oOutputColl = Nothing
    Set oObj = Nothing
    Set oPlateFromChild = Nothing
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
    GoTo CleanUp
    
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
    Dim oObj As Object
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
    
    Dim lCount As Long
    Set oOutputColl = New Collection
    
    '** 1. Validate that this is a LeafPlateSystem by checking that 2 levels above the
    'PlatePart are PlateSystems. Get the RootPlateSystem
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
        GoTo CleanUp
    End If
    
    
    '** 2. get children of RootPlateSystem and sort out the IJProfileSystems
    Set oProfSystCol = New Collection
    Set oSystem = oRootPlateSyst
    Set oSysChildren = oSystem.GetChildren
    For lCount = 1 To oSysChildren.Count
        Set oObj = oSysChildren.Item(lCount)
        If TypeOf oObj Is IJStiffenerSystem Then
            Call oProfSystCol.Add(oObj)
        End If
        Set oObj = Nothing
    Next 'lCount
    Set oSystem = Nothing
    Set oSysChildren = Nothing
    
    
    '** 3. Get systemchildren of the profilesystem and detect IJStructConnections
    Set oStructConCol = New Collection
    For Each oStiffSyst In oProfSystCol
        Set oSystem = oStiffSyst
        Set oSysChildren = oSystem.GetChildren
        For lCount = 1 To oSysChildren.Count
            Set oObj = oSysChildren.Item(lCount)
            If TypeOf oObj Is IJStructConnection Then
                Call oStructConCol.Add(oObj)
            End If
            Set oObj = Nothing
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
            Set oObj = oSysChildren.Item(lCount)
            If TypeOf oObj Is IJStructConnection Then
                 oStructConChildColl.Add oObj
            End If
            Set oObj = Nothing
        Next 'lCount
        Set oSystem = Nothing
        Set oSysChildren = Nothing
    Next
    
    '** 5. Now we have a collection of leaf StructConnections. Each of them are connected
    'to a profile and a plate. Loop though and if the plate they are connected to is the
    'passed in platesystem, then add the profile to the outputcoll
    For Each oObj In oStructConChildColl
        Set oAppCon = oObj
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
        
CleanUp:
    Set oRootPlateSyst = Nothing
    Set oLeafPlateSyst = Nothing
    Set oSystem = Nothing
    Set oSysChildren = Nothing
    Set oOutputColl = Nothing
    Set oObj = Nothing
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
    GoTo CleanUp
    
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
    Dim oObj As Object
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
    
    Dim oAsChild As IJAssemblyChild
    
    If Not oConObjects Is Nothing Then
        For Each oObj In oConObjects
            If (TypeOf oObj Is IJProfile) And (TypeOf oObj Is IJAssemblyChild) Then
            
                Set oAsChild = oObj
                
                If TypeOf oObj Is IJStiffenerPart Then   'for stiffener profile
                    Set oSDProfileWrapper.object = oObj
                    oSDProfileWrapper.GetStiffenedPlate oStiffenedPlate, bIsSystem
                    'add to collection, first convert to object to be able to extract as object

                    If oStiffenedPlate Is oPlate Then
                        oOutputColl.Add oAsChild
                    End If
                Else   'for beam profile
                    oOutputColl.Add oAsChild
                End If
                Set oObj = Nothing
            End If
        Next
    End If

    Set FindDetailedProfilesOnPlate = oOutputColl
    'MsgBox "Detail: " & FindDetailedProfilesOnPlate.Count
    
CleanUp:
    Set oPartSupport = Nothing
    Set oPlatePartSupport = Nothing
    Set oPlateDisp = Nothing
    Set oConObjects = Nothing
    Set oConns = Nothing
    Set oThisPartPort = Nothing
    Set oOtherPartPort = Nothing
    Set oObj = Nothing
    Set oOutputColl = Nothing
    Set oSDProfileWrapper = Nothing
    Set oStiffenedPlate = Nothing
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
    GoTo CleanUp
    
End Function

    
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
        
    Set GetValidProfilesOnPlate = oProfiles

CleanUp:
    Set oTempProfiles = Nothing
    Set oProfiles = Nothing
    Set oValidProfiles = Nothing
    Set oProfile = Nothing
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
    GoTo CleanUp
    
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
        GoTo CleanUp
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
    
    Dim oObj As Object
    Dim oChild As Object
    Dim oChildCol As IJDObjectCollection
    Dim oBeamObj As IJBeamPart
    
    If Not oBeamSystemsCol Is Nothing And oBeamSystemsCol.Count > 0 Then
        For Each oObj In oBeamSystemsCol
            Set oBeamObj = GetBeamPartFromBeamSys(oObj)
            If Not oBeamObj Is Nothing Then
                oBeamPartsCol.Add oBeamObj
            End If
            Set oBeamObj = Nothing
            Set oObj = Nothing
        Next
    End If
    Set FindBeamsOnPlate = oBeamPartsCol

CleanUp:
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
    GoTo CleanUp
    
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
                GoTo CleanUp
            End If
        Next
CleanUp:
    Set oTemp = Nothing
    Set oMiddleLevelBeamSys = Nothing
    Set oMiddleLevelBeamSysCol = Nothing
    Set oDesParent = Nothing
    Set oChildcoll = Nothing
    Exit Function
End Function

Private Function GetProfileLengthOnThePlate(oProfilePart As IJProfilePart, oPlatePart As IJPlatePart) As Double
 Const METHOD As String = "GetProfileLengthOnThePlate"
 On Error GoTo ErrorHelper
 
    Dim oPlatePartGeomState         As IJPartGeometryState
    Dim oProfilePartGeomState       As IJPartGeometryState
    Dim oSDOProfilePart             As StructDetailObjects.ProfilePart
    Dim oSDOPlatePart               As StructDetailObjects.PlatePart
    Dim dStiffeningLength           As Double

    Set oProfilePartGeomState = oProfilePart
    Set oPlatePartGeomState = oPlatePart

    If oPlatePartGeomState.PartGeometryState = DetailedPart And _
        oProfilePartGeomState.PartGeometryState = DetailedPart Then
        
        Set oSDOProfilePart = New StructDetailObjects.ProfilePart
        Set oSDOProfilePart.object = oProfilePart

        dStiffeningLength = oSDOProfilePart.ProfileLengthOnThePlate(oPlatePart)

    Else
'        If oPlatePart.Curved = PLATE_CURVATURE_Flat Then
'            'Use Gtypes to get the length
'        Else
            dStiffeningLength = ProjectWireAndGetLength(oProfilePart, oPlatePart)
'        End If
        
    End If
    
    GetProfileLengthOnThePlate = dStiffeningLength
    
    Set oPlatePartGeomState = Nothing
    Set oProfilePartGeomState = Nothing
    Set oSDOProfilePart = Nothing
    Set oSDOPlatePart = Nothing

    Exit Function
ErrorHelper:
    MsgBox Err.Description
End Function

Private Function ProjectWireAndGetLength(oProfilePart As IJProfilePart, oPlatePart As IJPlatePart) As Double
 Const METHOD As String = "ProjectWireAndGetLength"
 On Error GoTo ErrorHelper
    
    Dim oSDOProfilePart     As StructDetailObjects.ProfilePart
    Dim oSDOPlatePart       As StructDetailObjects.PlatePart
    Dim oLandCurve          As IJWireBody
    Dim oThicknessDir       As IJDVector
    Dim bCentered           As Boolean
    Dim oPlateport          As IJPort
    Dim oPortGeom           As Object
    Dim oProjectUtil        As IMSModelGeomOps.IProject
    Dim oLandCrvProj        As IJDModelBody
    Dim dReqAccuracy        As Double
    Dim dAchievedAcc        As Double
    Dim dLength             As Double
    Dim dArea               As Double
    Dim dVolume             As Double
    
    
    ProjectWireAndGetLength = 0
    
    Set oSDOProfilePart = New StructDetailObjects.ProfilePart
    Set oSDOProfilePart.object = oProfilePart
    
    oSDOProfilePart.LandingCurve oLandCurve, oThicknessDir, bCentered

    Set oSDOPlatePart = New StructDetailObjects.PlatePart
    Set oSDOPlatePart.object = oPlatePart
    
    Set oPlateport = oSDOPlatePart.BasePort(BPT_Base)
    Set oPortGeom = oPlateport.Geometry
    
    Set oProjectUtil = New IMSModelGeomOps.Project
    
    If Not oLandCurve Is Nothing Then

        oProjectUtil.CurveAlongVectorOnToSurface Nothing, _
                                                 oPortGeom, _
                                                 oLandCurve, _
                                                 Nothing, _
                                                 Nothing, _
                                                 oLandCrvProj
    End If

    If Not oLandCrvProj Is Nothing Then
        oLandCrvProj.GetDimMetrics dReqAccuracy, dAchievedAcc, dLength, dArea, dVolume
        ProjectWireAndGetLength = dLength
    End If
    
    Exit Function
ErrorHelper:
    MsgBox Err.Description
End Function
