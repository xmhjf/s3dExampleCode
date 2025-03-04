VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssignmentMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************
'  Copyright (C) 2008, Intergraph Corporation. All rights reserved.
'
'  Project     : Planning\Data\Rules\BlockAssignment\
'  File        : AssignmentMethods.cls
'
'  Description : Block Assignment part validation rules
'
'  History     :
'   16th July 2008      Kishore     Initial creation
'   1st Sep 2008        Kishore     TR-CP-148705    Block assignment is not working for spools
'   9th Sep 2008        Kishore     TR-CP-147526    Block assignment service is incorrect if parts
'                                                   are slanted to cutting surface
'   6th Nov 2008        Kishore     DI-CP-149905  Provide an option to process parts by Minimum Bounding Box
'   9th Sep 2008        Kishore     TR-CP-147526    Block assignment service is incorrect if parts
'                                                   are slanted to cutting surface
'   20th Nov 2008       Kishore     TR-CP-154585    Invalid Log file processing when IJBlockAssnRules_IsCandidateObject
'                                                   fails
'   7th Jul 2009        Apparao     TR-CP-167031  Intersections for the selected parts are not listed in Manage Block Intersectio
'**************************************************************************************

Option Explicit

Implements IJBlockAssnRules
Private Const IID_IJAssemblyChild As String = "{B447C9B4-FB74-11D1-8A49-00A0C9065DF6}"
Private Const strTracePath = "\HKEY_LOCAL_MACHINE\SOFTWARE\Intergraph\Applications\Environments\Planning\Debug"


'Checks whether the part can be processed
Private Function IJBlockAssnRules_CanObjectBeAddedToBucket(ByVal oPart As Object) As Boolean
On Error GoTo ErrorHandler

    Dim oPlnTimer       As New PlnTimer
    Dim oParent         As Object
    Dim oIJDObject      As IJDObject
    Dim oAsChild        As IJAssemblyChild
    Dim strReason       As String
    Dim oNamedItem      As IJNamedItem
    Dim oIJDPOM         As IJDPOM
    Dim oMoniker        As IUnknown
    Dim oPlnIntHelper   As IJDPlnIntHelper
    
    IJBlockAssnRules_CanObjectBeAddedToBucket = False
    
    Set oIJDObject = oPart
    'Check approval status and acess rights on the part
    If (oIJDObject.ApprovalStatus <> Working) Or _
        (oIJDObject.AccessControl And acUpdate) <> acUpdate Then
        
        oPlnTimer.PrintObjectDetails "BUCKET_IGNORED [Part's status]:", oPart
        Exit Function
    End If
        
    Set oIJDPOM = oIJDObject.ResourceManager
    Set oMoniker = oIJDPOM.GetObjectMoniker(oPart)
    
    'Part not assignable if it's an Assembly,
    'Planning Joint or Foundation Component
    If (False = oIJDPOM.SupportsInterface(oMoniker, "IJNamedItem")) Then
        oPlnTimer.PrintObjectDetails "BUCKET_IGNORED [No Name]:", oPart
        
    ElseIf (True = oIJDPOM.SupportsInterface(oMoniker, "IJAssembly")) Then
        'Structural System can be assigned
        If (True = oIJDPOM.SupportsInterface(oMoniker, "IJStructuralSystem")) Then
            IJBlockAssnRules_CanObjectBeAddedToBucket = True
        ElseIf (True = oIJDPOM.SupportsInterface(oMoniker, "IJDSpool")) Then
        
            Set oPlnIntHelper = New CPlnIntHelper
            If oPlnIntHelper.GetStoredProcAssemblyChildren(oPart, "IJAssemblyChild", False, Nothing, False).Count > 0 Then
                GoTo CheckParent
            End If
        Else
            'Part is an Assembly, hence not assignable.
            oPlnTimer.PrintObjectDetails "BUCKET_IGNORED [Assembly]:", oPart
        End If
            
    ElseIf (True = oIJDPOM.SupportsInterface(oMoniker, "IJPlnJoint")) Then
        oPlnTimer.PrintObjectDetails "BUCKET_IGNORED [Planning Joint]:", oPart

    ElseIf (True = oIJDPOM.SupportsInterface(oMoniker, "ISPSFoundationComponent")) Then
        'Part is a foundation component, hence not assignable.
        oPlnTimer.PrintObjectDetails "BUCKET_IGNORED [Foundation. Component]:", oPart
        
    Else
    
CheckParent:
        Set oAsChild = oPart
        Set oParent = oAsChild.Parent
        
        'Part has a parent
        If Not oParent Is Nothing Then
            
            'check approval status and acess rights on the part's parent
            Set oIJDObject = oParent
            If (oIJDObject.ApprovalStatus <> Working) Or _
                (oIJDObject.AccessControl And acUpdate) <> acUpdate Then
                
                oPlnTimer.PrintObjectDetails "BUCKET_IGNORED [Parent's status]:", oIJDObject
                
            'Part assignable only when its parent is Failed Parts folder or,
            'Unassigned parts folder or Project root.
            ElseIf TypeOf oParent Is IJPlnFailedParts Or _
                TypeOf oParent Is IJPlnUnprocessedParts Or _
                    TypeOf oParent Is IJConfigProjectRoot Then
                    
                IJBlockAssnRules_CanObjectBeAddedToBucket = True
            Else
                strReason = "BUCKET_IGNORED "
                strReason = strReason + "[Vaild parent - "
                
                Set oNamedItem = oParent
                strReason = strReason + oNamedItem.Name
                strReason = strReason + "]:"
                
                oPlnTimer.PrintObjectDetails strReason, oPart
            End If
        Else
            'Part has no parent, hence assignable
            IJBlockAssnRules_CanObjectBeAddedToBucket = True
        End If
    End If
    
wrapup:
    Set oPlnIntHelper = Nothing
    Set oIJDPOM = Nothing
    Set oMoniker = Nothing
    Set oNamedItem = Nothing
    Set oAsChild = Nothing
    Set oPart = Nothing
    Set oIJDObject = Nothing

Exit Function
ErrorHandler:
    IJBlockAssnRules_CanObjectBeAddedToBucket = False
    oPlnTimer.PrintObjectDetails "BUCKET_IGNORED (UnExpected Error)", oPart
    GoTo wrapup
End Function

'checks whether the part can be assigned and finds out appropriate assignment method.
'When the assignment has to take place by Range it returns the overlap ratio to be used
'to consider a Block for creating intersection object.
Private Function IJBlockAssnRules_IsCandidateObject(ByVal oPart As Object, eAssignMethod As PlnBlkAssnMethods, dOverlapFactor As Double) As Boolean
On Error GoTo ErrorHandler

    Dim oPlnTimer                   As New PlnTimer
    Dim bPlnSeamsPresent            As Boolean
    Dim varOverlap                  As Variant
    Dim oRegistry                   As IJRegistry
    Dim varValue                    As Variant
    
    Set oRegistry = New Registry
    oPlnTimer.InitTimer "IJBlockAssnRules_IsCandidateObject"
    oPlnTimer.StartTimer "IJBlockAssnRules_IsCandidateObject"
    
    If IsPartAssignable(oPart) = False Then
        IJBlockAssnRules_IsCandidateObject = False
        
        oPlnTimer.StopTimer
        oPlnTimer.ExitTimer
        Exit Function
    Else
        oPlnTimer.PrintInfo "Part is eligible for processing."
    End If
    
    bPlnSeamsPresent = False
    eAssignMethod = ByRangeOrGeometry(oPart, bPlnSeamsPresent)
    
    On Error Resume Next
    varValue = oRegistry.GetValue(strTracePath, "OverLappingValue")
    
    Err.Clear
    On Error GoTo ErrorHandler
    
    If Not varValue = vbNullString Then
        dOverlapFactor = varValue
    Else
        If eAssignMethod = PlnAssignByMinRange Then
            If bPlnSeamsPresent Then
                dOverlapFactor = 0.8
            Else
                dOverlapFactor = 0.95
            End If
        ElseIf eAssignMethod = PlnAssignByRange Then
            dOverlapFactor = 0.95
        Else
            dOverlapFactor = 1#
        End If
    End If
    
    oPlnTimer.StopTimer
    oPlnTimer.ExitTimer
    IJBlockAssnRules_IsCandidateObject = True

    Set oRegistry = Nothing
Exit Function
ErrorHandler:
    Err.Clear
    IJBlockAssnRules_IsCandidateObject = False
    oPlnTimer.PrintInfo "Unexpected error, ignoring part."
End Function

Private Function ByRangeOrGeometry(oPart As Object, bPlnSeamsPresent As Boolean) As PlnBlkAssnMethods
    
    Dim oHierarchyHlper             As SASSIGNMENTENGINELib.HierarchyHlper
    Dim ePlnApplicationDomains      As PlnApplicationDomains
    Dim oDesignChild                As IJDesignChild
    Dim oParentEntity               As Object
    
    Set oHierarchyHlper = New SASSIGNMENTENGINELib.HierarchyHlper
    ByRangeOrGeometry = PlnAssignByRange
    bPlnSeamsPresent = False

    ' Only allow split for structure.
    ePlnApplicationDomains = oHierarchyHlper.GetApplicationDomain(oPart)
    
    If ePlnApplicationDomains = PlnDomainStructure Then
        
        ByRangeOrGeometry = PlnAssignByGeometry
        
        If IsPartCreatedByPlanningSeam(oPart) Then
            ByRangeOrGeometry = PlnAssignBySmallVol
            bPlnSeamsPresent = True
        End If

    ElseIf ePlnApplicationDomains = PlnDomainOutfitting Then
    
        ' Get system parent from entity
        If TypeOf oPart Is IJDesignChild Then
            Set oDesignChild = oPart
            Set oParentEntity = oDesignChild.GetParent
        End If
    
        If Not oParentEntity Is Nothing Then
            If TypeOf oParentEntity Is IJRouteSplit Then
                ByRangeOrGeometry = PlnAssignByGeometry
            End If
        End If
    Else
        ByRangeOrGeometry = PlnAssignByRange
    End If
    
    ' Clean up

    Set oHierarchyHlper = Nothing
    Set oDesignChild = Nothing
    Set oParentEntity = Nothing

    Exit Function
End Function

Private Function IsPartAssignable(oPart As IJAssemblyChild) As Boolean

    Dim oParent         As Object
    Dim oIJDObject      As IJDObject
    Dim oPlnTimer       As New PlnTimer
    
    IsPartAssignable = False
    
    Set oIJDObject = oPart
    'Check approval status and acess rights on the part
    If (oIJDObject.ApprovalStatus <> Working) Or _
        (oIJDObject.AccessControl And acUpdate) <> acUpdate Then
        
        oPlnTimer.PrintInfo "Part is not in working status."
        Exit Function
    End If
    
    Set oParent = oPart.Parent
    
    'Part has a parent
    If Not oParent Is Nothing Then
        
        'check approval status and acess rights on the part's parent
        Set oIJDObject = oParent
        If (oIJDObject.ApprovalStatus <> Working) Or _
            (oIJDObject.AccessControl And acUpdate) <> acUpdate Then
            
            oPlnTimer.PrintInfo "Part's parent is not in working status."
            Exit Function
            
        'Part assignable only when its parent is Failed Parts folder or,
        'Unassigned parts folder or Project root.
        ElseIf TypeOf oParent Is IJPlnFailedParts Or _
            TypeOf oParent Is IJPlnUnprocessedParts Or _
                TypeOf oParent Is IJConfigProjectRoot Then
                
            IsPartAssignable = True
        Else
        
            oPlnTimer.PrintInfo "Part is already under an Assembly."
            IsPartAssignable = False
        End If
    Else
        'Part has no parent, hence assignable
        IsPartAssignable = True
    End If

End Function

Private Function IsPartCreatedByPlanningSeam(oPart As Object) As Boolean
On Error GoTo ErrorHandler

    Dim oAssocRel           As IJDAssocRelation
    Dim oTargetObjCol       As IJDTargetObjectCol
    Dim oIntersection       As IJIntersection
    Dim i                   As Long

    Set oAssocRel = oPart
    
    IsPartCreatedByPlanningSeam = False
    
    Set oTargetObjCol = oAssocRel.CollectionRelations(IID_IJAssemblyChild, "IntersectedObject_DEST")
    
    If oTargetObjCol.Count > 0 Then 'Means Entity is split at intersection and Seam is present
        For i = 1 To oTargetObjCol.Count
            Set oIntersection = oTargetObjCol.Item(i)
            If oIntersection.Status = PlnStatusType_Seam Then
                IsPartCreatedByPlanningSeam = True
                GoTo wrapup
            End If
        Next
    End If
    

wrapup:

    Set oTargetObjCol = Nothing
    Set oIntersection = Nothing
    Set oAssocRel = Nothing

    Exit Function
    
ErrorHandler:
Err.Clear
    IsPartCreatedByPlanningSeam = False
End Function
