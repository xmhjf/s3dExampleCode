VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SDStiffNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2001-2003 Global Research and Development, Inc.
'  All rights reserved.
'
'  Project: SDProfilePartNameRules
'
'  Abstract: The file contains name rules for Stiffener Profile Parts
'
'  NOTE: This uses the parent system to build the system O.name portion
'  of the part name, but does not use the parent system name directly.
'  It is assumed that the information used to build the system O.name portion
'  of the name IS part of the system name, but as the system name is not used
'  directly then no assumptions are made based on system name format.
'
'******************************************************************
Option Explicit

Implements IJNameRule
                                                        
Private Const MODULE = "SDProfilePartNameRules.SDStiffNameRule."
Private Const NO_BLOCK = "NoBlk"
Private Const NOSYMM_PORT_ADDON = 10000
Private Const NOSYMM_STBD_ADDON = 20000
Private Const NOSYMM_CENTER_ADDON = 30000
Private Const CL_TOLERANCE = 0.1
Private Const NEG_CL_TOLERANCE = -0.1

Private Enum SDPrinProfileAxis 'X Y profile on plate
    xdir = 0
    ydir = 1
    zdir = 2
    other = 3
End Enum

Private Enum SDSymPriority
    Neither = 0
    Port = 1
    Starboard = 2
    Center = 3
    Both = 4
End Enum

Private m_oErrors As IJEditErrors
Private m_oNameUtil As SDNameRulesUtilHelper
Private m_oNameHelp As SDNameRuleHelper
Private m_oSDOHelper As StructDetailObjects.Helper

Private Sub Class_Initialize()
    On Error GoTo label

    Set m_oNameUtil = New GSCADSDNameRulesUtil.SDNameRulesUtilHelper
    Set m_oNameHelp = New SDNameRulesHelper.SDNameRuleHelper
    Set m_oErrors = New IMSErrorLog.JServerErrors
    Set m_oSDOHelper = New StructDetailObjects.Helper
Exit Sub

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "Class_Initialize", Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
    Set m_oNameUtil = Nothing
    Set m_oNameHelp = Nothing
    Set m_oSDOHelper = Nothing
End Sub

'********************************************************************
' Function:
'   IJNameRule_ComputeName
' Description:
'   Refer to Naming.doc command spec for naming rules
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oEntity As Object, _
                                   ByVal oParents As IJElements, _
                                   ByVal oActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    Dim oNamedItem As IJNamedItem 'stores name
    Dim oStiffPart As StructDetailObjects.ProfilePart 'wrapper
    Dim oResourceMgr As IUnknown
    Dim oProfileParent As IJSystem
    Dim oPlatePart As IJPlate
    Dim oGroupMonitor As IJGroupMon
    Dim iPos As Long
    Dim strLocationID As String
    On Error GoTo errLabel
    
    If oEntity Is Nothing Then 'error, no entity to name
        Err.Raise E_INVALIDARG, MODULE, METHOD
    Else 'entity valid, set up wrapper and name interface
        Set oStiffPart = New StructDetailObjects.ProfilePart
        Set oStiffPart.object = oEntity
        Set oNamedItem = oEntity
    End If
    
    Call SetEnviron(oParents, oProfileParent, oPlatePart, oGroupMonitor)
    Set oResourceMgr = m_oSDOHelper.GetResourceManagerFromObject(oEntity)
        
    If (oGroupMonitor Is Nothing) Then 'no position index, use counter
        Dim eStiffType As StructProfileType
        Dim strPlatePartName As String
        Dim strProfileSystemOname As String
        Dim strNamingParents As String 'detect if name change has occured
        Dim strName As String
        Dim strPos As String 'indexing result
        Dim strProfileCat As String 'profile catagory

        'use index naming fields to construct name
        Call GetNonpositionNamingFields(oStiffPart, oProfileParent, _
                                        oPlatePart, strPlatePartName, _
                                        strProfileSystemOname)
        strProfileCat = GetProfileCategory(oStiffPart, oResourceMgr)
        eStiffType = oStiffPart.ProfileType
        
        'edge reinforcement naming rule is simpler, just use category
        If (eStiffType = sptEdgeReinforcement) Then
            Dim strBlkName As String
            
            strBlkName = GetBlockName(oEntity)
            strNamingParents = strBlkName + strProfileCat
            
            'only create the name and get next index if name will be updated
            If (strNamingParents <> oActiveEntity.NamingParentsString) Then
                Call m_oNameHelp.SDIndexByCounter(strNamingParents, _
                                                  iPos, _
                                                  oResourceMgr)
                strName = "<" + strBlkName + ">-" + _
                          strProfileCat + "." + Format(iPos)
            End If
        ElseIf Not (TypeOf oEntity Is IJAssemblyChild) Then
            'root part but not leaf part
            'create a profile system style name w/ extra level of indexing
            strNamingParents = strProfileSystemOname
            
            'only create the name and get next index if name will be updated
            If (strNamingParents <> oActiveEntity.NamingParentsString) Then
                Call m_oNameHelp.SDIndexByCounter(strNamingParents, _
                                                  iPos, _
                                                  oResourceMgr)
                strName = strProfileSystemOname + "." + Format(iPos)
            End If
        Else 'its a leaf stiffener
            strNamingParents = strPlatePartName + _
                               strProfileSystemOname + _
                               strProfileCat
            
            'only create the name and get next index if name will be updated
            If (strNamingParents <> oActiveEntity.NamingParentsString) Then
                Call m_oNameHelp.SDIndexByCounter(strNamingParents, _
                                                  iPos, _
                                                  oResourceMgr)
                strName = strPlatePartName + "-" + _
                          strProfileSystemOname + _
                          strProfileCat + _
                          "." + Format(iPos)
            End If
        End If

        'if naming parents string different, get index and update name
        If (strNamingParents <> oActiveEntity.NamingParentsString) Then
            oActiveEntity.NamingParentsString = strNamingParents
            
            m_oNameHelp.GetWorkShareLocationID oResourceMgr, strLocationID
            If Len(strLocationID) > 0 Then
                strName = strName & "-" & strLocationID
            End If
            
            oNamedItem.Name = strName
        End If
    Else 'using position indexing to number profiles across a plate
        Dim eProfAxis As SDPrinProfileAxis 'not used on this call
        Dim iProfPos As Integer 'position to index on
        Dim bChanged As Boolean
        Dim strGroupCriteria As String
        Dim ePartSymmetry As JSBM_SYMMETRY
        Dim ePartSide As SDSymPriority
        Dim oUpdateEntity As Object
         Dim oPartPos As IJShpStrPartPosition
        Dim bUserUpdated As Boolean
        eProfAxis = GetProfileAxis(oStiffPart, iProfPos)
        ePartSymmetry = oStiffPart.Symmetry
        
        ePartSide = GetPartSide(oStiffPart, eProfAxis)
        
        If Not (ePartSymmetry = Symmetry_PortAndStarboard) Then
            If (ePartSide = Starboard) Then
                iProfPos = iProfPos + NOSYMM_STBD_ADDON
            ElseIf (ePartSide = Center) Then
                iProfPos = iProfPos + NOSYMM_CENTER_ADDON
            Else
                iProfPos = iProfPos + NOSYMM_PORT_ADDON
            End If
        Else 'symmetric
            If (ePartSide = Port) Then
'                Set oUpdateEntity = oStiffPart.SymmetricalPart
            End If
        End If
                                                     
        'the group criteria is good for name prefix
        strGroupCriteria = oGroupMonitor.Criteria
   
        'insert or update DB row
        If (oUpdateEntity Is Nothing) Then
            Set oUpdateEntity = oEntity
        End If
        Call m_oNameHelp.SDUpdateIndex(oUpdateEntity, _
                                       strGroupCriteria, _
                                       iProfPos, _
                                       0, _
                                       bChanged, _
                                       oResourceMgr, _
                                       m_oNameUtil)
         Set oPartPos = oEntity
        bUserUpdated = oPartPos.IsPartUpdated
                
        If (bChanged Or bUserUpdated) Then 'notify GroupMonitor that will sort then recompute
            Call oGroupMonitor.NotifyPositionChange(oUpdateEntity)
        Else 'Tables good - get name  and load name
            Dim nPos As Integer
            
            Call m_oNameHelp.SDGetIndex(oUpdateEntity, _
                                        nPos, _
                                        oResourceMgr, _
                                        m_oNameUtil)
                                        
            m_oNameHelp.GetWorkShareLocationID oResourceMgr, strLocationID
            If Len(strLocationID) > 0 Then
                oNamedItem.Name = strGroupCriteria + "." + Format(nPos) + "-" + strLocationID
            Else
                oNamedItem.Name = strGroupCriteria + "." + Format(nPos)
            End If
            
           
        End If
        
        Set oUpdateEntity = Nothing
    End If
           
    Set oStiffPart = Nothing
 Exit Sub

errLabel:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
End Sub

'********************************************************************
' Function:
'   IJNameRule_GetNamingParents
' Description
'   All the Naming Parents that need to participate in an objects naming are
'   added here to the IJElements collection. The parents added here are used
'   in computing the name of the object in ComputeName() of the same
'   interface. Both these methods are called from naming rule semantic.
'********************************************************************
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    Dim oStiffPart As StructDetailObjects.ProfilePart
    Dim oResourceMgr As IUnknown
    Dim oSysParent As IJSystem
    Dim oPlatePart As IJPlate
    Dim oWorkColl As IJElements
    Dim ProfileCatPriority As Long
    Dim strProfileCat As String
    Dim eStiffType As StructProfileType
    Dim ePlateAxis As SDPrinPlateAxis
    Dim eProfAxis As SDPrinProfileAxis
    Dim bIsSystem As Boolean
    On Error GoTo label
    
    Set oWorkColl = New IMSCoreCollections.JObjectCollection
    Set oStiffPart = New StructDetailObjects.ProfilePart
    Set oStiffPart.object = oEntity
    
    'stand alone parts don't have a root system
    If Not (TypeOf oEntity Is IJDStandAloneStiffenerPart) Then
        Dim oSDHelper As StructDetailObjects.Helper
        
        'root system is used to provide O.name for part
        Set oSDHelper = New StructDetailObjects.Helper
        Set oSysParent = oSDHelper.Object_RootParentSystem(oEntity)
        Set oSDHelper = Nothing
        If (Not oSysParent Is Nothing) Then 'still make sure not stand alone
            Call oWorkColl.Add(oSysParent)
        End If
    End If
   
    'root parts only need parent system so make sure it's a leaf
    'before collecting other information for name
    'also for edge reinforcement don't need plate or other info
    eStiffType = oStiffPart.ProfileType
    If (TypeOf oEntity Is IJAssemblyChild And _
        Not (eStiffType = sptEdgeReinforcement)) Then 'its a non-ER leaf part
    
        'try to find the plate part
        On Error Resume Next
        Call FindPlatePartWithLongestToeTrace(oStiffPart, oPlatePart)
        On Error GoTo label
        If (oPlatePart Is Nothing) Then 'not detailed, just use parent plate
            Call oStiffPart.GetStiffenedPlate(oPlatePart, bIsSystem)
        End If
            
        Call oWorkColl.Add(oPlatePart)
    
        ' profile needs position sorting if     1) it is a given type of catalog
        '                                       2) it is on a planar plate
        '                                       3) it is linear
        '                                       4) it is NOT edge reinforcement
        Set oResourceMgr = m_oSDOHelper.GetResourceManagerFromObject(oStiffPart.object)
        strProfileCat = GetProfileCategory(oStiffPart, oResourceMgr, ProfileCatPriority)
'       to improve performance - this check is done after the Part symmetry is checked
'       If (TypeOf oPlatePart Is IJPlatePart) Then 'need axis info
'            ePlateAxis = m_oNameHelp.PrincipalAxisOfPlate(oPlatePart, m_oNameUtil)
'            eProfAxis = GetProfileAxis(oStiffPart)
'        End If
        
        If (TypeOf oPlatePart Is IJPlatePart) Then

            Dim ePartSymmetry As JSBM_SYMMETRY
            
            'if symmetry hasn't been set, don't create group monitor yet
            ePartSymmetry = oStiffPart.Symmetry
            
            If Not (ePartSymmetry = Symmetry_NotSet) Then
                ePlateAxis = m_oNameHelp.PrincipalAxisOfPlate(oPlatePart, m_oNameUtil)
                eProfAxis = GetProfileAxis(oStiffPart)
                If (Not (ePlateAxis = Nonplanar)) And _
                    (Not (eProfAxis = other)) Then
            '       (ProfileCatPriority < 34 Or _
            '        (ProfileCatPriority = 34 And IsProfileSystem(oPlatePart))) And _

                    'grouping criteria
                    Dim strPlatePartName As String
                    Dim strProfileSystemOname As String
                    Dim strBlkName As String
                    Dim strGroupCriteria As String
                    Dim oGroupMonitor As Object
                    Call GetNonpositionNamingFields(oStiffPart, _
                                                    oSysParent, _
                                                    oPlatePart, _
                                                    strPlatePartName, _
                                                    strProfileSystemOname)
                    strBlkName = GetBlockName(oEntity)
                    
                    'make sure parent profile is named and part is assigned
                    'to an assembly block, otherwise not ready to use
                    'group monitor until non-indexed naming is complete
                    If Not (strProfileSystemOname = "") And _
                       Not (strBlkName = NO_BLOCK) Then
                        strGroupCriteria = strPlatePartName + "-" + _
                                           strProfileSystemOname + _
                                           strProfileCat
                        Call m_oNameHelp.FindOrCreateGroupMonitor(oEntity, _
                                                                  strGroupCriteria, _
                                                                  oGroupMonitor, _
                                                                  oResourceMgr, _
                                                                  m_oNameUtil)
                        If Not (oGroupMonitor Is Nothing) Then
                            oWorkColl.Add oGroupMonitor
                            Set oGroupMonitor = Nothing
                        End If
                    End If
                End If
            End If
        End If
    End If
           
    Set IJNameRule_GetNamingParents = oWorkColl
    Set oStiffPart = Nothing
    Set oWorkColl = Nothing
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function

'********************************************************************
' Function:
'   AbsMin
' Description
'   Calculate the absolute minimum of the two numbers
'********************************************************************
Private Function AbsMin(Low, High As Double) As Double
    On Error GoTo label
    If (Abs(Low) < Abs(High)) Then
        AbsMin = Low
    Else
        AbsMin = High
    End If
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "AbsMin", Err.Description
    Err.Raise E_FAIL
End Function

'********************************************************************
' Function:
'   GetNonpositionNamingFields
' Description
'   Get the non-position naming fields
'********************************************************************
Private Sub GetNonpositionNamingFields(ByVal oStiffPart As StructDetailObjects.ProfilePart, _
                                       ByVal oProfileParent As IJSystem, _
                                       ByVal oPlatePart As IJPlate, _
                                       strPlatePartName As String, _
                                       strProfileSystemOname As String)
    Dim oSysNamedItem As IJNamedItem
    On Error GoTo label
    
    'stand alone parts must recreate O.name since they don't have a root
    'system to get it from (except for ER, where no O.name used)
    If (TypeOf oStiffPart.object Is IJDStandAloneStiffenerPart And _
        Not (TypeOf oStiffPart.object Is IJProfileER)) Then 'stand alone
        strProfileSystemOname = GetReferencePlaneNameForStandaloneProfilePart(oStiffPart)
    'for root parts, just use the system name which will have an index added
    ElseIf Not (TypeOf oStiffPart.object Is IJAssemblyChild) Then 'root only
        Set oSysNamedItem = oProfileParent
        strProfileSystemOname = oSysNamedItem.Name
        Set oSysNamedItem = Nothing
    ElseIf Not (TypeOf oStiffPart.object Is IJProfileER) Then 'get O.name for root system
        strProfileSystemOname = GetRefPlaneONameInfo(oProfileParent)
        'strProfileSystemOname = GetProfileSystemOname(oProfileParent)
    End If

    If Not (oPlatePart Is Nothing) Then
        strPlatePartName = GetPlatePartName(oPlatePart)
    Else
        strPlatePartName = ""
    End If
Exit Sub

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetNonpositionNamingFields", Err.Description
    Err.Raise E_FAIL
End Sub

'********************************************************************
' Function:
'   GetPartSide
' Abstract:
'   Calculates the part side (Port, Stbd, Center)
'********************************************************************
Private Function GetPartSide(ByVal oStiffPart As StructDetailObjects.ProfilePart, _
                             ByVal eProfAxis As SDPrinProfileAxis) As SDSymPriority
    Dim ePartSide As SDSymPriority
    On Error GoTo label

    If Not (eProfAxis = other) Then
        Dim oLowPos As IJDPosition
        Dim oHighPos As IJDPosition
        
        Call m_oNameUtil.GetRangeCorners(oStiffPart.object, oLowPos, oHighPos)

        If (eProfAxis = ydir) Or (eProfAxis = zdir) Then
            If (oLowPos.y > NEG_CL_TOLERANCE) And _
               (oHighPos.y > NEG_CL_TOLERANCE) Then
                ePartSide = Port
            ElseIf (oLowPos.y < CL_TOLERANCE) And _
                   (oHighPos.y < CL_TOLERANCE) Then
                ePartSide = Starboard
            Else
                ePartSide = Center
            End If
        ElseIf eProfAxis = xdir Then
            If (Abs(oHighPos.y) < CL_TOLERANCE) Then
                ePartSide = Center
            ElseIf oLowPos.y > NEG_CL_TOLERANCE Then
                ePartSide = Port
            Else
                ePartSide = Starboard
            End If
        End If
    End If
    
    GetPartSide = ePartSide
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetPartSide", Err.Description
    Err.Raise E_FAIL
End Function

'********************************************************************
' Function:
'   GetPlatePartName
' Description
'   get the name of the given plate part
'********************************************************************
Private Function GetPlatePartName(ByVal oPlatePart As IJPlate) As String
    Dim oPlatePartItem As IJNamedItem
    On Error GoTo label
    
    Set oPlatePartItem = oPlatePart
    GetPlatePartName = oPlatePartItem.Name
    Set oPlatePartItem = Nothing
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetPlatePartName", Err.Description
    Err.Raise E_FAIL
End Function

'********************************************************************
' Function:
'   GetProfileAxis
' Description
'   get the profile axis
'********************************************************************
Private Function GetProfileAxis(ByVal oStiffPart As StructDetailObjects.ProfilePart, _
                                Optional iLowValue As Integer) As SDPrinProfileAxis
    Dim oGeomOpsToolBox As IMSModelGeomOps.DGeomOpsToolBox
    Dim oTopologyLocate As IJTopologyLocate
    Dim oWireBody As IJWireBody
    Dim oStiffUnk As IUnknown
    Dim oMiddlePoint As IJDPosition
    Dim oOutputPoint As IJDPosition
    Dim oTangentVector As IJDVector
    On Error GoTo label
    
    'get the wire body used to figure out axis
    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
    Set oStiffUnk = oStiffPart.object
    Set oWireBody = oTopologyLocate.GetProfileParentWireBody(oStiffUnk)
    Set oStiffUnk = Nothing
    Set oTopologyLocate = Nothing
    
    'Get the geometry utility to see if its linear
    Set oGeomOpsToolBox = New IMSModelGeomOps.DGeomOpsToolBox
    
    'Get the middle point of the WireBody
    Call oGeomOpsToolBox.GetMiddlePointOfCompositeCurve(oWireBody, _
                                                        oMiddlePoint)
    
    If Not oMiddlePoint Is Nothing Then
       'Project the middle point to the WireBody to get the tangent value back
       Call oGeomOpsToolBox.ProjectPointOnWireBody(oWireBody, _
                                                   oMiddlePoint, _
                                                   oOutputPoint, _
                                                   oTangentVector)
    End If
    
    Set oGeomOpsToolBox = Nothing

    'check to see if linear along an axis
    If (oTangentVector.x < 0.001) Or _
       (oTangentVector.y < 0.001) Or _
       (oTangentVector.z < 0.001) Then 'verify what axis and get low value
       
        Dim oStartPoint As IJDPosition
        Dim oEndPoint As IJDPosition
        
        Call oWireBody.GetEndPoints(oStartPoint, oEndPoint)
        
        If (Abs(oEndPoint.x - oStartPoint.x) < 0.001) And _
           (Abs(oEndPoint.y - oStartPoint.y) < 0.001) Then
            GetProfileAxis = zdir
            iLowValue = AbsMin(oStartPoint.z, oEndPoint.z)
        ElseIf (Abs(oEndPoint.x - oStartPoint.x) < 0.001) And _
               (Abs(oEndPoint.z - oStartPoint.z) < 0.001) Then
            GetProfileAxis = ydir
            iLowValue = AbsMin(oStartPoint.y, oEndPoint.y)
        ElseIf (Abs(oEndPoint.z - oStartPoint.z) < 0.001) And _
               (Abs(oEndPoint.y - oStartPoint.y) < 0.001) Then
            GetProfileAxis = xdir
            iLowValue = AbsMin(oStartPoint.x, oEndPoint.x)
        Else
            GetProfileAxis = other
        End If
    Else 'not linear along an axis
        GetProfileAxis = other
    End If
    
    Set oWireBody = Nothing
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetProfileAxis", Err.Description
    Err.Raise E_FAIL
End Function

'********************************************************************
' Function:
'   GetProfileCategory
' Description
'   get the category for the profile part
'********************************************************************
Private Function GetProfileCategory(ByVal oStiffPart As StructDetailObjects.ProfilePart, _
                                    ByVal oModelResourceMgr As IUnknown, _
                                    Optional ProfileCatPriority As Long) As String
    Dim objCategory As GSCADMetaDataHelper.CMetaDataCategoryQuery
    Dim LongNames() As String
    Dim ShortNames() As String
    Dim PriorityValues() As Long
    Dim eType As StructProfileType
    Dim lCat As Long
    Dim status As Long
    On Error GoTo label
    
    GetProfileCategory = "" 'make sure its blank in case no category
    
    Set objCategory = New GSCADMetaDataHelper.CMetaDataCategoryQuery
    eType = oStiffPart.ProfileType
    lCat = oStiffPart.NamingCategory
    
    If lCat < 0 Then 'check for no category returned (negative not valid)
        GoTo CatNotFound
    End If
    
    Select Case eType
        Case sptVertical
            status = objCategory.GetVerticalProfileCategoryInfo(oModelResourceMgr, LongNames, ShortNames, PriorityValues)
        Case sptTransversal
            status = objCategory.GetTransverseProfileCategoryInfo(oModelResourceMgr, LongNames, ShortNames, PriorityValues)
        Case sptLongitudinal
            status = objCategory.GetLongitudinalProfileCategoryInfo(oModelResourceMgr, LongNames, ShortNames, PriorityValues)
        Case sptEdgeReinforcement
            status = objCategory.GetEdgeReinforcementCategoryInfo(oModelResourceMgr, LongNames, ShortNames, PriorityValues)
        Case sptStandalone
            status = objCategory.GetCategoryInfo(oModelResourceMgr, _
                                                 "ProfilePartCategory", _
                                                 LongNames, _
                                                 ShortNames, _
                                                 PriorityValues)
        Case Else
            GoTo CatNotFound
    End Select
    
    ' lCat is the priorityvalue, NOT the index.  So search for the value to get the appropriate index
    ProfileCatPriority = lCat
    
    Dim i As Long
    
    For i = LBound(PriorityValues) To UBound(PriorityValues)
        If PriorityValues(i) = lCat Then
            GetProfileCategory = ShortNames(i)
            Exit For
        End If
    Next

    
CatNotFound:
    Set objCategory = Nothing
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetProfileCategory", Err.Description
    Err.Raise E_FAIL
End Function

'********************************************************************
' Function:
'   GetProfileSystemOname
' Description
'   Get the system O.name from the system name
'********************************************************************
Private Function GetProfileSystemOname(ByVal oProfileParent As IJSystem) As String
    Dim oParentItem As IJNamedItem
    Dim strName As String
    Dim strFullName As String
    On Error GoTo label
    
    'this triggers off existing struct entity naming, where a dash
    'separates pseudo-O.name from index information. this does NOT
    'work properly with negative ref planes, and will need to be updated
    'when struct entity naming is updated to properly assign O.name
    Set oParentItem = oProfileParent
    strName = oParentItem.Name
    Set oParentItem = Nothing
    
    'if parent system named, get O.name
    If Not (strName = "") Then
        strFullName = StrReverse(strName)
        strFullName = StrReverse(Split(strFullName, "-")(0))
        GetProfileSystemOname = Left(strFullName, Len(strFullName)) ' - 1)
    Else 'not named, but may be named later so don't throw an error
        GetProfileSystemOname = ""
    End If
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetProfileSystemOname", Err.Description
    Err.Raise E_FAIL
End Function

'********************************************************************
' Function:
'   GetReferencePlaneNameForStandaloneProfilePart
' Abstract:
'   Get the reference plane for a stand-alone profile part
'********************************************************************
Private Function GetReferencePlaneNameForStandaloneProfilePart(ByVal oStiffPart As StructDetailObjects.ProfilePart) As String
    Dim oNamedItem As IJNamedItem
    Dim oRefPlane As IUnknown
    Dim strFullName As String
    On Error GoTo label
    
    strFullName = ""
    Call m_oNameUtil.GetProfileRefPlane(oStiffPart.object, oRefPlane)
    If Not (oRefPlane Is Nothing) Then
        Set oNamedItem = oRefPlane
        strFullName = oNamedItem.Name
        Set oRefPlane = Nothing
        Set oNamedItem = Nothing
     Else
        'StrcutPlane is the reference for landing curve
        Dim oIJPlane As IJPlane

        'Get the reference StructPlane
        Call m_oNameUtil.GetSubstitutedDefPlaneForSAStiff(oStiffPart.object, oIJPlane)

        'Get the ref structplane location
        strFullName = GetRefPlaneName(oIJPlane)
    End If
    
    'now filter out negative sign/ref plane index
    GetReferencePlaneNameForStandaloneProfilePart = m_oNameHelp.GetFilteredRefPlaneName(strFullName)
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetReferencePlaneNameForStandaloneProfilePart", Err.Description
    Err.Raise E_FAIL
End Function

'********************************************************************
' Function:
'   GetRefPlaneONameInfo
' Abstract:
'   Get the ref plane O.name info using the system
'********************************************************************
Private Function GetRefPlaneONameInfo(ByVal oSysParent As IJSystem) As String
    Dim oItem As IJNamedItem
    Dim oUtil As IJProfileAttributes
    Dim oPlate As IJPlate
    Dim oRefPlane As IJPlane
    Dim strOName As String
    On Error GoTo label
    
    'get the ref plane from the root parent system
    'Call m_oNameUtil.GetRefPlaneFromRootSystem(oSysParent, oRefPlane)
    Set oUtil = New ProfileUtils
    oUtil.GetInputs_StiffenerByIntersect oSysParent, oPlate, oRefPlane
    If Not (oRefPlane Is Nothing) Then
        Set oItem = oRefPlane
        If Len(oItem.Name) = 0 Then
            strOName = GetRefPlaneName(oRefPlane)
        Else
            strOName = oItem.Name
        End If
        
        Set oItem = Nothing
    Else
       strOName = ""
    End If
    
    'now filter out negative sign/ref plane index
    GetRefPlaneONameInfo = m_oNameHelp.GetFilteredRefPlaneName(strOName)
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetRefPlaneONameInfo", Err.Description
    Err.Raise E_FAIL
End Function

'********************************************************************
' Function:
'   GetRefPlaneName
' Abstract:
'   Get the ref plane location
'********************************************************************
Private Function GetRefPlaneName(ByVal oPlane As IJPlane) As String
    If Not oPlane Is Nothing Then
        Dim strPositionString As String
        Dim oPlaneNormalVector As New DVector
        Dim dx As Double, dy As Double, dz As Double
        Dim Rx As Double, Ry As Double, Rz As Double
        Dim dDotx As Double, dDoty As Double, dDotz As Double
        Dim oXaxis As New DVector
        Dim oYaxis As New DVector
        Dim oZaxis As New DVector
        Dim iPosition As Integer
        
        oXaxis.Set 1, 0, 0
        oYaxis.Set 0, 1, 0
        oZaxis.Set 0, 0, 1
        
        oPlane.GetNormal dx, dy, dz
        oPlane.GetRootPoint Rx, Ry, Rz
        oPlaneNormalVector.Set dx, dy, dz
        
        dDotx = Math.Abs(oPlaneNormalVector.Dot(oXaxis))
        dDoty = Math.Abs(oPlaneNormalVector.Dot(oYaxis))
        dDotz = Math.Abs(oPlaneNormalVector.Dot(oZaxis))
        
        If (Abs(dDotx - 1) <= 0.000001) Then
            'roundoff the X position to nearest mm
            iPosition = Round(Rx)
        ElseIf (Abs(dDoty - 1) <= 0.000001) Then
            'roundoff the Y position to nearest mm
            iPosition = Round(Ry)
        ElseIf (Abs(dDotz - 1) <= 0.000001) Then
            'roundoff the Z position to nearest mm
            iPosition = Round(Rz)
        End If
        
        strPositionString = CStr(iPosition)
        
        Set oPlaneNormalVector = Nothing
        Set oXaxis = Nothing
        Set oYaxis = Nothing
        Set oZaxis = Nothing
    End If
    
    GetRefPlaneName = strPositionString
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetRefPlaneName", Err.Description
    Err.Raise E_FAIL
End Function
'********************************************************************
' Function:
'   IsProfileSystem
' Abstract:
'   Is the profile part parent a profile system
'********************************************************************
Private Function IsProfileSystem(ByVal oProfileParent As IJSystem) As Boolean
    On Error GoTo label
    
    If TypeOf oProfileParent Is IJProfile Then
        IsProfileSystem = True
    Else 'nope
        IsProfileSystem = False
    End If
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IsProfileSystem", Err.Description
    Err.Raise E_FAIL
End Function

'********************************************************************
' Function:
'   SetEnviron
' Abstract:
'   Extract the key environment objects from the naming parents collection
'********************************************************************
Private Sub SetEnviron(ByVal oParents As IJElements, _
                       ByRef oProfileParent As IJSystem, _
                       ByRef oPlatePart As IJPlate, _
                       ByRef oGroupMonitor As Object)
    Dim oElement As Object
    Dim oPlateOnly As IJPlate
    On Error GoTo label
    
    'get all the relevant parent objects used for naming
    For Each oElement In oParents
        'must search for PlatePart first since Plate might also be the parent
        If (TypeOf oElement Is IJPlatePart) Then 'got our plate part
            Set oPlatePart = oElement
        ElseIf (TypeOf oElement Is IJPlate) Then
            'store it cause it could be the "platepart" or the parent system
            If (oPlateOnly Is Nothing) Then
                Set oPlateOnly = oElement
            End If
        ElseIf (TypeOf oElement Is IJSystem) Then 'no "plate-ness", just system
            Set oProfileParent = oElement
        ElseIf (TypeOf oElement Is IJGroupMon) Then
            Set oGroupMonitor = oElement
        End If
        
        Set oElement = Nothing
    Next
    
    'make sure we got a parent system, cause the parent might be a plate
    If (oProfileParent Is Nothing) And (Not (oPlateOnly Is Nothing)) Then
        Set oProfileParent = oPlateOnly
    End If
        
    If (oPlatePart Is Nothing) Then 'plate part is just a plate right now
        Set oPlatePart = oPlateOnly
    End If
Exit Sub

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "SetEnviron", Err.Description
    Err.Raise E_FAIL
End Sub

'********************************************************************
' Function:
'   FindPlatePartWithLongestToeTrace
' Abstract:
'   Finds the Connected Plate Part with the Longest Toe Trace for the
'   Stiffener
'********************************************************************
Private Sub FindPlatePartWithLongestToeTrace(ByVal oStiffPart As StructDetailObjects.ProfilePart, _
                                             ByRef oPlatePart As IJPlate)
    On Error GoTo label
    Dim oPartInfo As PartInfo
    Set oPartInfo = New PartInfo
    Set oPlatePart = oPartInfo.FindPlatePartWithLongestStiffenerPartToeTrace(oStiffPart.object)
    Exit Sub
label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "FindPlatePartWithLongestToeTrace", Err.Description
    Err.Raise E_FAIL
End Sub

'********************************************************************
' Function:
'   GetBlockName
' Abstract:
'   Retrieves the parent assembly to access the block and get the block name
'********************************************************************
Private Function GetBlockName(oEntity As Object) As String
    Dim oAsmChild As IJAssemblyChild
    Dim oAsmParent As IJAssembly
    On Error GoTo label
    
    If (TypeOf oEntity Is IJAssemblyChild) Then
        'if not detailed then not in an assembly, so make sure
        Set oAsmChild = oEntity
        On Error Resume Next
        Set oAsmParent = oAsmChild.Parent
        On Error GoTo label
        Set oAsmChild = Nothing
        If Not (oAsmParent Is Nothing) Then
            GetBlockName = m_oNameHelp.GetBlockName(oAsmParent, m_oNameUtil)
            Set oAsmParent = Nothing
        Else
            GetBlockName = NO_BLOCK
        End If
    End If
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetBlockName", Err.Description
    Err.Raise E_FAIL
End Function

'********************************************************************
' Function:
'   GetParentAssyConn
' Abstract:
'   Retrieves the parent assembly connection given a physical connection
'********************************************************************
Private Function GetParentAssyConn(oChild As Object) As Object
    Dim oParentUnk As IUnknown
    Dim oChildUnk As IUnknown
    Dim oIJSystemChild As IJSystemChild
    
    'if not a physical connection, skip it
    If Not (TypeOf oChild Is IJStructPhysicalConnection) Then
        Set GetParentAssyConn = Nothing
        Exit Function
    End If
    
    Set oChildUnk = oChild
    
    'look up the tree for the parent assembly connection
    Do
        If (TypeOf oChildUnk Is IJSystemChild) Then
            Set oIJSystemChild = oChildUnk
            Set oParentUnk = oIJSystemChild.GetParent
            Set oIJSystemChild = Nothing
        Else 'oops, unexpected so just stop so don't get in infinite loop
            Set oChildUnk = Nothing
            Exit Do
        End If
        
        'if current parent is assembly connection, set it and quit loop
        If (TypeOf oParentUnk Is IJAssemblyConnection) Then
            Set GetParentAssyConn = oParentUnk
            Exit Do
        Else 'keep looking
            Set oChildUnk = Nothing
            Set oChildUnk = oParentUnk
            Set oParentUnk = Nothing
        End If
    Loop While (oParentUnk Is Nothing)
    
End Function
