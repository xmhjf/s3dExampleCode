VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SDPlateNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2001-2003 Global Research and Development, Inc.
'  All rights reserved.
'
'  Project: SDPlatePartNameRules
'
'  Abstract: The file contains name rules for PlateParts
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
                                                        
Private Const MODULE = "SDPlatePartNameRules.SDPlateNameRule."
Private Const NO_BLOCK = "NoBlk"
Private Const NOSYMM_PORT_ADDON = 10000
Private Const NOSYMM_STBD_ADDON = 20000
Private Const NOSYMM_CENTER_ADDON = 30000
Private Const POSITION_TOL = 0.5
Private Const NEG_POSITION_TOL = -0.5

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
   
    'set globals
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
    'clear globals
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
    Dim oPlatePart As StructDetailObjects.PlatePart 'wrapper
    Dim oResourceMgr As IUnknown
    Dim oSysParent As IJSystem
    Dim oAsmParent As IJAssembly
    Dim oGroupMonitor As IJGroupMon
    Dim strLocationID As String
    On Error GoTo errLabel
            
    If oEntity Is Nothing Then 'error, no entity to name
        Err.Raise E_INVALIDARG, MODULE, METHOD
    Else 'entity valid, set up wrapper and name interface
        Set oPlatePart = New StructDetailObjects.PlatePart
        Set oPlatePart.object = oEntity
        Set oNamedItem = oEntity  'for setting the name
    End If
    
    'Finds objects involved in naming structure
    Call SetEnviron(oParents, oSysParent, oAsmParent, oGroupMonitor)
    Set oResourceMgr = m_oSDOHelper.GetResourceManagerFromObject(oPlatePart.object)
                
    If (oGroupMonitor Is Nothing) Then 'not using positional naming
        Dim strPos As String 'indexing
        Dim strSysONameInfo As String
        Dim strBlkParentName As String
        Dim strCategoryName As String
        Dim strNamedParentsString As String
        Dim strName As String
        Dim iPos As Long
        
        'use index naming fields to construct name
        Call GetNonpositionNamingFields(oPlatePart, _
                                        oResourceMgr, _
                                        oSysParent, _
                                        oAsmParent, _
                                        strSysONameInfo, _
                                        strBlkParentName, _
                                        strCategoryName)
                                        
        If (TypeOf oEntity Is IJCollarPart) Then 'its a collar
            strNamedParentsString = strBlkParentName + _
                                    strCategoryName
                                    
            'only create the name and get next index if name will be updated
            If (strNamedParentsString <> oActiveEntity.NamingParentsString) Then
                Call m_oNameHelp.SDIndexByCounter(strNamedParentsString, _
                                                  iPos, _
                                                  oResourceMgr)
                strName = "<" + strBlkParentName + ">-" + _
                          strCategoryName + "." + Format(iPos)
            End If
        ElseIf (TypeOf oEntity Is IJSmartPlate) Then
            Dim oSmartPlate As IJSmartPlate
            Dim PlateType As SmartPlateTypes
            
            'currently just name these plates using naming category
            Set oSmartPlate = oEntity
            PlateType = oSmartPlate.SmartPlateType
            If (PlateType = spType_BRACKET) Then
                strNamedParentsString = strBlkParentName + _
                                        strCategoryName
            Else 'no other specific recognized types (note currently same)
                strNamedParentsString = strBlkParentName + _
                                        strCategoryName
            End If
                                    
            'only create the name and get next index if name will be updated
            If (strNamedParentsString <> oActiveEntity.NamingParentsString) Then
                Call m_oNameHelp.SDIndexByCounter(strNamedParentsString, _
                                                  iPos, _
                                                  oResourceMgr)
                strName = "<" + strBlkParentName + ">-" + _
                          strCategoryName + "." + Format(iPos)
            End If
        ElseIf Not (TypeOf oEntity Is IJAssemblyChild) Then
            'root part but not leaf part
            'create a plate system style name w/ extra level of indexing
            strNamedParentsString = strSysONameInfo
            
            'only create the name and get next index if name will be updated
            If (strNamedParentsString <> oActiveEntity.NamingParentsString) Then
                Call m_oNameHelp.SDIndexByCounter(strNamedParentsString, _
                                                  iPos, _
                                                  oResourceMgr)
                strName = strSysONameInfo + "." + Format(iPos)
            End If
        Else 'not yet detailed in a block or its symmetry has not been set
            strNamedParentsString = strSysONameInfo + strBlkParentName + _
                                    strCategoryName
            
            'only create the name and get next index if name will be updated
            If (strNamedParentsString <> oActiveEntity.NamingParentsString) Then
                Call SDPartIndexByCount(oPlatePart, _
                                        oResourceMgr, _
                                        strNamedParentsString, _
                                        strPos)
                strName = "<" + strBlkParentName + ">-" + _
                          strSysONameInfo + strCategoryName + "." + strPos
            End If
        End If
        
        'if naming parents string different, update name and parents string
        If (strNamedParentsString <> oActiveEntity.NamingParentsString) Then
            oActiveEntity.NamingParentsString = strNamedParentsString
            m_oNameHelp.GetWorkShareLocationID oResourceMgr, strLocationID
            If Len(strLocationID) > 0 Then
                strName = strName & "-" & strLocationID
            End If
            oNamedItem.Name = strName
        End If
    Else 'use positional naming
        Dim oLowPos As IJDPosition
        Dim oHighPos As IJDPosition
        Dim px As Integer
        Dim py As Integer
        Dim bChanged As Boolean
        Dim strGroupCriteria As String
        Dim eAxis As SDPrinPlateAxis
        Dim ePartSymmetry As JSBM_SYMMETRY
        Dim ePartSide As SDSymPriority
        Dim oUpdateEntity As Object
        Dim oPartPos As IJShpStrPartPosition
        Dim bUserUpdated As Boolean
        Call m_oNameUtil.GetRangeCorners(oEntity, oLowPos, oHighPos)
        
        'note that we are rounding up to meter
        eAxis = m_oNameHelp.PrincipalAxisOfPlate(oEntity, m_oNameUtil)
        ePartSymmetry = oPlatePart.Symmetry
        
        ePartSide = GetPartSide(eAxis, oLowPos, oHighPos)
        If eAxis = Deck Then
                px = oLowPos.x + POSITION_TOL
            If Not (ePartSymmetry = Symmetry_PortAndStarboard) Then
                If (ePartSide = Starboard) Then
                    px = px + NOSYMM_STBD_ADDON
                        py = Abs(oLowPos.y + POSITION_TOL)
                ElseIf (ePartSide = Center) Then
                    px = px + NOSYMM_CENTER_ADDON
                        py = Abs(oLowPos.y + POSITION_TOL)
                Else
                    px = px + NOSYMM_PORT_ADDON
                        py = Abs(oHighPos.y)
                End If
            Else 'symmetric
                If (ePartSide = Starboard) Then
                        py = Abs(oLowPos.y)
                ElseIf (ePartSide = Port) Then
                        py = Abs(oHighPos.y)
                    ' Set oUpdateEntity = oPlatePart.SymmetricalPart
                End If
            End If
        ElseIf eAxis = Longt Then
                px = oHighPos.z + POSITION_TOL
                py = oLowPos.x + POSITION_TOL
            If Not (ePartSymmetry = Symmetry_PortAndStarboard) Then
                If (ePartSide = Starboard) Then
                    px = px + NOSYMM_STBD_ADDON
                ElseIf (ePartSide = Port) Then
                    px = px + NOSYMM_PORT_ADDON
                Else
                    px = px + NOSYMM_CENTER_ADDON
                End If
            Else 'symmetric
                '************************************************************
                'NOTE: For longitudinal, current system naming keeps negative
                'sign from ref plane as part of system name, so we cannot
                'use the "other" part because the criteria won't match the
                'entity being updated and group monitor semantic gets stuck
                'in a loop never thinking its done updating names.
                'Symmetrical sorting should still work, this just removes
                'the efficiency enhancement of having only 1 line in table
                'for 2 symmetrical parts
                '************************************************************
                Set oUpdateEntity = oEntity
            End If
        ElseIf eAxis = Transv Then
                py = oHighPos.z + POSITION_TOL
            If Not (ePartSymmetry = Symmetry_PortAndStarboard) Then
                If (ePartSide = Starboard) Then
                        px = Abs(oLowPos.y + POSITION_TOL) + _
                         NOSYMM_STBD_ADDON
                ElseIf (ePartSide = Center) Then
                        px = Abs(oLowPos.y + POSITION_TOL) + _
                         NOSYMM_CENTER_ADDON
                Else
                        px = Abs(oHighPos.y) + NOSYMM_PORT_ADDON
                End If
            Else 'symmetric
                If (ePartSide = Starboard) Then
                        px = Abs(oLowPos.y)
                ElseIf (ePartSide = Port) Then
                        px = Abs(oHighPos.y)
                    ' Set oUpdateEntity = oPlatePart.SymmetricalPart
                End If
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
                                       px, py, _
                                       bChanged, _
                                       oResourceMgr, _
                                       m_oNameUtil)
                
        Set oPartPos = oEntity
        bUserUpdated = oPartPos.IsPartUpdated
        
        If (bChanged Or bUserUpdated) Then 'notigyfy GroupMonitor that will sort then recompute
            Call oGroupMonitor.NotifyPositionChange(oUpdateEntity)
        Else 'table is good, load name
            Call SDPlateGetIndex(oUpdateEntity, oResourceMgr, strPos)
            
            m_oNameHelp.GetWorkShareLocationID oResourceMgr, strLocationID
            If Len(strLocationID) > 0 Then
                oNamedItem.Name = strGroupCriteria + "." + strPos + "-" + strLocationID
            Else
                oNamedItem.Name = strGroupCriteria + "." + strPos
            End If
            
        End If
        
        Set oUpdateEntity = Nothing
    End If
         
    Set oPlatePart = Nothing
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
'   in computing the name of the Object in ComputeName() of the same
'   interface. Both these methods are called from naming rule semantic.
'********************************************************************
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    Dim oPlatePart As StructDetailObjects.PlatePart
    Dim oSysParent As IJSystem
    Dim oAsmParent As IJAssembly
    Dim oWorkColl As IJElements
    Dim PrinPlateAxis As SDPrinPlateAxis
    On Error GoTo label

    Set oWorkColl = New IMSCoreCollections.JObjectCollection
    Set oPlatePart = New StructDetailObjects.PlatePart
    Set oPlatePart.object = oEntity
    
    'stand alone parts, collars, & smart plates don't have a root plate system
    If Not (TypeOf oEntity Is IJDStandAlonePlatePart) And _
       Not (TypeOf oEntity Is IJCollarPart) And _
       Not (TypeOf oEntity Is IJSmartPlate) Then
        Dim oSDHelper As StructDetailObjects.Helper
        Dim oIJPlate As IJPlate
        Dim PlateType As StructPlateType
        
        'root system is used to provide O.name for part
        Set oSDHelper = New StructDetailObjects.Helper
        Set oSysParent = oSDHelper.Object_RootParentSystem(oEntity)
        Set oSDHelper = Nothing
        
        'see if its a hull, if so don't need it
        If Not (oSysParent Is Nothing) Then
            Set oIJPlate = oSysParent
            PlateType = oIJPlate.PlateType
        
            Set oIJPlate = Nothing
        
            If Not (PlateType = Hull) Then
                oWorkColl.Add oSysParent
            End If
        End If
    ElseIf (IsLappedPlate(oPlatePart) = True) Then 'lapped plate, get system
        Dim oPlateUnk As IUnknown
        Dim oSysUnk As IUnknown
        
        Set oPlateUnk = oEntity
        Call m_oNameUtil.GetSystemFromLappedPlate(oPlateUnk, oSysUnk)
        Set oPlateUnk = Nothing
        
        If Not (oSysUnk Is Nothing) Then
            Set oSysParent = oSysUnk
            Set oSysUnk = Nothing
            oWorkColl.Add oSysParent
        End If
    End If
    
    If (TypeOf oEntity Is IJAssemblyChild) Then 'its a leaf part
        Dim oAsmChild As IJAssemblyChild
        
        'if not detailed then not in an assembly, so make sure
        Set oAsmChild = oEntity
        On Error Resume Next
        Set oAsmParent = oAsmChild.Parent
        On Error GoTo label
        Set oAsmChild = Nothing
        If Not (oAsmParent Is Nothing) Then
            oWorkColl.Add oAsmParent
        End If
    End If
'    to improve performance - this check is done after the Part symmetry is checked
'    'since VB always evaluates EVERYTHING, must separate out this check
'    If Not (TypeOf oEntity Is IJSmartPlate) And _
'       Not (TypeOf oEntity Is IJCollarPart) And _
'       Not (TypeOf oEntity Is IJDStandAlonePlatePart) And _
'       Not (oPlatePart.PlateType = Hull) And _
'           (TypeOf oEntity Is IJAssemblyChild) Then
'        PrinPlateAxis = m_oNameHelp.PrincipalAxisOfPlate(oEntity, m_oNameUtil)
'    Else 'initialize, particularly if its a hull
'        PrinPlateAxis = Nonplanar
'    End If

    'make sure its a leaf part and not a Collar or SmartPlate (e.g. bracket)
    'that will not use positional naming before creating GroupMonitor
    If Not (TypeOf oEntity Is IJSmartPlate) And _
       Not (TypeOf oEntity Is IJCollarPart) And _
       Not (oPlatePart.PlateType = Hull) And _
           (TypeOf oEntity Is IJAssemblyChild) Then 'And _
       ' Not (PrinPlateAxis = Nonplanar) Then

        Dim ePartSymmetry As JSBM_SYMMETRY
        
        'if symmetry hasn't been set, don't create group monitor yet
        ePartSymmetry = oPlatePart.Symmetry
        
        If Not (ePartSymmetry = Symmetry_NotSet) Then
            PrinPlateAxis = m_oNameHelp.PrincipalAxisOfPlate(oEntity, m_oNameUtil)
            If Not (PrinPlateAxis = Nonplanar) Then
            
                Dim oResourceMgr As IUnknown
                Dim strSysONameInfo As String
                Dim strBlkParentName As String
                Dim strCategoryName As String
                Dim strGroupCriteria As String
                Dim oGroupMonitor As Object
                
                Set oResourceMgr = m_oSDOHelper.GetResourceManagerFromObject(oPlatePart.object)
                Call GetNonpositionNamingFields(oPlatePart, _
                                                oResourceMgr, _
                                                oSysParent, _
                                                oAsmParent, _
                                                strSysONameInfo, _
                                                strBlkParentName, _
                                                strCategoryName)
        
                'if block assigned, activate group monitor for positional naming
                'otherwise wait till part is assigned to a block
                If Not (strBlkParentName = NO_BLOCK) Then
                    strGroupCriteria = "<" + strBlkParentName + ">-" + _
                                       strSysONameInfo + _
                                       strCategoryName
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
    
                Set oResourceMgr = Nothing
            End If
        End If
    End If
    
    Set IJNameRule_GetNamingParents = oWorkColl
    Set oPlatePart = Nothing
    Set oWorkColl = Nothing
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function

'********************************************************************
' Function:
'   GetNonpositionNamingFields
' Abstract:
'   Get the non-position naming fields
'********************************************************************
Private Sub GetNonpositionNamingFields(ByVal oPlatePart As StructDetailObjects.PlatePart, _
                                       ByVal oResourceMgr As IUnknown, _
                                       ByVal oSysParent As IJSystem, _
                                       ByVal oAsmParent As IJAssembly, _
                                       strSysONameInfo As String, _
                                       strBlkParentName As String, _
                                       strCategoryName As String)
    Dim oSysNamedItem As IJNamedItem
    On Error GoTo label
    
    'stand alone standard (i.e. not smart plate) parts must recreate O.name
    'since they don't have a root plate system to get it from
    If (TypeOf oPlatePart.object Is IJDStandAlonePlatePart) And _
       (IsLappedPlate(oPlatePart) = False) And _
       Not (TypeOf oPlatePart.object Is IJSmartPlate) Then 'stand alone
       
        strSysONameInfo = GetReferencePlaneNameForStandalonePart(oPlatePart)
        
    'for root parts, just use the system name which will have an index added
    ElseIf Not (TypeOf oPlatePart.object Is IJAssemblyChild) Then 'root only
        Set oSysNamedItem = oSysParent
        strSysONameInfo = oSysNamedItem.Name
        Set oSysNamedItem = Nothing
        
    'for standard plate parts, use plate system parent to get O.name
    'Note that collar and smart plate don't need O.name, as well as
    'plates on hull shell, which won't have parent specified
    ElseIf Not (TypeOf oPlatePart.object Is IJCollarPart) And _
           Not (TypeOf oPlatePart.object Is IJSmartPlate) And _
           Not (oSysParent Is Nothing) Then
        strSysONameInfo = GetRefPlaneONameInfo(oSysParent)
    Else 'not used, blank it out
        strSysONameInfo = ""
    End If
    
    If Not (oAsmParent Is Nothing) Then 'find the block
        strBlkParentName = m_oNameHelp.GetBlockName(oAsmParent, m_oNameUtil)
    Else 'no block
        strBlkParentName = NO_BLOCK
    End If
    
    strCategoryName = GetCat(oPlatePart, oResourceMgr)
Exit Sub

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetNonpositionNamingFields", Err.Description
    Err.Raise E_FAIL
End Sub

'********************************************************************
' Function:
'   GetCat
' Abstract:
'   Get the category for the plate part
'********************************************************************
Private Function GetCat(ByVal oPlatePart As StructDetailObjects.PlatePart, _
                        ByVal oModelResourceMgr As IUnknown) As String
    Dim objCategory As GSCADMetaDataHelper.CMetaDataCategoryQuery
    Dim LongNames() As String
    Dim ShortNames() As String
    Dim PriorityValues() As Long
    Dim eType As StructPlateType
    Dim lCat As Long
    Dim status As Long
    On Error GoTo label
    
    GetCat = "" 'make sure its blank in case no category
    
    Set objCategory = New GSCADMetaDataHelper.CMetaDataCategoryQuery
    eType = oPlatePart.PlateType
    lCat = oPlatePart.NamingCategory
    
    If lCat < 0 Then 'check for no category returned (negative not valid)
        GoTo CatNotFound
    End If
       
    Select Case eType
        Case DeckPlate
            status = objCategory.GetDeckCategoryInfo(oModelResourceMgr, LongNames, ShortNames, PriorityValues)
        Case TBulkheadPlate
            status = objCategory.GetTransverseBulkheadCategoryInfo(oModelResourceMgr, LongNames, ShortNames, PriorityValues)
        Case LBulkheadPlate
            status = objCategory.GetLongitudinalBulkheadCategoryInfo(oModelResourceMgr, LongNames, ShortNames, PriorityValues)
        Case Hull
            status = objCategory.GetShellCategoryInfo(oModelResourceMgr, LongNames, ShortNames, PriorityValues)
        Case Else
            'must use other means to figure out what it is
            If (TypeOf oPlatePart.object Is IJCollarPart) Then
                status = objCategory.GetCategoryInfo(oModelResourceMgr, "CollarCategory", LongNames, ShortNames, PriorityValues)
            ElseIf (TypeOf oPlatePart.object Is IJSmartPlate) Then
                Dim oSmartPlate As IJSmartPlate
                Dim PlateType As SmartPlateTypes
                
                'check for bracket
                Set oSmartPlate = oPlatePart.object
                PlateType = oSmartPlate.SmartPlateType
                If (PlateType = spType_BRACKET) Then
                    status = objCategory.GetBracketCategoryInfo(oModelResourceMgr, LongNames, ShortNames, PriorityValues)
                Else 'no category to look up
                    GoTo CatNotFound
                End If
            Else
                GoTo CatNotFound
            End If
    End Select
    
    ' lCat is the priorityvalue, NOT the index.  So search for the value to get the appropriate index
    Dim i As Long
    
    For i = LBound(PriorityValues) To UBound(PriorityValues)
        If PriorityValues(i) = lCat Then
            GetCat = ShortNames(i)
            Exit For
        End If
    Next
    
CatNotFound:
    Set objCategory = Nothing
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetCat", Err.Description
    Err.Raise E_FAIL
End Function

'********************************************************************
' Function:
'   GetPartSide
' Abstract:
'   Calculates the part side (Port, Stbd, Center)
'********************************************************************
Private Function GetPartSide(ByVal eAxis As SDPrinPlateAxis, _
                             ByVal oLowPos As IJDPosition, _
                             ByVal oHighPos As IJDPosition) As SDSymPriority
    Dim ePartSide As SDSymPriority
    On Error GoTo label

    If Not (eAxis = Nonplanar) Then ' calculate the axis plane
        If (eAxis = Deck) Or (eAxis = Transv) Then
                If (oLowPos.y > NEG_POSITION_TOL) And _
               (oHighPos.y > NEG_POSITION_TOL) Then
                    ePartSide = Port
                ElseIf (oLowPos.y < POSITION_TOL) And _
                       (oHighPos.y < POSITION_TOL) Then
                    ePartSide = Starboard
                Else
                    ePartSide = Center
                End If
        ElseIf eAxis = Longt Then
                If (Abs(oHighPos.y) < POSITION_TOL) Then
                    ePartSide = Center
                ElseIf (oLowPos.y > NEG_POSITION_TOL) Then
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
'   GetReferencePlaneNameForStandalonePart
' Abstract:
'   Get the reference plane for a stand-alone plate part
'********************************************************************
Private Function GetReferencePlaneNameForStandalonePart(ByVal oPlatePart As StructDetailObjects.PlatePart) As String
    Dim oItem As IJNamedItem
    Dim oRefPlane As IUnknown
    Dim strFullName As String
    Dim strName As String
         
    strName = ""
    
    On Error Resume Next 'do not propagate errors getting ref plane, may not have geom yet
    Call m_oNameUtil.GetPlateRefPlane(oPlatePart.object, oRefPlane)
    If (Not (oRefPlane Is Nothing)) Then
        On Error GoTo label
        Set oItem = oRefPlane
        strFullName = oItem.Name
        Set oItem = Nothing
        Set oRefPlane = Nothing
        
        strName = m_oNameHelp.GetFilteredRefPlaneName(strFullName)
    Else
        'StrcutPlane is the ref plane for Plate system
        Dim oPlatePartPort As IJPort
        
        'Get the baseport of the plate part
        Set oPlatePartPort = oPlatePart.BasePort(BPT_Base)
        
        'Get the ref plane location
        strName = GetRefPlaneNameFromPort(oPlatePartPort, True)
        
        Set oPlatePartPort = Nothing
    End If
    
    GetReferencePlaneNameForStandalonePart = strName
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetReferencePlaneNameForStandalonePart", Err.Description
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
    Dim oRefPlane As IUnknown
    Dim strOName As String
    Dim fltResult() As String
    On Error GoTo label
    
    strOName = ""
    
    'get the ref plane from the root parent system
    Call m_oNameUtil.GetRefPlaneFromRootSystem(oSysParent, oRefPlane)
    If Not (oRefPlane Is Nothing) Then
        Set oItem = oRefPlane
        strOName = oItem.Name
        Set oItem = Nothing
    Else
        'StrcutPlane is the ref plane for Plate system
        Dim oSurfGeomType As SurGeoType
        Dim oPlate As IJPlate
        Dim oPlateUtils As IJPlateAttributes
        
        If (TypeOf oSysParent Is IJPlate) Then
            Set oPlate = oSysParent
            Set oPlateUtils = New PlateUtils
            
            oSurfGeomType = oPlateUtils.SurfaceGeoType(oPlate)
            
            'Check if the plate system is Planar
            If oSurfGeomType = Planar Then
                Dim oPlateConnectable As IJConnectable
                Dim oEnumPorts As IJElements
                Dim oPlatePort As IJPort
                
                If (TypeOf oSysParent Is IJConnectable) Then
                    'Get the Face port of the plate system
                    Set oPlateConnectable = oSysParent
                    
                    If Not oPlateConnectable Is Nothing Then
                    
                        On Error Resume Next
                        oPlateConnectable.enumPorts oEnumPorts, PortFace
                        
                        If Not oEnumPorts Is Nothing Then
                            Dim ocount As Long
                            ocount = oEnumPorts.Count
                            
                            If (ocount > 0) Then
                                Set oPlatePort = oEnumPorts.Item(1)
                                
                                'Get the ref plane location
                                strOName = GetRefPlaneNameFromPort(oPlatePort, False)
                            End If
                        End If
                    End If
                End If
                
                Set oEnumPorts = Nothing
                Set oPlateConnectable = Nothing
                Set oPlatePort = Nothing
            End If
        End If
        
        Set oPlateUtils = Nothing
        Set oPlate = Nothing
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
'   GetRefPlaneNameFromPort
' Abstract:
'   Get the ref plane from port of the plate/plate part
'********************************************************************
Private Function GetRefPlaneNameFromPort(ByVal oPort As IJPort, ByVal isStandAlonePlate As Boolean) As String
    Dim oPlatePartPlane As IJPlane
    Dim strPositionString As String
    Set oPlatePartPlane = oPort
    
    If Not oPlatePartPlane Is Nothing Then
        Dim oPlateNormalVector As New DVector
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
        
        oPlatePartPlane.GetNormal dx, dy, dz
        oPlatePartPlane.GetRootPoint Rx, Ry, Rz
        oPlateNormalVector.Set dx, dy, dz
        
        dDotx = Math.Abs(oPlateNormalVector.Dot(oXaxis))
        dDoty = Math.Abs(oPlateNormalVector.Dot(oYaxis))
        dDotz = Math.Abs(oPlateNormalVector.Dot(oZaxis))
        
        Dim iPos As Double
        If (Abs(dDotx - 1) <= 0.000001) Then
            iPos = Rx
        ElseIf (Abs(dDoty - 1) <= 0.000001) Then
            iPos = Ry
        ElseIf (Abs(dDotz - 1) <= 0.000001) Then
            iPos = Rz
        End If
        
        If (isStandAlonePlate) Then
            If (iPos < 0) Then
                iPosition = Abs(iPos)
            Else
                iPosition = Fix(iPos)
                iPosition = Abs(iPosition)
            End If
        Else
            iPosition = Round(iPos * 1000)
        End If
        
        strPositionString = CStr(iPosition)
        
        Set oPlateNormalVector = Nothing
        Set oXaxis = Nothing
        Set oYaxis = Nothing
        Set oZaxis = Nothing
    End If
    
    GetRefPlaneNameFromPort = strPositionString
    Set oPlatePartPlane = Nothing
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "GetRefPlaneNameFromPort", Err.Description
    Err.Raise E_FAIL
End Function
'********************************************************************
' Function:
'   IsLappedPlate
' Abstract:
'   Checks the stand-alone plate part to see if its a lapped plate
'********************************************************************
Private Function IsLappedPlate(ByVal oPlatePart As StructDetailObjects.PlatePart) As Boolean
    Dim oIJPlatePart As IJPlatePart
    Dim oSDUtil As StructDetailConnectionUtil
    Dim bLappedPlate As Boolean
    On Error GoTo label

    'default
    IsLappedPlate = False
    
    'if its stand-alone, see if its a lapped plate part
    If (TypeOf oPlatePart.object Is IJDStandAlonePlatePart) Then
        Set oSDUtil = New StructDetailConnectionUtil
        Set oIJPlatePart = oPlatePart.object
        
        On Error Resume Next 'do not propagate error, may not have geom yet
        IsLappedPlate = oSDUtil.IsLappedPlate(oIJPlatePart)
        Set oIJPlatePart = Nothing
        Set oSDUtil = Nothing
    ' Else not stand-alone, not a lapped plate
    End If
Exit Function

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IsLappedPlate", Err.Description
    Err.Raise E_FAIL
End Function

'********************************************************************
' Function:
'   SDPartIndexByCount
' Abstract:
'   Get next plate index from the counter index
'********************************************************************
Private Sub SDPartIndexByCount(ByVal oPlatePart As StructDetailObjects.PlatePart, _
                               ByVal oResourceMgr As IUnknown, _
                               strName As String, _
                               strResult As String)
    Dim iIndex As Long
    Dim strSel As String
    On Error GoTo label
    
    strSel = "SDPlte" + strName
    Call m_oNameHelp.SDIndexByCounter(strSel, iIndex, oResourceMgr)
 
    strResult = Format(iIndex)
Exit Sub

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "SDPartIndexByCount", Err.Description
    Err.Raise E_FAIL
End Sub

'********************************************************************
' Function:
'   SDPlateGetIndex
' Abstract:
'   Get next plate index from the positional index
'********************************************************************
Private Sub SDPlateGetIndex(ByVal oEntity As Object, _
                            ByVal oResourceMgr As IUnknown, _
                            strPos As String)
    Dim iIndex As Integer
    On Error GoTo label
    
    Call m_oNameHelp.SDGetIndex(oEntity, iIndex, oResourceMgr, m_oNameUtil)

    If iIndex <> 0 Then
        Dim alpha As String, aStrPos As String
        alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        aStrPos = Mid(alpha, (iIndex Mod 26), 1)
        If iIndex > 26 Then
            aStrPos = aStrPos + Mid(alpha, ((iIndex / 26) Mod 26), 1)
        End If
        strPos = aStrPos
    Else
        strPos = "!!"
    End If
Exit Sub

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "SDGetIndex", Err.Description
    Err.Raise E_FAIL
End Sub

'********************************************************************
' Function:
'   SetEnviron
' Abstract:
'   Extract the key environment objects from the naming parents collection
'********************************************************************
Private Sub SetEnviron(ByVal oParents As IJElements, _
                       ByRef oSysParent As IJSystem, _
                       ByRef oAsmParent As IJAssembly, _
                       ByRef oGroupMonitor As IJGroupMon)
    Dim oElement As Object
    On Error GoTo label
    
    'get all the relevant parent objects used for naming
    For Each oElement In oParents
        If (TypeOf oElement Is IJSystem) Then
            Set oSysParent = oElement
        ElseIf (TypeOf oElement Is IJAssembly) Then
            Set oAsmParent = oElement
        ElseIf (TypeOf oElement Is IJGroupMon) Then
            Set oGroupMonitor = oElement
        End If
        
        Set oElement = Nothing
    Next
Exit Sub

label:
    'log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "SetEnviron", Err.Description
    Err.Raise E_FAIL
End Sub
