VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndCutDefCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'**********************************************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : EndCutDefCM.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    06/Sep/2011 - pnalugol
'           For supporting Generic ACs for Tube standard(Bounding)Vs non-tube Standard member(Bounded)
'    19/Sep/2011 -  CM - TR-CP-202051 "EdgeToFlange and EdgeToOutSide shapes doesn't work for ShapeAtEdge"
'                                     In this file change was done in CM_IsAxisTopEdgeInsideCornerRequested()
'                                     and CM_IsAxisBottomEdgeInsideCornerRequested(). "EdgeToFlnage" codelist
'                                     value was not considered in the logic for creating CF.
'    21/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    22/Sep/2011 - svsmylav TR-202526:
'                  CMConstructTopEdgeInsideCorner and CMConstructBottomEdgeInsideCorner are modified to
'                  use top/bottom port as input to CF incase EdgeToOutside/InsideToOutsideCorner are chosen.
'    22/Sep/2011  - GH/CM - TR-201985 "Inside" answer to "ShapeAtFace" question has no effect
'                           Modified CMConstructFace(Top/Btm)Corner and CM_IsAxisFace(Top/Btm)CornerRequested
'                           Method were enhanced to detect "Inside" CF condition and act accordingly
'    27/Sep/2011  - pnalugol - Changes made in ConstructNonPenTopOrBottomPC() to fix issue related to missing PCs when bounding is a plate
'    03/Oct/2011  - GH TR-203298 NeedToDeferCornerFeatureToFlangeCut() method is modified to Get Plate Normal from Plate Goemetry
'                        instead of oWLGeomwtry which was mistakenly used.
'    19/Oct/2011 - mpulikol
'           CR-CP-203633 Performance: Increase speed of generic member assembly connections
'   22/Nov/2011 - pnalugol
'           CR-CP-205624 - Removing functionality to defer corner feature from webcut to flange cut
'    30/Nov/2011 - svsmylav TR-205302: Replaced '.Subport' method call with 'GetLateralSubPortBeforeTrim'.
'    3/Feb/2012 - CM - TR-202532 : Modified CMNeedToCompute() for updating Corner Feature when necessary.
'    5/Apr/2012 - TR-CP-212760     ShapeAtFace fails at Bottom for member ACs with TDL
'    7/4/ 2013 - gkharrin\dsmamidi
'               DI-CP-235071 Improve performance by caching edge mapping data
'               Reduced the number of calls to edgemapping rule
'               Replaced GetEdgeMapping with GetEdgeMap, From GetEdgeMap,getting edgemap data if already cached, if not call the edgemapping rule to get the same
'   21/02/2014 - TR-244709 Modified CMNeedToCompute method to update the object based on the modifications on the property page.
'    22/April/2015 - modified MDT TR-271041.Added a method to update the PCs when bounded member is rotated after AC placement
'    03/August/2015 - MDT TR-274273 Change SM orientation from vertical to horizontal- deletes PCs under GenMemACs
'   03/Nov/15 -  pkakula TR-CP-278336 Generic Assebly connection fails when bounding is split
'                Updated UpdateFlangeCutInput()
'***********************************************************************************************************************
Const m_sClassName As String = "EndCutDefCM"
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

'------------------------------------------------------------------------------
' Is PC Needed & Construct PC Methods
'------------------------------------------------------------------------------
'*********************************************************************************************
' Method      : CM_IsTopPCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsTopPCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsTopPCNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sAnswer As String
    
    GetSelectorAnswer pMemberDescription.CAO, "ConnectTop", sAnswer

    If sAnswer = "" Then
        bIsNeeded = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsBottomPCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsBottomPCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsBottomPCNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    Dim sAnswer As String
    
    GetSelectorAnswer pMemberDescription.CAO, "ConnectBottom", sAnswer

    If sAnswer = "" Then
        bIsNeeded = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsFaceOutsidePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsFaceOutsidePCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsFaceOutsidePCNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    Dim sAnswer As String
    
    GetSelectorAnswer pMemberDescription.CAO, "ConnectFaceOutside", sAnswer

    If sAnswer = "" Then
        bIsNeeded = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsFacePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsFacePCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsFacePCNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
            
    Dim sAnswer As String
    
    GetSelectorAnswer pMemberDescription.CAO, "ConnectFace", sAnswer

    If sAnswer = "" Then
        bIsNeeded = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsInsidePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsInsidePCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsInsidePCNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
            
    Dim sAnswer As String
    GetSelectorAnswer pMemberDescription.CAO, "ConnectInside", sAnswer

    If sAnswer = "" Then
        bIsNeeded = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsOutsidePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsOutsidePCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsOutsidePCNeeded"
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
            
    bIsNeeded = False
    
    Dim sError As String
'
'    Dim sAnswer As String
'    GetSelectorAnswer pMemberDescription.CAO, "ConnectOutside", sAnswer
'
'    If sAnswer = "" Then
'        bIsNeeded = False
'    Else
'        If LCase(sAnswer) = LCase("Yes") Then
'            bIsNeeded = True
'        Else
'            bIsNeeded = False
'        End If
'    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsTopInsidePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsTopInsidePCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsTopInsidePCNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
            
    Dim sAnswer As String
    GetSelectorAnswer pMemberDescription.CAO, "ConnectTopInside", sAnswer

    If sAnswer = "" Then
        bIsNeeded = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsBottomInsidePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsBottomInsidePCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsBottomInsidePCNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
            
    Dim sAnswer As String
    GetSelectorAnswer pMemberDescription.CAO, "ConnectBottomInside", sAnswer

    If sAnswer = "" Then
        bIsNeeded = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsEdgePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsEdgePCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsEdgePCNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
            
    Dim sAnswer As String
    GetSelectorAnswer pMemberDescription.CAO, "ConnectEdge", sAnswer

    If sAnswer = "" Then
        bIsNeeded = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsTopEdgePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsTopEdgePCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsTopEdgePCNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
            
    Dim sAnswer As String
    
    GetSelectorAnswer pMemberDescription.CAO, "ConnectTopEdge", sAnswer

    If sAnswer = "" Then
        bIsNeeded = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsBottomEdgePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsBottomEdgePCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsBottomEdgePCNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
            
    Dim sAnswer As String
    GetSelectorAnswer pMemberDescription.CAO, "ConnectBottomEdge", sAnswer

    If sAnswer = "" Then
        bIsNeeded = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_ConstructTopPC
' Description : Add the Top Connection selector as a child item
'
'*********************************************************************************************
Public Sub CM_ConstructTopPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    Const METHOD = m_DefinitionProgid & "::CM_ConstructTopPC"
    On Error GoTo ErrorHandler
    
    Dim sError As String
    
    Dim sMsg As String
    sMsg = "Creating Physical Connection " & "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
           
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMemberDescription.CAO
    
    Dim oBounded As Object
    Dim oBounding As Object
    
    Set oBounded = oSDO_WebCut.BoundedPort
    Set oBounding = oSDO_WebCut.BoundingPort
    
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pMemberDescription.CAO, oBounding, oBounded, lSectionAlias, bPenetratesWeb)
    Dim eTop As JXSEC_CODE
    eTop = oEdgeMapColl.Item(CStr(JXSEC_TOP))
    
    Set pObject = CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", eTop, JXSEC_TOP)

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number

End Sub

'*********************************************************************************************
' Method      : CM_ConstructFacePC
' Description : Add the Face connection selector as a child item
'
'*********************************************************************************************
Public Sub CM_ConstructFacePC(ByVal pMemberDescription As IJDMemberDescription, _
                           ByVal pResourceManager As IUnknown, _
                           ByRef pObject As Object)
    Dim sMsg As String
    sMsg = "Creating Physical Connection " & _
           "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)

    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMemberDescription.CAO
    
    Dim oBounded As Object
    Dim oBounding As Object
    
    Set oBounded = oSDO_WebCut.BoundedPort
    Set oBounding = oSDO_WebCut.BoundingPort

    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pMemberDescription.CAO, oBounding, oBounded, lSectionAlias, bPenetratesWeb)
    Dim eWebRight As JXSEC_CODE
    eWebRight = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT))

    Set pObject = CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", eWebRight, JXSEC_WEB_RIGHT)

End Sub

'*********************************************************************************************
' Method      : CM_ConstructTopFacePC
' Description : Add the Top Face connection selector as a child item
'
'*********************************************************************************************
Public Sub CM_ConstructTopFacePC(ByVal pMemberDescription As IJDMemberDescription, _
                           ByVal pResourceManager As IUnknown, _
                           ByRef pObject As Object)
    Dim sMsg As String
    sMsg = "Creating Physical Connection " & _
           "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
           
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMemberDescription.CAO
    
    Dim oBounded As Object
    Dim oBounding As Object
    
    Set oBounded = oSDO_WebCut.BoundedPort
    Set oBounding = oSDO_WebCut.BoundingPort
    
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pMemberDescription.CAO, oBounding, oBounded, lSectionAlias, bPenetratesWeb)
    Dim eWebRightTop As JXSEC_CODE
    eWebRightTop = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT_TOP))

    Set pObject = CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", eWebRightTop, JXSEC_WEB_RIGHT_TOP)

End Sub

'*********************************************************************************************
' Method      : CM_ConstructBottomFacePC
' Description : Add the Bottom Face connection selector as a child item
'
'*********************************************************************************************
Public Sub CM_ConstructBottomFacePC(ByVal pMemberDescription As IJDMemberDescription, _
                           ByVal pResourceManager As IUnknown, _
                           ByRef pObject As Object)
    Dim sMsg As String
    sMsg = "Creating Physical Connection " & _
           "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
           
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMemberDescription.CAO
    
    Dim oBounded As Object
    Dim oBounding As Object
    
    Set oBounded = oSDO_WebCut.BoundedPort
    Set oBounding = oSDO_WebCut.BoundingPort
    
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pMemberDescription.CAO, oBounding, oBounded, lSectionAlias, bPenetratesWeb)
    Dim eWebRightBottom As JXSEC_CODE
    eWebRightBottom = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT_BOTTOM))

    Set pObject = CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", eWebRightBottom, JXSEC_WEB_RIGHT_BOTTOM)

End Sub

'*********************************************************************************************
' Method      : CM_ConstructTopInsidePC
' Description : Add the Top Inside connection selector as a child item
'
'*********************************************************************************************
Public Sub CM_ConstructTopInsidePC(ByVal pMemberDescription As IJDMemberDescription, _
                           ByVal pResourceManager As IUnknown, _
                           ByRef pObject As Object)
    Dim sMsg As String
    sMsg = "Creating Physical Connection " & _
           "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
           
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMemberDescription.CAO
    
    Dim oBounded As Object
    Dim oBounding As Object
    
    Set oBounded = oSDO_WebCut.BoundedPort
    Set oBounding = oSDO_WebCut.BoundingPort

    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pMemberDescription.CAO, oBounding, oBounded, lSectionAlias, bPenetratesWeb)
    Dim eTopFlangeRightBottom As JXSEC_CODE
    eTopFlangeRightBottom = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT_BOTTOM))

    Set pObject = CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", eTopFlangeRightBottom, JXSEC_TOP_FLANGE_RIGHT_BOTTOM)

End Sub

'*********************************************************************************************
' Method      : CM_ConstructTopOutsidePC
' Description : Construct Top Outside Physical Connection
'
'*********************************************************************************************
Public Sub CM_ConstructTopOutsidePC(ByVal pMemberDescription As IJDMemberDescription, _
                           ByVal pResourceManager As IUnknown, _
                           ByRef pObject As Object)
'    Dim sMsg As String
'    sMsg = "Creating Physical Connection " & _
'           "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
'
'    Dim oEndCutMappingRule As IJEndCutEdgeMappingRule
'    Dim lSectionAlias As Long
'    Dim bPenetratesWeb As Boolean
'    Dim oEdgeMapColl As JCmnShp_CollectionAlias
'
'    Dim oSDO_WebCut As StructDetailObjects.WebCut
'    Set oSDO_WebCut = New StructDetailObjects.WebCut
'    Set oSDO_WebCut.object = pMemberDescription.CAO
'
'    Dim oBounded As Object
'    Dim oBounding As Object
'
'    Set oBounded = oSDO_WebCut.BoundedPort
'    Set oBounding = oSDO_WebCut.BoundingPort
'
'    Set oEndCutMappingRule = CreateEdgeMappingRuleSymbolInstance
'    Set oEdgeMapColl = New Collection
'    oEndCutMappingRule.GetEdgeMapping oBounding, oBounded, lSectionAlias, bPenetratesWeb, oEdgeMapColl
'
'    Dim eTopFlangeRightTop As JXSEC_CODE
'    eTopFlangeRightTop = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT_TOP))
'
'    Set pObject = CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", eTopFlangeRightTop, JXSEC_TOP_FLANGE_RIGHT_TOP)

End Sub

'*********************************************************************************************
' Method      : CM_ConstructBottomInsidePC
' Description : Add the Bottom Inside connection selector as a child item
'
'*********************************************************************************************
Public Sub CM_ConstructBottomInsidePC(ByVal pMemberDescription As IJDMemberDescription, _
                           ByVal pResourceManager As IUnknown, _
                           ByRef pObject As Object)
    Dim sMsg As String
    sMsg = "Creating Physical Connection " & _
           "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
           
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMemberDescription.CAO
    
    Dim oBounded As Object
    Dim oBounding As Object
    
    Set oBounded = oSDO_WebCut.BoundedPort
    Set oBounding = oSDO_WebCut.BoundingPort
    
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pMemberDescription.CAO, oBounding, oBounded, lSectionAlias, bPenetratesWeb)
    Dim eBottomFlangeRightTop As JXSEC_CODE
    eBottomFlangeRightTop = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT_TOP))

    Set pObject = CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", eBottomFlangeRightTop, JXSEC_BOTTOM_FLANGE_RIGHT_TOP)

End Sub

'*********************************************************************************************
' Method      : CM_ConstructBottomOutsidePC
' Description : Construct Bottom Outside Physical Connection
'
'*********************************************************************************************
Public Sub CM_ConstructBottomOutsidePC(ByVal pMemberDescription As IJDMemberDescription, _
                           ByVal pResourceManager As IUnknown, _
                           ByRef pObject As Object)
'    Dim sMsg As String
'    sMsg = "Creating Physical Connection " & _
'           "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
'
'    Dim oEndCutMappingRule As IJEndCutEdgeMappingRule
'    Dim lSectionAlias As Long
'    Dim bPenetratesWeb As Boolean
'    Dim oEdgeMapColl As JCmnShp_CollectionAlias
'
'    Dim oSDO_WebCut As StructDetailObjects.WebCut
'    Set oSDO_WebCut = New StructDetailObjects.WebCut
'    Set oSDO_WebCut.object = pMemberDescription.CAO
'
'    Dim oBounded As Object
'    Dim oBounding As Object
'
'    Set oBounded = oSDO_WebCut.BoundedPort
'    Set oBounding = oSDO_WebCut.BoundingPort
'
'    Set oEndCutMappingRule = CreateEdgeMappingRuleSymbolInstance
'    Set oEdgeMapColl = New Collection
'    oEndCutMappingRule.GetEdgeMapping oBounding, oBounded, lSectionAlias, bPenetratesWeb, oEdgeMapColl
'
'    Dim eBottomFlangeRightBottom As JXSEC_CODE
'    eBottomFlangeRightBottom = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM))
'
'    Set pObject = CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", eBottomFlangeRightBottom, JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM)

End Sub

'*********************************************************************************************
' Method      : CM_ConstructTopEdgePC
' Description : Add the Top Edge connection selector as a child item
'
'*********************************************************************************************
Public Sub CM_ConstructTopEdgePC(ByVal pMemberDescription As IJDMemberDescription, _
                           ByVal pResourceManager As IUnknown, _
                           ByRef pObject As Object)
    Dim sMsg As String
    sMsg = "Creating Physical Connection " & _
           "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
           
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMemberDescription.CAO
    
    Dim oBounded As Object
    Dim oBounding As Object
    
    Set oBounded = oSDO_WebCut.BoundedPort
    Set oBounding = oSDO_WebCut.BoundingPort
    
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pMemberDescription.CAO, oBounding, oBounded, lSectionAlias, bPenetratesWeb)
    Dim eTopFlangeRight As JXSEC_CODE
    eTopFlangeRight = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT))

    Set pObject = CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", eTopFlangeRight, JXSEC_TOP_FLANGE_RIGHT)

End Sub

'*********************************************************************************************
' Method      : CM_ConstructBottomEdgePC
' Description : Add the Bottom Edge connection selector as a child item
'
'*********************************************************************************************
Public Sub CM_ConstructBottomEdgePC(ByVal pMemberDescription As IJDMemberDescription, _
                           ByVal pResourceManager As IUnknown, _
                           ByRef pObject As Object)
    Dim sMsg As String
    sMsg = "Creating Physical Connection " & _
           "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMemberDescription.CAO
    
    Dim oBounded As Object
    Dim oBounding As Object
    
    Set oBounded = oSDO_WebCut.BoundedPort
    Set oBounding = oSDO_WebCut.BoundingPort
    
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pMemberDescription.CAO, oBounding, oBounded, lSectionAlias, bPenetratesWeb)
    Dim eBottomFlangeRight As JXSEC_CODE
    eBottomFlangeRight = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT))

    Set pObject = CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", eBottomFlangeRight, JXSEC_BOTTOM_FLANGE_RIGHT)

End Sub

'*********************************************************************************************
' Method      : CM_ConstructBottomPC
' Description : Add the Bottom connection selector as a child item
'
'*********************************************************************************************
Public Sub CM_ConstructBottomPC(ByVal pMemberDescription As IJDMemberDescription, _
                           ByVal pResourceManager As IUnknown, _
                           ByRef pObject As Object)
    Dim sMsg As String
    sMsg = "Creating Physical Connection " & _
           "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
           
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMemberDescription.CAO
    
    Dim oBounded As Object
    Dim oBounding As Object
    
    Set oBounded = oSDO_WebCut.BoundedPort
    Set oBounding = oSDO_WebCut.BoundingPort
    
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pMemberDescription.CAO, oBounding, oBounded, lSectionAlias, bPenetratesWeb)
    Dim eBottom As JXSEC_CODE
    eBottom = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM))

    Set pObject = CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", eBottom, JXSEC_BOTTOM)

End Sub

'*********************************************************************************************
' Method      : CreateModify_PhysConns
' Description :
'
'*********************************************************************************************
Public Function CreateModify_PhysConns(ByVal pMemberDescription As IJDMemberDescription, _
                                       ByVal pResourceManager As IUnknown, _
                                       strStartClass As String, _
                                       ByVal eBoundingSubPort As JXSEC_CODE, _
                                       ByVal eBoundedSubPort As JXSEC_CODE) As Object
    Const METHOD = m_DefinitionProgid & "::CreateModify_PhysConns"
    
    Dim sError As String
    ' Get Class
    Dim pWebCut As StructDetailObjects.WebCut
    ' create new web cut wrapper class
    Set pWebCut = New StructDetailObjects.WebCut
    sError = "Setting web cut Inputs"
    Set pWebCut.object = pMemberDescription.CAO
    
    ' Get Bounding Part's Port
    Dim oBoundingPart As Object ' This is the Bounding part
    Dim oBoundingPort As IJPort
    sError = "Getting Bounding object"
    Set oBoundingPart = pWebCut.Bounding ' Get Bounding part from wrapper
    Dim phelper As StructDetailObjects.Helper
    Set phelper = New StructDetailObjects.Helper
    Dim oBounding As Object
    
    If (TypeOf oBoundingPart Is IJPlate) Or (TypeOf oBoundingPart Is SPSSlabEntity) Or (TypeOf oBoundingPart Is SPSWallPart) Then
        Dim pSDOHelper As New StructDetailObjects.Helper
        Dim oLastPlatePort As IJPort
        Set oLastPlatePort = pWebCut.BoundingPort
        Set oBoundingPort = pSDOHelper.GetEquivalentLastPort(oLastPlatePort)
    ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        sError = "Setting Profile Part wrapper object to penetrating object"
        Set oProfilePart.object = oBoundingPart
        sError = "Getting Bounding object's port"
        
        If eBoundingSubPort <> -1 Then
            Set oBoundingPort = oProfilePart.SubPort(eBoundingSubPort)
        Else
        'If Bounding is Stiffener then currently set BoundingPort as Lateral(Global Port)
            Set oBoundingPort = oProfilePart.baseport(BPT_Lateral)
        End If
        
    ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_BEAM Then
        Dim oBeamPart As New StructDetailObjects.BeamPart
        sError = "Setting Beam Part wrapper object to penetrating object"
        Set oBeamPart.object = oBoundingPart
        sError = "Getting Bounding object's port"
        Set oBoundingPort = oBeamPart.SubPort(eBoundingSubPort)
        Set oBoundingPort = phelper.GetEquivalentLastPort(oBoundingPort)
    ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_MEMBER Then
        Dim oMemberPart As New StructDetailObjects.MemberPart
        sError = "Setting Member Part wrapper object to penetrating object"
        Set oMemberPart.object = oBoundingPart
        sError = "Getting Bounding object's port"
        Set oBoundingPort = GetLateralSubPortBeforeTrim(oBoundingPart, eBoundingSubPort)
        Set oBoundingPort = phelper.GetEquivalentLastPort(oBoundingPort)
        If oBoundingPort Is Nothing Then
            Dim oBoundingPortBeforeTrim As IJStructPort
            Set oBoundingPortBeforeTrim = GetLateralSubPortBeforeTrim(oBoundingPart, eBoundingSubPort)
            Dim oGraphConn As IJStructGraphConnectable
            Set oGraphConn = oBoundingPart
            Dim oACTools As AssemblyConnectionTools
            Set oACTools = New AssemblyConnectionTools
            oACTools.GetBindingPort oBoundingPart, JS_TOPOLOGY_PROXY_LFACE, oBoundingPortBeforeTrim.OperationID, oBoundingPortBeforeTrim.OperatorID, oBoundingPortBeforeTrim.ContextID, oBoundingPortBeforeTrim.SectionID, "", oBoundingPort
        End If
    End If
   
    ' Get bounded profile Port
        Dim oBoundedPort As IJPort
    If phelper.ObjectType(pWebCut.Bounded) = SDOBJECT_STIFFENER Then
        Dim oBoundedPart As StructDetailObjects.ProfilePart
        Set oBoundedPart = New StructDetailObjects.ProfilePart
        Set oBoundedPart.object = pWebCut.Bounded
        Set oBoundedPort = oBoundedPart.CutoutSubPort(pWebCut.object, eBoundedSubPort)
    ElseIf phelper.ObjectType(pWebCut.Bounded) = SDOBJECT_BEAM Then
        Dim oBoundedBeamPart As StructDetailObjects.BeamPart
        Set oBoundedBeamPart = New StructDetailObjects.BeamPart
        Set oBoundedBeamPart.object = pWebCut.Bounded
        Set oBoundedPort = oBoundedBeamPart.CutoutSubPort(pWebCut.object, eBoundedSubPort)
    ElseIf phelper.ObjectType(pWebCut.Bounded) = SDOBJECT_MEMBER Then
        Dim oStructProfilePart As IJStructProfilePart
        Dim oStructEndCutUtil As IJStructEndCutUtil
        Set oStructProfilePart = pWebCut.Bounded
        Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
        oStructEndCutUtil.GetLatePortForFeatureSegment pWebCut.object, eBoundedSubPort, oBoundedPort
    End If
        
    ' Get the web cut, make it a parent of the PC
    Dim pSystemParent As IJSystemChild ' this is the slot's ca
    sError = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
    ' Create physical connection
    If Not oBoundedPort Is Nothing Then
    
        Dim oPhysicalConnection As New PhysicalConn
        sError = "Creating Physical Connection"
        oPhysicalConnection.Create pResourceManager, _
                                   oBoundedPort, _
                                   oBoundingPort, _
                                   strStartClass, _
                                   pSystemParent, ConnectionStandard
                                   
        sError = "Setting Physical Connection to private variable"
        Set CreateModify_PhysConns = oPhysicalConnection.object
        
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number

End Function

'------------------------------------------------------------------------------
' Is Corner Feature Needed & Construct Corner Feature Methods
'------------------------------------------------------------------------------
'*********************************************************************************************
' Method      : CM_IsFaceInsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsFaceInsideCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsFaceInsideCornerNeeded"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
            
    bIsNeeded = Me.IsGenericFaceInsideCornerRequested(oMD.CAO)
    
    If bIsNeeded Then
    
        Dim oFeature As IJStructFeature
        Set oFeature = oMD.CAO
        
        If oFeature.get_StructFeatureType = SF_WebCut Then
            ' Get the alias to determine if this is a bottom or top inside corner
            Dim lSectionAlias As Long
            Dim bPenetratesWeb As Boolean
            Dim oEdgeMapColl As JCmnShp_CollectionAlias
            Dim oSDOWebCut As New StructDetailObjects.WebCut
            
            Set oSDOWebCut.object = oMD.CAO
            Set oEdgeMapColl = New Collection
            Set oEdgeMapColl = GetEdgeMap(oMD.CAO, oSDOWebCut.BoundingPort, oSDOWebCut.BoundedPort, lSectionAlias, bPenetratesWeb)
            Dim portID1 As JXSEC_CODE
            Dim portID2 As JXSEC_CODE
            
            Select Case lSectionAlias
            
                ' WebTopFlangeRight = 1,WebBuiltUpTopFlangeRight = 2, WebTopFlange = 6
                Case 1, 2, 6
                    portID1 = JXSEC_WEB_RIGHT
                    portID2 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM
                ' WebBottomFlangeRight = 3,WebBuiltUpBottomFlangeRight = 4, WebBottomFlange = 7
                Case 3, 4, 7
                    portID1 = JXSEC_WEB_RIGHT
                    portID2 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP
                ' Unexpected cases, return false
                Case Else
            End Select
            
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsFaceOutsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsFaceOutsideCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsFaceOutsideCornerNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = False
    
'    bIsNeeded = Me.IsGenericFaceTopOutsideCornerRequested(oMD.CAO)
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsFaceTopInsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsFaceTopInsideCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsFaceTopInsideCornerNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = Me.IsGenericFaceTopInsideCornerRequested(oMD.CAO)
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
End Sub

'*********************************************************************************************
' Method      : CM_IsFaceBottomInsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsFaceBottomInsideCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsFaceBottomInsideCornerNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = Me.IsGenericFaceBottomInsideCornerRequested(oMD.CAO)
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsFaceTopOutsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsFaceTopOutsideCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsFaceTopOutsideCornerNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = Me.IsGenericFaceTopOutsideCornerRequested(oMD.CAO)
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsFaceBottomOutsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsFaceBottomOutsideCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsFaceBottomOutsideCornerNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = Me.IsGenericFaceBottomOutsideCornerRequested(oMD.CAO)
        
    Exit Sub
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsEdgeInsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsEdgeInsideCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsEdgeInsideCornerNeeded"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = Me.IsGenericEdgeInsideCornerRequested(oMD.CAO)
    
    If bIsNeeded Then
    
        Dim oFeature As IJStructFeature
        Set oFeature = oMD.CAO
        
        If oFeature.get_StructFeatureType = SF_WebCut Then
            ' Get the alias to determine if this is a bottom or top inside corner
            Dim lSectionAlias As Long
            Dim bPenetratesWeb As Boolean
            Dim oEdgeMapColl As JCmnShp_CollectionAlias
            Dim oSDOWebCut As New StructDetailObjects.WebCut
            
            Set oSDOWebCut.object = oMD.CAO
            Set oEdgeMapColl = New Collection
            Set oEdgeMapColl = GetEdgeMap(oMD.CAO, oSDOWebCut.BoundingPort, oSDOWebCut.BoundedPort, lSectionAlias, bPenetratesWeb)
            Dim portID1 As JXSEC_CODE
            Dim portID2 As JXSEC_CODE
            
            Select Case lSectionAlias
            
                ' WebTopFlangeRight = 1,WebBuiltUpTopFlangeRight = 2, WebTopFlange = 6
                Case 1, 2, 6
                    portID1 = JXSEC_TOP_FLANGE_RIGHT_BOTTOM
                    portID2 = JXSEC_TOP_FLANGE_RIGHT
                ' WebBottomFlangeRight = 3,WebBuiltUpBottomFlangeRight = 4, WebBottomFlange = 7
                Case 3, 4, 7
                    portID1 = JXSEC_BOTTOM_FLANGE_RIGHT_TOP
                    portID2 = JXSEC_BOTTOM_FLANGE_RIGHT
                ' Unexpected cases, return false
                Case Else
            End Select
            
        End If
    End If
        
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsTopEdgeInsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsTopEdgeInsideCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsTopEdgeInsideCornerNeeded"

    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = Me.IsGenericTopEdgeInsideCornerRequested(oMD.CAO)
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsBottomEdgeInsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsBottomEdgeInsideCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsBottomEdgeInsideCornerNeeded"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = Me.IsGenericBottomEdgeInsideCornerRequested(oMD.CAO)
        
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsEdgeOutsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsEdgeOutsideCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsEdgeOutsideCornerNeeded"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = Me.IsGenericEdgeOutsideCornerRequested(oMD.CAO)
        
    If bIsNeeded Then
    
        Dim oFeature As IJStructFeature
        Set oFeature = oMD.CAO
        
        If oFeature.get_StructFeatureType = SF_WebCut Then
            ' Get the alias to determine if this is a bottom or top inside corner
            Dim lSectionAlias As Long
            Dim bPenetratesWeb As Boolean
            Dim oEdgeMapColl As JCmnShp_CollectionAlias
            Dim oSDOWebCut As New StructDetailObjects.WebCut
            
            Set oSDOWebCut.object = oMD.CAO
            Set oEdgeMapColl = New Collection
            Set oEdgeMapColl = GetEdgeMap(oMD.CAO, oSDOWebCut.BoundingPort, oSDOWebCut.BoundedPort, lSectionAlias, bPenetratesWeb)
            Dim portID1 As JXSEC_CODE
            Dim portID2 As JXSEC_CODE
            
            Select Case lSectionAlias
            
                ' WebTopFlangeRight = 1,WebBuiltUpTopFlangeRight = 2, WebTopFlange = 6
                Case 1, 2, 6
                    portID1 = JXSEC_TOP_FLANGE_RIGHT_TOP
                    portID2 = JXSEC_TOP_FLANGE_RIGHT
                ' WebBottomFlangeRight = 3,WebBuiltUpBottomFlangeRight = 4, WebBottomFlange = 7
                Case 3, 4, 7
                    portID1 = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM
                    portID2 = JXSEC_BOTTOM_FLANGE_RIGHT
                ' Unexpected cases, return false
                Case Else
            End Select
            
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsTopEdgeOutsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsTopEdgeOutsideCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsTopEdgeOutsideCornerNeeded"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = Me.IsGenericTopEdgeOutsideCornerRequested(oMD.CAO)
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsBottomEdgeOutsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsBottomEdgeOutsideCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsBottomEdgeOutsideCornerNeeded"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = Me.IsGenericBottomEdgeOutsideCornerRequested(oMD.CAO)
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_ConstructFaceTopCorner
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructFaceTopCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
                           
    Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_WEB_RIGHT, JXSEC_TOP)
    
End Sub

'*********************************************************************************************
' Method      : CM_ConstructFaceTopOutsideCorner
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructFaceTopOutsideCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
                           
    Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_TOP_FLANGE_RIGHT_TOP, JXSEC_WEB_RIGHT_TOP)
    
End Sub

'*********************************************************************************************
' Method      : CM_ConstructTopEdgeOutsideCorner
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructTopEdgeOutsideCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
                           
    Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_TOP_FLANGE_RIGHT_TOP, JXSEC_TOP_FLANGE_RIGHT)
    
End Sub

'*********************************************************************************************
' Method      : CM_ConstructTopEdgeInsideCorner
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructTopEdgeInsideCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

    Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_TOP_FLANGE_RIGHT_BOTTOM, JXSEC_TOP_FLANGE_RIGHT)
    
End Sub

'*********************************************************************************************
' Method      : CM_ConstructTopEdgeCorner
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructTopEdgeCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
                           
    Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_TOP, JXSEC_TOP_FLANGE_RIGHT)
    
End Sub

'*********************************************************************************************
' Method      : CM_ConstructFaceTopInsideCorner
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructFaceTopInsideCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
                           
    Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_TOP_FLANGE_RIGHT_BOTTOM, JXSEC_WEB_RIGHT)
    
End Sub

'*********************************************************************************************
' Method      : CM_ConstructFaceBottomInsideCorner
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructFaceBottomInsideCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
                           
    Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_BOTTOM_FLANGE_RIGHT_TOP, JXSEC_WEB_RIGHT)
    
End Sub

'*********************************************************************************************
' Method      : CM_ConstructBottomEdgeInsideCorner
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructBottomEdgeInsideCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
                           
    Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_BOTTOM_FLANGE_RIGHT_TOP, JXSEC_BOTTOM_FLANGE_RIGHT)
    
End Sub

'*********************************************************************************************
' Method      : CM_ConstructBottomEdgeOutsideCorner
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructBottomEdgeOutsideCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
                           
    Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM, JXSEC_BOTTOM_FLANGE_RIGHT)
    
End Sub

'*********************************************************************************************
' Method      : CM_ConstructBottomEdgeCorner
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructBottomEdgeCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
                           
    Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_BOTTOM, JXSEC_BOTTOM_FLANGE_RIGHT)
    
End Sub

'*********************************************************************************************
' Method      : CM_ConstructFaceBottomOutsideCorner
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructFaceBottomOutsideCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
                           
    Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM, JXSEC_WEB_RIGHT_BOTTOM)
    
End Sub

'*********************************************************************************************
' Method      : CM_ConstructFaceBottomCorner
' Description :
'
'*********************************************************************************************
Public Sub CM_ConstructFaceBottomCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
                           
    Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_WEB_RIGHT, JXSEC_BOTTOM)
    
End Sub

'*********************************************************************************************
' Method      : CAConstruct_MbrCornerFeature
' Description : Add the root corner selector as a child item
'
'*********************************************************************************************
Public Function CAConstruct_MbrCornerFeature(ByVal oMD As IJDMemberDescription, _
                                             ByVal oResMgr As IUnknown, _
                                             portXid1 As JXSEC_CODE, _
                                             portXid2 As JXSEC_CODE) As Object

    Dim sMETHOD As String
    Dim sError As String
    sMETHOD = "CAConstruct_MbrCornerFeature"
    
    'Retrieve Ports to be used to create Corner Feature
    sError = "Retrieve Ports to be used to create WebCut Corner Feature"
    Dim oProfFacePort As IJPort
    Dim oProfEdgePort1 As IJPort
    Dim oProfEdgePort2 As IJPort
    
    MbrCornerFeatureData ByVal oMD, portXid1, portXid2, oProfFacePort, oProfEdgePort1, oProfEdgePort2

'    ReplaceFacePortForCornerFeatureIfNeeded oMD.CAO, portXid1, portXid2, oProfFacePort

    ' Get the End cut, make it a parent of the Corner Feature
    sError = "Setting system parent to Member Description Custom Assembly"
    Dim oSystemParent As IJSystemChild
    Set oSystemParent = oMD.CAO
    
    ' Create Corner Feature
    If Not oProfFacePort Is Nothing And Not oProfEdgePort1 Is Nothing And Not oProfEdgePort2 Is Nothing Then
        
        sError = "Creating Member Corner Feature"
        Dim oSDO_CornerFeature As StructDetailObjects.CornerFeature
        Set oSDO_CornerFeature = New StructDetailObjects.CornerFeature
        oSDO_CornerFeature.Create oResMgr, _
                                  oProfFacePort, _
                                  oProfEdgePort1, _
                                  oProfEdgePort2, _
                                  "MbrEndCutCorner", _
                                  oSystemParent
    
        sError = "Returning Member CornerFeature just created"
        Set CAConstruct_MbrCornerFeature = oSDO_CornerFeature.object
    
    End If
                               
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

'------------------------------------------------------------------------------
' Is Cut Needed & Construct Cut Methods
'------------------------------------------------------------------------------
'*********************************************************************************************
' Method      : CM_IsTopWebCutNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsTopWebCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsTopWebCutNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_WebCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sTopWebCutNeeded As String
    GetSelectorAnswer pMemberDescription.CAO, "CreateTopWebCut", sTopWebCutNeeded

    If sTopWebCutNeeded = "" Then
        bIsNeeded = False
    Else
        If LCase(sTopWebCutNeeded) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsBottomWebCutNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsBottomWebCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsBottomWebCutNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_WebCut) Then
        bIsNeeded = False
        Exit Sub
    End If

    Dim sBottomWebCutNeeded As String
    GetSelectorAnswer pMemberDescription.CAO, "CreateBottomWebCut", sBottomWebCutNeeded

    If sBottomWebCutNeeded = "" Then
        bIsNeeded = False
    Else
        If LCase(sBottomWebCutNeeded) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsTopFlangeCutNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsTopFlangeCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsTopFlangeCutNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sTopFlangeCutNeeded As String
    GetSelectorAnswer pMemberDescription.CAO, "CreateTopFlangeCut", sTopFlangeCutNeeded

    If sTopFlangeCutNeeded = "" Then
        bIsNeeded = False
    Else
        If LCase(sTopFlangeCutNeeded) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsBottomFlangeCutNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsBottomFlangeCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsBottomFlangeCutNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sBottomFlangeCutNeeded As String
    GetSelectorAnswer pMemberDescription.CAO, "CreateBottomFlangeCut", sBottomFlangeCutNeeded

    If sBottomFlangeCutNeeded = "" Then
        bIsNeeded = False
    Else
        If LCase(sBottomFlangeCutNeeded) = LCase("Yes") Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_ConstructTopWebCut
' Description : Add the Bottom connection selector as a child item
'
'*********************************************************************************************
Public Sub CM_ConstructTopWebCut(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    
    Dim sMsg As String
        
    Dim lDispId As Long
    
    Dim sEndCutSel As String
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort

    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    
    sMsg = "Creating EndCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
    lDispId = pMemberDescription.dispid

    Dim oEndCutObject As Object
    Set oEndCutObject = pMemberDescription.CAO
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oEndCutObject
    
    Set oBoundedPort = oSDO_WebCut.BoundedPort
    Set oBoundingPort = oSDO_WebCut.BoundingPort
    
    ' Need to get the IJSystem Interface
    sMsg = "Retreiving Parent System for WebCut"
    If TypeOf oEndCutObject Is IJDesignParent Then
        Set oDesignParent = oEndCutObject
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If

    sMsg = "StructDetailObjects.WebCut::Create ...Creating Web Cut Feature"
    sEndCutSel = "TopWebSel"
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    oSDO_WebCut.Create pResourceManager, oBoundingPort, oBoundedPort, _
                       sEndCutSel, oSystemParent
                               
    sMsg = "Return the created Web Cut"
    Set pObject = oSDO_WebCut.object
           
End Sub

'*********************************************************************************************
' Method      : CM_ConstructBottomWebCut
' Description : Add the Bottom connection selector as a child item
'
'*********************************************************************************************
Public Sub CM_ConstructBottomWebCut(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    Dim sMsg As String

    Dim lDispId As Long
    
    Dim sEndCutSel As String
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort

    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    
    sMsg = "Creating EndCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)
    lDispId = pMemberDescription.dispid

    Dim oEndCutObject As Object
    Set oEndCutObject = pMemberDescription.CAO
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oEndCutObject
    
    Set oBoundedPort = oSDO_WebCut.BoundedPort
    Set oBoundingPort = oSDO_WebCut.BoundingPort
    
    ' Need to get the IJSystem Interface
    sMsg = "Retreiving Parent System for WebCut"
    If TypeOf oEndCutObject Is IJDesignParent Then
        Set oDesignParent = oEndCutObject
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If

    sMsg = "StructDetailObjects.WebCut::Create ...Creating Web Cut Feature"
    sEndCutSel = "BottomWebSel"
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    oSDO_WebCut.Create pResourceManager, oBoundingPort, oBoundedPort, _
                       sEndCutSel, oSystemParent
                               
    sMsg = "Return the created Web Cut"
    Set pObject = oSDO_WebCut.object
           
End Sub

'*********************************************************************************************
' Method      : CM_ConstructTopFlangeCut
' Description : Add the Bottom connection selector as a child item
'
'*********************************************************************************************
Public Sub CM_ConstructTopFlangeCut(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
                           
    Const METHOD = m_DefinitionProgid & "::CM_ConstructTopFlangeCut"
    Dim sMsg As String

    ConstructFlangeCut pMemberDescription, pResourceManager, pObject, False

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
    
End Sub

'*********************************************************************************************
' Method      : CM_ConstructBottomFlangeCut
' Description : Add the Bottom connection selector as a child item
'
'*********************************************************************************************
Public Sub CM_ConstructBottomFlangeCut(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

    Const METHOD = m_DefinitionProgid & "::CM_ConstructBottomFlangeCut"
    Dim sMsg As String

    ConstructFlangeCut pMemberDescription, pResourceManager, pObject, True

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsPhysConnsNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsPhysConnsNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsPhysConnsNeeded"
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sMsg As String
    Dim oEndCutObject As Object
    Dim sAnswer As String
    
    Set oEndCutObject = pMemberDescription.CAO
    GetSelectorAnswer oEndCutObject, "CreatePhysicalConnection", sAnswer
    
    If UCase(sAnswer) = "YES" Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_FinalConstructEndCut
' Description :
'
'*********************************************************************************************
Public Sub CM_FinalConstructEndCut(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = m_DefinitionProgid & "::CM_FinalConstructEndCut"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    EndCut_FinalConstruct pAggregatorDescription

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_MigrateAggregator
' Description :
'
'*********************************************************************************************
Public Sub CM_MigrateAggregator(pAggregatorDescription As IJDAggregatorDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = m_DefinitionProgid & "::CM_MigrateAggregator"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Mirgrating EndCut Object"
    
    MigrateEndCutObject pAggregatorDescription, pMigrateHelper

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_MigratePhysConn
' Description :
'
'*********************************************************************************************
Public Sub CM_MigratePhysConn(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = m_DefinitionProgid & "::CM_MigratePhysConn"
    On Error GoTo ErrorHandler
    
    ' The Physical Connection Inputs have been Migrated at same time with the EndCut Inputs
    ' Therefore, there should be nothing to do here
    Dim sMsg As String
    sMsg = "Migrating Physical Connection Inputs"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_CopyQuestionsFromFlangeCenterCutToTopCut
' Description :
'
'*********************************************************************************************
Public Sub CM_CopyQuestionsFromFlangeCenterCutToTopCut(pMemberDescription As IJDMemberDescription)
    Const METHOD = "MbrACDefCM::CM_CopyQuestionsFromFlangeCenterCutToTopCut"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Copy questions from flange center cut to top cut"

    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription

    pCopyAnswerHelper.CopyAnswer CUSTOMERID + "MbrEndCut.TopFlangeSel", "BottomFlange", CUSTOMERID + "MbrEndCut.RootGenFlangeSel", "BottomFlange"

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_CopyQuestionsFromFlangeCenterCutToBottomCut
' Description :
'
'*********************************************************************************************
Public Sub CM_CopyQuestionsFromFlangeCenterCutToBottomCut(pMemberDescription As IJDMemberDescription)
    Const METHOD = "MbrACDefCM::CM_CopyQuestionsFromFlangeCenterCutToBottomCut"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Copy questions from flange center cut to bottom cut"

    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription

    pCopyAnswerHelper.CopyAnswer CUSTOMERID + "MbrEndCut.BottomFlangeSel", "BottomFlange", CUSTOMERID + "MbrEndCut.RootGenFlangeSel", "BottomFlange"
        
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_CopyQuestionsToTopFlangeCut
' Description :
'
'*********************************************************************************************
Public Sub CM_CopyQuestionsToTopFlangeCut(pMemberDescription As IJDMemberDescription)
    Const METHOD = "MbrACDefCM::CM_CopyQuestionsToTopFlangeCut"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Set Top flange"

    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription

    pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.RootGenFlangeSel", "BottomFlange", "No"
        
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_CopyQuestionsToBottomFlangeCut
' Description :
'
'*********************************************************************************************
Public Sub CM_CopyQuestionsToBottomFlangeCut(pMemberDescription As IJDMemberDescription)
    Const METHOD = "MbrACDefCM::CM_CopyQuestionsToBottomFlangeCut"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Set bottom flange"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    
    pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.RootGenFlangeSel", "BottomFlange", "Yes"
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*** Corner features for Axis member end cuts
'*********************************************************************************************
' Method      : CMConstructFaceTopCorner
' Description :
'
'*********************************************************************************************
Public Sub CMConstructFaceTopCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
     
    Dim oAppConnection As Object
    AssemblyConnection_SmartItemName pMemberDescription.CAO, , oAppConnection
    
    Dim eTop As ConnectedEdgeInfo
    Dim eBtm As ConnectedEdgeInfo
    Dim eITF As ConnectedEdgeInfo
    Dim eIBF As ConnectedEdgeInfo
    Dim bPenetratesWeb As Boolean
    Dim oBoundingPort As IJPort
    Dim oBoundedPort As IJPort
    
    GetBoundingAndBounded pMemberDescription.CAO, oBoundingPort, oBoundedPort
    GetConnectedEdgeInfo oAppConnection, oBoundedPort, oBoundingPort, eTop, eBtm, eITF, eIBF, , bPenetratesWeb
    
    If bPenetratesWeb And _
       eTop.IntersectingEdge = Web_Right And _
       eITF.IntersectingEdge = Web_Right Then
        Set pObject = CreateCornerFeatureBetweenTwoEndCutsByDispID(pMemberDescription, pResourceManager, 1, 4)
    Else
        Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_WEB_RIGHT, JXSEC_TOP_FLANGE_RIGHT_BOTTOM)
    End If
   
End Sub

'*** Corner features for Axis member end cuts
'*********************************************************************************************
' Method      : CMConstructFaceBottomCorner
' Description :
'
'*********************************************************************************************
Public Sub CMConstructFaceBottomCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
                           
    Dim oAppConnection As Object
    AssemblyConnection_SmartItemName pMemberDescription.CAO, , oAppConnection
    
    Dim eTop As ConnectedEdgeInfo
    Dim eBtm As ConnectedEdgeInfo
    Dim eITF As ConnectedEdgeInfo
    Dim eIBF As ConnectedEdgeInfo
    Dim bPenetratesWeb As Boolean
    Dim oBoundingPort As IJPort
    Dim oBoundedPort As IJPort
    
    GetBoundingAndBounded pMemberDescription.CAO, oBoundingPort, oBoundedPort
    GetConnectedEdgeInfo oAppConnection, oBoundedPort, oBoundingPort, eTop, eBtm, eITF, eIBF, , bPenetratesWeb
    
    If bPenetratesWeb And _
       eBtm.IntersectingEdge = Web_Right And _
       eIBF.IntersectingEdge = Web_Right Then
        Set pObject = CreateCornerFeatureBetweenTwoEndCutsByDispID(pMemberDescription, pResourceManager, 1, 8)
    Else
        Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_WEB_RIGHT, JXSEC_BOTTOM_FLANGE_RIGHT_TOP)
    End If
    
End Sub

'*** Corner features for Axis member end cuts
'
'*********************************************************************************************
' Method      : CMConstructTopEdgeInsideCorner
' Description :
'
'*********************************************************************************************
Public Sub CMConstructTopEdgeInsideCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

    'We need to check if CF is needed on top port and top-flange-right
    Dim sParentItemName As String
    Dim sShapeAtEdge As String
    Dim oParentObj As Object
    
    Parent_SmartItemName pMemberDescription.CAO, sParentItemName, oParentObj
    
    Select Case LCase(sParentItemName)
    
    Case LCase(gsMbrAxisToEdgeAndOutSide2Edge), (gsStiffEndToMbrEdgeAndOutSide2Edge)
        GetSelectorAnswer oParentObj, "ShapeAtEdgeOverlap", sShapeAtEdge
    Case LCase(gsMbrAxisToFaceAndOutSide1Edge), LCase(gsMbrAxisToOutSideAndOutSide1Edge), LCase(gsMbrAxisToEdge), _
         LCase(gsStiffEndToMbrFaceAndOutSide1Edge), LCase(gsStiffEndToMbrOutSideAndOutSide1Edge), LCase(gsStiffEndToMbrEdge)
        GetSelectorAnswer oParentObj, "ShapeAtEdge", sShapeAtEdge
    Case LCase(gsMbrAxisToOutSideAndOutSide2Edge), LCase(gsStiffEndToMbrOutSideAndOutSide2Edge)
        GetSelectorAnswer oParentObj, "ShapeAtTopEdge", sShapeAtEdge
    End Select

    If LCase(sParentItemName) = LCase(gsMbrAxisToEdge) Then
        Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_TOP, JXSEC_TOP_FLANGE_RIGHT)
    Else
        Select Case sShapeAtEdge
            Case gsEdgeToOutside, gsInsideToOutsideCorner
                Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_TOP, JXSEC_TOP_FLANGE_RIGHT)
            Case Else
                Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_TOP_FLANGE_RIGHT_BOTTOM, JXSEC_TOP_FLANGE_RIGHT)
        End Select
    End If
End Sub

'*********************************************************************************************
' Method      : CMConstructBottomEdgeInsideCorner
' Description :
'
'*********************************************************************************************
Public Sub CMConstructBottomEdgeInsideCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    
    'We need to check if CF is needed on bottom port and bottom-flange-right
    Dim sParentItemName As String
    Dim sShapeAtEdge As String
    Dim oParentObj As Object
    
    Parent_SmartItemName pMemberDescription.CAO, sParentItemName, oParentObj
    
    Select Case LCase(sParentItemName)
    
    Case LCase(gsMbrAxisToEdgeAndOutSide2Edge), LCase(gsStiffEndToMbrEdgeAndOutSide2Edge)
        GetSelectorAnswer oParentObj, "ShapeAtEdgeOverlap", sShapeAtEdge
    Case LCase(gsMbrAxisToFaceAndOutSide1Edge), LCase(gsMbrAxisToOutSideAndOutSide1Edge), LCase(gsMbrAxisToEdge), _
         LCase(gsStiffEndToMbrFaceAndOutSide1Edge), LCase(gsStiffEndToMbrOutSideAndOutSide1Edge), LCase(gsStiffEndToMbrEdge)
        GetSelectorAnswer oParentObj, "ShapeAtEdge", sShapeAtEdge
    Case LCase(gsMbrAxisToOutSideAndOutSide2Edge), LCase(gsStiffEndToMbrOutSideAndOutSide2Edge)
        GetSelectorAnswer oParentObj, "ShapeAtBottomEdge", sShapeAtEdge
    End Select

    If LCase(sParentItemName) = LCase(gsMbrAxisToEdge) Then
        Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_BOTTOM, JXSEC_BOTTOM_FLANGE_RIGHT)
    Else
        Select Case sShapeAtEdge
            Case gsEdgeToOutside, gsInsideToOutsideCorner
                Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_BOTTOM, JXSEC_BOTTOM_FLANGE_RIGHT)
            Case Else
                Set pObject = CAConstruct_MbrCornerFeature(pMemberDescription, pResourceManager, JXSEC_BOTTOM_FLANGE_RIGHT_TOP, JXSEC_BOTTOM_FLANGE_RIGHT)
        End Select
    End If
    
End Sub

'*********************************************************************************************
' Method      : CMConstructPhysConns
' Description :
'
'*********************************************************************************************
Public Function CMConstructPhysConns(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    Const METHOD = m_DefinitionProgid & "::CMConstructPhysConns"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Creating Physical Connection " & "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
           
    Dim lDispId As Long
    
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
   
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort

    'Get bounding and bounded ports and also objects
    GetBoundingAndBounded pMemberDescription.CAO, oBoundingPort, oBoundedPort
    Set oEdgeMapColl = New Collection
    
    Set oEdgeMapColl = GetEdgeMap(pMemberDescription.CAO, oBoundingPort, oBoundedPort, lSectionAlias, bPenetratesWeb)
    Dim eBoundingSubPort As JXSEC_CODE
    Dim eBoundedSubPort As JXSEC_CODE
   
    lDispId = pMemberDescription.dispid
    If lDispId = 1 Then
        eBoundedSubPort = JXSEC_WEB_RIGHT
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT))
    ElseIf lDispId = 2 Then
        eBoundedSubPort = JXSEC_TOP_FLANGE_RIGHT
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT))
    ElseIf lDispId = 3 Then
        eBoundedSubPort = JXSEC_BOTTOM_FLANGE_RIGHT
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT))
    ElseIf lDispId = 4 Then
        eBoundedSubPort = JXSEC_TOP_FLANGE_RIGHT_BOTTOM
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT_BOTTOM))
    ElseIf lDispId = 5 Then
        eBoundedSubPort = JXSEC_BOTTOM_FLANGE_RIGHT_TOP
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT_TOP))
    ElseIf lDispId = 6 Then
        eBoundedSubPort = JXSEC_TOP
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP))
    ElseIf lDispId = 7 Then
        eBoundedSubPort = JXSEC_BOTTOM
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM))
    ElseIf lDispId = 8 Then
        eBoundedSubPort = JXSEC_TOP_FLANGE_RIGHT_TOP
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT_TOP))
    ElseIf lDispId = 9 Then
        eBoundedSubPort = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM))
    ElseIf lDispId = 10 Then
        eBoundedSubPort = JXSEC_WEB_RIGHT_TOP
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT_TOP))
    ElseIf lDispId = 11 Then
        eBoundedSubPort = JXSEC_WEB_RIGHT_BOTTOM
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT_BOTTOM))
    Else
    End If

    sMsg = "Creating Physical Connection " & _
           "...pMemberDescription.dispid = " & Str(lDispId)
           
    
    Set pObject = CreateModify_PhysConns(pMemberDescription, pResourceManager, "TeeWeld", eBoundingSubPort, eBoundedSubPort)
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

'*********************************************************************************************
' Method      : ConstructFlangeCut
' Description :
'
'*********************************************************************************************
Private Sub ConstructFlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                               ByVal pResourceManager As IUnknown, _
                               ByRef pObject As Object, _
                               IsBottomFlange As Boolean)
                           
    Const METHOD = m_DefinitionProgid & "::ConstructFlangeCut"

    Dim sMsg As String
    sMsg = "Creating EndCut ...pMemberDescription.index = " & Str(pMemberDescription.Index)

    ' --------------------------------------------------------------------
    ' Determine if it is a web or flange cut creating the child flange cut
    ' --------------------------------------------------------------------
    ' This method is used by penetrated web cuts that need to create standard flange cuts, and penetrated flange cuts
    ' that need to create top and bottom flange cuts to add more shape
    Dim oEndCutObject As Object
    Dim oInputWebCut As Object
    Set oEndCutObject = pMemberDescription.CAO
    Set oInputWebCut = pMemberDescription.CAO
    
    Dim oFeature As IJStructFeature
    Dim eFeatureType As StructFeatureTypes
    
    eFeatureType = SF_CornerFeature
    If TypeOf oEndCutObject Is IJStructFeature Then
        Set oFeature = oEndCutObject
        eFeatureType = oFeature.get_StructFeatureType
    End If
    
    ' ------------------
    ' If it is a web cut
    ' ------------------
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    Select Case eFeatureType
        Case SF_WebCut
            Dim oSDO_WebCut As New StructDetailObjects.WebCut
            Set oSDO_WebCut.object = oEndCutObject

            ' ------------------------------------------------------------------------------
            ' If flange-penetrated, or not bounded by a member, use the web cut bounded port
            ' ------------------------------------------------------------------------------
            Dim bPenetratesWeb As Boolean
            bPenetratesWeb = False
            
            If (TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic Or TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
                bPenetratesWeb = IsWebPenetrated(oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort)
            End If
            
            If (Not bPenetratesWeb) Then
                Set oBoundingPort = oSDO_WebCut.BoundingPort
                Set oBoundedPort = oSDO_WebCut.BoundedPort
            ' ---------------------------------------------------
            ' If bounded by a member, find the port it intersects
            ' ---------------------------------------------------
            Else
                'To support Generic ACs when tube is bounding
                If IsTubularMember(oSDO_WebCut.Bounding) Then
                    Set oBoundingPort = oSDO_WebCut.BoundingPort
                Else
                    Dim boundingEdgeId As eBounding_Edge
                    Dim mappedEdgeId As JXSEC_CODE
                    
                    GetNonPenetratedIntersectedEdge oSDO_WebCut.object, oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort, boundingEdgeId, mappedEdgeId, IsBottomFlange
    
                    Select Case boundingEdgeId
                        Case Above, Below, None
                            Set oBoundingPort = oSDO_WebCut.BoundingPort
                        Case Else
                            Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounding, mappedEdgeId)
                    End Select
                End If
                
                Set oBoundedPort = oSDO_WebCut.BoundedPort
            End If
    ' -----------------------------------------------------------------------------------------------
    ' If it is a flange cut, the item being created is a top/bottom cut.  Use the parent flange ports
    ' -----------------------------------------------------------------------------------------------
        Case SF_FlangeCut
            Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
            Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
            Set oSDO_FlangeCut.object = oEndCutObject
            Set oBoundedPort = oSDO_FlangeCut.BoundedPort
            Set oBoundingPort = oSDO_FlangeCut.BoundingPort
            
            '---------------------------------------------------------------------------------
            ' Use the Center Flange Cuts Input Web Cut as the Input Web Cut
            '---------------------------------------------------------------------------------
            Set oInputWebCut = oSDO_FlangeCut.WebCut
            
        Case Else
            sMsg = "Unexpected object type"
            GoTo ErrorHandler
    End Select
    
    ' ---------------------------------
    ' Set the parent/child relationship
    ' ---------------------------------
    sMsg = "Retreiving Parent System for FlangeCut"
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    If TypeOf oEndCutObject Is IJDesignParent Then
        Set oDesignParent = oEndCutObject
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If

    ' ---------------------------------------------------------------------------------
    ' If this is a web cut, we want the root generic flange selector
    ' If this is a flange cut, we want the penetrated flange selector for bottom shapes
    ' ---------------------------------------------------------------------------------
    Dim sEndCutSel As String
    
    Select Case eFeatureType
        Case SF_WebCut
            sEndCutSel = "RootGenFlangeSel"
        Case SF_FlangeCut
            If IsBottomFlange Then
                sEndCutSel = "BottomFlangeSel"
            Else
                sEndCutSel = "TopFlangeSel"
            End If
    End Select
    
    sMsg = "StructDetailObjects.FlangeCut::Create ...Creating Flange Cut Feature"
    
    Dim oHelper As New StructDetailObjects.Helper
    If Not oHelper.GetEquivalentLastPort(oBoundingPort) Is Nothing Then
        Set oBoundingPort = oHelper.GetEquivalentLastPort(oBoundingPort)
    Else
        '
    End If
    
    
    Dim oSDO_NewFlangeCut As StructDetailObjects.FlangeCut
    Set oSDO_NewFlangeCut = New StructDetailObjects.FlangeCut
    oSDO_NewFlangeCut.Create pResourceManager, oBoundingPort, oBoundedPort, oInputWebCut, _
                             sEndCutSel, oSystemParent
                               
    sMsg = "Return the created Flange Cut"

    Set pObject = oSDO_NewFlangeCut.object
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
    
End Sub

'*********************************************************************************************
' Method      : UpdateFlangeCutInputByDispID
' Description :
'
'*********************************************************************************************
Public Sub UpdateFlangeCutInputByDispID(outsideCutDispID As Long, oCenterCut As IJDMemberObjects, oFlangeCut As Object)

    Const METHOD = "EndCutDefCM.UpdateFlangeCutInputByDispID"
    
    On Error GoTo ErrorHandler
  
    Dim oOutsideCut As Object
    Set oOutsideCut = oCenterCut.ItemByDispid(outsideCutDispID)
    
    UpdateFlangeCutInput oOutsideCut, oCenterCut, oFlangeCut
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
 

End Sub

'*********************************************************************************************
' Method      : UpdateFlangeCutInput
' Description :
'
'*********************************************************************************************
Public Sub UpdateFlangeCutInput(oOutsideCut As Object, oCenterCut As Object, oFlangeCut As Object)

    Const METHOD = "EndCutDefCM.UpdateFlangeCutInput"
    
    On Error GoTo ErrorHandler
    
    ' ------------------------------------------------------------------------------------------
    ' Set the bottom cut as the input, if it exists.  Otherwise, set the center cut as the input
    ' ------------------------------------------------------------------------------------------
    Dim oFlangeInputWebCut As Object

    If Not oOutsideCut Is Nothing Then
        Set oFlangeInputWebCut = oOutsideCut
    Else
        Set oFlangeInputWebCut = oCenterCut
    End If
   
    Dim oSmartOccurance As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    Set oSmartOccurance = oFlangeCut
    Set oSmartItem = oSmartOccurance.ItemObject
    Dim oCurFlangeCut As New StructDetailObjects.FlangeCut
    
    If Not oSmartItem Is Nothing Then
    
        Dim oSDO_WebCut As New StructDetailObjects.WebCut
        Set oSDO_WebCut.object = oFlangeInputWebCut
            
        ' ------------------------------------------------------------------------------
        ' If flange-penetrated, or not bounded by a member, use the web cut bounded port
        ' ------------------------------------------------------------------------------
        Dim bPenetratesWeb As Boolean
        bPenetratesWeb = False
        
        If (TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic Or TypeOf oSDO_WebCut.Bounding Is IJProfile) Then
            bPenetratesWeb = IsWebPenetrated(oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort)
        End If
        
        Set oBoundedPort = oSDO_WebCut.BoundedPort
        
        If (Not bPenetratesWeb) Or IsTubularMember(oSDO_WebCut.Bounding) Then
            Set oBoundingPort = oSDO_WebCut.BoundingPort
        Else
            Dim sBottomFlange As String
            Dim bIsBottomFlange As Boolean
            GetSelectorAnswer oSmartOccurance, "BottomFlange", sBottomFlange
            If sBottomFlange = "No" Then
                bIsBottomFlange = False
            Else
                bIsBottomFlange = True
            End If
               
            Dim boundingEdgeId As eBounding_Edge
            Dim mappedEdgeId As JXSEC_CODE
            ' If bounded by a member, find the port it intersects
            GetNonPenetratedIntersectedEdge oSDO_WebCut.object, oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort, boundingEdgeId, mappedEdgeId, bIsBottomFlange
            Select Case boundingEdgeId
                Case Above, Below, None
                    Set oBoundingPort = oSDO_WebCut.BoundingPort
                Case Else
                    Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounding, mappedEdgeId)
            End Select
        
        End If
    Else
        Set oCurFlangeCut.object = oFlangeCut
        Set oBoundingPort = oCurFlangeCut.BoundingPort
        Set oBoundedPort = oCurFlangeCut.BoundedPort
    End If
    
    Dim oFeatureUtils As GSCADSDCreateModifyUtilities.SDFeatureUtils
    Dim oResMgr As Object
    Dim oModFlangeCut As Object

    Set oFeatureUtils = New GSCADSDCreateModifyUtilities.SDFeatureUtils
    Set oResMgr = GetResourceManagerFromObject(oFlangeCut)
    
    Set oModFlangeCut = oFeatureUtils.CreateFlangeCut(oResMgr, _
                                                      oBoundingPort, _
                                                      oBoundedPort, _
                                                      oFlangeInputWebCut, _
                                                      vbNullString, _
                                                      Nothing, _
                                                      oFlangeCut)
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

'*********************************************************************************************
' Method      : ConstructNonPenTopOrBottomPC
' Description :
'
'*********************************************************************************************
Public Function ConstructNonPenTopOrBottomPC(ByVal oMD As IJDMemberDescription, ByVal oResMgr As IUnknown, isBottom As Boolean) As Object
    Const METHOD = m_DefinitionProgid & "::ConstructTopOrBottomPC"
    On Error GoTo ErrorHandler
        
    Set ConstructNonPenTopOrBottomPC = Nothing
        
    Dim sMsg As String
    
    ' -------------------
    ' Get a helper object
    ' -------------------
    If Not TypeOf oMD.CAO Is IJStructFeature Then
        sMsg = "Input is not a feature"
        GoTo ErrorHandler
    End If
    
    Dim oStructFeature As IJStructFeature
    Dim oBoundedPart As Object
    Dim oBoundingPart As Object
    Dim oBoundingPort As IJStructPort
    
    Set oStructFeature = oMD.CAO
    
    If oStructFeature.get_StructFeatureType = SF_WebCut Then
        Dim oSDOWebCut As New StructDetailObjects.WebCut
        Set oSDOWebCut.object = oMD.CAO
        Set oBoundedPart = oSDOWebCut.Bounded
        Set oBoundingPart = oSDOWebCut.Bounding
        Set oBoundingPort = oSDOWebCut.BoundingPort
    ElseIf oStructFeature.get_StructFeatureType = SF_FlangeCut Then
        Dim oSDOFlangeCut As New StructDetailObjects.FlangeCut
        Set oSDOFlangeCut.object = oMD.CAO
        Set oBoundedPart = oSDOFlangeCut.Bounded
        Set oBoundingPart = oSDOFlangeCut.Bounding
        Set oBoundingPort = oSDOFlangeCut.BoundingPort
    Else
       sMsg = "Input is not a web or flange cut"
       GoTo ErrorHandler
    End If
    
    Dim eBoundedSubPort As JXSEC_CODE
    
    If isBottom Then
        eBoundedSubPort = JXSEC_BOTTOM
    Else
        eBoundedSubPort = JXSEC_TOP
    End If
    
    ' ------------------------------------
    ' Get the top bounded port before trim
    ' ------------------------------------
    Dim oTopPort As IJPort, oBottomPort As IJPort
    Set oTopPort = GetLateralSubPortBeforeTrim(oBoundedPart, JXSEC_TOP)
    Set oBottomPort = GetLateralSubPortBeforeTrim(oBoundedPart, JXSEC_BOTTOM)
    
    ' ---------------------
    ' If bounded by a plate
    ' ---------------------
    Dim eBoundingPort As eUSER_CTX_FLAGS

    If TypeOf oBoundingPart Is IJPlate Then
        Dim oSDOPlate As New StructDetailObjects.PlatePart
        Set oSDOPlate.object = oBoundingPart
        
        ' ---------------------------------------------------------------------
        ' If bounded by a lateral port, use the closest between base and offset
        ' ---------------------------------------------------------------------
        Set oBoundingPort = oBoundingPort
        
        If (oBoundingPort.ContextID And CTX_LATERAL_LFACE) = CTX_LATERAL_LFACE Then
            Dim oBasePort As IJPort
            Dim oOffsetPort As IJPort
            Set oBasePort = oSDOPlate.BasePortBeforeChamfer(BPT_Base)
            Set oOffsetPort = oSDOPlate.BasePortBeforeChamfer(BPT_Offset)
        
            Dim oModelUtil As IJSGOModelBodyUtilities
            Dim oPtOnTop As IJDPosition
            Dim oPtOnBounding As IJDPosition
            Dim distBase As Double
            Dim distOffset As Double
            Set oModelUtil = New SGOModelBodyUtilities
            
            If Not isBottom Then
                oModelUtil.GetClosestPointsBetweenTwoBodies oBasePort.Geometry, oTopPort.Geometry, oPtOnTop, oPtOnBounding, distBase
                oModelUtil.GetClosestPointsBetweenTwoBodies oOffsetPort.Geometry, oTopPort.Geometry, oPtOnTop, oPtOnBounding, distOffset
            Else
                oModelUtil.GetClosestPointsBetweenTwoBodies oBasePort.Geometry, oBottomPort.Geometry, oPtOnTop, oPtOnBounding, distBase
                oModelUtil.GetClosestPointsBetweenTwoBodies oOffsetPort.Geometry, oBottomPort.Geometry, oPtOnTop, oPtOnBounding, distOffset
            End If
            
            If distBase < distOffset Then
                eBoundingPort = CTX_BASE
            Else
                eBoundingPort = CTX_OFFSET
            End If
        Else
            eBoundingPort = CTX_LATERAL
        End If
    Else
        ' only expected plate for now
    End If
                               
                                                              
    If Not oBoundingPort Is Nothing Then
        If oResMgr Is Nothing Then
            sMsg = "Modifying Physical Connection " & "...dispid = " & Str(oMD.dispid)
            Modify_PhysConn oMD, eBoundingPort, eBoundedSubPort
        Else
            sMsg = "Creating Physical Connection " & "...dispid = " & Str(oMD.dispid)
            Set ConstructNonPenTopOrBottomPC = Create_PhysConn(oMD, oResMgr, "TeeWeld", eBoundingPort, eBoundedSubPort)
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Function

'*********************************************************************************************
' Method      : CM_IsGenericTopCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsTopCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsGenericTopCornerNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = Me.IsGenericTopCornerRequested(oMD.CAO)
        
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsGenericBottomCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsBottomCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsGenericBottomCornerNeeded"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = Me.IsGenericBottomCornerRequested(oMD.CAO)
        
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_AxisIsFaceTopCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsAxisFaceTopCornerRequested(oFeature As IJStructFeature, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsAxisFaceTopCornerRequested"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    bIsNeeded = False  'Initial value, make it to True only for valid cases
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort

    Dim oBounded As Object
    Dim oBounding As Object

    'Get bounding and bounded ports and also objects
    GetBoundingAndBounded oFeature, oBoundingPort, oBoundedPort, _
        oBounding, oBounded
    
    'Check if bounding member is tube and exit
    If IsTubularMember(oBounding) Then
        bIsNeeded = False
        Exit Sub
    End If

'1st check:- findout cases for which this CF is not applicable and exit
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName oFeature, sParentItemName, oParentObj
    Select Case LCase(sParentItemName)
        Case LCase(gsMbrAxisToEdge), LCase(gsMbrAxisToOnMember), _
                LCase(gsMbrAxisToOutSideAndOutSideNoEdge), _
             LCase(gsStiffEndToMbrEdge), LCase(gsStiffEndToOnMember), _
                LCase(gsStiffEndToMbrOutSideAndOutSideNoEdge)
            Exit Sub                         '*** Exit ***
            '
    End Select
    
'2nd check:-
'Need to process no-edge, 1-edge and 2-edge cases. However, for 2-edge case we
'need only one step so it is handled first:
'    Dim oBoundedPort As IJPort
'    Dim oBoundingPort As IJPort

'    Dim oBounded As New StructDetailObjects.MemberPart
'    Dim oBounding As New StructDetailObjects.MemberPart

    'Get bounding and bounded ports and also objects
'    GetBoundingAndBounded oFeature, oBoundingPort, oBoundedPort, _
'        oBounding, oBounded

    Select Case LCase(sParentItemName)
        Case LCase(gsMbrAxisToEdgeAndEdge), LCase(gsMbrAxisToEdgeAndOutSide2Edge), _
                LCase(gsMbrAxisToOutSideAndOutSide2Edge), _
             LCase(gsStiffEndToMbrEdgeAndEdge), LCase(gsStiffEndToMbrEdgeAndOutSide2Edge), _
                LCase(gsStiffEndToMbrOutSideAndOutSide2Edge)
            'Two-edge cases: check for top inside corner selector answer
            Dim sTopInsideCornerAnswer As String
            GetSelectorAnswer oParentObj, "TopInsideCorner", sTopInsideCornerAnswer
            If Len(sTopInsideCornerAnswer) > 0 And StrComp(sTopInsideCornerAnswer, gsNone, vbTextCompare) <> 0 Then
                bIsNeeded = True
            End If
        Case Else
            
            'Get the Edge Mapping and Section Alias
            Dim lSectionAlias As Long
            Dim bPenetratesWeb As Boolean
            Dim oEdgeMapColl As JCmnShp_CollectionAlias
            Set oEdgeMapColl = New Collection
        
            Set oEdgeMapColl = GetEdgeMap(oParentObj, oBoundingPort, oBoundedPort, lSectionAlias, bPenetratesWeb)
            Dim cTopOrWL As ConnectedEdgeInfo
            Dim cBottomOrWR As ConnectedEdgeInfo
            Dim cInsideTForTFL As ConnectedEdgeInfo
            Dim cInsideBForTFR As ConnectedEdgeInfo
    
            GetConnectedEdgeInfo oParentObj, oBoundedPort, oBoundingPort, cTopOrWL, cBottomOrWR, cInsideTForTFL, cInsideBForTFR
                
            'One-edge cases: check for inside corner selector answer
            Dim bFurtherCheckOneEdgeCase As Boolean
            bFurtherCheckOneEdgeCase = False
            If StrComp(sParentItemName, gsMbrAxisToEdgeAndOutSide1Edge, vbTextCompare) = 0 Or _
                StrComp(sParentItemName, gsMbrAxisToFaceAndEdge, vbTextCompare) = 0 Or _
                StrComp(sParentItemName, gsStiffEndToMbrEdgeAndOutSide1Edge, vbTextCompare) = 0 Or _
                StrComp(sParentItemName, gsStiffEndToMbrFaceAndEdge, vbTextCompare) = 0 Then
'               1a. top bounding edge touches top bounded edge
                If (cTopOrWL.IntersectingEdge = eBounding_Edge.Top_Flange_Right Or _
                    cInsideTForTFL.IntersectingEdge = eBounding_Edge.Top_Flange_Right) Then
                    bFurtherCheckOneEdgeCase = True
                ElseIf (cTopOrWL.IntersectingEdge = Above Or cTopOrWL.IntersectingEdge = Top) And _
                       cInsideTForTFL.IntersectingEdge = Web_Right Then
                    Select Case lSectionAlias
                        Case 1, 2, 5, 6, 13, 19
                        'WebTopFlangeRight = 1, WebBuiltUpTopFlangeRight = 2, WebTopAndBottomRightFlanges = 5,
                        'WebTopFlange = 6, WebTopAndBottomFlanges = 13, TwoWebsBetweenFlanges = 19
                            bFurtherCheckOneEdgeCase = True
                    End Select
                End If
            ElseIf StrComp(sParentItemName, gsMbrAxisToFaceAndOutSide1Edge, vbTextCompare) = 0 Or _
                   StrComp(sParentItemName, gsStiffEndToMbrFaceAndOutSide1Edge, vbTextCompare) = 0 Then
                '1b. For following need to confirm that top bounding edge touches top bounded edge
                'if Top & InsideTopFlange is hitting FACE then Question it needs to respond is "ShapeAtFace" but not "InsideCorner"
                'OR if (extreme case)Top is Coplanar with Top_Flange_Right_Bottom even then Question it needs to respond is "ShapeAtFace" but not "InsideCorner"
                'Else it needs to respond "InsideCorner" answer
                If cTopOrWL.IntersectingEdge = Web_Right And cInsideTForTFL.IntersectingEdge = Web_Right Then
                    bFurtherCheckOneEdgeCase = False
                ElseIf cTopOrWL.IntersectingEdge = Top_Flange_Right And cTopOrWL.CoplanarEdge = Top_Flange_Right_Bottom Then
                    bFurtherCheckOneEdgeCase = False
                Else
                    bFurtherCheckOneEdgeCase = True
                End If
            ElseIf StrComp(sParentItemName, gsMbrAxisToOutSideAndOutSide1Edge, vbTextCompare) = 0 Or _
                   StrComp(sParentItemName, gsStiffEndToMbrOutSideAndOutSide1Edge, vbTextCompare) = 0 Then
                '1c. For following need to confirm that top bounding edge touches top bounded edge
                Select Case lSectionAlias
                    Case 1, 2, 5, 6, 13, 19
                    'WebTopFlangeRight = 1, WebBuiltUpTopFlangeRight = 2, WebTopAndBottomRightFlanges = 5,
                    'WebTopFlange = 6, WebTopAndBottomFlanges = 13, TwoWebsBetweenFlanges = 19
                        bFurtherCheckOneEdgeCase = True
                End Select
            End If
            If bFurtherCheckOneEdgeCase = True Then
                Dim sInsideCornerAnswer As String
                GetSelectorAnswer oParentObj, "InsideCorner", sInsideCornerAnswer
                If Len(sInsideCornerAnswer) > 0 And StrComp(sInsideCornerAnswer, gsNone, vbTextCompare) <> 0 Then
                    bIsNeeded = True
                End If
            Else
'               Remaining are No-edge cases: check for top shape at face selector answer
                Dim sTopShapeAtFaceAnswer As String
                GetSelectorAnswer oParentObj, "ShapeAtFace", sTopShapeAtFaceAnswer
                If sTopShapeAtFaceAnswer = "" Then
                    GetSelectorAnswer oParentObj, "TopShapeAtFace", sTopShapeAtFaceAnswer
                End If
                If Len(sTopShapeAtFaceAnswer) > 0 And StrComp(sTopShapeAtFaceAnswer, gsNone, vbTextCompare) <> 0 And _
                    cTopOrWL.IntersectingEdge = eBounding_Edge.Web_Right Then
                    If bPenetratesWeb = True Then
                        'Check if top flange exists and intersects with web of bounding
                        Dim bTopFL As Boolean
                        Dim bTopFR As Boolean
                        Dim bBtmFL As Boolean
                        Dim bBtmFR As Boolean
                        
                        CrossSection_Flanges oBounded, bTopFL, bBtmFL, bTopFR, bBtmFR
                        If (bTopFL) Or (bTopFR) Then
                            If StrComp(sTopShapeAtFaceAnswer, gsInside, vbTextCompare) = 0 Then
                                'Shape at Face: Inside
                                bIsNeeded = True
                            End If
                        End If
                    Else
                        'Shape at Face: Cope and Snipe
                        bIsNeeded = True
                    End If
                End If
            End If
    End Select
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsAxisFaceBottomCornerRequested
' Description :
'
'*********************************************************************************************
Public Sub CM_IsAxisFaceBottomCornerRequested(oFeature As IJStructFeature, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsAxisFaceBottomCornerRequested"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    bIsNeeded = False  'Initial value, make it to True only for valid cases
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort

    Dim oBounded As Object
    Dim oBounding As Object

    'Get bounding and bounded ports and also objects
    GetBoundingAndBounded oFeature, oBoundingPort, oBoundedPort, oBounding, oBounded
    
    'Check if bounding member is tube and exit
    If IsTubularMember(oBounding) Then
        bIsNeeded = False
        Exit Sub
    End If

    '1st check:- findout cases for which this CF is not applicable and exit
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName oFeature, sParentItemName, oParentObj
    Select Case LCase(sParentItemName)
        Case LCase(gsMbrAxisToEdge), LCase(gsMbrAxisToOnMember), _
                LCase(gsMbrAxisToOutSideAndOutSideNoEdge), _
             LCase(gsStiffEndToMbrEdge), LCase(gsStiffEndToOnMember), _
                LCase(gsStiffEndToMbrOutSideAndOutSideNoEdge)
            Exit Sub                         '*** Exit ***
            '
    End Select
    
'2nd check:-
'Need to process no-edge, 1-edge and 2-edge cases. However, for 2-edge case we
'need only one step so it is handled first:
'    Dim oBoundedPort As IJPort
'    Dim oBoundingPort As IJPort
'
'    Dim oBounded As New StructDetailObjects.MemberPart
'    Dim oBounding As New StructDetailObjects.MemberPart

    'Get bounding and bounded ports and also objects
'    GetBoundingAndBounded oFeature, oBoundingPort, oBoundedPort, _
        oBounding, oBounded

    Select Case LCase(sParentItemName)
        Case LCase(gsMbrAxisToEdgeAndEdge), LCase(gsMbrAxisToEdgeAndOutSide2Edge), _
                LCase(gsMbrAxisToOutSideAndOutSide2Edge), _
             LCase(gsStiffEndToMbrEdgeAndEdge), LCase(gsStiffEndToMbrEdgeAndOutSide2Edge), _
                LCase(gsStiffEndToMbrOutSideAndOutSide2Edge)
            'Two-edge cases: check for top inside corner selector answer
            Dim sBottomInsideCornerAnswer As String
            GetSelectorAnswer oParentObj, "BottomInsideCorner", sBottomInsideCornerAnswer
            If Len(sBottomInsideCornerAnswer) > 0 And StrComp(sBottomInsideCornerAnswer, gsNone, vbTextCompare) <> 0 Then
                bIsNeeded = True
            End If
        Case Else
            
            'Get the Edge Mapping and Section Alias
            Dim lSectionAlias As Long
            Dim bPenetratesWeb As Boolean
            Dim oEdgeMapColl As JCmnShp_CollectionAlias
            Set oEdgeMapColl = New Collection
            Set oEdgeMapColl = GetEdgeMap(oParentObj, oBoundingPort, oBoundedPort, lSectionAlias, bPenetratesWeb)
            Dim cTopOrWL As ConnectedEdgeInfo
            Dim cBottomOrWR As ConnectedEdgeInfo
            Dim cInsideTForTFL As ConnectedEdgeInfo
            Dim cInsideBForTFR As ConnectedEdgeInfo
    
            GetConnectedEdgeInfo oParentObj, oBoundedPort, oBoundingPort, cTopOrWL, cBottomOrWR, _
                cInsideTForTFL, cInsideBForTFR
                
            'One-edge cases: check for inside corner selector answer
            Dim bFurtherCheckOneEdgeCase As Boolean
            bFurtherCheckOneEdgeCase = False
            If StrComp(sParentItemName, gsMbrAxisToEdgeAndOutSide1Edge, vbTextCompare) = 0 Or _
                StrComp(sParentItemName, gsMbrAxisToFaceAndEdge, vbTextCompare) = 0 Or _
               StrComp(sParentItemName, gsStiffEndToMbrEdgeAndOutSide1Edge, vbTextCompare) = 0 Or _
                StrComp(sParentItemName, gsStiffEndToMbrFaceAndEdge, vbTextCompare) = 0 Then
'               1a. Bottom bounding edge touches bottom bounded edge
                If (cBottomOrWR.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right Or _
                    cInsideBForTFR.IntersectingEdge = eBounding_Edge.Bottom_Flange_Right) Then
                    bFurtherCheckOneEdgeCase = True
                ElseIf (cBottomOrWR.IntersectingEdge = Below Or cBottomOrWR.IntersectingEdge = Bottom) And _
                       cInsideBForTFR.IntersectingEdge = Web_Right Then
                    Select Case lSectionAlias
                        Case 3, 4, 5, 7, 13, 19
                        'WebBottomFlangeRight = 3,WebBuiltUpBottomFlangeRight = 4,WebTopAndBottomRightFlanges = 5,WebBottomFlange = 7,WebTopAndBottomFlanges = 13,TwoWebsBetweenFlanges = 19
                        bFurtherCheckOneEdgeCase = True
                    End Select
                End If
            ElseIf StrComp(sParentItemName, gsMbrAxisToFaceAndOutSide1Edge, vbTextCompare) = 0 Or _
                   StrComp(sParentItemName, gsStiffEndToMbrFaceAndOutSide1Edge, vbTextCompare) = 0 Then
                '1b. For following need to confirm that top bounding edge touches top bounded edge
                'if Btm & InsideBtmFlange is hitting FACE then Question it needs to respond is "ShapeAtFace" but not "InsideCorner"
                'OR if (extreme case)Btm is Coplanar with Bottom_Flange_Right_Top even then Question it needs to respond is "ShapeAtFace" but not "InsideCorner"
                'Else it needs to respond "InsideCorner" answer
                If cBottomOrWR.IntersectingEdge = Web_Right And cInsideBForTFR.IntersectingEdge = Web_Right Then
                    bFurtherCheckOneEdgeCase = False
                ElseIf cBottomOrWR.IntersectingEdge = Bottom_Flange_Right And cBottomOrWR.CoplanarEdge = Bottom_Flange_Right_Top Then
                    bFurtherCheckOneEdgeCase = False
                Else
                    bFurtherCheckOneEdgeCase = True
                End If
            ElseIf StrComp(sParentItemName, gsMbrAxisToOutSideAndOutSide1Edge, vbTextCompare) = 0 Or _
                   StrComp(sParentItemName, gsStiffEndToMbrOutSideAndOutSide1Edge, vbTextCompare) = 0 Then
'               1c. For following need to confirm that top bounding edge touches top bounded edge
                Select Case lSectionAlias
                    Case 3, 4, 5, 7, 13, 19
                    'WebBottomFlangeRight = 3,WebBuiltUpBottomFlangeRight = 4,WebTopAndBottomRightFlanges = 5,WebBottomFlange = 7,WebTopAndBottomFlanges = 13,TwoWebsBetweenFlanges = 19
                     bFurtherCheckOneEdgeCase = True
                End Select
            End If
            If bFurtherCheckOneEdgeCase = True Then
                Dim sInsideCornerAnswer As String
                GetSelectorAnswer oParentObj, "InsideCorner", sInsideCornerAnswer
                If Len(sInsideCornerAnswer) > 0 And StrComp(sInsideCornerAnswer, gsNone, vbTextCompare) <> 0 Then
                    bIsNeeded = True
                End If
            Else
'               Remaining are No-edge cases: check for top shape at face selector answer
                Dim sBottomShapeAtFaceAnswer As String
                GetSelectorAnswer oParentObj, "ShapeAtFace", sBottomShapeAtFaceAnswer
                If sBottomShapeAtFaceAnswer = "" Then
                    GetSelectorAnswer oParentObj, "BottomShapeAtFace", sBottomShapeAtFaceAnswer
                End If

                If Len(sBottomShapeAtFaceAnswer) > 0 And StrComp(sBottomShapeAtFaceAnswer, gsNone, vbTextCompare) <> 0 And _
                    cBottomOrWR.IntersectingEdge = eBounding_Edge.Web_Right Then
                    If bPenetratesWeb = True Then
                        'Check if top flange exists and intersects with web of bounding
                        Dim bTopFL As Boolean
                        Dim bTopFR As Boolean
                        Dim bBtmFL As Boolean
                        Dim bBtmFR As Boolean
                        
                        CrossSection_Flanges oBounded, bTopFL, bBtmFL, bTopFR, bBtmFR
                        If (bBtmFL) Or (bBtmFR) Then
                            If StrComp(sBottomShapeAtFaceAnswer, gsInside, vbTextCompare) = 0 Then
                                'Shape at Face: Inside
                                bIsNeeded = True
                            End If
                        End If
                    Else
                        'Shape at Face: Cope and Snipe
                        bIsNeeded = True
                    End If
                End If
            End If
    End Select
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : ComputeGeneric_IsTopEdgeInsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Function IsGenericTopEdgeInsideCornerRequested(oFeature As IJStructFeature) As Boolean
    Const METHOD = m_DefinitionProgid & "::CM_IsTopEdgeInsideCornerNeeded"

    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim sAnswer As String
    GetSelectorAnswer oFeature, "TopEdgeInsideCorner", sAnswer
    
    If sAnswer = "" Then
        IsGenericTopEdgeInsideCornerRequested = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            IsGenericTopEdgeInsideCornerRequested = True
        Else
            IsGenericTopEdgeInsideCornerRequested = False
        End If
    End If

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Function

'*********************************************************************************************
' Method      : ComputeGeneric_IsBottomEdgeInsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Function IsGenericBottomEdgeInsideCornerRequested(oFeature As IJStructFeature) As Boolean
    Const METHOD = m_DefinitionProgid & "::IsGenericBottomEdgeInsideCornerRequested"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim sAnswer As String
    GetSelectorAnswer oFeature, "BottomEdgeInsideCorner", sAnswer
    
    If sAnswer = "" Then
        IsGenericBottomEdgeInsideCornerRequested = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            IsGenericBottomEdgeInsideCornerRequested = True
        Else
            IsGenericBottomEdgeInsideCornerRequested = False
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Function

'*********************************************************************************************
' Method      : IsGenericFaceBottomInsideCornerRequested
' Description :
'
'*********************************************************************************************
Public Function IsGenericFaceBottomInsideCornerRequested(oFeature As IJStructFeature) As Boolean
    Const METHOD = m_DefinitionProgid & "::IsGenericFaceBottomInsideCornerRequested"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim sAnswer As String
    GetSelectorAnswer oFeature, "FaceBottomInsideCorner", sAnswer
    
    If sAnswer = "" Then
        IsGenericFaceBottomInsideCornerRequested = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            IsGenericFaceBottomInsideCornerRequested = True
        Else
            IsGenericFaceBottomInsideCornerRequested = False
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Function

'*********************************************************************************************
' Method      : IsGenericFaceTopInsideCornerRequested
' Description :
'
'*********************************************************************************************
Public Function IsGenericFaceTopInsideCornerRequested(oFeature As IJStructFeature) As Boolean
    Const METHOD = m_DefinitionProgid & "::IsGenericFaceTopInsideCornerRequested"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim sAnswer As String
    GetSelectorAnswer oFeature, "FaceTopInsideCorner", sAnswer
    
    If sAnswer = "" Then
        IsGenericFaceTopInsideCornerRequested = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            IsGenericFaceTopInsideCornerRequested = True
        Else
            IsGenericFaceTopInsideCornerRequested = False
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Function

'*********************************************************************************************
' Method      : IsGenericTopCornerRequested
' Description :
'
'*********************************************************************************************
Public Function IsGenericTopCornerRequested(oFeature As IJStructFeature) As Boolean
    Const METHOD = m_DefinitionProgid & "::IsGenericTopCornerRequested"
    On Error GoTo ErrorHandler
    
    Dim sError As String
    Dim sAnswer As String
    
    GetSelectorAnswer oFeature, "TopCorner", sAnswer

    If sAnswer = "" Then
        IsGenericTopCornerRequested = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            IsGenericTopCornerRequested = True
        Else
            IsGenericTopCornerRequested = False
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Function

Public Function IsGenericBottomCornerRequested(oFeature As IJStructFeature) As Boolean
    Const METHOD = m_DefinitionProgid & "::IsGenericBottomCornerRequested"
    On Error GoTo ErrorHandler
    
    Dim sError As String
    Dim sAnswer As String
    
    GetSelectorAnswer oFeature, "BottomCorner", sAnswer

    If sAnswer = "" Then
        IsGenericBottomCornerRequested = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            IsGenericBottomCornerRequested = True
        Else
            IsGenericBottomCornerRequested = False
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Function

Public Function IsGenericEdgeInsideCornerRequested(oFeature As IJStructFeature) As Boolean
    Const METHOD = m_DefinitionProgid & "::IsEdgeInsideCornerRequested"
    On Error GoTo ErrorHandler
    Dim sError As String
    Dim sAnswer As String

    GetSelectorAnswer oFeature, "EdgeInsideCorner", sAnswer
    
    If sAnswer = "" Then
        IsGenericEdgeInsideCornerRequested = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            IsGenericEdgeInsideCornerRequested = True
        Else
            IsGenericEdgeInsideCornerRequested = False
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Function

'*********************************************************************************************
' Method      : IsGenericFaceInsideCornerRequested
' Description :
'
'*********************************************************************************************
Public Function IsGenericFaceInsideCornerRequested(oFeature As IJStructFeature) As Boolean

    Const METHOD = m_DefinitionProgid & "::IsGenericFaceInsideCornerRequested"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim sAnswer As String
    GetSelectorAnswer oFeature, "FaceInsideCorner", sAnswer
    
    If sAnswer = "" Then
        IsGenericFaceInsideCornerRequested = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            IsGenericFaceInsideCornerRequested = True
        Else
            IsGenericFaceInsideCornerRequested = False
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
End Function

'*********************************************************************************************
' Method      : IsGenericFaceTopOutsideCornerRequested
' Description :
'
'*********************************************************************************************
Public Function IsGenericFaceTopOutsideCornerRequested(oFeature As IJStructFeature) As Boolean
    Const METHOD = m_DefinitionProgid & "::IsGenericFaceTopOutsideCornerRequested"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim sAnswer As String
    GetSelectorAnswer oFeature, "FaceTopOutsideCorner", sAnswer
    
    If sAnswer = "" Then
        IsGenericFaceTopOutsideCornerRequested = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            IsGenericFaceTopOutsideCornerRequested = True
        Else
            IsGenericFaceTopOutsideCornerRequested = False
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Function

'*********************************************************************************************
' Method      : IsGenericFaceBottomOutsideCornerRequested
' Description :
'
'*********************************************************************************************
Public Function IsGenericFaceBottomOutsideCornerRequested(oFeature As IJStructFeature) As Boolean
    Const METHOD = m_DefinitionProgid & "::IsFaceBottomOutsideCornerRequested"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim sAnswer As String
    GetSelectorAnswer oFeature, "FaceBottomOutsideCorner", sAnswer
    
    If sAnswer = "" Then
        IsGenericFaceBottomOutsideCornerRequested = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            IsGenericFaceBottomOutsideCornerRequested = True
        Else
            IsGenericFaceBottomOutsideCornerRequested = False
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Function

'*********************************************************************************************
' Method      : IsGenericEdgeOutsideCornerRequested
' Description :
'
'*********************************************************************************************
Public Function IsGenericEdgeOutsideCornerRequested(oFeature As IJStructFeature) As Boolean
    Const METHOD = m_DefinitionProgid & "::IsGenericEdgeOutsideCornerRequested"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim sAnswer As String

    GetSelectorAnswer oFeature, "EdgeOutsideCorner", sAnswer
    
    If sAnswer = "" Then
        IsGenericEdgeOutsideCornerRequested = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            IsGenericEdgeOutsideCornerRequested = True
        Else
            IsGenericEdgeOutsideCornerRequested = False
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Function

'*********************************************************************************************
' Method      : IsGenericTopEdgeOutsideCornerRequested
' Description :
'
'*********************************************************************************************
Public Function IsGenericTopEdgeOutsideCornerRequested(oFeature As IJStructFeature) As Boolean
    Const METHOD = m_DefinitionProgid & "::IsGenericTopEdgeOutsideCornerRequested"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim sAnswer As String
    GetSelectorAnswer oFeature, "TopEdgeOutsideCorner", sAnswer
    
    If sAnswer = "" Then
        IsGenericTopEdgeOutsideCornerRequested = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            IsGenericTopEdgeOutsideCornerRequested = True
        Else
            IsGenericTopEdgeOutsideCornerRequested = False
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Function

'*********************************************************************************************
' Method      : IsGenericBottomEdgeOutsideCornerRequested
' Description :
'
'*********************************************************************************************
Public Function IsGenericBottomEdgeOutsideCornerRequested(oFeature As IJStructFeature) As Boolean
    Const METHOD = m_DefinitionProgid & "::IsGenericBottomEdgeOutsideCornerRequested"
    
    On Error GoTo ErrorHandler
    Dim sError As String
    
    Dim sAnswer As String
    GetSelectorAnswer oFeature, "BottomEdgeOutsideCorner", sAnswer
    
    If sAnswer = "" Then
        IsGenericBottomEdgeOutsideCornerRequested = False
    Else
        If LCase(sAnswer) = LCase("Yes") Then
            IsGenericBottomEdgeOutsideCornerRequested = True
        Else
            IsGenericBottomEdgeOutsideCornerRequested = False
        End If
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number
 
End Function

'******************************************************************************************
' Method:
' CMNeedToCompute
'
' Description: Updates the respective smart occurences/also helps in modifying
'              the inputs of CF(if needed)
'
' *******************************************************************************************
Public Sub CMNeedToCompute(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "EndCutDefCM::CMNeedToCompute"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Updating a smart occurence"
    
    If Not pObject Is Nothing Then
        If TypeOf pObject Is IJStructFeature Then
            Dim oCF As IJStructFeature
            Dim oNewFacePort As IJPort
            Dim oNewEdgePort1 As IJPort
            Dim oNewEdgePort2 As IJPort
            Dim bModifyCornerFeature As Boolean
            Dim oACObject As Object
            Dim sACItemName As String
            
            Set oCF = pObject
            bModifyCornerFeature = False
            Dim OSC As IJSmartOccurrence
            
            
            If oCF.get_StructFeatureType = SF_CornerFeature Then
                bModifyCornerFeature = IsCornerFeatureInputsModificatonNeeded(pPropertyDescriptions, oCF, _
                                                                oNewFacePort, oNewEdgePort1, oNewEdgePort2)
                '--------------------------------------------------------------------
                'For specific cases in Mbr-Mbr EC and Stiff-Mbr EC we need to modify
                'CF ports otherwise atleast update the existing only for Stiff AC's
                '--------------------------------------------------------------------
                If bModifyCornerFeature Then
                    Dim oSDFeatureUtils As New SDFeatureUtils
                    oSDFeatureUtils.CreateCornerCut GetResourceManagerFromObject(pPropertyDescriptions.CAO), _
                                                                oNewFacePort, oNewEdgePort1, oNewEdgePort2, _
                                                                "MbrEndCutCorner", pPropertyDescriptions.CAO, pObject
                    Set OSC = pObject
                    OSC.Update
                Else
                    '--------------------------------------------------------------------
                    'for atleast stiffner AC we need to update Smart-Occ since the semantic doesnt trigger
                    'CF rules when there is any modification on Stiffener End Cuts(but this is not
                    'the case with Mbr End Cuts).So below code neeeds to run everytime for Stiff AC
                    '--------------------------------------------------------------------
                    AssemblyConnection_SmartItemName pObject, sACItemName, oACObject
                    'Only Update if the Assembly Connection is of Type Stiffener AC.
                    If GetMbrAssemblyConnectionType(oACObject) = eACType.ACType_Bounded Then
                        Set OSC = pObject
                        Dim oSelectionRuleInfo As IJSelectionRuleInfo
                        Set oSelectionRuleInfo = oACObject
                        Dim eSelectionModificationType As SelectionRuleModificationType
                        eSelectionModificationType = oSelectionRuleInfo.IsSelectionRuleModified
                        If eSelectionModificationType = SelectionRule_CurrentRuleModified Then
                            'selection rule is modified, need to update the object on Interface IJStructGeometry
                            OSC.Update
                        End If
                    End If
                End If
            End If
        End If
    End If
  
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub


'*********************************************************************************************
' Method      : CM_IsPCWithPlateNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsPCWithPlateNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)

    Const METHOD = m_DefinitionProgid & "::CM_IsPCWithPlateNeeded"
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sError As String
    Dim sACItemName As String
    Dim oSmartOccurance As IJSmartOccurrence
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim oSDO_PlatePart As StructDetailObjects.PlatePart
    Dim oBoundingPart As Object
    Dim oBaseSurface As IJSurfaceBody
    Dim oOffsetSurface As IJSurfaceBody
    Dim oWebLeftSurface As IJSurfaceBody
    Dim oWebRightSurface As IJSurfaceBody
    Dim oSGOModelBodyUtilities As SGOModelBodyUtilities
    Dim oClosestPosition1 As IJDPosition
    Dim oClosestPosition2 As IJDPosition
    Dim oClosestPosition3 As IJDPosition
    Dim dMinDistToWebLeft As Double
    Dim dMinDistToWebRight As Double
    Dim dLength As Double
    Dim WLNormal As IJDVector
    Dim WRNormal As IJDVector
    Dim oNormalToBase As IJDVector
    Dim oNormalToOffset As IJDVector
    Dim dWLDotProduct As Double
    Dim dWRDotProduct As Double
    
    bIsNeeded = False
    
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMemberDescription.CAO
    
    Set oClosestPosition1 = New DPosition
    Set oClosestPosition2 = New DPosition
    Set oClosestPosition3 = New DPosition
    Set oSGOModelBodyUtilities = New SGOModelBodyUtilities
    
    AssemblyConnection_SmartItemName pMemberDescription.CAO, sACItemName
    
     If pMemberDescription.dispid = 8 Then
     
        If sACItemName = "Generic_2B_01" And TypeOf oSDO_WebCut.Bounding Is IJPlatePart Then
       
            Set oBoundingPart = oSDO_WebCut.Bounding
            Set oSDO_PlatePart = New StructDetailObjects.PlatePart
            Set oSDO_PlatePart.object = oBoundingPart
            
            Set oBaseSurface = oSDO_PlatePart.baseport(BPT_Base).Geometry
            Set oWebLeftSurface = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_WEB_LEFT).Geometry
            Set oWebRightSurface = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_WEB_RIGHT).Geometry
            
            oSGOModelBodyUtilities.GetClosestPointsBetweenTwoBodies oBaseSurface, oWebLeftSurface, oClosestPosition1, oClosestPosition2, dMinDistToWebLeft
            oSGOModelBodyUtilities.GetClosestPointsBetweenTwoBodies oBaseSurface, oWebRightSurface, oClosestPosition1, oClosestPosition3, dMinDistToWebRight
            
            
            oBaseSurface.GetNormalFromPosition oClosestPosition1, oNormalToBase
            dLength = oNormalToBase.Length
            oNormalToBase.Length = -dLength
            
             Set WLNormal = New DVector
             Set WRNormal = New DVector
             
             WLNormal.Set oClosestPosition1.x - oClosestPosition2.x, oClosestPosition1.y - oClosestPosition2.y, oClosestPosition1.z - oClosestPosition2.z
             WRNormal.Set oClosestPosition1.x - oClosestPosition3.x, oClosestPosition1.y - oClosestPosition3.y, oClosestPosition1.z - oClosestPosition3.z
             dWLDotProduct = WLNormal.Dot(oNormalToBase)
             dWRDotProduct = WRNormal.Dot(oNormalToBase)
             
             If dMinDistToWebLeft < dMinDistToWebRight Then
                If dWLDotProduct < 0# Then
                        bIsNeeded = True
                ElseIf dWLDotProduct > 0# And dMinDistToWebLeft <= 0.003 Then
                        bIsNeeded = True
                End If
             ElseIf dMinDistToWebRight < dMinDistToWebLeft Then
                If dWRDotProduct < 0# Then
                    bIsNeeded = True
                ElseIf dWRDotProduct > 0# And dMinDistToWebRight <= 0.003 Then
                        bIsNeeded = True
                End If
             Else
                bIsNeeded = True
             End If
        End If
             
     ElseIf pMemberDescription.dispid = 9 Then
    
        If sACItemName = "Generic_2B_01" And TypeOf oSDO_WebCut.Bounding Is IJPlatePart Then
            
            Set oBoundingPart = oSDO_WebCut.Bounding
            Set oSDO_PlatePart = New StructDetailObjects.PlatePart
            Set oSDO_PlatePart.object = oBoundingPart
                
            Set oOffsetSurface = oSDO_PlatePart.baseport(BPT_Offset).Geometry
            Set oWebLeftSurface = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_WEB_LEFT).Geometry
            Set oWebRightSurface = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_WEB_RIGHT).Geometry
            
            oSGOModelBodyUtilities.GetClosestPointsBetweenTwoBodies oOffsetSurface, oWebLeftSurface, oClosestPosition1, oClosestPosition2, dMinDistToWebLeft
            oSGOModelBodyUtilities.GetClosestPointsBetweenTwoBodies oOffsetSurface, oWebRightSurface, oClosestPosition1, oClosestPosition3, dMinDistToWebRight
            
            
            oOffsetSurface.GetNormalFromPosition oClosestPosition1, oNormalToOffset
            dLength = oNormalToOffset.Length
            oNormalToOffset.Length = -dLength
             
             Set WLNormal = New DVector
             Set WRNormal = New DVector
             
             WLNormal.Set oClosestPosition1.x - oClosestPosition2.x, oClosestPosition1.y - oClosestPosition2.y, oClosestPosition1.z - oClosestPosition2.z
             WRNormal.Set oClosestPosition1.x - oClosestPosition3.x, oClosestPosition1.y - oClosestPosition3.y, oClosestPosition1.z - oClosestPosition3.z
             dWLDotProduct = WLNormal.Dot(oNormalToOffset)
             dWRDotProduct = WRNormal.Dot(oNormalToOffset)
            
             If dMinDistToWebLeft < dMinDistToWebRight Then
                If dWLDotProduct < 0# Then
                        bIsNeeded = True
                ElseIf dWLDotProduct > 0# And dMinDistToWebLeft <= 0.003 Then
                        bIsNeeded = True
                End If
             ElseIf dMinDistToWebRight < dMinDistToWebLeft Then
                If dWRDotProduct < 0# Then
                    bIsNeeded = True
                ElseIf dWRDotProduct > 0# And dMinDistToWebRight <= 0.003 Then
                        bIsNeeded = True
                End If
             Else
                bIsNeeded = True
             End If
        End If
     ElseIf pMemberDescription.dispid = 10 Then
            bIsNeeded = False
     ElseIf pMemberDescription.dispid = 11 Then
            bIsNeeded = False
     End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number

End Sub



'*********************************************************************************************
' Method      : CM_ConstructPCForGussetPlateAndMbr
'
' Description : Creates Physical connection between Flange of Bounded member and Gusset plate face
'
'*********************************************************************************************
Public Function CM_ConstructPCForGussetPlateAndMbr(ByVal pMemberDescription As IJDMemberDescription, _
                                      ByVal pResourceManager As IUnknown, _
                                      ByRef pObject As Object)
    Const METHOD = m_DefinitionProgid & "::CM_ConstructPCForGussetPlateAndMbr"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Creating Physical Connection " & "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
           

    Dim oSDO_PhysicalConn As StructDetailObjects.PhysicalConn
    Set oSDO_PhysicalConn = New StructDetailObjects.PhysicalConn


    Dim oEndCutObject As Object
    Set oEndCutObject = pMemberDescription.CAO

    'Fail if the endcutobject passing is not feature

    If Not TypeOf oEndCutObject Is IJStructFeature Then
        sMsg = "EndCut is not IJStructFeature ...Type:" & TypeName(oEndCutObject)
        GoTo ErrorHandler
    End If

    Dim oBoundedPart As Object
    Dim oBoundingPort As IJPort
    Dim oBoundedPort As IJPort
    Dim oBoundingPart As Object
    Dim oBounded_SplitAxisPort As ISPSSplitAxisPort
    Dim oBounding_SplitAxisPort As ISPSSplitAxisPort

    Dim oStructFeature As IJStructFeature
    Dim eStructFeature As StructFeatureTypes

    Set oStructFeature = oEndCutObject
    eStructFeature = oStructFeature.get_StructFeatureType

    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut

    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut

    Set oSDO_WebCut.object = oEndCutObject

    sMsg = "Getting Bounded objects from WebCut"

    Set oBoundedPart = oSDO_WebCut.Bounded
    Set oBoundingPart = oSDO_WebCut.Bounding

    Set oBoundedPort = oSDO_WebCut.BoundedPort
    Set oBoundingPort = oSDO_WebCut.BoundingPort

    If TypeOf oSDO_WebCut.BoundedPort Is ISPSSplitAxisPort Then
    Set oBounded_SplitAxisPort = oSDO_WebCut.BoundedPort
    End If

    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim oPartObject As Object
    Dim lStatus As Long

    InitEndCutConnectionData oBounded_SplitAxisPort, oBoundingPort, oBoundedData, oBoundingData, lStatus, sMsg
    Set oPartObject = oBoundedData.MemberPart

    Dim oFeature As Object
    Set oFeature = oSDO_WebCut.object

    Dim oCutoutBndedPort As IJPort

    Set oCutoutBndedPort = CutoutSubPort(oPartObject, oFeature, -1)

    Dim oHelper As StructDetailObjects.Helper
    Dim oLastBoundingPort As IJPort

    Dim oBasePort As IJPort
    Dim oOffsetPort As IJPort
    Dim oSDO_PlatePart As StructDetailObjects.PlatePart

    Set oSDO_PlatePart = New StructDetailObjects.PlatePart
    Set oSDO_PlatePart.object = oBoundingPart
    Set oBasePort = oSDO_PlatePart.baseport(BPT_Base)
    Set oOffsetPort = oSDO_PlatePart.baseport(BPT_Offset)
    Set oHelper = New StructDetailObjects.Helper

    If pMemberDescription.dispid = 8 Then
        Set oLastBoundingPort = oHelper.GetEquivalentLastPort(oBasePort)
    ElseIf pMemberDescription.dispid = 9 Then
        Set oLastBoundingPort = oHelper.GetEquivalentLastPort(oOffsetPort)
    End If

    If oLastBoundingPort Is Nothing Then
        Set oLastBoundingPort = oBoundingPort
    End If

    Dim oSystemParent As IJSystemChild
    Set oSystemParent = oEndCutObject

    oSDO_PhysicalConn.Create pResourceManager, oCutoutBndedPort, oLastBoundingPort, _
                              "TeeWeld", oSystemParent, ConnectionStandard
    
    Set pObject = oSDO_PhysicalConn.object
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Function

Public Function Construct_PhysConn(oMemberDescription As IJDMemberDescription, _
                                oResourceManager As IUnknown, _
                                sRootSmartClass As String, _
                                lMappedBdingPortEdgeId As Long, _
                                lBoundedSubPortEdgeId As Long) As Object

                                
Const METHOD = m_DefinitionProgid & "::Construct_PhysConn"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Set Construct_PhysConn = Nothing

    ' ----------------------------------------------------------
    ' Create StructDetailObjects wrapper class for given End Cut
    ' ----------------------------------------------------------
    sMsg = "Setting StructDeatilObjects Wrapper Class for EndCut"
    
    Dim oEndCutObject As Object
    Set oEndCutObject = oMemberDescription.CAO
    
    ' ------------------------------------------
    ' Fail, if the input object is not a feature
    ' ------------------------------------------
    If Not TypeOf oEndCutObject Is IJStructFeature Then
        sMsg = "EndCut is not IJStructFeature ...Type:" & TypeName(oEndCutObject)
        GoTo ErrorHandler
    End If
        
    sMsg = "Trying to retrieve End Cut Bounded Port for creating PC"
    
    Dim oBoundedPort As IJPort
    Dim oFeature As IJStructFeature
    Dim eFeatureType As StructFeatureTypes
    Dim oBoundedPart As Object
    Dim eEndCutType As eEndCutTypes
    
    Set oFeature = oEndCutObject
    eFeatureType = oFeature.get_StructFeatureType
    
    Select Case eFeatureType
        Case SF_WebCut
            Dim oWebCut As New StructDetailObjects.WebCut
            Set oWebCut.object = oFeature
            Set oBoundedPort = oWebCut.BoundedPort
            Set oBoundedPart = oWebCut.Bounded
            eEndCutType = WebCut
        Case SF_FlangeCut
            Dim oFlangeCut As New StructDetailObjects.FlangeCut
            Set oFlangeCut.object = oFeature
            Set oBoundedPort = oFlangeCut.BoundedPort
            Set oBoundedPart = oFlangeCut.Bounded
            If oFlangeCut.IsTopFlange Then
                eEndCutType = FlangeCutTop
            Else
                eEndCutType = FlangeCutBottom
            End If
        Case Else
            GoTo ErrorHandler
    End Select
    
    Dim oBoundingPCPort As IJPort
    Dim oHelper As StructDetailObjects.Helper
    Dim lEdgeId As Long
    
    sMsg = "From Mapped Port Collection retrieving the needed Bounding Port based on lEdgeID"
    
    Dim oMappedPortsColl As Collection
    ' ---------------------------------------------------------
    ' Get Mapped Bounding Port, from Mapped COllection of Ports
    ' from AC, based on given lEdgeID
    ' ---------------------------------------------------------
    GetMultiBoundingEdgeMap oEndCutObject, eEndCutType, oMappedPortsColl
    
    Set oBoundingPCPort = GetMappedPortFromGivenEdgeID(oMappedPortsColl, lMappedBdingPortEdgeId)
    
    If oBoundingPCPort Is Nothing Then
        sMsg = "Could not retrieve Bounding Port from Mapped Ports for Given Edge ID:" & lMappedBdingPortEdgeId
        GoTo ErrorHandler
    End If
    
    Set oHelper = New StructDetailObjects.Helper
    Set oBoundingPCPort = oHelper.GetEquivalentLastPort(oBoundingPCPort)
    
    sMsg = "Get Bounded Late Sub Port for PC creation based on lEdgeID"
    ' ---------------------------------------------------------
    ' Get Bounded Late Sub Port for PC creation based on lEdgeID
    ' ---------------------------------------------------------
    Dim oStructProfilePart As IJStructProfilePart
    Dim oStructEndCutUtil As IJStructEndCutUtil
    Dim oBoundedPCPort As IJPort
    
    Set oStructProfilePart = oBoundedPart
    Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
    oStructEndCutUtil.GetLatePortForFeatureSegment oEndCutObject, lBoundedSubPortEdgeId, oBoundedPCPort
    
    Dim pSystemParent As IJSystemChild
    Set pSystemParent = oEndCutObject

    ' Create physical connection
    If Not oBoundedPCPort Is Nothing Then

        Dim oPhysicalConnection As New PhysicalConn
        sMsg = "Create Physical Connection"
        oPhysicalConnection.Create oResourceManager, _
                                   oBoundedPCPort, _
                                   oBoundingPCPort, _
                                   sRootSmartClass, _
                                   pSystemParent, ConnectionStandard

        
        Set Construct_PhysConn = oPhysicalConnection.object

    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

Public Function Construct_CornerFeature(oMemberDescription As IJDMemberDescription, _
                                oResourceManager As IUnknown, _
                                sRootSmartClass As String, _
                                lEdgePortId1 As Long, _
                                lEdgePortId2 As Long) As Object

                                
Const METHOD = m_DefinitionProgid & "Construct_CornerFeature """
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Creating the Corner Feature"
        
    Dim oProfFacePort As IJPort
    Dim oProfEdgePort1 As IJPort
    Dim oProfEdgePort2 As IJPort
    
    sMsg = "Getting the valid ports for Corner Feature based on Edge ID's passed on End Cut object"
    MemberCornerFeatureDataByObject ByVal oMemberDescription.CAO, lEdgePortId1, lEdgePortId2, oProfFacePort, oProfEdgePort1, oProfEdgePort2
    
    ' Get the End cut, make it a parent of the Corner Feature
    sMsg = "Setting system parent to Member Description Custom Assembly"
    Dim oSystemParent As IJSystemChild
    Set oSystemParent = oMemberDescription.CAO
    
    ' Create Corner Feature
    If Not oProfFacePort Is Nothing And Not oProfEdgePort1 Is Nothing And Not oProfEdgePort2 Is Nothing Then
        
        sMsg = "Creating Member Corner Feature"
        Dim oSDO_CornerFeature As StructDetailObjects.CornerFeature
        Set oSDO_CornerFeature = New StructDetailObjects.CornerFeature
        oSDO_CornerFeature.Create oResourceManager, _
                                  oProfFacePort, _
                                  oProfEdgePort1, _
                                  oProfEdgePort2, _
                                  sRootSmartClass, _
                                  oSystemParent
    
        sMsg = "Returning Member CornerFeature just created"
        Set Construct_CornerFeature = oSDO_CornerFeature.object
    
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function


Public Sub MemberCornerFeatureDataByObject(ByVal oEndCutObject As Object, _
                              portXid1 As Long, _
                              portXid2 As Long, _
                              oFacePort As IJPort, _
                              oEdgePort1 As IJPort, _
                              oEdgePort2 As IJPort)
    Dim sMETHOD As String
    Dim sError As String
    sMETHOD = m_DefinitionProgid & " :: MemberCornerFeatureDataByObject"
    
    On Error GoTo ErrorHandler
    
    'Determine feature type and fill bounded and bounding ports
    Dim oFeature As IJStructFeature
    Dim eFeatureType As StructFeatureTypes
    Dim oEndCutBoundedPort As IJPort
    Dim oBoundedPart As Object

    Set oFeature = oEndCutObject
    eFeatureType = oFeature.get_StructFeatureType
    Select Case eFeatureType
        Case SF_WebCut
            Dim oWebCut As New StructDetailObjects.WebCut
            Set oWebCut.object = oFeature
            Set oEndCutBoundedPort = oWebCut.BoundedPort
            Set oBoundedPart = oWebCut.Bounded
        Case SF_FlangeCut
            Dim oFlangeCut As New StructDetailObjects.FlangeCut
            Set oFlangeCut.object = oFeature
            Set oEndCutBoundedPort = oFlangeCut.BoundedPort
            Set oBoundedPart = oFlangeCut.Bounded
        Case Else
            GoTo ErrorHandler
    End Select
    
    ' Verify the following objects are valid
    If (oEndCutBoundedPort.Connectable Is Nothing) Then
        sError = "oEndCutBoundedPort.Connectable Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    If (oBoundedPart Is Nothing) Then
        sError = "oBoundedPart Object is not Valid : is NOTHING"
        GoTo ErrorHandler
    End If
    
    If eFeatureType = SF_WebCut Then
        'Web penetrated case
        Set oFacePort = GetLateralSubPortBeforeTrim(oEndCutBoundedPort.Connectable, JXSEC_WEB_LEFT)
    Else
        'Flange penetrated case: check if this CF is on top and bottom flange
        Dim bBottomFlange As Boolean
        Dim strAnswer As String
        
        GetSelectorAnswer oEndCutObject, "BottomFlange", strAnswer
        If strAnswer = "Yes" Then
            Set oFacePort = GetLateralSubPortBeforeTrim(oEndCutBoundedPort.Connectable, JXSEC_BOTTOM)
        Else
            Set oFacePort = GetLateralSubPortBeforeTrim(oEndCutBoundedPort.Connectable, JXSEC_TOP)
        End If
    End If
    
    GetMbrEndCutCornerPorts oEndCutObject, oBoundedPart, portXid1, portXid2, oFacePort, oEdgePort1, oEdgePort2
    
    Exit Sub
ErrorHandler:
      Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

'******************************************************************************************
' Method:
' CMNeedtoComputePhysicalConn
'
' Description: Replaces the old ports with new ones
' *******************************************************************************************

Public Sub CMNeedtoComputePhysicalConn(pPropertyDescriptions As IJDPropertyDescription, pObject As Object, lBoundingEdgeID As Long, _
                                                                            lBoundedEdgeID As Long)
                                            
    Const METHOD = "EndCutDefCM::CMNeedtoComputePC"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Modifying the PC "
            
    Dim oSD_PhysicalConn As New StructDetailObjects.PhysicalConn
    Dim oStructFeature As IJStructFeature
    Set oSD_PhysicalConn.object = pObject
    Dim sBottomFlange As String
    Dim oSD_WebCut As New StructDetailObjects.WebCut
    Dim eEndCutType As eEndCutTypes
    ' Getting the endcut type
    If TypeOf pPropertyDescriptions.CAO Is IJStructFeature Then
        Set oStructFeature = pPropertyDescriptions.CAO
        Set oSD_WebCut.object = pPropertyDescriptions.CAO
        If oStructFeature.get_StructFeatureType = SF_WebCut Then
            eEndCutType = WebCut
        ElseIf oStructFeature.get_StructFeatureType = SF_FlangeCut Then
            GetSelectorAnswer pPropertyDescriptions.CAO, "BottomFlange", sBottomFlange
            If sBottomFlange = gsNo Then
                eEndCutType = FlangeCutTop
            Else
                eEndCutType = FlangeCutBottom
            End If
        End If
    End If
    ' get the mapped port collection based on endcut type
    Dim oMappedPortColl As JCmnShp_CollectionAlias
    GetMultiBoundingEdgeMap pPropertyDescriptions.CAO, eEndCutType, oMappedPortColl
    Dim oBoundingPort As IJPort
     
    Set oBoundingPort = GetMappedPortFromGivenEdgeID(oMappedPortColl, lBoundingEdgeID)
    Dim oBoundedPort As IJPort
    Dim oStructProfilePart As IJStructProfilePart
    Dim oStructEndCutUtil As IJStructEndCutUtil

    Set oStructProfilePart = oSD_WebCut.Bounded
    Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
    'Get the bounded edge ID and the corresponding port from endcut
    oStructEndCutUtil.GetLatePortForFeatureSegment pPropertyDescriptions.CAO, lBoundedEdgeID, oBoundedPort

    If Not oBoundingPort Is oSD_PhysicalConn.Port2 Then '  Checking whether the previous and new bounding ports are same or not
        Dim oConnectionDefinition As GSCADSDCreateModifyUtilities.IJSDConnectionDefinition
        Set oConnectionDefinition = New GSCADSDCreateModifyUtilities.SDConnectionUtils
        oConnectionDefinition.ReplacePhysicalConnectionPort pObject, _
                                                      oSD_PhysicalConn.Port2, oBoundingPort
        oConnectionDefinition.ReplacePhysicalConnectionPort pObject, _
                                                     oSD_PhysicalConn.Port1, oBoundedPort
    End If
Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
'******************************************************************************************
' Method:
' CM_NeedToComputePenPC
'
' Description: Replaces the old ports with new ones
' Optional Arguments Usage :
'        HasTopEdge = true indicates that the bounding member has only top edge.
'        HasBtmEdge = true indicates that bounding member has only Btm edge
'        If both HasTopEdge And HasBtmEdge are true, then it represents that the bounding member has both top and btm edge
'        HasFace indicates that bounded member faces to face which has no edges.
' *******************************************************************************************
Public Sub CM_NeedToComputePenPC(pPropertyDescriptions As IJDPropertyDescription, pObject As Object, Optional HasTopEdge As Boolean = False, _
                                                            Optional HasBtmEdge As Boolean = False, Optional HasFace As Boolean = False)
                                            
    Const METHOD = "CM_NeedToComputePenPC"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Modifying the PC "
            
    If TypeOf pObject Is IJStructPhysicalConnection Then
     
        Dim oSmartOccurance As IJSmartOccurrence
        Dim oSmartItem As IJSmartItem
        
        Dim oSD_PhysicalConn As New StructDetailObjects.PhysicalConn
        
        
        Set oSmartOccurance = pObject
        Set oSmartItem = oSmartOccurance.ItemObject
        
        Dim oMemDesc As IJDMemberDescription
        
        If Not oSmartItem Is Nothing Then
            Dim oPCObject As IJStructPhysicalConnection
            Set oPCObject = pObject
             
            ' Inputs can be nothing in the case as explained below.
            ' For Ex, when user changes the corner feature question Create Bottom Inside corner to yes.
            ' And changed the CF smart item to edge to face rathole and where the material is cut by CF.Inputs of the bottom inside
            ' physical connection will be no more in that case. Inorder to avoid the error and asserts in those cases,
            'just safely exiting the sub if inputs are less than 2
            Dim AppConn As IJAppConnection
            Set AppConn = oSmartOccurance
            Dim ppenumPorts As IJElements
            AppConn.enumPorts ppenumPorts
            If ppenumPorts.Count < 2 Then Exit Sub

            Dim lDispId As Long
            'Get the correct dispid of the object
            Set oMemDesc = GetMemberDescriptionFromPropertyDescs(pPropertyDescriptions, pObject)
            lDispId = oMemDesc.dispid
            Set oSD_PhysicalConn.object = pObject
            Dim oBoundingSubPort As JXSEC_CODE
            Dim oBoundedSubPort As JXSEC_CODE
            
            ' get the bounding and bounded sub ports based on the dispid
            If HasTopEdge = True And HasBtmEdge = True Then
                GetEdgeIDsForPC_TopAndBtmEdge pPropertyDescriptions, lDispId, oBoundingSubPort, oBoundedSubPort
            ElseIf HasTopEdge = True And HasBtmEdge = False Then
                GetEdgeIDsForPC_TopEdge pPropertyDescriptions, lDispId, oBoundingSubPort, oBoundedSubPort
            ElseIf HasTopEdge = False And HasBtmEdge = True Then
                GetEdgeIDsForPC_BtmEdge pPropertyDescriptions, lDispId, oBoundingSubPort, oBoundedSubPort
            ElseIf HasFace = True Then
                GetEdgeIDsForPC_Face pPropertyDescriptions, lDispId, oBoundingSubPort, oBoundedSubPort
            End If
            
            Dim oSD_WebCut As New StructDetailObjects.WebCut
            Set oSD_WebCut.object = pPropertyDescriptions.CAO
            
            Dim phelper As StructDetailObjects.Helper
            Set phelper = New StructDetailObjects.Helper
            
            Dim oBoundingPort As IJPort
            Dim oBoundedPort As IJPort
            
            If phelper.ObjectType(oSD_WebCut.Bounding) = SDOBJECT_MEMBER Then
                Dim oMemberPart As New StructDetailObjects.MemberPart
                Set oMemberPart.object = oSD_WebCut.Bounding
                
                Set oBoundingPort = GetLateralSubPortBeforeTrim(oSD_WebCut.Bounding, oBoundingSubPort)
                If Not oBoundingPort Is Nothing Then
                    Set oBoundingPort = phelper.GetEquivalentLastPort(oBoundingPort)
                End If
                If oBoundingPort Is Nothing Then
                    Dim oBoundingPortBeforeTrim As IJStructPort
                    Set oBoundingPortBeforeTrim = GetLateralSubPortBeforeTrim(oSD_WebCut.Bounding, oBoundingSubPort)
                    Dim oGraphConn As IJStructGraphConnectable
                    Set oGraphConn = oSD_WebCut.Bounding
                    Dim oACTools As AssemblyConnectionTools
                    Set oACTools = New AssemblyConnectionTools
                    oACTools.GetBindingPort oSD_WebCut.Bounding, JS_TOPOLOGY_PROXY_LFACE, oBoundingPortBeforeTrim.OperationID, oBoundingPortBeforeTrim.OperatorID, oBoundingPortBeforeTrim.ContextID, oBoundingPortBeforeTrim.SectionID, "", oBoundingPort
                End If
            ElseIf phelper.ObjectType(oSD_WebCut.Bounding) = SDOBJECT_STIFFENER Then
                Dim oProfilePart As New StructDetailObjects.ProfilePart
                sMsg = "Setting Profile Part wrapper object to penetrating object"
                Set oProfilePart.object = oSD_WebCut.Bounding
                sMsg = "Getting Bounding object's port"
                Set oBoundingPort = oProfilePart.baseport(BPT_Lateral)
            Else
                Exit Sub
            End If
            
            If phelper.ObjectType(oSD_WebCut.Bounded) = SDOBJECT_MEMBER Then
                Dim oStructProfilePart As IJStructProfilePart
                Dim oStructEndCutUtil As IJStructEndCutUtil
                Set oStructProfilePart = oSD_WebCut.Bounded
                Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
                oStructEndCutUtil.GetLatePortForFeatureSegment oSD_WebCut.object, oBoundedSubPort, oBoundedPort
            Else
                Exit Sub
            End If
            
            Dim oBoundingPort_old As IJPort
            Dim oBoundedPort_old As IJPort
            Set oBoundingPort_old = oSD_PhysicalConn.Port2
            Set oBoundedPort_old = oSD_PhysicalConn.Port1
        
            Dim oConnectionDefinition As GSCADSDCreateModifyUtilities.IJSDConnectionDefinition
            Set oConnectionDefinition = New GSCADSDCreateModifyUtilities.SDConnectionUtils
            If Not oBoundingPort_old Is oBoundingPort Then
               
                oConnectionDefinition.ReplacePhysicalConnectionPort pObject, _
                                                              oSD_PhysicalConn.Port2, oBoundingPort
                                                              
            End If
            
            If Not oBoundedPort_old Is oBoundedPort Then
            
                oConnectionDefinition.ReplacePhysicalConnectionPort pObject, _
                                                            oSD_PhysicalConn.Port1, oBoundedPort
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
'******************************************************************************************
' Method:
' CM_NeedToComputeNonPenPC
'
' Description: Replaces the old ports with new ones
' *******************************************************************************************
Public Sub CM_NeedToComputeNonPenPC(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
                                            
    Const METHOD = "CM_NeedToComputeNonPenPC"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Modifying the PC "
            
    If TypeOf pObject Is IJStructPhysicalConnection Then
     
        Dim oSmartOccurance As IJSmartOccurrence
        Dim oSmartItem As IJSmartItem
        
        Set oSmartOccurance = pObject
        Set oSmartItem = oSmartOccurance.ItemObject
        
        If Not oSmartItem Is Nothing Then
            ' --------------------------------------------
            ' Find out if this is the top or bottom flange
            ' --------------------------------------------
            Dim oSD_PhysicalConn As StructDetailObjects.PhysicalConn
            Set oSD_PhysicalConn = New PhysicalConn
            
            ' Inputs can be nothing in the case as explained below.
            ' For Ex, when user changes the corner feature question Create Bottom Inside corner to yes.
            ' And changed the CF smart item to edge to face rathole and where the material is cut by CF.Inputs of the bottom inside
            ' physical connection will be nothing in that case. Inorder to avoid the error and asserts in those cases, just safely exiting the sub if inputs are nothing
            
            Dim pPortColl As IJElements
            Dim AppConnection As IJAppConnection
            Set AppConnection = oSmartOccurance
            AppConnection.enumPorts pPortColl
            
            If pPortColl.Count < 2 Then Exit Sub
            
            Dim isBottom As Boolean
            Dim IsBottomFlange As String
            isBottom = False
            Dim oTempFeature As IJStructFeature
            Set oTempFeature = pPropertyDescriptions.CAO
            

            ' ----------------------------------------------------------------------
            ' If the bounding object is a member, find out which edge is intersected
            ' ----------------------------------------------------------------------
            
            Dim oSDO_WebCut As New StructDetailObjects.WebCut
            Set oSDO_WebCut.object = pPropertyDescriptions.CAO

            Dim boundingEdgeId As eBounding_Edge
            Dim mappedEdgeId As JXSEC_CODE
            
            boundingEdgeId = None
            
            If oTempFeature.get_StructFeatureType = SF_FlangeCut Then
                GetSelectorAnswer oSDO_WebCut.object, "BottomFlange", IsBottomFlange
                If IsBottomFlange = gsYes Then
                    isBottom = True
                End If
            End If
            ' Get the intersected edge
            If TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic Then
                GetNonPenetratedIntersectedEdge oSDO_WebCut.object, oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort, boundingEdgeId, mappedEdgeId, isBottom
            End If
            
            '***************
            
            Dim phelper As StructDetailObjects.Helper
            Set phelper = New StructDetailObjects.Helper
            
            Dim oBoundingPort As IJPort
            Dim oBoundedPort As IJPort
            ' Handled only for cases when the bounded is member and bounding can be profile or member.
            If phelper.ObjectType(oSDO_WebCut.Bounding) = SDOBJECT_MEMBER Then
                Dim oMemberPart As New StructDetailObjects.MemberPart
                Set oMemberPart.object = oSDO_WebCut.Bounding
                
                Set oBoundingPort = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounding, mappedEdgeId)
                Set oBoundingPort = phelper.GetEquivalentLastPort(oBoundingPort)
                
                If oBoundingPort Is Nothing Then
                    Dim oBoundingPortBeforeTrim As IJStructPort
                    Set oBoundingPortBeforeTrim = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounding, mappedEdgeId)
                    Dim oGraphConn As IJStructGraphConnectable
                    Set oGraphConn = oSDO_WebCut.Bounding
                    Dim oACTools As AssemblyConnectionTools
                    Set oACTools = New AssemblyConnectionTools
                    oACTools.GetBindingPort oSDO_WebCut.Bounding, JS_TOPOLOGY_PROXY_LFACE, oBoundingPortBeforeTrim.OperationID, oBoundingPortBeforeTrim.OperatorID, oBoundingPortBeforeTrim.ContextID, oBoundingPortBeforeTrim.SectionID, "", oBoundingPort
                End If
            ElseIf phelper.ObjectType(oSDO_WebCut.Bounding) = SDOBJECT_STIFFENER Then
                Dim oProfilePart As New StructDetailObjects.ProfilePart
                sMsg = "Setting Profile Part wrapper object to penetrating object"
                Set oProfilePart.object = oSDO_WebCut.Bounding
                sMsg = "Getting Bounding object's port"
                'If Bounding is Stiffener then currently set BoundingPort as Lateral(Global Port)
                Set oBoundingPort = oProfilePart.baseport(BPT_Lateral)
            Else
                Exit Sub
            End If
            
            If phelper.ObjectType(oSDO_WebCut.Bounded) = SDOBJECT_MEMBER Then
                Dim oStructProfilePart As IJStructProfilePart
                Dim oStructEndCutUtil As IJStructEndCutUtil
                Set oStructProfilePart = oSDO_WebCut.Bounded
                Set oStructEndCutUtil = oStructProfilePart.StructEndCutUtil
                oStructEndCutUtil.GetLatePortForFeatureSegment oSDO_WebCut.object, 8193, oBoundedPort
            Else
                Exit Sub
            End If
            
            Dim oBoundingPort_old As IJPort
            Dim oBoundedPort_old As IJPort
            Set oSD_PhysicalConn.object = pObject
            Set oBoundingPort_old = oSD_PhysicalConn.Port2
            Set oBoundedPort_old = oSD_PhysicalConn.Port1
            Dim oConnectionDefinition As GSCADSDCreateModifyUtilities.IJSDConnectionDefinition
            Set oConnectionDefinition = New GSCADSDCreateModifyUtilities.SDConnectionUtils
            If Not oBoundingPort_old Is oBoundingPort Then
                
                oConnectionDefinition.ReplacePhysicalConnectionPort pObject, _
                                                              oSD_PhysicalConn.Port2, oBoundingPort
                                                              
            End If
            If Not oBoundedPort_old Is oBoundedPort Then
                oConnectionDefinition.ReplacePhysicalConnectionPort pObject, _
                                                            oSD_PhysicalConn.Port1, oBoundedPort
            End If
        End If
    End If
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
'******************************************************************************************
' Method:
' GetEdgeIDsForPC_TopEdge
'
' Description: gets the bounding and bounded subports based on the dispid
' *******************************************************************************************
Public Sub GetEdgeIDsForPC_TopEdge(pPropertyDescriptions As IJDPropertyDescription, lDispId As Long, oBoundingSubPort As JXSEC_CODE, oBoundedSubPort As JXSEC_CODE)
    
    Const METHOD = "GetEdgeIDsForPC_TopEdge"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Getting the Bounded and bounding Ports based on DispId"
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pPropertyDescriptions.CAO
    
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Dim oBounded As Object
    Dim oBounding As Object
    Set oBounded = oSDO_WebCut.BoundedPort
    Set oBounding = oSDO_WebCut.BoundingPort
    
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pPropertyDescriptions.CAO, oBounding, oBounded, lSectionAlias, bPenetratesWeb)
    
    If lDispId = 0 Then
        GoTo ErrorHandler
    ElseIf lDispId = 1 Then ' Top PC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP))
        oBoundedSubPort = JXSEC_TOP
        
    ElseIf lDispId = 2 Then 'TopEdgePC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT))
        oBoundedSubPort = JXSEC_TOP_FLANGE_RIGHT
        
    ElseIf lDispId = 3 Then 'TopInsidePC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT_BOTTOM))
        oBoundedSubPort = JXSEC_TOP_FLANGE_RIGHT_BOTTOM
        
    ElseIf lDispId = 4 Then  'FacePC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT))
        oBoundedSubPort = JXSEC_WEB_RIGHT
    Else
        Exit Sub
    End If
      
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
'******************************************************************************************
' Method:
' GetEdgeIDsForPC_TopEdge
'
' Description: gets the bounding and bounded subports based on the dispid
' *******************************************************************************************
Public Sub GetEdgeIDsForPC_TopAndBtmEdge(pPropertyDescriptions As IJDPropertyDescription, lDispId As Long, oBoundingSubPort As JXSEC_CODE, oBoundedSubPort As JXSEC_CODE)
    
    Const METHOD = "GetEdgeIDsForPC_TopAndBtmEdge"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Getting the Bounded and bounding Ports based on DispId"
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pPropertyDescriptions.CAO
    
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Dim oBounded As Object
    Dim oBounding As Object
    Set oBounded = oSDO_WebCut.BoundedPort
    Set oBounding = oSDO_WebCut.BoundingPort
    
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pPropertyDescriptions.CAO, oBounding, oBounded, lSectionAlias, bPenetratesWeb)
    
    If lDispId = 0 Then
        GoTo ErrorHandler
    ElseIf lDispId = 1 Then ' Top PC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP))
        oBoundedSubPort = JXSEC_TOP
        
    ElseIf lDispId = 2 Then 'TopEdgePC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT))
        oBoundedSubPort = JXSEC_TOP_FLANGE_RIGHT
        
    ElseIf lDispId = 3 Then 'TopInsidePC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT_BOTTOM))
        oBoundedSubPort = JXSEC_TOP_FLANGE_RIGHT_BOTTOM
        
    ElseIf lDispId = 4 Then  'FacePC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT))
        oBoundedSubPort = JXSEC_WEB_RIGHT
        
    ElseIf lDispId = 5 Then 'BottomInsidePC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT_TOP))
        oBoundedSubPort = JXSEC_BOTTOM_FLANGE_RIGHT_TOP
        
    ElseIf lDispId = 6 Then 'BottomEdgePC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT))
        oBoundedSubPort = JXSEC_BOTTOM_FLANGE_RIGHT
        
    ElseIf lDispId = 7 Then 'BottomPC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM))
        oBoundedSubPort = JXSEC_BOTTOM
        
    Else
        Exit Sub
    End If
      
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
'******************************************************************************************
' Method:
' GetEdgeIDsForPC_BtmEdge
'
' Description: gets the bounding and bounded subports based on the dispid
' *******************************************************************************************
Public Sub GetEdgeIDsForPC_BtmEdge(pPropertyDescriptions As IJDPropertyDescription, lDispId As Long, oBoundingSubPort As JXSEC_CODE, oBoundedSubPort As JXSEC_CODE)
    
    Const METHOD = "GetEdgeIDsForPC_BtmEdge"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Getting the Bounded and bounding Ports based on DispId"
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pPropertyDescriptions.CAO
    
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Dim oBounded As Object
    Dim oBounding As Object
    Set oBounded = oSDO_WebCut.BoundedPort
    Set oBounding = oSDO_WebCut.BoundingPort
    
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pPropertyDescriptions.CAO, oBounding, oBounded, lSectionAlias, bPenetratesWeb)
    
    If lDispId = 0 Then
        GoTo ErrorHandler
    ElseIf lDispId = 1 Then 'FacePC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT))
        oBoundedSubPort = JXSEC_WEB_RIGHT
    ElseIf lDispId = 2 Then 'BottomInsidePC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT_TOP))
        oBoundedSubPort = JXSEC_BOTTOM_FLANGE_RIGHT_TOP
        
    ElseIf lDispId = 3 Then 'BottomEdgePC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT))
        oBoundedSubPort = JXSEC_BOTTOM_FLANGE_RIGHT
        
    ElseIf lDispId = 4 Then 'BottomPC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM))
        oBoundedSubPort = JXSEC_BOTTOM
    Else
        Exit Sub
    End If
      
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
'******************************************************************************************
' Method:
' GetEdgeIDsForPC_Face
'
' Description: gets the bounding and bounded subports based on the dispid
' *******************************************************************************************
Public Sub GetEdgeIDsForPC_Face(pPropertyDescriptions As IJDPropertyDescription, lDispId As Long, oBoundingSubPort As JXSEC_CODE, oBoundedSubPort As JXSEC_CODE)
    
    Const METHOD = "GetEdgeIDsForPC_Face"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Getting the Bounded and bounding Ports based on DispId"
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pPropertyDescriptions.CAO
    
    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    Dim oBounded As Object
    Dim oBounding As Object
    Set oBounded = oSDO_WebCut.BoundedPort
    Set oBounding = oSDO_WebCut.BoundingPort
    
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pPropertyDescriptions.CAO, oBounding, oBounded, lSectionAlias, bPenetratesWeb)
    
    If lDispId = 0 Then
        GoTo ErrorHandler
    ElseIf lDispId = 1 Then ' Top PC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP))
        oBoundedSubPort = JXSEC_TOP
    ElseIf lDispId = 2 Then  'FacePC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT))
        oBoundedSubPort = JXSEC_WEB_RIGHT
    ElseIf lDispId = 3 Then 'BottomPC
        oBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM))
        oBoundedSubPort = JXSEC_BOTTOM
    Else
        Exit Sub
    End If
      
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'******************************************************************************************
' Method:
' CMNeedtoComputeCF_MB
'
' Description: Replaces corner feature old ports with new ones
' *******************************************************************************************

Public Sub CMNeedtoComputeCF_MB(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
                                            
    Const METHOD = "CMNeedtoComputeCF_MB"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Modifying the CF "
            
    If pObject Is Nothing Then Exit Sub
        
    If TypeOf pObject Is IJStructFeature Then
        
        Dim oFeature As IJStructFeature
        Set oFeature = pObject
        
        If oFeature.get_StructFeatureType = SF_CornerFeature Then
                
            Dim oFeatureUtils As IJSDFeatureAttributes
            Dim oFacePort As IJPort
            Dim oEdgePort1 As IJPort
            Dim oEdgePort2 As IJPort
            Dim oStructEdgePort1 As IJStructPort
            Dim oStructEdgePort2 As IJStructPort
            
            'get the CF inputs
            Set oFeatureUtils = New SDFeatureUtils
            oFeatureUtils.get_CornerCutInputsEx oFeature, oFacePort, oEdgePort1, oEdgePort2
                                
            Set oStructEdgePort1 = oEdgePort1
            Set oStructEdgePort2 = oEdgePort2
            
            Dim eXidOldEdgePort1 As JXSEC_CODE
            Dim eXidOldEdgePort2 As JXSEC_CODE
            Dim eXidNewEdgePort1 As JXSEC_CODE
            Dim eXidNewEdgePort2 As JXSEC_CODE
            
            eXidOldEdgePort1 = oStructEdgePort1.SectionID
            eXidOldEdgePort2 = oStructEdgePort2.SectionID
                                
            Dim oBoundedPort As IJPort
            Dim oBoundingPort As IJPort
            Dim eEndCutType As eEndCutTypes
            
            'Get bounding and bounded ports and also objects
            GetBoundingAndBounded pPropertyDescriptions.CAO, oBoundingPort, oBoundedPort, , , eEndCutType
            
            Dim oMappedPortsColl As Collection
            Dim oAnglesColl As Collection
            
            GetMultiBoundingEdgeMap pPropertyDescriptions.CAO, eEndCutType, oMappedPortsColl, oAnglesColl
                                       
            'For Three Surfaces Bounding there is possibility of 2 Corners
            'Between 1st and 2nd and other one is between 2nd and 3rd.
            'Out of these two whichever corner forms angle < 180 deg
            'we take those two ports as inputs for CF
            
            If LessThan(oAnglesColl.Item(1), GetPI) Then
                eXidNewEdgePort1 = e_JXSEC_MultipleBounding_5001
                eXidNewEdgePort2 = e_JXSEC_MultipleBounding_5002
            Else
                eXidNewEdgePort1 = e_JXSEC_MultipleBounding_5003
                eXidNewEdgePort2 = e_JXSEC_MultipleBounding_5002
            End If
            
            'Check if old and new ports IDs are same
            If Not ((eXidOldEdgePort1 = eXidNewEdgePort1 And eXidOldEdgePort2 = eXidNewEdgePort2) Or _
                    (eXidOldEdgePort1 = eXidNewEdgePort2 And eXidOldEdgePort2 = eXidNewEdgePort1)) Then
                
                Dim oProfFacePort As IJPort
                Dim oProfEdgePort1 As IJPort
                Dim oProfEdgePort2 As IJPort
                
                sMsg = "Getting the valid ports for Corner Feature based on Edge ID's passed on End Cut object"
                
                'Get updated edge ports by passing new Edge IDs
                MemberCornerFeatureDataByObject ByVal pPropertyDescriptions.CAO, eXidNewEdgePort1, _
                                                            eXidNewEdgePort2, oProfFacePort, oProfEdgePort1, oProfEdgePort2
                                
                Dim oSmartOccurance As IJSmartOccurrence
                Set oSmartOccurance = pObject
                
                'Update Corner Feature with new Ports
                Dim oSDFeatureUtils As New SDFeatureUtils
                oSDFeatureUtils.CreateCornerCutEx GetResourceMgr(), oFacePort, oProfEdgePort1, oProfEdgePort2, _
                                                            oSmartOccurance.Item, pPropertyDescriptions.CAO, pObject
            
            End If
            
        End If
        
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub



