VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenWebSnipedTubeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : GenWebSnipedTubeDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    05/oct/2012 - Created
'   05/November - modified the CM_IsphysConnNeeded and PC construct method
'
'*********************************************************************************************
Const m_sClassName As String = "GenWebSnipedTubeDef"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ItemInputs
' Description : List any graphic Inputs that the Definition has here
'
'*********************************************************************************************
Public Sub ItemInputs(pIH As IJDInputsHelper)
    Const METHOD = m_DefinitionProgid & "::ItemInputs"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining Definition Inputs"

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : ItemAggregator
' Description :
'
'*********************************************************************************************
Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    Const METHOD = m_DefinitionProgid & "::ItemAggregator"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining ItemAggregator"
    
    pAD.UserTypeClsid = CA_WEBCUT
    pAD.AggregatorClsid = CA_AGGREGATE     ' CStructFeature
    
    pAD.SetCMFinalConstruct CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_FinalConstructEndCut"
    pAD.SetCMMigrate CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_MigrateAggregator"
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : ItemMembers
' Description : List all the Item members that this SmartClass can create
'
'*********************************************************************************************
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    Const METHOD = m_DefinitionProgid & "::ItemMembers"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim sDispId As String
    
    Dim iDispId As Long
    Dim oMemDesc As IJDMemberDescription
    
    'Physical Connection for the Web Cut
    iDispId = 1
    sDispId = "GenericWeb_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPCForGussetPlateAndTubeMbr", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnsNeeded"
    oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigratePhysConn"
    Set oMemDesc = Nothing
    'with platebase
    iDispId = 2
    sDispId = "GenericNonPenTopAndBtmPC_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPCForGussetPlateAndTubeMbr", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnsNeeded"
    Set oMemDesc = Nothing
   
    'With plateoffset
    iDispId = 3
    sDispId = "GenericNonPenTopAndBtmPC_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPCForGussetPlateAndTubeMbr", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnsNeeded"
    Set oMemDesc = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Smart Definition as a regular symbol definition
' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    
    IJDUserSymbolServices_GetDefinitionName = m_DefinitionName
    
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)
    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pDefinition.IJDInputs.RemoveAllInput
    pDefinition.IJDRepresentations.RemoveAllRepresentation

    pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
      
    ' define the inputs
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pDefinition
    pIH.InitAs m_FamilyProgid
    ItemInputs pIH
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Dim pAPDs As IJDPropertyDescriptions
    Set pAD = pDefinition
    Set pAPDs = pDefinition
    pAPDs.RemoveAll ' Removes all the previous property descriptions
    ItemAggregator pAD
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    pMDs.RemoveAll ' Removes all the previous Member descriptions
    ItemMembers pMDs
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, _
                                                            ByVal defParams As Variant, _
                                                            ByVal ActiveConnection As Object) As Object
    Dim pDefinition As IJDSymbolDefinition
    Dim pCAFactory As New CAFactory
    
    Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
    
    ' Set definition progId and codebase
    pDefinition.ProgId = m_DefinitionProgid
    pDefinition.CodeBase = CodeBase
    pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
      
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
    
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                                                      ByVal repName As String, _
                                                      ByVal outputcoll As Object, _
                                                      ByRef arrayOfInputs())
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_EditOccurence(pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean

End Function
'*********************************************************************************************
' Method      : CM_ConstructPCForGussetPlateAndTubeMbr
'
' Description : Creates Physical connection between Flange of Tube and Gusset plate face
'
'*********************************************************************************************
Public Function CM_ConstructPCForGussetPlateAndTubeMbr(ByVal pMemberDescription As IJDMemberDescription, _
                                      ByVal pResourceManager As IUnknown, _
                                      ByRef pObject As Object)
    Const METHOD = m_DefinitionProgid & "::CM_ConstructPCForGussetPlateAndTubeMbr"
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
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut

    Set oSDO_WebCut.object = oEndCutObject

    sMsg = "Getting Bounded objects from WebCut"

    Set oBoundedPart = oSDO_WebCut.Bounded
    Set oBoundingPart = oSDO_WebCut.Bounding

    Set oBoundedPort = oSDO_WebCut.BoundedPort
    Set oBoundingPort = oSDO_WebCut.BoundingPort

    Dim oCutoutBndedPort As IJPort

    Set oCutoutBndedPort = CutoutSubPort(oBoundedPart, oEndCutObject, -1)

    Dim oHelper As StructDetailObjects.Helper
    Dim oLastBoundingPort As IJPort

    Dim oBasePort As IJPort
    Dim oOffsetPort As IJPort
    Dim oSDO_PlatePart As StructDetailObjects.PlatePart

    Set oSDO_PlatePart = New StructDetailObjects.PlatePart
    Set oSDO_PlatePart.object = oBoundingPart
    Set oBasePort = oSDO_PlatePart.BasePort(BPT_Base)
    Set oOffsetPort = oSDO_PlatePart.BasePort(BPT_Offset)
    Set oHelper = New StructDetailObjects.Helper

    If pMemberDescription.dispid = 2 Then
        Set oLastBoundingPort = oHelper.GetEquivalentLastPort(oBasePort)
    ElseIf pMemberDescription.dispid = 3 Then
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
    
    Dim sError As String
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
    Dim WLSurfaceNormal As IJDVector
    Dim WRSurfaceNormal As IJDVector
    Dim oNormalToBase As IJDVector
    Dim oNormalToOffset As IJDVector
    Dim dWLDotProduct As Double
    Dim dWRDotProduct As Double
    Dim DotProduct1 As Double
    Dim DotProduct2 As Double
    Dim IsBndTube As Boolean
    
    bIsNeeded = False
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMemberDescription.CAO
    
    Dim eFeatureType As StructFeatureTypes
    eFeatureType = 0
    Dim oFeature As IJStructFeature
    Set oFeature = oSDO_WebCut.object
    eFeatureType = oFeature.get_structfeaturetype
    Dim vCuttingBehaviorType As Variant
    
    If eFeatureType = SF_FlangeCut Then
        vCuttingBehaviorType = ToFlangeInnerSurface
        Set_AttributeValue oSDO_WebCut.object, "IJShpStrFlangeCut", "CuttingBehavior", vCuttingBehaviorType
    End If
    
    Set oClosestPosition1 = New DPosition
    Set oClosestPosition2 = New DPosition
    Set oClosestPosition3 = New DPosition
    Set oSGOModelBodyUtilities = New SGOModelBodyUtilities
    Dim sAnswer As String
    Dim HasIntersectGeom As Boolean
    Dim oExtSurface As IJSurfaceBody
    Dim oGeomOffset As IJGeometryOffset
    Set oGeomOffset = New DGeomOpsOffset

    GetSelectorAnswer pMemberDescription.CAO, "CreatePhysicalConnection", sAnswer
    IsBndTube = IsTubularMember(oSDO_WebCut.BoundedPort)
    If sAnswer = "No" Then
        bIsNeeded = False
        Exit Sub
    Else
        If IsBndTube Then
            bIsNeeded = True
            Exit Sub
        End If
    End If
        If pMemberDescription.dispid = 1 And Not IsBndTube Then
            bIsNeeded = True
      
        ElseIf pMemberDescription.dispid = 2 And Not IsBndTube Then
         
            If TypeOf oSDO_WebCut.Bounding Is IJPlatePart Then
           
                Set oBoundingPart = oSDO_WebCut.Bounding
                Set oSDO_PlatePart = New StructDetailObjects.PlatePart
                Set oSDO_PlatePart.object = oBoundingPart
                
                Set oBaseSurface = oSDO_PlatePart.BasePort(BPT_Base).Geometry
                ' InOrder to avoid creation of Pc for the cases where there is no material to create PC.It is checking by HasIntersectingGeometry
                oGeomOffset.CreateExtendedSheetBody Nothing, oBaseSurface, Nothing, 0.5, Nothing, oExtSurface
                HasIntersectGeom = oSGOModelBodyUtilities.HasIntersectingGeometry(oExtSurface, oSDO_WebCut.Bounded)
                If Not HasIntersectGeom Then
                    bIsNeeded = False
                    Exit Sub
                End If
                
                Set oWebLeftSurface = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_WEB_LEFT).Geometry
                Set oWebRightSurface = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_WEB_RIGHT).Geometry
                
                oSGOModelBodyUtilities.GetClosestPointsBetweenTwoBodies oBaseSurface, oWebLeftSurface, oClosestPosition1, oClosestPosition2, dMinDistToWebLeft
                oSGOModelBodyUtilities.GetClosestPointsBetweenTwoBodies oBaseSurface, oWebRightSurface, oClosestPosition1, oClosestPosition3, dMinDistToWebRight
                
                oWebLeftSurface.GetNormalFromPosition oClosestPosition2, WLSurfaceNormal
                oWebRightSurface.GetNormalFromPosition oClosestPosition3, WRSurfaceNormal
                
                oBaseSurface.GetNormalFromPosition oClosestPosition1, oNormalToBase
                dLength = oNormalToBase.Length
                oNormalToBase.Length = -dLength
                
                 Set WLNormal = New dVector
                 Set WRNormal = New dVector
                 
                 WLNormal.Set oClosestPosition1.x - oClosestPosition2.x, oClosestPosition1.y - oClosestPosition2.y, oClosestPosition1.z - oClosestPosition2.z
                 WRNormal.Set oClosestPosition1.x - oClosestPosition3.x, oClosestPosition1.y - oClosestPosition3.y, oClosestPosition1.z - oClosestPosition3.z
                 
                 dWLDotProduct = WLNormal.Dot(oNormalToBase)
                 dWRDotProduct = WRNormal.Dot(oNormalToBase)
                 
                 DotProduct1 = WLSurfaceNormal.Dot(oNormalToBase)
                 DotProduct2 = WRSurfaceNormal.Dot(oNormalToBase)

                 If LessThan(dMinDistToWebLeft, dMinDistToWebRight) Then
                    If dWLDotProduct < 0# And DotProduct1 < 0# Then
                            bIsNeeded = True
                    ElseIf dWLDotProduct > 0# And dMinDistToWebLeft <= 0.002 Then
                            bIsNeeded = True
                    End If
                 ElseIf LessThan(dMinDistToWebRight, dMinDistToWebLeft) Then
                    If dWRDotProduct < 0# And DotProduct2 < 0# Then
                        bIsNeeded = True
                    ElseIf dWRDotProduct > 0# And dMinDistToWebRight <= 0.002 Then
                            bIsNeeded = True
                    End If
                 Else
                    bIsNeeded = False
                 
                 End If
            End If
                 
        ElseIf pMemberDescription.dispid = 3 And Not IsBndTube Then
        
            If TypeOf oSDO_WebCut.Bounding Is IJPlatePart Then
                
                Set oBoundingPart = oSDO_WebCut.Bounding
                Set oSDO_PlatePart = New StructDetailObjects.PlatePart
                Set oSDO_PlatePart.object = oBoundingPart
                    
                Set oOffsetSurface = oSDO_PlatePart.BasePort(BPT_Offset).Geometry
                ' InOrder to avoid creation of Pc for the cases where there is no material to create PC.It is checking by HasIntersectingGeometry
                oGeomOffset.CreateExtendedSheetBody Nothing, oOffsetSurface, Nothing, 0.5, Nothing, oExtSurface
                HasIntersectGeom = oSGOModelBodyUtilities.HasIntersectingGeometry(oExtSurface, oSDO_WebCut.Bounded)
                If Not HasIntersectGeom Then
                    bIsNeeded = False
                    Exit Sub
                End If
                Set oWebLeftSurface = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_WEB_LEFT).Geometry
                Set oWebRightSurface = GetLateralSubPortBeforeTrim(oSDO_WebCut.Bounded, JXSEC_WEB_RIGHT).Geometry
                
                oSGOModelBodyUtilities.GetClosestPointsBetweenTwoBodies oOffsetSurface, oWebLeftSurface, oClosestPosition1, oClosestPosition2, dMinDistToWebLeft
                oSGOModelBodyUtilities.GetClosestPointsBetweenTwoBodies oOffsetSurface, oWebRightSurface, oClosestPosition1, oClosestPosition3, dMinDistToWebRight
                
                oWebLeftSurface.GetNormalFromPosition oClosestPosition2, WLSurfaceNormal
                oWebRightSurface.GetNormalFromPosition oClosestPosition3, WRSurfaceNormal
                
                oOffsetSurface.GetNormalFromPosition oClosestPosition1, oNormalToOffset
                dLength = oNormalToOffset.Length
                oNormalToOffset.Length = -dLength
                 
                 Set WLNormal = New dVector
                 Set WRNormal = New dVector
                 
                 WLNormal.Set oClosestPosition1.x - oClosestPosition2.x, oClosestPosition1.y - oClosestPosition2.y, oClosestPosition1.z - oClosestPosition2.z
                 WRNormal.Set oClosestPosition1.x - oClosestPosition3.x, oClosestPosition1.y - oClosestPosition3.y, oClosestPosition1.z - oClosestPosition3.z
                 dWLDotProduct = WLNormal.Dot(oNormalToOffset)
                 dWRDotProduct = WRNormal.Dot(oNormalToOffset)
                
                DotProduct1 = WLSurfaceNormal.Dot(oNormalToOffset)
                DotProduct2 = WRSurfaceNormal.Dot(oNormalToOffset)

                 If LessThan(dMinDistToWebLeft, dMinDistToWebRight) Then
                    If dWLDotProduct < 0# And DotProduct1 < 0# Then
                            bIsNeeded = True
                    ElseIf dWLDotProduct > 0# And dMinDistToWebLeft <= 0.002 Then
                            bIsNeeded = True
                    End If
                 ElseIf LessThan(dMinDistToWebRight, dMinDistToWebLeft) Then
                    If dWRDotProduct < 0# And DotProduct2 < 0# Then
                        bIsNeeded = True
                    ElseIf dWRDotProduct > 0# And dMinDistToWebRight <= 0.002 Then
                            bIsNeeded = True
                    End If
                 Else
                    bIsNeeded = False
                 End If
            End If
        
        End If
   
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sError).Number

End Sub


