VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberEncaseDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------
'    Copyright (C) 2008 Intergraph Corporation. All rights reserved.
'
'
'
'   Custom assembly definition for the Member Encasement Definition
'   --------------------------------------------
'
'
'History
'    SS         10/29/08      Creation
'    MH         01/07/09      TR 156809 removed creation of BspCurve for curved member
'
'---------------------------------------------------------------------------------------

Option Explicit

Private Const MODULE = "MemberEncaseDef"

Const m_ItemProgId As String = "SPSMemberEncasements.MemberEncaseDef"
Const m_ItemName As String = "SPSMemberEncasements.MemberEncaseDef"

Const DisplayAspect = 32        ' aspect 5 = Insuilation.  2 **5 = 32

Dim bOnPreLoad As Boolean

Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements IStructInsulationGraphicInputHelper
Implements IStructCustomFoulCheck2
Implements IStructInsulationValidationHelper



Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

    IJDUserSymbolServices_EditOccurence = False

End Function


Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  
    IJDUserSymbolServices_GetDefinitionName = m_ItemName

End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IMSSymbolEntities.IJDSymbolDefinition)
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler
    
    Dim pAD As IJDAggregatorDescription
    Dim pPropertyDescriptions As IJDPropertyDescriptions
    Dim pMemberDescriptions As IJDMemberDescriptions
    Dim pMemberDescription As IJDMemberDescription
    
    pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
    Set pAD = pDefinition
    pAD.AggregatorClsid = "{FEE1FE1B-A358-47CB-A93D-9F55ED83ED00}" 'CLSID_StructInsulation
    pAD.SetCMConstruct imsCOOKIE_ID_USS_LIB, "CMConstructAggregator"

    ''' removed this so that SO will create the symbol references relation   pAD.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputAggregator"
    ''' removed  pAD.SetCMRemoveInputs -1, -1
    
    pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateAggregator"
    pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAggregator"
    Set pAD = Nothing
    
   ' Aggregator property
    Set pPropertyDescriptions = pDefinition
    pPropertyDescriptions.RemoveAll ' Remove all the previous property description

    'the setback interface is an input interface
    pPropertyDescriptions.AddProperty "MemberEncaseSetback", 1, SETBACK_INTERFACE

    pPropertyDescriptions.AddProperty "MemberEncaseWCG", 2, IJWeightCG, "CMEvaluateMemberEncaseWCG", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
    pPropertyDescriptions.AddProperty "MemberEncaseVolume", 3, IJGenericVolume, "CMEvaluateMemberEncaseVolume", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
    pPropertyDescriptions.AddProperty "MemberEncaseSurfaceArea", 4, IJSurfaceArea, "CMEvaluateMemberEncaseSurfaceArea", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
    pPropertyDescriptions.AddProperty "MemberEncaseLocalCoordinateSystem", 5, IJLocalCoordinateSystem, "CMEvaluateMemberEncaseLCS", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
    pPropertyDescriptions.AddProperty "MemberEncaseAggregatedGeometry", 6, IJGeometry, "CMEvaluateMemberEncaseAggGeometry", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
    pPropertyDescriptions.AddProperty "MemberEncaseAggregatedGeometry", 7, IJGraphicDataCache, "CMEvaluateMemberEncaseGTypeCache", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
                                        
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub


Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal pResourceMgr As Object) As Object
Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
On Error GoTo ErrorHandler

  Dim pDefinition As IJDSymbolDefinition
  Dim pFact As IJCAFactory
  Set pFact = New CAFactory
  Set pDefinition = pFact.CreateCAD(pResourceMgr)
  
  ' Set definition progId and codebase
  pDefinition.ProgId = m_ItemProgId
  pDefinition.CodeBase = CodeBase
  
  pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(definitionParameters)
    
  ' Initialize the definition
  IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
  
  ' Persistence behavior
  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
  'returned symbol definition
  Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
  
  Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

Public Sub CMFinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMFinalConstructAggregator"
    On Error GoTo ErrorHandler
    Dim mySO As IJSmartOccurrence
    Dim mySI As IJSmartItem
    
    Set mySO = pAggregatorDescription.CAO
    Set mySI = mySO.ItemObject
    
    CopyValuesFromItemToSO GetAttributeCollection(mySO, SETBACK_INTERFACE), _
                           GetAttributeCollection(mySI, SETBACK_INTERFACE)
    
        
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMConstructAggregator"
    On Error GoTo ErrorHandler

    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMMigrateAggregator(pAggregatorDescription As IJDAggregatorDescription, oMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateAggregator"
    On Error GoTo ErrorHandler
  
    Dim oChild As IJDesignChild
    Dim oInsulation As IStructInsulation
    Dim oReplacedMember As ISPSMemberPartPrismatic
    Dim oReplacingMembers As IJDObjectCollection
    Dim oReplacedInsulation As IJDObjectCollection
    Dim oReplacingInsulation As IJDObjectCollection
    Dim oInsulationServices As IInsulationServices
    Dim toBeDeleted As Boolean

    Set oInsulation = pAggregatorDescription.CAO

    Set oChild = oInsulation
    Set oReplacedMember = oChild.GetParent

    oMigrateHelper.ObjectsReplacing oReplacedMember, oReplacingMembers, toBeDeleted
    
    Set oInsulationServices = New InsulationServices
    
    oInsulationServices.ProcessSplit oInsulation, oReplacedMember, oReplacingMembers, oReplacingInsulation

    ' we have replacements?
    If oReplacingInsulation.count > 0 Then
        Set oReplacedInsulation = New JObjectCollection
        oReplacedInsulation.Add oInsulation
        
        oMigrateHelper.ObjectsReplaced oReplacedInsulation, oReplacingInsulation, True
        
        Set oReplacedInsulation = Nothing
    End If

    Set oInsulation = Nothing
    Set oInsulationServices = Nothing
    Set oReplacingInsulation = Nothing
    Set oReplacedMember = Nothing
    Set oReplacingMembers = Nothing
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateMemberEncaseWCG(pPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateMemberEncaseWCG"
    On Error GoTo ErrorHandler

    Dim iInsulation As IStructInsulation
    Dim oRC As IJDReferencesCollection
    Dim iMemberPart As ISPSMemberPartPrismatic
    
    Dim oSetbackAxis As IJCurve
    Dim servicesStatus As StructInsulationServicesStatus
    
    Dim iTransform As IJDT4x4
    Dim iSPSCrossSection As iSPSCrossSection
    Dim iAxisRotation As ISPSAxisRotation
    Dim dCPx As Double, dCPy As Double
    Dim dOVx As Double, dOVy As Double, dOVz As Double
    Dim dVolume As Double, dSurfaceArea As Double, dCGx As Double, dCGy As Double, dCGz As Double, dWeight As Double
    Dim startx As Double, starty As Double, startz As Double, endx As Double, endy As Double, endz As Double
    Dim vecOV As IJDVector
    Dim oSolid As Object
    Dim iMaterial As IJDMaterial

    Dim iiInsulationServices As IInsulationServices
    Set iiInsulationServices = New StructInsulations.InsulationServices

    Set iInsulation = pPropertyDescription.CAO
    Set oRC = GetRefCollection(iInsulation)

    Set iMemberPart = oRC.IJDEditJDArgument.GetEntityByIndex(3)
    Set iSPSCrossSection = iMemberPart.CrossSection
    Set iAxisRotation = iMemberPart.Rotation

    If ValidateClassWithSpec(iInsulation) = 0 Then      ' encasement class and spec encasement class do not agree
        iInsulation.ComputeStatus = StructInsulationInputHelper_IncompatibleEncasement
        SPSToDoErrorNotify INSULATIONERROR_TDLCODELISTNAME, INSULATIONERROR_NOSELECTION, iInsulation, Nothing
        Err.Raise E_FAIL
    End If
    
    Dim graphicInputHelperStatus As StructInsulationGraphicInputHelperStatus
    graphicInputHelperStatus = ValidateMemberFPCriteria(iMemberPart, iInsulation.definition)
    
      If graphicInputHelperStatus <> StructInsulationInputHelper_Ok Then
         iInsulation.ComputeStatus = graphicInputHelperStatus
         SPSToDoErrorNotify INSULATIONERROR_TDLCODELISTNAME, INSULATIONERROR_BADATTRIBUTEVALUES, iInsulation, Nothing
         Err.Raise E_FAIL
      End If
        
    iSPSCrossSection.GetCardinalPointOffset iSPSCrossSection.CardinalPoint, dCPx, dCPy

    GetSetBackAxis iMemberPart, iInsulation, iiInsulationServices, oSetbackAxis, servicesStatus
    If servicesStatus <> StructInsulationServices_Ok Then
        iInsulation.ComputeStatus = StructInsulationInputHelper_BadAttributeValues
        SPSToDoErrorNotify INSULATIONERROR_TDLCODELISTNAME, INSULATIONERROR_BADATTRIBUTEVALUES, iInsulation, Nothing
        Err.Raise E_FAIL
    End If
    
    oSetbackAxis.EndPoints startx, starty, startz, endx, endy, endz
    
    iAxisRotation.GetTransformAtPosition startx, starty, startz, iTransform, Nothing

    dOVx = iTransform.IndexValue(8)
    dOVy = iTransform.IndexValue(9)
    dOVz = iTransform.IndexValue(10)
    Set vecOV = New DVector
    vecOV.Set dOVx, dOVy, dOVz
    
    Dim iInsulationPM As IStructInsulationPersistMethods
    Set iInsulationPM = iInsulation

    iInsulationPM.StoreAxisCPOrientation oSetbackAxis, dCPx, dCPy, iAxisRotation.Mirror, vecOV

    servicesStatus = iiInsulationServices.ComputeSolidProperties(iInsulation, iMemberPart, dVolume, dSurfaceArea, dCGx, dCGy, dCGz, oSolid)
    If servicesStatus <> StructInsulationServices_Ok Then
        iInsulation.ComputeStatus = StructInsulationInputHelper_UnexpectedError
        SPSToDoErrorNotify INSULATIONERROR_TDLCODELISTNAME, INSULATIONERROR_UNEXPECTED, iInsulation, Nothing
        Err.Raise E_FAIL
    End If

    Set iMaterial = iInsulation.Material
    dWeight = dVolume * iMaterial.Density

    iInsulationPM.StoreSolidProperties dVolume, dSurfaceArea, dWeight, dCGx, dCGy, dCGz

    iInsulationPM.StoreDisplayGraphic DisplayAspect, oSolid

    Exit Sub
    
ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    Else
        HandleError MODULE, METHOD
    End If
End Sub

Public Sub CMEvaluateMemberEncaseVolume(pPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateMemberEncaseVolume"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateMemberEncaseSurfaceArea(pPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateMemberEncaseSurfaceArea"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateMemberEncaseLCS(pPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateMemberEncaseLocalCoordinateSystem"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateMemberEncaseAggGeometry(pPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateMemberEncaseAggregatedGeometry"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateMemberEncaseGTypeCache(pPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateMemberEncaseGTypeCache"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub IStructCustomFoulCheck2_GetConnectedParts(ByVal pInsPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements)
Const METHOD = "IJStructCustomFoulCheck_GetConnectedParts"
On Error GoTo ErrorHandler
                
    ' get inputs related to insulation
    Dim pInputMemColl As IJElements
    IStructInsulationGraphicInputHelper_GetGraphicInputs pInsPartObject, pInputMemColl
    
    Dim pInputMember As ISPSMemberPartPrismatic
    If Not pInputMemColl Is Nothing Then
        Set pInputMember = pInputMemColl.Item(1)
    End If
    pInputMemColl.Clear
        
    ' add input member to the list of monikers
    If Not pInputMember Is Nothing Then
        pIJMonUnks.Add pInputMember
    Else
        Exit Sub
    End If
    
    ' get its connected parts
    Dim pInputMemFoulCheck As IJFoulCheck
    Set pInputMemFoulCheck = pInputMember
    
    Dim pConnMemObjsColl As IJDObjectCollection
    Set pConnMemObjsColl = New JObjectCollection
    
    pInputMemFoulCheck.GetConnectedParts pConnMemObjsColl
    
    If pConnMemObjsColl.count > 0 Then
    
        ' resolve the monikers
        Dim iInsServices As IInsulationServices
        Set iInsServices = New StructInsulations.InsulationServices
        
        Dim pConnMemEles As IJElements
        iInsServices.ResolveMonikerList pInsPartObject, pConnMemObjsColl, pConnMemEles

        Set iInsServices = Nothing
        
        ' loop each of connected members add it to the list as well as its insulations
        Dim index As Long
        Dim pConnMember As ISPSMemberPartPrismatic
        Dim pConnMemInsulations As IJElements
        Dim oStructHasIns As IStructHasInsulation

        
        For index = 1 To pConnMemEles.count
        
            If TypeOf pConnMemEles.Item(index) Is ISPSMemberPartPrismatic Then
            
                Set pConnMember = pConnMemEles.Item(index)
                pIJMonUnks.Add pConnMember
                
                ' get its insulations
                Set oStructHasIns = pConnMember
                Set pConnMemInsulations = oStructHasIns.Insulations
                
                If Not pConnMemInsulations Is Nothing Then
                    pIJMonUnks.AddElements pConnMemInsulations


                End If

            End If
            
            Set pConnMember = Nothing
            pConnMemInsulations.Clear
            
        Next index
    
    End If
    
    pConnMemObjsColl.Clear
    
    'get the input member system
    Dim pInputMemSys As ISPSMemberSystem
    Set pInputMemSys = pInputMember.MemberSystem
    
    If pInputMemSys Is Nothing Then
        Exit Sub
    End If
    
    ' get parts that are at it input member joints
    Dim count As Integer
    Dim pMemSysAtJoints As IJElements
    Dim pPtOnJointsForMemSys As IJElements
    Dim pAxisJoint As ISPSAxisJoint
    Dim pMemSys As ISPSMemberSystem
    Dim ePortIndex As SPSMemberAxisPortIndex
    Dim pMemPart As ISPSMemberPartPrismatic
    Dim pMemEndPart As ISPSMemberPartPrismatic
    
    'get all the point on joints
    Set pPtOnJointsForMemSys = pInputMemSys.PointOnJoints
    
    ' for each joint
    For count = 1 To pPtOnJointsForMemSys.count
        
        Set pAxisJoint = pPtOnJointsForMemSys.Item(count)
        
        Set pMemSysAtJoints = pAxisJoint.EndMemberSystems
        For Each pMemSys In pMemSysAtJoints
            ePortIndex = pMemSys.ResolveEnd(pAxisJoint)
            Set pMemPart = pMemSys.DesignPartAtEnd(ePortIndex)
            If Not pMemPart Is Nothing Then
                pIJMonUnks.Add pMemPart
            End If
        Next
        pMemSysAtJoints.Clear
        Set pMemSysAtJoints = Nothing
    Next count
    
    ' repeat the same for both the ends
    Set pAxisJoint = pInputMemSys.JointAtEnd(SPSMemberAxisStart)
    
    Set pMemSysAtJoints = pAxisJoint.EndMemberSystems
    
    For Each pMemSys In pMemSysAtJoints
        If Not pMemSys Is pInputMemSys Then
            ePortIndex = pMemSys.ResolveEnd(pAxisJoint)
            Set pMemPart = pMemSys.DesignPartAtEnd(ePortIndex)
            If Not pMemPart Is Nothing Then
                pIJMonUnks.Add pMemPart
            End If
        End If
    Next
        
    pMemSysAtJoints.Clear
    Set pMemSysAtJoints = Nothing
    
    Set pAxisJoint = pInputMemSys.JointAtEnd(SPSMemberAxisEnd)
    
    Set pMemSysAtJoints = pAxisJoint.EndMemberSystems
    
    For Each pMemSys In pMemSysAtJoints
        If Not pMemSys Is pInputMemSys Then
            ePortIndex = pMemSys.ResolveEnd(pAxisJoint)
            Set pMemPart = pMemSys.DesignPartAtEnd(ePortIndex)
            If Not pMemPart Is Nothing Then
                pIJMonUnks.Add pMemPart
            End If
        End If
    Next
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub IStructCustomFoulCheck2_GetFoulInterfaceType(ByVal pItem As Object, pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
Const METHOD = "IJStructCustomFoulCheck_GetFoulInterfaceType"
On Error GoTo ErrorHandler

    Dim oInsulation As IStructInsulation
    Set oInsulation = pItem
    
    Dim oMaterial As IJDMaterial
    Set oMaterial = oInsulation.Material
    
    If oMaterial.MaterialType = "Fibrous" Then
        pFoulInterfaceType = NonParticipant
    Else
        pFoulInterfaceType = StandardGraphicEntity
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SP3DStructInterfaces.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
    Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
    On Error GoTo ErrorHandler

    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
    Const METHOD = "IJUserAttributeMgmt_OnPreLoad"
    On Error GoTo ErrorHandler

    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Sub GetSetBackAxis(oMemberPart As ISPSMemberPartPrismatic, oStructInsulation As Object, iiInsulationServices As IInsulationServices, _
                        ByRef oSetbackAxis As IJCurve, ByRef servicesStatus As StructInsulationServicesStatus)
    Const METHOD = "GetSetBackAxis"
    On Error GoTo ErrorHandler

    Dim tol As Double
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oAttrColl As Object
    Dim vAttrValue As Variant
    Dim setbackRef1 As Long, setbackRef2 As Long
    Dim setbackDir1 As Boolean, setbackDir2 As Boolean
    Dim setbackDist1 As Double, setbackDist2 As Double

    Dim oNewCurve As Object
    Dim oMemberAxis As IJCurve
    Dim bSetbackDistancesTooBig As Boolean
    
    Dim paramMid As Double
    Dim paramRef1 As Double, paramRef2 As Double
    Dim paramSetback1 As Double, paramSetback2 As Double

    servicesStatus = StructInsulationServices_UnexpectedError

    tol = oGeomFactory.GeometryServices.DistTolerance
    
    ' read the setback distance and setback reference attributes
    Set oAttrColl = GetAttributeCollection(oStructInsulation, SETBACK_INTERFACE)
    If oAttrColl Is Nothing Then
        setbackDist1 = 0
        setbackDist2 = 0
        setbackRef1 = SETBACKREFERENCE_AXISSTART
        setbackRef2 = SETBACKREFERENCE_AXISEND
    Else
        If GetAttributeValue(oAttrColl, SETBACK_DISTANCE1, vAttrValue) Then
            setbackDist1 = vAttrValue
        Else
            servicesStatus = StructInsulationServices_InvalidInput
            Exit Sub
        End If
        If GetAttributeValue(oAttrColl, SETBACK_DISTANCE2, vAttrValue) Then
            setbackDist2 = vAttrValue
        Else
            servicesStatus = StructInsulationServices_InvalidInput
            Exit Sub
        End If
        If GetAttributeValue(oAttrColl, SETBACK_REFERENCE1, vAttrValue) Then
            setbackRef1 = vAttrValue
        Else
            servicesStatus = StructInsulationServices_InvalidInput
            Exit Sub
        End If
        If GetAttributeValue(oAttrColl, SETBACK_REFERENCE2, vAttrValue) Then
            setbackRef2 = vAttrValue
        Else
            servicesStatus = StructInsulationServices_InvalidInput
            Exit Sub
        End If
    End If

    Set oMemberAxis = oMemberPart.Axis

    If setbackRef1 = SETBACKREFERENCE_PARTSTART Or setbackRef1 = SETBACKREFERENCE_PARTEND Or _
       setbackRef2 = SETBACKREFERENCE_PARTSTART Or setbackRef2 = SETBACKREFERENCE_PARTEND Then
        
        Dim sminx As Double, sminy As Double, sminz As Double, smaxx As Double, smaxy As Double, smaxz As Double
        Dim eminx As Double, eminy As Double, eminz As Double, emaxx As Double, emaxy As Double, emaxz As Double
        
        oMemberPart.GetEndFacePositions sminx, sminy, sminz, smaxx, smaxy, smaxz, eminx, eminy, eminz, emaxx, emaxy, emaxz
        
        If setbackRef1 = SETBACKREFERENCE_PARTSTART Then
            oMemberAxis.Parameter sminx, sminy, sminz, paramRef1
            setbackDir1 = True
        ElseIf setbackRef1 = SETBACKREFERENCE_PARTEND Then
            oMemberAxis.Parameter emaxx, emaxy, emaxz, paramRef1
            setbackDir1 = False
        End If
        If setbackRef2 = SETBACKREFERENCE_PARTSTART Then
            oMemberAxis.Parameter sminx, sminy, sminz, paramRef2
            setbackDir2 = True
        ElseIf setbackRef2 = SETBACKREFERENCE_PARTEND Then
            oMemberAxis.Parameter emaxx, emaxy, emaxz, paramRef2
            setbackDir2 = False
        End If
    End If

    If setbackRef1 = SETBACKREFERENCE_AXISSTART Or setbackRef1 = SETBACKREFERENCE_AXISEND Or _
       setbackRef2 = SETBACKREFERENCE_AXISSTART Or setbackRef2 = SETBACKREFERENCE_AXISEND Then
        
        Dim paramStart As Double, paramEnd As Double

        oMemberAxis.ParamRange paramStart, paramEnd
    
        If setbackRef1 = SETBACKREFERENCE_AXISSTART Then
            paramRef1 = paramStart
            setbackDir1 = True
        ElseIf setbackRef1 = SETBACKREFERENCE_AXISEND Then
            paramRef1 = paramEnd
            setbackDir1 = False
        End If
        If setbackRef2 = SETBACKREFERENCE_AXISSTART Then
            paramRef2 = paramStart
            setbackDir2 = True
        ElseIf setbackRef2 = SETBACKREFERENCE_AXISEND Then
            paramRef2 = paramEnd
            setbackDir2 = False
        End If
    End If

    ' measure setback distance along axis, relative to params at reference params
    servicesStatus = iiInsulationServices.CurveParameterAtDistanceAlong(oMemberAxis, setbackDir1, paramRef1, setbackDist1, paramSetback1)
    If servicesStatus <> StructInsulationServices_Ok Then
        Exit Sub
    End If

    servicesStatus = iiInsulationServices.CurveParameterAtDistanceAlong(oMemberAxis, setbackDir2, paramRef2, setbackDist2, paramSetback2)
    If servicesStatus <> StructInsulationServices_Ok Then
        Exit Sub
    End If

    If Abs(paramSetback1 - paramSetback2) < tol Then    'degenerate case
        servicesStatus = StructInsulationServices_DistanceExceedsCurveLength
        Exit Sub
    End If
        
    paramMid = 0.5 * (paramSetback1 + paramSetback2)

    Set oNewCurve = oGeomFactory.GeometryServices.CreateByPartOfCurve(Nothing, oMemberAxis, paramSetback1, paramMid, paramSetback2)
''    Dim iCurve As IJCurve
''    Dim tanx As Double, tany As Double, tanz As Double
''    Dim tan2x As Double, tan2y As Double, tan2z As Double
''    Set iCurve = oNewCurve
''    iCurve.Evaluate 0#, sminx, sminy, sminz, tanx, tany, tanz, tan2x, tan2y, tan2z
''    MsgBox "tanx=" & tanx & ", tany=" & tany

    Set oSetbackAxis = oNewCurve
    servicesStatus = StructInsulationServices_Ok
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Function IStructInsulationGraphicInputHelper_GetGraphicInputs(ByVal pInsulation As StructInsulations.IStructInsulation, ByRef ppElesInputObjects As IJElements) As StructInsulations.StructInsulationGraphicInputHelperStatus
Const METHOD = "IStructInsulationGraphicInputHelper_GetGraphicInputs"
On Error GoTo ErrorHandler

    Dim status As StructInsulations.StructInsulationGraphicInputHelperStatus
    Dim oRC As IJDReferencesCollection
    Dim oObj As Object

    status = StructInsulationInputHelper_UnexpectedError

    Set oRC = GetRefCollection(pInsulation)
    If oRC.IJDEditJDArgument.GetCount > 2 Then
        Set oObj = oRC.IJDEditJDArgument.GetEntityByIndex(3)
        If TypeOf oObj Is ISPSMemberPartPrismatic Then
            Set ppElesInputObjects = New JObjectCollection
            ppElesInputObjects.Add oObj
            status = StructInsulationInputHelper_Ok
        Else
            status = StructInsulationInputHelper_InvalidTypeOfObject
        End If
    Else
        status = StructInsulationInputHelper_BadNumberOfObjects
    End If
    
    IStructInsulationGraphicInputHelper_GetGraphicInputs = status
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub IStructInsulationGraphicInputHelper_InitializeEntity(ByVal pInsulation As StructInsulations.IStructInsulation)
Const METHOD = "IStructInsulationGraphicInputHelper_InitializeEntity"
On Error GoTo ErrorHandler              ' this is called when the business object definition is set.
    pInsulation.DisplayAspects = DisplayAspect     ' initialize the display to use only the insulation aspect  32 = 2**5
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Function IStructInsulationGraphicInputHelper_SetGraphicInputs(ByVal pInsulation As StructInsulations.IStructInsulation, ByVal elesInputObjects As IJElements) As StructInsulations.StructInsulationGraphicInputHelperStatus
Const METHOD = "IStructInsulationGraphicInputHelper_SetGraphicInputs"
On Error GoTo ErrorHandler

    Dim status As StructInsulations.StructInsulationGraphicInputHelperStatus
    Dim oRC As IJDReferencesCollection
    Dim oObj As Object

    status = StructInsulationInputHelper_UnexpectedError

    Set oRC = GetRefCollection(pInsulation)
    If elesInputObjects.count > 0 Then
        Set oObj = elesInputObjects(1)
        If TypeOf oObj Is ISPSMemberPartPrismatic Then
    
            oRC.IJDEditJDArgument.SetEntity 3, oObj, IJGeometry, "StructInsulGeometryRefColl_D" ' update on Geometry
            oRC.IJDEditJDArgument.SetEntity 4, oObj, ISPSMemberType, "MemberTypeRC_DEST"        ' update on MemberType
            
            AddRelationship oObj, IStructHasInsulation, pInsulation, "Insulation_O", vbNull, False
            
            status = StructInsulationInputHelper_Ok

        Else
            status = StructInsulationInputHelper_InvalidTypeOfObject
        End If
        
    Else
        status = StructInsulationInputHelper_BadNumberOfObjects
    End If

    IStructInsulationGraphicInputHelper_SetGraphicInputs = status
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub SetNameRule(BOProgId As String, obj As Object)
Const METHOD = "SetNameRule"
On Error GoTo ErrorHandler
 
    Dim NameRule As String
    Dim found As Boolean
    found = False
    On Error Resume Next
      
    Dim NamingRules As IJElements
    Dim oNameRuleHolder As GSCADGenericNamingRulesFacelets.IJDNameRuleHolder
    Dim oActiveNRHolder As GSCADGenericNamingRulesFacelets.IJDNameRuleHolder
    Dim oNameRuleHlpr As GSCADNameRuleSemantics.IJDNamingRulesHelper
    Set oNameRuleHlpr = New GSCADNameRuleHlpr.NamingRulesHelper
    Call oNameRuleHlpr.GetEntityNamingRulesGivenProgID(BOProgId, NamingRules)
    Dim ncount As Integer
    Dim oNameRuleAE As GSCADNameRuleSemantics.IJNameRuleAE
      
    For ncount = 1 To NamingRules.count
        Set oNameRuleHolder = NamingRules.Item(1)
    Next ncount
 
    Call oNameRuleHlpr.AddNamingRelations(obj, oNameRuleHolder, oNameRuleAE)
    Set oNameRuleHolder = Nothing
    
    Set oActiveNRHolder = Nothing
    Set oNameRuleHolder = Nothing
    Set oNameRuleAE = Nothing
 
    Exit Sub
 
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Function ValidateClassWithSpec(iInsulation As IStructInsulation) As Long
Const METHOD = "ValidateClassWithSpec"
On Error GoTo ErrorHandler
    
    ' The name of the item's parent smart class should be same as spec encasement rule.
    ' If it is not, then probably selection rule could not find an item for the currently specified spec.
    ' eg: one encasement class has items for member types that another does not.
    
    Dim strSpecEncasementRule As String
    Dim strClassName As String
    Dim structSpecification As IStructInsulationSpec

    ValidateClassWithSpec = 0

    Set structSpecification = iInsulation.Specification

    If Not structSpecification Is Nothing Then
        strSpecEncasementRule = structSpecification.EncasementRule
        strClassName = iInsulation.DefinitionClassName
    
        If strClassName = strSpecEncasementRule Then
            ValidateClassWithSpec = 1
        End If
    End If

    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function
Private Sub DebugExercise(testCase As Long, ByRef setbackRef1 As Long, ByRef setbackRef2 As Long)

    If testCase = 1 Then
        setbackRef1 = SETBACKREFERENCE_AXISSTART
        setbackRef2 = SETBACKREFERENCE_AXISSTART
    ElseIf testCase = 2 Then
        setbackRef1 = SETBACKREFERENCE_AXISSTART
        setbackRef2 = SETBACKREFERENCE_PARTSTART
    ElseIf testCase = 3 Then
        setbackRef1 = SETBACKREFERENCE_AXISSTART
        setbackRef2 = SETBACKREFERENCE_AXISEND
    ElseIf testCase = 4 Then
        setbackRef1 = SETBACKREFERENCE_AXISSTART
        setbackRef2 = SETBACKREFERENCE_PARTEND
    ElseIf testCase = 5 Then
        setbackRef1 = SETBACKREFERENCE_PARTSTART
        setbackRef2 = SETBACKREFERENCE_AXISSTART
    ElseIf testCase = 6 Then
        setbackRef1 = SETBACKREFERENCE_PARTSTART
        setbackRef2 = SETBACKREFERENCE_PARTSTART
    ElseIf testCase = 7 Then
        setbackRef1 = SETBACKREFERENCE_PARTSTART
        setbackRef2 = SETBACKREFERENCE_AXISEND
    ElseIf testCase = 8 Then
        setbackRef1 = SETBACKREFERENCE_PARTSTART
        setbackRef2 = SETBACKREFERENCE_PARTEND
    ElseIf testCase = 9 Then
        setbackRef1 = SETBACKREFERENCE_AXISEND
        setbackRef2 = SETBACKREFERENCE_AXISSTART
    ElseIf testCase = 10 Then
        setbackRef1 = SETBACKREFERENCE_AXISEND
        setbackRef2 = SETBACKREFERENCE_PARTSTART
    ElseIf testCase = 11 Then
        setbackRef1 = SETBACKREFERENCE_AXISEND
        setbackRef2 = SETBACKREFERENCE_AXISEND
    ElseIf testCase = 12 Then
        setbackRef1 = SETBACKREFERENCE_AXISEND
        setbackRef2 = SETBACKREFERENCE_PARTEND
    ElseIf testCase = 13 Then
        setbackRef1 = SETBACKREFERENCE_PARTEND
        setbackRef2 = SETBACKREFERENCE_AXISSTART
    ElseIf testCase = 14 Then
        setbackRef1 = SETBACKREFERENCE_PARTEND
        setbackRef2 = SETBACKREFERENCE_PARTSTART
    ElseIf testCase = 15 Then
        setbackRef1 = SETBACKREFERENCE_PARTEND
        setbackRef2 = SETBACKREFERENCE_AXISEND
    ElseIf testCase = 16 Then
        setbackRef1 = SETBACKREFERENCE_PARTEND
        setbackRef2 = SETBACKREFERENCE_PARTEND
    Else
        setbackRef1 = SETBACKREFERENCE_AXISSTART
        setbackRef2 = SETBACKREFERENCE_AXISEND
    End If

End Sub

Private Function IStructInsulationValidationHelper_ValidateInputs(ByVal oGraphicObject As Object, ByVal oDefinition As Object) As StructInsulations.StructInsulationGraphicInputHelperStatus
    IStructInsulationValidationHelper_ValidateInputs = ValidateMemberFPCriteria(oGraphicObject, oDefinition)
End Function
