Attribute VB_Name = "FeatureCM"
'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  File        : FeatureCM.bas
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'   20/11/2012 - GH -CR-222498
'           Added New method 'GetFeaturePositionAttr'
'
'*********************************************************************************************

Option Explicit

Private sMETHOD As String
Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\SharedVB\FeatureCM.bas"

Public Function CMConstructFETForFeature(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown) As Object
 On Error GoTo ErrorHandler
    sMETHOD = "CMConstructFETForFeature"
    On Error GoTo ErrorHandler
    
    Dim oFeature As Object
    Dim oPartWithFeature As Object
    Dim oPort1 As IJPort
    
    ' Get edge feature
    Set oFeature = pMemberDescription.CAO
    
    ' Get the cut out part
     Set oPartWithFeature = GetFeaturePartFromFeature(oFeature)
    
    ' Get port created by edge cut
    If TypeOf oPartWithFeature Is IJProfilePart Then
        Dim oProfilePartWrapper As New StructDetailObjects.ProfilePart
        Set oProfilePartWrapper.object = oPartWithFeature
        Set oPort1 = oProfilePartWrapper.CutoutPort(oFeature)
    ElseIf TypeOf oPartWithFeature Is IJPlatePart Then
        Dim oPlatePartWrapper As New StructDetailObjects.PlatePart
        Set oPlatePartWrapper.object = oPartWithFeature
        Set oPort1 = oPlatePartWrapper.CutoutPort(oFeature)
    ElseIf TypeOf oPartWithFeature Is ISPSMemberPartPrismatic Then
        Dim oMbrPartWrapper As New StructDetailObjectsEx.SDOMemberPart
        Set oMbrPartWrapper.object = oPartWithFeature
        Set oPort1 = oMbrPartWrapper.CutoutPort(oFeature)
    End If
    
    If Not oPort1 Is Nothing Then
        ' Create Free Edge Treatment
        Dim strStartClass As String
        Dim oFreeEdgeTreatmentWrapper As New StructDetailObjects.FreeEdgeTreatment
        
        strStartClass = "RootEdgeTreatment"
        oFreeEdgeTreatmentWrapper.Create pResourceManager, _
                                          oPort1, _
                                          Nothing, Nothing, _
                                          strStartClass, _
                                          oFeature
        
        Set CMConstructFETForFeature = oFreeEdgeTreatmentWrapper.object
        
        Set oFreeEdgeTreatmentWrapper = Nothing
        Set oPort1 = Nothing
    End If
    
    Set oFeature = Nothing

  Exit Function
  
ErrorHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

Public Function GetAnswer_ApplyTreatment(ByRef pMD As IJDMemberDescription, eFeatureType As SmartClassType) As String
    sMETHOD = "GetAnswer_ApplyTreatment"
   
    On Error GoTo ErrorHandler
    
    Dim vAnswer As Variant
    Dim sApplyTreatment As String
    
    Dim oObject As IJDObject
    Dim oResourceManagerUnknown As IUnknown
    
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim oParentSmartOccurrence As DEFINITIONHELPERSINTFLib.IJSmartOccurrence
    
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    
    sApplyTreatment = ""
    
    Set oParentSmartOccurrence = pMD.CAO
 
    Dim oParentSmartClass As IJSmartClass
    Set oParentSmartClass = GetSOFeatureRootClass(eFeatureType)
    
    Set oSymbolDefinition = oParentSmartClass.SelectionRuleDef
     
       
    On Error GoTo ErrorHandler

    Set oCommonHelper = New DefinitionHlprs.CommonHelper
    vAnswer = oCommonHelper.GetAnswer(oParentSmartOccurrence, _
                                      oSymbolDefinition, _
                                      "ApplyTreatment")
    sApplyTreatment = vAnswer
    
    GetAnswer_ApplyTreatment = sApplyTreatment
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

Private Function GetFeaturePartFromFeature(oFeature As Object) As Object
    On Error GoTo ErrorHandler
    sMETHOD = "GetFeaurePartFromFeature"
    
    If TypeOf oFeature Is IJStructFeature Then
        Dim oStructFeature As IJStructFeature
        Dim nFeatureType As StructFeatureTypes
        Set oStructFeature = oFeature
        nFeatureType = oStructFeature.get_StructFeatureType
        
        Set oStructFeature = Nothing
         
        Select Case nFeatureType
        
        Case SF_EdgeFeature
            ' Set up edge feature wrapper
            Dim oEdgeFeatureWrapper As StructDetailObjects.EdgeFeature
            Set oEdgeFeatureWrapper = New StructDetailObjects.EdgeFeature
            
            Set oEdgeFeatureWrapper.object = oFeature
            Set GetFeaturePartFromFeature = oEdgeFeatureWrapper.GetPartObject
            
            Set oEdgeFeatureWrapper = Nothing
        Case SF_CornerFeature
            Dim oCornerFeatureWrapper As StructDetailObjects.CornerFeature
            Set oCornerFeatureWrapper = New StructDetailObjects.CornerFeature
            
            Set oCornerFeatureWrapper.object = oFeature
            Set GetFeaturePartFromFeature = oCornerFeatureWrapper.GetPartObject
            
            Set oCornerFeatureWrapper = Nothing
        Case SF_Slot
            Dim oSlotFeatureWrapper As StructDetailObjects.Slot
            Set oSlotFeatureWrapper = New StructDetailObjects.Slot
            
            Set oSlotFeatureWrapper.object = oFeature
            Set GetFeaturePartFromFeature = oSlotFeatureWrapper.Penetrated
            
            Set oSlotFeatureWrapper = Nothing
        Case SF_FlangeCut
            Dim oFlangeCutFeatureWrapper As StructDetailObjects.FlangeCut
            Set oFlangeCutFeatureWrapper = New StructDetailObjects.FlangeCut
            
            Set oFlangeCutFeatureWrapper.object = oFeature
            Set GetFeaturePartFromFeature = oFlangeCutFeatureWrapper.Bounded
            
            Set oFlangeCutFeatureWrapper = Nothing
        
        Case SF_WebCut
            Dim oWebCutFeatureWrapper As StructDetailObjects.WebCut
            Set oWebCutFeatureWrapper = New StructDetailObjects.WebCut
            
            Set oWebCutFeatureWrapper.object = oFeature
            Set GetFeaturePartFromFeature = oWebCutFeatureWrapper.Bounded
            
            Set oWebCutFeatureWrapper = Nothing
        
        End Select
        
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function
Private Function GetSOFeatureRootClass(eFeatureType As SmartClassType) As IJSmartClass
    Const METHOD = "GetSOFeatureRootClass"
    
    Dim oCatalogQuery As IJSRDQuery
    Dim oSOFeatureClassQuery As IJSmartQuery
    Dim oSORootClass As IJSmartClass
    
    On Error GoTo ErrorHandler
    
    'Check if we are ready to compute.  If so, proceed.
    Set oCatalogQuery = New SRDQuery
    Set oSOFeatureClassQuery = oCatalogQuery
    Set oSORootClass = oSOFeatureClassQuery.GetRootClass(eFeatureType)
    
    Set GetSOFeatureRootClass = oSORootClass
    
    Set oCatalogQuery = Nothing
    Set oSOFeatureClassQuery = Nothing
    Set oSORootClass = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

Public Sub GetFeaturePositionAttr(ByVal oFeature As IJDObject, ByRef eFeaturePositionAttr As SD_FeaturePositionAttr, _
                                                        Optional ByRef bIsRightSideExtended As Boolean = False)
    Const METHOD = "GetFeaturePositionAttr"
    On Error GoTo ErrorHandler
    
    Dim oEdgePort As IJPort
    Dim oLocation As Object
    
    Dim oFeatureUtils As IJSDFeatureAttributes
    Set oFeatureUtils = New SDFeatureUtils

    'Get Feature Input
    oFeatureUtils.Get_Inputs_EdgeCut oFeature, oEdgePort, oLocation
    
    Dim dx As Double
    Dim dy As Double
    Dim dz As Double
    
    Dim oLocationPos As IJDPosition
    Dim oRefGeometry As Object
    If TypeOf oLocation Is IJPoint Then
        'the graphical input, "Point" is a Point3d Object which does not support IJDPosition
        'get the coordinates from IJPoint, which is the default interface of the Point3d object
        'put the coordinates obtained above into IJDPosition
        Dim oPoint3d As IJPoint
        Set oPoint3d = oLocation
        oPoint3d.GetPoint dx, dy, dz
        Set oLocationPos = New AutoMath.DPosition
        oLocationPos.Set dx, dy, dz
    ElseIf TypeOf oLocation Is IJDPosition Then
        Set oLocationPos = oLocation
    Else
        If TypeOf oLocation Is IJPort Then
            Dim oPort As IJPort
            Set oPort = oLocation
            Set oRefGeometry = oPort.Geometry
        Else
            ' Given Edge Location is NOT IJPoint or IJPort object
            ' ... Expect Edge Location to be a IJSurfaceBody or IJWireBody object
            Set oRefGeometry = oLocation
        End If
        
        If oRefGeometry Is Nothing Then
            '
        ElseIf TypeOf oRefGeometry Is IJSurfaceBody Or TypeOf oRefGeometry Is IJWireBody Then
            Dim oModelBodyUtil As IJSGOModelBodyUtilities
            Set oModelBodyUtil = New SGOModelBodyUtilities
            
            Dim oEdgePortGeom As IJWireBody
            Set oEdgePortGeom = oEdgePort.Geometry
            Dim oPointOnSurface As IJDPosition
            Dim dDistance As Double
            'Get point on EdgePort Closest to Input Port
            oModelBodyUtil.GetClosestPointsBetweenTwoBodies oEdgePortGeom, _
                                                   oRefGeometry, _
                                                   oLocationPos, _
                                                   oPointOnSurface, _
                                                   dDistance
        End If
    End If
    
                    
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
    
    'Get Feature Position Attribute
    oTopologyLocate.EdgeFeatureIsAtSeam oFeature, oEdgePort, oLocationPos, eFeaturePositionAttr, bIsRightSideExtended
    
  Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
            
End Sub

