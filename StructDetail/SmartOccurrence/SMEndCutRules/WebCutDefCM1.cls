VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebCutDefCM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMEndCutRules
'  File        : WebCutDefCM1.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    12/Jan/2012 - svsmylav
'           CR-174918: Updated CMCornerFeatureRadiusCon not to create CF for
'            for profile pseudo-knuckle, Convex case.
'    15/Feb/2012 - svsmylav
'           Removed earlier changes done for CR-174918 since selector custom method takes care of this.
'*********************************************************************************************

Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "EndCutRules\WebCutDefCM1.cls"

Private sError As String
Private sMETHOD As String

'***********************************************************************
' METHOD:  CMCornerFeature
'
' DESCRIPTION:  Create CornerFeature
'
'
'***********************************************************************
Public Sub CMCornerFeatureRadius(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_CornerFeature(pMemberDescription, pResourceManager, "RootCorner")

End Sub

'***********************************************************************
' METHOD:  CAConstruct_CornerFeature
'
' DESCRIPTION:  Create Scallop Corner Feature
'
'
'***********************************************************************
Public Function CAConstruct_CornerFeature(ByVal pMemberDescription As IJDMemberDescription, _
                                    ByVal pResourceManager As IUnknown, _
                                    strStartClass As String) As Object

    sMETHOD = "CAConstruct_CornerFeature"
    
    Dim bIsTopCorner As Boolean
    
    Dim oSystemParent As IJSystemChild
    Dim oProfFacePort As IJPort
    Dim oProfEdgePort1 As IJPort
    Dim oProfEdgePort2 As IJPort
    Dim oCornerPosition As IJDPosition
    
    Dim oEndCutObject As Object
    Dim oStructFeature As IJStructFeature
       
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim oSDO_CornerFeature As StructDetailObjects.CornerFeature
        
    Set oEndCutObject = pMemberDescription.CAO
    ' Initialize WebCut wrapper class
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oEndCutObject

    'Retrieve Ports to be used to create Corner Feature
    sError = "Retrieve Ports to be used to create WebCut Corner Feature"
    bIsTopCorner = False
    oSDO_WebCut.CornerFeatureData bIsTopCorner, oCornerPosition, _
                                  oProfFacePort, oProfEdgePort1, oProfEdgePort2

    ' Get the End cut, make it a parent of the Corner Feature
    sError = "Setting system parent to Member Description Custom Assembly"
    Set oSystemParent = oEndCutObject
    
    ' Create Corner Feature
    sError = "Creating Corner Feature"
    Set oSDO_CornerFeature = New StructDetailObjects.CornerFeature
    oSDO_CornerFeature.Create pResourceManager, _
                               oProfFacePort, _
                               oProfEdgePort1, _
                               oProfEdgePort2, _
                               strStartClass, _
                               oSystemParent
                               
    sError = "Returning CornerFeature just created"
    Set CAConstruct_CornerFeature = oSDO_CornerFeature.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
        
End Function


'***********************************************************************
' METHOD:  CMCornerFeatureRadiusCon
'
' DESCRIPTION:  Conditional for Radius Corner Feature
'
'***********************************************************************
Public Sub CMCornerFeatureRadiusCon(ByRef pMD As IJDMemberDescription, _
                              ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    
    bIsNeeded = False 'Initialize
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_StiffenerEndCutCornerFeature) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sSectionType As String
    
    Dim oCutObject As Object
    
    'get the question answer for feature type
    Dim vAnswer As Variant
    Dim sFeatureType As String
    
    Dim oObject As IJDObject
    Dim oResourceManagerUnknown As IUnknown
    
    Dim oDefinitions As IMSSymbolEntities.IJDDefinitions
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim oParentSmartOccurrence As DEFINITIONHELPERSINTFLib.IJSmartOccurrence
    Dim oSymbolEntitiesFactory As IMSSymbolEntities.IJDSymbolEntitiesFactory
    
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    
    sFeatureType = ""
    
    Set oParentSmartOccurrence = pMD.CAO
 
    Dim oSmartItem As IJSmartItem
    Dim oParentSmartClass As IJSmartClass
    
    Set oSmartItem = oParentSmartOccurrence.ItemObject
    Set oParentSmartClass = oSmartItem.Parent
    
    Set oSymbolDefinition = oParentSmartClass.SelectionRuleDef
        
    On Error GoTo ErrorHandler

    Set oCommonHelper = New DefinitionHlprs.CommonHelper
    vAnswer = oCommonHelper.GetAnswer(oParentSmartOccurrence, _
                                      oSymbolDefinition, _
                                      "FeatureType")
    sFeatureType = vAnswer
    
    sMETHOD = "CMCornerFeatureCon"
    sError = "Setting Conditional value"
    bIsNeeded = True
    Set oCutObject = pMD.CAO
    
    If TypeOf oCutObject Is IJStructFeature Then
        Dim oFeature As IJStructFeature
        Set oFeature = oCutObject
        
        If oFeature.get_StructFeatureType = SF_WebCut Then
            If LCase(sFeatureType) = LCase("None") Then
                bIsNeeded = False
            ElseIf LCase(sFeatureType) = LCase("Radius") Then
                bIsNeeded = True
            ElseIf LCase(sFeatureType) = LCase("Snipe") Then
                bIsNeeded = False
            End If
        Else 'the feature is a flange cut
            bIsNeeded = False
        End If
    End If
 
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
    End Sub
    
    
Public Sub CMGetEdgeTreatment(ByRef pMD As IJDMemberDescription, _
                          ByRef bIsNeeded As Boolean)
 On Error GoTo ErrorHandler
    bIsNeeded = False
     
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_FreeEdgeTreatment) Then
        bIsNeeded = False
        Exit Sub
    End If
     
     
    Dim vAnswer As Variant
    Dim sApplyTreatment As String
    
    Dim oObject As IJDObject
    Dim oResourceManagerUnknown As IUnknown
    
    Dim oDefinitions As IMSSymbolEntities.IJDDefinitions
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim oParentSmartOccurrence As DEFINITIONHELPERSINTFLib.IJSmartOccurrence
    Dim oSymbolEntitiesFactory As IMSSymbolEntities.IJDSymbolEntitiesFactory
    
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    
    sApplyTreatment = ""
    
    Set oParentSmartOccurrence = pMD.CAO
 
    Dim oSmartItem As IJSmartItem
    Dim oParentSmartClass As IJSmartClass
    
    Set oSmartItem = oParentSmartOccurrence.ItemObject
    Set oParentSmartClass = oSmartItem.Parent
    
    Set oSymbolDefinition = oParentSmartClass.SelectionRuleDef
        
    On Error GoTo ErrorHandler

    Set oCommonHelper = New DefinitionHlprs.CommonHelper
    vAnswer = oCommonHelper.GetAnswer(oParentSmartOccurrence, _
                                      oSymbolDefinition, _
                                      "ApplyTreatment")
    sApplyTreatment = vAnswer
    
    If LCase(sApplyTreatment) = LCase("No") Then
        bIsNeeded = False
    ElseIf LCase(sApplyTreatment) = LCase("Yes") Then
        bIsNeeded = True
    End If
    
    Exit Sub
  
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "CMGetEdgeTreatment").Number
End Sub


'***********************************************************************
' METHOD:  CMCornerFeatureSnipe
'
' DESCRIPTION:  Create sniped CornerFeature
'
'
'***********************************************************************
Public Sub CMCornerFeatureSnipe(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_CornerFeature(pMemberDescription, pResourceManager, "SmartSnipe")

End Sub

'***********************************************************************
' METHOD:  CMCornerFeatureSnipeCon
'
' DESCRIPTION:  Conditional for Snipe Corner Feature
'
'***********************************************************************
Public Sub CMCornerFeatureSnipeCon(ByRef pMD As IJDMemberDescription, _
                              ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_StiffenerEndCutCornerFeature) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sSectionType As String
    
    Dim oCutObject As Object
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    
    'get the question answer for feature type
    Dim vAnswer As Variant
    Dim sFeatureType As String
    
    Dim oObject As IJDObject
    Dim oResourceManagerUnknown As IUnknown
    
    Dim oDefinitions As IMSSymbolEntities.IJDDefinitions
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim oParentSmartOccurrence As DEFINITIONHELPERSINTFLib.IJSmartOccurrence
    Dim oSymbolEntitiesFactory As IMSSymbolEntities.IJDSymbolEntitiesFactory
    
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    
    sFeatureType = ""
    
    Set oParentSmartOccurrence = pMD.CAO
 
    Dim oSmartItem As IJSmartItem
    Dim oParentSmartClass As IJSmartClass
    
    Set oSmartItem = oParentSmartOccurrence.ItemObject
    Set oParentSmartClass = oSmartItem.Parent
    
    Set oSymbolDefinition = oParentSmartClass.SelectionRuleDef
        
    On Error GoTo ErrorHandler

    Set oCommonHelper = New DefinitionHlprs.CommonHelper
    vAnswer = oCommonHelper.GetAnswer(oParentSmartOccurrence, _
                                      oSymbolDefinition, _
                                      "FeatureType")
    sFeatureType = vAnswer
    
    sMETHOD = "CMCornerFeatureCon"
    sError = "Setting Conditional value"
    bIsNeeded = True
    Set oCutObject = pMD.CAO
    
    If TypeOf oCutObject Is IJStructFeature Then
        Dim oFeature As IJStructFeature
        Set oFeature = oCutObject
        
        If oFeature.get_StructFeatureType = SF_WebCut Then
            If LCase(sFeatureType) = LCase("None") Then
                bIsNeeded = False
            ElseIf LCase(sFeatureType) = LCase("Radius") Then
                bIsNeeded = False
            ElseIf LCase(sFeatureType) = LCase("Snipe") Then
                bIsNeeded = True
            End If
        Else 'the feature is a flange cut
            bIsNeeded = False
        End If
    End If
 
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
    End Sub

