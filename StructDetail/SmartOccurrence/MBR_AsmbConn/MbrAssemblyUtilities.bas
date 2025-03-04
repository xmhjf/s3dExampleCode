Attribute VB_Name = "MbrAssemblyUtilities"
'*******************************************************************
'
'Copyright (C) 2007 Intergraph Corporation. All rights reserved.
'
'File : MbrAssemblyUtilities.bas
'
'Author : D.A. Trent
'
'Description :
'   Utilites for used by the Member Assembly Connection Rules
'
'History:
'*****************************************************************************
Option Explicit
Private Const MODULE = "StructDetail\Data\SmartOccurrence\MemberAssyConn\MbrAssemblyUtilities"
'


'*************************************************************************
'Function
'Create_WebCut
'
'Abstract
'   Create WebCut given :
'       the Given the MemberDescription and Root Selection Rule
'
'input
'   pMemberDescription
'   pResourceManager
'   sEndCutSelRule
'   bUseBoundingEndPort
'   lEndCutSet
'
'Return
'   pEndCutObject
'
'Exceptions
'
'***************************************************************************
Public Sub Create_WebCut(pMemberDescription As IJDMemberDescription, _
                         pResourceManager As IUnknown, _
                         sEndCutSelRule As String, _
                         bUseBoundingEndPort As Boolean, _
                         bUseBoundedAlongPort As Boolean, _
                         lEndCutSet As Long, _
                         pEndCutObject As Object)
Const METHOD = "MbrAssemblyUtilities::Create_WebCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim lDispId As Long
    Dim lStatus As Long
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    Dim oPort As Object
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oAppConnection As IJAppConnection

    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    Dim oSDO_WebCut As StructDetailObjects.WebCut

    sMsg = "Creating WebCut ...pMemberDescription.index = " & Str(pMemberDescription.index)
    lDispId = pMemberDescription.dispid
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    sMsg = "Initializing End Cut data from IJAppConnection"
    Set oAppConnection = pMemberDescription.CAO
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, _
                             lStatus, sMsg
    
    Set oBoundedPort = oBoundedData.AxisPort
    If bUseBoundedAlongPort Then
        ' By default, the Bounded Port is the AxisStart or AxisEnd Port
        ' For ShortBox cases, need the Bounded AxisLong Port
        Set oBoundedPort = oBoundedData.MemberPart.AxisPort(SPSMemberAxisAlong)
    End If
    
    Set oBoundingPort = oBoundingData.AxisPort
    If bUseBoundingEndPort Then
        ' By default, the Bounding Port is the AlongAxis Port
        ' For End to End cases, need the Bounding End Port
        GetSupportingEndPort oBoundedData, oBoundingData, oBoundingPort
    End If
    
    If lEndCutSet = 2 Then
        ' For the Second set of End Cuts
        ' Switch the Bounding and Bounded Ports used to create the EndCut
        Set oPort = oBoundedPort
        Set oBoundedPort = oBoundingPort
        Set oBoundingPort = oPort
        Set oPort = Nothing
    End If
    
    ' Need to get the IJSystem Interface from ths CommonStruct AssemblyConnection
    sMsg = "Retreiving Parent System for WebCut"
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
    
    ' Create the Web Cut
    sMsg = "StructDetailObjects.WebCut::Create ...Creating Web Cut Feature"
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    oSDO_WebCut.Create pResourceManager, oBoundingPort, oBoundedPort, _
                       sEndCutSelRule, oSystemParent
                               
    sMsg = "Return the created Web Cut"
    Set pEndCutObject = oSDO_WebCut.object
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'Create_FlangeCut
'
'Abstract
'   Create FlangeCut given :
'       the Given the MemberDescription and Root Selection Rule
'
'input
'   pMemberDescription
'   pResourceManager
'   sEndCutSelRule
'   bUseBoundingEndPort
'   lEndCutSet
'
'Return
'   pEndCutObject
'
'Exceptions
'
'***************************************************************************
Public Sub Create_FlangeCut(pMemberDescription As IJDMemberDescription, _
                            pResourceManager As IUnknown, _
                            sEndCutSelRule As String, _
                            bUseBoundingEndPort As Boolean, _
                            bUseBoundedAlongPort As Boolean, _
                            lEndCutSet As Long, _
                            pEndCutObject As Object)
Const METHOD = "MbrAssemblyUtilities::Create_FlangeCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim lDispId As Long
    Dim lStatus As Long
    Dim nWebCuts As Long
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    Dim oPort As Object
    Dim oWebCut As Object
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oAppConnection As IJAppConnection
    Dim oMemberObjects As IJDMemberObjects
    Dim oStructFeature As IJStructFeature

    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut

    sMsg = "Creating FlangeCut ...pMemberDescription.index = " & Str(pMemberDescription.index)
    lDispId = pMemberDescription.dispid
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    sMsg = "Initializing End Cut data from IJAppConnection"
    Set oAppConnection = pMemberDescription.CAO
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, _
                             lStatus, sMsg
    
    Set oBoundedPort = oBoundedData.AxisPort
    If bUseBoundedAlongPort Then
        ' By default, the Bounded Port is the AxisStart or AxisEnd Port
        ' For ShortBox cases, need the Bounded AxisLong Port
        Set oBoundedPort = oBoundedData.MemberPart.AxisPort(SPSMemberAxisAlong)
    End If
    
    Set oBoundingPort = oBoundingData.AxisPort
    If bUseBoundingEndPort Then
        ' By default, the Bounding Port is the AlongAxis Port
        ' For End to End cases, need the Bounding End Port
        GetSupportingEndPort oBoundedData, oBoundingData, oBoundingPort
    End If
    

    If lEndCutSet = 2 Then
        ' For the Second set of End Cuts
        ' Switch the Bounding and Bounded Ports used to create the EndCut
        Set oPort = oBoundedPort
        Set oBoundedPort = oBoundingPort
        Set oBoundingPort = oPort
        Set oPort = Nothing
    End If
    
    ' Need to get the IJSystem Interface from ths CommonStruct AssemblyConnection
    sMsg = "Retreiving Parent System for FlangeCut"
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
    
    ' Create the Flange Cut
    sMsg = "StructDetailObjects.FlangeCut::Create ...Creating Flange Cut Feature"
    nWebCuts = 0
    Set oMemberObjects = oAppConnection
    For iIndex = 1 To oMemberObjects.Count
        If Not oMemberObjects.Item(iIndex) Is Nothing Then
            If TypeOf oMemberObjects.Item(iIndex) Is IJStructFeature Then
                Set oStructFeature = oMemberObjects.Item(iIndex)
                If oStructFeature.get_StructFeatureType = SF_WebCut Then
                    nWebCuts = nWebCuts + 1
                    If nWebCuts = lEndCutSet Then
                        Set oWebCut = oStructFeature
                        Exit For
                    End If
                End If
            End If
        End If
    Next iIndex
        
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    oSDO_FlangeCut.Create pResourceManager, oBoundingPort, oBoundedPort, _
                          oWebCut, sEndCutSelRule, oSystemParent
                               
    sMsg = "Return the created Flange Cut"
    Set pEndCutObject = oSDO_FlangeCut.object
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'Set_WebCutQuestions
'
'Abstract
'   Copy Answers from Parent to End Cut given :
'       the Given the MemberDescription and
'       EndCut Selection Rule Name and
'       EndCut Selection/Item Rule Prog Id and
'       Parent Assembly Connection Selection/Item Rule Prog Id
'
'input
'   pMemberDescription
'   sEndCutSel
'   sEndCutProgId
'   sParentProgId
'   bSwitchWeldPart
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub Set_WebCutQuestions(pMemberDescription As IJDMemberDescription, _
                               sEndCutSel As String, _
                               sEndCutProgId As String, _
                               sParentProgId As String, _
                               bSwitchWeldPart As Boolean, _
                               bForceUpdate As Boolean)
Const METHOD = "MbrAssemblyUtilities::Set_WebCutQuestions"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sEndCut As String
    Dim sWeldPart As String
    
    Dim lDispId As Long

    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    Dim oCopyAnswerHelper As DefinitionHlprs.CopyAnswerHelper
    
    sMsg = "Initializing End Cut data"
    lDispId = pMemberDescription.dispid
    
    sMsg = "Creating CopyAnswerHelper object"
    Set oCommonHelper = New DefinitionHlprs.CommonHelper
    Set oCopyAnswerHelper = New DefinitionHlprs.CopyAnswerHelper
    Set oCopyAnswerHelper.MemberDescription = pMemberDescription
    
    sMsg = "Copying EndCutType Question/Answer to " & sEndCutSel & _
           " .... MemberDescription.index = " & lDispId
    sEndCut = oCommonHelper.GetSelectorAnswer(pMemberDescription.CAO, _
                                              sParentProgId, "EndCutType")
    If Len(Trim(sEndCut)) < 1 Then
        ' EndCutType is not set in Parent Selection Rule
        ' Assume must have selected Item manually
        ' Default EndCutType to "W"
        sEndCut = "W"
    End If
    
    oCopyAnswerHelper.PutAnswer sEndCutProgId, "EndCutType", sEndCut
    
    sMsg = "Copying WeldPart Question/Answer to " & sEndCutSel & _
           " .... MemberDescription.index = " & lDispId
    sWeldPart = oCommonHelper.GetSelectorAnswer(pMemberDescription.CAO, _
                                                sParentProgId, "WeldPart")
    If Len(Trim(sWeldPart)) < 1 Then
        ' WeldPart is not set in Parent Selection Rule
        ' Assume must have selected Item manually
        ' Default Weldpart to "First"
        ' in these cases:
        ' bSwitchWeldPart argument should be able to control "First"/"Second"
        sWeldPart = "First"
    End If
    
    If bSwitchWeldPart Then
        If Trim(LCase(sWeldPart)) = LCase("First") Then
            sWeldPart = "Second"
        Else
            sWeldPart = "First"
        End If
    End If
    
    oCopyAnswerHelper.PutAnswer sEndCutProgId, "WeldPart", sWeldPart
    
    If bForceUpdate Then
        WebCut_ForceUpdate pMemberDescription
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'Set_FlangeCutQuestions
'
'Abstract
'   Copy Answers from Parent to End Cut given :
'       the Given the MemberDescription and
'       EndCut Selection Rule Name and
'       EndCut Selection/Item Rule Prog Id and
'       Parent Assembly Connection Selection/Item Rule Prog Id
'
'input
'   pMemberDescription
'   sEndCutSel
'   sEndCutProgId
'   sParentProgId
'   bSwitchWeldPart
'   bBottomFlange
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub Set_FlangeCutQuestions(pMemberDescription As IJDMemberDescription, _
                                  sEndCutSel As String, _
                                  sEndCutProgId As String, _
                                  sParentProgId As String, _
                                  bSwitchWeldPart As Boolean, _
                                  bBottomFlange As Boolean)
Const METHOD = "MbrAssemblyUtilities::Set_FlangeCutQuestions"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim lDispId As Long

    Dim oCopyAnswerHelper As DefinitionHlprs.CopyAnswerHelper
    
    sMsg = "Initializing End Cut data"
    lDispId = pMemberDescription.dispid
    
    Set_WebCutQuestions pMemberDescription, sEndCutSel, _
                        sEndCutProgId, sParentProgId, bSwitchWeldPart, False

    sMsg = "Creating CopyAnswerHelper object"
    Set oCopyAnswerHelper = New DefinitionHlprs.CopyAnswerHelper
    Set oCopyAnswerHelper.MemberDescription = pMemberDescription
    
    sMsg = "Setting BottomFlange Question/Answer to " & sEndCutSel & _
           " .... MemberDescription.index = " & lDispId
    If bBottomFlange Then
        oCopyAnswerHelper.PutAnswer sEndCutProgId, "BottomFlange", "Yes"
    Else
        oCopyAnswerHelper.PutAnswer sEndCutProgId, "BottomFlange", "No"
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'WebCut_ForceUpdate
'
'Abstract
' Force an Update on the WebCut using the same interface,IJStructGeometry,
' as is used when placing the WebCut as an input to the FlangeCuts
' This allows Assoc to always recompute the WebCut before FlangeCuts
'
'
'
'input
'   pMemberDescription
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub WebCut_ForceUpdate(pMemberDescription As IJDMemberDescription)
Const METHOD = "MbrAssemblyUtilities::WebCut_ForceUpdate"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim jIndex As Long
    Dim lDispId As Long
    
    Dim oWebCut As Object
    Dim oMemberItems As IJElements
    Dim oMemberObjects As IJDMemberObjects
    Dim oStructFeature As IJStructFeature
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    
    ' Force an Update on the WebCut using the same interface,IJStructGeometry,
    ' as is used when placing the WebCut as an input to the FlangeCuts
    ' This allows Assoc to always recompute the WebCut before FlangeCuts
    sMsg = "Calling Structdetailobjects.WebCut::ForceUpdateForFlangeCuts"
    lDispId = pMemberDescription.dispid
    Set oMemberObjects = pMemberDescription.CAO
        For iIndex = 1 To oMemberObjects.Count
            If Not oMemberObjects.Item(iIndex) Is Nothing Then
                If iIndex = lDispId Then
                    If TypeOf oMemberObjects.Item(iIndex) Is IJStructFeature Then
                        Set oStructFeature = oMemberObjects.Item(iIndex)
                        If oStructFeature.get_StructFeatureType = SF_WebCut Then
                            Set oWebCut = oStructFeature
                            Exit For
                        End If
                    End If
                End If
            End If
        Next iIndex
    
    If Not oWebCut Is Nothing Then
        Set oSDO_WebCut = New StructDetailObjects.WebCut
        Set oSDO_WebCut.object = oWebCut
        oSDO_WebCut.ForceUpdateForFlangeCuts
    
        Set oSDO_WebCut = Nothing
        Set oMemberObjects = Nothing
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'Create_BearingPlate
'
'Abstract
'   Create WebCut given :
'       the Given the MemberDescription and Root Selection Rule
'
'input
'   pMemberDescription
'   pResourceManager
'   sEndCutSelRule
'   bUseBoundingEndPort
'
'Return
'   pEndCutObject
'
'Exceptions
'
'***************************************************************************
Public Sub Create_BearingPlate(pMemberDescription As IJDMemberDescription, _
                               pResourceManager As IUnknown, _
                               sEndCutSelRule As String, _
                               bUseBoundingEndPort As Boolean, _
                               pBearingPlateObject As Object)
Const METHOD = "MbrAssemblyUtilities::Create_BearingPlate"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim lDispId As Long
    Dim lStatus As Long
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    Dim oBearingPlate As IJSmartPlate
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oAppConnection As IJAppConnection
    Dim oGraphicInputs As JCmnShp_CollectionAlias

    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    Dim oSPDefinition As GSCADSDCreateModifyUtilities.IJSDSmartPlateDefinition

    sMsg = "Creating WebCut ...pMemberDescription.index = " & Str(pMemberDescription.index)
    lDispId = pMemberDescription.dispid
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    sMsg = "Initializing End Cut data from IJAppConnection"
    Set oAppConnection = pMemberDescription.CAO
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, _
                             lStatus, sMsg
    
    Set oBoundedPort = oBoundedData.AxisPort
    Set oBoundingPort = oBoundingData.AxisPort
    If bUseBoundingEndPort Then
        GetSupportingEndPort oBoundedData, oBoundingData, oBoundingPort
    End If
    
    ' Need to get the IJSystem Interface from ths CommonStruct AssemblyConnection
    sMsg = "Retreiving Parent System for WebCut"
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
    
    ' Create the Bearing Plate
    sMsg = "...Creating Bearing Plate Object"
    Set oGraphicInputs = New Collection
    oGraphicInputs.Add oBoundingPort
    oGraphicInputs.Add oBoundedPort
    
    Set oSPDefinition = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
    Set oBearingPlate = oSPDefinition.CreateBearingPlatePart(pResourceManager, _
                                                             sEndCutSelRule, _
                                                             oGraphicInputs, _
                                                             oSystemParent)
    
    sMsg = "...Setting Bearing Plate Properties"
    SetPlatePartProperties oBearingPlate, pResourceManager, _
                           "CPlatePart", "BearingPlateCategory", _
                           NonTight, Standalone, _
                           "Steel - Carbon", "A", 0.01
            
    sMsg = "Return the created Bearing Plate"
    Set pBearingPlateObject = oBearingPlate
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'Create_BearingPlateWebCut
'
'Abstract
'   Create WebCut between the Bearing Plate and the Bounded member given :
'       the Given the MemberDescription and Root Selection Rule
'
'input
'   pMemberDescription
'   pResourceManager
'   sEndCutSelRule
'
'Return
'   pEndCutObject
'
'Exceptions
'
'***************************************************************************
Public Sub Create_BearingPlateWebCut(pMemberDescription As IJDMemberDescription, _
                                     pResourceManager As IUnknown, _
                                     sEndCutSelRule As String, _
                                     pEndCutObject As Object)
Const METHOD = "MbrAssemblyUtilities::Create_BearingPlateWebCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim lDispId As Long
    Dim lStatus As Long
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    Dim oPort As Object
    Dim oBearingPlate As Object
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oAppConnection As IJAppConnection
    Dim oMemberObjects As IJDMemberObjects

    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    Dim oSDO_WebCut As StructDetailObjects.WebCut

    sMsg = "Creating BearingPlate WebCut ...pMemberDescription.index = " & Str(pMemberDescription.index)
    lDispId = pMemberDescription.dispid
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    sMsg = "Initializing End Cut data from IJAppConnection"
    Set oAppConnection = pMemberDescription.CAO
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, _
                             lStatus, sMsg
    
    Set oBoundedPort = oBoundedData.AxisPort
    
    ' The End Cut is between the Bearing Plate and the Bounded Member
    ' Use the Offset Port from the BearingPlate (SmartPlate)
    sMsg = "Retreiving Offset Port from bearing Plate"
    Set oMemberObjects = oAppConnection
    For iIndex = 1 To oMemberObjects.Count
        If Not oMemberObjects.Item(iIndex) Is Nothing Then
            If TypeOf oMemberObjects.Item(iIndex) Is IJSmartPlate Then
                Set oBearingPlate = oMemberObjects.Item(iIndex)
                Exit For
            End If
        End If
    Next iIndex
    
    ' Get the Offset Port from the Bearing Plate
    ' This is the Bounding Port for the EndCut
    ' (require using Late Port Binding to retrieve the Port Moniker)
    BearingPlate_BasePort oBearingPlate, _
                          JS_TOPOLOGY_FILTER_SOLID_OFFSET_LFACE, _
                          oBoundingPort
    
    ' Need to get the IJSystem Interface from ths CommonStruct AssemblyConnection
    sMsg = "Retreiving Parent System for WebCut"
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
    
    ' Create the Web Cut
    sMsg = "StructDetailObjects.WebCut::Create ...Creating Web Cut Feature"
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    oSDO_WebCut.Create pResourceManager, oBoundingPort, oBoundedPort, _
                       sEndCutSelRule, oSystemParent
                               
    sMsg = "Return the created Web Cut"
    Set pEndCutObject = oSDO_WebCut.object
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'Create_BearingPlateFlangeCut
'
'Abstract
'   Create FlangeCut between the Bearing Plate and the Bounded member given :
'       the Given the MemberDescription and Root Selection Rule
'
'input
'   pMemberDescription
'   pResourceManager
'   sEndCutSelRule
'
'Return
'   pEndCutObject
'
'Exceptions
'
'***************************************************************************
Public Sub Create_BearingPlateFlangeCut(pMemberDescription As IJDMemberDescription, _
                                        pResourceManager As IUnknown, _
                                        sEndCutSelRule As String, _
                                        pEndCutObject As Object)
Const METHOD = "MbrAssemblyUtilities::Create_BearingPlateFlangeCut"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim iIndex As Long
    Dim lDispId As Long
    Dim lStatus As Long
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    Dim oWebCut As Object
    Dim oBearingPlate As Object
    Dim oSystemParent As IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oAppConnection As IJAppConnection
    Dim oMemberObjects As IJDMemberObjects
    Dim oStructFeature As IJStructFeature

    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut

    sMsg = "Creating FlangeCut ...pMemberDescription.index = " & Str(pMemberDescription.index)
    lDispId = pMemberDescription.dispid
    
    ' Get the Assembly Connection Ports from the IJAppConnection
    sMsg = "Initializing End Cut data from IJAppConnection"
    Set oAppConnection = pMemberDescription.CAO
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, _
                             lStatus, sMsg
    
    Set oBoundedPort = oBoundedData.AxisPort
    ' The End Cut is between the Bearing Plate and the Bounded Member
    ' Use the Offset Port from the BearingPlate (SmartPlate)
    sMsg = "Retreiving Offset Port from bearing Plate"
    Set oMemberObjects = oAppConnection
    For iIndex = 1 To oMemberObjects.Count
        If Not oMemberObjects.Item(iIndex) Is Nothing Then
            If TypeOf oMemberObjects.Item(iIndex) Is IJSmartPlate Then
                Set oBearingPlate = oMemberObjects.Item(iIndex)
                Exit For
            End If
        End If
    Next iIndex
    
    ' Get the Offset Port from the Bearing Plate
    ' This is the Bounding Port for the EndCut
    ' (require using Late Port Binding to retrieve the Port Moniker)
    BearingPlate_BasePort oBearingPlate, _
                          JS_TOPOLOGY_FILTER_SOLID_OFFSET_LFACE, _
                          oBoundingPort
    
    ' Need to get the IJSystem Interface from ths CommonStruct AssemblyConnection
    sMsg = "Retreiving Parent System for FlangeCut"
    If TypeOf oAppConnection Is IJDesignParent Then
        Set oDesignParent = oAppConnection
        If TypeOf oDesignParent Is IJSystem Then
            Set oSystemParent = oDesignParent
        End If
    End If
    
    ' Create the Flange Cut
    sMsg = "StructDetailObjects.FlangeCut::Create ...Creating Flange Cut Feature"
    Set oMemberObjects = oAppConnection
    For iIndex = 1 To oMemberObjects.Count
        If Not oMemberObjects.Item(iIndex) Is Nothing Then
            If TypeOf oMemberObjects.Item(iIndex) Is IJStructFeature Then
                Set oStructFeature = oMemberObjects.Item(iIndex)
                If oStructFeature.get_StructFeatureType = SF_WebCut Then
                    Set oWebCut = oStructFeature
                    Exit For
                End If
            End If
        End If
    Next iIndex
        
    Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
    oSDO_FlangeCut.Create pResourceManager, oBoundingPort, oBoundedPort, _
                          oWebCut, sEndCutSelRule, oSystemParent
                               
    sMsg = "Return the created Flange Cut"
    Set pEndCutObject = oSDO_FlangeCut.object
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'BearingPlate_BasePort
'
'Abstract
'   Given the BearingPlate, retrieve its Base or Offset Port
'   to be used in creating EndCuts and/or Physical Connections
'
'input
'   oBearingPlate
'   BasePortType
'           JS_TOPOLOGY_FILTER_SOLID_BASE_LFACE
'           JS_TOPOLOGY_FILTER_SOLID_OFFSET_LFACE
'Return
'   oPort
'
'Exceptions
'
'***************************************************************************
Public Sub BearingPlate_BasePort(oBearingPlate As Object, _
                                 BasePortType As JS_TOPOLOGY_FILTER_TYPE, _
                                 oLatePort As IJPort)
Const METHOD = "::BearingPlate_BasePort"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim oSDSmartPlateAtt As GSCADSDCreateModifyUtilities.IJSDSmartPlateAttributes
    
    Set oSDSmartPlateAtt = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
    oSDSmartPlateAtt.GetBasePort oBearingPlate, BasePortType, oLatePort
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'SetPlatePartProperties
'
'Abstract
'   Initialize/Set required PlatePart Properties
'
'input
'   oBearingPlate
'   oResourceManager
'   strEntity
'   strNamingCategoryTable
'   eTightness
'   ePlateType
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub SetPlatePartProperties(oBearingPlate As Object, _
                    oResourceManager As IUnknown, _
                    strEntity As String, _
                    strNamingCategoryTable As String, _
                    eTightness As GSCADShipGeomOps.StructPlateTightness, _
                    ePlateType As GSCADShipGeomOps.StructPlateType, _
                    strMatl As String, _
                    strGrade As String, _
                    dThickness As Double)
Const METHOD = "MbrAssemblyUtilities::SetPlatePartProperties"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim strLongNames() As String
    Dim strShortNames() As String

    Dim iIndex As Long
    Dim lPriority() As Long

    Dim oPlate As IJPlate
    Dim oMoldedConv As IJDPlateMoldedConventions
    
    Dim oRules As IJElements
    Dim oDummyAE As IJNameRuleAE
    Dim oQueryUtil As IJMetaDataCategoryQuery
    Dim oNamingObject As IJDNamingRulesHelper
    
    'Retrieve first default naming rule
    Set oNamingObject = New NamingRulesHelper
    oNamingObject.GetEntityNamingRulesGivenName strEntity, oRules
    If oRules.Count >= 1 Then
        oNamingObject.AddNamingRelations oBearingPlate, oRules.Item(1), oDummyAE
    End If
    Set oDummyAE = Nothing
    Set oNamingObject = Nothing
    
    ' Default naming category to first non-negative value
    Set oQueryUtil = New CMetaDataCategoryQuery
    oQueryUtil.GetCategoryInfo oResourceManager, _
                               strNamingCategoryTable, _
                               strLongNames, _
                               strShortNames, _
                               lPriority
    Set oQueryUtil = Nothing

    Set oPlate = oBearingPlate
    oPlate.NamingCategory = -1
    For iIndex = LBound(lPriority) To UBound(lPriority)
        If lPriority(iIndex) >= 0 Then
            oPlate.NamingCategory = lPriority(iIndex)
            Exit For
        End If
    Next iIndex

    Erase strLongNames
    Erase strShortNames
    Erase lPriority
            
    'Set Plate Type
    'Set Plate Tightness
    oPlate.PlateType = ePlateType
    oPlate.Tightness = eTightness
            
    'Set Plate Material, Grade and Thickness
    SetMatlGradeThickness oBearingPlate, strMatl, strGrade, dThickness
  
  Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*************************************************************************
'Function
'SetMatlGradeThickness
'
'Abstract
'   set Plate's Material Type, Grade, and Thickness
'
'input
'   oStructMaterial
'   strMatl
'   strGrade
'   dThickness
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub SetMatlGradeThickness(oStructMaterial As IJStructureMaterial, _
                                 strMatl As String, _
                                 strGrade As String, _
                                 dThickness As Double)
Const METHOD = "MbrAssemblyUtilities::InitializePlatePartProperties"
    Dim sMsg As String
  
    Dim iIndex As Long
    Dim nCount As Long
    
    Dim oPlate As IJPlate
    Dim oMatlObj As IJDMaterial
    Dim oPlateDims As IJDPlateDimensions
    Dim oRefDataQuery As RefDataMiddleServices.RefdataSOMMiddleServices
    Dim matlThickCol As IJDCollection
    Dim oMatProxy As Object
    Dim oPlateDimProxy As Object
    On Error GoTo ErrorHandler

    Dim oResMgr As IJDPOM

    Set oResMgr = GetResourceMgr
    
    
    Set oRefDataQuery = New RefDataMiddleServices.RefdataSOMMiddleServices
    Set oMatlObj = oRefDataQuery.GetMaterialByGrade(strMatl, strGrade)
    
    Set oMatProxy = oResMgr.GetProxy(oMatlObj)
    
    oStructMaterial.Material = oMatProxy
    
    If Not TypeOf oStructMaterial Is IJPlate Then
        Exit Sub
    End If
    
    Set oPlate = oStructMaterial
    Set matlThickCol = oRefDataQuery.GetPlateDimensions(oMatlObj.MaterialType, _
                                                        oMatlObj.MaterialGrade)
  
    nCount = matlThickCol.Size
    For iIndex = 1 To nCount
        Set oPlateDims = matlThickCol.Item(iIndex)
        If Abs(oPlateDims.thickness - dThickness) < 0.000005 Then
            Exit For
        End If
        Set oPlateDims = Nothing
    Next
    
    If oPlateDims Is Nothing Then
        Set oPlateDims = matlThickCol.Item(1)
    End If
    
    Set oPlateDimProxy = oResMgr.GetProxy(oPlateDims)
    oPlate.Dimensions = oPlateDimProxy
    Set oPlateDims = Nothing
    Set oPlateDimProxy = Nothing
    Set oRefDataQuery = Nothing
  
  Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Private Function GetResourceMgr() As IJDPOM

    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim pConnMiddle As IJDConnectMiddle
    Dim pAccessMiddle As IJDAccessMiddle
    
    Dim jContext As IJContext
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
 
    Set pConnMiddle = jContext.GetService("ConnectMiddle")
 
    Set pAccessMiddle = pConnMiddle
 
    Dim strModelDB As String
    strModelDB = oDBTypeConfig.get_DataBaseFromDBType("Model")
    Set GetResourceMgr = pAccessMiddle.GetResourceManager(strModelDB)
  
      
    Set jContext = Nothing
    Set oDBTypeConfig = Nothing
    Set pConnMiddle = Nothing
    Set pAccessMiddle = Nothing
End Function

