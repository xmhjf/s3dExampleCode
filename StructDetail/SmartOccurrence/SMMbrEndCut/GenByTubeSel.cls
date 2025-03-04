VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenByTubeSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011-12, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : GenByTubeSel.cls
'
'Description :
'   Selector for Selection of Member bounded by Generic Port/Object End Cuts
'   Expect Number of Inputs to be two(2) from the AppConnection.enumPorts
'       one Port/Object will be: (Bounding Object)
'           Plate Base/Offset/Lateral Face Port
'           Profile Base/Offset/ Lateral SubPort
'           MemberPart Base/Offset/ Lateral SubPort
'           Reference(Grid) Plane (IJPlane)
'           Point (IJPoint)
'       one Port will be SPSMemberAxisStart or SPSMemberAxisEnd (Bounded Member)
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    06/Sep/2011 - pnalugol
'           for supporting Generic ACs for Tube standard(Bounding)Vs non-tube Standard member(Bounded)
'    21/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    11/Jul/2012 - svsmylav
'            TR-206472(DM-216590):(i) For BUTube, leaf plate part need to be passed to 'GetSelForMbrBoundedToTube'
'     method (ii) Updated selector logic to handle border cases of 'To-Center'.
'*********************************************************************************************

Const m_sClassName As String = "GenByTubeSel"
Const m_FamilyProgid As String = ""
Const m_SelectorProgid As String = m_sProjectName + "." + m_sClassName
Const m_SelectorName As String = m_SelectorProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : SelectorInputs
' Description : List any graphic Inputs that the Selector has here
'
'*********************************************************************************************
Public Sub SelectorInputs(pIH As IJDInputsHelper)
    Const METHOD = m_SelectorProgid & "::SelectorInputs"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Defining Selector Inputs"

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED

    Exit Sub
ErrorHandler:
    pIH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorQuestions
' Description : List any Questions/Answers that the Selector uses here
'
'*********************************************************************************************
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
    Const METHOD = m_SelectorProgid & "::SelectorQuestions"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String

    ' TODO - Add your question below
    '===============================
    sMsg = "Defining/Initializing Selector Questions/Answers"
    'added questions as part of TR 202055- Generic ACs when Tube is Bounding
    pQH.SetQuestion "ConnectTube", "No", "BooleanCol", "CMSetConnectTubeAnswer", CUSTOMERID & "MbrEndCut.EndCutSelCM"
    pQH.SetQuestion "CreateTopWebCut", "No", "BooleanCol", "CMSetTopWebCutAnswer", CUSTOMERID & "MbrEndCut.GenWebPenetratedSel"
    pQH.SetQuestion "CreateBottomWebCut", "No", "BooleanCol", "CMSetBottomWebCutAnswer", CUSTOMERID & "MbrEndCut.GenWebPenetratedSel"
    pQH.SetQuestion "CreateTopFlangeCut", "No", "BooleanCol", "CMSetTopFlangeCutAnswer", CUSTOMERID & "MbrEndCut.GenWebPenetratedSel"
    pQH.SetQuestion "CreateBottomFlangeCut", "No", "BooleanCol", "CMSetBottomFlangeCutAnswer", CUSTOMERID & "MbrEndCut.GenWebPenetratedSel"

    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description : Select the appropriate standard/normal cut
'*********************************************************************************************
Public Sub SelectorLogic(oSL As IJDSelectorLogic)
    
    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    sMsg = "Selection logic"
    
    'For tubebytube cases
    Dim bIsBoxCut As Boolean
        
    'determine if the current cut is a part of box cut
    'get the bounding object and see if it has long box item and see if it's bounding
    'is same as bounded of current object.
    
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    
    If TypeOf oSL.InputObject(INPUT_BOUNDED) Is IJPort Then
        Dim oBoundedPort As IJPort
        Set oBoundedPort = oSL.InputObject(INPUT_BOUNDED)
        Set oBoundedObject = oBoundedPort.Connectable
    Else
        Set oBoundedObject = oSL.InputObject(INPUT_BOUNDED)
    End If
    
    Dim oSDO_Bounded As New StructDetailObjects.MemberPart
    Set oSDO_Bounded.object = oBoundedPort.Connectable
    
    If TypeOf oSL.InputObject(INPUT_BOUNDING) Is IJPort Then
        Dim oBoundingPort As IJPort
        Set oBoundingPort = oSL.InputObject(INPUT_BOUNDING)
        Set oBoundingObject = oBoundingPort.Connectable
    Else
        Set oBoundingObject = oSL.InputObject(INPUT_BOUNDING)
    End If
    
    Dim bIsBoundingTube As Boolean
    
    Dim selString As String
    bIsBoundingTube = IsTubularMember(oBoundingPort.Connectable)
    
    If Not bIsBoundingTube Then
        Exit Sub
    End If
    
    Dim dDistanceFromBounding1 As Double, dDistanceFromBounding2 As Double

    Dim bPenetratesWeb As Boolean
    bPenetratesWeb = IsWebPenetrated(oBoundingPort, oBoundedPort)
    
    Dim oBounding As Object, oBounded As Object
    Set oBounding = oBoundingObject
    Set oBounded = oBoundedObject
    
    Dim bBuiltup As Boolean
    Dim oBUMember As ISPSDesignedMember
    IsFromBuiltUpMember oBoundingObject, bBuiltup, oBUMember
        
    Dim iACFlag As Integer
    If bBuiltup Then
        Dim oSDO_WebCut As StructDetailObjects.WebCut
        Set oSDO_WebCut = New StructDetailObjects.WebCut
        Set oSDO_WebCut.object = oSL.SmartOccurrence
        Dim oBoundedLoc As IJDPosition
        Set oBoundedLoc = oSDO_WebCut.BoundedLocation
        
        Dim oPlateAttributes As GSCADCreateModifyUtilities.USSHelper
        Set oPlateAttributes = New GSCADCreateModifyUtilities.USSHelper
        Dim oLeafPlate As Object
        Dim bIsPlateSystem As Boolean
        Dim bIsSplitted As Boolean
        'Get Plate at given position
        oPlateAttributes.GetLeafPlateNearestToPoint oBoundingObject, oBoundedLoc, _
                                                    oLeafPlate, bIsSplitted
        If Not oLeafPlate Is Nothing Then
            Set oBoundingObject = oLeafPlate
        End If
    End If

    Dim eCase As eMemberBoundingCase
    eCase = GetBoundingCaseForTube(oBoundedObject, oBoundingObject, bPenetratesWeb, iACFlag)
        
    ' -----------------------------------------
    ' Make selection based on the bounding case
    ' -----------------------------------------
    Select Case eCase
    
        Case OSBottomAndOSTop, FCAndOSBottom, FCAndOSTop
        
            If bPenetratesWeb Then
            
                dDistanceFromBounding1 = GetDistanceFromBounding(oBoundingObject, oBoundedObject, JXSEC_TOP)
                dDistanceFromBounding2 = GetDistanceFromBounding(oBoundingObject, oBoundedObject, JXSEC_BOTTOM)
                
                Dim dTopClearance As Double
                Dim dBtmClearance As Double
                Dim dBoundedFlangeThickness As Double
                dBoundedFlangeThickness = oSDO_Bounded.FlangeThickness
                dTopClearance = dDistanceFromBounding1 - dBoundedFlangeThickness
                dBtmClearance = dDistanceFromBounding2 - dBoundedFlangeThickness
            
                 If eCase <> FCAndOSBottom And eCase <> FCAndOSTop Then
                    If dTopClearance > 0.005 Or _
                        dBtmClearance > 0.005 Then '5mm
                        oSL.Add "Generic_Tube_Web"
                    Else
                        oSL.Add "Generic_Tube_Web_Ext" 'Straight
                    End If
                Else
                    If Sgn(dTopClearance) = 1 And dTopClearance < 0.005 Then
                        oSL.Add "Generic_Tube_Web_FCAndOS_TOP"
                    ElseIf Sgn(dBtmClearance) = 1 And dBtmClearance < 0.005 Then
                        oSL.Add "Generic_Tube_Web_FCAndOS_BTM"
                    Else
                        oSL.Add "Generic_Tube_Web"
                    End If
                End If
                
            Else
            
                'Check distance from left and right side of web
                dDistanceFromBounding1 = GetDistanceFromBounding(oBoundingObject, oBoundedObject, JXSEC_WEB_LEFT)
                'Check distance below the bounding and set bottom nose value
                dDistanceFromBounding2 = GetDistanceFromBounding(oBoundingObject, oBoundedObject, JXSEC_WEB_RIGHT)
                
                If dDistanceFromBounding1 < 0.0001 And dDistanceFromBounding2 < 0.0001 Then
                    oSL.Add "Generic_Tube_Web"
                Else
                    oSL.Add "Generic_Tube_Web_Ext" 'Straight
                End If
            End If
            
        Case Else
            'ToCenter Case
            'Identify border Web-Penetrated cases and choose appropriate smart item
            Select Case iACFlag
             Case BorderAC_OSOS
                oSL.Add "Generic_Tube_Web_OSAndOS"
             Case BorderAC_FCOS_TOP
                oSL.Add "Generic_Tube_Web_FCAndOS_TOP"
             Case BorderAC_FCOS_BTM
                oSL.Add "Generic_Tube_Web_FCAndOS_BTM"
             Case BorderAC_ToCenter, 0
                oSL.Add "Generic_Tube_Web_ToCenter"
             Case Else
                'Need to handle flange penetrated ToCenter case
                oSL.Add "Generic_Tube_Web"
            End Select
            oSL.Add "Generic_Tube_Web_Ext" 'Straight
        
    End Select


    Exit Sub
ErrorHandler:
    oSL.ReportError sMsg, METHOD
    
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
' Remove all existing defined Input and Output (Representations)
' before defining the current Inputs and Outputs
    pSelector.IJDInputs.RemoveAllInput
    pSelector.IJDRepresentations.RemoveAllRepresentation

    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pSelector
    pIH.InitAs m_FamilyProgid
    SelectorInputs pIH

    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function

'*********************************************************************************************
' Method      : CMSelector
' Description :
'*********************************************************************************************
Public Sub CMSelector(pRep As IJDRepresentation)
    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

