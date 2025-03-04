VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StiffEndToLongBoxDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMAssyConRul
'  File        : StiffEndToLongBoxDef.cls
'
'  Description :
'
'   Item definition: basic template to define an item
'
'  Author      : Alligators
'
'  History     :
'    12/Jan/2012 - svsmylav
'           CR-174918: Added code in CMConstruct_WebCut to use plates subport
'           when the EndCut is for long profile and pseudo-knuckle.
'    16/Feb/2012 - svsmylav
'           CR-174918: 'CMConstruct_FlangeCut' and 'Set_CMInputs'  methods are updated:
'           for LongBox case, "FlangeCuts" SmartClass is used for SplitAndExtend Mfg. method
'           option and Convex knuckle. Also, "KnuckleEndCut" question is NOT copied to child
'           selector for above case.
'    17/04/2012 - GH -CR-174918
'           Replaced 'GetPlateSubPortForSplitAndExtend' method with 'GetStiffnedPlatePort_At_Position' in CMConstruct_WebCut
'*********************************************************************************************

Const m_ItemProgid As String = CUSTOMERID + "AssyConRul.StiffEndToLongBoxDef"
Const m_ItemName As String = CUSTOMERID + "AssyConRul.StiffEndToLongBoxDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\StiffEndToLongBoxDef.cls"

Implements IJDUserSymbolServices
'

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    Dim strError As String

    strError = "Setting bounded and bounding objects as inputs."
    pIH.SetInput INPUT_BOUNDED_OR_PENETRATED_OBJECT
    pIH.SetInput INPUT_BOUNDING_OR_PENETRATING_OBJECT
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs", strError).Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    ' TODO - Add aggregator description below
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
zMsgBox "StiffEndToLongBoxDef::ItemMembers"
    
    Dim strError As String
    Dim pMemDesc As IJDMemberDescription
       
    '--------------------------
    strError = "Constructing first Web Cut."
    Set pMemDesc = pMDs.AddMember("StiffEndToLongWebCut1", 1, _
                                  "CMConstruct_WebCut1", _
                                  imsCOOKIE_ID_USS_LIB)

    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMInputsToWebCut1"
    pMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMTest_WebCut1"
    
    '--------------------------
    strError = "Constructing Top Flange Cut for first Web Cut."
    Set pMemDesc = Nothing
    Set pMemDesc = pMDs.AddMember("StiffEndToLongFlangeCut11", 2, _
                                  "CMConstruct_FlangeCut11", _
                                  imsCOOKIE_ID_USS_LIB)

    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMInputsFlangeCut11"
    pMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMTest_FlangeCut11"
    
    '--------------------------
    strError = "Constructing Bottom Flange Cut for first Web Cut."
    Set pMemDesc = Nothing
    Set pMemDesc = pMDs.AddMember("StiffEndToLongFlangeCut12", 3, _
                                  "CMConstruct_FlangeCut12", _
                                  imsCOOKIE_ID_USS_LIB)

    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMInputsFlangeCut12"
    pMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMTest_FlangeCut12"
    
    '--------------------------
    '--------------------------
    '--------------------------
    strError = "Constructing second Web Cut."
    Set pMemDesc = pMDs.AddMember("StiffEndToShortWebCut2", 4, _
                                  "CMConstruct_WebCut2", _
                                  imsCOOKIE_ID_USS_LIB)
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMInputsToWebCut2"
    pMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMTest_WebCut2"
    
    '--------------------------
    strError = "Constructing Top Flange Cut for second Web Cut."
    Set pMemDesc = Nothing
    Set pMemDesc = pMDs.AddMember("StiffEndToShortFlangeCut21", 5, _
                                  "CMConstruct_FlangeCut21", _
                                  imsCOOKIE_ID_USS_LIB)
                                  
    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMInputsFlangeCut21"
    pMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMTest_FlangeCut21"

    '--------------------------
    strError = "Constructing Bottom Flange Cut for second Web Cut."
    Set pMemDesc = Nothing
    Set pMemDesc = pMDs.AddMember("StiffEndToShortFlangeCut22", 6, _
                                  "CMConstruct_FlangeCut22", _
                                  imsCOOKIE_ID_USS_LIB)

    pMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMInputsFlangeCut22"
    pMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMTest_FlangeCut22"
    
    Set pMemDesc = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers", strError).Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Methods HERE
' *******************************************************************************************

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub CMInputsToWebCut1(pMemberDescription As IJDMemberDescription)
Const sMETHOD = "CMInputsToWebCut1"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    sError = "Copying questions for first Miter WebCut"
    Set_CMInputs pMemberDescription, "", True

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

Public Sub CMInputsFlangeCut11(pMemberDescription As IJDMemberDescription)
Const sMETHOD = "CMInputsFlangeCut11"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    sError = "Copying questions for first Miter Top FlangeCut"
    Set_CMInputs pMemberDescription, "No", True

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

Public Sub CMInputsFlangeCut12(pMemberDescription As IJDMemberDescription)
Const sMETHOD = "CMInputsFlangeCut12"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    sError = "Copying questions for first Miter Bottom FlangeCut"
    Set_CMInputs pMemberDescription, "Yes", True

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

Public Sub CMInputsToWebCut2(pMemberDescription As IJDMemberDescription)
Const sMETHOD = "CMInputsToWebCut2"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    sError = "Copying questions for second Miter WebCut"
    Set_CMInputs pMemberDescription, "", False

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

Public Sub CMInputsFlangeCut21(pMemberDescription As IJDMemberDescription)
Const sMETHOD = "CMInputsFlangeCut21"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    sError = "Copying questions for second Miter Top FlangeCut"
    Set_CMInputs pMemberDescription, "No", False

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

Public Sub CMInputsFlangeCut22(pMemberDescription As IJDMemberDescription)
Const sMETHOD = "CMInputsFlangeCut22"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    sError = "Copying questions for second Miter Bottom FlangeCut"
    Set_CMInputs pMemberDescription, "Yes", False

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub CMTest_FlangeCut11(ByRef pMD As IJDMemberDescription, _
                              ByRef bIsNeeded As Boolean)
Const sMETHOD = "CMTest_FlangeCut11"
    On Error GoTo ErrorHandler
    
    Dim sSectionType As String
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = False
    sSectionType = Get_CrossSectionType(pMD)
        
    bIsNeeded = True
    If sSectionType = "FB" Or _
       sSectionType = "HalfR" Or _
       sSectionType = "P" Or _
       sSectionType = "R" Or _
       sSectionType = "SB" Or _
       sSectionType = "SqTu" Or _
       sSectionType = "RT" Then
            bIsNeeded = False
    End If
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

Public Sub CMTest_FlangeCut12(ByRef pMD As IJDMemberDescription, _
                              ByRef bIsNeeded As Boolean)
Const sMETHOD = "CMTest_FlangeCut12"
    On Error GoTo ErrorHandler
    
    Dim sSectionType As String
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = False
    sSectionType = Get_CrossSectionType(pMD)
    
    If sSectionType = "I" Or _
       sSectionType = "ISType" Or _
       sSectionType = "H" Or _
       sSectionType = "C_SS" Or _
       sSectionType = "CSType" Then
            bIsNeeded = True
        End If
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

Public Sub CMTest_FlangeCut21(ByRef pMD As IJDMemberDescription, _
                              ByRef bIsNeeded As Boolean)
Const sMETHOD = "CMTest_FlangeCut21"
    On Error GoTo ErrorHandler
    
    Dim sSectionType As String
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    bIsNeeded = False
    sSectionType = Get_CrossSectionType(pMD, False)
        
    bIsNeeded = True
    If sSectionType = "FB" Or _
       sSectionType = "HalfR" Or _
       sSectionType = "P" Or _
       sSectionType = "R" Or _
       sSectionType = "SB" Or _
       sSectionType = "SqTu" Or _
       sSectionType = "RT" Then
            bIsNeeded = False
    End If
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

Public Sub CMTest_FlangeCut22(ByRef pMD As IJDMemberDescription, _
                              ByRef bIsNeeded As Boolean)
Const sMETHOD = "CMTest_FlangeCut22"
    On Error GoTo ErrorHandler
    
    Dim sSectionType As String
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = False
    sSectionType = Get_CrossSectionType(pMD, False)
    
    If sSectionType = "I" Or _
       sSectionType = "ISType" Or _
       sSectionType = "H" Or _
       sSectionType = "C_SS" Or _
       sSectionType = "CSType" Then
            bIsNeeded = True
        End If
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Function CMConstruct_WebCut1(ByVal pMemberDescription As IJDMemberDescription, _
                                    ByVal pResourceManager As IUnknown, _
                                    ByRef pObject As Object)
Const sMETHOD = "CMConstruct_WebCut1"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    sError = "Creation of first WebCut"
    CMConstruct_WebCut pMemberDescription, pResourceManager, pObject, True
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function CMConstruct_FlangeCut11(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObject As Object)
Const sMETHOD = "CMConstruct_FlangeCut11"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    sError = "Creation of first Top FlangeCut"
    CMConstruct_FlangeCut pMemberDescription, pResourceManager, pObject, True
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function CMConstruct_FlangeCut12(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObject As Object)
Const sMETHOD = "CMConstruct_FlangeCut12"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    sError = "Creation of first Bottom FlangeCut"
    CMConstruct_FlangeCut pMemberDescription, pResourceManager, pObject, True
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function CMConstruct_WebCut2(ByVal pMemberDescription As IJDMemberDescription, _
                                    ByVal pResourceManager As IUnknown, _
                                    ByRef pObject As Object)
Const sMETHOD = "CMConstruct_WebCut2"
    On Error GoTo ErrorHandler
    Dim sError As String

    sError = "Creation of second WebCut"
    CMConstruct_WebCut pMemberDescription, pResourceManager, pObject, False
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function CMConstruct_FlangeCut21(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObject As Object)
Const sMETHOD = "CMConstruct_FlangeCut21"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    sError = "Creation of second Top FlangeCut"
    CMConstruct_FlangeCut pMemberDescription, pResourceManager, pObject, False
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function CMConstruct_FlangeCut22(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObject As Object)
Const sMETHOD = "CMConstruct_FlangeCut22"
    On Error GoTo ErrorHandler
    Dim sError As String
    
    sError = "Creation of second Bottom FlangeCut"
    CMConstruct_FlangeCut pMemberDescription, pResourceManager, pObject, False
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

'=============================================================================
'=============================================================================
'=============================================================================
' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_ItemName
End Function

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
    
    Dim pPDs As IJDPropertyDescriptions
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Set pAD = pDefinition
    Set pPDs = pAD
    pPDs.RemoveAll ' Remove all the previous property descriptions
    ItemAggregator pAD
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    pMDs.RemoveAll ' Remove all the previous Member descriptions
    ItemMembers pMDs
End Sub

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
    Dim pDefinition As IJDSymbolDefinition
    Dim pCAFactory As New CAFactory
    
    Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
    
    ' Set definition progId and codebase
    pDefinition.ProgId = m_ItemProgid
    pDefinition.CodeBase = CodeBase
    pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
      
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
End Function

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Private Sub Set_CMInputs(pMemberDescription As IJDMemberDescription, _
                         sBottomFlange As String, _
                         bLongBox As Boolean)

Const sMETHOD = "Set_CMInputs"
    
    On Error GoTo ErrorHandler
    
    Dim sError As String
    Dim sToSelector As String
    Dim sSetWeldPart As String
    Dim sUseWeldPart As String
    Dim pCopyAnswerHelper As CopyAnswerHelper
    Dim oProfileKnuckle As Object
    Dim iKnuckleMfgMethod As Integer
    Dim bIsConvex As Boolean
    Dim bCnvxSplitAndExtend As Boolean

    bIsConvex = False
    bCnvxSplitAndExtend = False
    
    GetProfileKnuckleType pMemberDescription.CAO, iKnuckleMfgMethod, oProfileKnuckle, bIsConvex
    If iKnuckleMfgMethod = pkmmSplitAndExtend And bIsConvex Then bCnvxSplitAndExtend = True
        
    If bLongBox Then
        If Len(Trim(sBottomFlange)) > 0 Then
            If bCnvxSplitAndExtend = False Then
                sToSelector = CUSTOMERID + "EndCutRules.EndToLongBoxFlangeCutSel"
            Else
                ''FlangeCut_LongBoxW1' did not resymbolize, so
                ' to make flange cut work regular FlangeCut Selector is chosen
                ' as a temporary fix
                sToSelector = CUSTOMERID + "EndCutRules.FlangeCutSel"
            End If
        Else
            sToSelector = CUSTOMERID + "EndCutRules.EndToLongBoxWebCutSel"
        End If
    
    Else
        If Len(Trim(sBottomFlange)) > 0 Then
            sToSelector = CUSTOMERID + "EndCutRules.FlangeCutSel"
        Else
            sToSelector = CUSTOMERID + "EndCutRules.WebCutSel"
        End If
    End If
        
    ' set Answers used by the Web Cut and Flange Cut Selectors
    Set pCopyAnswerHelper = New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    
    sError = "setting EndCut Type Question/Answer"
    If bLongBox Then
        pCopyAnswerHelper.CopyAnswer sToSelector, _
                                     QUES_ENDCUTTYPE, _
                                     CUSTOMERID + "AssyConRul.StiffEndByKnuckleSel", _
                                     QUES_ENDCUTTYPE
                                     
    
        sError = "setting KnuckleEndCut Type Question/Answer"
        If bCnvxSplitAndExtend = False Then
            pCopyAnswerHelper.CopyAnswer sToSelector, _
                                    "KnuckleEndCut", _
                                     CUSTOMERID + "AssyConRul.StiffEndByKnuckleSel", _
                                     "KnuckleEndCut"
        End If
    Else
        ' the Short Boxed End is always a Weld (W) type
        pCopyAnswerHelper.PutAnswer sToSelector, QUES_ENDCUTTYPE, gsW
    End If
    
    If Len(Trim(sBottomFlange)) > 0 Then
        sError = "setting TheBottomFlange Question/Answer"
        pCopyAnswerHelper.PutAnswer sToSelector, _
                                    "TheBottomFlange", _
                                    sBottomFlange
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function CMConstruct_WebCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object, _
                                   bLongBox As Boolean)
Const sMETHOD = "CMConstruct_WebCut"
    On Error GoTo ErrorHandler
    Dim sError As String
    Dim sWebCutClass As String

    ' Get wrapper class
    Dim oBounded As Object
    Dim oBounding As Object
    Dim oLateralPort As Object
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
    Dim oSDO_AssemblyConn As StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    sError = "Setting assembly connection inputs."
    Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
    Set oSDO_AssemblyConn = New StructDetailObjects.AssemblyConn
    Set oSDO_AssemblyConn.object = pMemberDescription.CAO
    
    sError = "Getting assembly connection ports."
    If bLongBox Then
        sWebCutClass = "WebCutsLongBox"
        Set oBounded = oSDO_AssemblyConn.Port1
        Set oBounding = oSDO_AssemblyConn.Port2
'''$$$
'''Debug/Test Only
zTest_KnuckleBoundPorts oBounded, oBounding, oSDO_AssemblyConn.BoundGlobalShipLocation
'''$$$
        Dim oProfileKnuckle As Object
        Dim iKnuckleMfgMethod As Integer
        Dim bIsConvex As Boolean
        
        GetProfileKnuckleType oSDO_AssemblyConn.object, iKnuckleMfgMethod, oProfileKnuckle, bIsConvex

        'Profile Knuckle with "SplitAndExtend" as Mfg Method and is Convex: BoundingPort is PlateSubPort
        If bIsConvex And iKnuckleMfgMethod = pkmmSplitAndExtend Then
            Set oSDO_ProfilePart.object = oSDO_AssemblyConn.ConnectedObject1
            'Use Short Profile's Plate Sub Port as Bounding Port for WebCut
            Set oLateralPort = GetStiffnedPlatePort_At_Position(oSDO_AssemblyConn.ConnectedObject1, _
                        oSDO_AssemblyConn.BoundGlobalShipLocation)
        Else
            '
            Set oSDO_ProfilePart.object = oSDO_AssemblyConn.ConnectedObject2
            Set oLateralPort = oSDO_ProfilePart.BasePortBeforeTrim(BPT_Lateral)
        End If
    Else
        sWebCutClass = "WebCuts"
        Set oBounded = oSDO_AssemblyConn.Port2
        Set oBounding = oSDO_AssemblyConn.Port1
'''$$$
'''Debug/Test Only
zTest_KnuckleBoundPorts oBounded, oBounding, oSDO_AssemblyConn.BoundGlobalShipLocation
'''$$$
        Set oSDO_ProfilePart.object = oSDO_AssemblyConn.ConnectedObject1
        Set oLateralPort = oSDO_ProfilePart.BasePortBeforeTrim(BPT_Lateral)
    End If
    
    ' Create webcut
    Set oSDO_WebCut = New WebCut
    sError = "Creating web cut."
    oSDO_WebCut.Create pResourceManager, oLateralPort, oBounded, _
                       sWebCutClass, oSDO_AssemblyConn.object
                               
    sError = "Return the webcut"
    Set pObject = oSDO_WebCut.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function CMConstruct_FlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                      ByVal pResourceManager As IUnknown, _
                                      ByRef pObject As Object, _
                                      bLongBox As Boolean)
Const sMETHOD = "CMConstruct_FlangeCut"
    On Error GoTo ErrorHandler
    Dim sError As String
    Dim sFlangeCutClass As String
    
    Dim iCount As Long
    Dim iIndex As Long
    Dim bFound As Boolean
    
    Dim oBounded As Object
    Dim oBounding As Object
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim oSDO_FlangeCut As StructDetailObjects.FlangeCut
    
    Dim oObject As Object
    Dim oWebCut As IJStructFeature
    Dim oMemberObjects As IJDMemberObjects
    Dim oSelectorLogic As IJDSelectorLogic
    
    Dim oProfileKnuckle As Object
    Dim iKnuckleMfgMethod As Integer
    Dim bIsConvex As Boolean
    
    bIsConvex = False
    
    GetProfileKnuckleType pMemberDescription.CAO, iKnuckleMfgMethod, oProfileKnuckle, bIsConvex
    'Profile Knuckle having "SplitAndExtend" as Mfg Method and is on Convex: BoundingPort is PlateSubPort
    If bLongBox Then
        If iKnuckleMfgMethod = pkmmSplitAndExtend And bIsConvex Then
            sFlangeCutClass = "FlangeCuts"
        Else
            sFlangeCutClass = "FlangeCutsLongBox"
        End If
    Else
        sFlangeCutClass = "FlangeCuts"
    End If
    
    ' need the Web Cut as input to the Flange Cut
    ' Loop thru current member Items searching for Web Cut
    ' For bReversePorts = True,
    '   creating second set of Flange Cuts
    '   want to use the second Web Cut (first from bottom of list)
    ' For bReversePorts = False,
    '   creating first set of Flange Cuts
    '   want to use the first Web Cut (first from top of list)
    
    Set oMemberObjects = pMemberDescription.CAO
    iCount = oMemberObjects.Count
    bFound = False
    
    For iIndex = 1 To iCount
        If bLongBox Then
            Set oObject = oMemberObjects.Item(iIndex)
        Else
            Set oObject = oMemberObjects.Item(iCount - iIndex + 1)
        End If
        
        If Not oObject Is Nothing Then
            If TypeOf oObject Is IJStructFeature Then
                Set oWebCut = oObject
                If (oWebCut.get_StructFeatureType = SF_WebCut) Then
                    bFound = True
                    Exit For
                Else
                    Set oWebCut = Nothing
                End If
            End If
           
           Set oObject = Nothing
        End If
    Next iIndex
    
    If bFound Then
        sError = "Retrieving Bounded/Bounding Ports from WebCut."
        Set oSDO_WebCut = New StructDetailObjects.WebCut
        Set oSDO_WebCut.object = oWebCut
        Set oBounded = oSDO_WebCut.BoundedPort
        Set oBounding = oSDO_WebCut.BoundingPort
        
        ' Create Miter Flange Cut
        sError = "Creating Flange Cut."
        Set oSDO_FlangeCut = New StructDetailObjects.FlangeCut
        Call oSDO_FlangeCut.Create(pResourceManager, _
                                   oBounding, oBounded, _
                                   oWebCut, _
                                   sFlangeCutClass, _
                                   pMemberDescription.CAO)
                               
         sError = "Return the flange cut "
         Set pObject = oSDO_FlangeCut.object
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Sub CMTest_WebCut1(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    bIsNeeded = True
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_WebCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMTest_WebCut1").Number
End Sub
Public Sub CMTest_WebCut2(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    bIsNeeded = True
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_WebCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMTest_WebCut2").Number
End Sub