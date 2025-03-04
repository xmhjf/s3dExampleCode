VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrAxisFlangeDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011-12, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrEndCut
'  File        : MbrAxisFlangeDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    22/Aug/2011 - svsmylav
'           B-25298(CR-179926): added logic to determine if PC is needed for outside-to-outside
'           configuration, for standard member bounded to tube case.
'    21/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    30/Nov/2011 - svsmylav TR-205302: Replaced '.Subport' method call with 'GetLateralSubPortBeforeTrim'.
'    07/Feb/2012 - svsmylav
'           TR-CP-209564: Updated 'CM_IsPhysConnsNeeded' method to return 'True'
'           when sniped-welded new FlangeCut smart items are selected.
'    11/Jul/2012 - svsmylav
'            TR-206472(DM-216590): Updated selector logic to handle border cases of 'To-Center'.
'*********************************************************************************************
Const m_sClassName As String = "MbrAxisFlangeDef"
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
    pIH.ReportError sMsg, METHOD
 
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
    
    Dim pPDs As IJDPropertyDescriptions
    Set pPDs = pAD
    pPDs.RemoveAll
    
    ' Set the FlangeCutting Behaviour
    sMsg = "Setting Flange Cutting Behaviour property."
    pPDs.AddProperty "CuttingBehavior", 1, "IJUAMbrEndCuts", "CM_SetFlangeCuttingBehaviour", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
    
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
    
    pMDs.RemoveAll  ' Removes all the previous cached member descriptions(best practice)
    
    Dim iDispId As Long
    Dim oMemDesc As IJDMemberDescription
    
    iDispId = 1
    sDispId = "AxisFlange_PC_" & Trim(Str(iDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructPhysConns", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPhysConnsNeeded"
    Set oMemDesc = Nothing
    
    'deleted methods to create corner feature that are deferred to flangecut by webcut as part of 205624
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Methods HERE
' *******************************************************************************************
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
    
    bIsNeeded = False 'Intial value, set it to True only for valid cases
    
    'Check if bounding is tube member and set conditional
    Dim oParentObj As Object
    Dim sParentItemName As String
    Dim strAnswer As String
    Parent_SmartItemName pMemberDescription.CAO, sParentItemName, oParentObj
        
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    Dim selString As String
    
    Set oAppConnection = oParentObj

    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    

    'For TopFlangeRight exists in all aliases, this PC is independent of bounding alias
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Set oBoundingPort = oBoundingData.AxisPort
    Set oBoundedPort = oBoundedData.AxisPort
    
    Dim bPenetratesWeb As Boolean
    
    If Not oBoundedPort Is Nothing And Not oBoundingPort Is Nothing Then
        bPenetratesWeb = IsWebPenetrated(oBoundingPort, oBoundedPort)
    Else
        Exit Sub
    End If
    
    Dim sBottomFlange As String
    GetSelectorAnswer pMemberDescription.CAO, "BottomFlange", sBottomFlange

    If IsTubularMember(oBoundingData.MemberPart) Then

        If StrComp(sParentItemName, gsMbrAxisToOutSideAndOutSideNoEdge, vbTextCompare) = 0 Or _
           StrComp(sParentItemName, gsStiffEndToMbrOutSideAndOutSideNoEdge, vbTextCompare) = 0 Then
            If bPenetratesWeb Then
                'Physical conection is not needed
                bIsNeeded = False
            Else
                'Need to add further checks to find flange-to-bounding overlap exists
                Dim eCode As JXSEC_CODE
                If StrComp(sBottomFlange, gsNo, vbTextCompare) = 0 Then
                    eCode = JXSEC_TOP
                Else
                    eCode = JXSEC_BOTTOM
                End If
                If GetDistanceFromBounding(oBoundingPort.Connectable, oBoundedPort.Connectable, eCode) < 0.00001 Then
                    'distance is 0.01 mm
                    bIsNeeded = True
                Else
                    bIsNeeded = False
                End If
            End If
            Exit Sub
        Else
            'Need to check shape-at-face answer first
            Dim strShapeAtFaceAnswer As String
            If StrComp(sBottomFlange, gsNo, vbTextCompare) = 0 Then
                'Check for Bounded Top
                GetSelectorAnswer oParentObj, "TopShapeAtFace", strShapeAtFaceAnswer
            Else
                'Check for Bounded Bottom
                GetSelectorAnswer oParentObj, "BottomShapeAtFace", strShapeAtFaceAnswer
            End If
            

            If StrComp(strShapeAtFaceAnswer, gsCope, vbTextCompare) = 0 Or _
                StrComp(strShapeAtFaceAnswer, gsSnipe, vbTextCompare) = 0 Then
                'Physical conection is not needed, so exit
                Exit Sub '*** Exit sub
            End If
            

            'Bounded-to-center case with remaining shape-at-face
            ' answers needs physical connection: but check for border cases
            If StrComp(sParentItemName, gsMbrAxisToCenter, vbTextCompare) = 0 Then
                Dim sSelector As String
                Dim iACFlag As Integer
                
                GetBoundingCaseForTube oBoundedData.MemberPart, oBoundingData.MemberPart, bPenetratesWeb, iACFlag
                
                If iACFlag = BorderAC_ToCenter Or iACFlag = 0 Then
                    bIsNeeded = True
                ElseIf iACFlag = BorderAC_FCOS_TOP And StrComp(sBottomFlange, gsYes, vbTextCompare) = 0 Then
                    bIsNeeded = True
                ElseIf iACFlag = BorderAC_FCOS_BTM And StrComp(sBottomFlange, gsNo, vbTextCompare) = 0 Then
                    bIsNeeded = True
                End If
                Exit Sub '*** Exit sub
            End If
            
            'Below geometry check is needed for face-and-outside case
            
            'Find bounding outer port
            Dim oBoundingOuterPort As IJPort
            Dim oSDO_Bounding As StructDetailObjects.MemberPart
            Set oSDO_Bounding = New StructDetailObjects.MemberPart
            Set oSDO_Bounding.object = oBoundingData.MemberPart
            Set oBoundingOuterPort = GetLateralSubPortBeforeTrim(oSDO_Bounding.object, JXSEC_OUTER_TUBE)
            
            Dim oBounded As Object
            Dim oPort As IJPort
            Dim oBoundedMbr As StructDetailObjects.MemberPart
            Dim oBoundedProfile As StructDetailObjects.ProfilePart
            
            
            Set oBounded = oBoundedData.MemberPart
            
            If TypeOf oBounded Is ISPSMemberPartCommon Then
                Set oBoundedMbr = New MemberPart
                Set oBoundedMbr.object = oBounded
                If StrComp(sBottomFlange, gsNo, vbTextCompare) = 0 Then
                    'Check for Bounded Top
                    Set oPort = oBoundedMbr.SubPort(JXSEC_TOP)
                Else
                    'Check for Bounded Bottom
                    Set oPort = oBoundedMbr.SubPort(JXSEC_BOTTOM)
                End If
            ElseIf TypeOf oBounded Is IJProfile Then
                Set oBoundedProfile = New ProfilePart
                Set oBoundedProfile.object = oBounded
                If StrComp(sBottomFlange, gsNo, vbTextCompare) = 0 Then
                    'Check for Bounded Top
                    Set oPort = oBoundedProfile.SubPort(JXSEC_TOP)
                Else
                    'Check for Bounded Bottom
                    Set oPort = oBoundedProfile.SubPort(JXSEC_BOTTOM)
                End If
            Else
            
            End If
                
            If StrComp(sBottomFlange, gsNo, vbTextCompare) = 0 Then
                'Check for Bounded Top
                Set oPort = GetLateralSubPortBeforeTrim(oBounded, JXSEC_TOP)
            Else
                'Check for Bounded Bottom
                Set oPort = GetLateralSubPortBeforeTrim(oBounded, JXSEC_BOTTOM)
            End If
            Dim oPlane As IJPlane
            If TypeOf oPort Is IJPlane Then
                Set oPlane = GetInfPlaneFromPort(oPort)
            End If
            Dim oIJSurfBdg As IJSurface
            Set oIJSurfBdg = oBoundingOuterPort

            Dim oElems As IJElements
            Dim gicCode As Geom3dIntersectConstants
            oIJSurfBdg.Intersect oPlane, oElems, gicCode
            If gicCode = ISECT_NOSOLUTION Then
                'Physical conection is not needed, so exit
                Exit Sub '*** Exit sub
            Else
                'Physical conection is needed
                bIsNeeded = True
                Exit Sub '*** Exit sub
            End If
        End If
    End If
    
    Dim oEndCutObject As Object
    Dim oSmartItem As IJSmartItem
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim sSelection As String

    sSelection = ""
    Set oEndCutObject = pMemberDescription.CAO

    If TypeOf oEndCutObject Is IJSmartOccurrence Then
        Set oSmartOccurrence = oEndCutObject
        Set oSmartItem = oSmartOccurrence.SmartItemObject
        sSelection = oSmartItem.Name
    End If

    If sSelection = "Axis_Flange_Flush_LR" Or _
        sSelection = "Axis_Flange_Variable_Bevel" Or _
        sSelection = "Axis_Flange_Flush_R" Or _
        sSelection = "Axis_Flange_Sniped_LR" Or _
        sSelection = "Axis_Flange_Sniped_R" Then
        bIsNeeded = True
    End If
   
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_ConstructPhysConns
' Description :
'
'*********************************************************************************************
Public Function CM_ConstructPhysConns(ByVal pMemberDescription As IJDMemberDescription, _
                                      ByVal pResourceManager As IUnknown, _
                                      ByRef pObject As Object)
    Const METHOD = m_DefinitionProgid & "::CM_ConstructPhysConns"
    On Error GoTo ErrorHandler

    Dim sMsg As String
   
    ' Get Bounding Part's Port
    Dim eBoundingPort As eUSER_CTX_FLAGS
    Dim eBoundedSubPort As JXSEC_CODE
    
    Dim oSDO_FlangeCut As New StructDetailObjects.FlangeCut
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_FlangeCut.object = pMemberDescription.CAO
    Set oSDO_WebCut.object = oSDO_FlangeCut.WebCut
        
    'TR 206472 - Change bounding port to use Base port of the bounding plate part when the Bounding is a built-up
    
    Dim oBUMember As ISPSDesignedMember
    Dim bBuiltup As Boolean
    
    IsFromBuiltUpMember oSDO_FlangeCut.Bounding, bBuiltup, oBUMember
    If bBuiltup Then
        eBoundingPort = CTX_BASE
    Else
        eBoundingPort = CTX_LATERAL
    End If
    eBoundedSubPort = 8193
    
    If pResourceManager Is Nothing Then
        sMsg = "Modifying Physical Connection " & _
               "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
        Set pObject = Modify_PhysConn(pMemberDescription, eBoundingPort, eBoundedSubPort)
        
    Else
        sMsg = "Creating Physical Connection " & _
               "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
                         
        ' --------------------------------------------
        ' Find out if this is the top or bottom flange
        ' --------------------------------------------
    
        Dim sBottomFlange As String
        Dim isBottom As Boolean
        isBottom = False
    
        GetSelectorAnswer pMemberDescription, "BottomFlange", sBottomFlange
        If sBottomFlange = "Yes" Then
            isBottom = True
        End If
    
        ' ----------------------------------------------------------------------
        ' If the bounding object is a member, find out which edge is intersected
        ' ----------------------------------------------------------------------
       
        Dim boundingEdgeId As eBounding_Edge
        Dim mappedEdgeId As JXSEC_CODE
    
        boundingEdgeId = None
    
        If TypeOf oSDO_FlangeCut.Bounding Is ISPSMemberPartPrismatic And Not IsTubularMember(oSDO_WebCut.Bounding) Then
            GetNonPenetratedIntersectedEdge oSDO_WebCut.object, oSDO_WebCut.BoundingPort, oSDO_WebCut.BoundedPort, boundingEdgeId, mappedEdgeId, isBottom
        End If
    
        ' -----------------------------------------------------------
        ' Create a Tee or Butt weld, depending on bounding edge width
        ' -----------------------------------------------------------
        If boundingEdgeId = eBounding_Edge.Bottom_Flange_Right Or boundingEdgeId = eBounding_Edge.Top_Flange_Right Then
            Set pObject = Create_PhysConn(pMemberDescription, pResourceManager, _
                                      "ButtWeld", eBoundingPort, eBoundedSubPort)
        ElseIf Not (boundingEdgeId = eBounding_Edge.Above Or boundingEdgeId = eBounding_Edge.Below Or boundingEdgeId = eBounding_Edge.None) Then
            Set pObject = Create_PhysConn(pMemberDescription, pResourceManager, _
                                      "TeeWeld", eBoundingPort, eBoundedSubPort)
        Else
            Set pObject = Create_PhysConn(pMemberDescription, pResourceManager, _
                                      "TeeWeld", eBoundingPort, eBoundedSubPort)
        End If

    End If
                                      
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Function
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
    ' Name should be unique
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
Public Function IJDUserSymbolServices_EditOccurence(pSymbolOccurence As Object, _
                                                    ByVal transactionMgr As Object) As Boolean
End Function

'******************************************************************************************
' Method:
' CM_SetFlangeCuttingBehaviour
'
' *******************************************************************************************
Public Sub CM_SetFlangeCuttingBehaviour(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CM_SetFlangeCuttingBehaviour"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Updating a Flange Cutting Behaviour"
            
    Dim oMemberDescriptions As IJDMemberDescriptions
    Set oMemberDescriptions = pPropertyDescriptions.definition
    
    Dim oMemberDescription As IJDMemberDescription
    Dim bIsNeeded As Boolean
    bIsNeeded = True
    
    Dim IsProExists As Boolean
    
    'Check if the attribute is exists on flange cut
    IsProExists = Has_Attribute(pPropertyDescriptions.CAO, "CuttingBehavior")
    
    'Exit if attribute is not bulkloaded
    If Not IsProExists Then Exit Sub
        
    'Call PC Item member conditional and check weather it is True or False
    'Based on this Set FlangeCut Cutting Behaviour
    Set oMemberDescription = oMemberDescriptions.ItemByDispid(1)
    oMemberDescription.CAO = pPropertyDescriptions.CAO

    'Adding Error Resume Next to overcome DotNet ATPs where they are creating FlangeCuts alone.
    
    CM_IsPhysConnsNeeded oMemberDescription, bIsNeeded
    
    'Set Proper Flange Cutting Behaviour
    Set_FlangeCuttingBehavior oMemberDescription.CAO, bIsNeeded
    
    oMemberDescription.CAO = Nothing
  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
