VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stiff_WP_2B_CF_Def"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'*******************************************************************
'
'Copyright (C) 2007 Intergraph Corporation. All rights reserved.
'
'File : Stiff_WP_2B_CF_Def.cls
'
'Author :
'
'Description :
'   Control creation of the Member Items for the GenericStiffConnSel Selector Class
'
'
'History:
'
'********************************************************************


Const m_ItemProgid As String = CUSTOMERID + "AssyConRul.Stiff_WP_2B_CF_Def"
Const m_ItemName As String = CUSTOMERID + "AssyConRul.Stiff_WP_2B_CF_Def"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\Stiff_WP_2B_CF_Def.cls"

Implements IJDUserSymbolServices

' List any graphic Inputs that the Definition has here
Public Sub ItemInputs(pIH As IJDInputsHelper)
Const METHOD = m_ItemProgid & "::ItemInputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining Selector Inputs"
    
    pIH.SetInput INPUT_BOUNDED_OR_PENETRATED_OBJECT
    pIH.SetInput INPUT_BOUNDING_OR_PENETRATING_OBJECT
    
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' List any graphic Inputs that the Definition has here
Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
Const METHOD = m_ItemProgid & "::ItemAggregator"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ItemAggregator"

    'CStructAssemblyConnection
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' List all the Item members that this SmartClass can create
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)

    Const METHOD = m_ItemProgid & "::ItemMembers"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String

    Dim oMemDesc As IJDMemberDescription

    Set oMemDesc = pMDs.AddMember("WebCut1", 1, "CM_CreateGenericWebCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsGenericWebCutNeeded"
    Set oMemDesc = Nothing
   
    Set oMemDesc = pMDs.AddMember("WebCut2", 2, "CM_CreateGenericWebCut", CUSTOMERID + "MbrAC.MbrACDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsGenericWebCutNeeded"
    Set oMemDesc = Nothing
        
    Set oMemDesc = pMDs.AddMember("CornerFeature", 3, "CM_ConstructCornerFeature", CUSTOMERID + "MbrAC.WP_2B_CF_Def")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsCornerFeatureCutNeeded"
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
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique

    IJDUserSymbolServices_GetDefinitionName = m_ItemName

End Function

' ********************************************************************************************
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
    Dim pPDs As IJDPropertyDescriptions
    Dim pAD As IJDAggregatorDescription
    Set pAD = pDefinition
    Set pPDs = pAD
    pPDs.RemoveAll
    ItemAggregator pAD
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    pMDs.RemoveAll
    ItemMembers pMDs
    
End Sub

' ********************************************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, _
                                                            ByVal defParams As Variant, _
                                                            ByVal ActiveConnection As Object) As Object
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

' ********************************************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                                                      ByVal repName As String, _
                                                      ByVal outputcoll As Object, _
                                                      ByRef arrayOfInputs())
End Sub

Public Function IJDUserSymbolServices_EditOccurence(pSymbolOccurence As Object, _
                                                    ByVal transactionMgr As Object) As Boolean
End Function

' *******************************************************************************************
Public Sub CM_IsCornerFeatureCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)

    Const METHOD = m_ItemProgid & "::CM_IsCornerFeatureCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    bIsNeeded = False

    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_StiffenerEndCutCornerFeature) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    ' --------------------------------------------------------------------
    ' Collect information on the relationship between bounding and bounded
    ' --------------------------------------------------------------------
    Dim i As Long

    Dim oAppConnection As IJAppConnection
    Dim oConnAttrbs As GSCADSDCreateModifyUtilities.IJSDConnectionAttributes
    
    Set oAppConnection = pMemberDescription.CAO
    Set oConnAttrbs = New SDConnectionUtils

    Dim oBoundingObjectColl As JCmnShp_CollectionAlias
    Set oBoundingObjectColl = New Collection
    
    For i = 1 To oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Count
        oBoundingObjectColl.Add oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Item(i)
    Next
    
    Dim oAssyConn As StructDetailObjects.AssemblyConn
    Set oAssyConn = New AssemblyConn
    Set oAssyConn.object = pMemberDescription.CAO
    
    oBoundingObjectColl.Add oAssyConn.Port2

    Dim oBoundingObject As Object
    Dim oPlateObjects As New Collection
    Dim oMemberObjects As New Collection

    For i = 1 To oBoundingObjectColl.Count

        Set oBoundingObject = oBoundingObjectColl.Item(i)

        If TypeOf oBoundingObject.Connectable Is IJConnectable Then

            If TypeOf oBoundingObject.Connectable Is IJPlate Then
                oPlateObjects.Add oBoundingObject
            ElseIf TypeOf oBoundingObject.Connectable Is ISPSMemberPartPrismatic Or _
                   TypeOf oBoundingObject.Connectable Is IJProfile Then
                oMemberObjects.Add oBoundingObject
            End If
        End If
    Next i

    ' ----------------------------------------------------------
    ' If there are two members, make sure they are not collinear
    ' ----------------------------------------------------------
    If oMemberObjects.Count = 2 Then

        Dim oBoundedData As MemberConnectionData
        Dim oBoundingData As MemberConnectionData
        Dim lStatus As Long

        InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg

        Dim bColinear As Boolean
        Dim bEndToEnd As Boolean
        Dim bRightAngle As Boolean

        If lStatus <> 0 Then
            Exit Sub
        End If

        CheckEndToEndConnection oBoundedData.MemberPart, oBoundingData.MemberPart, bEndToEnd, bColinear, bRightAngle

        If Not bColinear Then
            bIsNeeded = True
        End If
    ' --------------------------------------------------------
    ' If there are two plates, make sure they are not coplanar
    ' --------------------------------------------------------
    ' To Do

    ' ---------------------------------------
    ' Otherwise, assume the feature is needed
    ' ---------------------------------------
    Else
        bIsNeeded = True
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Sub CM_IsGenericWebCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)

    Const METHOD = m_ItemProgid & "::CM_IsCornerFeatureCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    sMsg = "Setting the conditional for web cut"

    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_WebCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
     bIsNeeded = True

Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

