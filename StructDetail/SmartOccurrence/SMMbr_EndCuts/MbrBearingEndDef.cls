VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrBearingEndDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2007-2008 Intergraph Corporation. All rights reserved.
'
'File : MbrBearingEndDef.cls
'
'Author : D.A. Trent
'
'Description :
'   Control creation of the Member Items for the MbrAxisWebSel Selector Class
'
'
'History:
'
'********************************************************************

Option Explicit

Const m_sClassName As String = "MbrBearingEndDef"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Const IID_IJPlate = "{53CF4EA0-91BF-11D1-BE56-080036B3A103}"
Const IID_IJStructureMaterial = "{E790A7C0-2DBA-11D2-96DC-0060974FF15B}"

Implements IJDUserSymbolServices
'

' List any graphic Inputs that the Definition has here
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

' List any graphic Inputs that the Definition has here
Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
Const METHOD = m_DefinitionProgid & "::ItemAggregator"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim pPDs As IJDPropertyDescriptions
    
    sMsg = "Defining ItemAggregator"
    
    pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructBearingPlate"
    pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateAggregator"
    
    ' The Bearing Plate Matrial and Thickness is controlled (set)
    ' by the IJPlate and IJStructureMaterial interfaces
    ' We want to know when these are changed changed so that the
    ' dependent Web Cut can be updated if need be
        
    ' Set the thickness
    Set pPDs = pAD
    sMsg = "Setting Thickness property."
    pPDs.AddProperty "Thickness", 1, IID_IJPlate, _
                     "CM_Thickness", imsCOOKIE_ID_USS_LIB
    
    ' Set the material and grade
    sMsg = "Setting MaterialAndGrade property."
    pPDs.AddProperty "MatAndGrade", 2, IID_IJStructureMaterial, _
                     "CM_MatAndGrade", imsCOOKIE_ID_USS_LIB
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' List all the Item members that this SmartClass can create
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
Const METHOD = m_DefinitionProgid & "::ItemMembers"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sDispId As String
    
    Dim iDispId As Long
    Dim oMemDesc As IJDMemberDescription
    
    ' Bearing Plate consist of (3) Member Items,
    '   (1) Web Cut and (2) Physical Connections
    ' Index(1) - Physical Connection between Bounding Member and bearing Plate
    ' Index(2) - Web Cut between Bounded Member and bearing Plate
    For iDispId = 1 To 2
        sDispId = "BearingPlate_" & Trim(Str(iDispId))
        Set oMemDesc = pMDs.AddMember(sDispId, iDispId, "CM_ConstructItems", imsCOOKIE_ID_USS_LIB)
        oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsMemberItemNeeded"
        oMemDesc.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CM_CopyQuestionsToMember"
        oMemDesc.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CM_MigrateMemberItem"
        Set oMemDesc = Nothing
    Next iDispId

    
    Set oMemDesc = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Methods HERE
' *******************************************************************************************
Public Sub CM_MemberItemCount(pMemberDescription As IJDMemberDescription, nCount As Long)
Const METHOD = m_DefinitionProgid & "::CM_MemberItemCount"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim lStatus As Long
    nCount = 1
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Sub CM_IsMemberItemNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
Const METHOD = m_DefinitionProgid & "::CM_IsMemberItemNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim sSelection As String
    Dim sEndCutType As String
    
    Dim lDispId As Long
    
    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oSmartOccurrence As IJSmartOccurrence
    Dim oSymbolDefinition As IJDSymbolDefinition

    Dim oWebCutObject As Object

    Dim oCommonHelper As DefinitionHlprs.CommonHelper

    ' Get the Current Selection
    bIsNeeded = False
    lDispId = pMemberDescription.dispid
    
    ' Bearing Plate consist of (3) Member Items,
    '   (1) Web Cut and (2) Physical Connections
    ' Index(1) - Physical Connection between Bounding Member and bearing Plate
    ' Index(2) - Web Cut between Bounded Member and bearing Plate
    If lDispId < 3 Then
        bIsNeeded = True
    
    Else
        bIsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Sub CM_CopyQuestionsToMember(pMemberDescription As IJDMemberDescription)
Const METHOD = m_DefinitionProgid & "::CM_CopyQuestionsToMember"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Dim oItemObject As Object
    Dim oStructFeature As IJStructFeature
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    
    sMsg = "Setting Questions on Bearing Plate Items " & _
           "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
    
    Set oItemObject = pMemberDescription.object
    If oItemObject Is Nothing Then
        Exit Sub
    
    ElseIf TypeOf oItemObject Is IJStructFeature Then
        ' Force an Update on the EndCut (WebCut) based on the Bearing Plate updates
        Set oStructFeature = oItemObject
        If oStructFeature.get_StructFeatureType = SF_WebCut Then
            Set oSDO_WebCut = New StructDetailObjects.WebCut
            Set oSDO_WebCut.object = oItemObject
            oSDO_WebCut.ForceUpdateForFlangeCuts
            Set oSDO_WebCut = Nothing
        End If
        
        Set oStructFeature = Nothing

    End If

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Function CM_ConstructItems(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
Const METHOD = m_DefinitionProgid & "::CM_ConstructItems"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Creating Bearing Plate Items " & _
           "...pMemberDescription.dispid = " & Str(pMemberDescription.dispid)
    CreateModify_MemberItems pMemberDescription, pResourceManager, pObject
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

Public Sub CMFinalConstructBearingPlate(ByVal pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = m_DefinitionProgid & "::CMFinalConstructBearingPlate"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    BearingPlate_FinalConstruct pAggregatorDescription
    
    Dim oPlate As IJPlate
    Dim oBearingPlate As Object
    
    Dim oSDSmartPlateOps As GSCADSDCreateModifyUtilities.IJSDSmartPlateOperations
    
    Set oBearingPlate = pAggregatorDescription.CAO

    Set oSDSmartPlateOps = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
    oSDSmartPlateOps.PlaceBearingPlate oBearingPlate
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' *******************************************************************************************
Public Sub CM_MigrateAggregator(pAggregatorDescription As IJDAggregatorDescription, _
                                pMigrateHelper As IJMigrateHelper)
Const METHOD = m_DefinitionProgid & "::CM_MigrateAggregator"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    sMsg = "Mirgrating EndCut Object"
    MigrateEndCutObject pAggregatorDescription, pMigrateHelper

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

' *******************************************************************************************
Public Sub CM_MigrateMemberItem(pMemberDesc As IJDMemberDescription, _
                              pMigrateHelper As IJMigrateHelper)
Const METHOD = m_DefinitionProgid & "::CM_MigratePhysConn"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    ' The Pyhsical Connection Inputs have been Mirgrated at same time with the EndCut Inputs
    ' Therefore, there should be nothing to do here
    sMsg = "Mirgrating Pyhsical Connection Inputs"

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'***********************************************************************
Public Sub CM_Thickness(pPD As IJDPropertyDescription, pObject As Object)
Const METHOD = m_DefinitionProgid & "::CM_Thickness"
    On Error GoTo ErrorHandler
    Dim sMsg As String
       
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'***********************************************************************
Public Sub CM_MatAndGrade(pPD As IJDPropertyDescription, pObject As Object)
Const METHOD = m_DefinitionProgid & "::CM_MatAndGrade"
    On Error GoTo ErrorHandler
    Dim sMsg As String
       
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
    IJDUserSymbolServices_GetDefinitionName = m_DefinitionName
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
    pMDs.RemoveAll ' Remove all the previous member descriptions
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
    pDefinition.ProgId = m_DefinitionProgid
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

' ********************************************************************************************
Public Function IJDUserSymbolServices_EditOccurence(pSymbolOccurence As Object, _
                                                    ByVal transactionMgr As Object) As Boolean
End Function

' ********************************************************************************************
' ********************************************************************************************
' ********************************************************************************************
Private Function CreateModify_MemberItems(ByVal oMemberDescription As IJDMemberDescription, _
                                          ByVal oResourceManager As IUnknown, _
                                          ByRef oMemberObject As Object)
Const METHOD = m_DefinitionProgid & "::CreateModify_MemberItems"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sSmartItem As String
    
    Dim lDispId As Long
    
    Dim oTmpObject As Object
    Dim oBearingPlate As Object
    Dim oBearingParent As Object
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim oSystemChild As IJSystemChild
    Dim oSystemParent As Object 'IJSystem
    Dim oDesignParent As IJDesignParent
    Dim oGraphicInputs As JCmnShp_CollectionAlias
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim oSDO_PhysicalConn As StructDetailObjects.PhysicalConn
    Dim oSDSmartPlateAtt As GSCADSDCreateModifyUtilities.IJSDSmartPlateAttributes
    
    Set oBearingPlate = oMemberDescription.CAO
    Set oSDSmartPlateAtt = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
    
    oSDSmartPlateAtt.GetInputs_BearingPlate oBearingPlate, oGraphicInputs
    Set oBoundingObject = oGraphicInputs.Item(1)
    Set oBoundedObject = oGraphicInputs.Item(2)

    ' Bearing Plate consist of (3) Member Items,
    '   (1) Web Cut and (2) Physical Connections
    ' Index(1) - Physical Connection between Bounding Member and bearing Plate
    ' Index(2) - Web Cut between Bounded Member and Bearing Plate
    lDispId = oMemberDescription.dispid
    
    If lDispId = 1 Then


        ' Get Base Port from Bearing Plate (Use Late Port Binding)
        sMsg = "Getting Bearing Plate Bounded Port for Physical Connection"
        oSDSmartPlateAtt.GetLateBindPort oBearingPlate, _
                                         JS_TOPOLOGY_FILTER_SOLID_BASE_LFACE, _
                                         True, _
                                         oBoundedPort
        
        ' The Bounding Port will be: Lateral, Base, or Offset
        sMsg = "Getting Member Bounding Port for Physical Connection"
        Set oBoundingPort = Member_GetSolidPort(oBoundingObject)
                                     
        ' The Bearing Plate can not have children
        ' (it can not be the Physical Connections Parent system)
        ' Use the Assembly Connection has the Parent
        sMsg = "Setting Bearing Plate Physical Connection System Parent"
        If TypeOf oBearingPlate Is IJSystemChild Then
            Set oSystemChild = oBearingPlate
            Set oTmpObject = oSystemChild.GetParent
            If TypeOf oTmpObject Is IJDesignParent Then
                Set oDesignParent = oTmpObject
                If TypeOf oDesignParent Is IJSystem Then
                    Set oSystemParent = oDesignParent
                End If
            End If
        End If
       
        ' Create physical connection
        sMsg = "Creating Physical Connection"
        Set oSDO_PhysicalConn = New StructDetailObjects.PhysicalConn
        oSDO_PhysicalConn.Create oResourceManager, oBoundedPort, oBoundingPort, _
                                 "LapWeld", oSystemParent, ConnectionStandard
                               
        sMsg = "Setting Physical Connection to private variable"
        Set oMemberObject = oSDO_PhysicalConn.object
        
    ElseIf lDispId = 2 Then
        ' The Bounded Port will be: Base, or Offset
        sMsg = "Getting Bounded Port for the Bearing Web Cut"
        Set oBoundedPort = oBoundedObject
        
        ' Use the Bounding Member Port from the Assembly Connection
        ' This is required so that the Web Cut contour will match the
        ' the bearing Plate Surface (even for Tubular Cross Sections)
        ' The Web Cut contour MUST be Offset by the Bearing Plate Thickness
        Set oBoundingPort = oBoundingObject
        ' Get Offset Port from Bearing Plate (Use Late Port Binding)
        sMsg = "Getting Bearing Plate Bounding Port for Web Cut"
        oSDSmartPlateAtt.GetLateBindPort oBearingPlate, _
                                         JS_TOPOLOGY_FILTER_SOLID_OFFSET_LFACE, _
                                         False, _
                                         oBoundingPort
        
        ' The Bearing Plate can not have children
        ' (it can not be the Web Cut Parent system)
        ' Use the Assembly Connection has the Parent
        sMsg = "Setting Bearing Plate Web Cut System Parent"
        If TypeOf oBearingPlate Is IJSystemChild Then
            Set oSystemChild = oBearingPlate
            Set oTmpObject = oSystemChild.GetParent
            If TypeOf oTmpObject Is IJDesignParent Then
                Set oDesignParent = oTmpObject
                If TypeOf oDesignParent Is IJSystem Then
                    Set oSystemParent = oDesignParent
                End If
            End If
        End If
       
        sMsg = "Check if Bounding Member Cross Section is Tubular"
        If IsTubularMember(oBoundingObject) Then
            sSmartItem = "M_BearingPlate_Tube_01"
            
        ElseIf IsTubularMember(oBoundedObject) Then
            sSmartItem = "M_Web_Weld_Tube00"
        Else
            sSmartItem = "M_BearingPlate_Weld_01"
        End If
       
        sMsg = "Creating The Bearing Plate Web Cut"
        Set oSDO_WebCut = New StructDetailObjects.WebCut
        oSDO_WebCut.Create oResourceManager, oBoundingPort, oBoundedPort, _
                           sSmartItem, oSystemParent
                               
        sMsg = "Return the created Web Cut"
        Set oMemberObject = oSDO_WebCut.object
        
    Else
    End If

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function
