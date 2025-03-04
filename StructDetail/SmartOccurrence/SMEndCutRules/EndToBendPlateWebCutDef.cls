VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EndToBendPlateWebCutDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMEndCutRules
'  File        : EndToBendPlateWebCutDef.cls
'
'  Description :
'
'   Item definition: basic template to define an item
'
'  Author      : Alligators
'
'  History     :
'    12/Jan/2012 - svsmylav
'           CR-174918: Added method for PC creation for profile (pseudo-knuckle case).
'*********************************************************************************************

Implements IJDUserSymbolServices

'------------- Give the Selector Name and ProgID
Const m_CAName As String = "EndToBendPlateWebCutDef"
Const m_CAProjectName As String = CUSTOMERID + "EndCutRules"

Const m_CAProgid As String = m_CAProjectName + "." + m_CAName
Const m_FamilyProgid As String = ""

'User names for item members
Const CA_WEBCUT = "{6441B309-DD8B-47CA-BB23-6FC6C0605628}"
Const CA_AGGREGATE = "{727935F4-EBB7-11D4-B124-080036B9BD03}"   ' CLSID of JCSmartOccurrence

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "EndCutRules\EndToBendPlateWebCutDef.cls"

'Newly created member objects
Private m_oPhysConn As Object
'

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    
    pAD.UserTypeClsid = CA_WEBCUT       '
    pAD.AggregatorClsid = CA_AGGREGATE     ' CStructFeature
    pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructWebCut"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

'******************************************************************************
'******************************************************************************
'******************************************************************************
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    
    ' Set Q/A to create free edge treatments
    Dim pMemDesc As IJDMemberDescription
    Set pMemDesc = pMDs.AddMember("WebCutsFET", 1, "CMConstructFET", CMLIBRARY_ENDCUTRULES)
    pMemDesc.SetCMConditional CMLIBRARY_ENDCUTRULES, "CMGetEdgeTreatment"
    Set pMemDesc = Nothing

    ' Set Q/A to Create WebCut PC
    Set pMemDesc = pMDs.AddMember("WebCutPC", 2, "CMConstructPC", imsCOOKIE_ID_USS_LIB)
    pMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsPCNeeded"
    Set pMemDesc = Nothing

  Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMFinalConstructWebCut(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Dim clsWebCutCM As New WebCutDefCM
    
    clsWebCutCM.WebCut_FinalConstruct pAggregatorDescription
    
    Set clsWebCutCM = Nothing
    Set m_oPhysConn = Nothing
End Sub
'***********************************************************************
' METHOD:  CM_IsPCNeeded
'
' DESCRIPTION:  Conditional for physical connection
'
'***********************************************************************
Public Sub CM_IsPCNeeded(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = True

    Exit Sub

'ErrorHandler:
'    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

'***********************************************************************
' METHOD:  CMConstructPC
'
' DESCRIPTION:  Creates PC between Plate and Stiffener
'
'***********************************************************************
Public Sub CMConstructPC(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, ByRef pObject As Object)
                                        
    Dim clsWebCutCM As New WebCutDefCM
    Dim sRootClass As String

    'go get the port of the bounding object
    Dim dPortValue As Long
    dPortValue = clsWebCutCM.GetBoundingProfilePort(pMemberDescription, pResourceManager)

    ' Butt Weld Physical Connection is required:
    '   for Profile bounded by Plate Edge cases
    ' Tee Weld Physical Connection is required:
    '   for other cases
    '
    If IsProfileBoundedByPlateEdge(pMemberDescription) Then
        sRootClass = "ButtWeld"
    Else
        sRootClass = "TeeWeld"
    End If

    Set m_oPhysConn = CAConstruct_PhysConnNames(pMemberDescription, _
                                                        pResourceManager, _
                                                        sRootClass, _
                                                        dPortValue, _
                                                        JXSEC_IDEALIZED_BOUNDARY)
    Set pObject = m_oPhysConn

    Set clsWebCutCM = Nothing

End Sub

' ** End CM **

' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_CAProgid
End Function

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, _
                                    ByVal defParams As Variant, _
                                    ByVal ActiveConnection As Object) As Object
Const sMETHOD = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler
    
    Dim pDefinition As IJDSymbolDefinition
    Dim pCAFactory As New CAFactory
    
    Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
    
    ' Set definition progId and codebase
    pDefinition.ProgId = m_CAProgid
    pDefinition.CodeBase = CodeBase
    pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
      
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)
Const sMETHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
      
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
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
End Function


Private Function CAConstruct_PhysConnNames(ByVal pMemberDescription As IJDMemberDescription, _
                                    ByVal pResourceManager As IUnknown, _
                                    strStartClass As String, _
                                    ByVal eBoundingSubPort As JXSEC_CODE, _
                                    ByVal eBoundedSubPort As JXSEC_CODE) As Object
    
    ' Get Class
    Dim pWebCut As New StructDetailObjects.WebCut
    Dim oFlangCut As New StructDetailObjects.FlangeCut
    Dim sError As String
    sError = "Setting web cut Inputs"

    
    ' Get Bounding Part's Port
    Dim oBoundingPart As Object ' This is the Bounding part
    Dim oBoundingPort As IJPort
    sError = "Getting Bounding object"
    ' Get Bounding part from wrapper
    Dim phelper As StructDetailObjects.Helper
    Set phelper = New StructDetailObjects.Helper
    
    Dim oBounding As Object

    Dim oFeature As GSCADSDCreateModifyUtilities.IJStructFeature
    Set oFeature = pMemberDescription.CAO
    
    Dim oSmOcc As IJSmartOccurrence
    Dim oSmItem As IJSmartItem
    Dim sItemName As String
    
    If oFeature.get_StructFeatureType = SF_WebCut Then
        Set pWebCut.object = pMemberDescription.CAO
        Set oBoundingPart = pWebCut.Bounding
        Set oSmOcc = pWebCut.object
        Set oSmItem = oSmOcc.SmartItemObject
        sItemName = oSmItem.Name
    End If
    
    If TypeOf oBoundingPart Is IJPlate Then
        Dim pSDOHelper As New StructDetailObjects.Helper
        Dim oLastPlatePort As IJPort
        If oFeature.get_StructFeatureType = SF_WebCut Then
            Set oLastPlatePort = pWebCut.BoundingPort
            Set oBoundingPort = pSDOHelper.GetEquivalentLastPort(oLastPlatePort)
        End If
    End If
   
    ' Get bounded profile Port
        Dim oBoundedPort As IJPort
        Dim oBoundedPart As New StructDetailObjects.ProfilePart
        Dim oBoundedBeamPart As New StructDetailObjects.BeamPart
        If oFeature.get_StructFeatureType = SF_WebCut Then
            If phelper.ObjectType(pWebCut.Bounded) = SDOBJECT_STIFFENER Then
                Set oBoundedPart = New StructDetailObjects.ProfilePart
                Set oBoundedPart.object = pWebCut.Bounded
                Set oBoundedPort = oBoundedPart.CutoutSubPort(pWebCut.object, eBoundedSubPort)
            End If
        End If
    
    ' Get the web cut, make it a parent of the PC
    Dim pSystemParent As IJSystemChild ' this is the slot's ca
    sError = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
    
    ' Create physical connection
    Dim oPhysicalConnection As New PhysicalConn
    sError = "Creating Physical Connection"
    oPhysicalConnection.Create pResourceManager, _
                                oBoundedPort, _
                               oBoundingPort, _
                               strStartClass, _
                               pSystemParent, ConnectionStandard
    
    sError = "Setting Physical Connection to private variable"
    Set CAConstruct_PhysConnNames = oPhysicalConnection.object

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number 'LogError(Err, MODULE, sMETHOD, sError).Number
        
End Function

