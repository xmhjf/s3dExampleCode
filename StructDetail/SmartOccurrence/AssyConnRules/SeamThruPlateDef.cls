VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeamThruPlateDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = "AssyConnRules.SeamThruPlateDef"
Const m_ItemName As String = "AssyConnRules.SeamThruPlateDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\AssyConnRules\SeamThruPlateDef.bas"

Implements IJDUserSymbolServices

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

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    ' TODO - Add aggregator description below
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    strError = "Constructing seam edge feature."
    pMDs.AddMember "SeamThruPlateEdgeFeature1", 1, "CMConstructEdgeFeature", imsCOOKIE_ID_USS_LIB
    
    Dim pMD As IJDMemberDescription
    strError = "Constructing collar."
    Set pMD = pMDs.AddMember("SeamThruPlateCollar1", 2, "CMConstructCollar", imsCOOKIE_ID_USS_LIB)
    
    ' Copy the Q/A in the Assembly that must be sent to the Collar
    pMD.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMCopyQuestionsToMembers"
    
    '  there are three properties on the collar to be set
    '  We could create three Custom Methods, but why reproduce code to do close to the same thing
    '  A good convention is to create a Property custom method based on the interface
    
    Dim pPDs As IJDPropertyDescriptions
    Set pPDs = pMD
    
    ' Set the thickness
    strError = "Setting Thickness property."
    pPDs.AddProperty "Thickness", 1, IID_IJPlate, "CMUpdateThickness", imsCOOKIE_ID_USS_LIB
    
    ' Set the material and grade
    strError = "Setting MaterialAndGrade property."
    pPDs.AddProperty "MatAndGrade", 2, IID_IJStructureMaterial, "CMUpdateMatAndGrade", imsCOOKIE_ID_USS_LIB
    
    ' Set the Side Of Plate
    strError = "Setting Side Of Plate property."
    pPDs.AddProperty "SideOfPlate", 3, IID_IJCollarPart, "CMSetSideOfPlate", imsCOOKIE_ID_USS_LIB
     

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers", strError).Number
End Sub

' ** Start CM **
Public Sub CMConstructCollar(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM
    Dim pMemberObjects As IJDMemberObjects
    
    strError = "Retrieving CAO and constructing collar."
    Set pMemberObjects = pMemberDescription.CAO
    Set pObject = clsAssyConnCM.CMConstruct_Collar(pMemberObjects.Item(1), "RootClip")
    
    Set pMemberObjects = pMemberDescription.CAO
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructCollar", strError).Number
End Sub

Public Sub CMUpdateThickness(pPD As IJDPropertyDescription, pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM
    
    strError = "Updating Thickness property."
    Call clsAssyConnCM.CMCopyParentThickness(pObject)
    
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMUpdateThickness", strError).Number
End Sub

Public Sub CMUpdateMatAndGrade(pPD As IJDPropertyDescription, pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM
    
    strError = "Updating MaterialAndGrade property."
    Call clsAssyConnCM.CMCopyParentMatAndGrade(pObject)
    
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMUpdateMatAndGrade", strError).Number
End Sub

Public Sub CMConstructEdgeFeature(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM
    
    strError = "Constructing seam edge feature."
    ' TO BE DETERMINED ***** Set pObject = clsAssyConnCM.CMConstruct_EdgeFeature(?, ?, "?RootEdgeFeature?")
    
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructEdgeFeature", strError).Number
End Sub
Public Sub CMSetSideOfPlate(pPD As IJDPropertyDescription, pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM
    
    strError = "Updating SideOfPlate property."
    Call clsAssyConnCM.CMSetSideOfPlate(pObject)
    
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMSetSideOfPlate", strError).Number
End Sub
Public Sub CMCopyQuestionsToMembers(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    strError = "Calling Copy answer to copy assembly questions to collar"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.CopyAnswer "CollarRules.RootClipSel", "StressLevel", _
                                 "AssyConnRules.PenetrationConnSel", "StressLevel"
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCopyQuestionsToMemebers", strError).Number
End Sub


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
    pIH.Definition = pDefinition
    pIH.InitAs m_FamilyProgid
    ItemInputs pIH
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Set pAD = pDefinition
    ItemAggregator pAD
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
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

