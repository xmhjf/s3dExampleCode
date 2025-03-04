VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateThruPlateCutoutDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = CUSTOMERID + "AssyConRul.PlateThruPlateCutoutDef"
Const m_ItemName As String = CUSTOMERID + "AssyConRul.PlateThruPlateCutoutDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\PlateThruPlateCutoutDef.bas"

Implements IJDUserSymbolServices

Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    Dim strError As String

    strError = "Setting bounded and bounding objects as inputs."
    pIH.SetInput INPUT_BOUNDED_OR_PENETRATED_OBJECT
    pIH.SetInput INPUT_BOUNDING_OR_PENETRATING_OBJECT
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs").Number
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

    strError = "Constructing physical connection."
    Dim pMember1Description As IJDMemberDescription
    Set pMember1Description = pMDs.AddMember("PlateThruPlateCutoutDef0", 1, "CMCreateCutout", imsCOOKIE_ID_USS_LIB)
    
    Call pMember1Description.SetCMConditional(imsCOOKIE_ID_USS_LIB, "CMPenetratorDetailed")
    
    'Add the compute property
    Dim pPropertyDescriptions As IJDPropertyDescriptions
    Set pPropertyDescriptions = pMember1Description
        pPropertyDescriptions.AddProperty "ComputeCutout", 1, "IJGeometry", "CMComputeCutout", imsCOOKIE_ID_USS_LIB

    Dim pMember2Description As IJDMemberDescription
    Set pMember2Description = pMDs.AddMember("PlateThruPlateCutoutDef1", 2, "CMConstructPC1", imsCOOKIE_ID_USS_LIB)
    
    Call pMember2Description.SetCMConditional(imsCOOKIE_ID_USS_LIB, "CMPenetratorDetailed")
    
    Dim pMember3Description As IJDMemberDescription
    Set pMember3Description = pMDs.AddMember("PlateThruPlateCutoutDef2", 3, "CMConstructPC2", imsCOOKIE_ID_USS_LIB)
    
    Call pMember3Description.SetCMConditional(imsCOOKIE_ID_USS_LIB, "CMPenetratorDetailed")
    
    Dim pMember4Description As IJDMemberDescription
    Set pMember4Description = pMDs.AddMember("PlateThruPlateCutoutDef3", 4, "CMConstructPC3", imsCOOKIE_ID_USS_LIB)
    
    Call pMember4Description.SetCMConditional(imsCOOKIE_ID_USS_LIB, "CMPenetratorDetailed")
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers", strError).Number
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CMPenetratorDetailed is a CM conditionnal function
' bIsDetailed is true if the penetrator is detailed
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMPenetratorDetailed(ByRef pMemberDescription As IJDMemberDescription, _
                                ByRef bIsDetailed As Boolean)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    
    ' Get the penetrating plate from the AC port
    strError = " Getting the penetrating part"
    
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
 
    Set pAssyConn.object = pMemberDescription.CAO
    
    Dim oPart As IJDStructApplyOperation
    Set oPart = pAssyConn.ConnectedObject2 ' must be the penetrating by rules
        
    Dim strTrimProgId As String
    strTrimProgId = "ThickenTrim.TrimPlate_AE.1"
    
    bIsDetailed = DoesOperationExists(strTrimProgId, oPart)
    
    If bIsDetailed Then
        If Not TypeOf oPart Is IJModelBody Then
            bIsDetailed = False
        End If
    End If
        
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMComputeMember", strError).Number
End Sub
' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMCreateCutout(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM

    strError = "Creating"
    
    Set pObject = clsAssyConnCM.CMCreateCutout(pMemberDescription, pResourceManager, BPT_Base)
    
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCreateCutout", strError).Number
End Sub
Public Sub CMComputeCutout(pPropertyDescription As IJDPropertyDescription, pObject As Object)
    Dim strError As String
    
    strError = "computing cutout"
    Dim oCutout As New StructDetailObjects.Cutout
    Set oCutout.object = pObject
    
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    strError = "Setting Assembly Connection Inputs"
    Set pAssyConn.object = pPropertyDescription.CAO
    
    strError = "Getting Assembly Connection Parts associated to ports"
    Dim oPort1Part As Object
    Dim oPort2Part As Object
    Set oPort1Part = pAssyConn.ConnectedObject1 ' must be the penetrated by the rules
    Set oPort2Part = pAssyConn.ConnectedObject2 ' must be the penetrating by rules
    
    strError = "Call the update cutout operation"
    Call oCutout.UpdateCutout(oPort2Part, oPort1Part, pAssyConn.object)
    
    Set oCutout = Nothing
Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCreateCutout", strError).Number
End Sub
Public Sub CMConstructPC1(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM

    strError = "Constructing physical connection1"
    Set pObject = clsAssyConnCM.CMConstruct_PhysConnBetweenLastPorts(pMemberDescription, _
                                                                     pResourceManager, _
                                                                     BPT_Base, "TeeWeld")
    
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructPC1", strError).Number
End Sub
Public Sub CMConstructPC2(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM

    strError = "Constructing physical connection2"
    Set pObject = clsAssyConnCM.CMConstruct_PhysConnBetweenLastPorts(pMemberDescription, _
                                                                     pResourceManager, _
                                                                     BPT_Offset, "TeeWeld")
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructPC2", strError).Number
End Sub
Public Sub CMConstructPC3(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM

    strError = "Constructing physical connection3"
    Set pObject = clsAssyConnCM.CMConstruct_PhysConnBetweenLastPorts(pMemberDescription, _
                                                                     pResourceManager, _
                                                                     BPT_Lateral, "TeeWeld")
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructPC3", strError).Number
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

Private Function DoesOperationExists(oOpProgID As String, oApplyOp As IJDStructApplyOperation) As Boolean
On Error GoTo ErrorHandler
    
    Dim strError As String
    
    Dim bResult As Boolean
    Dim oOpColl As New Collection
    Dim oOpLoop As Long
    Dim oProgIDObject As IJDStructOperationProgID
    bResult = False
    
    ' in-parameter test
    If Not oApplyOp Is Nothing And oOpProgID <> "" Then
    
        strError = "Getting operations ProgIds"
        ' fill collection with operation progIDs
        oApplyOp.GetOperationProgIDs oOpColl
        
        strError = "Looping throug ProgIds"
        ' loop to find the requested operation progID
        For oOpLoop = 1 To oOpColl.Count
            Set oProgIDObject = oOpColl.Item(oOpLoop)
            
            ' compare progIDs
            If oProgIDObject.ProgId = oOpProgID Then
                bResult = True
                Exit For
            End If
        Next oOpLoop
    End If
    
    ' release objects
    Set oOpColl = Nothing
    Set oProgIDObject = Nothing
    
    DoesOperationExists = bResult
    
Exit Function
ErrorHandler:
Err.Raise LogError(Err, MODULE, "DoesOperationExists", strError).Number
End Function

