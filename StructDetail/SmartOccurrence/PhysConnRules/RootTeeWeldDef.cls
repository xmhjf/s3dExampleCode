VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RootTeeWeldDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = "PhysConnRules.RootTeeWeldDef"
Const m_ItemName As String = "PhysConnRules.RootTeeWeldDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\PhysConnRules\RootTeeWeldDef.cls"

Implements IJDUserSymbolServices

Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    pIH.SetInput INPUT_CONN_OBJECT1
    pIH.SetInput INPUT_CONN_OBJECT2
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs").Number
End Sub

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    
    Dim oMemDesc As IJDMemberDescription
        
    Set oMemDesc = pMDs.AddMember("AutomaticSplitter", 1, "CMConstructAutomaticSplitter", imsCOOKIE_ID_USS_LIB)
    oMemDesc.SetCMConditional "PhysConnRules.RootTeeWeldDef", "CMIsAutomaticSplitterNeeded"
    oMemDesc.SetCMCount "PhysConnRules.RootTeeWeldDef", "CMAutomaticSplitterCount"

    Set oMemDesc = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers").Number
End Sub
' ** Start CM **
Public Sub CMAutomaticSplitterCount(ByVal pMemberDescription As IJDMemberDescription, _
                                    ByRef nCount As Long)
    On Error GoTo ErrorHandler
    
    Dim oPCWrapper As New StructDetailObjects.PhysicalConn
    
    Set oPCWrapper.object = pMemberDescription.CAO
    oPCWrapper.GetAutomaticSplitterLocationCount nCount
    Set oPCWrapper = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMAutomaticSplitterCount").Number
End Sub
Public Sub CMIsAutomaticSplitterNeeded(ByVal pMD As IJDMemberDescription, _
                                       ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    
    ' This method is only called on root pc.
    
    ' Note:
    ' 1. Custom assembly still calls this method with index = 0
    '    (thus no corresponding object) when CMCount is 0.
    '    In this case, return false
    ' 2. Custom assembly should make sure index is not greater than CMCount.
    
    ' If index is invalid, return False
    ' If original object does not exist yet, return True
    ' Otherwise,need to check if it still valid
    '    If still valid, update it if needed
    '    Otherwise, return False to delete it
    
    If pMD.Index < 1 Then
        ' Invalid input
        bIsNeeded = False
        Exit Sub
    End If
    
    If pMD.object Is Nothing Then
        bIsNeeded = True
        Exit Sub
    End If
       
    bIsNeeded = True
    '
    ' Perform additional check here if needed
    ' Customer additional check begins
    '    Check...
    ' Customer additional check begins
    
    ' If this member is needed, update it if needed
    If bIsNeeded = True And Not pMD.object Is Nothing Then
        Dim dOrgX As Double
        Dim dOrgY As Double
        Dim dOrgZ As Double
        Dim oPoint As IJPoint
        
        Set oPoint = pMD.object
        oPoint.GetPoint dOrgX, dOrgY, dOrgZ
        
        Dim oPhysConnWrapper As New StructDetailObjects.PhysicalConn
        Dim oLocation As Object
        Dim oLocationPosition As IJDPosition
        
        Dim dCurX As Double
        Dim dCurY As Double
        Dim dCurZ As Double
        
        Set oPhysConnWrapper.object = pMD.CAO
        oPhysConnWrapper.GetAutomaticSplitterLocation pMD.Index, oLocation
        Set oLocationPosition = oLocation
        Set oLocation = Nothing
        oLocationPosition.Get dCurX, dCurY, dCurZ
        Set oLocationPosition = Nothing
        If Abs(dOrgX - dCurX) > 0.0001 Or _
           Abs(dOrgY - dCurY) > 0.0001 Or _
           Abs(dOrgZ - dCurZ) > 0.0001 Then
            oPoint.SetPoint dCurX, dCurY, dCurZ
        End If
           
        Set oPhysConnWrapper = Nothing
        Set oPoint = Nothing
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMIsAutomaticSplitterNeeded").Number
End Sub

Public Sub CMConstructAutomaticSplitter(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObject As Object)
    On Error GoTo ErrorHandler

    ' Construct splitter
    Dim oPhysConn As Object
    Dim oPCWrapper As New StructDetailObjects.PhysicalConn
    Dim oOriginalObject As Object
    Dim oCalculatedLocation As Object
    
    Set oPhysConn = pMemberDescription.CAO
    Set oPCWrapper.object = oPhysConn
    Set oOriginalObject = pMemberDescription.object

    oPCWrapper.GetAutomaticSplitterLocation pMemberDescription.Index, oCalculatedLocation
    
    If oOriginalObject Is Nothing Then
        ' Not exist yet,create a new one
        oPCWrapper.PlaceAutomaticSplitter oCalculatedLocation, pObject
    Else
        ' Existing one, should have already been updated in CMIsAutomaticSplitterNeeded method
        Set pObject = oOriginalObject
    End If
    Set oCalculatedLocation = Nothing
    Set oOriginalObject = Nothing
    Set oPhysConn = Nothing
    Set oPCWrapper = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructAutomaticSplitter").Number

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
      
    Dim oMemberDescriptions As IJDMemberDescriptions
    
    Set oMemberDescriptions = pDefinition
    oMemberDescriptions.RemoveAll
      
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

Public Sub ReportError(Optional ByVal sFunctionName As String, Optional ByVal sErrorName As String)
    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description _
                & " - " & "::" & sFunctionName & " - " & sErrorName
End Sub

