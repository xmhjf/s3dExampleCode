VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImportedAsmConnDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : ImportedAsmConnDef.cls
'
'Author : R. Peringoth
'
'Description :
'    Macro for assembly connection created during CIMSteel import
'
'History:
'
' May/03/07  RP   created
'********************************************************************

Option Explicit

Private Const MODULE = "ImportedAsmConnDef"

Const m_ItemProgId As String = "SPSACMacros.ImportedAsmConnDef"
Const m_ItemName As String = "SPSACMacros.ImportedAsmConnDef"
Private Const strSourceFile As String = "ImportedAsmConnDef.cls"

Private m_oLocalizer As IJLocalizer

Implements ISPSFACInputHelper
Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements IJUserAttributeMgmtParent
Implements IJStructCustomFoulCheck


'*************************************************************************
'Function
'DefinitionInputs
'
'Abstract
'Sets up the inputs for the assembly connection
'
'Arguments
'pIH As IJDInputsHelper - Input
'
'Exceptions
'
'***************************************************************************

Public Sub DefinitionInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:
  pIH.ReportError
End Sub




'*************************************************************************
'Function
'IJDUserSymbolServices_InitializeSymbolDefinition
'
'Abstract
'Bulk load called method that establishes the Base Plate assembly connection
'definition
'
'Arguments
'pDefinition - Symbol definition created for the base plate assembly connection
'
'Exceptions
'
'***************************************************************************
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pDefinition As IJDSymbolDefinition)
  Const MT = "IJDUserSymbolServices_InitializeSymbolDefinition"
  On Error GoTo ErrorHandler
  
  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
  ' Define the inputs for the base plate connection (the helper class is initialized to the
  ' symbol definition and is an aid to help set the inputs)
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.Definition = pDefinition
  DefinitionInputs pIH
  
  ' Aggregator Type
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  pAD.AggregatorClsid = "{E43FD681-1B37-4CC1-BD94-F399F43F946F}"     'CStructAssemblyConnection
  pAD.SetCMSetInputs -1, -1
  pAD.SetCMRemoveInputs -1, -1
  pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAggregator"
  pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateAggregator"
  
  Set pAD = Nothing
  
  Dim pCADefinition As IJCADefinition
  Set pCADefinition = pDefinition
  Let pCADefinition.CopyBackwardFlag = igCOPY_BACKWARD_TRIM
  Set pCADefinition = Nothing
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  
  ' add the interfaces here. so that semantic here is executed when those interfaces are modified
  pAPDs.AddProperty "ImportedAsmConn", 1, "IJDAttributes"
  Set pAPDs = Nothing
  
  
  '''' no children for this assembly connection
 

  Exit Sub
ErrorHandler:
  HandleError MODULE, MT
End Sub


'*************************************************************************
'Function
'CMFinalConstructAggregator
'
'Abstract
'Permits one-time initialization of the user-defined aggregator class.
'
'Arguments
'IJDAggregatorDescription interface of the CustomAssemblyDefinition wrapper
'
'Exceptions
'
'***************************************************************************
Public Sub CMFinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMFinalConstructAggregator"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub



Public Sub CMMigrateAggregator(pAggregatorDescription As IJDAggregatorDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateAggregator"
    On Error GoTo ErrorHandler
    

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub




'*************************************************************************
'Function
'IJDUserSymbolServices
'   InstanciateDefinition
'
'Abstract
'Instantiates a persistent symbol definition object and initializes it for the first time,
'returning a pointer (ppSymbolDefDisp) to the IDispatch interface of the initialized symbol definition.
'
'Arguments:
'codeBase specifies the URL (or UNC) of the .cab file that can provides the dll associated to the symbol definition object (ActiveX® control packaging).
'definitionParameters  Definition parameters.
'pResourceMgr  resource manager to which the symbol definition will be connected.
'
'Return:
'S_OK  Operation succeeded.
'E_FAIL  Operation failed (no detail).
'
'Exceptions:
'
'***************************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
  Const MT = "IJDUserSymbolServices_InstanciateDefinition"
  On Error GoTo ErrorHandler
  
  Dim pDefinition As IJDSymbolDefinition
  Dim pFact As IJCAFactory
  Set pFact = New CAFactory
  Set pDefinition = pFact.CreateCAD(ActiveConnection)
  
  ' Set definition progId and codebase
  pDefinition.ProgId = m_ItemProgId
  pDefinition.CodeBase = CodeBase
    
  ' Initialize the definition
  IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
  pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
  
  ' Persistence behavior
  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
  ' Returned symbol definition
  Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
  
  Exit Function
ErrorHandler:
  HandleError MODULE, MT
End Function


'*************************************************************************
'Function
'IJDUserSymbolServices
'   GetDefinitionName
'
'Abstract
'Used during the execution of IJDDefinitionCollection::GetDefinitionByProgId to get the definition name
'based upon the definitionParameters passed in. It returns the definition name (pDefName) if it already
'exists within the collection. The name of a definition is the identifier of the definition object
'in the definition collection and assures its uniqueness in the given resource manager.
'
'Arguments
'definitionParameters
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_ItemName
End Function


'*************************************************************************
'Function
'
'IJDUserSymbolServices
'   InvokeRepresentation
'
'Abstract
'Obsolete method
'
'***************************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub


'*************************************************************************
'Function
'IJDUserSymbolServices_EditOccurence
'
'Abstract
'Obsolete method --instead you can record your custom command within the
'definition (see IJDCommandDescription interface)
'
'***************************************************************************
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function


'*************************************************************************
'Function
'IJStructCustomFoulCheck
'   GetConnectedParts
'
'Abstract
'Adds objects that are related with the SmartOccurrence ( input objects ) to the list of connected objects
'
'Arguments
'pPartObject is the SmartOccurrence object
'pIJMonUnks is the list of monikers
'
'Return
'
'Exceptions
'
'***************************************************************************
Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements)
End Sub


'*************************************************************************
'Function
'IJStructCustomFoulCheck
'   GetFoulInterfaceType
'
'Abstract
'Returns type of interference participant that this object is.
'
'Arguments
'pFoulInterfaceType is the InterferenceChecking type
'
'Return
'
'Exceptions
'
'***************************************************************************
Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
    pFoulInterfaceType = NonParticipant
End Sub


'*************************************************************************
'Function
'IJUserAttributeMgmt
'   OnAttributeChange
'
'Abstract
'Gets called for each attribute change on the property page
'
'Arguments
'pIJDAttrs is the list of all persistent attributes of the BusinessObject
'CollAllDisplayedValues is the list of attributes as currently displayed ( prior to Commit )
'pAttrToChange is which attribute is being edited
'varNewAttrValue is the value given by the user.
'
'Return
'String value should be "" for no error, and an error string to be displayed to the user
'if erroneous input was given.
'
'Exceptions
'
'***************************************************************************
Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
End Function


'*************************************************************************
'Function
'IJUserAttributeMgmt
'   OnPreCommit
'
'Abstract
'Gets called before the attribute changes are committed to allow a check of validity.
'
'Arguments
'pIJDAttrs is the list of all persistent attributes of the BusinessObject
'CollAllDisplayedValues is the list of attributes as currently displayed ( prior to Commit )
'
'Return
'String value should be "" for no error, and an error string to be displayed to the user
'if erroneous input was given.
'
'Exceptions
'
'***************************************************************************
Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
End Function


'*************************************************************************
'Function
'IJUserAttributeMgmt
'   OnPreLoad
'
'Abstract
'Gets called prior to display of attributes on the property page to set readOnly status
'
'Arguments
'pIJDAttrs is the list of all persistent attributes of the BusinessObject
'CollAllDisplayedValues is the list of IJAttributeDescriptor's
'
'
'Return
'String value should be "" for no error, and an error string to be displayed to the user
'if erroneous input was given.
'
'Exceptions
'
'***************************************************************************
Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
    Dim ii As Long
    Dim bWithPad As Boolean
    Dim oCollProxy As CollectionProxy
    
    Set oCollProxy = pIJDAttrs.CollectionOfAttributes("IJUAGenSurfAsmConnPad")
    
    bWithPad = oCollProxy.Item("WithPad").Value
    If Not bWithPad = True Then
        'If not with pad then gray out the attributes regarding pad
        Dim pAttrColl As Collection
        Set pAttrColl = CollAllDisplayedValues
        Dim pAttrDescr As IJAttributeDescriptor
        For ii = 1 To pAttrColl.count
            Set pAttrDescr = pAttrColl.Item(ii)
            If pAttrDescr.InterfaceName = "IJUAGenSurfAsmConnPad" Then
                If pAttrDescr.AttrName = "SizingRule" Or pAttrDescr.AttrName = "PadType" Or pAttrDescr.AttrName = "Offset" Then
                    pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
                End If
            End If
        Next
    End If
    
End Function


Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SPSMembers.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String
End Function


Private Function IJUserAttributeMgmtParent_OnAttributeChange(ByVal pParent As Object, ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SP3DStructInterfaces.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
'
End Function

Private Function IJUserAttributeMgmtParent_OnPreCommit(ByVal pParent As Object, ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
'
End Function

Private Function IJUserAttributeMgmtParent_OnPreLoad(ByVal pParent As Object, ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

'*************************************************************************
'Function
'ISPSFACInputHelper
'   UserAttributeMgmt
'
'Abstract
'Function on ISPSFACInputHelper to return the UserAttributeMgmt interface
'
'Arguments
'
'Return
'The interface is returned.
'
'Exceptions
'
'***************************************************************************
Private Property Get ISPSFACInputHelper_UserAttributeMgmt() As SP3DStructInterfaces.IJUserAttributeMgmt
End Property
'*************************************************************************
'Function
'ISPSFACInputHelper
'   ValidateObjects
'
'Abstract
'Validates the given objects and returns the objects to be used for assemblyConnection inputs
'
'Arguments
'inputObjs is a collection of objects
'relationObjs is the collection to be used to build relationships
'
'Return
'String value should be "" for no error, and an error string to be displayed to the user
'if erroneous input was given.
'
'Exceptions
'
'***************************************************************************


Private Property Get ISPSFACInputHelper_ValidateObjects(ByVal inputObjs As SP3DStructInterfaces.IJElements, relationObjs As SP3DStructInterfaces.IJElements) As SP3DStructInterfaces.SPSFACInputHelperStatus
    Const MT = "ISPSFACInputHelper_ValidateObjects"
    On Error GoTo ErrorHandler
   
    ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_Ok
    Exit Property
ErrorHandler:
    HandleError MODULE, MT
End Property



Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub

