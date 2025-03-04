VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemberContourTESym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---------------------------------------------------------------------------
'    Copyright (C) 2008 Intergraph Corporation. All rights reserved.
'
'
'
'History
'    SS         10/29/08      Creation
'---------------------------------------------------------------------------------------
Option Explicit

Private Const MODULE = "MemberContourTESym"
Private Const MYSYMBOL_PROGID = "SPSProgrammedShapes.MemberContourTESym"
Private Const MYSYMBOL_LIBNAME = "SPSProgrammedShapes_MemberContourTESym"

Implements IJDUserSymbolServices
Implements IStructInsulationGraphicInputFilter

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    Const METHOD = "IJDUserSymbolServices_GetDefinitionName"
    On Error GoTo ErrorHandler

    IJDUserSymbolServices_GetDefinitionName = MYSYMBOL_PROGID
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
    'Define the inputs
    
    DefineEncasementSymbolInputs pSymbolDefinition
   
    DefineEncasementSymbolOutput MYSYMBOL_LIBNAME, MYSYMBOL_PROGID, pSymbolDefinition
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal pResourceMgr As Object) As Object
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler

    Dim pSymbolDefinition As IJDSymbolDefinition
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    
    Set pSymbolDefinition = oSymbolFactory.CreateEntity(definition, pResourceMgr)
    Set oSymbolFactory = Nothing
    
    pSymbolDefinition.ProgId = MYSYMBOL_PROGID
    pSymbolDefinition.CodeBase = CodeBase
    pSymbolDefinition.Name = IJDUserSymbolServices_GetDefinitionName(definitionParameters)
    pSymbolDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    pSymbolDefinition.SupportOnlyOption = igSYMBOL_SUPPORT_ONLY    'go away when no occurences
    
    IJDUserSymbolServices_InitializeSymbolDefinition pSymbolDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = pSymbolDefinition

    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

Public Sub CMEvaluateEncasementCrossSection(pRepSCM As IJDRepresentationStdCustomMethod)
    Const METHOD = "CMEvaluateEncasementCrossSection"
    On Error GoTo ErrorHandler
 
    EvaluateContourCrossSection pRepSCM, True, False, False, False

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
    Err.Raise E_FAIL
End Sub

Private Function IStructInsulationGraphicInputFilter_FilterGraphicInput(ByVal oGraphicObject As Object, ByVal oDefinition As Object) As Long
Const METHOD = "IStructInsulationGraphicInputFilter_FilterGraphicInput"
On Error GoTo ErrorHandler

    IStructInsulationGraphicInputFilter_FilterGraphicInput = FilterMembers(oGraphicObject, oDefinition)

    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Function IStructInsulationGraphicInputFilter_GraphicInputFilterCriteria() As String
    IStructInsulationGraphicInputFilter_GraphicInputFilterCriteria = "[" & MYSYMBOL_PROGID & ",FilterMembers]"
End Function

Private Function IStructInsulationGraphicInputFilter_SingleGraphicInput() As Boolean
    IStructInsulationGraphicInputFilter_SingleGraphicInput = True
End Function

Public Function FilterMembers(oGraphicObject As Object, oDefinition As Object) As Integer
Const METHOD = "FilterMembers"
On Error GoTo ErrorHandler

    FilterMembers = ValidateMemberFPCriteria(oGraphicObject, oDefinition)
    If FilterMembers = StructInsulationInputHelper_Ok Then
        FilterMembers = 1
    Else
        FilterMembers = 0
    End If
    
    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function

