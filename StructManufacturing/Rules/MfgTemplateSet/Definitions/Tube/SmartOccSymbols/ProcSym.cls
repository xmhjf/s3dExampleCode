VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProcSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'/*******************************************************************
'Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'
' Project       : StrMfgTemplateSOSymbolTube.vbp
'
' File          : ProcSym.cls
'
' Description   : Symbol for Tube Template Process.
'
'  History      :
'      Siva        20th October 2011    created
'
'*******************************************************************/

Option Explicit
Private Const MODULE = "StrMfgTemplateSOSymbolTube.ProcSym"
Private Const m_progID = "StrMfgTemplateSOSymbolTube.ProcSym"

Private Const REPRESENTATION_1 = "Process"

Private Const NUM_INPUTS = 11

Private Const INPUT_1 = "MinHeight"
Private Const INPUT_2 = "MaxHeight"
Private Const INPUT_3 = "Extension"
Private Const INPUT_4 = "Side"
Private Const INPUT_5 = "SideEnd"
Private Const INPUT_6 = "Type"
Private Const INPUT_7 = "BasePlane"
Private Const INPUT_8 = "TemplateService"
Private Const INPUT_9 = "UserDefinedValues"
Private Const INPUT_10 = "TemplateNaming"
Private Const INPUT_11 = "Bevel"

Implements IJDUserSymbolServices

'////////////////////////////////////////////////////////////////////
'/*******************************************************************
' Method: EditOccurence
'
' Interface: IJDUserSymbolServices
'
' Description:
'
'*******************************************************************/
Public Function IJDUserSymbolServices_EditOccurence(ByRef SymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
    Const METHOD = "IJDUserSymbolServices_EditOccurence"
    On Error GoTo ErrorHandler
    IJDUserSymbolServices_EditOccurence = False
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

'////////////////////////////////////////////////////////////////////
'/*******************************************************************
' Method: InitializeCommonInformation
'
' Description:
'
'       1. Declare overriding user custom methods.
'
'*******************************************************************/
Private Sub InitializeCommonInformation(ByVal SymbolDefinition As DSymbolDefinition)
    Const METHOD = "InitializeCommonInformation"
    On Error GoTo ErrorHandler
    
    '2. Declare user custom methods.
    Dim oMethodDesc As New DMethodDescription
    Dim oLibraryDesc As IJDLibraryDescription
    Set oLibraryDesc = New DLibraryDescription
    Dim oUserMethods As IJDUserMethods

    ' Query the IJDUserMethods interface
    Set oUserMethods = SymbolDefinition

    ' Define the standard library to overriden the above std methods
    oLibraryDesc.Name = "StdLib"
    oLibraryDesc.Type = imsLIBRARY_IS_ACTIVEX
    oLibraryDesc.Source = m_progID

    ' Set it to the symbol definition
    oUserMethods.SetLibrary oLibraryDesc, False
    oUserMethods.SetDefaultLibrary "StdLib"
  
    ' ReturnResult method
    oMethodDesc.Name = "ReturnResult"
    oMethodDesc.Properties = imsMETHOD_OVERRIDE
    oMethodDesc.Cookie = imsCOOKIE_IS_NULL
    oUserMethods.SetMethod oMethodDesc

Cleanup:
    Set oMethodDesc = Nothing
    Set oLibraryDesc = Nothing
    Set oUserMethods = Nothing

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
    GoTo Cleanup
End Sub

'////////////////////////////////////////////////////////////////////
'/*******************************************************************
' Method: GetDefinitionName
'
' Interface: IJDUserSymbolServices
'
' Description:
'       Generate symbol definition name from definitionParameters.
'
'*******************************************************************/
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    Const METHOD = "IJDUserSymbolServices_GetDefinitionName"
    On Error GoTo ErrorHandler
    
    IJDUserSymbolServices_GetDefinitionName = m_progID

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

'////////////////////////////////////////////////////////////////////
'/*******************************************************************
' Method: InstanciateDefinition
'
' Interface: IJDUserSymbolServices
'
' Description:
'       Create symbol definition. And initialize it.
'
'*******************************************************************/
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal DefParams As Variant, ByVal ActiveConnection As Object) As Object
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler
    
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Dim oSymbolDefinition As IJDSymbolDefinition

    ' Instanciate the symbol definition
    Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)

    ' Set the definition ProgId and CodeBase
    oSymbolDefinition.ProgId = m_progID
    oSymbolDefinition.CodeBase = CodeBase

    ' Feed the symbol definition with dynamic information (Inputs, Representation, Outputs)
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition

    oSymbolDefinition.Name = IJDUserSymbolServices_GetDefinitionName(DefParams)

    'returned symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition

Cleanup:
    Set oSymbolFactory = Nothing
    Set oSymbolDefinition = Nothing

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
    GoTo Cleanup
End Function

'////////////////////////////////////////////////////////////////////
'/*******************************************************************
' Method: InvokeRepresentation
'
' Interface: IJDUserSymbolServices
'
' Description:
'       This method isn't used. ReturnResult method create the outputs.
'
'*******************************************************************/
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal OutputColl As Object, ByRef arrayOfInputs())
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation"
    ' Do nothing: ReturnResult method is overriden
    Exit Sub
End Sub

'////////////////////////////////////////////////////////////////////
'/*******************************************************************
' Method: InitializeSymbolDefinition
'
' Interface: IJDUserSymbolServices
'
' Description:
'       Initialize symbol definition. Create outputs definition.
'
'*******************************************************************/
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
        
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    'pSymbolDefinition.EllipsisInputsOption = igSYMBOL_ELLIPSIS_INPUTS
    
    pSymbolDefinition.IJDUserMethods.RemoveAllMethods
    
    ' Set the common information on the symbol definition (Overriden methods, serverFileName)
    InitializeCommonInformation pSymbolDefinition

    'For P2R2 ( ICDPP2R2-PC00012 )
    pSymbolDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    '
    ' Define inputs
    '
    ' Create a new input by new operator
    Dim Coord(1 To NUM_INPUTS) As IMSSymbolEntities.IJDInput
    Dim ii As Integer
    
     'Add the custom validate method
    Dim ValidateMethodCookie As Long
    Dim oUserMethods As IJDUserMethods
    Set oUserMethods = pSymbolDefinition

    ValidateMethodCookie = oUserMethods.GetMethodCookie("CheckParameterValues", imsCOOKIE_ID_USS_LIB)
    
    Dim oCustomMethodIntf As IJDInputStdCustomMethod
    
    For ii = 1 To NUM_INPUTS
        Set Coord(ii) = New IMSSymbolEntities.DInput
        Set oCustomMethodIntf = Coord(ii).IJDInputStdCustomMethod
        oCustomMethodIntf.SetCMCheck imsCOOKIE_ID_USS_LIB, ValidateMethodCookie
    Next
    Set oCustomMethodIntf = Nothing
    ' The two first arguments are in fact the ports of the connection
    SetTemplateCommonInputs pSymbolDefinition

    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.IJDParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
    
    PC.Type = igValue
    
    Coord(1).Name = INPUT_1 'MinHeight
    Coord(1).Description = INPUT_1
    Coord(1).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 7100 '"Fixed"
    Coord(1).DefaultParameterValue = PC

    Coord(2).Name = INPUT_2 'MaxHeight
    Coord(2).Description = INPUT_2
    Coord(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 7100 '"Fixed"
    Coord(2).DefaultParameterValue = PC

    Coord(3).Name = INPUT_3 'Extension
    Coord(3).Description = INPUT_3
    Coord(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 7100 '"ExtnBothEnds_Fixed"
    Coord(3).DefaultParameterValue = PC

    Coord(4).Name = INPUT_4 'Side
    Coord(4).Description = INPUT_4
    Coord(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 7100  '"Outer"
    Coord(4).DefaultParameterValue = PC
    
    Coord(5).Name = INPUT_5 'SideEnd
    Coord(5).Description = INPUT_5
    Coord(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 7100  '"Base"
    Coord(5).DefaultParameterValue = PC

    Coord(6).Name = INPUT_6 'Type
    Coord(6).Description = INPUT_6
    Coord(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 7100  '"TypeShortBCL"
    Coord(6).DefaultParameterValue = PC

    Coord(7).Name = INPUT_7 'BasePlane
    Coord(7).Description = INPUT_7
    Coord(7).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 7100  '"BySystem"
    Coord(7).DefaultParameterValue = PC

    Coord(8).Name = INPUT_8 'Service
    Coord(8).Description = INPUT_8
    Coord(8).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 7100  '"Default"
    Coord(8).DefaultParameterValue = PC
    
    Coord(9).Name = INPUT_9 'UserDefinedValues
    Coord(9).Description = INPUT_9
    Coord(9).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 7100  '"Default"
    Coord(9).DefaultParameterValue = PC
    
    Coord(10).Name = INPUT_10 'TemplateNaming
    Coord(10).Description = INPUT_10
    Coord(10).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 7100  '"Default"
    Coord(10).DefaultParameterValue = PC

    Coord(11).Name = INPUT_11 'Bevel
    Coord(11).Description = INPUT_11
    Coord(11).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 7101  '"Ignore"
    Coord(11).DefaultParameterValue = PC
        
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    ' Offset the inputs to account for the template reference collection inputs
    For ii = 1 To NUM_INPUTS
      InputsIf.SetInput Coord(ii), ii + igTemplateIndex
      Set Coord(ii) = Nothing
    Next
 
    pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED

    ' Set the different representations
    Dim oReps As IMSSymbolEntities.IJDRepresentations
    Dim IJDOutputs As IMSSymbolEntities.IJDOutputs
    Dim oRep As IMSSymbolEntities.IJDRepresentation
    Set oRep = New IMSSymbolEntities.DRepresentation

    oRep.RepresentationId = 1  'define a aspect 0 (Simple_physical)
    oRep.Name = REPRESENTATION_1
    oRep.Description = REPRESENTATION_1
    oRep.Properties = igREPRESENTATION_VBFUNCTION

    Set oReps = pSymbolDefinition
    Set IJDOutputs = oRep
    Dim oOutput As IMSSymbolEntities.IJDOutput
    Set oOutput = New IMSSymbolEntities.DOutput
    oOutput.Properties = 0

    oReps.SetRepresentation oRep
    oRep.Reset

    'RepEval--------------------------------------
    Dim oRepEval As IJDRepresentationEvaluation

    Set oRepEval = New DRepresentationEvaluation
    oRepEval.Name = "TemplateProcess"
    oRepEval.Description = "evaluation function identification for the hanger representation"
    oRepEval.Properties = igREPRESENTATION_HIDDEN
    oRepEval.Type = igREPRESENTATION_VBFUNCTION
    oRepEval.ProgId = m_progID

    'RepEvals---------------------------------------------
    Dim oRepEvals As IJDRepresentationEvaluations

    Set oRepEvals = pSymbolDefinition
    oRepEvals.AddRepresentationEvaluation oRepEval
        
    '===========================================================================
    'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
    'WHICH ARE GRAPHIC ENTITIES.
    '===========================================================================
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

Cleanup:
    Set IJDOutputs = Nothing
    Set oOutput = Nothing
    Set oRep = Nothing
    Set oReps = Nothing
    Set oRepEval = Nothing
    Set oRepEvals = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
    GoTo Cleanup
End Sub

'////////////////////////////////////////////////////////////////////
'/*******************************************************************
' Method: ReturnResult
'
' Description:
'    This is the method of IJDSymbolDefinitionPlayer which is implemented
'   by symbol definition.
'    Input objects are processed in this method and create the outputs.
'
'*******************************************************************/
Public Sub ReturnResult(ByVal SymbolDef As IJDSymbolDefinition, ByVal SymbolOcc As IJDSymbol, _
                ByVal repName As String, ByVal OutputColl As IJDOutputCollection)
    Const METHOD = "ReturnResult"
    On Error GoTo ErrorHandler
       
'     'set symbol as invisible.
    Dim oControlFlag As IJControlFlags
    Set oControlFlag = SymbolOcc
'    oControlFlag.ControlFlags(CTL_FLAG_SYSTEM_MASK) = CTL_FLAG_SYSTEM_MASK
    'make the oObject invisible - so it is not highlited on mousemove
    oControlFlag.ControlFlags(CTL_FLAG_NO_DRAW Or CTL_FLAG_NO_LOCATE Or CTL_FLAG_NOT_IN_SPATIAL_INDEX) = CTL_FLAG_NO_DRAW Or CTL_FLAG_NO_LOCATE Or CTL_FLAG_NOT_IN_SPATIAL_INDEX
  
    Exit Sub
ErrorHandler:
End Sub

Public Sub CheckParameterValues(ByVal pInputCM As IMSSymbolEntities.IJDInputStdCustomMethod, ByRef ppArgument As Object, ByRef InputIsValid As Boolean, ByRef errMsg As String)
On Error Resume Next
    'Method called from the command to check that the argument given is a parameter
    Dim PC As IJDParameterContent
    InputIsValid = True
    Set PC = ppArgument
    Dim oInput As IJDInput
    Set oInput = pInputCM
End Sub
  
' There is  one reference input(PlatePart) for the template symbols
Private Sub SetTemplateCommonInputs(pDef As IJDSymbolDefinition)
    Dim pInput As IJDInput
    Set pInput = New DInput

    pInput.Name = "Tube"
    pInput.Description = "Tube"
    pInput.Index = igPlateIndex
    pDef.IJDInputs.Add pInput
    pInput.Reset
End Sub

 
