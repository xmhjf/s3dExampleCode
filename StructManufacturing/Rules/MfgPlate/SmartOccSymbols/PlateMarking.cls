VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateMarking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'/*******************************************************************
'Copyright (C) 1998, Intergraph Corporation.  All rights reserved.
'
'
'Project: S:\StructManufacturing\Middle\Symbols\ProcessAndMarkingSymbols\StructMfgSymbols.vbp
'
'File: S:\StructManufacturing\Middle\Symbols\ProcessAndMarkingSymbols\PlateMarking.cls
'
'Revision:
'     02/07/01 GDreybus.
'
'Description:
'     Create the symbol parameters for plate marking.
'
'*******************************************************************/

Option Explicit
Private Const MODULE = "StructMfgSymbols.PlateMarking(PlateMarking.cls)"

Private Const m_progID = "StructMfgSymbols.PlateMarking"

Private Const REPRESENTATION_1 = "PlateMarking"

Private Const NUM_INPUTS = 33

Private Const INPUT_1 = "PlPlateLocation"
Private Const INPUT_2 = "PlProfileLocation"
Private Const INPUT_3 = "PlCollarLocation"
Private Const INPUT_4 = "PlBracketLocation"
Private Const INPUT_5 = "PlLapConnection"
Private Const INPUT_6 = "PlEndConnection"

Private Const INPUT_8 = "PlFittingPL2PLTJoint"
Private Const INPUT_9 = "PlFittingPL2PrPen"
Private Const INPUT_10 = "PlFittingPL2Pr"
Private Const INPUT_11 = "PlRefPlanesFR"
Private Const INPUT_12 = "PlRefPlanesWL"
Private Const INPUT_13 = "PlRefPlanesBL"
Private Const INPUT_14 = "PlBlockControl"
Private Const INPUT_15 = "PlSeamControl"
Private Const INPUT_16 = "PlTemplateLine"
Private Const INPUT_17 = "PlBaseControlLine"
Private Const INPUT_18 = "PlRollLines"
Private Const INPUT_19 = "PlKnuckleLines"
Private Const INPUT_20 = "PlNavalArchLines"
Private Const INPUT_21 = "PlRollBoundaries"
Private Const INPUT_22 = "PlShipDirection"
Private Const INPUT_23 = "PlDiagonals"
Private Const INPUT_24 = "PlRobotMark"
Private Const INPUT_25 = "PlPaintingLine"
Private Const INPUT_26 = "PlMark"
Private Const INPUT_27 = "PlPlateEdgeChkLines"
Private Const INPUT_28 = "PlWebFramesChkLines"
Private Const INPUT_29 = "PlBridge"
Private Const INPUT_30 = "PlFeatureTab"
Private Const INPUT_31 = "PlWeldingTab"
Private Const INPUT_32 = "PlBendingControlLine"
Private Const INPUT_33 = "PlMargin"
Private Const INPUT_7 = "PlUserDefined"

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

Public Function IJDUserSymbolServices_EditOccurence(ByRef SymbolOccurence As Object, _
                                                    ByVal TransactionMgr As Object) As Boolean
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
    Dim oLibraryDesc As New DLibraryDescription
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
    Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)

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
    
    'Add the custom validate method
    Dim ValidateMethodCookie As Long
    Dim oUserMethods As IJDUserMethods
    Set oUserMethods = pSymbolDefinition
    ValidateMethodCookie = oUserMethods.GetMethodCookie("CheckParameterValues", imsCOOKIE_ID_USS_LIB)
    Dim oCustomMethodIntf As IJDInputStdCustomMethod
        
    Dim ii As Integer
    For ii = 1 To NUM_INPUTS
        Set Coord(ii) = New IMSSymbolEntities.DInput
        Set oCustomMethodIntf = Coord(ii).IJDInputStdCustomMethod
        oCustomMethodIntf.SetCMCheck imsCOOKIE_ID_USS_LIB, ValidateMethodCookie
    Next
    Set oCustomMethodIntf = Nothing

    ' The two first arguments are in fact the ports of the connection
    SetPlateCommonInputs pSymbolDefinition

    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.IJDParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
    PC.Type = igValue
    ' PC.String = "Apply"
    
    Coord(1).Name = INPUT_1     'PlateLocation
    Coord(1).Description = INPUT_1
    Coord(1).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12010
    Coord(1).DefaultParameterValue = PC

    Coord(2).Name = INPUT_2     'ProfileLocation
    Coord(2).Description = INPUT_2
    Coord(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12020
    Coord(2).DefaultParameterValue = PC

    Coord(3).Name = INPUT_3     'CollarLocation
    Coord(3).Description = INPUT_3
    Coord(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12030
    Coord(3).DefaultParameterValue = PC

    Coord(4).Name = INPUT_4     'BracketLocation
    Coord(4).Description = INPUT_4
    Coord(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12040
    Coord(4).DefaultParameterValue = PC
        
    Coord(5).Name = INPUT_5     'LapConnection
    Coord(5).Description = INPUT_5
    Coord(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12060
    Coord(5).DefaultParameterValue = PC

    Coord(6).Name = INPUT_6     'EndConnection
    Coord(6).Description = INPUT_6
    Coord(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12050
    Coord(6).DefaultParameterValue = PC
    
    Coord(8).Name = INPUT_8     'FittingPL2PLTJoint
    Coord(8).Description = INPUT_8
    Coord(8).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12080
    Coord(8).DefaultParameterValue = PC

    Coord(9).Name = INPUT_9     'FittingPL2PrPen
    Coord(9).Description = INPUT_9
    Coord(9).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12090
    Coord(9).DefaultParameterValue = PC

    Coord(10).Name = INPUT_10     'FittingPL2Pr
    Coord(10).Description = INPUT_10
    Coord(10).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12100
    Coord(10).DefaultParameterValue = PC

    Coord(11).Name = INPUT_11     'RefPlanesFR
    Coord(11).Description = INPUT_11
    Coord(11).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12110
    Coord(11).DefaultParameterValue = PC

    Coord(12).Name = INPUT_12     'RefPlanesWL
    Coord(12).Description = INPUT_12
    Coord(12).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12120
    Coord(12).DefaultParameterValue = PC

    Coord(13).Name = INPUT_13     'RefPlanesBL
    Coord(13).Description = INPUT_13
    Coord(13).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12130
    Coord(13).DefaultParameterValue = PC

    Coord(14).Name = INPUT_14     'BlockControl
    Coord(14).Description = INPUT_14
    Coord(14).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12140
    Coord(14).DefaultParameterValue = PC

    Coord(15).Name = INPUT_15     'SeamControl
    Coord(15).Description = INPUT_15
    Coord(15).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12150
    Coord(15).DefaultParameterValue = PC

    Coord(16).Name = INPUT_16     'TemplateLine
    Coord(16).Description = INPUT_16
    Coord(16).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12160
    Coord(16).DefaultParameterValue = PC

    Coord(17).Name = INPUT_17     'BaseControlLine
    Coord(17).Description = INPUT_17
    Coord(17).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12170
    Coord(17).DefaultParameterValue = PC

    Coord(18).Name = INPUT_18     'RollLines
    Coord(18).Description = INPUT_18
    Coord(18).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12180
    Coord(18).DefaultParameterValue = PC

    Coord(19).Name = INPUT_19     'KnuckleLines
    Coord(19).Description = INPUT_19
    Coord(19).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12190
    Coord(19).DefaultParameterValue = PC

    Coord(20).Name = INPUT_20     'NavalArchLines
    Coord(20).Description = INPUT_20
    Coord(20).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12200
    Coord(20).DefaultParameterValue = PC

    Coord(21).Name = INPUT_21     'RollBoundaries
    Coord(21).Description = INPUT_21
    Coord(21).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12210
    Coord(21).DefaultParameterValue = PC

    Coord(22).Name = INPUT_22     'ShipDirection
    Coord(22).Description = INPUT_22
    Coord(22).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12220
    Coord(22).DefaultParameterValue = PC

    Coord(23).Name = INPUT_23     'Diagonals
    Coord(23).Description = INPUT_23
    Coord(23).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12230
    Coord(23).DefaultParameterValue = PC

    Coord(24).Name = INPUT_24     'RobotMark
    Coord(24).Description = INPUT_24
    Coord(24).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12240
    Coord(24).DefaultParameterValue = PC

    Coord(25).Name = INPUT_25     'PaintingLine
    Coord(25).Description = INPUT_25
    Coord(25).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12250
    Coord(25).DefaultParameterValue = PC

    Coord(26).Name = INPUT_26     'Mark
    Coord(26).Description = INPUT_26
    Coord(26).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12260
    Coord(26).DefaultParameterValue = PC

    Coord(27).Name = INPUT_27     'PlateEdgeChkLines
    Coord(27).Description = INPUT_27
    Coord(27).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12270
    Coord(27).DefaultParameterValue = PC

    Coord(28).Name = INPUT_28     'WebFramesChkLines
    Coord(28).Description = INPUT_28
    Coord(28).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12280
    Coord(28).DefaultParameterValue = PC

    Coord(29).Name = INPUT_29     'Bridge
    Coord(29).Description = INPUT_29
    Coord(29).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12290
    Coord(29).DefaultParameterValue = PC

    Coord(30).Name = INPUT_30     'FeatureTab
    Coord(30).Description = INPUT_30
    Coord(30).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12300
    Coord(30).DefaultParameterValue = PC

    Coord(31).Name = INPUT_31     'WeldingTab
    Coord(31).Description = INPUT_31
    Coord(31).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12310
    Coord(31).DefaultParameterValue = PC

    Coord(32).Name = INPUT_32     'BendingControlLine
    Coord(32).Description = INPUT_32
    Coord(32).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12320
    Coord(32).DefaultParameterValue = PC

    Coord(33).Name = INPUT_33     'Margin
    Coord(33).Description = INPUT_33
    Coord(33).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12330
    Coord(33).DefaultParameterValue = PC

    Coord(7).Name = INPUT_7     'UserDefined
    Coord(7).Description = INPUT_7
    Coord(7).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 12360
    Coord(7).DefaultParameterValue = PC

     ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    ' Offset the inputs to account for the plate reference collection inputs
    For ii = 1 To NUM_INPUTS
      InputsIf.SetInput Coord(ii), ii + igPlateIndex
      Set Coord(ii) = Nothing
    Next

    pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED
    'pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_SHARED

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
    oRepEval.Name = "PlateMarking"
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
'    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE

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
       
     'set symbol as invisible.
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
'    Select Case oInput.Name
'
'        'Distance parameters - negative values not applicable
'        Case Is = INPUT_1, INPUT_3, INPUT_4, INPUT_7, INPUT_9, _
'                  INPUT_11, INPUT_13, INPUT_14, INPUT_16
'
'            If PC.UomValue < 0 Then
'                InputIsValid = False
'                errMsg = oInput.Name & " cannot be negative"
'            End If
'
'        'string parameters
'        Case Is = INPUT_6
'            If LCase(PC.String) <> "base" Or LCase(PC.String) <> "offset" Then
'                InputIsValid = False
'                errMsg = INPUT_3 & " must be Base or Offset"
'            End If
'
'        Case Is = INPUT_5
'            If LCase(PC.String) <> "leg" Or LCase(PC.String) <> "throat" Then
'                InputIsValid = False
'                errMsg = INPUT_3 & " must be Leg or Throat"
'            End If
'
'        Case Else
'            InputIsValid = False
'            errMsg = "Unexpected input"
'    End Select
            
        
    
  End Sub



 