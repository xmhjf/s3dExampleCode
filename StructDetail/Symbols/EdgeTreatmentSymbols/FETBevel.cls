VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FETBevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'/*******************************************************************
'Copyright (C) 2005, Intergraph Corporation.  All rights reserved.
'
'
'Project: S:\StructDetail\Middle\Symbols\EdgeTreatmentSymbols\FreeEdgeTreatmentSym.vbp
'
'File: S:\StructDetail\Middle\Symbols\EdgeTreatmentSymbols\FETBevel.cls
'
'Revision:
'     08/24/05 Venu Kasarla.
'
'Description:
'     Create trace curve.
'
'
' !!!   !!!   !!!   !!!   !!!   !!!   !!!   !!!   !!!   !!!   !!!   !!!   !!!
' NOTE TO USERS:
'     Changes to the following fields are not permitted:
'            ReferencePart
'            RootGap
'            Radius
'            Nose
'            NoseOrientationAngle
'            RefSideFirstBevelDepth
'            RefSideSecondBevelDepth
'            RefSideFirstBevelAngle
'            RefSideSecondBevelAngle
'            AntiRefSideFirstBevelDepth
'            AntiRefSideSecondBevelDepth
'            AntiRefSideFirstBevelAngle
'            AntiRefSideSecondBevelAngle
'            NoseMethod
'            RefSideFirstBevelMethod
'            AntiRefSideFirstBevelMethod
'            RefSideSecondBevelMethod
'            AntiRefSideSecondBevelMethod
' !!!   !!!   !!!   !!!   !!!   !!!   !!!   !!!   !!!   !!!   !!!   !!!   !!!
'
'*******************************************************************/

Option Explicit
Private Const MODULE = "S:\StructDetail\Middle\Symbols\EdgeTreatmentSymbols\FETBevel.cls"

Private Const m_progID = "FreeEdgeTreatmentSym.FETBevel"

Private Const REPRESENTATION_1 = "FETBevel"

Private Const OUTPUT_1 = "Curve_1"

Private Const INPUT_1 = "RootGap"
Private Const INPUT_2 = "ReferencePart"
Private Const INPUT_3 = "ReferenceSide"
Private Const INPUT_4 = "Nose"
Private Const INPUT_5 = "NoseMethod"
Private Const INPUT_6 = "NoseOrientationAngle"
Private Const INPUT_7 = "RefSideFirstBevelDepth"
Private Const INPUT_8 = "RefSideFirstBevelMethod"
Private Const INPUT_9 = "RefSideFirstBevelAngle"

Private Const INPUT_10 = "AntiRefSideFirstBevelDepth"
Private Const INPUT_11 = "AntiRefSideFirstBevelMethod"
Private Const INPUT_12 = "AntiRefSideFirstBevelAngle"
Private Const INPUT_13 = "Radius"
Private Const INPUT_14 = "RefSideSecondBevelDepth"
Private Const INPUT_15 = "RefSideSecondBevelMethod"
Private Const INPUT_16 = "RefSideSecondBevelAngle"

Private Const INPUT_17 = "AntiRefSideSecondBevelDepth"
Private Const INPUT_18 = "AntiRefSideSecondBevelMethod"
Private Const INPUT_19 = "AntiRefSideSecondBevelAngle"

Private Const INPUT_20 = "RefPartName"

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
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
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
Private Sub InitializeCommonInformation(ByVal SymbolDefinition As DSymbolDefinition, _
                                        ByVal SymbolDefName As String)
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
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
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

    ' Set the common information on the symbol definition (Overriden methods, serverFileName)
    InitializeCommonInformation oSymbolDefinition, DefParams

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

    'For P2R2 ( ICDPP2R2-PC00012 )
    pSymbolDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    '
    ' Define inputs
    '
    ' Create a new input by new operator
    Dim Coord(1 To 20) As IMSSymbolEntities.IJDInput
    Dim ii As Integer
    
     'Add the custom validate method
    Dim ValidateMethodCookie As Long
    Dim oUserMethods As IJDUserMethods
    Set oUserMethods = pSymbolDefinition

    ValidateMethodCookie = oUserMethods.GetMethodCookie("CheckParameterValues", imsCOOKIE_ID_USS_LIB)
    
    Dim oCustomMethodIntf As IJDInputStdCustomMethod
     
    
    For ii = 1 To 20
        Set Coord(ii) = New IMSSymbolEntities.DInput
        Set oCustomMethodIntf = Coord(ii).IJDInputStdCustomMethod
        oCustomMethodIntf.SetCMCheck imsCOOKIE_ID_USS_LIB, ValidateMethodCookie
    Next
    Set oCustomMethodIntf = Nothing
    ' The two first arguments are in fact the ports of the connection
    SetCommonInputs pSymbolDefinition

    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.IJDParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
    
    PC.Type = igValue
    Coord(1).Name = INPUT_1 'root gap
    Coord(1).Description = INPUT_1
    Coord(1).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    PC.UomType = UNIT_DISTANCE
    Coord(1).DefaultParameterValue = PC
    
    PC.Reset
    Coord(2).Name = INPUT_2 'reference Part
    Coord(2).Description = INPUT_2
    Coord(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 1#
    PC.UomType = 0 'not a measurement unit
    Coord(2).DefaultParameterValue = PC

    PC.Reset
    PC.Type = igString
    Coord(3).Name = INPUT_3 'reference side
    Coord(3).Description = INPUT_3
    Coord(3).Properties = igINPUT_IS_A_PARAMETER
    PC.String = "Molded"
    Coord(3).DefaultParameterValue = PC

    PC.Reset
    PC.Type = igValue
    Coord(4).Name = INPUT_4 ' nose
    Coord(4).Description = INPUT_4
    Coord(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomType = UNIT_DISTANCE
    PC.UomValue = 0#
    Coord(4).DefaultParameterValue = PC
    
    PC.Reset
    Coord(5).Name = INPUT_5 'nose orientation angle method
    Coord(5).Description = INPUT_5
    Coord(5).Properties = igINPUT_IS_A_PARAMETER
    PC.UomType = 0
    PC.UomValue = 5
    Coord(5).DefaultParameterValue = PC

    PC.Reset
    PC.Type = igValue
    Coord(6).Name = INPUT_6 'nose orientation angle
    Coord(6).Description = INPUT_6
    Coord(6).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    PC.UomType = UNIT_ANGLE
    Coord(6).DefaultParameterValue = PC

    PC.Reset
    PC.Type = igValue
    Coord(7).Name = INPUT_7 'first bevel depth
    Coord(7).Description = INPUT_7
    Coord(7).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    PC.UomType = UNIT_DISTANCE
    Coord(7).DefaultParameterValue = PC
    
    PC.Reset
    PC.Type = igValue
    Coord(8).Name = INPUT_8 'RefSideFirstBevelAngleMethod
    Coord(8).Description = INPUT_8
    Coord(8).Properties = igINPUT_IS_A_PARAMETER
    PC.UomType = 0
    PC.UomValue = 5
    Coord(8).DefaultParameterValue = PC

    PC.Reset
    PC.Type = igValue
    Coord(9).Name = INPUT_9 'ref first bevel angle
    Coord(9).Description = INPUT_9
    Coord(9).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    PC.UomType = UNIT_ANGLE
    Coord(9).DefaultParameterValue = PC

    PC.Reset
    PC.Type = igValue
    Coord(10).Name = INPUT_10 'AntiRefSideFirstBevelDepth
    Coord(10).Description = INPUT_10
    Coord(10).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    PC.UomType = UNIT_DISTANCE
    Coord(10).DefaultParameterValue = PC
    
    PC.Reset
    PC.Type = igValue
    Coord(11).Name = INPUT_11 'AntiRefSideFirstBevelAngleMethod
    Coord(11).Description = INPUT_11
    Coord(11).Properties = igINPUT_IS_A_PARAMETER
    PC.UomType = 0
    PC.UomValue = 5
    Coord(11).DefaultParameterValue = PC

    PC.Reset
    PC.Type = igValue
    Coord(12).Name = INPUT_12 'AntiRefSideFirstBevelAngle
    Coord(12).Description = INPUT_12
    Coord(12).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    PC.UomType = UNIT_ANGLE
    Coord(12).DefaultParameterValue = PC

    PC.Reset
    PC.Type = igValue
    Coord(13).Name = INPUT_13 'Radius
    Coord(13).Description = INPUT_13
    Coord(13).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    PC.UomType = UNIT_DISTANCE
    Coord(13).DefaultParameterValue = PC

    PC.Reset
    PC.Type = igValue
    Coord(14).Name = INPUT_14 'RefSideSecondBevelDepth
    Coord(14).Description = INPUT_14
    Coord(14).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    PC.UomType = UNIT_DISTANCE
    Coord(14).DefaultParameterValue = PC
    
    PC.Reset
    PC.Type = igValue
    Coord(15).Name = INPUT_15 'RefSideSecondBevelAngleMethod
    Coord(15).Description = INPUT_15
    Coord(15).Properties = igINPUT_IS_A_PARAMETER
    PC.UomType = 0
    PC.UomValue = 5
    Coord(15).DefaultParameterValue = PC

    PC.Reset
    PC.Type = igValue
    Coord(16).Name = INPUT_16 'RefSideSecondBevelAngle
    Coord(16).Description = INPUT_16
    Coord(16).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    PC.UomType = UNIT_ANGLE
    Coord(16).DefaultParameterValue = PC

    PC.Reset
    PC.Type = igValue
    Coord(17).Name = INPUT_17 'AntiRefSideSecondBevelDepth
    Coord(17).Description = INPUT_17
    Coord(17).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    PC.UomType = UNIT_DISTANCE
    Coord(17).DefaultParameterValue = PC
    
    PC.Reset
    PC.Type = igValue
    Coord(18).Name = INPUT_18 'AntiRefSideSecondBevelAngleMethod
    Coord(18).Description = INPUT_18
    Coord(18).Properties = igINPUT_IS_A_PARAMETER
    PC.UomType = 0
    PC.UomValue = 5
    Coord(18).DefaultParameterValue = PC

    PC.Reset
    PC.Type = igValue
    Coord(19).Name = INPUT_19 'AntiRefSideSecondBevelAngle
    Coord(19).Description = INPUT_19
    Coord(19).Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0#
    PC.UomType = UNIT_ANGLE
    Coord(19).DefaultParameterValue = PC
    
    PC.Reset
    PC.Type = igString
    Coord(20).Name = INPUT_20 'ref part name
    Coord(20).Description = INPUT_20
    Coord(20).Properties = igINPUT_IS_A_PARAMETER
    PC.String = "Part1"
    Coord(20).DefaultParameterValue = PC
    
    ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition
'
    ' Offset the inputs to account for the one input ports
    For ii = 1 To 20
      InputsIf.SetInput Coord(ii), ii + igRefPartIndex
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
    'Representation name is Hanger
    oRep.Name = REPRESENTATION_1
    oRep.Description = REPRESENTATION_1
    oRep.Properties = igREPRESENTATION_VBFUNCTION

    Set oReps = pSymbolDefinition
    Set IJDOutputs = oRep
    Dim oOutput As IMSSymbolEntities.IJDOutput
    Set oOutput = New IMSSymbolEntities.DOutput
    oOutput.Properties = 0

'    oOutput.Name = OUTPUT_1
'    oOutput.Description = OUTPUT_1
'    'oRep.SetOutput oOutput
'    IJDOutputs.SetOutput oOutput

    oReps.SetRepresentation oRep
    oRep.Reset

    'RepEval--------------------------------------
    Dim oRepEval As IJDRepresentationEvaluation

    Set oRepEval = New DRepresentationEvaluation
    oRepEval.Name = "FETBevel"
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
'    Dim oControlFlag As IJControlFlags
'    Set oControlFlag = SymbolOcc
'    oControlFlag.ControlFlags(&HFF) = &HFF

    'Filter inputs
    Dim oEnumArg As IEnumJDArgument
    Set oEnumArg = SymbolOcc.IJDValuesArg.GetValues(igINPUT_ARGUMENTS_MERGE)
    
'    Dim oLine3d As ILines3d
'    Dim oGeomFactory As IJGeometryFactory
'    Set oGeomFactory = New GeometryFactory
'    Set oLine3d = oGeomFactory.Lines3d
'    Dim oLine As IJLine
'    Set oLine = oLine3d.CreateBy2Points(OutputColl.ResourceManager, 0#, 0#, 0#, 0#, 0#, 1#)
'
'    ' Mark the plane as support-only
'    UpdateControlFlags oLine, CTL_FLAG_CACHE
'
'    OutputColl.AddOutput OUTPUT_1, oLine

    
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
    Select Case oInput.Name
        
        'Distance parameters - negative values not applicable
        Case Is = INPUT_1, INPUT_7, INPUT_10, _
                  INPUT_13, INPUT_14, INPUT_17
            
            If PC.Type <> igValue Then
                InputIsValid = False
                errMsg = oInput.Name & " cannot be negative"
            End If
            
        ' Angle parameters - negative or positive
        Case Is = INPUT_6, INPUT_9, INPUT_12, INPUT_16, _
                  INPUT_19
            If PC.Type <> igValue Then
                InputIsValid = False
                errMsg = oInput.Name & " must be numeric"
            End If
        
        ' numeric values with a specific range
        Case Is = INPUT_2
'            If PC.UomValue < 1 Or PC.UomValue > 2 Then
'                InputIsValid = False
'                errMsg = oInput.Name & "must be 1 or 2"
'            End If
        
        'plain old numeric parameters
        Case Is = INPUT_4
            If PC.Type <> igValue Then
                InputIsValid = False
                errMsg = oInput.Name & " must be numeric"
            End If
        
        'string parameters
        Case Is = INPUT_3
            If LCase(PC.String) <> "base" And LCase(PC.String) <> "offset" And _
            LCase(PC.String) <> "molded" And LCase(PC.String) <> "antimolded" And _
            LCase(PC.String) <> "starboard" And LCase(PC.String) <> "port" And _
            LCase(PC.String) <> "outboard" And LCase(PC.String) <> "inboard" And _
            LCase(PC.String) <> "aft" And LCase(PC.String) <> "fore" And _
            LCase(PC.String) <> "above" And LCase(PC.String) <> "below" Then
                InputIsValid = False
                errMsg = oInput.Name & "must be Base or Offset"
            End If
            
        Case Is = INPUT_5, INPUT_8, INPUT_11, INPUT_15, INPUT_18, INPUT_20
                'these are codelisted values which may be customized by the user,
                'so for now, we won't check them

                    
        Case Else
            InputIsValid = False
            errMsg = "Unexpected input"
    End Select
        
    
  End Sub

