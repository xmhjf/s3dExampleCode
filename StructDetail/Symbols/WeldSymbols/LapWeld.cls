VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LapWeld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'/*******************************************************************
'Copyright (C) 1998, Intergraph Corporation.  All rights reserved.
'
'
'Project: S:\StructDetail\Middle\Symbols\WeldSymbols\WeldSymbols.vbp
'
'File: S:\StructDetail\Middle\Symbols\WeldSymbols\TeeWeld.vbp
'
'Revision:
'     02/07/01 GDreybus.
'
'Description:
'     Create trace curve.
'
'*******************************************************************/

Option Explicit
Private Const MODULE = "S:\StructDetail\Middle\Symbols\WeldSymbols\LapWeld.cls"

Private Const m_progID = "WeldSymbols.LapWeld"

Private Const REPRESENTATION_1 = "LapWeld"

Private Const OUTPUT_1 = "Curve_1"

Private Const INPUT_1 = "MoldedFillet"
Private Const INPUT_2 = "AntiMoldedFillet"
Private Const INPUT_3 = "FilletMeasureMethod"
Private Const INPUT_4 = "Category"

Private m_nInputCount As Long

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
    Dim Coord() As IMSSymbolEntities.IJDInput
    Dim arrayWeldingSymInputInfo() As WELDING_SYMBOL_INPUT_INFO
    Dim nWeldingSymInputCount As Long
    
    ' Returned arrayWeldingSymInputInfo is 1 based
    GetWeldingSymbolInputsDescription _
                            arrayWeldingSymInputInfo, _
                            WELD_TYPE_LAP, _
                            nWeldingSymInputCount
    m_nInputCount = 4 + nWeldingSymInputCount
    ReDim Coord(1 To m_nInputCount)
    
    'Add the custom validate method
    Dim ValidateMethodCookie As Long
    Dim oUserMethods As IJDUserMethods
    Set oUserMethods = pSymbolDefinition
    ValidateMethodCookie = oUserMethods.GetMethodCookie("CheckParameterValues", imsCOOKIE_ID_USS_LIB)
    Dim oCustomMethodIntf As IJDInputStdCustomMethod
    
    Dim ii As Integer
    For ii = 1 To m_nInputCount
        Set Coord(ii) = New IMSSymbolEntities.DInput
        Set oCustomMethodIntf = Coord(ii).IJDInputStdCustomMethod
        oCustomMethodIntf.SetCMCheck imsCOOKIE_ID_USS_LIB, ValidateMethodCookie
    Next
    Set oCustomMethodIntf = Nothing

    ' The two first arguments are in fact the ports of the connection
    SetWeldCommonInputs pSymbolDefinition

    ' Create a defaultValue
    Dim PC As IMSSymbolEntities.IJDParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC
    PC.Type = igValue
    PC.UomValue = 1#
    
    Coord(1).Name = INPUT_1
    Coord(1).Description = INPUT_1
    Coord(1).Properties = igINPUT_IS_A_PARAMETER
    PC.UomType = UNIT_DISTANCE
    Coord(1).DefaultParameterValue = PC

    Coord(2).Name = INPUT_2
    Coord(2).Description = INPUT_2
    Coord(2).Properties = igINPUT_IS_A_PARAMETER
    PC.UomType = UNIT_DISTANCE
    Coord(2).DefaultParameterValue = PC
    
    Coord(3).Name = INPUT_3
    Coord(3).Description = INPUT_3
    Coord(3).Properties = igINPUT_IS_A_PARAMETER
    PC.UomType = 0
    PC.UomValue = 65536
    Coord(3).DefaultParameterValue = PC

    Coord(4).Name = INPUT_4
    Coord(4).Description = INPUT_4
    Coord(4).Properties = igINPUT_IS_A_PARAMETER
    PC.UomType = 0
    PC.UomValue = 65536
    Coord(4).DefaultParameterValue = PC
    
    ' Add inputs from IJWeldingSymbol
    Dim nInputIndex As Long
    Dim nWeldingSymInputIndex As Long
    
    nInputIndex = 4
    If nWeldingSymInputCount > 0 Then
       For nWeldingSymInputIndex = 1 To nWeldingSymInputCount
          Dim stWSID As WELDING_SYMBOL_INPUT_INFO
          
          nInputIndex = nInputIndex + 1
          stWSID = arrayWeldingSymInputInfo(nWeldingSymInputIndex)
          PC.Reset
          PC.Type = stWSID.eType
          Coord(nInputIndex).Name = stWSID.sInputName
          Coord(nInputIndex).Description = stWSID.sInputName
          Coord(nInputIndex).Properties = igINPUT_IS_A_PARAMETER
          If stWSID.eType = igString Then
              
          ElseIf stWSID.eType = igValue Then
             PC.UomValue = stWSID.dUomValue
             PC.UomType = stWSID.nUomType
          End If
          Coord(nInputIndex).DefaultParameterValue = PC
       Next
    End If

     ' Set the input to the definition
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition

    ' Offset the inputs to account for the two input ports
    For ii = 1 To m_nInputCount
      InputsIf.SetInput Coord(ii), ii + igNonRefPartIndex
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
    oRepEval.Name = "LapWeld"
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
    
    ' Rule based IJWeldingSymbol based attributes were stored on parameter rule interface
    ' Copy them to IJWeldingSymbol if they are not overridden so that they can be retrieved via IJWeldingSymbol
    CopyAttributesToIJWeldingSymbol _
                             WELD_TYPE_LAP, _
                             SymbolOcc

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
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
        Case Is = INPUT_1, INPUT_2
            
            If PC.UomValue < 0 Then
                InputIsValid = False
                errMsg = oInput.Name & " cannot be negative"
            End If
        
        Case Is = INPUT_4
'            If LCase(PC.String) <> "Normal" Then
'                InputIsValid = False
'                errMsg = INPUT_4 & " must be Normal or "
'            End If
        
        Case Is = INPUT_3
'            If LCase(PC.String) <> "leg" Or LCase(PC.String) <> "throat" Then
'                InputIsValid = False
'                errMsg = INPUT_3 & " must be Leg or Throat"
'            End If
                    
        ' Following are IJWeldingSymbol inputs
        Case Is = PRIMARY_SIDE_SYMBOL, SECONDARY_SIDE_SYMBOL, _
                  PRIMARY_SIDE_GROOVE, SECONDARY_SIDE_GROOVE, _
                  PRIMARY_SIDE_SUPPLEMENTARY_SYMBOL, SECONDARY_SIDE_SUPPLEMENTARY_SYMBOL, _
                  PRIMARY_SIDE_CONTOUR, SECONDARY_SIDE_CONTOUR, _
                  PRIMARY_SIDE_FINISH_METHOD, SECONDARY_SIDE_FINISH_METHOD
            ' Integer data
            If PC.Type <> igValue Then
                InputIsValid = False
                errMsg = oInput.Name & " is not a value"
            End If

        Case Is = PRIMARY_SIDE_GROOVE_SIZE, SECONDARY_SIDE_GROOVE_SIZE, _
                  PRIMARY_SIDE_ACTUAL_THROAT_THICKNESS, SECONDARY_SIDE_ACTUAL_THROAT_THICKNESS, _
                  PRIMARY_SIDE_NOMINAL_THROAT_THICKNESS, SECONDARY_SIDE_NOMINAL_THROAT_THICKNESS, _
                  PRIMARY_SIDE_LENGTH, SECONDARY_SIDE_LENGTH, _
                  PRIMARY_SIDE_PITCH, SECONDARY_SIDE_PITCH, _
                  PRIMARY_SIDE_ROOT_OPENING, SECONDARY_SIDE_ROOT_OPENING, _
                  PRIMARY_SIDE_GROOVE_ANGLE, SECONDARY_SIDE_GROOVE_ANGLE, _
                  PRIMARY_SIDE_ACTUAL_LEG_LENGTH, SECONDARY_SIDE_ACTUAL_LEG_LENGTH
            ' Double data
            If PC.Type <> igValue Then
                InputIsValid = False
                errMsg = oInput.Name & " is not a value"
            End If
            
        Case Is = FIELD_WELD, ALL_AROUND, TAIL_NOTE_IS_REFERENCE
            ' Boolean data
            If PC.UomValue < 0 Or PC.UomValue > 2 Then
                InputIsValid = False
                errMsg = oInput.Name & "must be a boolean value"
            End If
            
        Case Is = TAIL_NOTES
        
        Case Else
            InputIsValid = False
            errMsg = "Unexpected input"
    End Select
            
        
    
  End Sub


