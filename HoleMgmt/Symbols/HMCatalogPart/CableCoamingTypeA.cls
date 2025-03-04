VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CableCoamingTypeA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'Copyright (C) 1998, Intergraph Corporation. All rights reserved.
'
'Project: Hole Management - HoleMgmt\Middle\Symbols\HMCatalogPart
'
'File: CableCoamingTypeA.cls
'
'Abstract: implementation of CatalogPart symbol for CableCoamingTypeA
'
'Author: sypark@ship.samsung.co.kr
'
'Revision:
'     01/28/02 - sypark@ship.samsung.co.kr - initial release
'
'Note: ReportAndRaiseUnanticipatedError method is located in
'      CommonApp/Client/Bin/GSCADUtilities.dll and may have to
'      be replaced
'*******************************************************************

Option Explicit

Private Const MODULE = "HMCatalogPart.CableCoamingTypeA(CableCoamingTypeA.cls)"

    Dim m_outputColl As IJDOutputCollection

Private m_oErrors As New IMSErrorLog.JServerErrors

Implements IJDUserSymbolServices

Private Const E_FAIL = -2147467259
Private Const PI = 3.14159265

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: EditOccurence
'
'Interface: IJDUserSymbolServices
'
'Abstract: performs the edition of the passed symbol occurrence
'********************************************************************
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
    
    Const METHOD = "IJDUserSymbolServices_EditOccurence"
    On Error GoTo ErrorHandler

    IJDUserSymbolServices_EditOccurence = False
    
    Exit Function

ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & METHOD, Err.Description
End Function

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: GetDefinitionName
'
'Interface: IJDUserSymbolServices
'
'Abstract: get the symbol definition name (progID)
'********************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String

    Const METHOD = "IJDUserSymbolServices_GetDefinitionName"
    On Error GoTo ErrorHandler
    
      IJDUserSymbolServices_GetDefinitionName = "HMCatalogPart.CableCoamingTypeA"

    Exit Function
    
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & METHOD, Err.Description
End Function
'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: InstanciateDefinition
'
'Interface: IJDUserSymbolServices
'
'Abstract: instantiates and returns an initialized symbol definition
'********************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
   
  Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
  On Error GoTo ErrorHandler
  
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(Definition, ActiveConnection)
  Set oSymbolFactory = Nothing
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = "HMCatalogPart.CableCoamingTypeA"
  oSymbolDefinition.CodeBase = CodeBase

  ' Give a unique name to the symbol definition
  oSymbolDefinition.Name = oSymbolDefinition.ProgId
  
  'returned symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  Set oSymbolDefinition = Nothing
  
  Exit Function

ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & METHOD, Err.Description
End Function

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: InitializeSymbolDefinition
'
'Interface: IJDUserSymbolServices
'
'Abstract: constructs the symbol definition by (re)defining the inputs, outputs,
'          representations, and representation
'********************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

  Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
  On Error GoTo ErrorHandler
  
  ' Remove all previous Symbol Definition information
  pSymbolDefinition.IJDInputs.RemoveAllInput
  pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
  pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations

'==============================================================================
' Sleeve Definition
' Inputs:
'           1. "Nothing"  ( Catalog Part )
'           2. "CWidth"     0.280
'           3. "CDepth"     0.120
'           4. "CLength     0.200
'           5. "CThickness" 0.009
'           6. "Offsetfromwall"( Offset distance  ) 0.100
'           7. "Horientation"( Horientation ) 0.0
' The representations "Physical" is set to the definition
' Physical :    (RepresentationId = 1)
'
'==============================================================================

'======================================================================
'DEFINE INPUTS FOR ALL REPRESENTATIONS
'======================================================================
  
  'Set the input to the definition
  Dim InputsIf As IMSSymbolEntities.IJDInputs
  Set InputsIf = pSymbolDefinition
  
  Dim oSymbolCache As New CustomCache
  oSymbolCache.SetupCustomCache pSymbolDefinition
  
  Dim Inputs(1 To 6) As IMSSymbolEntities.IJDInput

  Dim PC As IMSSymbolEntities.DParameterContent
  Set PC = New IMSSymbolEntities.DParameterContent 'not persistent PC

  PC.Type = igValue

  Dim Index As Integer

  Set Inputs(1) = New IMSSymbolEntities.DInput
  Inputs(1).Name = "CWidth"
  Inputs(1).Description = "Width of the Round coaming"
  Inputs(1).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.28
  Inputs(1).DefaultParameterValue = PC

  Set Inputs(2) = New IMSSymbolEntities.DInput
  Inputs(2).Name = "CDepth"
  Inputs(2).Description = "Depth of the Round Coaming"
  Inputs(2).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.12
  Inputs(2).DefaultParameterValue = PC

 
  Set Inputs(3) = New IMSSymbolEntities.DInput
  Inputs(3).Name = "CLength"
  Inputs(3).Description = "Length of the Round Coaming"
  Inputs(3).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.2
  Inputs(3).DefaultParameterValue = PC
  
   Set Inputs(4) = New IMSSymbolEntities.DInput
  Inputs(4).Name = "CThickness"
  Inputs(4).Description = "Thickness of the Round Coaming"
  Inputs(4).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.009
  Inputs(4).DefaultParameterValue = PC

  Set Inputs(5) = New IMSSymbolEntities.DInput
  Inputs(5).Name = "OffsetFromWall"
  Inputs(5).Description = "Offset from wall"
  Inputs(5).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0.1
  Inputs(5).DefaultParameterValue = PC

  Set Inputs(6) = New IMSSymbolEntities.DInput
  Inputs(6).Name = "HOrientation"
  Inputs(6).Description = "orientation"
  Inputs(6).Properties = igINPUT_IS_A_PARAMETER
  PC.UomValue = 0
  Inputs(6).DefaultParameterValue = PC
  
  For Index = 1 To 6
    InputsIf.SetInput Inputs(Index), Index + 1
    Set Inputs(Index) = Nothing
  Next


'===================================================================
'DEFINE OUTPUTS FOR Physical REPRESENTATION
'===================================================================
  Dim O1 As IMSSymbolEntities.IJDOutput
  Set O1 = New IMSSymbolEntities.DOutput

  O1.Name = "UpperBody"
  O1.Description = "Upper body"
  O1.Properties = 0
  
  Dim O2 As IMSSymbolEntities.IJDOutput
  Set O2 = New IMSSymbolEntities.DOutput

  O2.Name = "BottomBody"
  O2.Description = "Upper body"
  O2.Properties = 0
  
  Dim O3 As IMSSymbolEntities.IJDOutput
  Set O3 = New IMSSymbolEntities.DOutput

  O3.Name = "RightCylinder"
  O3.Description = "Right Cylinder"
  O3.Properties = 0
  
  Dim O4 As IMSSymbolEntities.IJDOutput
  Set O4 = New IMSSymbolEntities.DOutput

  O4.Name = "LeftCylinder"
  O4.Description = "Left Cylinder"
  O4.Properties = 0
  
'''
  Dim ConnectPort1 As IMSSymbolEntities.IJDOutput
  Set ConnectPort1 = New IMSSymbolEntities.DOutput
  ConnectPort1.Name = "ConnectPort1"
  ConnectPort1.Description = "Connect Port one"
  ConnectPort1.Properties = 0

'===================================================================
'DEFINE THE "Physical" REPRESENTATION
'===================================================================
  Dim rep1 As IMSSymbolEntities.IJDRepresentation
  Set rep1 = New IMSSymbolEntities.DRepresentation

  rep1.Name = "Physical"
  rep1.Description = "Physical Representation of Sleeve"
  'Set the repID to SimplePhysical. See GSCADSymbolServices library to see
  'different repIDs available.
  rep1.RepresentationId = SimplePhysical
  rep1.Properties = igREPRESENTATION_ISVBFUNCTION

  Dim oRep1Outputs As IMSSymbolEntities.IJDOutputs
  Set oRep1Outputs = rep1

  'Set the outputs
  oRep1Outputs.SetOutput O1
  Set O1 = Nothing
  
  'Set the outputs
  oRep1Outputs.SetOutput O2
  Set O2 = Nothing

    'Set the outputs
  oRep1Outputs.SetOutput O3
  Set O3 = Nothing
  
    'Set the outputs
  oRep1Outputs.SetOutput O4
  Set O4 = Nothing
 
  oRep1Outputs.SetOutput ConnectPort1
  Set ConnectPort1 = Nothing

  'Set the symbolic representation to definition
  Dim RepsIf As IMSSymbolEntities.IJDRepresentations
  Set RepsIf = pSymbolDefinition
  RepsIf.SetRepresentation rep1

  Set rep1 = Nothing
  Set RepsIf = Nothing
  Set oRep1Outputs = Nothing


  'Define evaluation for Physical representation
  Dim PhysicalRepEval As IJDRepresentationEvaluation
  Set PhysicalRepEval = New DRepresentationEvaluation
  PhysicalRepEval.Name = "Physical"
  PhysicalRepEval.Description = "Physical representation"
  PhysicalRepEval.Properties = igREPRESENTATION_HIDDEN
  PhysicalRepEval.Type = igREPRESENTATION_VBFUNCTION
  PhysicalRepEval.ProgId = "HMCatalogPart.CableCoamingTypeA"
    
  'Set the evaluations for the Symbolic and Physical representations on the definition
  Dim RepEvalsIf As IMSSymbolEntities.IJDRepresentationEvaluations
  Set RepEvalsIf = pSymbolDefinition
  RepEvalsIf.AddRepresentationEvaluation PhysicalRepEval
  Set PhysicalRepEval = Nothing
  Set RepEvalsIf = Nothing
    
'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================

    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE
 

'===========================================================================
'THE FOLLOWING STATEMENT SPECIFIES THAT THERE ARE NO INPUTS TO THE SYMBOL
'WHICH ARE GRAPHIC ENTITIES.
'===========================================================================


  Exit Sub

ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & METHOD, Err.Description
End Sub

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: InvokeRepresentation
'
'Interface: IJDUserSymbolServices
'
'Abstract: performs the evaluation function associated with the representation
'          with the set of passed inputs as arguments. the function has to compute
'          the symbol graphic outputs that will be attached to the specified
'          Output Collection object
'
'Outline:
'    * gather inputs
'********************************************************************

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation"
    On Error GoTo ErrorHandler
    
    Set m_outputColl = outputcoll
    If StrComp(repName, "Physical") = 0 Then
        PhysicalRep arrayOfInputs
    End If
    
    Exit Sub
    
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & METHOD, Err.Description
End Sub

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: PhysicalRep
'
'Interface:
'
'Abstract: CREATION OF PHYSICAL REPRESENTATION OF Cable coaming
'
'Outline:
'
'********************************************************************
Private Sub PhysicalRep(ByRef arrayOfInputs())
        
    Const METHOD = "PhysicalRep"
    On Error GoTo ErrorHandler
    
    Dim oWidth As Double, oDepth As Double, oLength As Double, oThickness As Double
    Dim Offsetfromwall As Double
    Dim OOrientation As Double
    Dim oAngle As Double
    
    Dim geomFactory As New IngrGeom3D.GeometryFactory

    oWidth = arrayOfInputs(2)
    oDepth = arrayOfInputs(3)
    oLength = arrayOfInputs(4)
    oThickness = arrayOfInputs(5)
    Offsetfromwall = arrayOfInputs(6)
    OOrientation = arrayOfInputs(7)
    oAngle = 0.06
    
    'If the orientation is not 0, which means that 90 degree, Width and Depth should be changed.
    If OOrientation <> 0 Then
        Dim oTempValue As Double
        oTempValue = oWidth
        oWidth = oDepth
        oDepth = oTempValue
    End If
    
    Dim SP1_X As Double, SP1_Y As Double, SP1_Z As Double
    Dim SP2_X As Double, SP2_Y As Double, SP2_Z As Double
    Dim SP3_X As Double, SP3_Y As Double, SP3_Z As Double
    Dim SP4_X As Double, SP4_Y As Double, SP4_Z As Double
    
    Dim oLine As IngrGeom3D.Line3d
    Dim Elems As IMSCoreCollections.IJElements
    Set Elems = New JObjectCollection
    
    
    '
    '               (1)=================      +
    '             +                            +
    '            +                              +
    '             +                            +
    '               +    =================(2)+
    '
    ' 1. Find Out (1) then, Create half cylinder and box
    ' 2. Find Out (2) then, Create half cylinder and box
    ' 3. Representation of two boxs and two half cylinders
    
    
    SP1_X = -oWidth / 2 + oAngle
    SP1_Y = oDepth / 2
    SP1_Z = Offsetfromwall
    
    SP2_X = -oWidth / 2 + oAngle
    SP2_Y = oDepth / 2 - oThickness
    SP2_Z = Offsetfromwall
    
    SP3_X = -oWidth / 2 + oAngle
    SP3_Y = oDepth / 2 - oThickness
    SP3_Z = SP2_Z - oLength

    SP4_X = -oWidth / 2 + oAngle
    SP4_Y = oDepth / 2
    SP4_Z = SP2_Z - oLength
    
    Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, SP1_X, SP1_Y, SP1_Z, SP2_X, SP2_Y, SP2_Z)
    Elems.Add oLine

    Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, SP2_X, SP2_Y, SP2_Z, SP3_X, SP3_Y, SP3_Z)
    Elems.Add oLine

    Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, SP3_X, SP3_Y, SP3_Z, SP4_X, SP4_Y, SP4_Z)
    Elems.Add oLine

    Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, SP4_X, SP4_Y, SP4_Z, SP1_X, SP1_Y, SP1_Z)
    Elems.Add oLine
    
    Dim oComplexString As IngrGeom3D.ComplexString3d
    Set oComplexString = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, Elems)
                
    '============================================================================
    'Create Upper body
    '============================================================================
    Dim Projection3d As IngrGeom3D.Projection3d
    Set Projection3d = geomFactory.Projections3d.CreateByCurve( _
                                m_outputColl.ResourceManager, _
                                oComplexString, _
                                1, 0, 0, _
                                oWidth - 2 * oAngle, True)
    m_outputColl.AddOutput "UpperBody", Projection3d
    Set Projection3d = Nothing

    '============================================================================
    'Create Left Cylinder
    '============================================================================
    Dim Revolution As IngrGeom3D.Revolution3d
    Set Revolution = geomFactory.Revolutions3d.CreateByCurve( _
                                m_outputColl.ResourceManager, _
                                oComplexString, _
                                0, 0, 1, _
                                -oWidth / 2 + oAngle, 0, Offsetfromwall - oLength / 2, _
                                4 * Atn(1), True)
    m_outputColl.AddOutput "LeftCylinder", Revolution
    Set oComplexString = Nothing
    Set Revolution = Nothing

    
    'Set element collection clear
    Elems.Clear
    
    SP1_X = oWidth / 2 - oAngle
    SP1_Y = -oDepth / 2
    SP1_Z = Offsetfromwall
    
    SP2_X = oWidth / 2 - oAngle
    SP2_Y = -oDepth / 2 + oThickness
    SP2_Z = Offsetfromwall
    
    SP3_X = oWidth / 2 - oAngle
    SP3_Y = -oDepth / 2 + oThickness
    SP3_Z = SP2_Z - oLength

    SP4_X = oWidth / 2 - oAngle
    SP4_Y = -oDepth / 2
    SP4_Z = SP2_Z - oLength
    
    Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, SP1_X, SP1_Y, SP1_Z, SP2_X, SP2_Y, SP2_Z)
    Elems.Add oLine

    Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, SP2_X, SP2_Y, SP2_Z, SP3_X, SP3_Y, SP3_Z)
    Elems.Add oLine

    Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, SP3_X, SP3_Y, SP3_Z, SP4_X, SP4_Y, SP4_Z)
    Elems.Add oLine

    Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, SP4_X, SP4_Y, SP4_Z, SP1_X, SP1_Y, SP1_Z)
    Elems.Add oLine
    
    Set oComplexString = geomFactory.ComplexStrings3d.CreateByCurves(Nothing, Elems)
                
    Set oLine = Nothing
    Set Elems = Nothing

    '============================================================================
    'Create Bottom body
    '============================================================================
    Set Projection3d = geomFactory.Projections3d.CreateByCurve( _
                                m_outputColl.ResourceManager, _
                                oComplexString, _
                                -1, 0, 0, _
                                oWidth - 2 * oAngle, True)
    m_outputColl.AddOutput "BottomBody", Projection3d
    Set Projection3d = Nothing

    '============================================================================
    'Create Create Right Cylinder
    '============================================================================
    Set Revolution = geomFactory.Revolutions3d.CreateByCurve( _
                                m_outputColl.ResourceManager, _
                                oComplexString, _
                                0, 0, 1, _
                                oWidth / 2 - oAngle, 0, Offsetfromwall - oLength / 2, _
                                4 * Atn(1), True)
    m_outputColl.AddOutput "RightCylinder", Revolution
    Set oComplexString = Nothing
    Set Revolution = Nothing


    '============================================================================
    'BUILD CONNECT PORT ONE
    '============================================================================

    Dim oPoint As IngrGeom3D.Point3d
    Set oPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, 0, _
                                                    0, Offsetfromwall - oLength / 2)

    m_outputColl.AddOutput "ConnectPort1", oPoint
    Set oPoint = Nothing
    
    
    Exit Sub
        
ErrorHandler:
'    ReportAndRaiseUnanticipatedError MODULE, METHOD
    m_oErrors.Add Err.Number, MODULE & " - " & METHOD, Err.Description
End Sub
 