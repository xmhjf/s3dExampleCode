VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         MA
'   Creation Date:  Monday, Sep 24 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol is developed for CR-127634
'   Source:
'   For Part data Basis value -416: Saunders Aseptic Diaphragm Valves, Close-Coupled
'                                                           Branch Valves.
'   For Part data Basis value -417: Saunders Aseptic Diaphragm Valves,Point-of-Use
'                                      Options (Close-Coupled Branch Valve ‘U’ Bend).
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private Const E_NOOPERATORDEFINED = &H80040400 'Corresponds to E_PIPECOMP_NOOPERATORDEFINED in RefDataPipingFacelets.idl
Private Const E_NOOPERATORPARTFOUND = &H80040401 'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim iOutput     As Double
    
    Dim parFacetoFace As Double
    Dim parOpCentertoEnd As Double
    Dim parValCentertoOpCenter As Double
    Dim parPortVOffset As Double
    Dim parPorttoPortLenOff As Double
    Dim parPorttoCenterHeiOff As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle As Double
    Dim parPositionerOffset As Double
    Dim parWidth As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parOpCentertoEnd = arrayOfInputs(3)
'    parValCentertoOpCenter = arrayOfInputs(4)
'    parPortVOffset = arrayOfInputs(5)
'    parPorttoPortLenOff = arrayOfInputs(6)
'    parPorttoCenterHeiOff = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
    parHandwheelAngle = arrayOfInputs(9)
    parPositionerOffset = arrayOfInputs(10)
    parWidth = arrayOfInputs(11)
    
    Dim objCyl As Object
    Dim objBox As Object
    Dim dCyldia As Double
    
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
   
    iOutput = 0
 
' This symbol is based on the following Part data basis values that govern its geometry
'   (Part data Basis value -416): Close coupled branch valve
'   (Part data Basis value -417): Close coupled branch valve with use point
 
 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Select Case lPartDataBasis

        Case Is <= 1, 416
        
            parFacetoFace = arrayOfInputs(2)
            parOpCentertoEnd = arrayOfInputs(3)
            parValCentertoOpCenter = arrayOfInputs(4)
            parPortVOffset = arrayOfInputs(5)
            
            'Insert your code for output 1(Cylinder 1)
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -parFacetoFace / 2, 0, 0
            oEnPoint.Set parFacetoFace / 2, 0, 0
            dCyldia = pipeDiam

            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Insert your code for output 2(Box 1)
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.3 * parFacetoFace, 2.5 * parPortVOffset, _
                                                -parValCentertoOpCenter
            oEnPoint.Set 0.3 * parFacetoFace, -parValCentertoOpCenter, _
                                               0.8 * parValCentertoOpCenter
            
            Set objBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBox
            Set objBox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Insert your code for output 3(Box 2)
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.3 * parFacetoFace, 2.5 * parPortVOffset, _
                                                  0.8 * parValCentertoOpCenter
            oEnPoint.Set 0.3 * parFacetoFace, -0.4 * parPortVOffset, _
                                    (parValCentertoOpCenter + 0.6 * parOpCentertoEnd)
            
            Set objBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBox
            Set objBox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Insert your code for output 4(Cylinder 2)
            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, parPortVOffset, (parValCentertoOpCenter + 0.6 * parOpCentertoEnd)
            oEnPoint.Set 0, parPortVOffset, (parValCentertoOpCenter + parOpCentertoEnd)
            dCyldia = pipeDiam

            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Place Nozzle 1
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            Set oPlacePoint = New AutoMath.DPosition
            Set oDir = New AutoMath.DVector
            oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
            oDir.Set -1, 0, 0

            Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
            Set oPlacePoint = Nothing
            Set oDir = Nothing

            'Place Nozzle 2
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            Set oPlacePoint = New AutoMath.DPosition
            Set oDir = New AutoMath.DVector
            oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
            oDir.Set 1, 0, 0

            Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
            Set oPlacePoint = Nothing
            Set oDir = Nothing

            'Place Nozzle 3
            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            Set oPlacePoint = New AutoMath.DPosition
            Set oDir = New AutoMath.DVector
            oPlacePoint.Set 0, parPortVOffset, _
                             (parValCentertoOpCenter + parOpCentertoEnd) + sptOffset - depth
            oDir.Set 0, 0, 1

            Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
            Set oPlacePoint = Nothing
            Set oDir = Nothing

        Case 417

            parOpCentertoEnd = arrayOfInputs(3)
            parValCentertoOpCenter = arrayOfInputs(4)
            parPortVOffset = arrayOfInputs(5)
            parPorttoPortLenOff = arrayOfInputs(6)
            parPorttoCenterHeiOff = arrayOfInputs(7)
            
            'Insert your code for output 1(Box 1)
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.15 * parPorttoPortLenOff, 2.5 * parPortVOffset, _
                                             -parValCentertoOpCenter
            oEnPoint.Set 0.15 * parPorttoPortLenOff, -parValCentertoOpCenter, _
                                               0.8 * parValCentertoOpCenter

            Set objBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBox
            Set objBox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Insert your code for output 2(Box 2)
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.15 * parPorttoPortLenOff, 2.5 * parPortVOffset, _
                                                    0.8 * parValCentertoOpCenter
            oEnPoint.Set 0.15 * parPorttoPortLenOff, -0.4 * parPortVOffset, _
                                (parValCentertoOpCenter + 0.6 * parOpCentertoEnd)

            Set objBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBox
            Set objBox = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Insert your code for output 3(Cylinder 2)
            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0, parPortVOffset, (parValCentertoOpCenter + 0.6 * parOpCentertoEnd)
            oEnPoint.Set 0, parPortVOffset, (parValCentertoOpCenter + parOpCentertoEnd)
            dCyldia = pipeDiam

            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCyldia, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing

            'Insert your code for output 4
            Dim oHorLine   As IngrGeom3D.Line3d
            Dim oArc  As IngrGeom3D.Arc3d
            Dim oVerLine   As IngrGeom3D.Line3d
    
            'Horizontal line
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set -0.15 * parPorttoPortLenOff, 0, 0
            oEnPoint.Set -0.25 * parPorttoPortLenOff, 0, 0
         
            Set oHorLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
            'Arc
            Dim oCenPoint As AutoMath.DPosition
            Set oCenPoint = New AutoMath.DPosition
            oStPoint.Set -0.25 * parPorttoPortLenOff, 0, 0
            oEnPoint.Set -0.5 * parPorttoPortLenOff, 0, -0.25 * parPorttoPortLenOff
            oCenPoint.Set -0.25 * parPorttoPortLenOff, 0, -0.25 * parPorttoPortLenOff
            
            Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
            
            'Vertical line
            oStPoint.Set -0.5 * parPorttoPortLenOff, 0, -0.25 * parPorttoPortLenOff
            oEnPoint.Set -0.5 * parPorttoPortLenOff, 0, -parPorttoCenterHeiOff
            
            Set oVerLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
            Dim oSweepCurve         As IngrGeom3D.ComplexString3d
            Dim oSweepCurveCol      As Collection
            Set oSweepCurveCol = New Collection
            oSweepCurveCol.Add oHorLine
            oSweepCurveCol.Add oArc
            oSweepCurveCol.Add oVerLine
    
            Dim oStartBC   As AutoMath.DPosition
            Set oStartBC = New AutoMath.DPosition
            oStartBC.Set -0.15 * parPorttoPortLenOff, 0, 0
    
            Set oSweepCurve = PlaceTrCString(oStartBC, oSweepCurveCol)
            Dim oDirProj As AutoMath.DVector
            Set oDirProj = New AutoMath.DVector
    
            Dim oCircle As IngrGeom3D.Circle3d
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            -0.15 * parPorttoPortLenOff, 0, 0, _
                                -1, 0, 0, pipeDiam / 2)
    
            Dim Surfset   As IngrGeom3D.IJElements
            Dim stnorm() As Double
            Dim ednorm() As Double
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
                             oSweepCurve, oCircle, CircularCorner, 0, stnorm, ednorm, False)
          
            Dim ObjSurface As Object
            For Each ObjSurface In Surfset
                iOutput = iOutput + 1
                m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSurface
            Next ObjSurface
    
            'Insert your code for output 5
            'Horizontal line
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            oStPoint.Set 0.15 * parPorttoPortLenOff, 0, 0
            oEnPoint.Set 0.25 * parPorttoPortLenOff, 0, 0
         
            Set oHorLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
            'Arc
            Set oCenPoint = New AutoMath.DPosition
            oStPoint.Set 0.25 * parPorttoPortLenOff, 0, 0
            oEnPoint.Set 0.5 * parPorttoPortLenOff, 0, -0.25 * parPorttoPortLenOff
            oCenPoint.Set 0.25 * parPorttoPortLenOff, 0, -0.25 * parPorttoPortLenOff
            
            Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
            
            'Vertical line
            oStPoint.Set 0.5 * parPorttoPortLenOff, 0, -0.25 * parPorttoPortLenOff
            oEnPoint.Set 0.5 * parPorttoPortLenOff, 0, -parPorttoCenterHeiOff
            
            Set oVerLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
    
            Set oSweepCurveCol = New Collection
            oSweepCurveCol.Add oHorLine
            oSweepCurveCol.Add oArc
            oSweepCurveCol.Add oVerLine
    
            Set oStartBC = New AutoMath.DPosition
            oStartBC.Set 0.15 * parPorttoPortLenOff, 0, 0
    
            Set oSweepCurve = PlaceTrCString(oStartBC, oSweepCurveCol)
            Set oDirProj = New AutoMath.DVector
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            0.15 * parPorttoPortLenOff, 0, 0, _
                                1, 0, 0, pipeDiam / 2)
    
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
                             oSweepCurve, oCircle, CircularCorner, 0, stnorm, ednorm, False)
    
            For Each ObjSurface In Surfset
                m_OutputColl.AddOutput "Surface", ObjSurface
            Next ObjSurface
    
            Set oHorLine = Nothing
            Set oArc = Nothing
            Set oVerLine = Nothing
            Set oSweepCurve = Nothing
            Set oSweepCurveCol = Nothing
            Set oStartBC = Nothing
            Set oDirProj = Nothing
            Set oCircle = Nothing
            Set Surfset = Nothing
            Set oGeomFactory = Nothing
            Set oCenPoint = Nothing

            'Place Nozzle 1
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            Set oPlacePoint = New AutoMath.DPosition
            Set oDir = New AutoMath.DVector
            oPlacePoint.Set -0.5 * parPorttoPortLenOff, 0, -parPorttoCenterHeiOff - sptOffset + depth
            oDir.Set 0, 0, -1
            
            Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
            Set oPlacePoint = Nothing
            Set oDir = Nothing

            'Place Nozzle 2
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            Set oPlacePoint = New AutoMath.DPosition
            Set oDir = New AutoMath.DVector
            oPlacePoint.Set 0.5 * parPorttoPortLenOff, 0, -parPorttoCenterHeiOff - sptOffset + depth
            oDir.Set 0, 0, -1
            
            Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
            Set oPlacePoint = Nothing
            Set oDir = Nothing

            'Place Nozzle 3
            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
            Set oPlacePoint = New AutoMath.DPosition
            Set oDir = New AutoMath.DVector
            oPlacePoint.Set 0, parPortVOffset, _
                                parValCentertoOpCenter + parOpCentertoEnd + sptOffset - depth
            oDir.Set 0, 0, 1
            
            Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
            Set oPlacePoint = Nothing
            Set oDir = Nothing

        Case Else
            GoTo ErrorLabel:
    End Select
  
' Insert your code for output 17 (Valve Operator)
    Dim oSymbolHelper As IJSymbolGeometryHelper
    Set oSymbolHelper = New SP3DSymbolHelper.SymbolServices
    oSymbolHelper.OutputCollection = m_OutputColl

    On Error Resume Next
    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector
    Set oDirX = New DVector
    Set oDirY = New DVector
    Set oDirZ = New DVector

    oDirX.Set Sin(parHandwheelAngle), 0, -Cos(parHandwheelAngle)
    oDirY.Set 0, 1, 0
    oDirZ.Set Cos(parHandwheelAngle), Sin(parHandwheelAngle), 0

    On Error GoTo ErrorLabel
    Dim oOperatorPart As IJDPart
    Dim oOperatorOcc   As IJPartOcc
    If Not oPipeComponent Is Nothing Then
'       Note: oPipeComponent.GetValveOperatorPart method call may fail in following two cases:
'        1. If valve operator catalog part number is not defined (to check for E_NOOPERATORDEFINED) or
'        2. If valve operator part is not found (to check for E_NOOPERATORPARTFOUND)
'        symbol should raise error for case 2.
        On Error GoTo HandleOperatorCodeLabel
        Set oOperatorPart = oPipeComponent.GetValveOperatorPart
        If Not oOperatorPart Is Nothing Then
            Dim OpOrigin As IJDPosition
            Set OpOrigin = New DPosition
            OpOrigin.Set 0, 0, parValCentertoOpCenter
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
        End If
    End If
HandleOperatorCodeLabel:
'   The symbol code execution should not raise error for case 1 i.e. If Err.Number = E_NOOPERATORDEFINED
        If Err.Number = E_NOOPERATORPARTFOUND Then    'Case 2
        Err.Raise Err.Number, MODULE & METHOD, "No operator part found."
    End If

    '' Update Occurrence Attributes of Operator.
    Dim oIJDInputsArg As IJDInputsArg
    Dim oIJDEditJDArgument As IJDEditJDArgument

    Set oIJDEditJDArgument = GetIJDEditJDArgument(oOperatorOcc)

    UpdateOperatorAttributes oOperatorOcc, "IJUAInstrumentPositioner", "PositionerOffset", 2.5 * parPortVOffset, oIJDEditJDArgument

    If lPartDataBasis <= 1 Or lPartDataBasis = 416 Then
    UpdateOperatorAttributes oOperatorOcc, "IJUAWidth", "Width", 0.6 * parFacetoFace, oIJDEditJDArgument
    ElseIf lPartDataBasis = 417 Then
    UpdateOperatorAttributes oOperatorOcc, "IJUAWidth", "Width", 0.3 * parPorttoPortLenOff, oIJDEditJDArgument
    End If
    
    Set oIJDInputsArg = oOperatorOcc
    oIJDInputsArg.Update

    Set oSymbolHelper = Nothing
    Set oOperatorPart = Nothing
    Set oPipeComponent = Nothing
    Set oOperatorOcc = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

Private Sub UpdateOperatorAttributes(oOccurrence As Object, strInterfaceName As String, strAttribNameOnInterface As String, _
                            vInputParameterValue As Variant, oIJDEditJDArgument As IJDEditJDArgument)

    Const METHOD = "UpdateOperatorAttributes"
    On Error GoTo ErrorLabel

    Dim oIJDAttributes As IJDAttributes
    Dim oIJDAttr As IJDAttribute
    Dim oAttribCol As IJDAttributesCol
'    Dim vNewValue As Variant
    Dim lInputIndex As Long

    Dim oSymbol As IJDSymbol
    Dim oInputs As IJDInputs
    Dim oInput As IJDInput
    Dim lCount As Long
    Dim lCount1 As Long

    ''To find the index number of the Attibute to be updated in the operator code.
    Set oSymbol = oOccurrence
    Set oInputs = oSymbol.IJDSymbolDefinition(1).IJDInputs
    lCount = oInputs.Count
    For lCount1 = 1 To lCount
        Set oInput = oInputs.Item(lCount1)
        If StrComp(oInput.name, strAttribNameOnInterface, vbTextCompare) = 0 Then
            lInputIndex = oInput.index
            Exit For
        End If
    Next lCount1

    ''If lInputIndex  is ZERO then operator doesnt have the attribute to be updated as it's InputParameter,
    ''so that no need for futher excecution.
    If lInputIndex = 0 Then Exit Sub

    ' Set value of the user attribute
    Set oIJDAttributes = oOccurrence
    Set oAttribCol = oIJDAttributes.CollectionOfAttributes(strInterfaceName)
    Set oIJDAttr = oAttribCol.Item(strAttribNameOnInterface)
'    vNewValue = dInputParameter
    oIJDAttr.Value = vInputParameterValue

    ' Provide this value as input argument to the operator symbol
    UpdateSymbolArgument CDbl(vInputParameterValue), lInputIndex, oIJDEditJDArgument

    Exit Sub

ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD

End Sub

Private Sub UpdateSymbolArgument(dblValue As Double, lOperatorAttributeArgumentIndex As Long, _
                    oIJDEditJDArgument As IJDEditJDArgument)

    Const METHOD = "UpdateSymbolArgument"
    On Error GoTo ErrorLabel

    Dim oPC As IJDParameterContent
    Dim oArgument As IJDArgument

    Set oPC = New DParameterContent
    oPC.Type = igValue
    oPC.UomType = 0
    oPC.uomValue = dblValue

    Set oArgument = New DArgument
    oArgument.index = lOperatorAttributeArgumentIndex
    oArgument.Entity = oPC

    oIJDEditJDArgument.SetArg oArgument

    Set oArgument = Nothing
    Set oPC = Nothing

    Exit Sub

ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD

End Sub

Private Function GetIJDEditJDArgument(oOperatorOcc As IJPartOcc) As IJDEditJDArgument

    Const METHOD = "GetIJDEditJDArgument"
    On Error GoTo ErrorLabel

    Dim oIJDEditJDArgument As IJDEditJDArgument
    Dim oIEnumJDArgument As IEnumJDArgument
    Dim oSymbol As IJDSymbol
    Dim oIJDValuesArg As IJDValuesArg

    Set oSymbol = oOperatorOcc
    Set oIJDValuesArg = oSymbol.IJDValuesArg
    Set oIEnumJDArgument = oIJDValuesArg.GetValues(igINPUT_ARGUMENTS_SET)
    Set oIJDEditJDArgument = oIEnumJDArgument
    Set GetIJDEditJDArgument = oIJDEditJDArgument

    Exit Function

ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD

End Function



