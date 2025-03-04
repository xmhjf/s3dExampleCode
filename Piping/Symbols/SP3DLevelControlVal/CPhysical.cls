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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   Physical.cls
'   Author:         MA
'   Creation Date:  Wednesday, Jan 30 2008
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------    -----               ------------------
'  30.JAN.2008      MA   CR-135564  Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double
Private Const E_NOOPERATORDEFINED = &H80040400 'Corresponds to E_PIPECOMP_NOOPERATORDEFINED in RefDataPipingFacelets.idl
Private Const E_NOOPERATORPARTFOUND = &H80040401 'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
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
    Dim parConnectionHeight As Double
    Dim parBodyOutsideDiameter As Double
    Dim parCoverDiameter As Double
    Dim parCoverHeight As Double
    Dim parHandwheelAngle As Double
    Dim parInsulationThickness As Double
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim oAxisVec As AutoMath.DVector
    Dim oCenPoint As AutoMath.DPosition
    Dim objCyl As Object
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
'    parConnectionHeight = arrayOfInputs(3)
'    parBodyOutsideDiameter = arrayOfInputs(4)
'    parCoverDiameter = arrayOfInputs(5)
'    parCoverHeight = arrayOfInputs(6)
    parHandwheelAngle = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)

    iOutput = 0

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                            sptOffset, depth
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    ' This symbol is based on the following three Part data basis values that govern its geometry,
    '   Level control valve, cylindrical type (Default case) - 740
    '   Level control valve, globe type - 741
    '   Level control valve, Y type - 742

    Select Case lPartDataBasis
        
        Case Is <= 1, 740
            parConnectionHeight = arrayOfInputs(3)
            parBodyOutsideDiameter = arrayOfInputs(4)
            
            'Assumption: Body Outside Diameter is an optional Input.
            ' When not specified Body Outside Diameter = 1.6 x ConnectionHeight
            If CmpDblLessThanOrEqualTo(parBodyOutsideDiameter, 0) Then _
                                    parBodyOutsideDiameter = 1.6 * parConnectionHeight
                                    
            'Place Cylinder 1
            oStPoint.Set -0.5 * parFacetoFace, 0, 0
            oEnPoint.Set -0.3 * parFacetoFace, 0, 0
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
                                   
            'Place Cylinder 2
            oStPoint.Set 0.5 * parFacetoFace, 0, 0
            oEnPoint.Set 0.3 * parFacetoFace, 0, 0
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
                                    
            'Insert your code for middle portion of the body of the valve
            Dim oPoints(1 To 8) As New AutoMath.DPosition
            Dim objBsplineCurve As IngrGeom3D.BSplineCurve3d
            Dim objMiddlePortion As Object
            
            'Points specifying the B-Spline curve are approximated to produce closer geometry.
            oPoints(1).Set -0.3 * parFacetoFace, 0.5 * pipeDiam, 0
            oPoints(2).Set -0.27 * parFacetoFace, 0.43 * parBodyOutsideDiameter, 0
            oPoints(3).Set -0.25 * parFacetoFace, 0.45 * parBodyOutsideDiameter, 0
            oPoints(4).Set -0.2 * parFacetoFace, 0.5 * parBodyOutsideDiameter, 0
            oPoints(5).Set 0.2 * parFacetoFace, 0.5 * parBodyOutsideDiameter, 0
            oPoints(6).Set 0.25 * parFacetoFace, 0.45 * parBodyOutsideDiameter, 0
            oPoints(7).Set 0.27 * parFacetoFace, 0.43 * parBodyOutsideDiameter, 0
            oPoints(8).Set 0.3 * parFacetoFace, 0.5 * pipeDiam, 0

            Set objBsplineCurve = PlaceTrBspline(7, oPoints)
            
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0, 0
        
            Set objMiddlePortion = PlaceRevolution(m_OutputColl, objBsplineCurve, oAxisVec, _
                                                        oCenPoint, 2 * PI, True)
 
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objMiddlePortion
            Set objMiddlePortion = Nothing
            Set oCenPoint = Nothing
            Set oAxisVec = Nothing
            Set objBsplineCurve = Nothing
         
            'Place Box
            oStPoint.Set -0.2 * parFacetoFace, 0.45 * parBodyOutsideDiameter, _
                                                0.1 * parBodyOutsideDiameter
            oEnPoint.Set -0.1 * parFacetoFace, 0.45 * parBodyOutsideDiameter _
                            + 0.15 * parConnectionHeight, -0.1 * parBodyOutsideDiameter
            
            Dim objBox As Object
            Set objBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBox
            Set objBox = Nothing
        
            'Place Cylinder 3
            oStPoint.Set -0.15 * parFacetoFace, 0.45 * parBodyOutsideDiameter _
                                            + 0.15 * parConnectionHeight, 0
            oEnPoint.Set -0.15 * parFacetoFace, 0.5 * parBodyOutsideDiameter _
                                            + 0.2 * parConnectionHeight, 0
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                0.1 * parBodyOutsideDiameter, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
        
        Case 741
            parConnectionHeight = arrayOfInputs(3)
            parCoverDiameter = arrayOfInputs(5)
            
            'Place Cylinder 1
            oStPoint.Set -0.5 * parFacetoFace, 0, 0
            oEnPoint.Set -0.3 * parFacetoFace, 0, 0
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
                                   
            'Place Cylinder 2
            oStPoint.Set 0.5 * parFacetoFace, 0, 0
            oEnPoint.Set 0.3 * parFacetoFace, 0, 0
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)
                                                                    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            
            'Place Ellipsoid
            Dim dEllipseMajorX   As Double
            Dim dEllipseMajorY   As Double
            Dim dEllipseMajorZ   As Double
            Dim dEllipseMMRatio  As Double
            Dim objEllipticalArc As IngrGeom3D.EllipticalArc3d
            Dim objEllipsoid As Object
            
            dEllipseMMRatio = (0.3 * parConnectionHeight) / (0.4 * parFacetoFace)
            dEllipseMajorX = 0.4 * parFacetoFace
            dEllipseMajorY = 0
            dEllipseMajorZ = 0
            
            Set objEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                                    Nothing, 0, 0, 0, _
                                                    0, 0, 1, dEllipseMajorX, dEllipseMajorY, dEllipseMajorZ, _
                                                    dEllipseMMRatio, 0, PI)
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0, 0
          
            Set objEllipsoid = PlaceRevolution(m_OutputColl, objEllipticalArc, oAxisVec, _
                                                    oCenPoint, 2 * PI, True)
          
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEllipsoid
            Set objEllipsoid = Nothing
        
            Set oCenPoint = Nothing
            Set oAxisVec = Nothing
            Dim sp3dElem1 As IJDObject
            Set sp3dElem1 = objEllipticalArc
            Set objEllipticalArc = Nothing
            sp3dElem1.Remove
            Set sp3dElem1 = Nothing
            
            'Place Cylinder 3
            oStPoint.Set 0, 0.28 * parConnectionHeight, 0
            oEnPoint.Set 0, 0.4 * parConnectionHeight, 0
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                        0.4 * parCoverDiameter, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            
            'Place Cylinder 4
            oStPoint.Set 0, 0.4 * parConnectionHeight, 0
            oEnPoint.Set 0, 0.6 * parConnectionHeight, 0
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                parCoverDiameter, True)
      
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            
        Case 742
            parCoverDiameter = arrayOfInputs(5)
            parCoverHeight = arrayOfInputs(6)
 
            'Assumption: Cover Diameter is an optional Input.
            ' When not specified Cover Diameter = 0.75 x Face to Face, and
            If CmpDblLessThanOrEqualTo(parCoverDiameter, 0) Then _
                                    parCoverDiameter = 0.75 * parFacetoFace
                                            
            'These values are used in updating the valve Operator.
            parHandwheelAngle = PI / 2.4 'Assumption: the Operator is at inclined at 75Deg to x Axis.
            Dim dInclinationAngle As Double
            dInclinationAngle = parHandwheelAngle
            
            Dim dEffectiveHeightofCover As Double
            dEffectiveHeightofCover = (parCoverHeight - 0.5 * pipeDiam) - parCoverDiameter * 0.5 * Cos(dInclinationAngle)
            
            Dim dOperatorHeight As Double 'used for updating the valve Operator height.
            Dim dOperatorDiameter As Double 'used for updating the valve Operator diameter.
            dOperatorHeight = dEffectiveHeightofCover * (2 - Sin(parHandwheelAngle))  'The operator places above the cover.
            dOperatorDiameter = 0.7 * parCoverDiameter
            
            Dim dCylDia As Double
            'The Ellipsoid shoud not engulf the upper cover on which operator appears,
            'the diameter of the stem reaching upto the cover should be less than cover width,
            'the base ellipsoid (created below) should not exceed face to face dimension and
            'the section of horizontal cone that protrudes into the inclined ellipsoid is less than Pipe dia.
            dCylDia = ReturnMin(parFacetoFace / 4, pipeDiam, parCoverHeight, parCoverDiameter)
            
            'Create horizontal Cone adjacent to horizontal cylinder at port1 and protruding into inclined ellipsoid.
            oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
            oEnPoint.Set 0, 0, 0
                                        
            Dim objLeftCone As Object
            Set objLeftCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, _
                                                        pipeDiam / 2, dCylDia / 2)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLeftCone
            Set objLeftCone = Nothing
            
            'Create the inclined Ellipsoid
            Dim dHalfMajorAxisInclinedEllipArc As Double
            dHalfMajorAxisInclinedEllipArc = dCylDia * 1.1
                                           
            Dim dHalfMinorAxisInclinedEllipArc As Double
            dHalfMinorAxisInclinedEllipArc = dCylDia * 0.85
                
            Dim objInclinedEllipticalArc As Object
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set objInclinedEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                0, 0, 0, _
                0, 0, 1, _
                dHalfMinorAxisInclinedEllipArc * Cos(dInclinationAngle), dHalfMinorAxisInclinedEllipArc * Sin(dInclinationAngle), 0, _
                dHalfMajorAxisInclinedEllipArc / dHalfMinorAxisInclinedEllipArc, _
                                    0, PI)

            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set Cos(dInclinationAngle), Sin(dInclinationAngle), 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0, 0

            Dim objInclinedEllipsoid As Object
            Set objInclinedEllipsoid = PlaceRevolution(m_OutputColl, objInclinedEllipticalArc, _
                                oAxisVec, oCenPoint, 2 * PI, True)

            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInclinedEllipsoid
            Set objInclinedEllipsoid = Nothing
            Set objInclinedEllipticalArc = Nothing
            Set oAxisVec = Nothing
            Set oCenPoint = Nothing
            Set oGeomFactory = Nothing
            
            'Create horizontal Cylinder at port 2
            oStPoint.Set parFacetoFace / 2 - flangeThick, 0, 0
            oEnPoint.Set 0.4 * parFacetoFace - flangeThick, 0, 0
            
            Set objCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            pipeDiam, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCyl
            Set objCyl = Nothing
            
            'Create horizontal Cone adjacent to horizontal cylinder at port2 and protruding into inclined ellipsoid.
            oStPoint.Set 0.4 * parFacetoFace - flangeThick, 0, 0
            oEnPoint.Set 0, 0, 0
                                        
            Dim objRightCone As Object
            Set objRightCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, _
                                                    pipeDiam / 2, dCylDia / 2)
                                                                                               
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objRightCone
            Set objRightCone = Nothing
            
            'Create Cover
            oStPoint.Set dEffectiveHeightofCover / Tan(dInclinationAngle), _
                                            dEffectiveHeightofCover, 0
            oEnPoint.Set oStPoint.x - 0.15 * parCoverHeight / Tan(dInclinationAngle), _
                                    oStPoint.y - 0.15 * parCoverHeight, 0
            
            
            Dim objCover As Object
            Set objCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            parCoverDiameter, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCover
            Set objCover = Nothing
            
            'Create stem reaching upto the bottom of the cover
            oStPoint.Set 0, 0, 0
            oEnPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z     'End point of the previous cylinder
            
            Dim objStem As Object
            Set objStem = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                        dCylDia, True)
            
            'Set the output
            m_OutputColl.AddOutput "Stem", objStem
            Set objStem = Nothing
            
        Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.
        
    End Select
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    'Insert your code for Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    
    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Dim objNozzle1 As GSCADNozzleEntities.IJDNozzle
    Set objNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
    Set objNozzle1 = Nothing
    
    'Insert your code for Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                        sptOffset, depth
    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    
    Dim objNozzle2 As GSCADNozzleEntities.IJDNozzle
    Set objNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
    Set objNozzle2 = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    
'   Insert your code for Valve Actuator
'-----------------------------------------------------------------------------------------------------
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

    If lPartDataBasis = 741 Then
        oDirX.Set Cos(parHandwheelAngle), 0, Sin(parHandwheelAngle)
        oDirY.Set 0, 1, 0
        oDirZ.Set -Sin(parHandwheelAngle), 0, Cos(parHandwheelAngle)
    ElseIf lPartDataBasis = 742 Then
        oDirX.Set Sin(parHandwheelAngle), -Cos(parHandwheelAngle), 0
        oDirY.Set Cos(parHandwheelAngle), Sin(parHandwheelAngle), 0
        oDirZ.Set 0, 0, 1
    End If
    
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
            
            'For the follwoing cases the operator is at an offset from the origin of the valve.
            If lPartDataBasis = 741 Then
                OpOrigin.Set 0, 0, 0
            ElseIf lPartDataBasis = 742 Then
                OpOrigin.Set (dEffectiveHeightofCover / Tan(parHandwheelAngle)) - dEffectiveHeightofCover * Cos(parHandwheelAngle), 0, 0
            End If
            
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
        
            '   Update Occurrence Attributes of Operator.
            Dim oIJDInputsArg As IJDInputsArg
            Dim oIJDEditJDArgument As IJDEditJDArgument
        
            Set oIJDEditJDArgument = GetIJDEditJDArgument(oOperatorOcc)
        
            If lPartDataBasis = 741 Then
                UpdateOperatorAttributes oOperatorOcc, "IJUAValveOperator", "OperatorDiameter", 0.6 * parCoverDiameter, oIJDEditJDArgument
                UpdateOperatorAttributes oOperatorOcc, "IJUAValveOperator", "OperatorHeight1", 0.6 * parConnectionHeight, oIJDEditJDArgument
                UpdateOperatorAttributes oOperatorOcc, "IJUAValveOperator", "OperatorHeight", parConnectionHeight, oIJDEditJDArgument
            ElseIf lPartDataBasis = 742 Then
                UpdateOperatorAttributes oOperatorOcc, "IJUAValveOperator", "OperatorDiameter", dOperatorDiameter, oIJDEditJDArgument
                UpdateOperatorAttributes oOperatorOcc, "IJUAValveOperator", "OperatorHeight1", dOperatorHeight, oIJDEditJDArgument
            End If
            
            Set oIJDInputsArg = oOperatorOcc
            oIJDInputsArg.Update
        End If
    End If
    
HandleOperatorCodeLabel:
'   The symbol code execution should not raise error for case 1 i.e. If Err.Number = E_NOOPERATORDEFINED
    If Err.Number = E_NOOPERATORPARTFOUND Then          'Case 2
        Err.Raise Err.Number, MODULE & METHOD, "No operator part found."
    End If
   
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

Private Function ReturnMin(A#, B#, C#, Optional D# = -1) As Double
    'Returns the minimum value among four input values.
    Const METHOD = "ReturnMin"
    On Error GoTo ErrorLabel
    
    Dim dMinValue As Double
    dMinValue = A
    If CmpDblLessThan(B, dMinValue) Then dMinValue = B
    
    If CmpDblLessThan(C, dMinValue) Then dMinValue = C
    
    If Not CmpDblLessThanOrEqualTo(D, 0) Then
        If CmpDblLessThan(D, dMinValue) Then dMinValue = D
    End If
    
    ReturnMin = dMinValue
    
    Exit Function
ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD
End Function

