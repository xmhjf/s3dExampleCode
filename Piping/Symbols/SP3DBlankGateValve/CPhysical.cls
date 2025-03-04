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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          BG
'   Creation Date:  Friday, Jan 11 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle and modified valve operator direction vectors.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   02.Jan.2006     KKK       CR-91106  Updated the symbol to be based on either face-to-face
'                                           or face-to-center(Implementation of PartDataBasis Property)
'   31.Mar.2006     SymbolTeam(India)       Modified to accept default value of lPartDataBasis to be less than or equal to 1.
'   25.sep.06       KKC            TR-102869 Modified symbol code to use Face to Center or Face to Face to Dimension.
'                                            Modified symbol code to use Face to Face as a input parameter.
'   2.JUL.2007      dkl            TR-123363: Commented out the code for recieving inputs from the 'arrayOfInputs' outside the respective Part Data Basis Case block
'                                     to enable compatibility with V6.1 version.
'   13.Jul.2007     MA             CR-123497 Implemented Part data basis to use Face 1 to Center and Face 2 to Center Dimensions.
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
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim parFacetoFace As Double
    Dim parActuatorOffset As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim ObjNozzle1 As Object
    Dim ObjLeftCone As Object
    Dim ObjTriangularDisc As Object
    Dim ObjActuatorCylinder As Object
    Dim ObjRightCone As Object
    Dim ObjNozzle2 As Object
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle   As Double
    Dim parFacetoCenter As Double
    Dim iOutput     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
    parActuatorOffset = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    parHandwheelAngle = arrayOfInputs(5)
'    parFacetoCenter = arrayOfInputs(6)
'    parFace1toCenter = arrayOfInputs(7)
'    parFace2toCenter = arrayOfInputs(8)
    
    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                sptOffset, depth
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                sptOffset2, depth2
        
 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    
' This symbol is based on the following Part data basis values that govern its geometry
'   (Part data Basis value -5): Face-to-face dimension basis
'   (Part data Basis value -10): Face-to-center dimension basis
'   (Part data Basis value -13): Asymmetrical Face-to-Center dimension basis
    
    Select Case lPartDataBasis
    
    Case Is <= 1, 5
        parFacetoFace = arrayOfInputs(2)
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
    
    Case 10
        parFacetoCenter = arrayOfInputs(6)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
    
    Case 13
        parFace1toCenter = arrayOfInputs(7)
        parFace2toCenter = arrayOfInputs(8)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
 
    Case Else
        GoTo ErrorLabel:
        
    End Select
        
' Insert your code for output 2(Cone)
    Dim oBaseCenPoint As AutoMath.DPosition
    Dim oTopCenPoint As AutoMath.DPosition
    Set oBaseCenPoint = New AutoMath.DPosition
    Set oTopCenPoint = New AutoMath.DPosition
    
    oBaseCenPoint.Set -dFace1toCenter + flangeThick, 0, 0
    oTopCenPoint.Set 0, 0, 0
    Set ObjLeftCone = PlaceCone(m_OutputColl, oBaseCenPoint, oTopCenPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftCone
    Set ObjLeftCone = Nothing
    
 ' Insert your code for output 3(Triangle)
    Dim dActuatorDia As Double
    Dim dActuatorDia1 As Double
    Dim dActuatorDia2 As Double
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        dActuatorDia1 = flangeDiam
    Else
        dActuatorDia1 = pipeDiam
    End If
    
    If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
        dActuatorDia2 = flangeDiam2
    Else
        dActuatorDia2 = pipeDiam2
    End If
    
    If CmpDblGreaterthan(dActuatorDia1, dActuatorDia2) Then
        dActuatorDia = dActuatorDia1
    Else
        dActuatorDia = dActuatorDia2
    End If
    
    'Form a Triangle using Point Strings
    'Triangle Prism  Base is considered as 1.5 times the greater of the Pipe Dia and Flange Dia
    'and it's Y-Co ordinate will start from FacetoFace / 4 Below the Origin
    
    Dim dblPnts(0 To 11)  As Double
    Dim oProjectionPoint As AutoMath.DPosition
    Set oProjectionPoint = New AutoMath.DPosition
    Dim oLineString As IngrGeom3D.LineString3d
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    'Lower Left Corner
    dblPnts(0) = -(dFace1toCenter + dFace2toCenter) / 8
    dblPnts(1) = -(dFace1toCenter + dFace2toCenter) / 4
    dblPnts(2) = -0.75 * dActuatorDia
    'Lower Right Corner
    dblPnts(3) = -(dFace1toCenter + dFace2toCenter) / 8
    dblPnts(4) = -(dFace1toCenter + dFace2toCenter) / 4
    dblPnts(5) = 0.75 * dActuatorDia
    'Apex
    dblPnts(6) = -(dFace1toCenter + dFace2toCenter) / 8
    dblPnts(7) = parActuatorOffset - dActuatorDia / 2
    dblPnts(8) = 0
    'Lower Left Corner
    dblPnts(9) = -(dFace1toCenter + dFace2toCenter) / 8
    dblPnts(10) = -(dFace1toCenter + dFace2toCenter) / 4
    dblPnts(11) = -0.75 * dActuatorDia
    oProjectionPoint.x = 1
    oProjectionPoint.y = 0
    oProjectionPoint.z = 0
    'Place Line String forming Triangle
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 4, dblPnts)
    'Project the Line String to the required Direction
    Set ObjTriangularDisc = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                        oLineString, _
                                                        oProjectionPoint.x, oProjectionPoint.y, oProjectionPoint.z, _
                                                        (dFace1toCenter + dFace2toCenter) / 4, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTriangularDisc
    Set ObjTriangularDisc = Nothing
    Dim sp3dElem As IJDObject
    Set sp3dElem = oLineString
    Set oLineString = Nothing
    Set oProjectionPoint = Nothing
    sp3dElem.Remove
    
 ' Insert your code for output 4(Cylinder)
    Dim oStartPoint As AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    Set oStartPoint = New AutoMath.DPosition
    Set oEndPoint = New AutoMath.DPosition
    
    oStartPoint.Set -(dFace1toCenter + dFace2toCenter) / 16, parActuatorOffset, 0
    oEndPoint.Set (dFace1toCenter + dFace2toCenter) / 16, parActuatorOffset, 0
    Set ObjActuatorCylinder = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, dActuatorDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorCylinder
    Set ObjActuatorCylinder = Nothing
    Set oStartPoint = Nothing
    Set oEndPoint = Nothing
    
 ' Insert your code for output 5(Cone)
    oBaseCenPoint.Set dFace2toCenter - flangeThick2, 0, 0
    oTopCenPoint.Set 0, 0, 0
    Set ObjRightCone = PlaceCone(m_OutputColl, oBaseCenPoint, oTopCenPoint, pipeDiam2 / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightCone
    Set ObjRightCone = Nothing
    Set oBaseCenPoint = Nothing
    Set oTopCenPoint = Nothing

' Insert your code for output 1(Nozzle)
    Dim oPlacePoint As AutoMath.DPosition
    Set oPlacePoint = New AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -dFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set ObjNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1
    Set ObjNozzle1 = Nothing
 
 ' Insert your code for output 6(Nozzle)
    oPlacePoint.Set dFace2toCenter + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set ObjNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2
    Set ObjNozzle2 = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

' Insert your code for output 8 (Valve Operator)
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
    
    oDirX.Set Cos(parHandwheelAngle), 0, Sin(parHandwheelAngle)
    oDirY.Set 0, 1, 0
    oDirZ.Set -Sin(parHandwheelAngle), 0, Cos(parHandwheelAngle)

'    Set oPipeComponent = oPartFclt
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
            OpOrigin.Set 0, 0, 0
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
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
