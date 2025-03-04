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
'   Creation Date:  Wednesday, Mar 12 2008
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------    -----               ------------------
'  12.MAR.2008      MA         CR-136876  Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double
Private Const SINGLE_TAP = 4
Private Const TWOTAPS_180DEGREES_APART = 10
Private Const THREETAPS_TAP3_45DEGREESCCW_FROMTAP1 = 16

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
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    Dim pipeDiam3        As Double
    Dim flangeThick3     As Double
    Dim sptOffset3       As Double
    Dim flangeDiam3      As Double
    Dim depth3           As Double
    Dim pipeDiam4        As Double
    Dim flangeThick4     As Double
    Dim sptOffset4       As Double
    Dim flangeDiam4      As Double
    Dim depth4           As Double
    Dim pipeDiam5        As Double
    Dim flangeThick5     As Double
    Dim sptOffset5       As Double
    Dim flangeDiam5      As Double
    Dim depth5           As Double
    
    Dim iOutput     As Double
    
    Dim parFacetoFace As Double
    Dim parPortHOffset As Double
    Dim parPortVOffset As Double
    Dim parTapOrientation As Long
    Dim parInsulationThickness As Double
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parPortHOffset = arrayOfInputs(3)
    parPortVOffset = arrayOfInputs(4)
    parTapOrientation = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)

    iOutput = 0

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                            sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                            sptOffset2, depth2
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    ' This symbol is based on the following three Part data basis value that govern its geometry,
    '   Orifice hub, butt weld - 871

    Dim dLineStringPoints(0 To 32) As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oCenterPoint As AutoMath.DPosition
    Set oCenterPoint = New AutoMath.DPosition
    Dim oAxisVect As AutoMath.DVector
    Set oAxisVect = New AutoMath.DVector
    Dim dBodyOutsideDiameter As Double
    Dim dConst As Double
    Dim dAngle As Double
    dAngle = 25 * (PI / 180) '25 degrees
    dConst = parFacetoFace * Tan(dAngle)
    dBodyOutsideDiameter = (2 * dConst) + pipeDiam2
    
    Select Case lPartDataBasis

        Case Is <= 1, 871
            
            'Insert your code for output 2(Body of Orifice Hub)
            'Point 1
            dLineStringPoints(0) = flangeThick
            dLineStringPoints(1) = 0
            dLineStringPoints(2) = 0
            
            'Point 2
            dLineStringPoints(3) = flangeThick
            dLineStringPoints(4) = flangeDiam / 2
            dLineStringPoints(5) = 0
            
            'Point 3
            dLineStringPoints(6) = flangeThick + 0.1 * parFacetoFace
            dLineStringPoints(7) = 0.4 * flangeDiam
            dLineStringPoints(8) = 0
            
            'Point 4
            dLineStringPoints(9) = flangeThick + 0.2 * parFacetoFace
            dLineStringPoints(10) = 0.4 * flangeDiam
            dLineStringPoints(11) = 0
            
            'Point 5
            dLineStringPoints(12) = flangeThick + 0.4 * parFacetoFace
            dLineStringPoints(13) = 0.4 * dBodyOutsideDiameter
            dLineStringPoints(14) = 0
            
            'Point 6
            dLineStringPoints(15) = flangeThick + 0.4 * parFacetoFace
            dLineStringPoints(16) = 0.5 * dBodyOutsideDiameter
            dLineStringPoints(17) = 0
            
            'Point 7
            dLineStringPoints(18) = flangeThick + 0.5 * parFacetoFace
            dLineStringPoints(19) = 0.5 * dBodyOutsideDiameter
            dLineStringPoints(20) = 0
            
            'Point 8
            dLineStringPoints(21) = 0.9 * parFacetoFace
            dLineStringPoints(22) = pipeDiam2 / 2
            dLineStringPoints(23) = 0
            
            'Point 9
            dLineStringPoints(24) = parFacetoFace
            dLineStringPoints(25) = pipeDiam2 / 2
            dLineStringPoints(26) = 0
            
            'Point 10
            dLineStringPoints(27) = parFacetoFace
            dLineStringPoints(28) = 0
            dLineStringPoints(29) = 0
            
            'Point 11
            dLineStringPoints(30) = flangeThick
            dLineStringPoints(31) = 0
            dLineStringPoints(32) = 0
            
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 11, dLineStringPoints)
            oAxisVect.Set 1, 0, 0
            oCenterPoint.Set 0, -0.0001, 0
            Dim objOrificeHub As Object
       
           Set objOrificeHub = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, oCenterPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objOrificeHub
            Set objOrificeHub = Nothing
            Set oLineString = Nothing

        Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.

    End Select

    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    'Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    
    oPlacePoint.Set -sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    'Place Nozzle 2
    oPlacePoint.Set parFacetoFace + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    Dim dTangentLength As Double
    'Tap length is assumed to be 3 inches
    dTangentLength = 0.0762 '3 inches
    
    If parTapOrientation = SINGLE_TAP Or parTapOrientation = TWOTAPS_180DEGREES_APART _
                Or parTapOrientation = THREETAPS_TAP3_45DEGREESCCW_FROMTAP1 Then
        'Place Nozzle 3
        'Tap 1
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, _
                                                        flangeDiam3, sptOffset3, depth3
        oPlacePoint.Set (parFacetoFace + sptOffset3 - depth3), _
                        (dBodyOutsideDiameter / 2 + sptOffset3 - depth3), 0
        oDir.Set Cos(dAngle), Sin(dAngle), 0
        
        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, _
                                                            oPlacePoint, dTangentLength)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    End If

    If parTapOrientation = TWOTAPS_180DEGREES_APART Or _
                    parTapOrientation = THREETAPS_TAP3_45DEGREESCCW_FROMTAP1 Then
        'Place Nozzle 4
        'Tap 2 (is oriented 180 degrees from Tap 1)
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick4, _
                                                        flangeDiam4, sptOffset4, depth4
        oPlacePoint.Set (parFacetoFace + sptOffset4 - depth4), _
                        -(dBodyOutsideDiameter / 2 + sptOffset4 - depth4), 0
        oDir.Set Cos(dAngle), -Sin(dAngle), 0
        
        Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, oDir, _
                                                            oPlacePoint, dTangentLength)

        'Set the output
        m_OutputColl.AddOutput "Noz4", objNozzle
        Set objNozzle = Nothing
    End If

    If parTapOrientation = THREETAPS_TAP3_45DEGREESCCW_FROMTAP1 Then   'Orifice Hub With 3 Taps
        'Place Nozzle 5
        'Tap 3
        RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam5, flangeThick5, _
                                                        flangeDiam5, sptOffset5, depth5
        'Tap 3 is oriented 45 degrees counter-clockwise from vertical up (Tap 1)
        Dim dAngle2 As Double
        dAngle2 = 45 * (PI / 180) '45 degrees
     
        Dim oTempCircle As Circle3d
        Set oTempCircle = New Circle3d
        oCenterPoint.Set parFacetoFace + sptOffset3 - depth3, _
                            (dBodyOutsideDiameter / 2 + sptOffset3 - depth3), 0
        oAxisVect.Set Cos(dAngle), Sin(dAngle), 0

        Set oTempCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                oAxisVect.x, oAxisVect.y, oAxisVect.z, pipeDiam3 / 2)
        
        Dim oTransMatrix As DT4x4
        Dim oRotVector As AutoMath.DVector
        Set oTransMatrix = New DT4x4
        Set oRotVector = New AutoMath.DVector
        
        oRotVector.Set 1, 0, 0
        oTransMatrix.LoadIdentity
        oTransMatrix.Rotate dAngle2, oRotVector
        oTempCircle.Transform oTransMatrix
        Set oDir = New DVector
        Dim x As Double
        Dim y As Double
        Dim z As Double
        oTempCircle.GetNormal x, y, z
        oDir.Set x, y, z
        
        oPlacePoint.Set (parFacetoFace + sptOffset5 - depth5), _
                        (dBodyOutsideDiameter / 2 + sptOffset5 - depth5) * Cos(dAngle2), _
                        (dBodyOutsideDiameter / 2 + sptOffset5 - depth5) * Sin(dAngle2)

        Set objNozzle = CreateNozzleWithLength(5, oPartFclt, m_OutputColl, oDir, _
                                                            oPlacePoint, dTangentLength)
        
        'Set the output
        m_OutputColl.AddOutput "Noz5", objNozzle
        Set objNozzle = Nothing
        Set oTempCircle = Nothing
        Set oTransMatrix = Nothing
        Set oRotVector = Nothing
    End If
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oCenterPoint = Nothing
    Set oAxisVect = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
