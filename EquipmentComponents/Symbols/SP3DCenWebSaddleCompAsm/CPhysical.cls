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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:02:00 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RUK
'   Creation Date:  Friday, May 4 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Six Outputs:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private PI As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    Set m_oSymGeomHelper = New SymbolServices
    PI = 4 * Atn(1)
     
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parVesselDiameter As Double
    Dim parBaseplateLength As Double
    Dim parCenToSaddleBottom As Double
    Dim parBackingPlateThick As Double
    Dim parBasePlateThick As Double
    Dim parXBoltHole As Double
    Dim parBaseplateWidth As Double
    Dim parRibWidth As Double
    Dim parBackingPlateWidth As Double
    Dim parWebThickness As Double
    Dim parRibThick As Double
    Dim parYBoltHole As Double
    Dim parNoOfRibs As Long

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parBaseplateLength = arrayOfInputs(3)       'A
    parCenToSaddleBottom = arrayOfInputs(4)     'B
    parBackingPlateThick = arrayOfInputs(5)     'C
    parBasePlateThick = arrayOfInputs(6)        'D
    parXBoltHole = arrayOfInputs(7)             'E
    parBaseplateWidth = arrayOfInputs(8)        'F
    parRibWidth = arrayOfInputs(9)              'G
    parBackingPlateWidth = arrayOfInputs(10)    'H
    parWebThickness = arrayOfInputs(11)         'J
    parRibThick = arrayOfInputs(12)             'L
    parYBoltHole = arrayOfInputs(13)            'M
    parNoOfRibs = arrayOfInputs(14)

    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    
    'Create the Backing Plate (Output 1)
            'Backing plate is created such that it's center is origin and it always started R/10 th distance from the support plate
            'So angle between support plate point and start point of the arc is a=(R/10)*(1/(R+t))
            'Angle Betwwen backing plate start point and X is 90-(60+a).Because the angle between supports is 120.
            If CmpDblGreaterthan(parVesselDiameter, LINEAR_TOLERANCE) Then
        Dim oComplexStr As IngrGeom3D.ComplexString3d
        Dim oAxisVec As New AutoMath.DVector
        Dim oArc As IngrGeom3D.Arc3d
        Dim oLine As IngrGeom3D.Line3d
        Dim oEleCollection As Collection
        Dim oCenter As New AutoMath.DPosition
        Dim dArcStartAng As Double
        
        If CmpDblGreaterthan(parBackingPlateThick, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBackingPlateWidth, LINEAR_TOLERANCE) Then
            Dim objBackingPlate As Object
            
            Set oEleCollection = New Collection
            dArcStartAng = (PI / 2) - ((PI / 3) + (parVesselDiameter / 2) / (10 * ((parVesselDiameter / 2) + parBackingPlateThick)))
            oStPoint.Set ((parVesselDiameter / 2) + parBackingPlateThick) * Cos(dArcStartAng), -parBackingPlateWidth / 2, _
                                                        -((parVesselDiameter / 2) + parBackingPlateThick) * Sin(dArcStartAng)
            oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
            oCenter.Set 0, -parBackingPlateWidth / 2, -((parVesselDiameter / 2) + parBackingPlateThick)
            Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, oCenter.x, oCenter.y, oCenter.z, _
                                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            oEleCollection.Add oArc
            Set oArc = Nothing
            
            oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
            oEnPoint.Set oStPoint.x + parBackingPlateThick * Cos(dArcStartAng), oStPoint.y, oStPoint.z + parBackingPlateThick * Sin(dArcStartAng)
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                                oEnPoint.x, oEnPoint.y, oEnPoint.z)
            oEleCollection.Add oLine
            Set oLine = Nothing
            
            oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
            oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
            oCenter.Set 0, -parBackingPlateWidth / 2, -(parVesselDiameter / 2)
            Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, oCenter.x, oCenter.y, oCenter.z, _
                                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            oEleCollection.Add oArc
            Set oArc = Nothing
            
            oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
            oEnPoint.Set oStPoint.x + parBackingPlateThick * Cos(dArcStartAng), oStPoint.y, oStPoint.z - parBackingPlateThick * Sin(dArcStartAng)
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                                oEnPoint.x, oEnPoint.y, oEnPoint.z)
            oEleCollection.Add oLine
            Set oLine = Nothing
            
            oAxisVec.Set 0, 1, 0
            oStPoint.Set ((parVesselDiameter / 2) + parBackingPlateThick) * Cos(dArcStartAng), -parBackingPlateWidth / 2, _
                                                        -((parVesselDiameter / 2) + parBackingPlateThick) * Sin(dArcStartAng)
            Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)
            Set objBackingPlate = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, parBackingPlateWidth, True)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBackingPlate
            Set objBackingPlate = Nothing
            
            Dim iCount As Integer
            oComplexStr.RemoveCurve True
            For iCount = 1 To oEleCollection.Count
                oEleCollection.Remove 1
            Next iCount
        End If
        
        'Create the Web (Output 2)
        If CmpDblGreaterthan(parWebThickness, LINEAR_TOLERANCE) Then
            Dim objSupport As Object
            Set oComplexStr = New ComplexString3d
            Set oEleCollection = New Collection
            dArcStartAng = PI / 2 - PI / 3
            oCenter.Set 0, -parWebThickness / 2, 0
            oStPoint.Set ((parVesselDiameter / 2) + parBackingPlateThick) * Cos(dArcStartAng), -parWebThickness / 2, _
                                                        -((parVesselDiameter / 2) + parBackingPlateThick) * Sin(dArcStartAng)
            oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
        
            oCenter.Set 0, -parWebThickness / 2, -((parVesselDiameter / 2) + parBackingPlateThick)
            Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, oCenter.x, oCenter.y, oCenter.z, _
                                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            oEleCollection.Add oArc
            Set oArc = Nothing
        
            oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
            oEnPoint.Set oStPoint.x, oStPoint.y, -(parCenToSaddleBottom - parBasePlateThick)
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                                oEnPoint.x, oEnPoint.y, oEnPoint.z)
            oEleCollection.Add oLine
            Set oLine = Nothing
            
            oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
            oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                                oEnPoint.x, oEnPoint.y, oEnPoint.z)
            oEleCollection.Add oLine
            Set oLine = Nothing
        
            oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
            oEnPoint.Set oStPoint.x, oStPoint.y, -((parVesselDiameter / 2) + parBackingPlateThick) * Sin(dArcStartAng)
            Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                                oEnPoint.x, oEnPoint.y, oEnPoint.z)
            oEleCollection.Add oLine
            Set oLine = Nothing
        
            oStPoint.Set ((parVesselDiameter / 2) + parBackingPlateThick) * Cos(dArcStartAng), -parWebThickness / 2, _
                                                        -((parVesselDiameter / 2) + parBackingPlateThick) * Sin(dArcStartAng)
            Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)
            oAxisVec.Set 0, 1, 0
            Set objSupport = PlaceProjection(m_OutputColl, oComplexStr, oAxisVec, parWebThickness, True)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objSupport
            Set objSupport = Nothing
        
            oComplexStr.RemoveCurve True
            For iCount = 1 To oEleCollection.Count
                oEleCollection.Remove 1
            Next iCount
        End If
        Set oComplexStr = Nothing
        Set oEleCollection = Nothing
        Set oLine = Nothing
        Set oArc = Nothing
        Set oCenter = Nothing
        Set oAxisVec = Nothing
    End If

    'Create the Ribs (Output 3)
    If parNoOfRibs > 0 And CmpDblGreaterthan(parRibThick, LINEAR_TOLERANCE) And CmpDblGreaterthan(parRibWidth, LINEAR_TOLERANCE) Then
        Dim objRib As Object
        Dim dRibStartAng As Double
        Dim dAngleBetRib As Double
        dRibStartAng = dArcStartAng
        If parNoOfRibs = 1 Then
            dRibStartAng = PI / 2
        Else
            dAngleBetRib = 2 * PI / (3 * (parNoOfRibs - 1))
        End If
        For iCount = 1 To parNoOfRibs
            oStPoint.Set ((parVesselDiameter / 2) + parBackingPlateThick) * Cos(dRibStartAng), _
                                parRibWidth / 2, _
                        -((parVesselDiameter / 2) + parBackingPlateThick) * Sin(dRibStartAng)
            If CmpDblLessThan(dRibStartAng, PI / 2) Then
                oEnPoint.Set oStPoint.x - parRibThick, -parRibWidth / 2, -(parCenToSaddleBottom - parBasePlateThick)
            Else
                oEnPoint.Set oStPoint.x + parRibThick, -parRibWidth / 2, -(parCenToSaddleBottom - parBasePlateThick)
            End If
                Set objRib = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the Output
            If iCount = 1 Then
                iOutput = iOutput + 1
            End If
            m_OutputColl.AddOutput "Rib_", objRib
            Set objRib = Nothing
            dRibStartAng = dRibStartAng + dAngleBetRib
        Next iCount
    End If

    'Create the Baseplate (Output 4)
    If CmpDblGreaterthan(parBaseplateLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parBaseplateWidth, LINEAR_TOLERANCE) And _
                CmpDblGreaterthan(parBasePlateThick, LINEAR_TOLERANCE) Then
        Dim objBaseplate As Object
        oStPoint.Set parBaseplateLength / 2, parBaseplateWidth / 2, -parCenToSaddleBottom
        oEnPoint.Set -parBaseplateLength / 2, -parBaseplateWidth / 2, -(parCenToSaddleBottom - parBasePlateThick)
        Set objBaseplate = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBaseplate
        Set objBaseplate = Nothing
    
        'Create the Default Surface at the Bottom of the Baseplate (Output 5)
        Dim objPlane As IngrGeom3D.Plane3d
        Dim dPoints(0 To 11) As Double
    
        dPoints(0) = parBaseplateLength / 2
        dPoints(1) = parBaseplateWidth / 2
        dPoints(2) = -parCenToSaddleBottom
    
        dPoints(3) = dPoints(0)
        dPoints(4) = -dPoints(1)
        dPoints(5) = dPoints(2)
    
        dPoints(6) = -dPoints(0)
        dPoints(7) = dPoints(4)
        dPoints(8) = dPoints(2)
        
        dPoints(9) = dPoints(6)
        dPoints(10) = -dPoints(7)
        dPoints(11) = dPoints(2)
        
        'Create persistent default surface plane - the plane can mate ---
        Set objPlane = oGeomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, dPoints())
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
        Set objPlane = Nothing
    End If
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oGeomFactory = Nothing

    'Create Equipment Foundation Port at the Bottom of Base Plate (Output 5)
    '
    '                      |-------|
    '   Y                  |       |
    '   ^                  |       |
    '   |                  |   X   |
    '   |                  |   ^   |
    '   |                  |   |   |
    '   -----> X           |   |   |Port CS
    '   Symbol CS          |   |----> Y
    '                      |       |
    '                      |       |
    '                      |       |
    '                      |       |
    '                      |       |
    '                      |-------|
    Dim objFoundationPort As IJEqpFoundationPort
    Dim oNozzlePHFactory As NozzlePHFactory
    Set oNozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
   'The origin of the port is taken to be at the centre point of the support base.
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = -parCenToSaddleBottom
    
    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#

    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set objFoundationPort = oNozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                    False, m_OutputColl.ResourceManager)
    Dim holes() As Variant
    Call objFoundationPort.GetHoles(holes())

    holes(0, 1) = -parXBoltHole
    holes(0, 2) = -parYBoltHole
    holes(1, 1) = parXBoltHole
    holes(1, 2) = -parYBoltHole
    holes(2, 1) = parXBoltHole
    holes(2, 2) = parYBoltHole
    holes(3, 1) = -parXBoltHole
    holes(3, 2) = parYBoltHole

    Call objFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))

    Call objFoundationPort.SetHoles(holes)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objFoundationPort
    Set objFoundationPort = Nothing
    Set oNozzlePHFactory = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
