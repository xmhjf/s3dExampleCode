VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:04:26 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Friday, Apr 25, 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This is a Shell and Tube Exchanger(E-162).
'   In this Symbol there are total 16 outputs.14 outputs are physical out of which 6 are pipe Nozzles.
'   There are 2 Reference Geometry objects which are Point at origin and a reference plane at support bottom.
'   This is a Dow symbol Equipment
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private PI As Double
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Physical:" 'Used for error messages
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oSymGeomHelper = New SymbolServices
    PI = Atn(1) * 4
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
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim parVesselDiameter As Double
    Dim parVesselLength As Double
    Dim parVesselTantoTan As Double
    Dim parVesselLength1 As Double
    Dim parVesselStartPoint As Double
    Dim parSupportLength As Double
    Dim parSupportWidth As Double
    Dim parSupportHeight  As Double
    Dim parFirstSupportLocation  As Double
    Dim parSecondSupportLocation As Double
    

    Dim iOutput     As Double
    
    Dim ObjFrontHeadBonnet As Object
    Dim ObjRearHeadBonnet As Object
    Dim ObjExchangerSupport1 As Object
    Dim ObjExchangerSupport2 As Object
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselLength = arrayOfInputs(3)
    parVesselTantoTan = arrayOfInputs(4)
    parVesselLength1 = arrayOfInputs(5)
    parVesselStartPoint = arrayOfInputs(6)
    parSupportLength = arrayOfInputs(7)
    parSupportWidth = arrayOfInputs(8)
    parSupportHeight = arrayOfInputs(9)
    parFirstSupportLocation = arrayOfInputs(10)
    parSecondSupportLocation = arrayOfInputs(11)
    
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    
    Dim dShellLength As Double
    Dim dDomeHeight As Double
    Dim dFlangeThickness As Double
    Dim dFlangeDiameter As Double
    
    dShellLength = parVesselTantoTan - parVesselLength1
    dDomeHeight = (parVesselLength - parVesselTantoTan) / 2
'    'Assumptions
    dFlangeThickness = 0.0508                                             ' 2 Inch
    dFlangeDiameter = parVesselDiameter + 0.1016       'Vessel OD +4 inch
    
    
' Insert your code for outputs 1(ObjFrontHeadBonnet)
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oAxisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition
    Dim cenX As Double
    Dim cenY As Double
    Dim cenZ As Double
    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    Dim MMRatio As Double
    Dim StartAngle As Double
    Dim SweepAngle As Double
    Dim norX As Double
    Dim norY As Double
    Dim norZ As Double
    Dim Linepts(0 To 5) As Double
    
    Dim oExchangerArc1 As IngrGeom3D.EllipticalArc3d
    Dim oExchangerArc2 As IngrGeom3D.EllipticalArc3d

    'First Curve
    cenX = -parVesselStartPoint
    cenY = 0
    cenZ = 0

    MajorX = 0
    MajorY = 0
    MajorZ = parVesselDiameter / 2

    MMRatio = dDomeHeight / (parVesselDiameter / 2)
    StartAngle = PI
    SweepAngle = PI / 2
    
    norX = 0
    norY = 1
    norZ = 0
    
    Set oExchangerArc1 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                            cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, MMRatio, _
                                            StartAngle, SweepAngle)
                                            
    oAxisVect.Set 1, 0, 0
    revCenPt.Set 0, 0, 0
    'Revolve it about X-Axiz
    Set ObjFrontHeadBonnet = PlaceRevolution(m_OutputColl, oExchangerArc1, oAxisVect, revCenPt, PI * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontHeadBonnet
    Set ObjFrontHeadBonnet = Nothing
    
' Insert your code for output 2(ObjFrontHeadBodyCyl)As a Cylinder
    oStPoint.Set -parVesselStartPoint, 0, 0
    oEnPoint.Set parVesselLength1 - parVesselStartPoint - dFlangeThickness, 0, 0
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
     m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parVesselDiameter
        
' Insert your code for output 3(ObjFrontHeadFlangeTop)As a Cylinder
    oStPoint.Set parVesselLength1 - parVesselStartPoint - dFlangeThickness, 0, 0
    oEnPoint.Set oStPoint.x + dFlangeThickness, 0, 0
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dFlangeDiameter
    
' Insert your code for output 4(ObjFrontHeadFlangeBottom)As a Cylinder
    oStPoint.Set parVesselLength1 - parVesselStartPoint, 0, 0
    oEnPoint.Set oStPoint.x + dFlangeThickness, 0, 0
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dFlangeDiameter
    
' Insert your code for output 5(ObjShellBodyCyl)As a Cylinder
    oStPoint.Set parVesselLength1 - parVesselStartPoint + dFlangeThickness, 0, 0
    oEnPoint.Set parVesselLength1 - parVesselStartPoint + dShellLength, 0, 0
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parVesselDiameter

' Insert your code for outputs 6(ObjRearHeadBonnet)
    cenX = parVesselLength1 - parVesselStartPoint + dShellLength
    cenY = 0
    cenZ = 0

    MajorX = 0
    MajorY = 0
    MajorZ = parVesselDiameter / 2

    MMRatio = dDomeHeight / (parVesselDiameter / 2)
    StartAngle = PI / 2
    SweepAngle = PI / 2
    
    norX = 0
    norY = 1
    norZ = 0
    
    Set oExchangerArc2 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                            cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, MMRatio, _
                                            StartAngle, SweepAngle)
                                            
    oAxisVect.Set 1, 0, 0
    revCenPt.Set 0, 0, 0
                        
    'Revolve it about X-Axiz
    Set ObjRearHeadBonnet = PlaceRevolution(m_OutputColl, oExchangerArc2, oAxisVect, revCenPt, PI * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRearHeadBonnet
    Set ObjRearHeadBonnet = Nothing
    
    Dim Objcurves As IJDObject
    Set Objcurves = oExchangerArc1
    Objcurves.Remove
    Set Objcurves = Nothing
    Set Objcurves = oExchangerArc2
    Objcurves.Remove
    
    Set Objcurves = Nothing
    Set revCenPt = Nothing
    
' Insert your code for output 7(ObjExchangerSupport1)
    Dim oComplexStr As IngrGeom3D.ComplexString3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim oLine As IngrGeom3D.Line3d
    Dim oEleCollection           As Collection
    Set oEleCollection = New Collection
    
    Dim dAngle As Double
    Dim dMeetPoint As Double
    
    dMeetPoint = Sqr(((parVesselDiameter / 2) ^ 2) - ((parSupportWidth / 2) ^ 2))

    Dim dlinepts() As Double
    ReDim dlinepts(0 To 14) As Double
    
    ''Support Points
    dlinepts(0) = parFirstSupportLocation - parSupportLength / 2
    dlinepts(1) = parSupportWidth / 2
    dlinepts(2) = -dMeetPoint

    dlinepts(3) = dlinepts(0)
    dlinepts(4) = parSupportWidth / 2
    dlinepts(5) = -parSupportHeight
    
    dlinepts(6) = dlinepts(0)
    dlinepts(7) = -parSupportWidth / 2
    dlinepts(8) = dlinepts(5)
    
    dlinepts(9) = dlinepts(0)
    dlinepts(10) = -parSupportWidth / 2
    dlinepts(11) = -dMeetPoint
    
    dlinepts(12) = dlinepts(0)
    dlinepts(13) = 0
    dlinepts(14) = -parVesselDiameter / 2
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(0), dlinepts(1), dlinepts(2), _
                                                      dlinepts(3), dlinepts(4), dlinepts(5))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(3), dlinepts(4), dlinepts(5), _
                                                      dlinepts(6), dlinepts(7), dlinepts(8))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, dlinepts(6), dlinepts(7), dlinepts(8), _
                                                      dlinepts(9), dlinepts(10), dlinepts(11))
    
    oEleCollection.Add oLine
    Set oLine = Nothing
    
    Set oArc = oGeomFactory.Arcs3d.CreateBy3Points(Nothing, dlinepts(9), dlinepts(10), dlinepts(11), _
                                                            dlinepts(12), dlinepts(13), dlinepts(14), _
                                                            dlinepts(0), dlinepts(1), dlinepts(2))
    oEleCollection.Add oArc
    Set oArc = Nothing
    
    oStPoint.Set dlinepts(0), dlinepts(1), dlinepts(2)
    Set oComplexStr = PlaceTrCString(oStPoint, oEleCollection)
    
    oAxisVect.Set 1, 0, 0
    Set ObjExchangerSupport1 = PlaceProjection(m_OutputColl, oComplexStr, oAxisVect, parSupportLength, True)
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerSupport1
    Set ObjExchangerSupport1 = Nothing
    
    Dim Count As Integer
    For Count = 1 To oEleCollection.Count
        oEleCollection.Remove 1
    Next Count
    
' Insert your code for output 8(ObjExchangerSupport2)
' The complex string created for the Support1 is translated along X-Dir to use it for projection for 2nd Support.
    Dim translateMat     As New AutoMath.DT4x4   'For complexstring translation
    Dim xTransVec As New AutoMath.DVector
    xTransVec.Set 1, 0, 0
    xTransVec.Length = parSecondSupportLocation - parFirstSupportLocation                      'C.L. to C.L. of supports = 19 feet
    translateMat.LoadIdentity
    translateMat.Translate xTransVec
    
    'Transform the complex string to the new location
    oComplexStr.Transform translateMat
    
    Set ObjExchangerSupport2 = PlaceProjection(m_OutputColl, oComplexStr, oAxisVect, parSupportLength, True)
     ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerSupport2
    Set ObjExchangerSupport2 = Nothing
    
    oComplexStr.RemoveCurve True
    Set oComplexStr = Nothing
    Set translateMat = Nothing
    Set xTransVec = Nothing

    Set oAxisVect = Nothing
    Set oEleCollection = Nothing
    Set oGeomFactory = Nothing
    Exit Sub
    
ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
