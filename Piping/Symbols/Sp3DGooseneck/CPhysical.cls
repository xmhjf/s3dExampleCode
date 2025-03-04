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
' Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          PS
'   Creation Date:  Tuesday, NOV 13 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol
''''''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Physical:"
Private PI           As Double
Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
       
    Dim iOutput     As Double
    Dim objCone     As Object
    Dim objDome     As Object
    Dim objcyl      As Object
    Dim ocbody      As Object
    
    
    Dim oArc        As IngrGeom3D.EllipticalArc3d
    Dim oCircle     As IngrGeom3D.Circle3d
    Dim oline       As IngrGeom3D.Line3d
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim BaseCenPoint As New AutoMath.DPosition
    Dim TopCenPoint  As New AutoMath.DPosition
    Dim surfcyl      As Object
   
    Dim startPoint   As New AutoMath.DPosition
    Dim endPoint     As New AutoMath.DPosition
    Dim axisVect     As New AutoMath.DVector
    Dim RevCenPt     As New AutoMath.DPosition
    Dim geomFactory  As New IngrGeom3D.GeometryFactory
        
    
    Dim pardimen1 As Double
    Dim pardimen2 As Double
    Dim pardimen3 As Double
    Dim pardimen4 As Double
    Dim pardimen5 As Double
    Dim parInsulationThickness As Double
        
    Set oPartFclt = arrayOfInputs(1)
    pardimen1 = arrayOfInputs(2)     '  length of the part'
    pardimen2 = arrayOfInputs(3)     '  Diameter of the part'
    pardimen3 = arrayOfInputs(4)     '  height1 of the part'
    pardimen4 = arrayOfInputs(5)     '  offset distance '
    pardimen5 = arrayOfInputs(6)     '  heigh2 of the part '
    parInsulationThickness = arrayOfInputs(7)
    iOutput = 0
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    ' OutPut 1: Cone1
    
    BaseCenPoint.Set 0, 0, 0
    TopCenPoint.Set pardimen1 * 0.15, 0, 0
    Set objCone = PlaceCone(m_OutputColl, BaseCenPoint, TopCenPoint, pipeDiam / 2, pardimen2 / 2, True)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone
    Set objCone = Nothing
       
    'OutPut 2:Placing Cylinder in Cone 1
    
    BaseCenPoint.Set pardimen1 * 0.05, 0, 0
    TopCenPoint.Set pardimen1 * 0.1, 0, 0
    Set objcyl = PlaceCylinder(m_OutputColl, BaseCenPoint, TopCenPoint, pardimen2 * 0.9, True)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objcyl
    Set objcyl = Nothing
    
    'OutPut 3: Placing the Body Cylinder
    
    startPoint.Set pardimen1 * 0.15, 0, 0
    endPoint.Set pardimen1 - pardimen2 / 2, 0, 0
    Set objcyl = PlaceCylinder(m_OutputColl, startPoint, endPoint, pardimen2, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objcyl
    Set objcyl = Nothing
       
    'OutPut4: Placing of top Dome

    Dim cenX As Double, cenY As Double, cenZ As Double
    Dim MajorX As Double, MajorY As Double, MajorZ As Double
    Dim norX As Double, norY As Double, norZ As Double
    Dim mMRatio As Double, StartAngle As Double, SweepAngle As Double
    cenX = pardimen1 - pardimen2 / 2
    cenY = 0
    cenZ = 0
    MajorX = 0
    MajorY = 0
    MajorZ = pardimen2 / 2
    mMRatio = 1
    StartAngle = PI / 2
    SweepAngle = PI / 2
    norX = 0
    norY = 1
    norZ = 0
    
    Set oArc = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
    axisVect.Set -1, 0, 0
    RevCenPt.Set pardimen1 / 2, 0, 0
    Set objDome = PlaceRevolution(m_OutputColl, oArc, axisVect, RevCenPt, 2 * PI, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objDome
    Set oArc = Nothing
    Set objDome = Nothing
       
    'OutPut 5: Placing the neck
    
    startPoint.Set pardimen3, 0, -pardimen4 / 2
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, startPoint.x, startPoint.y, startPoint.z, 0, 0, 1, pardimen2 / 2)
    RevCenPt.Set pardimen5, 0, -pardimen4 / 2
    axisVect.Set 0, 1, 0
    Set ocbody = PlaceRevolution(m_OutputColl, oCircle, axisVect, RevCenPt, 7 * PI / 18, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ocbody
    Set ocbody = Nothing
    Set axisVect = Nothing
    Set RevCenPt = Nothing
     
    ' OutPut 6: Placing the conections i.e the neck to the cylider body
     
    endPoint.Set pardimen3, 0, 0
    Set surfcyl = PlaceCylinder(m_OutputColl, startPoint, endPoint, pardimen2, True)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), surfcyl
    Set surfcyl = Nothing
    Set ocbody = Nothing
    Set geomFactory = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    'Output 7: Place Cone2
    
    Dim radius As Double
    radius = pardimen3 - pardimen5
    BaseCenPoint.Set pardimen5 + radius * Cos(7 * PI / 18), 0, -(pardimen4 / 2) - radius * Sin(7 * PI / 18)
    TopCenPoint.Set BaseCenPoint.x - 0.15 * pardimen1 * Sin(7 * PI / 18), 0, BaseCenPoint.z - 0.15 * pardimen1 * Cos(7 * PI / 18)
    Set objCone = PlaceCone(m_OutputColl, TopCenPoint, BaseCenPoint, pipeDiam / 2, pardimen2 / 2, True)
    ' Set the output
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone
    Set objCone = Nothing
    
    'OutPut 8: Place Cylinder in Cone 2
    
    Dim TopCenPoint1 As New AutoMath.DPosition
    TopCenPoint.Set BaseCenPoint.x - 0.05 * pardimen1 * Sin(7 * PI / 18), 0, BaseCenPoint.z - 0.05 * pardimen1 * Cos(7 * PI / 18)
    TopCenPoint1.Set BaseCenPoint.x - 0.1 * pardimen1 * Sin(7 * PI / 18), 0, BaseCenPoint.z - 0.1 * pardimen1 * Cos(7 * PI / 18)
    Set objcyl = PlaceCylinder(m_OutputColl, TopCenPoint, TopCenPoint1, pardimen2 * 0.9, True)
    ' Set the output
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objcyl
    Set objcyl = Nothing
       
    'OutPut 9: Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -cptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    ' Set the output
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    'OutPut 10: Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    oPlacePoint.Set BaseCenPoint.x - 0.15 * (pardimen1 + cptOffset - depth) * Sin(7 * PI / 18), 0, BaseCenPoint.z - 0.15 * (pardimen1 + cptOffset - depth) * Cos(7 * PI / 18)
    oDir.Set -Sin(7 * PI / 18), 0, -Cos(7 * PI / 18)
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
    ' Set the output
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set BaseCenPoint = Nothing
    Set TopCenPoint1 = Nothing
    Set TopCenPoint = Nothing
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub

