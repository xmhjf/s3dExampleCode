VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:         svsmylav
'   Creation Date:  Wednesday 30, Apr 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol is vessel skirt based on "Tower T-181" symbol
'   of The DOW Chemical company.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private Const POS_RIGHT = 3
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim pipeDiam(1 To 1) As Double
    Dim sptOffset(1 To 1) As Double
    Dim flangeDiam(1 To 1)    As Double
    Dim depth(1 To 1) As Double
    Dim flangeThick(1 To 1) As Double

    Dim iOutput     As Double
    Dim ObjMaintenaceEnvelope As Object
    
    Dim parFacetoCenter As Double
    Dim parVesselDiameter As Double
    Dim parNozzleLength As Double
    Dim parDavitPosition  As Long
    
'   Dynamic Port Input parameters variable declaration
    Dim parNPD As Double
    Dim parEndPreparation As Long
    Dim parScheduleThickness As Long
    Dim parEndStandard As Long
    Dim parPressureRating As Long
    Dim parFlowDirection As DistribFlow
    Dim parNPDUnitType As String

    Dim parPortIndex1 As Long
    Dim parNpd1 As Double
    Dim parEndPreparation1 As Long
    Dim parScheduleThickness1 As Long
    Dim parEndStandard1 As Long
    Dim parPressureRating1 As Long
    Dim parFlowDirection1 As DistribFlow
    Dim parId1 As String
    Dim parNPDUnitType1 As String
   
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2) 'Redundant, not used
    parVesselDiameter = arrayOfInputs(3)
    parNozzleLength = arrayOfInputs(4)
    parDavitPosition = arrayOfInputs(5)
    parNPD = arrayOfInputs(6)
    parEndPreparation = arrayOfInputs(7)
    parScheduleThickness = arrayOfInputs(8)
    parEndStandard = arrayOfInputs(9)
    parPressureRating = arrayOfInputs(10)
    parFlowDirection = arrayOfInputs(11)
    parPortIndex1 = arrayOfInputs(12)
    parNpd1 = arrayOfInputs(13)
    parEndPreparation1 = arrayOfInputs(14)
    parScheduleThickness1 = arrayOfInputs(15)
    parEndStandard1 = arrayOfInputs(16)
    parPressureRating1 = arrayOfInputs(17)
    parFlowDirection1 = arrayOfInputs(18)
    parId1 = arrayOfInputs(19)
    parNPDUnitType = arrayOfInputs(20)
    parNPDUnitType1 = arrayOfInputs(21)
    
    iOutput = 0
    
'   At this stage ComponentRotation is not exposed as input parameter.
'   It is set to default 0 Degree. If needed it can be exposed as Input
'   parameter. Transformation matrix is used to take care of it.
    Dim ComponentRotation As Double
    ComponentRotation = 0
    
'   Prepare rotation transformation matrix
    Dim transMatObj     As New AutoMath.DT4x4
    Dim zVec As New AutoMath.DVector
    transMatObj.LoadIdentity
    zVec.Set 0, 0, -1
    transMatObj.Rotate ComponentRotation, zVec
    
   
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
' Insert your code for output 1(Maintenace Envelope)
    Dim LineStrPoints(0 To 14)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim angle As Double
    Dim extraDist As Double  'Use this to have additional maintenance envelope
    
    extraDist = 0.1
    
    CreateRetrieveDynamicNozzle parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
                        
    Dim geomFactory As New IngrGeom3D.GeometryFactory
                        
    LineStrPoints(0) = -flangeDiam(parPortIndex1) / 2 - extraDist
    LineStrPoints(1) = parVesselDiameter / 2 + parNozzleLength - flangeThick(parPortIndex1) / 2
    LineStrPoints(2) = -flangeDiam(parPortIndex1) / 2 - extraDist
    
    LineStrPoints(3) = LineStrPoints(0)
    LineStrPoints(4) = LineStrPoints(1)
    LineStrPoints(5) = flangeDiam(parPortIndex1) / 2 + 0.15 + extraDist

    LineStrPoints(6) = flangeDiam(parPortIndex1) / 2 + 0.05 + extraDist
    LineStrPoints(7) = LineStrPoints(1)
    LineStrPoints(8) = LineStrPoints(5)
    
    LineStrPoints(9) = LineStrPoints(6)
    LineStrPoints(10) = LineStrPoints(1)
    LineStrPoints(11) = LineStrPoints(2)

    LineStrPoints(12) = LineStrPoints(0)
    LineStrPoints(13) = LineStrPoints(1)
    LineStrPoints(14) = LineStrPoints(2)
    
    Dim tmpVec As New AutoMath.DVector
    Dim HingeAxisPoint As New AutoMath.DPosition
    Dim TowerCenPoint As New AutoMath.DPosition
    'Use 0.0001 to avoid intersection of linestring with axis
    HingeAxisPoint.Set flangeDiam(parPortIndex1) / 2 + 0.05 + extraDist + 0.0001, _
                parVesselDiameter / 2 + parNozzleLength - flangeThick(parPortIndex1) / 2, 0
    TowerCenPoint.Set 0, 0, 0
    tmpVec.Set (HingeAxisPoint.x - TowerCenPoint.x), _
                (HingeAxisPoint.y - TowerCenPoint.y), _
                (HingeAxisPoint.z - TowerCenPoint.z)
    Set TowerCenPoint = Nothing
    
    Dim CenToCen As Double
    CenToCen = tmpVec.Length
    
    zVec.Set 0, 0, 1
    Dim yVec As New AutoMath.DVector
    yVec.Set 0, 1, 0
    Dim angCen As Double    'Angle between Y-axis and vector joining hinge-vessel centers
    Dim ang2 As Double      'Angle between vector joining hinge-vessel centers and tangent to vessel
    angCen = Abs(tmpVec.angle(yVec, zVec))
    ang2 = Atn(Sqr(CenToCen ^ 2 - (parVesselDiameter / 2) ^ 2) / _
                    (parVesselDiameter / 2))
                    
    angle = PI + angCen + ang2

    If parDavitPosition = POS_RIGHT Then
        Dim iCount As Integer
        For iCount = 1 To 5
            LineStrPoints(3 * iCount - 3) = -LineStrPoints(3 * iCount - 3)
        Next iCount
        HingeAxisPoint.x = -HingeAxisPoint.x
    Else
        zVec.Set 0, 0, -1
    End If
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPoints)
    Set geomFactory = Nothing
    
    Set ObjMaintenaceEnvelope = PlaceRevolution(m_OutputColl, oLineString, zVec, _
            HingeAxisPoint, angle, True)
            
'   Apply transformations to the object
    ObjMaintenaceEnvelope.Transform transMatObj
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintenaceEnvelope
    Set ObjMaintenaceEnvelope = Nothing
    Set transMatObj = Nothing
    Set yVec = Nothing
    Set zVec = Nothing
    Set tmpVec = Nothing
    Set HingeAxisPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
