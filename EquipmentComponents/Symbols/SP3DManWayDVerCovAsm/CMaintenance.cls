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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:10:54 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:         svsmylav
'   Creation Date:   Tuesday, April 222003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' This symbol is based on "Davit for Vertical Cover" symbol
' of The DOW Chemical company.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Maintenance:" 'Used for error messages
Private Const POS_RIGHT = 3
Private PI       As Double

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
    PI = Atn(1) * 4
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
''
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
    Dim ObjMaintenaceEnvelope As Object

    Dim parCentertoDavitCover As Double
    Dim parVesselDiameter As Double
    Dim parDavitHinge    As Long
    Dim parInsulationThickness As Double


' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCentertoDavitCover = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parDavitHinge = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)

    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    
' Insert your code for output 1(Maintenace Envelope)
    Dim LineStrPoints(0 To 14)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim angle As Double
    Dim MaintenanceClearance As Double  'Use this to have additional maintenance envelope

    MaintenanceClearance = 0.1

    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim NozzleLength As Double
    NozzleLength = parCentertoDavitCover - parVesselDiameter / 2

    LineStrPoints(0) = -flangeDiam / 2 - MaintenanceClearance
    LineStrPoints(1) = parCentertoDavitCover - flangeThick / 2
    LineStrPoints(2) = -flangeDiam / 2 - MaintenanceClearance

    LineStrPoints(3) = LineStrPoints(0)
    LineStrPoints(4) = LineStrPoints(1)
    LineStrPoints(5) = flangeDiam / 2 + 0.15 + MaintenanceClearance

    'Use 0.0001 to avoid intersection of linestring with axis
    LineStrPoints(6) = flangeDiam / 2 + 0.05 - 0.0001
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
    HingeAxisPoint.Set flangeDiam / 2 + 0.05, _
                parCentertoDavitCover - flangeThick / 2, 0
    TowerCenPoint.Set 0, 0, 0
    tmpVec.Set (HingeAxisPoint.x - TowerCenPoint.x), _
                (HingeAxisPoint.y - TowerCenPoint.y), _
                (HingeAxisPoint.z - TowerCenPoint.z)
    Set TowerCenPoint = Nothing

    Dim CenToCen As Double
    CenToCen = tmpVec.Length

    Dim zVec As New AutoMath.DVector
    zVec.Set 0, 0, 1
    Dim yVec As New AutoMath.DVector
    yVec.Set 0, 1, 0
    Dim angCen As Double    'Angle between Y-axis and vector joining hinge-vessel centers
'   ang2 is angle between vector joining hinge-vessel centers, and radial line upto tangent to vessel.
    Dim ang2 As Double
    angCen = Abs(tmpVec.angle(yVec, zVec))
    ang2 = Atn(Sqr(CenToCen ^ 2 - (parVesselDiameter / 2) ^ 2) / _
                    (parVesselDiameter / 2))

    angle = PI + angCen + ang2

    If parDavitHinge = POS_RIGHT Then
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

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMaintenaceEnvelope
    Set ObjMaintenaceEnvelope = Nothing
    Set yVec = Nothing
    Set zVec = Nothing
    Set tmpVec = Nothing
    Set HingeAxisPoint = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext

End Sub
