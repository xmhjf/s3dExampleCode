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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:03:47 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, Apr 22, 2003
'   Description:
 'This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   This is a Falling Film S & T Exchanger Symbol.
'   This symbol consists of 29 inputs and 21 outputs.It has 2 aspects one is Simple Physical
'   another is ReferenceGeometry.Among them 20 are physical and 1 is ReferenceGeometry.
'   This is a Dow document Equipment symbol E-161
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'   28.Sep.2006     RRK  TR-84724           If variables of type double are compared, appropriate CmpDblXXX functions are used.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double

Dim m_oSymGeomHelper As IJSymbolGeometryHelper
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
    Dim parVesselHeight As Double
    Dim parVesselHeight1 As Double
    Dim parDomeHeight As Double
    Dim parVesselStartPoint As Double
        
    Dim iOutput     As Double
    
    Dim ObjFrontHeadBonnet As Object
    Dim ObjRearHeadMiterTop15degBody As Object
    Dim ObjRearHeadMiterTop30degBody As Object
    Dim ObjRearHeadMiterBottom30degBody As Object
    Dim ObjRearHeadMiterBottom15degBody As Object


' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselHeight = arrayOfInputs(3)
    parVesselHeight1 = arrayOfInputs(4)
    parDomeHeight = arrayOfInputs(5)
    parVesselStartPoint = arrayOfInputs(6)
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New GeometryFactory
    
''Assumed Origin is at Control Point.
    Dim dRearHeadHeight As Double
    Dim dFlangeDiameter As Double
    Dim dFlangeHeight As Double
    

    dRearHeadHeight = (parVesselHeight1 * 1.459600348) - (parVesselHeight - parVesselStartPoint) - parVesselHeight1
    dFlangeDiameter = parVesselDiameter + 0.102 '' parVesselDiameter + 4 inch (Assumed flange diameter be 4 inch more the Exchanger diameter.
    dFlangeHeight = 0.05        '' Assumed as 2 inch.
    
'Insert your code for output 1(Elliptical Front Head Bonnet Body)
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
    Dim oCenterPoint As New AutoMath.DPosition
    Dim oAxisVec As New AutoMath.DVector
    Dim oEllipticalArc As IngrGeom3D.EllipticalArc3d
        
    Dim dEllipseMajorAxis As Double
    Dim dMinorToMajorRatio As Double
    Dim dPoints() As Double
    
    dEllipseMajorAxis = parVesselDiameter / 2
    dMinorToMajorRatio = parDomeHeight / dEllipseMajorAxis
    
    oCenterPoint.Set 0, 0, parVesselStartPoint
    
    Set oEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                         oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, 0, 1, 0, dEllipseMajorAxis, 0, 0, _
                         dMinorToMajorRatio, PI, PI / 2)
    
    oAxisVec.Set 0, 0, 1

    Set ObjFrontHeadBonnet = PlaceRevolution(m_OutputColl, oEllipticalArc, oAxisVec, oCenterPoint, 2 * PI, True)

'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontHeadBonnet
    Set ObjFrontHeadBonnet = Nothing
    
'Insert your code for output 2(Cylindrical Front Head Body)
    oStPoint.Set 0, 0, parVesselStartPoint
    oEndPoint.Set 0, 0, -(parVesselHeight - parVesselStartPoint - dFlangeHeight)
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEndPoint, parVesselDiameter
        
'Insert your code for output 3(Top Flange on Front head side)

    oStPoint.Set 0, 0, -(parVesselHeight - parVesselStartPoint - dFlangeHeight)
    oEndPoint.Set 0, 0, -(parVesselHeight - parVesselStartPoint)
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEndPoint, dFlangeDiameter
    
        
'Insert your code for output 4(Bottom Flange on Front head side)

    oStPoint.Set 0, 0, -(parVesselHeight - parVesselStartPoint)
    oEndPoint.Set 0, 0, -(parVesselHeight - parVesselStartPoint + dFlangeHeight)
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEndPoint, dFlangeDiameter
    
        
'Insert your code for output 5(Cylindrical Exchanger Shell Body)

    oStPoint.Set 0, 0, -(parVesselHeight - parVesselStartPoint + dFlangeHeight)
    oEndPoint.Set 0, 0, -(parVesselHeight - parVesselStartPoint + parVesselHeight1 - dFlangeHeight)
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEndPoint, parVesselDiameter
    
    
'Insert your code for output 6(Top Flange on Rear head side)

    oStPoint.Set 0, 0, -(parVesselHeight - parVesselStartPoint + parVesselHeight1 - dFlangeHeight)
    oEndPoint.Set 0, 0, -(parVesselHeight - parVesselStartPoint + parVesselHeight1)
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEndPoint, dFlangeDiameter
    
        
'Insert your code for output 7(Bottom Flange on Rear head side)

    oStPoint.Set 0, 0, -(parVesselHeight - parVesselStartPoint + parVesselHeight1)
    oEndPoint.Set 0, 0, -(parVesselHeight - parVesselStartPoint + parVesselHeight1 + dFlangeHeight)
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEndPoint, dFlangeDiameter
        

''REAR HEAD MITER CODE

''Transformation matrix to rotate Miter, with respect to change in Nozzle B orientation angle
''which is an Occurance Attribute

    Dim oTransformationMatrix     As New AutoMath.DT4x4
    Dim oRotationAxisVector As New AutoMath.DVector
    Dim dRotationAngle As Double
    
    dRotationAngle = 0  ''This is kept as ZERO, since Transformation angle is not exposed as Input parameter.
    
    oTransformationMatrix.LoadIdentity
    oRotationAxisVector.Set 0, 0, -1
    oTransformationMatrix.Rotate dRotationAngle, oRotationAxisVector

    Dim dRateOfChangeOfMiterDia As Double
    Dim dMiterRadius As Double
    Dim oEllipse(1 To 5) As IngrGeom3D.Ellipse3d      'Four miter bodies need Five profiles
    Dim dCenterX   As Double
    Dim dCenterY   As Double
    Dim dCenterZ   As Double
    Dim dMajorX As Double
    Dim dMajorY As Double
    Dim dMajorZ As Double
    Dim dNormalX As Double
    Dim dNormalY As Double
    Dim dNormalZ As Double
    
''To Find out Miter Radius.
''Miter radius is considered as the minimum of Rear Head(Miter) Height and CP to Miter nozzle Distance.
''This Helps, when occurance attributes Nozzle B position or Orientation is changed. As per that change Miter size gets changed.

    If CmpDblGreaterthan(dRearHeadHeight, parVesselHeight * 0.8) Then
        dMiterRadius = (parVesselHeight * 0.8)
    Else
        dMiterRadius = dRearHeadHeight
    End If
    
''Assumed 4inch is a factor which is to be deducted from the Miter Radius to get effective Miter radius.
''This is because to adjust the Miter Top cylider height or Length of Nozzle B placed at the Miter end,
''Incase occurance attributes Nozzle B position or Orientation is changed.
    dMiterRadius = dMiterRadius - 0.102
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    dRateOfChangeOfMiterDia = (parVesselDiameter / 2 - pipeDiam / 2) / 3 'For 30 degree miter segments

' Insert your code for output 8 (Rear Head Miter Bottom 15 degree portion)

    dCenterX = 0    'Face aligned to the center of Miter Bend
    dCenterY = -dMiterRadius
    dCenterZ = -(parVesselHeight1 * 1.459600348)
    dNormalX = 0
    dNormalY = -1
    dNormalZ = 0
    dMajorX = 0
    dMajorY = 0
    dMajorZ = pipeDiam / 2
    dMinorToMajorRatio = 1
    
    ''This is a Circle only.
    Set oEllipse(1) = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                        dCenterX, dCenterY, dCenterZ, _
                                                        dNormalX, dNormalY, dNormalZ, _
                                                        dMajorX, dMajorY, dMajorZ, dMinorToMajorRatio)

    dCenterX = 0
    dCenterY = -dMiterRadius * (1 - Sin(15 * PI / 180))
    dCenterZ = -(parVesselHeight1 * 1.459600348) + dMiterRadius * (1 - Cos(15 * PI / 180))
    dNormalX = 0
    dNormalY = -Cos(15 * PI / 180)
    dNormalZ = -Sin(15 * PI / 180)

    'Use 0.5 * dRateOfChangeOfMiterDia for 15 degrees miter segment
    
    dEllipseMajorAxis = (0.5 * pipeDiam + 0.5 * dRateOfChangeOfMiterDia) / Cos(15 * PI / 180)
    dMajorX = dEllipseMajorAxis * Sin(15 * PI / 180)
    dMajorY = 0
    dMajorZ = dEllipseMajorAxis * Cos(15 * PI / 180)
    dMinorToMajorRatio = Cos(15 * PI / 180)
    Set oEllipse(2) = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                        dCenterX, dCenterY, dCenterZ, _
                                                        dNormalX, dNormalY, dNormalZ, _
                                                        dMajorX, dMajorY, dMajorZ, dMinorToMajorRatio)

    Set ObjRearHeadMiterBottom15degBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                        oEllipse(1), oEllipse(2), True)
   
'Apply the transformation to the object
    ObjRearHeadMiterBottom15degBody.Transform oTransformationMatrix

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRearHeadMiterBottom15degBody
    Set ObjRearHeadMiterBottom15degBody = Nothing

' Insert your code for output 9 (Rear Head Miter Bottom 30 degree portion)
'    Current segment angle is 30 + first segment's angle is 15 =45 degrees
    dCenterX = 0
    dCenterY = -dMiterRadius * (1 - Sin(45 * PI / 180))
    dCenterZ = -(parVesselHeight1 * 1.459600348) + dMiterRadius * (1 - Cos(45 * PI / 180))
    dNormalX = 0
    dNormalY = -Cos(45 * PI / 180)
    dNormalZ = -Sin(45 * PI / 180)
    'Use (45 * dRateOfChangeOfMiterDia/30)  for current segment
    dEllipseMajorAxis = (0.5 * pipeDiam + 1.5 * dRateOfChangeOfMiterDia) / Cos(30 * PI / 180)
    dMajorX = 0
    dMajorY = -dEllipseMajorAxis * Sin(45 * PI / 180)
    dMajorZ = dEllipseMajorAxis * Cos(45 * PI / 180)
    dMinorToMajorRatio = Cos(30 * PI / 180)  'Current segment s angle is 30 degrees
    Set oEllipse(3) = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                        dCenterX, dCenterY, dCenterZ, _
                                                        dNormalX, dNormalY, dNormalZ, _
                                                        dMajorX, dMajorY, dMajorZ, dMinorToMajorRatio)

    Set ObjRearHeadMiterBottom30degBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                        oEllipse(2), oEllipse(3), True)

'Apply the transformation to the object
    ObjRearHeadMiterBottom30degBody.Transform oTransformationMatrix

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRearHeadMiterBottom30degBody
    Set ObjRearHeadMiterBottom30degBody = Nothing

' Insert your code for output 10 (Rear Head Miter Top 30 degree portion)
'    Current segment angle is 30 + (first two segment's angles are 15 + 30 =45) = 75 degrees
    dCenterX = 0
    dCenterY = -dMiterRadius * (1 - Sin(75 * PI / 180))
    dCenterZ = -(parVesselHeight1 * 1.459600348) + dMiterRadius * (1 - Cos(75 * PI / 180))
    dNormalX = 0
    dNormalY = -Cos(75 * PI / 180)
    dNormalZ = -Sin(75 * PI / 180)
    
    'Use (75 * dRateOfChangeOfMiterDia/30)  for current segment
    dEllipseMajorAxis = (0.5 * pipeDiam + 2.5 * dRateOfChangeOfMiterDia) / Cos(15 * PI / 180)
    dMajorX = 0
    dMajorY = dEllipseMajorAxis * Sin(75 * PI / 180)
    dMajorZ = dEllipseMajorAxis * Cos(75 * PI / 180)
    dMinorToMajorRatio = Cos(30 * PI / 180) 'Current segment s angle is 30 degrees
    
    Set oEllipse(4) = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                        dCenterX, dCenterY, dCenterZ, _
                                                        dNormalX, dNormalY, dNormalZ, _
                                                        dMajorX, dMajorY, dMajorZ, dMinorToMajorRatio)

    Set ObjRearHeadMiterTop30degBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                        oEllipse(3), oEllipse(4), True)

'Apply the transformation to the object
    ObjRearHeadMiterTop30degBody.Transform oTransformationMatrix

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRearHeadMiterTop30degBody
    Set ObjRearHeadMiterTop30degBody = Nothing

' Insert your code for output 11 (Rear Head Miter Top 15 degree portion)
'    Current segment angle is 90 degrees
    dCenterX = 0
    dCenterY = 0
    dCenterZ = -(parVesselHeight1 * 1.459600348) + dMiterRadius
    dNormalX = 0
    dNormalY = 0
    dNormalZ = 1
    dMajorX = 0
    dMajorY = parVesselDiameter / 2
    dMajorZ = 0
    dMinorToMajorRatio = 1
    
    ''This is Circle only.
    Set oEllipse(5) = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                        dCenterX, dCenterY, dCenterZ, _
                                                        dNormalX, dNormalY, dNormalZ, _
                                                        dMajorX, dMajorY, dMajorZ, dMinorToMajorRatio)

    Set ObjRearHeadMiterTop15degBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                        oEllipse(4), oEllipse(5), True)

'Apply the transformation to the object
    ObjRearHeadMiterTop15degBody.Transform oTransformationMatrix

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRearHeadMiterTop15degBody
    Set ObjRearHeadMiterTop15degBody = Nothing
    
    Set oTransformationMatrix = Nothing
    Set oRotationAxisVector = Nothing
    
    '   Cleanup
    Dim i As Integer
    For i = 1 To 5
        Set oEllipse(i) = Nothing
    Next i
    
'Insert your code for output 12(Rear Head Miter Top Cylinder portion)
    oStPoint.Set 0, 0, -(parVesselHeight - parVesselStartPoint) - parVesselHeight1 - dFlangeHeight
    oEndPoint.Set 0, 0, -(parVesselHeight1 * 1.459600348) + dMiterRadius

' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEndPoint, parVesselDiameter

    Set oCenterPoint = Nothing
    Set oAxisVec = Nothing
    Set oGeomFactory = Nothing
    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

