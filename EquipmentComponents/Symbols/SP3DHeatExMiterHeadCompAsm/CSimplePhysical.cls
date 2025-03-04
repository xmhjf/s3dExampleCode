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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:11:13 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          ACM
'   Creation Date:  Friday, May 16, 2003
'   Description:
 'This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003  SymbolTeam(India)   Copyright Information, Header  is added.
'   22.Aug.2003  SymbolTeam(India)   TR 46728 Change Port Index of a nozzle to 15 used to corrupt Equipment.
'                                    Modified port index logic.
'   23.Aug.2003  SymbolTeam(India)       Updated Error Handling Code
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   20.May.2005    MS  CR-76070: Modified the nozzle creation code(as members),added datum points.
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'   This is an Exchanger Miter Head Symbol.
'   This symbol is a Fully parametic symbol, consists of 23 inputs and 9 outputs.It has 2 aspects one is Simple Physical
'   another is ReferenceGeometry. 7 are physical and 2 is ReferenceGeometry.
'   This is a Dow Equipment E-161 exchanger Miter Head symbol

Option Explicit
Private PI As Double

Dim m_oSymGeomHelper As IJSymbolGeometryHelper


Private Const MODULE = "SimplePhysical:" 'Used for error messages

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
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim parVesselDiameter As Double
    Dim parFaceDiameter As Double
    Dim parFaceThickness As Double
    Dim parCPtoFace As Double
    
        
    Dim iOutput     As Double
   
    Dim ObjMiterTop15degBody As Object
    Dim ObjMiterTop30degBody As Object
    Dim ObjMiterBottom30degBody As Object
    Dim ObjMiterBottom15degBody As Object

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parFaceDiameter = arrayOfInputs(3)
    parFaceThickness = arrayOfInputs(4)
    parCPtoFace = arrayOfInputs(5)
    


    
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New GeometryFactory
    
''Assumed Origin is at Vessel Control Point.
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim dActualVesselDiameter As Double
    ' 'Assumptions
    'The input parameter for Vessel Diameter is Internal diameter .We need to take care for Thickness of Body.
    dActualVesselDiameter = parVesselDiameter + 0.0508   ' 2 Inch
    
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
    Dim dMinorToMajorRatio As Double
    Dim dEllipseMajorAxis As Double
    
''Miter Radius is assumed to 1.2 times the Vessel Diameter
    dMiterRadius = parVesselDiameter
''Assumed 4inch is a factor which is to be deducted from the Miter Radius to get effective Miter radius.
''This is because to adjust the Miter Top cylider height or Length of Nozzle B placed at the Miter end,
''Incase occurance attributes Nozzle B position or Orientation is changed.
    dMiterRadius = dMiterRadius - 0.102
    Dim nozzPos As Double
    nozzPos = parCPtoFace + parFaceThickness + 1.25 * dMiterRadius
'Insert your code for output 1(Flange on Miter)

    Dim oStPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition

    oStPoint.Set 0, 0, -parCPtoFace
    oEndPoint.Set 0, 0, -parCPtoFace - parFaceThickness
    
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEndPoint, parFaceDiameter
    
    dRateOfChangeOfMiterDia = (dActualVesselDiameter / 2 - pipeDiam / 2) / 3 'For 30 degree miter segments

' Insert your code for output 2 (Rear Head Miter Bottom 15 degree portion)

    dCenterX = 0    'Face aligned to the center of Miter Bend
    dCenterY = -dMiterRadius
    dCenterZ = -nozzPos
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
    dCenterZ = -nozzPos + dMiterRadius * (1 - Cos(15 * PI / 180))
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

    Set ObjMiterBottom15degBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                        oEllipse(1), oEllipse(2), True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMiterBottom15degBody
    Set ObjMiterBottom15degBody = Nothing

' Insert your code for output 3 (Rear Head Miter Bottom 30 degree portion)
'    Current segment angle is 30 + first segment's angle is 15 =45 degrees
    dCenterX = 0
    dCenterY = -dMiterRadius * (1 - Sin(45 * PI / 180))
    dCenterZ = -nozzPos + dMiterRadius * (1 - Cos(45 * PI / 180))
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

    Set ObjMiterBottom30degBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                        oEllipse(2), oEllipse(3), True)


'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMiterBottom30degBody
    Set ObjMiterBottom30degBody = Nothing

' Insert your code for output 4 (Rear Head Miter Top 30 degree portion)
'    Current segment angle is 30 + (first two segment's angles are 15 + 30 =45) = 75 degrees
    dCenterX = 0
    dCenterY = -dMiterRadius * (1 - Sin(75 * PI / 180))
    dCenterZ = -nozzPos + dMiterRadius * (1 - Cos(75 * PI / 180))
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

    Set ObjMiterTop30degBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                        oEllipse(3), oEllipse(4), True)


'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMiterTop30degBody
    Set ObjMiterTop30degBody = Nothing

' Insert your code for output 5 (Rear Head Miter Top 15 degree portion)
'    Current segment angle is 90 degrees
    dCenterX = 0
    dCenterY = 0
    dCenterZ = -nozzPos + dMiterRadius
    dNormalX = 0
    dNormalY = 0
    dNormalZ = 1
    dMajorX = 0
    dMajorY = dActualVesselDiameter / 2
    dMajorZ = 0
    dMinorToMajorRatio = 1
    
    ''This is Circle only.
    Set oEllipse(5) = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                        dCenterX, dCenterY, dCenterZ, _
                                                        dNormalX, dNormalY, dNormalZ, _
                                                        dMajorX, dMajorY, dMajorZ, dMinorToMajorRatio)

    Set ObjMiterTop15degBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                        oEllipse(4), oEllipse(5), True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMiterTop15degBody
    Set ObjMiterTop15degBody = Nothing
    
    '   Cleanup
    Dim iCount As Integer
    For iCount = 1 To 5
        Set oEllipse(iCount) = Nothing
    Next iCount
    
'Insert your code for output 6(Miter Top Cylinder portion)

    oStPoint.Set 0, 0, -parCPtoFace - parFaceThickness
    oEndPoint.Set 0, 0, -parCPtoFace - parFaceThickness - 0.25 * dMiterRadius

' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEndPoint, dActualVesselDiameter
    
    Set oStPoint = Nothing
    Set oEndPoint = Nothing
    Set oGeomFactory = Nothing
        
    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

