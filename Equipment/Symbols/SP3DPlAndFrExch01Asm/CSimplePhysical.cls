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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:03:33 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Wed, May 27, 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This is a Plate Exchanger(E-183).
'   In this Symbol there are total 22 outputs.20 outputs are physical out of which 4 are pipe Nozzles and one
'   Foundation Port.
'   Nozzles are fully Parametric.
'   There are 2 Reference Geometry objects which are Control-Point at origin and a reference plane at support bottom.
'   This is a Dow symbol Plate Exchanger.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private PI As Double
Private m_oSymGeomHelper As IJSymbolGeometryHelper
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
    
    Dim parEquipmentXLength As Double              'total Length
    Dim parEquipmentYLength As Double
    Dim parEquipmentZLength As Double
    Dim parFrontPlateThickness As Double           'FrontPlate thick
    Dim parRearPlateThickness As Double            'RearPlate thick
    Dim parPlatePackLength As Double
    Dim parCPtoFaceY As Double
    Dim parCPtoFaceZ As Double
    Dim parSupportLength As Double
    Dim parSupportWidth As Double
    Dim parSupportThickness As Double
    Dim parXboltHole As Double
    Dim parYboltHole As Double
   
    
    Dim iOutput     As Double
    
    Dim ObjFrontPlate As Object
    Dim ObjRearPlate As Object
    Dim ObjPlatePack As Object
    Dim ObjSupport1 As Object
    Dim ObjSupport2 As Object
    Dim ObjSupport3 As Object
    Dim ObjSupportClamp1 As Object
    Dim ObjSupportClamp2 As Object
    Dim ObjSupportClamp3 As Object
    Dim objEqpFoundationPort As Object
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parEquipmentXLength = arrayOfInputs(2)
    parEquipmentYLength = arrayOfInputs(3)
    parEquipmentZLength = arrayOfInputs(4)
    parFrontPlateThickness = arrayOfInputs(5)
    parRearPlateThickness = arrayOfInputs(6)
    parPlatePackLength = arrayOfInputs(7)
    parCPtoFaceY = arrayOfInputs(8)
    parCPtoFaceZ = arrayOfInputs(9)
    parSupportLength = arrayOfInputs(10)
    parSupportWidth = arrayOfInputs(11)
    parSupportThickness = arrayOfInputs(12)
    parXboltHole = arrayOfInputs(13)
    parYboltHole = arrayOfInputs(14)
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim dHorSupportLength As Double
    
    dHorSupportLength = parEquipmentXLength - 2 * parXboltHole - _
                                    parFrontPlateThickness - parRearPlateThickness - parPlatePackLength

    Dim botLeftCornerPt As New AutoMath.DPosition
    Dim topRightCornerPt As New AutoMath.DPosition
' Insert your code for outputs 1(ObjFrontPlate)
    botLeftCornerPt.Set 0, -parCPtoFaceY, -parCPtoFaceZ
    topRightCornerPt.Set parFrontPlateThickness, botLeftCornerPt.y + parEquipmentYLength, _
                                            botLeftCornerPt.z + parEquipmentZLength
    Set ObjFrontPlate = PlaceBox(m_OutputColl, botLeftCornerPt, topRightCornerPt)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFrontPlate
    Set ObjFrontPlate = Nothing
    
 'Insert your code for outputs 2(ObjPlatePack)
    botLeftCornerPt.Set parFrontPlateThickness, -parCPtoFaceY, -0.8 * parCPtoFaceZ
    topRightCornerPt.Set botLeftCornerPt.x + parPlatePackLength, botLeftCornerPt.y + parEquipmentYLength, _
                                 parEquipmentZLength - 1.2 * parCPtoFaceZ
    Set ObjPlatePack = PlaceBox(m_OutputColl, botLeftCornerPt, topRightCornerPt)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlatePack
    Set ObjPlatePack = Nothing
    
' Insert your code for outputs 3(ObjRearPlate)
    botLeftCornerPt.Set parFrontPlateThickness + parPlatePackLength, -parCPtoFaceY, -0.8 * parCPtoFaceZ
    topRightCornerPt.Set botLeftCornerPt.x + parRearPlateThickness, botLeftCornerPt.y + parEquipmentYLength, _
                                            parEquipmentZLength - 1.2 * parCPtoFaceZ
    Set ObjRearPlate = PlaceBox(m_OutputColl, botLeftCornerPt, topRightCornerPt)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRearPlate
    Set ObjRearPlate = Nothing
    Set botLeftCornerPt = Nothing
    Set topRightCornerPt = Nothing
    
' Insert your code for outputs 4(ObjInternalRod1)
    Dim dRodDiameter As Double
    Dim dRodLength As Double
    Dim dRodZSpacing As Double
    Dim dextrRodProj As Double
'Assumption
'    dRodDiameter = 0.8 * parFrontPlateThickness
'    dRodLength = 0.7 * parEquipmentXLength
'    dextrRodProj=dRodDiameter           are taken
    dRodDiameter = 0.8 * parFrontPlateThickness
    dRodLength = 0.7 * parEquipmentXLength
    dextrRodProj = dRodDiameter
    
    dRodZSpacing = (parEquipmentZLength - parCPtoFaceZ) / 4
    
    oStPoint.Set -dextrRodProj, -parCPtoFaceY + dRodDiameter, dRodZSpacing
    oEnPoint.Set dRodLength, -parCPtoFaceY + dRodDiameter, dRodZSpacing
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dRodDiameter
    
' Insert your code for outputs 5(ObjInternalRod2)
    oStPoint.Set -dextrRodProj, -parCPtoFaceY + parEquipmentYLength - dRodDiameter, dRodZSpacing
    oEnPoint.Set dRodLength, -parCPtoFaceY + parEquipmentYLength - dRodDiameter, dRodZSpacing
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dRodDiameter
    
' Insert your code for outputs 6(ObjInternalRod3)
    oStPoint.Set -dextrRodProj, -parCPtoFaceY + dRodDiameter, 2 * dRodZSpacing
    oEnPoint.Set dRodLength, -parCPtoFaceY + dRodDiameter, 2 * dRodZSpacing
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dRodDiameter
    
' Insert your code for outputs 7(ObjInternalRod4)
    oStPoint.Set -dextrRodProj, -parCPtoFaceY + parEquipmentYLength - dRodDiameter, 2 * dRodZSpacing
    oEnPoint.Set dRodLength, -parCPtoFaceY + parEquipmentYLength - dRodDiameter, 2 * dRodZSpacing
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dRodDiameter
    
' Insert your code for outputs 8(ObjInternalRod5)
    oStPoint.Set -dextrRodProj, -parCPtoFaceY + dRodDiameter, 3 * dRodZSpacing
    oEnPoint.Set dRodLength, -parCPtoFaceY + dRodDiameter, 3 * dRodZSpacing
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dRodDiameter
    
' Insert your code for outputs 9(ObjInternalRod6)
    oStPoint.Set -dextrRodProj, -parCPtoFaceY + parEquipmentYLength - dRodDiameter, 3 * dRodZSpacing
    oEnPoint.Set dRodLength, -parCPtoFaceY + parEquipmentYLength - dRodDiameter, 3 * dRodZSpacing
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dRodDiameter
    
' Insert your code for outputs 10(ObjSupport1)
    Dim dBackSupportWidth As Double
'Assumption
    'All three back supports are considered to be Square in Shape.
    'dBackSupportWidth = parFrontPlateThickness
    
    dBackSupportWidth = parFrontPlateThickness
    oStPoint.Set (parFrontPlateThickness + parPlatePackLength + parRearPlateThickness), _
                            -parCPtoFaceY + parEquipmentYLength / 2 - dBackSupportWidth / 2, _
                            -parCPtoFaceZ / 2 - dBackSupportWidth / 2
    oEnPoint.Set oStPoint.x + dHorSupportLength - dBackSupportWidth, oStPoint.y + dBackSupportWidth, _
                            oStPoint.z + dBackSupportWidth
    Set ObjSupport1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport1
    Set ObjSupport1 = Nothing
    
' Insert your code for outputs 11(ObjSupport2)
    oStPoint.Set (parFrontPlateThickness + parPlatePackLength + parRearPlateThickness), _
                            -parCPtoFaceY + parEquipmentYLength / 2 - dBackSupportWidth / 2, _
                            parEquipmentZLength - 1.5 * parCPtoFaceZ - dBackSupportWidth / 2
    oEnPoint.Set oStPoint.x + dHorSupportLength - dBackSupportWidth, oStPoint.y + dBackSupportWidth, _
                            oStPoint.z + dBackSupportWidth
    Set ObjSupport2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport2
    Set ObjSupport2 = Nothing

' Insert your code for outputs 12(ObjSupport3)
'Create the Bottom Rectangle of the Transition Piece using line string
    Dim obaseRect As IngrGeom3D.LineString3d
    'Create Line String
    
    Dim lineStrPts(0 To 26) As Double
    'point 1
    lineStrPts(0) = parFrontPlateThickness + parPlatePackLength + parRearPlateThickness + dHorSupportLength
    lineStrPts(1) = -parCPtoFaceY + parEquipmentYLength / 2 - parSupportWidth / 2 - 0.1 * dBackSupportWidth
    lineStrPts(2) = parEquipmentZLength - 1.5 * parCPtoFaceZ + dBackSupportWidth / 2
    
     'point  2
    lineStrPts(3) = lineStrPts(0) - dBackSupportWidth
    lineStrPts(4) = lineStrPts(1)
    lineStrPts(5) = lineStrPts(2)
    
     'point3
    lineStrPts(6) = lineStrPts(3)
    lineStrPts(7) = -parCPtoFaceY + parEquipmentYLength / 2 + parSupportWidth / 2 + 0.1 * dBackSupportWidth
    lineStrPts(8) = lineStrPts(2)
    
     'point4
    lineStrPts(9) = lineStrPts(0)
    lineStrPts(10) = lineStrPts(7)
    lineStrPts(11) = lineStrPts(2)
    
     'point5
    lineStrPts(12) = lineStrPts(0)
    lineStrPts(13) = -parCPtoFaceY + parEquipmentYLength / 2 + parSupportWidth / 2
    lineStrPts(14) = lineStrPts(2)
    
     'point6
    lineStrPts(15) = lineStrPts(3) + 0.1 * dBackSupportWidth
    lineStrPts(16) = lineStrPts(13)
    lineStrPts(17) = lineStrPts(2)
    
     'point7
    lineStrPts(18) = lineStrPts(15)
    lineStrPts(19) = -parCPtoFaceY + parEquipmentYLength / 2 - parSupportWidth / 2
    lineStrPts(20) = lineStrPts(2)
    
     'point8
    lineStrPts(21) = parFrontPlateThickness + parPlatePackLength + parRearPlateThickness + dHorSupportLength
    lineStrPts(22) = lineStrPts(19)
    lineStrPts(23) = lineStrPts(2)
    
    'point9
    lineStrPts(24) = lineStrPts(0)
    lineStrPts(25) = lineStrPts(1)
    lineStrPts(26) = lineStrPts(2)
    
    Set obaseRect = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 9, lineStrPts)
    
    Dim oAxisVect As New AutoMath.DVector
    oAxisVect.Set 0, 0, -1
    
    Dim dVertSupportHeight As Double
    dVertSupportHeight = (parEquipmentZLength - 1.5 * parCPtoFaceZ + dBackSupportWidth / 2) + _
                                        parCPtoFaceZ
    Set ObjSupport3 = PlaceProjection(m_OutputColl, obaseRect, oAxisVect, dVertSupportHeight, True)
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport3
    Set ObjSupport3 = Nothing

 'Insert your code for outputs 13(ObjSupportClamp1 )
 
    'point 1
    lineStrPts(0) = -parSupportLength
    lineStrPts(1) = -parCPtoFaceY
    lineStrPts(2) = -parCPtoFaceZ + parSupportThickness
    
     'point  2
    lineStrPts(3) = -parSupportThickness
    lineStrPts(4) = lineStrPts(1)
    lineStrPts(5) = lineStrPts(2)
    
     'point3
    lineStrPts(6) = -parSupportThickness
    lineStrPts(7) = lineStrPts(1)
    lineStrPts(8) = -parCPtoFaceZ + parSupportLength
    
     'point4
    lineStrPts(9) = 0
    lineStrPts(10) = lineStrPts(1)
    lineStrPts(11) = -parCPtoFaceZ + parSupportLength
    
     'point5
    lineStrPts(12) = 0
    lineStrPts(13) = lineStrPts(1)
    lineStrPts(14) = -parCPtoFaceZ
    
     'point6
    lineStrPts(15) = -parSupportLength
    lineStrPts(16) = lineStrPts(1)
    lineStrPts(17) = -parCPtoFaceZ
    
     'point7
    lineStrPts(18) = lineStrPts(0)
    lineStrPts(19) = lineStrPts(1)
    lineStrPts(20) = lineStrPts(2)
    
    Set obaseRect = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, lineStrPts)

    oAxisVect.Set 0, 1, 0
    Set ObjSupportClamp1 = PlaceProjection(m_OutputColl, obaseRect, oAxisVect, parSupportWidth, True)
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupportClamp1
    Set ObjSupportClamp1 = Nothing
    
'Insert your code for outputs 14(ObjSupportClamp2 )
    ' The line string 3 D created for the clamp1 is translated along y-Dir to use it for projection for 2nd clamp.
    Dim translateMat     As New AutoMath.DT4x4   'For complexstring translation
    Dim oTransVec As New AutoMath.DVector
    oTransVec.Set 0, 1, 0
    oTransVec.Length = parEquipmentYLength - parSupportWidth
    translateMat.LoadIdentity
    translateMat.Translate oTransVec
    
    'Transform the complex string to the new location
    obaseRect.Transform translateMat
    Set ObjSupportClamp2 = PlaceProjection(m_OutputColl, obaseRect, oAxisVect, parSupportWidth, True)
    
     ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupportClamp2
    Set ObjSupportClamp2 = Nothing
    
    Set oTransVec = Nothing
    Set translateMat = Nothing

'Insert your code for outputs 15(ObjSupportClamp3 )
    'point 1
    lineStrPts(0) = parFrontPlateThickness + parPlatePackLength + parRearPlateThickness + dHorSupportLength + _
                                -0.9 * dBackSupportWidth + parSupportLength
    lineStrPts(1) = -parCPtoFaceY + parEquipmentYLength / 2 - parSupportWidth / 2
    lineStrPts(2) = -parCPtoFaceZ + parSupportThickness
    
     'point  2
    lineStrPts(3) = lineStrPts(0) - parSupportLength + parSupportThickness
    lineStrPts(4) = lineStrPts(1)
    lineStrPts(5) = lineStrPts(2)
    
     'point3
    lineStrPts(6) = lineStrPts(3)
    lineStrPts(7) = lineStrPts(1)
    lineStrPts(8) = -parCPtoFaceZ + parSupportLength
    
     'point4
    lineStrPts(9) = parFrontPlateThickness + parPlatePackLength + parRearPlateThickness + dHorSupportLength + _
                                -0.9 * dBackSupportWidth
    lineStrPts(10) = lineStrPts(1)
    lineStrPts(11) = lineStrPts(8)
    
     'point5
    lineStrPts(12) = lineStrPts(9)
    lineStrPts(13) = lineStrPts(1)
    lineStrPts(14) = -parCPtoFaceZ
    
     'point6
    lineStrPts(15) = lineStrPts(9) + parSupportLength
    lineStrPts(16) = lineStrPts(1)
    lineStrPts(17) = -parCPtoFaceZ
    
     'point7
    lineStrPts(18) = lineStrPts(0)
    lineStrPts(19) = lineStrPts(1)
    lineStrPts(20) = lineStrPts(2)
    
    Set obaseRect = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, lineStrPts)

    oAxisVect.Set 0, 1, 0

    Set ObjSupportClamp3 = PlaceProjection(m_OutputColl, obaseRect, oAxisVect, parSupportWidth, True)
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupportClamp3
    Set ObjSupportClamp3 = Nothing
    
    Set oAxisVect = Nothing
    Set obaseRect = Nothing

   
'==========================================
'Construction of Equipment Foundatiopn Port
'==========================================
    Dim objExchangerFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory   As New NozzlePHFactory
    Set objExchangerFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "EqpFoundationPort", _
                                                False, m_OutputColl.ResourceManager)
    Dim holes() As Variant
    Call objExchangerFoundationPort.GetHoles(holes())
 
    holes(0, 1) = -parCPtoFaceY + parEquipmentYLength / 2 - parYboltHole / 2
    holes(0, 2) = -parXboltHole
    holes(1, 1) = -parCPtoFaceY + parEquipmentYLength / 2 + parYboltHole / 2
    holes(1, 2) = -parXboltHole
    holes(2, 1) = -parCPtoFaceY + parEquipmentYLength / 2
    holes(2, 2) = parFrontPlateThickness + parPlatePackLength + parRearPlateThickness + dHorSupportLength _
                                - 0.9 * dBackSupportWidth + parXboltHole

    Call objExchangerFoundationPort.PutCS(0, 0, -parCPtoFaceZ, 0, 1, 0, 0, 0, -1)
    Call objExchangerFoundationPort.SetHoles(holes)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objExchangerFoundationPort
    Set objExchangerFoundationPort = Nothing

   
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oAxisVect = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
