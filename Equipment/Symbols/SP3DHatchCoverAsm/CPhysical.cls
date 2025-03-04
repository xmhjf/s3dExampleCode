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
'   Creation Date:  Thursday, December 6 2007
'
'   Description:
'       CR-131706  Provide manhole and hatch cover equipment symbols.
'               Hatch Cover Assembly.
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Six Outputs:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   06.12.2007      RUK     Created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double
Private Const dInch = 2.54 / 100

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    PI = 4 * Atn(1)
     
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parCoverLength As Double
    Dim parCoverWidth As Double
    Dim parCoverThickness As Double
    Dim parCoverHeight As Double
    Dim parDistBetBoltHoles As Double
    Dim parHoleLength As Double
    Dim parHoleWidth As Double
    Dim parHoleThickness As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCoverLength = arrayOfInputs(2)
    parCoverWidth = arrayOfInputs(3)
    parCoverThickness = arrayOfInputs(4)
    parCoverHeight = arrayOfInputs(5)
    parDistBetBoltHoles = arrayOfInputs(6)
    parHoleLength = arrayOfInputs(7)
    parHoleWidth = arrayOfInputs(8)
    parHoleThickness = arrayOfInputs(9)
    
    'Origin is at DP1/PP1
    iOutput = 0

    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Dim oCenPoint As AutoMath.DPosition
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oVector As AutoMath.DVector
    Dim oTransMat As DT4x4
    Dim iCount As Integer
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    Set oCenPoint = New DPosition
    Set oVector = New DVector
    Set oTransMat = New DT4x4
    
    'Create the Hatch Cover
    Dim oOuterCompStr As IngrGeom3D.ComplexString3d
    Dim oInnerCompStr As IngrGeom3D.ComplexString3d
    Dim oCollection As Collection
    Dim oPlane As IngrGeom3D.Plane3d
    Dim oLine As IngrGeom3D.Line3d
    Dim oArc As IngrGeom3D.Arc3d
    
    Set oOuterCompStr = New ComplexString3d
    Set oInnerCompStr = New ComplexString3d
    Set oCollection = New Collection
    Set oPlane = New Plane3d
    Set oLine = New Line3d
    Set oArc = New Arc3d
    
    'Create the Outer Complex String
    'Assume that Radius of the arc in outer string  should greater than 6 Inches
    Dim dOuterArcRad As Double
    If CmpDblGreaterthan(parCoverWidth, parCoverLength) Then
        dOuterArcRad = (12 + 5 * parCoverLength) * dInch / 2
    Else
        dOuterArcRad = (12 + 5 * parCoverWidth) * dInch / 2
    End If
    'Line 1
    oStPoint.Set parCoverLength / 2 - dOuterArcRad, parCoverWidth / 2, 0
    oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 1
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set -parCoverLength / 2, (parCoverWidth / 2 - dOuterArcRad), oStPoint.z
    oCenPoint.Set oStPoint.x, oEnPoint.y, oEnPoint.z
    Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    'Line 2
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, -(parCoverWidth / 2 - dOuterArcRad), oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 2
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set -(parCoverLength / 2 - dOuterArcRad), -parCoverWidth / 2, oStPoint.z
    oCenPoint.Set oEnPoint.x, oStPoint.y, oStPoint.z
    Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    'Line 3
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set parCoverLength / 2 - dOuterArcRad, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 3
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set parCoverLength / 2, -(parCoverWidth / 2 - dOuterArcRad), oStPoint.z
    oCenPoint.Set oStPoint.x, oEnPoint.y, oEnPoint.z
    Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    'Line 4
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, (parCoverWidth / 2 - dOuterArcRad), oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 4
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set parCoverLength / 2 - dOuterArcRad, parCoverWidth / 2, oStPoint.z
    oCenPoint.Set oEnPoint.x, oStPoint.y, oStPoint.z
    Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    oStPoint.Set parCoverLength / 2 - dOuterArcRad, parCoverWidth / 2, 0
    Set oOuterCompStr = PlaceTrCString(oStPoint, oCollection)
    
    Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oOuterCompStr)
        
    'Create the Default surface at the bottom of the top plate
    Dim objDefSurface As Object
    Set objDefSurface = oGeomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, _
                                                        oOuterCompStr)
    oTransMat.LoadIdentity
    oVector.Set 0, 1, 0
    oTransMat.Rotate PI, oVector
    objDefSurface.Transform oTransMat
    
    oTransMat.LoadIdentity
    oVector.Set 0, 0, -2 * parCoverThickness
    oTransMat.Translate oVector
    objDefSurface.Transform oTransMat
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objDefSurface
    Set objDefSurface = Nothing
    
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    
    'Create the Inner Plate
    'Create the Ineer Complex String
    'Let the offset between Inner and Outer Complex strings be 6 Inches
    'Radius of the Arcs in Inner Complex string =
    '                       Radius of the Arcs in Outer Complex String - 6 Inches
    Dim objInnerPlate As Object
    Dim dInnerArcRad As Double
    Dim dOffset As Double       'Offset between Inner and Outer Complex strings
    Dim dLengthTemp As Double   'Length of the Inner Complex String
    Dim dWidthTemp As Double    'Width of the Inner Complex String
    
    dOffset = 6 * dInch
    dInnerArcRad = dOuterArcRad - dOffset
    dLengthTemp = parCoverLength - 2 * dOffset
    dWidthTemp = parCoverWidth - 2 * dOffset
    
    'Line 1
    oStPoint.Set dLengthTemp / 2 - dInnerArcRad, dWidthTemp / 2, 0
    oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 1
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set -dLengthTemp / 2, dWidthTemp / 2 - dInnerArcRad, oStPoint.z
    oCenPoint.Set oStPoint.x, oEnPoint.y, oEnPoint.z
    Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    'Line 2
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, -(dWidthTemp / 2 - dInnerArcRad), oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 2
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set -(dLengthTemp / 2 - dInnerArcRad), -dWidthTemp / 2, oStPoint.z
    oCenPoint.Set oEnPoint.x, oStPoint.y, oStPoint.z
    Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    'Line 3
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set dLengthTemp / 2 - dInnerArcRad, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 3
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set dLengthTemp / 2, -(dWidthTemp / 2 - dInnerArcRad), oStPoint.z
    oCenPoint.Set oStPoint.x, oEnPoint.y, oEnPoint.z
    Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    'Line 4
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, (dWidthTemp / 2 - dInnerArcRad), oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 4
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set dLengthTemp / 2 - dInnerArcRad, dWidthTemp / 2, oStPoint.z
    oCenPoint.Set oEnPoint.x, oStPoint.y, oStPoint.z
    Set oArc = PlaceTrArcByCenter(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    oStPoint.Set dLengthTemp / 2 - dInnerArcRad, dWidthTemp / 2, 0
    Set oInnerCompStr = PlaceTrCString(oStPoint, oCollection)
    
    oVector.Set 0, 0, -1
    Set objInnerPlate = PlaceProjection(m_OutputColl, oInnerCompStr, oVector, parCoverThickness, True)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInnerPlate
    Set objInnerPlate = Nothing
    
    oPlane.AddHole oInnerCompStr
    
    Set oInnerCompStr = Nothing
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    
    'Create the Bolts and Bolt Holes on the Cover
    'Let the Diameter of the Bolt Hole is 2.5 inch
    Dim oHoleCompStr As IngrGeom3D.ComplexString3d
    Dim dHoleRad As Double
    
    Set oHoleCompStr = New ComplexString3d
    dHoleRad = 2.5 * dInch / 2
    
    'Arc 1
    oStPoint.Set 0, parCoverWidth / 2 - dOffset / 2 + dHoleRad, 0
    oEnPoint.Set oStPoint.x, parCoverWidth / 2 - dOffset / 2 - dHoleRad, oStPoint.z
    oCenPoint.Set 0, parCoverWidth / 2 - dOffset / 2, 0
    oVector.Set 0, 0, 1
    Set oArc = PlaceTrArcByCenterNorm(oStPoint, oEnPoint, oCenPoint, oVector)
    oCollection.Add oArc
    
    'Arc 2
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, parCoverWidth / 2 - dOffset / 2 + dHoleRad, oStPoint.z
    oCenPoint.Set 0, parCoverWidth / 2 - dOffset / 2, 0
    oVector.Set 0, 0, 1
    Set oArc = PlaceTrArcByCenterNorm(oStPoint, oEnPoint, oCenPoint, oVector)
    oCollection.Add oArc
    
    'Let the Bolts are M20
    Dim objBolt As Object
    Dim oBoltStPt As AutoMath.DPosition
    Dim oBoltEnPt As AutoMath.DPosition
    Dim dBoltDia As Double
    Dim jCount As Integer
    Dim kCount As Integer
    Dim iNoOfBoltHoles As Integer
    
    Set oBoltStPt = New DPosition
    Set oBoltEnPt = New DPosition
    
    dBoltDia = 0.02
    oBoltStPt.Set 0, 0, -(1 / 8) * dInch
    oBoltEnPt.Set oBoltStPt.x, oBoltStPt.y, _
                    -(2 * parCoverThickness + (3 / 8 + 3 / 16 + 11 / 8 - 1 / 4) * dInch)
    oStPoint.Set 0, dHoleRad, 0
    
    Dim oRotVector As AutoMath.DVector
    Dim dRotAngle As Double
    
    Set oRotVector = New DVector
    oRotVector.Set 0, 0, 1
    dRotAngle = 0
    
    iNoOfBoltHoles = GetFloorValue(((parCoverLength - 2 * dOuterArcRad - 2 * dHoleRad) _
                                            / parDistBetBoltHoles)) + 1
    dLengthTemp = parCoverLength
    dWidthTemp = parCoverWidth
    
    For iCount = 0 To 1
        For jCount = 0 To 1
            For kCount = 0 To iNoOfBoltHoles - 1
                Set oHoleCompStr = PlaceTrCString(oStPoint, oCollection)
                Set objBolt = PlaceCylinder(m_OutputColl, oBoltStPt, oBoltEnPt, dBoltDia, True)
    
                'Transform
                oTransMat.LoadIdentity
                oVector.Set 0, dWidthTemp / 2 - dOffset / 2, 0
                If (iCount Mod 2) Then
                    oVector.Set oVector.y, oVector.x, oVector.z
                End If
                oTransMat.Translate oVector
                oHoleCompStr.Transform oTransMat
                objBolt.Transform oTransMat
                
                'Rotate
                oTransMat.LoadIdentity
                oTransMat.Rotate dRotAngle, oRotVector
                oHoleCompStr.Transform oTransMat
                objBolt.Transform oTransMat
                
                'Transform
                oTransMat.LoadIdentity
                oVector.Set (dLengthTemp / 2 - dOuterArcRad - dHoleRad) - _
                                                kCount * parDistBetBoltHoles, 0, 0
                If (iCount Mod 2) Then
                    oVector.Set oVector.y, oVector.x, oVector.z
                End If
                oTransMat.Translate oVector
                oHoleCompStr.Transform oTransMat
                objBolt.Transform oTransMat
                
                
                oPlane.AddHole oHoleCompStr
                m_OutputColl.AddOutput "Bolt_", objBolt
            Next kCount
            dRotAngle = dRotAngle + PI
        Next jCount
        dLengthTemp = parCoverWidth
        dWidthTemp = parCoverLength
        iNoOfBoltHoles = GetFloorValue(((parCoverWidth - 2 * dOuterArcRad - 2 * dHoleRad) _
                                            / parDistBetBoltHoles)) + 1
    Next iCount
    
    'Create the Bolts and Holes at the corners of the plate
    dLengthTemp = parCoverLength - 12 * dInch
    dWidthTemp = parCoverWidth - 12 * dInch
    oRotVector.Set 0, 0, 1
    iOutput = iOutput + 1
    
    For iCount = 0 To 1
        For jCount = 0 To 1
            Set oHoleCompStr = PlaceTrCString(oStPoint, oCollection)
            Set objBolt = PlaceCylinder(m_OutputColl, oBoltStPt, oBoltEnPt, dBoltDia, True)
            
            oTransMat.LoadIdentity
            oVector.Set ((-1) ^ iCount) * (dLengthTemp / 2 - dInnerArcRad + _
                                                (dInnerArcRad + dOffset / 2) * Cos(PI / 4)), _
                        dWidthTemp / 2 - dInnerArcRad + _
                                                (dInnerArcRad + dOffset / 2) * Sin(PI / 4), 0
            oTransMat.Translate oVector
            oHoleCompStr.Transform oTransMat
            objBolt.Transform oTransMat
            
            oTransMat.LoadIdentity
            oTransMat.Rotate jCount * PI, oRotVector
            oHoleCompStr.Transform oTransMat
            objBolt.Transform oTransMat
            
            oPlane.AddHole oHoleCompStr
            m_OutputColl.AddOutput "Bolt_", objBolt
        Next jCount
    Next iCount
    
    Set objBolt = Nothing
    Set oBoltStPt = Nothing
    Set oBoltEnPt = Nothing
    Set oRotVector = Nothing
    Set oHoleCompStr = Nothing
    
    'Sweep the Plane vertically downwords
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, 0, -2 * parCoverThickness
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    
    Dim stnorm() As Double
    Dim ednorm() As Double
    Dim oOuterPlate As Object
    Dim oSurfset   As IngrGeom3D.IJElements
    
    Set oSurfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                            m_OutputColl.ResourceManager, oLine, oPlane, _
                            CircularCorner, 0, stnorm, ednorm, True)
    'Set the Output
    iOutput = iOutput + 1
    For Each oOuterPlate In oSurfset
        m_OutputColl.AddOutput "OuterPlate_", oOuterPlate
    Next oOuterPlate
    
    Set oOuterPlate = Nothing
    For iCount = 1 To oSurfset.Count
        oSurfset.Remove 1
    Next iCount
    
    'The bottom plate was created as per the dimensions given in the pdf docment
    Dim objBotPlate As Object
    Dim oBotPlStr As IngrGeom3D.LineString3d
    Dim dBotPlPts() As Double
    
    ReDim dBotPlPts(0 To 20) As Double
    Set oBotPlStr = New LineString3d
    
    'Point 1
    dBotPlPts(0) = parCoverLength / 2 - dOuterArcRad
    dBotPlPts(1) = parCoverWidth / 2
    dBotPlPts(2) = -2 * parCoverThickness
    
    'Point 2
    dBotPlPts(3) = dBotPlPts(0)
    dBotPlPts(4) = parCoverWidth / 2 + (1 / 4) * dInch - (75 / 16) * dInch
    dBotPlPts(5) = dBotPlPts(2)
    
    'Point 3
    dBotPlPts(6) = dBotPlPts(0)
    dBotPlPts(7) = dBotPlPts(4)
    dBotPlPts(8) = -(2 * parCoverThickness + (3 / 8 + 3 / 16 + 11 / 8) * dInch)
    
    'Point 4
    dBotPlPts(9) = dBotPlPts(0)
    dBotPlPts(10) = parCoverWidth / 2 + (1 / 4) * dInch
    dBotPlPts(11) = dBotPlPts(8)
    
    'Point 5
    dBotPlPts(12) = dBotPlPts(0)
    dBotPlPts(13) = dBotPlPts(10)
    dBotPlPts(14) = -(2 * parCoverThickness + (3 / 8 + 3 / 16) * dInch)
    
    'Point 6
    dBotPlPts(15) = dBotPlPts(0)
    dBotPlPts(16) = parCoverWidth / 2 + (1 / 4) * dInch - (75 / 16 - 33 / 8) * dInch
    dBotPlPts(17) = dBotPlPts(14)
    
    'Point 7
    dBotPlPts(18) = dBotPlPts(0)
    dBotPlPts(19) = dBotPlPts(16)
    dBotPlPts(20) = dBotPlPts(2)
    
    
    Set oBotPlStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dBotPlPts)
    
    oTransMat.LoadIdentity
    oVector.Set 0, 0, -(2 * parCoverThickness)
    oTransMat.Translate oVector
    oOuterCompStr.Transform oTransMat
    
    Set oSurfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                            m_OutputColl.ResourceManager, oOuterCompStr, oBotPlStr, _
                            CircularCorner, 0, stnorm, ednorm, True)
    'Set the Output
    iOutput = iOutput + 1
    For Each objBotPlate In oSurfset
        m_OutputColl.AddOutput "BottomPlate_", objBotPlate
    Next objBotPlate
    
    Set objBotPlate = Nothing
    Set oBotPlStr = Nothing
    For iCount = 1 To oSurfset.Count
        oSurfset.Remove 1
    Next iCount
        
    'Create the Bottom part of the cover
    Dim objBotPart As Object
    Dim oBotPartStr As IngrGeom3D.LineString3d
    
    Set oBotPartStr = New LineString3d
    ReDim dBotPlPts(0 To 26) As Double
    
    'Point 1
    dBotPlPts(0) = parCoverLength / 2 - dOuterArcRad
    dBotPlPts(1) = parCoverWidth / 2
    dBotPlPts(2) = -2 * parCoverThickness
    
    'Point 2
    dBotPlPts(3) = dBotPlPts(0)
    dBotPlPts(4) = parCoverWidth / 2 - (75 / 16 - 1 / 4 + 13 / 16 + 3 / 8) * dInch
    dBotPlPts(5) = dBotPlPts(2)
    
    'Point 3
    dBotPlPts(6) = dBotPlPts(0)
    dBotPlPts(7) = dBotPlPts(4)
    dBotPlPts(8) = -(2 * parCoverThickness + parCoverHeight - 2 * (1 / 4) * dInch)
    
    'Point 4
    dBotPlPts(9) = dBotPlPts(0)
    If CmpDblGreaterthan(parCoverLength, parCoverWidth) Then
        dBotPlPts(10) = 0
    Else
        dBotPlPts(10) = (parCoverWidth - parCoverLength) / 2
    End If
    dBotPlPts(11) = dBotPlPts(8)
    
    'Point 5
    dBotPlPts(12) = dBotPlPts(0)
    dBotPlPts(13) = dBotPlPts(10)
    dBotPlPts(14) = -(2 * parCoverThickness + parCoverHeight - (1 / 4) * dInch)
    
    'Point 6
    dBotPlPts(15) = dBotPlPts(0)
    dBotPlPts(16) = parCoverWidth / 2 - (75 / 16 - 1 / 4 + 13 / 16 + 3 / 8) * dInch
    dBotPlPts(17) = dBotPlPts(14)
    
    'Point 7
    dBotPlPts(18) = dBotPlPts(0)
    dBotPlPts(19) = dBotPlPts(16)
    dBotPlPts(20) = -(2 * parCoverThickness + parCoverHeight)
    
    'Point 8
    dBotPlPts(21) = dBotPlPts(0)
    dBotPlPts(22) = dBotPlPts(4) + (3 / 8) * dInch
    dBotPlPts(23) = dBotPlPts(20)
    
    'Point 9
    dBotPlPts(24) = dBotPlPts(0)
    dBotPlPts(25) = dBotPlPts(22)
    dBotPlPts(26) = dBotPlPts(2)
    
    Set oBotPartStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 9, dBotPlPts)
    
    Set oSurfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                            m_OutputColl.ResourceManager, oOuterCompStr, oBotPartStr, _
                            CircularCorner, 0, stnorm, ednorm, False)
    
    'Set the Output
    iOutput = iOutput + 1
    For Each objBotPart In oSurfset
        m_OutputColl.AddOutput "BottomPart_", objBotPart
    Next objBotPart
    
    Set oBotPartStr = Nothing
    Set objBotPart = Nothing
    For iCount = 1 To oSurfset.Count
        oSurfset.Remove 1
    Next iCount
    Set oSurfset = Nothing
    
    'Remove the References
    Set oArc = Nothing
    Set oLine = Nothing
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    Set oCollection = Nothing
    Set oPlane = Nothing
    Set oOuterCompStr = Nothing
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenPoint = Nothing
    Set oVector = Nothing
    Set oTransMat = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Private Function GetFloorValue(Number As Double) As Double
    If CmpDblGreaterthan(CInt(Number), Number) Then
        GetFloorValue = CInt(Number) - 1
    Else
        GetFloorValue = CInt(Number)
    End If
End Function
