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
'               Manhole Cover Assembly.
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
    
    'To protect the symbol
    If CmpDblGreaterthanOrEqualTo(parCoverWidth, parCoverLength) Then
        parCoverWidth = parCoverLength - 0.0001
    End If
        
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
    Dim dOuterArcRad As Double
    dOuterArcRad = parCoverWidth / 2
    
    'Line 1
    oStPoint.Set parCoverLength / 2 - dOuterArcRad, parCoverWidth / 2, 0
    oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 1
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, -oStPoint.y, oStPoint.z
    oCenPoint.Set -parCoverLength / 2, 0, oEnPoint.z
    Set oArc = PlaceTrArcBy3Pts(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    'Line 2
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 2
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, -oStPoint.y, oStPoint.z
    oCenPoint.Set parCoverLength / 2, 0, oStPoint.z
    Set oArc = PlaceTrArcBy3Pts(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    oStPoint.Set parCoverLength / 2 - dOuterArcRad, parCoverWidth / 2, 0
    Set oOuterCompStr = PlaceTrCString(oStPoint, oCollection)
    
    Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oOuterCompStr)
    
    'Create the Default surface at the bottom of the outer plate
    Dim objDefSurface As Object
    Set objDefSurface = oGeomFactory.Planes3d.CreateByOuterBdry(m_OutputColl.ResourceManager, _
                        oOuterCompStr)
    
    oTransMat.LoadIdentity
    oVector.Set 0, 1, 0
    oTransMat.Rotate PI, oVector
    objDefSurface.Transform oTransMat
    
    oTransMat.LoadIdentity
    oVector.Set 0, 0, -(parCoverThickness + (1 + 1 / 8) * dInch)
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
    'Let the offset between Inner and Outer Complex strings be 3 1/4 Inches
    'Radius of the Arcs in Inner Complex string =
    '                       Radius of the Arcs in Outer Complex String - 3 1/4 Inches
    Dim objInnerPlate As Object
    Dim dInnerArcRad As Double
    Dim dOffset As Double       'Offset between Inner and Outer Complex strings
    Dim dLengthTemp As Double   'Length of the Inner Complex String
    Dim dWidthTemp As Double    'Width of the Inner Complex String

    dOffset = (15 / 4) * dInch
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
    oEnPoint.Set oStPoint.x, -oStPoint.y, oStPoint.z
    oCenPoint.Set -dLengthTemp / 2, 0, oEnPoint.z
    Set oArc = PlaceTrArcBy3Pts(oStPoint, oEnPoint, oCenPoint)
    oCollection.Add oArc
    
    'Line 2
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Arc 2
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, -oStPoint.y, oStPoint.z
    oCenPoint.Set dLengthTemp / 2, 0, oStPoint.z
    Set oArc = PlaceTrArcBy3Pts(oStPoint, oEnPoint, oCenPoint)
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
    'Let the Diameter of the Bolt Hole is 1.5 inch
    Dim oHoleCompStr As IngrGeom3D.ComplexString3d
    Dim dHoleRad As Double

    Set oHoleCompStr = New ComplexString3d
    dHoleRad = 1.5 * dInch / 2
    
    'Distance between two holes should not less than the diameter of the hole. If so the holes will overlap which may cause symbol to fail.
    If CmpDblGreaterthanOrEqualTo(2 * dHoleRad, parDistBetBoltHoles) Then
        parDistBetBoltHoles = 2 * dHoleRad + 0.002
    End If
    
    'Arc 1
    oStPoint.Set 0, dHoleRad, 0
    oEnPoint.Set oStPoint.x, -dHoleRad, oStPoint.z
    oCenPoint.Set 0, 0, 0
    oVector.Set 0, 0, 1
    Set oArc = PlaceTrArcByCenterNorm(oStPoint, oEnPoint, oCenPoint, oVector)
    oCollection.Add oArc

    'Arc 2
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set oStPoint.x, dHoleRad, oStPoint.z
    oCenPoint.Set 0, 0, 0
    oVector.Set 0, 0, 1
    Set oArc = PlaceTrArcByCenterNorm(oStPoint, oEnPoint, oCenPoint, oVector)
    oCollection.Add oArc
    
    oStPoint.Set 0, dHoleRad, 0
    Set oHoleCompStr = PlaceTrCString(oStPoint, oCollection)
    
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
    oBoltStPt.Set 0, 0, (1 / 2) * dInch
    oBoltEnPt.Set oBoltStPt.x, oBoltStPt.y, _
                        -(parCoverThickness + ((1 / 8) + 1 + (1 / 2)) * dInch)
                    
    oStPoint.Set 0, dHoleRad, 0

    'Create the Bolts and Holes along the arcs
    dLengthTemp = parCoverLength - dOffset
    dWidthTemp = parCoverWidth - dOffset
    iNoOfBoltHoles = (GetFloorValue((PI * dWidthTemp / 2) / parDistBetBoltHoles)) + 1
    
    Dim dAngle As Double 'Angle made by each bolt at the center of the arc
    dAngle = PI / iNoOfBoltHoles
    iOutput = iOutput + 1
    
    For iCount = 0 To 1
        If (iCount = 1 And CmpDblLessThanOrEqualTo( _
                            (dLengthTemp - dWidthTemp) / 2, 2 * dHoleRad)) Then
            For jCount = 1 To iNoOfBoltHoles - 1
                Set oHoleCompStr = PlaceTrCString(oStPoint, oCollection)
                Set objBolt = PlaceCylinder(m_OutputColl, oBoltStPt, oBoltEnPt, dBoltDia, True)
                
                oTransMat.LoadIdentity
                oVector.Set ((-1) ^ iCount) * _
                        (((dLengthTemp - dWidthTemp) / 2) + (dWidthTemp / 2 * Sin(jCount * dAngle))), _
                        (dWidthTemp / 2 * Cos(jCount * dAngle)), 0
                oTransMat.Translate oVector
                oHoleCompStr.Transform oTransMat
                objBolt.Transform oTransMat
                    
                oPlane.AddHole oHoleCompStr
                m_OutputColl.AddOutput "Bolt_", objBolt
            Next jCount
        Else
            For jCount = 0 To iNoOfBoltHoles
                Set oHoleCompStr = PlaceTrCString(oStPoint, oCollection)
                Set objBolt = PlaceCylinder(m_OutputColl, oBoltStPt, oBoltEnPt, dBoltDia, True)
                
                oTransMat.LoadIdentity
                oVector.Set ((-1) ^ iCount) * _
                        (((dLengthTemp - dWidthTemp) / 2) + (dWidthTemp / 2 * Sin(jCount * dAngle))), _
                        (dWidthTemp / 2 * Cos(jCount * dAngle)), 0
                oTransMat.Translate oVector
                oHoleCompStr.Transform oTransMat
                objBolt.Transform oTransMat
                    
                oPlane.AddHole oHoleCompStr
                m_OutputColl.AddOutput "Bolt_", objBolt
            Next jCount
        End If
    Next iCount
    
    'Create the Holes and bolts along the Length and width of the cover
    iNoOfBoltHoles = (GetFloorValue((parCoverLength - parCoverWidth - 2 * dHoleRad - 2 * parDistBetBoltHoles) / parDistBetBoltHoles)) + 1
    For iCount = 0 To 1
        For jCount = 1 To iNoOfBoltHoles + 1
            Set oHoleCompStr = PlaceTrCString(oStPoint, oCollection)
            Set objBolt = PlaceCylinder(m_OutputColl, oBoltStPt, oBoltEnPt, dBoltDia, True)
            
            oTransMat.LoadIdentity
            oVector.Set -((parCoverLength - parCoverWidth) / 2 - jCount * parDistBetBoltHoles), _
                            ((-1) ^ iCount) * (parCoverWidth / 2 - dOffset / 2), 0
            oTransMat.Translate oVector
            
            oHoleCompStr.Transform oTransMat
            objBolt.Transform oTransMat
            
            m_OutputColl.AddOutput "Bolt_", objBolt
            oPlane.AddHole oHoleCompStr
        Next jCount
    Next iCount
    
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    Set objBolt = Nothing
    Set oBoltStPt = Nothing
    Set oBoltEnPt = Nothing
    Set oHoleCompStr = Nothing

    'Sweep the Plane vertically downwords
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, 0, -(parCoverThickness + (1 / 8 + 1) * dInch)
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
    
    'Create the Bottom part of the cover
    oTransMat.LoadIdentity
    oVector.Set 0, 0, -(parCoverThickness + (1 / 8 + 1) * dInch)
    oTransMat.Translate oVector
    oOuterCompStr.Transform oTransMat
    
    Dim objBotPart As Object
    Dim oLineStr As IngrGeom3D.LineString3d

    Set oLineStr = New LineString3d
    Dim dLineStrPts() As Double
    
    ReDim dLineStrPts(0 To 14) As Double

    'Point 1
    dLineStrPts(0) = parCoverLength / 2 - dOuterArcRad
    dLineStrPts(1) = parCoverWidth / 2
    dLineStrPts(2) = -(parCoverThickness + (1 / 8 + 1) * dInch)
    
    'Point 2
    dLineStrPts(3) = dLineStrPts(0)
    dLineStrPts(4) = (parHoleWidth / 2)
    dLineStrPts(5) = dLineStrPts(2)
    
    'Point 3
    dLineStrPts(6) = dLineStrPts(0)
    dLineStrPts(7) = dLineStrPts(4)
    dLineStrPts(8) = dLineStrPts(2) - parCoverHeight

    'Point 4
    dLineStrPts(9) = dLineStrPts(0)
    dLineStrPts(10) = dLineStrPts(7) + (1 / 2) * dInch
    dLineStrPts(11) = dLineStrPts(8)

    'Point 5
    dLineStrPts(12) = dLineStrPts(0)
    dLineStrPts(13) = dLineStrPts(10)
    dLineStrPts(14) = dLineStrPts(2)
    
    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPts)

    Set oSurfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                            m_OutputColl.ResourceManager, oOuterCompStr, oLineStr, _
                            CircularCorner, 0, stnorm, ednorm, False)

    'Set the Output
    iOutput = iOutput + 1
    For Each objBotPart In oSurfset
        m_OutputColl.AddOutput "BottomPart_", objBotPart
    Next objBotPart

    Set oLineStr = Nothing
    Set objBotPart = Nothing
    For iCount = 1 To oSurfset.Count
        oSurfset.Remove 1
    Next iCount
    
    'Create the Handles
    'Handles are created as per the dimensions in the pdf document.
    Dim objHandle As Object
    Dim oSweepCurve As IngrGeom3D.Circle3d

    Dim dHandleHt As Double
    Dim dHandleLen As Double
    Dim dCSDia As Double
    ReDim dLineStrPts(0 To 14) As Double

    Set oSweepCurve = New Circle3d
    Set oLineStr = New LineString3d

    dHandleHt = 3 * dInch
    dHandleLen = 7 * dInch
    dCSDia = (1 / 2) * dInch

    'Point 1
    dLineStrPts(0) = 0
    dLineStrPts(1) = dHandleLen / 2
    dLineStrPts(2) = 0

    'Point 2
    dLineStrPts(3) = dLineStrPts(0)
    dLineStrPts(4) = dLineStrPts(1)
    dLineStrPts(5) = dHandleHt

    'Point 3
    dLineStrPts(6) = dLineStrPts(0)
    dLineStrPts(7) = -dHandleLen / 2
    dLineStrPts(8) = dLineStrPts(5)

    'Point 4
    dLineStrPts(9) = dLineStrPts(0)
    dLineStrPts(10) = dLineStrPts(7)
    dLineStrPts(11) = 0

    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 4, dLineStrPts)

    oCenPoint.Set 0, dHandleLen / 2 - dCSDia / 2, 0
    oVector.Set 0, 0, 1
    Set oSweepCurve = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                oCenPoint.x, oCenPoint.y, oCenPoint.z, _
                                oVector.x, oVector.y, oVector.z, dCSDia / 2)
    iOutput = iOutput + 1
    For iCount = 1 To 2
        Set oSurfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                m_OutputColl.ResourceManager, oLineStr, oSweepCurve, _
                                CircularCorner, 0, stnorm, ednorm, False)
        
        oTransMat.LoadIdentity
        oVector.Set ((-1) ^ iCount) * ((1 / 1.90625) * parCoverLength / 2), 0, 0 '1.90625 is as per the pdf doucument (1'- 3 1/4" / 8")
        oTransMat.Translate oVector
        
        'Set the Output
        For Each objHandle In oSurfset
            objHandle.Transform oTransMat
            m_OutputColl.AddOutput "Handle_", objHandle
        Next objHandle
        Set objHandle = Nothing
        For jCount = 1 To oSurfset.Count
            oSurfset.Remove 1
        Next jCount
    Next iCount
    
    Set oSweepCurve = Nothing
    Set oLineStr = Nothing
    Set oSurfset = Nothing
    
    'Remove the References
    Set oArc = Nothing
    Set oLine = Nothing
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
