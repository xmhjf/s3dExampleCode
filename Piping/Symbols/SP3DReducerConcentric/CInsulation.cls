VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          NN
'   Creation Date:  Saturday, Nov 4 2000
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols.
'  9.Feb.2007      dkl  CR 112819 Implemented Part data basis to modify symbol for HDPE spec.
' 27.June.2007     MA CR-121834  Implemented Part data basis for value 377
'  17.Dec.2007     KKC   CR-121408  Implemented the part data basis value 447
'  7.Jan.2008     svsmylav CR-131511  Implemented the part data basis values 620,621,623,625,626,628,630,632,634,636.
'                                     Added EndLength1 and EndLength2 for part data basis value 447.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Sub Class_Initialize()
'''
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim CptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim pipeDiam2       As Double
    Dim flangeThick2    As Double
    Dim CptOffset2      As Double
    Dim flangeDiam2     As Double
    Dim depth2          As Double
    Dim parInsulationThickness As Double
    Dim parEndLength1 As Double
    Dim parEndLength2 As Double
    Dim parTangentLength As Double
    Dim parReducerLength As Double
    Dim parSeattoSeat As Double
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
'   parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
'   parFace1toCenter = arrayOfInputs(4)
'   parFace2toCenter = arrayOfInputs(5)
'   parLayingLength = arrayOfInputs(6)
'   parEndLength = arrayOfInputs(7)
'   ParLeftFlatEndLength = arrayOfInputs(8)
'   ParRightFlatEndLength = arrayOfInputs(9)
'    parEndLength1 = arrayOfInputs(10)
'    parEndLength2 = arrayOfInputs(11)
'    parTangentLength = arrayOfInputs(12)
'    parReducerLength = arrayOfInputs(13)
'    parSeattoSeat = arrayOfInputs(14)
    
'   Local variables
    Dim dReducerLength As Double
    Dim dReductionRatio As Double
    Dim dTangentLength As Double
    Dim dFacetoFace As Double
    Dim dNozzleLength1 As Double
    Dim dNozzleLength2 As Double
    dReductionRatio = 0.15 'Reducer length is 15% of face-to-face
    
  RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                    CptOffset, depth
  RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                    CptOffset2, depth2

    Dim iOutput     As Long
    Dim objInsulatedBody As Object
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
       
    iOutput = 0
    
    Dim LineStrPoints(0 To 20)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim oAxisVect As AutoMath.DVector
    Dim oCenPoint As AutoMath.DPosition
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    Set oEndPoint = New AutoMath.DPosition
    Dim PI
    
' This symbol is based on the following Part data basis values that govern its geometry,
' 'Concentric reducer, standard' - 67 or 1 and 'Concentric reducer, stepped' - 70.

 ' Insert your code for output 2(Insulated Body)
    Select Case lPartDataBasis
        Case Is <= 1, 67
            Dim parFacetoFace   As Double
            parFacetoFace = arrayOfInputs(2)
            If CmpDblLessThanOrEqualTo(depth, 0) Then
               oStPoint.Set -(parFacetoFace / 2 - flangeThick), 0, 0
               oEndPoint.Set parFacetoFace / 2 - flangeThick2, 0, 0
        
               Set objInsulatedBody = PlaceCone(m_OutputColl, oStPoint, oEndPoint, _
                pipeDiam / 2 + parInsulationThickness, pipeDiam2 / 2 + parInsulationThickness)
            
            Else
               oStPoint.Set -parFacetoFace / 2, 0, 0
               oEndPoint.Set parFacetoFace / 2, 0, 0
               Dim dInsulationDiameter As Double
               dInsulationDiameter = InsulationDiameter(flangeDiam, flangeDiam2, _
                                                                parInsulationThickness)
                                                                
               Set objInsulatedBody = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, _
                                                        dInsulationDiameter, True)
            End If
            
            Set oStPoint = Nothing
            Set oEndPoint = Nothing
    
        Case 70
            Dim parFace1toCenter As Double
            Dim parFace2toCenter As Double
            parFace1toCenter = arrayOfInputs(4)
            parFace2toCenter = arrayOfInputs(5)
            Set oAxisVect = New AutoMath.DVector
            Set oCenPoint = New AutoMath.DPosition
            
            'Point 1
            LineStrPoints(0) = -parFace1toCenter
            If CmpDblLessThanOrEqualTo(depth, 0) Then 'if end is not female end.
                LineStrPoints(1) = pipeDiam / 2 + parInsulationThickness
            Else 'if end is female end, consider the flange/socket diameter.
                LineStrPoints(1) = flangeDiam / 2 + parInsulationThickness
            End If
            LineStrPoints(2) = 0
            
            'Point 2, the end of the larger cylinder.
            LineStrPoints(3) = 0
            LineStrPoints(4) = LineStrPoints(1)
            LineStrPoints(5) = LineStrPoints(2)
            
            'Point 3, the begining point of the smaller cylinder.
            LineStrPoints(6) = LineStrPoints(3)
            If CmpDblLessThanOrEqualTo(depth, 0) Then 'if end is not female end.
                LineStrPoints(7) = pipeDiam2 / 2 + parInsulationThickness
            Else 'if end is female end, consider the flange/socket diameter.
                LineStrPoints(7) = flangeDiam2 / 2 + parInsulationThickness
            End If
            LineStrPoints(8) = LineStrPoints(2)
            
            'Point 4,the end of the smaller cylinder.
            LineStrPoints(9) = parFace2toCenter
            LineStrPoints(10) = LineStrPoints(7)
            LineStrPoints(11) = LineStrPoints(2)
            
            'Point 5
            LineStrPoints(12) = LineStrPoints(9)
            LineStrPoints(13) = 0
            LineStrPoints(14) = LineStrPoints(2)
                
            'Point 6
            LineStrPoints(15) = LineStrPoints(0)
            LineStrPoints(16) = LineStrPoints(13)
            LineStrPoints(17) = LineStrPoints(2)
            
            'Point 7
            LineStrPoints(18) = LineStrPoints(0)
            LineStrPoints(19) = LineStrPoints(1)
            LineStrPoints(20) = LineStrPoints(2)
            
            Set geomFactory = New IngrGeom3D.GeometryFactory
            Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, _
                                                                    LineStrPoints)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 1, 0, 0
            oCenPoint.Set 0, -0.00001, 0
            PI = 4 * Atn(1)
            
            Set objInsulatedBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                            oCenPoint, 2 * PI, True)
                                    
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set geomFactory = Nothing
    
        Case 377
            Dim parLayingLength As Double
            Dim parEndLength As Double
            parLayingLength = arrayOfInputs(6)
            parEndLength = arrayOfInputs(7)
            
            Set oAxisVect = New AutoMath.DVector
            Set oCenPoint = New AutoMath.DPosition
            
            'Point 1
            LineStrPoints(0) = -0.9 * parLayingLength
            If CmpDblLessThanOrEqualTo(depth, 0) Then 'if end is not female end.
                LineStrPoints(1) = pipeDiam / 2 + parInsulationThickness
            Else 'if end is female end, consider the flange/socket diameter.
                LineStrPoints(1) = flangeDiam / 2 + parInsulationThickness
            End If
            LineStrPoints(2) = 0
            
            'Point 2
            LineStrPoints(3) = -0.1 * parLayingLength
            LineStrPoints(4) = LineStrPoints(1)
            LineStrPoints(5) = LineStrPoints(2)
            
            'Point 3
            LineStrPoints(6) = 0.1 * parLayingLength
            If CmpDblLessThanOrEqualTo(depth, 0) Then 'if end is not female end.
                LineStrPoints(7) = pipeDiam2 / 2 + parInsulationThickness
            Else 'if end is female end, consider the flange/socket diameter.
                LineStrPoints(7) = flangeDiam2 / 2 + parInsulationThickness
            End If
            LineStrPoints(8) = LineStrPoints(2)
            
            'Point 4
            LineStrPoints(9) = parEndLength + 0.1 * parLayingLength
            LineStrPoints(10) = LineStrPoints(7)
            LineStrPoints(11) = LineStrPoints(2)
            
            'Point 5
            LineStrPoints(12) = LineStrPoints(9)
            LineStrPoints(13) = 0
            LineStrPoints(14) = LineStrPoints(2)
                
            'Point 6
            LineStrPoints(15) = LineStrPoints(0)
            LineStrPoints(16) = LineStrPoints(13)
            LineStrPoints(17) = LineStrPoints(2)
            
            'Point 7
            LineStrPoints(18) = LineStrPoints(0)
            LineStrPoints(19) = LineStrPoints(1)
            LineStrPoints(20) = LineStrPoints(2)
            
            Set geomFactory = New IngrGeom3D.GeometryFactory
            Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 7, _
                                                                    LineStrPoints)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 1, 0, 0
            oCenPoint.Set 0, -0.00001, 0
            PI = 4 * Atn(1)
            
            Set objInsulatedBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                            oCenPoint, 2 * PI, True)
                                    
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set geomFactory = Nothing
            
        Case 447
            Dim ParLeftFlatEndLength As Double
            Dim ParRightFlatEndLength As Double
            
            parFacetoFace = arrayOfInputs(2)
            ParLeftFlatEndLength = arrayOfInputs(8)
            ParRightFlatEndLength = arrayOfInputs(9)
            
            If UBound(arrayOfInputs) >= 10 Then
                ParLeftFlatEndLength = arrayOfInputs(10)
                If UBound(arrayOfInputs) >= 11 Then ParRightFlatEndLength = arrayOfInputs(11)
                'EndLengths will be computed as 15% of the Face-to-Face dimension.
                If CmpDblLessThan(ParLeftFlatEndLength, LINEAR_TOLERANCE) Then
                    ParLeftFlatEndLength = 0.15 * parFacetoFace
                End If
                'User can specify only Endlength 1 - then assign it to Endlength 2
                If CmpDblLessThan(ParRightFlatEndLength, LINEAR_TOLERANCE) Then
                    ParRightFlatEndLength = ParLeftFlatEndLength
                End If
            End If
        
            Dim dReducingLength As Double
            dReducingLength = parFacetoFace - ParLeftFlatEndLength - ParRightFlatEndLength
            'Point 1
            LineStrPoints(0) = -parFacetoFace / 2
            LineStrPoints(1) = 0
            LineStrPoints(2) = 0
            
            'Point 2
            LineStrPoints(3) = -parFacetoFace / 2
            LineStrPoints(4) = pipeDiam / 2 + parInsulationThickness
            LineStrPoints(5) = LineStrPoints(2)
            
            'Point 3
            LineStrPoints(6) = -parFacetoFace / 2 + ParLeftFlatEndLength
            LineStrPoints(7) = pipeDiam / 2 + parInsulationThickness
            LineStrPoints(8) = LineStrPoints(2)
            
            'Point 4
            LineStrPoints(9) = parFacetoFace / 2 - ParRightFlatEndLength
            LineStrPoints(10) = pipeDiam2 / 2 + parInsulationThickness
            LineStrPoints(11) = LineStrPoints(2)
            
            'Point 5
            LineStrPoints(12) = parFacetoFace / 2
            LineStrPoints(13) = pipeDiam2 / 2 + parInsulationThickness
            LineStrPoints(14) = LineStrPoints(2)
                
            'Point 6
            LineStrPoints(15) = parFacetoFace / 2
            LineStrPoints(16) = 0
            LineStrPoints(17) = LineStrPoints(2)
            
            'Point 7
            LineStrPoints(18) = LineStrPoints(0)
            LineStrPoints(19) = LineStrPoints(1)
            LineStrPoints(20) = LineStrPoints(2)
        
            Set geomFactory = New IngrGeom3D.GeometryFactory
            Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                                    7, LineStrPoints)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, -0.00001, 0
            PI = 4 * Atn(1)
            
            Set objInsulatedBody = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)
                                    
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set geomFactory = Nothing
            
       Case 620
            parFacetoFace = arrayOfInputs(2)
            parTangentLength = arrayOfInputs(12)
            dFacetoFace = parFacetoFace
'           Assumption: use 40 available length as reducer length. Accept only if it is not negative.
            If CmpDblLessThanOrEqualTo(dFacetoFace, (parTangentLength + flangeThick2)) Then GoTo ErrorLabel
            'Compute Reducer Length, Nozzle Length 1 and Nozzle Length 2
            dNozzleLength1 = parTangentLength
            dReducerLength = (dFacetoFace - parTangentLength - flangeThick2) * 0.4
            dNozzleLength2 = flangeThick2 + (dFacetoFace - parTangentLength - flangeThick2) * 0.6
            'Points for cone
            oStPoint.Set -0.5 * dFacetoFace + parTangentLength, 0, 0
            oEndPoint.Set oStPoint.x + dReducerLength, 0, 0
            
       Case 621
            parTangentLength = arrayOfInputs(12)
            parSeattoSeat = arrayOfInputs(14)
            dFacetoFace = parSeattoSeat + depth + depth2 - CptOffset - CptOffset2
'           Assumption: use 50 available length as reducer length. Accept only if it is not negative.
            If CmpDblLessThanOrEqualTo(dFacetoFace, (parTangentLength + flangeThick2)) Then GoTo ErrorLabel
            'Compute Reducer Length, Nozzle Length 1 and Nozzle Length 2
            dNozzleLength1 = parTangentLength
            dReducerLength = (dFacetoFace - parTangentLength - flangeThick2) * 0.4
            dNozzleLength2 = flangeThick2 + (dFacetoFace - parTangentLength - flangeThick2) * 0.6
            'Points for cone
            oStPoint.Set -0.5 * dFacetoFace + parTangentLength, 0, 0
            oEndPoint.Set oStPoint.x + dReducerLength, 0, 0
            
       Case 623
            parSeattoSeat = arrayOfInputs(14)
            'Assumption: Nozzle Length 2 is taken as 1.1 times flangeThick2
            dNozzleLength2 = flangeThick2 * 1.1
            dFacetoFace = parSeattoSeat + depth + depth2 - CptOffset - CptOffset2
            dReducerLength = dReductionRatio * dFacetoFace
            oStPoint.Set dFacetoFace / 2 - dNozzleLength2 - dReducerLength, 0, 0
            oEndPoint.Set oStPoint.x + dReducerLength, 0, 0
            dNozzleLength1 = dFacetoFace - dNozzleLength2 - dReducerLength
            
       Case 625
            parFacetoFace = arrayOfInputs(2)
            parTangentLength = arrayOfInputs(12)
            dFacetoFace = parFacetoFace
            'Compute Reducer Length, Nozzle Length 1 and Nozzle Length 2
            dNozzleLength1 = parTangentLength
            dNozzleLength2 = flangeThick2
            dReducerLength = dFacetoFace - dNozzleLength1 - dNozzleLength2
            oStPoint.Set dFacetoFace / 2 - dNozzleLength2 - dReducerLength, 0, 0
            oEndPoint.Set oStPoint.x + dReducerLength, 0, 0

       Case 626
            parTangentLength = arrayOfInputs(12)
            parSeattoSeat = arrayOfInputs(14)
            dFacetoFace = parSeattoSeat + depth + depth2 - CptOffset - CptOffset2
            'Compute Reducer Length, Nozzle Length 1 and Nozzle Length 2
            dNozzleLength1 = parTangentLength
            dNozzleLength2 = flangeThick2
            dReducerLength = dFacetoFace - dNozzleLength1 - dNozzleLength2
            oStPoint.Set dFacetoFace / 2 - dNozzleLength2 - dReducerLength, 0, 0
            oEndPoint.Set oStPoint.x + dReducerLength, 0, 0
       Case 628
            'Close to PDB 632 - but wanted current approach helps fine tuning graphics of 628 case
            parSeattoSeat = arrayOfInputs(14)
            dFacetoFace = parSeattoSeat + depth + depth2 - CptOffset - CptOffset2
            'Assumption: dReducerLength = 30% of dFacetoFace
            dReducerLength = dFacetoFace * 0.3
            dNozzleLength2 = flangeThick2
            'Compute Reducer Length, Nozzle Length 1 and Nozzle Length 2
            dNozzleLength1 = dFacetoFace - dNozzleLength2 - dReducerLength
            oStPoint.Set dFacetoFace / 2 - dNozzleLength2 - dReducerLength, 0, 0
            oEndPoint.Set oStPoint.x + dReducerLength, 0, 0
       Case 630
            'Sanitary increaser, long, specified by face-to-face dimension, tangent length, and reducer length’
            parFacetoFace = arrayOfInputs(2)
            'Applicable for E-32.1f (ASTM A 74, Table 25)
            parTangentLength = arrayOfInputs(12)
            parReducerLength = arrayOfInputs(13)

            dFacetoFace = parFacetoFace
            dTangentLength = parTangentLength
            dReducerLength = parReducerLength

            'For E-32.1g(Tyler Soil Pipe Catalog)
            If CmpDblLessThanOrEqualTo(dTangentLength, LINEAR_TOLERANCE) And _
                CmpDblLessThanOrEqualTo(dReducerLength, LINEAR_TOLERANCE) Then
                'Assumptions: Tangent Length is equal to 25% of dFacetoFace
                dTangentLength = 0.25 * dFacetoFace
                dReducerLength = dReductionRatio * dFacetoFace
            End If
            'Compute Reducer Length, Nozzle Length 1 and Nozzle Length 2
            dNozzleLength1 = parTangentLength
            dNozzleLength2 = dFacetoFace - dNozzleLength1 - dReducerLength
            oStPoint.Set dFacetoFace / 2 - dNozzleLength2 - dReducerLength, 0, 0
            oEndPoint.Set oStPoint.x + dReducerLength, 0, 0

       Case 632
            parSeattoSeat = arrayOfInputs(14)
            dFacetoFace = parSeattoSeat + depth + depth2 - CptOffset - CptOffset2
            'Assumptions and cases are listed below:
            If CmpDblEqual(flangeThick, 0) And CmpDblEqual(flangeThick2, 0) Then
                'Case 1. Flange thickness at port 1 and port 2 are zero
                'Apportion 1/3 of dFacetoFace for tangents and reducer
                dReducerLength = dFacetoFace / 3
                dNozzleLength1 = dReducerLength
                dNozzleLength2 = dFacetoFace - dNozzleLength1 - dReducerLength
            ElseIf CmpDblEqual(flangeThick, 0) Or CmpDblEqual(flangeThick2, 0) Then
                'Case 2. Flange thickness at port 1 OR port 2 are zero
                'Take 50% remaining length as ReducerLength
                dReducerLength = (dFacetoFace - flangeThick - flangeThick2) / 2
                If CmpDblGreaterthan(flangeThick, LINEAR_TOLERANCE) Then
                    dNozzleLength1 = flangeThick
                    dNozzleLength2 = dReducerLength
                End If
                If CmpDblGreaterthan(flangeThick2, LINEAR_TOLERANCE) Then
                    dNozzleLength2 = flangeThick2
                    dNozzleLength1 = dReducerLength
                End If
            Else
                'Case 3. Flange thickness at port 1 and port 2 are non-zero
                dNozzleLength1 = flangeThick
                dNozzleLength2 = flangeThick2
                dReducerLength = dFacetoFace - flangeThick - flangeThick2
            End If
            oStPoint.Set dFacetoFace / 2 - dNozzleLength2 - dReducerLength, 0, 0
            oEndPoint.Set oStPoint.x + dReducerLength, 0, 0
            
       Case 634
            parFacetoFace = arrayOfInputs(2)
            dFacetoFace = parFacetoFace
            'Compute Reducer Length, Nozzle Length 1 and Nozzle Length 2
            dReductionRatio = 0.2
            dReducerLength = dReductionRatio * dFacetoFace
            dNozzleLength1 = (dFacetoFace - dReducerLength) / 2
            dNozzleLength2 = dNozzleLength1
            oStPoint.Set -dReducerLength / 2, 0, 0
            oEndPoint.Set oStPoint.x + dReducerLength, 0, 0

       Case 636
            'Close to PDB 628 - but wanted current approach helps fine tuning graphics of 636 case
            parFacetoFace = arrayOfInputs(2)
            dFacetoFace = parFacetoFace
            'Assumption: Nozzle Length 2 is taken as flangeThick2
            dNozzleLength2 = flangeThick2
            dReductionRatio = 0.2
            dReducerLength = dReductionRatio * dFacetoFace
            oStPoint.Set dFacetoFace / 2 - dNozzleLength2 - dReducerLength, 0, 0
            oEndPoint.Set oStPoint.x + dReducerLength, 0, 0
            dNozzleLength1 = dFacetoFace - dNozzleLength2 - dReducerLength
        Case Else:   GoTo ErrorLabel: ' Invalid Part data basis.
    End Select
    
    Select Case lPartDataBasis
        Case 620, 621, 623, 625, 626, 628, 630, 632, 634, 636
            Dim oStPointPortIns As AutoMath.DPosition
            Dim oEnPointPortIns As AutoMath.DPosition

            Set oStPointPortIns = New AutoMath.DPosition
            Set oEnPointPortIns = New AutoMath.DPosition
            
            'Create insulation at port 1 (Hub/Flange)
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                dInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            Else
                dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            End If
        
            oStPointPortIns.Set -dFacetoFace / 2, 0, 0
            If CmpDblLessThanOrEqualTo(flangeThick, NEGLIGIBLE_THICKNESS) Then
                oEnPointPortIns.Set oStPointPortIns.x + NEGLIGIBLE_THICKNESS, oStPointPortIns.y, oStPointPortIns.z
            Else
                oEnPointPortIns.Set oStPointPortIns.x + flangeThick + parInsulationThickness, _
                        oStPointPortIns.y, oStPointPortIns.z
            End If
            Dim objInsulationAtPort1
            Set objInsulationAtPort1 = PlaceCylinder(m_OutputColl, oStPointPortIns, oEnPointPortIns, _
                                                        dInsulationDiameter, True)
            ' Set the output
            m_OutputColl.AddOutput "InsulationAtPort1", objInsulationAtPort1
            Set objInsulationAtPort1 = Nothing
            
            'Create insulation at port 1 body (Hub/Flange)
            'Insualte only when Tangent length 1 is non-zero.
            If Not CmpDblEqual(oEnPointPortIns.x, oStPoint.x) Then  ' displacement along Y and Z is zero hance comparing along X.
                oStPointPortIns.Set oEnPointPortIns.x, oEnPointPortIns.y, oEnPointPortIns.z
                oEnPointPortIns.Set oStPoint.x, oStPoint.y, oStPoint.z 'Reducer's staring point
            
                Dim objInsulationBodyAtPort1
                Set objInsulationBodyAtPort1 = PlaceCylinder(m_OutputColl, oStPointPortIns, oEnPointPortIns, _
                                                            dInsulationDiameter, True)
                ' Set the output
                m_OutputColl.AddOutput "InsulationBodyAtPort1", objInsulationBodyAtPort1
                Set objInsulationBodyAtPort1 = Nothing
            End If
            
            ' Insert your code for output 2(Insulated Body)
            Dim dInsulationDiameter2 As Double
            If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then
                dInsulationDiameter2 = flangeDiam2 + 2 * parInsulationThickness
            Else
                dInsulationDiameter2 = pipeDiam2 + 2 * parInsulationThickness
            End If
            Set objInsulatedBody = PlaceCone(m_OutputColl, oStPoint, oEndPoint, _
                                dInsulationDiameter / 2, dInsulationDiameter2 / 2, True)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody
            Set objInsulatedBody = Nothing
            
            'Create insulation at port 2 (Hub/Flange)
            oStPointPortIns.Set dFacetoFace / 2, 0, 0
            If CmpDblLessThanOrEqualTo(flangeThick2, NEGLIGIBLE_THICKNESS) Then
                oEnPointPortIns.Set oStPointPortIns.x + NEGLIGIBLE_THICKNESS, oStPointPortIns.y, oStPointPortIns.z
            Else
                oEnPointPortIns.Set oStPointPortIns.x - flangeThick2 - parInsulationThickness, _
                        oStPointPortIns.y, oStPointPortIns.z
            End If
            Dim objInsulationAtPort2
            Set objInsulationAtPort2 = PlaceCylinder(m_OutputColl, oStPointPortIns, oEnPointPortIns, _
                                                        dInsulationDiameter2, True)
            
            m_OutputColl.AddOutput "InsulationAtPort2", objInsulationAtPort2
            Set objInsulationAtPort2 = Nothing
            
            'Create insulation at port 2 body.
            'Prevent failure when Tangent length 2 is zero.
            If Not CmpDblEqual(oEnPointPortIns.x, oEndPoint.x) Then  ' displacement along Y and Z is zero hance comparing along X.
                oStPointPortIns.Set oEnPointPortIns.x, oEnPointPortIns.y, oEnPointPortIns.z
                oEnPointPortIns.Set oEndPoint.x, oEndPoint.y, oEndPoint.z 'Reducer's end point
                        
                Dim objInsulationBodyAtPort2
                Set objInsulationBodyAtPort2 = PlaceCylinder(m_OutputColl, oStPointPortIns, oEnPointPortIns, _
                                                            dInsulationDiameter2, True)
                ' Set the output
                m_OutputColl.AddOutput "InsulationBodyAtPort2", objInsulationBodyAtPort2
                Set objInsulationBodyAtPort2 = Nothing
            End If
            
            Set oStPointPortIns = Nothing
            Set oEnPointPortIns = Nothing
            Set oStPoint = Nothing
            Set oEndPoint = Nothing
            
        Case Else:
            'Handle earlier Part data basis values
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody
            Set objInsulatedBody = Nothing
    End Select
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
