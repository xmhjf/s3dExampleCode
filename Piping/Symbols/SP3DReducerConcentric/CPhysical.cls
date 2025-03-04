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
'
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  Saturday, Nov 4 2000
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  9.Feb.2007      dkl  CR 112819 Implemented Part data basis to modify symbol for HDPE spec.
' 27.June.2007     MA CR-121834  Implemented Part data basis for value 377
'  17.Dec.2007     KKC   CR-121408  Implemented the part data basis value 447
'  7.Jan.2008     svsmylav CR-131511  Implemented the part data basis values 620,621,623,625,626,628,630,632,634,636.
'                                     Added EndLength1 and EndLength2 for part data basis value 447.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
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
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                CptOffset, depth
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                CptOffset2, depth2
    
    Dim iOutput     As Long
    Dim objReducer As Object
 
    iOutput = 0
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Dim LineStrPoints(0 To 20)  As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Dim oAxisVect As AutoMath.DVector
    Dim oCenPoint As AutoMath.DPosition
    Dim PI As Double

    Dim oPlacePoint1 As AutoMath.DPosition
    Dim oDir1        As AutoMath.DVector
    Dim oPlacePoint2 As AutoMath.DPosition
    Dim oDir2       As AutoMath.DVector
    Dim objNozzle1   As GSCADNozzleEntities.IJDNozzle
    Dim objNozzle2   As GSCADNozzleEntities.IJDNozzle
    Dim oStPoint As AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
    Set oPlacePoint1 = New AutoMath.DPosition
    Set oDir1 = New AutoMath.DVector
    Set oPlacePoint2 = New AutoMath.DPosition
    Set oDir2 = New AutoMath.DVector
    Set oStPoint = New AutoMath.DPosition
    Set oEndPoint = New AutoMath.DPosition
    
'   Local variables
    Dim dReducerLength As Double
    Dim dReductionRatio As Double
    Dim dTangentLength As Double
    Dim dFacetoFace As Double
    Dim dNozzleLength1 As Double
    Dim dNozzleLength2 As Double
    dReductionRatio = 0.15 'Reducer length is 15% of face-to-face
    
' This symbol is based on the following Part data basis values that govern its geometry,
' 'Concentric reducer, standard' - 67 or 1 and 'Concentric reducer, stepped' - 70
' 'Concentric Reducer, Tubing' - 377
' Cocentric Reducer, Flat Ends- 447
    
' Insert your code for output 1(Body of Reducer)
    Select Case lPartDataBasis
    
        Case Is <= 1, 67
            Dim parFacetoFace   As Double
            parFacetoFace = arrayOfInputs(2)
            
            oStPoint.Set -parFacetoFace / 2 + flangeThick, 0, 0
            oEndPoint.Set parFacetoFace / 2 - flangeThick2, 0, 0
            
            Set objReducer = PlaceCone(m_OutputColl, oStPoint, oEndPoint, pipeDiam / 2, _
                                                                     pipeDiam2 / 2)
            
            Set oStPoint = Nothing
            Set oEndPoint = Nothing
            
            oPlacePoint1.Set -parFacetoFace / 2 - CptOffset + depth, 0, 0
            oDir1.Set -1, 0, 0
            Set objNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDir1, oPlacePoint1)
            
            oPlacePoint2.Set parFacetoFace / 2 + CptOffset2 - depth2, 0, 0
            oDir2.Set 1, 0, 0
            Set objNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir2, oPlacePoint2)
    
        Case 70
            Dim parFace1toCenter As Double
            Dim parFace2toCenter As Double
            parFace1toCenter = arrayOfInputs(4)
            parFace2toCenter = arrayOfInputs(5)
            
            'Point 1
            LineStrPoints(0) = -parFace1toCenter + flangeThick
            LineStrPoints(1) = pipeDiam / 2
            LineStrPoints(2) = 0
            
            'Point 2, the end of the larger cylinder.
            LineStrPoints(3) = 0
            LineStrPoints(4) = LineStrPoints(1)
            LineStrPoints(5) = LineStrPoints(2)
            
            'Point 3, the begining point of the smaller cylinder.
            LineStrPoints(6) = LineStrPoints(3)
            LineStrPoints(7) = pipeDiam2 / 2
            LineStrPoints(8) = LineStrPoints(2)
            
            'Point 4,the end of the smaller cylinder.
            LineStrPoints(9) = parFace2toCenter - flangeThick2
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
            Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                                    7, LineStrPoints)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, -0.00001, 0
            PI = 4 * Atn(1)
            
            Set objReducer = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)
                                    
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set geomFactory = Nothing
            
            oPlacePoint1.Set -parFace1toCenter - CptOffset + depth, 0, 0
            oDir1.Set -1, 0, 0
            Set objNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDir1, oPlacePoint1)
            
            oPlacePoint2.Set parFace2toCenter + CptOffset2 - depth2, 0, 0
            oDir2.Set 1, 0, 0
            Set objNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir2, oPlacePoint2)
        
        Case 377
            Dim parLayingLength As Double
            Dim parEndLength As Double
            parLayingLength = arrayOfInputs(6)
            parEndLength = arrayOfInputs(7)
            
            'Point 1
            LineStrPoints(0) = -0.9 * parLayingLength
            LineStrPoints(1) = pipeDiam / 2
            LineStrPoints(2) = 0
            
            'Point 2
            LineStrPoints(3) = -0.1 * parLayingLength
            LineStrPoints(4) = LineStrPoints(1)
            LineStrPoints(5) = LineStrPoints(2)
            
            'Point 3
            LineStrPoints(6) = 0.1 * parLayingLength
            LineStrPoints(7) = pipeDiam2 / 2
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
            Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                                    7, LineStrPoints)
            Set oAxisVect = New AutoMath.DVector
            oAxisVect.Set 1, 0, 0
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, -0.00001, 0
            PI = 4 * Atn(1)
            
            Set objReducer = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)
                                    
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set geomFactory = Nothing
            
            oPlacePoint1.Set -0.9 * parLayingLength - CptOffset + depth, 0, 0
            oDir1.Set -1, 0, 0
            Set objNozzle1 = CreateNozzleJustaCircle(1, oPartFclt, m_OutputColl, oDir1, oPlacePoint1)
            
            oPlacePoint2.Set parEndLength + 0.1 * parLayingLength + CptOffset2 - depth2, 0, 0
            oDir2.Set 1, 0, 0
            Set objNozzle2 = CreateNozzleJustaCircle(2, oPartFclt, m_OutputColl, oDir2, oPlacePoint2)
            
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
            LineStrPoints(4) = pipeDiam / 2
            LineStrPoints(5) = LineStrPoints(2)
            
            'Point 3
            LineStrPoints(6) = -parFacetoFace / 2 + ParLeftFlatEndLength
            LineStrPoints(7) = pipeDiam / 2
            LineStrPoints(8) = LineStrPoints(2)
            
            'Point 4
            LineStrPoints(9) = parFacetoFace / 2 - ParRightFlatEndLength
            LineStrPoints(10) = pipeDiam2 / 2
            LineStrPoints(11) = LineStrPoints(2)
            
            'Point 5
            LineStrPoints(12) = parFacetoFace / 2
            LineStrPoints(13) = pipeDiam2 / 2
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
            
            Set objReducer = PlaceRevolution(m_OutputColl, oLineString, oAxisVect, _
                                                oCenPoint, 2 * PI, True)
                                    
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oLineString = Nothing
            Set geomFactory = Nothing
            
            oPlacePoint1.Set -parFacetoFace / 2 - CptOffset + depth, 0, 0
            oDir1.Set -1, 0, 0
            Set objNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDir1, oPlacePoint1)
            
            oPlacePoint2.Set parFacetoFace / 2 + CptOffset2 - depth2, 0, 0
            oDir2.Set 1, 0, 0
            Set objNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir2, oPlacePoint2)
            
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
            Dim oCone As Object
            Set oCone = PlaceCone(m_OutputColl, oStPoint, oEndPoint, pipeDiam / 2, pipeDiam2 / 2, True)
            
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCone
            Set oCone = Nothing
            Set oStPoint = Nothing
            Set oEndPoint = Nothing
            
            'Place Nozzle 1
            If CmpDblLessThan(dNozzleLength1, flangeThick) Then dNozzleLength1 = flangeThick
            oPlacePoint1.Set -0.5 * dFacetoFace - CptOffset + depth, 0, 0
            oDir1.Set -1, 0, 0
            Set objNozzle1 = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir1, _
                        oPlacePoint1, dNozzleLength1)
            
            'Place Nozzle 2
            If CmpDblLessThan(dNozzleLength2, flangeThick2) Then dNozzleLength2 = flangeThick2
            oPlacePoint2.Set 0.5 * dFacetoFace + CptOffset2 - depth2, 0, 0
            oDir2.Set 1, 0, 0
            Set objNozzle2 = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir2, _
                        oPlacePoint2, dNozzleLength2)
        Case Else:
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objReducer
            Set objReducer = Nothing
    End Select
    
    ' Place Nozzle 1
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle1
    Set objNozzle1 = Nothing
    Set oDir1 = Nothing
    Set oPlacePoint1 = Nothing
        
    ' Place Nozzle 2
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle2
    Set objNozzle2 = Nothing
    Set oDir2 = Nothing
    Set oPlacePoint2 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
