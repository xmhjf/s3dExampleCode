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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Wednesday 30, Jul 2003
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The Details were taken from PDS Piping Component Data - Reference Guide at Page no D-55
'    SN=F63 and MC=E45LT. Symbol is created with 8 Outputs - five are Insulation aspect
'    outputs which are as follows:
'      ObjInsulatedBody using 'PlaceRevolution' and  other 4 outputs - 'PlaceCylinder' function.
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'  02.Sep-2003     SSP             CR-47620  Modified 45LongTangent Elbow to Remove Union part.
'  19.Aug.2005     svsmylav        TR-83739: Added check to ensure non-zero length for
'                                  cylinder to avoid PlaceCylinder function failure.
'  08.SEP.2006     KKC             DI-95670  Replace names with initials in all revision history sheets and symbols.
'  13.Feb.2007     dkl             CR-113390  Modified the description. Implemented part data basis as per the design document.
'                                  Used CmpDblXXX methods where appropriate.
'  14-May-2007     svsmylav        TR-119696: Made Case 180 as default instead of case 160 and updated default case like PDS.
'  2.JUL.2007      dkl             TR-123363: Recieved inputs from the 'arrayOfInputs' in the respective Part Data Basis Cases.
'                            Used 'LINEAR_TOLERANCE' for comparing with zero in with function 'CmpDblGreaterthan'.
'  10.Oct.2007     dkl             CR-126668: Implemented Part data basis case 181. Updated port insulation to selectively
'                                     insulate ports when flange diameter is non zero.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
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

    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim flangeDiam2      As Double

    Dim iOutput     As Integer
    Dim ObjInsulatedTangent1 As Object
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedTangent2 As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedPort2 As Object
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parTangentLength1 As Double
    Dim parTangentLength2 As Double
    Dim parFacetoCenter As Double
    Dim parTangentLength As Double
    Dim parBendRadius As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFace1toCenter = arrayOfInputs(2)
'    parTangentLength1 = arrayOfInputs(3)
'    parFace2toCenter = arrayOfInputs(4)
'    parTangentLength2 = arrayOfInputs(5)
'    parFacetoCenter = arrayOfInputs(6)
'    parTangentLength = arrayOfInputs(7)
'    parBendRadius = arrayOfInputs(8)
'    parInsulationThickness = arrayOfInputs(9)

    'For supporting V6.1 and V7 versions.
    If UBound(arrayOfInputs) >= 9 Then
        parInsulationThickness = arrayOfInputs(9)
    Else
        parInsulationThickness = arrayOfInputs(5)
    End If
    
    iOutput = 0
    
    Dim dTangentLength1 As Double
    Dim dTangentLength2 As Double
    Dim dElbowRadius As Double
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dAngle As Double
    dAngle = PI / 4
      
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
' This symbol is based on the following Part data basis values that govern its geometry,
' Symmetrical Elbow-
'   (Part data Basis value -160): TangentLength and BendRadius
'   (Part data Basis value -165): FacetoCenter and TangentLength
' Unsymmetrical Elbow-
'   (Part data Basis value -170): TangentLength1, TangentLength2 and BendRadius
'   (Part data Basis value -175): Face1toCenter, TangentLength1 and BendRadius
'   (Part data Basis value -180): Face1toCenter, TangentLength1, Face2toCenter,
'                                           TangentLength2 and BendRadius
'   (Part data Basis value -185): Face1toCenter, TangentLength1 and TangentLength2
    
    Select Case lPartDataBasis
    
        Case Is <= 1, 180
            'Face1toCenter, TangentLength1, Face2toCenter, TangentLength2 and BendRadius
                If UBound(arrayOfInputs) >= 9 Then
                    parFace1toCenter = arrayOfInputs(2)
                    parTangentLength1 = arrayOfInputs(3)
                    parFace2toCenter = arrayOfInputs(4)
                    parTangentLength2 = arrayOfInputs(5)
                    parTangentLength = arrayOfInputs(7) 'For supporting V7 version.
                    parBendRadius = arrayOfInputs(8)
                Else
                'for supporting V6.1
                    parFace1toCenter = arrayOfInputs(2)
                    parFace2toCenter = arrayOfInputs(3)
                    parTangentLength = arrayOfInputs(4)
                End If
            'Notes:-
            '   The default case should work for following:
            '   i. PDS EDEN uses four parameters Face1toCenter, TangentLength1, Face2toCenter and
            '      TangentLength2. It computes bend radius.
                    dTangentLength1 = parTangentLength1
                    dTangentLength2 = parTangentLength2
            '   ii. Prior to CR-113390(SP1) fix, the symbol had three inputs viz. Face1toCenter,
            '       Face2toCenter, TangentLength(for tangent length at port1).It computes bend radius and tangent length at port2.
                    If CmpDblGreaterthan(parTangentLength, LINEAR_TOLERANCE) Then dTangentLength1 = parTangentLength
            '   user provided values for Face1toCenter/Face2toCenter/BendRadius should not be overwritten.
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            
            'Check if user provided bend radius
            If Not CmpDblEqual(parBendRadius, 0) Then
                dElbowRadius = parBendRadius
                dTangentLength1 = parFace1toCenter - parBendRadius * Tan(dAngle / 2)
                dTangentLength2 = parFace2toCenter - parBendRadius * Tan(dAngle / 2)
            Else
                'PDS EDEN approach
                Dim dDelta1 As Double
                Dim dDelta2 As Double
                Dim dDelta As Double
                
                dDelta1 = (parFace1toCenter - dTangentLength1)
                dDelta2 = (parFace2toCenter - dTangentLength2)
                dDelta = dDelta1 - dDelta2
                dElbowRadius = dDelta1 / Tan(dAngle / 2)
                If CmpDblLessThan(dDelta, -0.0001) Then
                    dTangentLength2 = dTangentLength2 + Abs(dDelta)
                Else
                    dTangentLength1 = dTangentLength1 + Abs(dDelta)
                    dElbowRadius = dDelta2 / Tan(dAngle / 2)
                End If
            End If
        Case 160 'TangentLength and BendRadius
            parTangentLength = arrayOfInputs(7)
            parBendRadius = arrayOfInputs(8)
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dElbowRadius = parBendRadius
            dFace1toCenter = parBendRadius * Tan(dAngle / 2) + parTangentLength
            dFace2toCenter = dFace1toCenter
        Case 165    'FacetoCenter and TangentLength
            parFacetoCenter = arrayOfInputs(6)
            parTangentLength = arrayOfInputs(7)
            dTangentLength1 = parTangentLength
            dTangentLength2 = parTangentLength
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            ' ElbowRadius is computed based on Face to Center and Tangent Length.
            dElbowRadius = (parFacetoCenter - parTangentLength) / Tan(dAngle / 2)
        Case 170   'TangentLength1, TangentLength2 and BendRadius
            parTangentLength1 = arrayOfInputs(3)
            parTangentLength2 = arrayOfInputs(5)
            parBendRadius = arrayOfInputs(8)
            dElbowRadius = parBendRadius
            dTangentLength1 = parTangentLength1
            dTangentLength2 = parTangentLength2
            dFace1toCenter = parBendRadius * Tan(dAngle / 2) + parTangentLength1
            dFace2toCenter = parBendRadius * Tan(dAngle / 2) + parTangentLength2
        Case 175   'Face1toCenter, TangentLength1 and BendRadius
            parFace1toCenter = arrayOfInputs(2)
            parTangentLength1 = arrayOfInputs(3)
            parBendRadius = arrayOfInputs(8)
            dElbowRadius = parBendRadius
            dTangentLength1 = parTangentLength1
            dFace1toCenter = parTangentLength1 + parBendRadius * Tan(dAngle / 2)
            dTangentLength2 = 0
            dFace2toCenter = parBendRadius * Tan(dAngle / 2)
        Case 181    'Face1toCenter and Face2toCenter
            parFace1toCenter = arrayOfInputs(2)
            parFace2toCenter = arrayOfInputs(4)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dElbowRadius = parFace2toCenter / Tan(dAngle / 2)
            dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dAngle / 2)
            dTangentLength2 = 0
        Case 185   'Face1toCenter, TangentLength1 and TangentLength2
            parFace1toCenter = arrayOfInputs(2)
            parTangentLength1 = arrayOfInputs(3)
            parTangentLength2 = arrayOfInputs(5)
            dFace1toCenter = parFace1toCenter
            dTangentLength1 = parTangentLength1
            dTangentLength2 = parTangentLength2
            ' ElbowRadius is computed based on Face to Center and Tangent Length.
            dElbowRadius = (parFace1toCenter - parTangentLength1) / Tan(dAngle / 2)
            dFace2toCenter = dElbowRadius * Tan(dAngle / 2) + parTangentLength2
        
        Case Else
            GoTo ErrorLabel 'Invalid specification.
    
    End Select

' Insert your code for output 1(Insulated Tangent At Port1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                    sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                    sptOffset, depth
'   If tangent length is zero and flange thickness is zero then insulation for tangent will fail (Zero length
'   cylinder). We use negligible thickness to protect symbol.
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = 0.0001
    Dim ArcStPointToCenter  As Double
    Dim InsulationRadius As Double
    
    ArcStPointToCenter = dFace1toCenter - dTangentLength1
    InsulationRadius = pipeDiam / 2 + parInsulationThickness

'   Ensure that the hub radius is not greater than the distance from the center
'   of rotation to the elbow center line
    If CmpDblGreaterthan(InsulationRadius, dElbowRadius) Then
        InsulationRadius = dElbowRadius
    End If
    
    stPoint.Set -dFace1toCenter + flangeThick, 0, 0
    enPoint.Set -ArcStPointToCenter, 0, 0
    
    Set ObjInsulatedTangent1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                                2 * InsulationRadius, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedTangent1
    Set ObjInsulatedTangent1 = Nothing

' Insert your code for output 2(Insulated Body)

'   Construct a circle that will be used to create the revolution
    Dim objCircleI   As IngrGeom3D.Circle3d

    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Set objCircleI = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -ArcStPointToCenter, 0, 0, _
                        1, 0, 0, _
                      InsulationRadius)
                      
'   Revolve the circle about the Z axis passing by the origin
    Dim CP As New AutoMath.DPosition 'revolution center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    CP.Set -ArcStPointToCenter, dElbowRadius, 0
    CV.Set 0, 0, 1
    Set ObjInsulatedBody = PlaceRevolution(m_OutputColl, objCircleI, _
                    CV, CP, dAngle, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set geomFactory = Nothing
    Set CP = Nothing
    Set CV = Nothing
    Set objCircleI = Nothing
    Set geomFactory = Nothing

' Insert your code for output 3(Insulated Tangent At Port2)
    ArcStPointToCenter = dFace2toCenter - dTangentLength2
    
    stPoint.Set (dFace2toCenter - flangeThick2) * Cos(dAngle), _
                (dFace2toCenter - flangeThick2) * Sin(dAngle), 0
    enPoint.Set ArcStPointToCenter * Cos(dAngle), _
                 ArcStPointToCenter * Sin(dAngle), 0
    InsulationRadius = pipeDiam2 / 2 + parInsulationThickness
    
    Set ObjInsulatedTangent2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                                2 * InsulationRadius, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedTangent2
    Set ObjInsulatedTangent2 = Nothing

' Insert your code for output 4(Insulated Port1)
    If CmpDblGreaterthan(flangeDiam, LINEAR_TOLERANCE) Then
        Dim InsDiameterPort As Double
        stPoint.Set -dFace1toCenter, 0, 0
        enPoint.Set -dFace1toCenter + flangeThick + parInsulationThickness, 0, 0
        InsDiameterPort = flangeDiam + parInsulationThickness * 2
                
        Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                                                            InsDiameterPort, True)
    
        ' Set the output
        m_OutputColl.AddOutput "InsulatedPort1", ObjInsulatedPort1
        Set ObjInsulatedPort1 = Nothing
    End If

' Insert your code for output 5(Insulated Port2)
    If CmpDblGreaterthan(flangeDiam2, LINEAR_TOLERANCE) Then
        stPoint.Set dFace2toCenter * Cos(dAngle), dFace2toCenter * Sin(dAngle), 0
        enPoint.Set (dFace2toCenter - flangeThick2 - parInsulationThickness) * Cos(dAngle), _
                    (dFace2toCenter - flangeThick2 - parInsulationThickness) * Sin(dAngle), 0
        InsDiameterPort = flangeDiam2 + parInsulationThickness * 2
                
        Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                                                            InsDiameterPort, True)
    
        ' Set the output
        m_OutputColl.AddOutput "InsulatedPort2", ObjInsulatedPort2
        Set ObjInsulatedPort2 = Nothing
    End If
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
