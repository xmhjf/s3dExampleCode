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
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         SSP
'   Creation Date:  Wednesday, Aug 27 2003
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The Details are taken from design document, Appendix E-26.
'   Change History:
'   dd.mmm.yyyy    who       change description
'   -----------    -----     ------------------
'  08.SEP.2006     KKC       DI-95670  Replace names with initials in all revision history sheets and symbols.
'  08.Nov.2006     dkl       TR-109032 Introduced CmpDblXXX functions where required.
'  13.Feb.2007     dkl       CR-113390  Modified the description. Implemented part data basis as per the design document.
'  11-May-2007     svsmylav  TR-119696: Made Case 180 as default instead of case 160 and updated default case like PDS.
'  2.JUL.2007      dkl       TR-123363: Recieved inputs from the 'arrayOfInputs' in the respective Part Data Basis Cases.
'                            Used 'LINEAR_TOLERANCE' for comparing with zero in with function 'CmpDblGreaterthan'.
'  10.Oct.2007     dkl       CR-126668: Implemented Part data basis case 181.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "SimplePhysical:" 'Used for error messages
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

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parTangentLength1 As Double
    Dim parTangentLength2 As Double
    Dim parFacetoCenter As Double
    Dim parTangentLength As Double
    Dim parBendRadius As Double
    Dim parInsulationThickness As Double
    
    Dim axisVect    As New AutoMath.DVector
    Dim ArcStPointToCenter  As Double
    Dim iCount As Integer
    
    Dim iOutput     As Integer

    Dim ObjElbow As Object
    
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

   
    iOutput = 0
    
    Dim dTangentLength1 As Double
    Dim dTangentLength2 As Double
    Dim dElbowRadius As Double
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
    Dim dAngle As Double
    dAngle = PI / 2
      
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
            'Face1toCenter, TangentLength1, Face2toCenter, TangentLength2 and BendRadius.
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
                dTangentLength1 = parFace1toCenter - parBendRadius
                dTangentLength2 = parFace2toCenter - parBendRadius
            Else
                'PDS EDEN approach
                Dim dR1 As Double
                Dim dR2 As Double
                Dim dDelta As Double
                
                dR1 = parFace1toCenter - dTangentLength1
                dR2 = parFace2toCenter - dTangentLength2
                dDelta = dR1 - dR2
                dElbowRadius = dR1
                If CmpDblLessThan(dDelta, -0.0001) Then
                    dTangentLength2 = dTangentLength2 + Abs(dDelta)
                Else
                    dTangentLength1 = dTangentLength1 + Abs(dDelta)
                    dElbowRadius = dR2
                End If
            End If
        Case 160   'TangentLength and BendRadius
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
            dTangentLength1 = dFace1toCenter - dElbowRadius
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
    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory

' Insert your code for output 6(Elbow)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    ArcStPointToCenter = dFace1toCenter - dTangentLength1
'   Construct a circle that will be used to create the revolution
    Dim objCircle   As IngrGeom3D.Circle3d
    Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         -ArcStPointToCenter, 0, 0, _
                        1, 0, 0, _
                        pipeDiam / 2)
'   Revolve the circle about the Z axis
    Dim CP As New AutoMath.DPosition 'arc center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    CP.Set -ArcStPointToCenter, dElbowRadius, 0
    CV.Set 0, 0, 1
    Set ObjElbow = PlaceRevolution(m_OutputColl, objCircle, _
                    CV, CP, dAngle, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjElbow
    Set ObjElbow = Nothing
    Set CP = Nothing
    Set CV = Nothing
    Set objCircle = Nothing
    Set geomFactory = Nothing
    
' Place Nozzle 1
    Dim dNozzleLength As Double
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    Dim PorttoOrigin As Double
    PorttoOrigin = -dFace1toCenter - sptOffset + depth
    oPlacePoint.Set PorttoOrigin, 0, 0
    oDir.Set -1, 0, 0

    dNozzleLength = dTangentLength1
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dNozzleLength)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    PorttoOrigin = dFace2toCenter + sptOffset - depth
    oPlacePoint.Set 0, PorttoOrigin, 0
    oDir.Set 0, 1, 0

    dNozzleLength = dTangentLength2
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dNozzleLength)
                        
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
