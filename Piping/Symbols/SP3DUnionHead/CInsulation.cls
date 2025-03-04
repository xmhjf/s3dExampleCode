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
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         MS
'   Creation Date:  Tuesday, Dec07 2004
'   Description:
'   The following Part data basis cases are addressed in this symbol
'   Default
'   Union head, Type 2 (1034) (Source: Design document - M15 version, section E-12)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----        ------------------
'  08.SEP.2006     KKC      DI-95670  Replace names with initials in all revision history sheets and symbols
'  24.OCT.2007     KKC      TR-129956  User is not able to place delivered symbol for Union Tail
'   23.July.2008    RUK     CR-CP-145836  Enhance Union symbol to be specified by length 1, length 2 and nut thickness
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:"    'Used for error messages
Private PI As Double
Private Const NEGLIGIBLE_THICKNESS = 0.0001

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam1 As Double
    Dim flangeThick As Double
    Dim cptOffset As Double
    Dim flangeDiam1 As Double
    Dim depth As Double
    Dim pipeDiam2 As Double
    Dim flangeDiam2 As Double

    Dim iOutput As Double
    Dim ObjInsulatedBody As Object

    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)

    iOutput = 0

    'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam1, cptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, cptOffset, depth
    
    Select Case lPartDataBasis
    Case lPartDataBasis <= 1
        parFacetoFace = arrayOfInputs(2)
        
        Dim CenterPos As New AutoMath.DPosition
        CenterPos.Set 0, 0, 0

        ' Insert your code for output 1(Insulated Body)
        Dim parInsulationDiameter As Double
    
        parInsulationDiameter = ReturnMax4(pipeDiam1, pipeDiam2, flangeDiam1, flangeDiam2) + parInsulationThickness * 2
        Dim parInsulationLength As Double
        'parInsulationLength = InsulationLength(parFacetoFace, flangeThick, parInsulationThickness)
        'Kept the insulation cylinder length to be the FacetoFace,
        'as for higher insulation thickness values the InsulationLength function
        'returns values greater than the FacetoFace, due to which the Insulation graphics become longer than the flange.
        Dim endofInsulation As New AutoMath.DPosition
        endofInsulation.Set parFacetoFace, 0, 0
    
        Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, CenterPos, endofInsulation, parInsulationDiameter, True)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
        Set ObjInsulatedBody = Nothing

    Case 1034
        Dim parLength1 As Double
        Dim parLength2 As Double
        Dim parHalfNutThickness As Double
        Dim parNutWidth As Double
        
        'Inputs
        parLength1 = arrayOfInputs(4)
        parLength2 = arrayOfInputs(5)
        parHalfNutThickness = arrayOfInputs(6)
        parNutWidth = arrayOfInputs(7)
    
        Dim oStPoint As AutoMath.DPosition
        Dim oEnpoint As New DPosition
        Dim oVector As AutoMath.DVector
        
        Set oStPoint = New DPosition
        Set oEnpoint = New DPosition
        Set oVector = New DVector
        
        Dim dHalfNutThickness As Double
        Dim dNozzLen As Double
        
        'Compute the value of half Nut thickness when it's value was not specified
        If Not CmpDblEqual(parHalfNutThickness, LINEAR_TOLERANCE) Then
            dHalfNutThickness = parHalfNutThickness
        Else
            dHalfNutThickness = parLength1 / 3
        End If
                                                
        'Create the Indulated Body
        Dim objInsHead As Object
        oStPoint.Set 0, 0, 0
        oEnpoint.Set -parLength1, 0, 0
        Set objInsHead = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, _
                                    1.2 * pipeDiam1 + 2 * parInsulationThickness, True)
                
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsHead
        Set objInsHead = Nothing
        
        'Create the Nut
        Dim oGeomFact As IngrGeom3D.GeometryFactory
        Dim objNut As Object
        Dim oLineStr As LineString3d
        Dim oScaleMat As AutoMath.DT4x4
        Dim dNutSideLen As Double
        Dim dPoints(0 To 20) As Double
        Dim dScaleFact As Double
        
        Set oGeomFact = New GeometryFactory
        Set oScaleMat = New DT4x4
        
        If Not CmpDblEqual(parNutWidth, 0) Then
            dNutSideLen = (parNutWidth / 2) / Tan(PI / 3)
        Else
            dNutSideLen = 1.2 * pipeDiam1
        End If
        
        'Point 1
        dPoints(0) = 0
        dPoints(1) = dNutSideLen / 2
        dPoints(2) = (dNutSideLen / 2) * Tan(PI / 3)
        
        'Point 2
        dPoints(3) = dPoints(0)
        dPoints(4) = dPoints(1) + dNutSideLen * Cos(PI / 3)
        dPoints(5) = 0
        
        'Point 3
        dPoints(6) = dPoints(0)
        dPoints(7) = dPoints(1)
        dPoints(8) = -dPoints(2)
        
        'Point 4
        dPoints(9) = dPoints(0)
        dPoints(10) = -dPoints(7)
        dPoints(11) = dPoints(8)
        
        'Point 5
        dPoints(12) = dPoints(0)
        dPoints(13) = -dPoints(4)
        dPoints(14) = dPoints(5)
        
        'Point 6
        dPoints(15) = dPoints(0)
        dPoints(16) = -dPoints(1)
        dPoints(17) = dPoints(2)
        
        'Point 7
        dPoints(18) = 0
        dPoints(19) = dNutSideLen / 2
        dPoints(20) = (dNutSideLen / 2) * Tan(PI / 3)
        
        oVector.Set -1, 0, 0
        Set oLineStr = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 7, dPoints)
        
        'Scale the Line string
        dScaleFact = (dNutSideLen / 2 + dNutSideLen * Cos(PI / 3) + parInsulationThickness) / _
                        (dNutSideLen / 2 + dNutSideLen * Cos(PI / 3))
        oScaleMat.LoadIdentity
        oScaleMat.[Scale] dScaleFact
        oLineStr.Transform oScaleMat
        
        Set objNut = PlaceProjection(m_OutputColl, oLineStr, oVector, dHalfNutThickness, True)
        
        'Set the Output
        m_OutputColl.AddOutput "InsulatedNut", objNut
        Set objNut = Nothing
        Set oLineStr = Nothing
        Set oGeomFact = Nothing
        
        'Create the Insulation for Port 2
        Dim objInsPort2 As Object
        Dim dPort2Len As Double
        
        Dim oPipePort As IJCatalogPipePort
        Dim oNozzColl As IJDCollection
        Set oNozzColl = oPartFclt.GetNozzles()
        Set oPipePort = oNozzColl.Item(2)
        
        If oPipePort.TerminationClass = 10 Then '"Male" Then
            dPort2Len = parLength2
        End If
        oStPoint.Set -(parLength1 + dPort2Len), 0, 0
        oEnpoint.Set -(parLength1 + dPort2Len - _
                            BodyOD(dPort2Len, NEGLIGIBLE_THICKNESS)), 0, 0
        Set objInsPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, _
                         BodyOD(pipeDiam2, flangeDiam2) + 2 * parInsulationThickness, True)
        
        'Set the Output
        m_OutputColl.AddOutput "InuslationPort2", objInsPort2
        Set objInsPort2 = Nothing
        Set oPipePort = Nothing
        Set oNozzColl = Nothing
                                
        'Remove the references
        Set oStPoint = Nothing
        Set oEnpoint = Nothing
        Set oVector = Nothing
    End Select
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub

Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If CmpDblLessThan(MaxValue, B) Then MaxValue = B
    If CmpDblLessThan(MaxValue, C) Then MaxValue = C
    If CmpDblLessThan(MaxValue, D) Then MaxValue = D
    ReturnMax4 = MaxValue
End Function

