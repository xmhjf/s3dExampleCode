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
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         MS
'   Creation Date:  Tuesday, Dec07 2004
'   Description:
'   The following Part data basis cases are addressed in this symbol
'   Default
'   Union tail, Type 2 (1036) (Source: Design document - M15 version, section E-12)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   08.SEP.2006     KKC     DI-95670  Replace names with initials in all revision history sheets and symbols
'   23.July.2008    RUK     CR-CP-145836  Enhance Union symbol to be specified by length 1, length 2 and nut thickness
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:"    'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam1 As Double
    Dim pipeDiam2 As Double
    Dim flangeThick As Double
    Dim cptOffset1 As Double
    Dim flangeDiam As Double
    Dim depth1 As Double
    Dim cptOffset2 As Double
    Dim depth2 As Double
    
    Dim iOutput As Long

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

    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oPlacePoint As AutoMath.DPosition
    
    Set oDir = New AutoMath.DVector
    Set oPlacePoint = New DPosition
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, cptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, cptOffset2, depth2
    
    Select Case lPartDataBasis
    Case Is <= 1
        parFacetoFace = arrayOfInputs(2)
        
        'Place Nozzle 1
        oDir.Set -1, 0, 0
        oPlacePoint.Set -cptOffset1 + depth1, 0, 0
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFacetoFace)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
        'Place Nozzle 2
        oPlacePoint.Set parFacetoFace + cptOffset2 - depth2, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
    
    Case 1036
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
        Dim oEnpoint As AutoMath.DPosition
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
                                                
        'Create the Body
        Dim objTail As Object
        oStPoint.Set 0, 0, 0
        oEnpoint.Set parLength1, 0, 0
        Set objTail = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, 1.2 * pipeDiam1, True)
                
        'Set the Output
        m_OutputColl.AddOutput "Tail", objTail
        Set objTail = Nothing
        
        'Create the Nut
        Dim oGeomFact As IngrGeom3D.GeometryFactory
        Dim objNut As Object
        Dim oLineStr As LineString3d
        Dim dNutSideLen As Double
        Dim dPoints(0 To 20) As Double
        
        Set oGeomFact = New GeometryFactory
        
        If Not CmpDblEqual(parNutWidth, 0) Then
            dNutSideLen = 2 * (parNutWidth / 2) / Tan(PI / 3)
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
        
        oVector.Set 1, 0, 0
        Set oLineStr = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 7, dPoints)
        Set objNut = PlaceProjection(m_OutputColl, oLineStr, oVector, dHalfNutThickness, True)
        
        'Set the Output
        m_OutputColl.AddOutput "Nut", objNut
        Set objNut = Nothing
        Set oLineStr = Nothing
        Set oGeomFact = Nothing
                
        'Create Port 1
        oDir.Set -1, 0, 0
        oPlacePoint.Set -(cptOffset1 - depth1), 0, 0
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Create Port 2
        oDir.Set 1, 0, 0
        
        Dim oPipePort As IJCatalogPipePort
        Dim oNozzColl As IJDCollection
        Set oNozzColl = oPartFclt.GetNozzles()
        Set oPipePort = oNozzColl.Item(2)
        
        If oPipePort.TerminationClass = 10 Then '"Male" Then
            oPlacePoint.Set parLength1 + parLength2 + cptOffset2 - depth2, 0, 0
            Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                                                            oPlacePoint, parLength2)
        Else
            oPlacePoint.Set parLength1 + cptOffset2 - depth2, 0, 0
            Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        End If
                
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        Set oPipePort = Nothing
        Set oNozzColl = Nothing
                                
        'Remove the references
        Set oStPoint = Nothing
        Set oEnpoint = Nothing
        Set oVector = Nothing
    End Select

    Set oDir = Nothing
    Set oPlacePoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
