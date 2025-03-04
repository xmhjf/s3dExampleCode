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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  Sunday, Feb 4 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   09.Jul.2003     SymbolTeam(India)   Copyright Information, Header  is added.
'   08.SEP.2006     KKC  DI-95670       Replace names with initials in all revision history sheets and symbols
'   12.Feb.2008     RUK                 CR-136268  Enhance the latrolet symbol to be more realistic per Bonney Forge catalog
'   03.Jul.2008     RUK                 TR-142779  Latrolet symbol does not support trimming for insertion on sloped line
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim cptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double
    Dim CenterPos As New AutoMath.DPosition
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    CenterPos.Set 0, 0, 0

    Dim iOutput As Double
    Dim ObjBody As Object

'    Dim parFacetoFace As Double
'    Dim parInsulationThickness As Double
'    Dim parLatroletHeight As Double
'    Dim parLatroletWidth As Double
'    Dim parLatroletLength As Double
                
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parInsulationThickness = arrayOfInputs(3)
'    parLatroletHeight = arrayOfInputs(4)
'    parLatroletWidth = arrayOfInputs(5)
'    parLatroletLength = arrayOfInputs(6)
    iOutput = 0
    
    Dim parAngle As Double  'TR-142779
    If UBound(arrayOfInputs) >= 7 Then
        parAngle = arrayOfInputs(7)
    Else
        parAngle = PI / 4
    End If
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

    ' Insert your code for output 1(Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim pipeDiam2 As Double
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, cptOffset, depth
    
    Select Case lPartDataBasis
    Case Is <= 1    'Default
        Dim parFacetoFace As Double
        Dim parInsulationThickness As Double
        
        'Inputs
        parFacetoFace = arrayOfInputs(2)
        parInsulationThickness = arrayOfInputs(3)

        Dim parLatroletRad As Double
        parLatroletRad = BodyOD(pipeDiam2, flangeDiam) / 2
        Dim parLatroletBaseRad As Double
        '    parLatroletBaseRad = BodyOD(pipeDiam * 0.625, parLatroletRad)
        parLatroletBaseRad = BodyOD(pipeDiam * 0.5, parLatroletRad)
        Dim Face2Pos As New AutoMath.DPosition
        Face2Pos.Set (parFacetoFace + cptOffset - depth) * Cos(parAngle), (parFacetoFace + cptOffset - depth) * Sin(parAngle), 0
        Set ObjBody = PlaceCone(m_OutputColl, CenterPos, Face2Pos, parLatroletBaseRad, parLatroletRad)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
        Set ObjBody = Nothing
        
    Case 762    'Latrolet, specified by latrolet length
        Dim parLatroletHeight As Double
        Dim parLatroletWidth As Double
        Dim parLatroletLength As Double
                
        'Inputs
        parLatroletHeight = arrayOfInputs(4)
        parLatroletWidth = arrayOfInputs(5)
        parLatroletLength = arrayOfInputs(6)
        
        Dim dLatroletLength As Double 'Length from pipe durface to port 2 along the latrolet axis
        dLatroletLength = parLatroletLength
        
        'Calculate Latrolet length based on Latrolet Width or Latrolet Height
        If CmpDblEqual(parLatroletLength, LINEAR_TOLERANCE) Then
            If Not CmpDblEqual(parLatroletWidth, LINEAR_TOLERANCE) Then
                dLatroletLength = (parLatroletWidth - _
                            (pipeDiam2 * (Sin(parAngle) + 1 / Sin(parAngle)) / 2)) / Cos(parAngle)
            ElseIf Not CmpDblEqual(parLatroletHeight, LINEAR_TOLERANCE) Then
                dLatroletLength = (parLatroletHeight - pipeDiam2 * Cos(parAngle) / 2) / Sin(parAngle)
            Else
                GoTo ErrorLabel
            End If
        End If
        
        Dim dWallThick As Double
        RetrievePipeWallThick 2, oPartFclt, dWallThick
        
        Dim oGeomFact As IngrGeom3D.GeometryFactory
        Dim oLineStr As IngrGeom3D.LineString3d
        Dim dPoints(0 To 11) As Double
        
        Set oGeomFact = New GeometryFactory
        Set oLineStr = New LineString3d
        
        Dim dLatUpSurLen As Double  'Length of the Latrolet Upper surface upto Pipe Surface
        dLatUpSurLen = dLatroletLength + pipeDiam2 / 2
        
        'Divide the Latrolet upper surface into three parts having lengths 20%, 30%, 40% of
        'Latrolet Upper surface Length respectively.
        'Assume that projection at the surface is 3 times of Wall thickness of the pipe.
        Dim dTempAng As Double 'Slope for second and third lines w.r.t latrolet axis
        Dim dTempAng2 As Double 'parAngle + dTempAng 'Angle made by second line with pipe axis
        Dim dTempAng3 As Double '(parAngle - dTempAng)
        Dim dTempLength1 As Double  'Length of the line which has 30% contribution in Latrolet Upper surface Length
        
        dTempAng = Atn((3 * dWallThick) / 0.4 * dLatUpSurLen)
        dTempAng2 = parAngle + dTempAng
        dTempAng3 = parAngle - dTempAng
        dTempLength1 = 0.3 * dLatUpSurLen / Cos(dTempAng)
        
        'Create the Line string along the upper surface of the latrolet and revolve it.
        'Point 1
        dPoints(0) = (pipeDiam / 2) / Tan(parAngle) - (pipeDiam2 / 2) / Sin(parAngle)
        dPoints(1) = pipeDiam / 2
        dPoints(2) = 0
        
        'Point 2
        dPoints(3) = dPoints(0)
        dPoints(4) = pipeDiam / 2 + 0.2 * dLatUpSurLen * Sin(parAngle)
        dPoints(5) = dPoints(2)
        
        'Point 3
        dPoints(6) = dPoints(3) + dTempLength1 * Cos(dTempAng2)
        dPoints(7) = dPoints(4) + dTempLength1 * Sin(dTempAng2)
        dPoints(8) = dPoints(2)
        
        'Point 4
        dPoints(9) = (pipeDiam / 2) / Tan(parAngle) + _
                    dLatroletLength * Cos(parAngle) - (pipeDiam2 / 2) * Sin(parAngle)
        dPoints(10) = pipeDiam / 2 + dLatroletLength * Sin(parAngle) + _
                        (pipeDiam2 / 2) * Cos(parAngle)
        dPoints(11) = dPoints(2)
        Set oLineStr = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 4, dPoints)
        
        Dim oVector As DVector
        Dim oCenter As DPosition
        
        Set oVector = New DVector
        Set oCenter = New DPosition
        
        oCenter.Set (pipeDiam / 2) / Tan(parAngle) + dLatroletLength * Cos(parAngle), _
                                (pipeDiam / 2 + dLatroletLength * Sin(parAngle)), 0
        oVector.Set Cos(parAngle), Sin(parAngle), 0
        
        Set ObjBody = PlaceRevolution(m_OutputColl, oLineStr, oVector, oCenter, 2 * PI, True)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
        Set ObjBody = Nothing
        
        Set oVector = Nothing
        Set oCenter = Nothing
        Set oLineStr = Nothing
        Set oGeomFact = Nothing
    End Select

    'Place Nozzle 1
    oPlacePoint.Set 0, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    'Place Nozzle 2
    If lPartDataBasis <= 1 Then 'Default
        oPlacePoint.Set (parFacetoFace + cptOffset - depth) * Cos(parAngle), _
                            (parFacetoFace + cptOffset - depth) * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0
    ElseIf lPartDataBasis = 762 Then    ''Latrolet, specified by latrolet length
        oPlacePoint.Set ((pipeDiam / 2) / Tan(parAngle) + (dLatroletLength + cptOffset - depth) * Cos(parAngle)), _
                        (pipeDiam / 2 + (dLatroletLength + cptOffset - depth) * Sin(parAngle)), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0
    End If
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
        
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
