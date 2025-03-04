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
Private Const MODULE = "Insulation:" 'Used for error messages
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

    CenterPos.Set 0, 0, 0

    Dim iOutput As Double
    Dim ObjBody As Object
    Dim ObjInsulatedBody As Object

    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double

    Dim parAngle As Double  'TR-142779
    If UBound(arrayOfInputs) >= 7 Then
        parAngle = arrayOfInputs(7)
    Else
        parAngle = PI / 4
    End If
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parInsulationThickness = arrayOfInputs(3)

    iOutput = 0
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    Dim pipeDiam2 As Double
    Dim flangeDiam2 As Double
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam2, cptOffset, depth
    
    Select Case lPartDataBasis
    Case Is <= 1    'Default
        parFacetoFace = arrayOfInputs(2)
        parInsulationThickness = arrayOfInputs(3)
        'Insert your code for output 2(Insulated Body)
        Dim parLatroletRad As Double
        parLatroletRad = BodyOD(pipeDiam2, flangeDiam) / 2
        Dim parLatroletBaseRad As Double
        'parLatroletBaseRad = BodyOD(pipeDiam * 0.625, parLatroletRad)
        parLatroletBaseRad = BodyOD(pipeDiam * 0.5, parLatroletRad)
        Dim Face2Pos As New AutoMath.DPosition
        Face2Pos.Set (parFacetoFace + cptOffset - depth) * Cos(parAngle), _
                            (parFacetoFace + cptOffset - depth) * Sin(parAngle), 0
    
        Set ObjInsulatedBody = PlaceCone(m_OutputColl, CenterPos, Face2Pos, _
            parLatroletBaseRad + parInsulationThickness, parLatroletRad + parInsulationThickness)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
        Set ObjInsulatedBody = Nothing
    
    Case 762    'Latrolet, specified by latrolet length
        Dim parLatroletHeight As Double
        Dim parLatroletWidth As Double
        Dim parLatroletLength As Double
                
        'Inputs
        parInsulationThickness = arrayOfInputs(3)
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
        
        Dim dInsDiam2 As Double
        dInsDiam2 = BodyOD(pipeDiam2, flangeDiam2) + 2 * parInsulationThickness
        
        Dim oGeomFact As IngrGeom3D.GeometryFactory
        Dim oStPoint As DPosition
        Dim oEnPoint As DPosition
        
        Set oGeomFact = New GeometryFactory
        Set oStPoint = New DPosition
        Set oEnPoint = New DPosition
        
        Dim dBaseCenX As Double
        Dim dBaseCenY As Double
        
        dBaseCenX = ((dLatroletLength + (pipeDiam / 2) / Sin(parAngle)) - _
                        (dLatroletLength + ((pipeDiam2 / 2) / Tan(parAngle)) + _
                        parInsulationThickness)) * Cos(parAngle)
        dBaseCenY = ((dLatroletLength + (pipeDiam / 2) / Sin(parAngle)) - _
                        (dLatroletLength + ((pipeDiam2 / 2) / Tan(parAngle)) + _
                        parInsulationThickness)) * Sin(parAngle)
        
        oStPoint.Set dBaseCenX, dBaseCenY, 0
        oEnPoint.Set (pipeDiam / 2) / Tan(parAngle) + dLatroletLength * Cos(parAngle), _
                                    pipeDiam / 2 + dLatroletLength * Sin(parAngle), 0
        Set ObjBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                    dInsDiam2 + 6 * dWallThick, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
        Set ObjBody = Nothing
        
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oGeomFact = Nothing
    End Select
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
