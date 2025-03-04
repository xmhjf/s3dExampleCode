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
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         ACM
'   Creation Date:  Monday, July 12, 2004
'   Description:
'       This is a Closure Palte symbol. Sketch is taken from Design document.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   08.SEP.2006      KKC    DI-95670  Replace names with initials in all revision history sheets and symbols
'   15.Feb.2008     RUK         CR-33401  Create symbols, as required, for jacketed piping specs
'                               Sketch is taken from design document appendix E 105
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private PI As Double
Private Const MODULE = "SimplePhysical:" 'Used for error messages

Private Sub Class_Terminate()
    '''
End Sub

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

    Dim iOutput As Integer

'    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
'    Dim parPlateInsideDia As Double
'    Dim parPlateThick As Double
'    Dim parJackPipeOffset As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parInsulationThickness = arrayOfInputs(3)
'    parPlateOutsideDia = arrayOfInputs(3)
'    parPlateInsideDia = arrayOfInputs(4)
'    parPlateThick = arrayOfInputs(5)
'    parJackPipeOffset = arrayOfInputs(6)
    iOutput = 0
    
     'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim dPlateThick As Double
    Dim dPlateOutSideDia As Double
    Dim dPlateInsideDia As Double
    Dim dJackPipeOffset As Double
    
    Set oPlacePoint = New AutoMath.DPosition
    'Retrieve the nozzle parameters
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                        cptOffset, depth
    Select Case lPartDataBasis
    Case Is <= 1
        Dim parFacetoFace As Double
        
        'Inputs
        parFacetoFace = arrayOfInputs(2)
        parInsulationThickness = arrayOfInputs(3)
        
        dPlateOutSideDia = pipeDiam
        dPlateInsideDia = 0.5 * dPlateOutSideDia
        dPlateThick = parFacetoFace
        dJackPipeOffset = parFacetoFace / 2
        oPlacePoint.Set parFacetoFace / 2 + cptOffset - depth, 0, 0
    
    Case 817, 818   'Jacket closure plate, outside diameter fits flush with inside diameter of jacket pipe (817).
                    'Jacket closure plate, outside diameter larger than outside diameter of jacket pipe (818).
        Dim parPlateOutsideDia As Double
        Dim parPlateInsideDia As Double
        Dim parPlateThick As Double
        Dim parJackPipeOffset As Double
        
        'Inputs
        parInsulationThickness = arrayOfInputs(3)
        parPlateOutsideDia = arrayOfInputs(4)
        parPlateInsideDia = arrayOfInputs(5)
        parPlateThick = arrayOfInputs(6)
        parJackPipeOffset = arrayOfInputs(7)
        
        If CmpDblEqual(parPlateOutsideDia, 0) Then
            Dim dPipeOD As Double
            Dim dPipeWallThick As Double
            RetrievePipeOD 1, oPartFclt, m_OutputColl, dPipeOD
            RetrievePipeWallThick 1, oPartFclt, dPipeWallThick
            dPlateOutSideDia = dPipeOD - 2 * dPipeWallThick
        Else
            dPlateOutSideDia = parPlateOutsideDia
        End If
        If CmpDblEqual(parPlateInsideDia, 0) Then
            dPlateInsideDia = 0.5 * parPlateOutsideDia
        Else
            dPlateInsideDia = parPlateInsideDia
        End If
        dPlateThick = parPlateThick
        If lPartDataBasis = 817 Then
            dJackPipeOffset = parJackPipeOffset
        ElseIf lPartDataBasis = 818 Then
            dJackPipeOffset = 0
        End If
        oPlacePoint.Set 0, 0, 0
    Case Else
        GoTo ErrorLabel
    End Select
    
    'Create the closure plate Body
    Dim oGeomFact As IngrGeom3D.GeometryFactory
    Dim oCenter As AutoMath.DPosition
    Dim oRevAxis As AutoMath.DVector
    Dim oLineStr As IngrGeom3D.LineString3d
    Dim ObjClosurePlateBody As Object
    Dim dPoints(0 To 14) As Double

    Set oGeomFact = New GeometryFactory
    Set oCenter = New DPosition
    Set oRevAxis = New DVector
    Set oLineStr = New LineString3d
    
    dPoints(0) = dJackPipeOffset
    dPoints(1) = dPlateInsideDia / 2
    dPoints(2) = 0
    
    dPoints(3) = dPoints(0)
    dPoints(4) = dPlateOutSideDia / 2
    dPoints(5) = dPoints(2)
    
    dPoints(6) = dPoints(0) - dPlateThick
    dPoints(7) = dPoints(4)
    dPoints(8) = dPoints(2)
    
    dPoints(9) = dPoints(6)
    dPoints(10) = dPoints(1)
    dPoints(11) = dPoints(2)
    
    dPoints(12) = dPoints(0)
    dPoints(13) = dPoints(1)
    dPoints(14) = dPoints(2)
    
    Set oLineStr = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
    oCenter.Set 0, 0, 0
    oRevAxis.Set 1, 0, 0
    Set ObjClosurePlateBody = PlaceRevolution(m_OutputColl, oLineStr, oRevAxis, _
                                                oCenter, 2 * PI, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjClosurePlateBody
    Set ObjClosurePlateBody = Nothing
    Set oLineStr = Nothing
    Set oRevAxis = Nothing
    Set oCenter = Nothing
    Set oGeomFact = Nothing
    
    ' Place Nozzle
    Dim oDir As AutoMath.DVector
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
        
    Set oDir = New AutoMath.DVector
        
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
        
    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext
End Sub
