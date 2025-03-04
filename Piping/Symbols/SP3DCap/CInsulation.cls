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
'   Creation Date:  Wednesday, Nov 22 2000
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
        '   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'  09.Jul.2003      SymbolTeam(India)   Copyright Information, Header  is added.
'  08.SEP.2006      KKC                 DI-95670  Replace names with initials in all revision history sheets and symbols.
'  30.Jan.2007      dkl                 CR 112819 Implemented Part data basis to modify cap symbol for HDPE spec.
'  19.Feb.2008      RUK                 CR-33401  Create symbols, as required, for jacketed piping specs
'                                       Implemented the PDB:"Cap, jacket, specified by cap length, tangent length, and inside diameter (819)"
'   25.Sep.2008     MP     CR-150009    Implemented part data basis values 1087, 1088 and 1089.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double
Private Const NEGLIGIBLE_THICKNESS = 0.0001

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
    Dim objInsulatedBody As Object
    Dim parFacetoEnd As Double
    Dim parInsulationThickness As Double
    Dim parFacetoFace As Double
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    '   parFacetoEnd = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    '   parFace1toCenter = arrayOfInputs(4)
    '    parCapInsideDia = arrayOfInputs(5)
    '    parCapLength = arrayOfInputs(6)
    '    parTangentLength = arrayOfInputs(7)


    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis

    If lPartDataBasis = 1087 Or lPartDataBasis = 1088 Or lPartDataBasis = 1089 Then
        parFacetoFace = arrayOfInputs(2)
    Else
        parFacetoEnd = arrayOfInputs(2)
    End If

    Dim parInsulationDiameter As Double
    Dim oEndOfInsulation As AutoMath.DPosition
    Set oEndOfInsulation = New AutoMath.DPosition
    Dim oStOfInsulation As AutoMath.DPosition
    Set oStOfInsulation = New AutoMath.DPosition
    Dim oCenterPos As AutoMath.DPosition
    Set oCenterPos = New AutoMath.DPosition
    oCenterPos.Set 0, 0, 0

    Dim dFaceToEndIns As Double

    iOutput = 0


    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                       cptOffset, depth

    ' The Cap symbol is based on the following three Part data basis values that govern its geometry,
    ' 'End Cap, Welded' - 59, 'End Cap, machined' - 60, and 'End Cap, molded' - 65 .

    Select Case lPartDataBasis
    Case Is <= 1, 57, 60, 65
        'The value of Port Depth defines if the Cap is Male:BE or Female: SW or FT
        If CmpDblEqual(depth, 0) Then
            oStOfInsulation.Set 0, 0, 0
            dFaceToEndIns = parFacetoEnd + parInsulationThickness
            oEndOfInsulation.Set dFaceToEndIns, 0, 0
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        Else
            oStOfInsulation.Set -(depth - cptOffset), oCenterPos.y, oCenterPos.z
            oEndOfInsulation.Set parFacetoEnd - (depth - cptOffset) + parInsulationThickness, _
                                 oCenterPos.y, oCenterPos.z
            parInsulationDiameter = flangeDiam + parInsulationThickness * 2
        End If
        ' Insert your code for output 2(Insulated Body)
        Set objInsulatedBody = PlaceCylinder(m_OutputColl, oStOfInsulation, oEndOfInsulation, _
                                             parInsulationDiameter, True)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody
        Set objInsulatedBody = Nothing
    
    Case 819    ''Cap, jacket, specified by cap length, tangent length, and inside diameter
        'CR-33401  Create symbols, as required, for jacketed piping specs
        Dim parCapInsideDia As Double
        Dim parCapLength As Double
        Dim parTangentLength As Double

        'Inputs
        parCapInsideDia = arrayOfInputs(5)
        parCapLength = arrayOfInputs(6)
        parTangentLength = arrayOfInputs(7)

        Dim oGeomFact As IngrGeom3D.GeometryFactory
        Dim oCompStr As IngrGeom3D.ComplexString3d
        Dim oCollection As Collection
        Dim oLine As IngrGeom3D.Line3d
        Dim oEllArc As IngrGeom3D.EllipticalArc3d
        Dim oStPoint As AutoMath.DPosition
        Dim oEndPoint As AutoMath.DPosition
        Dim oCenPoint As AutoMath.DPosition
        Dim oMajor As AutoMath.DPosition
        Dim oNormal As AutoMath.DVector
        Dim dMMRatio As Double
        Dim iCount As Integer

        Set oGeomFact = New GeometryFactory
        Set oCompStr = New ComplexString3d
        Set oCollection = New Collection
        Set oLine = New Line3d
        Set oEllArc = New EllipticalArc3d
        Set oStPoint = New DPosition
        Set oEndPoint = New DPosition
        Set oCenPoint = New DPosition
        Set oMajor = New DPosition
        Set oNormal = New DVector

        PI = 4 * Atn(1)

        'Assume Tangent Length = 3/4 of Cap Length, if its value = 0
        Dim dTangentLength As Double
        dTangentLength = parTangentLength
        If CmpDblEqual(parTangentLength, LINEAR_TOLERANCE) Then
            dTangentLength = (3 / 4) * parCapLength
        End If

        'Line 1
        oStPoint.Set -(depth - cptOffset), pipeDiam / 2 + parInsulationThickness, 0
        oEndPoint.Set dTangentLength, oStPoint.y, oStPoint.z
        Set oLine = PlaceTrLine(oStPoint, oEndPoint)
        oCollection.Add oLine
        Set oLine = Nothing

        'Elliptical Arc
        oCenPoint.Set -(depth - cptOffset) + dTangentLength, parCapInsideDia / 2, 0
        oNormal.Set 0, 0, -1
        oMajor.Set 0, (pipeDiam - parCapInsideDia) / 2 + parInsulationThickness, 0
        dMMRatio = (parCapLength - dTangentLength + parInsulationThickness) / _
                            ((pipeDiam - parCapInsideDia) / 2 + parInsulationThickness)
        Set oEllArc = oGeomFact.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                    oCenPoint.x, oCenPoint.y, oCenPoint.z, _
                                    oNormal.x, oNormal.y, oNormal.z, _
                                    oMajor.x, oMajor.y, oMajor.z, dMMRatio, 0, PI / 2)
        oCollection.Add oEllArc
        Set oEllArc = Nothing
        Set oMajor = Nothing

        'Create the Complex String
        oStPoint.Set -(depth - cptOffset), pipeDiam / 2 + parInsulationThickness, 0
        Set oCompStr = PlaceTrCString(oStPoint, oCollection)

        'Revolve the complex string about the center
        oNormal.Set -1, 0, 0
        oCenPoint.Set 0, 0, 0
        Set objInsulatedBody = PlaceRevolution(m_OutputColl, oCompStr, oNormal, oCenPoint, 2 * PI, True)
        
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody
        Set objInsulatedBody = Nothing
        
        'Create the insulation for Port
        Dim objInsPort As Object
        Dim dInsDia As Double
        Dim dInsThick As Double
        
        dInsDia = BodyOD(pipeDiam, flangeDiam) + 2 * parInsulationThickness
        If CmpDblEqual(flangeThick, 0) Then
            dInsThick = NEGLIGIBLE_THICKNESS
        Else
            dInsThick = flangeThick + parInsulationThickness
        End If
        oStPoint.Set -(depth - cptOffset), 0, 0
        oEndPoint.Set dInsThick, 0, 0
        Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, dInsDia, True)
        
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "InsPort", objInsPort
        Set objInsPort = Nothing
        
        'Remove the References
        Set oGeomFact = Nothing
        Set oCompStr = Nothing
        For iCount = 1 To oCollection.Count
            oCollection.Remove 1
        Next iCount
        Set oCollection = Nothing
        Set oNormal = Nothing
        Set oStPoint = Nothing
        Set oEndPoint = Nothing
        Set oCenPoint = Nothing
    Case 1087, 1088, 1089
         oStOfInsulation.Set -parInsulationThickness, 0, 0
         oEndOfInsulation.Set parFacetoFace + parInsulationThickness, 0, 0
         parInsulationDiameter = flangeDiam + parInsulationThickness * 2
         
         Set objInsulatedBody = PlaceCylinder(m_OutputColl, oStOfInsulation, oEndOfInsulation, parInsulationDiameter, True)
            
         ' Set the output
         iOutput = iOutput + 1
         m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody
         Set objInsulatedBody = Nothing
        
         Set oStOfInsulation = Nothing
         Set oEndOfInsulation = Nothing
    
    Case Else
        GoTo ErrorLabel:    ' Invalid Part data basis.
    End Select

    Set oStOfInsulation = Nothing
    Set oEndOfInsulation = Nothing
    Set oCenterPos = Nothing
    Set oPipeComponent = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
