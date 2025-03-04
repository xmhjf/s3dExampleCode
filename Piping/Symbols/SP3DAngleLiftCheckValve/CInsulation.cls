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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RUK
'   Creation Date:  Monday, Nov 26 2007
'   Description:
'           CR-130389  Provide symbols for angle pressure seal lift check valve and operator
'           Source:
'   Change History:
'   dd.mmm.yyyy     who      change description
'   -----------     -----    ------------------
'   26.Nov.2007     RUK       Created
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
    Dim PipeDia As Double
    Dim flangeThick As Double
    Dim CptOffset As Double
    Dim FlangeDia As Double
    Dim depth As Double
    
    Dim iOutput As Double
    
    Dim parFacetoCenter As Double
    Dim parFacetoEnd As Double
    Dim parInsulationThickness As Double
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parFacetoEnd = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)

    iOutput = 0
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oCenPoint As AutoMath.DPosition
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oVector As AutoMath.DVector

    Set oGeomFactory = New GeometryFactory
    Set oCenPoint = New DPosition
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    Set oVector = New DVector
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, PipeDia, flangeThick, FlangeDia, _
                                                CptOffset, depth
                                                    
    'Create the Insulation for body
    Dim objInsBody As Object
    Dim dX As Double
    dX = Sqr(((0.8 * PipeDia) ^ 2) - ((0.6 * PipeDia) ^ 2))
    oCenPoint.Set -dX, 0, 0
    Set objInsBody = PlaceSphere(m_OutputColl, oCenPoint, 0.8 * PipeDia + parInsulationThickness)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBody
    Set objInsBody = Nothing
    
    'Create the insulation for Lift Mechanism
    Dim objLiftMech As Object
    Dim oLineStr As IngrGeom3D.LineString3d
    Dim dLineStrPts(0 To 17) As Double
    
    dLineStrPts(0) = 0
    dLineStrPts(1) = 0
    dLineStrPts(2) = 0

    dLineStrPts(3) = 0
    dLineStrPts(4) = 0.6 * PipeDia + parInsulationThickness
    dLineStrPts(5) = dLineStrPts(2)

    dLineStrPts(6) = 0.5 * parFacetoEnd
    dLineStrPts(7) = dLineStrPts(4)
    dLineStrPts(8) = dLineStrPts(2)

    dLineStrPts(9) = 0.9 * parFacetoEnd
    dLineStrPts(10) = 0.8 * PipeDia + parInsulationThickness
    dLineStrPts(11) = dLineStrPts(2)

    dLineStrPts(12) = parFacetoEnd + parInsulationThickness
    dLineStrPts(13) = dLineStrPts(10)
    dLineStrPts(14) = dLineStrPts(2)

    dLineStrPts(15) = dLineStrPts(12)
    dLineStrPts(16) = dLineStrPts(1)
    dLineStrPts(17) = dLineStrPts(2)
    
    Set oLineStr = New LineString3d
    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 6, dLineStrPts)

    oVector.Set 1, 0, 0
    oCenPoint.Set 0, 0, 0
    Set objLiftMech = PlaceRevolution(m_OutputColl, oLineStr, oVector, oCenPoint, 2 * PI, True)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLiftMech
    Set oLineStr = Nothing
    Set objLiftMech = Nothing
    
    'Create the insulation for Inlet
    Dim objInsInlet As Object
    oStPoint.Set 0, 0, 0
    oEnPoint.Set -parFacetoCenter, 0, 0
    Set objInsInlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                            PipeDia + 2 * parInsulationThickness, True)
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsInlet
    Set objInsInlet = Nothing
    
    'Create the Insulation for Outlet
    Dim objInsOutlet As Object
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, parFacetoCenter, 0
    Set objInsOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                        PipeDia + 2 * parInsulationThickness, True)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsOutlet
    Set objInsOutlet = Nothing
    
    'Create the Insulation for Inlet Nozzle
    Dim dInsDia As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, PipeDia, flangeThick, FlangeDia, _
                                                                    CptOffset, depth
    oStPoint.Set -parFacetoCenter, 0, 0
    If CmpDblEqual(flangeThick, LINEAR_TOLERANCE) Then
        oEnPoint.Set -(parFacetoCenter - NEGLIGIBLE_THICKNESS), 0, 0
        dInsDia = PipeDia + 2 * parInsulationThickness
    Else
        oEnPoint.Set -(parFacetoCenter - flangeThick), 0, 0
        dInsDia = FlangeDia + 2 * parInsulationThickness
    End If
    Set objInsInlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsInlet
    Set objInsInlet = Nothing
    
    'Create the Insulation for Outlet Nozzle
    RetrieveParameters 2, oPartFclt, m_OutputColl, PipeDia, flangeThick, FlangeDia, _
                                                                    CptOffset, depth
    oStPoint.Set 0, parFacetoCenter, 0
    If CmpDblEqual(flangeThick, LINEAR_TOLERANCE) Then
        oEnPoint.Set 0, (parFacetoCenter - NEGLIGIBLE_THICKNESS), 0
        dInsDia = PipeDia + 2 * parInsulationThickness
    Else
        oEnPoint.Set 0, (parFacetoCenter - flangeThick), 0
        dInsDia = FlangeDia + 2 * parInsulationThickness
    End If
    Set objInsOutlet = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
        
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsOutlet
    Set objInsOutlet = Nothing
        
    'Remove the References
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenPoint = Nothing
    Set oVector = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
