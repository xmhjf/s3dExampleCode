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
Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double

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
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, PipeDia, flangeThick, FlangeDia, _
                                                CptOffset, depth
 
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
    
    ' Place body of sphere
    ' -------------------------------------------------------------------------------------
    Dim objBodySphere As IngrGeom3D.Sphere3d
    Dim dX As Double
    dX = Sqr(((0.8 * PipeDia) ^ 2) - ((0.6 * PipeDia) ^ 2))
    oCenPoint.Set -dX, 0, 0
    Set objBodySphere = PlaceSphere(m_OutputColl, oCenPoint, PipeDia * 0.8)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBodySphere
    Set objBodySphere = Nothing
    
    'Create the Lift mechanism
    '-----------------------------------------------------------------------------------
    Dim objLiftMech As Object
    Dim oLineStr As IngrGeom3D.LineString3d
    Dim dLineStrPts(0 To 17) As Double
    
    dLineStrPts(0) = 0
    dLineStrPts(1) = 0
    dLineStrPts(2) = 0
    
    dLineStrPts(3) = 0
    dLineStrPts(4) = 0.6 * PipeDia
    dLineStrPts(5) = dLineStrPts(2)
    
    dLineStrPts(6) = 0.5 * parFacetoEnd
    dLineStrPts(7) = dLineStrPts(4)
    dLineStrPts(8) = dLineStrPts(2)
    
    dLineStrPts(9) = 0.9 * parFacetoEnd
    dLineStrPts(10) = 0.8 * PipeDia
    dLineStrPts(11) = dLineStrPts(2)
    
    dLineStrPts(12) = parFacetoEnd
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
    
    'Remove the References
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenPoint = Nothing
    Set oVector = Nothing
    Set oGeomFactory = Nothing
    
    
    ' Place Nozzle 1 (Inlet)
    ' ------------------------------------------------------------------------------------------
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -(parFacetoCenter + CptOffset - depth), 0, 0
    oDir.Set -1, 0, 0

    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                                oPlacePoint, parFacetoCenter)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    
    Set objNozzle = Nothing

    ' Place Nozzle 2 (Outlet)
    ' ------------------------------------------------------------------------------------------
    RetrieveParameters 2, oPartFclt, m_OutputColl, _
                       PipeDia, flangeThick, FlangeDia, CptOffset, depth
    
    oPlacePoint.Set 0, parFacetoCenter + CptOffset - depth, 0
    oDir.Set 0, 1, 0

    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                                                    oPlacePoint, parFacetoCenter)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oCenPoint = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
