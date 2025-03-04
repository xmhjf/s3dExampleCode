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
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Monday, Apr 24 2006
'   Description:
'        Angle Valve
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   24.Apr.2006     SymbolTeam(India)       DI-94663  New symbol is prepared from existing
'                                           GSCAD symbol.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim PipeDia As Double
    Dim FlangeThick As Double
    Dim CptOffset As Double
    Dim FlangeDia As Double
    Dim Depth As Double
    
    Dim iOutput As Double
    
    Dim OperatorWidth As Double
    Dim OperatorHeight As Double
    Dim OperatorDia As Double
    Dim RadiusBody As Double

    Dim HandleThick As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    OperatorWidth = arrayOfInputs(2)
    OperatorHeight = arrayOfInputs(3)
    OperatorDia = arrayOfInputs(4)
'''    RadiusBody = arrayOfInputs(5)

    RetrieveParameters 1, oPartFclt, m_OutputColl, _
                       PipeDia, FlangeThick, FlangeDia, CptOffset, Depth
    
    HandleThick = OperatorDia * 0.1
    
    iOutput = 0

    Dim CenterPos As New AutoMath.DPosition
    CenterPos.Set 0, 0, 0
    
 ' Place body of valve
 ' -------------------------------------------------------------------------------------

    Dim objBodyV As IngrGeom3D.Sphere3d
    RadiusBody = PipeDia * 0.7
    Set objBodyV = PlaceSphere(m_OutputColl, CenterPos, RadiusBody)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBodyV
    Set objBodyV = Nothing
    
 ' Place Handle
 ' --------------------------------------------------------------------------------
    
    Dim CenBase As New AutoMath.DPosition
    Dim CenTop As New AutoMath.DPosition
    
    CenBase.Set CenterPos.x, CenterPos.y + OperatorHeight - HandleThick, CenterPos.z
    
    CenTop.Set CenterPos.x, CenterPos.y, CenterPos.z

    Dim objHandle As IngrGeom3D.Cone3d
    Set objHandle = PlaceCone(m_OutputColl, CenBase, CenTop, OperatorDia / 2, 0)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandle
    
    Set objHandle = Nothing
    Set CenBase = Nothing
    Set CenTop = Nothing
    
 ' Place Handle Plate
 '---------------------------------------------------------------------------------------

    Dim StartPos As New AutoMath.DPosition
    Dim EndPos As New AutoMath.DPosition
    
    StartPos.Set CenterPos.x, CenterPos.y + OperatorHeight - HandleThick, CenterPos.z
    EndPos.Set StartPos.x, StartPos.y + HandleThick, StartPos.z
    
    Dim objHandlePlate As IngrGeom3D.Projection3d
    Set objHandlePlate = PlaceCylinder(m_OutputColl, StartPos, EndPos, OperatorDia, True)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHandlePlate
    
    Set objHandlePlate = Nothing
    
 ' Place Inlet
 '---------------------------------------------------------------------------------------
 
    StartPos.Set CenterPos.x, CenterPos.y, CenterPos.z
    EndPos.Set StartPos.x + OperatorWidth - FlangeThick, StartPos.y, StartPos.z
    
    Dim objInlet As IngrGeom3D.Projection3d
    Set objInlet = PlaceCylinder(m_OutputColl, StartPos, EndPos, PipeDia, True)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInlet
    
    Set objInlet = Nothing
    
 ' Place Outlet
 '---------------------------------------------------------------------------------------
 
    EndPos.Set StartPos.x, StartPos.y - OperatorWidth + FlangeThick, StartPos.z
    
    Dim objOutlet As IngrGeom3D.Projection3d
    Set objOutlet = PlaceCylinder(m_OutputColl, StartPos, EndPos, PipeDia, True)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objOutlet
    
    Set objOutlet = Nothing
    
    Set StartPos = Nothing
    Set EndPos = Nothing
    
 ' Place Nozzle 1
 ' ------------------------------------------------------------------------------------------
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set OperatorWidth + CptOffset - Depth, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

 ' Place Nozzle 2
 ' ------------------------------------------------------------------------------------------
    RetrieveParameters 2, oPartFclt, m_OutputColl, _
                       PipeDia, FlangeThick, FlangeDia, CptOffset, Depth
    
    oPlacePoint.Set 0, -OperatorWidth - CptOffset + Depth, 0
    oDir.Set 0, -1, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
