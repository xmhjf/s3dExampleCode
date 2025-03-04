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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          Jang Choi SHI
'   Creation Date:  Thursday, Oct 04 1999
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   08.Aug.2003     ACM(caddagud)         TR- 45826  A value of 0.0000001 is used in Z direction to
'                                           avoid symbol placement failure(gives assertion errors) for all three cones.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
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
    
' For Test ----------------------------------------------------------------
'''    OperatorWidth = arrayOfInputs(1)
'''    OperatorHeight = arrayOfInputs(2)
'''    OperatorDia = arrayOfInputs(3)
'''    RadiusBody = arrayOfInputs(4)
'''
'''    PipeDia = 0.1
'''    FlangeThick = 0.024
'''    FlangeDia = 0.21
'''    CptOffset = 0.001
' -------------------------------------------------------------------------
    
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
    
    Dim objInlet As IngrGeom3D.Cone3d
    Set objInlet = PlaceCone(m_OutputColl, StartPos, EndPos, 0.02, PipeDia / 2)
    
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInlet
    
    Set objInlet = Nothing
    
 ' Place Outlet
 '---------------------------------------------------------------------------------------
    EndPos.Set StartPos.x, StartPos.y - OperatorWidth + FlangeThick, StartPos.z
    
    Dim objOutlet As IngrGeom3D.Cone3d
    Set objOutlet = PlaceCone(m_OutputColl, StartPos, EndPos, 0.02, PipeDia / 2)
    
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
