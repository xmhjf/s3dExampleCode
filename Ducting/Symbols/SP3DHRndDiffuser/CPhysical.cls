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

'   Author:         kkk
'   Creation Date:  Thursday 2, March 2006
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       HVAC Round Fixed Discharge Diffuser.
'       The Symbol is taken from Carnes Catalog and Model Number is Fixed Discharge
'       Model SSEA
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
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
    Const NEGLIGIBLE_THICKNESS = 0.0001
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CPos       As New AutoMath.DPosition
    Dim Collarheight As Double
    Dim parNeckDia As Double
      
    CPos.Set 0, 0, 0

    Dim iOutput     As Double
    Dim ObjFormProjection As Object
    Dim ObjNeck As Object

    Dim Width As Double
    Dim ThroatRadius As Double
    Dim dblOpening As Double
    Dim FrameOuterDia As Double
    Dim dblFormProjecHeight As Double
    Dim FormProjection As Double
    Dim dblFormProjecHalfHeight As Double
    Dim OverallProjection As Double
    Dim dblConeLength1 As Double
    Dim dblConeLength2 As Double
    Dim dTheta As Double
    Dim PI As Double
    Dim pw As Double
    Dim ph2 As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Width = arrayOfInputs(2) 'B' - Dimension
    FrameOuterDia = arrayOfInputs(3) 'C - Dimension
    FormProjection = arrayOfInputs(4) 'D -Dimension
    OverallProjection = arrayOfInputs(5) 'E -Dimension
    Collarheight = arrayOfInputs(6) 'G-Dimension
    
    parNeckDia = Width
    FrameOuterDia = FrameOuterDia
    
    dblOpening = (FrameOuterDia - parNeckDia) / 2
    dblFormProjecHalfHeight = FrameOuterDia / 2
    
    pw = FormProjection
    ph2 = dblFormProjecHalfHeight
    
    dblConeLength2 = OverallProjection
    dblConeLength1 = dblConeLength2 * 0.6
    
    dTheta = Atn(ph2 / pw)
    iOutput = 0
    
' Insert your code for output 1(Bell Mouth)
    Dim NeckStart As New AutoMath.DPosition
    Dim NeckEnd As New AutoMath.DPosition

    NeckStart.Set CPos.X, CPos.Y, CPos.Z
    NeckEnd.Set CPos.X + Collarheight, CPos.Y, CPos.Z
    
    Set ObjNeck = PlaceCylinder(m_OutputColl, NeckStart, NeckEnd, parNeckDia, False)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNeck
    Set ObjNeck = Nothing
    
    
 ' Insert your code for output 2(Diffuser Outer Profile)
    Dim p(1 To 19)   As New AutoMath.DPosition
    Dim objBspline  As IngrGeom3D.BSplineCurve3d
    
    Dim CP As New AutoMath.DPosition 'revolution center point
    Dim CV As New AutoMath.DVector 'rotation vector for rotation
    CP.Set Collarheight, 0, 0
    CV.Set 1, 0, 0
       
    PI = 4 * Atn(1)
    
    p(1).Set CPos.X + Collarheight, CPos.Y + parNeckDia / 2, CPos.Z
    p(2).Set p(1).X, p(1).Y + dblOpening / 2, p(1).Z
    p(3).Set p(1).X, p(2).Y + dblOpening / 2, p(2).Z
    
    p(4).Set p(3).X + pw / 5, p(3).Y + pw / 5 * Tan(dTheta), p(3).Z
    p(5).Set p(3).X + pw / 4, p(3).Y + pw / 4 * Tan(dTheta), p(3).Z
    p(6).Set p(3).X + pw / 3, p(3).Y + pw / 3 * Tan(dTheta), p(3).Z
    p(7).Set p(3).X + pw / 2, p(3).Y + pw / 2 * Tan(dTheta), p(3).Z
    
    p(8).Set p(3).X + (pw * 2 / 3), p(3).Y + (pw * 2 / 3) * Tan(dTheta), p(3).Z
    p(9).Set p(3).X + (pw * 3 / 4), p(3).Y + (pw * 3 / 4) * Tan(dTheta), p(3).Z
    p(10).Set p(3).X + (pw * 4 / 5), p(3).Y + (pw * 4 / 5) * Tan(dTheta), p(3).Z
    p(11).Set p(3).X + (pw), p(3).Y + (pw) * Tan(dTheta), p(3).Z
    
    p(12).Set p(3).X + (pw * 4 / 5), p(3).Y + ph2 + (pw * 1 / 5) * Tan(dTheta), p(3).Z
    p(13).Set p(3).X + (pw * 3 / 4), p(3).Y + ph2 + (pw * 1 / 4) * Tan(dTheta), p(3).Z
    p(14).Set p(3).X + (pw * 2 / 3), p(3).Y + ph2 + (pw * 1 / 3) * Tan(dTheta), p(3).Z
    p(15).Set p(3).X + (pw / 2), p(3).Y + ph2 + (pw / 2) * Tan(dTheta), p(3).Z
    p(16).Set p(3).X + (pw / 3), p(3).Y + ph2 + (pw * 2 / 3) * Tan(dTheta), p(3).Z
    p(17).Set p(3).X + (pw / 4), p(3).Y + ph2 + (pw * 3 / 4) * Tan(dTheta), p(3).Z
    p(18).Set p(3).X + (pw / 5), p(3).Y + ph2 + (pw * 4 / 5) * Tan(dTheta), p(3).Z
    p(19).Set p(3).X, p(3).Y + ph2 * 2, p(3).Z
    
     
    Set objBspline = PlaceTrBspline(5, p)
    
    Set ObjFormProjection = PlaceRevolution(m_OutputColl, objBspline, CV, CP, PI * 2, False)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFormProjection
    Set ObjFormProjection = Nothing
    Set objBspline = Nothing
   
' Insert your code for output 3(Diffuser Cone - 1)

    Dim pt(1 To 11)   As New AutoMath.DPosition
    Dim objDiffCone1  As IngrGeom3D.BSplineCurve3d
    Dim ObjDiffConeProj1 As Object
    
    Dim conex As Double
    conex = (dblOpening + ph2) * 0.7
    
    Dim alpha As Double
    alpha = Atn(conex / dblConeLength1)
    
    
    Dim cl1 As Double
    cl1 = dblConeLength1
    
    CP.Set Collarheight, 0, 0
    CV.Set 1, 0, 0
    
    pt(1).Set CPos.X + Collarheight, CPos.Y + (parNeckDia / 2) * 0.5, CPos.Z
    pt(2).Set pt(1).X + Collarheight * 0.1, pt(1).Y, pt(1).Z
    pt(3).Set pt(1).X + Collarheight * 0.15, pt(1).Y, pt(1).Z
    pt(4).Set pt(3).X, pt(1).Y + (parNeckDia / 2) * 0.6, pt(1).Z
    pt(5).Set pt(3).X, pt(1).Y + (parNeckDia / 2) * 0.8, pt(1).Z
    
    pt(6).Set pt(5).X + cl1 / 4, pt(5).Y + (cl1 / 4) * Tan(alpha), pt(5).Z
    pt(7).Set pt(5).X + cl1 / 2, pt(5).Y + (cl1 / 2) * Tan(alpha), pt(5).Z
    pt(8).Set pt(5).X + cl1 * 3 / 4, pt(5).Y + (cl1 * 3 / 4) * Tan(alpha), pt(5).Z
    pt(9).Set pt(5).X + cl1, pt(5).Y + cl1 * Tan(alpha), pt(5).Z
    pt(10).Set pt(9).X, pt(9).Y + parNeckDia * 0.05, pt(5).Z
    pt(11).Set pt(9).X, pt(9).Y + parNeckDia * 0.1, pt(5).Z
    
    Set objDiffCone1 = PlaceTrBspline(5, pt)
    
    Set ObjDiffConeProj1 = PlaceRevolution(m_OutputColl, objDiffCone1, CV, CP, PI * 2, False)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffConeProj1
    Set ObjDiffConeProj1 = Nothing
    Set objDiffCone1 = Nothing
    
' Insert your code for output 5(Diffuser Cone - 2)

    Dim pnt(1 To 9)   As New AutoMath.DPosition
    Dim objDiffCone2  As IngrGeom3D.BSplineCurve3d
    Dim ObjDiffConeProj2 As Object
    Dim cl2 As Double
    cl2 = dblConeLength2
    
    Dim conex2 As Double
    conex2 = conex * 0.8
    
    Dim alpha2 As Double
    alpha2 = Atn(conex2 / dblConeLength2)
    
    CP.Set Collarheight, 0, 0
    CV.Set 1, 0, 0
    
    pnt(1).Set Collarheight * 1.1, (parNeckDia / 2) * 0.1, 0
    pnt(2).Set pnt(1).X, pnt(1).Y + (parNeckDia / 2) * 0.2, pt(1).Z
    pnt(3).Set pnt(1).X, pnt(1).Y + (parNeckDia / 2) * 0.4, pt(1).Z
    
    pnt(4).Set pnt(1).X + cl2 / 4, pnt(3).Y + (cl2 / 4) * Tan(alpha2), pnt(1).Z
    pnt(5).Set pnt(1).X + cl2 / 2, pnt(3).Y + (cl2 / 2) * Tan(alpha2), pnt(1).Z
    pnt(6).Set pnt(1).X + cl2 * 3 / 4, pnt(3).Y + (cl2 * 3 / 4) * Tan(alpha2), pnt(1).Z
    pnt(7).Set pnt(1).X + cl2, pnt(3).Y + cl2 * Tan(alpha2), pnt(5).Z
    pnt(8).Set pnt(7).X, pnt(7).Y + parNeckDia * 0.1, pnt(1).Z
    pnt(9).Set pnt(7).X, pnt(7).Y + parNeckDia * 0.15, pnt(1).Z
    
    Set objDiffCone2 = PlaceTrBspline(5, pnt)
    Set ObjDiffConeProj2 = PlaceRevolution(m_OutputColl, objDiffCone2, CV, CP, PI * 2, False)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDiffConeProj2
    Set ObjDiffConeProj2 = Nothing
    Set objDiffCone2 = Nothing
    
'=====================================
'BUILD HVAC NOZZLE: Diffuser Port
'=====================================
    Dim NozzleFactory As New GSCADNozzleEntities.NozzleFactory
    Dim oHvacNozzle As GSCADNozzleEntities.HvacNozzle
    Dim iNozzle As GSCADNozzleEntities.IJDNozzle
    Dim iDistribPort As GSCADNozzleEntities.IJDistribPort
    Dim Depth As Double
    Dim EndPrep As Long
    Dim FlowDir As DistribFlow
    Dim PortStatus As DistribPortStatus
    Dim DimBaseOuter As Boolean
    Dim PortDepth As Double
    Dim cptOffset As Double
    Dim pos As New AutoMath.DPosition
    Dim dir As New AutoMath.DVector
    Dim iPortIndex As Integer
    Dim CornerRadius As Double
    Dim FlangeWidth As Double
    Dim NozzleLength As Double, Thickness As Double

    'Set HVAC nozzle parameters
    iPortIndex = 1
    EndPrep = 301
    CornerRadius = 0#
    Thickness = 0#
    NozzleLength = Thickness + NEGLIGIBLE_THICKNESS
    PortDepth = 0#
    cptOffset = 0#
    Depth = Width
    ' To construct nozzle as crosssection only, use FlangeWidth of 0, and
    ' some non-zero value for flange thickness
    ' Flange thickness assumed to be negigible thickness
    FlowDir = DistribFlow_IN

    ' Depth and Width of crosssection will be the same as diameter of the nozzle which
    ' is the inletDia

    DimBaseOuter = True
    PortStatus = DistribPortStatus_BASE
    Set oHvacNozzle = NozzleFactory.CreateHvacNozzle(iPortIndex, "SymbDefn", GSCADNozzleEntities.Round, EndPrep, _
                                            Thickness, FlangeWidth, FlowDir, Width, _
                                            Depth, CornerRadius, DimBaseOuter, PortStatus, _
                                            "DiffuserPort", PortDepth, cptOffset, False, m_OutputColl.ResourceManager)
    'Position of the nozzle should be the conenct point of the nozzle
    Dim X As Double, Y As Double, Z As Double
    X = 0#
    Y = 0#
    Z = 0#
    pos.Set X, Y, Z
    Set iDistribPort = oHvacNozzle
    iDistribPort.SetPortLocation pos

    'Direction specified here of the nozzle should be the direction in which pipe will be routed.
    'Graphics of the nozzle will appear in opposite direction to the direction specified on the nozzle.
    dir.Set -1, 0, 0
    iDistribPort.SetDirectionVector dir

    dir.Set 0, 1, 0
    iDistribPort.SetRadialOrient dir

    Set iNozzle = oHvacNozzle
    iNozzle.Length = NozzleLength

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHvacNozzle
    Set oHvacNozzle = Nothing
    Set iNozzle = Nothing
    Set iDistribPort = Nothing
    Set CPos = Nothing
    Set NeckStart = Nothing
    Set NeckEnd = Nothing
    Set CP = Nothing
    Set CV = Nothing
    Set NozzleFactory = Nothing
    Set pos = Nothing
    Set dir = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    End Sub
