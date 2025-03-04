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
'   Author:          svsmylav
'   Creation Date:  Friday, Jan 4 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                         change description
'   -----------    -----                        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav                Defined new input parameter HandwheelAngle.
'   08.SEP.2006     KKC                     DI-95670   Replace names with initials in all revision history sheets and symbols
'   30.JAN.2008     PK                      CR-135510  Enhanced the existing 3 way ball valve to support dimensional basis for
'                                               (i)Face-to-center dimension basis
'                                               (ii)Face-to-Face dimension and Face-to-Center dimension for port 3
'                                               (iii)Asymmetrical Face-to-Center dimension basis
' 05.May.2008      VRK                      CR-CP-141546  Enhance master 3-way ball valve symbol for Y-type, L-type, and T-type
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI As Double
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Private Function ReturnMax3(A#, B#, C#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If CmpDblLessThan(MaxValue, B) Then MaxValue = B
    If CmpDblLessThan(MaxValue, C) Then MaxValue = C
    ReturnMax3 = MaxValue
End Function

Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If CmpDblLessThan(MaxValue, B) Then MaxValue = B
    If CmpDblLessThan(MaxValue, C) Then MaxValue = C
    If CmpDblLessThan(MaxValue, D) Then MaxValue = D
    ReturnMax4 = MaxValue
End Function

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim pipeDiam2       As Double
    Dim flangeThick2    As Double
    Dim sptOffset2      As Double
    Dim flangeDiam2     As Double
    Dim depth2          As Double
    Dim pipeDiam3       As Double
    Dim flangeThick3    As Double
    Dim sptOffset3      As Double
    Dim flangeDiam3     As Double
    Dim depth3          As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjInsulatedCylinder1 As Object
    Dim ObjInsulatedCylinder2 As Object
    Dim objInsCylinder As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double
    
    Dim oStPoint                 As New AutoMath.DPosition
    Dim oEnPoint                 As New AutoMath.DPosition
    Dim BallDiameter            As Double

    Dim parInsulationDiameter   As Double
    Dim dInsulationDiameter2    As Double
    Dim dInsulationDiameter3    As Double
    Dim parHandwheelAngle   As Double
    
    Dim parFacetoFace       As Double
    Dim parFacetoCenter     As Double
    
    Dim dFace1toCenter      As Double
    Dim dFace2toCenter      As Double
    Dim dFace3toCenter      As Double
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New GeometryFactory
    Dim oEllipse As IngrGeom3D.Ellipse3d
    Dim objEllTop As Object
    Dim dThickness As Double
    Dim dEllipseMajor As Double
    Dim dProjLength As Double
    Dim dMMRatio As Double
    Dim ObjBox As Object
    Dim dLineStrPoints(0 To 14) As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oAxisvec As New AutoMath.DVector

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFace1toCenter = arrayOfInputs(2)
    'parFace2toCenter = arrayOfInputs(3)
    'parFace3toCenter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    parHandwheelAngle = arrayOfInputs(6)
    'parFacetoFace = arrayOfInputs(7)
    'parFacetoCenter = arrayOfInputs(8)
    
    dEllipseMajor = 0.06
    
    iOutput = 0
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                            flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, _
                            flangeThick2, flangeDiam2, sptOffset2, depth2
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, _
                            flangeThick3, flangeDiam3, sptOffset3, depth3
                            
      
    Dim oPipeComponent      As IJDPipeComponent
    Dim lPartdatabasis      As Long
    Set oPipeComponent = oPartFclt
    lPartdatabasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Dim dFlangediam As Double
    Dim dFlangediam2 As Double
    Dim dFlangediam3 As Double
    dFlangediam = flangeDiam
    dFlangediam2 = flangeDiam2
    dFlangediam3 = flangeDiam3
    
    If lPartdatabasis = 963 Or lPartdatabasis = 964 Or _
                lPartdatabasis = 965 Then
        'Checking for zero flange diameter and thickness
        If CmpDblEqual(flangeDiam, 0) Then
            dFlangediam = 1.75 * pipeDiam
        ElseIf CmpDblEqual(flangeDiam2, 0) Then
           dFlangediam2 = 1.75 * pipeDiam2
        ElseIf CmpDblEqual(flangeDiam3, 0) Then
            dFlangediam3 = 1.75 * pipeDiam3
        End If
        
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        ElseIf CmpDblEqual(flangeThick2, 0) Then
            flangeThick2 = NEGLIGIBLE_THICKNESS
        ElseIf CmpDblEqual(flangeThick3, 0) Then
            flangeThick3 = NEGLIGIBLE_THICKNESS
        End If
   End If
 
    Select Case lPartdatabasis
    
        Case Is <= 1, 13                        'Asymmetrical Face-to-Center dimension basis
            parFace1toCenter = arrayOfInputs(2)
            parFace2toCenter = arrayOfInputs(3)
            parFace3toCenter = arrayOfInputs(4)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dFace3toCenter = parFace3toCenter
            If dFace1toCenter = 0 Then
                dFace1toCenter = dFace2toCenter
            ElseIf dFace2toCenter = 0 Then
                dFace2toCenter = dFace1toCenter
            End If
            
        Case 10                                 'Face-to-center dimension basis
            parFacetoCenter = arrayOfInputs(8)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dFace3toCenter = parFacetoCenter
                    
        Case 11                                 'Face-to-Face dimension and Face-to-Center dimension for port 3
            parFace3toCenter = arrayOfInputs(4)
            parFacetoFace = arrayOfInputs(7)
            dFace1toCenter = parFacetoFace / 2
            dFace2toCenter = parFacetoFace / 2
            dFace3toCenter = parFace3toCenter
            
        Case 963                                  '3 Way Ball Valve, T-Type
            parFace3toCenter = arrayOfInputs(4)
            parFacetoFace = arrayOfInputs(7)
            dFace1toCenter = parFacetoFace / 2
            dFace2toCenter = parFacetoFace / 2
            dFace3toCenter = parFace3toCenter
            parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            dInsulationDiameter2 = pipeDiam2 + 2 * parInsulationThickness
            dInsulationDiameter3 = pipeDiam3 + 2 * parInsulationThickness
            
            If CmpDblGreaterthan(dEllipseMajor, dFace1toCenter / 2) Then
                dEllipseMajor = 0.35 * dFace1toCenter
            End If
            dMMRatio = (0.5 * dEllipseMajor + parInsulationThickness) / (dEllipseMajor + parInsulationThickness)
            Set oEllipse = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(m_OutputColl.ResourceManager, _
            CenterPos.x, CenterPos.y + 0.4 * parFacetoFace, CenterPos.z, 0, 1, 0, _
             dEllipseMajor + parInsulationThickness, 0, 0, dMMRatio)
            oAxisvec.Set 0, 1, 0
            Set objEllTop = PlaceProjection(m_OutputColl, oEllipse, oAxisvec, 0.05 * parFacetoFace, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEllTop
            Set oEllipse = Nothing
            Set objEllTop = Nothing
            
            'Top Disc
            oStPoint.Set 0, 0.3 * parFacetoFace, 0
            oEnPoint.Set 0, 0.4 * parFacetoFace, 0
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.5 * parFacetoFace, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCylinder
            Set objInsCylinder = Nothing
        
            oStPoint.Set -0.75 * (dFace1toCenter - flangeThick), 0, 0
            oEnPoint.Set 0.75 * (dFace2toCenter - flangeThick2), 0, 0
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                        1.2 * pipeDiam + 2 * parInsulationThickness, True)
            'Set the output
            m_OutputColl.AddOutput "objInsCylinder", objInsCylinder
            Set objInsCylinder = Nothing
            
            oStPoint.Set 0, 0.3 * parFacetoFace, 0
            oEnPoint.Set 0, -0.3 * parFacetoFace, 0
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.4 * parFacetoFace, True)
            'Set the output
            m_OutputColl.AddOutput "objInsCylinder", objInsCylinder
            Set objInsCylinder = Nothing
            
            oStPoint.Set -0.75 * (dFace1toCenter - flangeThick), 0, 0
            oEnPoint.Set -dFace1toCenter + flangeThick, 0, 0
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
            'Set the output
            m_OutputColl.AddOutput "objInsCylinder", objInsCylinder
            Set objInsCylinder = Nothing
            
            oStPoint.Set 0.75 * (dFace2toCenter - flangeThick2), 0, 0
            oEnPoint.Set dFace2toCenter - flangeThick2, 0, 0
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter2, True)
            'Set the output
            m_OutputColl.AddOutput "objInsCylinder", objInsCylinder
            Set objInsCylinder = Nothing
            
            oStPoint.Set 0, 0, 0.3 * parFacetoFace
            oEnPoint.Set 0, 0, -0.3 * parFacetoFace
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.4 * parFacetoFace, True)
            'Set the output
            m_OutputColl.AddOutput "objInsCylinder", objInsCylinder
            Set objInsCylinder = Nothing
            
            oStPoint.Set 0, 0, 0.3 * parFacetoFace
            oEnPoint.Set 0, 0, oStPoint.z + 0.75 * ((dFace3toCenter - flangeThick3) - 0.3 * parFacetoFace)
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                1.2 * pipeDiam3 + 2 * parInsulationThickness, True)
            'Set the output
            m_OutputColl.AddOutput "objInsCylinder", objInsCylinder
            Set objInsCylinder = Nothing
            
            oStPoint.Set 0, 0, oEnPoint.z
            oEnPoint.Set 0, 0, (dFace3toCenter - flangeThick3)
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter3, True)
            'Set the output
            m_OutputColl.AddOutput "objInsCylinder", objInsCylinder
            Set objInsCylinder = Nothing
                
            'Nozzle Insulation
            oStPoint.Set -dFace1toCenter, 0, 0
            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set -dFace1toCenter + NEGLIGIBLE_THICKNESS, 0, 0
            Else
                If CmpDblLessThan((dFace1toCenter - flangeThick - parInsulationThickness), _
                                                                    LINEAR_TOLERANCE) Then
                    oEnPoint.Set 0, 0, 0
                Else
                    oEnPoint.Set -(dFace1toCenter - flangeThick - parInsulationThickness), 0, 0
                End If
            End If
            
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            If CmpDblGreaterthan(pipeDiam, flangeDiam) Then parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                             parInsulationDiameter, True)
            'Set the output
            m_OutputColl.AddOutput "NozzleInsCylinder_", objInsCylinder
            Set objInsCylinder = Nothing
            
            'Nozzle 2
            oStPoint.Set dFace2toCenter, 0, 0
            If CmpDblEqual(flangeThick2, 0) Then
                oEnPoint.Set (dFace2toCenter - NEGLIGIBLE_THICKNESS), 0, 0
            Else
                If CmpDblLessThan((dFace2toCenter - flangeThick2 - parInsulationThickness), _
                                                                    LINEAR_TOLERANCE) Then
                    oEnPoint.Set 0, 0, 0
                Else
                    oEnPoint.Set (dFace2toCenter - flangeThick2 - parInsulationThickness), 0, 0
                End If
            End If
            parInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
            If CmpDblGreaterthan(pipeDiam2, flangeDiam2) Then parInsulationDiameter = pipeDiam2 + 2 * parInsulationThickness
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                             parInsulationDiameter, True)
            'Set the output
            m_OutputColl.AddOutput "NozzleInsCylinder_", objInsCylinder
            Set objInsCylinder = Nothing
            
            'Nozzle 3
            oStPoint.Set 0, 0, dFace3toCenter
            If CmpDblEqual(flangeThick3, 0) Then
                oEnPoint.Set 0, 0, (dFace3toCenter - NEGLIGIBLE_THICKNESS)
            Else
                If CmpDblLessThan((dFace3toCenter - flangeThick3 - parInsulationThickness), _
                                                                    LINEAR_TOLERANCE) Then
                    oEnPoint.Set 0, 0, 0
                Else
                    oEnPoint.Set 0, 0, (dFace3toCenter - flangeThick3 - parInsulationThickness)
                End If
            End If
            parInsulationDiameter = flangeDiam3 + 2 * parInsulationThickness
            If CmpDblGreaterthan(pipeDiam3, flangeDiam3) Then parInsulationDiameter = pipeDiam3 + 2 * parInsulationThickness
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                             parInsulationDiameter, True)
            'Set the output
            m_OutputColl.AddOutput "NozzleInsCylinder_", objInsCylinder
            Set objInsCylinder = Nothing
            
        Case 964                    '3 Way Ball Valve, L-Type
            parFace3toCenter = arrayOfInputs(4)
            parFacetoFace = arrayOfInputs(7)
            dFace1toCenter = parFacetoFace / 2
            dFace2toCenter = parFacetoFace / 2
            dFace3toCenter = parFace3toCenter
            
            parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            dInsulationDiameter2 = pipeDiam2 + 2 * parInsulationThickness
            dInsulationDiameter3 = pipeDiam3 + 2 * parInsulationThickness
            
            If CmpDblGreaterthan(dEllipseMajor, dFace1toCenter / 2) Then
                dEllipseMajor = 0.35 * dFace1toCenter
            End If
            dMMRatio = (0.5 * dEllipseMajor + parInsulationThickness) / (dEllipseMajor + parInsulationThickness)
            Set oEllipse = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(m_OutputColl.ResourceManager, _
                             CenterPos.x, CenterPos.y + 0.5 * (pipeDiam + dFlangediam) / 2, _
                CenterPos.z, 0, 1, 0, 0, 0, _
                dEllipseMajor + parInsulationThickness, dMMRatio)
            oAxisvec.Set 0, 1, 0
            Set objEllTop = PlaceProjection(m_OutputColl, oEllipse, oAxisvec, 0.05 * parFacetoFace, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objEllTop
            Set oEllipse = Nothing
            Set objEllTop = Nothing
        
            oStPoint.Set -dFace1toCenter + flangeThick, 0, 0
            oEnPoint.Set -0.5 * (dFace1toCenter + flangeThick), 0, 0
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCylinder
            Set objInsCylinder = Nothing
            
            oStPoint.Set dFace2toCenter - flangeThick2, 0, 0
            oEnPoint.Set 0.6 * (dFace2toCenter - flangeThick2), 0, 0
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter2, True)
            'Set the output
            m_OutputColl.AddOutput "objInsCylinder", objInsCylinder
            Set objInsCylinder = Nothing
            
            oStPoint.Set -0.5 * (dFace1toCenter + flangeThick), 0, 0
            oEnPoint.Set 0.4 * (dFace2toCenter - flangeThick2), 0, 0
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.5 * (parInsulationDiameter + dFlangediam + 2 * parInsulationThickness), True)
            'Set the output
            m_OutputColl.AddOutput "objInsCylinder", objInsCylinder
            Set objInsCylinder = Nothing
            
            oStPoint.Set 0, 0, 0
            oEnPoint.Set 0, 0, dFace3toCenter - flangeThick3
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.5 * (dInsulationDiameter3 + flangeDiam3 + 2 * parInsulationThickness), True)
            'Set the output
            m_OutputColl.AddOutput "objInsCylinder", objInsCylinder
            Set objInsCylinder = Nothing
            
            'Box
            dThickness = 0.0625 * 0.0254 '0.0625in
    
            dLineStrPoints(0) = 0.4 * (dFace2toCenter - flangeThick2) - parInsulationThickness
            dLineStrPoints(1) = -0.5 * (flangeDiam2 + 2 * parInsulationThickness)
            dLineStrPoints(2) = 0.5 * (flangeDiam2 + 2 * parInsulationThickness)
            
            dLineStrPoints(3) = dLineStrPoints(0)
            dLineStrPoints(4) = 0.5 * (flangeDiam2 + 2 * parInsulationThickness)
            dLineStrPoints(5) = 0.5 * (flangeDiam2 + 2 * parInsulationThickness)
    
            'back points
            dLineStrPoints(6) = dLineStrPoints(0)
            dLineStrPoints(7) = 0.5 * (flangeDiam2 + 2 * parInsulationThickness)
            dLineStrPoints(8) = -0.5 * (flangeDiam2 + 2 * parInsulationThickness)
    
            dLineStrPoints(9) = dLineStrPoints(0)
            dLineStrPoints(10) = -0.5 * (flangeDiam2 + 2 * parInsulationThickness)
            dLineStrPoints(11) = -0.5 * (flangeDiam2 + 2 * parInsulationThickness)
    
            dLineStrPoints(12) = dLineStrPoints(0)
            dLineStrPoints(13) = dLineStrPoints(1)
            dLineStrPoints(14) = dLineStrPoints(2)
    
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                            5, dLineStrPoints)
            oAxisvec.Set 1, 0, 0
            Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisvec, _
             (0.1 * (dFace2toCenter - flangeThick2) - 0.5 * dThickness) + parInsulationThickness, True)
            'Set the output
            m_OutputColl.AddOutput "ObjInsbox_", ObjBox
            Set ObjBox = Nothing
            
            'Box 2
            dLineStrPoints(0) = 0.5 * (dFace2toCenter - flangeThick2) + 0.5 * dThickness
            dLineStrPoints(1) = -0.5 * (flangeDiam2 + 2 * parInsulationThickness)
            dLineStrPoints(2) = 0.5 * (flangeDiam2 + 2 * parInsulationThickness)
            
            dLineStrPoints(3) = dLineStrPoints(0)
            dLineStrPoints(4) = 0.5 * (flangeDiam2 + 2 * parInsulationThickness)
            dLineStrPoints(5) = 0.5 * (flangeDiam2 + 2 * parInsulationThickness)
    
            'back points
            dLineStrPoints(6) = dLineStrPoints(0)
            dLineStrPoints(7) = 0.5 * (flangeDiam2 + 2 * parInsulationThickness)
            dLineStrPoints(8) = -0.5 * (flangeDiam2 + 2 * parInsulationThickness)
    
            dLineStrPoints(9) = dLineStrPoints(0)
            dLineStrPoints(10) = -0.5 * (flangeDiam2 + 2 * parInsulationThickness)
            dLineStrPoints(11) = -0.5 * (flangeDiam2 + 2 * parInsulationThickness)
    
            dLineStrPoints(12) = dLineStrPoints(0)
            dLineStrPoints(13) = dLineStrPoints(1)
            dLineStrPoints(14) = dLineStrPoints(2)
    
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                            5, dLineStrPoints)
            oAxisvec.Set 1, 0, 0
            dProjLength = (0.1 * (dFace2toCenter - flangeThick2) - 0.5 * dThickness) + parInsulationThickness
            If CmpDblGreaterthan(dLineStrPoints(0) + dProjLength, dFace2toCenter) Then
                dProjLength = dFace2toCenter - dLineStrPoints(0)
            End If
            Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisvec, _
                dProjLength, True)
            'Set the output
            m_OutputColl.AddOutput "ObjInsbox_", ObjBox
            Set ObjBox = Nothing
    
            
            'Nozzle Insulation
            oStPoint.Set -dFace1toCenter, 0, 0
            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set -dFace1toCenter + NEGLIGIBLE_THICKNESS, 0, 0
            Else
                If CmpDblLessThan((dFace1toCenter - flangeThick - parInsulationThickness), _
                                                                    LINEAR_TOLERANCE) Then
                    oEnPoint.Set 0, 0, 0
                Else
                    oEnPoint.Set -(dFace1toCenter - flangeThick - parInsulationThickness), 0, 0
                End If
            End If
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            If CmpDblGreaterthan(pipeDiam, flangeDiam) Then parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                             parInsulationDiameter, True)
            'Set the output
            m_OutputColl.AddOutput "NozzleInsCylinder_", objInsCylinder
            Set objInsCylinder = Nothing
            
            'Nozzle 2
            oStPoint.Set dFace2toCenter, 0, 0
            If CmpDblEqual(flangeThick2, 0) Then
                oEnPoint.Set (dFace2toCenter - NEGLIGIBLE_THICKNESS), 0, 0
            Else
                If CmpDblLessThan((dFace2toCenter - flangeThick2 - parInsulationThickness), _
                                                                    LINEAR_TOLERANCE) Then
                    oEnPoint.Set 0, 0, 0
                Else
                    oEnPoint.Set (dFace2toCenter - flangeThick2 - parInsulationThickness), 0, 0
                End If
            End If
            parInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
            If CmpDblGreaterthan(pipeDiam2, flangeDiam2) Then parInsulationDiameter = pipeDiam2 + 2 * parInsulationThickness
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                             parInsulationDiameter, True)
            'Set the output
            m_OutputColl.AddOutput "NozzleInsCylinder_", objInsCylinder
            Set objInsCylinder = Nothing
            
            'Nozzle 3
            oStPoint.Set 0, 0, dFace3toCenter
            If CmpDblEqual(flangeThick3, 0) Then
                oEnPoint.Set 0, 0, (dFace3toCenter - NEGLIGIBLE_THICKNESS)
            Else
                If CmpDblLessThan((dFace3toCenter - flangeThick3 - parInsulationThickness), _
                                                                    LINEAR_TOLERANCE) Then
                    oEnPoint.Set 0, 0, 0
                Else
                    oEnPoint.Set 0, 0, (dFace3toCenter - flangeThick3 - parInsulationThickness)
                End If
            End If
            parInsulationDiameter = flangeDiam3 + 2 * parInsulationThickness
            If CmpDblGreaterthan(pipeDiam3, flangeDiam3) Then parInsulationDiameter = pipeDiam3 + 2 * parInsulationThickness
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                             parInsulationDiameter, True)
            'Set the output
            m_OutputColl.AddOutput "NozzleInsCylinder_", objInsCylinder
            Set objInsCylinder = Nothing
        
        Case 965                '3 Way Ball Valve, Y-Type
            parFacetoCenter = arrayOfInputs(8)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dFace3toCenter = parFacetoCenter
            
            Dim dInsFlangeDia1 As Double
            Dim dInsFlangeDia2 As Double
            Dim dInsFlangeDia3 As Double
            
            parInsulationDiameter = dFlangediam + parInsulationThickness * 2
            dInsulationDiameter2 = flangeDiam2 + parInsulationThickness * 2
            dInsulationDiameter3 = flangeDiam3 + parInsulationThickness * 2
            dInsFlangeDia1 = parInsulationDiameter / Sqr(2)
            dInsFlangeDia2 = dInsulationDiameter2 / Sqr(2)
            dInsFlangeDia3 = dInsulationDiameter3 / Sqr(2)
    
            dThickness = 0.0625 * 0.0254 '0.0625in
            
            oStPoint.Set -dFace1toCenter + flangeThick, 0, 0
            oEnPoint.Set -(0.5 * dFace1toCenter + flangeThick), 0, 0
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                            pipeDiam + 2 * parInsulationThickness, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCylinder
            Set objInsCylinder = Nothing
    
            dProjLength = flangeThick + parInsulationThickness
            If CmpDblGreaterthan(dProjLength, dFace1toCenter / 2) Then
                dProjLength = dFace1toCenter / 2
            End If
            oStPoint.Set -(0.5 * dFace1toCenter + dProjLength), -0.5 * dInsFlangeDia1, 0.5 * dInsFlangeDia1
            oEnPoint.Set -(0.5 * dFace1toCenter + 0.5 * dThickness), 0.5 * dInsFlangeDia1, -0.5 * dInsFlangeDia1
            Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the output
            m_OutputColl.AddOutput "InsObjbox_", ObjBox
            Set ObjBox = Nothing
    
            oStPoint.Set -0.5 * dFace1toCenter + flangeThick + parInsulationThickness, -0.5 * dInsFlangeDia1, 0.5 * dInsFlangeDia1
            oEnPoint.Set -(0.5 * dFace1toCenter - 0.5 * dThickness), _
                0.5 * dInsFlangeDia1, -0.5 * dInsFlangeDia1
            Set ObjBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            'Set the output
            m_OutputColl.AddOutput "InsObjbox_", ObjBox
            Set ObjBox = Nothing
    
            oStPoint.Set -0.5 * dFace1toCenter + flangeThick, 0, 0
            oEnPoint.Set 0, 0, 0
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 0.5 * (dInsFlangeDia1 + (pipeDiam + 2 * parInsulationThickness)), True)
            'Set the output
            m_OutputColl.AddOutput "InsobjInsCylinder_", objInsCylinder
            Set objInsCylinder = Nothing
    
            'Port 2 body
            oStPoint.Set (dFace2toCenter - flangeThick2) * Cos(PI / 3), 0, -(dFace2toCenter - flangeThick2) * Sin(PI / 3)
            oEnPoint.Set (0.5 * dFace2toCenter + flangeThick2) * Cos(PI / 3), 0, -(0.5 * dFace2toCenter + flangeThick2) * Sin(PI / 3)
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    pipeDiam2 + 2 * parInsulationThickness, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCylinder
            Set objInsCylinder = Nothing
            
            'Line
            dLineStrPoints(0) = (0.5 * dFace2toCenter - 0.5 * dThickness)
            dLineStrPoints(1) = -0.5 * dInsFlangeDia2
            dLineStrPoints(2) = 0.5 * dInsFlangeDia2
            
            dLineStrPoints(3) = dLineStrPoints(0)
            dLineStrPoints(4) = 0.5 * dInsFlangeDia2
            dLineStrPoints(5) = 0.5 * dInsFlangeDia2
    
            'back points
            dLineStrPoints(6) = dLineStrPoints(0)
            dLineStrPoints(7) = 0.5 * dInsFlangeDia2
            dLineStrPoints(8) = -0.5 * dInsFlangeDia2
    
            dLineStrPoints(9) = dLineStrPoints(0)
            dLineStrPoints(10) = -0.5 * dInsFlangeDia2
            dLineStrPoints(11) = -0.5 * dInsFlangeDia2
    
            dLineStrPoints(12) = dLineStrPoints(0)
            dLineStrPoints(13) = dLineStrPoints(1)
            dLineStrPoints(14) = dLineStrPoints(2)
    
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                            5, dLineStrPoints)
            oAxisvec.Set -1, 0, 0
            Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisvec, _
                (flangeThick2 - 0.5 * dThickness) + parInsulationThickness, True)
            Dim oTransMatrix As IJDT4x4
            Set oTransMatrix = New DT4x4
            oTransMatrix.LoadIdentity
            oAxisvec.Set 0, 1, 0
            oTransMatrix.Rotate PI / 3, oAxisvec
            ObjBox.Transform oTransMatrix
            'Set the output
            m_OutputColl.AddOutput "InsObjbox_", ObjBox
            Set ObjBox = Nothing
            
            'Box 2
            dLineStrPoints(0) = (0.5 * dFace2toCenter + 0.5 * dThickness)
            dLineStrPoints(1) = -0.5 * dInsFlangeDia2
            dLineStrPoints(2) = 0.5 * dInsFlangeDia2
            
            dLineStrPoints(3) = dLineStrPoints(0)
            dLineStrPoints(4) = 0.5 * dInsFlangeDia2
            dLineStrPoints(5) = 0.5 * dInsFlangeDia2
    
            'back points
            dLineStrPoints(6) = dLineStrPoints(0)
            dLineStrPoints(7) = 0.5 * dInsFlangeDia2
            dLineStrPoints(8) = -0.5 * dInsFlangeDia2
    
            dLineStrPoints(9) = dLineStrPoints(0)
            dLineStrPoints(10) = -0.5 * dInsFlangeDia2
            dLineStrPoints(11) = -0.5 * dInsFlangeDia2
    
            dLineStrPoints(12) = dLineStrPoints(0)
            dLineStrPoints(13) = dLineStrPoints(1)
            dLineStrPoints(14) = dLineStrPoints(2)
    
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                            5, dLineStrPoints)
            oAxisvec.Set 1, 0, 0
            Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisvec, _
                (flangeThick2 - 0.5 * dThickness) + parInsulationThickness, True)
            Set oTransMatrix = New DT4x4
            oTransMatrix.LoadIdentity
            oAxisvec.Set 0, 1, 0
            oTransMatrix.Rotate PI / 3, oAxisvec
            ObjBox.Transform oTransMatrix
            'Set the output
            m_OutputColl.AddOutput "InsObjbox_", ObjBox
            Set ObjBox = Nothing
            
            oStPoint.Set (0.5 * dFace2toCenter - flangeThick2) * Cos(PI / 3), 0, -(0.5 * dFace2toCenter - flangeThick2) * Sin(PI / 3)
            oEnPoint.Set 0, 0, 0
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                0.5 * (dInsFlangeDia2 + pipeDiam2 + 2 * parInsulationThickness), True)
            'Set the output
            m_OutputColl.AddOutput "InsobjInsCylinder_", objInsCylinder
            Set objInsCylinder = Nothing
    
            'Port 3
            oStPoint.Set (dFace3toCenter - flangeThick3) * Sin(PI / 6), 0, (dFace3toCenter - flangeThick3) * Cos(PI / 6)
            oEnPoint.Set (0.5 * dFace3toCenter + flangeThick3) * Sin(PI / 6), 0, (0.5 * dFace3toCenter + flangeThick3) * Cos(PI / 6)
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    pipeDiam3 + 2 * parInsulationThickness, True)
            'Set the output
            m_OutputColl.AddOutput "ObjInsCylinder", objInsCylinder
            Set objInsCylinder = Nothing
            
            dLineStrPoints(0) = (0.5 * dFace3toCenter - 0.5 * dThickness)
            dLineStrPoints(1) = -0.5 * dInsFlangeDia3
            dLineStrPoints(2) = 0.5 * dInsFlangeDia3
            
            dLineStrPoints(3) = dLineStrPoints(0)
            dLineStrPoints(4) = 0.5 * dInsFlangeDia3
            dLineStrPoints(5) = 0.5 * dInsFlangeDia3
    
            'back points
            dLineStrPoints(6) = dLineStrPoints(0)
            dLineStrPoints(7) = 0.5 * dInsFlangeDia3
            dLineStrPoints(8) = -0.5 * dInsFlangeDia3
    
            dLineStrPoints(9) = dLineStrPoints(0)
            dLineStrPoints(10) = -0.5 * dInsFlangeDia3
            dLineStrPoints(11) = -0.5 * dInsFlangeDia3
    
            dLineStrPoints(12) = dLineStrPoints(0)
            dLineStrPoints(13) = dLineStrPoints(1)
            dLineStrPoints(14) = dLineStrPoints(2)
    
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                            5, dLineStrPoints)
            oAxisvec.Set -1, 0, 0
            Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisvec, _
                (flangeThick3 - 0.5 * dThickness) + parInsulationThickness, True)
            Set oTransMatrix = New DT4x4
            oTransMatrix.LoadIdentity
            oAxisvec.Set 0, -1, 0
            oTransMatrix.Rotate PI / 3, oAxisvec
            ObjBox.Transform oTransMatrix
            'Set the output
            m_OutputColl.AddOutput "InsObjbox_", ObjBox
            Set ObjBox = Nothing
            
            'Box 2
            dLineStrPoints(0) = (0.5 * dFace3toCenter + 0.5 * dThickness)
            dLineStrPoints(1) = -0.5 * dInsFlangeDia3
            dLineStrPoints(2) = 0.5 * dInsFlangeDia2
            
            dLineStrPoints(3) = dLineStrPoints(0)
            dLineStrPoints(4) = 0.5 * dInsFlangeDia3
            dLineStrPoints(5) = 0.5 * dInsFlangeDia3
    
            'back points
            dLineStrPoints(6) = dLineStrPoints(0)
            dLineStrPoints(7) = 0.5 * dInsFlangeDia3
            dLineStrPoints(8) = -0.5 * dInsFlangeDia3
    
            dLineStrPoints(9) = dLineStrPoints(0)
            dLineStrPoints(10) = -0.5 * dInsFlangeDia3
            dLineStrPoints(11) = -0.5 * dInsFlangeDia3
    
            dLineStrPoints(12) = dLineStrPoints(0)
            dLineStrPoints(13) = dLineStrPoints(1)
            dLineStrPoints(14) = dLineStrPoints(2)
    
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                            5, dLineStrPoints)
            oAxisvec.Set 1, 0, 0
            Set ObjBox = PlaceProjection(m_OutputColl, oLineString, oAxisvec, _
                (flangeThick3 - 0.5 * dThickness) + parInsulationThickness, True)
            Set oTransMatrix = New DT4x4
            oTransMatrix.LoadIdentity
            oAxisvec.Set 0, -1, 0
            oTransMatrix.Rotate PI / 3, oAxisvec
            ObjBox.Transform oTransMatrix
            'Set the output
            m_OutputColl.AddOutput "InsObjbox_", ObjBox
            Set ObjBox = Nothing
            
            oStPoint.Set (0.5 * dFace3toCenter - flangeThick3) * Sin(PI / 6), 0, (0.5 * dFace3toCenter - flangeThick3) * Cos(PI / 6)
            oEnPoint.Set 0, 0, 0
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                        0.5 * (dInsFlangeDia3 + pipeDiam3 + 2 * parInsulationThickness), True)
            'Set the output
            m_OutputColl.AddOutput "InsobjInsCylinder_", objInsCylinder
            Set objInsCylinder = Nothing
            
            
            If CmpDblGreaterthan(dEllipseMajor, dFace1toCenter / 2) Then
                dEllipseMajor = 0.35 * dFace1toCenter
            End If
            dMMRatio = (0.5 * dEllipseMajor + parInsulationThickness) / (dEllipseMajor + parInsulationThickness)
            Set oEllipse = oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(m_OutputColl.ResourceManager, _
                CenterPos.x, CenterPos.y + 0.5 * (dFlangediam / Sqr(2) + pipeDiam) / 2, CenterPos.z, 0, 1, 0, 0, _
                0, dEllipseMajor + parInsulationThickness, dMMRatio)
            oAxisvec.Set 0, 1, 0
            Set objEllTop = PlaceProjection(m_OutputColl, oEllipse, oAxisvec, 0.05 * dFace3toCenter, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput "Elliptical Top", objEllTop
            Set oEllipse = Nothing
            Set objEllTop = Nothing
            
            'Nozzle
            oStPoint.Set -dFace1toCenter, 0, 0
            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set -dFace1toCenter + NEGLIGIBLE_THICKNESS, 0, 0
            Else
                If CmpDblLessThan((dFace1toCenter - flangeThick - parInsulationThickness), _
                                                                    LINEAR_TOLERANCE) Then
                    oEnPoint.Set 0, 0, 0
                Else
                    oEnPoint.Set -(dFace1toCenter - flangeThick - parInsulationThickness), 0, 0
                End If
            End If
            parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
            If CmpDblGreaterthan(pipeDiam, flangeDiam) Then parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                             parInsulationDiameter, True)
            'Set the output
            m_OutputColl.AddOutput "NozzleInsCylinder_", objInsCylinder
            Set objInsCylinder = Nothing
            
            'Nozzle 2
            oStPoint.Set dFace2toCenter * Cos(PI / 3), 0, -dFace2toCenter * Sin(PI / 3)
            If CmpDblEqual(flangeThick2, 0) Then
                oEnPoint.Set (dFace2toCenter - NEGLIGIBLE_THICKNESS) * Cos(PI / 3), 0, _
                           -(dFace2toCenter - NEGLIGIBLE_THICKNESS) * Sin(PI / 3)
            Else
                If CmpDblLessThan((dFace2toCenter - flangeThick2 - parInsulationThickness), LINEAR_TOLERANCE) Then
                    oEnPoint.Set 0, 0, 0
                Else
                    oEnPoint.Set (dFace2toCenter - flangeThick2 - parInsulationThickness) * Cos(PI / 3), 0, _
                        -(dFace2toCenter - flangeThick2 - parInsulationThickness) * Sin(PI / 3)
                End If
            End If
            parInsulationDiameter = flangeDiam2 + 2 * parInsulationThickness
            If CmpDblGreaterthan(pipeDiam2, flangeDiam2) Then parInsulationDiameter = pipeDiam2 + 2 * parInsulationThickness
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                             parInsulationDiameter, True)
            'Set the output
            m_OutputColl.AddOutput "NozzleInsCylinder_", objInsCylinder
            Set objInsCylinder = Nothing
            
            'Nozzle 3
            oStPoint.Set dFace3toCenter * Sin(PI / 6), 0, dFace3toCenter * Cos(PI / 6)
            If CmpDblEqual(flangeThick3, 0) Then
                oEnPoint.Set (dFace3toCenter - NEGLIGIBLE_THICKNESS) * Sin(PI / 6), 0, _
                            (dFace3toCenter - NEGLIGIBLE_THICKNESS) * Cos(PI / 6)
            Else
                If CmpDblLessThan((dFace3toCenter - flangeThick3 - parInsulationThickness), LINEAR_TOLERANCE) Then
                    oEnPoint.Set 0, 0, 0
                Else
                    oEnPoint.Set (dFace3toCenter - flangeThick3 - parInsulationThickness) * Sin(PI / 6), 0, _
                        (dFace3toCenter - flangeThick3 - parInsulationThickness) * Cos(PI / 6)
                End If
            End If
            parInsulationDiameter = flangeDiam3 + 2 * parInsulationThickness
            If CmpDblGreaterthan(pipeDiam3, flangeDiam3) Then parInsulationDiameter = pipeDiam3 + 2 * parInsulationThickness
            Set objInsCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                             parInsulationDiameter, True)
            'Set the output
            m_OutputColl.AddOutput "NozzleInsCylinder_", objInsCylinder
            Set objInsCylinder = Nothing
            
        Case Else
            GoTo ErrorLabel:
            
    End Select
    
    If lPartdatabasis <= 1 Or lPartdatabasis = 10 Or _
       lPartdatabasis = 11 Or lPartdatabasis = 13 Then

        'Insert your code for output 5(Insulation for Cylinder1)
        'Maximum of pipeDiam1, flangeDiam1, pipeDiam2 and flangeDiam2 is used to compute
        'insulation diameter
        parInsulationDiameter = ReturnMax4(pipeDiam, flangeDiam, pipeDiam2, flangeDiam2) + _
                                2 * parInsulationThickness
        oStPoint.Set CenterPos.x - dFace1toCenter, _
                    CenterPos.y, _
                    CenterPos.z
        oEnPoint.Set CenterPos.x + dFace2toCenter, _
                    CenterPos.y, _
                    CenterPos.z
        Set ObjInsulatedCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCylinder1
        Set ObjInsulatedCylinder1 = Nothing
        
        'Insert your code for output 6(Insulation for Cylinder2)
        oStPoint.Set CenterPos.x, _
                    CenterPos.y + dFace3toCenter, _
                    CenterPos.z
        'Extend the insulation along negative Y axis by Ball radius plus parInsulationThickness
        '(Ball diameter is considered to be half of the maximum face to center distance)
        BallDiameter = 0.5 * ReturnMax3(dFace1toCenter, dFace2toCenter, dFace3toCenter)
        oEnPoint.Set CenterPos.x, _
                    CenterPos.y - BallDiameter / 2 - parInsulationThickness, _
                    CenterPos.z
                    
        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        End If
        
        'Insulation diameter for port 3 is equated to twice the minimum of Face 1 to Center
        'and Face 2 to Center
        Dim dMinValue As Double
        If CmpDblLessThan(dFace1toCenter, dFace2toCenter) Then
            dMinValue = dFace1toCenter
        Else
            dMinValue = dFace2toCenter
        End If
        
        If CmpDblGreaterthan(parInsulationDiameter / 2, dMinValue) Then
            parInsulationDiameter = 2 * dMinValue
        End If
        
        Set ObjInsulatedCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCylinder2
        Set ObjInsulatedCylinder2 = Nothing
    End If
    
        Set oGeomFactory = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oTransMatrix = Nothing
        Set oAxisvec = Nothing
        Set CenterPos = Nothing
        Set oLineString = Nothing
        
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
