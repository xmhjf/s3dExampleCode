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
'   Copyright (c) 2003-04, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Friday, Jan 4 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   03.Dec.2004     svsmylav                Defined new input parameter HandwheelAngle.
'   08.SEP.2006      KKC  DI-95670          Replace names with initials in all revision history sheets and symbols
'   11.OCT.2007      PK   CR-128045         Enhance tank drain valve to provide option for
'                                           Saunders tank bottom valve.
'   28.FEB.2008      PK   CR-135703         Update the tank drain valve per following options
'                                            1)Tank drain valve, specified by face-to-center 1, face-to-center 2, and angle
'                                            2)Tank drain valve, specified by tangent length 1, face-to-center 2, and angle
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI  As Double
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput                 As Double
    Dim ObjBodyCone1            As Object
    Dim ObjBodyCone2            As Object
    Dim ObjInsulatedCylinder1   As Object
    Dim ObjInsulatedCylinder2   As Object
    
    Dim parFace1toCenter        As Double
    Dim parFace2toCenter        As Double
    Dim parInsulationThickness  As Double

    'Variables defined by svsmylav
    Dim stPoint                 As New AutoMath.DPosition
    Dim enPoint                 As New AutoMath.DPosition
    Dim parInsDiameterPort      As Double
    Dim parInsulationDiameter   As Double
    
    Dim parHandwheelAngle       As Double
    
    Dim parAngle                As Double
    Dim parTangentLength1       As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFace1toCenter = arrayOfInputs(2)
    'parFace2toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    parHandwheelAngle = arrayOfInputs(5)
    'parTangentLength1 = arrayOfInputs(6)
    'parAngle = arrayOfInputs(7)
    
    iOutput = 0
    
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Integer
    
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    If lPartDataBasis <= 1 Or lPartDataBasis = 468 Then
    
        'Inputs
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
        parAngle = arrayOfInputs(7)
        
        Dim dAngle As Double
        If lPartDataBasis <= 1 Then
            dAngle = PI / 4
        ElseIf lPartDataBasis = 468 Then
            dAngle = parAngle
        End If
        
        'Insert your code for output 3(Insulation for Cylinder1)
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        stPoint.Set CenterPos.x - parFace1toCenter, _
                    CenterPos.y, _
                    CenterPos.z
        enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
        parInsulationDiameter = BodyOD(pipeDiam, flangeDiam) + 2 * parInsulationThickness
        
        Set ObjInsulatedCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                        parInsulationDiameter, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCylinder1
        Set ObjInsulatedCylinder1 = Nothing
        
        'Insert your code for output 5(Insulation for Cylinder2)
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        stPoint.Set CenterPos.x + parFace2toCenter * Cos(dAngle), _
                    CenterPos.y - parFace2toCenter * Sin(dAngle), _
                    CenterPos.z
        enPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
        
        parInsulationDiameter = BodyOD(pipeDiam, flangeDiam) + 2 * parInsulationThickness
               
        Set ObjInsulatedCylinder2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                        parInsulationDiameter, True)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCylinder2
        Set ObjInsulatedCylinder2 = Nothing
    
    ElseIf lPartDataBasis = 469 Then
        
        'Dim parCylinderDiameter As Double
        'Dim parCylinderLength   As Double
        'Dim parFacetoCenter     As Double
        
        'Inputs
        'parCylinderDiameter = arrayOfInputs(6)
        'parCylinderLength = arrayOfInputs(7)
        'parFacetoCenter = arrayOfInputs(8)
        parFace2toCenter = arrayOfInputs(3)
        parTangentLength1 = arrayOfInputs(6)
        parAngle = arrayOfInputs(7)
                       
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        Dim dDiameter As Double
        If CmpDblEqual(flangeDiam, 0) Then
            dDiameter = pipeDiam
        Else
            dDiameter = flangeDiam
        End If
        
        'Insert your code for Port1
        Dim oInsPort    As Object
        
        stPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
        enPoint.Set CenterPos.x - parTangentLength1, CenterPos.y, CenterPos.z
                
        parInsulationDiameter = BodyOD(pipeDiam, flangeDiam) + 2 * parInsulationThickness
        
        Set oInsPort = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
        
        'Set the Output(Insualtion Port1)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oInsPort
        Set oInsPort = Nothing
        
        'Insert your code for Port2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        
        Dim dValue4   As Double
        dValue4 = 0.5 * dDiameter * (Sin(parAngle) * _
                                    Tan(parAngle) - 1 + Cos(parAngle))
              
        parInsulationDiameter = BodyOD(pipeDiam, flangeDiam) + 2 * parInsulationThickness
        
        stPoint.Set CenterPos.x + (dValue4 * Cos(parAngle) + parFace2toCenter) * Sin(parAngle), _
                                    CenterPos.y - (dValue4 * Cos(parAngle) + parFace2toCenter) * Cos(parAngle), _
                                    CenterPos.z
                                    
        If CmpDblEqual(flangeThick, 0) Then
            enPoint.Set CenterPos.x + (dValue4 * Cos(parAngle) + parFace2toCenter - NEGLIGIBLE_THICKNESS) * Sin(parAngle), _
                                    CenterPos.y - (dValue4 * Cos(parAngle) + parFace2toCenter - NEGLIGIBLE_THICKNESS) * Cos(parAngle), _
                                    CenterPos.z
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
                parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
            End If
        Else
            enPoint.Set CenterPos.x + (dValue4 * Cos(parAngle) + parFace2toCenter - flangeThick - parInsulationThickness) * Sin(parAngle), _
                                    CenterPos.y - (dValue4 * Cos(parAngle) + parFace2toCenter - flangeThick - parInsulationThickness) * Cos(parAngle), _
                                    CenterPos.z
        End If
        
        Set oInsPort = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
        
        'Set the Output(Insualtion Port2)
        m_OutputColl.AddOutput "InsulationPort2", oInsPort
        Set oInsPort = Nothing
        
        'Insert your code for Port2 Cylinder
        Dim oCylinder2  As Object
        stPoint.Set CenterPos.x + (dValue4 * Cos(parAngle) + parFace2toCenter) * Sin(parAngle), _
                                    CenterPos.y - (dValue4 * Cos(parAngle) + parFace2toCenter) * Cos(parAngle), _
                                    CenterPos.z
        enPoint.Set CenterPos.x + 2 * dValue4 * Sin(parAngle) * Cos(parAngle), _
                                    CenterPos.y - 2 * dValue4 * Cos(parAngle) * Cos(parAngle), _
                                    CenterPos.z
                                            
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        Set oCylinder2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
        
        'Set the Output(Insualtion Port2)
        m_OutputColl.AddOutput "Cylinder2", oCylinder2
        Set oCylinder2 = Nothing
        
        'Insert your code for body
        Dim oCircle1        As IngrGeom3D.Circle3d
        Dim oCircle2        As IngrGeom3D.Circle3d
        Dim oDir            As AutoMath.DVector
        Dim oGeomFactory    As IngrGeom3D.GeometryFactory
        Dim oInsBody        As Object
        Dim dValue1         As Double
        
        Dim dHeight1        As Double
       
        dHeight1 = dDiameter / 2 * Sin(parAngle)
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        dValue1 = 0.5 * dDiameter * (Sin(parAngle) * _
                    Tan(parAngle) - 1 + Cos(parAngle))
        
        Set oDir = New AutoMath.DVector
        oDir.Set 1, 0, 0
       
        Set oCircle1 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                            CenterPos.x, CenterPos.y, CenterPos.z, _
                            oDir.x, oDir.y, oDir.z, dDiameter / 2 + _
                            parInsulationThickness)
        
        oDir.Set Cos(parAngle), Sin(parAngle), 0
                 
        Set oCircle2 = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                             CenterPos.x + dHeight1 + parInsulationThickness * Cos(parAngle), _
                             CenterPos.y + dDiameter / 2 * (1 - Cos(parAngle)) + parInsulationThickness * Sin(parAngle), _
                             CenterPos.z, oDir.x, oDir.y, oDir.z, 3 / 8 * dDiameter + parInsulationThickness)
        
        Set oInsBody = oGeomFactory.RuledSurfaces3d.CreateByCurves( _
                            m_OutputColl.ResourceManager, oCircle1, oCircle2, True)
        
        'Set the Output(Insulation Body)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oInsBody
        
        Set oCircle1 = Nothing
        Set oCircle2 = Nothing
        Set oDir = Nothing
        Set oInsBody = Nothing
        Set oGeomFactory = Nothing
        Set stPoint = Nothing
        Set enPoint = Nothing
    End If
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
