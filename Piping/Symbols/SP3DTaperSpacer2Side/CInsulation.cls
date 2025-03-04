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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          KKC
'   Creation Date:  Monday, Sep 3 2007
'   Description:
'   For the Partdatabasis option 381 and 382, the source is from
'    http://www.pep-plastic.com/manufacturers/edlon/pdf/cat-plpfa.pdf
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----        ------------------
'  03.SEP.2007        KKC         CR-121718 Added new Insulation Aspect
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
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
    
    Dim iOutput     As Double
    Dim ObjInsulationBody As Object
    
    Dim parSpacerThickness As Double
    Dim parAngle As Double
    Dim parBodyOutsideDiamter As Double
    Dim parSpacerThickness1 As Double
    Dim parInsulationThickness As Double
    Dim lPartDataBasis As Long
    Dim dInsulationDiameter As Double
    Dim TanHalfB        As Double
    Dim SinHalfB        As Double
    Dim CosHalfB        As Double
    Dim Face1toCenter   As Double
    Dim Face2Offset1    As Double
    Dim Face2Offset2    As Double
    
'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    If lPartDataBasis <= 1 Or lPartDataBasis = 380 Then
        parSpacerThickness = arrayOfInputs(2)
        parAngle = arrayOfInputs(3)
        If UBound(arrayOfInputs) >= 4 Then
             parInsulationThickness = arrayOfInputs(6)
        End If
        
    ElseIf lPartDataBasis = 381 Then
        parSpacerThickness = arrayOfInputs(2)
        parSpacerThickness1 = arrayOfInputs(4)
        parBodyOutsideDiamter = arrayOfInputs(5)
        parInsulationThickness = arrayOfInputs(6)
    ElseIf lPartDataBasis = 382 Then
        parSpacerThickness = arrayOfInputs(2)
        parAngle = arrayOfInputs(3)
        parBodyOutsideDiamter = arrayOfInputs(5)
        parInsulationThickness = arrayOfInputs(6)
    End If
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    iOutput = 0
    If CmpDblEqual(parAngle, 0) Then
       parAngle = Atn((parSpacerThickness1 - parSpacerThickness) / parBodyOutsideDiamter)
    End If
         
    If CmpDblEqual(parBodyOutsideDiamter, 0) Then parBodyOutsideDiamter = pipeDiam
    
    TanHalfB = Tan(0.5 * parAngle)
    SinHalfB = Sin(0.5 * parAngle)
    CosHalfB = Cos(0.5 * parAngle)
    
    Dim dLimitMaxthick As Double
    dLimitMaxthick = parSpacerThickness / (2 * TanHalfB)
    If CmpDblGreaterthan(parInsulationThickness, dLimitMaxthick) Then
        parInsulationThickness = dLimitMaxthick
    End If
    iOutput = 0
'   Note: Coordinate system is considered in such a way that the X-axis is collinear with
'   normal of the port1 for Partdatabasis 1.

' Insert your code for output 1(Body of Spacer)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If lPartDataBasis <= 1 Or lPartDataBasis = 380 Then
        Face1toCenter = (parSpacerThickness / 2 + TanHalfB * pipeDiam / 2) / CosHalfB
    ElseIf lPartDataBasis = 381 Or lPartDataBasis = 382 Then
        Face1toCenter = parSpacerThickness / 2 + TanHalfB * parBodyOutsideDiamter / 2
    End If
    
'   Place Left Hand Side ellipse

    Dim ellNormalX As Double, ellNormalY As Double, ellNormalZ As Double
    Dim ell1CenterX As Double, ell1CenterY As Double, ell1CenterZ As Double
    Dim MajorX As Double, MajorY As Double, MajorZ As Double
    If lPartDataBasis <= 1 Or lPartDataBasis = 380 Then
        ellNormalX = -1
        ellNormalY = 0
        ellNormalZ = 0
        ell1CenterX = CenterPos.x - Face1toCenter + flangeThick
        ell1CenterY = CenterPos.y
        ell1CenterZ = CenterPos.z
        MajorX = 0
        MajorY = (pipeDiam + 2 * parInsulationThickness) / (2 * CosHalfB)
        MajorZ = 0
    ElseIf lPartDataBasis = 381 Or lPartDataBasis = 382 Then
        ellNormalX = -CosHalfB
        ellNormalY = SinHalfB
        ellNormalZ = 0
        ell1CenterX = CenterPos.x - Face1toCenter + flangeThick / Cos(parAngle)
        MajorX = TanHalfB * (parBodyOutsideDiamter + 2 * parInsulationThickness) / 2
        MajorY = parBodyOutsideDiamter / 2 + parInsulationThickness
        MajorZ = 0
    End If
    
    Dim MMRatio As Double
    MMRatio = CosHalfB
    
    Dim ObjEllipse As Object
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Set ObjEllipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                        ell1CenterX, ell1CenterY, ell1CenterZ, _
                                        ellNormalX, ellNormalY, ellNormalZ, _
                                        MajorX, MajorY, MajorZ, MMRatio)

'   Place Right Hand Side ellipse
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If lPartDataBasis <= 1 Or lPartDataBasis = 380 Then
        Face2Offset1 = (parSpacerThickness + TanHalfB * pipeDiam) * CosHalfB - Face1toCenter
        Face2Offset2 = (parSpacerThickness + TanHalfB * pipeDiam) * SinHalfB
    
        ellNormalX = Cos(parAngle)
        ellNormalY = Sin(parAngle)
        ellNormalZ = 0
        ell1CenterX = CenterPos.x + Face2Offset1 - flangeThick
        ell1CenterY = CenterPos.y + Face2Offset2 - flangeThick * TanHalfB
        ell1CenterZ = CenterPos.z
        MajorX = -(0.5 * (pipeDiam + 2 * parInsulationThickness) / CosHalfB) * Sin(parAngle)
        MajorY = (0.5 * (pipeDiam + 2 * parInsulationThickness) / CosHalfB) * Cos(parAngle)
        MajorZ = 0
        
    ElseIf lPartDataBasis = 381 Or lPartDataBasis = 382 Then
        ellNormalX = CosHalfB
        ellNormalY = SinHalfB
        ellNormalZ = 0
        ell1CenterX = CenterPos.x + parSpacerThickness / 2 + TanHalfB * parBodyOutsideDiamter / 2 - flangeThick / Cos(parAngle)
        ell1CenterY = CenterPos.y
        ell1CenterZ = CenterPos.z
       
        MajorX = TanHalfB * (parBodyOutsideDiamter + 2 * parInsulationThickness) / 2
        MajorY = parBodyOutsideDiamter / 2 + parInsulationThickness
        MajorZ = 0
    End If
    
    Dim ObjEllipse2 As Object
    Set ObjEllipse2 = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                        ell1CenterX, ell1CenterY, ell1CenterZ, _
                                        ellNormalX, ellNormalY, ellNormalZ, _
                                        MajorX, MajorY, MajorZ, MMRatio)
'   Created Ruled surface
    Set ObjInsulationBody = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                            ObjEllipse, ObjEllipse2, True)
                                          
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulationBody
    Set ObjInsulationBody = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
