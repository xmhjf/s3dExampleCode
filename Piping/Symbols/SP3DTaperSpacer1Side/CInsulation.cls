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
    Dim ObjInsulationCylinder As Object
    
    Dim parSpacerThickness As Double
    Dim parAngle As Double
    Dim parBodyOutsideDiamter As Double
    Dim parSpacerThickness1 As Double
    Dim parInsulationThickness As Double
    Dim lPartDataBasis As Long
    Dim dInsulationDiameter As Double
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
    If CmpDblEqual(parAngle, 0) Then
       parAngle = Atn((parSpacerThickness1 - parSpacerThickness) / parBodyOutsideDiamter)
    End If
    
    Dim dlLimitMaxThk As Double
    dlLimitMaxThk = parSpacerThickness / Tan(parAngle)
    
    If CmpDblGreaterthan(parInsulationThickness, dlLimitMaxThk) Then
        parInsulationThickness = dlLimitMaxThk
    End If
    
    If CmpDblEqual(parBodyOutsideDiamter, 0) Then parBodyOutsideDiamter = pipeDiam
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    iOutput = 0
    If CmpDblEqual(parAngle, 0) Then
       parAngle = Atn((parSpacerThickness1 - parSpacerThickness) / parBodyOutsideDiamter)
    End If
       
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

'   Place circle
    Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
    cirNormalX = 1
    cirNormalY = 0
    cirNormalZ = 0
    Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
    If lPartDataBasis <= 1 Or lPartDataBasis = 380 Then
        cir1CenterX = CenterPos.x - parSpacerThickness / 2 - Tan(parAngle) * pipeDiam / 4 + flangeThick
        cir1CenterY = CenterPos.y
        cir1CenterZ = CenterPos.z
        
        Dim objCircle1 As IngrGeom3D.Circle3d
        Dim geomFactory As New IngrGeom3D.GeometryFactory
        Set objCircle1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                    Nothing, cir1CenterX, _
                                                    cir1CenterY, cir1CenterZ, cirNormalX, _
                                                    cirNormalY, cirNormalZ, (pipeDiam / 2 + parInsulationThickness))
    ElseIf lPartDataBasis = 381 Or lPartDataBasis = 382 Then
        cir1CenterX = CenterPos.x - parSpacerThickness - (parBodyOutsideDiamter / 2) * Tan(parAngle) + flangeThick
        cir1CenterY = CenterPos.y
        cir1CenterZ = CenterPos.z
        Set objCircle1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                    Nothing, cir1CenterX, _
                                                    cir1CenterY, cir1CenterZ, cirNormalX, _
                                                    cirNormalY, cirNormalZ, (parBodyOutsideDiamter / 2 + parInsulationThickness))
    End If
    
'   Place ellipse
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    Dim ellNormalX As Double, ellNormalY As Double, ellNormalZ As Double
    If lPartDataBasis <= 1 Or lPartDataBasis = 380 Then
        ellNormalX = Cos(parAngle)
        ellNormalY = Sin(parAngle)
        ellNormalZ = 0
        Dim ell1CenterX As Double, ell1CenterY As Double, ell1CenterZ As Double
        ell1CenterX = CenterPos.x + parSpacerThickness / 2 + Tan(parAngle) * pipeDiam / 4 _
                                  - flangeThick / Cos(parAngle)
        ell1CenterY = CenterPos.y
        ell1CenterZ = CenterPos.z
        Dim MajorX As Double, MajorY As Double, MajorZ As Double
        MajorX = -Tan(parAngle) * (pipeDiam + 2 * parInsulationThickness) / 2
        MajorY = pipeDiam / 2
        MajorZ = 0
        
    ElseIf lPartDataBasis = 381 Or lPartDataBasis = 382 Then
        ellNormalX = Cos(parAngle)
        ellNormalY = Sin(parAngle)
        ellNormalZ = 0
        ell1CenterX = CenterPos.x - flangeThick / Cos(parAngle)
        ell1CenterY = CenterPos.y
        ell1CenterZ = CenterPos.z
       
        MajorX = -Tan(parAngle) * (parBodyOutsideDiamter + 2 * parInsulationThickness) / 2
        MajorY = parBodyOutsideDiamter / 2 + parInsulationThickness
        MajorZ = 0
        
    End If
    Dim ObjEllipse As Object
    Set ObjEllipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                        ell1CenterX, ell1CenterY, ell1CenterZ, _
                                        ellNormalX, ellNormalY, ellNormalZ, _
                                        MajorX, MajorY, MajorZ, Cos(parAngle))
    
'   Created Ruled surface
    Dim ObjInsBodySpacer As Object
    Set ObjInsBodySpacer = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                            objCircle1, ObjEllipse, True)
    Set geomFactory = Nothing

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBodySpacer
    Set ObjInsBodySpacer = Nothing

    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
