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
'   Author:          svsmylav
'   Creation Date:  Monday, Feb 4 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No. is: F14A, Page No. D-41 of PDS Piping Component Data Reference Guide.
' SP3DTaperSpacer1Side Symbol is created with three Outputs as follows:
'   1) ObjBodySpacer: RuledSurfaces3d using 'CreateByCurves' method,
'   2 & 3 are nozzles prepared using 'CreateNozzle' function.
' For the Partdatabasis option 381 and 382, the source is from
'    http://www.pep-plastic.com/manufacturers/edlon/pdf/cat-plpfa.pdf
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  03.SEP.2007        KKC         CR-121718 Provided the Taper Spacer symbol as per specification
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

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

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjBodySpacer As Object
    
    Dim parSpacerThickness As Double
    Dim parAngle As Double
    Dim parBodyOutsideDiamter As Double
    Dim parSpacerThickness1 As Double
    Dim lPartDataBasis As Long
'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    If lPartDataBasis <= 1 Or lPartDataBasis = 380 Then
        parSpacerThickness = arrayOfInputs(2)
        parAngle = arrayOfInputs(3)
    ElseIf lPartDataBasis = 381 Then
        parSpacerThickness = arrayOfInputs(2)
        parSpacerThickness1 = arrayOfInputs(4)
        parBodyOutsideDiamter = arrayOfInputs(5)
    ElseIf lPartDataBasis = 382 Then
        parSpacerThickness = arrayOfInputs(2)
        parAngle = arrayOfInputs(3)
        parBodyOutsideDiamter = arrayOfInputs(5)
    End If
    iOutput = 0
    If CmpDblEqual(parAngle, 0) Then
       parAngle = Atn((parSpacerThickness1 - parSpacerThickness) / parBodyOutsideDiamter)
    End If
       
' Insert your code for output 1(Body of Spacer)
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
                                                    cirNormalY, cirNormalZ, pipeDiam / 2)
    ElseIf lPartDataBasis = 381 Or lPartDataBasis = 382 Then
        cir1CenterX = CenterPos.x - parSpacerThickness - (parBodyOutsideDiamter / 2) * Tan(parAngle) + flangeThick
        cir1CenterY = CenterPos.y
        cir1CenterZ = CenterPos.z
        Set objCircle1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                    Nothing, cir1CenterX, _
                                                    cir1CenterY, cir1CenterZ, cirNormalX, _
                                                    cirNormalY, cirNormalZ, parBodyOutsideDiamter / 2)
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
        MajorX = -Tan(parAngle) * pipeDiam / 2
        MajorY = pipeDiam / 2
        MajorZ = 0
        
    ElseIf lPartDataBasis = 381 Or lPartDataBasis = 382 Then
        ellNormalX = Cos(parAngle)
        ellNormalY = Sin(parAngle)
        ellNormalZ = 0
        ell1CenterX = CenterPos.x - flangeThick / Cos(parAngle)
        ell1CenterY = CenterPos.y
        ell1CenterZ = CenterPos.z
       
        MajorX = -Tan(parAngle) * parBodyOutsideDiamter / 2
        MajorY = parBodyOutsideDiamter / 2
        MajorZ = 0
        
    End If
    Dim ObjEllipse As Object
    Set ObjEllipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                        ell1CenterX, ell1CenterY, ell1CenterZ, _
                                        ellNormalX, ellNormalY, ellNormalZ, _
                                        MajorX, MajorY, MajorZ, Cos(parAngle))
    
'   Created Ruled surface
    Set ObjBodySpacer = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                            objCircle1, ObjEllipse, True)
    Set geomFactory = Nothing

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodySpacer
    Set ObjBodySpacer = Nothing

    
' Remove curves
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCircle1
    Set objCircle1 = Nothing
    sp3dElem1.Remove
    
'   Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = ObjEllipse
    Set ObjEllipse = Nothing
    sp3dElem1.Remove
    
' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    If lPartDataBasis <= 1 Or lPartDataBasis = 380 Then
        oPlacePoint.Set CenterPos.x - parSpacerThickness / 2 - Tan(parAngle) * pipeDiam / 4 - sptOffset + depth, _
                        CenterPos.y, _
                        CenterPos.z
    ElseIf lPartDataBasis = 381 Or lPartDataBasis = 382 Then
'          oPlacePoint.Set CenterPos.x - parSpacerThickness - sptOffset + depth, _
'                    CenterPos.y, _
'                    CenterPos.z
           oPlacePoint.Set CenterPos.x - parSpacerThickness - (parBodyOutsideDiamter / 2) * Tan(parAngle) - sptOffset + depth, _
                    CenterPos.y, _
                    CenterPos.z
    End If
    oDir.Set -1, 0, 0

    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If lPartDataBasis <= 1 Or lPartDataBasis = 380 Then
        oPlacePoint.Set CenterPos.x + parSpacerThickness / 2 + Tan(parAngle) * pipeDiam / 4 + (sptOffset - depth) * Cos(parAngle), _
                        CenterPos.y + (sptOffset - depth) * Sin(parAngle), _
                        CenterPos.z
    ElseIf lPartDataBasis = 381 Or lPartDataBasis = 382 Then
'        oPlacePoint.Set CenterPos.x + (parBodyOutsideDiamter / 2) * Tan(parAngle) + (sptOffset - depth) * Cos(parAngle), _
'                CenterPos.y + (sptOffset - depth) * Sin(parAngle), _
'                CenterPos.z
          oPlacePoint.Set CenterPos.x + (sptOffset - depth) * Cos(parAngle), _
                CenterPos.y + (sptOffset - depth) * Sin(parAngle), _
                CenterPos.z
    End If
    oDir.Set Cos(parAngle), Sin(parAngle), 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set CenterPos = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
   
End Sub
