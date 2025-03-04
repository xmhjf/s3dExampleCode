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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  Saturday, Jul 21 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy        who                           change description
'   -----------       -----                             ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   20.Oct.2004     MS                    CR 60899 Updated symbol to compute offset based on OD or ID of the pipes.
'                                           An Input parameter EccentricOffsetBasis is added which is an codelist value.
'   01.Aug.2006     KKK                     Replaced 'm_OutputColl.ResourceManager' with "Nothing'.
'  08.SEP.2006      KKC                   DI-95670  Replace names with initials in all revision history sheets and symbols
'   22.Feb.2007     RRK                   TR-113129 Changed the parEccentricOffsetBasis type from Double to Long
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
    Dim pipeDiam1        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim pipeInsideDiam1 As Double
    Dim pipeInsideDiam2 As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim parEccentricOffsetBasis As Long

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    parEccentricOffsetBasis = arrayOfInputs(4)
    

    iOutput = 0

 ' Insert your code for output 1(Body)
  ' Create Circle1

    RetrieveParametersWithInsidePipeDiameter 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, sptOffset, depth, pipeInsideDiam1
    
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    RetrieveParametersWithInsidePipeDiameter 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2, pipeInsideDiam2
    
    Dim dOffsetBetweemCentreLines As Double
    
    If parEccentricOffsetBasis = 2 Then 'based on inside diameter
        dOffsetBetweemCentreLines = (pipeInsideDiam1 - pipeInsideDiam2) / 2
    Else 'considering parEccentricOffsetBasis = 1 case based on outside diameter
        dOffsetBetweemCentreLines = (pipeDiam1 - pipeDiam2) / 2
    End If
    
    
    Dim cirNormalX As Double
    Dim cirNormalY As Double
    Dim cirNormalZ As Double
    
    cirNormalX = 1
    cirNormalY = 0
    cirNormalZ = 0
    
    Dim objCircle1 As IngrGeom3D.Circle3d
    
    Dim cir1CenterX As Double
    Dim cir1CenterY As Double
    Dim cir1CenterZ As Double
    
    cir1CenterX = -parFacetoFace / 2 + flangeThick
    cir1CenterY = dOffsetBetweemCentreLines / 2
    cir1CenterZ = 0
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Set objCircle1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                Nothing, cir1CenterX, _
                                                cir1CenterY, cir1CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, pipeDiam1 * 0.5)
    

'Create Circle2
    Dim objCircle2 As IngrGeom3D.Circle3d
    Dim cir2CenterX As Double
    Dim cir2CenterY As Double
    Dim cir2CenterZ As Double
    
    cir2CenterX = parFacetoFace / 2 - flangeThick2
    cir2CenterY = -dOffsetBetweemCentreLines / 2
    cir2CenterZ = 0
        
    Set objCircle2 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                Nothing, cir2CenterX, _
                                                cir2CenterY, cir2CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, pipeDiam2 * 0.5)
    
    
    'Create Ruled Surface for body
    Dim ObjBody As IngrGeom3D.RuledSurface3d
    
    Set ObjBody = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                            objCircle1, objCircle2, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
' remove curve
    Dim sp3dElem1 As IJDObject
    Dim sp3dElem2 As IJDObject
    
    Set sp3dElem1 = objCircle1
    Set sp3dElem2 = objCircle2
    Set objCircle1 = Nothing
    Set objCircle2 = Nothing
    sp3dElem1.Remove
    sp3dElem2.Remove
    
    
' Place Nozzle 1
    
'    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    faceToFace = arrayOfInputs(2)
    oPlacePoint.Set -faceToFace / 2 - sptOffset + depth, dOffsetBetweemCentreLines / 2, 0
    oDir.Set -1, 0, 0
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
'    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set faceToFace / 2 + sptOffset2 - depth2, -dOffsetBetweemCentreLines / 2, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set CenterPos = Nothing
    Set geomFactory = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set sp3dElem1 = Nothing
    Set sp3dElem2 = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    
End Sub


