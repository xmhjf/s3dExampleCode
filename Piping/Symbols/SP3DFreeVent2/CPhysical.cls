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
'   Author:          svsmylav
'   Creation Date:   Wednesday 1, Feb 2006
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per Model 5000 in the Growth corporation's Catalog GrowthC.pdf
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'
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
    
    Dim iOutput     As Double
    Dim ObjFreeeVentBody As Object
    
'**
    Dim parPipingSpecialtyWidth As Double
    Dim parPipingSpecialtyHeight As Double
    Dim parVentAngle As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parPipingSpecialtyWidth = arrayOfInputs(2)  'A
    parPipingSpecialtyHeight = arrayOfInputs(3) 'B
    parVentAngle = arrayOfInputs(4)

    iOutput = 0

'   Assumption: Origin of the symbol is assumed to be at the center of the flange face

 ' Insert your code for output 1(FreeeVentBody)
    Dim dBaseDiam  As Double
    Dim dBendRadius  As Double
    Dim dNozztoCenter As Double

    Dim cirNormalX As Double
    Dim cirNormalY As Double
    Dim cirNormalZ As Double
    Dim cir1CenterX As Double
    Dim cir1CenterY As Double
    Dim cir1CenterZ As Double
    
    Dim objCirc1 As IngrGeom3D.Circle3d
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim centPoint   As New AutoMath.DPosition
    Dim axis        As New AutoMath.DVector

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    dBaseDiam = flangeDiam
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dBaseDiam = pipeDiam 'Check to ensure atleast value equal to pipeDiam
    dBendRadius = (parPipingSpecialtyWidth - dBaseDiam / 2 - pipeDiam / 2) / 2
    dNozztoCenter = parPipingSpecialtyHeight - (dBendRadius + pipeDiam / 2)
    If CmpDblLessThan(dNozztoCenter, flangeThick) Then dNozztoCenter = flangeThick

    cirNormalX = -1
    cirNormalY = 0
    cirNormalZ = 0

    cir1CenterX = dNozztoCenter
    cir1CenterY = 0
    cir1CenterZ = 0

    Set objCirc1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                Nothing, cir1CenterX, _
                                                cir1CenterY, cir1CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, pipeDiam / 2)
    centPoint.Set dNozztoCenter, dBendRadius, 0
    axis.Set 0, 0, 1
    Set ObjFreeeVentBody = PlaceRevolution(m_OutputColl, objCirc1, _
                                    axis, centPoint, parVentAngle, _
                                    True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFreeeVentBody
    Set ObjFreeeVentBody = Nothing
    Set geomFactory = Nothing
    Set centPoint = Nothing
    Set axis = Nothing
    Set objCirc1 = Nothing
 
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
 
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozztoCenter)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
   
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
