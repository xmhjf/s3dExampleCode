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
'   Author:          MS
'   Creation Date:  Tuesday, Jun 4 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No. is: S11BZ Page No. D-102 PDS Piping Component Data Reference Guide.
' SP3DFreeVent Symbol is created with Two Outputs
' The two are physical aspect outputs created as follows:
'   ObjFreeeVentBody - Using 'PlaceRevolution' function and
'   One ObjNozzle objects  by using 'CreateNozzleWithLength' function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   01.Aug.2006     Kishore(kkkottur)       TR: 80254 Replaced 'm_OutputColl.ResourceManager' with "Nothing'.

'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private PI           As Double
Private Const MODULE = "Physical:" 'Used for error messages

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
    
    Dim iOutput     As Double
    Dim ObjFreeeVentBody As Object
    
    Dim parNozztoCenter As Double
    Dim parFacetoCenter As Double
    Dim parVentAngle As Double
    

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parNozztoCenter = arrayOfInputs(2) 'P2
    parFacetoCenter = arrayOfInputs(3) 'P3
    parVentAngle = arrayOfInputs(4) 'P4
    iOutput = 0


 ' Insert your code for output 1(FreeeVentBody)
     RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    Dim objCirc1 As IngrGeom3D.Circle3d
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    
    Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
    cirNormalX = 0
    cirNormalY = 0
    cirNormalZ = -1
    
    Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
    cir1CenterX = 0
    cir1CenterY = 0
    cir1CenterZ = parNozztoCenter - parFacetoCenter
    
    Set objCirc1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                Nothing, cir1CenterX, _
                                                cir1CenterY, cir1CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, pipeDiam / 2)
 
    Dim centPoint   As New AutoMath.DPosition
    Dim axis        As New AutoMath.DVector
 
    centPoint.Set parFacetoCenter, 0, parNozztoCenter - parFacetoCenter
    axis.Set 0, 1, 0
    
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
 
' Place Nozzle 1
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set 0, 0, -sptOffset + depth
    oDir.Set 0, 0, -1
 
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, parNozztoCenter - parFacetoCenter)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    ' remove curve
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCirc1
    Set objCirc1 = Nothing
    sp3dElem1.Remove
   
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
