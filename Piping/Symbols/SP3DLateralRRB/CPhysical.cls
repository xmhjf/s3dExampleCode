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
'   Creation Date:  Friday, Dec 28 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy   who               change description
'   -----------   -----             ------------------
'   09.Jul.2003   SymbolTeam(India) Copyright Information, Header  is added.
'   08.SEP.2006   KKC               DI-95670  Replace names with initials in all revision history sheets and symbols
'   20.DEC.2007   KKC               CR-130970 Enhance Lateral symbols for underground routing to be trimmable for sloped routing
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI                  As Double
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
    
    Dim iOutput     As Double
    Dim ObjInsulatedBody1 As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedBody2 As Object
    Dim ObjInsulatedPort2 As Object
    Dim ObjInsulatedBranch As Object
    Dim ObjInsulatedPort3 As Object
    Dim objTruncCone      As Object
    
'   Variables defined by svsmylav
    Dim ConeCenterBasePt    As New AutoMath.DPosition
    Dim ConeCenterTopPt     As New AutoMath.DPosition
    Dim ConeBaseDia         As Double
    Dim ConeTopDia          As Double
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double

    Dim parAngle As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    If UBound(arrayOfInputs) >= 6 Then
        parAngle = arrayOfInputs(6)
    Else
        parAngle = PI / 4 'To suppport the prior to V7 instances
    End If
    
    
    iOutput = 0
    
    ConeCenterBasePt.Set CenterPos.x, CenterPos.y, CenterPos.z
    ConeCenterTopPt.Set CenterPos.x - parFace2toCenter, CenterPos.y, CenterPos.z

' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
'   Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
'   faceToFace = arrayOfInputs(2)

    oPlacePoint.Set CenterPos.x - parFace1toCenter - sptOffset + depth, _
                    CenterPos.y, CenterPos.z
    oDir.Set -1, 0, 0
    
'   Set oPartFclt = arrayOfInputs(1)

    Set objNozzle = CreateNozzleWithLength( _
                        1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, parFace1toCenter)
    ConeBaseDia = pipeDiam
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set CenterPos.x + parFace2toCenter + sptOffset - depth, _
                    CenterPos.y, CenterPos.z
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set CenterPos.x + (parFace3toCenter + sptOffset - depth) * Cos(parAngle), _
                    CenterPos.y + (parFace3toCenter + sptOffset - depth) * Sin(parAngle), _
                    CenterPos.z
    oDir.Set Cos(parAngle), Sin(parAngle), 0

    Set objNozzle = CreateNozzleWithLength( _
                        3, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, parFace3toCenter)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
        
' Place truncated cone
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    ConeCenterBasePt.Set CenterPos.x, CenterPos.y, CenterPos.z
    ConeCenterTopPt.Set CenterPos.x + parFace2toCenter - flangeThick, _
                        CenterPos.y, CenterPos.z
    ConeTopDia = pipeDiam
    Set objTruncCone = PlaceCone(m_OutputColl, ConeCenterBasePt, ConeCenterTopPt, ConeBaseDia / 2, ConeTopDia / 2)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTruncCone
    Set objNozzle = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' This is a sample of placing cylinder and 2 nozzles
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' Use nozzle 1 as pipe port for flange and pipe parameters
'''    Dim pipeDiam        As Double
'''    Dim flangeThick     As Double
'''    Dim sptOffset       As Double
'''    Dim flangeDiam      As Double
'''    Dim depth           As Double
'''    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
'''    Dim parOperatorWidth    As Double
'''    parOperatorWidth = arrayOfInputs(2)
'''    Dim stPoint   As New AutoMath.DPosition
'''    Dim enPoint   As New AutoMath.DPosition
'''    stPoint.Set -parOperatorWidth / 2 + flangeThick, 0, 0
'''    enPoint.Set parOperatorWidth / 2 - flangeThick, 0, 0
'''    Dim objCylinder As Object
'''
'''    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
'''    Set objCylinder = Nothing
'''
'''    Dim oPlacePoint As AutoMath.DPosition
'''    Dim oDir        As AutoMath.DVector
'''    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
'''
'''    Set oPlacePoint = New AutoMath.DPosition
'''    Set oDir = New AutoMath.DVector
'''    oPlacePoint.Set -parOperatorWidth / 2 - sptOffset + depth, 0, 0
'''    oDir.Set -1, 0, 0
'''    Set oPartFclt = arrayOfInputs(1)
'''    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
'''    oPlacePoint.Set parOperatorWidth / 2 + sptOffset - depth, 0, 0
'''    oDir.Set 1, 0, 0
'''
'''    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
