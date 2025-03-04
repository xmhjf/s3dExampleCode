VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CSimplePhysical.cls
'   Author:         MS
'   Creation Date:  Monday June 14, 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

    Private m_oSymGeomHelper As IJSymbolGeometryHelper
    Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oSymGeomHelper = New SymbolServices
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
    Private Sub Class_Terminate()
        Set m_oSymGeomHelper = Nothing
    End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
   
    Dim iOutput     As Double

    Dim parPostBaseToValveCenter As Double
    Dim parWallThickness As Double
    Dim parPostHeight As Double
    Dim parPostDiameter As Double
    Dim parHandWheelDiameter As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parPostBaseToValveCenter = arrayOfInputs(2)
    parWallThickness = arrayOfInputs(3)
    parPostHeight = arrayOfInputs(4)
    parPostDiameter = arrayOfInputs(5)
    parHandWheelDiameter = arrayOfInputs(6)
  
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    iOutput = 0
    Dim StemDia As Double
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

' Insert your code for output 1( Stem)
    'Assumed stem diameter to be 0.015m
    'Assuming the bottom portion of post below the inside the wall to be equal to the wall thickness
    StemDia = 0.015
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parPostBaseToValveCenter - parWallThickness, 0
    ' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, StemDia


' Insert your code for output 2(Bottom portion of post below the inside the wall )
    stPoint.Set 0, parPostBaseToValveCenter - parWallThickness, 0
    enPoint.Set 0, parPostBaseToValveCenter + parWallThickness, 0
    ' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parPostDiameter

' Insert your code for output 3(Base Plate Flange)
    Dim PostBasePlateThickness As Double
    Dim PostBasePlateDiameter As Double
    ' Assuming Post Base Plate Thickness to be 1in and Post Base Plate Diameter to be 12in
    PostBasePlateThickness = 0.0254
    PostBasePlateDiameter = 0.3048
    
    stPoint.Set 0, parPostBaseToValveCenter + parWallThickness, 0
    enPoint.Set 0, parPostBaseToValveCenter + parWallThickness + PostBasePlateThickness, 0
    ' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, PostBasePlateDiameter

' Insert your code for output 4(Post Body)
    stPoint.Set 0, parPostBaseToValveCenter + parWallThickness + PostBasePlateThickness, 0
    enPoint.Set 0, parPostBaseToValveCenter + parWallThickness + parPostHeight, 0
    ' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parPostDiameter

' Insert your code for output 5(Hand Wheel of Post )
    Dim TorOrigin       As New AutoMath.DPosition
    Dim TorNorAxis         As New AutoMath.DVector
    Dim TorMajAxis         As New AutoMath.DVector
    Dim TorPriRad       As Double
    Dim TorSecRad       As Double
    
    TorOrigin.Set 0, parPostBaseToValveCenter + parWallThickness + parPostHeight, 0
'   Prepare Torus Radii, the Torus radius is assumed to be 0.5 of StemDiam
    TorSecRad = StemDia / 2
    TorPriRad = parHandWheelDiameter / 2 - TorSecRad
    
'   Prepare axis vector for Torus
    TorNorAxis.Set 0, 1, 0
    TorMajAxis.Set 1, 0, 0
    
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateTorus arrayOfOutputs(iOutput), TorOrigin, TorNorAxis, TorMajAxis, TorPriRad, 360, 2 * TorSecRad
    
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

