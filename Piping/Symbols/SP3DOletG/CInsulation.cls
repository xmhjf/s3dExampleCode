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
'   Copyright (c) 2004-07, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CInsulation.cls
'   Author:         ACM
'   Creation Date:  Monday, Feb 9, 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       This is a Generic Olet symbol. It takes inputs from the Generic Data tables.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  28.SEP.2006     svsmylav TR-97093  Updated the symbol inline with the F173 (PDS EDEN).
'  18.APR.2007     KKC  CR-117893 Modified Olet symbol to be Trimmable
'  29.Jun.2007     svsmylav/dkl  Assigned parAngle = PI / 2 'For V6.1 symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private PI           As Double
Private RAD          As Double

Dim m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub

Private Sub Class_Initialize()
    Set m_oGeomHelper = New SymbolServices

    PI = 4 * Atn(1)
    RAD = 180 / PI
    
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
   
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1       As Double
    Dim pipeDiam2        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjInsulatedBody As Object
    Dim parA As Double
    Dim parInsulationThickness As Double
    Dim parAngle As Double
    
    Dim CentertoInsu As Double
    Dim oInsuPoint         As New AutoMath.DPosition
    Dim oFacePoint         As New AutoMath.DPosition
   
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parA = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    parAngle = PI / 2 'For V6.1 symbol
    If UBound(arrayOfInputs) >= 4 Then 'Not to fail the prior to V7 instances by accessing index 31
        parAngle = arrayOfInputs(4)
    End If
    
    m_oGeomHelper.OutputCollection = m_OutputColl

    iOutput = 0

' Insert your code for output 2(Insulated Body)
    Dim dMajorBodyDia As Double
    
    'get the Header data
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, _
                        flangeThick, flangeDiam, cptOffset, depth

    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, _
                        flangeThick, flangeDiam, cptOffset, depth
                        
''Waist Dia is is assumed as 1.5 times the Branch PipeOD as per the PDS Eden 'F-173'
    Dim dHoleinHeader As Double
    Dim dOffset As Double
    
'   Compute offset
    If CmpDblLessThan(pipeDiam2, flangeDiam) Then
       dHoleinHeader = flangeDiam
    Else
       dHoleinHeader = pipeDiam2
    End If

    If CmpDblGreaterthan(dHoleinHeader, pipeDiam1) Then
        dOffset = 0.125 * pipeDiam1
    Else
        dOffset = 0.5 * Sqr(pipeDiam1 ^ 2 - dHoleinHeader ^ 2)
    End If

'   Compute Major Body Diameter
    If CmpDblLessThan(pipeDiam2, flangeDiam) Then
       dMajorBodyDia = flangeDiam * 1.5
    Else
       dMajorBodyDia = pipeDiam2 * 1.5
    End If

    oInsuPoint.Set dOffset * Cos(parAngle), dOffset * Sin(parAngle), 0
    oFacePoint.Set (pipeDiam1 / 2 + parA) * Cos(parAngle), (pipeDiam1 / 2 + parA) * Sin(parAngle), 0
    
    Dim oDirVector As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotAboutZaxis As Double
    Set oDirVector = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    dRotAboutZaxis = PI / 2 - parAngle
    oDirVector.Set 0, 0, -1
    oTransformationMat.Rotate dRotAboutZaxis, oDirVector
    
    'Set the Output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oInsuPoint, oFacePoint, _
                                        dMajorBodyDia + parInsulationThickness * 2
    Set oTransformationMat = Nothing
    Set oDirVector = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
End Sub
