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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         ACM
'   Creation Date:  Monday, July 12, 2004
'   Description:
'      This is a Generic Pipet symbol. Technical details are taken from the site http://www.wfi-intl.com/bwp/bwp1.html
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim parA As Double
    Dim parInsulationThickness As Double
    
    Dim CentertoInsu As Double
    Dim oInsuPoint         As New AutoMath.DPosition
    Dim oFacePoint         As New AutoMath.DPosition

   
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parA = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    m_oGeomHelper.OutputCollection = m_OutputColl

    iOutput = 0

' Insert your code for output 2(Insulated Body)
    Dim dMajorBodyDia As Double
    
    'get the Header(Jacketed Pipe) data
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth


    CentertoInsu = (0.5 * pipeDiam) * Cos(40# / RAD)
    oInsuPoint.Set 0, CentertoInsu, 0
   
    oFacePoint.Set 0, pipeDiam / 2 + parA, 0

    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth
                        
''Waist Dia is is assumed as 1.5 times the Branch PipeOD as per the PDS Eden 'F-173'
    If pipeDiam < flangeDiam Then
       dMajorBodyDia = flangeDiam * 1.5
    Else
       dMajorBodyDia = pipeDiam * 1.5
    End If
    
    'Set the Output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oInsuPoint, oFacePoint, _
                                        dMajorBodyDia + parInsulationThickness * 2

    Set oInsuPoint = Nothing
    Set oFacePoint = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub



