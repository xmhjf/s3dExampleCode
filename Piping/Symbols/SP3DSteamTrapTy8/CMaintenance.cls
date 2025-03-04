VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:         svsmylav
'   Creation Date:  Fridday, Feb 3 2006
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol constructed as per the Model No. BPC32Y (Type: Balanced Pressure Thermostatic Steam Trap)
'    in Spirax Sarco's catalog available at URL http://www.spiraxsarco.com
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  06.Mar.2008     VRK  TR-136850  Modified lower Trap Body
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Maintenance:" 'Used for error messages
Private PI       As Double
Dim m_oSymGeomHelper As IJSymbolGeometryHelper

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
     Set m_oSymGeomHelper = New SymbolServices
      PI = 4 * Atn(1)
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
    
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    Dim objMaintUpperTrapBody As Object

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    Dim parTrapLength1 As Double
    Dim parFacetoFace As Double
    Dim parBodyWidth As Double
    Dim parTrapHeightToptoCen As Double
    Dim parWithdrawalDistance As Double
    Dim parTrapHeightBottoCen As Double
    Dim parWithdrawalDistance1 As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parTrapLength1 = arrayOfInputs(2)    'A
    parFacetoFace = arrayOfInputs(3)     'A1
    parBodyWidth = arrayOfInputs(4)  'B
    parTrapHeightToptoCen = arrayOfInputs(5)     'C
    parWithdrawalDistance = arrayOfInputs(6)     'E
    parTrapHeightBottoCen = arrayOfInputs(7)     'F
    parWithdrawalDistance1 = arrayOfInputs(8)    'G
    parInsulationThickness = arrayOfInputs(9)
   
    iOutput = 0

    m_oSymGeomHelper.OutputCollection = m_OutputColl

' Insert your code for output 8(Maintenance of Upper Trap Body)
    Dim dUpperTrapBodyDia  As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

'   Assumptions:
'     1. Upper Trap Body (along with bolts) is considered for maintenance envelope, assumed to be box
'     2. Upper trap body diameter is taken to be 60 percent of A
'     3. The base for Upper Trap Body is taken as a box starting from half-of-pipeDiam

    dUpperTrapBodyDia = 0.6 * parTrapLength1
    stPoint.Set -dUpperTrapBodyDia / 2, -pipeDiam / 2, -parBodyWidth / 2
    enPoint.Set stPoint.x + dUpperTrapBodyDia, _
        stPoint.y + (parTrapHeightToptoCen + parWithdrawalDistance), _
        stPoint.z + parBodyWidth
    Set objMaintUpperTrapBody = PlaceBox(m_OutputColl, stPoint, enPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objMaintUpperTrapBody
    Set objMaintUpperTrapBody = Nothing
    
' Insert your code for output 9(Maintenance of Lower Trap Body
'   Assumptions:
'     1. Lower trap body diameter is taken to be 30 percent of A
'     2. Start point of the Lower Trap Body is taken to be at A /4 distance from the origin
    Dim dLowerTrapBodyDia  As Double
    dLowerTrapBodyDia = 0.3 * parTrapLength1
    stPoint.Set -parTrapLength1 / 4 + parTrapHeightBottoCen * Cos(PI / 4), _
                -parTrapHeightBottoCen * Sin(PI / 4), 0
    enPoint.Set stPoint.x + parWithdrawalDistance1 * Cos(PI / 4), _
                stPoint.y - parWithdrawalDistance1 * Sin(PI / 4), _
                stPoint.z

' Set the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dLowerTrapBodyDia
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
