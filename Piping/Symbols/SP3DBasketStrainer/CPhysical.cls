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
'   Author:         BG
'   Creation Date:  Wednesday 30, Jul 2003
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This Symbol details are available in PDS Piping Data Reference Guide at page no D-106 and SN=S53AC
'    and S53AZ
'    Consist of Physical and Insulation Aspects. Physical aspect is made up of Two Nozzles, One Collar of Cylinder,
'    one Collar extension of Box and Filter body of truncated cone. Insulation aspect consist of a Cylinder
'    for Collar Insulation.

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   26.OCT.2004     MS     TR 65841 Corrected the code to generate Collar Cylinder. Corrected the code to be as per PDS Eden
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

''''

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
    Dim ObjStrainer As Object
    Dim ObjCollar As Object
    Dim ObjCollarExtn As Object

    Dim parCollarThickness As Double
    Dim parBasketLength As Double
    Dim parBasketEndDiameter As Double
    Dim parCollarEndtoCenter As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parCollarThickness = arrayOfInputs(2)   'P2
    parBasketLength = arrayOfInputs(3)          'P3
    parBasketEndDiameter = arrayOfInputs(4)     'P4
    parCollarEndtoCenter = arrayOfInputs(5)         'P5
    parInsulationThickness = arrayOfInputs(6)

    iOutput = 0

 ' Insert your code for output 1(Strainer)
    Dim baseCenPt As AutoMath.DPosition
    Dim topCenPt As AutoMath.DPosition
    Set baseCenPt = New AutoMath.DPosition
    Set topCenPt = New AutoMath.DPosition

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    baseCenPt.Set -parCollarThickness / 2, 0, 0
    topCenPt.Set parCollarThickness / 2 + parBasketLength, 0, 0
    Set ObjStrainer = PlaceCone(m_OutputColl, baseCenPt, topCenPt, pipeDiam / 2, parBasketEndDiameter / 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStrainer
    Set ObjStrainer = Nothing
    Set baseCenPt = Nothing
    Set topCenPt = Nothing

 ' Insert your code for output 2(Collar)
    Dim startPt As AutoMath.DPosition
    Dim endPt As AutoMath.DPosition
    Set startPt = New AutoMath.DPosition
    Set endPt = New AutoMath.DPosition

    startPt.Set -parCollarThickness / 2, 0, 0
    endPt.Set parCollarThickness / 2, 0, 0
    
    'If the flangeDiam is zero, then set the collar diameter as pipediam
    Dim CollarDiameter As Double
    CollarDiameter = flangeDiam
    If flangeDiam = 0 Then CollarDiameter = pipeDiam
    
    Set ObjCollar = PlaceCylinder(m_OutputColl, startPt, endPt, CollarDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCollar
    Set ObjCollar = Nothing
    Set startPt = Nothing
    Set endPt = Nothing

 ' Insert your code for output 3(CollarExtn)
    Dim lowerPt As AutoMath.DPosition
    Dim upperPt As AutoMath.DPosition
    Set lowerPt = New AutoMath.DPosition
    Set upperPt = New AutoMath.DPosition
    
    Dim factor As Double
    ' The factors have been taken similar to that in the PDS Eden code. S53AC
    
    If (pipeDiam < 0.1) Then 'Assuming 3in pipediam not more than 0.1m and 10in pipediam not more than 0.3m
        factor = 1.25
        If flangeDiam = 0 Then factor = 0.75
    Else
        If (pipeDiam > 0.1 And pipeDiam < 0.3) Then
            factor = 0.55
        Else
            factor = 0.3
        End If
    End If
    
    
    Dim CollarExtnWidth As Double
    Dim CollarMeetHeight As Double
    CollarExtnWidth = pipeDiam * factor
    
    CollarMeetHeight = Sqr(CollarDiameter / 2 * CollarDiameter / 2 - CollarExtnWidth / 2 * CollarExtnWidth / 2)
    
    lowerPt.Set -parCollarThickness / 2, CollarMeetHeight, -CollarExtnWidth / 2
    upperPt.Set parCollarThickness / 2, parCollarEndtoCenter, CollarExtnWidth / 2
    
    Set ObjCollarExtn = PlaceBox(m_OutputColl, lowerPt, upperPt)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCollarExtn
    Set ObjCollarExtn = Nothing
    Set lowerPt = Nothing
    Set upperPt = Nothing

' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parCollarThickness / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parCollarThickness / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
