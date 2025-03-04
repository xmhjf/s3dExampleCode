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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Sunday, Jul 18 2004
'   Description:
'     This class module is the place for user to implement graphical part of VBSymbol for this aspect
'     Eccentric Reducing Tee symbol taken from BPE BIO-TECH catalog. It consists of eight outputs
'     Three Nozzles (Physical) and others are insulation aspect outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        -----------------
'   22.Aug.2005     svsmylav        TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
'   08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   30.Jul.2008     KKC      CR- 146404 Enhance Eccentric Reducing Tee symbol for seat-to-seat dimension per JIS G 5527
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim pipeDiam1        As Double
    Dim flangeThick     As Double
    Dim cptoffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parSeat1toCenter As Double
    Dim parSeat2toCenter As Double
    Dim parInsulationThickness As Double
    
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double
     
    Set oPartFclt = arrayOfInputs(1)
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

' Inputs
    If lPartDataBasis <= 1 Then
        parFace1toCenter = arrayOfInputs(2)
        parFace2toCenter = arrayOfInputs(3)
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
        parInsulationThickness = arrayOfInputs(4)
        
    ElseIf lPartDataBasis = 25 Then
         parSeat1toCenter = arrayOfInputs(5)
         parSeat2toCenter = arrayOfInputs(6)
         parInsulationThickness = arrayOfInputs(4)
         RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, _
                                    flangeDiam, cptoffset, depth
         dFace1toCenter = parSeat1toCenter + depth - cptoffset
         RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, _
                                    flangeDiam, cptoffset, depth
         dFace2toCenter = parSeat2toCenter + depth - cptoffset
    End If
    iOutput = 0
     
' Insert your code for output 1(Insulated Body)
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim dInsulationDiameter As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, cptoffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set -(dFace1toCenter - flangeThick), 0, 0
    enPoint.Set dFace1toCenter - flangeThick, 0, 0
    dInsulationDiameter = pipeDiam1 + parInsulationThickness * 2
    Dim ObjInsualatedCylinder As Object
    Set ObjInsualatedCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsualatedCylinder
    Set ObjInsualatedCylinder = Nothing
       
' Insert your code for output 2(Insulated Port1)
    stPoint.Set -dFace1toCenter, 0, 0
    enPoint.Set -(dFace1toCenter - flangeThick - parInsulationThickness), 0, 0
    dInsulationDiameter = pipeDiam1 + parInsulationThickness * 2
    If flangeDiam > pipeDiam1 Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2

    Dim ObjInsualatedCylinder1 As Object
    Set ObjInsualatedCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsualatedCylinder1
    Set ObjInsualatedCylinder1 = Nothing
      
' Insert your code for output 3(Insulated Port2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set dFace1toCenter, 0, 0
    enPoint.Set (dFace1toCenter - flangeThick - parInsulationThickness), 0, 0
    dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2

    Dim ObjInsualatedCylinder2 As Object
    Set ObjInsualatedCylinder2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsualatedCylinder2
    Set ObjInsualatedCylinder2 = Nothing
      
  ' Insert your code for output 4(Insulated Branch)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    stPoint.Set 0, 0, (pipeDiam1 - pipeDiam) / 2
    enPoint.Set 0, dFace2toCenter, (pipeDiam1 - pipeDiam) / 2
    dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    
    Dim ObjInsualatedCylinder3 As Object
    Set ObjInsualatedCylinder3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsualatedCylinder3
    Set ObjInsualatedCylinder3 = Nothing
     
' Insert your code for output 5 (Insulated Port3)
    stPoint.Set 0, dFace2toCenter, (pipeDiam1 - pipeDiam) / 2
    enPoint.Set 0, (dFace2toCenter - flangeThick - parInsulationThickness), (pipeDiam1 - pipeDiam) / 2
    dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If flangeDiam > pipeDiam Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2
    
    Dim ObjInsualatedCylinder4 As Object
    Set ObjInsualatedCylinder4 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsualatedCylinder4
    Set ObjInsualatedCylinder4 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
     
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub
