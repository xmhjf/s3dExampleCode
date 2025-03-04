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
'   Copyright (c) 2005-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.Cls
'   ProgID:         SP3DSweepolet.CSweepolet
'   Author:         kkk
'   Creation Date:  Wednesday, November 16 2005
'   Description:
'   The Details were taken from PDS Piping Component Data - Reference Guide at Page no D-75
'   SN=F172 and MC=SWOL. Symbol consist of Physical aspect only. Physical aspect is
'   made up of 1.Cylinder as Base, 2.Another Cylinder as Body, 3.Nozzle-1 at the
'   center of Pipe-1, Nozzle-2 along +ive Y-axis.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  18.APR.2007     KKC  CR-117893 Modified Olet symbol to be Trimmable
'   2.Jul.2007     svsmylav  TR-123363: Added PI/2 value to parAngle to support V6.1 instances.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double


Private Sub Class_Initialize()

 PI = 4 * Atn(1)

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim pipeDiam2       As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim iOutput     As Double
    
    Dim ObjBaseCylinderIns As Object
    Dim ObjNozCylinIns As Object
    Dim parFacetoHeaderCenter As Double
    Dim parSweepoletWidth As Double
    Dim parAngle As Double
    Dim dBaseDiameter As Double
    Dim dBasetoOutlet As Double
    Dim dCentertoOutlet As Double
    Dim parInsulationThickness As Double
    Dim BaseHalfTk As Double
    Dim InsulationDiameter As Double
    
' Inputs
   Set oPartFclt = arrayOfInputs(1)
    parFacetoHeaderCenter = arrayOfInputs(2)
    parSweepoletWidth = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    If UBound(arrayOfInputs) >= 5 Then 'Not to fail the prior to V7 instances by accessing index 31
        parAngle = arrayOfInputs(5)
    Else
        parAngle = PI / 2 'To suppport V6.1 instances
    End If
    
    iOutput = 0

  ' Insert your code for output 1(BaseCylinder Insulation)
    Dim startPt As AutoMath.DPosition
    Dim endPt As AutoMath.DPosition
    Set startPt = New AutoMath.DPosition
    Set endPt = New AutoMath.DPosition
    dBasetoOutlet = parFacetoHeaderCenter - (pipeDiam / 2)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset, depth
    ' According to PDS EDEN Sweepolet Base = 1/4th of Pipe diameter.
    startPt.Set 0, (pipeDiam / 2) * 3 / 4, 0
    endPt.Set 0, pipeDiam / 2, 0
    dBaseDiameter = pipeDiam2 + parSweepoletWidth * 2
    InsulationDiameter = dBaseDiameter + 2 * parInsulationThickness
    Set ObjBaseCylinderIns = PlaceCylinder(m_OutputColl, startPt, endPt, InsulationDiameter, True)
    
    Dim oDirVector As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotAboutZaxis As Double
    Set oDirVector = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    dRotAboutZaxis = PI / 2 - parAngle
    oDirVector.Set 0, 0, -1
    oTransformationMat.Rotate dRotAboutZaxis, oDirVector
    ObjBaseCylinderIns.Transform oTransformationMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBaseCylinderIns
    Set ObjBaseCylinderIns = Nothing
    
'Insert your code for Output 2 (Outlet Insulation)

    startPt.Set 0, (pipeDiam / 2) + sptOffset - depth, 0
    endPt.Set 0, parFacetoHeaderCenter, 0
    
    InsulationDiameter = pipeDiam2 + 2 * parInsulationThickness
    Set ObjNozCylinIns = PlaceCylinder(m_OutputColl, startPt, endPt, InsulationDiameter, True)
    ObjNozCylinIns.Transform oTransformationMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozCylinIns
    Set ObjNozCylinIns = Nothing
    Set startPt = Nothing
    Set endPt = Nothing
    Set oDirVector = Nothing
    Set oTransformationMat = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
  
End Sub

