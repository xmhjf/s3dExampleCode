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
'   Copyright (c) 2005-07, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   ProgID:         SP3D4WayInstRootVal.C4WayInstRootVal
'   Author:         kkk
'   Creation Date:  Wednesday, November 16 2005
'   Description:
'   The Details about this Symbol is taken from PDS Piping Component Data Reference Guide. Symbol
'   SN=V82 and MC=4WRV1, at Page No.  D-33  The Symbol consist of Physical and Insulation aspects.
'   Physical Aspect is made up of 1.Nozzle with Length, 2.Cylinder of diameter equal to nozzle 2
'   flange Diameter, 3.Nozzle at port 2, 4. Nozzle 3 and 5. Nozzle 4
'   Insulation Aspect consist of Insulation Cylinder for Nozzle1 (Including Body), Nozzle2(Including Body),
'   Nozzle3, Nozzle4
'
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   08.SEP.2006     KKC     DI-95670  Replace names with initials in all revision history sheets and symbols
'   27.Aug.2007     VRG     TR-124959 Insulation for ports was ensured for different end preparations
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()

End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt        As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim pipeDiam3        As Double
    Dim pipeDiam4        As Double
    
    Dim flangeThick1     As Double
    Dim flangeThick2     As Double
    Dim flangeThick3     As Double
    Dim flangeThick4     As Double
    
    Dim sptOffset1       As Double
    Dim sptOffset2       As Double
    Dim sptOffset3       As Double
    Dim sptOffset4       As Double
    
    Dim flangeDiam1      As Double
    Dim flangeDiam2      As Double
    Dim flangeDiam3      As Double
    Dim flangeDiam4      As Double
    
    Dim depth1           As Double
    Dim depth2           As Double
    Dim depth3           As Double
    Dim depth4           As Double
        
    Dim iOutput     As Integer
    Dim ObjBodyCylinIns As Object
    Dim ObjBodyPipeDiaIns As Object
    Dim ObjNozzle1Ins As Object
    Dim ObjNozzle2Ins As Object
    Dim ObjNozzle3Ins As Object
    Dim ObjNozzle4Ins As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parCentertoTap As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parCentertoTap = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0

 ' Insert your code for output 2(BodyCylinIns)
    
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    Dim dInsDiameter As Double
    Dim dBodyDia As Double
   
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, sptOffset3, depth3
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick4, flangeDiam4, sptOffset4, depth4
    
    dBodyDia = ReturnMax4(pipeDiam1, pipeDiam2, flangeDiam1, flangeDiam2)
    dInsDiameter = dBodyDia + parInsulationThickness * 2
    
    startPoint.Set -parFace1toCenter, 0, 0
    endPoint.Set 0, 0, 0
    Set ObjBodyPipeDiaIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, dInsDiameter, True)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyPipeDiaIns
    Set ObjBodyPipeDiaIns = Nothing
      
' Insert your code for output 4(Nozzle2 and Rest of Body)
    startPoint.Set 0, 0, 0
    endPoint.Set parFace2toCenter, 0, 0
    Set ObjBodyCylinIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, dInsDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCylinIns
    Set ObjBodyCylinIns = Nothing
        
 ' Insert your code for output 5(Nozzle3Ins)
    If CmpDblEqual(flangeThick3, 0) Then
            flangeThick3 = NEGLIGIBLE_THICKNESS
    End If
    startPoint.Set parCentertoTap, dBodyDia / 2, 0
    endPoint.Set parCentertoTap, dBodyDia / 2 - flangeThick3 - parInsulationThickness, 0
    dInsDiameter = flangeDiam3 + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam3, flangeDiam3) Then dInsDiameter = pipeDiam3 + parInsulationThickness * 2
    Set ObjNozzle3Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, dInsDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle3Ins
    Set ObjNozzle3Ins = Nothing
    
 ' Insert your code for output 6(Nozzle4Ins)
    If CmpDblEqual(flangeThick4, 0) Then
            flangeThick4 = NEGLIGIBLE_THICKNESS
    End If
    startPoint.Set parCentertoTap, -dBodyDia / 2, 0
    endPoint.Set parCentertoTap, -dBodyDia / 2 + flangeThick4 + parInsulationThickness, 0
    dInsDiameter = flangeDiam4 + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam4, flangeDiam4) Then dInsDiameter = pipeDiam4 + parInsulationThickness * 2
    Set ObjNozzle4Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, dInsDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle4Ins
    Set ObjNozzle4Ins = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
 
End Sub

Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If CmpDblLessThan(MaxValue, B) Then MaxValue = B
    If CmpDblLessThan(MaxValue, C) Then MaxValue = C
    If CmpDblLessThan(MaxValue, D) Then MaxValue = D
    ReturnMax4 = MaxValue
End Function
