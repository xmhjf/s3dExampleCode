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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         RRK
'   Creation Date:  Tuesday 17, Apr 2007
'   Description:
'   This implements the Insulation aspect of JIS-F7121 Water Filter symbol to
'   support multiple dimensional bases as per CR 118311.
'   The symbol supports the following dimensional bases
'   1.)Part Data Basis Value 240:-Water filter, outlet port at 180 degrees
'   2.)Part Data Basis Value 245:-Water filter, outlet port at 90 degrees, clockwise
'   3.)Part Data Basis Value 250:-Water filter, outlet port at 90 degrees, counter-clockwise

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
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
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    
    Dim parFacetoFace As Double
    Dim parFacetoCenter As Double
    Dim parHeight As Double
    Dim parHeight1 As Double
    Dim parHeight2 As Double
    Dim parInsulationThickness As Double
    
    Dim dInsulationDia As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parFacetoCenter = arrayOfInputs(3)
    parHeight = arrayOfInputs(4)
    parHeight1 = arrayOfInputs(5) 'height of the Large vertical cylinder below the center-line.
    parHeight2 = arrayOfInputs(6) 'height of the Large vertical cylinder above the center-line.
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0

'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Dim dFacetoCenter As Double
    
    Select Case lPartDataBasis
        Case Is <= 1, 240
            parFacetoFace = arrayOfInputs(2)
            dFacetoCenter = parFacetoFace / 2
            
        Case 245, 250
            parFacetoCenter = arrayOfInputs(3)
            dFacetoCenter = parFacetoCenter
            
        Case Else
            GoTo ErrorLabel 'Invalid specification.
    End Select
    
'Place Output 7(Insulation for vertical cylinder)

    Dim objInsVerCylinder As Object
    Dim oStPoint   As Object
    Dim oEnPoint   As Object
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
' Assumption: Length of the horizontal cylinder protrusion on either sides of the Large
'             vertical cylinder is twice of the flange thickness.
    Dim dCylinderDiam As Double
    Dim dFlangeThickness As Double
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    dFlangeThickness = flangeThick
' Check for the case where flange thickness is zero
    If CmpDblEqual(flangeThick, 0) Then dFlangeThickness = 0.1 * dFacetoCenter
        
    'Diameter of the larger cylinder = Face to Face - Twice Flange Thickness
    ' - Twice length of the smaller horizontal cylinders.
    dCylinderDiam = 2 * dFacetoCenter - 6 * dFlangeThickness
    
    oStPoint.Set 0, -parHeight1 - parInsulationThickness, 0
    oEnPoint.Set 0, parHeight2, 0
        
    dInsulationDia = dCylinderDiam + 2 * parInsulationThickness
    
    Set objInsVerCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsVerCylinder
    Set objInsVerCylinder = Nothing

'   Place Output 8(Insulation for Cover)
    Dim ObjInsCover As Object
    Dim dCylTopDia As Double
    Dim dInsCoverThickness As Double
    
'   Assumptions
    dCylTopDia = 1.1 * dCylinderDiam
    dInsCoverThickness = 0.45 * parHeight2 + parInsulationThickness
    
    oStPoint.Set 0, parHeight2, 0
    oEnPoint.Set 0, parHeight2 + dInsCoverThickness, 0
    
    dInsulationDia = dCylTopDia + 2 * parInsulationThickness
    Set ObjInsCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsCover
    Set ObjInsCover = Nothing

'   Place Output 9(Insulation for Top of the Cover)
    Dim ObjInsTopCover As Object
    Dim dCoverTopDia As Double
    Dim dTopCylThickness As Double
    
    dCoverTopDia = 0.2 * dCylinderDiam
    dTopCylThickness = 0.25 * parHeight2
    
    oStPoint.Set 0, parHeight2 + dTopCylThickness, 0
    oEnPoint.Set 0, parHeight - parHeight1 + parInsulationThickness, 0
    
    dInsulationDia = dCoverTopDia + 2 * parInsulationThickness
    Set ObjInsTopCover = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsTopCover
    Set ObjInsTopCover = Nothing
    
'   Place Output 10(Insulation for Horizontal Cylinder at Port 1)
    
    Dim objInsPort1HorCyl As Object

    oStPoint.Set -dFacetoCenter, 0, 0
    oEnPoint.Set 0, 0, 0
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    
    Set objInsPort1HorCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort1HorCyl
    Set objInsPort1HorCyl = Nothing

'   Place Output 11(Insulation for Horizontal Cylinder at Port 2)
    
    Dim objInsPort2HorCyl As Object
    
    Select Case lPartDataBasis
        Case Is <= 1, 240
        oStPoint.Set dFacetoCenter, 0, 0
        
        Case 245
        oStPoint.Set 0, 0, -dFacetoCenter
        
        Case 250
        oStPoint.Set 0, 0, dFacetoCenter
    End Select
    
    oEnPoint.Set 0, 0, 0
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    
    Set objInsPort2HorCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort2HorCyl
    Set objInsPort2HorCyl = Nothing

 ' Insert your code for output 12(ObjInsPort1)
    Dim ObjInsPort1 As Object
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    dFlangeThickness = flangeThick
' Check for the case where flange thickness is zero
    If CmpDblEqual(flangeThick, 0) Then dFlangeThickness = 0.001
    
    oStPoint.Set -dFacetoCenter, 0, 0
    oEnPoint.Set -dFacetoCenter + dFlangeThickness + parInsulationThickness, 0, 0
    
    dInsulationDia = flangeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing
    
 ' Insert your code for output 13(ObjInsPort2)
    Dim ObjInsPort2 As Object
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    dFlangeThickness = flangeThick
' Check for the case where flange thickness is zero
    If CmpDblEqual(flangeThick, 0) Then dFlangeThickness = 0.001
    
    Select Case lPartDataBasis
        Case Is <= 1, 240
        oStPoint.Set dFacetoCenter, 0, 0
        oEnPoint.Set dFacetoCenter - dFlangeThickness - parInsulationThickness, 0, 0
    
        Case 245
        oStPoint.Set 0, 0, -dFacetoCenter
        oEnPoint.Set 0, 0, -dFacetoCenter + dFlangeThickness + parInsulationThickness
        
        Case 250
        oStPoint.Set 0, 0, dFacetoCenter
        oEnPoint.Set 0, 0, dFacetoCenter - dFlangeThickness - parInsulationThickness
    End Select

    dInsulationDia = flangeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDia = pipeDiam + parInsulationThickness * 2
    Set ObjInsPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort2
    Set ObjInsPort2 = Nothing

    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
