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
'**********************************************************************************
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   File:   SimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, Jul 4 2006
'   Description:
'        Mass Flow Controller(Source: Model SL7950S in Mass Flow Meter.pdf at www.emersonprocess.com)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CSimplePhysical" 'Used for error messages


Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
'
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs() As Variant, ByRef arrayOfOutputs() As String)
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
    Dim parInstrumentLength As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentHeight2 As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInstrumentLength = arrayOfInputs(3)
    parInstrumentHeight = arrayOfInputs(4)
    parInstrumentWidth = arrayOfInputs(5)
    parInstrumentHeight1 = arrayOfInputs(6)
    parInstrumentHeight2 = arrayOfInputs(7)
    
    iOutput = 0

' Insert your code for output 1(Mass Flow Controller Body)
    Dim ObjMassFlowControllerBody As Object
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
'   Assumptions:
'   (i) Length is taken to be equal to 90% of parInstrumentLength
'   (ii) Height is taken to be equal to 1.1 times the (2 * parInstrumentHeight)
    Dim dMassFCBodyLength As Double
    Dim dMassFCBodyHeight As Double
    
    dMassFCBodyLength = 0.9 * parInstrumentLength
    dMassFCBodyHeight = 1.1 * (2 * parInstrumentHeight)

    stPoint.Set -dMassFCBodyLength / 2, -parInstrumentHeight, -parInstrumentWidth / 2
    enPoint.Set stPoint.x + dMassFCBodyLength, stPoint.y + dMassFCBodyHeight, _
                    stPoint.z + parInstrumentWidth
    Set ObjMassFlowControllerBody = PlaceBox(m_OutputColl, stPoint, enPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMassFlowControllerBody
    Set ObjMassFlowControllerBody = Nothing
    
' Insert your code for output 2(Body 1(Rectangular Body))
    Dim ObjBody1 As Object
    Dim dLength As Double
    
'   Assumption: Body 1 length is taken to be equal to 65% of dMassFCBodyLength
    dLength = 0.65 * dMassFCBodyLength
    stPoint.Set -dMassFCBodyLength / 2, dMassFCBodyHeight / 2, -parInstrumentWidth / 2
    enPoint.Set stPoint.x + dLength, parInstrumentHeight1, stPoint.z + parInstrumentWidth
    Set ObjBody1 = PlaceBox(m_OutputColl, stPoint, enPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody1
    Set ObjBody1 = Nothing

' Insert your code for output 3(Body 2(Cylindrical Body))
    Dim ObjCylindricalBody As Object
    Dim dCylinderDiameter As Double
    
    dCylinderDiameter = dMassFCBodyLength - dLength

    stPoint.Set dMassFCBodyLength / 2 - dCylinderDiameter / 2, dMassFCBodyHeight / 2, 0
    enPoint.Set stPoint.x, parInstrumentHeight2, stPoint.z
    Set ObjCylindricalBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, dCylinderDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylindricalBody
    Set ObjCylindricalBody = Nothing
    
' Insert your code for output 4 (Left Hand Side Nozzzle Mounting Box)
    Dim ObjLHSNozMountingBox As Object

    stPoint.Set -parInstrumentLength / 2, -parInstrumentHeight, -parInstrumentWidth / 2
    enPoint.Set -dMassFCBodyLength / 2, stPoint.y + 2 * parInstrumentHeight, stPoint.z + parInstrumentWidth
    Set ObjLHSNozMountingBox = PlaceBox(m_OutputColl, stPoint, enPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLHSNozMountingBox
    Set ObjLHSNozMountingBox = Nothing
    
' Insert your code for output 5 (Right Hand Side Nozzzle Mounting Box)
    Dim ObjRHSNozMountingBox As Object
    stPoint.Set dMassFCBodyLength / 2, -parInstrumentHeight, -parInstrumentWidth / 2
    enPoint.Set parInstrumentLength / 2, stPoint.y + 2 * parInstrumentHeight, stPoint.z + parInstrumentWidth
    Set ObjRHSNozMountingBox = PlaceBox(m_OutputColl, stPoint, enPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRHSNozMountingBox
    Set ObjRHSNozMountingBox = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Dim dNozzleLength As Double
    dNozzleLength = (parFacetoFace - parInstrumentLength) / 2
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    dNozzleLength = (parFacetoFace - parInstrumentLength) / 2
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext

End Sub
