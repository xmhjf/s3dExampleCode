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
'
'   CPhysical.cls
'   Author:        MS
'   Creation Date: Tuesday, Oct 12 2004
'   Description:

'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "SimplePhysical:" 'Used for error messages

Private Sub Class_Initialize()
''
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double

    Dim parEqpDimension1 As Double
    Dim parEqpDimension2 As Double
    Dim parEqpDimension3 As Double
    Dim parEqpDimension4 As Double
    Dim parEqpDimension5 As Double
    Dim parEqpDimension6 As Double
    Dim parEqpDimension7 As Double
    Dim parEqpDimension8 As Double
    Dim parIsCornerSkewed As Double
    Dim parPlatformHeight As Double
    
    Dim objProjBody As Object
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parEqpDimension1 = arrayOfInputs(2)
    parEqpDimension2 = arrayOfInputs(3)
    parEqpDimension3 = arrayOfInputs(4)
    parEqpDimension4 = arrayOfInputs(5)
    parEqpDimension5 = arrayOfInputs(6)
    parEqpDimension6 = arrayOfInputs(7)
    parEqpDimension7 = arrayOfInputs(8)
    parEqpDimension8 = arrayOfInputs(9)
    parIsCornerSkewed = arrayOfInputs(10)
    parPlatformHeight = arrayOfInputs(11)

    iOutput = 0
    
    
' Insert your code for outputs
    Dim oLine           As IngrGeom3D.Line3d
    Dim oNormalVector   As New AutoMath.DVector
    Dim oLinesCollection          As New Collection
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    
    'Line 1
    stPoint.Set -parEqpDimension2, parEqpDimension7 + parEqpDimension8, parPlatformHeight
    enPoint.Set parEqpDimension3, parEqpDimension7 + parEqpDimension8, parPlatformHeight
        
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
  
    'Line 2
    If (parIsCornerSkewed <> 0) Then ' Not skewed
        stPoint.Set parEqpDimension3, parEqpDimension7 + parEqpDimension8, parPlatformHeight
        enPoint.Set parEqpDimension3, parEqpDimension7, parPlatformHeight
    Else
        stPoint.Set parEqpDimension3, parEqpDimension7 + parEqpDimension8, parPlatformHeight
        enPoint.Set parEqpDimension3 + parEqpDimension4 - 0.001, parEqpDimension7, parPlatformHeight
    End If
    
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
    'Line 3
    If (parIsCornerSkewed <> 0) Then ' Not skewed
        stPoint.Set parEqpDimension3, parEqpDimension7, parPlatformHeight
        enPoint.Set parEqpDimension3 + parEqpDimension4, parEqpDimension7, parPlatformHeight
    Else
        stPoint.Set parEqpDimension3 + parEqpDimension4 - 0.001, parEqpDimension7, parPlatformHeight
        enPoint.Set parEqpDimension3 + parEqpDimension4, parEqpDimension7, parPlatformHeight
    End If
    
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
    'Line 4
    stPoint.Set enPoint.x, enPoint.y, enPoint.z
    enPoint.Set parEqpDimension3 + parEqpDimension4, -parEqpDimension6, parPlatformHeight
    
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
    'Line 5
    If (parIsCornerSkewed <> 0) Then ' Not skewed
        stPoint.Set parEqpDimension3 + parEqpDimension4, -parEqpDimension6, parPlatformHeight
        enPoint.Set parEqpDimension3, -parEqpDimension6, parPlatformHeight
    Else
        stPoint.Set parEqpDimension3 + parEqpDimension4, -parEqpDimension6, parPlatformHeight
        enPoint.Set parEqpDimension3 - 0.001, -parEqpDimension6 - parEqpDimension5, parPlatformHeight
    End If
    
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
    'Line 6
    If (parIsCornerSkewed <> 0) Then ' Not skewed
        stPoint.Set parEqpDimension3, -parEqpDimension6, parPlatformHeight
        enPoint.Set parEqpDimension3, -parEqpDimension6 - parEqpDimension5, parPlatformHeight
    Else
        stPoint.Set parEqpDimension3 - 0.001, -parEqpDimension6 - parEqpDimension5, parPlatformHeight
        enPoint.Set parEqpDimension3, -parEqpDimension6 - parEqpDimension5, parPlatformHeight
    End If
    
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
    'Line 7
    stPoint.Set enPoint.x, enPoint.y, enPoint.z
    enPoint.Set -parEqpDimension2, -parEqpDimension5 - parEqpDimension6, parPlatformHeight
    
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
    
    'Line 8
    If (parIsCornerSkewed <> 0) Then ' Not skewed
        stPoint.Set -parEqpDimension2, -parEqpDimension5 - parEqpDimension6, parPlatformHeight
        enPoint.Set -parEqpDimension2, -parEqpDimension6, parPlatformHeight
    Else
        stPoint.Set -parEqpDimension2, -parEqpDimension5 - parEqpDimension6, parPlatformHeight
        enPoint.Set -parEqpDimension2 - parEqpDimension1 - 0.001, -parEqpDimension6, parPlatformHeight
    End If
    
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
    'Line 9
    If (parIsCornerSkewed <> 0) Then ' Not skewed
        stPoint.Set -parEqpDimension2, -parEqpDimension6, parPlatformHeight
        enPoint.Set -parEqpDimension2 - parEqpDimension1, -parEqpDimension6, parPlatformHeight
    Else
        stPoint.Set -parEqpDimension2 - parEqpDimension1 - 0.001, -parEqpDimension6, parPlatformHeight
        enPoint.Set -parEqpDimension2 - parEqpDimension1, -parEqpDimension6, parPlatformHeight
    End If
    
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
    'Line 10
    stPoint.Set enPoint.x, enPoint.y, enPoint.z
    enPoint.Set -parEqpDimension2 - parEqpDimension1, parEqpDimension7, parPlatformHeight

    
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
    'Line 11
    If (parIsCornerSkewed <> 0) Then ' Not skewed
        stPoint.Set -parEqpDimension2 - parEqpDimension1, parEqpDimension7, parPlatformHeight
        enPoint.Set -parEqpDimension2, parEqpDimension7, parPlatformHeight
    Else
        stPoint.Set -parEqpDimension2 - parEqpDimension1, parEqpDimension7, parPlatformHeight
        enPoint.Set -parEqpDimension2 - 0.001, parEqpDimension7 + parEqpDimension8, parPlatformHeight
    End If
    
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
    'Line 12
    If (parIsCornerSkewed <> 0) Then ' Not skewed
        stPoint.Set -parEqpDimension2, parEqpDimension7, parPlatformHeight
        enPoint.Set -parEqpDimension2, parEqpDimension7 + parEqpDimension8, parPlatformHeight
    Else
        stPoint.Set -parEqpDimension2 - 0.001, parEqpDimension7 + parEqpDimension8, parPlatformHeight
        enPoint.Set -parEqpDimension2, parEqpDimension7 + parEqpDimension8, parPlatformHeight
    End If
    
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
    
    Dim comStpoint As New AutoMath.DPosition
    Dim oComplexString As IngrGeom3D.ComplexString3d
    
    comStpoint.Set -parEqpDimension2, parEqpDimension7 + parEqpDimension8, parPlatformHeight
    Set oComplexString = PlaceTrCString(comStpoint, oLinesCollection)
    Set oLinesCollection = Nothing

' Project the complex string formed
    oNormalVector.Set 0, 0, -1
    Set objProjBody = PlaceProjection(m_OutputColl, oComplexString, oNormalVector, parPlatformHeight, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objProjBody
    Set objProjBody = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext

End Sub

''' This function creates persistent/Transient Line based on
''' start and end points of the line
'''<{(Line begin)}>
Private Function PlaceLine(pResourceMgr As IUnknown, ByRef startPoint As IJDPosition, _
                            ByRef endPoint As IJDPosition) _
                            As IngrGeom3D.Line3d

    Const METHOD = "PlaceLine"
    On Error GoTo ErrorHandler
        
    Dim oLine As IngrGeom3D.Line3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    ' Create Line object
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(pResourceMgr, _
                                startPoint.x, startPoint.y, startPoint.z, _
                                endPoint.x, endPoint.y, endPoint.z)
    
    
    Set PlaceLine = oLine
    Set oLine = Nothing
    Set oGeomFactory = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Function
'''<{(Line end)}>


