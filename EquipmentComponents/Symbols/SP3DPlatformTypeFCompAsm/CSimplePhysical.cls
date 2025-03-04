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
Private PI As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
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
    Dim parPlatformAngle As Double
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
    parPlatformAngle = arrayOfInputs(8)
    parPlatformHeight = arrayOfInputs(9)
    
    iOutput = 0
    
    
' Insert your code for outputs
    Dim oLine           As IngrGeom3D.Line3d
    Dim oNormalVector   As New AutoMath.DVector
    Dim oLinesCollection          As New Collection
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    
    'Line 1
    stPoint.Set -parEqpDimension1, parEqpDimension6, parPlatformHeight
    enPoint.Set parEqpDimension2, parEqpDimension6, parPlatformHeight
        
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
    'Line 2
    
    
    
    stPoint.Set enPoint.x, enPoint.y, enPoint.z
    enPoint.x = parEqpDimension2
    enPoint.y = ((1 / (Tan(parPlatformAngle))) * parEqpDimension2) + (parEqpDimension4 / (Sin(PI - parPlatformAngle)))
    enPoint.z = parPlatformHeight
      
    
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
    'Line 3
    stPoint.Set enPoint.x, enPoint.y, enPoint.z
    enPoint.x = (parEqpDimension5 * Cos(parPlatformAngle - PI / 2)) + (parEqpDimension4 * Cos(PI - parPlatformAngle))
    enPoint.y = -(parEqpDimension5 * Sin(parPlatformAngle - PI / 2)) + (parEqpDimension4 * Sin(PI - parPlatformAngle))
    enPoint.z = parPlatformHeight
       
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
 
    'Line 4
    stPoint.Set enPoint.x, enPoint.y, enPoint.z
    
    enPoint.x = (parEqpDimension5 * Cos(parPlatformAngle - PI / 2)) - (parEqpDimension3 * Cos(PI - parPlatformAngle))
    enPoint.y = -(parEqpDimension5 * Sin(parPlatformAngle - PI / 2)) - (parEqpDimension3 * Sin(PI - parPlatformAngle))
    enPoint.z = parPlatformHeight
    
    
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
    
    
    'Line 5
    stPoint.Set enPoint.x, enPoint.y, enPoint.z
    
    enPoint.x = -parEqpDimension1
    enPoint.y = (-(1 / (Tan(parPlatformAngle))) * parEqpDimension1) - (parEqpDimension3 / (Sin(PI - parPlatformAngle)))
    enPoint.z = parPlatformHeight
    
    
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    
    
    'Line 6
    stPoint.Set enPoint.x, enPoint.y, enPoint.z
    enPoint.Set -parEqpDimension1, parEqpDimension6, parPlatformHeight
    
    
    Set oLine = PlaceLine(m_OutputColl.ResourceManager, stPoint, enPoint)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    oLinesCollection.Add oLine
    

    Dim comStpoint As New AutoMath.DPosition
    Dim oComplexString As IngrGeom3D.ComplexString3d
    
    comStpoint.Set -parEqpDimension1, parEqpDimension6, parPlatformHeight
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


