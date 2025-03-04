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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.

'   ProgID          :  SP3DCableRiserTray.RiserTray
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Friday, 25 April 2008
'   Description     :  Cable Tray Riser Box
'   Change History  :
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   25-APR-2008      PK        CR-136329:Created the symbol.
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
''''

End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput                 As Double
    Dim ActualWidth             As Double
    Dim ActualDepth             As Double
    Dim parWidth                As Double
    Dim parDepth                As Double
    Dim parHeight               As Double
    Dim parPortOffset           As Double
    Dim parThickness            As Double
                  
    Dim oPort1   As AutoMath.DPosition
    Dim oPort2   As AutoMath.DPosition
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    Set oPort1 = New AutoMath.DPosition
    Set oPort2 = New AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    iOutput = 0
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parHeight = arrayOfInputs(4)
    parPortOffset = arrayOfInputs(5)
    parThickness = arrayOfInputs(6)
    
    Dim lCableTrayPartDataBasis As Long
    Dim oTrayPart As IJCableTrayPart
    Set oTrayPart = oPartFclt
    lCableTrayPartDataBasis = oTrayPart.PartDataBasis
    
    Dim oRiserBox           As Object
    Dim oDir                As AutoMath.DVector
    Dim oRadialOrient       As AutoMath.DVector
    Dim oCableTrayPort      As GSCADNozzleEntities.IJCableTrayPortOcc
    
    If lCableTrayPartDataBasis = 138 Then
    
        'Create Riser Box
        oStPoint.Set 0, parWidth / 2, parDepth / 2
        oEnPoint.Set parHeight, -parWidth / 2, -parDepth / 2
        Set oRiserBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRiserBox
        Set oRiserBox = Nothing
   
        'Place Port 1
        Set oDir = New AutoMath.DVector
        Set oRadialOrient = New AutoMath.DVector
        oDir.Set -1, 0, 0
        oRadialOrient.Set 0, 0, 1
        oPort1.Set -parPortOffset - parThickness, 0, 0
            
        Set oCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPort1, oDir, _
                                                oRadialOrient, m_OutputColl)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCableTrayPort
        Set oCableTrayPort = Nothing
        Set oPort1 = Nothing
        Set oDir = Nothing
        Set oRadialOrient = Nothing
     
        'Place Port 2
        Set oDir = New AutoMath.DVector
        Set oRadialOrient = New AutoMath.DVector
        oPort2.Set parHeight, 0, 0
        oDir.Set 1, 0, 0
        oRadialOrient.Set 0, 0, 1
        
        Set oCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPort2, oDir, _
                                                    oRadialOrient, m_OutputColl)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCableTrayPort
        Set oPort2 = Nothing
        Set oCableTrayPort = Nothing
        Set oDir = Nothing
        Set oRadialOrient = Nothing
    Else
        
        If lCableTrayPartDataBasis = 139 Then
        
            'Create Riser Box
            oStPoint.Set -parHeight / 2, parWidth / 2, parDepth / 2
            oEnPoint.Set parHeight / 2, -parWidth / 2, -parDepth / 2
            Set oRiserBox = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRiserBox
            Set oRiserBox = Nothing
                      
            'Place Port 1
            Set oDir = New AutoMath.DVector
            Set oRadialOrient = New AutoMath.DVector
            oDir.Set -1, 0, 0
            oRadialOrient.Set 0, 0, 1
            oPort1.Set -parHeight / 2 - parPortOffset - parThickness, 0, 0
            
            Set oCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPort1, oDir, _
                                                        oRadialOrient, m_OutputColl)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCableTrayPort
            Set oCableTrayPort = Nothing
            Set oPort1 = Nothing
            Set oDir = Nothing
            Set oRadialOrient = Nothing
                    
            'Place Port 2
            Set oDir = New AutoMath.DVector
            Set oRadialOrient = New AutoMath.DVector
            oPort2.Set 0, 0, parDepth / 2
            oDir.Set 0, 0, 1
            oRadialOrient.Set -1, 0, 0
                
            Set oCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPort2, oDir, _
                                                            oRadialOrient, m_OutputColl)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCableTrayPort
        
            Set oCableTrayPort = Nothing
            Set oDir = Nothing
            Set oRadialOrient = Nothing
            Set oPort2 = Nothing
        End If
    End If
    Exit Sub
    
ErrorLabel:

    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
