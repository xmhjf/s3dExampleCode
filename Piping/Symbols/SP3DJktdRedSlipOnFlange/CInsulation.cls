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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         ACM
'   Creation Date:  Wednesday, July 14 2004
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       This is a Reducing Slip-on Flange for Jacked-Piping.
'       On One side it has,One for Jacketed Pipe and One port of core pipe.
'       Other side it has one port.'
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub

Private Sub Class_Initialize()
    Set m_oGeomHelper = New SymbolServices
    
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim oCenterPos       As New AutoMath.DPosition
    Dim oEndPos As New AutoMath.DPosition

    oCenterPos.Set 0, 0, 0
    
    m_oGeomHelper.OutputCollection = m_OutputColl

    Dim iOutput     As Double
    
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInsulationThickness = arrayOfInputs(2)
    
    iOutput = 0

 ' Insert your code for output 1(Insulated Body)
 ''Place cylider with the Jacked pipe or flange diameter whichever is greater.
    Dim dInsulationDiameter As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        
    If flangeDiam > pipeDiam Then
        dInsulationDiameter = flangeDiam + parInsulationThickness * 2
    Else
        dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    End If
                
    oEndPos.Set flangeThick + parInsulationThickness, 0, 0

' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oCenterPos, oEndPos, dInsulationDiameter
     
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
