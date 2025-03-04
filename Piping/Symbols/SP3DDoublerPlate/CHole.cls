VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   ProgId        :        SP3DDoublerPlate.DoublerPlate
'   File          :        CHole.cls
'   Author        :        PK
'   Creation Date :        Wednesday 5, May 2007
'   Description   :        Doubler Plate Penetration
'   Source        :        B-PP-1A.pdf,B-PP-7.pdf and B-PP-1D.pdf
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   05.DEC.07       PK     CR-131488:Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Hole:"    'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart

    Dim parInDia                As Double
    Dim parOutDia               As Double
    Dim parThickness            As Double
    Dim parThicknessMultiplier  As Double
    Dim parHoleDia              As Double
    
    Dim iOutput                 As Integer
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInDia = arrayOfInputs(2)
    parOutDia = arrayOfInputs(3)
    parThickness = arrayOfInputs(4)
    parThicknessMultiplier = arrayOfInputs(5)
    parHoleDia = arrayOfInputs(6)
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    iOutput = 0
    
    'Insert your code for Hole
    Dim oHole As Object
    oStPoint.Set 0, 0, 0
    oEnPoint.Set -parThickness, 0, 0
    Set oHole = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parHoleDia, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHole
    
    Set oHole = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub
ErrorLabel:
   Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
