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
'   ProgId        :        SP3DSMPipeSleeveWSeal.SleeveWSeal
'   File          :        CHole.cls
'   Author        :        PK
'   Creation Date :        Wednesday 6, May 2007
'   Description   :        Penetration Pipe Sleeve with Seal
'   Source        :        B-PP-13/13A.pdf
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   06.DEC.2007     PK     CR-131488 Created the symbol
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Hole:"    'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim parOutDia         As Double
    Dim parInDia          As Double
    Dim parSealWidth      As Double
    Dim parSleeveWidth    As Double
       
    Dim iOutput  As Integer
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInDia = arrayOfInputs(2)
    parOutDia = arrayOfInputs(3)
    parSleeveWidth = arrayOfInputs(4)
    parSealWidth = arrayOfInputs(5)
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    iOutput = 0
    
    'Insert your code for Hole
    Dim oHole As Object
    oStPoint.Set -parSleeveWidth / 2, 0, 0
    oEnPoint.Set parSleeveWidth / 2, 0, 0
    Set oHole = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parOutDia, True)
    
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

