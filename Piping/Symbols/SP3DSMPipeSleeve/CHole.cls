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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   ProgId        :        SP3DSMPipeSleeve.SMPipeSleeve
'   File          :        CHole.cls
'   Author        :        PK
'   Creation Date :        Wednesday 6, May 2007
'   Description   :        Penetration Pipe Sleeve
'   Source        :        B-PP-3.pdf
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   06.DEC.2007     PK     CR-131488 Created the symbol
'   11.AUG.2008     BS     CR-147688 Enhanced the symbol to Short Sleeve
'   11.AUG.2008     BS     CR-147689 Enhanced the symbol to  Long Sleeve
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Hole:"    'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim parInDia As Double
    Dim parOutDia As Double
    Dim parSleeveWidth1 As Double
    Dim parSleeveWidth2 As Double
    Dim parHoleDia As Double
    Dim parThickness As Double
    Dim parLength As Double
    
    Dim iOutput  As Integer
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parInDia = arrayOfInputs(2)
    parOutDia = arrayOfInputs(3)
'    parSleeveWidth1 = arrayOfInputs(4)
'    parSleeveWidth2 = arrayOfInputs(5)
'    parThickness = arrayOfInputs(6)
'    parHoleDia = arrayOfInputs(7)
'    parLength = arrayOfInputs(8)
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    iOutput = 0
    
    Dim oHole As Object
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Select Case lPartDataBasis
        Case Is <= 1
            parInDia = arrayOfInputs(2)
            parSleeveWidth1 = arrayOfInputs(4)
            parSleeveWidth2 = arrayOfInputs(5)
            
            'Insert your code for Hole
            oStPoint.Set parSleeveWidth1, 0, 0
            oEnPoint.Set -parSleeveWidth2, 0, 0
            Set oHole = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parOutDia, True)

        Case 1065           'long sleeve
            parThickness = arrayOfInputs(6)
            parHoleDia = arrayOfInputs(7)
            parLength = arrayOfInputs(8)
            
            'Insert your code for Hole
            oStPoint.Set parLength / 2, 0, 0
            oEnPoint.Set -parLength / 2, 0, 0
            Set oHole = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parHoleDia, True)
        Case 1066           'short sleeve
            parThickness = arrayOfInputs(6)
            parLength = arrayOfInputs(8)

            'Insert your code for Hole
            oStPoint.Set parLength / 2, 0, 0
            oEnPoint.Set -parLength / 2, 0, 0
            Set oHole = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parOutDia, True)
            
        Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.
        
    End Select
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oHole
    Set oHole = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oPipeComponent = Nothing
    Exit Sub
ErrorLabel:
   Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub



