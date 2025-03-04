VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:17:40 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:          BG
'   Creation Date:  Thursday, Feb 6 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' This Symbol is a replicate of "Tank.TankServices" Symbol, except that the symbol is in line with
' Symbol created using wizard.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   18.May.2005     kkk         CR-76070: Removed the Nozzle related parameters and code.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Maintenance:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    Dim ObjTankEnvelope As Object
    
    Dim parVesselLength As Double
    Dim parVesselDiameter As Double
    Dim parSupportLength As Double
    Dim parSupportHeight As Double
    Dim parSupportThickness As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselLength = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parSupportLength = arrayOfInputs(4)
    parSupportHeight = arrayOfInputs(5)
    parSupportThickness = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0
    
    Dim lowerPts As AutoMath.DPosition
    Dim upperPts As AutoMath.DPosition
    Dim dNozLen As Double
    
    Set lowerPts = New AutoMath.DPosition
    Set upperPts = New AutoMath.DPosition
    
'    Assumption: Nozzle length is taken to be 0.1 times the tank diameter
    dNozLen = parVesselDiameter * 0.1
'Elliptical head minor radius is 0.5 times the Tank Radius
    lowerPts.Set -0.5 * 0.5 * parVesselDiameter - (parVesselLength + 0.5 * parVesselDiameter) * 0.1, _
                            -(parVesselDiameter / 2 * 1.1), _
                            -((parSupportHeight + parVesselDiameter + dNozLen) * 0.1)
    upperPts.Set (parVesselLength + 0.5 * 0.5 * parVesselDiameter) * 1.1, _
                            parVesselDiameter / 2 * 1.1, _
                            (parSupportHeight + parVesselDiameter + dNozLen) * 1.1
    
    Set ObjTankEnvelope = PlaceBox(m_OutputColl, lowerPts, upperPts)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTankEnvelope
    Set ObjTankEnvelope = Nothing
    Set lowerPts = Nothing
    Set upperPts = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
