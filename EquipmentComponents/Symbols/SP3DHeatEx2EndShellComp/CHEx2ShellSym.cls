VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHEx2ShellSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CHEx2ShellSym.cls
'   ProgID:         SP3DHeatEx2EndShellComp.CHEx2ShellSym
'   Author:         BBK
'   Creation Date:  Monday, Nov 01, 2003

'   This is an Exchanger Shell component Body Symbol.
'   This symbol is Fully parametric symbol, consists of 77 inputs and 10 outputs.It has 2 aspects one is Simple Physical
'   another is ReferenceGeometry. Among them 8 are physical and 2 is ReferenceGeometry. It has Five nozzles and all
'   are fully parametric.
'   This is an equipment component symbol using the Smart Occurrence instead of Part occurrence. The symbol would be of
'   two parts, one smart occurrence symbol and the other is the custom assembly definition. This file defines the symbol
'   part of the complete component and the nozzles are created in the definition(CHEx2ShellDef.cls)
'   This is a Dow Equipment E-161 Exchanger Shell symbol.
'
'   Change History:
'   dd.mmm.yyyy         who             change description
'   -----------         ---             ------------------
'
'   01.Nov.2004     BBK     Initial Creation from CHEx2EndShell.cls
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CHEx2ShellSym:"  'Used for error messages

Private m_oSymbolHelper As IJSymbolHelper
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices


Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DHeatEx2EndShellComp"
    m_oSymbolHelper.ClassName = "CHEx2ShellSym"
 
'' Inputs
    m_oSymbolHelper.NumInputs = 77
    m_oSymbolHelper.AddInputDef 1, "VesselDiameter", "Exchanger Outer Diameter", 1.4986
    m_oSymbolHelper.AddInputDef 2, "VesselHeight", "Shell Body Height", 7.30885
    m_oSymbolHelper.AddInputDef 3, "CPtoFace", "Control point to Shell Top flange", 1.524
    m_oSymbolHelper.AddInputDef 4, "FaceDiameter", "Flange Diameter", 0.9398
    m_oSymbolHelper.AddInputDef 5, "FaceThickness", "Flange Thickness", 0.0508
    m_oSymbolHelper.AddInputDef 6, "NozzlePosition", "Nozzle C Position w.r.t Control Point", 1.905
    m_oSymbolHelper.AddInputDef 7, "NozzleOrientation", "Nozzle C Orientation", 0.61
    m_oSymbolHelper.AddInputDef 8, "RadialPosition", "Nozzle C Length", 1.016
    m_oSymbolHelper.AddInputDef 9, "Offset", "Nozzle C Offset", 0
    m_oSymbolHelper.AddInputDef 10, "NozzlePosition1", "Nozzle D Position w.r.t Control Point", 8.509
    m_oSymbolHelper.AddInputDef 11, "NozzleOrientation1", "Nozzle D Orientation", 0.785
    m_oSymbolHelper.AddInputDef 12, "RadialPosition1", "Nozzle D Length", 1.016
    m_oSymbolHelper.AddInputDef 13, "Offset1", "Nozzle D Offset", 0
    m_oSymbolHelper.AddInputDef 14, "NozzlePosition2", "Nozzle F1 Position w.r.t Control Point", 8.80745
    m_oSymbolHelper.AddInputDef 15, "NozzleOrientation2", "Nozzle F1 Orientation", 2.356
    m_oSymbolHelper.AddInputDef 16, "RadialPosition2", "Nozzle F1 Length", 1.016
    m_oSymbolHelper.AddInputDef 17, "Offset2", "Nozzle F1 Offset", 0
    m_oSymbolHelper.AddInputDef 18, "NozzlePosition3", "Nozzle F2 Position w.r.t Control Point", 1.5494
    m_oSymbolHelper.AddInputDef 19, "NozzleOrientation3", "Nozzle F2 Orientation", 2.356
    m_oSymbolHelper.AddInputDef 20, "RadialPosition3", "Nozzle F2 Length", 1.016
    m_oSymbolHelper.AddInputDef 21, "Offset3", "Nozzle F2 Offset", 0
    m_oSymbolHelper.AddInputDef 22, "NozzlePosition4", "Nozzle F3 Position w.r.t Control Point", 1.5494
    m_oSymbolHelper.AddInputDef 23, "NozzleOrientation4", "Nozzle F3 Orientation", 3.926
    m_oSymbolHelper.AddInputDef 24, "RadialPosition4", "Nozzle F3 Length", 1.016
    m_oSymbolHelper.AddInputDef 25, "Offset4", "Nozzle F3 Offset", 0
    
    m_oSymbolHelper.AddInputDef 26, "Npd", "NPD", 100
    m_oSymbolHelper.AddInputDef 27, "EndPreparation", "End Preparation", 21
    m_oSymbolHelper.AddInputDef 28, "ScheduleThickness", "Schedule Thickness", 100
    m_oSymbolHelper.AddInputDef 29, "EndStandard", "End Standard", 5
    m_oSymbolHelper.AddInputDef 30, "PressureRating", "Pressure Rating", 35
    m_oSymbolHelper.AddInputDef 31, "FlowDirection", "Flow Direction", 3
    m_oSymbolHelper.AddInputDef 32, "PortIndex1", "Port Index1", 1
    m_oSymbolHelper.AddInputDef 33, "Npd1", "NPD1", 150 ''6inch
    m_oSymbolHelper.AddInputDef 34, "EndPreparation1", "End Preparation1", 0
    m_oSymbolHelper.AddInputDef 35, "ScheduleThickness1", "Schedule Thickness1", 0
    m_oSymbolHelper.AddInputDef 36, "EndStandard1", "End Standard1", 0
    m_oSymbolHelper.AddInputDef 37, "PressureRating1", "Pressure Rating1", 0
    m_oSymbolHelper.AddInputDef 38, "FlowDirection1", "Flow Direction1", 0
    m_oSymbolHelper.AddInputDef 39, "PortIndex2", "Port Index2", 2
    m_oSymbolHelper.AddInputDef 40, "Npd2", "NPD2", 100
    m_oSymbolHelper.AddInputDef 41, "EndPreparation2", "End Preparation2", 0
    m_oSymbolHelper.AddInputDef 42, "ScheduleThickness2", "Schedule Thickness2", 0
    m_oSymbolHelper.AddInputDef 43, "EndStandard2", "End Standard2", 0
    m_oSymbolHelper.AddInputDef 44, "PressureRating2", "Pressure Rating2", 0
    m_oSymbolHelper.AddInputDef 45, "FlowDirection2", "Flow Direction2", 0
    m_oSymbolHelper.AddInputDef 46, "PortIndex3", "Port Index3", 3
    m_oSymbolHelper.AddInputDef 47, "Npd3", "NPD3", 25  ''1 inch
    m_oSymbolHelper.AddInputDef 48, "EndPreparation3", "End Preparation3", 0
    m_oSymbolHelper.AddInputDef 49, "ScheduleThickness3", "Schedule Thickness3", 0
    m_oSymbolHelper.AddInputDef 50, "EndStandard3", "End Standard3", 0
    m_oSymbolHelper.AddInputDef 51, "PressureRating3", "Pressure Rating3", 0
    m_oSymbolHelper.AddInputDef 52, "FlowDirection3", "Flow Direction3", 0
    m_oSymbolHelper.AddInputDef 53, "PortIndex4", "Port Index4", 4
    m_oSymbolHelper.AddInputDef 54, "Npd4", "NPD4", 25
    m_oSymbolHelper.AddInputDef 55, "EndPreparation4", "End Preparation4", 0
    m_oSymbolHelper.AddInputDef 56, "ScheduleThickness4", "Schedule Thickness4", 0
    m_oSymbolHelper.AddInputDef 57, "EndStandard4", "End Standard4", 0
    m_oSymbolHelper.AddInputDef 58, "PressureRating4", "Pressure Rating4", 0
    m_oSymbolHelper.AddInputDef 59, "FlowDirection4", "Flow Direction4", 0
    m_oSymbolHelper.AddInputDef 60, "PortIndex5", "Port Index5", 5
    m_oSymbolHelper.AddInputDef 61, "Npd5", "NPD5", 25
    m_oSymbolHelper.AddInputDef 62, "EndPreparation5", "End Preparation5", 0
    m_oSymbolHelper.AddInputDef 63, "ScheduleThickness5", "Schedule Thickness5", 0
    m_oSymbolHelper.AddInputDef 64, "EndStandard5", "End Standard5", 0
    m_oSymbolHelper.AddInputDef 65, "PressureRating5", "Pressure Rating5", 0
    m_oSymbolHelper.AddInputDef 66, "FlowDirection5", "Flow Direction5", 0
    m_oSymbolHelper.AddInputDef 67, "Id1", "Id1", "C"
    m_oSymbolHelper.AddInputDef 68, "Id2", "Id2", "D"
    m_oSymbolHelper.AddInputDef 69, "Id3", "Id3", "F1"
    m_oSymbolHelper.AddInputDef 70, "Id4", "Id4", "F2"
    m_oSymbolHelper.AddInputDef 71, "Id5", "Id5", "F3"
    m_oSymbolHelper.AddInputDef 72, "NpdUnitType", "Npd Unit Type", "mm"
    m_oSymbolHelper.AddInputDef 73, "NpdUnitType1", "Npd Unit Type1", ""
    m_oSymbolHelper.AddInputDef 74, "NpdUnitType2", "Npd Unit Type2", ""
    m_oSymbolHelper.AddInputDef 75, "NpdUnitType3", "Npd Unit Type3", ""
    m_oSymbolHelper.AddInputDef 76, "NpdUnitType4", "Npd Unit Type4", ""
    m_oSymbolHelper.AddInputDef 77, "NpdUnitType5", "Npd Unit Type5", ""

' Outputs
    m_oSymbolHelper.NumOutputs = 10
    m_oSymbolHelper.AddOutputDef 1, "ShellFlange", "Shell Flange", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "ShellBodyCyl", "Cylindrical Exchanger Shell Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "ShellBottomFlange", "Shell Bottom Flange", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "C", "Nozzle A1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "D", "Nozzle A2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "F1", "Nozzle B1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "F2", "Nozzle B2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "F3", "Nozzle E1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "VesselAxisPoint", "Reference Point on axis of vessel at Coontrol Point", ReferenceGeometry
    m_oSymbolHelper.AddOutputDef 10, "DefaultSurface", "Reference Surface at the CP of vessel", ReferenceGeometry

' Aspects
    m_oSymbolHelper.NumAspects = 2
    m_oSymbolHelper.AddAspectDef 1, "SimplePhysical", "Physical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "RefGeometry", "ReferenceGeometry", ReferenceGeometry
    
    Exit Sub
    
ErrorHandler:
Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext

End Sub

Private Sub Class_Terminate()
    Set m_oSymbolHelper = Nothing
End Sub

' note that the rest of the code is generated by the Symbol Wizard and should not need to be modified
Public Function IJDUserSymbolServices_InstanciateDefinition( _
                ByVal CodeBase As String, _
                ByVal defParameters As Variant, _
                ByVal ActiveConnection As Object) As Object
    ' call symbol services default implementation of this method
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition:"
On Error GoTo Errx
    Set IJDUserSymbolServices_InstanciateDefinition = m_oSymbolHelper.InstanciateDefinition(CodeBase, defParameters, ActiveConnection)
    
    Exit Function

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Function
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
Const METHOD = "IJDUserSymbolServices_GetDefinitionName:"
  On Error GoTo Errx
      IJDUserSymbolServices_GetDefinitionName = m_oSymbolHelper.ProjectName + "." + m_oSymbolHelper.ClassName
    
    Exit Function

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition:"
On Error GoTo Errx
    m_oSymbolHelper.InitializeSymbolDefinition pSymbolDefinition
Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
      
End Sub

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                    ByVal repName As String, _
                    ByVal outputcoll As Object, _
                    ByRef arrayOfInputs())
Const METHOD = "IJDUserSymbolServices_InvokeRepresentation:"
On Error GoTo Errx
    m_oSymbolHelper.InvokeRepresentation sblOcc, repName, outputcoll, arrayOfInputs
        
Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
            Err.HelpFile, Err.HelpContext
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
Const METHOD = "IJDUserSymbolServices_EditOccurence:"
On Error GoTo Errx
    ' The definition uses the generic EditOccurrence command
    IJDUserSymbolServices_EditOccurence = False
        Exit Function

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
            Err.HelpFile, Err.HelpContext

End Function




