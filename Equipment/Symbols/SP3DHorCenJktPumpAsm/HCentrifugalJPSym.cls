VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HCentrifugalJPSym"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"SYM"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:14:28 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   HCentrifugalJPSym.cls
'   ProgID:         SP3DHorCenJktPumpAsm.HCentrifugalJPSym
'   Author:         svsmylav
'   Creation Date:  Tuesday, 26 April 2003
'   Description:
'   The symbol is prepared based on symbol of The DOW Chemical company.
'   The symbol consists of 111 inputs, 24 Physical outputs and 2 Reference Aspect outputs(control point and default surface)
'   The pump has a suction, discharge, drain, flush connection, Jacket In, Lantern In, Jacket Out, Lantern Out Nozzles.
'   All the nozzles are fully parametric.
'   The symbol has one foundation port with the port origin at the center of the base of support. The holes
'   location is governed by the XboltHole and YboltHole Input parameters.
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "HCentrifugalJPSym:"  'Used for error messages

Private m_oSymbolHelper As IJSymbolHelper
    
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices

Private Sub Class_Initialize()
    
    Const METHOD = "Class_Initialize:"
    On Error GoTo ErrorHandler
    
    Set m_oSymbolHelper = New SymbolServices
    m_oSymbolHelper.ProjectName = "SP3DHorCenJktPumpAsm"
    m_oSymbolHelper.ClassName = "HCentrifugalJPSym"
    
' Inputs
    m_oSymbolHelper.NumInputs = 14
    m_oSymbolHelper.AddInputDef 1, "DisNozzletoFace", "CL Discharge Nozzle to Face", 0.506984
    m_oSymbolHelper.AddInputDef 2, "DisNozzletoChFace", "CL Discharge Nozzle to Chamber Face", 0.355092
    m_oSymbolHelper.AddInputDef 3, "DischargetoFaceSuction", "CL of Discharge to Face of Suction Nozzle", 0.102108
    m_oSymbolHelper.AddInputDef 4, "BaseplatetoSuction", "Bottom of Pump Baseplate to CL of Suction", 0.210058
    m_oSymbolHelper.AddInputDef 5, "PumptoFaceDisNozzle", "CL of Pump to Face of Discharge Nozzle", 0.242062
    m_oSymbolHelper.AddInputDef 6, "DisNozzletoFaceBaseplate", "CL of Discharge Nozzle to Right Face of Support", 0.254508
    m_oSymbolHelper.AddInputDef 7, "MotorLength", "Motor Length", 0.735076
    m_oSymbolHelper.AddInputDef 8, "CouplingLength", "Coupling Length", 0.092202
    m_oSymbolHelper.AddInputDef 9, "SupportLength", "Support Total Length", 1.5748
    m_oSymbolHelper.AddInputDef 10, "SupportHeight", "Support Total Height", 0.15748
    m_oSymbolHelper.AddInputDef 11, "SupportWidth", "Support Total Width", 0.6096
    m_oSymbolHelper.AddInputDef 12, "XboltHole", "Bolt X Hole Location", 0.2667
    m_oSymbolHelper.AddInputDef 13, "YboltHole", "Bolt Y Hole Location", 1.4478
   
'    Base Type
    m_oSymbolHelper.AddInputDef 14, "BaseType", "Type of Support Base: PolyShield or Structure Steel", "Structure Steel"
        
' Outputs
    m_oSymbolHelper.NumOutputs = 24
    m_oSymbolHelper.AddOutputDef 1, "MotorBody", "Motor Body", SimplePhysical
    m_oSymbolHelper.AddOutputDef 2, "MotorEllipticalCap1", "Motor Elliptical Rear Cap", SimplePhysical
    m_oSymbolHelper.AddOutputDef 3, "MotorEllipticalCap2", "Motor Elliptical Near Cap", SimplePhysical
    m_oSymbolHelper.AddOutputDef 4, "Coupling", "Shaft couping", SimplePhysical
    m_oSymbolHelper.AddOutputDef 5, "Shaft", "Shaft", SimplePhysical
    m_oSymbolHelper.AddOutputDef 6, "PumpCylinder1", "Pump Cylinder 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 7, "PumpCylinder2", "Pump Cylinder 2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 8, "PumpCylinder3", "Pump Cylinder 3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 9, "PumpCylinder4", "Pump Cylinder 4", SimplePhysical
    
''Support(Base) is a set of planes
    m_oSymbolHelper.AddOutputDef 10, "SupportBodyTopPlane", "Support Body Top Plane", SimplePhysical
    m_oSymbolHelper.AddOutputDef 11, "DefaultSurface", "Support Body Bottom Plane", SimplePhysical
    m_oSymbolHelper.AddOutputDef 12, "SupportBodySidePlane1", "Support Body side plane1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 13, "SupportBodySidePlane2", "Support Body side plane2", SimplePhysical
    m_oSymbolHelper.AddOutputDef 14, "SupportBodySidePlane3", "Support Body side plane3", SimplePhysical
    m_oSymbolHelper.AddOutputDef 15, "SupportBodySidePlane4", "Support Body side plane4", SimplePhysical
    m_oSymbolHelper.AddOutputDef 16, "SupportBodyRefPlane1", "Support Body Referance Plane 1", SimplePhysical
    m_oSymbolHelper.AddOutputDef 17, "SupportBodyRefPlane2", "Support Body Referance Plane 2", SimplePhysical
       
    m_oSymbolHelper.AddOutputDef 18, "SupportBox1", "support under the pump discharge nozzle", SimplePhysical
    m_oSymbolHelper.AddOutputDef 19, "SupportBox2", "support under the intermediate chamber", SimplePhysical
    m_oSymbolHelper.AddOutputDef 20, "SupportBox3", "support under the motor", SimplePhysical
    m_oSymbolHelper.AddOutputDef 21, "EqpFoundationPort", "Foundation Port under support", SimplePhysical
    m_oSymbolHelper.AddOutputDef 22, "DrainBend", "Bend of the Drain Nozzle", SimplePhysical
    m_oSymbolHelper.AddOutputDef 23, "DrainBendCyl", "Cylinder of the Drain Nozzle", SimplePhysical
    m_oSymbolHelper.AddOutputDef 24, "PumpControlPoint", "Point at origin", ReferenceGeometry
    
' Aspects
    m_oSymbolHelper.NumAspects = 2
    m_oSymbolHelper.AddAspectDef 1, "SimplePhysical", "Physical", SimplePhysical
    m_oSymbolHelper.AddAspectDef 2, "ReferenceGeometry", "ReferenceGeometry", ReferenceGeometry
    
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






