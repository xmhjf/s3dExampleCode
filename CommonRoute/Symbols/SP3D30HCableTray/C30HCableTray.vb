'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Copyright (C) 2009, Intergraph PPO. All rights reserved.
'
'File
'  C30HCableTray.vb
'
'Abstract
'	This is a 30 deg horizontal Cabletray symbol. This class subclasses from CustomSymbolDefinition.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit On
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Equipment.Middle
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.Common.Exceptions

'-----------------------------------------------------------------------------------
'Namespace of this class is Ingr.SP3D.Content.CableTraySample.
'It is recommended that customers specify namespace of their symbols to be
'<CompanyName>.SP3D.Content.<Specialization>.
'It is also recommended that if customers want to change this symbol to suit their
'requirements, they should change namespace/symbol name so the identity of the modified
'symbol will be different from the one delivered by Intergraph.
'-----------------------------------------------------------------------------------

Public Class C30HCableTray : Inherits CustomSymbolDefinition

    '----------------------------------------------------------------------------------
    'DefinitionName/ProgID of this symbol is "SP3D30HCableTray,Ingr.SP3D.Content.CableTraySample.C30HCableTray"
    '----------------------------------------------------------------------------------


#Region "Definition of Inputs"

    <InputCatalogPart(1)> _
        Public m_oPartInput As InputCatalogPart

#End Region

#Region "Definitions of Aspects and their outputs"

    ' Physical Aspect 
    <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
    <SymbolOutput("HoriTangent", "Horizontal Tangent")> _
    <SymbolOutput("Elbow", "Elbow")> _
    <SymbolOutput("IncTangent", "Inclined Tangent")> _
    <SymbolOutput("TrayPort1", "Tray Port 1")> _
    <SymbolOutput("TrayPort2", "Tray Port 2")> _
        Public m_oPhysicalAspect As AspectDefinition

#End Region

#Region "Construction of outputs of all aspects"

    Protected Overrides Sub ConstructOutputs()        
        Try
            Dim oTrayPart As Part, oConnection As SP3DConnection
            Dim oWarningColl As New Collection(Of SymbolWarningException)

            ' Get Input values
            Try
                oTrayPart = m_oPartInput.Value
                If oTrayPart Is Nothing Then
                    Throw New CmnException("Unable to retrieve tray catalog part.")
                End If
                oConnection = OccurrenceConnection  ' Get the connection where outputs will be created.
                If oConnection Is Nothing Then
                    Throw New CmnException("Unable to retrieve connection.")
                End If
            Catch oEx As Exception
                If (TypeOf oEx Is SymbolWarningException) Then
                    oWarningColl.Add(oEx)
                Else
                    Throw
                End If
            End Try
            '=================================================
            ' Construction of Physical Aspect
            '=================================================

            Dim oPropertyValve As PropertyValueDouble, dBendRadius As Double
            Dim dInsertionDepth As Double, dTangentLength As Double
            Dim dActualWidth As Double, dActualDepth As Double
            Dim oTrayPort1 As CableTrayPortDef, oTrayPort2 As CableTrayPortDef
            Dim oPropertyDouble As PropertyValueDouble

            oPropertyValve = oTrayPart.GetPropertyValue("IJCableTrayPart", "BendRadius")
            dBendRadius = oPropertyValve.PropValue
            oPropertyValve = oTrayPart.GetPropertyValue("IJCableTrayPart", "TangentLength")
            dTangentLength = oPropertyValve.PropValue
            oPropertyValve = oTrayPart.GetPropertyValue("IJCableTrayPart", "InsertionDepth")
            dInsertionDepth = oPropertyValve.PropValue

            oTrayPort1 = oTrayPart.PortDefinitions.Item(0)
            oTrayPort2 = oTrayPart.PortDefinitions.Item(1)

            oPropertyDouble = oTrayPort1.GetPropertyValue("IJCableTrayPort", "ActualWidth")
            dActualWidth = oPropertyDouble.PropValue
            oPropertyDouble = oTrayPort1.GetPropertyValue("IJCableTrayPort", "ActualDepth")
            dActualDepth = oPropertyDouble.PropValue

            'variable for relocating the port considering insertion depth.
            'Check to validate that if the tangentLength is zero, set it to a very small value
            If dTangentLength = 0 Then dTangentLength = 0.000001

            ' Insert your code for output 1(U Shape Horizontal Tangent)
            Dim dFacetoCenter As Double
            Dim dHorDepth As Double, dHorWidth As Double, dAngle As Double
            Dim oPointsColl As New Collection(Of Position)

            dAngle = Math.PI / 6
            dHorDepth = dActualDepth / 2
            dHorWidth = dActualWidth / 2
            dFacetoCenter = (dBendRadius + dHorWidth) * Tan(dAngle / 2) + dTangentLength
            Dim oPort1 As New Position(-dFacetoCenter, 0, 0)

            ' Initialize SymbolGeometryhelper, set the active position and the orientation. 
            ' Primary as Global X and Secondary as Global Z.  
            Dim oSymbolGeomHlpr As New SymbolGeometryHelper()
            oSymbolGeomHlpr.ActivePosition = oPort1
            oSymbolGeomHlpr.SetOrientation(New Vector(1, 0, 0), New Vector(0, 0, 1))

            oPointsColl.Add(New Position(dHorWidth, dHorDepth, 0))
            oPointsColl.Add(New Position(dHorWidth, -dHorDepth, 0))
            oPointsColl.Add(New Position(-dHorWidth, -dHorDepth, 0))
            oPointsColl.Add(New Position(-dHorWidth, dHorDepth, 0))
            Dim oLineStr3d As New LineString3d(oPointsColl)
            Dim oHoriTangent = oSymbolGeomHlpr.CreateProjectedPolygonFromCrossSection(oConnection, oLineStr3d, dTangentLength, False)
            ' Set the output
            m_oPhysicalAspect.Outputs("HoriTangent") = oHoriTangent
            oPointsColl.Clear() ' Clearing oPointsColl

            ' Insert your code for output 2(U Shape Elbow)
            oSymbolGeomHlpr.ActivePosition = New Position(-dFacetoCenter + dTangentLength, dBendRadius + dHorWidth, 0)
            oSymbolGeomHlpr.SetOrientation(New Vector(0, 0, 1), New Vector(1, 0, 0))
            Dim LineStrCP As New Position(-dFacetoCenter + dTangentLength, 0, 0)

            oPointsColl.Add(New Position(LineStrCP.X, LineStrCP.Y + dHorWidth, LineStrCP.Z + dHorDepth))
            oPointsColl.Add(New Position(LineStrCP.X, LineStrCP.Y + dHorWidth, LineStrCP.Z - dHorDepth))
            oPointsColl.Add(New Position(LineStrCP.X, LineStrCP.Y - dHorWidth, LineStrCP.Z - dHorDepth))
            oPointsColl.Add(New Position(LineStrCP.X, LineStrCP.Y - dHorWidth, LineStrCP.Z + dHorDepth))
            Dim oElbowLineStr = New LineString3d(oPointsColl)
            Dim oElbow = oSymbolGeomHlpr.CreateSurfaceofRevolution(oConnection, oElbowLineStr, dAngle)
            ' Set the output
            m_oPhysicalAspect.Outputs("Elbow") = oElbow
            oPointsColl.Clear() ' Clearing oPointsColl

            ' Insert your code for output 3(U Shape Inclined Tangent)
            Dim oPort2 As New Position(dFacetoCenter * Cos(dAngle), dFacetoCenter * Sin(dAngle), 0)
            oSymbolGeomHlpr.ActivePosition = oPort2
            oSymbolGeomHlpr.SetOrientation(New Vector(-Cos(dAngle), -Sin(dAngle), 0), New Vector(0, 0, 1))
            Dim oInclTangent = oSymbolGeomHlpr.CreateProjectedPolygonFromCrossSection(oConnection, oLineStr3d, dTangentLength, False)
            ' Set the output
            m_oPhysicalAspect.Outputs("IncTangent") = oInclTangent

            ' Place Nozzle 1
            Dim oDir As New Vector(-1, 0, 0)
            Dim oPortLocation As New Position(oPort1.X - dInsertionDepth * oDir.X, oPort1.Y - dInsertionDepth * oDir.Y, oPort1.Z - dInsertionDepth * oDir.Z)
            Dim oCableTrayPort = New CableTrayPort(oTrayPart, oConnection, 1, oPortLocation, oDir)
            oCableTrayPort.RadialVector = New Vector(0, 0, 1)
            ' Set the output
            m_oPhysicalAspect.Outputs("TrayPort1") = oCableTrayPort

            ' Place Nozzle 2
            oDir.Set(Cos(dAngle), Sin(dAngle), 0)
            oPortLocation.Set(oPort2.X - dInsertionDepth * oDir.X, oPort2.Y - dInsertionDepth * oDir.Y, oPort2.Z - dInsertionDepth * oDir.Z)
            oCableTrayPort = New CableTrayPort(oTrayPart, oConnection, 2, oPortLocation, oDir)
            oCableTrayPort.RadialVector = New Vector(0, 0, 1)
            ' Set the output
            m_oPhysicalAspect.Outputs("TrayPort2") = oCableTrayPort

            If (oWarningColl.Count > 0) Then
                Throw oWarningColl.Item(0)
            End If

        Catch Ex As Exception ' General Unhandled exception 
            Throw
        End Try

    End Sub
#End Region

End Class
