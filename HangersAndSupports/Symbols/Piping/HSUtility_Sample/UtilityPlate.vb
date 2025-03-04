''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'Copyright (C) 2010, Intergraph PPO. All rights reserved.
'
'File
'  cUtility_Plate.vb
'
'Abstract
'	This is a sample Utility Plate Symbol.
'
'23-03-2015     Chethan TR-CP-268570  Namespace inconsistency in .NET content for few H&S project  
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit On
Imports System.Math
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.ReferenceData.Middle.Services
Imports Ingr.SP3D.ReferenceData.Middle
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Exceptions
Imports Ingr.SP3D.Structure.Middle.Services
Imports Ingr.SP3D.Support.Middle
Imports Ingr.SP3D.Content.Support.Symbols
'-----------------------------------------------------------------------------------
'Namespace of this class is Ingr.SP3D.Support.Content.Symbols
'It is recommended that customers specify namespace of their symbols to be
'<CompanyName>.SP3D.Content.<Specialization>.
'It is also recommended that if customers want to change this symbol to suit their
'requirements, they should change namespace/symbol name so the identity of the modified
'symbol will be different from the one delivered by Intergraph.
'-----------------------------------------------------------------------------------
Public Class UtilityPlate : Inherits CustomSymbolDefinition
    Implements ICustomHgrBOMDescription
    '----------------------------------------------------------------------------------
    'DefinitionName/ProgID of this symbol is "HSNET_Utility,Ingr.SP3D.Support.Content.cUtility_Plate"
    '----------------------------------------------------------------------------------
    Private Const DENSITY As Double = 7849  'kg/m^3


#Region "Definition of Inputs"

    <InputCatalogPart(1)> Public m_oPartInput As InputCatalogPart
    <InputDouble(2, "THICKNESS", "Thickness", 0)> Public m_dThickness As InputDouble
    <InputDouble(3, "WIDTH", "Width", 0)> Public m_dWidth As InputDouble
    <InputDouble(4, "DEPTH", "Depth", 0)> Public m_dDepth As InputDouble
#End Region

#Region "Definitions of Aspects and their outputs"

    'Physical Aspect
    <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
    <SymbolOutput("BODY", "BODY")> _
    <SymbolOutput("LINE", "LINE")> _
    <SymbolOutput("Port1", "Port1")> _
    <SymbolOutput("Port2", "Port2")> _
    Public m_oPhysicalAspect As AspectDefinition

#End Region

#Region "Construction of outputs of all aspects"

    Protected Overrides Sub ConstructOutputs()

        Try
            Dim oPart As Part = Nothing, oConnection As SP3DConnection
            Dim dThickness As Double, dWidth As Double, dDepth As Double
            Dim oWarningColl As New Collection(Of SymbolWarningException)

            ' Get Input values
            Try
                oPart = m_oPartInput.Value
                dThickness = m_dThickness.Value
                dWidth = m_dWidth.Value
                dDepth = m_dDepth.Value
            Catch oEx As Exception
                If (TypeOf oEx Is SymbolWarningException) Then
                    oWarningColl.Add(oEx)
                Else
                    Throw
                End If
            End Try

            oConnection = OccurrenceConnection ' Get the connection where outputs will be created.

            '=================================================
            ' Construction of Physical Aspect 
            '=================================================
            'Create Ports
            Dim oPort1 = New Port(oConnection, oPart, "TopStructure", New Position(0, 0, 0), New Vector(1, 0, 0), New Vector(0, 0, 1))
            m_oPhysicalAspect.Outputs("Port1") = oPort1
            Dim oPort2 = New Port(oConnection, oPart, "BotStructure", New Position(0, 0, dThickness), New Vector(1, 0, 0), New Vector(0, 0, 1))
            m_oPhysicalAspect.Outputs("Port2") = oPort2

            'Initialize SymbolGeometryHelper. Set the active position and orientation 
            Dim oSymbolGeomHlpr As New SymbolGeometryHelper()
            oSymbolGeomHlpr.ActivePosition = New Position(0, 0, 0)
            oSymbolGeomHlpr.SetOrientation(New Vector(0, 0, 1), New Vector(1, 0, 0))

            Dim oPlate = oSymbolGeomHlpr.CreateBox(oConnection, dThickness, dDepth, dWidth)
            Dim oLinePoints As New List(Of Position)
            oLinePoints.Add(New Position(-dDepth / 2, 0, dThickness / 2))
            oLinePoints.Add(New Position(dDepth / 2, 0, dThickness / 2))
            Dim oLine = oSymbolGeomHlpr.CreateLineString(oConnection, oLinePoints, False)

            m_oPhysicalAspect.Outputs("BODY") = oPlate
            m_oPhysicalAspect.Outputs("LINE") = oLine

            'Dim oHgrHelper As New HSHelper
            'If dWidth > 0.5 And dWidth < 1 Then
            '    oHgrHelper.Notify(Err, "Width is between 0.5 and 1", "ConstructOutputs", "CUtility_Plate", HSHelper.WarnOrError.WarnOnly)
            'End If

            'If dWidth > 1 Then
            '    oHgrHelper.Notify(Err, "Width greater than 1", "ConstructOutputs", "CUtility_Plate", HSHelper.WarnOrError.ErrorOnly)
            'End If

        Catch Ex As Exception 'General Unhandled exception 
            Throw
        End Try

    End Sub

#End Region

    Public Function BOMDescription(ByVal oSupportOrComponent As Common.Middle.BusinessObject) As String Implements ICustomHgrBOMDescription.BOMDescription

        Dim oSupportComp As SupportComponent = oSupportOrComponent
        Dim oPart = oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects(0)

        Dim THICKNESS As Double = DirectCast(oPart.GetPropertyValue("IJUAHgrUtility_PLATE", "THICKNESS"), PropertyValueDouble).PropValue
        Dim WIDTH As Double = DirectCast(oSupportComp.GetPropertyValue("IJOAHgrUtility_PLATE", "WIDTH"), PropertyValueDouble).PropValue
        Dim DEPTH As Double = DirectCast(oSupportComp.GetPropertyValue("IJOAHgrUtility_PLATE", "DEPTH"), PropertyValueDouble).PropValue

        Dim strDepth As String = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, DEPTH, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.UNIT_NOT_SET)
        Dim strWidth As String = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, WIDTH, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.UNIT_NOT_SET)
        Dim strTHICKNESS As String = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, THICKNESS, UnitName.DISTANCE_FOOT, UnitName.DISTANCE_INCH, UnitName.UNIT_NOT_SET)


        Return "From Symbol:  " & strTHICKNESS & " Plate Steel, " & strWidth & " X " & strDepth

    End Function
End Class