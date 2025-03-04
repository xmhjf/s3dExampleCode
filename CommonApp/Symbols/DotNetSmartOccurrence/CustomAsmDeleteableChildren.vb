Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services

Namespace Symbols
    <SymbolVersion("1.0.0.0")> _
    <OutputNotification("IJDAttributes")> _
    <InputNotification("IJDAttributes", True)> _
    <OutputNotification("IJDGeometry", True)> _
    Public Class CustomAsmDeleteableChildren : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "Diameter"
        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart

        Private Const CONST_Base As String = "Base"
        Private Const CONST_Body As String = "Body"
        Private Const CONST_Head As String = "Head"

        Private Const CONST_IJUATestDotNetSphere As String = "IJUATestDotNetSphere"
        Private Const CONST_PropertyDiameter As String = "Diameter"

        <AssemblyOutput(1, CONST_Base)> _
        Public m_objBase As AssemblyOutput

        <AssemblyOutput(2, CONST_Body)> _
        Public m_objBody As AssemblyOutput

        <AssemblyOutput(3, CONST_Head)> _
        Public m_objHead As AssemblyOutput

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs (construct them if they
        '''  do not exist / update them if they do exist)
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            Dim oBusinessObject As BusinessObject
            oBusinessObject = Me.Occurrence
            If oBusinessObject Is Nothing Then
                Throw New CmnException("Occurrence property is null!")
            End If

            Dim oConnection As SP3DConnection
            oConnection = oBusinessObject.DBConnection

            Dim dblBaseRadius As Double
            Dim oPropVal As PropertyValueDouble
            oPropVal = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyDiameter), PropertyValueDouble)
            dblBaseRadius = oPropVal.PropValue().Value / 2.0
            Dim dblBodyRadius As Double
            dblBodyRadius = dblBaseRadius * 2.0 / 3.0
            Dim dblHeadRadius As Double
            dblHeadRadius = dblBaseRadius / 2.0

            ' Allow the head to be removed by the user
            m_objHead.CanDeleteIndependently = True

            ' Set the position of the base
            Dim origin As New Position(0.0, 0.0, dblBaseRadius)

            '*************************************************************
            ' Construct the base if it doesn't exist otherwise update it
            '*************************************************************
            If m_objBase.Output Is Nothing Then
                m_objBase.Output = New Sphere3d(oConnection, origin, dblBaseRadius, True)
            Else
                '*******************
                ' Update base
                '*******************
            End If

            ' Set the position of the body
            origin.Z = origin.Z + dblBaseRadius + dblBodyRadius

            '*************************************************************
            ' Construct the body if it doesn't exist otherwise update it
            '*************************************************************
            If m_objBody.Output Is Nothing Then
                '**************************************
                ' Construct body if not removed by user
                '**************************************
                If Not m_objBody.HasBeenDeletedByUser Then
                    m_objBody.Output = New Sphere3d(oConnection, origin, dblBodyRadius, True)
                Else
                    ' Forcibly allow the output to be regenerated. We won't re-create the output
                    '   during this evaluate but will wait for the next.
                    m_objBody.HasBeenDeletedByUser = False
                End If

            Else
                '*******************
                ' Update body
                '*******************
            End If

            ' Set the position of the head
            origin.Z = origin.Z + dblBodyRadius + dblHeadRadius

            '*************************************************************
            ' Construct the head if head doesn't exist otherwise update it
            '*************************************************************
            If m_objHead.Output Is Nothing Then
                '**************************************
                ' Construct head if not removed by user
                '**************************************
                If Not m_objHead.HasBeenDeletedByUser Then
                    m_objHead.Output = New Sphere3d(oConnection, origin, dblHeadRadius, True)
                End If
            Else
                '*******************
                ' Update head
                '*******************
            End If

            ' Once the head is removed then allow the body to be removed
            If m_objHead.Output Is Nothing Then
                m_objBody.CanDeleteIndependently = True
            End If
        End Sub
    End Class
End Namespace

