Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Common.Exceptions

Namespace Symbols
    <SymbolVersion("1.0.0.0")> _
    <OutputNotification("IJDAttributes")> _
    <InputNotification("IJDAttributes", True)> _
    <OutputNotification("IJDGeometry", True)> _
    Public Class CustomAsmWithOnlyMemberOutputs : Inherits CustomAssemblyDefinition
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
        ''' Evaluate assembly outputs
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
            oPropVal = DirectCast(Me.Occurrence.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyDiameter), PropertyValueDouble)
            dblBaseRadius = oPropVal.PropValue().Value / 2.0
            Dim dblBodyRadius As Double
            dblBodyRadius = dblBaseRadius * 2.0 / 3.0
            Dim dblHeadRadius As Double
            dblHeadRadius = dblBaseRadius / 2.0

            ' Define origin of base
            Dim origin As New Position(0.0, 0.0, dblBaseRadius)

            If m_objBase.Output Is Nothing Then
                '=================================================
                ' Construct base
                '=================================================
                m_objBase.Output = New Sphere3d(oConnection, origin, dblBaseRadius, True)
            End If

            ' Define origin of body
            origin.Z = origin.Z + dblBaseRadius + dblBodyRadius

            If m_objBody.Output Is Nothing Then
                '=================================================
                ' Construct body
                '=================================================
                m_objBody.Output = New Sphere3d(oConnection, origin, dblBodyRadius, True)
            End If

            ' Define origin of head
            origin.Z = origin.Z + dblBodyRadius + dblHeadRadius

            If m_objHead.Output Is Nothing Then
                '=================================================
                ' Construct head
                '=================================================
                m_objHead.Output = New Sphere3d(oConnection, origin, dblHeadRadius, True)
            End If

            '===========================================
            ' Verify that we can get access to the
            ' assembly output for a given BO
            '===========================================
            Dim oBO As BusinessObject
            oBO = m_objBody.Output
            Dim oAssemblyOutput As AssemblyOutput
            oAssemblyOutput = GetAssemblyOutput(oBO)
            If Not oAssemblyOutput.Equals(m_objBody) Then
                Throw New CmnException("Expected the assembly output to be the body of the snowman")
            End If


            '===========================================
            ' Verify that we throw the correct exception
            ' if no assembly output for a given BO
            '===========================================
            oAssemblyOutput = Nothing
            Try
                oAssemblyOutput = GetAssemblyOutput(oBusinessObject)
            Catch ex As CustomAssemblyNoAssemblyOutputForBOException
                ' Got what we expected
            End Try
            If Not oAssemblyOutput Is Nothing Then
                Throw New CmnException("Expected AssemblyOutput to be null")
            End If

        End Sub

#Region "Implement property Management methods"
        ''' <summary>
        ''' Preload property pages
        ''' </summary>
        ''' <param name="oBusinessObject"></param>
        ''' <param name="boProperties"></param>
        ''' <remarks></remarks>
        Public Overrides Sub OnPreLoad(ByVal oBusinessObject As BusinessObject, ByVal boProperties As ReadOnlyCollection(Of PropertyDescriptor))

            ' Check if the business object is the CustomAssembly occurrence
            If oBusinessObject.Equals(Me.Occurrence) Then
            Else
                Dim oAssemblyOutput As AssemblyOutput = GetAssemblyOutput(oBusinessObject)

                Dim sAssemblyOutputName As String = oAssemblyOutput.Name
                Select Case sAssemblyOutputName
                    Case CONST_Base
                    Case CONST_Body
                    Case CONST_Head
                    Case Else
                        Throw New CmnException("Unexpected assembly output BO.")
                End Select
            End If
        End Sub

        ''' <summary>
        ''' Method to verify a custom assembly or output changed property
        ''' </summary>
        ''' <param name="oBusinessObject"></param>
        ''' <param name="CollAllDisplayedValues"></param>
        ''' <param name="oPropToChange"></param>
        ''' <param name="oNewPropValue"></param>
        ''' <param name="strErrorMsg"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function OnPropertyChange(ByVal oBusinessObject As BusinessObject, _
                                  ByVal CollAllDisplayedValues As ReadOnlyCollection(Of PropertyDescriptor), _
                                  ByVal oPropToChange As PropertyDescriptor, _
                                  ByVal oNewPropValue As PropertyValue, _
                                  ByRef strErrorMsg As String) As Boolean

            Dim oPropInfo As PropertyInformation = oNewPropValue.PropertyInfo

            ' Check if the business object is the customassembly occurrence
            If oBusinessObject.Equals(Me.Occurrence) Then
            Else
                Dim oAssemblyOutput As AssemblyOutput = GetAssemblyOutput(oBusinessObject)

                Dim sAssemblyOutputName As String = oAssemblyOutput.Name
                Select Case sAssemblyOutputName
                    Case CONST_Base
                    Case CONST_Body
                    Case CONST_Head
                    Case Else
                        Throw New CmnException("Unexpected assembly output BO.")
                End Select
            End If

            Return True
        End Function
#End Region

    End Class
End Namespace

