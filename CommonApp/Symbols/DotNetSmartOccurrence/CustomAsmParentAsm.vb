Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports System.Runtime.InteropServices
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports SP3DPIA.SmartOccurrence
Imports Ingr.SP3D.Common.Middle.Services.Hidden
Imports OLEENGINELib
Imports AutoMath

Namespace Symbols
    <CacheOption(CacheOptionType.NonCached)> _
    <SymbolVersion("1.0.0.0")> _
    <OutputNotification("IJDAttributes")> _
    <OutputNotification("IJDGeometry")> _
    <OutputNotification("IJDAttributes", True)> _
    <OutputNotification("IJDGeometry", True)> _
    Public Class CustomAsmParentAsm : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "Diameter"
        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
        Public m_dblDiameter As InputDouble

        Private Const CONST_MediumSnowmanHeadSmartItem As String = "Medium_SnowmanHead"
        Private Const CONST_MediumSnowmanBodySmartItem As String = "Medium_SnowmanBody"
        Private Const CONST_Base As String = "Base"
        Private Const CONST_Body As String = "Body"
        Private Const CONST_Head As String = "Head"

        Private Const CONST_IJUATestDotNetSphere As String = "IJUATestDotNetSphere"
        Private Const CONST_PropertyDiameter As String = "Diameter"
        Private Const CONST_PropertyOriginX As String = "OriginX"
        Private Const CONST_PropertyOriginY As String = "OriginY"
        Private Const CONST_PropertyOriginZ As String = "OriginZ"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base of snowman")> _
        Public m_physicalAspect As AspectDefinition

        <AssemblyOutput(1, CONST_Body)> _
        Public m_objBody As AssemblyOutput

        <AssemblyOutput(2, CONST_Head)> _
        Public m_objHead As AssemblyOutput

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            ' Get the connection
            Dim oSP3DConnection As SP3DConnection
            oSP3DConnection = Me.OccurrenceConnection

            Dim dblBaseRadius As Double
            dblBaseRadius = m_dblDiameter.Value / 2.0

            '=================================================
            ' Construct base
            '=================================================
            Dim objBase As Sphere3d

            objBase = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, dblBaseRadius), dblBaseRadius, True)

            m_physicalAspect.Outputs(CONST_Base) = objBase
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

            Dim oPropVal As PropertyValueDouble
            Dim oVector As New Vector(0.0, 0.0, 0.0)
            oPropVal = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyOriginX), PropertyValueDouble)
            If oPropVal.PropValue().HasValue Then
                oVector.X = oPropVal.PropValue().Value
            End If
            oPropVal = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyOriginY), PropertyValueDouble)
            If oPropVal.PropValue().HasValue Then
                oVector.Y = oPropVal.PropValue().Value
            End If
            oPropVal = DirectCast(oBusinessObject.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyOriginZ), PropertyValueDouble)
            If oPropVal.PropValue().HasValue Then
                oVector.Z = oPropVal.PropValue().Value
            End If

            ' Test updating the diameter property in the Evaluate
            Dim dblBaseRadius As Double
            dblBaseRadius = m_dblDiameter.Value / 2.0


            '=================================================
            ' Construct body
            '=================================================
            If m_objBody.Output Is Nothing Then
                m_objBody.Output = CreateSubComponent(oConnection, CONST_MediumSnowmanBodySmartItem)
            End If

            '********************************
            ' Update body
            '********************************
            Dim oBody As BusinessObject
            oBody = m_objBody.Output
            Dim dblBodyRadius As Double
            dblBodyRadius = 0.0
            oPropVal = DirectCast(oBody.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyDiameter), PropertyValueDouble)
            If oPropVal.PropValue().HasValue Then
                dblBodyRadius = oPropVal.PropValue().Value / 2.0
            End If
            oVector.Z = oVector.Z + dblBaseRadius * 2.0 + dblBodyRadius
            Translate(oBody, oVector)

            '=================================================
            ' Construct head
            '=================================================
            If m_objHead.Output Is Nothing Then
                m_objHead.Output = CreateSubComponent(oConnection, CONST_MediumSnowmanHeadSmartItem)
            End If

            '********************************
            ' Update head
            '********************************
            Dim oHead As BusinessObject
            oHead = m_objHead.Output
            Dim dblHeadRadius As Double
            dblHeadRadius = 0.0
            oPropVal = DirectCast(oHead.GetPropertyValue(CONST_IJUATestDotNetSphere, CONST_PropertyDiameter), PropertyValueDouble)
            If oPropVal.PropValue().HasValue Then
                dblHeadRadius = oPropVal.PropValue().Value / 2.0
            End If
            oVector.Z = oVector.Z + dblBodyRadius + dblHeadRadius
            Translate(oHead, oVector)
        End Sub

        Private Function CreateSubComponent(ByRef oConnection As SP3DConnection, ByVal rootSelection As String) As BusinessObject
            Dim oSmartOccBO As BusinessObject = Nothing
            Dim oFactory As IJSmartOccurrenceFactory
            oFactory = New SmartOccurrenceFactory()

            If Not oFactory Is Nothing And Not oConnection Is Nothing Then
                Dim oResourceMgr As Object
                oResourceMgr = MiddleUtilities.GetResourceManagerFromSP3DConnection(oConnection)

                Dim oSmartOcc As IJSmartOccurrence = oFactory.CreateSmartOccurrence(oResourceMgr)
                If Not oSmartOcc Is Nothing Then
                    ' Initialize the smart occurrence
                    oSmartOcc.Class = rootSelection
                    oSmartOcc.RootSelectorClass = rootSelection
                    oSmartOcc.Update()
                    oSmartOccBO = COMConverters.ConvertCOMBOToBO(oSmartOcc)
                End If
            End If

            Marshal.ReleaseComObject(oFactory)

            Return oSmartOccBO
        End Function

        Private Sub Translate(ByRef oObject As BusinessObject, ByVal translationVector As Vector)
            If Not oObject Is Nothing Then
                Dim oIJDSymbolOccurrence As IJDOccurrence
                oIJDSymbolOccurrence = DirectCast(COMConverters.ConvertBOToCOMBO(oObject), IJDOccurrence)
                Dim oVec As OLEENGINELib.IJDVector
                oVec = DirectCast(New AutoMath.DVector(), OLEENGINELib.IJDVector)
                oVec.Set(translationVector.X, translationVector.Y, translationVector.Z)
                Dim oMtx As OLEENGINELib.IJDT4x4
                oMtx = oIJDSymbolOccurrence.Matrix
                oMtx.IndexValue(12) = translationVector.X
                oMtx.IndexValue(13) = translationVector.Y
                oMtx.IndexValue(14) = translationVector.Z
                oIJDSymbolOccurrence.Matrix = oMtx
            End If
        End Sub

    End Class
End Namespace

