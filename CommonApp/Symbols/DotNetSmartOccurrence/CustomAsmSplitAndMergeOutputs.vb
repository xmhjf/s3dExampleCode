Option Strict On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Common.Exceptions

Namespace Symbols
    <SymbolVersion("1.0.0.0")> _
    <InputNotification("IJDAttributes", True)> _
    <OutputNotification("IJDGeometry", True)> _
    Public Class CustomAsmSplitAndMergeOutputs : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "PointCount"
        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart

        Private Const CONST_Points As String = "Points"
        Private Const CONST_MorePoints As String = "MorePoints"

        Private Const CONST_IJUATestDotNetPointMigrate As String = "IJUATestDotNetPointMigrate"
        Private Const CONST_PropertyPointCount As String = "PointCount"

        <AssemblyOutput(1, CONST_Points)> _
        Public m_objPoints As AssemblyOutputs

        <AssemblyOutput(2, CONST_MorePoints)> _
        Public m_objMorePoints As AssemblyOutputs

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

            Dim oPropVal As PropertyValueInt
            oPropVal = DirectCast(Me.Occurrence.GetPropertyValue(CONST_IJUATestDotNetPointMigrate, CONST_PropertyPointCount), PropertyValueInt)
            Dim iPointCount As Integer = oPropVal.PropValue().Value
            Dim bDeletedInstead As Boolean = False
            If iPointCount < 0 Then
                iPointCount = Abs(iPointCount)
                bDeletedInstead = True
            End If

            If m_objPoints.Count = iPointCount Then ' Count has not changed so do nothing
            ElseIf m_objPoints.Count < iPointCount Then ' Need to split and add some points
                Dim startZ As Double = 0.0
                Dim increment As Double = 12.0 / Max(iPointCount, 1)
                If m_objPoints.Count = 0 Then ' Create for the first time so no migrate
                    For iCnt As Integer = 1 To iPointCount
                        m_objPoints.Add(New Point3d(oConnection, New Position(0, 0, startZ)))
                        startZ = startZ + increment
                    Next iCnt
                Else ' Adding additional points
                    Dim replacingObjects As Collection(Of BusinessObject) = New Collection(Of BusinessObject)()
                    ' Remove the orignal point or add a new one
                    If Not bDeletedInstead Then
                        Dim replacedObjects As Collection(Of BusinessObject) = New Collection(Of BusinessObject)()
                        ' First update the existing points
                        For iCnt As Integer = 1 To m_objPoints.Count
                            Dim oPointBO As BusinessObject = m_objPoints(iCnt - 1)
                            Dim oPoint As Point3d = DirectCast(oPointBO, Point3d)
                            oPoint.Position = New Position(0, 0, startZ)
                            replacedObjects.Add(oPoint)
                            startZ = startZ + increment
                        Next iCnt

                        ' Copy over the replaced collection to the replacing collection
                        For Each oRepObject As BusinessObject In replacedObjects
                            replacingObjects.Add(oRepObject)
                        Next oRepObject

                        ' Add the new points
                        For iCnt As Integer = m_objPoints.Count + 1 To iPointCount
                            Dim oNewPoint3d As Point3d = New Point3d(oConnection, New Position(0.0, 0.0, startZ))
                            m_objPoints.Add(oNewPoint3d)
                            replacingObjects.Add(oNewPoint3d)
                            startZ = startZ + increment
                        Next iCnt
                        ' Replace the original with itself and a set of new ones
                        ReplaceObjects(replacedObjects, replacingObjects)
                    Else
                        Dim deletedObjects As Collection(Of BusinessObject) = New Collection(Of BusinessObject)()
                        ' First update the existing points
                        For iCnt As Integer = m_objPoints.Count To 1 Step -1
                            Dim oPtBO As BusinessObject = m_objPoints(iCnt - 1)
                            deletedObjects.Add(oPtBO)
                            m_objPoints.RemoveAt(iCnt - 1)
                        Next iCnt

                        ' Add the new points
                        For iCnt As Integer = 1 To iPointCount
                            Dim oNewPt3d As Point3d = New Point3d(oConnection, New Position(0.0, 0.0, startZ))
                            m_objPoints.Add(oNewPt3d)
                            replacingObjects.Add(oNewPt3d)
                            startZ = startZ + increment
                        Next iCnt

                        Try
                            ' Delete the original objects and replace it with new ones
                            DeleteObjects(deletedObjects, replacingObjects)
                        Catch ex As CustomAssemblyInvalidPropertyDescriptionIndexException
                            ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "SmartOccurrenceATP", 987, "Attempted to migrate many to many")
                        End Try
                    End If
                End If
            Else ' Remove some points
                Dim objectsToBeDeleted As Collection(Of BusinessObject) = New Collection(Of BusinessObject)()

                For iCnt As Integer = m_objPoints.Count To iPointCount + 1 Step -1
                    Dim oPtBO As BusinessObject = m_objPoints(iCnt - 1)
                    objectsToBeDeleted.Add(oPtBO)
                    m_objPoints.RemoveAt(iCnt - 1)
                Next iCnt

                Dim startZ As Double = 0.0
                Dim increment As Double = 12.0 / Max(iPointCount, 1)
                Dim endZ As Double = increment
                Dim replacingObjects As Collection(Of BusinessObject) = New Collection(Of BusinessObject)()
                ' Update the remaining points
                For iCnt As Integer = 1 To m_objPoints.Count
                    Dim oPointBO As BusinessObject = m_objPoints(iCnt - 1)
                    Dim oPoint As Point3d = DirectCast(oPointBO, Point3d)
                    oPoint.Position = New Position(0, 0, startZ)
                    replacingObjects.Add(oPoint)
                    startZ = startZ + increment
                Next iCnt

                ' Deleting some of the original objects - will map each deleted
                '   object to the set of replacingObjects
                Dim deletedObjects As Collection(Of BusinessObject) = New Collection(Of BusinessObject)()
                For Each objBusiness As BusinessObject In objectsToBeDeleted
                    deletedObjects.Clear()
                    deletedObjects.Add(objBusiness)
                    DeleteObjects(deletedObjects, replacingObjects)
                Next objBusiness
            End If

            ' Create another set of points
            '    for every multiple of four
            Dim iModulo4 As Integer = Convert.ToInt32((iPointCount - 1) / 4)
            If iModulo4 > 0 Then
                ' If the point count is greater than the number needed then we need to remove some
                If m_objMorePoints.Count > iModulo4 Then
                    Dim deletedObjects2 As Collection(Of BusinessObject) = New Collection(Of BusinessObject)()
                    ' Delete some of the points in the second assembly output
                    For iRemoveCnt As Integer = m_objMorePoints.Count To iModulo4 + 1 Step -1
                        Dim objBO As BusinessObject = m_objMorePoints(iRemoveCnt - 1)
                        deletedObjects2.Add(objBO)
                        m_objMorePoints.RemoveAt(iRemoveCnt - 1)
                    Next iRemoveCnt
                    Dim replacingObjects2 As Collection(Of BusinessObject) = New Collection(Of BusinessObject)()
                    For iRemainingCnt As Integer = m_objMorePoints.Count To 1 Step -1
                        Dim objBO As BusinessObject = m_objMorePoints(iRemainingCnt - 1)
                        replacingObjects2.Add(objBO)
                    Next iRemainingCnt

                    DeleteObjects(deletedObjects2, replacingObjects2)
                Else
                    Dim startZ As Double = 0.0
                    For iAddCnt As Integer = m_objMorePoints.Count + 1 To iModulo4
                        m_objMorePoints.Add(New Point3d(oConnection, 1.0, 0.0, startZ))
                        startZ = iAddCnt * 10.0
                    Next iAddCnt
                End If
            ElseIf m_objMorePoints.Count > 0 Then
                ' Delete all the points in the second assembly output
                For iRemoveCnt As Integer = m_objMorePoints.Count - 1 To 0 Step -1
                    m_objMorePoints.RemoveAt(iRemoveCnt)
                Next iRemoveCnt
            End If
        End Sub

    End Class

    <SymbolVersion("1.0.0.0")> _
    <InputNotification("IJDAttributes", True)> _
    <OutputNotification("IJDGeometry", True)> _
    Public Class CustomAsmMigrator : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart

        Private Const CONST_Sphere As String = "Sphere"


        <AssemblyOutput(1, CONST_Sphere)> _
        Public m_objSphere As AssemblyOutput

        ''' <summary>
        ''' Construct assembly outputs
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

            ' Location of the sphere is dependant on the input point
            Dim oPoint As Point3d = PointInput
            Dim oPos As Position = oPoint.Position

            '=================================================
            ' Construct sphere if it does not alreasdy exist
            '=================================================
            If m_objSphere.Output = Nothing Then
                m_objSphere.Output = New Sphere3d(oConnection, oPos, 0.1, True)
            Else
                Dim objSphere As Sphere3d = DirectCast(m_objSphere.Output, Sphere3d)
                objSphere.Center = oPos
            End If

        End Sub

        Public Overrides Sub InputsReplaced()
            MyBase.InputsReplaced()

            Dim oPoint As Point3d = PointInput

            Dim bWasInputReplaced As ReplacementStatus = GetInputReplacementStatus(oPoint)
            If bWasInputReplaced = ReplacementStatus.Replaced Or _
               bWasInputReplaced = ReplacementStatus.Deleted Then
                ' Make the point with the largest Z the new input to our symbol
                Dim oNewPoint As Point3d = Nothing
                Dim dMaxZ As Double = -1.0
                Dim oReplacedObjects As ReadOnlyCollection(Of BusinessObject) = GetReplacingObjects(oPoint)
                For Each oObject As BusinessObject In oReplacedObjects
                    Dim oNextPoint As Point3d = DirectCast(oObject, Point3d)
                    Dim oPos As Position = oNextPoint.Position
                    If dMaxZ < oPos.Z Then
                        oNewPoint = oNextPoint
                        dMaxZ = oPos.Z
                    End If
                Next oObject

                If Not oNewPoint Is Nothing Then
                    PointInput = oNewPoint
                End If
            End If

        End Sub

        Private Property PointInput() As Point3d
            Get
                Dim oPoint3d As Point3d = Nothing
                Dim oComOccurrence As Object = Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO(Occurrence)
                Dim oSO As SP3DPIA.SmartOccurrence.IJSmartOccurrence = DirectCast(oComOccurrence, SP3DPIA.SmartOccurrence.IJSmartOccurrence)
                Dim oInputHelper As SP3DPIA.SmartOccurrence.IJSmartOccurrenceInputHelper = DirectCast(oSO, SP3DPIA.SmartOccurrence.IJSmartOccurrenceInputHelper)

                Dim oComObject As Object = oInputHelper.GetObjectInput(1)
                Dim oBO As BusinessObject = Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertCOMBOToBO(oComObject)
                oPoint3d = DirectCast(oBO, Point3d)

                Return oPoint3d
            End Get

            Set(ByVal value As Point3d)
                Dim oPointCOM As Object = Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO(value)
                Dim oComOccurrence As Object = Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO(Occurrence)
                Dim oInputHelper As SP3DPIA.SmartOccurrence.IJSmartOccurrenceInputHelper = DirectCast(oComOccurrence, SP3DPIA.SmartOccurrence.IJSmartOccurrenceInputHelper)

                oInputHelper.RemoveAllInputs()
                oInputHelper.AddObjectInputEx(1, "IJPoint", "CtrlPt1_RefColl_DEST", "", oPointCOM)
            End Set
        End Property

    End Class

    <SymbolVersion("1.0.0.0")> _
    <CacheOption(CacheOptionType.NonCached)> _
    Public Class CustomAsmSymbolOnlyMigrator : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        <InputCatalogPart(1)> _
            Public m_catalogPart As InputCatalogPart

        Private Const CONST_Sphere As String = "Sphere"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Sphere, CONST_Sphere)> _
        Public m_physicalAspect As AspectDefinition

        ''' <summary>
        ''' Construct symbol output
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            Dim oBusinessObject As BusinessObject
            oBusinessObject = Me.Occurrence
            If oBusinessObject Is Nothing Then
                Throw New CmnException("Occurrence property is null!")
            End If

            Dim oConnection As SP3DConnection
            oConnection = oBusinessObject.DBConnection

            ' Location of the sphere is dependant on the point position
            Dim oPoint As Point3d = PointInput
            Dim oPos As Position = oPoint.Position

            '=================================================
            ' Construct sphere
            '=================================================
            Dim objSphere As Sphere3d = New Sphere3d(oConnection, oPos, 0.1, True)

            m_physicalAspect.Outputs(CONST_Sphere) = objSphere

        End Sub


        Public Overrides Sub InputsReplaced()
            MyBase.InputsReplaced()

            Dim oPoint As Point3d = PointInput

            Dim bWasInputReplaced As ReplacementStatus = GetInputReplacementStatus(oPoint)
            If bWasInputReplaced = ReplacementStatus.Replaced Or _
               bWasInputReplaced = ReplacementStatus.Deleted Then
                ' Make the point with the largest Z the new input to our symbol
                Dim oNewPoint As Point3d = Nothing
                Dim dMaxZ As Double = -1.0
                Dim oReplacedObjects As ReadOnlyCollection(Of BusinessObject) = GetReplacingObjects(oPoint)
                For Each oObject As BusinessObject In oReplacedObjects
                    Dim oNextPoint As Point3d = DirectCast(oObject, Point3d)
                    Dim oPos As Position = oNextPoint.Position
                    If dMaxZ < oPos.Z Then
                        oNewPoint = oNextPoint
                        dMaxZ = oPos.Z
                    End If
                Next oObject

                If Not oNewPoint Is Nothing Then
                    PointInput = oNewPoint
                End If
            End If

        End Sub

        Private Property PointInput() As Point3d
            Get
                Dim oPoint3d As Point3d = Nothing
                Dim oComOccurrence As Object = Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO(Occurrence)
                Dim oSO As SP3DPIA.SmartOccurrence.IJSmartOccurrence = DirectCast(oComOccurrence, SP3DPIA.SmartOccurrence.IJSmartOccurrence)
                Dim oInputHelper As SP3DPIA.SmartOccurrence.IJSmartOccurrenceInputHelper = DirectCast(oSO, SP3DPIA.SmartOccurrence.IJSmartOccurrenceInputHelper)

                Dim oComObject As Object = oInputHelper.GetObjectInput(1)
                Dim oBO As BusinessObject = Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertCOMBOToBO(oComObject)
                oPoint3d = DirectCast(oBO, Point3d)

                Return oPoint3d
            End Get

            Set(ByVal value As Point3d)
                Dim oPointCOM As Object = Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO(value)
                Dim oComOccurrence As Object = Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO(Occurrence)
                Dim oInputHelper As SP3DPIA.SmartOccurrence.IJSmartOccurrenceInputHelper = DirectCast(oComOccurrence, SP3DPIA.SmartOccurrence.IJSmartOccurrenceInputHelper)

                oInputHelper.RemoveAllInputs()
                oInputHelper.AddObjectInputEx(1, "IJFullObject", "ReferencesCollectionToReferencesRelationDestination", "", oPointCOM)
            End Set
        End Property

    End Class

End Namespace

