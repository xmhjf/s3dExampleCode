Imports Microsoft.VisualBasic
Imports Ingr.SP3D.Common.Middle
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Grids.Middle

Public Class IndexNameRule
    Inherits NameRuleBase

    Private Const MODELDATABASE = "Model"

    Private Const strDecimalFormat As String = "0.000"    'Precision of the Position 

    ''' <summary>
    '''  Creates a name for the object passed in. The name is based on the parents
    '''  name and object name.The Naming Parents are added in AddNamingParents().
    '''  Both these methods are called from naming rule semantic.
    ''' </summary>
    ''' <param name="oEntity">Child object that needs to have the naming rule naming.</param>
    ''' <param name="oParents">Naming parents collection.</param>
    ''' <param name="oActiveEntity">Naming rules active entity on which the NamingParentsString is stored.</param>
    ''' <exception cref="ArgumentNullException">The Grids entity is null.</exception> 
    Public Overrides Sub ComputeName(ByVal oEntity As BusinessObject, ByVal oParents As ReadOnlyCollection(Of BusinessObject), _
                                     ByVal oActiveEntity As BusinessObject)

        Dim oGridEntity As GridPlaneBase = Nothing
        Dim oGridCylinderEntity As GridCylinder = Nothing
        Dim secondaryindex As Long
        Dim tertiaryindex As Long
        Dim name As String = Nothing
        Dim childNamedItem As INamedItem = oEntity

        Try
            'check whether oEntity is existing, if it is nothing throw ArgumentNullException.
            If oEntity Is Nothing Then
                Throw New ArgumentNullException()
            End If

            Dim axis As GridAxis = Nothing
            'Check for the type of oEntity and proceed.
            If TypeOf (oEntity) Is GridPlaneBase Then
                Dim gridPlane As GridPlaneBase = oEntity
                axis = gridPlane.Axis
                Dim position As Position
                position = gridPlane.Origin
                'Name the grid planes according to the axis type.
                Select Case axis.AxisType
                    Case AxisType.X
                        name = "X" + Math.Round(position.X, 3).ToString
                    Case AxisType.Y
                        name = "Y" + Math.Round(position.Y, 3).ToString
                    Case AxisType.Z
                        name = "Z" + Math.Round(position.Z, 3).ToString
                End Select
            End If

            'Set user defined Name.
            childNamedItem.SetUserDefinedName(name.ToString())
        Catch ex As Exception
            Throw New Exception("GridsNamingRulesNetVB.IndexNameRule.ComputeName " + ex.Message)
        End Try

    End Sub
    ''' <summary>
    ''' Gets the index.
    ''' </summary>
    ''' <param name="entity">The entity.</param>
    ''' <returns></returns>
    Private Function GetIndex(ByVal entity As GridPlane) As Integer
        Dim currGridPlane As GridPlane
        Dim prevGridPlane As GridPlane
        Dim index As Integer
        index = 0
        currGridPlane = entity
        'Traverse through all the planes until very first plane on the axis is got.
        Do
            Try
                prevGridPlane = currGridPlane.GetReferenceGridPlane(ReferenceType.Previous, NestingLevelType.Primary)
                'if there is no reference to the index, return the current index i.e. index = 0.
            Catch NoReferenceException As Exception
                Return index
            End Try
            index = index + 1
            currGridPlane = prevGridPlane
        Loop
    End Function

    ''' <summary>
    ''' All the Naming Parents that need to participate in an objects naming are added here to the
    ''' Collection(Of BusinessObject). The parents added here are used in computing the name of the object in
    ''' ComputeName(). Both these methods are called from naming rule semantic.
    ''' </summary>
    ''' <param name="oEntity">Child object that needs to have the naming rule naming.</param>
    ''' <returns> Collection of parents that participate </returns>

    Public Overrides Function GetNamingParents(ByVal oEntity As BusinessObject) As Collection(Of BusinessObject)

        Dim oParentsColl As New Collection(Of BusinessObject)

        GetNamingParents = oParentsColl
    End Function
    
End Class
