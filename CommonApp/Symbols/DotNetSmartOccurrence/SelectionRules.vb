Option Explicit On
Imports System
Imports System.Math
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.ReferenceData.Middle
Imports Ingr.SP3D.Common.Exceptions
'Imports Ingr.SP3D.Structure.Middle

Namespace Symbols
    Friend Enum RuleInputs
        Object1 = 2
        Object2 = 3
    End Enum

    <RuleVersion("1.0.0.0")> _
    Public Class TestInputsSelectionRule : Inherits SelectorRule
        Private Const CONST_ErrorCodeList As String = "DotNetCustomAsmErrorCodeList"
        Private Const CONST_ObsoletePlaneSupport As Integer = 9

        Public Overrides ReadOnly Property Selections() As ReadOnlyCollection(Of String)
            Get
                Dim choices As New Collection(Of String)()

                ' Access SmartOccurrence to verify that it is valid
                If Occurrence Is Nothing Then
                    Throw New CmnException("Occurrence is nothing!")
                End If

                ' Access PartClass to verify that it is valid
                If PartClass Is Nothing Then
                    Throw New CmnException("Not expecting the PartClass property to be nothing!")
                Else
                    Dim oPartClass As PartClass
                    oPartClass = DirectCast(Me.PartClass, PartClass)
                    If oPartClass.PartClassName = "TestSelectionRules" Then
                        ' Just verifying that we can get at it
                    End If
                End If

                ' If one input then chose between smart items based on the type of input
                If Inputs.Count = 0 Then
                    choices.Add("SmartItemWithNoInputObjects")
                ElseIf Inputs(RuleInputs.Object1) Is Nothing Then
                    choices.Add("SmartItemWithNoInputObjects")
                Else
                    Dim objInput As BusinessObject
                    objInput = Inputs(RuleInputs.Object1)
                    If TypeOf (objInput) Is Line3d Then
                        choices.Add("TestSelRulesWithLines")
                    ElseIf TypeOf (objInput) Is Point3d Then
                        choices.Add("TestSelRulesWithPnts")
                    ElseIf TypeOf (objInput) Is Plane3d Then
                        choices.Add("SmartItemWithPlaneInput")
                        ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, CONST_ErrorCodeList, CONST_ObsoletePlaneSupport, "Do not support plane as input.", Occurrence)
                    ElseIf TypeOf (objInput) Is Arc3d Then
                        choices.Add("TestSelRulesWithLines")
                        ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, CONST_ErrorCodeList, 101, "Lines are supported so we'll allow arcs too.", Occurrence)
                    ElseIf TypeOf (objInput) Is Cone3d Then
                        choices.Add("SmartItemWithPlaneInput")
                        Throw New CmnException("Invalid input into selection rule")
                    ElseIf TypeOf (objInput) Is Circle3d Then
                        ' Don't return anything
                    Else
                        Throw New CmnException("Unexpected failure")
                    End If
                End If

                Return New ReadOnlyCollection(Of String)(choices)
            End Get
        End Property
    End Class

    <RuleVersion("1.0.0.0")> _
    Public Class TestLineInputSelectionRule : Inherits SelectorRule
        Public Overrides ReadOnly Property Selections() As ReadOnlyCollection(Of String)
            Get
                Dim choices As New Collection(Of String)()

                ' Access SmartOccurrence to verify that it is valid
                If Occurrence Is Nothing Then
                    Throw New CmnException("Occurrence is nothing!")
                End If

                ' Access PartClass to verify that it is valid
                If PartClass Is Nothing Then
                    Throw New CmnException("Not expecting the PartClass property to be nothing!")
                Else
                    Dim oPartClass As PartClass
                    oPartClass = DirectCast(Me.PartClass, PartClass)
                    If oPartClass.PartClassName = "TestSelRulesWithLines" Then
                        ' Just verifying that we can get at it
                    End If
                End If

                ' If one input then chose between smart items based on the type of input
                If Inputs.Count = 0 Then
                    Throw New CmnException("Expected an input object")
                ElseIf Inputs.Count < 2 Then
                    Throw New CmnException("Expected two objects")
                Else
                    Dim objInput1 As BusinessObject
                    objInput1 = Inputs(RuleInputs.Object1)
                    If Not TypeOf (objInput1) Is Line3d And Not TypeOf (objInput1) Is Arc3d Then
                        Throw New CmnException("Expected first object to be a line or arc object")
                    Else
                        Dim objInput2 As BusinessObject
                        objInput2 = Inputs(RuleInputs.Object2)
                        If TypeOf (objInput2) Is Line3d Or TypeOf (objInput2) Is Arc3d Then
                            choices.Add("SmartItemWith2LinesInput")
                        Else
                            choices.Add("SmartItemWithLineInput")
                        End If
                    End If
                End If

                Return New ReadOnlyCollection(Of String)(choices)
            End Get
        End Property
    End Class

    <RuleVersion("1.0.0.0")> _
    Public Class TestPointInputSelectionRule : Inherits SelectorRule
        Public Overrides ReadOnly Property Selections() As ReadOnlyCollection(Of String)
            Get
                Dim choices As New Collection(Of String)()

                ' Access SmartOccurrence to verify that it is valid
                If Occurrence Is Nothing Then
                    Throw New CmnException("Occurrence is nothing!")
                End If

                ' Access PartClass to verify that it is valid
                If PartClass Is Nothing Then
                    Throw New CmnException("Not expecting the PartClass property to be nothing!")
                Else
                    Dim oPartClass As PartClass
                    oPartClass = DirectCast(Me.PartClass, PartClass)
                    If oPartClass.PartClassName = "TestSelRulesWithPnts" Then
                        ' Just verifying that we can get at it
                    End If
                End If

                ' If one input then chose between smart items based on the type of input
                If Inputs.Count = 0 Then
                    Throw New CmnException("Expected an input object")
                Else
                    Dim objInput As BusinessObject
                    objInput = Inputs(RuleInputs.Object1)
                    If TypeOf (objInput) Is Point3d Then
                        choices.Add("SmartItemWithPointInput")
                    Else
                        Throw New CmnException("Expected a point input object")
                    End If
                End If

                Return New ReadOnlyCollection(Of String)(choices)
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Root selector for selection rule testing
    ''' </summary>
    ''' <remarks></remarks>
    <RuleVersion("1.4.0.0")> _
    <RuleInterface("ITestRootSelector", "Test Selection Rule")> _
    Public Class TestRootSelector : Inherits SelectorRule
        <SelectorQuestionString(1, "BaseQuestion", "Base question", "Some string")> _
        Public m_BaseQuestionAnswer As SelectorQuestionString
        <SelectorQuestionCodelist(2, "CodelistQuestion", "Codelist Question", "QuestionCodelistOne", 1)> _
        Public m_CLQuestionAnswer As SelectorQuestionCodelist
        ''' <summary>
        ''' Return the choices for the selection rule
        ''' </summary>
        ''' <returns>Array of of Smart item names and/or Smart class names.</returns>
        ''' <remarks></remarks>
        Public Overrides ReadOnly Property Selections() As ReadOnlyCollection(Of String)
            Get
                Dim choices As New Collection(Of String)()
                If m_BaseQuestionAnswer.Value = "Some string" Then
                    choices.Add("TestSelRulesWithQs")
                ElseIf m_BaseQuestionAnswer.Value = "First1of2SelectionsA" Then
                    choices.Add("TestSelRulesWithQs")
                    choices.Add("FirstSmartItemForTestRootSelector")
                ElseIf m_BaseQuestionAnswer.Value = "First1of2SelectionsB" Then
                    choices.Add("FirstSmartItemForTestRootSelector")
                    choices.Add("TestSelRulesWithQs")
                ElseIf m_BaseQuestionAnswer.Value = "TestParentAnswer" Then
                    choices.Add("TestSelRulesWithQs")
                ElseIf m_BaseQuestionAnswer.Value = "PickFourthSmartItemForSelectionRule" Then
                    choices.Add("TestSelRulesWithQs")
                ElseIf m_BaseQuestionAnswer.Value = "ToDoError" Then
                    choices.Add("TestSelWithToDoMsg")
                    Me.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "SelectionRuleATP", 1001, "ToDOError A")
                ElseIf m_BaseQuestionAnswer.Value = "ToDoWarning" Then
                    choices.Add("TestSelWithToDoMsg")
                    Me.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "SelectionRuleATP", 2001, "ToDOWarning A")
                ElseIf m_BaseQuestionAnswer.Value = "PickWarningSmartItemForSelectionRule" Then
                    choices.Add("WarningSmartItemForSelectionRule")
                    Me.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "SelectionRuleATP", 2002, "ToDOWarning B")
                ElseIf m_BaseQuestionAnswer.Value = "Bad value" Then
                    Throw New CmnException("Incorrect anwser.")
                ElseIf m_BaseQuestionAnswer.Value = "NoQuestions" Then
                    choices.Add("TestNoQuestions")
                ElseIf m_BaseQuestionAnswer.Value = "NoQuestionsSelectGrandChild" Then
                    choices.Add("TestNoQuestions")
                ElseIf m_BaseQuestionAnswer.Value = "NoSelectionRule" Then
                    choices.Add("TestNoSelectionRule")
                ElseIf m_BaseQuestionAnswer.Value = "DefaultAnswer1" Then
                    choices.Add("TestDefaultAnswer")
                ElseIf m_BaseQuestionAnswer.Value = "SmartItemWithCustomSymbolDef" Then
                    choices.Add("TestSelRulesWithQs")
                ElseIf m_BaseQuestionAnswer.Value = "TestParentSelection" Then
                    choices.Add("TestSelRulesWithQs")
                ElseIf m_BaseQuestionAnswer.Value = "TestSiblingSelection" Then
                    choices.Add("TestSelRulesWithQs")
                ElseIf m_BaseQuestionAnswer.Value = "TestCOMSelection" Then
                    choices.Add("TestCOMSelRulesWithQs")
                Else
                    choices.Add("TestSelectorWithBadAnswer")
                End If


                Return New ReadOnlyCollection(Of String)(choices)
            End Get
        End Property


    End Class

    ''' <summary>
    ''' Test No Questions
    ''' </summary>
    ''' <remarks></remarks>
    <RuleVersion("1.0.0.0")> _
    Public Class TestNoQuestions : Inherits SelectorRule
        Public Overrides ReadOnly Property Selections() As ReadOnlyCollection(Of String)
            Get
                Const INTERFACE_ItestRootSelector As String = "ITestRootSelector"
                Dim oSymbolOcc As BusinessObject = Me.Occurrence
                Dim oPropValue As PropertyValueString = DirectCast(oSymbolOcc.GetPropertyValue(INTERFACE_ItestRootSelector, "BaseQuestion"), PropertyValueString)

                Dim choices As New Collection(Of String)()
                ' This property value is intended to test the condition where
                '   a selection rule choses a great grandchild smart item skipping
                '   selection rules of all the child SmartClasses
                If oPropValue.PropValue = "NoQuestionsSelectGrandChild" Then
                    choices.Add("WithParameterRuleNoSelRule")
                Else
                    choices.Add("TestWithQuestions")
                End If
                Return New ReadOnlyCollection(Of String)(choices)
            End Get
        End Property


    End Class

    ''' <summary>
    ''' Test No Selection Rule
    ''' </summary>
    ''' <remarks></remarks>
    <RuleVersion("1.0.0.0")> _
    Public Class TestNoSelectionRule : Inherits SelectorRule
        Public Overrides ReadOnly Property Selections() As ReadOnlyCollection(Of String)
            Get
                Dim choices As New Collection(Of String)()
                choices.Add("FirstSmartItemForTestRootSelector")
                Return New ReadOnlyCollection(Of String)(choices)
            End Get
        End Property


    End Class
    ''' <summary>
    ''' Test With Questions
    ''' </summary>
    ''' <remarks></remarks>
    <RuleVersion("1.0.0.0")> _
    <RuleInterface("ITestWithQuestions", "Test With Questions")> _
    Public Class TestWithQuestions : Inherits SelectorRule
        <SelectorQuestionString(1, "BaseQuestion", "Base question", "Some string")> _
        Public m_BaseQuestionAnswer As SelectorQuestionString

        ''' <summary>
        ''' Return the choices for the selection rule
        ''' </summary>
        ''' <returns>Array of of Smart item names and/or Smart class names.</returns>
        ''' <remarks></remarks>
        Public Overrides ReadOnly Property Selections() As ReadOnlyCollection(Of String)
            Get
                Dim choices As New Collection(Of String)()
                choices.Add("FirstSmartItemForTestWithQuestions")
                Return New ReadOnlyCollection(Of String)(choices)
            End Get
        End Property


    End Class

    Friend Enum QuestionCodelistOne
        AnswerA = 1
        AnswerB = 2
        AnswerC = 3
    End Enum

    ''' <summary>
    ''' Test selection rules with selector questions
    ''' </summary>
    ''' <remarks></remarks>
    <RuleVersion("1.0.0.0")> _
    <RuleInterface("IMySelectionRule", "My Selection Rule")> _
    Public Class SelectionRuleWithQuestions : Inherits SelectorRule
        <SelectorQuestionCodelist(1, "QuestionNumber1", "Question Number One", "QuestionCodelistOne", QuestionCodelistOne.AnswerA)> _
        Public m_AnswerNumber1 As SelectorQuestionCodelist
        <SelectorQuestionCodelist(2, "QuestionNumber2", "Question Number Two", "QuestionCodelistTwo", QuestionCodelistOne.AnswerB)> _
        Public m_AnswerNumber2 As SelectorQuestionCodelist
        <SelectorQuestionString(3, "QuestionNumber3", "Question Number Three", "Value 1")> _
        Public m_AnswerNumber3 As SelectorQuestionString
        <SelectorQuestionCodelist(4, "HiddenQuestionNumber1", "Hidden Question Number One", "QuestionCodelistOne", QuestionCodelistOne.AnswerB)> _
        <HiddenQuestion(True)> _
        Public m_AnswerNumber4 As SelectorQuestionCodelist

        ''' <summary>
        ''' Return the choices for the selection rule
        ''' </summary>
        ''' <returns>Array of of Smart item names and/or Smart class names.</returns>
        ''' <remarks></remarks>
        Public Overrides ReadOnly Property Selections() As ReadOnlyCollection(Of String)
            Get
                Dim choices As New Collection(Of String)()
                Dim oPropAnswer As PropertyValueString = Occurrence.GetPropertyValue("ITestRootSelector", "BaseQuestion")

                If (oPropAnswer.PropValue = "TestParentAnswer") Then
                    choices.Add("SecondSmartItemForSelectionRule")
                ElseIf (oPropAnswer.PropValue = "First1of2SelectionsA") Then
                    choices.Add("ThirdSmartItemForSelectionRule")
                ElseIf (oPropAnswer.PropValue = "PickFourthSmartItemForSelectionRule") Then
                    choices.Add("FourthSmartItemForSelectionRule")
                ElseIf (oPropAnswer.PropValue = "SmartItemWithCustomSymbolDef") Then
                    choices.Add("CustomSymWithOnlyCachedSymbolOutputs")
                ElseIf (oPropAnswer.PropValue = "TestParentSelection") Then
                    choices.Add("TestRootSelector")
                ElseIf (oPropAnswer.PropValue = "TestSiblingSelection") Then
                    choices.Add("TestDefaultAnswer")
                ElseIf (m_AnswerNumber1.Value = QuestionCodelistOne.AnswerA) Then
                    choices.Add("FirstSmartItemForSelectionRule")
                End If

                Return New ReadOnlyCollection(Of String)(choices)
            End Get
        End Property

        Public Overrides Sub OverrideDefaultAnswer(ByVal selectorQuestion As SelectorQuestion)
            Select Case selectorQuestion.Name
                Case "QuestionNumber2"
                    Dim oCodelistAnswer As SelectorQuestionCodelist = DirectCast(selectorQuestion, SelectorQuestionCodelist)
                    oCodelistAnswer.Value = QuestionCodelistOne.AnswerA
            End Select
        End Sub
    End Class
    ''' <summary>
    ''' Test Dedault Answer
    ''' </summary>
    ''' <remarks></remarks>
    <RuleVersion("1.1.0.0")> _
    <RuleInterface("ITestDefaultAnswer", "Test Dedault Answer")> _
    Public Class TestDefaultAnswer : Inherits SelectorRule
        <SelectorQuestionCodelist(1, "QuestionNumberOne", "Question Number One", "QuestionCodelistOne", QuestionCodelistOne.AnswerA)> _
        Public m_AnswerNumber1 As SelectorQuestionCodelist
        <SelectorQuestionString(2, "QuestionNumberTwo", "Question Number Two", "AQTwo")> _
        Public m_AnswerNumber2 As SelectorQuestionString
        <SelectorQuestionString(3, "QuestionNumberThree", "Question Number Three", "AQThree")> _
        Public m_AnswerNumber3 As SelectorQuestionString
        <SelectorQuestionString(4, "QuestionNumberFour", "Question Number Four", "AQFour")> _
        Public m_AnswerNumber4 As SelectorQuestionString

        ''' <summary>
        ''' Return the choices for the selection rule
        ''' </summary>
        ''' <returns>Array of of Smart item names and/or Smart class names.</returns>
        ''' <remarks></remarks>
        Public Overrides ReadOnly Property Selections() As ReadOnlyCollection(Of String)
            Get
                Dim choices As New Collection(Of String)()

                If (m_AnswerNumber1.Value = QuestionCodelistOne.AnswerA) Then
                    choices.Add("FirstSmartItemForDefaultAnswer")
                ElseIf (m_AnswerNumber1.Value = QuestionCodelistOne.AnswerB) Then
                    choices.Add("SecondSmartItemForDefaultAnswer")
                Else 'If (m_AnswerNumber1.Value = QuestionCodelistOne.AnswerC) Then
                    choices.Add("ThirdSmartItemForDefaultAnswer")
                End If

                Return New ReadOnlyCollection(Of String)(choices)
            End Get
        End Property

        Public Overrides Sub OverrideDefaultAnswer(ByVal answer As SelectorQuestion)
            Dim oStringAnswer As SelectorQuestionString
            Dim oCodelistAnswer As SelectorQuestionCodelist
            Select Case answer.Name
                Case "QuestionNumberOne"
                    oCodelistAnswer = DirectCast(answer, SelectorQuestionCodelist)
                    oCodelistAnswer.Value = QuestionCodelistOne.AnswerB
                Case "QuestionNumberTwo"
                    oStringAnswer = DirectCast(answer, SelectorQuestionString)
                    oStringAnswer.Value = "DefaultAnswer2"
                Case "QuestionNumberThree"
                    oStringAnswer = DirectCast(answer, SelectorQuestionString)
                    oStringAnswer.Value = "DefaultAnswer3"
                Case "QuestionNumberFour"
                    oStringAnswer = DirectCast(answer, SelectorQuestionString)
                    oStringAnswer.Value = "DefaultAnswer4"
            End Select
        End Sub
    End Class
    ''' <summary>
    ''' Symbol Definition to test rules with selector questions
    ''' </summary>
    ''' <remarks></remarks>
    <CacheOption(CacheOptionType.NonCached)> _
    <SymbolVersion("1.1.0.0")> _
    <OutputNotification("IJDGeometry")> _
    Public Class FirstSmartItemForSelectionRule : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   2. "Diameter"

        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.3)> _
        Public m_dblDiameter As InputDouble

        Private Const CONST_Base As String = "Base"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Base, "Base sphere")> _
        Public m_physicalAspect As AspectDefinition

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            Dim dblBaseRadius As Double
            m_dblDiameter.Value = m_dblDiameter.Value / 2.0
            m_dblDiameter.Value = m_dblDiameter.Value * 2.0
            dblBaseRadius = m_dblDiameter.Value / 2.0

            '=================================================
            ' Construct base
            '=================================================
            Dim objBase As Sphere3d

            objBase = New Sphere3d(oSP3DConnection, New Position(0.0, 0.0, dblBaseRadius), dblBaseRadius, True)

            m_physicalAspect.Outputs(CONST_Base) = objBase

        End Sub
    End Class

    ''' <summary>
    ''' Second Symbol Definition to test rules with selector questions
    ''' </summary>
    ''' <remarks></remarks>
    Public Class SecondSmartItemForSelectionRule : Inherits FirstSmartItemForSelectionRule
        Protected Overrides Sub ConstructOutputs()
            MyBase.ConstructOutputs()
            If DirectCast(m_catalogPart.Value, Part).PartNumber = "SecondSmartItemForDefaultAnswer" Then
                Dim oPropertyValue As PropertyValue = Occurrence.GetPropertyValue("ITestDefaultAnswer", "QuestionNumberOne")
                If oPropertyValue Is Nothing Then
                    ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "ITestDefaultAnswer interface's QuestionNumberOne property is null.")
                ElseIf Not TypeOf (oPropertyValue) Is PropertyValueCodelist Then
                    ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "ITestDefaultAnswer interface's QuestionNumberOne property is not a codelist property.")
                Else
                    Dim oPropertyCodelist As PropertyValueCodelist = DirectCast(oPropertyValue, PropertyValueCodelist)
                    If oPropertyCodelist.PropValue <> QuestionCodelistOne.AnswerB Then
                        ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "Expected default answer to be QuestionCodelistOne.AnswerB.")
                    End If
                End If
            End If
        End Sub
    End Class

    ''' <summary>
    ''' Third Symbol Definition to test rules with selector questions
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ThirdSmartItemForSelectionRule : Inherits FirstSmartItemForSelectionRule
    End Class

    ''' <summary>
    ''' Test selection rules with selector questions
    ''' </summary>
    ''' <remarks></remarks>
    <RuleVersion("1.0.0.0")> _
    <RuleInterface("IMyCOMToNetSelectionRule", "My selection rule to test mixing COM and .Net rules")> _
    Public Class TestDotNetFromCOMSelector : Inherits SelectorRule

        <SelectorQuestionCodelist(1, "OneQuestion", "Question One", "QuestionCodelistOne", QuestionCodelistOne.AnswerA)> _
        Public m_AnswerNumber1 As SelectorQuestionCodelist
        <SelectorQuestionDouble(2, "QuestionDouble", "Question with a  double", 0.1)> _
        <RuleDriven(False)> _
        Public m_doubleAnswer As SelectorQuestionDouble

        ''' <summary>
        ''' Return the choices for the selection rule
        ''' </summary>
        ''' <returns>Array of of Smart item names and/or Smart class names.</returns>
        ''' <remarks></remarks>
        Public Overrides ReadOnly Property Selections() As ReadOnlyCollection(Of String)
            Get
                Dim choices As New Collection(Of String)()

                choices.Add("LeafDotNetCOMSelector")

                Return New ReadOnlyCollection(Of String)(choices)
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Test selection rules with selector questions
    ''' </summary>
    ''' <remarks></remarks>
    <RuleVersion("1.0.0.0")> _
    <RuleInterface("IMyLeafSelectionRule", "My slection rule to test mixing COM and .Net rules")> _
    Public Class LeafDotNetCOMSelector : Inherits SelectorRule

        <SelectorQuestionCodelist(1, "LeafQuestionNumber1", "Question Number One", "QuestionCodelistOne", QuestionCodelistOne.AnswerA)> _
        Public m_AnswerNumber1 As SelectorQuestionCodelist
        <SelectorQuestionString(2, "LeafQuestionNumber2", "Question Number Two", "Leaf Value")> _
        Public m_AnswerNumber2 As SelectorQuestionString

        ''' <summary>
        ''' Return the choices for the selection rule
        ''' </summary>
        ''' <returns>Array of of Smart item names and/or Smart class names.</returns>
        ''' <remarks></remarks>
        Public Overrides ReadOnly Property Selections() As ReadOnlyCollection(Of String)
            Get
                Dim choices As New Collection(Of String)()

                If (m_AnswerNumber1.Value = QuestionCodelistOne.AnswerA) Then
                    choices.Add("SmartItem2FromLeafSelection")
                    choices.Add("SmartItem3FromLeafSelection")
                    choices.Add("SmartItem1FromLeafSelection")
                End If

                Return New ReadOnlyCollection(Of String)(choices)
            End Get
        End Property
    End Class


    ''' <summary>
    ''' Test selection rules with selector questions
    ''' </summary>
    ''' <remarks></remarks>
    <RuleVersion("1.0.0.0")> _
    <RuleInterface("IMyToDoSelectionRule", "My selection rule to test ToDo messages in the rule")> _
    Public Class TestSelectorWithToDoMessage : Inherits SelectorRule
        <SelectorQuestionString(1, "Question", "First Question", "None")> _
        Public m_AnswerNumber As SelectorQuestionString

        ''' <summary>
        ''' Return the choices for the selection rule
        ''' </summary>
        ''' <returns>Array of of Smart item names and/or Smart class names.</returns>
        ''' <remarks></remarks>
        Public Overrides ReadOnly Property Selections() As ReadOnlyCollection(Of String)
            Get
                Dim choices As New Collection(Of String)()

                choices.Add("CustomAssembly1cm")
                choices.Add("CustomAssembly2cm")

                Return New ReadOnlyCollection(Of String)(choices)
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Tests for parameter rules
    ''' </summary>
    ''' <remarks></remarks>
    <CacheOption(CacheOptionType.Cached)> _
    <SymbolVersion("1.0.0.0")> _
    Public Class CustomAsmForParameterRule : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   5. "Diameter"

        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.0)> _
        Public m_dblDiameter As InputDouble
        <InputDouble(3, "OriginX", "OriginX", 0.0)> _
        Public m_dblOriginX As InputDouble
        <InputDouble(4, "OriginY", "Origin Y", 0.0)> _
        Public m_dblOriginY As InputDouble
        <InputDouble(5, "OriginZ", "Origin Z", 0.0)> _
        Public m_dblOriginZ As InputDouble
        <InputDouble(6, "Volume", "Volume", 0.0)> _
        Public m_dblVolume As InputDouble
        <InputDouble(7, "SurfaceArea", "Surface Area", 0.0)> _
        Public m_dblSurfaceArea As InputDouble
        <InputDouble(8, "SecondarySideNominalThroatThickness", "Secondary Side Nominal Throat Thickness", 0.0001)> _
        Public m_dblSecondarySideNominalThroatThickness As InputDouble

        Private Const CONST_SymbolSphere As String = "SymbolSphere"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_SymbolSphere, "Symbol Sphere")> _
        Public m_physicalAspect As AspectDefinition

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            Dim dblBaseRadius As Double
            dblBaseRadius = m_dblDiameter.Value / 2.0

            '=================================================
            ' Construct base
            '=================================================
            Dim objBase As Sphere3d

            objBase = New Sphere3d(oSP3DConnection, New Position(m_dblOriginX.Value, m_dblOriginY.Value, m_dblOriginZ.Value), dblBaseRadius, True)

            m_physicalAspect.Outputs(CONST_SymbolSphere) = objBase
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            ' Verify that we can at least access the occurrence during the evaluate
            Dim sString As String = Occurrence.ToString()
        End Sub


#Region "Implement property Management methods"
        ''' <summary>
        ''' Preload property pages
        ''' </summary>
        ''' <param name="oBusinessObject"></param>
        ''' <param name="boProperties"></param>
        ''' <remarks></remarks>
        Public Overrides Sub OnPreLoad(ByVal oBusinessObject As BusinessObject, ByVal boProperties As ReadOnlyCollection(Of PropertyDescriptor))
            Dim propDescriptor As PropertyDescriptor

            ' Set all properties other than one as read-only
            For Each propDescriptor In boProperties
                If propDescriptor.Property.PropertyInfo.Name <> "SecondarySideActualLegLength" Then
                    propDescriptor.ReadOnly = True
                End If
            Next
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
            Dim retStatus As Boolean
            retStatus = True
            If oPropToChange.Property.PropertyInfo.Name = "Volume" Then
                Dim propDescriptor As PropertyDescriptor
                For Each propDescriptor In CollAllDisplayedValues
                    If propDescriptor.Property.PropertyInfo.Name = "OriginZ" Then
                        propDescriptor.ReadOnly = False
                    ElseIf propDescriptor.Property.PropertyInfo.Name = "OriginY" Then
                        Dim PropDouble As PropertyValueDouble = DirectCast(propDescriptor.Property, PropertyValueDouble)
                        PropDouble.PropValue = 4.4
                    End If
                Next
            ElseIf oPropToChange.Property.PropertyInfo.Name = "SurfaceArea" Then
                strErrorMsg = "Bad surface area."
                retStatus = False
            ElseIf oPropToChange.Property.PropertyInfo.Name = "OriginX" Then
                strErrorMsg = "Bad Origin X."
                retStatus = False
            End If

            Return retStatus
        End Function
#End Region

    End Class

    <RuleVersion("1.0.0.0")> _
    <RuleInterface("ITestParameterRule", "Test parameter rule interface")> _
    Public Class TestParameterRule : Inherits ParameterRule

        <ControlledParameter(1, "Diameter", "Diameter")> _
        Public m_Diameter As ControlledParameterDouble
        <ControlledParameter(2, "OriginX", "Origin X")> _
        Public m_OriginX As ControlledParameterDouble
        <ControlledParameter(3, "OriginY", "Origin Y")> _
        Public m_OriginY As ControlledParameterDouble
        <ControlledParameter(4, "OriginZ", "Origin Z")> _
        Public m_OriginZ As ControlledParameterDouble
        <ControlledParameter(5, "Volume", "Volume")> _
        Public m_Volume As ControlledParameterDouble
        <ControlledParameter(6, "SurfaceArea", "Surface Area")> _
        Public m_SurfaceArea As ControlledParameterDouble
        <ControlledParameter(7, "SecondarySideNominalThroatThickness", "Secondary Side Nominal Throat Thickness")> _
        Public m_WeldParamterRule As ControlledParameterDouble


        Public Overrides Sub Evaluate()
            ' Value of 1111 then create a ToDoError
            If Not m_Diameter.Value Is Nothing Then
                Dim diff As Double = m_Diameter.Value - 1111.0
                If Abs(diff) < 0.01 Then
                    Me.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "CustomAssemblyATP", 1111, "Throwing a ToDo Error in a parameter rule")
                End If
                ' Value of 2222 then throw an exception
                diff = m_Diameter.Value - 2222.0
                If Abs(diff) < 0.01 Then
                    Throw New CmnException("Test exception thrown in parameter rule")
                End If
                ' Value of 3333 then create a ToDo warning
                diff = m_Diameter.Value - 3333.0
                If Abs(diff) < 0.01 Then
                    Me.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, "CustomAssemblyATP", 3333, "Throwing a ToDo Error in a parameter rule")
                End If
                ' Value of 4444 then create a ToDo error with a "localized message" (i.e., non-codelisted)
                diff = m_Diameter.Value - 4444.0
                If Abs(diff) < 0.01 Then
                    Me.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "Bad diameter!")
                End If
            End If
            m_Diameter.Value = 2.0
            m_OriginX.Value = 1.0
            m_OriginY.Value = 1.1
            m_OriginZ.Value = 1.2
            m_Volume.Value = 5.0
            m_SurfaceArea.Value = 4.0
            m_WeldParamterRule.Value = 0.0002
        End Sub
    End Class

    <RuleVersion("1.1.0.0")> _
    <RuleInterface("ITestBadParameterRule", "Test bad parameter rule interface")> _
    Public Class TestBadParameterRule : Inherits ParameterRule

        <ControlledParameter(1, "Radius", "Radius")> _
        Public m_Radius As ControlledParameterDouble
        <ControlledParameter(2, "Diameter", "Diameter")> _
        Public m_Diameter As ControlledParameterDouble
        <ControlledParameter(3, "SomeString", "Some string")> _
        Public m_SomeString As ControlledParameterString

        Public Overrides Sub Evaluate()
            m_Radius.Value = 2.0
            m_Diameter.Value = 2.7
            m_SomeString.Value = "oops"
        End Sub
    End Class

    ''' <summary>
    ''' Tests for parameter rules driving undefaulted values from the bulkloaded spreadsheet
    ''' </summary>
    ''' <remarks></remarks>
    <CacheOption(CacheOptionType.NonCached)> _
    <SymbolVersion("1.0.0.0")> _
    Public Class CustomAsmWithNonDefaultedInputs : Inherits CustomAssemblyDefinition
        '   1. "Part"  ( Catalog part )
        '   5. "Diameter"

        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 0.0)> _
        Public m_dblDiameter As InputDouble
        <InputDouble(3, "OriginX", "OriginX", 0.0)> _
        Public m_dblOriginX As InputDouble
        <InputDouble(4, "OriginY", "Origin Y", 0.0)> _
        Public m_dblOriginY As InputDouble
        <InputDouble(5, "OriginZ", "Origin Z", 0.0)> _
        Public m_dblOriginZ As InputDouble
        <InputDouble(6, "Volume", "Volume", 0.0)> _
        Public m_dblVolume As InputDouble
        <InputDouble(7, "SurfaceArea", "Surface Area", 0.0)> _
        Public m_dblSurfaceArea As InputDouble
        <InputString(8, "Astring", "A string", "Hmm")> _
        Public m_sSomeString As InputString

        Private Const CONST_SymbolSphere As String = "SymbolSphere"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_SymbolSphere, "Symbol Sphere")> _
        Public m_physicalAspect As AspectDefinition

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            Dim dblBaseRadius As Double
            dblBaseRadius = m_dblDiameter.Value / 2.0

            If (m_sSomeString.Value <> "Something") Then
                Me.ToDoListMessage = New ToDoListMessage(ToDoMessageTypes.ToDoMessageError, "Expected symbol input to be set to parameter rule defined string.")
            End If
            '=================================================
            ' Construct base
            '=================================================
            Dim objBase As Sphere3d

            objBase = New Sphere3d(oSP3DConnection, New Position(m_dblOriginX.Value, m_dblOriginY.Value, m_dblOriginZ.Value), dblBaseRadius, True)

            m_physicalAspect.Outputs(CONST_SymbolSphere) = objBase
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            ' Verify that we can at least access the occurrence during the evaluate
            Dim sString As String = Occurrence.ToString()
        End Sub

    End Class

    <RuleVersion("1.0.0.0")> _
    <RuleInterface("IParameterRulesx", "Parameter rules setting inputs without default values")> _
    Public Class ParameterRuleToDefaultInputs : Inherits ParameterRule

        <ControlledParameter(1, "Diameter", "Diameter")> _
        Public m_Diameter As ControlledParameterDouble
        <ControlledParameter(2, "OriginX", "Origin X")> _
        Public m_OriginX As ControlledParameterDouble
        <ControlledParameter(3, "OriginY", "Origin Y")> _
        Public m_OriginY As ControlledParameterDouble
        <ControlledParameter(4, "OriginZ", "Origin Z")> _
        Public m_OriginZ As ControlledParameterDouble
        <ControlledParameter(5, "Volume", "Volume")> _
        Public m_Volume As ControlledParameterDouble
        <ControlledParameter(6, "SurfaceArea", "Surface Area")> _
        Public m_SurfaceArea As ControlledParameterDouble
        <ControlledParameter(7, "Astring", "A string")> _
        Public m_someString As ControlledParameterString

        Public Overrides Sub Evaluate()
            m_Diameter.Value = 2.0
            m_OriginX.Value = 2.0
            m_OriginY.Value = 2.0
            m_OriginZ.Value = 2.0
            m_Volume.Value = 3.14
            m_SurfaceArea.Value = 6.28
            m_someString.Value = "Something" ' Will be validated by the ConstructOutputs of assicated symbol
        End Sub
    End Class

    ''' <summary>
    ''' Root selector for testing changing the parameter rule where the rule
    ''' drives different values on the same interface.
    ''' </summary>
    ''' <remarks></remarks>
    <RuleVersion("1.0.0.0")> _
    <RuleInterface("ITestNo100RootSelector", "Test Selection Rule")> _
    Public Class TestNo100SelectionRule : Inherits SelectorRule
        <SelectorQuestionString(1, "WhichRule", "Which rule?", "Rule-A")> _
        Public m_whichRule As SelectorQuestionString
        ''' <summary>
        ''' Return the choices for the selection rule
        ''' </summary>
        ''' <returns>Array of of Smart item names and/or Smart class names.</returns>
        ''' <remarks></remarks>
        Public Overrides ReadOnly Property Selections() As ReadOnlyCollection(Of String)
            Get
                Dim choices As New Collection(Of String)()

                If m_whichRule.Value = "Rule-B" Then
                    choices.Add("TestNo100ParamRulesB")
                Else
                    choices.Add("TestNo100ParamRulesA")
                    choices.Add("TestNo100ParamRulesB")
                End If

                Return New ReadOnlyCollection(Of String)(choices)
            End Get
        End Property


    End Class

    ''' <summary>
    ''' Selector for picking the SmartItem in changing the parameter rule where the rule
    ''' drives different values on the same interface.
    ''' </summary>
    ''' <remarks></remarks>
    <RuleVersion("1.0.0.0")> _
    <RuleInterface("ITestNo100SelectionRuleA", "Test Selection Rule")> _
    Public Class TestNo100SelectionRuleA : Inherits SelectorRule
        <SelectorQuestionString(1, "WhichItem", "Which item?", "Double")> _
        Public m_whichItem As SelectorQuestionString
        ''' <summary>
        ''' Return the choices for the selection rule
        ''' </summary>
        ''' <returns>Array of of Smart item names and/or Smart class names.</returns>
        ''' <remarks></remarks>
        Public Overrides ReadOnly Property Selections() As ReadOnlyCollection(Of String)
            Get
                Dim choices As New Collection(Of String)()

                If m_whichItem.Value = "Double" Then
                    choices.Add("TestNo100RuleDrivingDouble")
                Else
                    choices.Add("TestNo100RuleDrivingLong")
                    choices.Add("TestNo100RuleDrivingDouble")
                End If

                Return New ReadOnlyCollection(Of String)(choices)
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Selector for picking the SmartItem in changing the parameter rule where the rule
    ''' drives different values on the same interface.
    ''' </summary>
    ''' <remarks></remarks>
    <RuleVersion("1.0.0.0")> _
    Public Class TestNo100SelectionRuleB : Inherits SelectorRule
        ''' <summary>
        ''' Return the choices for the selection rule
        ''' </summary>
        ''' <returns>Array of of Smart item names and/or Smart class names.</returns>
        ''' <remarks></remarks>
        Public Overrides ReadOnly Property Selections() As ReadOnlyCollection(Of String)
            Get
                Dim choices As New Collection(Of String)()

                choices.Add("TestNo100RuleDrivingFloat")

                Return New ReadOnlyCollection(Of String)(choices)
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Parameter rule driving some of the attributes (double and string) on the IJUGenericProperties interface
    ''' </summary>
    ''' <remarks></remarks>
    <RuleVersion("1.0.0.0")> _
    <RuleInterface("ITestNo100ParameterRuleDrivingDouble", "Parameter rule setting only the double and string attributes on interface IJUAGenericProperties")> _
    Public Class TestNo100ParameterRuleDrivingDouble : Inherits ParameterRule
        <ControlledParameter(1, "Diameter", "Diameter")> _
        Public m_Diameter As ControlledParameterDouble
        <ControlledParameter(2, "OriginX", "Origin X")> _
        Public m_OriginX As ControlledParameterDouble
        <ControlledParameter(3, "OriginY", "Origin Y")> _
        Public m_OriginY As ControlledParameterDouble
        <ControlledParameter(4, "OriginZ", "Origin Z")> _
        Public m_OriginZ As ControlledParameterDouble
        <ControlledParameter(5, "Volume", "Volume")> _
        Public m_Volume As ControlledParameterDouble
        <ControlledParameter(6, "SurfaceArea", "Surface Area")> _
        Public m_SurfaceArea As ControlledParameterDouble
        <ControlledParameter(7, "Adouble", "A double")> _
        Public m_someDouble As ControlledParameterDouble
        <ControlledParameter(8, "Astring", "A string")> _
        Public m_someString As ControlledParameterString
        <ControlledParameter(9, "PrimarySideSymbol", "Primar side symbol")> _
        Public m_primarySideSymbol As ControlledParameterDouble
        <ControlledParameter(10, "SecondarySideSymbol", "Secondary side symbol")> _
        Public m_secondarySideSymbol As ControlledParameterDouble
        <ControlledParameter(11, "PrimarySideGroove", "Primary side groove")> _
        Public m_primarySideGroove As ControlledParameterDouble
        <ControlledParameter(12, "SecondarySideGroove", "Secondary side groove")> _
        Public m_secondarySideGroove As ControlledParameterDouble
        <ControlledParameter(13, "PrimarySideGrooveSize", "Primary side groove size")> _
        Public m_primarySideGrooveSize As ControlledParameterDouble
        <ControlledParameter(14, "SecondarySideGrooveSize", "Secondary side groove size")> _
        Public m_secondarySideGrooveSize As ControlledParameterDouble
        <ControlledParameter(15, "PrimarySideActualThroatThickness", "Primary Side Actual Throat Thickness")> _
        Public m_primarySideActualThroatThickness As ControlledParameterDouble
        <ControlledParameter(16, "SecondarySideActualThroatThickness", "Secondary Side Actual Throat Thickness")> _
        Public m_secondarySideActualThroatThickness As ControlledParameterDouble
        <ControlledParameter(17, "PrimarySideNominalThroatThickness", "Primary Side Nominal Throat Thickness")> _
        Public m_primarySideNominalThroatThickness As ControlledParameterDouble
        <ControlledParameter(18, "SecondarySideNominalThroatThickness", "Secondary Side Nominal Throat Thickness")> _
        Public m_secondarySideNominalThroatThickness As ControlledParameterDouble
        <ControlledParameter(19, "FieldWeld", "Field Weld")> _
        Public m_fieldWeld As ControlledParameterDouble
        <ControlledParameter(20, "AllAround", "All Around")> _
        Public m_allAround As ControlledParameterDouble
        <ControlledParameter(21, "PrimarySideSupplementarySymbol", "Primary Side Supplementary Symbol")> _
        Public m_primarySideSupplementarySymbol As ControlledParameterDouble
        <ControlledParameter(22, "TailNotes", "Tail Notes")> _
        Public m_tailNotes As ControlledParameterString
        <ControlledParameter(23, "TailNoteIsReference", "Tail Note Is Reference")> _
        Public m_tailNoteIsReference As ControlledParameterDouble
        <ControlledParameter(24, "PrimarySideLength", "Primary Side Length")> _
        Public m_primarySideLength As ControlledParameterDouble
        <ControlledParameter(25, "SecondarySideLength", "Secondary Side Length")> _
        Public m_secondarySideLength As ControlledParameterDouble
        <ControlledParameter(26, "PrimarySidePitch", "Primary Side Pitch")> _
        Public m_primarySidePitch As ControlledParameterDouble
        <ControlledParameter(27, "SecondarySidePitch", "SecondarySidePitch")> _
        Public m_secondarySidePitch As ControlledParameterDouble
        <ControlledParameter(28, "PrimarySideDiameter", "Primary Side Diameter")> _
        Public m_primarySideDiameter As ControlledParameterDouble
        <ControlledParameter(29, "SecondarySideDiameter", "Secondary Side Diameter")> _
        Public m_secondarySideDiameter As ControlledParameterDouble
        <ControlledParameter(30, "PrimarySideContour", "Primary Side Contour")> _
        Public m_primarySideContour As ControlledParameterDouble
        <ControlledParameter(31, "SecondarySideContour", "Secondary Side Contour")> _
        Public m_secondarySideContour As ControlledParameterDouble
        <ControlledParameter(32, "PrimarySideFinishMethod", "Primary Side Finish Method")> _
        Public m_primarySideFinishMethod As ControlledParameterDouble
        <ControlledParameter(33, "SecondarySideFinishMethod", "Secondary Side Finish Method")> _
        Public m_secondarySideFinishMethod As ControlledParameterDouble
        <ControlledParameter(34, "PrimarySideRootOpening", "Primary Side Root Opening")> _
        Public m_primarySideRootOpening As ControlledParameterDouble
        <ControlledParameter(35, "SecondarySideRootOpening", "Secondary Side Root Opening")> _
        Public m_secondarySideRootOpening As ControlledParameterDouble
        <ControlledParameter(36, "PrimarySideGrooveAngle", "Primary Side Groove Angle")> _
        Public m_primarySideGrooveAngle As ControlledParameterDouble
        <ControlledParameter(37, "SecondarySideGrooveAngle", "Secondary Side Groove Angle")> _
        Public m_secondarySideGrooveAngle As ControlledParameterDouble
        <ControlledParameter(38, "PrimarySideNumberOfWelds", "Primary Side Number Of Welds")> _
        Public m_primarySideNumberOfWelds As ControlledParameterDouble
        <ControlledParameter(39, "SecondarySideNumberOfWelds", "Secondary Side Number Of Welds")> _
        Public m_secondarySideNumberOfWelds As ControlledParameterDouble
        <ControlledParameter(40, "PrimarySideActualLegLength", "Primary Side Actual Leg Length")> _
        Public m_primarySideActualLegLength As ControlledParameterDouble
        <ControlledParameter(41, "SecondarySideActualLegLength", "Secondary Side Actual Leg Length")> _
        Public m_secondarySideActualLegLength As ControlledParameterDouble

        Public Overrides Sub Evaluate()
            m_Diameter.Value = 2.0
            m_OriginX.Value = 2.0
            m_OriginY.Value = 2.0
            m_OriginZ.Value = 2.0
            m_Volume.Value = 2.0
            m_SurfaceArea.Value = 2.0
            m_someDouble.Value = 2.0
            m_someString.Value = "2.0"
            m_primarySideSymbol.Value = 2.0
            m_secondarySideSymbol.Value = 2.0
            m_primarySideGroove.Value = 2.0
            m_secondarySideGroove.Value = 2.0
            m_primarySideGrooveSize.Value = 2.0
            m_secondarySideGrooveSize.Value = 2.0
            m_primarySideActualThroatThickness.Value = 2.0
            m_secondarySideActualThroatThickness.Value = 2.0
            m_primarySideNominalThroatThickness.Value = 2.0
            m_secondarySideNominalThroatThickness.Value = 2.0
            m_fieldWeld.Value = 0.0
            m_allAround.Value = 0.0
            m_primarySideSupplementarySymbol.Value = 2.0
            m_tailNotes.Value = "2.0"
            m_tailNoteIsReference.Value = 0.0
            m_primarySideLength.Value = 2.0
            m_secondarySideLength.Value = 2.0
            m_primarySidePitch.Value = 2.0
            m_secondarySidePitch.Value = 2.0
            m_primarySideDiameter.Value = 2.0
            m_secondarySideDiameter.Value = 2.0
            m_primarySideContour.Value = 2.0
            m_secondarySideContour.Value = 2.0
            m_primarySideFinishMethod.Value = 2.0
            m_secondarySideFinishMethod.Value = 2.0
            m_primarySideRootOpening.Value = 2.0
            m_secondarySideRootOpening.Value = 2.0
            m_primarySideGrooveAngle.Value = 2.0
            m_secondarySideGrooveAngle.Value = 2.0
            m_primarySideNumberOfWelds.Value = 2.0
            m_secondarySideNumberOfWelds.Value = 2.0
            m_primarySideActualLegLength.Value = 2.0
            m_secondarySideActualLegLength.Value = 2.0
        End Sub
    End Class

    ''' <summary>
    ''' Parameter rule driving some of the attributes (long and short) on the IJUGenericProperties interface
    ''' </summary>
    ''' <remarks></remarks>
    <RuleVersion("1.0.0.0")> _
    <RuleInterface("ITestNo100ParameterRuleDrivingLong", "Parameter rule setting only the long and short attributes on interface IJUAGenericProperties")> _
    Public Class TestNo100ParameterRuleDrivingLong : Inherits ParameterRule
        <ControlledParameter(1, "Diameter", "Diameter")> _
        Public m_Diameter As ControlledParameterDouble
        <ControlledParameter(2, "OriginX", "Origin X")> _
        Public m_OriginX As ControlledParameterDouble
        <ControlledParameter(3, "OriginY", "Origin Y")> _
        Public m_OriginY As ControlledParameterDouble
        <ControlledParameter(4, "OriginZ", "Origin Z")> _
        Public m_OriginZ As ControlledParameterDouble
        <ControlledParameter(5, "Volume", "Volume")> _
        Public m_Volume As ControlledParameterDouble
        <ControlledParameter(6, "SurfaceArea", "Surface Area")> _
        Public m_SurfaceArea As ControlledParameterDouble
        <ControlledParameter(7, "Along", "A long")> _
        Public m_someLong As ControlledParameterDouble
        <ControlledParameter(8, "Ashort", "A short")> _
        Public m_someShort As ControlledParameterDouble
        <ControlledParameter(9, "PrimarySideSymbol", "Primar side symbol")> _
        Public m_primarySideSymbol As ControlledParameterDouble
        <ControlledParameter(10, "SecondarySideSymbol", "Secondary side symbol")> _
        Public m_secondarySideSymbol As ControlledParameterDouble
        <ControlledParameter(11, "PrimarySideGroove", "Primary side groove")> _
        Public m_primarySideGroove As ControlledParameterDouble
        <ControlledParameter(12, "SecondarySideGroove", "Secondary side groove")> _
        Public m_secondarySideGroove As ControlledParameterDouble
        <ControlledParameter(13, "PrimarySideGrooveSize", "Primary side groove size")> _
        Public m_primarySideGrooveSize As ControlledParameterDouble
        <ControlledParameter(14, "SecondarySideGrooveSize", "Secondary side groove size")> _
        Public m_secondarySideGrooveSize As ControlledParameterDouble
        <ControlledParameter(15, "PrimarySideActualThroatThickness", "Primary Side Actual Throat Thickness")> _
        Public m_primarySideActualThroatThickness As ControlledParameterDouble
        <ControlledParameter(16, "SecondarySideActualThroatThickness", "Secondary Side Actual Throat Thickness")> _
        Public m_secondarySideActualThroatThickness As ControlledParameterDouble
        <ControlledParameter(17, "PrimarySideNominalThroatThickness", "Primary Side Nominal Throat Thickness")> _
        Public m_primarySideNominalThroatThickness As ControlledParameterDouble
        <ControlledParameter(18, "SecondarySideNominalThroatThickness", "Secondary Side Nominal Throat Thickness")> _
        Public m_secondarySideNominalThroatThickness As ControlledParameterDouble
        <ControlledParameter(19, "FieldWeld", "Field Weld")> _
        Public m_fieldWeld As ControlledParameterDouble
        <ControlledParameter(20, "AllAround", "All Around")> _
        Public m_allAround As ControlledParameterDouble
        <ControlledParameter(21, "PrimarySideSupplementarySymbol", "Primary Side Supplementary Symbol")> _
        Public m_primarySideSupplementarySymbol As ControlledParameterDouble
        <ControlledParameter(22, "TailNotes", "Tail Notes")> _
        Public m_tailNotes As ControlledParameterString
        <ControlledParameter(23, "AnotherString", "Another string")> _
        Public m_someString As ControlledParameterString

        Public Overrides Sub Evaluate()
            m_Diameter.Value = 3.0
            m_OriginX.Value = 3.0
            m_OriginY.Value = 3.0
            m_OriginZ.Value = 3.0
            m_Volume.Value = 3.0
            m_SurfaceArea.Value = 3.0
            m_someLong.Value = 3.0
            m_someShort.Value = 3.0
            m_primarySideSymbol.Value = 3.0
            m_secondarySideSymbol.Value = 3.0
            m_primarySideGroove.Value = 3.0
            m_secondarySideGroove.Value = 3.0
            m_primarySideGrooveSize.Value = 3.0
            m_secondarySideGrooveSize.Value = 3.0
            m_primarySideActualThroatThickness.Value = 3.0
            m_secondarySideActualThroatThickness.Value = 3.0
            m_primarySideNominalThroatThickness.Value = 3.0
            m_secondarySideNominalThroatThickness.Value = 3.0
            m_fieldWeld.Value = 1.0
            m_allAround.Value = 1.0
            m_primarySideSupplementarySymbol.Value = 3.0
            m_tailNotes.Value = "3.0"
            m_someString.Value = "3.0"
        End Sub
    End Class

    ''' <summary>
    ''' Parameter rule driving some of the attributes (float and bool) on the IJUGenericProperties interface
    ''' </summary>
    ''' <remarks></remarks>
    <RuleVersion("1.0.0.0")> _
    <RuleInterface("ITestNo100ParameterRuleDrivingFloat", "Parameter rule setting only the float and bool attributes on interface IJUAGenericProperties")> _
    Public Class TestNo100ParameterRuleDrivingFloat : Inherits ParameterRule
        <ControlledParameter(1, "Diameter", "Diameter")> _
        Public m_Diameter As ControlledParameterDouble
        <ControlledParameter(2, "OriginX", "Origin X")> _
        Public m_OriginX As ControlledParameterDouble
        <ControlledParameter(3, "OriginY", "Origin Y")> _
        Public m_OriginY As ControlledParameterDouble
        <ControlledParameter(4, "OriginZ", "Origin Z")> _
        Public m_OriginZ As ControlledParameterDouble
        <ControlledParameter(5, "Volume", "Volume")> _
        Public m_Volume As ControlledParameterDouble
        <ControlledParameter(6, "SurfaceArea", "Surface Area")> _
        Public m_SurfaceArea As ControlledParameterDouble
        <ControlledParameter(7, "Afloat", "A float")> _
        Public m_someFloat As ControlledParameterDouble
        <ControlledParameter(8, "Abool", "A boolean")> _
        Public m_someBoolean As ControlledParameterDouble
        <ControlledParameter(9, "PrimarySideSymbol", "Primary side symbol")> _
        Public m_primarySideSymbol As ControlledParameterDouble
        <ControlledParameter(10, "SecondarySideSymbol", "Secondary side symbol")> _
        Public m_secondarySideSymbol As ControlledParameterDouble
        <ControlledParameter(11, "PrimarySideGroove", "Primary side groove")> _
        Public m_primarySideGroove As ControlledParameterDouble
        <ControlledParameter(12, "SecondarySideGroove", "Secondary side groove")> _
        Public m_secondarySideGroove As ControlledParameterDouble
        <ControlledParameter(13, "PrimarySideGrooveSize", "Primary side groove size")> _
        Public m_primarySideGrooveSize As ControlledParameterDouble
        <ControlledParameter(14, "SecondarySideGrooveSize", "Secondary side groove size")> _
        Public m_secondarySideGrooveSize As ControlledParameterDouble
        <ControlledParameter(15, "PrimarySideActualThroatThickness", "Primary Side Actual Throat Thickness")> _
        Public m_primarySideActualThroatThickness As ControlledParameterDouble
        <ControlledParameter(16, "SecondarySideActualThroatThickness", "Secondary Side Actual Throat Thickness")> _
        Public m_secondarySideActualThroatThickness As ControlledParameterDouble
        <ControlledParameter(17, "PrimarySideNominalThroatThickness", "Primary Side Nominal Throat Thickness")> _
        Public m_primarySideNominalThroatThickness As ControlledParameterDouble
        <ControlledParameter(18, "SecondarySideNominalThroatThickness", "Secondary Side Nominal Throat Thickness")> _
        Public m_secondarySideNominalThroatThickness As ControlledParameterDouble
        <ControlledParameter(19, "FieldWeld", "Field Weld")> _
        Public m_fieldWeld As ControlledParameterDouble
        <ControlledParameter(20, "AllAround", "All Around")> _
        Public m_allAround As ControlledParameterDouble
        <ControlledParameter(21, "PrimarySideSupplementarySymbol", "Primary Side Supplementary Symbol")> _
        Public m_primarySideSupplementarySymbol As ControlledParameterDouble
        <ControlledParameter(22, "TailNotes", "Tail Notes")> _
        Public m_tailNotes As ControlledParameterString
        <ControlledParameter(23, "TailNoteIsReference", "Tail Note Is Reference")> _
        Public m_tailNoteIsReference As ControlledParameterDouble
        <ControlledParameter(24, "PrimarySideLength", "Primary Side Length")> _
        Public m_primarySideLength As ControlledParameterDouble
        <ControlledParameter(25, "SecondarySideLength", "Secondary Side Length")> _
        Public m_secondarySideLength As ControlledParameterDouble
        <ControlledParameter(26, "PrimarySidePitch", "Primary Side Pitch")> _
        Public m_primarySidePitch As ControlledParameterDouble
        <ControlledParameter(27, "SecondarySidePitch", "SecondarySidePitch")> _
        Public m_secondarySidePitch As ControlledParameterDouble
        <ControlledParameter(28, "PrimarySideDiameter", "Primary Side Diameter")> _
        Public m_primarySideDiameter As ControlledParameterDouble
        <ControlledParameter(29, "SecondarySideDiameter", "Secondary Side Diameter")> _
        Public m_secondarySideDiameter As ControlledParameterDouble
        <ControlledParameter(30, "PrimarySideContour", "Primary Side Contour")> _
        Public m_primarySideContour As ControlledParameterDouble
        <ControlledParameter(31, "SecondarySideContour", "Secondary Side Contour")> _
        Public m_secondarySideContour As ControlledParameterDouble
        <ControlledParameter(32, "PrimarySideFinishMethod", "Primary Side Finish Method")> _
        Public m_primarySideFinishMethod As ControlledParameterDouble
        <ControlledParameter(33, "SecondarySideFinishMethod", "Secondary Side Finish Method")> _
        Public m_secondarySideFinishMethod As ControlledParameterDouble
        <ControlledParameter(34, "PrimarySideRootOpening", "Primary Side Root Opening")> _
        Public m_primarySideRootOpening As ControlledParameterDouble
        <ControlledParameter(35, "SecondarySideRootOpening", "Secondary Side Root Opening")> _
        Public m_secondarySideRootOpening As ControlledParameterDouble
        <ControlledParameter(36, "PrimarySideGrooveAngle", "Primary Side Groove Angle")> _
        Public m_primarySideGrooveAngle As ControlledParameterDouble
        <ControlledParameter(37, "SecondarySideGrooveAngle", "Secondary Side Groove Angle")> _
        Public m_secondarySideGrooveAngle As ControlledParameterDouble
        <ControlledParameter(38, "PrimarySideNumberOfWelds", "Primary Side Number Of Welds")> _
        Public m_primarySideNumberOfWelds As ControlledParameterDouble
        <ControlledParameter(39, "SecondarySideNumberOfWelds", "Secondary Side Number Of Welds")> _
        Public m_secondarySideNumberOfWelds As ControlledParameterDouble
        <ControlledParameter(40, "PrimarySideActualLegLength", "Primary Side Actual Leg Length")> _
        Public m_primarySideActualLegLength As ControlledParameterDouble
        <ControlledParameter(41, "SecondarySideActualLegLength", "Secondary Side Actual Leg Length")> _
        Public m_secondarySideActualLegLength As ControlledParameterDouble
        <ControlledParameter(42, "AnotherDouble", "Another double")> _
        Public m_anotherDouble As ControlledParameterDouble

        Public Overrides Sub Evaluate()
            m_Diameter.Value = 4.0
            m_OriginX.Value = 4.0
            m_OriginY.Value = 4.0
            m_OriginZ.Value = 4.0
            m_Volume.Value = 4.0
            m_SurfaceArea.Value = 4.0
            m_someFloat.Value = 4.0
            m_someBoolean.Value = 0.0
            m_primarySideSymbol.Value = 4.0
            m_secondarySideSymbol.Value = 4.0
            m_primarySideGroove.Value = 4.0
            m_secondarySideGroove.Value = 4.0
            m_primarySideGrooveSize.Value = 4.0
            m_secondarySideGrooveSize.Value = 4.0
            m_primarySideActualThroatThickness.Value = 4.0
            m_secondarySideActualThroatThickness.Value = 4.0
            m_primarySideNominalThroatThickness.Value = 4.0
            m_secondarySideNominalThroatThickness.Value = 4.0
            m_fieldWeld.Value = 0.0
            m_allAround.Value = 0.0
            m_primarySideSupplementarySymbol.Value = 4.0
            m_tailNotes.Value = "4.0"
            m_tailNoteIsReference.Value = 1.0
            m_primarySideLength.Value = 4.0
            m_secondarySideLength.Value = 4.0
            m_primarySidePitch.Value = 4.0
            m_secondarySidePitch.Value = 4.0
            m_primarySideDiameter.Value = 4.0
            m_secondarySideDiameter.Value = 4.0
            m_primarySideContour.Value = 4.0
            m_secondarySideContour.Value = 4.0
            m_primarySideFinishMethod.Value = 4.0
            m_secondarySideFinishMethod.Value = 4.0
            m_primarySideRootOpening.Value = 4.0
            m_secondarySideRootOpening.Value = 4.0
            m_primarySideGrooveAngle.Value = 4.0
            m_secondarySideGrooveAngle.Value = 4.0
            m_primarySideNumberOfWelds.Value = 4.0
            m_secondarySideNumberOfWelds.Value = 4.0
            m_primarySideActualLegLength.Value = 4.0
            m_secondarySideActualLegLength.Value = 4.0
            m_anotherDouble.Value = 4.0
        End Sub
    End Class

    ''' <summary>
    ''' CustomAssembly to test the parameter rules above
    ''' </summary>
    ''' <remarks></remarks>
    <CacheOption(CacheOptionType.Cached)> _
    Public Class TestNo100CustomAsm : Inherits CustomAssemblyDefinition
        <InputCatalogPart(1)> _
        Public m_catalogPart As InputCatalogPart
        <InputDouble(2, "Diameter", "Diameter", 5.0)> _
        Public m_dblDiameter As InputDouble
        <InputDouble(3, "OriginX", "OriginX", 5.0)> _
        Public m_dblOriginX As InputDouble
        <InputDouble(4, "OriginY", "Origin Y", 5.0)> _
        Public m_dblOriginY As InputDouble
        <InputDouble(5, "OriginZ", "Origin Z", 5.0)> _
        Public m_dblOriginZ As InputDouble
        <InputDouble(6, "Volume", "Volume", 5.0)> _
        Public m_dblVolume As InputDouble
        <InputDouble(7, "SurfaceArea", "Surface Area", 5.0)> _
        Public m_dblSurfaceArea As InputDouble
        <InputDouble(8, "Adouble", "A double", 5.0)> _
        Public m_dblSomeDouble As InputDouble
        <InputDouble(9, "Afloat", "A float", 5.0)> _
        Public m_dblSomeFloat As InputDouble
        <InputDouble(10, "Along", "A long", 5.0)> _
        Public m_dblSomeLong As InputDouble
        <InputDouble(11, "Ashort", "A short", 5.0)> _
        Public m_dblSomeShort As InputDouble
        <InputString(12, "Astring", "A string", "5.0")> _
        Public m_dblSomeString As InputString
        <InputDouble(13, "Abool", "A boolean", 5.0)> _
        Public m_dblSomeBoolean As InputDouble
        <InputDouble(14, "PrimarySideSymbol", "Primar side symbol", 5.0)> _
        Public m_dblPrimarySideSymbol As InputDouble
        <InputDouble(15, "SecondarySideSymbol", "Secondary side symbol", 5.0)> _
        Public m_dblSecondarySideSymbol As InputDouble
        <InputDouble(16, "PrimarySideGroove", "Primary side groove", 5.0)> _
        Public m_dblPrimarySideGroove As InputDouble
        <InputDouble(17, "SecondarySideGroove", "Secondary side groove", 5.0)> _
        Public m_dblSecondarySideGroove As InputDouble
        <InputDouble(18, "PrimarySideGrooveSize", "Primary side groove size", 5.0)> _
        Public m_dblPrimarySideGrooveSize As InputDouble
        <InputDouble(19, "SecondarySideGrooveSize", "Secondary side groove size", 5.0)> _
        Public m_dblSecondarySideGrooveSize As InputDouble
        <InputDouble(20, "PrimarySideActualThroatThickness", "Primary Side Actual Throat Thickness", 5.0)> _
        Public m_dblPrimarySideActualThroatThickness As InputDouble
        <InputDouble(21, "SecondarySideActualThroatThickness", "Secondary Side Actual Throat Thickness", 5.0)> _
        Public m_dblSecondarySideActualThroatThickness As InputDouble
        <InputDouble(22, "PrimarySideNominalThroatThickness", "Primary Side Nominal Throat Thickness", 5.0)> _
        Public m_dblPrimarySideNominalThroatThickness As InputDouble
        <InputDouble(23, "SecondarySideNominalThroatThickness", "Secondary Side Nominal Throat Thickness", 5.0)> _
        Public m_dblSecondarySideNominalThroatThickness As InputDouble
        <InputDouble(24, "FieldWeld", "Field Weld", 5.0)> _
        Public m_dblFieldWeld As InputDouble
        <InputDouble(25, "AllAround", "All Around", 5.0)> _
        Public m_dblAllAround As InputDouble
        <InputDouble(26, "PrimarySideSupplementarySymbol", "Primary Side Supplementary Symbol", 5.0)> _
        Public m_dblPrimarySideSupplementarySymbol As InputDouble
        <InputString(27, "TailNotes", "Tail Notes", "5.0")> _
        Public m_dblTailNotes As InputString
        <InputDouble(28, "TailNoteIsReference", "Tail Note Is Reference", 5.0)> _
        Public m_dblTailNoteIsReference As InputDouble
        <InputDouble(29, "PrimarySideLength", "Primary Side Length", 5.0)> _
        Public m_dblPrimarySideLength As InputDouble
        <InputDouble(30, "SecondarySideLength", "Secondary Side Length", 5.0)> _
        Public m_dblSecondarySideLength As InputDouble
        <InputDouble(31, "PrimarySidePitch", "Primary Side Pitch", 5.0)> _
        Public m_dblPrimarySidePitch As InputDouble
        <InputDouble(32, "SecondarySidePitch", "SecondarySidePitch", 5.0)> _
        Public m_dblSecondarySidePitch As InputDouble
        <InputDouble(33, "PrimarySideDiameter", "Primary Side Diameter", 5.0)> _
        Public m_dblPrimarySideDiameter As InputDouble
        <InputDouble(34, "SecondarySideDiameter", "Secondary Side Diameter", 5.0)> _
        Public m_dblSecondarySideDiameter As InputDouble
        <InputDouble(35, "PrimarySideContour", "Primary Side Contour", 5.0)> _
        Public m_dblPrimarySideContour As InputDouble
        <InputDouble(36, "SecondarySideContour", "Secondary Side Contour", 5.0)> _
        Public m_dblSecondarySideContour As InputDouble
        <InputDouble(37, "PrimarySideFinishMethod", "Primary Side Finish Method", 5.0)> _
        Public m_dblPrimarySideFinishMethod As InputDouble
        <InputDouble(38, "SecondarySideFinishMethod", "Secondary Side Finish Method", 5.0)> _
        Public m_dblSecondarySideFinishMethod As InputDouble
        <InputDouble(39, "PrimarySideRootOpening", "Primary Side Root Opening", 5.0)> _
        Public m_dblPrimarySideRootOpening As InputDouble
        <InputDouble(40, "SecondarySideRootOpening", "Secondary Side Root Opening", 5.0)> _
        Public m_dblSecondarySideRootOpening As InputDouble
        <InputDouble(41, "PrimarySideGrooveAngle", "Primary Side Groove Angle", 5.0)> _
        Public m_dblPrimarySideGrooveAngle As InputDouble
        <InputDouble(42, "SecondarySideGrooveAngle", "Secondary Side Groove Angle", 5.0)> _
        Public m_dblSecondarySideGrooveAngle As InputDouble
        <InputDouble(43, "PrimarySideNumberOfWelds", "Primary Side Number Of Welds", 5.0)> _
        Public m_dblPrimarySideNumberOfWelds As InputDouble
        <InputDouble(44, "SecondarySideNumberOfWelds", "Secondary Side Number Of Welds", 5.0)> _
        Public m_dblSecondarySideNumberOfWelds As InputDouble
        <InputDouble(45, "PrimarySideActualLegLength", "Primary Side Actual Leg Length", 5.0)> _
        Public m_dblPrimarySideActualLegLength As InputDouble
        <InputDouble(46, "SecondarySideActualLegLength", "Secondary Side Actual Leg Length", 5.0)> _
        Public m_dblSecondarySideActualLegLength As InputDouble
        <InputString(47, "AnotherString", "A string", "5.0")> _
        Public m_dblAnotherString As InputString
        <InputDouble(48, "AnotherDouble", "Another double", 5.0)> _
        Public m_dblAnotherDouble As InputDouble

        Private Const CONST_SymbolSphere As String = "SymbolSphere"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_SymbolSphere, "Symbol Sphere")> _
        Public m_physicalAspect As AspectDefinition

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            Dim oSP3DConnection As SP3DConnection

            ' Get the connection
            oSP3DConnection = Me.OccurrenceConnection

            Dim dblBaseRadius As Double
            dblBaseRadius = m_dblDiameter.Value / 2.0

            '=================================================
            ' Construct base
            '=================================================
            Dim objBase As Sphere3d

            objBase = New Sphere3d(oSP3DConnection, New Position(m_dblOriginX.Value, m_dblOriginY.Value, m_dblOriginZ.Value), dblBaseRadius, True)

            m_physicalAspect.Outputs(CONST_SymbolSphere) = objBase
        End Sub

        ''' <summary>
        ''' Evaluate assembly outputs
        ''' </summary>
        ''' <remarks></remarks>
        Public Overrides Sub EvaluateAssembly()
            ' Verify that we can at least access the occurrence during the evaluate
            Dim sString As String = Occurrence.ToString()
        End Sub
    End Class

    ''' <summary>
    ''' CustomAssembly to test the TestNo100 parameter rulesA
    ''' </summary>
    ''' <remarks></remarks>
     <SymbolVersion("1.0.0.0")> _
    Public Class TestNo100CustomAsmA : Inherits TestNo100CustomAsm
    End Class

    ''' <summary>
    ''' CustomAssembly to test the TestNo100 parameter rulesB
    ''' </summary>
    ''' <remarks></remarks>
    <SymbolVersion("1.0.0.0")> _
    Public Class TestNo100CustomAsmB : Inherits TestNo100CustomAsm
    End Class


    <CacheOption(CacheOptionType.NonCached)> _
    <SymbolVersion("1.0.0.0")> _
    Public Class SymbolWithInputObjects : Inherits CustomAssemblyDefinition
        '   1. "Bottom Plane"
        '   2. "OriginX"
        '   3. "OriginY"
        '   4. "OriginZ"
        '   5. "Diameter"
        '   6. "Height"

        <InputObject(1, "BottomPlane", "Bottom Plane")> _
        Public m_bottomPlane As InputObject
        <InputDouble(2, "OriginX", "OriginX", 0.0)> _
        Public m_dblOriginX As InputDouble
        <InputDouble(3, "OriginY", "OriginY", 0.0)> _
        Public m_dblOriginY As InputDouble
        <InputDouble(4, "OriginZ", "OriginZ", 0.0)> _
        Public m_dblOriginZ As InputDouble
        <InputDouble(5, "Diameter", "Diameter", 0.3)> _
        Public m_dblDiameter As InputDouble
        <InputDouble(6, "Height", "Height", 1.0)> _
        Public m_dblHeight As InputDouble

        Private Const CONST_Cylinder As String = "Cylinder"

        <Aspect("Physical", "Physical Aspect", AspectID.SimplePhysical)> _
        <SymbolOutput(CONST_Cylinder, "Cylinder")> _
        Public m_physicalAspect As AspectDefinition

        ''' <summary>
        ''' Construct symbol outputs/aspects
        ''' </summary>
        ''' <remarks></remarks>
        Protected Overrides Sub ConstructOutputs()
            ' Get the connection
            Dim oSP3DConnection As SP3DConnection = Me.OccurrenceConnection

            Dim bottomPlane As Plane3d = DirectCast(m_bottomPlane.Value, Plane3d)

            Dim originPosition As Position = bottomPlane.ProjectPoint(New Position(m_dblOriginX.Value, m_dblOriginY.Value, m_dblOriginZ.Value))
            Dim vectorNormal As Vector = New Vector(bottomPlane.Normal)
            vectorNormal.Scale(m_dblHeight.Value)
            Dim endPosition As Position = New Position(vectorNormal.X + originPosition.X, vectorNormal.Y + originPosition.Y, vectorNormal.Z + originPosition.Z)
            Dim vectorInPlane As Vector = New Vector(bottomPlane.VDirection)
            vectorInPlane.Scale(m_dblDiameter.Value / 2.0)
            Dim bottomPosition As Position = New Position(vectorInPlane.X + originPosition.X, vectorInPlane.Y + originPosition.Y, vectorInPlane.Z + originPosition.Z)
            Dim topPosition As Position = New Position(vectorNormal.X + bottomPosition.X, vectorNormal.Y + bottomPosition.Y, vectorNormal.Z + bottomPosition.Z)

            '=================================================
            ' Construct cylinder
            '=================================================
            Dim cylinder As Cone3d = New Cone3d(oSP3DConnection, originPosition, endPosition, bottomPosition, topPosition, True)
            m_physicalAspect.Outputs(CONST_Cylinder) = cylinder
        End Sub
    End Class
End Namespace
