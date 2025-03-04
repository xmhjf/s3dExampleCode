//-------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  EndCutsResourceIds.cs
//
//Abstract
//	EndCutsResourceIds is EndCuts custom assemblies resource identifiers.
//-------------------------------------------------------------------------------------------------------
namespace Ingr.SP3D.Content.Structure.Services
{
    /// <summary>
    /// EndCuts custom assemblies resource identifiers.
    /// </summary>
    internal static class EndCutsResourceIds
    {
        /// <summary>
        /// EndCuts localizer resource.
        /// </summary>
        internal const string EndCutsResource = "Ingr.SP3D.Content.Structure.Resources.EndCuts";

        /// <summary>
        /// EndCuts localizer assembly.
        /// </summary>
        internal const string EndCutsAssembly = "EndCuts";

        /// <summary>
        /// Unexpected error while getting the selections from {0}
        /// </summary>
        internal const int ToDoSelections = 101;

        /// <summary>
        /// Unexpected error while setting the default answer in {0}
        /// </summary>
        internal const int ToDoDefaultAnswer = 102;

        /// <summary>
        /// Unexpected error while evaluating the custom assembly of {0}
        /// </summary>
        internal const int ToDoEvaluateAssembly = 103;

        /// <summary>
        /// Unexpected error while inputs replaced the custom assembly of {0}
        /// </summary>
        internal const int ToDoInputsReplaced = 104;

        /// <summary>
        /// The existing port geometry must be point for replacing the inputs.
        /// </summary>
        internal const int ToDoPortGeometryNotPoint = 105;

        /// <summary>
        /// Unexpected error while evaluating the {0}
        /// </summary>
        internal const int ToDoParameterRule = 106;

        /// <summary>
        /// Unexpected error occured while evaluating the member web and flangecut custom assembly definition.
        /// </summary>
        internal const int MbrWebAndFlangeCutDefUnExpectedError = 107;

        /// <summary>
        /// Unexpected error on evaluating the conditional methods while replacing the inputs.
        /// </summary>
        internal const int ErrOnReplacingInputs = 108;

        /// <summary>
        /// Unexpected error while evaluating the StiffenerWebAndFlangeCutDefinition.
        /// </summary>
        internal const int ToDoStiffenerWebAndFlangeCutDefinition = 109;

        /// <summary>
        /// Unexpected error while creating the corner Feature.
        /// </summary>
        internal const int ErrCreateCornerFeature = 110;

        /// <summary>
        /// Unexpected failure while getting the bounded and bounding sub port type
        /// </summary>
        internal const int ErrInBoundingAndBoundedSubPortType = 111;
    }
}
