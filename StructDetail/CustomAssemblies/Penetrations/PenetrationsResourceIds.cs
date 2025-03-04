//-------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  PenetrationsResourceIds.cs
//
//Abstract
//	PenetrationsResourceIds is penetrations custom assemblies resource identifiers.
//-------------------------------------------------------------------------------------------------------
namespace Ingr.SP3D.Content.Structure.Services
{
    /// <summary>
    /// Penetrations custom assemblies resource identifiers.
    /// </summary>
    internal static class PenetrationsResourceIds
    {
        /// <summary>
        /// Penetrations localizer resource.
        /// </summary>
        internal const string PenetrationsResource = "Ingr.SP3D.Content.Structure.Resources.Penetrations";

        /// <summary>
        /// Penetrations localizer assembly.
        /// </summary>
        internal const string PenetrationsAssembly = "Penetrations";

        /// <summary>
        /// Unexpected error while getting the selections from {0}
        /// </summary>
        internal const int ToDoSelections = 101;

        /// <summary>
        /// Error while setting the default answer.
        /// </summary>
        internal const int ToDoDefaultAnswer = 102;

        /// <summary>
        /// Unexpected error while evaluating the {0}
        /// </summary>
        internal const int ToDoParameterRule = 103;

        /// <summary>
        /// Unexpected error while evaluating the custom assembly of {0}
        /// </summary>
        internal const int ToDoDefinition = 104;

        /// <summary>
        /// Unable to get penetrating section type name.
        /// </summary>
        internal const int ErrGetPenetratingSectionTypeName = 105;

        /// <summary>
        /// No equivalent Collar item for Profile.
        /// </summary>
        internal const int ErrNoCollarItemForProfile = 106;

        /// <summary>
        /// A new symbol have to create that does not include a radius.
        /// </summary>
        internal const int ErrNoCollarItemForPlate = 107;

        /// <summary>
        /// Invalid penetrating section type name.
        /// </summary>
        internal const int ErrInvalidSectionTypeName = 108;

        /// <summary>
        /// Invalid slot type.
        /// </summary>
        internal const int ErrInvalidSlotType = 109;

        /// <summary>
        /// $1 should not be evaluated for a profile with $2 section type name. Check if the correct catalog part was selected by the selection rule.
        /// </summary>
        internal const int ErrInvalidParamRule = 110;

        /// <summary>
        /// Interface description for IJPlate. 
        /// </summary>
        internal const int Thickness = 111;

        /// <summary>
        /// Interface description for IJStructureMaterial. 
        /// </summary>
        internal const int MatAndGrade = 112;

        /// <summary>
        /// Interface description for IJCollarPart. 
        /// </summary>
        internal const int SideOfPlate = 113;

    }
}
