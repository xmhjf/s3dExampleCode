//-------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  ChamfersResourceIds.cs
//
//Abstract
//	ChamfersResourceIds is chamfer custom assemblies resource identifiers.
//-------------------------------------------------------------------------------------------------------

namespace Ingr.SP3D.Content.Structure.Services
{
    /// <summary>
    /// Chamfer custom assemblies resource identifiers.
    /// </summary>
    internal static class ChamfersResourceIds
    {
        /// <summary>
        /// Chamfers localizer resource.
        /// </summary>
        internal const string ChamfersResources = "Ingr.SP3D.Content.Structure.Resources.Chamfers";
        /// <summary>
        /// Chamfers localizer assembly.
        /// </summary>
        internal const string ChamfersAssembly = "Chamfers";
        /// <summary>
        /// Unexpected error while getting the selections from SingleSided Chamfer selector rule.
        /// </summary>
        internal const int ToDoChamferSelections = 101;
        /// <summary>
        /// Error while evaluating chamfer parameter rule.
        /// </summary>
        internal const int ToDoChamferParameterRule = 102;
        /// <summary>
        /// Unexpected error while copying answers of chamfer to the physical connection.
        /// </summary>
        internal const int ErrCopyAnswers = 103;
        /// <summary>
        /// Unexpected exception in getting the thicknessDifference of plate part over member flange.
        /// </summary>
        internal const int GetThicknessDifferenceError = 104;
        /// <summary>
        /// Error while creating the physical connection.
        /// </summary>
        internal const int CreatePhysicalConnectionError = 105;
        /// <summary>
        /// Unexpected error while evaluating the chamfer physical connection definition.
        /// </summary>
        internal const int ToDoEvaluateAssembly = 106;
    }
}