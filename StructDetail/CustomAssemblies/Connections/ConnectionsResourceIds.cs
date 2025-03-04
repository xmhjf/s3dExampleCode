//-------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  ConnectionsResourceIds.cs
//
//Abstract
//	ConnectionsResourceIds contains all resource ids used in Connections custom assembly definitions.
//-------------------------------------------------------------------------------------------------------
namespace Ingr.SP3D.Content.Structure.Services
{
    /// <summary>
    /// Hold resouce string data for AssemblyConnections custom assembly definitions
    /// </summary>
    internal static class ConnectionsResourceIdentifiers
    {
        /// <summary>
        /// Connections localizer resource.
        /// </summary>
        internal const string ConnectionsResource = "Ingr.SP3D.Content.Structure.Resources.Connections";

        /// <summary>
        /// Connections localizer assembly.
        /// </summary>
        internal const string ConnectionsAssembly = "Connections";

        /// <summary>
        /// Unexpected error while getting the selections from StiffenerEndByPlate AssemblyConnection selector rule.
        /// </summary>
        internal const int ToDoSelections = 101;

        /// <summary>
        /// Unexpected error while setting the default answer of StiffenerEndByPlate AssemblyConnection selector rule.
        /// </summary>
        internal const int ToDoDefaultAnswer = 102;

        /// <summary>
        /// Unexpected error while evaluating of StiffenerEndToPlateFace AssemblyConnection custom assembly definition.
        /// </summary>
        internal const int ToDoEvaluateAssembly = 103;

        /// <summary>
        /// AssemblyConnection has more than one bounding port.
        /// </summary>
        internal const int ErrBoundingPortCount = 104;

        /// <summary>
        /// No bounded ports associated with the AssemblyConnection.
        /// </summary>
        internal const int ErrBoundedPortCount = 105;

        /// <summary>
        /// One of the input ports is not a TopologyPort.
        /// </summary>
        internal const int ErrACInputs = 106;

        /// <summary>
        /// Invalid bounding port, ContextId should be Base or Offset or Lateral.
        /// </summary>
        internal const int ErrInvalidBoundingPort = 107;
        
        /// <summary>
        /// Unexpected error while creating the CustomPlatePart.
        /// </summary>
        internal const int ErrCustomPlatePart = 110;
        /// <summary>
        /// Unexpected error while getting the selections from PhysicalConnection selector rule.
        /// </summary>
        internal const int ToDoPhysicalConnectionSelections = 111;

        /// <summary>
        /// Error while evaluating PhysicalConnection parameter rule.
        /// </summary>
        internal const int ToDoPhysicalConnectionParameterRule = 112;

        /// <summary>
        /// Error in getting the Category.
        /// </summary>
        internal const int ErrInvalidCategory = 113;

        /// <summary>
        /// Error in getting the Profile.
        /// </summary>
        internal const int ErrInvalidProfile = 114;

        /// <summary>
        /// Error in getting the SquareTrim.
        /// </summary>
        internal const int ErrSquareTrim = 115;

        /// <summary>
        /// Error in getting the EndCutType.
        /// </summary>
        internal const int ErrEndCutType = 116;

        /// <summary>
        /// Error  while evaluating PhysicalConnectionCustomAssemblyDefinition
        /// </summary>
        internal const int ErrEvaluateAssembly = 117;
    }
}
