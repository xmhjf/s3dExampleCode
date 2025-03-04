//-------------------------------------------------------------------------------------------------------
//Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
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
    /// Hold resource string data for AssemblyConnections custom assembly definitions
    /// </summary>
    internal static class MemberAssemblyConnectionsResourceIds
    {
        /// <summary>
        /// Connections localizer resource.
        /// </summary>
        internal const string DEFAULT_RESOURCE = "Ingr.SP3D.Content.Structure.Resources.MemberAssemblyConnections";

        /// <summary>
        /// Connections localizer assembly.
        /// </summary>
        internal const string DEFAULT_ASSEMBLY = "MemberAssemblyConnections";

        /// <summary>
        /// No bounded ports associated with the AssemblyConnection.
        /// </summary>
        internal const int ErrBoundedPortCount = 101;

        /// <summary>
        /// No bounding ports associated with the AssemblyConnection.
        /// </summary>
        internal const int ErrBoundingPortCount = 110;

        /// <summary>
        /// Unexpected error while getting the selections from StiffenerEndByPlate AssemblyConnection selector rule.
        /// </summary>
        internal const int ToDoSelections = 101;

        /// <summary>
        /// Unexpected error while setting the default answer of StiffenerEndByPlate AssemblyConnection selector rule.
        /// </summary>
        internal const int ToDoDefaultAnswer = 102;

        /// <summary>
        /// Unexpected error while evaluating of Fitted AssemblyConnection custom assembly definition.
        /// </summary>
        internal const int ToDoEvaluateAssembly = 103;

        /// <summary>
        /// Bounded MemberPart section type is Invalid.
        /// </summary>
        internal const int InValidBoundedSection = 104;

        /// <summary>
        /// Bounding MemberPart section type is Invalid.
        /// </summary>
        internal const int InValidBoundingSection = 105;

        /// <summary>
        /// Bounded port already in relation with other assembly connection.
        /// </summary>
        public const int InvalidInputForBounded = 106;

        /// <summary>
        /// Bounding port already in relation with other assembly connection.
        /// </summary>
        public const int InvalidInputForBounding = 107;

        /// <summary>
        /// One of the input ports is not a MemberPartAxisPort.
        /// </summary>
        internal const int ErrInvalidPortType = 108;

        /// <summary>
        /// Unexpected error while getting the selections from Fitted AssemblyConnection selector rule
        /// </summary>
        internal const int ErrUnExpectedInSelectionRule = 109;
    }
}