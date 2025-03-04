//-------------------------------------------------------------------------------------------------------
//Copyright 1992 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  FeaturesResourceIds.cs
//
//Abstract
//	FeaturesResourceIds contains all resource ids used in custom assembly definitions for all features.
//-------------------------------------------------------------------------------------------------------
namespace Ingr.SP3D.Content.Structure.Services
{
    /// <summary>
    /// Hold resource string data for custom assembly definitions for all features.
    /// </summary>
    internal static class FeaturesResourceIds
    {
        /// <summary>
        /// Connections localizer resource.
        /// </summary>
        internal const string DEFAULT_RESOURCE = "Ingr.SP3D.Content.Structure.Resources.Features";

        /// <summary>
        /// Connections localizer assembly.
        /// </summary>
        internal const string DEFAULT_ASSEMBLY = "Features";

        /// <summary>
        /// Unexpected error while inputs replaced the custom assembly of {0}
        /// </summary>
        internal const int ErrToDoInputsReplaced = 1;

        /// <summary>
        /// Unexpected error while evaluating custom assembly of {0}
        /// </summary>
        internal const int ErrToDoEvaluateAssembly = 2;

        /// <summary>
        /// Error when InsideFlangeClearance value of cope feature is set to more than half of bounding member's depth.
        /// </summary>
        internal const int ErrInvalidInsideFlangeClearance = 3;

        /// <summary>
        /// The assembly output used as cutback surface must be ISurface.
        /// </summary>
        internal const int ErrInvalidCutbackSurfaceType = 4;
    }
}