//--------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  MemberAssemblyConnectionShape.cs
//
//Abstract
//--------------------------------------------------------------------------------------------------------
namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Holds the information about answers set on the assembly connection for the shapes at top and bottom edges.
    /// </summary>
    internal class MemberAssemblyConnectionShape
    {
        /// <summary>
        /// Gets or sets the bottom shape.
        /// </summary>
        internal int BottomShape { get; set; }

        /// <summary>
        /// Gets or sets the top shape.
        /// </summary>
        internal int TopShape { get; set; }

        /// <summary>
        /// Gets or sets the face top inside corner shape.
        /// </summary>
        internal int FaceTopInsideCornerShape { get; set; }

        /// <summary>
        /// Gets or sets the face bottom inside corner shape.
        /// </summary>
        internal int FaceBottomInsideCornerShape { get; set; }

        /// <summary>
        /// Gets or sets the bottom answer.
        /// </summary>
        internal string BottomAnswer { get; set; }

        /// <summary>
        /// Gets or sets code list short descrition for the top answer.
        /// </summary>
        internal string TopAnswer { get; set; }

        /// <summary>
        /// Gets or sets the face top inside corner answer.
        /// </summary>
        internal string FaceTopInsideCornerAnswer { get; set; }

        /// <summary>
        /// Gets or sets the face bottom inside corner answer.
        /// </summary>
        internal string FaceBottomInsideCornerAnswer { get; set; }
    }
}