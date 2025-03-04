//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   ProcessDefinition is a .NET Symbol Definition for TemplateSet on Profile Face . 
//                 
//
//      Author:  
//
//      History:
//      August 15th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.ProfileFace
{
    /// <summary>
    /// Process Symbol Definition for TemplateSet on Profile Face.
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    public class ProcessDefinition : ManufacturingSymbolDefinition
    {
        //Define 12 inputs
        #region "Definition of Inputs"
        /// <summary>
        /// Minimum Height
        /// </summary>
        [InputDouble(1, SymbolsConstants.ProcessMinHeight, SymbolsConstants.ProcessMinHeight, 5100)] //Fixed
        public InputDouble minHeight;
        /// <summary>
        /// Maximum Height
        /// </summary>
        [InputDouble(2, SymbolsConstants.ProcessMaxHeight, SymbolsConstants.ProcessMaxHeight, 5110)] //Fixed
        public InputDouble maxHeight;
        /// <summary>
        /// Extension 
        /// </summary>
        [InputDouble(3, SymbolsConstants.ProcessExtension, SymbolsConstants.ProcessExtension, 5120)] //Fixed
        public InputDouble extension;
        /// <summary>
        /// Side
        /// </summary>
        [InputDouble(4, SymbolsConstants.ProcessSide, SymbolsConstants.ProcessSide, 5130)] //BaseSide
        public InputDouble side;
        /// <summary>
        /// Type
        /// </summary>
        [InputDouble(5, SymbolsConstants.ProcessType, SymbolsConstants.ProcessType, 5140)] //Frame
        public InputDouble type;
        /// <summary>
        /// Orientation
        /// </summary>
        [InputDouble(6, SymbolsConstants.ProcessOrientation, SymbolsConstants.ProcessOrientation, 5150)] //AlongFrame
        public InputDouble orientation;
        /// <summary>
        /// Position
        /// </summary>
        [InputDouble(7, SymbolsConstants.ProcessPosition, SymbolsConstants.ProcessPosition, 5160)] //FramesAndEdges
        public InputDouble position;
        /// <summary>
        /// Base Plane
        /// </summary>
        [InputDouble(8, SymbolsConstants.ProcessBasePlane, SymbolsConstants.ProcessBasePlane, 5170)] //Average Of Corners Plane
        public InputDouble basePlane;
        /// <summary>
        /// Direction
        /// </summary>
        [InputDouble(9, SymbolsConstants.ProcessDirection, SymbolsConstants.ProcessDirection, 5180)] //Transversal
        public InputDouble direction;
        /// <summary>
        /// Template Service
        /// </summary>
        [InputDouble(10, SymbolsConstants.ProcessTemplateService, SymbolsConstants.ProcessTemplateService, 5190)] //Default
        public InputDouble templateService;
        /// <summary>
        /// User Defined Values
        /// </summary>
        [InputDouble(11, SymbolsConstants.ProcessUserDefinedValues, SymbolsConstants.ProcessUserDefinedValues, 53100)] //Default
        public InputDouble userDefinedValues;
        /// <summary>
        /// Template Naming
        /// </summary>
        [InputDouble(12, SymbolsConstants.ProcessTemplateNaming, SymbolsConstants.ProcessTemplateNaming, 53110)] //TemplateNaming
        public InputDouble templateNaming;

        #endregion

        //Define Aspect
        #region Define aspect
        /// <summary>
        /// The simple physical
        /// </summary>
        [Aspect(SymbolsConstants.ProcessAspect, SymbolsConstants.ProcessAspectDesc, AspectID.SimplePhysical)]
        public AspectDefinition simplePhysical;
        #endregion
    }
}