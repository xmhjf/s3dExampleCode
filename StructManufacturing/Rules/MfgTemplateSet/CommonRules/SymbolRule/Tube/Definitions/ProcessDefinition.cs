//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   ProcessDefinition is a .NET Symbol Definition for TemplateSet on Plate 
//                 
//
//      Author:  
//
//      History:
//      August 5th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.Tube
{
    /// <summary>
    /// Process Symbol Definition for TemplateSet on Tube.
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    public class ProcessDefinition : ManufacturingSymbolDefinition
    {
        //Define 11 inputs
        #region "Definition of Inputs"
        /// <summary>
        /// Minimum Height
        /// </summary>
        [InputDouble(1, SymbolsConstants.ProcessMinHeight, SymbolsConstants.ProcessMinHeight, 7100)] //Fixed
        public InputDouble minHeight;
        /// <summary>
        /// Maximum Height
        /// </summary>
        [InputDouble(2, SymbolsConstants.ProcessMaxHeight, SymbolsConstants.ProcessMaxHeight, 7100)] //Fixed
        public InputDouble maxHeight;
        /// <summary>
        /// Extension
        /// </summary>
        [InputDouble(3, SymbolsConstants.ProcessExtension, SymbolsConstants.ProcessExtension, 7100)] //ExtnRad90
        public InputDouble extension;
        /// <summary>
        /// Side
        /// </summary>
        [InputDouble(4, SymbolsConstants.ProcessSide, SymbolsConstants.ProcessSide, 7100)] //Outer
        public InputDouble side;
        /// <summary>
        /// Side End
        /// </summary>
        [InputDouble(5, SymbolsConstants.ProcessSideEnd, SymbolsConstants.ProcessSideEnd, 7100)] //Base
        public InputDouble sideEnd;
        /// <summary>
        /// ProcessType
        /// </summary>
        [InputDouble(6, SymbolsConstants.ProcessType, SymbolsConstants.ProcessType, 7100)] //ShortDistanceBCL
        public InputDouble type;
        /// <summary>
        /// BasePlane
        /// </summary>
        [InputDouble(7, SymbolsConstants.ProcessBasePlane, SymbolsConstants.ProcessBasePlane, 7100)] //BySystem
        public InputDouble basePlane;
        /// <summary>
        /// Tube Service 
        /// </summary>
        [InputDouble(8, SymbolsConstants.ProcessTemplateService, SymbolsConstants.ProcessTemplateService, 7100)] //Default
        public InputDouble service;
        /// <summary>
        /// UserDefinedValues
        /// </summary>
        [InputDouble(9, SymbolsConstants.ProcessUserDefinedValues, SymbolsConstants.ProcessUserDefinedValues, 7100)] //Default
        public InputDouble userDefinedValues;
        /// <summary>
        /// Template Naming
        /// </summary>
        [InputDouble(10, SymbolsConstants.ProcessTemplateNaming, SymbolsConstants.ProcessTemplateNaming, 7100)] //Default
        public InputDouble templateNaming;
        /// <summary>
        /// Bevel
        /// </summary>
        [InputDouble(11, SymbolsConstants.ProcessBevel, SymbolsConstants.ProcessBevel, 7101)] //Fixed
        public InputDouble bevel;
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