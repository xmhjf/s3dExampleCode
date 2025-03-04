//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   ProcessDefinition is a .NET Symbol Definition for TemplateSet on ProfileEdge
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

namespace Ingr.SP3D.Content.Manufacturing.TemplateSet.ProfileEdge
{
    /// <summary>
    /// Process Symbol Definition for TemplateSet on ProfileEdge 
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    public class ProcessDefinition : ManufacturingSymbolDefinition
    {
        //Define 12 inputs
        #region "Definition of Inputs"
        /// <summary>
        /// Type
        /// </summary>
        [InputDouble(1, SymbolsConstants.ProcessType, SymbolsConstants.ProcessType, 7200)] //Standard
        public InputDouble type;
        /// <summary>
        /// Side
        /// </summary>
        [InputDouble(2, SymbolsConstants.ProcessSide, SymbolsConstants.ProcessSide, 7213)] //Bottom
        public InputDouble side;
        /// <summary>
        /// Offset
        /// </summary>
        [InputDouble(3, SymbolsConstants.ProcessOffset, SymbolsConstants.ProcessOffset, 7220)] //No_Offset
        public InputDouble offset;
        /// <summary>
        /// Extension
        /// </summary>
        [InputDouble(4, SymbolsConstants.ProcessExtension,SymbolsConstants.ProcessExtension, 7230)] //BothEnds_Fixed
        public InputDouble extension;
        /// <summary>
        /// Minimum Height
        /// </summary>
        [InputDouble(5, SymbolsConstants.ProcessMinHeight, SymbolsConstants.ProcessMinHeight, 7270)] //Fixed
        public InputDouble minHeight;
        /// <summary>
        /// Maximum Height
        /// </summary>
        [InputDouble(6, SymbolsConstants.ProcessMaxHeight, SymbolsConstants.ProcessMaxHeight, 7280)] //Fixed
        public InputDouble maxHeight;
        /// <summary>
        /// Service
        /// </summary>
        [InputDouble(7, SymbolsConstants.ProcessService, SymbolsConstants.ProcessService, 7240)] //Default
        public InputDouble service;
        /// <summary>
        /// User Defined Value
        /// </summary>
        [InputDouble(8, SymbolsConstants.ProcessUsrDefinedValue, SymbolsConstants.ProcessUsrDefinedValue, 7250)] //Default
        public InputDouble userDefVal;
        /// <summary>
        /// Template Name
        /// </summary>
        [InputDouble(9, SymbolsConstants.ProcessTemplateName, SymbolsConstants.ProcessTemplateName, 7260)] //Default
        public InputDouble templateName;
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