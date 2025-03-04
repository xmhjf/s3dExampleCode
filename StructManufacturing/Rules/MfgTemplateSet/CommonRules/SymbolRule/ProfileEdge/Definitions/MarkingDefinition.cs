//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   MarkingDefinition is a .NET Symbol Definition for TemplateSet on Plate
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
    /// Marking Symbol Definition for TemplateSet on Plate.
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    public class MarkingDefinition : ManufacturingSymbolDefinition
    {
        //Define 11 inputs
        #region "Definition of Inputs"
        /// <summary>
        /// Base Control Line Mark
        /// </summary>
        [InputDouble(1, SymbolsConstants.MarkingBaseCtlLineMark, SymbolsConstants.MarkingBaseCtlLineMark, 7300)] //Apply
        public InputDouble baseControlMark;
        /// <summary>
        /// Seam Marks
        /// </summary>
        [InputDouble(2, SymbolsConstants.MarkingSeamMark, SymbolsConstants.MarkingSeamMark, 7310)] //Apply
        public InputDouble seamMark;
        /// <summary>
        /// Ship Dir Mark
        /// </summary>
        [InputDouble(3, SymbolsConstants.MarkingShipDirectionMark, SymbolsConstants.MarkingShipDirectionMark, 7320)] //Apply
        public InputDouble shipDirMark;
        /// <summary>
        /// Custom Mark
        /// </summary>
        [InputDouble(4, SymbolsConstants.MarkingCustomMark, SymbolsConstants.MarkingCustomMark, 7331)] //Ignore
        public InputDouble customMark;
        #endregion
        //Define Aspect
        #region Define aspect
        /// <summary>
        /// The simple physical
        /// </summary>
        [Aspect(SymbolsConstants.MarkingAspect, SymbolsConstants.MarkingAspectDesc, AspectID.SimplePhysical)]
        public AspectDefinition simplePhysical;
        #endregion
    }
}