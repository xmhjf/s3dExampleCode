//-------------------------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  SurfaceTrimSelectorRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\SmartPlantStructure\Symbols\Release\SPSFeatureMacros.dll
//  Original Class Name: ‘SurfaceTrimSel’ in VB content
//
//Abstract
//	SurfaceTrimSelectorRule is a .NET selection rule for surface trim feature.
//-------------------------------------------------------------------------------------------------------------------------------
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using System.Collections.Generic;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Selector for Surface Trim feature, which selects the list of available items in the context of the Surface Trim feature.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [DefaultLocalizer(FeaturesResourceIds.DEFAULT_RESOURCE, FeaturesResourceIds.DEFAULT_ASSEMBLY)]
    [RuleInterface(StructureCustomAssembliesConstants.IASPSFeatureMacros_SurfaceTrimSel, StructureCustomAssembliesConstants.IASPSFeatureMacros_SurfaceTrimSel)]
    public class SurfaceTrimSelectorRule : SelectorRule
    {
        #region Public override properties and methods

        //=======================================================================================================
        //DefinitionName/ProgID of this symbol is "Features,Ingr.SP3D.Content.Structure.SurfaceTrimSelectorRule"
        //=======================================================================================================

        /// <summary>
        /// Processes the rules and returns the collection of named part items that are possible choices based on provided inputs.
        /// If multiple choices are provided by the selection rule, the first element in this collection will be the default selection made.
        /// </summary>
        public override ReadOnlyCollection<string> Selections
        {
            get
            {
                // Only one item available
                List<string> list = new List<string>() { StructureCustomAssembliesConstants.SurfaceTrim };
                return new ReadOnlyCollection<string>(list);
            }
        }

        #endregion Public override properties and methods
    }
}