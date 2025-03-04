//-------------------------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  CopeSelectorRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\SmartPlantStructure\Symbols\Release\SPSFeatureMacros.dll
//  Original Class Name: ‘CopeSel’ in VB content
//
//Abstract
//	CopeSelectorRule is a .NET selection rule for cope feature.
//-------------------------------------------------------------------------------------------------------------------------------
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Selector for cope feature, which selects the list of available items in the context of the corner cope feature.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [DefaultLocalizer(FeaturesResourceIds.DEFAULT_RESOURCE, FeaturesResourceIds.DEFAULT_ASSEMBLY)]
    [RuleInterface(StructureCustomAssembliesConstants.IASPSFeatureMacros_CopeSel, StructureCustomAssembliesConstants.IASPSFeatureMacros_CopeSel)]
    public class CopeSelectorRule : SelectorRule
    {
        //=====================================================================================================
        //DefinitionName/ProgID of this symbol is "Features,Ingr.SP3D.Content.Structure.CopeSelectorRule"
        //=====================================================================================================

        /// <summary>
        /// Processes the rules and returns the collection of named part items that are possible choices based on provided inputs.
        /// If multiple choices are provided by the selection rule, the first element in this collection will be the default selection made.
        /// </summary>
        public override ReadOnlyCollection<string> Selections
        {
            get
            {
                // Only one item available
                List<string> list = new List<string>() { StructureCustomAssembliesConstants.CopeFeature1 };
                return new ReadOnlyCollection<string>(list);
            }
        }
    }
}