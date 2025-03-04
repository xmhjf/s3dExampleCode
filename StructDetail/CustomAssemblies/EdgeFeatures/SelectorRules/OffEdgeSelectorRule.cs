//-------------------------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  OffEdgeSelectorRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMEdgeFeatureRules.dll
//  Original Class Name: ‘OffEdgeSel’ in VB content
//
//Abstract
//	OffEdgeSelectorRule is a .NET selector rule which selects the list of available items in the context of the Edge feature. 
//  This class subclasses from EdgeFeatureSelectorRule.
//
//Change History:
//  dd.mmm.yyyy    who    change description
//-------------------------------------------------------------------------------------------------------------------------------
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Selector for edge feature, which selects the list of available items in the context of this edge feature.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [DefaultLocalizer(EdgeFeaturesResourceIdentifiers.EdgeFeaturesResources, EdgeFeaturesResourceIdentifiers.EdgeFeaturesAssembly)]
    //New content does not need to specify the interface name, but this is specified here to support code written against old rules which directly uses interface name to get parameter value.
    [RuleInterface(DetailingCustomAssembliesConstants.IASMEdgeFeatureRules_OffEdgeSel, DetailingCustomAssembliesConstants.IASMEdgeFeatureRules_OffEdgeSel)]
    public class OffEdgeSelectorRule : EdgeFeatureSelectorRule
    {
        //=============================================================================================================
        //DefinitionName/ProgID of this selector rule is "EdgeFeatures,Ingr.SP3D.Content.Structure.OffEdgeSelectorRule"
        //=============================================================================================================

        /// <summary>
        /// Return the list of available items that are possible choices for this edge feature.
        /// </summary>
        public override ReadOnlyCollection<string> Selections
        {
            get 
            {
                Collection<string> choices = new Collection<string>();
                try
                {
                    //Validating inputs
                    ValidateEdgeFeatureInputs();

                    //Check if any ToDo message of type error has been created during the validation of inputs
                    if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        //ToDo list is created while validating the inputs
                        return new ReadOnlyCollection<string>(choices);
                    }
  
                    choices.Add(DetailingCustomAssembliesConstants.Circular_DxO);
                    choices.Add(DetailingCustomAssembliesConstants.Circular_25x1p5);
                    choices.Add(DetailingCustomAssembliesConstants.Circular_32x1p5);
                    choices.Add(DetailingCustomAssembliesConstants.Circular_35x1p5);
                    choices.Add(DetailingCustomAssembliesConstants.Circular_50x1p5);
                    choices.Add(DetailingCustomAssembliesConstants.Circular_75x1p5);
                    choices.Add(DetailingCustomAssembliesConstants.Circular_100x1p5);
                    choices.Add(DetailingCustomAssembliesConstants.Circular_150x1p5);

                    choices.Add(DetailingCustomAssembliesConstants.Oval_LxDxO);
                    choices.Add(DetailingCustomAssembliesConstants.Oval_60x30x1p5);
                    choices.Add(DetailingCustomAssembliesConstants.Oval_100x50x1p5);
                    choices.Add(DetailingCustomAssembliesConstants.Oval_130x60x1p5);
                    choices.Add(DetailingCustomAssembliesConstants.Oval_150x75x1p5);
                }
                catch
                {
                    if (base.ToDoListMessage == null)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(EdgeFeaturesResourceIdentifiers.ToDoEdgeFeatureSelections,
                            "Unexpected error while getting the selections from OffEdge selector rule."));
                    }
                }
                return new ReadOnlyCollection<string>(choices); 
            }
        }        
    }
}