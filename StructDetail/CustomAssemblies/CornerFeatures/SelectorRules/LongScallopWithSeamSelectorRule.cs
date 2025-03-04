//---------------------------------------------------------------------------------------------------------------------------------
//Copyright© 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  LongScallopWithSeamSelectorRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMCornerFeatRules.dll
//  Original Class Name: ‘LongScallopWSeamSel’ in VB content
//
//Abstract
//	LongScallopWithSeamSelectorRule is a .NET selector rule which selects the list of available items in the context of the Corner.
//  This class subclasses from CornerSelectorRule.
//
//Change History:
//  dd.mmm.yyyy    who    change description
//---------------------------------------------------------------------------------------------------------------------------------
using System;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Selector for corner feature, which selects the list of available items in the context of the corner feature.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [DefaultLocalizer(CornerFeaturesResourceIdentifiers.CornerFeaturesResources, CornerFeaturesResourceIdentifiers.CornerFeaturesAssembly)]
    //New content does not need to specify the interface name, but this is specified here to support code written against old rules which directly uses interface name to get selector question. 
    [RuleInterface(DetailingCustomAssembliesConstants.IASMCornerFeatRules_LongScallopWSeamSel, DetailingCustomAssembliesConstants.IASMCornerFeatRules_LongScallopWSeamSel)]
    public class LongScallopWithSeamSelectorRule : CornerSelectorRule
    {
        //====================================================================================================================
        //DefinitionName/ProgID of this symbol is "CornerFeatures,Ingr.SP3D.Content.Structure.LongScallopWithSeamSelectorRule"
        //====================================================================================================================

        #region Public override properties and methods

        /// <summary>
        /// Gets the list of named part items that are possible choices based on provided inputs.
        /// </summary>
        public override ReadOnlyCollection<string> Selections
        {
            get
            {
                Collection<string> choices = new Collection<string>();
                try
                {
                    //Validates the inputs of the corner feature
                    base.ValidateInputs();

                    //Check if any ToDo message of type error has been created during the validation of inputs
                    if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        //ToDo list is created while validating the inputs
                        return new ReadOnlyCollection<string>(choices);
                    }

                    choices.Add(DetailingCustomAssembliesConstants.LongScallopWithSeam);
                    choices.Add(DetailingCustomAssembliesConstants.LongScallopWithSeamWithCollar);
                }
                catch (Exception)
                {
                    //There could be specific ToDo records created before, so do not override it.
                    if (base.ToDoListMessage == null)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(CornerFeaturesResourceIdentifiers.ToDoSelections,
                            "Unexpected error while getting the selections from {0}"), this.ToString()));
                    }
                }

                return new ReadOnlyCollection<string>(choices);
            }
        }

        #endregion Public override properties and methods
    }
}