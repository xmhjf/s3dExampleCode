//------------------------------------------------------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  SingleSidedChamferSelectorRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMChamferRules.dll
//  Original Class Name: ‘SingleSidedSel’ in VB content
//
//Abstract
//	 SingleSidedChamferSelectorRule is a .NET selection rule for Chamfer which selects the list of available items in the context of the SingleSidedChamfer.
//   This class subclasses from ChamferSelectorRule.       
//-------------------------------------------------------------------------------------------------------------------------------------------------------------
using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Selector for SingleSidedChamfer, which selects the list of available items in the context of the Chamfer.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [DefaultLocalizer(ChamfersResourceIds.ChamfersResources, ChamfersResourceIds.ChamfersAssembly)]
    //New content does not need to specify the interface name, but this is specified here to support code written against old rules which directly uses interface name to get parameter value.
    [RuleInterface(DetailingCustomAssembliesConstants.IASMChamferRules_SingleSidedSel, DetailingCustomAssembliesConstants.IASMChamferRules_SingleSidedSel)]
    public class SingleSidedChamferSelectorRule : ChamferSelectorRule
    {
        //==============================================================================================================
        //DefinitionName/ProgID of this symbol is "Chamfers,Ingr.SP3D.Content.Structure.SingleSidedChamferSelectorRule"
        //==============================================================================================================
        #region Public override properties and methods

        /// <summary>
        /// Return the list of available items that are possible choices for this chamfer.
        /// </summary>
        public override ReadOnlyCollection<string> Selections
        {
            get
            {
                Collection<string> choices = new Collection<string>();
                try
                {
                    //Get the Chamfer
                    Feature chamfer = (Feature)base.Occurrence;

                    //Validating inputs
                    base.ValidateChamferInputs();

                    //Check if any ToDo message of type error has been created during the validation of inputs
                    if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        //ToDo list is created while validating the inputs
                        return new ReadOnlyCollection<string>(choices);
                    }

                    //Gets the property value for the first question that matches the given question name
                    //Chamfer type and Chamfer weld is obtained from the RootChamferSelectionRule
                    String chamferType = ((PropertyValueString)chamfer.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.ChamferType)).PropValue;
                    String chamferWeld = ((PropertyValueString)chamfer.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.ChamferWeld)).PropValue;

                    //Get the chamfer custom assembly parent it can be assembly connection or a feature
                    BusinessObject chamferCustomAssemblyParent = SymbolHelper.GetCustomAssemblyParent(chamfer);
                    Feature feature = chamferCustomAssemblyParent as Feature;

                    //If chamfer is created by assembly connection
                    if (chamferCustomAssemblyParent != null && chamferCustomAssemblyParent.SupportsInterface(MarineSymbolConstants.IJStructAssemblyConnection))
                    {
                        if (chamferType.Equals(DetailingCustomAssembliesConstants.ChamferTypeObj1Offset))
                        {
                            choices.Add(DetailingCustomAssembliesConstants.SingleSidedOffset);
                        }
                        else
                        {
                            choices.Add(DetailingCustomAssembliesConstants.SingleSidedBase);
                        }
                    }
                    //if the chamferparent is a feature(webcut)                   
                    else if (feature != null && feature.FeatureType == FeatureType.WebCut)
                    {
                        if (chamferType.Equals(DetailingCustomAssembliesConstants.ChamferTypeObj1Offset) || chamferType.Equals(DetailingCustomAssembliesConstants.ChamferTypeObj2Offset))
                        {
                            choices.Add(DetailingCustomAssembliesConstants.SingleSidedOffset);
                        }
                        else if (chamferType.Equals(DetailingCustomAssembliesConstants.ChamferTypeObj1Base) || chamferType.Equals(DetailingCustomAssembliesConstants.ChamferTypeObj2Base))
                        {
                            choices.Add(DetailingCustomAssembliesConstants.SingleSidedBase);
                        }
                    }
                    //select chamfer
                    else
                    {
                        switch (chamferType)
                        {
                            case DetailingCustomAssembliesConstants.ChamferTypeObj1Offset:
                            case DetailingCustomAssembliesConstants.ChamferTypeObj2Offset:
                                choices.Add(DetailingCustomAssembliesConstants.SingleSidedOffsetPC);
                                break;
                            case DetailingCustomAssembliesConstants.ChamferTypeObj1Base:
                            case DetailingCustomAssembliesConstants.ChamferTypeObj2Base:
                                choices.Add(DetailingCustomAssembliesConstants.SingleSidedBasePC);
                                break;
                            case DetailingCustomAssembliesConstants.ChamferTypeObj1BaseObj2Offset:
                                choices.Add((chamferWeld == DetailingCustomAssembliesConstants.ChamferWeldFirst) ?
                                    DetailingCustomAssembliesConstants.SingleSidedBasePC :
                                    DetailingCustomAssembliesConstants.SingleSidedOffset);
                                break;
                            case DetailingCustomAssembliesConstants.ChamferTypeObj1OffsetObj2Base:
                                choices.Add((chamferWeld == DetailingCustomAssembliesConstants.ChamferWeldFirst) ?
                                    DetailingCustomAssembliesConstants.SingleSidedOffsetPC :
                                    DetailingCustomAssembliesConstants.SingleSidedBase);
                                break;
                        }                        
                    }
                }
                catch
                {
                    if (base.ToDoListMessage == null)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(ChamfersResourceIds.ToDoChamferSelections,
                            "Unexpected error while getting the selections from SingleSided Chamfer selector rule."));
                    }
                }
                return new ReadOnlyCollection<string>(choices);
            }
        }

        #endregion Public override properties and methods
    }
}