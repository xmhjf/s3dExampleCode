/***********************************************************************************************************************/
// Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
// File
//  SlotCSelectorRule.cs
//
// Abstract
//	SlotCSelectorRule is a .NET selector rule which selects the list of available items in the context of the SlotC. 
//  This class subclasses from SlotSelectorRule.
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMSlotRules.dll
//  Original Class Name: ‘SlotCSel’ in VB content
//
// Change History:
//  dd.mmm.yyyy    who    change description
/***********************************************************************************************************************/
using System;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Content.Structure.Services;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Selector for slot, which selects the list of available items in the context of the slot.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    //New content does not need to specify the interface name, but this is specified here to support code written against old rules which directly uses interface name to get selector question. 
    [RuleInterface(DetailingCustomAssembliesConstants.IASMSlotRules_SlotCSel, DetailingCustomAssembliesConstants.IASMSlotRules_SlotCSel, DetailingCustomAssembliesConstants.IASlotCSelV)]
    [DefaultLocalizer(PenetrationsResourceIds.PenetrationsResource, PenetrationsResourceIds.PenetrationsAssembly)]
    public class SlotCSelectorRule : SlotSelectorRule
    {
        //=====================================================================================================
        //DefinitionName/ProgID of this symbol is "Penetrations,Ingr.SP3D.Content.Structure.SlotCSelectorRule"
        //=====================================================================================================

        [SelectorQuestionDouble(100, DetailingCustomAssembliesConstants.Clearance, DetailingCustomAssembliesConstants.Clearance, 0.1)]
        public SelectorQuestionDouble clearanceAnswer;
        [SelectorQuestionCodelist(101, DetailingCustomAssembliesConstants.BaseCorners, DetailingCustomAssembliesConstants.BaseCorners, DetailingCustomAssembliesConstants.BooleanCol, (int)Answer.No)]
        public SelectorQuestionCodelist baseCornersAnswer;
        [SelectorQuestionCodelist(102, DetailingCustomAssembliesConstants.OutsideCorners, DetailingCustomAssembliesConstants.OutsideCorners, DetailingCustomAssembliesConstants.BooleanCol, (int)Answer.No)]
        public SelectorQuestionCodelist outsideCornersAnswer;


        #region Public methods

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
                    //Get the slot feature
                    Feature slot = (Feature)base.Occurrence;

                    //Validating the inputs required to create the slot.
                    //This is a virtual method and can be overriden to add specific validation, if specific input is missing.
                    base.ValidateInputs();

                    //Check if any ToDo message of type error has been created during the validation of inputs
                    if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        //ToDo list is created while validating the inputs
                        return new ReadOnlyCollection<string>(choices); 
                    }

                    //Get penetrating profile part height and section type name
                    string sectionTypeName = GetPenetratingSectionTypeName();
                    double profilePartHeight = GetPenetratingHeight(MarineSymbolConstants.IJUAXSectionWeb, MarineSymbolConstants.WebLength);

                    //Get plate tightness either from penetrated plate part or base plate part
                    PlatePartBase platepart;
                    platepart = base.Penetrated as PlatePartBase;

                    if (platepart == null)
                    {
                        platepart = (PlatePartBase)slot.SlotBasePlatePort.Connectable;
                    }
                    Tightness plateTightness = platepart.Tightness;

                    //Get slot type from RootSlotSelectorRule
                    int slotType = ((PropertyValueCodelist)slot.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.SlotType)).PropValue;

                    //Consider intially it is a tight slot
                    bool tightSlot = true;
                    if (slotType == (int)SlotType.Default)
                    {
                        if (plateTightness == Tightness.NonTight || plateTightness == Tightness.UnSpecified)
                        {
                            tightSlot = false;
                        }
                    }
                    else if (slotType == (int)SlotType.NonTight)
                    {
                        tightSlot = false;
                    }

                    //Get slection choices based on the section type. This can be customizable by the content implementor.
                    switch (sectionTypeName)
                    {
                        case MarineSymbolConstants.EA:
                        case MarineSymbolConstants.UA:                            
                            if (tightSlot == false) //NonTightSlot
                            {
                                //Left connected
                                choices.Add(DetailingCustomAssembliesConstants.SlotAC_L_PAT_DTR);
                                //Left & Top Connected
                                choices.Add(DetailingCustomAssembliesConstants.SlotAC_LT_PAA);
                                choices.Add(DetailingCustomAssembliesConstants.SlotAC_LT_PAT);
                            }
                            else //TightSlot
                            {
                                //Left Top connected
                                choices.Add(DetailingCustomAssembliesConstants.SlotAC_LT_PAT);
                                choices.Add(DetailingCustomAssembliesConstants.SlotAC_LT_PAT2);
                                //Left connected
                                choices.Add(DetailingCustomAssembliesConstants.SlotAC_L_PAT_STR);
                            }                            
                            break;

                        case MarineSymbolConstants.BUTL2:
                            if (profilePartHeight > 0.1)
                            {
                                if (tightSlot == false) //NonTightSlot
                                {
                                    //Left connected
                                    choices.Add(DetailingCustomAssembliesConstants.SlotL2C_L_PAT_DTR); //Default selection
                                }
                                else //TightSlot
                                {
                                    //Left connected
                                    choices.Add(DetailingCustomAssembliesConstants.SlotL2C_L_PAT_DTR2);
                                }
                            }
                            break;
                        case MarineSymbolConstants.B:
                            if (tightSlot == false) //NonTightSlot
                            {
                                if (profilePartHeight > 0.12)
                                {
                                    choices.Add(DetailingCustomAssembliesConstants.SlotBC_L_PAT_DTR);
                                }
                                else
                                {
                                    choices.Add(DetailingCustomAssembliesConstants.SlotBC_L_PAT_STR);
                                }
                            }
                            else //TightSlot
                            {
                                //Left Connected
                                choices.Add(DetailingCustomAssembliesConstants.SlotBC_L_LTT_STR);
                                choices.Add(DetailingCustomAssembliesConstants.SlotBC_L_PAT_STR);

                                //Left Top Connected
                                choices.Add(DetailingCustomAssembliesConstants.SlotBC_LT_PAT);
                                choices.Add(DetailingCustomAssembliesConstants.SlotBC_LT_PAT2);
                            }
                            break;

                        case MarineSymbolConstants.BUT:
                            if (profilePartHeight > 0.1)
                            {
                                if (tightSlot == false) //NonTightSlot
                                {
                                    //Left Connected
                                    choices.Add(DetailingCustomAssembliesConstants.SlotTC_L_PAT_DTR);
                                    //Top Connected
                                    choices.Add(DetailingCustomAssembliesConstants.SlotTC_T_PAA_STR);
                                }
                                else //TightSlot
                                {
                                    //Left Connected
                                    choices.Add(DetailingCustomAssembliesConstants.SlotTC_L_PAT_STR);
                                }
                            }
                            break;
                        case MarineSymbolConstants.FB:
                            if (tightSlot == false) //NonTightSlot
                            {
                                //Left Connected
                                choices.Add(DetailingCustomAssembliesConstants.SlotFC_L_PAT_DTR);
                            }
                            else  //TightSlot
                            {
                                //Left Top connected
                                choices.Add(DetailingCustomAssembliesConstants.SlotFC_LT_PAT);
                                choices.Add(DetailingCustomAssembliesConstants.SlotFC_LT_PAT2);

                                //Left Connected
                                choices.Add(DetailingCustomAssembliesConstants.SlotFC_L_PAT_STR);
                            }
                            break;
                        default:
                            break;
                    }
                }
                catch (Exception)
                {
                    //may be there could be specific ToDo record is created before, so do not override it.
                    if (base.ToDoListMessage == null)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(PenetrationsResourceIds.ToDoSelections,
                            "Unexpected error while getting the selections from {0}"), this.ToString()));
                    }
                }
                return new ReadOnlyCollection<string>(choices);
            }
        }

        /// <summary>
        /// Sets the value of the selector question when it is controlled by the system. 
        /// The default answer will be the value provided via the SelectorQuestion attribute and
        /// will be invoked for each “system controlled” question prior to invoking the Selections rule method.
        /// </summary>
        /// <param name="selectorQuestion">Answer whose value can be defined</param>
        public override void OverrideDefaultAnswer(SelectorQuestion selectorQuestion)
        {
            try
            {
                //Get the slot feature
                Feature slot = (Feature)base.Occurrence;               
                
                //Validating inputs
                ValidateInputs();

                if (base.Penetrated == null || base.Penetrating == null)
                {
                    return;
                }

                if (selectorQuestion.Name.Equals(DetailingCustomAssembliesConstants.Clearance, StringComparison.OrdinalIgnoreCase))
                {
                    //gets the penetrating profile section type name
                    string sectionTypeName = GetPenetratingSectionTypeName();

                    //Set clearence value based on the section type name 
                    ((SelectorQuestionDouble)selectorQuestion).Value = PenetrationsServices.GetClearance(sectionTypeName);
                }
                else if (selectorQuestion.Name.Equals(DetailingCustomAssembliesConstants.BaseCorners, StringComparison.OrdinalIgnoreCase))
                {
                    //Set BaseCorners value to No 
                    ((SelectorQuestionCodelist)selectorQuestion).Value = (int)Answer.No;
                }
                else if (selectorQuestion.Name.Equals(DetailingCustomAssembliesConstants.OutsideCorners, StringComparison.OrdinalIgnoreCase))
                {
                    //Set OutsideCorners value to No 
                    ((SelectorQuestionCodelist)selectorQuestion).Value = (int)Answer.No;
                }

            }
            catch (Exception)
            {
                //may be there could be specific ToDo record is created before, so do not override it.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(PenetrationsResourceIds.ToDoDefaultAnswer,
                        "Error while setting the default answer of") + " " + this.ToString() + ".");
                }
            }
        }

        #endregion Public methods
    }

}
