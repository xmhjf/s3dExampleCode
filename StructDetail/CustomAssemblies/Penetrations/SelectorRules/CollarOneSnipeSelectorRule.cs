//***************************************************************************************************************************************/
// Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  CollarOneSnipeSelectorRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMCollarRules.dll
//  Original Class Name: ‘CollarOneSnipeSel’ in VB content
//
//Abstract
//	CollarOneSnipeSelectorRule is a .NET selector rule which selects the list of available items in the context of the CollarOneSnipe. 
//  This class subclasses from CollarSelectorRule.
//
// Change History:
//  dd.mmm.yyyy    who    change description
//****************************************************************************************************************************************/
using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Structure.Middle;
using System.Collections.ObjectModel;
using System.Collections.Generic;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    ///  Selector for CollarOneSnipe, which selects the list of available items in the context of the CollarOneSnipe.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    //New content does not need to specify the interface name, but this is specified here to support code written against old rules which directly uses interface name to get selector question. 
    [RuleInterface(DetailingCustomAssembliesConstants.IASMCollarRules_CollarOneSnipeSel, DetailingCustomAssembliesConstants.IASMCollarRules_CollarOneSnipeSel)]
    [DefaultLocalizer(PenetrationsResourceIds.PenetrationsResource, PenetrationsResourceIds.PenetrationsAssembly)]
    public class CollarOneSnipeSelectorRule : CollarSelectorRule
    {
        //=============================================================================================================
        //DefinitionName/ProgID of this symbol is "Penetrations,Ingr.SP3D.Content.Structure.CollarOneSnipeSelectorRule"
        //=============================================================================================================

        #region Selector Questions
        [SelectorQuestionCodelist(100, DetailingCustomAssembliesConstants.AddCornerSnipe, DetailingCustomAssembliesConstants.AddCornerSnipe, DetailingCustomAssembliesConstants.BooleanCol, (int)Answer.No)]
        public SelectorQuestionCodelist addCornerSnipe;
        #endregion Selector Questions

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
                    //Get the CollarPart
                    CollarPart collarPart = (CollarPart)base.Occurrence;

                    //Check supported Feature
                    base.ValidateInputs();

                    //Check if any ToDo message of type error has been created during the validation of inputs
                    if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        //ToDo list is created while validating the inputs
                        return new ReadOnlyCollection<string>(choices);
                    }

                    bool isPenetratingPlate = base.Penetrating is PlatePartBase ? true : false;

                    //Get the inputs
                    Feature slot = base.SupportedFeature;
                    string slotPartName = slot.PartName;
                    string slotPartClassName = ((PartClass)slot.Part.PartClass).PartClassName;

                    // A slot Feature can have:
                    //  One collar, regular collar or right side collar (A)
                    //  Two collars, one right side collar (A) and one left side collar(B)
                    if (slotPartClassName.Equals(DetailingCustomAssembliesConstants.SlotC))
                    {
                        switch (slotPartName)
                        {
                            case DetailingCustomAssembliesConstants.SlotAC_LT_PAT:
                                //If a Plate/Stiffener combination is used to form a EA or UA section alias then there is no fillet radius on the top flange right top corner.
                                //CollarACT_A expects a radius at the corner.
                                //Therefore, if the penetrated part is a plate then we cannot add CollarACT_A.
                                //Use CollarTCT_A instead, which does not have a radius.
                                if (isPenetratingPlate)
                                {
                                    choices.Add(DetailingCustomAssembliesConstants.CollarTCT_A);
                                }
                                else
                                {
                                    choices.Add(DetailingCustomAssembliesConstants.CollarACT_A);
                                }
                                break;
                            case DetailingCustomAssembliesConstants.SlotAC_LT_PAT2:
                                //If a Plate/Stiffener combination is used to form a EA or UA section alias then there is no fillet radius on the top flange right top corner.
                                //CollarACT_A2 expects a radius at the corner.
                                //Therefore, if the penetrated part is a plate then we cannot add CollarACT_A2.
                                //Use CollarTCT_A2 instead, which does not have a radius.
                                if (isPenetratingPlate)
                                {
                                    choices.Add(DetailingCustomAssembliesConstants.CollarTCT_A2);
                                }
                                else
                                {
                                    choices.Add(DetailingCustomAssembliesConstants.CollarACT_A2);
                                }
                                break;
                            case DetailingCustomAssembliesConstants.SlotBC_LT_PAT:
                                choices.Add(DetailingCustomAssembliesConstants.CollarBCT_A);
                                break;
                            case DetailingCustomAssembliesConstants.SlotBC_LT_PAT2:
                                choices.Add(DetailingCustomAssembliesConstants.CollarBCT_A2);
                                break;
                            case DetailingCustomAssembliesConstants.SlotFC_LT_PAT:
                                choices.Add(DetailingCustomAssembliesConstants.CollarFCT_A);
                                break;
                            case DetailingCustomAssembliesConstants.SlotFC_LT_PAT2:
                                choices.Add(DetailingCustomAssembliesConstants.CollarFCT_A2);
                                break;
                            case DetailingCustomAssembliesConstants.SlotTC_T_PAA_STR:
                            case DetailingCustomAssembliesConstants.SlotAC_LT_AAA:
                                // If user selects these items, when they change this to other item, they must validate it!
                                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(PenetrationsResourceIds.ErrInvalidSlotType,
                                    "Invalid slot type."));
                                break;
                            case DetailingCustomAssembliesConstants.SlotAC_LT_PAA:
                                //If Plate/Stiffener combination then use the Collar for BUT
                                if (isPenetratingPlate)
                                {
                                    choices.Add(DetailingCustomAssembliesConstants.CollarTCT_A3);
                                }
                                else
                                {
                                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(PenetrationsResourceIds.ErrNoCollarItemForProfile,
                                        "No equivalent Collar item for Profile."));
                                }
                                break;
                            default:
                                string penetratingSectionTypeName = base.PenetratingSectionTypeName;
                                if (string.IsNullOrWhiteSpace(penetratingSectionTypeName))
                                {
                                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(PenetrationsResourceIds.ErrGetPenetratingSectionTypeName,
                                        "Unable to get penetrating section type name."));
                                    return new ReadOnlyCollection<string>(choices);
                                }

                                switch (penetratingSectionTypeName)
                                {
                                    case MarineSymbolConstants.EA:
                                    case MarineSymbolConstants.UA:
                                        // If a Plate/Stiffener combination is used to form a EA or UA section alias then there is no fillet radius on the top flange right top corner.
                                        if (isPenetratingPlate)
                                        {
                                            base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(PenetrationsResourceIds.ErrNoCollarItemForPlate,
                                                "A new symbol have to create that does not include a radius."));
                                        }
                                        else
                                        {
                                            choices.Add(DetailingCustomAssembliesConstants.CollarACT_SM);
                                        }
                                        break;
                                    case MarineSymbolConstants.B:
                                        if (slotPartName.Equals(DetailingCustomAssembliesConstants.SlotBC_L_LTT_STR))
                                        {
                                            choices.Add(DetailingCustomAssembliesConstants.CollarBCT_A3);
                                        }
                                        choices.Add(DetailingCustomAssembliesConstants.CollarBCT_SM);
                                        break;
                                    case MarineSymbolConstants.FB:
                                        choices.Add(DetailingCustomAssembliesConstants.CollarFCT_SM);
                                        break;
                                    case MarineSymbolConstants.BUT:
                                    case MarineSymbolConstants.BUTL2:
                                        choices.Add(DetailingCustomAssembliesConstants.CollarTCT_SM);
                                        break;
                                    default:
                                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(PenetrationsResourceIds.ErrInvalidSectionTypeName,
                                            "Invalid penetrating section type name."));
                                        break;
                                }
                                break;
                        }
                    }
                    else if (slotPartClassName.Equals(DetailingCustomAssembliesConstants.SlotC2))
                    {
                        //get the CollarCreationOrder from IASMCollarRules_RootClipSel
                        string collarCreationOrder = ((PropertyValueString)collarPart.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.CollarCreationOrder)).PropValue;
                        switch (slotPartName)
                        {
                            case DetailingCustomAssembliesConstants.SlotL2C2_T_PTT_STR:
                            case DetailingCustomAssembliesConstants.SlotTC2_T_PTT_STR:
                                if (collarCreationOrder == DetailingCustomAssembliesConstants.Primary)
                                {
                                    choices.Add(DetailingCustomAssembliesConstants.CollarTCT_A);
                                }
                                else
                                {
                                    choices.Add(DetailingCustomAssembliesConstants.CollarTCT_B);
                                }
                                break;
                            case DetailingCustomAssembliesConstants.SlotL2C2_T_PTT_STR2:
                            case DetailingCustomAssembliesConstants.SlotTC2_T_PTT_STR2:
                                if (collarCreationOrder == DetailingCustomAssembliesConstants.Primary)
                                {
                                    choices.Add(DetailingCustomAssembliesConstants.CollarTCT_A2);
                                }
                                else
                                {
                                    choices.Add(DetailingCustomAssembliesConstants.CollarTCT_B2);
                                }
                                break;
                            case DetailingCustomAssembliesConstants.SlotTC2_T_PAA_STR:
                                if (collarCreationOrder == DetailingCustomAssembliesConstants.Primary)
                                {
                                    choices.Add(DetailingCustomAssembliesConstants.CollarTCT_A3);
                                }
                                else
                                {
                                    choices.Add(DetailingCustomAssembliesConstants.CollarTCT_B3);
                                }
                                break;
                            default:
                                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(PenetrationsResourceIds.ErrInvalidSlotType,
                                    "Invalid slot type."));
                                break;
                        }
                    }
                    else
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(PenetrationsResourceIds.ErrInvalidSlotType,
                            "Invalid slot type."));
                    }
                }
                catch (Exception)
                {
                    //There could be specific ToDo records created before, so do not override it.
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
        /// will be invoked for each “system controlled” question prior to invoking the 'Selections' method.
        /// </summary>
        /// <param name="selectorQuestion">Answer whose value can be defined</param>
        public override void OverrideDefaultAnswer(SelectorQuestion selectorQuestion)
        {
            ((SelectorQuestionCodelist)selectorQuestion).Value = (int)Answer.Yes;
        }

        #endregion Public override properties and methods
    }
}