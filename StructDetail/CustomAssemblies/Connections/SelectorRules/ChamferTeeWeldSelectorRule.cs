/***************************************************************************************************************************************/
// Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  ChamferTeeWeldSelectorRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMPhysConnRules.dll
//  Original Class Name: ‘ChamferTeeWeldSel’ in VB content
//
//Abstract
//	ChamferTeeWeldSelectorRule is a .NET selector rule which selects the list of available items in the context of the PhysicalConnections. 
//  This class subclasses from PhysicalConnectionSelectorRule.
/*****************************************************************************************************************************************/

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Selector for PhysicalConnections, which selects the list of available items in the context of the ChamferTeeWelds.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    //New content does not need to specify the interface name, but this is specified here to support code written against old rules which directly uses interface name to get selector question. 
    [RuleInterface(DetailingCustomAssembliesConstants.IASMPhysConnRules_ChamferTeeWeldSel, DetailingCustomAssembliesConstants.IASMPhysConnRules_ChamferTeeWeldSel)]
    [DefaultLocalizer(ConnectionsResourceIdentifiers.ConnectionsResource, ConnectionsResourceIdentifiers.ConnectionsAssembly)]
    public class ChamferTeeWeldSelectorRule : PhysicalConnectionSelectorRule
    {
        //=================================================================================================================
        //DefinitionName/ProgID of this symbol is "Connections,Ingr.SP3D.Content.Structure.ChamferTeeWeldSelectorRule"
        //=================================================================================================================

        #region Selector questions
        [SelectorQuestionCodelist(100, DetailingCustomAssembliesConstants.Category, DetailingCustomAssembliesConstants.Category, DetailingCustomAssembliesConstants.TeeWeldCategory, (int)TeeWeldCategory.Normal)]
        public SelectorQuestionCodelist categoryTypeAnswer;
        [SelectorQuestionCodelist(101, DetailingCustomAssembliesConstants.ClassSociety, DetailingCustomAssembliesConstants.ClassSociety, DetailingCustomAssembliesConstants.ClassSocietyCol, (int)ClassSociety.Lloyds)]
        public SelectorQuestionCodelist classSocietyAnswer;
        [SelectorQuestionCodelist(102, DetailingCustomAssembliesConstants.BevelAngleMethod, DetailingCustomAssembliesConstants.BevelAngleMethod, DetailingCustomAssembliesConstants.BevelMethod, (int)BevelMethod.Constant)]
        public SelectorQuestionCodelist bevelAngleMethodAnswer;
        [SelectorQuestionString(103, DetailingCustomAssembliesConstants.ChamferThickness, DetailingCustomAssembliesConstants.ChamferThickness, DetailingCustomAssembliesConstants.ChamferThicknessAnswer)]
        public SelectorQuestionString chamferTicknessAnswer;
        #endregion Selector questions
        /// <summary>
        ///  Selector for Chamfer Welds, which selects the list of available items in the context of the Physical Connections.
        /// </summary>
        public override ReadOnlyCollection<string> Selections
        {
            get
            {
                Collection<string> choices = new Collection<string>();
                try
                {
                    // Get Class Arguments
                    PhysicalConnection physicalConnection = (PhysicalConnection)base.Occurrence;

                    //Validating the inputs required to create the slot.
                    //This is a virtual method and can be overriden to add specific validation, if specific input is missing.
                    ValidateInputs();
                    //Check if any ToDo message of type error has been created during the validation of inputs
                    if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        //ToDo list is created while validating the inputs
                        return new ReadOnlyCollection<string>(choices);
                    }

                    // Use filter in place of selection logic, if set
                    choices = ConnectionServices.GetFilteredSelections(physicalConnection);
                    if (choices.Count == 0)
                    {
                        string category = ((PropertyValue)(physicalConnection).GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.Category)).ToString();
                        double mountingAngle = physicalConnection.GetConnectionAngle();
                        //If the returned Mounting Angle is greater than 90 degrees, subtract from 180 to get the smaller angle
                        if ((mountingAngle - StructHelper.DISTTOL) >= Math.PI / 2)
                        {
                            mountingAngle = Math.PI - mountingAngle;
                        }
                        double mountingAngleComplement = (Math.PI / 2) - mountingAngle;
                        //If this is a physical connection between a stiffener and the plate it is stiffening, and if the mounting
                        //angle is well off normal (beyond OFF_NORMAL_ANGLE_TOLERANCE), then stiffener attachment method
                        //must be examined in selecting the proper tee weld.
                        bool isStiffenerAttachmentNeedtoVerify = false;
                        double webThickness = 0;
                        StiffenerPartBase stiffenerPartBase = base.BoundedObject as StiffenerPartBase;
                        PlatePartBase platePartBase = base.BoundingObject as PlatePartBase;
                        if (stiffenerPartBase != null)
                        {
                            webThickness = DetailingCustomAssembliesServices.GetWebThickness(stiffenerPartBase.CrossSection);
                        }
                        if (stiffenerPartBase != null && platePartBase != null)//Possibly need to consider long point /short point issue (stiffener attachment)
                        {
                            //check if connected object 2 the stiffened plate of connected object 1
                            PlateSystemBase stiffenedPlate = stiffenerPartBase.RootStiffenerSystem.PlateToStiffen;
                            PlateSystemBase parentSystem = ((PlatePart)platePartBase).RootPlateSystem;
                            if (stiffenedPlate == parentSystem && mountingAngleComplement > Math.PI / 90)
                            {
                                isStiffenerAttachmentNeedtoVerify = true;
                            }
                        }
                        if (isStiffenerAttachmentNeedtoVerify)
                        {
                            choices = ConnectionServices.GetSelectionsForChamferWeld(base.BoundedObject, category, mountingAngleComplement, webThickness);
                        }
                        else
                        {
                            //Don't need to examine stiffener attachment.
                            //Either this
                            //   physical connection is not between a stiffener and its stiffened plate
                            //OR
                            //  the mounting angle between stiffener and it's stiffened plate is close to normal
                            double chamferThickness = Convert.ToDouble(((PropertyValue)(physicalConnection).GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.ChamferThickness)).ToString());
                            switch (category)
                            {
                                case DetailingCustomAssembliesConstants.Normal:
                                    double openAngle = chamferThickness * Math.Abs(Math.Tan(mountingAngleComplement));
                                    if (openAngle < Math3d.FitTolerance * 3)
                                    {
                                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer1);
                                    }
                                    else if (openAngle >= Math3d.FitTolerance * 3 && ((mountingAngle - StructHelper.DISTTOL) > Math.PI / 4))
                                    {
                                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer2);
                                    }
                                    else if (openAngle >= Math3d.FitTolerance * 3 && ((mountingAngle - StructHelper.DISTTOL) <= Math.PI / 4))
                                    {
                                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer3);
                                    }
                                    break;
                                case DetailingCustomAssembliesConstants.Full:
                                    if (chamferThickness < Math3d.FitTolerance * 4)
                                    {
                                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer4);
                                    }
                                    else
                                    {
                                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer5);
                                    }
                                    break;
                                case DetailingCustomAssembliesConstants.Deep:
                                    if (chamferThickness > StructHelper.DISTTOL * 25)
                                    {
                                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer6);
                                    }
                                    else
                                    {
                                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChamfer7);
                                    }
                                    break;
                                case DetailingCustomAssembliesConstants.Chain:
                                    {
                                        choices.Add(DetailingCustomAssembliesConstants.ChainWeldChamfer);
                                    }
                                    break;
                                case DetailingCustomAssembliesConstants.ZigZag:
                                    {
                                        choices.Add(DetailingCustomAssembliesConstants.ZigZagWeldChamfer);
                                    }
                                    break;
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    //may be there could be specific ToDo record is created before, so do not override it.
                    if (base.ToDoListMessage == null)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(ConnectionsResourceIdentifiers.ToDoPhysicalConnectionSelections,
                        "Unexpected error while evaluating the {0}"), this.ToString()));
                    }
                }
                return new ReadOnlyCollection<string>(choices);
            }
        }

        // <summary>
        /// Sets the answer for the selector question when it is controlled by the system. 
        /// The default answer will be the value provided via the SelectorQuestion attribute and
        /// will be invoked for each “system controlled” question prior to invoking the 'Selections' property.
        /// </summary>
        /// <param name="answer">Answer whose value can be defined</param>
        public override void OverrideDefaultAnswer(SelectorQuestion selectorQuestion)
        {
            string selectorQuestionName = selectorQuestion.Name;
            if (selectorQuestionName.Equals(DetailingCustomAssembliesConstants.BevelAngleMethod))
            {
                SelectorQuestionCodelist bevelAngleMethodAnswerCodelist = (SelectorQuestionCodelist)selectorQuestion;
                bevelAngleMethodAnswerCodelist.Value = (int)BevelMethod.Varying;
            }
        }
    }
}
