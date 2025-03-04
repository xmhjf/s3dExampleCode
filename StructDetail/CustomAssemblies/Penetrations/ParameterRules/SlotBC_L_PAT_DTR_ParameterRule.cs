/******************************************************************************************************************/
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  SlotBC_L_PAT_DTR_Parm.cs
//
//Abstract
//	SlotBC_L_PAT_DTR_parameterRule is a .NET parameter rule which is defining the custom parameter rule. 
//  This class subclasses from SlotParameterRule.
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMSlotRules.dll
//  Original Class Name: ‘SlotBC_L_PAT_DTR_Parm’ in VB content
//
// Parameter rule which defining the custom parameter rule.
// BC - Bulb cross section and Slot C type.
// L - Left connected.
// PAT - Connected at left bottom, left top, right top corner (Point-Arc-Tangent)
// DTR - Different tangent radius.
/******************************************************************************************************************/
using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Content.Structure.Services;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Parameter rule which defines the custom parameter rule for Slot C type and left connected penetrations with BULB cross section.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    ///New content does not need to specify the interface name, but this is specified here to support code written against old rules which directly uses interface name to get parameter value. 
    [RuleInterface(DetailingCustomAssembliesConstants.IASMSlotRules_SlotBC_L_PAT_DTR_Parm, DetailingCustomAssembliesConstants.IASMSlotRules_SlotBC_L_PAT_DTR_Parm, DetailingCustomAssembliesConstants.IASlotBC_L_PAT_DTR_ParmV)]
    [DefaultLocalizer(PenetrationsResourceIds.PenetrationsResource, PenetrationsResourceIds.PenetrationsAssembly)]
    public class SlotBC_L_PAT_DTR_ParameterRule : SlotParameterRule
    {
        //==============================================================================================================
        //DefinitionName/ProgID of this symbol is "Penetrations,Ingr.SP3D.Content.Structure.SlotBC_L_PAT_DTR_ParameterRule"
        //==============================================================================================================

        //These are the parameters which will be evaluated by this rule.
        [ControlledParameter(1, DetailingCustomAssembliesConstants.TopFlangeLeftTopCornerRadius, DetailingCustomAssembliesConstants.TopFlangeLeftTopCornerRadius)]
        public ControlledParameterDouble topFlangeLeftTopCornerRadius;
        [ControlledParameter(2, DetailingCustomAssembliesConstants.TopClearance, DetailingCustomAssembliesConstants.TopClearance)]
        public ControlledParameterDouble topClearance;
        [ControlledParameter(3, DetailingCustomAssembliesConstants.TopFlangeRightTopCornerRadius, DetailingCustomAssembliesConstants.TopFlangeRightTopCornerRadius)]
        public ControlledParameterDouble topFlangeRightTopCornerRadius;
        [ControlledParameter(4, DetailingCustomAssembliesConstants.SlotAngle, DetailingCustomAssembliesConstants.SlotAngle)]
        public ControlledParameterDouble slotAngle;

        /// <summary>
        /// Evaluates the parameter rule.
        /// It computes the item parameters in the context of the smart occurrence.
        /// </summary>
        public override void Evaluate()
        {
            try
            {
                //Get the slot feature
                Feature slot = (Feature)base.Occurrence;

                //Validating the inputs required to create the slot.
                base.ValidateInputs();

                //Check if any ToDo message of type error has been created during the validation of inputs
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    //ToDo list is created while validating the inputs
                    return;
                }

                //Gets the penetrating height on passing penetrated and penetrating objects(profile or plate) 
                double profilePartHeight = GetPenetratingHeight(MarineSymbolConstants.IJUAXSectionWeb, MarineSymbolConstants.WebLength);
                
                //Gets the slot angle onpassing obtained slot 
                slotAngle.Value = PenetrationsServices.GetSlotAngle(slot);

                // Set topFlangeLeftTopCornerRadius,topClearance,topFlangeRightTopCornerRadius values based on the penetrating height 
                // as one geometry is not appropriate for all the cross sections and its values must be changing based on the input geometry attribute changes.
                if (profilePartHeight >= 0.18 && profilePartHeight < 0.2)
                {
                    topFlangeLeftTopCornerRadius.Value = 0.025;
                    topClearance.Value = 0.025;
                    topFlangeRightTopCornerRadius.Value = 0.04;
                }
                else if (profilePartHeight >= 0.2)
                {
                    topFlangeLeftTopCornerRadius.Value = 0.04;
                    topClearance.Value = 0.04;
                    topFlangeRightTopCornerRadius.Value = 0.05;
                }
                else
                {
                    topFlangeLeftTopCornerRadius.Value = 0.025;
                    topClearance.Value = 0.025;
                    topFlangeRightTopCornerRadius.Value = 0.04;
                }

                //Sets the orientation of the slot depending upon the orientation of the assembly parent of the penetrated object
                PenetrationsServices.SetSlotAssemblyOrientation((Feature)base.Occurrence, base.Penetrated);
            }
            catch (Exception)
            {
                //may be there could be specific ToDo record is created before, so do not override it.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(PenetrationsResourceIds.ToDoParameterRule,
                        "Unexpected error while evaluating the {0}"), this.ToString()));
                }
            }
        }    
    }      
}