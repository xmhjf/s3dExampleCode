/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  SlotTC_L_PAT_DTR_ParameterRule.cs
//
//Abstract
//	SlotTC_L_PAT_DTR_ParameterRule is a .NET parameter rule which is defining the custom parameter rule. 
//  This class subclasses from SlotParameterRuleBase.
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMSlotRules.dll
//  Original Class Name: ‘SlotTC_L_PAT_DTR_Parm’ in VB content
//
//  Parameter rule which defining the custom parameter rule.
//  TC - BUT cross section and Slot C type.
//  L - Left connected.
//  PAT - Connected at left bottom, left top, right top corner (Point-Arc-Tangent)
//  DTR - Different tangent radius.
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Content.Structure.Services;

namespace Ingr.SP3D.Content.Structure
{  
    /// <summary>
    /// Parameter rule which defines the custom parameter rule for Slot C type and left connected penetrations with BUT cross section.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    //New content does not need to specify the interface name, but this is specified here to support code written against old rules which directly uses interface name to get parameter value. 
    [RuleInterface(DetailingCustomAssembliesConstants.IASMSlotRules_SlotTC_L_PAT_DTR_Parm, DetailingCustomAssembliesConstants.IASMSlotRules_SlotTC_L_PAT_DTR_Parm, DetailingCustomAssembliesConstants.IASlotTC_L_PAT_DTR_ParmV)]
    [DefaultLocalizer(PenetrationsResourceIds.PenetrationsResource, PenetrationsResourceIds.PenetrationsAssembly)]
    public class SlotTC_L_PAT_DTR_ParameterRule : SlotParameterRule
    {
        //=================================================================================================================
        //DefinitionName/ProgID of this symbol is "Penetrations,Ingr.SP3D.Content.Structure.SlotTC_L_PAT_DTR_ParameterRule"
        //=================================================================================================================

        //These are the parameters which will be evaluated by this rule.
        [ControlledParameter(1, DetailingCustomAssembliesConstants.TopFlangeLeftClearance, DetailingCustomAssembliesConstants.TopFlangeLeftClearance)]
        public ControlledParameterDouble topFlangeLeftClearance;
        [ControlledParameter(2, DetailingCustomAssembliesConstants.TopFlangeLeftTopCornerRadius, DetailingCustomAssembliesConstants.TopFlangeLeftTopCornerRadius)]
        public ControlledParameterDouble topFlangeLeftTopCornerRadius;
        [ControlledParameter(3, DetailingCustomAssembliesConstants.TopClearance, DetailingCustomAssembliesConstants.TopClearance)]
        public ControlledParameterDouble topClearance;
        [ControlledParameter(4, DetailingCustomAssembliesConstants.TopFlangeRightTopCornerRadius, DetailingCustomAssembliesConstants.TopFlangeRightTopCornerRadius)]
        public ControlledParameterDouble topFlangeRightTopCornerRadius;
        [ControlledParameter(5, DetailingCustomAssembliesConstants.SlotAngle, DetailingCustomAssembliesConstants.SlotAngle)]
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
                ValidateInputs();

                //Check if any ToDo message of type error has been created during the validation of inputs
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    //ToDo list is created while validating the inputs
                    return;
                }
                
                //Gets the slot angle onpassing obtained slot
                slotAngle.Value = PenetrationsServices.GetSlotAngle(slot);
                
                if (base.Penetrating is PlatePart)
                {
                    //Gets the penetrating width on passing penetrated and penetrating objects(plate part)
                    double profilePartWidth = Math.Round(base.SlotMappingRule.GetSectionWidth(base.Penetrating, base.Penetrated), 3);
                    // Set topFlangeLeftTopCornerRadius,topClearance,topFlangeRightTopCornerRadius values based on the penetrating width
                    // as one geometry is not appropriate for all the cross sections and its values must be changing based on the input geometry attribute changes.                    
                    if (profilePartWidth >= 0.05 && profilePartWidth < 0.07)
                    {
                        topFlangeLeftClearance.Value = 0.02;
                        topFlangeLeftTopCornerRadius.Value = 0.04;
                        topClearance.Value = 0.025;
                        topFlangeRightTopCornerRadius.Value = 0.04;
                    }
                    else if (profilePartWidth >= 0.07 && profilePartWidth < 0.1)
                    {
                        topFlangeLeftClearance.Value = 0.02;
                        topFlangeLeftTopCornerRadius.Value = 0.05;
                        topClearance.Value = 0.04;
                        topFlangeRightTopCornerRadius.Value = 0.05;
                    }
                    else if (profilePartWidth >= 0.1 && profilePartWidth < 0.2)
                    {
                        topFlangeLeftClearance.Value = 0.04;
                        topFlangeLeftTopCornerRadius.Value = 0.065;
                        topClearance.Value = 0.05;
                        topFlangeRightTopCornerRadius.Value = 0.1;
                    }
                    else if (profilePartWidth >= 0.2 && profilePartWidth <= 0.35)
                    {
                        topFlangeLeftClearance.Value = 0.05;
                        topFlangeLeftTopCornerRadius.Value = 0.1;
                        topClearance.Value = 0.05;
                        topFlangeRightTopCornerRadius.Value = 0.15;
                    }
                    else
                    {
                        topFlangeLeftClearance.Value = 0.075;
                        topFlangeLeftTopCornerRadius.Value = 0.15;
                        topClearance.Value = 0.075;
                        topFlangeRightTopCornerRadius.Value = 0.2;
                    }
                }
                else if (base.Penetrating is ProfilePart) 
                {
                    //Gets the penetrating height on passing penetrated and penetrating objects(profile part)
                    double profilePartHeight = Math.Round(base.GetPenetratingHeight(MarineSymbolConstants.IJUAXSectionWeb, MarineSymbolConstants.WebLength),3);
                    // Set topFlangeLeftTopCornerRadius,topClearance,topFlangeRightTopCornerRadius values based on the penetrating height
                    // as one geometry is not appropriate for all the cross sections and its values must be changing based on the input geometry attribute changes.
                    if (profilePartHeight > 0.1 && profilePartHeight <= 0.2)
                    {
                        topFlangeLeftClearance.Value = 0.02;
                        topFlangeLeftTopCornerRadius.Value = 0.04;
                        topClearance.Value = 0.025;
                        topFlangeRightTopCornerRadius.Value = 0.04;
                    }
                    else if (profilePartHeight > 0.2 && profilePartHeight <= 0.3)
                    {
                        topFlangeLeftClearance.Value = 0.02;
                        topFlangeLeftTopCornerRadius.Value = 0.05;
                        topClearance.Value = 0.04;
                        topFlangeRightTopCornerRadius.Value = 0.05;
                    }
                    else if (profilePartHeight > 0.3 && profilePartHeight <= 0.6)
                    {
                        topFlangeLeftClearance.Value = 0.04;
                        topFlangeLeftTopCornerRadius.Value = 0.065;
                        topClearance.Value = 0.05;
                        topFlangeRightTopCornerRadius.Value = 0.1;
                    }
                    else
                    {
                        topFlangeLeftClearance.Value = 0.05;
                        topFlangeLeftTopCornerRadius.Value = 0.1;
                        topClearance.Value = 0.05;
                        topFlangeRightTopCornerRadius.Value = 0.15;
                    }
                }

                //Set the slot assembly orientation to set the orientation of the slot depending upon the orientation of the assembly parent of the penetrated object
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