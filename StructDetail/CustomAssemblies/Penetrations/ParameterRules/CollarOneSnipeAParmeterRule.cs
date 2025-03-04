//************************************************************************************************************/
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  CollarOneSnipeAParmeterRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMCollarRules.dll
//  Original Class Name: ‘CollarOneSnipeAParm’ in VB content
//
//Abstract
//	CollarOneSnipeAParmeterRule is a .NET parameter rule which is defining the custom parameter rule. 
//  This class subclasses from CollarParameterRule.
//
// Change History:
//  dd.mmm.yyyy    who    change description
//************************************************************************************************************/
using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Parameter rule which defines the custom parameter rule. 
    /// </summary>
    [RuleVersion("1.0.0.0")]
    //New content does not need to specify the interface name, but this is specified here to support code written against old rules which directly uses interface name to get parameter value. 
    [RuleInterface(DetailingCustomAssembliesConstants.IASMCollarRules_CollarOneSnipeAParm, DetailingCustomAssembliesConstants.IASMCollarRules_CollarOneSnipeAParm)]
    [DefaultLocalizer(PenetrationsResourceIds.PenetrationsResource, PenetrationsResourceIds.PenetrationsAssembly)]
    public class CollarOneSnipeAParmeterRule : CollarParameterRule
    {
        //==============================================================================================================
        //DefinitionName/ProgID of this symbol is "Penetrations,Ingr.SP3D.Content.Structure.CollarOneSnipeAParmeterRule"
        //==============================================================================================================

        #region Parameters
        [ControlledParameter(1, DetailingCustomAssembliesConstants.SideOfPart, DetailingCustomAssembliesConstants.SideOfPart)]
        public ControlledParameterDouble sideOfPart;
        [ControlledParameter(2, DetailingCustomAssembliesConstants.LapDistance, DetailingCustomAssembliesConstants.LapDistance)]
        public ControlledParameterDouble lapDistance;
        [ControlledParameter(3, DetailingCustomAssembliesConstants.InnerCornerRadius, DetailingCustomAssembliesConstants.InnerCornerRadius)]
        public ControlledParameterDouble innerCornerRadius;
        [ControlledParameter(4, DetailingCustomAssembliesConstants.OuterCornerRadius, DetailingCustomAssembliesConstants.OuterCornerRadius)]
        public ControlledParameterDouble outerCornerRadius;
        #endregion Parameters

        #region Public override properties and methods

        /// <summary>
        /// Evaluates the parameter rule.
        /// It computes the parameters in the context of the CollarPart.
        /// </summary>
        public override void Evaluate()
        {
            try
            {  
                //Check the CollarPart supported Feature.
                base.ValidateInputs();

                //Check if any ToDo message of type error has been created during the validation of inputs
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    //ToDo list is created while validating the inputs
                    return;
                }

                //Get CollarSideOfPlate answer
                CollarPart collarPart = (CollarPart)base.Occurrence;
                PropertyValue collarSideOfPlatePropertyValue = collarPart.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.CollarSideOfPlate);
                int collarSideOfPlate = (collarSideOfPlatePropertyValue != null) ? ((PropertyValueCodelist)collarSideOfPlatePropertyValue).PropValue : 0;

                //set the SideOfPart value based on the CollarSideOfPlate answer and MoldedSide of penetrated object
                ContextTypes mostStiffenerSide = ContextTypes.Base;

                PlatePartBase penetratedPlate = this.Penetrated as PlatePartBase;
               
                //If the penetrated object is a profile, stiffener side is base. Will skip the below code and return true.
                if (penetratedPlate != null)
                {
                    mostStiffenerSide = ((PlatePart)penetratedPlate).MostStiffenerSide;
                }

                int sideOfPart = (int)SideOfPart.Molded;
                switch (collarSideOfPlate)
                {
                    case (int)CollarSide.NoFlip:
                        if (mostStiffenerSide == ContextTypes.Base)
                        {
                            sideOfPart = (int)SideOfPart.AntiMolded;
                        }
                        break;
                    case (int)CollarSide.Flip:
                        if (mostStiffenerSide == ContextTypes.Offset)
                        {
                            sideOfPart = (int)SideOfPart.AntiMolded;
                        }
                        break;
                    case (int)CollarSide.Centered:
                        sideOfPart = (int)SideOfPart.Centered;
                        break;
                }

                //Get the CollarPart’ s PartName and set the section face type accordingly
                int sectionFaceType = (collarPart.PartName.Contains("_B")) ? (int)SectionFaceType.Web_Left : (int)SectionFaceType.Web_Right;

                //Get slot profile penetrated angle
                //this angle can be used to calculate adjusted CollarPart width
                double slotProfilePenetratedAngle = base.GetSlotProfilePenetratedAngle(sideOfPart, sectionFaceType);

                double lapDistance = 0.055;
                double innerCornerRadius = 0.025;
                double outerCornerRadius = 0.09;

                //Calculate other parameters based on profile height
                //Calculate Inner/Outer corner radius based on angle between profile and stiffened plate
                string penetratingSectionTypeName = base.PenetratingSectionTypeName;
                if (string.IsNullOrWhiteSpace(penetratingSectionTypeName))
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(PenetrationsResourceIds.ErrGetPenetratingSectionTypeName,
                        "Unable to get penetrating section type name."));

                    return;
                }

                //Get penetrated object thickness
                double penetratedThickness = base.PenetratedThickness;

                //Get clip thickness
                double clipThickness = StructHelper.GetDoubleProperty(collarPart, DetailingCustomAssembliesConstants.IJUAClipProps, DetailingCustomAssembliesConstants.ClipThickness);

                bool calculateInnerRadius = true;
                double distance = 0.0;
                switch (penetratingSectionTypeName)
                {
                    case MarineSymbolConstants.EA:
                    case MarineSymbolConstants.UA:
                    case MarineSymbolConstants.BUTL2:
                    case MarineSymbolConstants.BUT:
                        distance = 0.1335 - (penetratedThickness + clipThickness) * Math.Sin(slotProfilePenetratedAngle) / 2.0;
                        break;
                    case MarineSymbolConstants.B:
                        distance = 0.1785 - base.PenetratingSectionWidth - (penetratedThickness + clipThickness) * Math.Sin(slotProfilePenetratedAngle) / 2.0;                        
                        break;
                    case MarineSymbolConstants.FB:
                        PlatePartBase penetratingPlate = base.Penetrating as PlatePartBase;
                        double penetratingSectionWidth = 0.0;
                        if (penetratingPlate != null)
                        {
                            penetratingSectionWidth = base.PenetratingSectionWidth;
                        }
                        distance = 0.126 - penetratingSectionWidth - (penetratedThickness + clipThickness) * Math.Sin(slotProfilePenetratedAngle) / 2.0;                        
                        break;
                    default:
                        calculateInnerRadius = false;
                        lapDistance = 0.04;
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning,
                            string.Format(base.GetString(PenetrationsResourceIds.ErrInvalidParamRule,
                            "$1 should not be evaluated for a profile with $2 section type name. Check if the correct catalog part was selected by the selection rule."),
                            this.ToString(), penetratingSectionTypeName));
                        break;
                }

                if (calculateInnerRadius)
                {
                    //gets the angle between profile top and web right, this angle can be used to calculate adjusted CollarPart width
                    double slotTopWebAngle = base.GetSlotTopWebAngle();

                    double delta = distance / Math.Sin(slotTopWebAngle) - innerCornerRadius / Math.Tan(slotTopWebAngle / 2.0) - outerCornerRadius * Math.Tan(slotTopWebAngle / 2.0);
                    if (delta < 0.003)
                    {
                        innerCornerRadius = innerCornerRadius - 0.003 + delta;
                    }
                }

                this.sideOfPart.Value = sideOfPart;
                this.lapDistance.Value = lapDistance;
                this.innerCornerRadius.Value = innerCornerRadius;
                this.outerCornerRadius.Value = outerCornerRadius;

            }
            catch (Exception)
            {
                //There could be specific ToDo records created before, so do not override it.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(PenetrationsResourceIds.ToDoParameterRule,
                        "Unexpected error while evaluating the {0}"), this.ToString()));
                }
            }
        }

        #endregion Public override properties and methods
    }
}