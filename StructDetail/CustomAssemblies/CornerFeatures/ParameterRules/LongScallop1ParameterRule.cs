//********************************************************************************************************************/
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  LongScallop1ParameterRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMCornerFeatRules.dll
//  Original Class Name: ‘LongScallop1Parm’ in VB content
//
//Abstract
//	LongScallop1ParameterRule is a .NET parameter rule which is defining the custom parameter rule. 
//  This class subclasses from CornerParameterRule.
//
// Change History:
//  dd.mmm.yyyy    who    change description
//********************************************************************************************************************/
using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Parameter rule for the LongScallop1, which defines the custom parameter rule in the context of Corner Feature.
    /// It computes the item parameters for the corner feature and assigns them.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [DefaultLocalizer(CornerFeaturesResourceIdentifiers.CornerFeaturesResources, CornerFeaturesResourceIdentifiers.CornerFeaturesAssembly)]
    //New content does not need to specify the interface name, but this is specified here to support code written against old rules which directly uses interface name to get selector question. 
    [RuleInterface(DetailingCustomAssembliesConstants.IASMCornerFeatRules_LongScallop1Parm, DetailingCustomAssembliesConstants.IASMCornerFeatRules_LongScallop1Parm)]
    public class LongScallop1ParameterRule : CornerParameterRule
    {
        //==============================================================================================================
        //DefinitionName/ProgID of this symbol is "CornerFeatures,Ingr.SP3D.Content.Structure.LongScallop1ParameterRule"
        //==============================================================================================================

        #region Parameters
        [ControlledParameter(1, DetailingCustomAssembliesConstants.Ulength, DetailingCustomAssembliesConstants.Ulength)]
        public ControlledParameterDouble uLength;
        [ControlledParameter(2, MarineSymbolConstants.Radius, MarineSymbolConstants.Radius)]
        public ControlledParameterDouble radius;
        [ControlledParameter(3, DetailingCustomAssembliesConstants.Flip, DetailingCustomAssembliesConstants.Flip)]
        public ControlledParameterDouble flip;
        #endregion Parameters

        #region Public override properties and methods

        /// <summary>
        /// Evaluates the parameter rule.
        /// It computes the parameters in the context of the Corner Feature.
        /// </summary>
        public override void Evaluate()
        {
            try
            {
                //Validates the inputs of the corner feature
                base.ValidateInputs();

                //Check if any ToDo message of type error has been created during the validation of inputs
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    //ToDo list is created while validating the inputs
                    return;
                }

                //Set default parameter values
                this.uLength.Value = 0.0;
                this.radius.Value = 0.05; //default (50mm)
                this.flip.Value = 0.0; //default NoFlip = 0, Flip = 1

                //Get part name
                Feature cornerFeature = (Feature)base.Occurrence;
                string partName = cornerFeature.PartName;

                //If the seam is found, then set the parameter values according to IJUACornerFeatureSeam
                int seamFound = StructHelper.GetIntProperty(cornerFeature, DetailingCustomAssembliesConstants.IJUACornerFeatureSeam, DetailingCustomAssembliesConstants.SeamFound);
                if (seamFound == 1 && (partName == DetailingCustomAssembliesConstants.LongScallopWithSeam || partName == DetailingCustomAssembliesConstants.LongScallopWithSeamWithCollar))
                {
                    this.flip.Value = StructHelper.GetIntProperty(cornerFeature, DetailingCustomAssembliesConstants.IJUACornerFeatureSeam, DetailingCustomAssembliesConstants.sFlip);
                    this.uLength.Value = StructHelper.GetDoubleProperty(cornerFeature, DetailingCustomAssembliesConstants.IJUACornerFeatureSeam, DetailingCustomAssembliesConstants.DistToSeam) + 0.015; //should be the same as the search distance
                    this.radius.Value = 0.05; //default (50mm)
                }
                else
                {
                    //Set general corner size
                    double gapTolerance = 0.003; //will not find gaps < 3mm
                    double gapAlongEdgeU, gapAlongEdgeV;

                    //Check if the corner gap exists, if there any gap existing, then drive the parameter values accordingly.
                    if (cornerFeature.DoesCornerGapExist(gapTolerance, out gapAlongEdgeU, out gapAlongEdgeV) && partName != DetailingCustomAssembliesConstants.LongScallop50x100 && partName != DetailingCustomAssembliesConstants.LongScallopAlongCorner50x100)
                    {
                        if (gapAlongEdgeU > 0.0)
                        {
                            //the Ulength value will be increased, no flip
                            this.uLength.Value = gapAlongEdgeU + 0.015;
                        }
                        else if (gapAlongEdgeV > 0.0)
                        {
                            //the Ulength value will be increased, flip
                            this.uLength.Value = gapAlongEdgeV + 0.015;
                            //flip the corner feature
                            this.flip.Value = 1.0;
                        }
                        this.radius.Value = 0.05; //default (50mm)
                    }
                    else
                    {
                        //Get corner "Filp" answer
                        int cornerFlip = ((PropertyValueCodelist)cornerFeature.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.CornerOrientation)).PropValue;
                        if (cornerFlip == (int)Flip.Yes)
                        {
                            this.flip.Value = 1.0;
                        }

                        //Get connectable on which corner feature is placed
                        IConnectable facePortConnectable = base.FacePort.Connectable;
                        if (facePortConnectable is PlatePartBase)
                        {
                            this.uLength.Value = 0.1; //default (100mm)
                            this.radius.Value = 0.05; //default (50mm)
                        }
                        else if (facePortConnectable is ICrossSection)
                        {
                            if (partName == DetailingCustomAssembliesConstants.LongScallop50x100 || partName == DetailingCustomAssembliesConstants.LongScallopAlongCorner50x100)
                            {
                                this.uLength.Value = 0.1; //default (100mm)
                                this.radius.Value = 0.05; //default (50mm)
                            }
                            else
                            {
                                CrossSection crossSection = ((ICrossSection)facePortConnectable).CrossSection;
                                double webLength = 0.0;
                                if (facePortConnectable is StiffenerPartBase)
                                {
                                    //Get web length from StiffenerPartBase cross-section
                                    webLength = DetailingCustomAssembliesServices.GetWebLength(crossSection);
                                }
                                else if (facePortConnectable is MemberPart)
                                {
                                    //Get web depth from MemberPart cross-section
                                    webLength = DetailingCustomAssembliesServices.GetWebDepth(crossSection);
                                    if (webLength < StructHelper.DISTTOL)
                                    {
                                        webLength = crossSection.Depth;
                                    }
                                }

                                if (webLength >= 0.0 || webLength <= 0.2)
                                {
                                    this.uLength.Value = 0.07; //default (70mm)
                                    this.radius.Value = 0.035; //default (35mm)
                                }
                                else if (webLength >= 0.2 || webLength <= 0.4)
                                {
                                    this.uLength.Value = 0.1; //default (100mm)
                                    this.radius.Value = 0.05; //default (50mm)
                                }
                                else if (webLength >= 0.4)
                                {
                                    this.uLength.Value = 0.15; //default (150mm)
                                    this.radius.Value = 0.075; //default (75mm)
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                //There could be specific ToDo records created before, so do not override it.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(CornerFeaturesResourceIdentifiers.ToDoParameterRule,
                        "Unexpected error while evaluating the {0}"), this.ToString()));
                }
            }
        }

        #endregion Public override properties and methods
    }
}