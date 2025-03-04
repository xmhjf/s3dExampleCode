//-------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  TeeWeldVParameterRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMPhysConnRules.dll
//  Original Class Name: ‘TeeWeldVParm’ in VB content
//
//Abstract
//     TeeWeldVParameterRule is a .NET parameter rule which is defining the custom parameter rule. 
//  This class subclasses from PhysicalConnectionParameterRule. 
//-----------------------------------------------------------------------------------------------------------

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;
using System;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Parameter rule which defines the custom parameter rule for PhysicalConnections. 
    /// </summary>
    [RuleVersion("1.0.0.0")]
    //New content does not need to specify the interface name, but this is specified here to support code written against old rules which directly uses interface name to get parameter value. 
    [RuleInterface(DetailingCustomAssembliesConstants.IASMPhysConnRules_TeeWeldVParm, DetailingCustomAssembliesConstants.IASMPhysConnRules_TeeWeldVParm)]
    [DefaultLocalizer(ConnectionsResourceIdentifiers.ConnectionsResource, ConnectionsResourceIdentifiers.ConnectionsAssembly)]
    public class TeeWeldVParameterRule : PhysicalConnectionParameterRule
    {
        //==============================================================================================================
        //DefinitionName/ProgID of this symbol is "Connections,Ingr.SP3D.Content.Structure.TeeWeldVParameterRule"
        //==============================================================================================================
        #region Parameters

        [ControlledParameter(1, DetailingCustomAssembliesConstants.MoldedFillet, DetailingCustomAssembliesConstants.MoldedFillet)]
        public ControlledParameterDouble moldedFillet;
        [ControlledParameter(2, DetailingCustomAssembliesConstants.AntiMoldedFillet, DetailingCustomAssembliesConstants.AntiMoldedFillet)]
        public ControlledParameterDouble antiMoldedFillet;
        [ControlledParameter(3, DetailingCustomAssembliesConstants.FilletMeasureMethod, DetailingCustomAssembliesConstants.FilletMeasureMethod)]
        public ControlledParameterDouble filletMeasuredMethod;
        [ControlledParameter(4, DetailingCustomAssembliesConstants.Category, DetailingCustomAssembliesConstants.Category)]
        public ControlledParameterDouble category;
        [ControlledParameter(5, DetailingCustomAssembliesConstants.ReferenceSide, DetailingCustomAssembliesConstants.ReferenceSide)]
        public ControlledParameterString referenceSide;
        [ControlledParameter(6, DetailingCustomAssembliesConstants.RefPartName, DetailingCustomAssembliesConstants.RefPartName)]
        public ControlledParameterString refPartName;
        [ControlledParameter(7, DetailingCustomAssembliesConstants.RefSideFirstBevelDepth, DetailingCustomAssembliesConstants.RefSideFirstBevelDepth)]
        public ControlledParameterDouble refSideFirstBevelDepth;
        [ControlledParameter(8, DetailingCustomAssembliesConstants.RefSideFirstBevelMethod, DetailingCustomAssembliesConstants.RefSideFirstBevelMethod)]
        public ControlledParameterDouble refSideFirstBevelMethod;
        [ControlledParameter(9, DetailingCustomAssembliesConstants.RefSideFirstBevelAngle, DetailingCustomAssembliesConstants.RefSideFirstBevelAngle)]
        public ControlledParameterDouble refSideFirstBevelAngle;
        [ControlledParameter(10, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelDepth, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelDepth)]
        public ControlledParameterDouble antiRefSideFirstBevelDepth;
        [ControlledParameter(11, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelMethod, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelMethod)]
        public ControlledParameterDouble antiRefSideFirstBevelMethod;
        [ControlledParameter(12, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelAngle, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelAngle)]
        public ControlledParameterDouble antiRefSideFirstBevelAngle;
        [ControlledParameter(13, DetailingCustomAssembliesConstants.PrimarySideGroove, DetailingCustomAssembliesConstants.PrimarySideGroove)]
        public ControlledParameterDouble primaySideGroove;
        [ControlledParameter(14, DetailingCustomAssembliesConstants.SecondarySideGroove, DetailingCustomAssembliesConstants.SecondarySideGroove)]
        public ControlledParameterDouble secondarySideGroove;
        [ControlledParameter(15, DetailingCustomAssembliesConstants.PrimarySideGrooveSize, DetailingCustomAssembliesConstants.PrimarySideGrooveSize)]
        public ControlledParameterDouble primarySideGrooveSize;
        [ControlledParameter(16, DetailingCustomAssembliesConstants.SecondarySideGrooveSize, DetailingCustomAssembliesConstants.SecondarySideGrooveSize)]
        public ControlledParameterDouble secondarySideGrooveSize;
        [ControlledParameter(17, DetailingCustomAssembliesConstants.PrimarySideActualThroatThickness, DetailingCustomAssembliesConstants.PrimarySideActualThroatThickness)]
        public ControlledParameterDouble primarySideActualThroatThickness;
        [ControlledParameter(18, DetailingCustomAssembliesConstants.SecondarySideActualThroatThickness, DetailingCustomAssembliesConstants.SecondarySideActualThroatThickness)]
        public ControlledParameterDouble secondarySideActualThroatThickness;
        [ControlledParameter(19, DetailingCustomAssembliesConstants.PrimarySideNominalThroatThickness, DetailingCustomAssembliesConstants.PrimarySideNominalThroatThickness)]
        public ControlledParameterDouble primarySideNominalThroatThickness;
        [ControlledParameter(20, DetailingCustomAssembliesConstants.SecondarySideNominalThroatThickness, DetailingCustomAssembliesConstants.SecondarySideNominalThroatThickness)]
        public ControlledParameterDouble secondarySideNominalThroatThickness;
        [ControlledParameter(21, DetailingCustomAssembliesConstants.PrimarySideLength, DetailingCustomAssembliesConstants.PrimarySideLength)]
        public ControlledParameterDouble primarySideLength;
        [ControlledParameter(22, DetailingCustomAssembliesConstants.SecondarySideLength, DetailingCustomAssembliesConstants.SecondarySideLength)]
        public ControlledParameterDouble secondarySideLength;
        [ControlledParameter(23, DetailingCustomAssembliesConstants.PrimarySidePitch, DetailingCustomAssembliesConstants.PrimarySidePitch)]
        public ControlledParameterDouble primarySidePitch;
        [ControlledParameter(24, DetailingCustomAssembliesConstants.SecondarySidePitch, DetailingCustomAssembliesConstants.SecondarySidePitch)]
        public ControlledParameterDouble secondarySidePitch;
        [ControlledParameter(25, DetailingCustomAssembliesConstants.PrimarySideRootOpening, DetailingCustomAssembliesConstants.PrimarySideRootOpening)]
        public ControlledParameterDouble primarySideRootOpening;
        [ControlledParameter(26, DetailingCustomAssembliesConstants.PrimarySideGrooveAngle, DetailingCustomAssembliesConstants.PrimarySideGrooveAngle)]
        public ControlledParameterDouble primarySideGrooveAngle;
        [ControlledParameter(27, DetailingCustomAssembliesConstants.SecondarySideGrooveAngle, DetailingCustomAssembliesConstants.SecondarySideGrooveAngle)]
        public ControlledParameterDouble secondarySideGrooveAngle;
        [ControlledParameter(28, DetailingCustomAssembliesConstants.TailNotes, DetailingCustomAssembliesConstants.TailNotes)]
        public ControlledParameterString tailNotes;
        [ControlledParameter(29, DetailingCustomAssembliesConstants.TailNoteIsReference, DetailingCustomAssembliesConstants.TailNoteIsReference)]
        public ControlledParameterDouble tailNotesIsReference;
        [ControlledParameter(30, DetailingCustomAssembliesConstants.PrimarySideActualLegLength, DetailingCustomAssembliesConstants.PrimarySideActualLegLength)]
        public ControlledParameterDouble primarySideActualLeglength;
        [ControlledParameter(31, DetailingCustomAssembliesConstants.SecondarySideActualLegLength, DetailingCustomAssembliesConstants.SecondarySideActualLegLength)]
        public ControlledParameterDouble secondarySideActualLegLength;
        [ControlledParameter(32, DetailingCustomAssembliesConstants.PrimarySideSymbol, DetailingCustomAssembliesConstants.PrimarySideSymbol)]
        public ControlledParameterDouble primarySideSymbol;
        [ControlledParameter(33, DetailingCustomAssembliesConstants.SecondarySideSymbol, DetailingCustomAssembliesConstants.SecondarySideSymbol)]
        public ControlledParameterDouble secondarySideSymbol;


        #endregion Parameters

        #region Public override properties and methods
        /// <summary>
        /// Evaluates the parameter rule.
        /// It calculates the item parameters in the context of the Physical Connection.
        /// </summary>
        public override void Evaluate()
        {
            try
            {
                // Get Class Arguments
                PhysicalConnection physicalConnection = (PhysicalConnection)base.Occurrence;

                //Validating inputs
                ValidateInputs();
                //Check if any ToDo message of type error has been created during the validation of inputs
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    //ToDo list is created while validating the inputs
                    return;
                }
                //Get data required for Parameter Rule
                double boundedPartThickness = 0.0, boundingPartThickness = 0.0;
                if (physicalConnection.PartName.Equals(DetailingCustomAssembliesConstants.TeeWeldV))
                {
                    try
                    {
                        ConnectionServices.GetPhysicalConnectionPartsThickness(physicalConnection, out boundedPartThickness, out boundingPartThickness);
                    }
                    catch (Exception ex)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning,
                        base.GetString(ConnectionsResourceIdentifiers.ErrEndCutType,
                        ex.Message + this.GetType().Name + " physicalconnection parameter rule."));
                    }
                }
                else
                {
                    boundedPartThickness = (double)((PropertyValueDouble)physicalConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.ChamferThickness)).PropValue;
                    boundingPartThickness = ConnectionServices.GetBoundingPartThickness(base.BoundingObject);
                }
                //Get the part names
                this.refPartName.Value = physicalConnection.BoundedObject.ToString();
                //set the reference side              
                string referenceSideName = ConnectionServices.GetReferenceSide(physicalConnection, base.BoundedObject);
                string moldedSideProfile = string.Empty;
                this.referenceSide.Value = ConnectionServices.GetMoldedType(referenceSideName);
                bool isRefSideMolded = ((this.referenceSide.Value).Equals(DetailingCustomAssembliesConstants.ReferenceSideAntimolded)) ? false : true;
                // variables to compute the values for the angle and depth on the two sides
                double refSideFirstBevelAngleValue = 0.0, refSideFirstBevelDepthValue = 0.0, antiRefSideFirstBevelAngleValue = 0.0, antiRefSideFirstBevelDepthValue = 0.0;
                //Get the answer for first welding side, i.e Molded or AntiMolded.
                int firstWeldingSide = (int)((PropertyValueCodelist)physicalConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.FirstWeldingSide)).PropValue;
                double mountingAngle = physicalConnection.GetConnectionAngle();
                bool isRefSideOffset = (referenceSideName.Equals(DetailingCustomAssembliesConstants.ReferenceSideOffset) || moldedSideProfile.Equals(DetailingCustomAssembliesConstants.WebRightMoldedSide) || moldedSideProfile.Equals(DetailingCustomAssembliesConstants.TopFlangeBottomFaceMoldedSide) || moldedSideProfile.Equals(DetailingCustomAssembliesConstants.BottomFlangeTopFaceMoldedSide)) ? true : false;
                bool isRefSideBase = (referenceSideName.Equals(DetailingCustomAssembliesConstants.ReferenceSideBase) || moldedSideProfile.Equals(DetailingCustomAssembliesConstants.WebLeftMoldedSide) || moldedSideProfile.Equals(DetailingCustomAssembliesConstants.TopFlangeTopFaceMoldedSide) || moldedSideProfile.Equals(DetailingCustomAssembliesConstants.BottomFlangeBottomFaceMoldedSide)) ? true : false;
                double angle = physicalConnection.BoundedObject is CollarPart ? 0.52359878 : 0.785398;
                //Get BevelMethod answer
                int bevelMethod = (int)((PropertyValueCodelist)physicalConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.BevelAngleMethod)).PropValue;
                if (firstWeldingSide == (int)FirstWeldingSide.MoldedSide)
                {
                    //calculate angles, depending on the bevel method
                    switch (bevelMethod)
                    {
                        case (int)BevelMethod.Constant:
                            refSideFirstBevelAngleValue = (isRefSideOffset) ? 0 : Math.PI - angle - mountingAngle;
                            refSideFirstBevelDepthValue = (isRefSideOffset) ? 0 : boundedPartThickness;
                            antiRefSideFirstBevelAngleValue = (isRefSideOffset) ? Math.PI - angle - mountingAngle : 0;
                            antiRefSideFirstBevelDepthValue = (isRefSideOffset) ? boundedPartThickness : 0;
                            this.refSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Leg;
                            this.antiRefSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Leg;
                            break;
                        case (int)BevelMethod.Varying:
                            refSideFirstBevelAngleValue = (isRefSideOffset) ? 0 : angle;
                            refSideFirstBevelDepthValue = (isRefSideOffset) ? 0 : boundedPartThickness;
                            antiRefSideFirstBevelAngleValue = (isRefSideOffset) ? angle : 0;
                            antiRefSideFirstBevelDepthValue = (isRefSideOffset) ? boundedPartThickness : 0;
                            this.refSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Throat;
                            this.antiRefSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Throat;
                            break;
                    }
                }
                else //the bevel is on the antimolded side
                {
                    //calculate Angle, depending on the bevel method
                    switch (bevelMethod)
                    {
                        case (int)BevelMethod.Constant:
                            refSideFirstBevelAngleValue = (isRefSideOffset) ? Math.PI - angle - mountingAngle : 0;
                            refSideFirstBevelDepthValue = (isRefSideOffset) ? boundedPartThickness : 0;
                            antiRefSideFirstBevelAngleValue = (isRefSideOffset) ? 0 : angle;
                            antiRefSideFirstBevelDepthValue = (isRefSideOffset) ? 0 : boundedPartThickness;
                            this.refSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Leg;
                            this.antiRefSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Leg;
                            break;
                        case (int)BevelMethod.Varying:
                            refSideFirstBevelAngleValue = (isRefSideOffset) ? angle : 0;
                            refSideFirstBevelDepthValue = (isRefSideOffset) ? boundedPartThickness : 0;
                            antiRefSideFirstBevelAngleValue = (isRefSideOffset) ? 0 : angle;
                            antiRefSideFirstBevelDepthValue = (isRefSideOffset) ? 0 : boundedPartThickness;
                            this.refSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Throat;
                            this.antiRefSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Throat;
                            break;
                    }
                }
                //set the actual values in the outputs
                this.refSideFirstBevelDepth.Value = refSideFirstBevelDepthValue;
                this.antiRefSideFirstBevelDepth.Value = antiRefSideFirstBevelDepthValue;
                this.refSideFirstBevelAngle.Value = refSideFirstBevelAngleValue;
                this.antiRefSideFirstBevelAngle.Value = antiRefSideFirstBevelAngleValue;
                //Calculate the fillet size
                double filletValue = boundedPartThickness <= boundingPartThickness ? boundedPartThickness * 0.2 : boundingPartThickness * 0.2;
                double antiMoldedFilletValue = boundedPartThickness <= boundingPartThickness ? boundedPartThickness * 0.2 : boundingPartThickness * 0.2;
                this.moldedFillet.Value = filletValue;
                this.antiMoldedFillet.Value = filletValue;
                this.filletMeasuredMethod.Value = (int)WeldFilletMeasure.Leg;
                double refSideGroove = (int)WeldGroove.Bevel, antiRefSideGroove = (int)WeldGroove.Bevel;
                //Get Category answer
                this.category.Value = (int)((PropertyValueCodelist)physicalConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.Category)).PropValue;
                // fill in the correct values for IJWeldSymbol
                // this method will include a check for any of the input parameters from the bevel
                // that have been overridden by the user
                if (physicalConnection.IsParameterOverridden(DetailingCustomAssembliesConstants.MoldedFillet))
                {
                    filletValue = (double)((PropertyValueDouble)physicalConnection.GetParameterValue(DetailingCustomAssembliesConstants.MoldedFillet)).PropValue;
                }
                if (physicalConnection.IsParameterOverridden(DetailingCustomAssembliesConstants.AntiMoldedFillet))
                {
                    antiMoldedFilletValue = (double)((PropertyValueDouble)physicalConnection.GetParameterValue(DetailingCustomAssembliesConstants.AntiMoldedFillet)).PropValue;
                }
                if (physicalConnection.IsParameterOverridden(DetailingCustomAssembliesConstants.RefSideFirstBevelDepth))
                {
                    refSideFirstBevelDepthValue = (double)((PropertyValueDouble)physicalConnection.GetParameterValue(DetailingCustomAssembliesConstants.RefSideFirstBevelDepth)).PropValue;
                }
                if (physicalConnection.IsParameterOverridden(DetailingCustomAssembliesConstants.RefSideFirstBevelAngle))
                {
                    refSideFirstBevelAngleValue = (double)((PropertyValueDouble)physicalConnection.GetParameterValue(DetailingCustomAssembliesConstants.RefSideFirstBevelAngle)).PropValue;
                }
                if (physicalConnection.IsParameterOverridden(DetailingCustomAssembliesConstants.AntiRefSideFirstBevelDepth))
                {
                    antiRefSideFirstBevelDepthValue = (double)((PropertyValueDouble)physicalConnection.GetParameterValue(DetailingCustomAssembliesConstants.AntiRefSideFirstBevelDepth)).PropValue;
                }
                if (physicalConnection.IsParameterOverridden(DetailingCustomAssembliesConstants.AntiRefSideFirstBevelAngle))
                {
                    antiRefSideFirstBevelAngleValue = (double)((PropertyValueDouble)physicalConnection.GetParameterValue(DetailingCustomAssembliesConstants.AntiRefSideFirstBevelAngle)).PropValue;
                }
                if (physicalConnection.IsParameterOverridden(DetailingCustomAssembliesConstants.ReferenceSide))
                {
                    isRefSideMolded = referenceSide.Value.Equals(DetailingCustomAssembliesConstants.molded) ? true : false;
                }
                if (refSideFirstBevelDepthValue < StructHelper.DISTTOL)
                {
                    refSideFirstBevelAngleValue = 0;
                    refSideGroove = (int)WeldGroove.None;
                }
                if (antiRefSideFirstBevelDepthValue < StructHelper.DISTTOL)
                {
                    antiRefSideFirstBevelDepthValue = 0;
                    antiRefSideGroove = (int)WeldGroove.None;
                }
                this.primarySideLength.Value = 0;
                this.primarySidePitch.Value = 0;
                this.secondarySideLength.Value = 0;
                this.secondarySidePitch.Value = 0;
                this.primarySideRootOpening.Value = 0;
                this.primarySideNominalThroatThickness.Value = filletValue;
                this.secondarySideNominalThroatThickness.Value = antiMoldedFilletValue;
                this.primaySideGroove.Value = (isRefSideMolded) ? refSideGroove : antiRefSideGroove;
                this.primarySideActualThroatThickness.Value = (isRefSideMolded) ? refSideFirstBevelDepthValue : antiRefSideFirstBevelDepthValue;
                this.primarySideGrooveAngle.Value = (isRefSideMolded) ? refSideFirstBevelAngleValue : antiRefSideFirstBevelAngleValue;
                this.secondarySideGroove.Value = (isRefSideMolded) ? antiRefSideGroove : refSideGroove;
                this.secondarySideActualThroatThickness.Value = (isRefSideMolded) ? antiRefSideFirstBevelDepthValue : refSideFirstBevelDepthValue;
                this.secondarySideGrooveAngle.Value = (isRefSideMolded) ? antiRefSideFirstBevelAngleValue : refSideFirstBevelAngleValue;
                this.primarySideSymbol.Value = filletValue > 0 ? (int)WeldSymbol.Fillet : (int)WeldSymbol.None;
                this.secondarySideSymbol.Value = filletValue > 0 ? (int)WeldSymbol.Fillet : (int)WeldSymbol.None;
                this.primarySideGrooveSize.Value = 0.0;
                this.secondarySideGrooveSize.Value = 0.0;
                this.tailNotes.Value = string.Empty;
                this.tailNotesIsReference.Value = 0.0;
                this.primarySideActualLeglength.Value = 0.0;
                this.secondarySideActualLegLength.Value = 0.0;

            }
            catch (Exception)
            {
                //There could be specific ToDo records created before, so do not override it.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(ConnectionsResourceIdentifiers.ToDoPhysicalConnectionParameterRule,
                        "Unexpected error while evaluating the {0}"), this.ToString()));
                }
            }
        }
        #endregion Public override properties and methods
    }
}

