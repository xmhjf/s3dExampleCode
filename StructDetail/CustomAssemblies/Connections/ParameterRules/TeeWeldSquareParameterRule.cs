//-------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  TeeWeldSquareParameterRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMPhysConnRules.dll
//  Original Class Name: ‘TeeWeldSquareParm’ in VB content
//
//Abstract
//	TeeWeldSquareParameterRule is a .NET parameter rule which is defining the custom parameter rule. 
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
    [RuleInterface(DetailingCustomAssembliesConstants.IASMPhysConnRules_TeeWeldSquareParm, DetailingCustomAssembliesConstants.IASMPhysConnRules_TeeWeldSquareParm)]
    [DefaultLocalizer(ConnectionsResourceIdentifiers.ConnectionsResource, ConnectionsResourceIdentifiers.ConnectionsAssembly)]
    public class TeeWeldSquareParameterRule : PhysicalConnectionParameterRule
    {
        //==============================================================================================================
        //DefinitionName/ProgID of this symbol is "Connections,Ingr.SP3D.Content.Structure.TeeWeldSquareParameterRule"
        //==============================================================================================================
        #region Parameters
        [ControlledParameter(1, DetailingCustomAssembliesConstants.Nose, DetailingCustomAssembliesConstants.Nose)]
        public ControlledParameterDouble nose;
        [ControlledParameter(2, DetailingCustomAssembliesConstants.NoseMethod, DetailingCustomAssembliesConstants.NoseMethod)]
        public ControlledParameterDouble noseMethod;
        [ControlledParameter(3, DetailingCustomAssembliesConstants.NoseOrientationAngle, DetailingCustomAssembliesConstants.NoseOrientationAngle)]
        public ControlledParameterDouble noseOrientationAngle;
        [ControlledParameter(4, DetailingCustomAssembliesConstants.RefSideFirstBevelDepth, DetailingCustomAssembliesConstants.RefSideFirstBevelDepth)]
        public ControlledParameterDouble refSideFirstBevelDepth;
        [ControlledParameter(5, DetailingCustomAssembliesConstants.RefSideFirstBevelMethod, DetailingCustomAssembliesConstants.RefSideFirstBevelMethod)]
        public ControlledParameterDouble refSideFirstBevelMethod;
        [ControlledParameter(6, DetailingCustomAssembliesConstants.RefSideFirstBevelAngle, DetailingCustomAssembliesConstants.RefSideFirstBevelAngle)]
        public ControlledParameterDouble refSideFirstBevelAngle;
        [ControlledParameter(7, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelDepth, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelDepth)]
        public ControlledParameterDouble antiRefSideFirstBevelDepth;
        [ControlledParameter(8, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelMethod, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelMethod)]
        public ControlledParameterDouble antiRefSideFirstBevelMethod;
        [ControlledParameter(9, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelAngle, DetailingCustomAssembliesConstants.AntiRefSideFirstBevelAngle)]
        public ControlledParameterDouble antiRefSideFirstBevelAngle;
        [ControlledParameter(10, DetailingCustomAssembliesConstants.MoldedFillet, DetailingCustomAssembliesConstants.MoldedFillet)]
        public ControlledParameterDouble moldedFillet;
        [ControlledParameter(11, DetailingCustomAssembliesConstants.AntiMoldedFillet, DetailingCustomAssembliesConstants.AntiMoldedFillet)]
        public ControlledParameterDouble antiMoldedFillet;
        [ControlledParameter(12, DetailingCustomAssembliesConstants.FilletMeasureMethod, DetailingCustomAssembliesConstants.FilletMeasureMethod)]
        public ControlledParameterDouble filletMeasuredMethod;
        [ControlledParameter(13, DetailingCustomAssembliesConstants.Category, DetailingCustomAssembliesConstants.Category)]
        public ControlledParameterDouble category;
        [ControlledParameter(14, DetailingCustomAssembliesConstants.ReferenceSide, DetailingCustomAssembliesConstants.ReferenceSide)]
        public ControlledParameterString referenceSide;
        [ControlledParameter(15, DetailingCustomAssembliesConstants.RefPartName, DetailingCustomAssembliesConstants.RefPartName)]
        public ControlledParameterString refPartName;
        [ControlledParameter(16, DetailingCustomAssembliesConstants.PrimarySideGroove, DetailingCustomAssembliesConstants.PrimarySideGroove)]
        public ControlledParameterDouble primaySideGroove;
        [ControlledParameter(17, DetailingCustomAssembliesConstants.SecondarySideGroove, DetailingCustomAssembliesConstants.SecondarySideGroove)]
        public ControlledParameterDouble secondarySideGroove;
        [ControlledParameter(18, DetailingCustomAssembliesConstants.PrimarySideGrooveSize, DetailingCustomAssembliesConstants.PrimarySideGrooveSize)]
        public ControlledParameterDouble primarySideGrooveSize;
        [ControlledParameter(19, DetailingCustomAssembliesConstants.SecondarySideGrooveSize, DetailingCustomAssembliesConstants.SecondarySideGrooveSize)]
        public ControlledParameterDouble secondarySideGrooveSize;
        [ControlledParameter(20, DetailingCustomAssembliesConstants.PrimarySideActualThroatThickness, DetailingCustomAssembliesConstants.PrimarySideActualThroatThickness)]
        public ControlledParameterDouble primarySideActualThroatThickness;
        [ControlledParameter(21, DetailingCustomAssembliesConstants.SecondarySideActualThroatThickness, DetailingCustomAssembliesConstants.SecondarySideActualThroatThickness)]
        public ControlledParameterDouble secondarySideActualThroatThickness;
        [ControlledParameter(22, DetailingCustomAssembliesConstants.PrimarySideNominalThroatThickness, DetailingCustomAssembliesConstants.PrimarySideNominalThroatThickness)]
        public ControlledParameterDouble primarySideNominalThroatThickness;
        [ControlledParameter(23, DetailingCustomAssembliesConstants.SecondarySideNominalThroatThickness, DetailingCustomAssembliesConstants.SecondarySideNominalThroatThickness)]
        public ControlledParameterDouble secondarySideNominalThroatThickness;
        [ControlledParameter(24, DetailingCustomAssembliesConstants.PrimarySideLength, DetailingCustomAssembliesConstants.PrimarySideLength)]
        public ControlledParameterDouble primarySideLength;
        [ControlledParameter(25, DetailingCustomAssembliesConstants.SecondarySideLength, DetailingCustomAssembliesConstants.SecondarySideLength)]
        public ControlledParameterDouble secondarySideLength;
        [ControlledParameter(26, DetailingCustomAssembliesConstants.PrimarySidePitch, DetailingCustomAssembliesConstants.PrimarySidePitch)]
        public ControlledParameterDouble primarySidePitch;
        [ControlledParameter(27, DetailingCustomAssembliesConstants.SecondarySidePitch, DetailingCustomAssembliesConstants.SecondarySidePitch)]
        public ControlledParameterDouble secondarySidePitch;
        [ControlledParameter(28, DetailingCustomAssembliesConstants.PrimarySideRootOpening, DetailingCustomAssembliesConstants.PrimarySideRootOpening)]
        public ControlledParameterDouble primarySideRootOpening;
        [ControlledParameter(29, DetailingCustomAssembliesConstants.PrimarySideGrooveAngle, DetailingCustomAssembliesConstants.PrimarySideGrooveAngle)]
        public ControlledParameterDouble primarySideGrooveAngle;
        [ControlledParameter(30, DetailingCustomAssembliesConstants.SecondarySideGrooveAngle, DetailingCustomAssembliesConstants.SecondarySideGrooveAngle)]
        public ControlledParameterDouble secondarySideGrooveAngle;
        [ControlledParameter(31, DetailingCustomAssembliesConstants.TailNotes, DetailingCustomAssembliesConstants.TailNotes)]
        public ControlledParameterString tailNotes;
        [ControlledParameter(32, DetailingCustomAssembliesConstants.TailNoteIsReference, DetailingCustomAssembliesConstants.TailNoteIsReference)]
        public ControlledParameterDouble tailNotesIsReference;
        [ControlledParameter(33, DetailingCustomAssembliesConstants.PrimarySideActualLegLength, DetailingCustomAssembliesConstants.PrimarySideActualLegLength)]
        public ControlledParameterDouble primarySideActualLeglength;
        [ControlledParameter(34, DetailingCustomAssembliesConstants.SecondarySideActualLegLength, DetailingCustomAssembliesConstants.SecondarySideActualLegLength)]
        public ControlledParameterDouble secondarySideActualLegLength;
        [ControlledParameter(35, DetailingCustomAssembliesConstants.PrimarySideSymbol, DetailingCustomAssembliesConstants.PrimarySideSymbol)]
        public ControlledParameterDouble primarySideSymbol;
        [ControlledParameter(36, DetailingCustomAssembliesConstants.SecondarySideSymbol, DetailingCustomAssembliesConstants.SecondarySideSymbol)]
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
                if (physicalConnection.PartName.Equals(DetailingCustomAssembliesConstants.TeeWeldSquare))
                {
                    try
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
                    catch (Exception ex)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning,
                        base.GetString(ConnectionsResourceIdentifiers.ToDoPhysicalConnectionParameterRule,
                        ex.Message + this.GetType().Name + " physicalconnection parameter rule."));
                    }
                }
                else
                {
                    boundedPartThickness = (double)((PropertyValueDouble)physicalConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.ChamferThickness)).PropValue;
                    boundingPartThickness = ConnectionServices.GetBoundingPartThickness(base.BoundingObject);
                }

                //calculate NoseOrientationAngle, depending on the bevel method 
                this.noseOrientationAngle.Value = Math.PI / 2;
                this.noseMethod.Value = (int)BevelMethod.Constant;
                this.nose.Value = boundedPartThickness;
                //Get the part names
                this.refPartName.Value = physicalConnection.BoundedObject.ToString();
                //get the reference side
                string referenceSideName = ConnectionServices.GetReferenceSide(physicalConnection, base.BoundedObject);
                double moldedFilletValue = 0, antiMoldedFilletValue = 0, refSideFirstBevelAngleValue = 0, antiRefSideFirstBevelAngleValue = 0, refSideFirstBevelDepthValue = 0, antiRefSideFirstBevelDepthValue = 0;
                this.referenceSide.Value = ConnectionServices.GetMoldedType(referenceSideName);
                bool isRefSideMolded = ((this.referenceSide.Value).Equals(DetailingCustomAssembliesConstants.ReferenceSideAntimolded)) ? false : true;
                moldedFilletValue = boundedPartThickness * DetailingCustomAssembliesConstants.MinimunTolerance * 82;
                antiMoldedFilletValue = boundedPartThickness * DetailingCustomAssembliesConstants.MinimunTolerance * 82;
                this.moldedFillet.Value = moldedFilletValue;
                this.antiMoldedFillet.Value = antiMoldedFilletValue;
                this.filletMeasuredMethod.Value = (int)WeldFilletMeasure.Leg;
                string categoryAnswer = physicalConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.Category).ToString();
                //Get Category answer
                this.category.Value = (int)((PropertyValueCodelist)physicalConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.Category)).PropValue;
                this.refSideFirstBevelDepth.Value = 0.0;
                this.refSideFirstBevelAngle.Value = 0.0;
                this.refSideFirstBevelMethod.Value = (int)BevelMethod.Constant;
                this.antiRefSideFirstBevelDepth.Value = 0.0;
                this.antiRefSideFirstBevelAngle.Value = 0.0;
                this.antiRefSideFirstBevelMethod.Value = (int)BevelMethod.Constant;
                // fill in the correct values for IJWeldSymbol
                // this method will include a check for any of the input parameters from the bevel
                // that have been overridden by the user
                if (physicalConnection.IsParameterOverridden(DetailingCustomAssembliesConstants.MoldedFillet))
                {
                    moldedFilletValue = (double)((PropertyValueDouble)physicalConnection.GetParameterValue(DetailingCustomAssembliesConstants.MoldedFillet)).PropValue;
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
                }
                if (antiRefSideFirstBevelDepthValue < StructHelper.DISTTOL)
                {
                    antiRefSideFirstBevelDepthValue = 0;
                }
                this.primarySideLength.Value = 0;
                this.primarySidePitch.Value = 0;
                this.secondarySideLength.Value = 0;
                this.secondarySidePitch.Value = 0;
                this.primarySideRootOpening.Value = 0;
                this.primarySideNominalThroatThickness.Value = moldedFilletValue;
                this.secondarySideNominalThroatThickness.Value = antiMoldedFilletValue;
                this.primaySideGroove.Value = 0;
                this.primarySideGrooveSize.Value = 0.0;
                this.secondarySideGrooveSize.Value = 0.0;
                this.primarySideActualThroatThickness.Value = isRefSideMolded ? refSideFirstBevelDepthValue : antiRefSideFirstBevelDepthValue;
                this.primarySideGrooveAngle.Value = isRefSideMolded ? refSideFirstBevelAngleValue : antiRefSideFirstBevelAngleValue;
                this.secondarySideGroove.Value = 0;
                this.secondarySideActualThroatThickness.Value = isRefSideMolded ? antiRefSideFirstBevelDepthValue : refSideFirstBevelDepthValue;
                this.secondarySideGrooveAngle.Value = isRefSideMolded ? antiRefSideFirstBevelAngleValue : refSideFirstBevelAngleValue;
                this.tailNotes.Value = string.Empty;
                this.tailNotesIsReference.Value = 0.0;
                this.primarySideActualLeglength.Value = 0.0;
                this.secondarySideActualLegLength.Value = 0.0;
                this.primarySideSymbol.Value = moldedFilletValue > 0 ? (int)WeldSymbol.Fillet : (int)WeldSymbol.None;
                this.secondarySideSymbol.Value = moldedFilletValue > 0 ? (int)WeldSymbol.Fillet : (int)WeldSymbol.None;
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
