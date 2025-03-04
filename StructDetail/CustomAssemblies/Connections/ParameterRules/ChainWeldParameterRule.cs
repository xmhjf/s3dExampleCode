//---------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  ChainWeldParameterRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMPhysConnRules.dll
//  Original Class Name: ‘ChainWeldParm’ in VB content
//
//Abstract
//	ChainWeldParameterRule is a .NET parameter rule which is defining the custom parameter rule. 
//  This class subclasses from PhysicalConnectionParameterRule. 
//--------------------------------------------------------------------------------------------------------

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
    [RuleInterface(DetailingCustomAssembliesConstants.IASMPhysConnRules_ChainWeldParm, DetailingCustomAssembliesConstants.IASMPhysConnRules_ChainWeldParm)]
    [DefaultLocalizer(ConnectionsResourceIdentifiers.ConnectionsResource, ConnectionsResourceIdentifiers.ConnectionsAssembly)]
    public class ChainWeldParameterRule : PhysicalConnectionParameterRule
    {
        //==============================================================================================================
        //DefinitionName/ProgID of this symbol is "Connections,Ingr.SP3D.Content.Structure.ChainWeldParameterRule"
        //==============================================================================================================
        #region Parameters
        [ControlledParameter(1, DetailingCustomAssembliesConstants.Nose, DetailingCustomAssembliesConstants.Nose)]
        public ControlledParameterDouble nose;
        [ControlledParameter(2, DetailingCustomAssembliesConstants.NoseOrientationAngle, DetailingCustomAssembliesConstants.NoseOrientationAngle)]
        public ControlledParameterDouble noseOrientationAngle;
        [ControlledParameter(3, DetailingCustomAssembliesConstants.MoldedFillet, DetailingCustomAssembliesConstants.MoldedFillet)]
        public ControlledParameterDouble moldedFillet;
        [ControlledParameter(4, DetailingCustomAssembliesConstants.AntiMoldedFillet, DetailingCustomAssembliesConstants.AntiMoldedFillet)]
        public ControlledParameterDouble antiMoldedFillet;
        [ControlledParameter(5, DetailingCustomAssembliesConstants.FilletMeasureMethod, DetailingCustomAssembliesConstants.FilletMeasureMethod)]
        public ControlledParameterDouble filletMeasuredMethod;
        [ControlledParameter(6, DetailingCustomAssembliesConstants.Category, DetailingCustomAssembliesConstants.Category)]
        public ControlledParameterDouble category;
        [ControlledParameter(7, DetailingCustomAssembliesConstants.Pitch, DetailingCustomAssembliesConstants.Pitch)]
        public ControlledParameterDouble pitch;
        [ControlledParameter(8, DetailingCustomAssembliesConstants.Length, DetailingCustomAssembliesConstants.Length)]
        public ControlledParameterDouble length;
        [ControlledParameter(9, DetailingCustomAssembliesConstants.ReferenceSide, DetailingCustomAssembliesConstants.ReferenceSide)]
        public ControlledParameterString referenceSide;
        [ControlledParameter(10, DetailingCustomAssembliesConstants.RefPartName, DetailingCustomAssembliesConstants.RefPartName)]
        public ControlledParameterString refPartName;
        [ControlledParameter(11, DetailingCustomAssembliesConstants.PrimarySideGroove, DetailingCustomAssembliesConstants.PrimarySideGroove)]
        public ControlledParameterDouble primarySideGroove;
        [ControlledParameter(12, DetailingCustomAssembliesConstants.SecondarySideGroove, DetailingCustomAssembliesConstants.SecondarySideGroove)]
        public ControlledParameterDouble secondarySideGroove;
        [ControlledParameter(13, DetailingCustomAssembliesConstants.PrimarySideGrooveSize, DetailingCustomAssembliesConstants.PrimarySideGrooveSize)]
        public ControlledParameterDouble primarySideGrooveSize;
        [ControlledParameter(14, DetailingCustomAssembliesConstants.SecondarySideGrooveSize, DetailingCustomAssembliesConstants.SecondarySideGrooveSize)]
        public ControlledParameterDouble secondarySideGrooveSize;
        [ControlledParameter(15, DetailingCustomAssembliesConstants.PrimarySideActualThroatThickness, DetailingCustomAssembliesConstants.PrimarySideActualThroatThickness)]
        public ControlledParameterDouble primarySideActualThroatThickness;
        [ControlledParameter(16, DetailingCustomAssembliesConstants.SecondarySideActualThroatThickness, DetailingCustomAssembliesConstants.SecondarySideActualThroatThickness)]
        public ControlledParameterDouble secondarySideActualThroatThickness;
        [ControlledParameter(17, DetailingCustomAssembliesConstants.PrimarySideNominalThroatThickness, DetailingCustomAssembliesConstants.PrimarySideNominalThroatThickness)]
        public ControlledParameterDouble primarySideNominalThroatThickness;
        [ControlledParameter(18, DetailingCustomAssembliesConstants.SecondarySideNominalThroatThickness, DetailingCustomAssembliesConstants.SecondarySideNominalThroatThickness)]
        public ControlledParameterDouble secondarySideNominalThroatThickness;
        [ControlledParameter(19, DetailingCustomAssembliesConstants.PrimarySideLength, DetailingCustomAssembliesConstants.PrimarySideLength)]
        public ControlledParameterDouble primarySideLength;
        [ControlledParameter(20, DetailingCustomAssembliesConstants.SecondarySideLength, DetailingCustomAssembliesConstants.SecondarySideLength)]
        public ControlledParameterDouble secondarySideLength;
        [ControlledParameter(21, DetailingCustomAssembliesConstants.PrimarySidePitch, DetailingCustomAssembliesConstants.PrimarySidePitch)]
        public ControlledParameterDouble primarySidePitch;
        [ControlledParameter(22, DetailingCustomAssembliesConstants.SecondarySidePitch, DetailingCustomAssembliesConstants.SecondarySidePitch)]
        public ControlledParameterDouble secondarySidePitch;
        [ControlledParameter(23, DetailingCustomAssembliesConstants.PrimarySideRootOpening, DetailingCustomAssembliesConstants.PrimarySideRootOpening)]
        public ControlledParameterDouble primarySideRootOpening;
        [ControlledParameter(24, DetailingCustomAssembliesConstants.PrimarySideGrooveAngle, DetailingCustomAssembliesConstants.PrimarySideGrooveAngle)]
        public ControlledParameterDouble primarySideGrooveAngle;
        [ControlledParameter(25, DetailingCustomAssembliesConstants.SecondarySideGrooveAngle, DetailingCustomAssembliesConstants.SecondarySideGrooveAngle)]
        public ControlledParameterDouble secondarySideGrooveAngle;
        [ControlledParameter(26, DetailingCustomAssembliesConstants.TailNotes, DetailingCustomAssembliesConstants.TailNotes)]
        public ControlledParameterString tailNotes;
        [ControlledParameter(27, DetailingCustomAssembliesConstants.TailNoteIsReference, DetailingCustomAssembliesConstants.TailNoteIsReference)]
        public ControlledParameterDouble tailNotesIsReference;
        [ControlledParameter(28, DetailingCustomAssembliesConstants.PrimarySideActualLegLength, DetailingCustomAssembliesConstants.PrimarySideActualLegLength)]
        public ControlledParameterDouble primarySideActualLeglength;
        [ControlledParameter(29, DetailingCustomAssembliesConstants.SecondarySideActualLegLength, DetailingCustomAssembliesConstants.SecondarySideActualLegLength)]
        public ControlledParameterDouble secondarySideActualLegLength;
        [ControlledParameter(30, DetailingCustomAssembliesConstants.PrimarySideSymbol, DetailingCustomAssembliesConstants.PrimarySideSymbol)]
        public ControlledParameterDouble primarySideSymbol;
        [ControlledParameter(31, DetailingCustomAssembliesConstants.SecondarySideSymbol, DetailingCustomAssembliesConstants.SecondarySideSymbol)]
        public ControlledParameterDouble secondarySideSymbol;
        #endregion Parameters

        #region Public override properties and methods
        /// <summary>
        /// Evaluates the parameter rule.
        /// It calculates the item parameters in the context of the Physical Connection.
        /// </summary>
        public override void Evaluate()
        {
            //*****************************************************
            //For SM Implementation
            //This is the Delivered ChainWeld, which has a pitch of 300 mm and length of
            //75 mm by default.  No requirements were provided for P and L.
            //The fillets are (0.2 * thickness)
            //*******************************************************
            try
            {
                //Get Class Arguments
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
                if (physicalConnection.PartName.Equals(DetailingCustomAssembliesConstants.ChainWeld))
                {
                    try
                    {
                        ConnectionServices.GetPhysicalConnectionPartsThickness(physicalConnection, out boundedPartThickness, out boundingPartThickness);
                    }
                    catch (CmnException ex)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning,
                        base.GetString(ConnectionsResourceIdentifiers.ErrEndCutType, "Unable to get endcut thickness") + " " + ex.Message + this.GetType().Name);
                    }
                }
                else
                {
                    boundedPartThickness = Convert.ToDouble(((PropertyValue)physicalConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.ChamferThickness)).ToString());
                    boundingPartThickness = ConnectionServices.GetBoundingPartThickness(base.BoundingObject);
                }

                this.nose.Value = boundedPartThickness;
                this.noseOrientationAngle.Value = Math.PI / 2;
                //Calculate Pitch and Length based on Class Society
                int classSociety = ((PropertyValueCodelist)physicalConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.ClassSociety)).PropValue;
                double pitchValue = 0.0, lengthValue = 0.0;
                ConnectionServices.GetPitchAndLengthValues(classSociety, out pitchValue, out lengthValue);
                //store the values, for now.  May be overridden by a user defined value later
                this.pitch.Value = pitchValue;
                this.length.Value = lengthValue;
                //get the reference side               
                string referenceSideName = ConnectionServices.GetReferenceSide(physicalConnection, base.BoundedObject);
                this.referenceSide.Value = ConnectionServices.GetMoldedType(referenceSideName);
                double moldedFilletValue = boundedPartThickness * 0.2;
                double antiMoldedFilletValue = moldedFilletValue;
                this.moldedFillet.Value = moldedFilletValue;
                this.antiMoldedFillet.Value = antiMoldedFilletValue;
                this.refPartName.Value = physicalConnection.BoundedObject.ToString();
                this.category.Value = (int)TeeWeldCategory.Chain;
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
                if (physicalConnection.IsParameterOverridden(DetailingCustomAssembliesConstants.Pitch))
                {
                    pitchValue = (double)((PropertyValueDouble)physicalConnection.GetParameterValue(DetailingCustomAssembliesConstants.Pitch)).PropValue;
                }
                if (physicalConnection.IsParameterOverridden(DetailingCustomAssembliesConstants.Length))
                {
                    lengthValue = (double)((PropertyValueDouble)physicalConnection.GetParameterValue(DetailingCustomAssembliesConstants.Length)).PropValue;
                }

                this.filletMeasuredMethod.Value = (int)WeldFilletMeasure.Leg;
                this.primarySideGroove.Value = (int)WeldGroove.None;
                this.secondarySideGroove.Value = (int)WeldGroove.None;
                this.primarySideGrooveSize.Value = 0.0;
                this.secondarySideGrooveSize.Value = 0.0;
                this.primarySideActualThroatThickness.Value = 0.0;
                this.secondarySideActualThroatThickness.Value = 0.0;
                this.primarySideRootOpening.Value = 0.0;
                this.primarySideGrooveAngle.Value = 0.0;
                this.secondarySideGrooveAngle.Value = 0.0;
                this.tailNotesIsReference.Value = 0.0;
                this.primarySideActualLeglength.Value = 0.0;
                this.secondarySideActualLegLength.Value = 0.0;
                this.primarySideLength.Value = lengthValue;
                this.secondarySideLength.Value = lengthValue;
                this.primarySidePitch.Value = pitchValue;
                this.tailNotes.Value = string.Empty;
                this.secondarySidePitch.Value = pitchValue;
                this.primarySideNominalThroatThickness.Value = moldedFilletValue;
                this.secondarySideNominalThroatThickness.Value = antiMoldedFilletValue;
                this.primarySideSymbol.Value = moldedFilletValue > 0.0 ? (int)WeldSymbol.Fillet : (int)WeldSymbol.None;
                this.secondarySideSymbol.Value = moldedFilletValue > 0.0 ? (int)WeldSymbol.Fillet : (int)WeldSymbol.None;
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
