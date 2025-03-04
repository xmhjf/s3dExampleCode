//-------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  TeeWeld5ParameterRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMPhysConnRules.dll
//  Original Class Name: ‘TeeWeld5Parm’ in VB content
//
//Abstract
//	TeeWeld5ParameterRule is a .NET parameter rule which is defining the custom parameter rule. 
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
    [RuleInterface(DetailingCustomAssembliesConstants.IASMPhysConnRules_TeeWeld5Parm, DetailingCustomAssembliesConstants.IASMPhysConnRules_TeeWeld5Parm)]
    [DefaultLocalizer(ConnectionsResourceIdentifiers.ConnectionsResource, ConnectionsResourceIdentifiers.ConnectionsAssembly)]
    public class TeeWeld5ParameterRule : PhysicalConnectionParameterRule
    {
        //==============================================================================================================
        //DefinitionName/ProgID of this symbol is "Connections,Ingr.SP3D.Content.Structure.TeeWeld5ParameterRule"
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
                if (physicalConnection.PartName.Equals(DetailingCustomAssembliesConstants.TeeWeld5))
                {
                    try
                    {
                        ConnectionServices.GetPhysicalConnectionPartsThickness(physicalConnection, out boundedPartThickness, out boundingPartThickness);
                    }
                    catch (Exception ex)
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
                this.nose.Value = 0;
                this.noseOrientationAngle.Value = 0;
                this.antiRefSideFirstBevelDepth.Value = boundedPartThickness / 2;
                this.refSideFirstBevelDepth.Value = boundedPartThickness / 2;
                //Get the part names
                this.refPartName.Value = physicalConnection.BoundedObject.ToString();
                //calculate angles, depending on the bevel method
                //Get BevelMethod answer
                int bevelMethod = (int)((PropertyValueCodelist)physicalConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.BevelAngleMethod)).PropValue;
                switch (bevelMethod)
                {
                    case (int)BevelMethod.Constant:
                        double mountingAngle = physicalConnection.GetConnectionAngle();
                        this.refSideFirstBevelAngle.Value = Math.Abs(0.6981317 + Math.Abs((Math.PI / 2) - mountingAngle));
                        this.antiRefSideFirstBevelAngle.Value = Math.Abs(0.6981317 - Math.Abs((Math.PI / 2) - mountingAngle));
                        this.refSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Leg;
                        this.antiRefSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Leg;
                        this.noseMethod.Value = bevelMethod;
                        break;
                    case (int)BevelMethod.Varying:
                        this.refSideFirstBevelAngle.Value = 0.6981317;
                        this.antiRefSideFirstBevelAngle.Value = 0.6981317;
                        this.refSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Throat;
                        this.antiRefSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Throat;
                        this.noseMethod.Value = bevelMethod;
                        break;
                }

                //get the reference side
                string referenceSideName = ConnectionServices.GetReferenceSide(physicalConnection, base.BoundedObject);
                this.referenceSide.Value = ConnectionServices.GetMoldedType(referenceSideName);
                this.moldedFillet.Value = this.antiMoldedFillet.Value = boundedPartThickness / 2 < 0.00952 ? boundedPartThickness / 2 : 0.00952;
                this.filletMeasuredMethod.Value = (int)WeldFilletMeasure.Leg;
                //Get Category answer
                this.category.Value = (int)((PropertyValueCodelist)physicalConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.Category)).PropValue;
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
