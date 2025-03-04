//-------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  TeeWeld7ParameterRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMPhysConnRules.dll
//  Original Class Name: ‘TeeWeld7Parm’ in VB content
//
//Abstract
//	TeeWeld7ParameterRule is a .NET parameter rule which is defining the custom parameter rule. 
//  This class subclasses from PhysicalConnectionParameterRule. 
// Change History:
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
    [RuleInterface(DetailingCustomAssembliesConstants.IASMPhysConnRules_TeeWeld7Parm, DetailingCustomAssembliesConstants.IASMPhysConnRules_TeeWeld7Parm)]
    [DefaultLocalizer(ConnectionsResourceIdentifiers.ConnectionsResource, ConnectionsResourceIdentifiers.ConnectionsAssembly)]
    public class TeeWeld7ParameterRule : PhysicalConnectionParameterRule
    {
        //==============================================================================================================
        //DefinitionName/ProgID of this symbol is "Connections,Ingr.SP3D.Content.Structure.TeeWeld7ParameterRule"
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
                if (physicalConnection.PartName.Equals(DetailingCustomAssembliesConstants.TeeWeld7))
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
                this.refPartName.Value = physicalConnection.BoundedObject.ToString();
                double mountingAngle = physicalConnection.GetConnectionAngle();
                double smallerMountingAngle = (mountingAngle >= Math.PI / 2) ? Math.PI - mountingAngle : mountingAngle;
                this.nose.Value = Math3d.FitTolerance * 3 * Math.Cos((Math.PI / 2) - smallerMountingAngle);
                //Get BevelMethod answer
                int bevelMethod = (int)((PropertyValueCodelist)physicalConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.BevelAngleMethod)).PropValue;
                switch (bevelMethod)
                {
                    case (int)BevelMethod.Constant:
                        {

                            this.refSideFirstBevelDepth.Value = (((Math.PI / 2) - mountingAngle) > 0) ? 0 : (boundedPartThickness - Math3d.FitTolerance * 3) * Math.Abs(Math.Cos((Math.PI / 2) - mountingAngle));
                            this.refSideFirstBevelAngle.Value = (((Math.PI / 2) - mountingAngle) > 0) ? 0 : Math.Abs(0.6981317 + Math.Abs((Math.PI / 2) - mountingAngle));
                            this.antiRefSideFirstBevelDepth.Value = (((Math.PI / 2) - mountingAngle) > 0) ? (boundedPartThickness - (Math3d.FitTolerance * 3)) * Math.Abs(Math.Cos((Math.PI / 2) - mountingAngle)) : 0;
                            this.antiRefSideFirstBevelAngle.Value = (((Math.PI / 2) - mountingAngle) > 0) ? Math.Abs(0.6981317 - ((Math.PI / 2) - mountingAngle)) : 0;
                            this.noseOrientationAngle.Value = Math.PI - mountingAngle;
                            this.refSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Leg;
                            this.antiRefSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Leg;
                            this.noseMethod.Value = bevelMethod;
                        }
                        break;
                    case (int)BevelMethod.Varying:
                        {

                            this.refSideFirstBevelDepth.Value = ((90 - mountingAngle) > 0) ? 0 : (boundedPartThickness - (Math3d.FitTolerance * 3)) * Math.Abs(Math.Cos((Math.PI / 2) - mountingAngle));
                            this.refSideFirstBevelAngle.Value = 0;
                            this.antiRefSideFirstBevelDepth.Value = ((90 - mountingAngle) > 0) ? (boundedPartThickness - (Math3d.FitTolerance * 3)) * Math.Abs(Math.Cos((Math.PI / 2) - mountingAngle)) : 0;
                            this.antiRefSideFirstBevelAngle.Value = ((90 - mountingAngle) > 0) ? 0.6981317 : 0;
                            this.noseOrientationAngle.Value = 0;
                            this.refSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Throat;
                            this.antiRefSideFirstBevelMethod.Value = (int)WeldFilletMeasure.Throat;
                            this.noseMethod.Value = bevelMethod;
                        }
                        break;
                }

                double fillet = boundedPartThickness <= boundingPartThickness ? boundedPartThickness / 4 : boundingPartThickness / 4;
                if (fillet < 0.00317)
                {
                    fillet = 0.00317;
                }
                else if (fillet > 0.00952)
                {
                    fillet = 0.00952;
                }
                string referenceSideName = ConnectionServices.GetReferenceSide(physicalConnection, base.BoundedObject);
                this.referenceSide.Value = ConnectionServices.GetMoldedType(referenceSideName);
                this.moldedFillet.Value = fillet;
                this.antiMoldedFillet.Value = fillet;
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
