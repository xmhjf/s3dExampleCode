//---------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  TeeWeld1ParameterRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMPhysConnRules.dll
//  Original Class Name: ‘TeeWeld1Parm’ in VB content
//
//Abstract
//	TeeWeld1ParameterRule is a .NET parameter rule which is defining the custom parameter rule. 
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
    [RuleInterface(DetailingCustomAssembliesConstants.IASMPhysConnRules_TeeWeld1Parm, DetailingCustomAssembliesConstants.IASMPhysConnRules_TeeWeld1Parm)]
    [DefaultLocalizer(ConnectionsResourceIdentifiers.ConnectionsResource, ConnectionsResourceIdentifiers.ConnectionsAssembly)]
    public class TeeWeld1ParameterRule : PhysicalConnectionParameterRule
    {
        //==============================================================================================================
        //DefinitionName/ProgID of this symbol is "Connections,Ingr.SP3D.Content.Structure.TeeWeld1ParameterRule"
        //==============================================================================================================
        #region Parameters
        [ControlledParameter(1, DetailingCustomAssembliesConstants.Nose, DetailingCustomAssembliesConstants.Nose)]
        public ControlledParameterDouble nose;
        [ControlledParameter(2, DetailingCustomAssembliesConstants.NoseMethod, DetailingCustomAssembliesConstants.NoseMethod)]
        public ControlledParameterDouble noseMethod;
        [ControlledParameter(3, DetailingCustomAssembliesConstants.NoseOrientationAngle, DetailingCustomAssembliesConstants.NoseOrientationAngle)]
        public ControlledParameterDouble noseOrientationAngle;
        [ControlledParameter(4, DetailingCustomAssembliesConstants.MoldedFillet, DetailingCustomAssembliesConstants.MoldedFillet)]
        public ControlledParameterDouble moldedFillet;
        [ControlledParameter(5, DetailingCustomAssembliesConstants.AntiMoldedFillet, DetailingCustomAssembliesConstants.AntiMoldedFillet)]
        public ControlledParameterDouble antiMoldedFillet;
        [ControlledParameter(6, DetailingCustomAssembliesConstants.FilletMeasureMethod, DetailingCustomAssembliesConstants.FilletMeasureMethod)]
        public ControlledParameterDouble filletMeasuredMethod;
        [ControlledParameter(7, DetailingCustomAssembliesConstants.Category, DetailingCustomAssembliesConstants.Category)]
        public ControlledParameterDouble category;
        [ControlledParameter(8, DetailingCustomAssembliesConstants.ReferenceSide, DetailingCustomAssembliesConstants.ReferenceSide)]
        public ControlledParameterString referenceSide;
        [ControlledParameter(9, DetailingCustomAssembliesConstants.RefPartName, DetailingCustomAssembliesConstants.RefPartName)]
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
                if (physicalConnection.PartName.Equals(DetailingCustomAssembliesConstants.TeeWeld1))
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
                //calculate NoseOrientationAngle, depending on the bevel method
                //Get BevelMethod answer
                int bevelMethod = (int)((PropertyValueCodelist)physicalConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.BevelAngleMethod)).PropValue;  
                switch (bevelMethod)
                {
                    case (int)BevelMethod.Constant:
                        this.noseOrientationAngle.Value = Math.PI / 2;
                        this.noseMethod.Value = bevelMethod;
                        break;
                    case (int)BevelMethod.Varying:
                        this.noseOrientationAngle.Value = Math.Abs((Math.PI / 2) - (physicalConnection.GetConnectionAngle() - StructHelper.DISTTOL));
                        this.noseMethod.Value = bevelMethod;
                        break;
                }

                this.nose.Value = boundedPartThickness;
                //Get the part names               
                this.refPartName.Value = physicalConnection.BoundedObject.ToString();
                //get the reference side
                string referenceSideName = ConnectionServices.GetReferenceSide(physicalConnection, base.BoundedObject);
                this.referenceSide.Value = ConnectionServices.GetMoldedType(referenceSideName);
                this.antiMoldedFillet.Value = boundedPartThickness > boundingPartThickness ? boundingPartThickness * 0.82 : boundedPartThickness * 0.82;
                this.moldedFillet.Value = boundedPartThickness > boundingPartThickness ? boundingPartThickness * 0.82 : boundedPartThickness * 0.82;
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
