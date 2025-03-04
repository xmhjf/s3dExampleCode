//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ElbowLug.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.ElbowLug
//   Author       :  Vijaya
//   Creation Date:  18-Feb-2013
//   Description:  CR-CP-222482 Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   18-Feb-2013    Vijaya  CR-CP-222482 Initial Creation
//   25/Mar/2013    Vijaya   DI-CP-228142  Modify the error handling for delivered H&S symbols
//   12-12-2014     PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]
    [VariableOutputs]
    public class ElbowLug : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.ElbowLug"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "RodDiameter", "Diameter of the Rod", 0)]
        public InputDouble m_RodDiameter;
        [InputDouble(3, "PipeOD", "Diameter of the Pipe", 0)]
        public InputDouble m_PipeDiameter;

        #endregion

        #region "Definition of Additional Inputs"

        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddElbowLugInputs(4, out endIndex, additionalInputs);
                return additionalInputs;
            }
        }

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        //Ports as Outputs
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("Lug", "Lug")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Definition of Additional Outputs"

        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddElbowLugOutputs(additionalOutputs);
            }
            return additionalOutputs;
        }
        #endregion
        #region "Construct Outputs"
        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                Double rodDiameter = m_RodDiameter.Value;
                Double pipeDiameter = m_PipeDiameter.Value;

                ElbowLugInputs elbowLug = LoadElbowLugData(4);
                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                Matrix4X4 matrix = new Matrix4X4();

                //Check to Face to Center. Assumed the Elbow is Short Radius
                if (HgrCompareDoubleService.cmpdbl(elbowLug.FaceToCenter, 0)==true)
                    elbowLug.FaceToCenter = elbowLug.ElbowRadius;

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Route"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Lug", new Position(0, 0, elbowLug.FaceToCenter + elbowLug.RodTakeOut), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Lug"] = port2;


                //Add Elbow geometry
                if (HgrCompareDoubleService.cmpdbl(elbowLug.Gap1, 0)==true) //For Single Lug
                    AddElbowLug(elbowLug, pipeDiameter, matrix, m_PhysicalAspect.Outputs, "ElbowLugShape1");
                else  //For Double Lug
                {
                    matrix.Translate(new Vector(0, elbowLug.Gap1 / 2 + elbowLug.Thickness1 / 2, 0));
                    AddElbowLug(elbowLug, pipeDiameter, matrix, m_PhysicalAspect.Outputs, "ElbowLugShape1");
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(0, -elbowLug.Gap1 / 2 - elbowLug.Thickness1 / 2, 0));
                    AddElbowLug(elbowLug, pipeDiameter, matrix, m_PhysicalAspect.Outputs, "ElbowLugShape2");

                    //Add Stiffener
                    if (HgrCompareDoubleService.cmpdbl(elbowLug.StiffenerHeight, 0)==false && HgrCompareDoubleService.cmpdbl(elbowLug.StiffenerLength, 0)==false)
                    {
                        symbolGeometryHelper.ActivePosition = new Position(-elbowLug.StiffenerLength / 2, -elbowLug.Gap1 / 2, elbowLug.FaceToCenter + elbowLug.RodTakeOut - elbowLug.StiffenerHeight - elbowLug.StiffenerOffset);
                        BusinessObject stiffener = symbolGeometryHelper.CreateBox(null, elbowLug.StiffenerLength, elbowLug.Gap1, elbowLug.StiffenerHeight, 9);
                        m_PhysicalAspect.Outputs["ElbowLugStiffener"] = stiffener;
                    }
                    //Add Pin
                    if (HgrCompareDoubleService.cmpdbl(elbowLug.Pin1Diameter, 0)==false || HgrCompareDoubleService.cmpdbl(elbowLug.Pin1Length, 0)==false)
                    {
                        symbolGeometryHelper = new SymbolGeometryHelper();
                        symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                        symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Rotate((Math.PI / 2), new Vector(1, 0, 0));
                        Projection3d elbowLugPin = new Projection3d(symbolGeometryHelper.CreateCylinder(null, elbowLug.Pin1Diameter / 2, elbowLug.Pin1Length));
                        elbowLugPin.Transform(matrix);
                        matrix = new Matrix4X4();
                        matrix.SetIdentity();
                        matrix.Translate(new Vector(0, elbowLug.Pin1Length / 2, elbowLug.FaceToCenter + elbowLug.RodTakeOut));
                        elbowLugPin.Transform(matrix);
                        m_PhysicalAspect.Outputs["ElbowLugPin"] = elbowLugPin;
                    }
                }

            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of ElbowLug"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                //System WCG Attributes
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = 0;
                }
                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = 0;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = 0;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
                {
                    cogZ = 0;
                }

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of ElbowLug"));
                    return;
                }
            }
        }
        #endregion


    }
}