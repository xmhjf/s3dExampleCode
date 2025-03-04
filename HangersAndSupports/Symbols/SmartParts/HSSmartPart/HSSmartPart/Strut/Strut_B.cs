//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Strut_B.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Strut_B
//   Author       :  Manikanth
//   Creation Date:  24/01/2013
//   Description:    CR-CP-222466 Creating StrutB Class

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   24/01/2013   Manikanth  CR-CP-222466 Creating StrutB Class 
//   25/Mar/2013  Sridhar    DI-CP-228142  Modify the error handling for delivered H&S symbols
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using System.Collections.Generic;

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
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Strut_B : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Strut_B"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Definition of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddStrutBInputs(2, out endIndex, additionalInputs);
                return additionalInputs;
            }
        }
        #endregion

        #region "Definition of Additional Outputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddStrutBOutputs(additionalOutputs);
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
                StrutBInputs strutB = LoadStrutBData(2);
                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;               
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                if (strutB.rodEndType == 1 || strutB.rodEndType == 2 || strutB.rodEndType == 3 || strutB.rodEndType == 4)
                {
                    Port port1 = new Port(OccurrenceConnection, part, "Port1", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_PhysicalAspect.Outputs["Port1"] = port1;
                }
                else if (strutB.rodEndType == 5)
                {
                    Port port1 = new Port(OccurrenceConnection, part, "Port1", new Position(0, 0, 0), new Vector(-1, 0, 0), new Vector(0, 0, -1));
                    m_PhysicalAspect.Outputs["Port1"] = port1;
                }

                Port port2 = new Port(OccurrenceConnection, part, "Port2", new Position(0, 0, strutB.length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port2"] = port2;

                Matrix4X4 matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, 0, 0));
                AddStrutB(strutB, matrix, m_PhysicalAspect.Outputs, "Strut");
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Strut_B"));
                }
            }
        }
        #endregion


        #region "ICustomHgrWeightCG Members"

        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                ////System WCG Attributes

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
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in WeightCG of Strut_B"));
                    return;
                }
            }
        }
        #endregion



    }

}
