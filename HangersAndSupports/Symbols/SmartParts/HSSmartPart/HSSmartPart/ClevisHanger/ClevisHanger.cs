//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ClevisHanger.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.ClevisHanger
//   Author       :  Manikanth
//   Creation Date:  15-02-2013
//   Description  :  CR CP-222479 Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   15-02-2013    Manikanth  CR CP-222479 Initial Creation
//   25/Mar/2013   Rajeswari  DI-CP-228142  Modify the error handling for delivered H&S symbols
//   05/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
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

    public class ClevisHanger : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.ClevisHanger"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;

        #endregion
        #region "Definition Of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                additionalInputs.Add(new InputDouble(2, "RodDiameter", "RodDiameter", 0, false));
                AddClevisHangerInputs(3, out endIndex, additionalInputs);
                additionalInputs.Add(new InputDouble(++endIndex, "PipeOD", "PipeOD", 0, false));
                return additionalInputs;

            }
        }
        #endregion
        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Route", "Route")]
        [SymbolOutput("RodEnd", "RodEnd")]
        [SymbolOutput("Bolt1", "Bolt1")]
        [SymbolOutput("Bolt2", "Bolt2")]
        public AspectDefinition m_PhysicalAspect;
        #endregion
        #region "Definition of Addiional Outputs"

        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOuputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddClevisHangerOutputs(additionalOuputs);
            }
            return additionalOuputs;
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
                int startIndex = 30;
                double rodDiameter = GetDoubleInputValue(2);
                ClevisHangerInputs cHanger = LoadClevisHangerData(3);
                if (base.ToDoListMessage != null)
                {
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        return;
                    }
                }
                double pipeOd = GetDoubleInputValue(startIndex);

                Matrix4X4 matrix = new Matrix4X4();              

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                AddClevisHanger(cHanger, pipeOd, rodDiameter, matrix, m_PhysicalAspect.Outputs, "ClevisHanger");

                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Route"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "RodEnd", new Position(0, 0, cHanger.RodTakeOut), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["RodEnd"] = port2;

                try
                {
                    Port port3 = new Port(OccurrenceConnection, part, "Bolt1", new Position(0, -cHanger.Pin1Length / 2, cHanger.Height2), new Vector(1, 0, 0), new Vector(0, -1, 0));
                    m_PhysicalAspect.Outputs["Bolt1"] = port3;

                    Port port4 = new Port(OccurrenceConnection, part, "Bolt2", new Position(0, cHanger.Pin1Length / 2, cHanger.Height2), new Vector(1, 0, 0), new Vector(0, 1, 0));
                    m_PhysicalAspect.Outputs["Bolt2"] = port4;

                }
                catch
                {
                    // the ports may not be defined in refdata.
                }

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of ClevisHanger"));
                    return;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in WeightCG of ClevisHanger"));
                    return;
                }
            }
        }
        #endregion



    }

}
