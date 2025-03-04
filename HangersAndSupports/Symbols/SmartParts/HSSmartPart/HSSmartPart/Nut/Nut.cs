//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Nut.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Nut
//   Author       :  
//   Creation Date:  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   27.Dec.2012     BS     CR-CP-221376 Initial Creation   
//   25/Mar/2013     BS     DI-CP-228142  Modify the error handling for delivered H&S symbols                    
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Support.Exceptions;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Content.Support.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Nut : SmartPartComponentDefinition, ICustomWeightCG 
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Nut"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "RodDiameter", "Diameter of the Rod", 0)]
        public InputDouble m_Diameter;

        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Top", "Top")]
        [SymbolOutput("Bottom", "Bottom")]


        public AspectDefinition m_PhysicalAspect;

        #endregion


        #region "Definition of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddNutInputs(3, out endIndex, additionalInputs);
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
                AddNutOutputs(additionalOutputs);
            }
            return additionalOutputs;
        }
        #endregion
        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;

                double rodDiameter = m_Diameter.Value;
                NutInputs nut = LoadNutData(3);
                if (base.ToDoListMessage != null)
                {
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {                        
                        return;                        
                    }                    
                }

              
                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Bottom", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Bottom"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Top", new Position(0, 0, nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Top"] = port2;

                Matrix4X4 matrix = new Matrix4X4();
                matrix.Origin = new Position(0, 0, 0);

                AddNut(nut, matrix, m_PhysicalAspect.Outputs, "Nut");

            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Nut"));
                }
            }
        }
        #endregion
        
        #region "ICustomWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                ////System WCG Attributes

                Part CatalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = 0;
                }
                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = 0;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = 0;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of Nut"));
                }
            }
        }
        #endregion
    }
}
