//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Clevis.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Clevis
//   Author       :  Pradeep
//   Creation Date:  23-01-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   23-01-2013     Pradeep   CR-222477- Initial Creation.
//   25/Mar/2013    Sridhar   DI-CP-228142  Modify the error handling for delivered H&S symbols
//   05/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//   13/Aug/2015     PR       TR 276067	Few smart parts are having cache option as non-cached when they should be cached
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;


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
    [SymbolVersion("1.0.0.1")]
    [VariableOutputs]
    public class Clevis : SmartPartComponentDefinition, ICustomWeightCG
    {
        //SmartPartHelper smarthelper = new SmartPartHelper();
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Clevis"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Pin", "Pin")]
        [SymbolOutput("RodEnd", "RodEnd")]
        [SymbolOutput("Surface", "Surface")]
        public AspectDefinition m_oSymbolic;


        #region "Definition of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddRodInputs(2, out endIndex, additionalInputs);
                AddClevisInputs(4, out endIndex, additionalInputs);
                return additionalInputs;
            }
        }
        #endregion
        #region "Definition of Additional Outputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            AddClevisOutputs(aspectName, additionalOutputs);

            return additionalOutputs;
        }
        #endregion
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

                RodInputs rod = LoadRodData(2);
                ClevisInputs clevis = LoadClevisData(4);
                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;
                try
                {
                    Port port1 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_oSymbolic.Outputs["Pin"] = port1;
                }
                catch (Ingr.SP3D.Support.Exceptions.PortCreationException)
                {
                    Port port1 = new Port(OccurrenceConnection, part, "Hole", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                    m_oSymbolic.Outputs["Hole"] = port1;
                }
                Port port2 = new Port(OccurrenceConnection, part, "RodEnd", new Position(0, 0, rod.RodTakeOut), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oSymbolic.Outputs["RodEnd"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "Surface", new Position(0, 0, ((clevis.Opening1) + (clevis.nut.ShapeLength))), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oSymbolic.Outputs["Surface"] = port3;
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                Matrix4X4 matrix = new Matrix4X4();
                matrix.Origin = new Position(0, 0, 0);
              //Clevis 
                AddClevis(clevis, matrix, m_oSymbolic.Outputs, "Clevis");

            }

            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Clevis"));
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of Clevis"));
                }
            }
        }
        #endregion

    }

}
