//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   LRParts_FIG253N.cs
//    LRParts,Ingr.SP3D.Content.Support.Symbols.LRParts_FIG253N
//   Author       :  Rajeswari
//   Creation Date:  25/10/2012  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  25/10/2012     Rajeswari Initial Creation
//  26/03/2013     Rajeswari DI-CP-228142  Modify the error handling for delivered H&S symbols
//  30/10/2013     Vijaya    CR-CP-242533  Provide the ability to store GType outputs as a single blob.
//  21/05/2015     PVK       TR-CP-243831	Error in WeightCG when some LRParts replaced with any LRParts_FIG253N parts
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
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
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class LRParts_FIG253N : CustomSymbolDefinition, ICustomWeightCG 
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "LRParts,Ingr.SP3D.Content.Support.Symbols.LRParts_FIG253N"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Length", "Length", 0.999999)]
        public InputDouble m_Length;
        [InputDouble(3, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_ROD_DIA;
        [InputDouble(4, "WEIGHT_PER_LENGTH", "WEIGHT_PER_LENGTH", 0.999999)]
        public InputDouble m_WEIGHT_PER_LENGTH;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("ROD", "ROD")]
        public AspectDefinition m_Symbolic;

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
                Double length = m_Length.Value;
                Double rodDiameter = m_ROD_DIA.Value;

                if (length == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidLength, "Length cannot be zero"));
                    return;
                }
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrInvalidRodDiameter, "Rod Diameter cannot be zero or negative"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));

                Port port1 = new Port(OccurrenceConnection, part, "ExThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "ExThdLH", new Position(0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                Projection3d rod = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2, length);
                m_Symbolic.Outputs["ROD"] = rod;
            }
            catch  //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of LRParts_FIG253N.cs."));
                return;
            }
        }
        #endregion
        
        #region "ICustomWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                double length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                double weightPerUnitlength = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrLRParts_FIG253N", "WEIGHT_PER_LENGTH")).PropValue;
                Double weight, cogX, cogY, cogZ;
                weight = weightPerUnitlength * length;
                cogX = 0;
                cogY = 0;
                cogZ = -length / 2;
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, LRPartsLocalizer.GetString(LRPartsResourceIDs.ErrWeightCG, "Error in WeightCG of LRParts_FIG253N.cs."));
            }
        }
        #endregion
    }
}
