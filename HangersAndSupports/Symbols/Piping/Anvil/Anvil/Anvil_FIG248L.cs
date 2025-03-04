//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG248L.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG248L
//   Author       : Vijaya 
//   Creation Date: 30-April-2013 
//   Description: Initial Creation-CR-CP-222292
//   
//   Anvil_FIG248L.cs is same for Anvil_FIG278L.cs

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30-April-2013  Vijaya   CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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

    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]
    public class Anvil_FIG248L : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG248L"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Length", "Length", 0.999999)]
        public InputDouble m_dLength;
        [InputDouble(3, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_dROD_DIA;
        [InputDouble(4, "WEIGHT_PER_LENGTH", "WEIGHT_PER_LENGTH", 0.999999)]
        public InputDouble m_dWEIGHT_PER_LENGTH;
        [InputDouble(5, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        [InputDouble(6, "D_USER", "D_USER", 0.999999)]
        public InputDouble m_dD_USER;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("ROD", "ROD")]
        [SymbolOutput("EYE", "EYE")]
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
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                Double length = m_dLength.Value, rodDiameter = m_dROD_DIA.Value;
                double holeClearance = 0.003175;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Eye", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "ExThdLH", new Position(0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidDiameterGTZero, "Diameter should be greater than zero"));
                    return;
                }
                if (rodDiameter > 0.036)
                    holeClearance = 0.00635;

                Vector normal = new Position(0, 0, length).Subtract(new Position(0, 0, holeClearance + rodDiameter));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, holeClearance + rodDiameter);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d rod = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2.0, normal.Length);
                m_Symbolic.Outputs["ROD"] = rod;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, rodDiameter / 2, (holeClearance + rodDiameter) / 2).Subtract(new Position(0, -rodDiameter / 2, (holeClearance + rodDiameter) / 2));
                symbolGeometryHelper.ActivePosition = new Position(0, -rodDiameter / 2, (holeClearance + rodDiameter) / 2);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d eye = symbolGeometryHelper.CreateCylinder(null, (rodDiameter + holeClearance) / 2 + rodDiameter, normal.Length);
                m_Symbolic.Outputs["EYE"] = eye;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG248L"));
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
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double weight, cogX, cogY, cogZ, weightPerUnitLength=0;
                double length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                if(supportComponentBO.SupportsInterface("IJOAHgrAnvil_FIG248L"))
                    weightPerUnitLength = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJOAHgrAnvil_FIG248L", "WEIGHT_PER_LENGTH")).PropValue;
                else if(supportComponentBO.SupportsInterface("IJOAHgrAnvil_FIG278L"))
                    weightPerUnitLength = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJOAHgrAnvil_FIG278L", "WEIGHT_PER_LENGTH")).PropValue;

                weight = weightPerUnitLength * length;
                cogX = 0;
                cogY = 0;
                cogZ = -length / 2;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrWeightCG, "Error while defining weightCG of Anvil_FIG248L"));
            }
        }
        #endregion
    }

}
