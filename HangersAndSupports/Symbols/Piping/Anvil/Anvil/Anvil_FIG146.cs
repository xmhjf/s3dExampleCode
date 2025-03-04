//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG146.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG146
//   Author       :  Vijay
//   Creation Date:  30-04-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30-04-2013     Vijay    CR-CP-222292 Convert HS_Anvil VB Project to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
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
    [SymbolVersion("1.0.0.0")]
    public class Anvil_FIG146 : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG146"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_dROD_DIA;
        [InputDouble(3, "Length", "Length", 0.999999)]
        public InputDouble m_dLength;
        [InputDouble(4, "WEIGHT_PER_LENGTH", "WEIGHT_PER_LENGTH", 0.999999)]
        public InputDouble m_dWEIGHT_PER_LENGTH;
        [InputDouble(5, "FINISH", "FINISH",1)]
        public InputDouble m_oFINISH;
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
              
                Double rodDiameter = m_dROD_DIA.Value;
                Double length = m_dLength.Value;
                Double weigthPerLength = m_dWEIGHT_PER_LENGTH.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "BotExThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "TopExThdRH", new Position(0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (length == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidLengthNZero, "Length  cannot be zero"));
                    return;
                }

                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidRodDiaGTZero, "Rod diameter  should be greater than zero"));
                    return;
                }

                Matrix4X4 matrix = new Matrix4X4();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Translate(new Vector(0, 0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d rodCylinder = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2, length);
                rodCylinder.Transform(matrix);
                m_Symbolic.Outputs["ROD"] = rodCylinder;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG146"));
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
                double weight, cogX, cogY, cogZ;
                double length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccLength", "Length")).PropValue;
                double weightPerUnitLength = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrAnvil_fig146", "WEIGHT_PER_LENGTH")).PropValue;

                weight = weightPerUnitLength * length;
                cogX = 0;
                cogY = 0;
                cogZ = -length / 2;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrWeightCG, "Error while defining weightCG of Anvil_FIG146"));
            }
        }
        #endregion
    }
}
