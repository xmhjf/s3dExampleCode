//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG291.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG291
//   Author       :  Hema
//   Creation Date:  30-04-2013
//   Description:    Converted Anvil_Parts VB Project to C#.Net Project

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30-04-2013     Hema    CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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

    public class Anvil_FIG291 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG291"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputString(2, "SIZE", "SIZE", "No Value")]
        public InputString m_oSIZE;
        [InputDouble(3, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(4, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(5, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("CYL", "CYL")]
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

                string size = m_oSIZE.Value;
                Double L = m_dL.Value;
                double fraction = FeetInchFractiontoMeter(size);
                if (base.ToDoListMessage != null)
                {
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        return;
                    }
                }
                double D = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, fraction, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);

                //Intializing SymbolGeomHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "BotPin", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "TopPin", new Position(0, 0, D), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidDGTZero, "D value should be greater than zero"));
                    return;
                }           

                Vector normal = new Position(0, L / 2, D / 2).Subtract(new Position(0, -L / 2, D / 2));
                symbolGeometryHelper.ActivePosition = new Position(0, -L / 2, D / 2);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d cylinder = symbolGeometryHelper.CreateCylinder(null, D / 2, normal.Length);
                m_Symbolic.Outputs["CYL"] = cylinder;

            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG291"));
                    return;
                }
            }
        }
        #endregion
        
        /// <summary>
        /// Converts fraction value to double value .
        /// </summary>
        /// <param name="fraction">fraction value-String.</param>
        /// <example>
        /// double c=FitInchfractiontoMeter(size);
        /// </example>
        double FeetInchFractiontoMeter(string fraction)
        {
            double result;

            if (double.TryParse(fraction, out result))
                return result;

            string[] split = fraction.Split(new char[] { ' ', '/' });
            int a, b, c, d;

            if (split.Length == 2 || split.Length == 3)
            {
                if (int.TryParse(split[0], out a) && int.TryParse(split[1], out b))
                {
                    if (split.Length == 2)
                        return (double)a / b;
                    if (int.TryParse(split[2], out c))
                        return a + (double)b / c;
                }
            }
            else if (split.Length == 4)
            {

                if (int.TryParse(split[0], out a) && int.TryParse(split[1], out b) && int.TryParse(split[2], out c))
                {
                    if (int.TryParse(split[3], out d))
                        return (a * 304.8 + (b + (double)c / d) * 25.4) / 25.4;
                }
            }
            ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFraction, "The Input Fraction is not in Correct Format"));
            return 0;
        }
    }
}
