//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG133.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG133
//   Author       :  Vijay
//   Creation Date:  30-04-2013
//   Description:
//   
//   Anvil_FIG133.cs is same for Anvil_FIG134.cs

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30-04-2013     Vijay    CR-CP-222292 Convert HS_Anvil VB Project to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
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

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class Anvil_FIG133 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG133"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputString(2, "SIZE", "SIZE", "No Vlaue")]
        public InputString m_oSIZE;
        [InputDouble(3, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(4, "BOLT_SIZE", "BOLT_SIZE", 0.999999)]
        public InputDouble m_dBOLT_SIZE;
        [InputDouble(5, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(6, "FT", "FT", 0.999999)]
        public InputDouble m_dFT;
        [InputDouble(7, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(8, "W", "W", 0.999999)]
        public InputDouble m_dW;
        [InputDouble(9, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BOLT", "BOLT")]
        [SymbolOutput("CLAMP1", "CLAMP1")]
        [SymbolOutput("CLAMP2", "CLAMP2")]
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

                String size = m_oSIZE.Value;
                Double B = m_dB.Value;
                Double boltSize = m_dBOLT_SIZE.Value;
                Double A = m_dA.Value;
                Double FT = m_dFT.Value;
                Double T = m_dT.Value;
                Double W = m_dW.Value;
                double fraction = FitInchfractiontoMeter(size);
                if (base.ToDoListMessage != null)
                {
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        return;
                    }
                }
                Double FW = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, fraction, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -(B - boltSize / 2)), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_Symbolic.Outputs["Port2"] = port2;
                
                //Validating Inputs
                if (boltSize <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidBoltSizeGTZero, "Bolt size should be greater than zero."));
                    return;
                }
                if (FT == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFTZero, "FT value cannot be zero"));
                    return;
                }

                Vector normal = new Position(0, -(A / 2 + 2 * T), -B).Subtract(new Position(0, (A / 2 + 2 * T), -B));
                symbolGeometryHelper.ActivePosition = new Position(0, (A / 2 + 2 * T), -B);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d boltCylinder = symbolGeometryHelper.CreateCylinder(null, boltSize / 2, normal.Length);
                m_Symbolic.Outputs["BOLT"] = boltCylinder;

                Projection3d halfClamp1 = CreateHalfClamp(B + T + W / 2, A, FW, FT, W, T, false);
                m_Symbolic.Outputs["CLAMP1"] = halfClamp1;

                Projection3d halfClamp2 = CreateHalfClamp(B + T + W / 2, A, FW, FT, W, T, true);
                m_Symbolic.Outputs["CLAMP2"] = halfClamp2;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG133"));
                    return;
                }
            }
        }

        /// <summary>
        /// Creates a half clamp using Flange Linear Height, Clearance Gap between flanges, Flange width, Flange Inside Arc Diameter, Flange Length and Flange Thickness
        /// </summary>        
        /// <param name="flangeLinearHeight">Flange Linear Height.</param>
        /// <param name="halfClearanceBtFlanges">Half the Clearance Gap between two flanges.</param>        
        /// <param name="flangeWidth"> This is Half of the Clamp Width(i.e One Flange Width from Centre).</param>        
        /// <param name="flangeInsideCircDiameter">Inside Flange Circular Arc Diameter.</param>
        /// <param name="flangeLength">Flange Length along Projected Side.</param>
        /// <param name="flangeThickness">Flange Thickness.</param>
        /// <param name="inverted">Inverted.</param>
        public Projection3d CreateHalfClamp(Double flangeLinearHeight, Double halfClearanceBtFlanges, Double flangeWidth, Double flangeInsideCircDiameter, Double flangeLength, Double flangeThickness, Boolean inverted)
        {
            Matrix4X4 transMatrix = new Matrix4X4();
            SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
            Collection<ICurve> collection = new Collection<ICurve>();
            collection.Add(new Line3d(new Position(-flangeLength / 2, (halfClearanceBtFlanges / 2), -flangeLinearHeight), new Position(-flangeLength / 2, (halfClearanceBtFlanges / 2 + flangeThickness), -flangeLinearHeight)));
            collection.Add(new Line3d(new Position(-flangeLength / 2, (halfClearanceBtFlanges / 2 + flangeThickness), -flangeLinearHeight), new Position(-flangeLength / 2, (halfClearanceBtFlanges / 2 + flangeThickness), -flangeThickness)));
            collection.Add(new Line3d(new Position(-flangeLength / 2, (halfClearanceBtFlanges / 2 + flangeThickness), -flangeThickness), new Position(-flangeLength / 2, (flangeWidth / 2), -flangeThickness)));

            symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
            Arc3d outerArc = symbolGeometryHelper.CreateArc(null, flangeInsideCircDiameter / 2 + flangeThickness, Math.PI);
            transMatrix.Rotate(-Math.PI / 2, new Vector(0, 0, 1));
            transMatrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
            transMatrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
            transMatrix.Translate(new Vector(-flangeLength / 2, (flangeWidth / 2), flangeInsideCircDiameter / 2));
            outerArc.Transform(transMatrix);
            collection.Add(outerArc);

            collection.Add(new Line3d(new Position(-flangeLength / 2, (flangeWidth / 2), flangeInsideCircDiameter + flangeThickness), new Position(-flangeLength / 2, (halfClearanceBtFlanges / 2), flangeInsideCircDiameter + flangeThickness)));
            collection.Add(new Line3d(new Position(-flangeLength / 2, (halfClearanceBtFlanges / 2), flangeInsideCircDiameter + flangeThickness), new Position(-flangeLength / 2, (halfClearanceBtFlanges / 2), flangeInsideCircDiameter)));
            collection.Add(new Line3d(new Position(-flangeLength / 2, (halfClearanceBtFlanges / 2), flangeInsideCircDiameter), new Position(-flangeLength / 2, (flangeWidth / 2), flangeInsideCircDiameter)));

            symbolGeometryHelper = new SymbolGeometryHelper();
            transMatrix = new Matrix4X4();
            symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
            Arc3d innerArc = symbolGeometryHelper.CreateArc(null, flangeInsideCircDiameter / 2, Math.PI);
            transMatrix.Rotate(-Math.PI / 2, new Vector(0, 0, 1));
            transMatrix.Rotate(Math.PI / 2, new Vector(0, 1, 0), new Position(0, 0, 0));
            transMatrix.Rotate(Math.PI / 2, new Vector(1, 0, 0), new Position(0, 0, 0));
            transMatrix.Translate(new Vector(-flangeLength / 2, (flangeWidth / 2), flangeInsideCircDiameter / 2));
            innerArc.Transform(transMatrix);
            collection.Add(innerArc);

            collection.Add(new Line3d(new Position(-flangeLength / 2, (flangeWidth / 2), 0), new Position(-flangeLength / 2, (halfClearanceBtFlanges / 2), 0)));
            collection.Add(new Line3d(new Position(-flangeLength / 2, (halfClearanceBtFlanges / 2), 0), new Position(-flangeLength / 2, (halfClearanceBtFlanges / 2), -flangeLinearHeight)));

            Projection3d halfClamp = new Projection3d(new ComplexString3d(collection), new Vector(1, 0, 0), flangeLength, true);

            if (inverted == true)
            {
                transMatrix = new Matrix4X4();
                transMatrix.Rotate(Math.PI, new Vector(0, 0, 1));
                halfClamp.Transform(transMatrix);
            }
            return halfClamp;
        }

        /// <summary>
        /// Converts fraction value to double value .
        /// </summary>
        /// <param name="fraction">fraction value-String.</param>
        /// <example>
        /// double c=FitInchfractiontoMeter(size);
        /// </example>
        double FitInchfractiontoMeter(string fraction)
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
        #endregion
    }
}
