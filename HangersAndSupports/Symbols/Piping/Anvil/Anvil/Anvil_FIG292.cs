//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG292.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG292
//   Author       :  Hema
//   Creation Date:  30-04-2013
//   Description:    Converted Anvil_Parts VB Project to C#.Net Project

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30-04-2013     Hema    CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

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

    public class Anvil_FIG292 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG292"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "FLANGE_WIDTH", "FLANGE_WIDTH", 1)]
        public InputDouble m_oFLANGE_WIDTH;
        [InputDouble(3, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble m_dROD_DIA;
        [InputString(4, "SIZE", "SIZE", "No Value")]
        public InputString m_oSIZE;
        [InputDouble(5, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(6, "V", "V", 0.999999)]
        public InputDouble m_dV;
        [InputDouble(7, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("EYE", "EYE")]
        [SymbolOutput("PIN", "PIN")]
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
                Double flangeWidthValue = m_oFLANGE_WIDTH.Value;
                String size = m_oSIZE.Value;
                Double B = m_dB.Value;
                Double V = m_dV.Value;
                String flangeWidth;

                if (flangeWidthValue < 1 || flangeWidthValue > 13)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFlangewCodelist, "FLANGE_WIDTH codelist value should be between 1 & 13."));
                    return;
                }
                Double E = 0, T = 0.01905, Q;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (metadataManager != null)
                    flangeWidth = metadataManager.GetCodelistInfo("Anvil_FIG292_FlangeW", "UDP").GetCodelistItem((int)flangeWidthValue).ShortDisplayName.Trim();
                else
                    flangeWidth = "3";

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass anvilfig292Parts = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG292_TO");
                ReadOnlyCollection<BusinessObject> fig292Class = anvilfig292Parts.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItems in fig292Class)
                {
                    if (((string)((PropertyValueString)classItems.GetPropertyValue("IJUAHgrAnvil_FIG292_TO", "SIZE")).PropValue == size))
                    {
                        E = (double)((PropertyValueDouble)classItems.GetPropertyValue("IJUAHgrAnvil_FIG292_TO", "E_" + flangeWidth)).PropValue;
                        break;
                    }
                }

                //This is because when E = 0, the ports lie on top of each other and the part won't place.
                if (E == 0)
                    E = 0.001;
                double flangeWidthDoubleValue=FeettInchFractiontoMeter(flangeWidth);
                if (base.ToDoListMessage != null)
                {
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        return;
                    }
                }
                Q = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, flangeWidthDoubleValue, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER) / 2;

                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, -E), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidRodDiaGTZero, "Rod diameter should be greater than zero"));
                    return;
                }
                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTGTZero, "T value should be greater than zero"));
                    return;
                }

                //This is because when E = 0, the ports lie on top of each other and the part won't place.
                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(0, 0, -E + B));
                pointCollection.Add(new Position(0, Q, 0));
                pointCollection.Add(new Position(0, Q - V, 0.6 * T));
                pointCollection.Add(new Position(0, Q, 0.9 * T));
                pointCollection.Add(new Position(0, Q + T, 0));
                pointCollection.Add(new Position(0, rodDiameter, rodDiameter - E));
                pointCollection.Add(new Position(0, -rodDiameter, rodDiameter - E));
                pointCollection.Add(new Position(0, -Q - T, 0));
                pointCollection.Add(new Position(0, -Q, 0.9 * T));
                pointCollection.Add(new Position(0, -Q + V, 0.6 * T));
                pointCollection.Add(new Position(0, -Q, 0));
                pointCollection.Add(new Position(0, 0, -E + B));

                Projection3d body = new Projection3d(new LineString3d(pointCollection), new Vector(1, 0, 0), 0.001, true);
                m_Symbolic.Outputs["BODY"] = body;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, 0, -E - B).Subtract(new Position(0, 0, -E + rodDiameter));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -E + rodDiameter);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d eye = symbolGeometryHelper.CreateCylinder(null, rodDiameter * 0.75, normal.Length);
                m_Symbolic.Outputs["EYE"] = eye;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal1 = new Position(0, -(Q + 1.3 * T), -T / 8).Subtract(new Position(0, Q + 1.3 * T, -T / 8));
                symbolGeometryHelper.ActivePosition = new Position(0, Q + 1.3 * T, -T / 8);
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());
                Projection3d pin = symbolGeometryHelper.CreateCylinder(null, T / 8, normal1.Length);
                m_Symbolic.Outputs["PIN"] = pin;

            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG292"));
                    return;
                }
            }
        }
        /// <summary>
        /// Converts fraction value to double value .
        /// </summary>
        /// <param name="fraction">fraction value-String.</param>
        /// <example>
        /// double c=FeettInchFractiontoMeter(size);
        /// </example>
        double FeettInchFractiontoMeter(string fraction)
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
            ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFraction, "Input fraction is not in the correct format"));
            return 0;
        }
        #endregion
    }
}
