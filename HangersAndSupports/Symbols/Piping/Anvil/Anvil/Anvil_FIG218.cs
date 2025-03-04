//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG218.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG218
//   Author       :  Rajeswari
//   Creation Date:  09-May-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   09-May-2013  Rajeswari  CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class Anvil_FIG218 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG218"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "FLANGE_WIDTH", "FLANGE_WIDTH", 1)]
        public InputDouble m_oFLANGE_WIDTH;
        [InputString(3, "SIZE", "SIZE", "No Value")]
        public InputString m_oSIZE;
        [InputDouble(4, "BOLT_DIA", "BOLT_DIA", 0.999999)]
        public InputDouble m_dBOLT_DIA;
        [InputDouble(5, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("BOLT", "BOLT")]
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

                Double boltDiameter = m_dBOLT_DIA.Value;
                Double flangeWidthValue = m_oFLANGE_WIDTH.Value;
                string flangeWidth;

                //Validating Inputs
                if (flangeWidthValue < 1 || flangeWidthValue > 6)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFlangewCodelist, "FLANGE_WIDTH codelist value should be between 1 & 6."));
                    return;
                }
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (metadataManager != null)
                    flangeWidth = metadataManager.GetCodelistInfo("Anvil_FIG218_FlangeW", "UDP").GetCodelistItem((int)flangeWidthValue).ShortDisplayName.Trim();
                else
                    flangeWidth = "3";

                Double E = 0, W = 0, T = 0;

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass anvilfig218TOParts = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG218_TO");
                ReadOnlyCollection<BusinessObject> fig218TOParts = anvilfig218TOParts.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject part1 in fig218TOParts)
                {
                    if (((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrAnvil_FIG218_TO", "FLANGE_WIDTH")).PropValue == flangeWidth))
                    {
                        E = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG218_TO", "TAKE_OUT")).PropValue;
                        double flangeWidthDoubleValue = FeettInchFractiontoMeter(flangeWidth);
                        if (base.ToDoListMessage != null)
                        {
                            if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                            {
                                return;
                            }
                        }
                        W = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, flangeWidthDoubleValue, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                        T = 0.01905;
                        break;
                    }
                }

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, -E), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTGTZero, "T value should be greater than zero"));
                    return;
                }
                if (boltDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidBoltDiaGTZero, "Bolt diameter should be greater than zero"));
                    return;
                }

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(0, 0, (-E) + boltDiameter - boltDiameter / 2));
                pointCollection.Add(new Position(0, W / 2, 0));
                pointCollection.Add(new Position(0, W / 2 - 0.8 * T, 0.6 * T));
                pointCollection.Add(new Position(0, W / 2 - 0.5 * T, 0.9 * T));
                pointCollection.Add(new Position(0, W / 2 + T, 0));
                pointCollection.Add(new Position(0, boltDiameter, -boltDiameter * 1.5 - E));
                pointCollection.Add(new Position(0, -boltDiameter, -boltDiameter * 1.5 - E));
                pointCollection.Add(new Position(0, -W / 2 - T, 0));
                pointCollection.Add(new Position(0, -W / 2 + 0.5 * T, 0.9 * T));
                pointCollection.Add(new Position(0, -W / 2 + 0.8 * T, 0.6 * T));
                pointCollection.Add(new Position(0, -W / 2, 0));
                pointCollection.Add(new Position(0, 0, (-E) + boltDiameter - boltDiameter / 2));
                Projection3d body = new Projection3d(new LineString3d(pointCollection), new Vector(1, 0, 0), 0.001, true);
                m_Symbolic.Outputs["BODY"] = body;

                Vector normal = new Position(-boltDiameter, 0, (-E) - boltDiameter / 2).Subtract(new Position(boltDiameter, 0, (-E) - boltDiameter / 2));
                symbolGeometryHelper.ActivePosition = new Position(boltDiameter, 0, (-E) - boltDiameter / 2);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d bolt = symbolGeometryHelper.CreateCylinder(null, boltDiameter / 2, normal.Length);
                m_Symbolic.Outputs["BOLT"] = bolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, -(W / 2 + 1.3 * T), -T / 8).Subtract(new Position(0, W / 2 + 1.3 * T, -T / 8));
                symbolGeometryHelper.ActivePosition = new Position(0, W / 2 + 1.3 * T, -T / 8);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d pin = symbolGeometryHelper.CreateCylinder(null, T / 8, normal.Length);
                m_Symbolic.Outputs["PIN"] = pin;
            }
            catch//General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG218"));
                return;
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
            ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFraction, " Input fraction is not in correct format"));
            return 0;
        }

        #endregion

    }

}
