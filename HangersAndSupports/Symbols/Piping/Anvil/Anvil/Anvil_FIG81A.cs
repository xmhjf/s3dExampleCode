//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG81A.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG81A
//   //   Author       :  Manikanth
//   Creation Date:  16-05-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   15-05-2013    Manikanth CR-CP-233113-Convert HS_Anvil VB Project to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
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
    public class Anvil_FIG81A : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG81A"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "SUSPENSION", "SUSPENSION", 1)]
        public InputDouble m_SUSPENSION;
        [InputDouble(3, "LOAD", "LOAD", 0.999999)]
        public InputDouble m_LOAD;
        [InputDouble(4, "P", "P", 0.999999)]
        public InputDouble m_P;
        [InputDouble(5, "TOTAL_TRAV", "TOTAL_TRAV", 0.999999)]
        public InputDouble m_TOTAL_TRAV;
        [InputDouble(6, "N", "N", 0.999999)]
        public InputDouble m_N;
        [InputString(7, "SIZE", "SIZE", "No Value")]
        public InputString m_SIZE;
        [InputDouble(8, "G", "G", 0.999999)]
        public InputDouble m_G;
        [InputDouble(9, "D", "D", 0.999999)]
        public InputDouble m_D;
        [InputDouble(10, "E", "E", 0.999999)]
        public InputDouble m_E;
        [InputDouble(11, "F", "F", 0.999999)]
        public InputDouble m_F;
        [InputDouble(12, "M", "M", 0.999999)]
        public InputDouble m_M;
        [InputDouble(13, "ACT_TRAVEL", "ACT_TRAVEL", 0.999999)]
        public InputDouble m_ACT_TRAVEL;
        [InputDouble(14, "TRAVEL_DIR", "TRAVEL_DIR", 1)]
        public InputDouble m_TRAVEL_DIR;
        #endregion

        #region "Definitions of Aspects && their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("SPRING", "SPRING")]
        [SymbolOutput("BOTTOM", "BOTTOM")]
        [SymbolOutput("THINGY", "THINGY")]
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

                Double load = m_LOAD.Value;
                Double P = m_P.Value;
                Double totalTravel = m_TOTAL_TRAV.Value;
                Double N = m_N.Value;
                Double G = m_G.Value;
                Double D = m_D.Value;
                Double E = m_E.Value;
                Double F = m_F.Value;
                Double M = m_M.Value;
                Double actTravel = m_ACT_TRAVEL.Value;
                Double suspension = m_SUSPENSION.Value, J = 0, B = 0, C = 0, factor = 0, takeOut = 0;
                string actualSuspension, size = m_SIZE.Value;
                double minsize, maxSize, totalTravelMin, totalTravelMax;
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (metadataManager != null)
                    actualSuspension = (metadataManager.GetCodelistInfo("Anvil_Constant_Sus", "UDP").GetCodelistItem((int)suspension).ShortDisplayName);
                else
                    actualSuspension = "Single";

                CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                PartClass anvil81AJ = (PartClass)cataloghelper.GetPartClass("Anvil_FIG81A_J");
                ReadOnlyCollection<BusinessObject> anvil81AJclassItems = anvil81AJ.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in anvil81AJclassItems)
                {
                    if (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81A_J", "LOAD_MAX")).PropValue > load) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81A_J", "LOAD_MIN")).PropValue <= load))
                    {
                        totalTravel = 0.5 * ((int)((totalTravel / 0.0254 + 0.499) * 2));
                        J = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81A_J", "J")).PropValue;
                        break;
                    }
                }

                PartClass anvil81AB = (PartClass)cataloghelper.GetPartClass("Anvil_FIG81A_B");
                ReadOnlyCollection<BusinessObject> anvil81ABclassItems = anvil81AB.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in anvil81ABclassItems)
                {
                    if (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81A_B", "TOTAL_TRAVEL")).PropValue > (totalTravel * 0.0254 - 0.001)) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81A_B", "TOTAL_TRAVEL")).PropValue <= (totalTravel * 0.0254 + 0.001)))
                    {
                        B = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81A_B", "B")).PropValue;
                        break;
                    }
                }

                Port port1 = new Port(OccurrenceConnection, part, "BotInThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                PartClass anvil81 = (PartClass)cataloghelper.GetPartClass("Anvil_FIG81A_CandFACTOR");
                ReadOnlyCollection<BusinessObject> anvil81AclassItems = anvil81.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in anvil81AclassItems)
                {
                    minsize = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, double.Parse(size), UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                    maxSize = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, double.Parse(size), UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                    totalTravelMax = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, totalTravel, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                    totalTravelMin = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, totalTravel, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);

                    if (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81", "MIN_SIZE")).PropValue) <= (minsize) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81", "MAX_SIZE")).PropValue) >= (maxSize) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81", "TOTAL_TRAVEL_MAX")).PropValue) >= (totalTravelMax) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81", "TOTAL_TRAVEL_MIN")).PropValue) <= (totalTravelMin) && ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81", "TYPE")).PropValue) == "A")
                    {
                        try
                        {
                            C = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81", "C")).PropValue;
                            factor = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81", "FACTOR")).PropValue;
                            takeOut = (factor - ((totalTravel * 0.0254) / 2));
                        }
                        catch
                        {
                            if (int.Parse(size) >= 1 && int.Parse(size) <= 9)
                            {
                                if (totalTravel <= 4)
                                {
                                    C = 6;
                                    factor = 12.75;
                                }
                                else if (totalTravel >= 4.5)
                                {
                                    C = 10;
                                    factor = 15.3125;
                                }
                            }

                            if (int.Parse(size) >= 10 && int.Parse(size) <= 18)
                            {
                                if (totalTravel <= 5)
                                {
                                    C = 8;
                                    factor = 10.875;
                                }
                                else if (totalTravel >= 5.5)
                                {
                                    C = 11;
                                    factor = 13.25;
                                }

                            }

                            if (int.Parse(size) >= 19 && int.Parse(size) <= 34)
                            {
                                if (totalTravel <= 5)
                                {
                                    C = 10;
                                    factor = 16.75;
                                }
                                else if (totalTravel >= 5.5)
                                {
                                    C = 10;
                                    factor = 16.75;
                                }
                            }

                            if (int.Parse(size) >= 35 && int.Parse(size) <= 49)
                            {
                                if (totalTravel <= 6)
                                {
                                    C = 11;
                                    factor = 21.125;
                                }
                                else if (totalTravel >= 6.5)
                                {
                                    C = 19;
                                    factor = 25.75;
                                }
                            }

                            if (int.Parse(size) >= 50 && int.Parse(size) <= 63)
                            {
                                if (totalTravel <= 8)
                                {
                                    C = 16;
                                    factor = 24.9375;
                                }
                                else if (totalTravel >= 8.5 && totalTravel <= 11)
                                {
                                    C = 24;
                                    factor = 24.9375;
                                }
                                else if (totalTravel >= 11.5)
                                {
                                    C = 24;
                                    factor = 30.25;
                                }
                            }

                            if (int.Parse(size) >= 64 && int.Parse(size) <= 74)
                            {
                                if (totalTravel <= 10.5)
                                {
                                    C = 15.75;
                                    factor = 34.4375;
                                }
                                else if (totalTravel >= 11)
                                {
                                    C = 21.25;
                                    factor = 34.5625;
                                }
                            }

                            if (int.Parse(size) >= 75 && int.Parse(size) <= 83)
                            {
                                if (totalTravel <= 10.5)
                                {
                                    C = 15.25;
                                    factor = 36.5;
                                }
                                else if (totalTravel >= 11)
                                {
                                    C = 20.25;
                                    factor = 45.875;
                                }
                            }
                            C = C * 0.0254;
                            takeOut = (factor - (totalTravel / 2)) * 0.0254;
                        }
                        break;
                    }
                }

                if (actualSuspension == "Single" && int.Parse(size) < 64)
                {
                    Port port2 = new Port(OccurrenceConnection, part, "TopInThdRH", new Position(0, 0, takeOut), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Port2"] = port2;
                }
                else
                {
                    Port port2 = new Port(OccurrenceConnection, part, "TopInThdRH", new Position(0, B + G - C, takeOut), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Port2"] = port2;

                    Port port3 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, B + G, takeOut), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Port3"] = port3;
                }

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //Validating Inputs
                if (M <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroMvalue, "M value should be greater than zero"));
                    return;
                }
                if (N <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroNvalue, "N value should be greater than zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroDvalue, "D value should be greater than zero"));
                    return;
                }
                if (J <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTotalLoad, "Total load value must be between LOAD_MAX and LOAD_MIN"));
                    return;
                }

                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 1, 0));
                matrix.Translate(new Vector(-B - G - E, 0, takeOut + P - M / 2 - F));
                matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d springcylinder = symbolGeometryHelper.CreateCylinder(null, M / 2, D);
                springcylinder.Transform(matrix);
                m_Symbolic.Outputs["SPRING"] = springcylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-B - G - E, -N / 2, takeOut + P - 0.9 * M));
                matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d bottomBox = symbolGeometryHelper.CreateBox(null, C + E * 2, N, 0.9 * M, 9);
                bottomBox.Transform(matrix);
                m_Symbolic.Outputs["BOTTOM"] = bottomBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(0, 0, -J * 1.5));
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d thingyCylinder = symbolGeometryHelper.CreateCylinder(null, J, takeOut + P - 0.9 * M + J * 1.5);
                thingyCylinder.Transform(matrix);
                m_Symbolic.Outputs["THINGY"] = thingyCylinder;
            }

            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG81A"));
                    return;
                }
            }
        }
        #endregion

    }

}
