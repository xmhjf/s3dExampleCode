//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG81C.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG81C
//   Author       :  Manikanth
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
    public class Anvil_FIG81C : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG81C"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "SUSPENSION", "SUSPENSION", 1)]
        public InputDouble m_SUSPENSION;
        [InputDouble(3, "LOAD", "LOAD", 0.999999)]
        public InputDouble m_LOAD;
        [InputDouble(4, "ACT_TRAVEL", "ACT_TRAVEL", 0.999999)]
        public InputDouble m_ACT_TRAVEL;
        [InputDouble(5, "TOTAL_TRAV", "TOTAL_TRAV", 0.999999)]
        public InputDouble m_TOTAL_TRAV;
        [InputDouble(6, "M", "M", 0.999999)]
        public InputDouble m_dM;
        [InputString(7, "SIZE", "SIZE", "No Value")]
        public InputString m_SIZE;
        [InputDouble(8, "C", "C", 0.999999)]
        public InputDouble m_C;
        [InputDouble(9, "G", "G", 0.999999)]
        public InputDouble m_G;
        [InputDouble(10, "FACTOR", "FACTOR", 0.999999)]
        public InputDouble m_FACTOR;
        [InputDouble(11, "N", "N", 0.999999)]
        public InputDouble m_N;
        [InputDouble(12, "D", "D", 0.999999)]
        public InputDouble m_D;
        [InputDouble(13, "E", "E", 0.999999)]
        public InputDouble m_E;
        [InputDouble(14, "F", "F", 0.999999)]
        public InputDouble m_F;
        [InputDouble(15, "H", "H", 0.999999)]
        public InputDouble m_H;
        [InputDouble(16, "TRAVEL_DIR", "TRAVEL_DIR", 1)]
        public InputDouble m_TRAVEL_DIR;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("LUG1", "LUG1")]
        [SymbolOutput("LUG2", "LUG2")]
        [SymbolOutput("LUG3", "LUG3")]
        [SymbolOutput("LUG4", "LUG4")]
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
                Double actTravel = m_ACT_TRAVEL.Value;
                Double totalTravel = m_TOTAL_TRAV.Value;
                Double M = m_dM.Value;
                Double C = m_C.Value;
                Double G = m_G.Value;
                Double factor = m_FACTOR.Value;
                Double N = m_N.Value;
                Double D = m_D.Value;
                Double E = m_E.Value;
                Double F = 0.0;//m_F.Value;
                double suspension = m_SUSPENSION.Value;
                Double H = m_H.Value, J = 0, K = 0, R = 0, B = 0, takeOut = 0, W = 0, S = 0, T = 0;
                string actualSuspension, size = m_SIZE.Value;
                double minsize, maxSize, totalTravelMin, totalTravelMax;
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;

                if (metadataManager != null)
                    actualSuspension = (metadataManager.GetCodelistInfo("Anvil_Constant_Sus", "UDP").GetCodelistItem((int)suspension).ShortDisplayName);
                else
                    actualSuspension = "Single";

                CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                PartClass anvil81CJ = (PartClass)cataloghelper.GetPartClass("Anvil_FIG81C_J");
                ReadOnlyCollection<BusinessObject> anvil81CJclassItems = anvil81CJ.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in anvil81CJclassItems)
                {
                    if (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81C_J", "LOAD_MAX")).PropValue > load) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81C_J", "LOAD_MIN")).PropValue <= load))
                    {
                        totalTravel = 0.5 * ((int)((totalTravel / 0.0254 + 0.499) * 2));
                        J = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81C_J", "J")).PropValue;
                        K = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81C_J", "K")).PropValue;
                        R = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81C_J", "R")).PropValue;
                        W = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81C_J", "W")).PropValue;
                        S = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81C_J", "S")).PropValue;
                        T = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81C_J", "T")).PropValue;
                        break;
                    }
                }

                PartClass anvil81CB = (PartClass)cataloghelper.GetPartClass("Anvil_FIG81C_B");
                ReadOnlyCollection<BusinessObject> anvil81CBclassItems = anvil81CB.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in anvil81CBclassItems)
                {
                    if (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81C_B", "TOTAL_TRAVEL")).PropValue > (totalTravel * 0.0254 - 0.001)) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81C_B", "TOTAL_TRAVEL")).PropValue < (totalTravel * 0.0254 + 0.001)))
                    {
                        B = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81C_B", "B")).PropValue;
                        break;
                    }
                }

                Port port1 = new Port(OccurrenceConnection, part, "InThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;


                PartClass anvil81B = (PartClass)cataloghelper.GetPartClass("Anvil_FIG81C_CandFACTOR");
                ReadOnlyCollection<BusinessObject> anvil81ClassItems = anvil81B.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in anvil81ClassItems)
                {
                    minsize = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, double.Parse(size), UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                    maxSize = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, double.Parse(size), UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                    totalTravelMax = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, totalTravel, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                    totalTravelMin = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, totalTravel, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);

                    if (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81", "MIN_SIZE")).PropValue) <= (minsize) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81", "MAX_SIZE")).PropValue) >= (maxSize) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81", "TOTAL_TRAVEL_MAX")).PropValue) >= (totalTravelMax) && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81", "TOTAL_TRAVEL_MIN")).PropValue) <= (totalTravelMin) && ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrAnvil_FIG81", "TYPE")).PropValue) == "C")
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
                                    C = 5.5;
                                    factor = 14.625;
                                }
                                else if (totalTravel >= 4.5)
                                {
                                    C = 9.5;
                                    factor = 17.1875;
                                }
                            }

                            if (int.Parse(size) >= 10 && int.Parse(size) <= 18)
                            {
                                if (totalTravel <= 5)
                                {
                                    C = 7.5;
                                    factor = 13.0625;
                                }
                                else if (totalTravel >= 5.5)
                                {
                                    C = 10.5;
                                    factor = 15.4375;
                                }

                            }

                            if (int.Parse(size) >= 19 && int.Parse(size) <= 34)
                            {
                                if (totalTravel <= 5)
                                {
                                    C = 8.5;
                                    factor = 19.875;
                                }
                                else if (totalTravel >= 5.5)
                                {
                                    C = 8.5;
                                    factor = 19.875;
                                }
                            }

                            if (int.Parse(size) >= 35 && int.Parse(size) <= 49)
                            {
                                if (totalTravel <= 6)
                                {
                                    C = 9.5;
                                    factor = 25.625;
                                }
                                else if (totalTravel >= 6.5)
                                {
                                    C = 17.5;
                                    factor = 30.125;
                                }
                            }

                            if (int.Parse(size) >= 50 && int.Parse(size) <= 63)
                            {
                                if (totalTravel <= 8)
                                {
                                    C = 13.375;
                                    factor = 30.6875;
                                }
                                else if (totalTravel >= 8.5 && totalTravel <= 11)
                                {
                                    C = 21.375;
                                    factor = 30.6875;
                                }
                                else if (totalTravel >= 11.5)
                                {
                                    C = 21.375;
                                    factor = 36;
                                }
                            }

                            if (int.Parse(size) >= 64 && int.Parse(size) <= 74)
                            {
                                if (totalTravel <= 10.5)
                                {
                                    C = 13.75;
                                    factor = 42.375;
                                }
                                else if (totalTravel >= 11)
                                {
                                    C = 19.25;
                                    factor = 42.5;
                                }
                            }

                            if (int.Parse(size) >= 75 && int.Parse(size) <= 83)
                            {
                                if (totalTravel <= 10.5)
                                {
                                    C = 12.75;
                                    factor = 45.75;
                                }
                                else if (totalTravel >= 11)
                                {
                                    C = 18.25;
                                    factor = 45.75;
                                }
                            }
                            C = C * 0.0254;
                            takeOut = (factor - (totalTravel / 2)) * 0.0254;
                        }
                        break;
                    }
                }

                //Validating Inputs
                if (J <= 0 || T <= 0 || W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTotalLoad, "Total load value must be between LOAD_MAX and LOAD_MIN"));
                    return;
                }
                if ((H + R) <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidHAndR, "H + R value should be greater than zero"));
                    return;
                }
                if (N <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroNvalue, "N value should be greater than zero"));
                    return;
                }
                if (M <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroMvalue, "M value should be greater than zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidZeroDvalue, "D value should be greater than zero"));
                    return;
                }
                if ((C + E) <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidCAndE, "C + E value should be greater than zero"));
                    return;
                }              

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                if (actualSuspension == "Single" && int.Parse(size) < 64)
                {
                    Port port2 = new Port(OccurrenceConnection, part, "LeftHole", new Position(0, 0, takeOut + K / 2), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Port2"] = port2;

                    Port port3 = new Port(OccurrenceConnection, part, "RightHole", new Position(0, W, takeOut + K / 2), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Port3"] = port3;

                    matrix.Translate(new Vector(-W * 1.5, S / 2, takeOut - H));
                    matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    Projection3d box1 = symbolGeometryHelper.CreateBox(null, W * 2, T, H + R, 9);
                    box1.Transform(matrix);
                    m_Symbolic.Outputs["LUG1"] = box1;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-W * 1.5, -S / 2 - T, takeOut - H));
                    matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    Projection3d box2 = symbolGeometryHelper.CreateBox(null, W * 2, T, H + R, 9);
                    box2.Transform(matrix);
                    m_Symbolic.Outputs["LUG2"] = box2;

                }
                else
                {
                    Port port2 = new Port(OccurrenceConnection, part, "LeftHole", new Position(0, B + G - C, takeOut + K / 2), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Port2"] = port2;

                    Port port3 = new Port(OccurrenceConnection, part, "RightHole", new Position(0, B + G, takeOut + K / 2), new Vector(1, 0, 0), new Vector(0, 0, -1));
                    m_Symbolic.Outputs["Port3"] = port3;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-B - G - W / 2, -S / 2 - T, takeOut - H));
                    matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    Projection3d lug1Box = symbolGeometryHelper.CreateBox(null, W, T, H + R, 9);
                    lug1Box.Transform(matrix);
                    m_Symbolic.Outputs["LUG1"] = lug1Box;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-B - G - W / 2, S / 2, takeOut - H));
                    matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    Projection3d lug2Box = symbolGeometryHelper.CreateBox(null, W, T, H + R, 9);
                    lug2Box.Transform(matrix);
                    m_Symbolic.Outputs["LUG2"] = lug2Box;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-B - G - W / 2 + C, -S / 2 - T, takeOut - H));
                    matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    Projection3d lug3Box = symbolGeometryHelper.CreateBox(null, W, T, H + R, 9);
                    lug3Box.Transform(matrix);
                    m_Symbolic.Outputs["LUG3"] = lug3Box;

                    symbolGeometryHelper = new SymbolGeometryHelper();
                    matrix = new Matrix4X4();
                    matrix.Translate(new Vector(-B - G - W / 2 + C, S / 2, takeOut - H));
                    matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                    Projection3d lug4Box = symbolGeometryHelper.CreateBox(null, W, T, H + R, 9);
                    lug4Box.Transform(matrix);
                    m_Symbolic.Outputs["LUG4"] = lug4Box;

                }

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), (new Vector(0, 0, 1)).GetOrthogonalVector());
                matrix = new Matrix4X4();
                matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 1, 0));
                matrix.Translate(new Vector(-B - G - E, 0, takeOut - H - M / 2 - F));
                matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d springCylinder = symbolGeometryHelper.CreateCylinder(null, M / 2, D);
                springCylinder.Transform(matrix);
                m_Symbolic.Outputs["SPRING"] = springCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-B - G - E, -N / 2, takeOut - H - 0.9 * M));
                matrix.Rotate((3 * Math.PI) / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d bottomBox = symbolGeometryHelper.CreateBox(null, C + E * 2, N, 0.9 * M, 9);
                bottomBox.Transform(matrix);
                m_Symbolic.Outputs["BOTTOM"] = bottomBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(0, 0, -J * 1.5));
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d thingyCylinder = symbolGeometryHelper.CreateCylinder(null, J, takeOut - H - 0.9 * M + J * 1.5);
                thingyCylinder.Transform(matrix);
                m_Symbolic.Outputs["THINGY"] = thingyCylinder;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG81C"));
                    return;
                }
            }
        }
        #endregion

    }

}
