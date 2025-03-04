//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Bline_Clamp_9_1249X.cs
//   Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_Clamp_9_1249X
//   Author       :  Vijaya
//   Creation Date:  30.July.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30.July.2012    Vijaya     CR-CP-219113 Converted HS_B-Line VB project to C#.Net Project 
//   26.Nov.2012     Rajeswari  CR-CP-219113 Modified the code with SymbolGeomHelper
//   26.Mar.2013     Vijaya     DI-CP-228142  Modify the error handling for delivered H&S symbols 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;


namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Support.Content.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Bline_Clamp_9_1249X : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_Clamp_9_1249X"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Gap", "Gap between Clamp and Web Rail", 0)]
        public InputDouble m_dGap;
        [InputDouble(3, "LP", "Length of Plate", 0)]
        public InputDouble m_dPlateLength;
        [InputDouble(4, "TP", "Thickness of Plate", 0)]
        public InputDouble m_dPlateThickness;
        [InputDouble(5, "WP", "Width of Plate", 0)]
        public InputDouble m_dPlateWidth;
        [InputDouble(6, "TC", "Thickness of Clamp", 0)]
        public InputDouble m_dClampThickness;
        [InputDouble(7, "WC", "Width of Clamp", 0)]
        public InputDouble m_dCWidth;
        [InputDouble(8, "HC", "Height of Clamp", 0)]
        public InputDouble m_dClampHeight;
        [InputDouble(9, "LU", "Length of Upper Arm of Clamp", 0)]
        public InputDouble m_dUpperArmLength;
        [InputDouble(10, "LB", "Length of Bottom Arm of Clamp", 0)]
        public InputDouble m_dBottomArmLength;
        [InputDouble(11, "Inside_Outside", "Inside or Outside", 0)]
        public InputDouble m_dInside_Outside;
        [InputDouble(12, "WT", "Tray Width", 0)]
        public InputDouble m_dTrayWidth;
        [InputDouble(13, "TrayWT", "Tray Web Thickness", 0)]
        public InputDouble m_dTrayWebThickness;
        [InputDouble(14, "HoleDiameter", "Hole Diameter", 0)]
        public InputDouble m_dHDiameter;
        [InputDouble(15, "Flip_Clamp", "Flip Clamp", 0)]
        public InputDouble m_dFlipClamp;
        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("PLATE", "PLATE")]
        [SymbolOutput("CLAMP", "CLAMP")]
        [SymbolOutput("BOT_BOLT", "BOT_BOLT")]

        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;

                double GA = m_dGap.Value;
                double LP = m_dPlateLength.Value;
                double TP = m_dPlateThickness.Value;
                double WP = m_dPlateWidth.Value;

                double TC = m_dClampThickness.Value;
                double WC = m_dCWidth.Value;
                double HC = m_dClampHeight.Value;
                double LU = m_dUpperArmLength.Value;
                double LB = m_dBottomArmLength.Value;
                double inOrOut = m_dInside_Outside.Value;
                double WT = m_dTrayWidth.Value;
                double TW = m_dTrayWebThickness.Value;
                double DO = m_dHDiameter.Value;
                double flipClamp = m_dFlipClamp.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //================================================= , , 
                if (DO <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidDO, "Hole Diameter value should  be greater than 0."));
                if (LP <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidPlateL, "Plate Length value should  be greater than 0."));
                if (WP <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidPlateW, "Plate Width  value should  be greater than 0."));
                if (TP <= 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidPlateT, "Plate Thickness value should  be greater than 0."));

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "TrayPort", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port1"] = port1;

                Collection<Position> pointCollection = new Collection<Position>();
                Vector projectionVector = new Vector();
                Vector normal1;

                PropertyValueCodelist insideOutsideCodelistValue = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrBlineInsideOutside", "Inside_Outside");
                if (inOrOut <= 0 || inOrOut > 2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidInorOut, "inOrOut value should be between 1 and 2"));
                PropertyValueCodelist flipClampCodelistValue = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrBlineFlipClamp", "Flip_Clamp");
                if (flipClamp <= 0 || flipClamp > 2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidflipClamp, "flipClamp value should be between 1 and 2"));
                string flipClampValue = flipClampCodelistValue.PropertyInfo.CodeListInfo.GetCodelistItem((int)flipClamp).DisplayName.ToLower();

                if (insideOutsideCodelistValue.PropertyInfo.CodeListInfo.GetCodelistItem((int)inOrOut).DisplayName.ToLower() == "outside")
                {
                    if (flipClampValue == "no")
                    {
                        symbolGeometryHelper.ActivePosition = new Position(0, WT / 2 + TW + GA, 1.5 * TC);
                        symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                        Projection3d plate = (Projection3d)symbolGeometryHelper.CreateBox(null, LP, WP, TP);
                        m_PhysicalAspect.Outputs["PLATE"] = plate;

                        //Add Extruded U-Shape for CLAMP
                        pointCollection.Add(new Position(WP / 2, WT / 2 + TW + GA + LP / 2 - WC / 2, TC));
                        pointCollection.Add(new Position(WP / 2 - LU, WT / 2 + TW + GA + LP / 2 - WC / 2, TC));
                        pointCollection.Add(new Position(WP / 2 - LU, WT / 2 + TW + GA + LP / 2 - WC / 2, 0));
                        pointCollection.Add(new Position(WP / 2 - TC, WT / 2 + TW + GA + LP / 2 - WC / 2, 0));
                        pointCollection.Add(new Position(WP / 2 - TC, WT / 2 + TW + GA + LP / 2 - WC / 2, -HC + 2 * TC));
                        pointCollection.Add(new Position(WP / 2 - LB, WT / 2 + TW + GA + LP / 2 - WC / 2, -HC + 2 * TC));
                        pointCollection.Add(new Position(WP / 2 - LB, WT / 2 + TW + GA + LP / 2 - WC / 2, -HC + TC));
                        pointCollection.Add(new Position(WP / 2, WT / 2 + TW + GA + LP / 2 - WC / 2, -HC + TC));
                        pointCollection.Add(new Position(WP / 2, WT / 2 + TW + GA + LP / 2 - WC / 2, TC));

                        projectionVector.Set(0, WC, 0);
                        Projection3d clamp1 = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                        m_PhysicalAspect.Outputs["CLAMP"] = clamp1;

                        normal1 = new Position(0, WT / 2 + TW + GA + LP / 2, -HC + TC + TC).Subtract(new Position(0, WT / 2 + TW + GA + LP / 2, -HC + TC - 0.25 * TC));
                        symbolGeometryHelper = new SymbolGeometryHelper();
                        symbolGeometryHelper.ActivePosition = new Position(0, WT / 2 + TW + GA + LP / 2, -HC + TC - 0.25 * TC);
                        symbolGeometryHelper.SetOrientation(normal1, new Vector(1, 0, 0));
                        Projection3d bottomBolt = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal1.Length);
                        m_PhysicalAspect.Outputs["BOT_BOLT"] = bottomBolt;
                    }

                    else if (flipClampValue == "yes")
                    {
                        symbolGeometryHelper.ActivePosition = new Position(0, -(WT / 2 + TW + GA), 1.5 * TC);
                        symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(1, 0, 0));
                        Projection3d plate = (Projection3d)symbolGeometryHelper.CreateBox(null, LP, WP, TP);
                        m_PhysicalAspect.Outputs["PLATE"] = plate;

                        pointCollection.Add(new Position(WP / 2, -(WT / 2 + TW + GA + LP / 2 - WC / 2), TC));
                        pointCollection.Add(new Position((WP / 2 - LU), -(WT / 2 + TW + GA + LP / 2 - WC / 2), TC));
                        pointCollection.Add(new Position((WP / 2 - LU), -(WT / 2 + TW + GA + LP / 2 - WC / 2), 0));
                        pointCollection.Add(new Position((WP / 2 - TC), -(WT / 2 + TW + GA + LP / 2 - WC / 2), 0));
                        pointCollection.Add(new Position((WP / 2 - TC), -(WT / 2 + TW + GA + LP / 2 - WC / 2), -HC + 2 * TC));
                        pointCollection.Add(new Position((WP / 2 - LB), -(WT / 2 + TW + GA + LP / 2 - WC / 2), -HC + 2 * TC));
                        pointCollection.Add(new Position((WP / 2 - LB), -(WT / 2 + TW + GA + LP / 2 - WC / 2), -HC + TC));
                        pointCollection.Add(new Position(WP / 2, -(WT / 2 + TW + GA + LP / 2 - WC / 2), -HC + TC));
                        pointCollection.Add(new Position(WP / 2, -(WT / 2 + TW + GA + LP / 2 - WC / 2), TC));

                        projectionVector.Set(0, -WC, 0);
                        Projection3d clamp2 = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                        m_PhysicalAspect.Outputs["CLAMP"] = clamp2;

                        normal1 = new Position(0, -WT / 2 - TW - GA - LP / 2, -HC + TC + TC).Subtract(new Position(0, -WT / 2 - TW - GA - LP / 2, -HC + TC - 0.25 * TC));
                        symbolGeometryHelper = new SymbolGeometryHelper();
                        symbolGeometryHelper.ActivePosition = new Position(0, -WT / 2 - TW - GA - LP / 2, -HC + TC - 0.25 * TC);
                        symbolGeometryHelper.SetOrientation(normal1, new Vector(1, 0, 0));
                        Projection3d bottomBolt = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal1.Length);
                        m_PhysicalAspect.Outputs["BOT_BOLT"] = bottomBolt;
                    }
                    pointCollection.Clear();
                }
                else
                {

                    if (flipClampValue == "no")
                    {
                        symbolGeometryHelper.ActivePosition = new Position(0, WT / 2 - GA, TC + TC / 2);
                        symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(1, 0, 0));
                        Projection3d plate = (Projection3d)symbolGeometryHelper.CreateBox(null, LP, WP, TP);
                        m_PhysicalAspect.Outputs["PLATE"] = plate;

                        pointCollection.Add(new Position(WP / 2, WT / 2 - GA - LP / 2 + WC / 2, TC));
                        pointCollection.Add(new Position(WP / 2 - LU, WT / 2 - GA - LP / 2 + WC / 2, TC));
                        pointCollection.Add(new Position(WP / 2 - LU, WT / 2 - GA - LP / 2 + WC / 2, 0));
                        pointCollection.Add(new Position(WP / 2 - TC, WT / 2 - GA - LP / 2 + WC / 2, 0));
                        pointCollection.Add(new Position(WP / 2 - TC, WT / 2 - GA - LP / 2 + WC / 2, -HC + 2 * TC));
                        pointCollection.Add(new Position(WP / 2 - LB, WT / 2 - GA - LP / 2 + WC / 2, -HC + 2 * TC));
                        pointCollection.Add(new Position(WP / 2 - LB, WT / 2 - GA - LP / 2 + WC / 2, -HC + TC));
                        pointCollection.Add(new Position(WP / 2, WT / 2 - GA - LP / 2 + WC / 2, -HC + TC));
                        pointCollection.Add(new Position(WP / 2, WT / 2 - GA - LP / 2 + WC / 2, TC));

                        projectionVector.Set(0, -WC, 0);
                        Projection3d clamp3 = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                        m_PhysicalAspect.Outputs["CLAMP"] = clamp3;

                        normal1 = new Position(0, WT / 2 - GA - LP / 2, -HC + TC + TC).Subtract(new Position(0, WT / 2 - GA - LP / 2, -HC + TC - 0.25 * TC));
                        symbolGeometryHelper = new SymbolGeometryHelper();
                        symbolGeometryHelper.ActivePosition = new Position(0, WT / 2 - GA - LP / 2, -HC + TC - 0.25 * TC);
                        symbolGeometryHelper.SetOrientation(normal1, new Vector(1, 0, 0));
                        Projection3d bottomBolt = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal1.Length);
                        m_PhysicalAspect.Outputs["BOT_BOLT"] = bottomBolt;
                    }
                    else if (flipClampValue == "yes")
                    {
                        symbolGeometryHelper.ActivePosition = new Position(0, -(WT / 2 - GA), TC + TC / 2);
                        symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                        Projection3d plate = (Projection3d)symbolGeometryHelper.CreateBox(null, LP, WP, TP);
                        m_PhysicalAspect.Outputs["PLATE"] = plate;

                        pointCollection.Add(new Position(WP / 2, -(WT / 2 - GA - LP / 2 + WC / 2), TC));
                        pointCollection.Add(new Position((WP / 2 - LU), -(WT / 2 - GA - LP / 2 + WC / 2), TC));
                        pointCollection.Add(new Position((WP / 2 - LU), -(WT / 2 - GA - LP / 2 + WC / 2), 0));
                        pointCollection.Add(new Position((WP / 2 - TC), -(WT / 2 - GA - LP / 2 + WC / 2), 0));
                        pointCollection.Add(new Position((WP / 2 - TC), -(WT / 2 - GA - LP / 2 + WC / 2), -HC + 2 * TC));
                        pointCollection.Add(new Position((WP / 2 - LB), -(WT / 2 - GA - LP / 2 + WC / 2), -HC + 2 * TC));
                        pointCollection.Add(new Position((WP / 2 - LB), -(WT / 2 - GA - LP / 2 + WC / 2), -HC + TC));
                        pointCollection.Add(new Position(WP / 2, -(WT / 2 - GA - LP / 2 + WC / 2), -HC + TC));
                        pointCollection.Add(new Position(WP / 2, -(WT / 2 - GA - LP / 2 + WC / 2), TC));

                        projectionVector.Set(0, WC, 0);
                        Projection3d clamp4 = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                        m_PhysicalAspect.Outputs["CLAMP"] = clamp4;

                        normal1 = new Position(0, -WT / 2 + GA + LP / 2, -HC + TC + TC).Subtract(new Position(0, -WT / 2 + GA + LP / 2, -HC + TC - 0.25 * TC));
                        symbolGeometryHelper = new SymbolGeometryHelper();
                        symbolGeometryHelper.ActivePosition = new Position(0, -WT / 2 + GA + LP / 2, -HC + TC - 0.25 * TC);
                        symbolGeometryHelper.SetOrientation(normal1, new Vector(1, 0, 0));
                        Projection3d bottomBolt = symbolGeometryHelper.CreateCylinder(null, (DO) / 2, normal1.Length);
                        m_PhysicalAspect.Outputs["BOT_BOLT"] = bottomBolt;
                    }
                }
            }
            catch
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Bline_Clamp_9_1249X."));
                    return;
                }
            }
        }
        #endregion

        //BOM Description.
        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)catalogPart.GetPropertyValue("IJUAHgrBlineFinish", "Finish");
                PropertyValueCodelist insideOutsideCodelistValue = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrBlineInsideOutside", "Inside_Outside");
                if (finishCodelist.PropValue < 0 || finishCodelist.PropValue > 8)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBlineFinish, "Bline Finish Code list value should be between 1 and 8"));
                if (insideOutsideCodelistValue.PropValue <= 0 || insideOutsideCodelistValue.PropValue > 2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidInside_Outside, "Bline Inside_Outside Code list value should be between 1 and 2"));

                string finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;
                string inOrOut = insideOutsideCodelistValue.PropertyInfo.CodeListInfo.GetCodelistItem(insideOutsideCodelistValue.PropValue).DisplayName;


                bomDescription = catalogPart.PartDescription + ",Installed " + inOrOut + ", Finish: " + finish;
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Bline_Clamp_9_1249X."));
                return "";

            }
        }

        #endregion
    }
}



