//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_FPR.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_FPR
//   Author       :  Manikanth
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013     Manikanth    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
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
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    public class PSL_F496 : HangerComponentSymbolDefinition, ICustomWeightCG, ICustomHgrBOMDescription
    {

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "Length", "Length", 0.999999)]
        public InputDouble Length;
        [InputDouble(3, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(4, "F", "F", 0.999999)]
        public InputDouble F;
        [InputDouble(5, "G", "G", 0.999999)]
        public InputDouble G;
        [InputDouble(6, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(7, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(8, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(9, "D", "D", 0.999999)]
        public InputDouble D;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOPPLATE", "TOPPLATE")]
        [SymbolOutput("BASEPLATE", "BASEPLATE")]
        [SymbolOutput("MIDDLE_CYL", "MIDDLE_CYL")]
        [SymbolOutput("HOLE1", "HOLE1")]
        [SymbolOutput("HOLE2", "HOLE2")]
        [SymbolOutput("HOLE3", "HOLE3")]
        [SymbolOutput("HOLE4", "HOLE4")]
        public AspectDefinition m_Symbolic;
        #endregion

        #region "Construct Outputs"
        protected override void ConstructOutputs()
        {
            //----------------------------------------------------------------------------------
            //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_F496"
            //----------------------------------------------------------------------------------
            try
            {
                double length = Length.Value;
                double e = E.Value;
                double f = F.Value;
                double g = G.Value;
                double a = A.Value;
                double c = C.Value;
                double d = D.Value;

                if (length > 1)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidOperatingLoad, "Operating load is out of range"));

                Part part = (Part)PartInput.Value;
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (g <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidGGTZ, "G value should be greater than zero"));
                    return;
                }
                if (c <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCGTZ, "C value should be greater than zero"));
                    return;
                }
                if (a <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidAGTZ, "A value should be greater than zero"));
                    return;
                }               
                if (d <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidDGTZ, "D value should be greater than zero"));
                    return;
                }
                if (e <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEGTZ, "E value should be greater than zero"));
                    return;
                }
                symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 rotateMatrix = new Matrix4X4();
                rotateMatrix.Translate(new Vector(-g / 2, -g / 2, 0));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d topPlateBox = symbolGeometryHelper.CreateBox(null, g, g, c, 9);
                topPlateBox.Transform(rotateMatrix);
                m_Symbolic.Outputs["TOPPLATE"] = topPlateBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Translate(new Vector(-d / 2, -d / 2, length - c));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d basePlateBox = symbolGeometryHelper.CreateBox(null, d, d, c, 9);
                basePlateBox.Transform(rotateMatrix);
                m_Symbolic.Outputs["BASEPLATE"] = basePlateBox;


                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, 0, length - c).Subtract(new Position(0, 0, c));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, c);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d middleCyl = symbolGeometryHelper.CreateCylinder(null, a / 2, normal.Length);
                m_Symbolic.Outputs["MIDDLE_CYL"] = middleCyl;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(-f / 2, f / 2, c).Subtract(new Position(-f / 2, f / 2, 0));
                symbolGeometryHelper.ActivePosition = new Position(-f / 2, f / 2, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d hole1 = symbolGeometryHelper.CreateCylinder(null, e / 2, normal.Length);
                m_Symbolic.Outputs["HOLE1"] = hole1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(-f / 2, -f / 2, c).Subtract(new Position(-f / 2, -f / 2, 0));
                symbolGeometryHelper.ActivePosition = new Position(-f / 2, -f / 2, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d hole2 = symbolGeometryHelper.CreateCylinder(null, e / 2, normal.Length);
                m_Symbolic.Outputs["HOLE2"] = hole2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(f / 2, -f / 2, c).Subtract(new Position(f / 2, -f / 2, 0));
                symbolGeometryHelper.ActivePosition = new Position(f / 2, -f / 2, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d hole3 = symbolGeometryHelper.CreateCylinder(null, e / 2, normal.Length);
                m_Symbolic.Outputs["HOLE3"] = hole3;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(f / 2, f / 2, c).Subtract(new Position(f / 2, f / 2, 0));
                symbolGeometryHelper.ActivePosition = new Position(f / 2, f / 2, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d hole4 = symbolGeometryHelper.CreateCylinder(null, e / 2, normal.Length);
                m_Symbolic.Outputs["HOLE4"] = hole4;


            }
            catch (Exception)
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_F496.cs."));
                return;
            }
        }

        #endregion

        #region "ICustomHgrBOMDescription Members"
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string bomDescrition = "";
            try
            {
                Part part = (Part)SupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                string partNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_PART_NUMBER", "PART_NUMBER")).PropValue;
                double length = (double)((PropertyValueDouble)SupportOrComponent.GetPropertyValue("IJUAHgrOccLength_mm", "Length")).PropValue;
                bomDescrition = "PSL " + partNumber + "  Standard Pedestal for Rigid Supports, L=" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length, UnitName.DISTANCE_MILLIMETER);
                return bomDescrition;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of PSL_276B."));
                return "";
            }
        }
        #endregion

        #region "ICustomHgrWeightCG Members"
        public void EvaluateWeightCG(BusinessObject oBusinessObject)
        {
            try
            {
                Part catalogPart = (Part)oBusinessObject.GetRelationship("madeFrom", "part").TargetObjects[0];
                double length = (double)((PropertyValueDouble)oBusinessObject.GetPropertyValue("IJUAHgrOccLength_mm", "Length")).PropValue;
                double plateWeight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrPSL_F496", "MASS_PLATES")).PropValue;
                double weightPerUnitLength = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrPSL_F496", "MASS_TUBE")).PropValue;
                double plateThickness = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrPSL_F496", "B")).PropValue;
                double weight = (weightPerUnitLength * (length - 2 * plateThickness)) + plateWeight;
                double cogZ = -length / 2;
                double cogY = 0;
                double cogX = 0;
                SupportComponent supportComponent = (SupportComponent)oBusinessObject;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);

            }
            catch (Exception)
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of PSL_F495.cs."));
                return;
            }
        }
        #endregion

    }
}
