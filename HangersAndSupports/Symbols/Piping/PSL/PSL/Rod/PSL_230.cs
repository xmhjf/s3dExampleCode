//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_230.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_230
//   Author       :  Vijay
//   Creation Date:  22-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   22-08-2013     Vijay    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net   
//   30-Dec-2014    PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report 
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
    public class PSL_230 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_230"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "R_T_O", "R_T_O", 0.999999)]
        public InputDouble R_T_O;
        [InputDouble(3, "Length", "Length", 0.999999)]
        public InputDouble Length;
        [InputDouble(4, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble ROD_DIA;
        [InputDouble(5, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(6, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(7, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(8, "D", "D", 0.999999)]
        public InputDouble D;
        [InputString(9, "PART_NUMBER", "PART_NUMBER", "No Value")]
        public InputString PART_NUMBER;
        [InputDouble(10, "END_THREAD", "END_THREAD", 0.999999)]
        public InputDouble END_THREAD;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("TOP_CYL_1", "TOP_CYL_1")]
        [SymbolOutput("TOP_CYL_2", "TOP_CYL_2")]
        [SymbolOutput("BOTTOM", "BOTTOM")]
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

                Double rodDiameter = ROD_DIA.Value;
                Double length = Length.Value;
                Double rto = R_T_O.Value;
                String partNumber = PART_NUMBER.Value;
                int threadCode = (int)END_THREAD.Value;
                String rightString = partNumber.Substring(partNumber.Length - 3);
                String[] arrayOfComponents = new String[8];
                Double threadLength = 0.0;
                Double a = A.Value;
                Double b = B.Value;
                Double c = C.Value;
                Double d = D.Value;

                //initialize array of componnents with column names in PSL_226_AUX table
                arrayOfComponents[1] = "VARIABLE_EFFORT_SUPPORT";
                arrayOfComponents[2] = "TURN_BUCKLE_NUT";
                arrayOfComponents[3] = "ROD_COUPLING_NUT";
                arrayOfComponents[4] = "CLEVIS_NUT";
                arrayOfComponents[5] = "WELDLESS_EYE_NUT";
                arrayOfComponents[6] = "WELDLESS_EYE";
                arrayOfComponents[7] = "TWO_NUTS";

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                if (!(threadCode == 8 || threadCode == 9))
                    threadLength = PSLSymbolServices.GetDataByCondition("PSL_226_AUX", "IJUAHgrPSL_226_AUX", arrayOfComponents[(int)threadCode], "IJUAHgrPSL_226_AUX", "PART_NUMBER", "F226" + rightString);

                //Warning only
                if (length < threadLength)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidThreadLength, "Overall length must exceed thread length"));

                Port port1 = new Port(OccurrenceConnection, part, "Eye", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "ExThdRH", new Position(0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidRodDiameterGTZ, "ROD_DIA should be greater than zero"));
                    return;
                }
                if (c <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCGTZ, "C value should be greater than zero"));
                    return;
                }

                double angle = Math.Atan((a / 2.0 - rodDiameter / 2.0) / (d + b - (a / 2.0) - 0.004)) * 180.0 / Math.PI;
                double y2 = (a / 2.0 + c / 2.0) * Math.Cos(angle * Math.PI / 180);
                double y1 = rodDiameter / 2.0 + c - (y2 * (c / 2.0) / (a / 2.0 + c / 2.0));
                double z1 = (rodDiameter / 2.0) + d - (Math.Sqrt(((a / 2.0 + c / 2.0) * (a / 2.0 + c / 2.0)) - (y2 * y2)) * (c / 2.0) / (a / 2.0 + c / 2.0));
                double z2 = rodDiameter / 2.0 - b + a / 2.0 + Math.Sqrt((a / 2.0 + c / 2.0) * (a / 2.0 + c / 2.0) - y2 * y2);

                symbolGeometryHelper.ActivePosition = new Position(0, 0, b);
                Ellipse3d curve = (Ellipse3d)symbolGeometryHelper.CreateEllipse(null, (rodDiameter / 2.0 + c), c, 2 * Math.PI);
                Projection3d top = new Projection3d(curve, new Vector(0, 0, 1), d, true);
                m_Symbolic.Outputs["TOP"] = top;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(y2, 0, z2 - rodDiameter / 2.0 + b).Subtract(new Position(y1, 0, z1 - rodDiameter / 2.0 + b));
                symbolGeometryHelper.ActivePosition = new Position(y1, 0, z1 - rodDiameter / 2.0 + b);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topCylinder1 = symbolGeometryHelper.CreateCylinder(null, c / 2.0, normal.Length);
                m_Symbolic.Outputs["TOP_CYL_1"] = topCylinder1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(-y2, 0, z2 - rodDiameter / 2.0 + b).Subtract(new Position(-y1, 0, z1 - rodDiameter / 2.0 + b));
                symbolGeometryHelper.ActivePosition = new Position(-y1, 0, z1 - rodDiameter / 2.0 + b);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topCylinder2 = symbolGeometryHelper.CreateCylinder(null, c / 2.0, normal.Length);
                m_Symbolic.Outputs["TOP_CYL_2"] = topCylinder2;

                matrix.Rotate((180.0 + angle) * Math.PI / 180.0, new Vector(0, 1, 0));
                matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1));
                matrix.Translate(new Vector(0, 0, rodDiameter / 2.0 - b + a / 2.0 - rodDiameter / 2.0 + b));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Revolution3d bend = new Revolution3d((new Circle3d(new Position(a / 2.0 + c / 2.0, 0, 0), new Vector(0, 0, 1), c / 2.0)), new Vector(0, -1, 0), new Position(0, 0, 0), (182.0 + (angle * 2.0)) * Math.PI / 180.0, true);
                bend.Transform(matrix);
                m_Symbolic.Outputs["BOTTOM"] = bend;

                matrix = new Matrix4X4();
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Translate(new Vector(0, 0, rto));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d rodCylinder = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2.0, length - rto);
                rodCylinder.Transform(matrix);
                m_Symbolic.Outputs["ROD"] = rodCylinder;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_230."));
                    return;
                }
            }
        }

        #endregion


        #region ICustomWeightCG Members

        public void EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double weight, cogX, cogY, cogZ;
                double rodDiametr = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrRod_Dia_mm", "ROD_DIA")).PropValue;
                double rto = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrPSL_230", "R_T_O")).PropValue;
                double length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccLength_mm", "Length")).PropValue;

                double weightEye = PSLSymbolServices.GetDataByCondition("PSL_239", "IJUAHgrPSL_239", "WEIGHT", "IJUAHgrRod_Dia_mm", "ROD_DIA", rodDiametr - 0.001, rodDiametr + 0.001);
                double weightPerLength = PSLSymbolServices.GetDataByCondition("PSL_226", "IJUAHgrPSL_226", "WEIGHT_PER_LENGTH", "IJUAHgrRod_Dia_mm", "ROD_DIA", rodDiametr - 0.001, rodDiametr + 0.001);
                weight = weightEye + (length - rto) * weightPerLength;
                cogX = 0;
                cogY = 0;
                cogZ = -length / 2;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of PSL_230.cs."));
            }
        }

        #endregion

        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                string finish = string.Empty;
                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrPSL_FINISH1", "FINISH");
                if ((finishCodelist.PropValue < 1 || finishCodelist.PropValue > 6) && (finishCodelist.PropValue != 11 || finishCodelist.PropValue != 12))
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFinish1Codelist, "FINISH codelist number should be 1 or 6 or 11 or 12"));
                    finish = "Self Colour";
                }
                else
                    finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).ShortDisplayName;

                string partNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_PART_NUMBER", "PART_NUMBER")).PropValue;
                double lengthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccLength_mm", "Length")).PropValue;
                double rodDiameterValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrRod_Dia_mm", "ROD_DIA")).PropValue;
                string length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_MILLIMETER);
                string rodDiameter = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, rodDiameterValue, UnitName.DISTANCE_MILLIMETER);

                //PSL F230M08 Eye Rod oa:  IJUAHgrOccLength_mm:: LENGTH
                bomDescription = "PSL " + partNumber + " " + rodDiameter + " Eye Rod, Fisish:" + finish + ", Length:" + length;
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of PSL_230.cs."));
                return "";
            }
        }

        #endregion
    }
}
