//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   HALFEN_HZL_63_63.cs
//    Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_HZL_63_63
//   Author       :Sasidhar  
//   Creation Date:23-11-2012  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   23-11-2012    Sasidhar   CR-CP-222275 Converted VB HS_HALFEN_PC Project to C#.Net
//   20-03-2013      Vijay    DI-CP-228142  Modify the error handling for delivered H&S symbols
//   10-03-2015     Chethan   TR-CP-269406  Make HgrBeam a Non-Cached Symbol  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Support.Content.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.1.0.0")]
    [VariableOutputs]
    public class HALFEN_HZL_63_63 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_HZL_63_63"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Length", "Length", 0.999999)]
        public InputDouble m_dLength;
        [InputDouble(3, "Open_Ends", "Open_Ends", 1)]
        public InputDouble m_oOpen_Ends;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Port4", "Port4")]
        [SymbolOutput("Port5", "Port5")]
        [SymbolOutput("Port6", "Port6")]
        [SymbolOutput("Port7", "Port7")]
        [SymbolOutput("Port8", "Port8")]
        [SymbolOutput("BOX", "BOX")]
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
                SP3DConnection connection = default(SP3DConnection);
                connection = OccurrenceConnection;

                Double length = m_dLength.Value;
                Double height = 0.063;
                Double width = 0.063;

                if (length > 3)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrNLTLengthArguments, "Length cannot exceed 3M."));
                }
                //ports

                Port port1 = new Port(connection, part, "BeginTop", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(connection, part, "BeginCorner", new Position(0, -width / 2.0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                Port port3 = new Port(connection, part, "BeginEdge", new Position(-height / 2.0, -width / 2.0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;

                Port port4 = new Port(connection, part, "BeginMiddle", new Position(-height / 2.0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port4"] = port4;

                Port port5 = new Port(connection, part, "EndTop", new Position(0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port5"] = port5;

                Port port6 = new Port(connection, part, "EndCorner", new Position(0, -width / 2.0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port6"] = port6;

                Port port7 = new Port(connection, part, "EndEdge", new Position(-height / 2.0, -width / 2.0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port7"] = port7;

                Port port8 = new Port(connection, part, "EndMiddle", new Position(-height / 2.0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port8"] = port8;
                if (length <=0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrNGTEqualToZeroLengthArguments, "Length can not be zero or negative"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-width / 2, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d box = (Projection3d)symbolGeometryHelper.CreateBox(null, length, width, height);
                m_Symbolic.Outputs["BOX"] = box;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of HALFEN_HZL_63_63.cs."));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                String stockNumber;
                Double lengthValue;
                String length;
                String capBom;
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                PropertyValueCodelist openendsCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrOpenEnds", "Open_Ends");
                String openEnds = openendsCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(openendsCodelist.PropValue).DisplayName;

                if (Double.Parse(openEnds) > 0)
                {
                    capBom = " w/ " + openEnds + " x HPE 63/63 - LLDPE - 0318.000-00010";
                }
                else
                {
                    capBom = "";
                }
                lengthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrOccLength", "Length")).PropValue;
                length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_INCH);
                stockNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrStock_Number", "Stock_Number")).PropValue;
                bomString = part.PartDescription + " - fv - " + length + " - " + stockNumber + capBom;

                return bomString;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrBOMDescription, "Error in BOMDescription of HALFEN_HZL_63_63.cs."));
                }
                return "";
            }
        }
        #endregion

        #region "ICustomWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                Double weight, cogX, cogY, cogZ;
                Double length = (Double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrOccLength", "Length")).PropValue;
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                weight = (Double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                Double weightperUnitlength = (Double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrWeight_Per_MM", "Weight_Per_MM")).PropValue;

                if (weight == 1.6E+308)
                    weight = 0.0;
                else
                    weight = weightperUnitlength * length * 1000;
                
                cogX = 0;
                cogY = 0;
                cogZ = -length/2;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrWeightCG, "Error in WeightCG of HALFEN_HZL_63_63.cs."));
                }
            }
        }
        #endregion
    }

}
