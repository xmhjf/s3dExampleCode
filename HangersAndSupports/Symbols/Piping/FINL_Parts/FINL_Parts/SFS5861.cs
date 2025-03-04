//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   SFS5861.cs
//    FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5861
//   Author       :   Rajeswari
//   Creation Date:  18-March-2013
//   Description:    CR-CP-222272-Initial Creation

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  18-March-2013 Rajeswari CR-CP-222272-Initial Creation
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
    public class SFS5861 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "FINL_Parts,Ingr.SP3D.Content.Support.Symbols.SFS5861"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(3, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(4, "TakeOut", "TakeOut", 0.999999)]
        public InputDouble m_dTakeOut;
        [InputDouble(5, "UserLength", "UserLength", 0.999999)]
        public InputDouble m_dUserLength;
        [InputDouble(6, "SteelSize", "SteelSize", 1)]
        public InputDouble m_oSteelSize;
        [InputString(7, "LoadClass", "LoadClass", "No Value")]
        public InputString m_oLoadClass;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        [SymbolOutput("Lug1", "Lug1")]
        [SymbolOutput("Lug2", "Lug2")]
        [SymbolOutput("Lug3", "Lug3")]
        [SymbolOutput("Lug4", "Lug4")]
        [SymbolOutput("Beam1", "Beam1")]
        [SymbolOutput("Beam2", "Beam2")]
        [SymbolOutput("Hole1", "Hole1")]
        [SymbolOutput("Hole2", "Hole2")]
        [SymbolOutput("Hole3", "Hole3")]
        [SymbolOutput("Hole4", "Hole4")]
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
                 
               

                Double H = m_dH.Value;
                Double D = m_dD.Value;
                Double TakeOut = m_dTakeOut.Value;
                Double length = m_dUserLength.Value;
                Double steelSizeValue = m_oSteelSize.Value;
                String loadClass = m_oLoadClass.Value;
                if (length <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidLengthGTZero, "Length should be greater than zero"));
                    return;
                }
                if (D <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidDGTZero, "D value should be greater than zero"));
                    return;
                }
                double height=0, depth=0,K=0;
                PropertyValueCodelist steelSizeCodelist = (PropertyValueCodelist)part.GetPropertyValue("IJUAFINL_SteelSize", "SteelSize");
                CodelistItem codelist;
                codelist = steelSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)steelSizeValue);
                if (codelist.Value < 1 || codelist.Value > 7)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidSteelSize, "SteelSize code list value should be between 1 and 7"));
                    return;
                }
                string steelSize = steelSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)steelSizeValue).ShortDisplayName;

                // Getting steel height and depth
                CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                PartClass auxTable = (PartClass)cataloghelper.GetPartClass("FINLCmpSrv_SFS5861");
                ReadOnlyCollection<BusinessObject> classItems = auxTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                Collection<PropertyValue> attributes = new Collection<PropertyValue>();
                foreach (BusinessObject classItem in classItems)
                {
                    ReadOnlyCollection<PropertyValue> properties = classItem.GetAllProperties();
                    if (classItem.GetPropertyValue("IJUAFINLCmpSrv_SFS5861", "Size").ToString() == steelSize.Trim())
                    {
                        height = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAFINLCmpSrv_SFS5861", "Height")).PropValue;
                        depth = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAFINLCmpSrv_SFS5861", "Depth")).PropValue;
                        K = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAFINLCmpSrv_SFS5861", "K"+loadClass)).PropValue;
                    }
                }
                if (K == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidKNZero, "SteelSize not compatible with Load Class"));
                    return;
                }
                if (K != 0 && K + 0.12 < length)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrInvalidKLTUserLenth, "Length has exceeded maximum allowable for SteelSize. Changing Cross Beam Length to 400 mm"));
                    length = 0.4;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Hole1", new Position(-length / 2 + 0.06, 0, -height - 0.01 - TakeOut), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                Port port3 = new Port(OccurrenceConnection, part, "Hole2", new Position(length / 2 - 0.06, 0, -height - 0.01 - TakeOut), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port3"] = port3;

                symbolGeometryHelper.ActivePosition = new Position(-length / 2 + 0.01, -0.05, 0);
                Projection3d lug1 = (Projection3d)symbolGeometryHelper.CreateBox(null, 0.1, 0.1, 0.01, 9);
                m_Symbolic.Outputs["Lug1"] = lug1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(length / 2 - 0.11, -0.05, 0);
                Projection3d lug2 = (Projection3d)symbolGeometryHelper.CreateBox(null, 0.1, 0.1, 0.01, 9);
                m_Symbolic.Outputs["Lug2"] = lug2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-length / 2 + 0.01, -0.05, -height - 0.01);
                Projection3d lug3 = (Projection3d)symbolGeometryHelper.CreateBox(null, 0.1, 0.1, 0.01, 9);
                m_Symbolic.Outputs["Lug3"] = lug3;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(length / 2 - 0.11, -0.05, -height - 0.01);
                Projection3d lug4 = (Projection3d)symbolGeometryHelper.CreateBox(null, 0.1, 0.1, 0.01, 9);
                m_Symbolic.Outputs["Lug4"] = lug4;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-length/2,-H/2.0-depth,-height);
                Projection3d beam1 = (Projection3d)symbolGeometryHelper.CreateBox(null, length,depth,height, 9);
                m_Symbolic.Outputs["Beam1"] = beam1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-length / 2.0, H / 2.0, -height);
                Projection3d beam2 = (Projection3d)symbolGeometryHelper.CreateBox(null, length, depth, height, 9);
                m_Symbolic.Outputs["Beam2"] = beam2;

                Vector v1 = new Vector(0, 0, 1);
                Matrix4X4 matrix = new Matrix4X4();

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0,0,0);
                symbolGeometryHelper.SetOrientation(v1,v1.GetOrthogonalVector());
                Projection3d hole1 = symbolGeometryHelper.CreateCylinder(null, D / 2, 0.01);
                matrix.Translate(new Vector(-length / 2 + 0.06, 0, 0));
                hole1.Transform(matrix);
                m_Symbolic.Outputs["Hole1"] = hole1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0,0,0);
                symbolGeometryHelper.SetOrientation(v1, v1.GetOrthogonalVector());
                Projection3d hole2 = symbolGeometryHelper.CreateCylinder(null, D / 2, 0.01);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(length / 2 - 0.06, 0, 0));
                hole2.Transform(matrix);
                m_Symbolic.Outputs["Hole2"] = hole2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0,0,0);
                symbolGeometryHelper.SetOrientation(v1, v1.GetOrthogonalVector());
                Projection3d hole3 = symbolGeometryHelper.CreateCylinder(null, D / 2, 0.01);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(-length / 2 + 0.06, 0, -height - 0.01));
                hole3.Transform(matrix);
                m_Symbolic.Outputs["Hole3"] = hole3;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0,0,0);
                symbolGeometryHelper.SetOrientation(v1, v1.GetOrthogonalVector());
                Projection3d hole4 = symbolGeometryHelper.CreateCylinder(null, D / 2, 0.01);
                matrix = new Matrix4X4();
                matrix.Translate(new Vector(length / 2 - 0.06, 0, -height - 0.01));
                hole4.Transform(matrix);
                m_Symbolic.Outputs["Hole4"] = hole4;
            }
            catch //General Unhandled exception 
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SFS5861.cs."));
                return;
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"
   
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                string loadClass = (string)((PropertyValueString)catalogPart.GetPropertyValue("IJUAFINL_LoadClass", "LoadClass")).PropValue;
                double lengthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAFINL_UserLength", "UserLength")).PropValue;
                string length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_MILLIMETER);
                string[] lengthArray = length.Split('.');
                bomDescription = "Heavy Cross Beam SFS 5861 - " + loadClass + " - " + lengthArray[0];
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FINL_PartsLocalizer.GetString(FINL_PartsSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of SFS5861.cs."));
                return ""; ;
            }
        }
        #endregion

    }

}
