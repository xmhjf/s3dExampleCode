//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_801.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_801
//   Author       :  Rajeswari
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013   Rajeswari  CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
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

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class PSL_801 : HangerComponentSymbolDefinition, ICustomWeightCG, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_801"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "GAGE", "GAGE", 0.999999)]
        public InputDouble GAGE;
        [InputDouble(3, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(4, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(5, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(6, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(7, "F", "F", 0.999999)]
        public InputDouble F;
        [InputDouble(8, "C", "C", 0.999999)]
        public InputDouble C;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BEND", "BEND")]
        [SymbolOutput("R", "R")]
        [SymbolOutput("L", "L")]
        [SymbolOutput("R2", "R2")]
        [SymbolOutput("L2", "L2")]
        [SymbolOutput("BASE", "BASE")]
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
                Part part = (Part)PartInput.Value;
                Double gage = GAGE.Value;
                Double b = B.Value;
                Double d = D.Value;
                Double e = E.Value;
                Double a = A.Value;
                Double f = F.Value;
                Double c = C.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -a / 2 - f), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (e <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEGTZ, "E value should be greater than zero"));
                    return;
                }
                if (f <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFGTZ, "F value should be greater than zero"));
                    return;
                }
                if (b <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBGTZ, "B value should be greater than zero"));
                    return;
                }
                if (d == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidDNEZ, "D value cannot be zero"));
                    return;
                }
                if (a == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidANEZ, "A value cannot be zero"));
                    return;
                }
                if (c <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCGTZ, "C value should be greater than zero."));
                    return;
                }
                Revolution3d bend = new Revolution3d((new Circle3d(new Position(c / 2, 0, 0), new Vector(0, 0, 1), c / 2 - a / 2)), new Vector(0, -1, 0), new Position(0, 0, 0), (Math.Atan(1) * 4) * 180 / 180, true);
                matrix.Translate(new Vector(0, gage, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                bend.Transform(matrix);
                m_Symbolic.Outputs["BEND"] = bend;

                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d r = symbolGeometryHelper.CreateCylinder(null, c / 2 - a / 2, a / 2);
                matrix.SetIdentity();
                matrix.Translate(new Vector(gage, c / 2, -a / 2));
                r.Transform(matrix);
                m_Symbolic.Outputs["R"] = r;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d l = symbolGeometryHelper.CreateCylinder(null, c / 2 - a / 2, a / 2);
                matrix.SetIdentity();
                matrix.Translate(new Vector(gage, -c / 2, -a / 2));
                l.Transform(matrix);
                m_Symbolic.Outputs["L"] = l;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d r2 = symbolGeometryHelper.CreateCylinder(null, b / 2, d);
                matrix.SetIdentity();
                matrix.Translate(new Vector(gage, c / 2, -d));
                r2.Transform(matrix);
                m_Symbolic.Outputs["R2"] = r2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d l2 = symbolGeometryHelper.CreateCylinder(null, b / 2, d);
                matrix.SetIdentity();
                matrix.Translate(new Vector(gage, -c / 2, -d));
                l2.Transform(matrix);
                m_Symbolic.Outputs["L2"] = l2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-e / 2 + gage, -(c + b + (c / 2 - a / 2) * 2) / 2, -a / 2);
                Projection3d baseBox = symbolGeometryHelper.CreateBox(null, e, c + b + (c / 2 - a / 2) * 2, f, 9);
                m_Symbolic.Outputs["BASE"] = baseBox;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_801.cs"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                string interfaceName = string.Empty;
                if (supportComponentBO.SupportsInterface("IJOAHgrPSL_801"))
                    interfaceName = "IJOAHgrPSL_801";
                else
                    interfaceName = "IJOAHgrPSL_802";

                int matGrade = ((int)((PropertyValueCodelist)supportComponentBO.GetPropertyValue(interfaceName, "MAT_GRADE")).PropValue);
                if (matGrade < 1 || matGrade > 2)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidMatGradeCodelist1and2, "Material Grade codelist value should be 1 or 2."));
                    matGrade = 1;
                }
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                string materialGrade = metadataManager.GetCodelistInfo("PSL_801_802_GRADE", "UDP").GetCodelistItem(matGrade).ShortDisplayName;
                string size = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_SIZE", "SIZE")).PropValue;

                double weight = PSLSymbolServices.GetDataByCondition("PSL_COMLIN_WEIGHT_AUX", "IJUAHgrPSL_COMLIN_WEIGHT_AUX", materialGrade, "IJUAHgrPSL_COMLIN_WEIGHT_AUX", "Size", size);

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, 0, 0, 0);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of PSL_801.cs."));
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                string partNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_SIZE", "SIZE")).PropValue;
                string interfaceName = string.Empty;
                if (oSupportOrComponent.SupportsInterface("IJOAHgrPSL_801"))
                    interfaceName = "IJOAHgrPSL_801";
                else
                    interfaceName = "IJOAHgrPSL_802";
                int matGrade = (int)((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue(interfaceName, "MAT_GRADE")).PropValue;
                if (matGrade < 1 || matGrade > 2)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidMatGradeCodelist1and2, "Material Grade codelist value should be 1 or 2."));
                    matGrade = 1;
                }
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                string materialGrade = metadataManager.GetCodelistInfo("PSL_801_802_GRADE", "UDP").GetCodelistItem(matGrade).ShortDisplayName;

                if (oSupportOrComponent.SupportsInterface("IJOAHgrPSL_801"))
                    bomDescription = "PSL " + partNumber + " Grip-Type U-Bolt for Stainless Steel & Galvanised Pipes, Material Grade: " + materialGrade;
                else
                    bomDescription = "PSL " + partNumber + " Grip-Type U-Bolt for Copper and Cupro-Nickel Pipes, Material Grade: " + materialGrade;

                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of PSL_801.cs."));
                return "";
            }
        }
        #endregion

    }
}
