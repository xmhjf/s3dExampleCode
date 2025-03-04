//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_916.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_916
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
    public class PSL_916 : HangerComponentSymbolDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_916"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(3, "Pipe_Dia", "Pipe_Dia", 0.999999)]
        public InputDouble Pipe_Dia;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("CYL", "CYL")]
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
                double actualPipeDiameter = 0;
                Double pipeDiameter = Pipe_Dia.Value;
                Double c = C.Value;

                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (metadataManager != null)
                    actualPipeDiameter = Convert.ToDouble(metadataManager.GetCodelistInfo("PSL_816_PipeDia", "UDP").GetCodelistItem((int)pipeDiameter).ShortDisplayName.Trim()) / 1000;
                else
                    actualPipeDiameter = 0.1143;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                if (c == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCNEZ, "C value cannot be zero"));
                    return;
                }
                if (actualPipeDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidActualPipeDiameterGTZ, "Pipe_Dia should be greater than zero"));
                    return;
                }
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d cylinder = symbolGeometryHelper.CreateCylinder(null, actualPipeDiameter / 2.0, c);
                Matrix4X4 matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -c / 2.0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                cylinder.Transform(matrix);
                m_Symbolic.Outputs["CYL"] = cylinder;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_916.cs"));
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
                double weight = 0;
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                string interfaceName = string.Empty;
                if (supportComponentBO.SupportsInterface("IJOAHgrPSL_916"))
                    interfaceName = "IJOAHgrPSL_916";
                else
                    interfaceName = "IJOAHgrPSL_917";

                int matGrade = (int)((PropertyValueCodelist)supportComponentBO.GetPropertyValue(interfaceName, "MAT_GRADE")).PropValue;
                if (matGrade < 1 || matGrade > 2)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidMatGradeCodelist1and2, "Material Grade codelist value should be 1 or 2."));
                    matGrade = 1;
                }
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                string materialGrade = metadataManager.GetCodelistInfo("PSL_916_917_GRADE", "UDP").GetCodelistItem(matGrade).ShortDisplayName;
                string size = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_SIZE", "SIZE")).PropValue;

                if (part.SupportsInterface("IJUAHgrPSL_916"))
                {
                    double C = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_916", "C")).PropValue;
                    weight = C * PSLSymbolServices.GetDataByCondition("PSL_COMLIN_WEIGHT_AUX", "IJUAHgrPSL_COMLIN_WEIGHT_AUX", materialGrade, "IJUAHgrPSL_COMLIN_WEIGHT_AUX", "Size", size);
                }
                else
                    weight = PSLSymbolServices.GetDataByCondition("PSL_COMLIN_WEIGHT_AUX", "IJUAHgrPSL_COMLIN_WEIGHT_AUX", materialGrade, "IJUAHgrPSL_COMLIN_WEIGHT_AUX", "Size", size);

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, 0, 0, 0);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of PSL_916.cs."));
            }
        }
        #endregion
    }
}
