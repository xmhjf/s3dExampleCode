//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RULEGuideSecSize.cs
//   Frame_Assy,Ingr.SP3D.Content.Support.Rules.Rule_Offset1ByNPD
//   Author       :Hema
//   Creation Date:30.July.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  30.July.2013     Hema    CR-CP-224470 Convert HS_S3DFrame to C# .Net
//  06.May.2015      PVK     CR-CP-270669	Modify the exsisting .Net Frame_Assy to honor attributes which are not in OA
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    public class Rule_Offset1ByNPD : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;

            BoundingBoxHelper boundingBoxHelper = new BoundingBoxHelper(support);
            SupportedHelper supportedHelper = new SP3D.Support.Middle.SupportedHelper(support);

            Boolean includeInsulation = (Boolean)((PropertyValueBoolean)FrameAssemblyServices.GetPropertyValue(support,"IJUAhsIncludeInsulation", "IncludeInsulation")).PropValue;
            int routeIndex = boundingBoxHelper.GetBoundaryRouteIndex("BBFrame", BoundingBoxEdge.Left);
            PipeObjectInfo pipeInfo = (PipeObjectInfo)supportedHelper.SupportedObjectInfo(routeIndex);
             double[] pipeDiameter=new double[1];
            
            if (includeInsulation)
                 pipeDiameter[0] = pipeInfo.OutsideDiameter + pipeInfo.InsulationThickness;
             else
                 pipeDiameter[0] = pipeInfo.OutsideDiameter;
             return pipeDiameter;
        }
    }
}