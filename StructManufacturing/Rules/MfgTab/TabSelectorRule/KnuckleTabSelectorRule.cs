//----------------------------------------------------------------------------------
//      Copyright (C) 2015 Intergraph Corporation.  All rights reserved.
//
//      Purpose: Selector Rule for KnuckleTab. It provides the catalog part for given tab candidate.
//               
//
//      Author:  Soumya Kottoori
//
//      History:
//      August 24th, 2015   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using System;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Manufacturing.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Content.Manufacturing.Services;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.Generic;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle.Services;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Selector rule for knuckle tab.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [RuleInterface(ManufacturingTabSymbolConstants.tabname, ManufacturingTabSymbolConstants.interfaceName)]
    public class KnuckleTabSelectorRule : TabSelectorRule
    {
        /// <summary>
        /// Returns  the TabInformation with catalog part, driven entity geometry, driving entity geometry. 
        /// In case of tabs along edge, it returns also the TabAlongEdge information through the TabInformation.
        /// </summary>
        /// <param name="tabSelectorInfo">This class holds the tab corner geometries and other needed information.</param>
        public override ReadOnlyCollection<TabInformation> GetInputs(TabSelectorInformation tabSelectorInfo)
        {

            IList<TabInformation> tabInputs = new List<TabInformation>();
            foreach (TabCandidate selectionInfo in tabSelectorInfo.Tabcandidates)
            {
                if (selectionInfo.NumberOfKnuckles > 0)
                {
                    TopologyPort port = (TopologyPort)selectionInfo.FirstGeometry.RepresentedEntity;
                    PlatePart platePart = (PlatePart)port.Connectable;
                    if (platePart == null)
                    {
                        port = (TopologyPort)selectionInfo.SecondGeometry.RepresentedEntity;
                        platePart = (PlatePart)port.Connectable;
                    }
                    BusinessObject firstObject = (BusinessObject)selectionInfo.FirstGeometry.RepresentedEntity;
                    BusinessObject secondObject = (BusinessObject)selectionInfo.SecondGeometry.RepresentedEntity;

                    BusinessObject firstOperator = PortService.GetPortOperator(platePart, (TopologyPort)firstObject);

                    if (secondObject is PlateKnuckle && firstOperator is Seam)
                    {
                        ManufacturingGeometry drivenGeometry = selectionInfo.FirstGeometry;
                        ManufacturingGeometry drivingGeometry = selectionInfo.SecondGeometry;
                        CatalogBaseHelper catalogHelper = new CatalogBaseHelper();
                        string partClassName = "SMTabKnuckleType3";
                        BusinessObject partClassObject = catalogHelper.GetPartClass(partClassName);
                        ReadOnlyCollection<BusinessObject> parts = ((PartClass)partClassObject).Parts;
                        Part part = (Part)catalogHelper.GetPart("MfgKnuckleTabLength_150");
                        TabInformation tabInput = new TabInformation(drivenGeometry, drivingGeometry, part, selectionInfo.Location);
                        tabInputs.Add(tabInput);
                        break;
                    }
                }

            }
            ReadOnlyCollection<TabInformation> tabs = new ReadOnlyCollection<TabInformation>(tabInputs);
            return tabs;
        }
    }
}
