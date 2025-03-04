//-----------------------------------------------------------------------------
//      Copyright (C) 2010-12 Intergraph Corporation.  All rights reserved.
//
//      Component:  SectionCircularMap class is designed to return a cross section
//                  map based on section types that are circular.  
//
//      Author:  3XCalibur
//
//      History:
//      November 03, 2010       WR                  Created
//
//   October 11,2012     Alliagtors      CR-207934/DM-CP-220895(for v2013) 'pointMap' collection is filled appropriately.
//-----------------------------------------------------------------------------

using System.Collections.Generic;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure.CrossSectionMappings
{
    /// <summary>
    /// The SectionCircularMap class is designed to return a cross section edge map based on circular type shape cross sections.
    /// </summary>
    internal class SectionCircularMap : ICrossSectionMap
    {
        #region ICrossSectionMap Members

        /// <summary>
        /// Gets the cross section map.
        /// </summary>
        /// <param name="flipLeftAndRight">if set to <c>true</c> [is mirrored].</param>
        /// <param name="quadrant">The quadrant.</param>
        /// <returns>The cross section mapping of edges based on the quadrant and the flipLeftAndRight flag.</returns>
        public Dictionary<SectionFaceType, SectionFaceType> GetCrossSectionMap(bool flipLeftAndRight, int quadrant, Dictionary<int, int> pointMap)
        {
            Dictionary<SectionFaceType, SectionFaceType> mappedEdges = new Dictionary<SectionFaceType, SectionFaceType>();

            //      IT = Inner_Tube
            //      OT = Outer_Tube
            //
            //          o o o o
            //        o         o
            //       o   o o o    o
            //       o  o     oIT o OT
            //       o  o     o   o
            //       o   o o o    o
            //        o         o
            //          o o o o

            mappedEdges.Add(SectionFaceType.Outer_Tube, SectionFaceType.Outer_Tube);
            mappedEdges.Add(SectionFaceType.Inner_Tube, SectionFaceType.Inner_Tube);

            if (!flipLeftAndRight)
            {
                switch (quadrant)
                {
                    case 1:

                        for (int i = 1; i < 28; i++)
                        {
                            pointMap.Add(i, i);
                        }

                        break;
                    
                    case 2:

                        pointMap.Add(1, 3);
                        pointMap.Add(2, 3);
                        pointMap.Add(3, 8);
                        pointMap.Add(4, 8);
                        pointMap.Add(5, 5);          //pointMap.Add(5, 8);
                        pointMap.Add(6, 3);
                        pointMap.Add(7, 7);
                        pointMap.Add(8, 15);
                        pointMap.Add(9, 15);
                        pointMap.Add(10, 15);
                        pointMap.Add(11, 15);
                        pointMap.Add(12, 15);
                        pointMap.Add(13, 15);
                        pointMap.Add(14, 15);
                        pointMap.Add(15, 1);
                        pointMap.Add(16, 1);
                        pointMap.Add(17, 1);
                        pointMap.Add(18, 15);
                        pointMap.Add(19, 19);
                        pointMap.Add(20, 3);
                        pointMap.Add(21, 3);
                        pointMap.Add(22, 3);
                        pointMap.Add(23, 3);
                        pointMap.Add(24, 3);
                        pointMap.Add(25, 8);
                        pointMap.Add(26, 1);
                        pointMap.Add(27, 27);

                        break;
                    
                    case 3:

                        pointMap.Add(1, 8);
                        pointMap.Add(2, 8);
                        pointMap.Add(3, 15);
                        pointMap.Add(4, 15);
                        pointMap.Add(5, 5);          //pointMap.Add(5, 15);
                        pointMap.Add(6, 8);
                        pointMap.Add(7, 7);
                        pointMap.Add(8, 1);
                        pointMap.Add(9, 1);
                        pointMap.Add(10, 1);
                        pointMap.Add(11, 1);
                        pointMap.Add(12, 1);
                        pointMap.Add(13, 1);
                        pointMap.Add(14, 1);
                        pointMap.Add(15, 3);
                        pointMap.Add(16, 3);
                        pointMap.Add(17, 3);
                        pointMap.Add(18, 1);
                        pointMap.Add(19, 19);
                        pointMap.Add(20, 8);
                        pointMap.Add(21, 8);
                        pointMap.Add(22, 8);
                        pointMap.Add(23, 8);
                        pointMap.Add(24, 8);
                        pointMap.Add(25, 15);
                        pointMap.Add(26, 3);
                        pointMap.Add(27, 27);

                        break;
                    
                    case 4:
                        
                        pointMap.Add(1, 15);
                        pointMap.Add(2, 15);
                        pointMap.Add(3, 1);
                        pointMap.Add(4, 1);
                        pointMap.Add(5, 5);          //pointMap.Add(5, 1);
                        pointMap.Add(6, 15);
                        pointMap.Add(7, 7);
                        pointMap.Add(8, 3);
                        pointMap.Add(9, 3);
                        pointMap.Add(10, 3);
                        pointMap.Add(11, 3);
                        pointMap.Add(12, 3);
                        pointMap.Add(13, 3);
                        pointMap.Add(14, 3);
                        pointMap.Add(15, 8);
                        pointMap.Add(16, 8);
                        pointMap.Add(17, 8);
                        pointMap.Add(18, 3);
                        pointMap.Add(19, 19);
                        pointMap.Add(20, 15);
                        pointMap.Add(21, 15);
                        pointMap.Add(22, 15);
                        pointMap.Add(23, 15);
                        pointMap.Add(24, 15);
                        pointMap.Add(25, 1);
                        pointMap.Add(26, 8);
                        pointMap.Add(27, 27);

                        break;
                    
                    default:
                        throw new EdgeMappingException(quadrant);
                }
            }
            else
            {
                switch (quadrant)
                {
                    case 1:

                        pointMap.Add(1, 1);
                        pointMap.Add(2, 24);
                        pointMap.Add(3, 23);
                        pointMap.Add(4, 22);
                        pointMap.Add(5, 5);          //pointMap.Add(5, 21);
                        pointMap.Add(6, 20);
                        pointMap.Add(7, 19);
                        pointMap.Add(8, 18);
                        pointMap.Add(9, 17);
                        pointMap.Add(10, 16);
                        pointMap.Add(11, 15);
                        pointMap.Add(12, 14);
                        pointMap.Add(13, 13);
                        pointMap.Add(14, 12);
                        pointMap.Add(15, 11);
                        pointMap.Add(16, 10);
                        pointMap.Add(17, 9);
                        pointMap.Add(18, 8);
                        pointMap.Add(19, 7);
                        pointMap.Add(20, 6);
                        pointMap.Add(21, 5);
                        pointMap.Add(22, 4);
                        pointMap.Add(23, 3);
                        pointMap.Add(24, 2);
                        pointMap.Add(25, 26);
                        pointMap.Add(26, 25);
                        pointMap.Add(27, 27);

                        break;
                    
                    case 2:

                        pointMap.Add(1, 15);
                        pointMap.Add(2, 15);
                        pointMap.Add(3, 8);
                        pointMap.Add(4, 8);
                        pointMap.Add(5, 5);          //pointMap.Add(5, 8);
                        pointMap.Add(6, 15);
                        pointMap.Add(7, 7);
                        pointMap.Add(8, 3);
                        pointMap.Add(9, 3);
                        pointMap.Add(10, 3);
                        pointMap.Add(11, 3);
                        pointMap.Add(12, 3);
                        pointMap.Add(13, 3);
                        pointMap.Add(14, 3);
                        pointMap.Add(15, 1);
                        pointMap.Add(16, 1);
                        pointMap.Add(17, 1);
                        pointMap.Add(18, 1);
                        pointMap.Add(19, 19);
                        pointMap.Add(20, 15);
                        pointMap.Add(21, 15);
                        pointMap.Add(22, 15);
                        pointMap.Add(23, 15);
                        pointMap.Add(24, 15);
                        pointMap.Add(25, 8);
                        pointMap.Add(26, 1);
                        pointMap.Add(27, 27);

                        break;
                    
                    case 3:

                        pointMap.Add(1, 8);
                        pointMap.Add(2, 8);
                        pointMap.Add(3, 3);
                        pointMap.Add(4, 3);
                        pointMap.Add(5, 5);          //pointMap.Add(5, 3);
                        pointMap.Add(6, 8);
                        pointMap.Add(7, 7);
                        pointMap.Add(8, 1);
                        pointMap.Add(9, 1);
                        pointMap.Add(10, 1);
                        pointMap.Add(11, 1);
                        pointMap.Add(12, 1);
                        pointMap.Add(13, 1);
                        pointMap.Add(14, 1);
                        pointMap.Add(15, 15);
                        pointMap.Add(16, 15);
                        pointMap.Add(17, 15);
                        pointMap.Add(18, 1);
                        pointMap.Add(19, 19);
                        pointMap.Add(20, 8);
                        pointMap.Add(21, 8);
                        pointMap.Add(22, 8);
                        pointMap.Add(23, 8);
                        pointMap.Add(24, 8);
                        pointMap.Add(25, 3);
                        pointMap.Add(26, 15);
                        pointMap.Add(27, 27);

                        break;
                    
                    case 4:

                        pointMap.Add(1, 3);
                        pointMap.Add(2, 3);
                        pointMap.Add(3, 1);
                        pointMap.Add(4, 1);
                        pointMap.Add(5, 5);          //pointMap.Add(5, 1);
                        pointMap.Add(6, 3);
                        pointMap.Add(7, 7);
                        pointMap.Add(8, 15);
                        pointMap.Add(9, 15);
                        pointMap.Add(10, 15);
                        pointMap.Add(11, 15);
                        pointMap.Add(12, 15);
                        pointMap.Add(13, 15);
                        pointMap.Add(14, 15);
                        pointMap.Add(15, 8);
                        pointMap.Add(16, 8);
                        pointMap.Add(17, 8);
                        pointMap.Add(18, 15);
                        pointMap.Add(19, 19);
                        pointMap.Add(20, 3);
                        pointMap.Add(21, 3);
                        pointMap.Add(22, 3);
                        pointMap.Add(23, 3);
                        pointMap.Add(24, 3);
                        pointMap.Add(25, 1);
                        pointMap.Add(26, 8);
                        pointMap.Add(27, 27);

                        break;
                    
                    default:
                        throw new EdgeMappingException(quadrant);
                }
            }

            return mappedEdges;
        }

        #endregion
    }
}
