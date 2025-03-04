//-----------------------------------------------------------------------------
//      Copyright (C) 2010-12 Intergraph Corporation.  All rights reserved.
//
//      Component:  SectionRectangularMap class is designed to return a cross section
//                  map based on section types that are rectangular.  
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
    /// The SectionRectangularMap class is designed to return a cross section edge map based on rectangular type shape cross sections.
    /// </summary>
    internal class SectionRectangularMap : ICrossSectionMap
    {
        #region ICrossSectionMap Members

        /// <summary>
        /// Gets the cross section map.
        /// </summary>
        /// <param name="flipLeftAndRight">if set to <c>true</c> [is mirrored].</param>
        /// <param name="quadrant">The quadrant.</param>
        /// <returns>The cross section mapping of edges based on the quadrant and the flipLeftAndRight flag.</returns>
        /// <exception cref="S3DEndCutMappingException">Unable to map cross section.</exception>
        public Dictionary<SectionFaceType, SectionFaceType> GetCrossSectionMap(bool flipLeftAndRight, int quadrant, Dictionary<int, int> pointMap)
        {
            Dictionary<SectionFaceType, SectionFaceType> mappedEdges = new Dictionary<SectionFaceType, SectionFaceType>();

            //      WR = Web_Right  
            //      Top
            //      WL = Web_Left            
            //      Bottom
            //
            //                      Top
            //           *************************
            //           *                       * 
            //           *    ***************    *
            //           *    *             *    *
            //           *    *             *    *
            //       WL  *    *             *    *  WR
            //           *    *             *    *
            //           *    *             *    *
            //           *    ***************    *
            //           *                       *  
            //           *************************
            //                    Bottom
            //
            // It was decided not to map the inner edges for this cross section.  They can easily be mapped if desired as 
            // InnerWebRight, InnerWebTop, InnerWebLeft, and InnerWebBottom.

            if (!flipLeftAndRight)
            {
                // Not mirrored case
                switch (quadrant)
                {
                    case 1:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Left);                        
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Bottom);

                        for (int i = 1; i < 28; i++)
                        {
                            pointMap.Add(i, i);
                        }                        
                        
                        break;
                    
                    case 2:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Web_Left);
                        
                        pointMap.Add(1, 25);
                        pointMap.Add(2, 25);
                        pointMap.Add(3, 9);
                        pointMap.Add(4, 9);
                        pointMap.Add(5, 9);
                        pointMap.Add(6, 25);
                        pointMap.Add(7, 7);
                        pointMap.Add(8, 26);
                        pointMap.Add(9, 15);
                        pointMap.Add(10, 15);
                        pointMap.Add(11, 15);
                        pointMap.Add(12, 26);
                        pointMap.Add(13, 26);
                        pointMap.Add(14, 26);
                        pointMap.Add(15, 21);
                        pointMap.Add(16, 21);
                        pointMap.Add(17, 21);
                        pointMap.Add(18, 26);
                        pointMap.Add(19, 19);
                        pointMap.Add(20, 25);
                        pointMap.Add(21, 3);
                        pointMap.Add(22, 3);
                        pointMap.Add(23, 3);
                        pointMap.Add(24, 25);
                        pointMap.Add(25, 8);
                        pointMap.Add(26, 1);
                        pointMap.Add(27, 27);

                        break;
                    
                    case 3:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top);
                        
                        pointMap.Add(1, 8);
                        pointMap.Add(2, 8);
                        pointMap.Add(3, 15);
                        pointMap.Add(4, 15);
                        pointMap.Add(5, 15);
                        pointMap.Add(6, 8);
                        pointMap.Add(7, 7);
                        pointMap.Add(8, 1);
                        pointMap.Add(9, 21);
                        pointMap.Add(10, 21);
                        pointMap.Add(11, 21);
                        pointMap.Add(12, 1);
                        pointMap.Add(13, 1);
                        pointMap.Add(14, 1);
                        pointMap.Add(15, 3);
                        pointMap.Add(16, 3);
                        pointMap.Add(17, 3);
                        pointMap.Add(18, 1);
                        pointMap.Add(19, 19);
                        pointMap.Add(20, 8);
                        pointMap.Add(21, 9);
                        pointMap.Add(22, 9);
                        pointMap.Add(23, 9);
                        pointMap.Add(24, 8);
                        pointMap.Add(25, 26);
                        pointMap.Add(26, 25);
                        pointMap.Add(27, 27);

                        break;

                    case 4:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Web_Right);
                        
                        pointMap.Add(1, 26);
                        pointMap.Add(2, 26);
                        pointMap.Add(3, 21);
                        pointMap.Add(4, 21);
                        pointMap.Add(5, 21);
                        pointMap.Add(6, 26);
                        pointMap.Add(7, 7);
                        pointMap.Add(8, 25);
                        pointMap.Add(9, 3);
                        pointMap.Add(10, 3);
                        pointMap.Add(11, 3);
                        pointMap.Add(12, 25);
                        pointMap.Add(13, 25);
                        pointMap.Add(14, 25);
                        pointMap.Add(15, 9);
                        pointMap.Add(16, 9);
                        pointMap.Add(17, 9);
                        pointMap.Add(18, 25);
                        pointMap.Add(19, 19);
                        pointMap.Add(20, 26);
                        pointMap.Add(21, 15);
                        pointMap.Add(22, 15);
                        pointMap.Add(23, 15);
                        pointMap.Add(24, 26);
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
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Bottom);                        
                        
                        pointMap.Add(1, 1);
                        pointMap.Add(2, 24);
                        pointMap.Add(3, 23);
                        pointMap.Add(4, 22);
                        pointMap.Add(5, 21);
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
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Web_Right);
                        
                        pointMap.Add(1, 26);
                        pointMap.Add(2, 26);
                        pointMap.Add(3, 15);
                        pointMap.Add(4, 15);
                        pointMap.Add(5, 15);
                        pointMap.Add(6, 26);
                        pointMap.Add(7, 7);
                        pointMap.Add(8, 25);
                        pointMap.Add(9, 9);
                        pointMap.Add(10, 9);
                        pointMap.Add(11, 9);
                        pointMap.Add(12, 25);
                        pointMap.Add(13, 25);
                        pointMap.Add(14, 25);
                        pointMap.Add(15, 3);
                        pointMap.Add(16, 3);
                        pointMap.Add(17, 3);
                        pointMap.Add(18, 25);
                        pointMap.Add(19, 19);
                        pointMap.Add(20, 26);
                        pointMap.Add(21, 21);
                        pointMap.Add(22, 21);
                        pointMap.Add(23, 21);
                        pointMap.Add(24, 26);
                        pointMap.Add(25, 8);
                        pointMap.Add(26, 1);
                        pointMap.Add(27, 27);

                        break;
                   
                    case 3:
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Web_Right);                       
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Bottom);                       
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Web_Left);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Top);
                        
                        pointMap.Add(1, 8);
                        pointMap.Add(2, 8);
                        pointMap.Add(3, 9);
                        pointMap.Add(4, 9);
                        pointMap.Add(5, 9);
                        pointMap.Add(6, 8);
                        pointMap.Add(7, 7);
                        pointMap.Add(8, 1);
                        pointMap.Add(9, 3);
                        pointMap.Add(10, 3);
                        pointMap.Add(11, 3);
                        pointMap.Add(12, 1);
                        pointMap.Add(13, 1);
                        pointMap.Add(14, 1);
                        pointMap.Add(15, 21);
                        pointMap.Add(16, 21);
                        pointMap.Add(17, 21);
                        pointMap.Add(18, 1);
                        pointMap.Add(19, 19);
                        pointMap.Add(20, 8);
                        pointMap.Add(21, 15);
                        pointMap.Add(22, 15);
                        pointMap.Add(23, 15);
                        pointMap.Add(24, 8);
                        pointMap.Add(25, 25);
                        pointMap.Add(26, 26);
                        pointMap.Add(27, 27);

                        break;
                    
                    case 4:    
                        mappedEdges.Add(SectionFaceType.Web_Right, SectionFaceType.Top);
                        mappedEdges.Add(SectionFaceType.Top, SectionFaceType.Web_Right);
                        mappedEdges.Add(SectionFaceType.Web_Left, SectionFaceType.Bottom);
                        mappedEdges.Add(SectionFaceType.Bottom, SectionFaceType.Web_Left);
                        
                        pointMap.Add(1, 25);
                        pointMap.Add(2, 25);
                        pointMap.Add(3, 3);
                        pointMap.Add(4, 3);
                        pointMap.Add(5, 3);
                        pointMap.Add(6, 25);
                        pointMap.Add(7, 7);
                        pointMap.Add(8, 26);
                        pointMap.Add(9, 21);
                        pointMap.Add(10, 21);
                        pointMap.Add(11, 21);
                        pointMap.Add(12, 26);
                        pointMap.Add(13, 26);
                        pointMap.Add(14, 26);
                        pointMap.Add(15, 15);
                        pointMap.Add(16, 15);
                        pointMap.Add(17, 15);
                        pointMap.Add(18, 26);
                        pointMap.Add(19, 19);
                        pointMap.Add(20, 25);
                        pointMap.Add(21, 21);
                        pointMap.Add(22, 21);
                        pointMap.Add(23, 21);
                        pointMap.Add(24, 25);
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
