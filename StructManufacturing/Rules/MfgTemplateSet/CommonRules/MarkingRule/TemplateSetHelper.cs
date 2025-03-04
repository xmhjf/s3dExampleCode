//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Common helper class for Template marking functions.    
//
//      Author:  
//
//      History:
//      June 27th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using System;
using Ingr.SP3D.Common.Middle;


namespace Ingr.SP3D.Content.Manufacturing
{
    internal class TemplateSetHelper
    {

        internal static void GetDirectionNames(Vector tangentVector, Position positionOnCurve, bool centerCross, out string directionName, out string oppositeDirectionName)
        {
            directionName = "";
            oppositeDirectionName = "";

            if( (tangentVector.X >= 0.5) && (tangentVector.X > Math.Abs(tangentVector.Y)) && (tangentVector.X > Math.Abs(tangentVector.Z)))
            {
                directionName = "F"; 
                oppositeDirectionName = "A";
            }
            else if ( (tangentVector.Y >= 0.5) && (tangentVector.Y > Math.Abs(tangentVector.X)) && (tangentVector.Y > Math.Abs(tangentVector.Z)))
            {
                if (positionOnCurve.Y > 0.001)
                {
                    directionName = "O"; // Outer
                    oppositeDirectionName = "I";
                }
                else
                {
                    directionName = "I"; //Inner
                    oppositeDirectionName = "O";
                }

                if (centerCross == true)
                {
                    directionName = "P";
                    oppositeDirectionName = "S";
                }
            }
            else if ((tangentVector.Z >= 0.5) && (tangentVector.Z > Math.Abs(tangentVector.X)) && (tangentVector.Z > Math.Abs(tangentVector.Y)))
            {
                directionName = "U";
                oppositeDirectionName = "D";
            }

            if ((tangentVector.X < -0.5) && (tangentVector.X < tangentVector.Y) && (tangentVector.X < tangentVector.Z))
            {
                directionName = "A";
                oppositeDirectionName = "F";
            }
            else if ((tangentVector.Y < -0.5) && (tangentVector.Y < tangentVector.Z))                     
            {
                if (positionOnCurve.Y > 0.001)
                {
                    directionName = "I"; // Inner
                    oppositeDirectionName = "O";
                }
                else
                {
                    directionName = "O"; //Outer
                    oppositeDirectionName = "I";
                }
                if (centerCross == true)
                {
                    directionName = "S";
                    oppositeDirectionName = "P";
                }
            }
            else if ((tangentVector.Z < -0.5) && (tangentVector.Z < tangentVector.X) && (tangentVector.Z < tangentVector.Y))
            {
                directionName = "D";
                oppositeDirectionName = "U";
            }

        }

    }
}
