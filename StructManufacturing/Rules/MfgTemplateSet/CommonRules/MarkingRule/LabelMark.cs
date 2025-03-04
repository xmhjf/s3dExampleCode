//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Label mark rule creates label marks on Template.    
//
//      Author:  
//
//      History:
//      June 27th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using System;
using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Label mark Rule.
    /// </summary>
    public class LabelMark : MarkingRule
    {
        /// <summary>
        /// Creates label marks.
        /// </summary>
        /// <param name="markingInfo">The marking info.</param>
        public override void Evaluate(MarkingInformation markingInfo)
        {
            try
            {

                //////// NO IMPLEMENTATION /////////////
                throw new NotImplementedException();

            }
            catch (Exception e)
            {
                LogForToDoList(e, 3017, "Call to TemplateSet Label Mark Rule failed with the error" + e.Message);
            }

        }

    }
}
