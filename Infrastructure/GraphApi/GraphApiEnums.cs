// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using ApplicationCore.Helpers;
using System;
using System.Collections.Generic;
using System.Text;

namespace Infrastructure.GraphApi
{
    class GraphApiEnums
    {
    }

    public class GraphClientContext : SmartEnum<GraphClientContext, int>
    {
        public static GraphClientContext User = new GraphClientContext(nameof(User), 0);
        public static GraphClientContext Application = new GraphClientContext(nameof(Application), 1);
        public static GraphClientContext OnBehalf = new GraphClientContext(nameof(OnBehalf), 2);

        protected GraphClientContext(string name, int value) : base(name, value)
        {
        }
    }
}
