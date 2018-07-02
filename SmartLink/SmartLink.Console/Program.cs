// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartLink.Console
{
    class Program
    {
        static void Main(string[] args)
        {
            var test = new DbTest();
            test.TestInsertSourceCategory();
        }
    }
}
