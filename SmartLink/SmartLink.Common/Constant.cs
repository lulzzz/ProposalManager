// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartLink.Common
{
    static public class Constant
    {
        public const int AZURETABLE_BATCH_COUNT = 100;
        public const string PUBLISH_QUEUE_NAME = "publishqueue";
        public const string PUBLISH_TABLE_NAME = "publishtable";
        public const string CHECK_TABLE_NAME = "checkdocumenttable";

        static public readonly string POINTTYPE_SOURCEPOINT = "Source Point";
        static public readonly string POINTTYPE_SOURCEPOINTHISTORY = "Source Point history";
        static public readonly string POINTTYPE_DESTINATIONPOINT = "Destination Point";
        static public readonly string POINTYTPE_DESTINATIONCATALOG = "Destination Catalog";
        static public readonly string POINTTYPE_DESTINATIONLIST = "Destination Points list";
        static public readonly string POINTTYPE_SOURCECATALOG = "Source Catalog";
        static public readonly string POINTTYPE_SOURCECATALOGLIST = "Source Catalog list";
        static public readonly string POINTTYPE_CLONE = "Clone Folder";

        static public readonly string ACTIONTYPE_GET = "Get";
        static public readonly string ACTIONTYPE_ADD  = "Add";
        static public readonly string ACTIONTYPE_EDIT = "Edit";
        static public readonly string ACTIONTYPE_DELETE = "Delete";
        static public readonly string ACTIONTYPE_PUBLISH = "Publish";
        static public readonly string ACTIONTYPE_CLONE = "Clone";
    }
}
