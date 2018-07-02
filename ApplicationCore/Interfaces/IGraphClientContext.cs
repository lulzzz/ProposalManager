// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Graph;

namespace ApplicationCore.Interfaces
{
    /// <summary>
    /// Interface to abstract a generic graph client context.
    /// </summary>
    public interface IGraphClientContext
    {
        GraphServiceClient GraphClient { get; }
    }

    public interface IGraphClientUserContext : IGraphClientContext { }

    public interface IGraphClientAppContext : IGraphClientContext { }
}
