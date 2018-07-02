// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Text;

namespace ApplicationCore.Helpers
{
    public static class JsonErrorResponse
    {
        public static JObject BadRequest(string message = "", string requestId = "")
        {
            dynamic innerError = new JObject();
            innerError.requestId = requestId;
            innerError.date = DateTimeOffset.Now;

            dynamic error = new JObject();
            error.code = "BadRequest";
            error.message = message ?? "BadRequest";
            error.innerError = innerError;

            dynamic response = new JObject();
            response.error = error;

            return JObject.FromObject(response);
        }

        public static JObject BadRequest(string code = "", string message = "", string requestId = "")
        {
            dynamic innerError = new JObject();
            innerError.requestId = requestId;
            innerError.date = DateTimeOffset.Now;

            dynamic error = new JObject();
            error.code = code ?? "BadRequest";
            error.message = message ?? "BadRequest";
            error.innerError = innerError;

            dynamic response = new JObject();
            response.error = error;

            return JObject.FromObject(response);
        }
    }
}
