// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

namespace ApplicationCore
{
    public static class CoreConstants
    {
        // TBD
    }

    public static class ErrorConstants
    {
        public static class Codes
        {
            public static string GeneralException = "generalException";

            public static string InvalidRequest = "invalidRequest";

            public static string ItemNotFound = "itemNotFound";

            public static string NotAllowed = "notAllowed";

            public static string Timeout = "timeout";

            public static string TooManyRedirects = "tooManyRedirects";
        }

        public static class Messages
        {
            public static string AuthenticationProviderMissing = "Authentication provider is required before sending a request.";

            public static string BaseUrlMissing = "Base URL cannot be null or empty.";

            public static string InvalidTypeForDateConverter = "DateConverter can only serialize objects of type Date.";

            public static string LocationHeaderNotSetOnRedirect = "Location header not present in redirection response.";

            public static string OverallTimeoutCannotBeSet = "Overall timeout cannot be set after the first request is sent.";

            public static string RequestTimedOut = "The request timed out.";

            public static string RequestUrlMissing = "Request URL is required to send a request.";

            public static string TooManyRedirectsFormatString = "More than {0} redirects encountered while sending the request.";

            public static string UnableToCreateInstanceOfTypeFormatString = "Unable to create an instance of type {0}.";

            public static string UnableToDeserializeDate = "Unable to deserialize the returned Date.";

            public static string UnexpectedExceptionOnSend = "An error occurred sending the request.";

            public static string UnexpectedExceptionResponse = "Unexpected exception returned from the service.";
        }
    }
}