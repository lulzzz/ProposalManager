// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using ApplicationCore.Helpers;
using Newtonsoft.Json;

namespace ApplicationCore
{
    public class ActionStatus : SmartEnum<ActionStatus, int>
    {
        public static ActionStatus NotStarted = new ActionStatus(nameof(NotStarted), 0);
        public static ActionStatus InProgress = new ActionStatus(nameof(InProgress), 1);
        public static ActionStatus Blocked = new ActionStatus(nameof(Blocked), 2);
        public static ActionStatus Completed = new ActionStatus(nameof(Completed), 3);

        [JsonConstructor]
        protected ActionStatus(string name, int value) : base(name, value)
        {
        }
    }

    public class ContentType : SmartEnum<ContentType, int>
    {
        public static ContentType NoneEmpty = new ContentType(nameof(NoneEmpty), 0);
        public static ContentType Opportunity = new ContentType(nameof(Opportunity), 1);
        public static ContentType Workflow = new ContentType(nameof(Workflow), 2);
        public static ContentType Document = new ContentType(nameof(Document), 3);
        public static ContentType ProposalDocument = new ContentType(nameof(ProposalDocument), 4);

        [JsonConstructor]
        protected ContentType(string name, int value) : base(name, value)
        {
        }
    }

    public class DocumentContext : SmartEnum<DocumentContext, int>
    {
        public static DocumentContext Attachment = new DocumentContext("Attachment", 0);
        public static DocumentContext ProposalTemplate = new DocumentContext("ProposalTemplate", 1);
        public static DocumentContext ChecklistDocument = new DocumentContext("ChecklistDocument", 2);

        [JsonConstructor]
        protected DocumentContext(string name, int value) : base(name, value)
        {
        }
    }

    public class OpportunityChannel : SmartEnum<OpportunityChannel, int>
    {
        public static OpportunityChannel General = new OpportunityChannel("General", 0);
        public static OpportunityChannel RiskAssessment = new OpportunityChannel("Risk Assessment", 1);
        public static OpportunityChannel CreditCheck = new OpportunityChannel("Credit Check", 2);
        public static OpportunityChannel Compliance = new OpportunityChannel("Compliance", 3);
        public static OpportunityChannel FormalProposal = new OpportunityChannel("Formal Proposal", 4);
        public static OpportunityChannel CustomerDecision = new OpportunityChannel("Customer Decision", 5);

        [JsonConstructor]
        protected OpportunityChannel(string name, int value) : base(name, value)
        {
        }
    }

    public class StatusCodes : SmartEnum<StatusCodes, int>
    {
        public static StatusCodes Status100Continue = new StatusCodes(nameof(Status100Continue), 100);
        public static StatusCodes Status101SwitchingProtocols = new StatusCodes(nameof(Status101SwitchingProtocols), 101);
        public static StatusCodes Status102Processing = new StatusCodes(nameof(Status102Processing), 102);

        public static StatusCodes Status200OK = new StatusCodes(nameof(Status200OK), 200);
        public static StatusCodes Status201Created = new StatusCodes(nameof(Status201Created), 201);
        public static StatusCodes Status202Accepted = new StatusCodes(nameof(Status202Accepted), 202);
        public static StatusCodes Status203NonAuthoritative = new StatusCodes(nameof(Status203NonAuthoritative), 203);
        public static StatusCodes Status204NoContent = new StatusCodes(nameof(Status204NoContent), 204);
        public static StatusCodes Status205ResetContent = new StatusCodes(nameof(Status205ResetContent), 205);
        public static StatusCodes Status206PartialContent = new StatusCodes(nameof(Status206PartialContent), 206);
        public static StatusCodes Status207MultiStatus = new StatusCodes(nameof(Status207MultiStatus), 207);
        public static StatusCodes Status208AlreadyReported = new StatusCodes(nameof(Status208AlreadyReported), 208);

        public static StatusCodes Status226IMUsed = new StatusCodes(nameof(Status226IMUsed), 226);

        public static StatusCodes Status300MultipleChoices = new StatusCodes(nameof(Status300MultipleChoices), 300);
        public static StatusCodes Status301MovedPermanently = new StatusCodes(nameof(Status301MovedPermanently), 301);
        public static StatusCodes Status302Found = new StatusCodes(nameof(Status302Found), 302);
        public static StatusCodes Status303SeeOther = new StatusCodes(nameof(Status303SeeOther), 303);
        public static StatusCodes Status304NotModified = new StatusCodes(nameof(Status304NotModified), 304);
        public static StatusCodes Status305UseProxy = new StatusCodes(nameof(Status305UseProxy), 305);
        public static StatusCodes Status306SwitchProxy = new StatusCodes(nameof(Status306SwitchProxy), 306);
        public static StatusCodes Status307TemporaryRedirect = new StatusCodes(nameof(Status307TemporaryRedirect), 307);
        public static StatusCodes Status308PermanentRedirect = new StatusCodes(nameof(Status308PermanentRedirect), 308);

        public static StatusCodes Status400BadRequest = new StatusCodes(nameof(Status400BadRequest), 400);
        public static StatusCodes Status401Unauthorized = new StatusCodes(nameof(Status401Unauthorized), 401);
        public static StatusCodes Status402PaymentRequired = new StatusCodes(nameof(Status402PaymentRequired), 402);
        public static StatusCodes Status403Forbidden = new StatusCodes(nameof(Status403Forbidden), 403);
        public static StatusCodes Status404NotFound = new StatusCodes(nameof(Status404NotFound), 404);
        public static StatusCodes Status405MethodNotAllowed = new StatusCodes(nameof(Status405MethodNotAllowed), 405);
        public static StatusCodes Status406NotAcceptable = new StatusCodes(nameof(Status406NotAcceptable), 406);
        public static StatusCodes Status407ProxyAuthenticationRequired = new StatusCodes(nameof(Status407ProxyAuthenticationRequired), 407);
        public static StatusCodes Status408RequestTimeout = new StatusCodes(nameof(Status408RequestTimeout), 408);
        public static StatusCodes Status409Conflict = new StatusCodes(nameof(Status409Conflict), 409);
        public static StatusCodes Status410Gone = new StatusCodes(nameof(Status410Gone), 410);
        public static StatusCodes Status411LengthRequired = new StatusCodes(nameof(Status411LengthRequired), 411);
        public static StatusCodes Status412PreconditionFailed = new StatusCodes(nameof(Status412PreconditionFailed), 412);
        public static StatusCodes Status413PayloadTooLarge = new StatusCodes(nameof(Status413PayloadTooLarge), 413);
        public static StatusCodes Status414UriTooLong = new StatusCodes(nameof(Status414UriTooLong), 414);
        public static StatusCodes Status415UnsupportedMediaType = new StatusCodes(nameof(Status415UnsupportedMediaType), 415);
        public static StatusCodes Status416RangeNotSatisfiable = new StatusCodes(nameof(Status416RangeNotSatisfiable), 416);
        public static StatusCodes Status417ExpectationFailed = new StatusCodes(nameof(Status417ExpectationFailed), 417);
        public static StatusCodes Status418ImATeapot = new StatusCodes(nameof(Status418ImATeapot), 418);
        public static StatusCodes Status419AuthenticationTimeout = new StatusCodes(nameof(Status419AuthenticationTimeout), 419);

        public static StatusCodes Status421MisdirectedRequest = new StatusCodes(nameof(Status421MisdirectedRequest), 421);
        public static StatusCodes Status422UnprocessableEntity = new StatusCodes(nameof(Status422UnprocessableEntity), 422);
        public static StatusCodes Status423Locked = new StatusCodes(nameof(Status423Locked), 423);
        public static StatusCodes Status424FailedDependency = new StatusCodes(nameof(Status424FailedDependency), 424);

        public static StatusCodes Status426UpgradeRequired = new StatusCodes(nameof(Status426UpgradeRequired), 426);

        public static StatusCodes Status428PreconditionRequired = new StatusCodes(nameof(Status428PreconditionRequired), 428);
        public static StatusCodes Status429TooManyRequests = new StatusCodes(nameof(Status429TooManyRequests), 429);

        public static StatusCodes Status431RequestHeaderFieldsTooLarge = new StatusCodes(nameof(Status431RequestHeaderFieldsTooLarge), 431);

        public static StatusCodes Status451UnavailableForLegalReasons = new StatusCodes(nameof(Status451UnavailableForLegalReasons), 451);

        public static StatusCodes Status500InternalServerError = new StatusCodes(nameof(Status500InternalServerError), 500);
        public static StatusCodes Status501NotImplemented = new StatusCodes(nameof(Status501NotImplemented), 501);
        public static StatusCodes Status502BadGateway = new StatusCodes(nameof(Status502BadGateway), 502);
        public static StatusCodes Status503ServiceUnavailable = new StatusCodes(nameof(Status503ServiceUnavailable), 503);
        public static StatusCodes Status504GatewayTimeout = new StatusCodes(nameof(Status504GatewayTimeout), 504);
        public static StatusCodes Status505HttpVersionNotsupported = new StatusCodes(nameof(Status505HttpVersionNotsupported), 505);
        public static StatusCodes Status506VariantAlsoNegotiates = new StatusCodes(nameof(Status506VariantAlsoNegotiates), 506);
        public static StatusCodes Status507InsufficientStorage = new StatusCodes(nameof(Status507InsufficientStorage), 507);
        public static StatusCodes Status508LoopDetected = new StatusCodes(nameof(Status508LoopDetected), 508);

        public static StatusCodes Status510NotExtended = new StatusCodes(nameof(Status510NotExtended), 510);
        public static StatusCodes Status511NetworkAuthenticationRequired = new StatusCodes(nameof(Status511NetworkAuthenticationRequired), 511);

        protected StatusCodes(string name, int value) : base(name, value)
        {
        }
    }
}
