// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Infrastructure.Authorization.Requirements;


namespace Infrastructure.Authorization.Handlers
{
    public class PermissionHandler : IAuthorizationHandler
    {
        public Task HandleAsync(AuthorizationHandlerContext context)
        {
            var pendingRequirements = context.PendingRequirements.ToList();

            foreach (var requirement in pendingRequirements)
            {
                if (requirement is ReadPermission)
                {
                    if (IsOwner(context.User, context.Resource) ||
                        IsSponsor(context.User, context.Resource))
                    {
                        context.Succeed(requirement);
                    }
                }
                else if (requirement is EditPermission ||
                         requirement is DeletePermission)
                {
                    if (IsOwner(context.User, context.Resource))
                    {
                        context.Succeed(requirement);
                    }
                }
            }

            //TODO: Use the following if targeting a version of
            //.NET Framework older than 4.6:
            //      return Task.FromResult(0);
            return Task.CompletedTask;
        }

        private bool IsOwner(ClaimsPrincipal user, object resource)
        {
            // Code omitted for brevity

            return true;
        }

        private bool IsSponsor(ClaimsPrincipal user, object resource)
        {
            // Code omitted for brevity

            return true;
        }
    }
}
