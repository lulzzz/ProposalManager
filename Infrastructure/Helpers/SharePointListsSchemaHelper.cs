// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;

namespace Infrastructure.Helpers
{
    public class SharePointListsSchemaHelper
    {

        public static string GetOpportunityJsonSchema(string displayName)
        {
            string json = @"
{
  'displayName': '" + displayName + @"',
  'columns': [
    {
      'name': 'ObjectIdentifier',
      'text': {},
      'indexed': true
    },
    {
      'name': 'CustomerName',
      'text': {},
      'indexed': true
    },
    {
      'name': 'OpportunityName',
      'text': {},
      'indexed': true
    },
    {
      'name': 'DealSize',
      'number': {}
    },
    {
      'name': 'AnnualRevenue',
      'number': {}
    },
    {
      'name': 'Industry',
      'text': {},
      'indexed': true
    },
    {
      'name': 'Region',
      'text': {},
      'indexed': true
    },
    {
      'name': 'OpportunityNotes',
      'text': {}
    },
    {
      'name': 'Margin',
      'number': {}
    },
    {
      'name': 'Rate',
      'number': {}
    },
    {
      'name': 'DebtRatio',
      'number': {}
    },
    {
      'name': 'Purpose',
      'text': {},
      'indexed': true
    },
    {
      'name': 'DisbursementSchedule',
      'text': {},
      'indexed': true
    },
    {
      'name': 'CollateralAmount',
      'number': {}
    },
    {
      'name': 'Guarantees',
      'text': {},
      'indexed': true
    },
    {
      'name': 'RiskRating',
      'number': {}
    },
    {
      'name': 'OpenedDate',
      'dateTime': {}
    },
    {
      'name': 'Status',
      'text': {},
      'indexed': true
    }
  ]
}";
            return json;
        }


        public static string GetNotificationsJsonSchema(string displayName)
        {
            string json = @"
{
  'displayName': '" + displayName + @"',
  'columns': [
    {
      'name': 'SentToMail',
      'text': {}
    },
    {
      'name': 'SentToName',
      'text': {}
    },
    {
      'name': 'SentDate',
      'dateTime': {}
    },
    {
      'name': 'SentFromMail',
      'text': {}
    },
    {
      'name': 'SentFromName',
      'text': {}
    },
    {
      'name': 'ReadDate',
      'dateTime': {}
    },
    {
      'name': 'MessageBody',
      'text': {}
    }
  ]
}
";
            return json;
        }

        public static string GetUserProfileJsonSchema(string displayName)
        {
            string json = @"
{
     'displayName': '" + displayName + @"',
    'columns': [
    {
      'name': 'ObjectIdentifier',
      'text': {},
      'indexed': true
    },
    {
      'name': 'UserMail',
      'text': {},
      'indexed': true
    },
    {
      'name': 'PhoneNumber',
      'text': {},
      'indexed': true
    },
    {
      'name': 'ProfileImageURL',
      'text': {},
      'indexed': true
    },
    {
      'name': 'UserRole',
      'text': {},
      'indexed': true
    }
  ]
}
";
            return json;
        }
    }

    public enum ListSchema
    {
        OpportunitiesListSchema,
        NotificationsListSchema,
        UsersListSchema
    }
}
