// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using CommLendingWeb.Extensions;
using CommLendingWeb.Helpers;
using CommLendingWeb.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace CommLendingWeb.Controllers
{
	[Authorize]
	public class DocumentController : BaseController
    {
		private readonly string SiteId;
		private readonly string ProposalManagerApiUrl;
		public DocumentController(IConfiguration configuration, IGraphSdkHelper graphSdkHelper) :
			base(configuration, graphSdkHelper)
		{
			// Get from config
			var appOptions = new AppOptions();
			configuration.Bind("AppOptions", appOptions);
			ProposalManagerApiUrl = appOptions.ProposalManagerApiUrl;
			SiteId = appOptions.SiteId;
		}

		[HttpPost]
		public async Task<IActionResult> UpdateTask(string opportunityId, string documentData)
		{
			try
			{
				if(string.IsNullOrWhiteSpace(opportunityId))
				{
					return BadRequest($"{nameof(opportunityId)} is required");
				}

				if (string.IsNullOrWhiteSpace(documentData))
				{
					return BadRequest($"{nameof(documentData)} is required");
				}

				var uri = $"{ProposalManagerApiUrl}/api/Opportunity";
				var client = GetAuthorizedWebClient();
				var content = new StringContent(documentData, Encoding.UTF8, "application/json");
				var request = new HttpRequestMessage(new HttpMethod("PATCH"), uri);
				request.Content = content;

				var response = await client.SendAsync(request);

				if(!response.IsSuccessStatusCode)
				{
					return BadRequest(response.ReasonPhrase);
				}

				return Ok();
			}
			catch (Exception ex)
			{
				return BadRequest($"Error updating Opportunity: {ex.Message}");
			}
		}

		[HttpGet]
		public async Task<IActionResult> GetFormalProposal(string id)
		{
			try
			{
				if(string.IsNullOrWhiteSpace(id))
				{
					return BadRequest($"{nameof(id)} is required");
				}

				var client = GetAuthorizedWebClient();
				
				var opportunity = await client.GetStringAsync($"{ProposalManagerApiUrl}/api/Opportunity?name={id}");

				return Ok(opportunity);
			}
			catch (Exception ex)
			{
				return BadRequest(ex.Message);
			}
		}

		//[HttpGet]
		//public async Task<string> GetOOXml(string id)
		//{
		//	if (string.IsNullOrWhiteSpace(id))
		//	{
		//		throw new ArgumentNullException(nameof(id));
		//	}

		//	var graphClient = GraphHelper.GetAuthenticatedClient();
		//	var stream = await graphClient.Sites[SiteId].Lists[ListId].Items[id].DriveItem.Content.Request().GetAsync();

		//	using (var wordDocument = WordprocessingDocument.Open(stream, false))
		//	{
		//		return wordDocument.ToFlatOpcString();
		//	}
		//}

		[HttpGet]
		public async Task<IActionResult> List(string id)
		{
			//TODO: try to use graph client proxy entities
			// if not feasible then filter in odata query by displayName eq 'Documents' to reduce payload
			//var items = await graphClient.Sites[$"{SiteId}"].Sites[id].Lists.Request().Expand("Items").GetAsync();
			// Initialize the GraphServiceClient.
			try
			{
				if (string.IsNullOrWhiteSpace(id))
				{
					return BadRequest($"{nameof(id)} is required");
				}

				var graphClient = GraphHelper.GetAuthenticatedClient();
				var uri = $"https://graph.microsoft.com/v1.0/sites/{SiteId}:/sites/{id}:/lists?$expand=items";
				var request = new HttpRequestMessage(HttpMethod.Get, uri);

				await graphClient.AuthenticationProvider.AuthenticateRequestAsync(request);

				var response = await graphClient.HttpProvider.SendAsync(request);

				if (!response.IsSuccessStatusCode)
				{
					throw new Exception($"Error retrieving documents: {response.ReasonPhrase}");
				}

				var json = await response.Content.ReadAsStringAsync();

				dynamic items = JsonConvert.DeserializeObject(json);

				if(!items.value.HasValues)
				{
					return Ok(Enumerable.Empty<Document>());
				}

				var result = new List<Document>();

				foreach (var list in items.value)
				{
					if (list.displayName == "Documents")
					{
						foreach (var item in list.items)
						{
							var webUrl = item.webUrl.ToString();
							result.Add(
								new Document()
								{
									Id = item.id,
									WebUrl = item.webUrl,
									CreatedByUser = new User() { Id = item.createdBy.user.id, DisplayName = item.createdBy.user.displayName },
									LastModifiedByUser = new User() { Id = item.lastModifiedBy.user.id, DisplayName = item.lastModifiedBy.user.displayName },
									LastModifiedDateTime = item.lastModifiedDateTime,
									CreatedDateTime = item.createdDateTime,
									Type = webUrl.Substring(webUrl.LastIndexOf('.') + 1),
									Name = webUrl.Substring(webUrl.LastIndexOf('/') + 1)
								});
						}
					}
				}

				return Ok(result.OrderBy(x => x.Name));
			}
			catch (Exception ex)
			{
				return BadRequest(ex.Message);
			}
		}

		private HttpClient GetAuthorizedWebClient()
		{
			var client = new HttpClient();
			var token = GraphHelper.GetProposalManagerToken().GetAwaiter().GetResult();

			client.DefaultRequestHeaders.Accept.Clear();
			client.DefaultRequestHeaders.Accept.Add(
				new MediaTypeWithQualityHeaderValue("application/json"));
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

			return client;
		}
	}
}
