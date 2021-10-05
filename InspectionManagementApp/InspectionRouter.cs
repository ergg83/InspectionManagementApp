using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using System.Threading.Tasks; 
using Xrm.Tools.WebAPI;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Xrm.Tools.WebAPI.Results; 
using System.Dynamic;
using Xrm.Tools.WebAPI.Requests; 
using System.Collections.Generic;

namespace InspectionManagementApp
{
    public static class InspectionRouter
    {
        [FunctionName("InspectionRouter")]
        public static void Run([TimerTrigger("0 0 0 * * *")]TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            CRMWebAPI api = GetCRMWebAPI(log).Result;
            dynamic whoami = api.ExecuteFunction("WhoAmI").Result;
            log.LogInformation($"UserID: {whoami.UserId}");

            var inspections = GetInspections(api).Result;
            var users = GetUsers(api).Result;

            Random rnd = new Random(); 
            int sequenceNumber = rnd.Next(10, 99);

            int currentUserIndex = 0;
            foreach (dynamic inspection in inspections.List)
            {
                log.LogInformation($"Routing inspection {inspection.contoso_name}");
                var inspectionResult = new CRMUpdateResult();
                //Your record assignment would like this. We will not assign records to different users in this lab
                // if (users.List.Count > (currentUserIndex))
                //{
                // dynamic currentUser = users.List[currentUserIndex];
                // inspectionResult = RouteInspection(api, inspection, currentUser.systemuserid.ToString(), sequenceNumber).Result;
                //currentUserIndex++;
                //}
                //We will instead assign inspections to the user you are currently logged in as
                inspectionResult = RouteInspection(api, inspection, whoami.UserId.ToString(), sequenceNumber).Result;
            }                
        }

        private static async Task<CRMWebAPI> GetCRMWebAPI(ILogger log)
        {
            var clientID = Environment.GetEnvironmentVariable("cdsclientid", EnvironmentVariableTarget.Process);
            var clientSecret = Environment.GetEnvironmentVariable("cdsclientsecret", EnvironmentVariableTarget.Process);
            var crmBaseUrl = Environment.GetEnvironmentVariable("cdsurl", EnvironmentVariableTarget.Process);
            var crmul = crmBaseUrl + "/api/data/v9.0/";
            
            AuthenticationParameters ap = await AuthenticationParameters.CreateFromUrlAsync(new Uri(crmul));
            var clientcred = new ClientCredential(clientID, clientSecret);

            var auth = ap.Authority.Replace("/oauth2/authorize", "");
            var authContext = new AuthenticationContext(auth);

            var authenticationResult = await authContext.AcquireTokenAsync(crmBaseUrl, clientcred);

            return new CRMWebAPI(crmul,authenticationResult.AccessToken);
        }

        private static Task<CRMGetListResult<ExpandoObject>> GetInspections(CRMWebAPI api)
        {

            var fetchxml = @"<fetch version=""1.0"" mapping=""logical"">
                              <entity name=""contoso_inspection"" >
                              <attribute name=""contoso_inspectionid"" />
                              <attribute name=""contoso_name"" />
                              <attribute name=""ownerid"" />
                              <attribute name=""contoso_inspectiontype"" />
                              <attribute name=""contoso_sequence"" />
                              <attribute name=""contoso_scheduleddate"" />
                              <attribute name=""statuscode"" />
                              <filter type=""and"" >
                              <condition value=""0"" operator=""eq"" attribute=""statecode"" />
                              <condition attribute=""statuscode"" operator=""in"" >
                                  <value>1</value>
                                  <value>100000000</value>
                              </condition>
                              </filter>
                              </entity>
                              </fetch>";

            var inspections = api.GetList("contoso_inspections", QueryOptions: new CRMGetListOptions() { FetchXml = fetchxml });

            return inspections;

        }

        private static Task<CRMGetListResult<ExpandoObject>> GetUsers(CRMWebAPI api)
        {
            var users = api.GetList<ExpandoObject>("systemusers");
            return users;
        }

        private static async Task<CRMUpdateResult> RouteInspection(CRMWebAPI api, dynamic inspection, string userId, int sequenceNumber)
        {
            dynamic updateObject = new ExpandoObject();
            ((IDictionary<string, object>)updateObject).Add("ownerid@odata.bind", "/systemusers(" + userId + ")");
            updateObject.contoso_sequence = sequenceNumber.ToString();
            return await api.Update("contoso_inspections", new Guid(inspection.contoso_inspectionid), updateObject);
        }        

    }//
}
