using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Alexa.NET.Request;
using Alexa.NET.Response;
using Alexa.NET.Request.Type;
using Alexa.NET;
using System.Security.Claims;
using O365Skill.Service;

namespace O365Skill.Functions
{
    public static class AlexaFunction
    {
        [FunctionName("O365AlexaSkill")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            
            string json = await req.ReadAsStringAsync();
            var skillRequest = JsonConvert.DeserializeObject<SkillRequest>(json);

            var requestType = skillRequest.GetRequestType();

            AlexaService alexaService = new AlexaService(skillRequest);

            if (requestType == typeof(LaunchRequest)) return new OkObjectResult(await alexaService.WelcomeMessage());
            if (requestType == typeof(SessionEndedRequest))
            {
                SessionEndedRequest sessionReq = skillRequest.Request as SessionEndedRequest;

                //log.LogError(sessionReq.Reason.ToString());
                return new OkObjectResult(String.Empty);
            }

            if (requestType == typeof(IntentRequest))
            {
                var intentRequest = skillRequest.Request as IntentRequest;
                return new OkObjectResult(await alexaService.ManageIntent(intentRequest));
            }
            return new OkObjectResult("OK");
        }

        

    }
}
