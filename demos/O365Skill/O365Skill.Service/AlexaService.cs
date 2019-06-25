using System;
using System.Threading.Tasks;
using O365Skill.Graph;
using Alexa.NET.Request;
using Alexa.NET.Response;
using Alexa.NET.Request.Type;
using Alexa.NET;

namespace O365Skill.Service
{
    public class AlexaService
    {
        private SkillRequest _request;
        private SkillResponse _response;
        public AlexaService(SkillRequest request)
        {
            _request = request;
        }

        public async Task<SkillResponse> ManageIntent(IntentRequest intentRequest)
        {
            switch (intentRequest.Intent.Name.ToUpper())
            {
                case "LEER_DOCUMENTO":
                    return await ReadDocument(intentRequest);
                case "AVISAR_MANAGER":
                    return await WarnManager(intentRequest);
                case "RESERVAR_CITA":
                    return await ScheduleMeeting(intentRequest);
                case "AMAZON.CANCELINTENT":
                    _response = _response = ResponseBuilder.Tell("OK, lo que tu digas!");
                    _response.Response.ShouldEndSession = false;
                    return _response;
                case "AMAZON.HELPINTENT":
                    _response = ResponseBuilder.Tell("Skill de demo del SharePoint Saturday Madrid");
                    _response.Response.ShouldEndSession = false;
                    return _response;
                case "AMAZON.STOPINTENT":
                    return await SayGoodbye(intentRequest);
                default:
                    _response = ResponseBuilder.Tell("Lo siento, no te he entendido.");
                    _response.Response.ShouldEndSession = false;
                    return _response;
            }
        }
        public async Task<SkillResponse> WelcomeMessage()
        {
            try
            {
                using (O365Client client = new O365Client())
                {
                    string name = await client.GetName();
                    _response = ResponseBuilder.Tell("Bienvenido a SharePoint Saturday Madrid, " + name);
                    _response.Response.ShouldEndSession = false;
                    return _response;
                }
            }
            catch (Exception)
            {
                _response = ResponseBuilder.Tell("Se ha producido un error con la Skill del SharePoint Saturday Madrid");
                _response.Response.ShouldEndSession = true;
                return _response;
            }
        }
        private async Task<SkillResponse> ReadDocument(IntentRequest intentRequest)
        {
            var id_documento = intentRequest.Intent.Slots["id_documento"].Value;

            using (O365Client _client = new O365Client())
            {
                string resultado=await _client.GetDocument("Eventos","Documentos",id_documento);
                _response = ResponseBuilder.Tell(resultado);
                _response.Response.ShouldEndSession = false;
                return _response;
            }
        }
        private async Task<SkillResponse> WarnManager(IntentRequest intentRequest)
        {
            using (O365Client _client = new O365Client())
            {
                string resultado = await _client.SendTeamAlert("ferran", "spsaturdaydemo.onmicrosoft.com");
                _response = ResponseBuilder.Tell(resultado);
                _response.Response.ShouldEndSession = false;
                return _response;
            }
        }
        private async Task<SkillResponse> ScheduleMeeting(IntentRequest intentRequest)
        {
            using (O365Client _client = new O365Client())
            {
                var dia = intentRequest.Intent.Slots["dia"].Value;
                var mes = intentRequest.Intent.Slots["mes"].Value;
                var horaInicio = intentRequest.Intent.Slots["hora_inicio"].Value;
                var horaFin= intentRequest.Intent.Slots["hora_fin"].Value;
                var tema= intentRequest.Intent.Slots["tema"].Value;

                int intDia = Convert.ToInt32(dia);
                int intHoraInicio = Convert.ToInt32(horaInicio);
                int intHoraFin= Convert.ToInt32(horaFin);

                bool resultado = await _client.CreateAppointment("Cita", tema, intDia, mes, 2019, intHoraInicio, intHoraFin);
                _response = ResponseBuilder.Tell("La cita se ha creado correctamente");
                _response.Response.ShouldEndSession = false;
                return _response;
            }
        }
        private async Task<SkillResponse> SayGoodbye(IntentRequest intentRequest)
        {

            string mensaje = "De parte de Ferran, muchas gracias por asistir a esta sesión";
            _response = ResponseBuilder.Tell(mensaje);
            _response.Response.ShouldEndSession = true;
            return _response;
        }

    }
}
