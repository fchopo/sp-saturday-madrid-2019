using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace O365Skill.Graph
{
    public class O365Client: IDisposable
    {
        GraphServiceClient _graphServiceClient;

        #region "Keys"
        private const string _clientId = "f721daa4-0661-4224-a2a9-9895a12a37ec";
        private const string _clientSecret = "qn.nU28vIWpvdGsNQ0y7njsIFc=jp:w-";
        private const string _tenantId = "e127ea86-f50b-4917-baff-b26ccd9c7be8";
        #endregion

        //https://github.com/microsoftgraph/msgraph-sdk-dotnet-auth

        public O365Client()
        {
            IConfidentialClientApplication clientApplication = ClientCredentialProvider.CreateClientApplication(_clientId, new ClientCredential(_clientSecret), null, _tenantId);
            ClientCredentialProvider authenticationProvider = new ClientCredentialProvider(clientApplication);

            _graphServiceClient = new GraphServiceClient(authenticationProvider);
        }

        public async Task<string> GetName()
        {
            IGraphServiceUsersCollectionPage users = await _graphServiceClient.Users.Request().GetAsync();

            return users.Where(x => x.DisplayName.Contains("Ferran")).FirstOrDefault().DisplayName;
        }

        public async Task<string> GetDocument(string siteName, string docLibraryName, string id)
        {
            List<Option> options = new List<Option>();
            options.Add(new QueryOption("search", siteName));

            Site site = (await _graphServiceClient.Sites.Request(options).GetAsync()).First();

            Drive docLibrary = (await (_graphServiceClient.Sites[site.Id].Drives.Request().GetAsync())).Where(x => x.Name == docLibraryName).First();

            IEnumerable<DriveItem> documentsList = await _graphServiceClient.Sites[site.Id].Drives[docLibrary.Id].Root.Children.Request().GetAsync();

            foreach (DriveItem document in documentsList)
            {
                Microsoft.Graph.ListItem item = await _graphServiceClient.Sites[site.Id].Drives[docLibrary.Id].Items[document.Id].ListItem.Request().GetAsync();
                if (item.Id == id)
                {
                    if ((document.File.MimeType.ToLower() == "application/vnd.openxmlformats-officedocument.wordprocessingml.document") || (document.File.MimeType.ToLower() == "application / msword"))
                    {
                        using (Stream stream = await _graphServiceClient.Sites[site.Id].Drives[docLibrary.Id].Items[document.Id].Content.Request().GetAsync())
                        {
                            WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, false);
                            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
                            wordprocessingDocument.Close();
                            wordprocessingDocument.Dispose();
                            return body.InnerText;
                        }

                    }

                    if (document.File.MimeType.ToLower() == "application/pdf")
                    {
                        using (Stream stream = await _graphServiceClient.Sites[site.Id].Drives[docLibrary.Id].Items[document.Id].Content.Request().GetAsync())
                        {
                            PdfReader _reader = new PdfReader(stream);
                            ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                            string currentText = PdfTextExtractor.GetTextFromPage(_reader, 1, strategy);
                            currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                            return currentText;
                        }
                    }
                }
            }
            return "Lo siento, no he encontrado el documento que buscas.";
        }

        public async Task<string> SendEmailAlert(string userName, string domain)
        {
            string upn = $"{userName}@{domain}";
            User manager = (await _graphServiceClient.Users[upn].Manager.Request().GetAsync()) as User;
            Message mailMessage = new Message();
            mailMessage.Subject = "Llego tarde!";
            mailMessage.Body = new ItemBody
            {
                ContentType = Microsoft.Graph.BodyType.Text,
                Content = "Pues sí... Llego tarde otra vez"
            };
            mailMessage.ToRecipients = new List<Recipient>()
            {
                new Recipient {EmailAddress=new EmailAddress{Address=manager.Mail}}
            };
            var saveToSentItems = true;
            await _graphServiceClient.Users[upn].SendMail(mailMessage, saveToSentItems).Request().PostAsync();
            return "Aviso entregado!";
        }

        public async Task<string> SendTeamAlert(string userName, string domain)
        {
            using(HttpClient httpClient=new HttpClient())
            {
                httpClient.DefaultRequestHeaders.Accept.Clear();
                httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                httpClient.BaseAddress = new Uri("https://prod-77.westeurope.logic.azure.com:443");
                TeamsMessage data = new TeamsMessage() { mensaje = "Llego tarde" };

                var stringPayload = JsonConvert.SerializeObject(data);
                var httpContent = new StringContent(stringPayload, Encoding.UTF8, "application/json");

                await httpClient.PostAsync("/workflows/28756ee1cdc54416bfd41935a2a3e18d/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=fRHR_A3vEZW9osNnr5BE-60OJW9zSTPjmJnBfGjvxrk",httpContent);
                return "Aviso publicado en Teams!";
            }
        }

        public async Task<bool> CreateAppointment(string subject, string content, int day, string monthName, int year, int startTime, int EndTime)
        {
            try
            {
                User user = (await _graphServiceClient.Users.Request().GetAsync()).Where(x => x.DisplayName.Contains("Ferran")).FirstOrDefault();

                Microsoft.Graph.Event @event = new Event();
                @event.Subject = subject;

                ItemBody body = new ItemBody();
                body.Content = content;
                body.ContentType = Microsoft.Graph.BodyType.Html;

                DateTimeTimeZone start = new DateTimeTimeZone();
                start.DateTime = new DateTime(year, DateTime.ParseExact(monthName, "MMMM", new CultureInfo("es-ES")).Month, day, startTime, 0, 0).ToString("yyyy-MM-ddTHH:mm:ss");
                start.TimeZone = "Europe/Madrid";

                DateTimeTimeZone end = new DateTimeTimeZone();
                end.DateTime = new DateTime(year, DateTime.ParseExact(monthName, "MMMM", new CultureInfo("es-ES")).Month, day, EndTime, 0, 0).ToString("yyyy-MM-ddTHH:mm:ss");
                end.TimeZone = "Europe/Madrid";

                Location location = new Location();
                location.DisplayName = "Microsoft Iberica Madrid";


                @event.Body = body;
                @event.Start = start;
                @event.End = end;
                @event.Location = location;

                await _graphServiceClient.Users[user.Id].Calendar.Events.Request().AddAsync(@event);

                return true;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        #region IDisposable Support
        private bool disposedValue = false; // Para detectar llamadas redundantes

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: elimine el estado administrado (objetos administrados).
                }

                // TODO: libere los recursos no administrados (objetos no administrados) y reemplace el siguiente finalizador.
                // TODO: configure los campos grandes en nulos.

                disposedValue = true;
            }
        }

        // TODO: reemplace un finalizador solo si el anterior Dispose(bool disposing) tiene código para liberar los recursos no administrados.
        // ~O365Client()
        // {
        //   // No cambie este código. Coloque el código de limpieza en el anterior Dispose(colocación de bool).
        //   Dispose(false);
        // }

        // Este código se agrega para implementar correctamente el patrón descartable.
        public void Dispose()
        {
            // No cambie este código. Coloque el código de limpieza en el anterior Dispose(colocación de bool).
            Dispose(true);
            // TODO: quite la marca de comentario de la siguiente línea si el finalizador se ha reemplazado antes.
            // GC.SuppressFinalize(this);
        }
        #endregion
    }

    public class TeamsMessage
    {
        public string mensaje { get; set; }
    }
}
