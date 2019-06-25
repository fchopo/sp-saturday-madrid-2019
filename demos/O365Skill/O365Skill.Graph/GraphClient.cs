using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using Newtonsoft.Json;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Text;
using System.Globalization;

namespace O365Skill.Graph
{
    public class GraphClient: IDisposable
    {
        GraphServiceClient _graphServiceClient;
        private const string _clientId = "2066c24b-b6ad-4292-ad4d-d180ba16aa4a";
        private const string _clientSecret = "SFMa/2]b3pdxw+@zPZ4IMMXvhIkGz6+0";
        private const string _tenantId = "829dabd8-9473-4518-a8f6-654c5dc49b6a";

        //https://github.com/microsoftgraph/msgraph-sdk-dotnet-auth
        public GraphClient()
        {
            IConfidentialClientApplication clientApplication = ClientCredentialProvider.CreateClientApplication(_clientId, new ClientCredential(_clientSecret),null,_tenantId);

            ClientCredentialProvider authenticationProvider = new ClientCredentialProvider(clientApplication);

            _graphServiceClient = new GraphServiceClient(authenticationProvider);
        }

        public async Task<string> GetName()
        {
            IGraphServiceUsersCollectionPage users = await _graphServiceClient.Users.Request().Select(x => x.DisplayName).GetAsync();

            return users[0].DisplayName;
        }

        public async Task<List<string>> GetPlayers()
        {
            try
            {
                List<string> data = new List<string>();
                List<Option> options = new List<Option>();

                options.Add(new QueryOption("search", "devsite"));

                Site site = (await _graphServiceClient.Sites.Request(options).GetAsync()).First();



                List playersList = (await (_graphServiceClient.Sites[site.Id].Lists.Request().GetAsync())).Where(x => x.Name == "Players").First();

                IEnumerable<Microsoft.Graph.ListItem> players = await _graphServiceClient.Sites[site.Id].Lists[playersList.Id].Items.Request().GetAsync();

                foreach (Microsoft.Graph.ListItem listItem in players)
                {
                    FieldValueSet fieldValueSet = (await _graphServiceClient.Sites[site.Id].Lists[playersList.Id].Items[listItem.Id].Fields.Request().GetAsync());
                    IDictionary<string, object> keyValuePairs = fieldValueSet.AdditionalData;

                    data.Add(keyValuePairs["Title"].ToString());
                }
                return data;
            }
            catch (Exception ex)
            {
                return null;
            }
            
        }

        public async Task<List<string>> GetDocument()
        {
            try
            {
                List<string> data = new List<string>();
                List<Option> options = new List<Option>();

                options.Add(new QueryOption("search", "devsite"));

                Site site = (await _graphServiceClient.Sites.Request(options).GetAsync()).First();

                Drive docContratos = (await (_graphServiceClient.Sites[site.Id].Drives.Request().GetAsync())).Where(x => x.Name == "Contratos").First();

                IEnumerable<DriveItem> contratos = await _graphServiceClient.Sites[site.Id].Drives[docContratos.Id].Root.Children.Request().GetAsync();

                foreach (DriveItem driveItem in contratos)
                {
                    if ((driveItem.File.MimeType.ToLower()== "application/vnd.openxmlformats-officedocument.wordprocessingml.document")||(driveItem.File.MimeType.ToLower()== "application / msword"))
                    { 
                        Stream stream = await _graphServiceClient.Sites[site.Id].Drives[docContratos.Id].Items[driveItem.Id].Content.Request().GetAsync();
                        WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, false);
                        Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
                        wordprocessingDocument.Close();
                        wordprocessingDocument.Dispose();
                        data.Add(body.InnerText);                    
                    }
                    if (driveItem.File.MimeType.ToLower() == "application/pdf")
                    {
                        Stream stream = await _graphServiceClient.Sites[site.Id].Drives[docContratos.Id].Items[driveItem.Id].Content.Request().GetAsync();
                        PdfReader _reader = new PdfReader(stream);
                        ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                        string currentText = PdfTextExtractor.GetTextFromPage(_reader, 1, strategy);
                        currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                        data.Add(currentText);
                    }

                }
                return data;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public async Task<string> GetUserInfo()
        {
            try
            {
                List<Option> options = new List<Option>();

                options.Add(new QueryOption("select", "MobilePhone,Skills"));

                User user = (await _graphServiceClient.Users.Request().GetAsync()).Where(x => x.DisplayName.Contains("Ferran")).FirstOrDefault();

                User detailedUser=await _graphServiceClient.Users[user.Id].Request().Expand("extensions").GetAsync();

                return detailedUser.Skills.FirstOrDefault();
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public async Task<bool> CreateAppointment(string subject, string content, int day, string monthName, int year, int startTime, int EndTime )
        {
            try
            {
                User user = (await _graphServiceClient.Users.Request().GetAsync()).Where(x => x.DisplayName.Contains("Ferran")).FirstOrDefault();

                

                Microsoft.Graph.Event @event = new Event();
                @event.Subject = subject;

                ItemBody body = new ItemBody();
                body.Content = content;
                body.ContentType =  Microsoft.Graph.BodyType.Html;

                DateTimeTimeZone start = new DateTimeTimeZone();
                start.DateTime = new DateTime(year, DateTime.ParseExact(monthName, "MMMM", new CultureInfo("es-ES")).Month, day, startTime, 0, 0).ToString("yyyy-MM-ddTHH:mm:ss");
                start.TimeZone = "Europe/Madrid";

                DateTimeTimeZone end = new DateTimeTimeZone();
                end.DateTime = new DateTime(year, DateTime.ParseExact(monthName, "MMMM", new CultureInfo("es-ES")).Month, day, EndTime, 0, 0).ToString("yyyy-MM-ddTHH:mm:ss");
                end.TimeZone = "Europe/Madrid";

                Location location = new Location();
                location.DisplayName= "Microsoft Iberica Madrid";
                

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
        // ~GraphClient()
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
}
