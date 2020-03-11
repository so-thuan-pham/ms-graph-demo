using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Calendar = Microsoft.Graph.Calendar;

namespace MSGraphCalendarApp1
{
    public static class GraphHelper
    {
        private static string instance = ConfigurationManager.AppSettings["AAD:Instance"];
        private static string tenant = ConfigurationManager.AppSettings["AAD:Tenant"];
        private static string clientId = ConfigurationManager.AppSettings["AAD:ClientId"];
        private static string clientSecret = ConfigurationManager.AppSettings["AAD:ClientSecret"];
        private static string authority = String.Format(CultureInfo.InvariantCulture, instance, tenant);
        private static string graphScopes = ConfigurationManager.AppSettings["AAD:AppScopes"];

        public static async Task<User> GetUserDetailsAsync(string accessToken)
        {
            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", accessToken);
                    }));

            return await graphClient.Me.Request().GetAsync();
        }

        public static async Task<IEnumerable<Event>> GetEventsAsync()
        {
            var graphClient = GetAuthenticatedClient();

            var events = await graphClient.Me.Events.Request()
                .Select("subject,organizer,start,end")
                .OrderBy("createdDateTime DESC")
                .GetAsync();

            return events.CurrentPage;
        }

        public static async void CreateCalendar()
        {
            var graphClient = GetAuthenticatedClient();

            var userId = "2d33c3c9-833e-4e73-91b8-fd374200d5a9";
            var userDisplayName = "Thuan Pham";
            var userMail = "thuan.pham@schooloutfitters.com";

            var sharedCalendar = graphClient.Users[userId].Calendars.Request().Filter("name eq 'Thuan Test Shared Calendar2'").GetAsync().Result.FirstOrDefault();

            if (sharedCalendar == null)
            {
                sharedCalendar = new Microsoft.Graph.Calendar
                {
                    Name = "Thuan Test Shared Calendar2",
                    CanEdit = false,
                    CanShare = true,
                    Color = CalendarColor.LightGreen,
                    Owner = new EmailAddress { Name = userDisplayName, Address = userMail }
                };

                sharedCalendar = await graphClient.Users[userId].Calendars.Request().AddAsync(sharedCalendar);
            }

            CreateEvents(userId, sharedCalendar.Id);
        }

        public static async void CreateEvents(string userId, string calendarId)
        {
            var graphClient = GetAuthenticatedClient();

            var events = new List<Event>
            {
                new Event
                {
                    Subject = "Event 1",
                    Body = new ItemBody { Content = "Event 1 Body" },
                    Categories = new string[] { "Information" },
                    Start = DateTimeTimeZone.FromDateTime(DateTime.Now.AddDays(1)),
                    End = DateTimeTimeZone.FromDateTime(DateTime.Now.AddDays(2))
                },
                new Event
                {
                    Subject = "Event 2",
                    Body = new ItemBody { Content = "Event 2 Body" },
                    Categories = new string[] { "Information" },
                    Start = DateTimeTimeZone.FromDateTime(DateTime.Now.AddDays(1)),
                    End = DateTimeTimeZone.FromDateTime(DateTime.Now.AddDays(2))
                }
            };

            var tasks = from myEvent in events select CreateEvent(userId, calendarId, myEvent);

            await Task.WhenAll(tasks);
        }

        private static async Task CreateEvent(string userId, string calendarId, Event myEvent)
        {
            var graphClient = GetAuthenticatedClient();

            await graphClient.Users[userId].Calendars[calendarId].Events.Request().AddAsync(myEvent);
        }

        private static GraphServiceClient GetAuthenticatedClient()
        {
            return new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        var confidentialClient = ConfidentialClientApplicationBuilder.Create(clientId)
                            .WithAuthority(new Uri(authority))
                            .WithClientSecret(clientSecret)
                            .Build();

                        //var tokenStore = new SessionTokenStore(idClient.UserTokenCache,
                        //    HttpContext.Current, ClaimsPrincipal.Current);

                        //var accounts = await idClient.GetAccountsAsync();

                        // By calling this here, the token can be refreshed
                        // if it's expired right before the Graph call is made
                        var scopes = new string[] { "https://graph.microsoft.com/.default" };
                        AuthenticationResult result = null;
                        try
                        {
                            result = await confidentialClient.AcquireTokenForClient(scopes).ExecuteAsync();
                        }
                        catch (Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Scope provided is not supported");
                            Console.ResetColor();
                            throw;
                        }

                        

                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", result.AccessToken);
                    }));
        }
    }
}
