using Azure;
using Azure.AI.TextAnalytics;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using Microsoft.Graph;
using Microsoft.Identity.Web;
using Personalized365.Web.Graph;
using Personalized365.Web.ViewModels;

namespace Personalized365.Web.Pages
{
    [AuthorizeForScopes(ScopeKeySection = "DownstreamApi:Scopes")]
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;
        private readonly GraphProfileClient _graphProfileClient;
        private readonly GraphEmailClient _graphEmailClient;
        private readonly GraphCalendarClient _graphCalendarClient;
        private readonly GraphTodoClient _graphTodoClient;
        private readonly AppSettings appSettings;

        public string UserDisplayName { get; private set; } = "";
        public string? UserPhoto { get; private set; }

        public IEnumerable<Message> Messages { get; private set; }
        public List<SummarizedMessage> SummarizedMessages { get; private set; }

        private MailboxSettings MailboxSettings { get; set; }
        public IEnumerable<Event> Events { get; private set; }

        public IEnumerable<TodoTask> TodoTasks { get; private set; }

        public IndexModel(
            ILogger<IndexModel> logger,
            ITokenAcquisition tokenAcquisition,
            GraphProfileClient graphProfileClient,
            GraphEmailClient graphEmailClient,
            GraphCalendarClient graphCalendarClient,
            GraphTodoClient graphTodoClient)
        {
            _logger = logger;
            _graphProfileClient = graphProfileClient;
            _graphEmailClient = graphEmailClient;
            _graphCalendarClient = graphCalendarClient;
            _graphTodoClient = graphTodoClient;

            appSettings = AppSettings.Load();
        }

        public async Task OnGetAsync()
        {
            var user = await _graphProfileClient.GetUserProfile();
            UserDisplayName = user.DisplayName.Split(' ')[0];
            UserPhoto = await _graphProfileClient.GetUserProfileImage();
            await GetEmailAsync();
            await GetCalendarEventAsync();
            await GetTodoAsync();
        }

        #region Email

        private async Task GetEmailAsync()
        {
            Messages = await _graphEmailClient.GetUserMessages();
            SummarizedMessages = new List<SummarizedMessage>();

            var textAnalyticsClient = GetTextAnalyticsClient();

            foreach (var message in Messages)
            {
                SummarizedMessages.Add(new SummarizedMessage
                {
                    Id = message.Id,
                    Subject = message.Subject,
                    ReceivedUtcDateTime = message.ReceivedDateTime.GetValueOrDefault().UtcDateTime,
                    BodyPreview = message.BodyPreview,
                    SummarySentences = await GetSummaryAsync(textAnalyticsClient, message.BodyPreview)
                });
            }
        }

        private TextAnalyticsClient GetTextAnalyticsClient()
        {
            var credential = new AzureKeyCredential(appSettings.TextAnalyticsCredential);
            var endpoint = new Uri(appSettings.TextAnalyticsEndpoint);

            return new TextAnalyticsClient(endpoint, credential);
        }

        private async Task<IList<string>> GetSummaryAsync(TextAnalyticsClient client, string emailBody)
        {
            var sentances = new List<string>();

            // Prepare analyze operation input. You can add multiple documents to this list and perform the same
            // operation to all of them.
            var batchInput = new List<string>
            {
                emailBody
            };

            TextAnalyticsActions actions = new TextAnalyticsActions()
            {
                ExtractSummaryActions = new List<ExtractSummaryAction>() { new ExtractSummaryAction() }
            };

            // Start analysis process.
            AnalyzeActionsOperation operation = await client.StartAnalyzeActionsAsync(batchInput, actions);
            await operation.WaitForCompletionAsync();
            // View operation status.
            Console.WriteLine($"AnalyzeActions operation has completed");
            Console.WriteLine();

            Console.WriteLine($"Created On   : {operation.CreatedOn}");
            Console.WriteLine($"Expires On   : {operation.ExpiresOn}");
            Console.WriteLine($"Id           : {operation.Id}");
            Console.WriteLine($"Status       : {operation.Status}");

            Console.WriteLine();
            // View operation results.
            await foreach (AnalyzeActionsResult documentsInPage in operation.Value)
            {
                IReadOnlyCollection<ExtractSummaryActionResult> summaryResults = documentsInPage.ExtractSummaryResults;

                foreach (ExtractSummaryActionResult summaryActionResults in summaryResults)
                {
                    if (summaryActionResults.HasError)
                    {
                        Console.WriteLine($"  Error!");
                        Console.WriteLine($"  Action error code: {summaryActionResults.Error.ErrorCode}.");
                        Console.WriteLine($"  Message: {summaryActionResults.Error.Message}");
                        continue;
                    }

                    foreach (ExtractSummaryResult documentResults in summaryActionResults.DocumentsResults)
                    {
                        if (documentResults.HasError)
                        {
                            Console.WriteLine($"  Error!");
                            Console.WriteLine($"  Document error code: {documentResults.Error.ErrorCode}.");
                            Console.WriteLine($"  Message: {documentResults.Error.Message}");
                            continue;
                        }

                        Console.WriteLine($"  Extracted the following {documentResults.Sentences.Count} sentence(s):");
                        Console.WriteLine();

                        foreach (SummarySentence sentence in documentResults.Sentences)
                        {
                            Console.WriteLine($"  Sentence: {sentence.Text}");
                            Console.WriteLine();

                            sentances.Add(sentence.Text);
                        }
                    }
                }
            }

            return sentances;
        }

        public async Task<PartialViewResult> OnGetViewEmail(string id)
        {
            var email = (await _graphEmailClient.GetUserMessages())
                .First(model => model.Id == id);

            return new PartialViewResult
            {
                ViewName = "ViewEmail",
                ViewData = new ViewDataDictionary<EmailViewModel>(ViewData, new EmailViewModel
                {
                    Subject = email.Subject,
                    Body = email.BodyPreview,
                    ReceivedUtcDateTime = email.ReceivedDateTime.GetValueOrDefault().UtcDateTime
                })
            };
        }

        public async Task<PartialViewResult> OnPostAddEmailToDo(string id)
        {
            var email = (await _graphEmailClient.GetUserMessages())
                .First(model => model.Id == id);

            var todoTask = new TodoTask
            {
                Title = $"From email: {email.Subject}",
                Categories = new List<string>
                {
                    "Important"
                },
                Importance = Importance.High
            };

            await _graphTodoClient.AddTodo(todoTask);
            await GetTodoAsync();

            return new PartialViewResult
            {
                ViewName = "DialogBox",
                ViewData = new ViewDataDictionary<DialogBoxViewModel>(ViewData, new DialogBoxViewModel
                {
                    Message = "Added successfully."
                })
            };
        }

        #endregion

        #region Calendar Event

        private async Task GetCalendarEventAsync()
        {
            MailboxSettings = await _graphCalendarClient.GetUserMailboxSettings();
            var userTimeZone = (string.IsNullOrEmpty(MailboxSettings.TimeZone))
                ? "Pacific Standard Time"
                : MailboxSettings.TimeZone;
            Events = await _graphCalendarClient.GetEvents(userTimeZone);
        }

        public string FormatDateTimeTimeZone(DateTimeTimeZone value)
        {
            // Parse the date/time string from Graph into a DateTime
            var graphDatetime = value.DateTime;
            if (DateTime.TryParse(graphDatetime, out DateTime dateTime))
            {
                var dateTimeFormat = $"{MailboxSettings.DateFormat} {MailboxSettings.TimeFormat}".Trim();
                if (!string.IsNullOrEmpty(dateTimeFormat))
                {
                    return dateTime.ToString(dateTimeFormat);
                }
                else
                {
                    return $"{dateTime.ToShortDateString()} {dateTime.ToShortTimeString()}";
                }
            }
            else
            {
                return graphDatetime;
            }
        }

        #endregion

        #region To do

        private async Task GetTodoAsync()
        {
            TodoTasks = await _graphTodoClient.GetTodos();
        }

        #endregion
    }
}