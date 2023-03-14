using Microsoft.Graph;

namespace Personalized365.Web.Graph
{
    public class GraphEmailClient
    {
        private readonly ILogger<GraphEmailClient> _logger = null;
        private readonly GraphServiceClient _graphServiceClient = null;

        public GraphEmailClient(
            ILogger<GraphEmailClient> logger,
            GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<IEnumerable<Message>> GetUserMessages()
        {
            try
            {
                var emails = await _graphServiceClient.Me.Messages
                    .Request()
                    .Select(msg => new
                    {
                        msg.Subject,
                        msg.BodyPreview,
                        msg.ReceivedDateTime
                    })
                    .OrderBy("receivedDateTime desc")
                    .Top(10)
                    .GetAsync();
                return emails.CurrentPage;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error calling Graph /me/messages: {ex.Message}");
                throw;
            }
        }

        public async Task<(IEnumerable<Message> Messages, string NextLink)> GetUserMessagesPage(
            string nextPageLink = null, int top = 2)
        {
            IUserMessagesCollectionPage pagedMessages;

            try
            {
                if (nextPageLink == null)
                {
                    // Get initial page of messages
                    pagedMessages = await _graphServiceClient.Me.Messages
                            .Request()
                            .Select(msg => new
                            {
                                msg.Subject,
                                msg.BodyPreview,
                                msg.ReceivedDateTime
                            })
                            .Top(top)
                            .OrderBy("receivedDateTime desc")
                            .GetAsync();
                }
                else
                {
                    // Use the next page request URI value to get the page of messages
                    var messagesCollectionRequest = new UserMessagesCollectionRequest(nextPageLink, _graphServiceClient, null);
                    pagedMessages = await messagesCollectionRequest.GetAsync();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error calling Graph /me/messages to page messages: {ex.Message}");
                throw;
            }

            return (Messages: pagedMessages, NextLink: GetNextLink(pagedMessages));
        }

        private string GetNextLink(IUserMessagesCollectionPage pagedMessages)
        {
            if (pagedMessages.NextPageRequest != null)
            {
                // Get the URL for the next batch of records
                return pagedMessages
                    .NextPageRequest
                    .GetHttpRequestMessage()
                    .RequestUri?
                    .OriginalString;
            }

            return null;
        }
    }
}
