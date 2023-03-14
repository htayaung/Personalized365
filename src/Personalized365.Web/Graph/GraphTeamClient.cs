using Microsoft.Graph;

namespace Personalized365.Web.Graph
{
    public class GraphTeamClient
    {
        private readonly ILogger<GraphTeamClient> _logger = null;
        private readonly GraphServiceClient _graphServiceClient = null;

        public GraphTeamClient(
            ILogger<GraphTeamClient> logger,
            GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<IEnumerable<Team>> GetMyTeams()
        {
            _logger.LogInformation("My joined teams.");

            try
            {
                // Use GraphServiceClient to call Me.CalendarView
                var teams = await _graphServiceClient
                    .Me
                    .JoinedTeams
                    .Request()
                    .GetAsync();

                await SetChannels(teams);

                return teams;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error calling Graph /me/joinedTeams: {ex.Message}");
                throw;
            }
        }

        private async Task SetChannels(IEnumerable<Team> teams)
        {
            foreach (var team in teams)
            {
                team.AllChannels = await _graphServiceClient
                    .Teams[team.Id]
                    .AllChannels
                    .Request()
                    .GetAsync();
            }
        }
    }
}
