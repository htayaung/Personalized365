using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Graph;
using Microsoft.Identity.Web;
using Personalized365.Web.Graph;

namespace Personalized365.Web.Pages
{
    [AuthorizeForScopes(ScopeKeySection = "DownstreamApi:Scopes")]
    public class TeamsModel : PageModel
    {
        private readonly GraphTeamClient _graphTeamClient;

        public IEnumerable<Team> MyTeams { get; private set; }

        public TeamsModel(GraphTeamClient graphTeamClient)
        {
            _graphTeamClient = graphTeamClient;
        }

        public async Task OnGetAsync()
        {
            MyTeams = await _graphTeamClient.GetMyTeams();
        }
    }
}
