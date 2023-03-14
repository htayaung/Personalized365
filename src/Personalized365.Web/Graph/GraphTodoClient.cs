using Microsoft.Graph;

namespace Personalized365.Web.Graph
{
    public class GraphTodoClient
    {
        private readonly ILogger<GraphTodoClient> _logger;
        private readonly GraphServiceClient _graphServiceClient;

        public GraphTodoClient(
            ILogger<GraphTodoClient> logger,
            GraphServiceClient graphServiceClient)
        {
            _logger = logger;
            _graphServiceClient = graphServiceClient;
        }

        public async Task<IEnumerable<TodoTask>> GetTodos()
        {
            _logger.LogInformation($"User todo list");
            var tasks = new List<TodoTask>();

            try
            {
                var todoCollection = await _graphServiceClient
                    .Me
                    .Todo
                    .Lists
                    .Request()
                    .GetAsync();

                foreach (var entity in todoCollection)
                {
                    var result = await _graphServiceClient
                        .Me
                        .Todo
                        .Lists[entity.Id]
                        .Tasks
                        .Request()
                        .OrderBy("createdDateTime")
                        .GetAsync();
                    tasks.AddRange(result.ToList());
                }

                return tasks;
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error calling Graph /me/todo/lists: {ex.Message}");
                throw;
            }
        }

        public async Task<TodoTask> AddTodo(TodoTask task, string? todoTaskListId = null)
        {
            if (todoTaskListId is null)
            {
                var taskListCollection = await _graphServiceClient
                    .Me
                    .Todo
                    .Lists
                    .Request()
                    .GetAsync();

                // TODO: Check further for collection id
                todoTaskListId = taskListCollection.First().Id;
            }

            return await _graphServiceClient
                .Me
                .Todo
                .Lists[todoTaskListId]
                .Tasks
                .Request()
                .AddAsync(task);
        }
    }
}
