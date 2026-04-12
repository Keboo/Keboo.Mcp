using System.Reflection;
using Keboo.Mcp.Configuration;
using Keboo.Mcp.Services;
using Keboo.Mcp.Tools;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

HostApplicationBuilder builder = Host.CreateApplicationBuilder(args);

builder.Configuration.AddUserSecrets(Assembly.GetExecutingAssembly(), optional: true);
builder.Logging.AddConsole(options => options.LogToStandardErrorThreshold = LogLevel.Trace);

builder.Services.Configure<TeamsGraphOptions>(builder.Configuration.GetSection(TeamsGraphOptions.SectionName));

builder.Services
    .AddSingleton<ITeamsGraphClient, MicrosoftGraphTeamsClient>()
    .AddSingleton<ITeamsChatService, TeamsChatService>();

builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithTools<TeamsChatTools>();

await builder.Build().RunAsync();
