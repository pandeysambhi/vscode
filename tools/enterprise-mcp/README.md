# Enterprise integrations MCP server

This dependency-free, local MCP server gives the VS Code fork's Copilot agent one private tool surface for GitHub, Jira, and Microsoft Teams. Requests go directly from the developer machine to the configured enterprise APIs. No public GitHub MCP server is involved.

## Configure the shared workspace

The shared registration is checked in at `.vscode/mcp.json` and loads all credentials from the repository-root `.env`. Create it from the safe, committed template, then fill in the local values:

```bash
cp .env.example .env
```

The real `.env` is ignored by Git. Do not commit it. Run **MCP: List Servers** and start **enterprise-integrations** after populating the file. The equivalent registration is shown below for reuse in another workspace.

```jsonc
{
	"servers": {
		"enterprise-integrations": {
			"type": "stdio",
			"command": "node",
			"args": ["${workspaceFolder}/tools/enterprise-mcp/server.mjs"],
			"envFile": "${workspaceFolder}/.env"
		}
	}
}
```

For GitHub Enterprise Server, set `GITHUB_API_URL` to `https://github.example.com/api/v3`. The PAT needs only the repository and gist permissions required by your policy. Jira Cloud uses email plus API token; Jira Data Center can use a personal access token by leaving `JIRA_EMAIL` empty. The Teams tool expects a Microsoft Graph bearer token with the tenant-approved permission needed to create channel messages.

Set `ENTERPRISE_MCP_ALLOW_WRITES` to `true` only when the agent should be able to create PRs, gists, Jira issues, and Teams messages. Reading Jira issues remains available while writes are disabled.

## Run and test

```bash
cd tools/enterprise-mcp
npm test
node server.mjs
```

The server communicates over MCP stdio, so the second command waits for a client. It never logs tokens or request headers.

## Production hardening

The local stdio process is suitable for a controlled enterprise workstation. For a centrally hosted deployment, put an authenticated Streamable HTTP gateway in front of equivalent handlers, replace long-lived Graph tokens with Entra ID OAuth/on-behalf-of exchange, restrict GitHub/Jira scopes, add audit logging with secret redaction, and enforce repository/project/team allowlists server-side.
