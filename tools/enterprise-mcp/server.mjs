/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import { createInterface } from 'node:readline';
import { pathToFileURL } from 'node:url';

const jsonHeaders = { 'Content-Type': 'application/json' };

function required(value, name) {
	if (!value) {
		throw new Error(`${name} is not configured`);
	}
	return value;
}

function writesAllowed(env) {
	return env.ENTERPRISE_MCP_ALLOW_WRITES?.toLowerCase() === 'true';
}

function requireWrites(env) {
	if (!writesAllowed(env)) {
		throw new Error('Write tools are disabled. Set ENTERPRISE_MCP_ALLOW_WRITES=true after reviewing the requested operation.');
	}
}

async function request(url, init, fetchImpl) {
	const response = await fetchImpl(url, init);
	const text = await response.text();
	let body;
	try {
		body = text ? JSON.parse(text) : undefined;
	} catch {
		body = text;
	}
	if (!response.ok) {
		const detail = typeof body === 'string' ? body : JSON.stringify(body);
		throw new Error(`${init.method ?? 'GET'} ${url} failed (${response.status}): ${detail}`);
	}
	return body;
}

function githubHeaders(env) {
	return {
		...jsonHeaders,
		Accept: 'application/vnd.github+json',
		Authorization: `Bearer ${required(env.GITHUB_TOKEN, 'GITHUB_TOKEN')}`,
		'X-GitHub-Api-Version': '2022-11-28',
		'User-Agent': 'code-oss-enterprise-mcp'
	};
}

function jiraHeaders(env) {
	const token = required(env.JIRA_API_TOKEN, 'JIRA_API_TOKEN');
	const authorization = env.JIRA_EMAIL
		? `Basic ${Buffer.from(`${env.JIRA_EMAIL}:${token}`).toString('base64')}`
		: `Bearer ${token}`;
	return { ...jsonHeaders, Accept: 'application/json', Authorization: authorization };
}

export const tools = [
	{
		name: 'github_create_pull_request',
		description: 'Create a pull request in GitHub or GitHub Enterprise.',
		inputSchema: {
			type: 'object', additionalProperties: false,
			properties: {
				owner: { type: 'string' }, repo: { type: 'string' }, title: { type: 'string' },
				head: { type: 'string', description: 'Branch containing the changes.' },
				base: { type: 'string', description: 'Branch to merge into.' },
				body: { type: 'string' }, draft: { type: 'boolean', default: false }
			},
			required: ['owner', 'repo', 'title', 'head', 'base']
		},
		annotations: { title: 'Create GitHub pull request', destructiveHint: false, idempotentHint: false }
	},
	{
		name: 'github_create_gist',
		description: 'Create a GitHub gist. File values are the desired text contents.',
		inputSchema: {
			type: 'object', additionalProperties: false,
			properties: {
				description: { type: 'string' }, public: { type: 'boolean', default: false },
				files: { type: 'object', additionalProperties: { type: 'string' }, minProperties: 1 }
			},
			required: ['files']
		},
		annotations: { title: 'Create GitHub gist', destructiveHint: false, idempotentHint: false }
	},
	{
		name: 'jira_get_issue',
		description: 'Read one Jira issue by key.',
		inputSchema: {
			type: 'object', additionalProperties: false,
			properties: { issueKey: { type: 'string' } }, required: ['issueKey']
		},
		annotations: { title: 'Read Jira issue', readOnlyHint: true, idempotentHint: true }
	},
	{
		name: 'jira_create_issue',
		description: 'Create a Jira issue using Atlassian Document Format for the description.',
		inputSchema: {
			type: 'object', additionalProperties: false,
			properties: {
				projectKey: { type: 'string' }, summary: { type: 'string' }, description: { type: 'string' },
				issueType: { type: 'string', default: 'Task' }
			},
			required: ['projectKey', 'summary']
		},
		annotations: { title: 'Create Jira issue', destructiveHint: false, idempotentHint: false }
	},
	{
		name: 'teams_send_channel_message',
		description: 'Send a plain-text message to a Microsoft Teams channel using Microsoft Graph.',
		inputSchema: {
			type: 'object', additionalProperties: false,
			properties: { teamId: { type: 'string' }, channelId: { type: 'string' }, message: { type: 'string' } },
			required: ['teamId', 'channelId', 'message']
		},
		annotations: { title: 'Send Teams channel message', destructiveHint: false, idempotentHint: false }
	}
];

export async function callTool(name, args, env = process.env, fetchImpl = fetch) {
	switch (name) {
		case 'github_create_pull_request': {
			requireWrites(env);
			const baseUrl = (env.GITHUB_API_URL || 'https://api.github.com').replace(/\/$/, '');
			return request(`${baseUrl}/repos/${encodeURIComponent(args.owner)}/${encodeURIComponent(args.repo)}/pulls`, {
				method: 'POST', headers: githubHeaders(env), body: JSON.stringify({ title: args.title, head: args.head, base: args.base, body: args.body, draft: args.draft ?? false })
			}, fetchImpl);
		}
		case 'github_create_gist': {
			requireWrites(env);
			const baseUrl = (env.GITHUB_API_URL || 'https://api.github.com').replace(/\/$/, '');
			const files = Object.fromEntries(Object.entries(args.files).map(([name, content]) => [name, { content }]));
			return request(`${baseUrl}/gists`, {
				method: 'POST', headers: githubHeaders(env), body: JSON.stringify({ description: args.description, public: args.public ?? false, files })
			}, fetchImpl);
		}
		case 'jira_get_issue': {
			const baseUrl = required(env.JIRA_BASE_URL, 'JIRA_BASE_URL').replace(/\/$/, '');
			return request(`${baseUrl}/rest/api/3/issue/${encodeURIComponent(args.issueKey)}`, { headers: jiraHeaders(env) }, fetchImpl);
		}
		case 'jira_create_issue': {
			requireWrites(env);
			const baseUrl = required(env.JIRA_BASE_URL, 'JIRA_BASE_URL').replace(/\/$/, '');
			const description = args.description ? { type: 'doc', version: 1, content: [{ type: 'paragraph', content: [{ type: 'text', text: args.description }] }] } : undefined;
			return request(`${baseUrl}/rest/api/3/issue`, {
				method: 'POST', headers: jiraHeaders(env), body: JSON.stringify({ fields: { project: { key: args.projectKey }, summary: args.summary, description, issuetype: { name: args.issueType ?? 'Task' } } })
			}, fetchImpl);
		}
		case 'teams_send_channel_message': {
			requireWrites(env);
			const graphUrl = (env.MS_GRAPH_URL || 'https://graph.microsoft.com/v1.0').replace(/\/$/, '');
			return request(`${graphUrl}/teams/${encodeURIComponent(args.teamId)}/channels/${encodeURIComponent(args.channelId)}/messages`, {
				method: 'POST', headers: { ...jsonHeaders, Authorization: `Bearer ${required(env.MS_GRAPH_TOKEN, 'MS_GRAPH_TOKEN')}` },
				body: JSON.stringify({ body: { contentType: 'text', content: args.message } })
			}, fetchImpl);
		}
		default:
			throw new Error(`Unknown tool: ${name}`);
	}
}

export async function handleMessage(message, env = process.env, fetchImpl = fetch) {
	if (message.method === 'initialize') {
		return { protocolVersion: '2025-06-18', capabilities: { tools: { listChanged: false } }, serverInfo: { name: 'enterprise-integrations', version: '0.1.0' } };
	}
	if (message.method === 'ping') {
		return {};
	}
	if (message.method === 'tools/list') {
		return { tools };
	}
	if (message.method === 'tools/call') {
		try {
			const result = await callTool(message.params.name, message.params.arguments ?? {}, env, fetchImpl);
			return { content: [{ type: 'text', text: JSON.stringify(result, undefined, 2) }] };
		} catch (error) {
			return { isError: true, content: [{ type: 'text', text: error instanceof Error ? error.message : String(error) }] };
		}
	}
	throw new Error(`Unsupported method: ${message.method}`);
}

export async function start(input = process.stdin, output = process.stdout) {
	const lines = createInterface({ input, crlfDelay: Infinity });
	for await (const line of lines) {
		if (!line.trim()) {
			continue;
		}
		let message;
		try {
			message = JSON.parse(line);
			if (message.id === undefined) {
				continue;
			}
			const result = await handleMessage(message);
			output.write(`${JSON.stringify({ jsonrpc: '2.0', id: message.id, result })}\n`);
		} catch (error) {
			const errorMessage = error instanceof Error ? error.message : String(error);
			output.write(`${JSON.stringify({ jsonrpc: '2.0', id: message?.id ?? null, error: { code: -32603, message: errorMessage } })}\n`);
		}
	}
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
	start().catch(error => {
		console.error(error);
		process.exitCode = 1;
	});
}
