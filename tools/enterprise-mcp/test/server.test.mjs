/*---------------------------------------------------------------------------------------------
 *  Copyright (c) Microsoft Corporation. All rights reserved.
 *  Licensed under the MIT License. See License.txt in the project root for license information.
 *--------------------------------------------------------------------------------------------*/

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { callTool, handleMessage } from '../server.mjs';

test('lists the unified enterprise tools', async () => {
	const result = await handleMessage({ method: 'tools/list' });
	assert.deepEqual(result.tools.map(tool => tool.name), [
		'github_create_pull_request',
		'github_create_gist',
		'jira_get_issue',
		'jira_create_issue',
		'teams_send_channel_message'
	]);
});

test('blocks write operations unless explicitly enabled', async () => {
	await assert.rejects(
		callTool('github_create_gist', { files: { 'note.txt': 'hello' } }, {}, async () => assert.fail('fetch should not run')),
		/Write tools are disabled/
	);
});

test('creates a pull request against a configurable GitHub API', async () => {
	let request;
	const fetchImpl = async (url, init) => {
		request = { url, init };
		return new Response(JSON.stringify({ number: 42, html_url: 'https://github.example/pull/42' }), { status: 201 });
	};
	const result = await callTool('github_create_pull_request', {
		owner: 'acme', repo: 'editor', title: 'Change', head: 'feature', base: 'main'
	}, {
		ENTERPRISE_MCP_ALLOW_WRITES: 'true', GITHUB_TOKEN: 'secret', GITHUB_API_URL: 'https://github.example/api/v3/'
	}, fetchImpl);
	assert.deepEqual(result, { number: 42, html_url: 'https://github.example/pull/42' });
	assert.equal(request.url, 'https://github.example/api/v3/repos/acme/editor/pulls');
	assert.equal(request.init.headers.Authorization, 'Bearer secret');
	assert.deepEqual(JSON.parse(request.init.body), { title: 'Change', head: 'feature', base: 'main', draft: false });
});

test('uses Jira basic authentication when an email is configured', async () => {
	let authorization;
	const fetchImpl = async (_url, init) => {
		authorization = init.headers.Authorization;
		return new Response(JSON.stringify({ key: 'ENG-1' }), { status: 200 });
	};
	await callTool('jira_get_issue', { issueKey: 'ENG-1' }, {
		JIRA_BASE_URL: 'https://acme.atlassian.net', JIRA_EMAIL: 'dev@acme.test', JIRA_API_TOKEN: 'token'
	}, fetchImpl);
	assert.equal(authorization, `Basic ${Buffer.from('dev@acme.test:token').toString('base64')}`);
});
