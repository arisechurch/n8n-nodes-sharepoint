import { INodeType, INodeTypeDescription } from 'n8n-workflow'
import { execute, listOptions, run, siteOptions } from './GenericFunctions'

export class MicrosoftSharepoint implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Microsoft Sharepoint',
		name: 'microsoftSharepoint',
		group: ['transform'],
		version: 1,
		description: 'Microsoft Sharepoint API',
		defaults: {
			name: 'Microsoft Sharepoint',
		},
		inputs: ['main'],
		outputs: ['main'],
		credentials: [
			{
				name: 'microsoftSharepointApi',
				required: true,
			},
		],
		properties: [
			{
				displayName: 'Sharepoint Site Name or ID',
				name: 'site',
				type: 'options',
				description: 'Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code-examples/expressions/">expression</a>',
				required: true,
				typeOptions: {
					loadOptionsMethod: 'getSites',
				},
				default: '',
			},
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				default: 'lists',
				options: [
					{
						name: 'List',
						value: 'lists',
					},
				],
			},
			{
				displayName: 'List Name or ID',
				name: 'list',
				type: 'options',
				description: 'Choose from the list, or specify an ID using an <a href="https://docs.n8n.io/code-examples/expressions/">expression</a>',
				noDataExpression: true,
				displayOptions: {
					hide: {
						site: [''],
					},
					show: {
						resource: ['lists'],
					},
				},
				typeOptions: {
					loadOptionsMethod: 'getLists',
				},
				default: '',
			},
			{
				displayName: 'Method',
				name: 'method',
				type: 'options',
				noDataExpression: true,
				default: 'GET',
				options: [
					{
						name: 'GET',
						value: 'GET',
					},
					{
						name: 'POST',
						value: 'POST',
					},
					{
						name: 'PATCH',
						value: 'PATCH',
					},
					{
						name: 'PUT',
						value: 'PUT',
					},
				],
			},
			{
				displayName: 'Path',
				name: 'path',
				description: 'URL path',
				type: 'string',
				required: true,
				default: '/',
			},
			{
				displayName: 'Body',
				name: 'body',
				description: 'JSON body to send',
				type: 'json',
				displayOptions: {
					hide: {
						method: ['GET'],
					},
				},
				default: '',
			},
		],
	}

	methods = {
		loadOptions: {
			getSites: execute(siteOptions),
			getLists: execute(listOptions),
		},
	}

	// The function below is responsible for actually doing whatever this node
	// is supposed to do. In this case, we're just appending the `myString` property
	// with whatever the user has entered.
	// You can make async calls and use `await`.
	execute = execute(run.map(_ => [[_]]))
}
