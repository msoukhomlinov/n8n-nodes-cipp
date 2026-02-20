import type { INodeProperties } from 'n8n-workflow';

export const quarantineOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['quarantine'],
			},
		},
		options: [
			{
				name: 'Deny',
				value: 'deny',
				description: 'Deny a quarantined message and prevent delivery',
				action: 'Deny quarantined message',
			},
			{
				name: 'Get Many',
				value: 'getMany',
				description: 'List quarantined email messages',
				action: 'List quarantined messages',
			},
			{
				name: 'Get Message',
				value: 'getMessage',
				description: 'Get the body content of a quarantined message',
				action: 'Get quarantined message body',
			},
			{
				name: 'Release',
				value: 'release',
				description: 'Release a quarantined message for delivery',
				action: 'Release quarantined message',
			},
		],
		default: 'getMany',
	},
];

export const quarantineFields: INodeProperties[] = [
	// Tenant selector for all quarantine operations
	{
		displayName: 'Tenant',
		name: 'tenantFilter',
		type: 'resourceLocator',
		default: { mode: 'list', value: '' },
		required: true,
		description: 'The tenant to perform the operation on',
		displayOptions: {
			show: {
				resource: ['quarantine'],
			},
		},
		modes: [
			{
				displayName: 'From List',
				name: 'list',
				type: 'list',
				typeOptions: {
					searchListMethod: 'tenantSearch',
					searchable: true,
				},
			},
			{
				displayName: 'By Domain',
				name: 'domain',
				type: 'string',
				placeholder: 'e.g. contoso.onmicrosoft.com',
			},
		],
	},

	// Message ID for release/deny operations
	{
		displayName: 'Message ID',
		name: 'messageId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['quarantine'],
				operation: ['release', 'deny', 'getMessage'],
			},
		},
		default: '',
		placeholder: 'e.g. abc123def456...',
		description: 'The unique identifier of the quarantined message',
	},

	// Allow Sender option for release operation
	{
		displayName: 'Allow Sender',
		name: 'allowSender',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['quarantine'],
				operation: ['release'],
			},
		},
		default: false,
		description: 'Whether to add the sender to the allowed list when releasing the message',
	},

	// Return All option for getMany
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['quarantine'],
				operation: ['getMany'],
			},
		},
		default: false,
		description: 'Whether to return all results or only up to a given limit',
	},
	{
		displayName: 'Limit',
		name: 'limit',
		type: 'number',
		displayOptions: {
			show: {
				resource: ['quarantine'],
				operation: ['getMany'],
				returnAll: [false],
			},
		},
		typeOptions: {
			minValue: 1,
			maxValue: 500,
		},
		default: 50,
		description: 'Max number of results to return',
	},
];
