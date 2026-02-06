import type { INodeProperties } from 'n8n-workflow';

export const gdapOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['gdap'],
			},
		},
		options: [
			{
				name: 'List Roles',
				value: 'listRoles',
				description: 'List GDAP roles and assignments',
				action: 'List GDAP roles',
			},
			{
				name: 'Send Invite',
				value: 'sendInvite',
				description: 'Send a GDAP invitation to a tenant',
				action: 'Send GDAP invite',
			},
		],
		default: 'listRoles',
	},
];

export const gdapFields: INodeProperties[] = [
	// Tenant selector
	{
		displayName: 'Tenant',
		name: 'tenantFilter',
		type: 'resourceLocator',
		default: { mode: 'list', value: '' },
		required: true,
		description: 'The tenant to perform the operation on',
		displayOptions: {
			show: {
				resource: ['gdap'],
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

	// Return All for list operation
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['gdap'],
				operation: ['listRoles'],
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
				resource: ['gdap'],
				operation: ['listRoles'],
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

	// Invite fields
	{
		displayName: 'Roles',
		name: 'gdapRoles',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['gdap'],
				operation: ['sendInvite'],
			},
		},
		default: '',
		placeholder: 'e.g. Exchange Administrator, Security Reader',
		description: 'Comma-separated list of GDAP roles to include in the invite',
	},
];
