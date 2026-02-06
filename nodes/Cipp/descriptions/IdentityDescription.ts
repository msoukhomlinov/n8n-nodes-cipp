import type { INodeProperties } from 'n8n-workflow';

export const identityOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['identity'],
			},
		},
		options: [
			{
				name: 'List Audit Logs',
				value: 'listAuditLogs',
				description: 'List audit logs for compliance monitoring',
				action: 'List audit logs',
			},
			{
				name: 'List Deleted Items',
				value: 'listDeletedItems',
				description: 'List deleted users, groups, and applications',
				action: 'List deleted items',
			},
			{
				name: 'List Roles',
				value: 'listRoles',
				description: 'List Azure AD roles and assignments',
				action: 'List roles',
			},
			{
				name: 'Restore Deleted',
				value: 'restoreDeleted',
				description: 'Restore a deleted object',
				action: 'Restore deleted object',
			},
		],
		default: 'listAuditLogs',
	},
];

export const identityFields: INodeProperties[] = [
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
				resource: ['identity'],
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

	// Return All for list operations
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['identity'],
				operation: ['listAuditLogs', 'listDeletedItems', 'listRoles'],
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
				resource: ['identity'],
				operation: ['listAuditLogs', 'listDeletedItems', 'listRoles'],
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

	// Object ID for restore
	{
		displayName: 'Object ID',
		name: 'objectId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['identity'],
				operation: ['restoreDeleted'],
			},
		},
		default: '',
		placeholder: 'e.g. 12345678-1234-1234-1234-123456789abc',
		description: 'The ID of the deleted object to restore',
	},
];
