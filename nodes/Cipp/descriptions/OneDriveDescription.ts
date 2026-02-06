import type { INodeProperties } from 'n8n-workflow';

export const onedriveOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['onedrive'],
			},
		},
		options: [
			{
				name: 'Add Shortcut',
				value: 'addShortcut',
				description: 'Add a OneDrive shortcut for a user',
				action: 'Add OneDrive shortcut',
			},
			{
				name: 'Provision',
				value: 'provision',
				description: 'Pre-provision OneDrive for a user',
				action: 'Provision OneDrive',
			},
		],
		default: 'provision',
	},
];

export const onedriveFields: INodeProperties[] = [
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
				resource: ['onedrive'],
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

	// User ID for all OneDrive operations
	{
		displayName: 'User ID',
		name: 'userId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['onedrive'],
			},
		},
		default: '',
		placeholder: 'user@domain.com',
		description: 'The User Principal Name (UPN) of the user',
	},

	// Shortcut URL for addShortcut
	{
		displayName: 'Shortcut URL',
		name: 'shortcutUrl',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['onedrive'],
				operation: ['addShortcut'],
			},
		},
		default: '',
		placeholder: 'e.g. https://contoso.sharepoint.com/sites/TeamSite',
		description: 'The SharePoint URL to create a shortcut to',
	},
	{
		displayName: 'Shortcut Name',
		name: 'shortcutName',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['onedrive'],
				operation: ['addShortcut'],
			},
		},
		default: '',
		placeholder: 'e.g. Team Documents',
		description: 'Display name for the shortcut (optional)',
	},
];
