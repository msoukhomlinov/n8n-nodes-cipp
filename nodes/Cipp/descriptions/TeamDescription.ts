import type { INodeProperties } from 'n8n-workflow';

export const teamOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['team'],
			},
		},
		options: [
			{
				name: 'Add',
				value: 'add',
				description: 'Create a new Microsoft Team',
				action: 'Add a team',
			},
			{
				name: 'Add SharePoint Site',
				value: 'addSite',
				description: 'Create a new SharePoint site',
				action: 'Add share point site',
			},
			{
				name: 'Add Sites Bulk',
				value: 'addSitesBulk',
				description: 'Create multiple SharePoint sites from a list',
				action: 'Add sites bulk',
			},
			{
				name: 'Get Activity',
				value: 'getActivity',
				description: 'Get Teams user activity',
				action: 'Get teams activity',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get a list of Teams',
				action: 'Get many teams',
			},
			{
				name: 'Get Sites',
				value: 'getSites',
				description: 'Get SharePoint sites or OneDrive accounts',
				action: 'Get share point sites',
			},
			{
				name: 'Manage Site Member',
				value: 'manageSiteMember',
				description: 'Add or remove a SharePoint site member',
				action: 'Manage site member',
			},
			{
				name: 'Manage Site Permissions',
				value: 'manageSitePermissions',
				description: 'Manage SharePoint or OneDrive permissions',
				action: 'Manage site permissions',
			},
		],
		default: 'getAll',
	},
];

export const teamFields: INodeProperties[] = [
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
				resource: ['team'],
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

	// Get Many fields
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['getAll', 'getSites', 'getActivity'],
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
				resource: ['team'],
				operation: ['getAll', 'getSites', 'getActivity'],
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

	// Get Sites type
	{
		displayName: 'Site Type',
		name: 'siteType',
		type: 'options',
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['getSites'],
			},
		},
		options: [
			{ name: 'SharePoint Sites', value: 'SharePointSiteUsage' },
			{ name: 'OneDrive Accounts', value: 'OneDriveUsageAccount' },
		],
		default: 'SharePointSiteUsage',
		description: 'The type of sites to list',
	},

	// Add Team fields
	{
		displayName: 'Team Name',
		name: 'displayName',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['add'],
			},
		},
		default: '',
		description: 'The display name of the team',
	},
	{
		displayName: 'Description',
		name: 'teamDescription',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['add'],
			},
		},
		default: '',
		description: 'Description of the team',
	},
	{
		displayName: 'Owner',
		name: 'owner',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['add'],
			},
		},
		default: '',
		placeholder: 'owner@domain.com',
		description: 'The UPN of the team owner',
	},
	{
		displayName: 'Visibility',
		name: 'visibility',
		type: 'options',
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['add'],
			},
		},
		options: [
			{ name: 'Public', value: 'public' },
			{ name: 'Private', value: 'private' },
		],
		default: 'private',
		description: 'The visibility of the team',
	},

	// Add SharePoint Site fields
	{
		displayName: 'Site Name',
		name: 'siteName',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['addSite'],
			},
		},
		default: '',
		description: 'The name of the SharePoint site',
	},
	{
		displayName: 'Site Description',
		name: 'siteDescription',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['addSite'],
			},
		},
		default: '',
		description: 'Description of the site',
	},
	{
		displayName: 'Site Owner',
		name: 'siteOwner',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['addSite'],
			},
		},
		default: '',
		placeholder: 'owner@domain.com',
		description: 'The UPN of the site owner',
	},
	{
		displayName: 'Template',
		name: 'templateName',
		type: 'options',
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['addSite'],
			},
		},
		options: [
			{ name: 'Team Site', value: 'team' },
			{ name: 'Communication Site', value: 'communication' },
		],
		default: 'team',
		description: 'The site template to use',
	},
	{
		displayName: 'Site Design',
		name: 'siteDesign',
		type: 'options',
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['addSite'],
				templateName: ['communication'],
			},
		},
		options: [
			{ name: 'Blank', value: 'blank' },
			{ name: 'Showcase', value: 'Showcase' },
			{ name: 'Topic', value: 'Topic' },
		],
		default: 'blank',
		description: 'The site design template',
	},

	// Manage Site Member fields
	{
		displayName: 'Site URL',
		name: 'siteUrl',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['manageSiteMember', 'manageSitePermissions'],
			},
		},
		default: '',
		placeholder: 'https://contoso.sharepoint.com/sites/sitename',
		description: 'The URL of the SharePoint site',
	},
	{
		displayName: 'User',
		name: 'siteUser',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['manageSiteMember', 'manageSitePermissions'],
			},
		},
		default: '',
		placeholder: 'user@domain.com',
		description: 'The UPN of the user to add/remove',
	},
	{
		displayName: 'Action',
		name: 'memberAction',
		type: 'options',
		required: true,
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['manageSiteMember'],
			},
		},
		options: [
			{ name: 'Add Member', value: 'add' },
			{ name: 'Remove Member', value: 'remove' },
		],
		default: 'add',
		description: 'Whether to add or remove the user',
	},
	{
		displayName: 'Remove Permission',
		name: 'removePermission',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['manageSitePermissions'],
			},
		},
		default: false,
		description: 'Whether to remove (true) or add (false) the permission',
	},

	// Add Sites Bulk field
	{
		displayName: 'Sites Configuration',
		name: 'sitesConfig',
		type: 'json',
		required: true,
		displayOptions: {
			show: {
				resource: ['team'],
				operation: ['addSitesBulk'],
			},
		},
		default: '[]',
		description: 'JSON array of site configurations to create in bulk',
	},
];

// Voice operations
export const voiceOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['voice'],
			},
		},
		options: [
			{
				name: 'Assign Number',
				value: 'assignNumber',
				description: 'Assign a phone number to a user',
				action: 'Assign phone number',
			},
			{
				name: 'Get Emergency Locations',
				value: 'getLocations',
				description: 'Get Teams emergency locations',
				action: 'Get emergency locations',
			},
			{
				name: 'Get Phone Numbers',
				value: 'getPhoneNumbers',
				description: 'Get Teams voice phone numbers',
				action: 'Get phone numbers',
			},
			{
				name: 'Unassign Number',
				value: 'unassignNumber',
				description: 'Unassign a phone number from a user',
				action: 'Unassign phone number',
			},
		],
		default: 'getPhoneNumbers',
	},
];

export const voiceFields: INodeProperties[] = [
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
				resource: ['voice'],
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

	// Get Many fields
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['voice'],
				operation: ['getPhoneNumbers', 'getLocations'],
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
				resource: ['voice'],
				operation: ['getPhoneNumbers', 'getLocations'],
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

	// Assign/Unassign fields
	{
		displayName: 'Phone Number',
		name: 'phoneNumber',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['voice'],
				operation: ['assignNumber', 'unassignNumber'],
			},
		},
		default: '',
		placeholder: '+1234567890',
		description: 'The phone number to assign/unassign',
	},
	{
		displayName: 'User',
		name: 'voiceUser',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['voice'],
				operation: ['assignNumber', 'unassignNumber'],
			},
		},
		default: '',
		placeholder: 'user@domain.com',
		description: 'The UPN of the user',
	},
	{
		displayName: 'Phone Number Type',
		name: 'phoneNumberType',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['voice'],
				operation: ['assignNumber', 'unassignNumber'],
			},
		},
		default: '',
		description: 'The type of phone number',
	},
	{
		displayName: 'Location Only',
		name: 'locationOnly',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['voice'],
				operation: ['assignNumber'],
			},
		},
		default: false,
		description: 'Whether to only set emergency location without assigning to user',
	},
];
