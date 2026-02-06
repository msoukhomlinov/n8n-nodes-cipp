import type { INodeProperties } from 'n8n-workflow';

export const tenantOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['tenant'],
			},
		},
		options: [
			{
				name: 'Clear Cache',
				value: 'clearCache',
				description: 'Clear the tenant cache in CIPP',
				action: 'Clear tenant cache',
			},
			{
				name: 'CSP License Action',
				value: 'cspLicenseAction',
				description: 'Add or remove CSP licenses for a tenant',
				action: 'Csp license action',
			},
			{
				name: 'Get CSP Licenses',
				value: 'getCspLicenses',
				description: 'Get CSP license information for a tenant',
				action: 'Get CSP licenses',
			},
			{
				name: 'Get Licenses',
				value: 'getLicenses',
				description: 'Get license information and usage for a tenant',
				action: 'Get licenses',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get a list of many managed tenants',
				action: 'Get many tenants',
			},
			{
				name: 'List CSP SKUs',
				value: 'listCspSkus',
				description: 'List available CSP license SKUs catalog',
				action: 'List CSP SKUs',
			},
			{
				name: 'List Defender State',
				value: 'listDefenderState',
				description: 'Get Defender security posture for a tenant',
				action: 'List Defender state',
			},
		],
		default: 'getAll',
	},
];

export const tenantFields: INodeProperties[] = [
	// Tenant selector for tenant-specific operations
	{
		displayName: 'Tenant',
		name: 'tenantFilter',
		type: 'resourceLocator',
		default: { mode: 'list', value: '' },
		required: true,
		description: 'The tenant to perform the operation on',
		displayOptions: {
			show: {
				resource: ['tenant'],
				operation: ['getLicenses', 'getCspLicenses', 'cspLicenseAction', 'listDefenderState', 'listCspSkus'],
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

	// Get Many options
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['tenant'],
				operation: ['getAll', 'getLicenses', 'getCspLicenses', 'listDefenderState', 'listCspSkus'],
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
				resource: ['tenant'],
				operation: ['getAll', 'getLicenses', 'getCspLicenses', 'listDefenderState', 'listCspSkus'],
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
	{
		displayName: 'Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: {
			show: {
				resource: ['tenant'],
				operation: ['getAll'],
			},
		},
		options: [
			{
				displayName: 'Include All Tenants Option',
				name: 'allTenantSelector',
				type: 'boolean',
				default: false,
				description: 'Whether to include the "All Tenants" option in the results',
			},
		],
	},
	{
		displayName: 'Options',
		name: 'licenseOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: {
			show: {
				resource: ['tenant'],
				operation: ['getLicenses'],
			},
		},
		options: [
			{
				displayName: 'Summary Only',
				name: 'summaryOnly',
				type: 'boolean',
				default: false,
				description: 'Whether to return only license counts without the full list of assigned users and groups (faster)',
			},
		],
	},

	// Clear Cache options
	{
		displayName: 'Clear Tenant Cache Only',
		name: 'clearCacheTenantOnly',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['tenant'],
				operation: ['clearCache'],
			},
		},
		default: false,
		description: 'Whether to only clear the tenant cache (not all caches)',
	},

	// CSP License Action fields
	{
		displayName: 'Action',
		name: 'cspAction',
		type: 'options',
		displayOptions: {
			show: {
				resource: ['tenant'],
				operation: ['cspLicenseAction'],
			},
		},
		options: [
			{ name: 'Add Licenses', value: 'add' },
			{ name: 'Remove Licenses', value: 'remove' },
		],
		default: 'add',
		description: 'Whether to add or remove licenses',
	},
	{
		displayName: 'License SKU',
		name: 'licenseSku',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['tenant'],
				operation: ['cspLicenseAction'],
			},
		},
		default: '',
		placeholder: 'e.g. O365_BUSINESS_PREMIUM',
		description: 'The license SKU to add or remove',
	},
	{
		displayName: 'Quantity',
		name: 'licenseQuantity',
		type: 'number',
		required: true,
		displayOptions: {
			show: {
				resource: ['tenant'],
				operation: ['cspLicenseAction'],
			},
		},
		typeOptions: {
			minValue: 1,
		},
		default: 1,
		description: 'Number of licenses to add or remove',
	},
];
