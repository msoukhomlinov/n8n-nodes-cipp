import type { INodeProperties } from 'n8n-workflow';

export const alertOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['alert'],
			},
		},
		options: [
			{
				name: 'Add',
				value: 'add',
				description: 'Create a new alert rule',
				action: 'Add an alert',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get alerts from the queue',
				action: 'Get many alerts',
			},
			{
				name: 'Get Security Alerts',
				value: 'getSecurityAlerts',
				description: 'Get Defender security alerts',
				action: 'Get security alerts',
			},
			{
				name: 'Get Security Incidents',
				value: 'getSecurityIncidents',
				action: 'Get security incidents',
			},
			{
				name: 'Set Security Alert Status',
				value: 'setSecurityAlertStatus',
				description: 'Update the status of a security alert',
				action: 'Set security alert status',
			},
			{
				name: 'Set Security Incident Status',
				value: 'setSecurityIncidentStatus',
				description: 'Update the status of a security incident',
				action: 'Set security incident status',
			},
		],
		default: 'getAll',
	},
];

export const alertFields: INodeProperties[] = [
	// Tenant selector for security operations
	{
		displayName: 'Tenant',
		name: 'tenantFilter',
		type: 'resourceLocator',
		default: { mode: 'list', value: '' },
		required: true,
		description: 'The tenant to perform the operation on',
		displayOptions: {
			show: {
				resource: ['alert'],
				operation: [
					'getSecurityAlerts',
					'getSecurityIncidents',
					'setSecurityAlertStatus',
					'setSecurityIncidentStatus',
				],
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
				resource: ['alert'],
				operation: ['getAll', 'getSecurityAlerts', 'getSecurityIncidents'],
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
				resource: ['alert'],
				operation: ['getAll', 'getSecurityAlerts', 'getSecurityIncidents'],
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

	// Add Alert fields
	{
		displayName: 'Alert Configuration',
		name: 'alertConfig',
		type: 'json',
		required: true,
		displayOptions: {
			show: {
				resource: ['alert'],
				operation: ['add'],
			},
		},
		default: '{\n  "tenants": [],\n  "excludedTenants": [],\n  "logSource": "",\n  "conditions": {},\n  "actions": []\n}',
		description: 'JSON configuration for the alert rule',
	},

	// Security Alert Status
	{
		displayName: 'Alert ID',
		name: 'alertId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['alert'],
				operation: ['setSecurityAlertStatus'],
			},
		},
		default: '',
		description: 'The ID of the security alert',
	},
	{
		displayName: 'Status',
		name: 'alertStatus',
		type: 'options',
		required: true,
		displayOptions: {
			show: {
				resource: ['alert'],
				operation: ['setSecurityAlertStatus'],
			},
		},
		options: [
			{ name: 'In Progress', value: 'inProgress' },
			{ name: 'Resolved', value: 'resolved' },
		],
		default: 'resolved',
		description: 'The new status for the alert',
	},
	{
		displayName: 'Additional Fields',
		name: 'alertAdditionalFields',
		type: 'collection',
		placeholder: 'Add Field',
		default: {},
		displayOptions: {
			show: {
				resource: ['alert'],
				operation: ['setSecurityAlertStatus'],
			},
		},
		options: [
			{
				displayName: 'Vendor Name',
				name: 'vendorName',
				type: 'string',
				default: '',
				description: 'The vendor name for the alert',
			},
			{
				displayName: 'Provider Name',
				name: 'providerName',
				type: 'string',
				default: '',
				description: 'The provider name for the alert',
			},
		],
	},

	// Security Incident Status
	{
		displayName: 'Incident ID',
		name: 'incidentId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['alert'],
				operation: ['setSecurityIncidentStatus'],
			},
		},
		default: '',
		description: 'The ID of the security incident',
	},
	{
		displayName: 'Status',
		name: 'incidentStatus',
		type: 'options',
		required: true,
		displayOptions: {
			show: {
				resource: ['alert'],
				operation: ['setSecurityIncidentStatus'],
			},
		},
		options: [
			{ name: 'Active', value: 'active' },
			{ name: 'In Progress', value: 'inProgress' },
			{ name: 'Resolved', value: 'resolved' },
		],
		default: 'resolved',
		description: 'The new status for the incident',
	},
	{
		displayName: 'Assign To',
		name: 'assignedTo',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['alert'],
				operation: ['setSecurityIncidentStatus'],
			},
		},
		default: '',
		placeholder: 'user@domain.com',
		description: 'The user to assign the incident to',
	},
];
