import type { INodeProperties } from 'n8n-workflow';

export const scheduledItemOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['scheduledItem'],
			},
		},
		options: [
			{
				name: 'Add',
				value: 'add',
				description: 'Create a new scheduled item',
				action: 'Add scheduled item',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get a list of scheduled items',
				action: 'Get many scheduled items',
			},
			{
				name: 'Remove',
				value: 'remove',
				description: 'Remove a scheduled item',
				action: 'Remove scheduled item',
			},
		],
		default: 'getAll',
	},
];

export const scheduledItemFields: INodeProperties[] = [
	// Get Many fields
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['scheduledItem'],
				operation: ['getAll'],
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
				resource: ['scheduledItem'],
				operation: ['getAll'],
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
				resource: ['scheduledItem'],
				operation: ['getAll'],
			},
		},
		options: [
			{
				displayName: 'Show Hidden',
				name: 'showHidden',
				type: 'boolean',
				default: false,
				description: 'Whether to show hidden system jobs',
			},
			{
				displayName: 'Filter by Name',
				name: 'name',
				type: 'string',
				default: '',
				description: 'Filter scheduled items by name',
			},
		],
	},

	// Add Scheduled Item fields
	{
		displayName: 'Tenant',
		name: 'tenantFilter',
		type: 'resourceLocator',
		default: { mode: 'list', value: '' },
		description: 'The tenant to run the job for (optional for all-tenant jobs)',
		displayOptions: {
			show: {
				resource: ['scheduledItem'],
				operation: ['add'],
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
	{
		displayName: 'Job Name',
		name: 'jobName',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['scheduledItem'],
				operation: ['add'],
			},
		},
		default: '',
		description: 'Name of the scheduled job',
	},
	{
		displayName: 'Command',
		name: 'command',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['scheduledItem'],
				operation: ['add'],
			},
		},
		default: '',
		placeholder: 'e.g. Get-CIPPUsers',
		description: 'The command to execute',
	},
	{
		displayName: 'Scheduled Time',
		name: 'scheduledTime',
		type: 'dateTime',
		displayOptions: {
			show: {
				resource: ['scheduledItem'],
				operation: ['add'],
			},
		},
		default: '',
		description: 'When the job should run',
	},
	{
		displayName: 'Recurrence',
		name: 'recurrence',
		type: 'options',
		displayOptions: {
			show: {
				resource: ['scheduledItem'],
				operation: ['add'],
			},
		},
		options: [
			{ name: 'Once', value: '0' },
			{ name: 'Daily', value: '1d' },
			{ name: 'Weekly', value: '7d' },
			{ name: 'Monthly', value: '30d' },
			{ name: 'Yearly', value: '365d' },
		],
		default: '0',
		description: 'How often the job should recur',
	},
	{
		displayName: 'Parameters',
		name: 'parameters',
		type: 'json',
		displayOptions: {
			show: {
				resource: ['scheduledItem'],
				operation: ['add'],
			},
		},
		default: '{}',
		description: 'JSON parameters for the command',
	},
	{
		displayName: 'Post-Execution Actions',
		name: 'postExecution',
		type: 'multiOptions',
		displayOptions: {
			show: {
				resource: ['scheduledItem'],
				operation: ['add'],
			},
		},
		options: [
			{ name: 'Webhook', value: 'Webhook' },
			{ name: 'Email', value: 'Email' },
			{ name: 'PSA', value: 'PSA' },
		],
		default: [],
		description: 'Actions to take after job execution',
	},

	// Remove Scheduled Item fields
	{
		displayName: 'Row Key',
		name: 'rowKey',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['scheduledItem'],
				operation: ['remove'],
			},
		},
		default: '',
		description: 'The RowKey of the scheduled item to remove',
	},
];

// Backup operations
export const backupOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['backup'],
			},
		},
		options: [
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get a list of backups',
				action: 'Get many backups',
			},
			{
				name: 'Restore',
				value: 'restore',
				description: 'Restore a backup',
				action: 'Restore backup',
			},
			{
				name: 'Run',
				value: 'run',
				description: 'Create a new backup',
				action: 'Run backup',
			},
			{
				name: 'Set Auto-Backup',
				value: 'setAutoBackup',
				description: 'Enable or disable automatic backups',
				action: 'Set auto backup',
			},
		],
		default: 'getAll',
	},
];

export const backupFields: INodeProperties[] = [
	// Get Many fields
	{
		displayName: 'Return All',
		name: 'returnAll',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['backup'],
				operation: ['getAll'],
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
				resource: ['backup'],
				operation: ['getAll'],
				returnAll: [false],
			},
		},
		typeOptions: {
			minValue: 1,
			maxValue: 100,
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
				resource: ['backup'],
				operation: ['getAll'],
			},
		},
		options: [
			{
				displayName: 'Names Only',
				name: 'namesOnly',
				type: 'boolean',
				default: false,
				description: 'Whether to return only backup names',
			},
			{
				displayName: 'Backup Name',
				name: 'backupName',
				type: 'string',
				default: '',
				description: 'Get a specific backup by name',
			},
		],
	},

	// Restore fields
	{
		displayName: 'Backup Data',
		name: 'backupData',
		type: 'json',
		required: true,
		displayOptions: {
			show: {
				resource: ['backup'],
				operation: ['restore'],
			},
		},
		default: '{}',
		description: 'The backup data to restore',
	},

	// Auto-Backup fields
	{
		displayName: 'Enable Auto-Backup',
		name: 'enableAutoBackup',
		type: 'boolean',
		required: true,
		displayOptions: {
			show: {
				resource: ['backup'],
				operation: ['setAutoBackup'],
			},
		},
		default: true,
		description: 'Whether to enable automatic backups',
	},
];

// Tools operations
export const toolsOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['tools'],
			},
		},
		options: [
			{
				name: 'Breach Search (Account)',
				value: 'breachAccount',
				description: 'Check breaches for an email or domain',
				action: 'Search breaches for account',
			},
			{
				name: 'Breach Search (Tenant)',
				value: 'breachTenant',
				description: 'Get breaches for a tenant',
				action: 'Search breaches for tenant',
			},
			{
				name: 'Execute Breach Search',
				value: 'executeBreachSearch',
				description: 'Execute a comprehensive breach search',
				action: 'Execute breach search',
			},
			{
				name: 'Exec Graph Request',
				value: 'execGraphRequest',
				description: 'Execute a Microsoft Graph request via POST /api/ExecGraphRequest',
				action: 'Exec graph request',
			},
			{
				name: 'Graph Request (Exec)',
				value: 'graphRequestExec',
				description: 'Execute Microsoft Graph GET/POST/PATCH requests via your CIPP fork',
				action: 'Execute graph request',
			},
			{
				name: 'Graph Request (List)',
				value: 'graphRequest',
				description: 'Make a custom Microsoft Graph GET/list request',
				action: 'Execute graph list request',
			},
		],
		default: 'graphRequest',
	},
];

export const toolsFields: INodeProperties[] = [
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
				resource: ['tools'],
				operation: [
					'breachTenant',
					'executeBreachSearch',
					'execGraphRequest',
					'graphRequest',
					'graphRequestExec',
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

	// Breach Account fields
	{
		displayName: 'Account',
		name: 'account',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['tools'],
				operation: ['breachAccount'],
			},
		},
		default: '',
		placeholder: 'user@domain.com or domain.com',
		description: 'The email address or domain to check for breaches',
	},

	// Graph Request fields
	{
		displayName: 'Endpoint',
		name: 'graphEndpoint',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['tools'],
				operation: ['graphRequest'],
			},
		},
		default: 'users',
		placeholder: 'e.g. users, groups, devices',
		description: 'The Graph API endpoint to call',
	},
	{
		displayName: 'Options',
		name: 'graphOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: {
			show: {
				resource: ['tools'],
				operation: ['graphRequest'],
			},
		},
		options: [
			{
				displayName: '$Count',
				name: 'count',
				type: 'boolean',
				default: false,
				description: 'Whether to include count',
			},
			{
				displayName: '$Filter',
				name: 'filter',
				type: 'string',
				default: '',
				placeholder: "startsWith(displayName,'John')",
				description: 'OData filter',
			},
			{
				displayName: '$Orderby',
				name: 'orderby',
				type: 'string',
				default: '',
				placeholder: 'displayName',
				description: 'Field to order by',
			},
			{
				displayName: '$Select',
				name: 'select',
				type: 'string',
				default: '',
				placeholder: 'ID,displayName,userPrincipalName',
				description: 'Fields to select',
			},
			{
				displayName: '$Top',
				name: 'top',
				type: 'number',
				default: 100,
				description: 'Number of records to return',
			},
		],
	},
	// Exec Graph Request fields
	{
		displayName: 'Endpoint',
		name: 'execEndpoint',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['tools'],
				operation: ['execGraphRequest'],
			},
		},
		default: '',
		placeholder: 'e.g. users, groups, devices',
		description: 'The Graph API endpoint to call',
	},
	{
		displayName: 'Method',
		name: 'execMethod',
		type: 'options',
		required: true,
		displayOptions: {
			show: {
				resource: ['tools'],
				operation: ['execGraphRequest'],
			},
		},
		options: [
			{ name: 'GET', value: 'GET' },
			{ name: 'PATCH', value: 'PATCH' },
			{ name: 'POST', value: 'POST' },
			{ name: 'DELETE', value: 'DELETE' },
		],
		default: 'GET',
		description: 'HTTP method for the Graph request',
	},
	{
		displayName: 'Body',
		name: 'execBody',
		type: 'json',
		displayOptions: {
			show: {
				resource: ['tools'],
				operation: ['execGraphRequest'],
				execMethod: ['POST', 'PATCH'],
			},
		},
		default: '{}',
		description: 'Request body as JSON',
	},

	// Graph Request (Exec) fields — Teams Shifts focused
	{
		displayName: 'Endpoint',
		name: 'graphExecEndpoint',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['tools'],
				operation: ['graphRequestExec'],
			},
		},
		default: 'teams/{team-ID}/schedule/shifts',
		placeholder: 'e.g. teams/{team-ID}/schedule/shifts',
		description:
			'Graph endpoint path to execute (relative path preferred, such as teams/{team-ID}/schedule/shifts)',
	},
	{
		displayName: 'Method',
		name: 'graphExecMethod',
		type: 'options',
		required: true,
		displayOptions: {
			show: {
				resource: ['tools'],
				operation: ['graphRequestExec'],
			},
		},
		options: [
			{
				name: 'GET',
				value: 'GET',
			},
			{
				name: 'PATCH',
				value: 'PATCH',
			},
			{
				name: 'POST',
				value: 'POST',
			},
		],
		default: 'GET',
		description: 'HTTP method for the Graph request',
	},
	{
		displayName: 'Headers',
		name: 'graphExecHeaders',
		type: 'json',
		displayOptions: {
			show: {
				resource: ['tools'],
				operation: ['graphRequestExec'],
			},
		},
		default: '{}',
		description: 'Optional Graph request headers as a JSON object',
	},
	{
		displayName: 'Body',
		name: 'graphExecBody',
		type: 'json',
		displayOptions: {
			show: {
				resource: ['tools'],
				operation: ['graphRequestExec'],
				graphExecMethod: ['POST', 'PATCH'],
			},
		},
		default: '{}',
		description: 'Graph request body as JSON',
	},
	{
		displayName: 'Exec Options',
		name: 'graphExecOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: {
			show: {
				resource: ['tools'],
				operation: ['graphRequestExec'],
			},
		},
		options: [
			{
				displayName: 'Enforce Teams Shifts Endpoint Pattern',
				name: 'enforceShiftsAllowlist',
				type: 'boolean',
				default: true,
				description: 'Whether to require an endpoint matching teams/{ID}/schedule/* before sending',
			},
			{
				displayName: 'Max Payload Bytes',
				name: 'maxPayloadBytes',
				type: 'number',
				typeOptions: {
					minValue: 1024,
					maxValue: 1048576,
				},
				default: 262144,
				description: 'Maximum serialized payload size in bytes sent to CIPP',
			},
		],
	},
];
