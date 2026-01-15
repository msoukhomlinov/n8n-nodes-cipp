import type { INodeProperties } from 'n8n-workflow';

export const applicationOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['application'],
			},
		},
		options: [
			{
				name: 'Add Chocolatey App',
				value: 'addChocolatey',
				description: 'Add a Chocolatey application',
				action: 'Add chocolatey app',
			},
			{
				name: 'Add MSP App',
				value: 'addMsp',
				description: 'Add an MSP/RMM application',
				action: 'Add MSP app',
			},
			{
				name: 'Add Office App',
				value: 'addOffice',
				description: 'Add Microsoft 365 Apps',
				action: 'Add office app',
			},
			{
				name: 'Add Store App',
				value: 'addStore',
				description: 'Add a Microsoft Store application',
				action: 'Add store app',
			},
			{
				name: 'Add WinGet App',
				value: 'addWinget',
				description: 'Add a WinGet application',
				action: 'Add win get app',
			},
			{
				name: 'Assign',
				value: 'assign',
				description: 'Assign an application to users or devices',
				action: 'Assign application',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get a list of applications',
				action: 'Get many applications',
			},
			{
				name: 'Get Queue',
				value: 'getQueue',
				description: 'Get the application deployment queue',
				action: 'Get application queue',
			},
			{
				name: 'Remove',
				value: 'remove',
				description: 'Remove an application',
				action: 'Remove application',
			},
			{
				name: 'Remove From Queue',
				value: 'removeFromQueue',
				description: 'Remove an application from the queue',
				action: 'Remove from queue',
			},
		],
		default: 'getAll',
	},
];

export const applicationFields: INodeProperties[] = [
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
				resource: ['application'],
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
				resource: ['application'],
				operation: ['getAll', 'getQueue'],
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
				resource: ['application'],
				operation: ['getAll', 'getQueue'],
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

	// Application ID for operations
	{
		displayName: 'Application ID',
		name: 'appId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['assign', 'remove'],
			},
		},
		default: '',
		description: 'The ID of the application',
	},

	// Queue Item ID
	{
		displayName: 'Queue Item ID',
		name: 'queueId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['removeFromQueue'],
			},
		},
		default: '',
		description: 'The ID of the queue item to remove',
	},

	// Assignment target
	{
		displayName: 'Assign To',
		name: 'assignTo',
		type: 'options',
		required: true,
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['assign', 'addWinget', 'addStore', 'addChocolatey', 'addMsp', 'addOffice'],
			},
		},
		options: [
			{ name: 'All Devices', value: 'AllDevices' },
			{ name: 'All Users', value: 'AllUsers' },
			{ name: 'Both', value: 'Both' },
			{ name: 'Custom Group', value: 'customGroup' },
			{ name: 'Do Not Assign', value: 'On' },
		],
		default: 'AllDevices',
		description: 'Who to assign the application to',
	},
	{
		displayName: 'Custom Group Names',
		name: 'customGroupNames',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['assign', 'addWinget', 'addStore', 'addChocolatey', 'addMsp'],
				assignTo: ['customGroup'],
			},
		},
		default: '',
		placeholder: 'Group1,Group2',
		description: 'Comma-separated list of group names to assign to',
	},

	// WinGet/Store App fields
	{
		displayName: 'Package ID',
		name: 'packageId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['addWinget', 'addStore'],
			},
		},
		default: '',
		placeholder: 'e.g. Google.Chrome or 9WZDNCRFJ3TJ',
		description: 'The WinGet package ID or Store product ID',
	},
	{
		displayName: 'Application Name',
		name: 'appName',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['addWinget', 'addStore', 'addChocolatey'],
			},
		},
		default: '',
		description: 'The display name for the application',
	},
	{
		displayName: 'Description',
		name: 'appDescription',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['addWinget', 'addStore', 'addChocolatey'],
			},
		},
		default: '',
		description: 'Description of the application',
	},
	{
		displayName: 'Mark for Uninstall',
		name: 'uninstall',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['addWinget', 'addStore', 'addChocolatey'],
			},
		},
		default: false,
		description: 'Whether to mark the app for uninstallation instead of installation',
	},

	// Chocolatey specific
	{
		displayName: 'Package Name',
		name: 'packageName',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['addChocolatey'],
			},
		},
		default: '',
		placeholder: 'e.g. googlechrome',
		description: 'The Chocolatey package name',
	},
	{
		displayName: 'Additional Options',
		name: 'chocoOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['addChocolatey'],
			},
		},
		options: [
			{
				displayName: 'Custom Repository URL',
				name: 'customRepo',
				type: 'string',
				default: '',
				description: 'Custom Chocolatey repository URL',
			},
			{
				displayName: 'Install as System',
				name: 'installAsSystem',
				type: 'boolean',
				default: true,
				description: 'Whether to install as SYSTEM user',
			},
			{
				displayName: 'Disable Restart',
				name: 'disableRestart',
				type: 'boolean',
				default: false,
				description: 'Whether to disable automatic restart',
			},
		],
	},

	// MSP App fields
	{
		displayName: 'RMM Tool',
		name: 'rmmTool',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['addMsp'],
			},
		},
		default: '',
		description: 'The RMM tool identifier',
	},
	{
		displayName: 'Display Name',
		name: 'mspDisplayName',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['addMsp'],
			},
		},
		default: '',
		description: 'The display name for the application',
	},
	{
		displayName: 'RMM Parameters',
		name: 'rmmParameters',
		type: 'json',
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['addMsp'],
			},
		},
		default: '{}',
		description: 'RMM-specific parameters as JSON',
	},

	// Office App fields
	{
		displayName: 'Excluded Apps',
		name: 'excludedApps',
		type: 'multiOptions',
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['addOffice'],
			},
		},
		options: [
			{ name: 'Access', value: 'access' },
			{ name: 'Excel', value: 'excel' },
			{ name: 'Groove', value: 'groove' },
			{ name: 'Lync', value: 'lync' },
			{ name: 'OneDrive', value: 'oneDrive' },
			{ name: 'OneNote', value: 'oneNote' },
			{ name: 'Outlook', value: 'outlook' },
			{ name: 'PowerPoint', value: 'powerPoint' },
			{ name: 'Publisher', value: 'publisher' },
			{ name: 'Teams', value: 'teams' },
			{ name: 'Word', value: 'word' },
		],
		default: [],
		description: 'Apps to exclude from the installation',
	},
	{
		displayName: 'Update Channel',
		name: 'updateChannel',
		type: 'options',
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['addOffice'],
			},
		},
		options: [
			{ name: 'Current', value: 'Current' },
			{ name: 'Monthly Enterprise', value: 'MonthlyEnterprise' },
			{ name: 'Semi-Annual', value: 'SemiAnnual' },
		],
		default: 'Current',
		description: 'The update channel for Office',
	},
	{
		displayName: 'Additional Options',
		name: 'officeOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: {
			show: {
				resource: ['application'],
				operation: ['addOffice'],
			},
		},
		options: [
			{
				displayName: 'Accept License',
				name: 'acceptLicense',
				type: 'boolean',
				default: true,
				description: 'Whether to auto-accept the license',
			},
			{
				displayName: 'Languages',
				name: 'languages',
				type: 'string',
				default: 'en-us',
				placeholder: 'en-us,es-es',
				description: 'Comma-separated list of language codes',
			},
			{
				displayName: 'Remove Other Versions',
				name: 'removeOtherVersions',
				type: 'boolean',
				default: true,
				description: 'Whether to remove other Office versions',
			},
			{
				displayName: 'Shared Computer Activation',
				name: 'sharedComputerActivation',
				type: 'boolean',
				default: false,
				description: 'Whether to enable shared computer activation',
			},
			{
				displayName: 'Use 64-Bit',
				name: 'use64bit',
				type: 'boolean',
				default: true,
				description: 'Whether to install 64-bit version',
			},
		],
	},
];
