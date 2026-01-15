import type { INodeProperties } from 'n8n-workflow';

export const deviceOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['device'],
			},
		},
		options: [
			{
				name: 'Execute Action',
				value: 'executeAction',
				description: 'Execute an action on a device (reboot, wipe, sync, etc.)',
				action: 'Execute device action',
			},
			{
				name: 'Get LAPS Password',
				value: 'getLapsPassword',
				description: 'Get the Local Admin Password (LAPS)',
				action: 'Get LAPS password',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get a list of devices',
				action: 'Get many devices',
			},
			{
				name: 'Get Recovery Key',
				value: 'getRecoveryKey',
				description: 'Get BitLocker recovery keys',
				action: 'Get recovery key',
			},
			{
				name: 'Manage',
				value: 'manage',
				description: 'Enable, disable, or delete a device',
				action: 'Manage device',
			},
		],
		default: 'getAll',
	},
];

export const deviceFields: INodeProperties[] = [
	// Tenant selector for all device operations
	{
		displayName: 'Tenant',
		name: 'tenantFilter',
		type: 'resourceLocator',
		default: { mode: 'list', value: '' },
		required: true,
		description: 'The tenant to perform the operation on',
		displayOptions: {
			show: {
				resource: ['device'],
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
				resource: ['device'],
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
				resource: ['device'],
				operation: ['getAll'],
				returnAll: [false],
			},
		},
		typeOptions: {
			minValue: 1,
			maxValue: 999,
		},
		default: 50,
		description: 'Max number of results to return',
	},

	// Device ID for single-device operations
	{
		displayName: 'Device ID',
		name: 'deviceId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['device'],
				operation: ['manage', 'executeAction', 'getRecoveryKey', 'getLapsPassword'],
			},
		},
		default: '',
		description: 'The ID of the device',
	},

	// Manage Device action
	{
		displayName: 'Action',
		name: 'manageAction',
		type: 'options',
		required: true,
		displayOptions: {
			show: {
				resource: ['device'],
				operation: ['manage'],
			},
		},
		options: [
			{ name: 'Enable', value: 'Enable' },
			{ name: 'Disable', value: 'Disable' },
			{ name: 'Delete', value: 'Delete' },
		],
		default: 'Enable',
	},

	// Execute Action options
	{
		displayName: 'Action',
		name: 'executeDeviceAction',
		type: 'options',
		required: true,
		displayOptions: {
			show: {
				resource: ['device'],
				operation: ['executeAction'],
			},
		},
		options: [
			{ name: 'Fresh Start', value: 'FreshStart' },
			{ name: 'Full Scan', value: 'FullScan' },
			{ name: 'Quick Scan', value: 'QuickScan' },
			{ name: 'Reboot', value: 'Reboot' },
			{ name: 'Remove From Autopilot', value: 'RemoveFromAutopilot' },
			{ name: 'Remove From Azure AD', value: 'RemoveFromAzure' },
			{ name: 'Remove From Defender', value: 'RemoveFromDefender' },
			{ name: 'Remove From Everywhere', value: 'RemoveFromEverywhere' },
			{ name: 'Remove From Intune', value: 'RemoveFromIntune' },
			{ name: 'Rename', value: 'Rename' },
			{ name: 'Reset to Autopilot', value: 'Autopilot' },
			{ name: 'Sync Device', value: 'SyncDevice' },
			{ name: 'Windows Update - All', value: 'WindowsUpdateAll' },
			{ name: 'Windows Update - Drivers', value: 'WindowsUpdateDrivers' },
			{ name: 'Windows Update - Features', value: 'WindowsUpdateFeatures' },
			{ name: 'Windows Update - Other', value: 'WindowsUpdateOther' },
			{ name: 'Windows Update - Reboot', value: 'WindowsUpdateReboot' },
			{ name: 'Windows Update - Scan', value: 'WindowsUpdateScan' },
			{ name: 'Wipe', value: 'Wipe' },
		],
		default: 'SyncDevice',
		description: 'The action to execute on the device',
	},

	// Rename device name
	{
		displayName: 'New Device Name',
		name: 'newDeviceName',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['device'],
				operation: ['executeAction'],
				executeDeviceAction: ['Rename'],
			},
		},
		default: '',
		description: 'The new name for the device',
	},
];

// Autopilot operations
export const autopilotOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['autopilot'],
			},
		},
		options: [
			{
				name: 'Assign',
				value: 'assign',
				description: 'Assign an Autopilot device to a user',
				action: 'Assign autopilot device',
			},
			{
				name: 'Get Configurations',
				value: 'getConfigurations',
				description: 'Get Autopilot configurations (ESP or profiles)',
				action: 'Get autopilot configurations',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get a list of Autopilot devices',
				action: 'Get many autopilot devices',
			},
			{
				name: 'Remove',
				value: 'remove',
				description: 'Remove an Autopilot device',
				action: 'Remove autopilot device',
			},
			{
				name: 'Sync',
				value: 'sync',
				description: 'Sync Autopilot devices',
				action: 'Sync autopilot devices',
			},
		],
		default: 'getAll',
	},
];

export const autopilotFields: INodeProperties[] = [
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
				resource: ['autopilot'],
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
				resource: ['autopilot'],
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
				resource: ['autopilot'],
				operation: ['getAll'],
				returnAll: [false],
			},
		},
		typeOptions: {
			minValue: 1,
			maxValue: 999,
		},
		default: 50,
		description: 'Max number of results to return',
	},

	// Device ID for single operations
	{
		displayName: 'Device ID',
		name: 'deviceId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['autopilot'],
				operation: ['assign', 'remove'],
			},
		},
		default: '',
		description: 'The ID of the Autopilot device',
	},

	// Assign fields
	{
		displayName: 'Serial Number',
		name: 'serialNumber',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['autopilot'],
				operation: ['assign'],
			},
		},
		default: '',
		description: 'The serial number of the device',
	},
	{
		displayName: 'User Principal Name',
		name: 'userPrincipalName',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['autopilot'],
				operation: ['assign'],
			},
		},
		default: '',
		placeholder: 'user@domain.com',
		description: 'The UPN of the user to assign the device to',
	},

	// Get Configurations type
	{
		displayName: 'Configuration Type',
		name: 'configType',
		type: 'options',
		displayOptions: {
			show: {
				resource: ['autopilot'],
				operation: ['getConfigurations'],
			},
		},
		options: [
			{ name: 'Enrollment Status Page', value: 'ESP' },
			{ name: 'Autopilot Profile', value: 'ApProfile' },
		],
		default: 'ApProfile',
		description: 'The type of configuration to retrieve',
	},
];
