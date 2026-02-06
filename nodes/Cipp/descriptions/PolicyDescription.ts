import type { INodeProperties } from 'n8n-workflow';

export const policyOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['policy'],
			},
		},
		options: [
			{
				name: 'Add',
				value: 'add',
				description: 'Add an Intune policy',
				action: 'Add policy',
			},
			{
				name: 'Assign',
				value: 'assign',
				description: 'Assign a policy to users or devices',
				action: 'Assign policy',
			},
			{
				name: 'Get Many',
				value: 'getMany',
				description: 'List Intune policies',
				action: 'Get many policies',
			},
			{
				name: 'List Defender TVM',
				value: 'listDefenderTvm',
				description: 'List Defender Threat and Vulnerability Management data',
				action: 'List Defender TVM',
			},
			{
				name: 'Remove',
				value: 'remove',
				description: 'Remove an Intune policy',
				action: 'Remove policy',
			},
		],
		default: 'getMany',
	},
];

export const policyFields: INodeProperties[] = [
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
				resource: ['policy'],
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
				resource: ['policy'],
				operation: ['getMany', 'listDefenderTvm'],
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
				resource: ['policy'],
				operation: ['getMany', 'listDefenderTvm'],
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

	// Policy ID for single-policy operations
	{
		displayName: 'Policy ID',
		name: 'policyId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['policy'],
				operation: ['assign', 'remove'],
			},
		},
		default: '',
		description: 'The ID of the policy',
	},

	// Assignment target for assign operation
	{
		displayName: 'Assign To',
		name: 'assignTo',
		type: 'options',
		displayOptions: {
			show: {
				resource: ['policy'],
				operation: ['assign'],
			},
		},
		options: [
			{ name: 'All Users', value: 'allUsers' },
			{ name: 'All Devices', value: 'allDevices' },
			{ name: 'Custom Group', value: 'customGroup' },
		],
		default: 'allUsers',
		description: 'Target for policy assignment',
	},
	{
		displayName: 'Group Names',
		name: 'customGroupNames',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['policy'],
				operation: ['assign'],
				assignTo: ['customGroup'],
			},
		},
		default: '',
		placeholder: 'e.g. Group1, Group2',
		description: 'Comma-separated list of group names to assign the policy to',
	},

	// Add policy fields
	{
		displayName: 'Policy Configuration',
		name: 'policyConfig',
		type: 'json',
		required: true,
		displayOptions: {
			show: {
				resource: ['policy'],
				operation: ['add'],
			},
		},
		default: '{}',
		description: 'JSON configuration for the new policy',
	},
];
