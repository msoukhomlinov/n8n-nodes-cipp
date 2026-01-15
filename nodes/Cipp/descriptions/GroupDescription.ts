import type { INodeProperties } from 'n8n-workflow';

export const groupOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['group'],
			},
		},
		options: [
			{
				name: 'Add',
				value: 'add',
				description: 'Create a new group',
				action: 'Add a group',
			},
			{
				name: 'Delete',
				value: 'delete',
				description: 'Delete a group',
				action: 'Delete a group',
			},
			{
				name: 'Edit Members',
				value: 'edit',
				description: 'Add or remove members/owners from a group',
				action: 'Edit group members',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get a list of groups',
				action: 'Get many groups',
			},
			{
				name: 'Hide From GAL',
				value: 'hideFromGal',
				description: 'Hide or unhide a group from the Global Address List',
				action: 'Hide group from GAL',
			},
			{
				name: 'Set Delivery Management',
				value: 'deliveryManagement',
				description: 'Manage group delivery settings',
				action: 'Set delivery management',
			},
		],
		default: 'getAll',
	},
];

export const groupFields: INodeProperties[] = [
	// Tenant selector for all group operations
	{
		displayName: 'Tenant',
		name: 'tenantFilter',
		type: 'resourceLocator',
		default: { mode: 'list', value: '' },
		required: true,
		description: 'The tenant to perform the operation on',
		displayOptions: {
			show: {
				resource: ['group'],
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
				resource: ['group'],
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
				resource: ['group'],
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
	{
		displayName: 'Options',
		name: 'options',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: {
			show: {
				resource: ['group'],
				operation: ['getAll'],
			},
		},
		options: [
			{
				displayName: 'Group ID',
				name: 'groupId',
				type: 'string',
				default: '',
				description: 'Get details for a specific group by ID',
			},
			{
				displayName: 'Include Members',
				name: 'members',
				type: 'boolean',
				default: false,
				description: 'Whether to include group members in the response',
			},
			{
				displayName: 'Include Owners',
				name: 'owners',
				type: 'boolean',
				default: false,
				description: 'Whether to include group owners in the response',
			},
		],
	},

	// Add Group fields
	{
		displayName: 'Group Name',
		name: 'groupName',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['group'],
				operation: ['add'],
			},
		},
		default: '',
		description: 'The display name of the group',
	},
	{
		displayName: 'Group Type',
		name: 'groupType',
		type: 'options',
		required: true,
		displayOptions: {
			show: {
				resource: ['group'],
				operation: ['add'],
			},
		},
		options: [
			{ name: 'Microsoft 365', value: 'Microsoft 365' },
			{ name: 'Security', value: 'Security' },
			{ name: 'Distribution', value: 'Distribution' },
			{ name: 'Mail-Enabled Security', value: 'Mail-Enabled Security' },
		],
		default: 'Microsoft 365',
		description: 'The type of group to create',
	},

	// Edit Group fields
	{
		displayName: 'Group ID',
		name: 'groupId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['group'],
				operation: ['edit', 'delete', 'hideFromGal', 'deliveryManagement'],
			},
		},
		default: '',
		description: 'The ID of the group to modify',
	},
	{
		displayName: 'Edit Options',
		name: 'editOptions',
		type: 'collection',
		placeholder: 'Add Edit Option',
		default: {},
		displayOptions: {
			show: {
				resource: ['group'],
				operation: ['edit'],
			},
		},
		options: [
			{
				displayName: 'Add Members',
				name: 'addMembers',
				type: 'string',
				default: '',
				placeholder: 'user1@domain.com,user2@domain.com',
				description: 'Comma-separated list of users to add as members',
			},
			{
				displayName: 'Add Owners',
				name: 'addOwners',
				type: 'string',
				default: '',
				placeholder: 'owner@domain.com',
				description: 'Comma-separated list of users to add as owners',
			},
			{
				displayName: 'Remove Members',
				name: 'removeMembers',
				type: 'string',
				default: '',
				placeholder: 'user1@domain.com,user2@domain.com',
				description: 'Comma-separated list of members to remove',
			},
			{
				displayName: 'Remove Owners',
				name: 'removeOwners',
				type: 'string',
				default: '',
				placeholder: 'owner@domain.com',
				description: 'Comma-separated list of owners to remove',
			},
		],
	},

	// Delete Group fields
	{
		displayName: 'Group Type',
		name: 'groupTypeForDelete',
		type: 'options',
		displayOptions: {
			show: {
				resource: ['group'],
				operation: ['delete', 'hideFromGal', 'deliveryManagement'],
			},
		},
		options: [
			{ name: 'Microsoft 365', value: 'Microsoft 365' },
			{ name: 'Security', value: 'Security' },
			{ name: 'Distribution', value: 'Distribution' },
			{ name: 'Mail-Enabled Security', value: 'Mail-Enabled Security' },
		],
		default: 'Microsoft 365',
		description: 'The type of the group',
	},
	{
		displayName: 'Group Email',
		name: 'groupEmail',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['group'],
				operation: ['hideFromGal', 'deliveryManagement'],
			},
		},
		default: '',
		description: 'The email address of the group',
	},

	// Hide from GAL option
	{
		displayName: 'Hide From GAL',
		name: 'hideFromGal',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['group'],
				operation: ['hideFromGal'],
			},
		},
		default: true,
		description: 'Whether to hide the group from the Global Address List',
	},

	// Delivery Management option
	{
		displayName: 'Only Allow Internal Messages',
		name: 'onlyInternal',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['group'],
				operation: ['deliveryManagement'],
			},
		},
		default: false,
		description: 'Whether to only allow messages from internal senders',
	},
];
