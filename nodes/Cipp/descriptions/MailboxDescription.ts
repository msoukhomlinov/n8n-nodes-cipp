import type { INodeProperties } from 'n8n-workflow';

export const mailboxOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['mailbox'],
			},
		},
		options: [
			{
				name: 'Convert',
				value: 'convert',
				description: 'Convert mailbox between shared and regular types',
				action: 'Convert mailbox',
			},
			{
				name: 'Enable Archive',
				value: 'enableArchive',
				description: 'Enable online archive for a user',
				action: 'Enable archive',
			},
			{
				name: 'Set Email Forwarding',
				value: 'setForwarding',
				description: 'Manage email forwarding settings',
				action: 'Set email forwarding',
			},
			{
				name: 'Set Out of Office',
				value: 'setOutOfOffice',
				description: 'Set or disable out of office message',
				action: 'Set out of office',
			},
		],
		default: 'convert',
	},
];

export const mailboxFields: INodeProperties[] = [
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
				resource: ['mailbox'],
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

	// User ID for all mailbox operations
	{
		displayName: 'User ID',
		name: 'userId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['mailbox'],
			},
		},
		default: '',
		placeholder: 'user@domain.com',
		description: 'The User Principal Name (UPN) of the mailbox owner',
	},

	// Convert Mailbox
	{
		displayName: 'Mailbox Type',
		name: 'mailboxType',
		type: 'options',
		required: true,
		displayOptions: {
			show: {
				resource: ['mailbox'],
				operation: ['convert'],
			},
		},
		options: [
			{ name: 'Shared', value: 'Shared' },
			{ name: 'Regular', value: 'Regular' },
		],
		default: 'Shared',
		description: 'The type to convert the mailbox to',
	},

	// Out of Office
	{
		displayName: 'Auto-Reply State',
		name: 'autoReplyState',
		type: 'options',
		required: true,
		displayOptions: {
			show: {
				resource: ['mailbox'],
				operation: ['setOutOfOffice'],
			},
		},
		options: [
			{ name: 'Enabled', value: 'Enabled' },
			{ name: 'Disabled', value: 'Disabled' },
		],
		default: 'Enabled',
		description: 'Whether to enable or disable the auto-reply',
	},
	{
		displayName: 'Auto-Reply Message',
		name: 'autoReplyMessage',
		type: 'string',
		typeOptions: {
			rows: 4,
		},
		displayOptions: {
			show: {
				resource: ['mailbox'],
				operation: ['setOutOfOffice'],
				autoReplyState: ['Enabled'],
			},
		},
		default: '',
		description: 'The out of office message to set',
	},

	// Email Forwarding
	{
		displayName: 'Forward To',
		name: 'forwardTo',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['mailbox'],
				operation: ['setForwarding'],
			},
		},
		default: '',
		placeholder: 'forward@domain.com',
		description: 'The email address to forward to (leave empty to disable)',
	},
	{
		displayName: 'Keep Copy',
		name: 'keepCopy',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['mailbox'],
				operation: ['setForwarding'],
			},
		},
		default: true,
		description: 'Whether to keep a copy of forwarded messages in the mailbox',
	},
];
