import type { INodeProperties } from 'n8n-workflow';

export const userOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['user'],
			},
		},
		options: [
			{
				name: 'Add',
				value: 'add',
				description: 'Create a new user',
				action: 'Add a user',
			},
			{
				name: 'Clear Immutable ID',
				value: 'clearImmutableId',
				description: 'Clear the immutable ID for a user',
				action: 'Clear immutable ID',
			},
			{
				name: 'Create TAP',
				value: 'createTap',
				description: 'Create a Temporary Access Password',
				action: 'Create temporary access password',
			},
			{
				name: 'Disable',
				value: 'disable',
				description: 'Block sign-in for a user',
				action: 'Disable a user',
			},
			{
				name: 'Enable',
				value: 'enable',
				description: 'Unblock sign-in for a user',
				action: 'Enable a user',
			},
			{
				name: 'Get Many',
				value: 'getAll',
				description: 'Get a list of users',
				action: 'Get many users',
			},
			{
				name: 'Offboard',
				value: 'offboard',
				description: 'Offboard a user with all offboarding tasks',
				action: 'Offboard a user',
			},
			{
				name: 'Remove',
				value: 'remove',
				description: 'Delete a user',
				action: 'Remove a user',
			},
			{
				name: 'Reset MFA',
				value: 'resetMfa',
				description: 'Re-require MFA registration for a user',
				action: 'Reset MFA',
			},
			{
				name: 'Reset Password',
				value: 'resetPassword',
				description: 'Reset a user password',
				action: 'Reset password',
			},
			{
				name: 'Revoke Sessions',
				value: 'revokeSessions',
				description: 'Revoke all active sessions',
				action: 'Revoke sessions',
			},
			{
				name: 'Send MFA Push',
				value: 'sendMfaPush',
				description: 'Send an MFA push notification',
				action: 'Send MFA push',
			},
			{
				name: 'Set Per-User MFA',
				value: 'setPerUserMfa',
				description: 'Set per-user MFA state',
				action: 'Set per user mfa',
			},
		],
		default: 'getAll',
	},
];

export const userFields: INodeProperties[] = [
	// Tenant selector for all user operations
	{
		displayName: 'Tenant',
		name: 'tenantFilter',
		type: 'resourceLocator',
		default: { mode: 'list', value: '' },
		required: true,
		description: 'The tenant to perform the operation on',
		displayOptions: {
			show: {
				resource: ['user'],
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
				hint: 'Enter the tenant default domain',
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
				resource: ['user'],
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
				resource: ['user'],
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

	// User ID for single-user operations
	{
		displayName: 'User ID',
		name: 'userId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['user'],
				operation: [
					'disable',
					'enable',
					'resetMfa',
					'resetPassword',
					'revokeSessions',
					'remove',
					'clearImmutableId',
					'createTap',
					'sendMfaPush',
					'setPerUserMfa',
				],
			},
		},
		default: '',
		placeholder: 'user@domain.com or GUID',
		description: 'The User Principal Name (UPN) or Object ID of the user',
	},

	// Per-User MFA state
	{
		displayName: 'MFA State',
		name: 'mfaState',
		type: 'options',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['setPerUserMfa'],
			},
		},
		options: [
			{ name: 'Enforced', value: 'Enforced' },
			{ name: 'Enabled', value: 'Enabled' },
			{ name: 'Disabled', value: 'Disabled' },
		],
		default: 'Enforced',
		description: 'The MFA state to set for the user',
	},

	// Reset Password options
	{
		displayName: 'Additional Options',
		name: 'passwordOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['resetPassword'],
			},
		},
		options: [
			{
				displayName: 'Must Change Password',
				name: 'mustChangePass',
				type: 'boolean',
				default: true,
				description: 'Whether the user must change password at next logon',
			},
		],
	},

	// Add User fields
	{
		displayName: 'First Name',
		name: 'firstName',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['add'],
			},
		},
		default: '',
		description: 'The first name of the user',
	},
	{
		displayName: 'Last Name',
		name: 'lastName',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['add'],
			},
		},
		default: '',
		description: 'The last name of the user',
	},
	{
		displayName: 'Domain',
		name: 'domain',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['add'],
			},
		},
		default: '',
		placeholder: 'e.g. contoso.com',
		description: 'The primary domain for the user email',
	},
	{
		displayName: 'Additional Fields',
		name: 'additionalFields',
		type: 'collection',
		placeholder: 'Add Field',
		default: {},
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['add'],
			},
		},
		options: [
			{
				displayName: 'Display Name',
				name: 'displayName',
				type: 'string',
				default: '',
				description: 'Custom display name (defaults to First Last)',
			},
			{
				displayName: 'Mail Nickname',
				name: 'mailNickname',
				type: 'string',
				default: '',
				description: 'Mail alias (defaults to first.last)',
			},
			{
				displayName: 'Usage Location',
				name: 'usageLocation',
				type: 'string',
				default: 'US',
				placeholder: 'e.g. US, GB, DE',
				description: 'ISO country code for license assignment',
			},
		],
	},

	// Offboard User fields
	{
		displayName: 'Users to Offboard',
		name: 'usersToOffboard',
		type: 'json',
		required: true,
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['offboard'],
			},
		},
		default: '[]',
		description: 'JSON array of user objects to offboard',
	},
	{
		displayName: 'Scheduled Offboard',
		name: 'scheduledOffboard',
		type: 'boolean',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['offboard'],
			},
		},
		default: false,
		description: 'Whether to schedule the offboarding for later',
	},

	// Get Many - Fields to return
	{
		displayName: 'Fields to Return',
		name: 'userFields',
		type: 'multiOptions',
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['getAll'],
			},
		},
		options: [
			{
				name: 'Account Enabled',
				value: 'accountEnabled',
				description: 'Whether the account is enabled',
			},
			{
				name: 'Assigned Licenses',
				value: 'assignedLicenses',
				description: 'Licenses assigned to the user',
			},
			{
				name: 'City',
				value: 'city',
				description: 'City from address',
			},
			{
				name: 'Company Name',
				value: 'companyName',
				description: 'Company name',
			},
			{
				name: 'Country',
				value: 'country',
				description: 'Country/region',
			},
			{
				name: 'Created Date',
				value: 'createdDateTime',
				description: 'When the user was created',
			},
			{
				name: 'Department',
				value: 'department',
				description: 'Department name',
			},
			{
				name: 'Display Name',
				value: 'displayName',
				description: 'Display name of the user',
			},
			{
				name: 'Employee ID',
				value: 'employeeId',
				description: 'Employee identifier',
			},
			{
				name: 'First Name',
				value: 'givenName',
				description: 'First/given name',
			},
			{
				name: 'ID',
				value: 'id',
				description: 'Unique identifier (GUID)',
			},
			{
				name: 'Job Title',
				value: 'jobTitle',
				description: 'Job title',
			},
			{
				name: 'Last Name',
				value: 'surname',
				description: 'Last/family name',
			},
			{
				name: 'Last Password Change',
				value: 'lastPasswordChangeDateTime',
				description: 'When password was last changed',
			},
			{
				name: 'License Details',
				value: 'licenseAssignmentStates',
				description: 'Details about license assignments',
			},
			{
				name: 'Mail',
				value: 'mail',
				description: 'Primary email address',
			},
			{
				name: 'Manager',
				value: 'manager',
				description: 'User manager',
			},
			{
				name: 'Mobile Phone',
				value: 'mobilePhone',
				description: 'Mobile phone number',
			},
			{
				name: 'Office Location',
				value: 'officeLocation',
				description: 'Office location',
			},
			{
				name: 'On-Premises Sync',
				value: 'onPremisesSyncEnabled',
				description: 'Whether synced from on-premises AD',
			},
			{
				name: 'Phone Number',
				value: 'businessPhones',
				description: 'Business phone numbers',
			},
			{
				name: 'Proxy Addresses',
				value: 'proxyAddresses',
				description: 'All email addresses including aliases',
			},
			{
				name: 'Sign-In Activity',
				value: 'signInActivity',
				description: 'Last sign-in date/time (requires Azure AD Premium)',
			},
			{
				name: 'State',
				value: 'state',
				description: 'State or province',
			},
			{
				name: 'Street Address',
				value: 'streetAddress',
				description: 'Street address',
			},
			{
				name: 'Usage Location',
				value: 'usageLocation',
				description: 'Country code for license assignment',
			},
			{
				name: 'User Principal Name',
				value: 'userPrincipalName',
				description: 'Sign-in name (email format)',
			},
			{
				name: 'User Type',
				value: 'userType',
				description: 'Member or Guest',
			},
		],
		default: ['id', 'displayName', 'userPrincipalName', 'mail', 'accountEnabled'],
		description: 'Select which user properties to return. Limiting fields improves performance.',
	},

	// Get Many filters
	{
		displayName: 'Options',
		name: 'filters',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: {
			show: {
				resource: ['user'],
				operation: ['getAll'],
			},
		},
		options: [
			{
				displayName: 'Additional Fields',
				name: 'select',
				type: 'string',
				default: '',
				placeholder: 'e.g. otherMails,employeeType',
				description: 'Additional fields not in the list above (comma-separated)',
			},
			{
				displayName: 'Filter Query',
				name: 'filter',
				type: 'string',
				default: '',
				placeholder: "e.g. startsWith(displayName,'John')",
				description: 'OData filter query to filter which users are returned',
			},
		],
	},
];
