import type { INodeProperties } from 'n8n-workflow';

// ────────────────── Shared helpers ──────────────────

const tenantSelector: INodeProperties = {
	displayName: 'Tenant',
	name: 'tenantFilter',
	type: 'resourceLocator',
	default: { mode: 'list', value: '' },
	required: true,
	description: 'The tenant to perform the operation on',
	displayOptions: {
		show: {
			resource: ['teamsShift'],
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
};

const teamIdField: INodeProperties = {
	displayName: 'Team ID',
	name: 'teamId',
	type: 'string',
	required: true,
	displayOptions: {
		show: {
			resource: ['teamsShift'],
		},
	},
	default: '',
	placeholder: 'e.g. 9f4808c6-e2ff-4c26-9154-bcaacc11ca95',
	description: 'The ID of the team whose schedule to manage',
};

// ────────────────── Operations ──────────────────

export const shiftOperations: INodeProperties[] = [
	{
		displayName: 'Operation',
		name: 'operation',
		type: 'options',
		noDataExpression: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
			},
		},
		options: [
			// ── Shifts ──
			{
				name: 'List Shifts',
				value: 'listShifts',
				description: 'List all shifts in a team schedule',
				action: 'List shifts',
			},
			{
				name: 'Create Shift',
				value: 'createShift',
				description: 'Create a new shift',
				action: 'Create shift',
			},
			{
				name: 'Update Shift',
				value: 'updateShift',
				description: 'Update an existing shift',
				action: 'Update shift',
			},
			{
				name: 'Delete Shift',
				value: 'deleteShift',
				description: 'Delete a shift',
				action: 'Delete shift',
			},
			// ── Open Shifts ──
			{
				name: 'List Open Shifts',
				value: 'listOpenShifts',
				description: 'List all open (unassigned) shifts',
				action: 'List open shifts',
			},
			{
				name: 'Create Open Shift',
				value: 'createOpenShift',
				description: 'Create a new open shift',
				action: 'Create open shift',
			},
			{
				name: 'Update Open Shift',
				value: 'updateOpenShift',
				description: 'Update an existing open shift',
				action: 'Update open shift',
			},
			{
				name: 'Delete Open Shift',
				value: 'deleteOpenShift',
				description: 'Delete an open shift',
				action: 'Delete open shift',
			},
			// ── Scheduling Groups ──
			{
				name: 'List Scheduling Groups',
				value: 'listSchedulingGroups',
				description: 'List scheduling groups',
				action: 'List scheduling groups',
			},
			{
				name: 'Create Scheduling Group',
				value: 'createSchedulingGroup',
				description: 'Create a scheduling group',
				action: 'Create scheduling group',
			},
			{
				name: 'Update Scheduling Group',
				value: 'updateSchedulingGroup',
				description: 'Update a scheduling group',
				action: 'Update scheduling group',
			},
			{
				name: 'Delete Scheduling Group',
				value: 'deleteSchedulingGroup',
				description: 'Mark a scheduling group as inactive',
				action: 'Delete scheduling group',
			},
			// ── Time Off Reasons ──
			{
				name: 'List Time Off Reasons',
				value: 'listTimeOffReasons',
				description: 'List time-off reasons',
				action: 'List time off reasons',
			},
			{
				name: 'Create Time Off Reason',
				value: 'createTimeOffReason',
				description: 'Create a time-off reason',
				action: 'Create time off reason',
			},
			{
				name: 'Update Time Off Reason',
				value: 'updateTimeOffReason',
				description: 'Update a time-off reason',
				action: 'Update time off reason',
			},
			{
				name: 'Delete Time Off Reason',
				value: 'deleteTimeOffReason',
				description: 'Mark a time-off reason as inactive',
				action: 'Delete time off reason',
			},
			// ── Time Off Requests ──
			{
				name: 'List Time Off Requests',
				value: 'listTimeOffRequests',
				description: 'List time-off requests',
				action: 'List time off requests',
			},
			{
				name: 'Create Time Off Request',
				value: 'createTimeOffRequest',
				description: 'Create a time-off request',
				action: 'Create time off request',
			},
			{
				name: 'Approve Time Off Request',
				value: 'approveTimeOffRequest',
				description: 'Approve a time-off request',
				action: 'Approve time off request',
			},
			{
				name: 'Decline Time Off Request',
				value: 'declineTimeOffRequest',
				description: 'Decline a time-off request',
				action: 'Decline time off request',
			},
			// ── Swap Shift Requests ──
			{
				name: 'List Swap Shift Requests',
				value: 'listSwapShiftRequests',
				description: 'List swap-shift change requests',
				action: 'List swap shift requests',
			},
			{
				name: 'Create Swap Shift Request',
				value: 'createSwapShiftRequest',
				description: 'Create a swap-shift change request',
				action: 'Create swap shift request',
			},
			{
				name: 'Approve Swap Shift Request',
				value: 'approveSwapShiftRequest',
				description: 'Approve a swap-shift change request',
				action: 'Approve swap shift request',
			},
			{
				name: 'Decline Swap Shift Request',
				value: 'declineSwapShiftRequest',
				description: 'Decline a swap-shift change request',
				action: 'Decline swap shift request',
			},
			// ── Offer Shift Requests ──
			{
				name: 'List Offer Shift Requests',
				value: 'listOfferShiftRequests',
				description: 'List offer-shift requests',
				action: 'List offer shift requests',
			},
			{
				name: 'Create Offer Shift Request',
				value: 'createOfferShiftRequest',
				description: 'Create an offer-shift request',
				action: 'Create offer shift request',
			},
			{
				name: 'Approve Offer Shift Request',
				value: 'approveOfferShiftRequest',
				description: 'Approve an offer-shift request',
				action: 'Approve offer shift request',
			},
			{
				name: 'Decline Offer Shift Request',
				value: 'declineOfferShiftRequest',
				description: 'Decline an offer-shift request',
				action: 'Decline offer shift request',
			},
		],
		default: 'listShifts',
	},
];

// ────────────────── Fields ──────────────────

export const shiftFields: INodeProperties[] = [
	// Shared: Tenant & Team ID
	tenantSelector,
	teamIdField,

	// Filters for all list operations
	{
		displayName: 'Filters',
		name: 'listFilters',
		type: 'collection',
		placeholder: 'Add Filter',
		default: {},
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: [
					'listShifts',
					'listOpenShifts',
					'listSchedulingGroups',
					'listTimeOffReasons',
					'listTimeOffRequests',
					'listSwapShiftRequests',
					'listOfferShiftRequests',
				],
			},
		},
		options: [
			{
				displayName: 'Start Date',
				name: 'startDate',
				type: 'dateTime',
				default: '',
				description: 'Filter results starting on or after this date/time (ISO 8601). Applied as sharedShift/startDateTime ge for shifts.',
			},
			{
				displayName: 'End Date',
				name: 'endDate',
				type: 'dateTime',
				default: '',
				description: 'Filter results ending on or before this date/time (ISO 8601). Applied as sharedShift/endDateTime le for shifts.',
			},
			{
				displayName: '$Filter (Raw OData)',
				name: 'rawFilter',
				type: 'string',
				default: '',
				placeholder: "sharedShift/startDateTime ge 2024-01-01T00:00:00Z",
				description: 'Raw OData $filter expression. Overrides Start/End Date if set.',
			},
		],
	},

	// ════════════════ SHIFT FIELDS ════════════════

	// Shift ID — for update / delete
	{
		displayName: 'Shift ID',
		name: 'shiftId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['updateShift', 'deleteShift'],
			},
		},
		default: '',
		description: 'The ID of the shift to update or delete',
	},

	// User ID — for create / update shift
	{
		displayName: 'User ID',
		name: 'userId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createShift'],
			},
		},
		default: '',
		placeholder: 'e.g. user-object-id or UPN',
		description: 'The user to assign the shift to',
	},

	// Start Date Time — shift create / update
	{
		displayName: 'Start Date Time',
		name: 'startDateTime',
		type: 'dateTime',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createShift'],
			},
		},
		default: '',
		description: 'Start date and time of the shift (ISO 8601)',
	},

	// End Date Time — shift create / update
	{
		displayName: 'End Date Time',
		name: 'endDateTime',
		type: 'dateTime',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createShift'],
			},
		},
		default: '',
		description: 'End date and time of the shift (ISO 8601)',
	},

	// Shift create options
	{
		displayName: 'Options',
		name: 'shiftOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createShift'],
			},
		},
		options: [
			{
				displayName: 'Display Name',
				name: 'displayName',
				type: 'string',
				default: '',
				description: 'The shift display name / label',
			},
			{
				displayName: 'Notes',
				name: 'notes',
				type: 'string',
				default: '',
				description: 'Notes for the shift visible to the user',
			},
			{
				displayName: 'Theme',
				name: 'theme',
				type: 'options',
				options: [
					{ name: 'White', value: 'white' },
					{ name: 'Blue', value: 'blue' },
					{ name: 'Green', value: 'green' },
					{ name: 'Purple', value: 'purple' },
					{ name: 'Pink', value: 'pink' },
					{ name: 'Yellow', value: 'yellow' },
					{ name: 'Gray', value: 'gray' },
					{ name: 'Dark Blue', value: 'darkBlue' },
					{ name: 'Dark Green', value: 'darkGreen' },
					{ name: 'Dark Purple', value: 'darkPurple' },
					{ name: 'Dark Pink', value: 'darkPink' },
					{ name: 'Dark Yellow', value: 'darkYellow' },
				],
				default: 'white',
				description: 'The shift color theme',
			},
			{
				displayName: 'Activities',
				name: 'activities',
				type: 'json',
				default: '[]',
				description: 'JSON array of shift activities [{displayName, startDateTime, endDateTime, code, theme}]',
			},
		],
	},

	// Shift update body (JSON) — flexible for PUT
	{
		displayName: 'Shift Data',
		name: 'shiftUpdateData',
		type: 'json',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['updateShift'],
			},
		},
		default: '{}',
		description: 'Full shift object JSON for the PUT update. Must include sharedShift with startDateTime, endDateTime, etc.',
	},

	// ════════════════ OPEN SHIFT FIELDS ════════════════

	{
		displayName: 'Open Shift ID',
		name: 'openShiftId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['updateOpenShift', 'deleteOpenShift'],
			},
		},
		default: '',
		description: 'The ID of the open shift',
	},

	// Open shift create fields
	{
		displayName: 'Scheduling Group ID',
		name: 'schedulingGroupId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createOpenShift'],
			},
		},
		default: '',
		description: 'The scheduling group the open shift belongs to',
	},
	{
		displayName: 'Start Date Time',
		name: 'openShiftStart',
		type: 'dateTime',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createOpenShift'],
			},
		},
		default: '',
		description: 'Start date and time of the open shift',
	},
	{
		displayName: 'End Date Time',
		name: 'openShiftEnd',
		type: 'dateTime',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createOpenShift'],
			},
		},
		default: '',
		description: 'End date and time of the open shift',
	},
	{
		displayName: 'Open Slot Count',
		name: 'openSlotCount',
		type: 'number',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createOpenShift'],
			},
		},
		default: 1,
		description: 'Number of slots available for this open shift',
	},
	{
		displayName: 'Options',
		name: 'openShiftOptions',
		type: 'collection',
		placeholder: 'Add Option',
		default: {},
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createOpenShift'],
			},
		},
		options: [
			{
				displayName: 'Display Name',
				name: 'displayName',
				type: 'string',
				default: '',
				description: 'Display name for the open shift',
			},
			{
				displayName: 'Notes',
				name: 'notes',
				type: 'string',
				default: '',
				description: 'Notes for the open shift',
			},
			{
				displayName: 'Theme',
				name: 'theme',
				type: 'options',
				options: [
					{ name: 'White', value: 'white' },
					{ name: 'Blue', value: 'blue' },
					{ name: 'Green', value: 'green' },
					{ name: 'Purple', value: 'purple' },
					{ name: 'Pink', value: 'pink' },
					{ name: 'Yellow', value: 'yellow' },
					{ name: 'Gray', value: 'gray' },
					{ name: 'Dark Blue', value: 'darkBlue' },
					{ name: 'Dark Green', value: 'darkGreen' },
					{ name: 'Dark Purple', value: 'darkPurple' },
					{ name: 'Dark Pink', value: 'darkPink' },
					{ name: 'Dark Yellow', value: 'darkYellow' },
				],
				default: 'white',
				description: 'The open shift color theme',
			},
			{
				displayName: 'Activities',
				name: 'activities',
				type: 'json',
				default: '[]',
				description: 'JSON array of activities',
			},
		],
	},

	// Open shift update (JSON)
	{
		displayName: 'Open Shift Data',
		name: 'openShiftUpdateData',
		type: 'json',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['updateOpenShift'],
			},
		},
		default: '{}',
		description: 'Full open shift object JSON for the PUT update',
	},

	// ════════════════ SCHEDULING GROUP FIELDS ════════════════

	{
		displayName: 'Scheduling Group ID',
		name: 'schedulingGroupUpdateId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['updateSchedulingGroup', 'deleteSchedulingGroup'],
			},
		},
		default: '',
		description: 'The ID of the scheduling group',
	},
	{
		displayName: 'Display Name',
		name: 'groupDisplayName',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createSchedulingGroup'],
			},
		},
		default: '',
		description: 'Display name for the scheduling group',
	},
	{
		displayName: 'User IDs',
		name: 'groupUserIds',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createSchedulingGroup'],
			},
		},
		default: '',
		placeholder: 'user-id-1, user-id-2',
		description: 'Comma-separated list of user IDs to include in the group',
	},

	// Scheduling group update (JSON)
	{
		displayName: 'Scheduling Group Data',
		name: 'schedulingGroupUpdateData',
		type: 'json',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['updateSchedulingGroup'],
			},
		},
		default: '{}',
		description: 'Full scheduling group JSON for the PUT update',
	},

	// ════════════════ TIME OFF REASON FIELDS ════════════════

	{
		displayName: 'Time Off Reason ID',
		name: 'timeOffReasonId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['updateTimeOffReason', 'deleteTimeOffReason'],
			},
		},
		default: '',
		description: 'The ID of the time-off reason',
	},
	{
		displayName: 'Display Name',
		name: 'reasonDisplayName',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createTimeOffReason'],
			},
		},
		default: '',
		placeholder: 'e.g. Vacation, Sick Leave',
		description: 'Display name for the time-off reason',
	},
	{
		displayName: 'Icon Type',
		name: 'iconType',
		type: 'options',
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createTimeOffReason'],
			},
		},
		options: [
			{ name: 'None', value: 'none' },
			{ name: 'Car', value: 'car' },
			{ name: 'Calendar', value: 'calendar' },
			{ name: 'Running', value: 'running' },
			{ name: 'Plane', value: 'plane' },
			{ name: 'First Aid', value: 'firstAid' },
			{ name: 'Doctor', value: 'doctor' },
			{ name: 'Not Working', value: 'notWorking' },
			{ name: 'Clock', value: 'clock' },
			{ name: 'Jury Duty', value: 'juryDuty' },
			{ name: 'Globe', value: 'globe' },
			{ name: 'Cup', value: 'cup' },
			{ name: 'Phone', value: 'phone' },
			{ name: 'Weather', value: 'weather' },
			{ name: 'Umbrella', value: 'umbrella' },
			{ name: 'Piggy Bank', value: 'piggyBank' },
			{ name: 'Dog', value: 'dog' },
			{ name: 'Cake', value: 'cake' },
			{ name: 'Traffic Cone', value: 'trafficCone' },
			{ name: 'Pin', value: 'pin' },
			{ name: 'Sunny', value: 'sunny' },
		],
		default: 'none',
		description: 'Icon associated with the time-off reason',
	},

	// Time off reason update (JSON)
	{
		displayName: 'Time Off Reason Data',
		name: 'timeOffReasonUpdateData',
		type: 'json',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['updateTimeOffReason'],
			},
		},
		default: '{}',
		description: 'Full time-off reason JSON for the PUT update',
	},

	// ════════════════ TIME OFF REQUEST FIELDS ════════════════

	{
		displayName: 'Time Off Request ID',
		name: 'timeOffRequestId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['approveTimeOffRequest', 'declineTimeOffRequest'],
			},
		},
		default: '',
		description: 'The ID of the time-off request',
	},
	{
		displayName: 'Start Date Time',
		name: 'timeOffStart',
		type: 'dateTime',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createTimeOffRequest'],
			},
		},
		default: '',
		description: 'Start date and time for the time off',
	},
	{
		displayName: 'End Date Time',
		name: 'timeOffEnd',
		type: 'dateTime',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createTimeOffRequest'],
			},
		},
		default: '',
		description: 'End date and time for the time off',
	},
	{
		displayName: 'Time Off Reason ID',
		name: 'timeOffReasonIdForRequest',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createTimeOffRequest'],
			},
		},
		default: '',
		description: 'The ID of the time-off reason for this request',
	},
	{
		displayName: 'Message',
		name: 'approvalMessage',
		type: 'string',
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: [
					'approveTimeOffRequest',
					'declineTimeOffRequest',
					'approveSwapShiftRequest',
					'declineSwapShiftRequest',
					'approveOfferShiftRequest',
					'declineOfferShiftRequest',
				],
			},
		},
		default: '',
		description: 'Optional message for the approval/decline',
	},

	// ════════════════ SWAP SHIFT REQUEST FIELDS ════════════════

	{
		displayName: 'Swap Shift Request ID',
		name: 'swapShiftRequestId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['approveSwapShiftRequest', 'declineSwapShiftRequest'],
			},
		},
		default: '',
		description: 'The ID of the swap-shift change request',
	},
	{
		displayName: 'Sender Shift ID',
		name: 'senderShiftId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createSwapShiftRequest'],
			},
		},
		default: '',
		description: 'The shift ID of the sender (your shift)',
	},
	{
		displayName: 'Recipient Shift ID',
		name: 'recipientShiftId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createSwapShiftRequest'],
			},
		},
		default: '',
		description: 'The shift ID of the recipient (the shift you want)',
	},
	{
		displayName: 'Recipient User ID',
		name: 'swapRecipientUserId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createSwapShiftRequest'],
			},
		},
		default: '',
		description: 'The user ID of the person you want to swap with',
	},

	// ════════════════ OFFER SHIFT REQUEST FIELDS ════════════════

	{
		displayName: 'Offer Shift Request ID',
		name: 'offerShiftRequestId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['approveOfferShiftRequest', 'declineOfferShiftRequest'],
			},
		},
		default: '',
		description: 'The ID of the offer-shift request',
	},
	{
		displayName: 'Sender Shift ID',
		name: 'offerSenderShiftId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createOfferShiftRequest'],
			},
		},
		default: '',
		description: 'The shift ID being offered',
	},
	{
		displayName: 'Recipient User ID',
		name: 'offerRecipientUserId',
		type: 'string',
		required: true,
		displayOptions: {
			show: {
				resource: ['teamsShift'],
				operation: ['createOfferShiftRequest'],
			},
		},
		default: '',
		description: 'The user ID of the person receiving the offer',
	},
];
