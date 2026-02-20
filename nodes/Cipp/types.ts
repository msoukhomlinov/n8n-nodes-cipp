export interface ICippCredentials {
	baseUrl: string;
	tenantId: string;
	clientId: string;
	clientSecret: string;
}

export interface IAuthToken {
	accessToken: string;
	expiresAt: number;
}

export interface ITenant {
	customerId: string;
	defaultDomainName: string;
	displayName: string;
	domains?: string[];
}

export interface ICippApiResponse<T = unknown> {
	Results?: T;
	Metadata?: {
		Heading?: string;
		Success?: boolean;
	};
}

export interface IGraphRequestParams {
	Endpoint: string;
	tenantFilter?: string;
	$select?: string;
	$filter?: string;
	$orderby?: string;
	$top?: number;
	$count?: boolean;
	manualPagination?: boolean;
	ReverseTenantLookup?: boolean;
}

export interface IGraphExecRequestParams {
	tenantFilter: string;
	endpoint: string;
	method: 'GET' | 'POST' | 'PATCH';
	headers?: Record<string, string>;
	body?: unknown;
}

export interface IUserCreateParams {
	tenantFilter: string;
	firstName: string;
	lastName: string;
	mailNickname?: string;
	domain: string;
	usageLocation?: string;
	displayName?: string;
	password?: string;
	mustChangePass?: boolean;
	licenses?: string[];
}

export interface IGroupParams {
	tenantFilter: string;
	groupName?: string;
	groupType?: 'Microsoft 365' | 'Security' | 'Distribution' | 'Mail-Enabled Security';
	groupId?: string;
	members?: string[];
	owners?: string[];
	addMember?: string;
	removeMember?: string;
}

export interface IDeviceActionParams {
	tenantFilter: string;
	deviceId: string;
	action:
		| 'Reboot'
		| 'Rename'
		| 'Autopilot'
		| 'Wipe'
		| 'RemoveFromAutopilot'
		| 'RemoveFromIntune'
		| 'RemoveFromDefender'
		| 'RemoveFromAzure'
		| 'RemoveFromEverywhere'
		| 'SyncDevice'
		| 'FreshStart'
		| 'QuickScan'
		| 'FullScan'
		| 'WindowsUpdateScan'
		| 'WindowsUpdateReboot'
		| 'WindowsUpdateDrivers'
		| 'WindowsUpdateFeatures'
		| 'WindowsUpdateOther'
		| 'WindowsUpdateAll';
	newDeviceName?: string;
}

export interface IScheduledItemParams {
	tenantFilter?: string;
	Name: string;
	Command: string;
	Parameters?: Record<string, unknown>;
	ScheduledTime?: number;
	Recurrence?: '0' | '1d' | '7d' | '30d' | '365d';
	PostExecution?: string[];
	hidden?: boolean;
}

export interface IAlertParams {
	tenantFilter?: string[];
	excludedTenants?: string[];
	logSource?: string;
	conditions?: Record<string, unknown>;
	actions?: string[];
}

export interface ISecurityAlertParams {
	tenantFilter: string;
	alertId: string;
	status: 'inProgress' | 'resolved';
	vendorName?: string;
	providerName?: string;
}

export interface ISecurityIncidentParams {
	tenantFilter: string;
	incidentId: string;
	status: 'active' | 'inProgress' | 'resolved';
	assignedTo?: string;
}

export interface IAppAssignParams {
	tenantFilter: string;
	appId: string;
	assignTo: 'AllUsers' | 'AllDevices' | 'Both' | 'customGroup';
	customGroupNames?: string[];
}

export interface IPolicyAssignParams {
	tenantFilter: string;
	policyId: string;
	assignTo: 'AllUsers' | 'AllDevices' | 'Both';
}

export interface ITeamParams {
	tenantFilter: string;
	displayName: string;
	description?: string;
	owner: string;
	visibility: 'public' | 'private';
}

export interface ISiteParams {
	tenantFilter: string;
	siteName: string;
	siteDescription?: string;
	owner: string;
	templateName: 'team' | 'communication';
	siteDesign?: 'blank' | 'Showcase' | 'Topic';
}

export interface IVoiceAssignParams {
	tenantFilter: string;
	phoneNumber: string;
	phoneNumberType?: string;
	locationOnly?: boolean;
	userPrincipalNameOrLocationId: string;
}

export interface INamedLocationParams {
	tenantFilter: string;
	displayName: string;
	locationType: 'country' | 'ip';
	countriesAndRegions?: string[];
	includeUnknownCountriesAndRegions?: boolean;
	ipRanges?: string[];
	isTrusted?: boolean;
}

export interface IConditionalAccessExclusionParams {
	tenantFilter: string;
	userId: string;
	policyId: string;
	startDate?: string;
	endDate?: string;
}
