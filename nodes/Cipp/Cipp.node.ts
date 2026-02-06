import type {
	ICredentialsDecrypted,
	ICredentialTestFunctions,
	IDataObject,
	IExecuteFunctions,
	ILoadOptionsFunctions,
	INodeCredentialTestResult,
	INodeExecutionData,
	INodeListSearchResult,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';

import {
	cippApiRequest,
	getResourceLocatorValue,
	getTenantList,
} from './GenericFunctions';

import { operationFields, resourceFields } from './descriptions';

export class Cipp implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'CIPP.app',
		name: 'cippApp',
		icon: 'file:cipp.png',
		group: ['transform'],
		version: 1,
		subtitle: '={{$parameter["operation"] + ": " + $parameter["resource"]}}',
		description: 'Manage Microsoft 365 tenants via CIPP.app',
		defaults: {
			name: 'CIPP.app',
		},
		inputs: ['main'],
		outputs: ['main'],
		credentials: [
			{
				name: 'cippApi',
				required: true,
				testedBy: 'cippApiCredentialTest',
			},
		],
		properties: [
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'Alert',
						value: 'alert',
						description: 'Manage alerts and security incidents',
					},
					{
						name: 'Application',
						value: 'application',
						description: 'Manage Intune applications',
					},
					{
						name: 'Autopilot',
						value: 'autopilot',
						description: 'Manage Autopilot devices',
					},
					{
						name: 'Backup',
						value: 'backup',
						description: 'Manage CIPP backups',
					},
					{
						name: 'Device',
						value: 'device',
						description: 'Manage Intune devices',
					},
					{
						name: 'GDAP',
						value: 'gdap',
						description: 'Manage GDAP partner relationships',
					},
					{
						name: 'Group',
						value: 'group',
						description: 'Manage Azure AD groups',
					},
					{
						name: 'Identity',
						value: 'identity',
						description: 'Manage audit logs, roles, and deleted items',
					},
					{
						name: 'Mailbox',
						value: 'mailbox',
						description: 'Manage Exchange mailboxes',
					},
					{
						name: 'OneDrive',
						value: 'onedrive',
						description: 'Provision and manage OneDrive',
					},
					{
						name: 'Policy',
						value: 'policy',
						description: 'Manage Intune policies and Defender TVM',
					},
					{
						name: 'Quarantine',
						value: 'quarantine',
						description: 'Manage quarantined email messages',
					},
					{
						name: 'Scheduled Item',
						value: 'scheduledItem',
						description: 'Manage scheduled jobs',
					},
					{
						name: 'Team',
						value: 'team',
						description: 'Manage Teams and SharePoint',
					},
					{
						name: 'Tenant',
						value: 'tenant',
						description: 'List and manage tenants',
					},
					{
						name: 'Tool',
						value: 'tools',
						description: 'Breach search and Graph requests',
					},
					{
						name: 'User',
						value: 'user',
						description: 'Manage Azure AD users',
					},
					{
						name: 'Voice',
						value: 'voice',
						description: 'Manage Teams Voice',
					},
				],
				default: 'tenant',
			},
			...operationFields,
			...resourceFields,
		],
		usableAsTool: true,
	};

	methods = {
		credentialTest: {
			async cippApiCredentialTest(
				this: ICredentialTestFunctions,
				credential: ICredentialsDecrypted,
			): Promise<INodeCredentialTestResult> {
				try {
					const creds = credential.data as IDataObject;
					const baseUrl = (creds.baseUrl as string).replace(/\/$/, '');
					const tenantId = creds.tenantId as string;
					const clientId = creds.clientId as string;
					const clientSecret = creds.clientSecret as string;

					// Step 1: Get OAuth token from Azure AD
					const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
					const scope = `api://${clientId}/.default`;

					const tokenBody = `grant_type=client_credentials&client_id=${encodeURIComponent(clientId)}&client_secret=${encodeURIComponent(clientSecret)}&scope=${encodeURIComponent(scope)}`;

					const tokenResponse = (await this.helpers.request({
						method: 'POST',
						uri: tokenUrl,
						headers: {
							'Content-Type': 'application/x-www-form-urlencoded',
						},
						body: tokenBody,
						json: true,
					})) as IDataObject;

					if (!tokenResponse.access_token) {
						return {
							status: 'Error',
							message: 'Failed to obtain access token from Azure AD. Check your Tenant ID, Client ID, and Client Secret.',
						};
					}

					// Step 2: Test API connection with the token
					await this.helpers.request({
						method: 'GET',
						uri: `${baseUrl}/api/ListTenants`,
						headers: {
							Authorization: `Bearer ${tokenResponse.access_token}`,
							Accept: 'application/json',
						},
						json: true,
					});

					return {
						status: 'OK',
						message: 'Connection successful!',
					};
				} catch (error) {
					const err = error as { message?: string; statusCode?: number };
					return {
						status: 'Error',
						message: `Connection failed: ${err.message || 'Unknown error'}`,
					};
				}
			},
		},
		listSearch: {
			async tenantSearch(
				this: ILoadOptionsFunctions,
				filter?: string,
			): Promise<INodeListSearchResult> {
				const tenants = await getTenantList.call(this);

				const results = tenants
					.filter((tenant) => {
						if (!filter) return true;
						const searchTerm = filter.toLowerCase();
						return (
							tenant.displayName?.toLowerCase().includes(searchTerm) ||
							tenant.defaultDomainName?.toLowerCase().includes(searchTerm)
						);
					})
					.map((tenant) => ({
						name: tenant.displayName || tenant.defaultDomainName,
						value: tenant.defaultDomainName,
						url: `https://portal.azure.com/${tenant.defaultDomainName}`,
					}));

				return { results };
			},
		},
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];

		const resource = this.getNodeParameter('resource', 0) as string;
		const operation = this.getNodeParameter('operation', 0) as string;

		for (let i = 0; i < items.length; i++) {
			try {
				let responseData: IDataObject | IDataObject[] = {};

				// Get tenant filter if applicable
				const getTenantFilter = (): string => {
					try {
						const tenantValue = this.getNodeParameter('tenantFilter', i) as IDataObject;
						return getResourceLocatorValue(tenantValue);
					} catch {
						return '';
					}
				};

				// ==================== TENANT ====================
				if (resource === 'tenant') {
					if (operation === 'getAll') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						const options = this.getNodeParameter('options', i, {}) as IDataObject;

						const qs: IDataObject = {};
						if (options.allTenantSelector) {
							qs.allTenantSelector = true;
						}

						responseData = await cippApiRequest.call(this, 'GET', '/api/ListTenants', {}, qs);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'clearCache') {
						const clearTenantOnly = this.getNodeParameter('clearCacheTenantOnly', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ListTenants',
							{
								ClearCache: true,
								TenantsOnly: clearTenantOnly,
							},
							{},
						);
					} else if (operation === 'getLicenses') {
						const tenantFilter = getTenantFilter();
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						const licenseOptions = this.getNodeParameter('licenseOptions', i, {}) as IDataObject;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListLicenses',
							{},
							{ tenantFilter },
						);

						// Strip out AssignedUsers and AssignedGroups if summaryOnly is enabled
						if (licenseOptions.summaryOnly && Array.isArray(responseData)) {
							responseData = (responseData as IDataObject[]).map((license) => {
								const { AssignedUsers, AssignedGroups, ...summary } = license;
								return summary;
							});
						}

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'getCspLicenses') {
						const tenantFilter = getTenantFilter();
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListCSPLicenses',
							{},
							{ tenantFilter },
						);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'cspLicenseAction') {
						const tenantFilter = getTenantFilter();
						const action = this.getNodeParameter('cspAction', i) as string;
						const licenseSku = this.getNodeParameter('licenseSku', i) as string;
						const quantity = this.getNodeParameter('licenseQuantity', i) as number;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecCSPLicense',
							{
								tenantFilter,
								LicenseSKU: licenseSku,
								Quantity: quantity,
								Action: action,
							},
							{},
						);
					} else if (operation === 'listDefenderState') {
						const tenantFilter = getTenantFilter();
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListDefenderState',
							{},
							{ tenantFilter },
						);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'listCspSkus') {
						const tenantFilter = getTenantFilter();
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListCSPSKUs',
							{},
							{ tenantFilter },
						);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					}
				}

				// ==================== USER ====================
				else if (resource === 'user') {
					const tenantFilter = getTenantFilter();

					if (operation === 'getAll') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						const userFields = this.getNodeParameter('userFields', i, []) as string[];
						const filters = this.getNodeParameter('filters', i, {}) as IDataObject;

						const qs: IDataObject = {
							tenantFilter,
							Endpoint: 'users',
						};

						// Build select fields from multi-select + additional fields
						const selectFields: string[] = [...userFields];
						if (filters.select) {
							const additionalFields = (filters.select as string).split(',').map((f) => f.trim());
							selectFields.push(...additionalFields);
						}
						if (selectFields.length > 0) {
							qs['$select'] = selectFields.join(',');
						}

						if (filters.filter) qs['$filter'] = filters.filter;

						if (!returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							qs['$top'] = limit;
						}

						responseData = await cippApiRequest.call(this, 'GET', '/api/ListGraphRequest', {}, qs);
					} else if (operation === 'add') {
						const firstName = this.getNodeParameter('firstName', i) as string;
						const lastName = this.getNodeParameter('lastName', i) as string;
						const domain = this.getNodeParameter('domain', i) as string;
						const additionalFields = this.getNodeParameter('additionalFields', i, {}) as IDataObject;

						const body: IDataObject = {
							tenantFilter,
							firstName,
							lastName,
							domain,
							...additionalFields,
						};

						responseData = await cippApiRequest.call(this, 'POST', '/api/AddUser', body, {});
					} else if (operation === 'disable' || operation === 'enable') {
						const userId = this.getNodeParameter('userId', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecDisableUser',
							{
								tenantFilter,
								ID: userId,
								Enable: operation === 'enable',
							},
							{},
						);
					} else if (operation === 'resetPassword') {
						const userId = this.getNodeParameter('userId', i) as string;
						const options = this.getNodeParameter('passwordOptions', i, {}) as IDataObject;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecResetPass',
							{
								tenantFilter,
								ID: userId,
								MustChange: options.mustChangePass !== false,
							},
							{},
						);
					} else if (operation === 'remove') {
						const userId = this.getNodeParameter('userId', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/RemoveUser',
							{
								tenantFilter,
								ID: userId,
							},
							{},
						);
					} else if (operation === 'resetMfa') {
						const userId = this.getNodeParameter('userId', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecResetMFA',
							{
								tenantFilter,
								ID: userId,
							},
							{},
						);
					} else if (operation === 'sendMfaPush') {
						const userId = this.getNodeParameter('userId', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecSendPush',
							{
								tenantFilter,
								UserPrincipalName: userId,
							},
							{},
						);
					} else if (operation === 'setPerUserMfa') {
						const userId = this.getNodeParameter('userId', i) as string;
						const mfaState = this.getNodeParameter('mfaState', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecPerUserMFA',
							{
								tenantFilter,
								userId,
								State: mfaState,
							},
							{},
						);
					} else if (operation === 'createTap') {
						const userId = this.getNodeParameter('userId', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecCreateTAP',
							{
								tenantFilter,
								ID: userId,
							},
							{},
						);
					} else if (operation === 'clearImmutableId') {
						const userId = this.getNodeParameter('userId', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecClrImmId',
							{
								tenantFilter,
								ID: userId,
							},
							{},
						);
					} else if (operation === 'revokeSessions') {
						const userId = this.getNodeParameter('userId', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecRevokeSessions',
							{
								tenantFilter,
								ID: userId,
							},
							{},
						);
					} else if (operation === 'offboard') {
						const users = this.getNodeParameter('usersToOffboard', i) as string;
						const scheduled = this.getNodeParameter('scheduledOffboard', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecOffboardUser',
							{
								tenantFilter,
								user: JSON.parse(users),
								Scheduled: { enabled: scheduled },
							},
							{},
						);
					} else if (operation === 'listInactiveAccounts') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListInactiveAccounts',
							{},
							{ tenantFilter },
						);
						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'listSignIns') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListSignIns',
							{},
							{ tenantFilter },
						);
						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'listMfaUsers') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListMFAUsers',
							{},
							{ tenantFilter },
						);
						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'dismissRiskyUser') {
						const userId = this.getNodeParameter('userId', i) as string;
						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecDismissRiskyUser',
							{
								tenantFilter,
								ID: userId,
							},
							{},
						);
					} else if (operation === 'listJitAdmin') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListJITAdmin',
							{},
							{ tenantFilter },
						);
						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'execJitAdmin') {
						const userId = this.getNodeParameter('userId', i) as string;
						const jitAdminRole = this.getNodeParameter('jitAdminRole', i) as string;
						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecJITAdmin',
							{
								tenantFilter,
								ID: userId,
								Role: jitAdminRole,
							},
							{},
						);
					}
				}

				// ==================== GROUP ====================
				else if (resource === 'group') {
					const tenantFilter = getTenantFilter();

					if (operation === 'getAll') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						const options = this.getNodeParameter('options', i, {}) as IDataObject;

						const qs: IDataObject = { tenantFilter };
						if (options.groupId) qs.groupId = options.groupId;
						if (options.members) qs.members = true;
						if (options.owners) qs.owners = true;

						responseData = await cippApiRequest.call(this, 'GET', '/api/ListGroups', {}, qs);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'add') {
						const groupName = this.getNodeParameter('groupName', i) as string;
						const groupType = this.getNodeParameter('groupType', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/AddGroup',
							{
								tenantFilter,
								displayName: groupName,
								groupType,
							},
							{},
						);
					} else if (operation === 'edit') {
						const groupId = this.getNodeParameter('groupId', i) as string;
						const editOptions = this.getNodeParameter('editOptions', i, {}) as IDataObject;

						const body: IDataObject = {
							tenantFilter,
							groupId,
						};

						if (editOptions.addMembers) {
							body.addMembers = (editOptions.addMembers as string).split(',').map((m) => m.trim());
						}
						if (editOptions.addOwners) {
							body.addOwners = (editOptions.addOwners as string).split(',').map((o) => o.trim());
						}
						if (editOptions.removeMembers) {
							body.removeMembers = (editOptions.removeMembers as string)
								.split(',')
								.map((m) => m.trim());
						}
						if (editOptions.removeOwners) {
							body.removeOwners = (editOptions.removeOwners as string)
								.split(',')
								.map((o) => o.trim());
						}

						responseData = await cippApiRequest.call(this, 'POST', '/api/EditGroup', body, {});
					} else if (operation === 'delete') {
						const groupId = this.getNodeParameter('groupId', i) as string;
						const groupType = this.getNodeParameter('groupTypeForDelete', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecGroupsDelete',
							{
								tenantFilter,
								ID: groupId,
								groupType,
							},
							{},
						);
					} else if (operation === 'hideFromGal') {
						const groupId = this.getNodeParameter('groupId', i) as string;
						const groupEmail = this.getNodeParameter('groupEmail', i) as string;
						const groupType = this.getNodeParameter('groupTypeForDelete', i) as string;
						const hideFromGal = this.getNodeParameter('hideFromGal', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecGroupsHideFromGAL',
							{
								tenantFilter,
								ID: groupId,
								groupEmail,
								groupType,
								HideFromGAL: hideFromGal,
							},
							{},
						);
					} else if (operation === 'deliveryManagement') {
						const groupId = this.getNodeParameter('groupId', i) as string;
						const groupEmail = this.getNodeParameter('groupEmail', i) as string;
						const groupType = this.getNodeParameter('groupTypeForDelete', i) as string;
						const onlyInternal = this.getNodeParameter('onlyInternal', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecGroupsDeliveryManagement',
							{
								tenantFilter,
								ID: groupId,
								groupEmail,
								groupType,
								OnlyAllowInternal: onlyInternal,
							},
							{},
						);
					}
				}

				// ==================== DEVICE ====================
				else if (resource === 'device') {
					const tenantFilter = getTenantFilter();

					if (operation === 'getAll') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListDevices',
							{},
							{ tenantFilter },
						);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'manage') {
						const deviceId = this.getNodeParameter('deviceId', i) as string;
						const action = this.getNodeParameter('manageAction', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecDeviceDelete',
							{
								tenantFilter,
								ID: deviceId,
								action,
							},
							{},
						);
					} else if (operation === 'executeAction') {
						const deviceId = this.getNodeParameter('deviceId', i) as string;
						const action = this.getNodeParameter('executeDeviceAction', i) as string;

						const body: IDataObject = {
							tenantFilter,
							GUID: deviceId,
							Action: action,
						};

						if (action === 'Rename') {
							body.newDeviceName = this.getNodeParameter('newDeviceName', i) as string;
						}

						responseData = await cippApiRequest.call(this, 'POST', '/api/ExecDeviceAction', body, {});
					} else if (operation === 'getRecoveryKey') {
						const deviceId = this.getNodeParameter('deviceId', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecGetRecoveryKey',
							{
								tenantFilter,
								GUID: deviceId,
							},
							{},
						);
					} else if (operation === 'getLapsPassword') {
						const deviceId = this.getNodeParameter('deviceId', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecGetLocalAdminPassword',
							{
								tenantFilter,
								GUID: deviceId,
							},
							{},
						);
					}
				}

				// ==================== AUTOPILOT ====================
				else if (resource === 'autopilot') {
					const tenantFilter = getTenantFilter();

					if (operation === 'getAll') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListAPDevices',
							{},
							{ tenantFilter },
						);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'assign') {
						const deviceId = this.getNodeParameter('deviceId', i) as string;
						const serialNumber = this.getNodeParameter('serialNumber', i) as string;
						const userPrincipalName = this.getNodeParameter('userPrincipalName', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecAssignAPDevice',
							{
								tenantFilter,
								ID: deviceId,
								serialNumber,
								userPrincipalName,
							},
							{},
						);
					} else if (operation === 'remove') {
						const deviceId = this.getNodeParameter('deviceId', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/RemoveAPDevice',
							{
								tenantFilter,
								ID: deviceId,
							},
							{},
						);
					} else if (operation === 'sync') {
						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecSyncAPDevices',
							{ tenantFilter },
							{},
						);
					} else if (operation === 'getConfigurations') {
						const configType = this.getNodeParameter('configType', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListAutopilotConfig',
							{},
							{ tenantFilter, type: configType },
						);
					}
				}

				// ==================== MAILBOX ====================
				else if (resource === 'mailbox') {
					const tenantFilter = getTenantFilter();
					const userId = this.getNodeParameter('userId', i) as string;

					if (operation === 'convert') {
						const mailboxType = this.getNodeParameter('mailboxType', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecConvertMailbox',
							{
								tenantFilter,
								ID: userId,
								MailboxType: mailboxType,
							},
							{},
						);
					} else if (operation === 'enableArchive') {
						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecEnableArchive',
							{
								tenantFilter,
								ID: userId,
							},
							{},
						);
					} else if (operation === 'setOutOfOffice') {
						const autoReplyState = this.getNodeParameter('autoReplyState', i) as string;
						const body: IDataObject = {
							tenantFilter,
							user: userId,
							AutoReplyState: autoReplyState,
						};

						if (autoReplyState === 'Enabled') {
							body.input = this.getNodeParameter('autoReplyMessage', i) as string;
						}

						responseData = await cippApiRequest.call(this, 'POST', '/api/ExecSetOOO', body, {});
					} else if (operation === 'setForwarding') {
						const forwardTo = this.getNodeParameter('forwardTo', i) as string;
						const keepCopy = this.getNodeParameter('keepCopy', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecEmailForward',
							{
								tenantFilter,
								user: userId,
								ForwardTo: forwardTo,
								KeepCopy: keepCopy,
							},
							{},
						);
					}
				}

				// ==================== QUARANTINE ====================
				else if (resource === 'quarantine') {
					const tenantFilter = getTenantFilter();

					if (operation === 'getMany') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListMailQuarantine',
							{},
							{ tenantFilter },
						);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'release') {
						const messageId = this.getNodeParameter('messageId', i) as string;
						const allowSender = this.getNodeParameter('allowSender', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecQuarantineManagement',
							{
								tenantFilter,
								ID: messageId,
								Type: 'Release',
								AllowSender: allowSender,
							},
							{},
						);
					} else if (operation === 'deny') {
						const messageId = this.getNodeParameter('messageId', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecQuarantineManagement',
							{
								tenantFilter,
								ID: messageId,
								Type: 'Deny',
							},
							{},
						);
					}
				}

				// ==================== ALERT ====================
				else if (resource === 'alert') {
					if (operation === 'getAll') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						responseData = await cippApiRequest.call(this, 'GET', '/api/ListAlertsQueue', {}, {});

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'add') {
						const alertConfig = this.getNodeParameter('alertConfig', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/AddAlert',
							JSON.parse(alertConfig),
							{},
						);
					} else if (operation === 'getSecurityAlerts') {
						const tenantFilter = getTenantFilter();
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ExecAlertsList',
							{},
							{ tenantFilter },
						);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'getSecurityIncidents') {
						const tenantFilter = getTenantFilter();
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ExecIncidentsList',
							{},
							{ tenantFilter },
						);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'setSecurityAlertStatus') {
						const tenantFilter = getTenantFilter();
						const alertId = this.getNodeParameter('alertId', i) as string;
						const status = this.getNodeParameter('alertStatus', i) as string;
						const additionalFields = this.getNodeParameter(
							'alertAdditionalFields',
							i,
							{},
						) as IDataObject;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecSetSecurityAlert',
							{
								tenantFilter,
								ID: alertId,
								Status: status,
								...additionalFields,
							},
							{},
						);
					} else if (operation === 'setSecurityIncidentStatus') {
						const tenantFilter = getTenantFilter();
						const incidentId = this.getNodeParameter('incidentId', i) as string;
						const status = this.getNodeParameter('incidentStatus', i) as string;
						const assignedTo = this.getNodeParameter('assignedTo', i, '') as string;

						const body: IDataObject = {
							tenantFilter,
							ID: incidentId,
							Status: status,
						};

						if (assignedTo) {
							body.AssignedTo = assignedTo;
						}

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecSetSecurityIncident',
							body,
							{},
						);
					}
				}

				// ==================== APPLICATION ====================
				else if (resource === 'application') {
					const tenantFilter = getTenantFilter();

					if (operation === 'getAll') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListApps',
							{},
							{ tenantFilter },
						);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'getQueue') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListApplicationQueue',
							{},
							{ tenantFilter },
						);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'assign') {
						const appId = this.getNodeParameter('appId', i) as string;
						const assignTo = this.getNodeParameter('assignTo', i) as string;

						const body: IDataObject = {
							tenantFilter,
							ID: appId,
							AssignTo: assignTo,
						};

						if (assignTo === 'customGroup') {
							const customGroups = this.getNodeParameter('customGroupNames', i, '') as string;
							body.customGroupNames = customGroups.split(',').map((g) => g.trim());
						}

						responseData = await cippApiRequest.call(this, 'POST', '/api/ExecAssignApp', body, {});
					} else if (operation === 'remove') {
						const appId = this.getNodeParameter('appId', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/RemoveApp',
							{
								tenantFilter,
								ID: appId,
							},
							{},
						);
					} else if (operation === 'removeFromQueue') {
						const queueId = this.getNodeParameter('queueId', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/RemoveQueuedApp',
							{
								tenantFilter,
								ID: queueId,
							},
							{},
						);
					} else if (operation === 'addWinget') {
						const packageId = this.getNodeParameter('packageId', i) as string;
						const appName = this.getNodeParameter('appName', i) as string;
						const appDescription = this.getNodeParameter('appDescription', i, '') as string;
						const uninstall = this.getNodeParameter('uninstall', i) as boolean;
						const assignTo = this.getNodeParameter('assignTo', i) as string;

						const body: IDataObject = {
							tenantFilter,
							PackageIdentifier: packageId,
							ApplicationName: appName,
							Description: appDescription,
							InstallAsSystem: true,
							UninstallApp: uninstall,
							AssignTo: assignTo,
						};

						if (assignTo === 'customGroup') {
							const customGroups = this.getNodeParameter('customGroupNames', i, '') as string;
							body.customGroupNames = customGroups.split(',').map((g) => g.trim());
						}

						responseData = await cippApiRequest.call(this, 'POST', '/api/AddWinGetApp', body, {});
					} else if (operation === 'addStore') {
						const packageId = this.getNodeParameter('packageId', i) as string;
						const appName = this.getNodeParameter('appName', i) as string;
						const appDescription = this.getNodeParameter('appDescription', i, '') as string;
						const uninstall = this.getNodeParameter('uninstall', i) as boolean;
						const assignTo = this.getNodeParameter('assignTo', i) as string;

						const body: IDataObject = {
							tenantFilter,
							PackageIdentifier: packageId,
							ApplicationName: appName,
							Description: appDescription,
							UninstallApp: uninstall,
							AssignTo: assignTo,
						};

						if (assignTo === 'customGroup') {
							const customGroups = this.getNodeParameter('customGroupNames', i, '') as string;
							body.customGroupNames = customGroups.split(',').map((g) => g.trim());
						}

						responseData = await cippApiRequest.call(this, 'POST', '/api/AddStoreApp', body, {});
					} else if (operation === 'addChocolatey') {
						const packageName = this.getNodeParameter('packageName', i) as string;
						const appName = this.getNodeParameter('appName', i) as string;
						const appDescription = this.getNodeParameter('appDescription', i, '') as string;
						const uninstall = this.getNodeParameter('uninstall', i) as boolean;
						const assignTo = this.getNodeParameter('assignTo', i) as string;
						const chocoOptions = this.getNodeParameter('chocoOptions', i, {}) as IDataObject;

						const body: IDataObject = {
							tenantFilter,
							PackageName: packageName,
							ApplicationName: appName,
							Description: appDescription,
							UninstallApp: uninstall,
							AssignTo: assignTo,
							...chocoOptions,
						};

						if (assignTo === 'customGroup') {
							const customGroups = this.getNodeParameter('customGroupNames', i, '') as string;
							body.customGroupNames = customGroups.split(',').map((g) => g.trim());
						}

						responseData = await cippApiRequest.call(this, 'POST', '/api/AddChocoApp', body, {});
					} else if (operation === 'addMsp') {
						const rmmTool = this.getNodeParameter('rmmTool', i) as string;
						const displayName = this.getNodeParameter('mspDisplayName', i) as string;
						const rmmParameters = this.getNodeParameter('rmmParameters', i) as string;
						const assignTo = this.getNodeParameter('assignTo', i) as string;

						const body: IDataObject = {
							tenantFilter,
							RMM: rmmTool,
							displayName,
							RMMParams: JSON.parse(rmmParameters),
							AssignTo: assignTo,
						};

						if (assignTo === 'customGroup') {
							const customGroups = this.getNodeParameter('customGroupNames', i, '') as string;
							body.customGroupNames = customGroups.split(',').map((g) => g.trim());
						}

						responseData = await cippApiRequest.call(this, 'POST', '/api/AddMSPApp', body, {});
					} else if (operation === 'addOffice') {
						const excludedApps = this.getNodeParameter('excludedApps', i) as string[];
						const updateChannel = this.getNodeParameter('updateChannel', i) as string;
						const assignTo = this.getNodeParameter('assignTo', i) as string;
						const officeOptions = this.getNodeParameter('officeOptions', i, {}) as IDataObject;

						const body: IDataObject = {
							tenantFilter,
							ExcludeApps: excludedApps,
							UpdateChannel: updateChannel,
							AssignTo: assignTo,
							...officeOptions,
						};

						responseData = await cippApiRequest.call(this, 'POST', '/api/AddOfficeApp', body, {});
					}
				}

				// ==================== TEAM ====================
				else if (resource === 'team') {
					const tenantFilter = getTenantFilter();

					if (operation === 'getAll') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListTeams',
							{},
							{ tenantFilter, type: 'list' },
						);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'add') {
						const displayName = this.getNodeParameter('displayName', i) as string;
						const description = this.getNodeParameter('teamDescription', i, '') as string;
						const owner = this.getNodeParameter('owner', i) as string;
						const visibility = this.getNodeParameter('visibility', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/AddTeam',
							{
								tenantFilter,
								displayName,
								description,
								owner,
								visibility,
							},
							{},
						);
					} else if (operation === 'getSites') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						const siteType = this.getNodeParameter('siteType', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListSites',
							{},
							{ tenantFilter, type: siteType },
						);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'addSite') {
						const siteName = this.getNodeParameter('siteName', i) as string;
						const siteDescription = this.getNodeParameter('siteDescription', i, '') as string;
						const owner = this.getNodeParameter('siteOwner', i) as string;
						const templateName = this.getNodeParameter('templateName', i) as string;

						const body: IDataObject = {
							tenantFilter,
							siteName,
							siteDescription,
							owner,
							TemplateName: templateName,
						};

						if (templateName === 'communication') {
							body.siteDesign = this.getNodeParameter('siteDesign', i) as string;
						}

						responseData = await cippApiRequest.call(this, 'POST', '/api/AddSite', body, {});
					} else if (operation === 'getActivity') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListTeamsActivity',
							{},
							{ tenantFilter, type: 'TeamsUserActivityUser' },
						);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'manageSiteMember') {
						const siteUrl = this.getNodeParameter('siteUrl', i) as string;
						const siteUser = this.getNodeParameter('siteUser', i) as string;
						const action = this.getNodeParameter('memberAction', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecSetSharePointMember',
							{
								tenantFilter,
								URL: siteUrl,
								AddMember: action === 'add',
								userPrincipalName: siteUser,
							},
							{},
						);
					} else if (operation === 'manageSitePermissions') {
						const siteUrl = this.getNodeParameter('siteUrl', i) as string;
						const siteUser = this.getNodeParameter('siteUser', i) as string;
						const removePermission = this.getNodeParameter('removePermission', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecSharePointPerms',
							{
								tenantFilter,
								URL: siteUrl,
								RemovePermission: removePermission,
								userPrincipalName: siteUser,
							},
							{},
						);
					} else if (operation === 'addSitesBulk') {
						const sitesConfig = this.getNodeParameter('sitesConfig', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/AddSiteBulk',
							{
								tenantFilter,
								sites: JSON.parse(sitesConfig),
							},
							{},
						);
					}
				}

				// ==================== VOICE ====================
				else if (resource === 'voice') {
					const tenantFilter = getTenantFilter();

					if (operation === 'getPhoneNumbers') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListTeamsVoice',
							{},
							{ tenantFilter },
						);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'getLocations') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListTeamsLisLocation',
							{},
							{ tenantFilter },
						);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'assignNumber') {
						const phoneNumber = this.getNodeParameter('phoneNumber', i) as string;
						const voiceUser = this.getNodeParameter('voiceUser', i) as string;
						const phoneNumberType = this.getNodeParameter('phoneNumberType', i, '') as string;
						const locationOnly = this.getNodeParameter('locationOnly', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecTeamsVoicePhoneNumberAssignment',
							{
								tenantFilter,
								PhoneNumber: phoneNumber,
								PhoneNumberType: phoneNumberType,
								LocationOnly: locationOnly,
								UserPrincipalNameOrLocationId: voiceUser,
							},
							{},
						);
					} else if (operation === 'unassignNumber') {
						const phoneNumber = this.getNodeParameter('phoneNumber', i) as string;
						const voiceUser = this.getNodeParameter('voiceUser', i) as string;
						const phoneNumberType = this.getNodeParameter('phoneNumberType', i, '') as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecRemoveTeamsVoicePhoneNumberAssignment',
							{
								tenantFilter,
								PhoneNumber: phoneNumber,
								PhoneNumberType: phoneNumberType,
								AssignedTo: voiceUser,
							},
							{},
						);
					}
				}

				// ==================== SCHEDULED ITEM ====================
				else if (resource === 'scheduledItem') {
					if (operation === 'getAll') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						const options = this.getNodeParameter('options', i, {}) as IDataObject;

						const qs: IDataObject = {};
						if (options.showHidden) qs.showHidden = true;
						if (options.name) qs.name = options.name;

						responseData = await cippApiRequest.call(this, 'GET', '/api/ListScheduledItems', {}, qs);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'add') {
						const tenantFilter = getTenantFilter();
						const jobName = this.getNodeParameter('jobName', i) as string;
						const command = this.getNodeParameter('command', i) as string;
						const scheduledTime = this.getNodeParameter('scheduledTime', i, '') as string;
						const recurrence = this.getNodeParameter('recurrence', i) as string;
						const parameters = this.getNodeParameter('parameters', i) as string;
						const postExecution = this.getNodeParameter('postExecution', i) as string[];

						const body: IDataObject = {
							TenantFilter: tenantFilter || 'AllTenants',
							Name: jobName,
							Command: command,
							Recurrence: recurrence,
							Parameters: JSON.parse(parameters),
							PostExecution: postExecution,
						};

						if (scheduledTime) {
							body.ScheduledTime = new Date(scheduledTime).getTime() / 1000;
						}

						responseData = await cippApiRequest.call(this, 'POST', '/api/AddScheduledItem', body, {});
					} else if (operation === 'remove') {
						const rowKey = this.getNodeParameter('rowKey', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/RemoveScheduledItem',
							{
								RowKey: rowKey,
							},
							{},
						);
					}
				}

				// ==================== BACKUP ====================
				else if (resource === 'backup') {
					if (operation === 'getAll') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						const options = this.getNodeParameter('options', i, {}) as IDataObject;

						const qs: IDataObject = {};
						if (options.namesOnly) qs.NameOnly = true;
						if (options.backupName) qs.BackupName = options.backupName;

						responseData = await cippApiRequest.call(this, 'GET', '/api/ExecListBackup', {}, qs);

						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'run') {
						responseData = await cippApiRequest.call(this, 'POST', '/api/ExecRunBackup', {}, {});
					} else if (operation === 'restore') {
						const backupData = this.getNodeParameter('backupData', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecRestoreBackup',
							JSON.parse(backupData),
							{},
						);
					} else if (operation === 'setAutoBackup') {
						const enableAutoBackup = this.getNodeParameter('enableAutoBackup', i) as boolean;

						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecSetCIPPAutoBackup',
							{
								Enabled: enableAutoBackup,
							},
							{},
						);
					}
				}

				// ==================== TOOLS ====================
				else if (resource === 'tools') {
					if (operation === 'breachAccount') {
						const account = this.getNodeParameter('account', i) as string;

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListBreachesAccount',
							{},
							{ account },
						);
					} else if (operation === 'breachTenant') {
						const tenantFilter = getTenantFilter();

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListBreachesTenant',
							{},
							{ tenantFilter },
						);
					} else if (operation === 'executeBreachSearch') {
						const tenantFilter = getTenantFilter();

						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ExecBreachSearch',
							{},
							{ tenantFilter },
						);
					} else if (operation === 'graphRequest') {
						const tenantFilter = getTenantFilter();
						const endpoint = this.getNodeParameter('graphEndpoint', i) as string;
						const graphOptions = this.getNodeParameter('graphOptions', i, {}) as IDataObject;

						const qs: IDataObject = {
							tenantFilter,
							Endpoint: endpoint,
						};

						if (graphOptions.select) qs['$select'] = graphOptions.select;
						if (graphOptions.filter) qs['$filter'] = graphOptions.filter;
						if (graphOptions.orderby) qs['$orderby'] = graphOptions.orderby;
						if (graphOptions.top) qs['$top'] = graphOptions.top;
						if (graphOptions.count) qs['$count'] = graphOptions.count;

						responseData = await cippApiRequest.call(this, 'GET', '/api/ListGraphRequest', {}, qs);
					}
				}

				// ==================== IDENTITY ====================
				else if (resource === 'identity') {
					const tenantFilter = getTenantFilter();

					if (operation === 'listAuditLogs') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListAuditLogs',
							{},
							{ tenantFilter },
						);
						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'listDeletedItems') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListDeletedItems',
							{},
							{ tenantFilter },
						);
						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'listRoles') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListRoles',
							{},
							{ tenantFilter },
						);
						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'restoreDeleted') {
						const objectId = this.getNodeParameter('objectId', i) as string;
						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecRestoreDeleted',
							{
								tenantFilter,
								ID: objectId,
							},
							{},
						);
					}
				}

				// ==================== POLICY ====================
				else if (resource === 'policy') {
					const tenantFilter = getTenantFilter();

					if (operation === 'getMany') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListIntunePolicy',
							{},
							{ tenantFilter },
						);
						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'add') {
						const policyConfig = this.getNodeParameter('policyConfig', i) as string;
						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/AddPolicy',
							{
								tenantFilter,
								...JSON.parse(policyConfig),
							},
							{},
						);
					} else if (operation === 'assign') {
						const policyId = this.getNodeParameter('policyId', i) as string;
						const assignTo = this.getNodeParameter('assignTo', i) as string;
						const body: IDataObject = {
							tenantFilter,
							ID: policyId,
							AssignTo: assignTo,
						};
						if (assignTo === 'customGroup') {
							body.customGroupNames = this.getNodeParameter('customGroupNames', i) as string;
						}
						responseData = await cippApiRequest.call(this, 'POST', '/api/AssignPolicy', body, {});
					} else if (operation === 'remove') {
						const policyId = this.getNodeParameter('policyId', i) as string;
						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/RemovePolicy',
							{
								tenantFilter,
								ID: policyId,
							},
							{},
						);
					} else if (operation === 'listDefenderTvm') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListDefenderTVM',
							{},
							{ tenantFilter },
						);
						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					}
				}

				// ==================== ONEDRIVE ====================
				else if (resource === 'onedrive') {
					const tenantFilter = getTenantFilter();
					const userId = this.getNodeParameter('userId', i) as string;

					if (operation === 'provision') {
						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecOneDriveProvision',
							{
								tenantFilter,
								UserPrincipalName: userId,
							},
							{},
						);
					} else if (operation === 'addShortcut') {
						const shortcutUrl = this.getNodeParameter('shortcutUrl', i) as string;
						const shortcutName = this.getNodeParameter('shortcutName', i, '') as string;
						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecOneDriveShortCut',
							{
								tenantFilter,
								UserPrincipalName: userId,
								URL: shortcutUrl,
								ShortcutName: shortcutName,
							},
							{},
						);
					}
				}

				// ==================== GDAP ====================
				else if (resource === 'gdap') {
					const tenantFilter = getTenantFilter();

					if (operation === 'listRoles') {
						const returnAll = this.getNodeParameter('returnAll', i) as boolean;
						responseData = await cippApiRequest.call(
							this,
							'GET',
							'/api/ListGDAPRoles',
							{},
							{ tenantFilter },
						);
						if (Array.isArray(responseData) && !returnAll) {
							const limit = this.getNodeParameter('limit', i) as number;
							responseData = responseData.slice(0, limit);
						}
					} else if (operation === 'sendInvite') {
						const gdapRoles = this.getNodeParameter('gdapRoles', i) as string;
						responseData = await cippApiRequest.call(
							this,
							'POST',
							'/api/ExecGDAPInvite',
							{
								tenantFilter,
								Roles: gdapRoles.split(',').map((r) => r.trim()),
							},
							{},
						);
					}
				}

				// Handle response
				const executionData = this.helpers.constructExecutionMetaData(
					this.helpers.returnJsonArray(responseData as IDataObject[]),
					{ itemData: { item: i } },
				);

				returnData.push(...executionData);
			} catch (error) {
				if (this.continueOnFail()) {
					const executionErrorData = this.helpers.constructExecutionMetaData(
						this.helpers.returnJsonArray({ error: (error as Error).message }),
						{ itemData: { item: i } },
					);
					returnData.push(...executionErrorData);
					continue;
				}
				throw error;
			}
		}

		return [returnData];
	}
}
