import type { Icon, ICredentialType, INodeProperties } from 'n8n-workflow';

export class CippApi implements ICredentialType {
	name = 'cippApi';
	displayName = 'CIPP.app API';
	icon: Icon = 'file:cipp.png';
	documentationUrl = 'https://docs.cipp.app/api-documentation/setup-and-authentication';

	properties: INodeProperties[] = [
		{
			displayName: 'CIPP Instance URL',
			name: 'baseUrl',
			type: 'string',
			default: '',
			required: true,
			placeholder: 'https://cipp.yourdomain.com',
			description: 'The base URL of your CIPP deployment (API URL)',
		},
		{
			displayName: 'Azure AD Tenant ID',
			name: 'tenantId',
			type: 'string',
			default: '',
			required: true,
			placeholder: 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
			description: 'Your Azure AD Tenant ID (where the CIPP-SAM app registration lives)',
		},
		{
			displayName: 'Application (Client) ID',
			name: 'clientId',
			type: 'string',
			default: '',
			required: true,
			placeholder: 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx',
			description: 'The Application (Client) ID from your CIPP-SAM Azure AD App Registration',
		},
		{
			displayName: 'Client Secret',
			name: 'clientSecret',
			type: 'string',
			typeOptions: { password: true },
			default: '',
			required: true,
			description: 'The Client Secret from your CIPP-SAM Azure AD App Registration',
		},
	];

	// Credential testing is handled by the node via credentialTest method
	// This uses our custom OAuth flow to properly authenticate
}
