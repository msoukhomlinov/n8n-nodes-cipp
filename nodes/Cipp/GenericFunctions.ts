import type {
	IDataObject,
	IExecuteFunctions,
	IHookFunctions,
	IHttpRequestMethods,
	IHttpRequestOptions,
	ILoadOptionsFunctions,
	JsonObject,
} from 'n8n-workflow';

import { NodeApiError, NodeOperationError } from 'n8n-workflow';

import type { IAuthToken, ICippCredentials, ITenant } from './types';

// Token cache to avoid repeated authentication calls
const tokenCache = new Map<string, IAuthToken>();

function getCacheKey(credentials: ICippCredentials): string {
	return `${credentials.clientId}:${credentials.tenantId}`;
}

/**
 * Normalizes and joins URL parts
 */
function joinUrl(baseUrl: string, endpoint: string): string {
	const normalizedBase = baseUrl.endsWith('/') ? baseUrl.slice(0, -1) : baseUrl;
	const normalizedEndpoint = endpoint.startsWith('/') ? endpoint : `/${endpoint}`;
	return `${normalizedBase}${normalizedEndpoint}`;
}

/**
 * Gets OAuth2 access token from Azure AD using client credentials flow
 */
export async function getAccessToken(
	this: IExecuteFunctions | ILoadOptionsFunctions | IHookFunctions,
	credentials: ICippCredentials,
): Promise<string> {
	const cacheKey = getCacheKey(credentials);
	const cached = tokenCache.get(cacheKey);

	// Return cached token if still valid (with 5-minute buffer)
	if (cached && cached.expiresAt > Date.now() + 300000) {
		return cached.accessToken;
	}

	const tokenUrl = `https://login.microsoftonline.com/${credentials.tenantId}/oauth2/v2.0/token`;
	const scope = `api://${credentials.clientId}/.default`;

	const body = `grant_type=client_credentials&client_id=${encodeURIComponent(credentials.clientId)}&client_secret=${encodeURIComponent(credentials.clientSecret)}&scope=${encodeURIComponent(scope)}`;

	const options: IHttpRequestOptions = {
		method: 'POST',
		url: tokenUrl,
		headers: {
			'Content-Type': 'application/x-www-form-urlencoded',
		},
		body,
		json: true,
	};

	try {
		const response = (await this.helpers.httpRequest(options)) as IDataObject;

		if (!response.access_token) {
			throw new NodeOperationError(this.getNode(), 'No access token in response');
		}

		const expiresIn = (response.expires_in as number) || 3600;
		const authToken: IAuthToken = {
			accessToken: response.access_token as string,
			expiresAt: Date.now() + expiresIn * 1000,
		};

		tokenCache.set(cacheKey, authToken);
		return authToken.accessToken;
	} catch (error) {
		const err = error as IDataObject;
		const errorResponse = (error || {}) as JsonObject;

		// Clear cache on auth failure
		tokenCache.delete(cacheKey);

		throw new NodeApiError(this.getNode(), errorResponse, {
			message: 'Failed to authenticate with CIPP',
			description:
				(err.error_description as string) ||
				(err.message as string) ||
				'Check your Azure AD credentials and ensure the app registration is configured correctly',
		});
	}
}

/**
 * Makes an authenticated request to the CIPP API
 */
export async function cippApiRequest(
	this: IExecuteFunctions | ILoadOptionsFunctions | IHookFunctions,
	method: IHttpRequestMethods,
	endpoint: string,
	body: IDataObject = {},
	query: IDataObject = {},
): Promise<IDataObject> {
	const credentials = (await this.getCredentials('cippApi')) as unknown as ICippCredentials;

	// Normalize base URL
	const baseUrl = credentials.baseUrl.replace(/\/$/, '');
	const accessToken = await getAccessToken.call(this, credentials);

	const url = joinUrl(baseUrl, endpoint);

	const options: IHttpRequestOptions = {
		method,
		url,
		headers: {
			Authorization: `Bearer ${accessToken}`,
			Accept: 'application/json',
			'Content-Type': 'application/json',
		},
		qs: query,
		json: true,
	};

	if (Object.keys(body).length > 0) {
		options.body = body;
	}

	try {
		const response = await this.helpers.httpRequest(options);
		return response as IDataObject;
	} catch (error: unknown) {
		const errorResponse = (error || {}) as JsonObject;
		const err = error as {
			statusCode?: number;
			response?: { headers?: IDataObject; status?: number; statusCode?: number };
			error?: { message?: string; error_description?: string };
			message?: string;
		};

		const statusCode = err.statusCode || err.response?.status || err.response?.statusCode;

		if (statusCode === 401) {
			// Clear token cache on auth failure
			tokenCache.delete(getCacheKey(credentials));
			throw new NodeApiError(this.getNode(), errorResponse, {
				message: 'Authentication failed',
				description:
					'Your access token has expired or is invalid. Check your CIPP API credentials.',
			});
		}

		if (statusCode === 403) {
			throw new NodeApiError(this.getNode(), errorResponse, {
				message: 'Permission denied',
				description:
					'Your API client does not have permission to perform this action. Check your CIPP API client role.',
			});
		}

		if (statusCode === 404) {
			throw new NodeApiError(this.getNode(), errorResponse, {
				message: 'Resource not found',
				description:
					err.error?.message ||
					'The requested resource does not exist. Check your tenant filter and IDs.',
			});
		}

		if (statusCode === 429) {
			throw new NodeApiError(this.getNode(), errorResponse, {
				message: 'Rate limit exceeded',
				description: 'Too many requests. Please wait before retrying.',
			});
		}

		const errorMessage =
			err.error?.message || err.error?.error_description || err.message || 'Unknown error';

		throw new NodeApiError(this.getNode(), errorResponse, {
			message: `CIPP API Error (Status ${statusCode || 'Unknown'})`,
			description: errorMessage,
		});
	}
}

/**
 * Makes an authenticated request and handles pagination for list endpoints
 */
export async function cippApiRequestAllItems(
	this: IExecuteFunctions | ILoadOptionsFunctions,
	method: IHttpRequestMethods,
	endpoint: string,
	body: IDataObject = {},
	query: IDataObject = {},
): Promise<IDataObject[]> {
	// CIPP API typically returns all results in a single response
	// but we handle potential pagination here for future compatibility
	const response = await cippApiRequest.call(this, method, endpoint, body, query);

	// Handle different response formats
	if (Array.isArray(response)) {
		return response as IDataObject[];
	}

	if (response.Results && Array.isArray(response.Results)) {
		return response.Results as IDataObject[];
	}

	if (response.value && Array.isArray(response.value)) {
		return response.value as IDataObject[];
	}

	// Single item response
	return [response];
}

/**
 * Fetches the list of tenants from CIPP
 */
export async function getTenantList(
	this: IExecuteFunctions | ILoadOptionsFunctions | IHookFunctions,
): Promise<ITenant[]> {
	const response = await cippApiRequest.call(this, 'GET', '/api/ListTenants', {}, {});

	if (Array.isArray(response)) {
		return response as ITenant[];
	}

	if (response.Results && Array.isArray(response.Results)) {
		return response.Results as ITenant[];
	}

	return [];
}

/**
 * Helper to get a resource locator value (handles both list and id modes)
 */
export function getResourceLocatorValue(value: unknown): string {
	if (typeof value === 'string') {
		return value;
	}

	if (value && typeof value === 'object') {
		const locator = value as { mode: string; value: string };
		return locator.value || '';
	}

	return '';
}
