import { alertFields, alertOperations } from './AlertDescription';
import { applicationFields, applicationOperations } from './ApplicationDescription';
import {
	backupFields,
	backupOperations,
	scheduledItemFields,
	scheduledItemOperations,
	toolsFields,
	toolsOperations,
} from './CippDescription';
import {
	autopilotFields,
	autopilotOperations,
	deviceFields,
	deviceOperations,
} from './DeviceDescription';
import { gdapFields, gdapOperations } from './GdapDescription';
import { groupFields, groupOperations } from './GroupDescription';
import { identityFields, identityOperations } from './IdentityDescription';
import { mailboxFields, mailboxOperations } from './MailboxDescription';
import { onedriveFields, onedriveOperations } from './OneDriveDescription';
import { policyFields, policyOperations } from './PolicyDescription';
import { quarantineFields, quarantineOperations } from './QuarantineDescription';
import { teamFields, teamOperations, voiceFields, voiceOperations } from './TeamDescription';
import { tenantFields, tenantOperations } from './TenantDescription';
import { userFields, userOperations } from './UserDescription';

export const operationFields = [
	...alertOperations,
	...applicationOperations,
	...autopilotOperations,
	...backupOperations,
	...deviceOperations,
	...gdapOperations,
	...groupOperations,
	...identityOperations,
	...mailboxOperations,
	...onedriveOperations,
	...policyOperations,
	...quarantineOperations,
	...scheduledItemOperations,
	...teamOperations,
	...tenantOperations,
	...toolsOperations,
	...userOperations,
	...voiceOperations,
];

export const resourceFields = [
	...alertFields,
	...applicationFields,
	...autopilotFields,
	...backupFields,
	...deviceFields,
	...gdapFields,
	...groupFields,
	...identityFields,
	...mailboxFields,
	...onedriveFields,
	...policyFields,
	...quarantineFields,
	...scheduledItemFields,
	...teamFields,
	...tenantFields,
	...toolsFields,
	...userFields,
	...voiceFields,
];
