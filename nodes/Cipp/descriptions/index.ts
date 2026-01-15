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
import { groupFields, groupOperations } from './GroupDescription';
import { mailboxFields, mailboxOperations } from './MailboxDescription';
import { teamFields, teamOperations, voiceFields, voiceOperations } from './TeamDescription';
import { tenantFields, tenantOperations } from './TenantDescription';
import { userFields, userOperations } from './UserDescription';

export const operationFields = [
	...alertOperations,
	...applicationOperations,
	...autopilotOperations,
	...backupOperations,
	...deviceOperations,
	...groupOperations,
	...mailboxOperations,
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
	...groupFields,
	...mailboxFields,
	...scheduledItemFields,
	...teamFields,
	...tenantFields,
	...toolsFields,
	...userFields,
	...voiceFields,
];
