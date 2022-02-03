import { Glue42Core } from '@glue42/core';
import GlueFactory from '@glue42/core';

export type Glue42Api  = Glue42Core.GlueCore;

export const initializeGlue = async (): Promise<Glue42Api> => GlueFactory({
	application: 'n8n-excel-node',
	gateway: {
		ws: 'ws://localhost:8385/',
		protocolVersion: 3
	},
	auth: {
		username: 'suzunov',
		password: ''
	}
});
