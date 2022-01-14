import { ITriggerFunctions } from 'n8n-core';
import {
	INodeType,
	INodeTypeDescription,
	ITriggerResponse,
} from 'n8n-workflow';

import GlueFactory from '@glue42/core';

export class Glue42Notifications implements INodeType {
	public readonly description: INodeTypeDescription = {
		displayName: 'Glue42 Notifications',
		name: 'Glue42Notifications',
		icon: 'file:glue42Notifications.svg',
		group: ['trigger'],
		version: 1,
		description: 'Triggers a when a Glue42 Notification is raised.',
		eventTriggerDescription: '',
		defaults: {
			name: 'Glue42Notifications',
			color: '#00FF00',
		},
		inputs: [],
		outputs: ['main'],
		properties: [
			// TODO: Supply options.
		],
	};


	async trigger(this: ITriggerFunctions): Promise<ITriggerResponse> {

		const glue = await GlueFactory({
			application: 'n8n-notifications-trigger-node',
			gateway: {
				ws: 'ws://localhost:8385/',
				protocolVersion: 3
			},
			auth: {
				username: 'suzunov',
				password: ''
			}
		});

		// The trigger function to execute when a notification got raised.
		const executeTrigger = (items: any[]) => {
			if(items && items.length) {
				this.emit([
					this.helpers.returnJsonArray(items)
				]);
			}
		};

		console.log('+++ subscribeFor T42.GNS.Subscribe.Notifications....', );
		const subscription = await glue.interop.subscribe('T42.GNS.Subscribe.Notifications');
		subscription.onData(({ data: { items } }) => {
			executeTrigger(items);
		});

		// Unsubscribe from receiving notifications.
		async function closeFunction() {
			console.log('+++ closeSubscription T42.GNS.Subscribe.Notifications....', );
			subscription?.close();
		}

		// async function manualTriggerFunction() {
		// 	executeTrigger();
		// }

		return {
			closeFunction,
			// manualTriggerFunction,
		};
	}
}
