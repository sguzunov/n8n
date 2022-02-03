import { ITriggerFunctions } from 'n8n-core';
import {
	INodeType,
	INodeTypeDescription,
	ITriggerResponse
} from 'n8n-workflow';

import { initializeGlue } from '../GlueUtils';

export class Glue42Notifications implements INodeType {

	public readonly description: INodeTypeDescription = {
		displayName: 'Glue42 Notifications',
		name: 'Glue42Notifications',
		icon: 'file:glue42Notifications.svg',
		group: ['trigger'],
		version: 1,
		description: 'Triggers the flow when a Glue42 Notification is raised.',
		eventTriggerDescription: '',
		defaults: {
			name: 'Glue42 Notifications',
			color: '#00FF00',
		},
		inputs: [],
		outputs: ['main'],
		outputNames: ['Notification'],
		properties: [],
	};


	async trigger(this: ITriggerFunctions): Promise<ITriggerResponse> {

		const glue = await initializeGlue();
		console.log(`Glue42 SDK initialized ,version ${glue.version}`);

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
