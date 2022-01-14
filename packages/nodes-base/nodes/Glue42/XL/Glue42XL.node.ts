import { IExecuteFunctions } from 'n8n-core';
import {
	INodeType,
	INodeTypeDescription,
	// LoggerProxy as Logger
} from 'n8n-workflow';

import GlueFactory from '@glue42/core';
import { updateExcelSheet } from './Glue42XL';
// import { updateExcelSheet } from './Glue42XL';

export class Glue42XL implements INodeType {
	public readonly description: INodeTypeDescription = {
		displayName: 'Glue42 XL',
		name: 'Glue42XL',
		icon: 'file:glue42XL.svg',
		group: ['input', 'output'],
		version: 1,
		description: 'Read, update and write data to Microsoft Excel via Glue42.',
		eventTriggerDescription: '',
		defaults: {
			name: 'Glue42XL',
			color: '#00FF00',
		},
		inputs: ['main'],
		outputs: ['main'],
		properties: [
			// TODO: Supply options.
		],
	};


	async execute(this: IExecuteFunctions): Promise<null> {

		console.log('+++ Glue42XL....', );

		const glue = await GlueFactory({
			application: 'n8n-XL-node',
			gateway: {
				ws: 'ws://localhost:8385/',
				protocolVersion: 3
			},
			auth: {
				username: 'suzunov',
				password: ''
			}
		});
		console.log('+++ Glue42XL done', glue.version);

		const inputData = this.getInputData();

		glue.interop.invoke('n8n-test', {
			time: Date.now(),
			source: 'Glue42XL',
			inputData
		}).catch(console.error);

		await updateExcelSheet(glue, {
			data: inputData[0].json,
			table: 'Notifications',
			workbook: 'C:\\Users\\suzunov\\Desktop\\MacroTemplate.xlsm',
			worksheet: 'Notifications'
		})
		.catch(console.error);

		return null;
	}
}
