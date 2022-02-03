import { IExecuteFunctions } from 'n8n-core';
import {
	INodeType,
	INodeTypeDescription,
	// LoggerProxy as Logger
} from 'n8n-workflow';

import { appendToWorksheet, startExcel } from './Glue42Excel';
import { initializeGlue } from '../GlueUtils';

export class Glue42Excel implements INodeType {
	public readonly description: INodeTypeDescription = {
		displayName: 'Glue42 Excel',
		name: 'Glue42Excel',
		icon: 'file:glue42Excel.svg',
		group: ['input', 'output'],
		version: 1,
		description: 'Read, update and write data to Microsoft Excel via Glue42.',
		eventTriggerDescription: '',
		defaults: {
			name: 'Glue42Excel',
			color: '#00FF00',
		},
		inputs: ['main'],
		outputs: ['main'],
		properties: [
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				options: [
					{
						name: 'Workbook',
						value: 'workbook',
					},
					{
						name: 'Worksheet',
						value: 'worksheet',
					},

				],
				default: 'worksheet'
			},

			// --------------
			// Operations for a workbook.
			// --------------
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				displayOptions: {
					show: {
						resource: [
							'workbook',
						],
					},
				},
				options: [
					{
						name: 'Create',
						value: 'create',
						description: 'Create a workbook.',
					},
				],
				default: 'create',
				description: 'The operation to perform.',
			},

			// --------------
			// Operations for a worksheet.
			// --------------
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				displayOptions: {
					show: {
						resource: [
							'worksheet',
						],
					},
				},
				options: [
					{
						name: 'Append',
						value: 'append',
						description: 'Append data to a sheet. If the sheet does not exists, it is automatically created.',
					},
					{
						name: 'Create',
						value: 'create',
						description: 'Create a new sheet',
					},
					{
						name: 'Delete',
						value: 'delete',
						description: 'Delete columns and rows from a sheet',
					},
					{
						name: 'Lookup',
						value: 'lookup',
						description: 'Look up a specific column value and return the matching row',
					},
					{
						name: 'Read',
						value: 'read',
						description: 'Read data from a sheet',
					},
					{
						name: 'Remove',
						value: 'remove',
						description: 'Remove a sheet',
					},
					{
						name: 'Update',
						value: 'update',
						description: 'Update rows in a sheet',
					},
				],
				default: 'append',
				description: 'The operation to perform.',
			},

			// --------------
			// All
			// --------------
			{
				displayName: 'File Location',
				name: 'fileLocation',
				type: 'string',
				default: 'C:\\Workspace\\PoC\\n8n\\files',
				required: true
			},
			{
				displayName: 'Workbook Name',
				name: 'workbookName',
				type: 'string',
				default: 'Social Media.xlsx',
				required: true
			},
			{
				displayName: 'Worksheet Name',
				name: 'worksheetName',
				type: 'string',
				default: 'Musk about TSLA',
				required: true,
				displayOptions: {
					show: {
						resource: [
							'worksheet',
						],
					},
				},
			},
			{
				displayName: 'Table Name',
				name: 'tableName',
				type: 'string',
				default: 'Tweets',
				required: true,
				displayOptions: {
					show: {
						resource: [
							'worksheet',
						],
					},
				},
			},
			{
				displayName: 'Table Columns',
				name: 'tableColumns',
				type: 'string',
				default: 'Date; Tweet',
				required: true,
				displayOptions: {
					show: {
						resource: [
							'worksheet',
						],
					},
				},
			}
		],
	};

	public async execute(this: IExecuteFunctions): Promise<any> {

		console.log('[Glue42Excel] executing...');

		const glue = await initializeGlue();

		console.log('[Glue42Excel] glue initialized ', glue.version);

		const resource = this.getNodeParameter('resource', 0) as string;
		const operation = this.getNodeParameter('operation', 0) as string;

		if(resource === 'workbook') {
			// Out of demo scope.
			return null;
		}

		// Selected resource is "worksheet".

		if(operation !== 'append') {
			// Out of demo scope. Only "append" is supported.
			return null;
		}

		const excelStarted = await startExcel(glue).catch(console.error)
		if(!excelStarted) {
			console.error('Excel not started');
			return null;
		}

		const fileLocation = this.getNodeParameter('fileLocation', 0) as string;
		const workbookName = this.getNodeParameter('workbookName', 0) as string;
		const worksheetName = this.getNodeParameter('worksheetName', 0) as string;
		const tableName = this.getNodeParameter('tableName', 0) as string;
		const tableColumns = this.getNodeParameter('tableColumns', 0) as string;

		const data = this.getInputData()[0].json;
		const columns = tableColumns.split(/;/).map((col) => col.trim())

		await appendToWorksheet(glue, {
			data,
			columns,
			table: tableName,
			workbookName,
			worksheetName,
			fileLocation,
		})
		.catch(console.error);

		return null;
	}
}
