import { Glue42Core } from '@glue42/core';

interface Glue42Api {
	interop: Glue42Core.Interop.API;
};

const isOpenWorkbook = async (
	{ interop }: Glue42Api,
	args: { workbook: string; worksheet: string; }
) => {
	return interop.invoke(
		'T42.ExcelScript.Workbook.IsOpen',
		{
			...args,
			includeUnsaved: true
		})
		.then(({ returned: { result } }) => JSON.parse(result));
}

const createWorkbook = async ({ interop }: Glue42Api, args: { workbook: string; }) => {
	return interop.invoke('T42.ExcelScript.Workbook.Create', args)
};

const createWorksheet = async (
	{ interop }: Glue42Api,
	args: { workbook: string; worksheet: string; }
) => {
	return interop.invoke('T42.ExcelScript.Worksheet.Create', args);
};

interface XLCreateTableArgs {
	columns: string[];
	positionRange: string;
	tableName: string;
	tableStyle?: string;
	value?: Array<[]>;
	workbook: string;
	worksheet: string;
}

const createTable = async ({ interop }: Glue42Api, args: XLCreateTableArgs) => {
	return interop.invoke('T42.XL.CreateTable', args);
};

interface XLWriteTableRowsArgs {
	tableName: string;
	value: Array<Array<any>>;
	workbook: string;
	worksheet: string;
	rowPosition?: number;
}

const writeTableRows = async ({ interop }: Glue42Api, args: XLWriteTableRowsArgs) => {
	return interop.invoke('T42.XL.WriteTableRows', args);
};

const tableExists = async ({ interop }: Glue42Api, { table, ...args }: any) => {
	return interop.invoke('T42.ExcelScript.Table.GetInformation', args)
		.then(({ returned: { result } }) => JSON.parse(result))
		.then(({ tables }) => tables.some(({ name }: any) => name === table));
};


export interface UpdateExcelSheetOptions {
	fileLocation: string;
	workbookName: string;
	worksheetName: string;
	table: string;
	data: Record<string, any>;
	columns: string[];
}

export const appendToWorksheet = async (glue: Glue42Api, options: UpdateExcelSheetOptions) => {
	const { data, workbookName, worksheetName: worksheet, table, fileLocation, columns } = options;
	const workbook = `${fileLocation}\\${workbookName}`;

	console.log('[Glue42Excel] appendToWorksheet() options: ', options);

	const { isOpen } = await isOpenWorkbook(glue, { workbook, worksheet });
	if (isOpen == false) {
		console.log(`Book ${workbook} will be created.`)

		await createWorkbook(glue, { workbook });

		console.log(`Book ${workbook} created - success.`)
	} else {
		console.log(`Book ${workbook} already exists.`)
	}


	const { hasSheet } = await isOpenWorkbook(glue, { workbook, worksheet });
	if (hasSheet == false) {
		console.log(`Sheet ${worksheet} will be created.`)

		await createWorksheet(glue, { workbook, worksheet });

		console.log(`Sheet ${worksheet} created - success.`);
	} else {
		console.log(`Sheet ${worksheet} already exists.`)
	}

	const hasTable = await tableExists(glue, {
		workbook,
		worksheet,
		table
	});

	const dataTransformed: any = columns.map((col: string) => data[col]);

	if (hasTable == false) {
		console.log('Table will be created.');

		await createTable(glue, {
			columns: columns,
			positionRange: 'A1',
			tableName: table,
			value: [dataTransformed],
			workbook: workbookName,
			worksheet
		});
		console.log('Table created.');

		return;
	}

	console.log('Table already created.');
	await writeTableRows(glue, {
		tableName: table,
		value: [dataTransformed],
		workbook: workbookName,
		worksheet
	});
};

// Common Glue42 utils.

const startApp = async ({ interop }: Glue42Api, name: string) => interop.invoke('T42.ACS.StartApplication', { Name: name });

const matchExcelProtocol = ({ name }: { name: string }) => name.toLocaleLowerCase().startsWith('t42.excel');

export const startExcel = async ({ interop }: Glue42Api): Promise<boolean> => {

	const excelStarted = interop.methods().some(matchExcelProtocol);

	return excelStarted ?
		Promise.resolve(true) :
		new Promise(async (resolve) => {
			await startApp({ interop }, 'excel');

			const timeout = setTimeout(() => {
				unsubscribe();
				clearTimeout(timeout);
				resolve(false);
			}, 5 * 1000);

			var unsubscribe = interop.serverMethodAdded(({ method }) => {
				if (matchExcelProtocol(method)) {
					unsubscribe();
					clearTimeout(timeout);
					resolve(true);
				}
			});
		});
};
