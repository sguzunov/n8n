import { Glue42Core } from '@glue42/core';

interface Glue42Api {
	interop: Glue42Core.Interop.API
};

interface UpdateExcelSheetOptions {
	workbook: string;
	worksheet: string;
	table: string;
	data: any
}

export const updateExcelSheet = async ({ interop }: Glue42Api, opts: UpdateExcelSheetOptions) => {
	const { data, workbook, worksheet, table } = opts;

  await interop.invoke('T42.ExcelScript.Table.AddRow', {
    table: table || 'Orders',
    values: JSON.stringify(
      [
				{
          "col": "Title",
          "value": data.title
        },
        {
          "col": "Description",
          "value": data.description
        },
        {
          "col": "Severity",
          "value": data.severity
        }
      ]),
    workbook: workbook,
    worksheet: worksheet
  })
}
