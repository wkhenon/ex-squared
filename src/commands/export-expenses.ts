import ExcelJS from 'exceljs';
import { Editor, Notice, Plugin, TFile, TFolder } from 'obsidian';
import { parseExpenses } from '../expense-parser';

const EXPORT_FOLDER = 'Exports';

export function registerExpenseExportCommand(plugin: Plugin): void {
	plugin.addCommand({
		id: 'export-selected-expenses-to-excel',
		name: 'Export selected expenses to Excel',
		editorCallback: (editor: Editor) => {
			void exportSelectedExpenses(plugin, editor);
		},
	});

	plugin.registerEvent(
		plugin.app.workspace.on('editor-menu', (menu, editor) => {
			if (!editor.getSelection().trim()) return;

			menu.addItem((item) =>
				item
					.setTitle('Export selected expenses to Excel')
					.setIcon('table')
					.onClick(() => {
						void exportSelectedExpenses(plugin, editor);
					}),
			);
		}),
	);
}

async function exportSelectedExpenses(plugin: Plugin, editor: Editor): Promise<void> {
	const selection = editor.getSelection();
	if (!selection.trim()) {
		new Notice('Select the budget heading and expenses to export.');
		return;
	}
	const linkPosition = editor.getCursor('to');

	let expenses;
	try {
		expenses = parseExpenses(selection);
	} catch (error) {
		new Notice(error instanceof Error ? error.message : 'Could not read the selected expenses.');
		return;
	}

	const note = plugin.app.workspace.getActiveFile();
	if (!note) {
		new Notice('Open a note before exporting expenses.');
		return;
	}

	try {
		await ensureExportFolder(plugin);
		const dateStamp = formatDateStamp(new Date());
		const summaries = aggregateExpenses(expenses);
		const workbook = new ExcelJS.Workbook();
		workbook.creator = 'Ex Squared';
		workbook.subject = 'Expense summary by card and payment source';
		const worksheet = workbook.addWorksheet('Summary', {
			views: [{ state: 'frozen', ySplit: 2, showGridLines: false }],
		});
		worksheet.columns = [
			{ header: 'Card / payment source', key: 'label', width: 30 },
			{ header: 'Amount', key: 'amount', width: 15 },
		];
		worksheet.mergeCells('A1:B1');
		const title = worksheet.getCell('A1');
		title.value = 'Expense summary';
		title.font = { name: 'Aptos', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
		title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF202020' } };
		title.alignment = { horizontal: 'center', vertical: 'middle' };
		title.border = cellBorder();
		worksheet.getRow(1).height = 22;

		const header = worksheet.getRow(2);
		header.values = ['Card / payment source', 'Amount'];
		header.height = 20;
		for (const cell of [header.getCell(1), header.getCell(2)]) {
			cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
			cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF303030' } };
			cell.border = cellBorder();
		}
		worksheet.getColumn(2).numFmt = '$#,##0.00;[Red]-$#,##0.00';

		for (const card of summaries) {
			const cardRow = worksheet.addRow([`#${card.card}`, card.totalCents / 100]);
			cardRow.height = 20;
			cardRow.getCell(1).fill = {
				type: 'pattern',
				pattern: 'solid',
				fgColor: { argb: 'FF888888' },
			};
			cardRow.getCell(2).fill = {
				type: 'pattern',
				pattern: 'solid',
				fgColor: { argb: 'FF193A64' },
			};
			cardRow.getCell(2).alignment = { horizontal: 'right' };
			for (const cell of [cardRow.getCell(1), cardRow.getCell(2)]) {
				cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
				cell.border = cellBorder();
			}

			for (const source of card.sources) {
				const sourceRow = worksheet.addRow([`#${source.paymentSource}`, source.totalCents / 100]);
				sourceRow.getCell(1).alignment = { vertical: 'middle' };
				sourceRow.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' };
				for (const cell of [sourceRow.getCell(1), sourceRow.getCell(2)]) {
					cell.font = { color: { argb: 'FF000000' } };
					cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
					cell.border = cellBorder();
				}
			}
		}

		const fileName = await uniqueExportPath(plugin, note, dateStamp);
		const data = await workbook.xlsx.writeBuffer() as ArrayBuffer;
		const exportedFile = await plugin.app.vault.createBinary(fileName, data);
		const noteLink = plugin.app.fileManager.generateMarkdownLink(
			exportedFile,
			note.path,
			undefined,
			`Expense summary - ${dateStamp}`,
		);
		const separator = selection.endsWith('\n') ? '\n' : '\n\n';
		editor.replaceRange(`${separator}${noteLink}\n`, linkPosition);
		new Notice(`Exported ${summaries.length} card summaries and added a link to the note.`);
	} catch (error) {
		console.error('Ex Squared: failed to save expense workbook', error);
		new Notice('Could not save the Excel workbook in the Exports folder.');
	}
}

function aggregateExpenses(expenses: ReturnType<typeof parseExpenses>) {
	const cards = new Map<string, Map<string, number>>();
	for (const expense of expenses) {
		let sources = cards.get(expense.card);
		if (!sources) cards.set(expense.card, (sources = new Map<string, number>()));
		sources.set(
			expense.paymentSource,
			(sources.get(expense.paymentSource) ?? 0) + Math.round(expense.amount * 100),
		);
	}
	return [...cards.entries()]
		.map(([card, sources]) => {
			const paymentSources = [...sources.entries()]
				.map(([paymentSource, totalCents]) => ({ paymentSource, totalCents }))
				.sort((a, b) => a.paymentSource.localeCompare(b.paymentSource));
			return {
				card,
				totalCents: paymentSources.reduce((total, source) => total + source.totalCents, 0),
				sources: paymentSources,
			};
		})
		.sort((a, b) => a.card.localeCompare(b.card));
}

async function ensureExportFolder(plugin: Plugin): Promise<void> {
	const existing = plugin.app.vault.getAbstractFileByPath(EXPORT_FOLDER);
	if (!existing) {
		await plugin.app.vault.createFolder(EXPORT_FOLDER);
	} else if (!(existing instanceof TFolder)) {
		throw new Error(`A file named “${EXPORT_FOLDER}” already exists in the vault.`);
	}
}

async function uniqueExportPath(plugin: Plugin, note: TFile, dateStamp: string): Promise<string> {
	const noteName = note.basename.replace(/[\\/:*?"<>|]/g, '-');
	const stem = `${noteName} - expense summary - ${dateStamp}`;
	let suffix = 1;
	let path = `${EXPORT_FOLDER}/${stem}.xlsx`;
	while (plugin.app.vault.getAbstractFileByPath(path)) {
		suffix += 1;
		path = `${EXPORT_FOLDER}/${stem} (${suffix}).xlsx`;
	}
	return path;
}

function formatDateStamp(date: Date): string {
	const year = date.getFullYear();
	const month = String(date.getMonth() + 1).padStart(2, '0');
	const day = String(date.getDate()).padStart(2, '0');
	return `${year}-${month}-${day}`;
}

function cellBorder() {
	const edge = { style: 'thin' as const, color: { argb: 'FF000000' } };
	return { top: edge, right: edge, bottom: edge, left: edge };
}
