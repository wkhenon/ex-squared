import { Editor, Notice, Plugin, TFile, TFolder } from 'obsidian';
import { parseExpenses } from '../expense-parser';
import { CardSummary, createExpenseSummaryWorkbook } from '../workbook/xlsx-writer';

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

		const fileName = await uniqueExportPath(plugin, note, dateStamp);
		const data = createExpenseSummaryWorkbook(summaries);
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

function aggregateExpenses(expenses: ReturnType<typeof parseExpenses>): CardSummary[] {
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
