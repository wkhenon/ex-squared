import { Plugin } from 'obsidian';
import { registerExpenseExportCommand } from './commands/export-expenses';

export default class ExSquaredPlugin extends Plugin {
	onload() {
		registerExpenseExportCommand(this);
	}
}
