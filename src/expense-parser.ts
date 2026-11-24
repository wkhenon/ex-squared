export interface ExpenseRow {
	date: string;
	paymentSource: string;
	card: string;
	amount: number;
	description: string;
}

const BUDGET_HEADING = /^\s*(?:#{1,6}\s*)?(?:[*_]{1,3}\s*)?Budget for Period\s*:/i;
const DAY_LINE = /^\s*[-*+]\s+(\d{1,2}\/\d{1,2})\s+[A-Za-z]+\b/;
const TAG = /(^|\s)#([\p{L}\p{N}_/-]+)/gu;
const AMOUNT = /(?:^|\s)\$?(\d[\d,]*(?:\.\d{1,2})?)(?=\s|$)/;

export function parseExpenses(selection: string): ExpenseRow[] {
	const lines = selection.split(/\r?\n/);
	const headingIndex = lines.findIndex((line) => BUDGET_HEADING.test(line));
	if (headingIndex < 0) {
		throw new Error('Selection must start with a “Budget for Period” heading.');
	}

	const expenses: ExpenseRow[] = [];
	let currentDate = '';
	for (const line of lines.slice(headingIndex + 1)) {
		const dayMatch = line.match(DAY_LINE);
		if (dayMatch?.[1]) {
			currentDate = dayMatch[1];
			continue;
		}

		const tags = [...line.matchAll(TAG)].map((match) => match[2]).filter(Boolean);
		if (tags.length < 2) continue;

		const withoutTags = line.replace(TAG, '$1').trim();
		const amountMatch = withoutTags.match(AMOUNT);
		if (!amountMatch || !currentDate) continue;

		const amount = Number(amountMatch[1]?.replaceAll(',', ''));
		if (!Number.isFinite(amount)) continue;
		const amountIndex = amountMatch.index ?? 0;
		const amountTextLength = amountMatch[0].length;
		const description = withoutTags.slice(amountIndex + amountTextLength).trim();
		expenses.push({
			date: currentDate,
			paymentSource: tags[0] ?? '',
			card: tags[1] ?? '',
			amount,
			description,
		});
	}

	if (expenses.length === 0) {
		throw new Error('No expenses found. Include dated day bullets and expense bullets with two tags and an amount.');
	}
	return expenses;
}
