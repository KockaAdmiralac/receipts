import { google, sheets_v4 } from 'googleapis';
import { mkdir, readdir, readFile, rename } from 'node:fs/promises';
import { join } from 'node:path';
import { PDFParse } from 'pdf-parse';

type ProductMap = Record<string, string>;
type Store = 'Coop' | 'Migros';

const DATE_REGEX: Record<Store, RegExp> = {
	Coop: /(\d{2})\.(\d{2})\.(\d{2})/,
	Migros: /(\d{2})\.(\d{2})\.20(\d{2})/,
};

interface Config {
	spreadsheetId: string;
	logTableName: string;
	stockTableName: string;
	mapTableName: string;
	transactionType: string;
	transactionCategory: string;
	transactionReason: string;
}

interface ReceiptItem {
	name: string;
	quantity: number;
	unitPrice: number;
}

interface ReceiptData {
	date: string;
	store: Store;
	items: ReceiptItem[];
	billTotal: number;
}

const {
	spreadsheetId,
	logTableName,
	stockTableName,
	mapTableName,
	transactionType,
	transactionCategory,
	transactionReason,
} = JSON.parse(await readFile('config.json', 'utf-8')) as Config;

async function parseReceipt(text: string, map: ProductMap): Promise<ReceiptData> {
	let store = 'Coop' as Store;
	if (text.includes('Migros')) {
		store = 'Migros';
	}
	const lines = text
		.split('\n')
		.map(l => l.trim())
		.filter(Boolean)
		.slice(3);
	const items: ReceiptItem[] = [];
	const dateLine = text.match(DATE_REGEX[store]);
	if (!dateLine) {
		throw new Error('Date not found in receipt');
	}
	const date = `${Number(dateLine[1])}/${Number(dateLine[2])}/20${dateLine[3]}`;
	for (const line of lines) {
		if (line.startsWith('Total CHF')) {
			const match = line.match(/Total CHF\s+([\d.]+)/);
			if (!match) {
				throw new Error('Total amount row unparseable');
			}
			return {
				date,
				store,
				billTotal: parseFloat(match[1]),
				items,
			};
		}
		const itemMatch = line
			.match(/^(.*)\s+(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)\s+(?:\d+(?:\.\d+)?\s+)?(\d+(?:\.\d+)?)\s+\d+\s*$/);
		if (!itemMatch) {
			continue;
		}
		const name = itemMatch[1].trim();
		items.push({
			name: map[name] || name,
			quantity: parseFloat(itemMatch[2]),
			unitPrice: parseFloat(itemMatch[3]),
		});
	}
	throw new Error('Total amount not found in receipt');
}

const writeReceiptToSheets = async (
	sheets: sheets_v4.Sheets,
	data: ReceiptData
) => sheets.spreadsheets.values.append({
	spreadsheetId,
	range: logTableName,
	valueInputOption: 'USER_ENTERED',
	requestBody: {
		values: [[
			transactionType,
			data.date,
			data.billTotal,
			transactionCategory,
			transactionReason,
			`${data.store}: ${data.items
				.map(i => (i.quantity === 1 ?
					i.name :
					`${i.name} x${i.quantity}`))
				.join(', ')}`,
		]],
	},
});

async function updateStock(sheets: sheets_v4.Sheets, items: ReceiptItem[]) {
	const res = await sheets.spreadsheets.values.get({
		spreadsheetId,
		range: `${stockTableName}!A:B`,
	});
	const rows = res.data.values ?? [];
	const index = Object.fromEntries(rows
		.map((row, index) => [row[0], index])
		.filter(([name, index]) => name && !isNaN(rows[index][1])));
	await sheets.spreadsheets.values.batchUpdate({
		spreadsheetId,
		requestBody: {
			valueInputOption: 'USER_ENTERED',
			data: items.filter(item => item.name in index).map(item => ({
				range: `${stockTableName}!B${index[item.name] + 1}`,
				values: [[
					Number(rows[index[item.name]][1] ?? 0) + item.quantity,
				]],
			})),
		},
	});
}

const sheets = google.sheets({
	auth: new google.auth.GoogleAuth({
		credentials: JSON.parse(await readFile('credentials.json', 'utf-8')),
		scopes: ['https://www.googleapis.com/auth/spreadsheets'],
	}),
	version: 'v4',
});
const map = Object.fromEntries((await sheets.spreadsheets.values.get({
	spreadsheetId,
	range: mapTableName,
})).data.values ?? []) as ProductMap;
await mkdir('processed', {recursive: true});
const items: ReceiptItem[] = [];
for (const file of await readdir('files', {
	withFileTypes: true
})) {
	if (!file.isFile() || !file.name.toLowerCase().endsWith('.pdf')) {
		continue;
	}
	const path = join('files', file.name);
	const parser = new PDFParse({
		data: await readFile(path),
	});
	const {fingerprints} = await parser.getInfo();
	if (!fingerprints) {
		throw new Error('PDF fingerprint not found');
	}
	const newPath = join('processed', `${fingerprints
		.filter(Boolean)
		.join('')}.pdf`);
	const {text} = await parser.getText();
	const receipt = await parseReceipt(text, map);
	await rename(path, newPath);
	console.info('Recording receipt', receipt);
	await writeReceiptToSheets(sheets, receipt);
	items.push(...receipt.items);
}
console.info('Updating stock...');
await updateStock(sheets, items);
