import { google, sheets_v4 } from 'googleapis';
import { mkdir, readdir, readFile, rename } from 'node:fs/promises';
import { join } from 'node:path';
import { PDFParse } from 'pdf-parse';

interface Config {
	spreadsheetId: string;
	logTableName: string;
	stockTableName: string;
	mapTableName: string;
	transactionType: string;
	transactionCategory: string;
	transactionReason: string;
}

type ProductMap = Record<string, string>;
type Store = 'Coop' | 'Migros';

interface Transaction {
	date: string;
	store: Store;
}

interface Item {
	name: string;
	quantity: number;
	unitPrice: number;
}

type ItemWithTransaction = Item & Transaction;

interface Receipt extends Transaction {
	filePath: string;
	id: string;
	items: Item[];
	billTotal: number;
}

const DATE_REGEX: Record<Store, RegExp> = {
	Coop: /(\d{2})\.(\d{2})\.(\d{2})/,
	Migros: /(\d{2})\.(\d{2})\.20(\d{2})/,
};

const {
	spreadsheetId,
	logTableName,
	stockTableName,
	mapTableName,
	transactionType,
	transactionCategory,
	transactionReason,
} = JSON.parse(await readFile('config.json', 'utf-8')) as Config;

async function parseReceipt(filePath: string, map: ProductMap): Promise<Receipt> {
	const parser = new PDFParse({
		data: await readFile(filePath),
	});
	const {fingerprints} = await parser.getInfo();
	if (!fingerprints) {
		throw new Error('PDF fingerprint not found');
	}
	const {text} = await parser.getText();
	let store = 'Coop' as Store;
	if (text.includes('Migros')) {
		store = 'Migros';
	}
	const lines = text
		.split('\n')
		.map(l => l.trim())
		.filter(Boolean)
		.slice(3);
	const items: Item[] = [];
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
				filePath,
				id: fingerprints.filter(Boolean).join(''),
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
		if (!(name in map)) {
			throw new Error(`Product "${name}" not found in product map!`);
		}
		items.push({
			name: map[name] || name,
			quantity: parseFloat(itemMatch[2]),
			unitPrice: parseFloat(itemMatch[3]),
		});
	}
	throw new Error('Total amount not found in receipt');
}

const updateLog = async (
	sheets: sheets_v4.Sheets,
	receipts: Receipt[]
) => sheets.spreadsheets.values.append({
	spreadsheetId,
	range: logTableName,
	valueInputOption: 'USER_ENTERED',
	requestBody: {
		values: receipts.map(receipt => [
			transactionType,
			receipt.date,
			receipt.billTotal,
			transactionCategory,
			transactionReason,
			`${receipt.store}: ${receipt.items
				.map(i => (i.quantity === 1 ?
					i.name :
					`${i.name} x${i.quantity}`))
				.join(', ')}`,
		]),
	},
});

async function updateStock(sheets: sheets_v4.Sheets, items: ItemWithTransaction[]) {
	const res = await sheets.spreadsheets.values.get({
		spreadsheetId,
		range: `${stockTableName}!A:E`,
	});
	const rows = res.data.values ?? [];
	const index = Object.fromEntries(rows
		.map((row, index) => [row[0], index])
		.filter(([name, index]) => name && !isNaN(rows[index][1])));
	const updates: Record<number, [number, string]> = {};
	for (const item of items) {
		if (!(item.name in index)) {
			continue;
		}
		if (updates[index[item.name] + 1]) {
			updates[index[item.name] + 1][0] += item.quantity;
			continue;
		}
		updates[index[item.name] + 1] = [
			Number(rows[index[item.name]][1] ?? 0) + item.quantity,
			item.date,
		];
		const oldStore = rows[index[item.name]][3];
		const oldUnitPrice = Number(rows[index[item.name]][4] ?? 0);
		if (item.unitPrice !== oldUnitPrice) {
			console.warn(`Unit price for "${item.name}" changed from ${oldUnitPrice} to ${item.unitPrice}.`);
		}
		if (item.store !== oldStore) {
			console.warn(`Item "${item.name}" was bought in ${item.store} rather than ${oldStore}.`);
		}
	}
	await sheets.spreadsheets.values.batchUpdate({
		spreadsheetId,
		requestBody: {
			valueInputOption: 'USER_ENTERED',
			data: Object.entries(updates).flatMap(([row, [quantity, date]]) => [
				{
					range: `${stockTableName}!B${row}`,
					values: [[quantity]],
				},
				{
					range: `${stockTableName}!G${row}`,
					values: [[date]],
				},
			]),
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
const receipts = await Promise.all((await readdir('files'))
	.filter(fileName => fileName.toLowerCase().endsWith('.pdf'))
	.map(fileName => parseReceipt(join('files', fileName), map)));
console.info('Parsed receipts:', receipts);
if (process.argv.includes('--dry-run')) {
	process.exit(0);
}
await updateLog(sheets, receipts);
await updateStock(sheets, receipts
	.flatMap(({date, items, store}) => items
		.map(item => ({...item, date, store}))));
await mkdir('processed', {recursive: true});
await Promise.all(receipts.map(r => rename(r.filePath, join('processed', `${r.id}.pdf`))));
