import { App, Editor, EditorPosition, Loc, MarkdownFileInfo, Plugin, TFile } from 'obsidian';
import { PowerPointFile } from './PowerPointFile';


export default class AliasPickerPlugin extends Plugin {

	async onload() {

		this.addCommand({
			id: 'powerpoint-notes:write-notes-to-slides',
			name: 'Write notes to slides',
			editorCheckCallback: (checking: boolean, editor: Editor, markdownFileInfo: MarkdownFileInfo) => {
				const file = markdownFileInfo.file;
				if (!file)
					return false;
				const headersWithBulletPoints = getHeadersWithBulletPointsInFile(file, this.app, editor);
				console.log(headersWithBulletPoints);
				if (headersWithBulletPoints === undefined)
					return false;
				if (checking)
					return headersWithBulletPoints !== undefined;

				const fileCache = this.app.metadataCache.getFileCache(file);
				const frontmatter = fileCache?.frontmatter;
				if (!frontmatter)
					return;
				const powerPointFilePath = frontmatter['powerPoint-file'];
				if (!powerPointFilePath || typeof powerPointFilePath !== 'string' || !powerPointFilePath.endsWith('.pptx'))
					return;
				async function writeToPowerPoint(headersWithBulletPoints: HeaderBulletPoints) {
					const powerPointFile = await PowerPointFile.loadAsync(powerPointFilePath);
					const notes = await powerPointFile.getAllNotesFromSlide(1);
					console.log(notes);
					const modifiedPath = powerPointFilePath.replace('.pptx', `_modified.pptx`);
					const clonedFile = await powerPointFile.copyAsync(modifiedPath);

					for (let i = 0; i < headersWithBulletPoints.length; i++) {
						const header = headersWithBulletPoints[i];
						const bulletPoints = header.bulletPoints;

						await clonedFile.writeNotesFileAsync(i + 1, bulletPoints);
					}
					await clonedFile.saveAsync();
				}
				writeToPowerPoint(headersWithBulletPoints);
			}
		});

	}

	onunload() {

	}
}

function toEditorPosition(pos: Loc): EditorPosition {
	return { line: pos.line, ch: pos.col };
}
function toEditorLineStartPosition(pos: Loc): EditorPosition {
	return { line: pos.line, ch: 0 };
}

type HeaderBulletPoints = {
	header: string;
	bulletPoints: string[];
}[];

function getHeadersWithBulletPointsInFile(file: TFile, app: App, editor: Editor): HeaderBulletPoints | undefined {
	const fileCache = app.metadataCache.getFileCache(file);
	if (!fileCache) return undefined;
	const headers = fileCache.headings;
	if (!headers || headers.length === 0) return [];
	const allListItems = fileCache.listItems;
	if (!allListItems || allListItems.length === 0) return [];
	const result: { header: string, bulletPoints: string[] }[] = [];
	for (const header of headers) {
		const nextHeader = headers.find(h => h.position.start.line > header.position.start.line);
		const listItems = allListItems.filter(item => item.position.start.line > header.position.start.line && item.position.start.line < (nextHeader?.position.start.line ?? Infinity));

		const listItemTexts = listItems.map(item => {
			const startPosition: EditorPosition = toEditorLineStartPosition(item.position.start);
			const endPosition: EditorPosition = toEditorPosition(item.position.end);
			return editor.getRange(startPosition, endPosition);
		});
		result.push({ header: header.heading, bulletPoints: listItemTexts });
	}
	return result;
}