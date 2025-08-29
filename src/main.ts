import { CachedMetadata, Editor, EditorPosition, HeadingCache, ListItemCache, Loc, MarkdownFileInfo, Notice, Plugin, Pos } from 'obsidian';
import { PowerPointFile } from './PowerPointFile';


export default class AliasPickerPlugin extends Plugin {

	async onload() {

		this.addCommand({
			id: 'powerpoint-notes:write-notes-to-slides',
			name: 'Write notes to slides',
			editorCheckCallback: (checking: boolean, editor: Editor, markdownFileInfo: MarkdownFileInfo) => {
				const file = markdownFileInfo.file; if (!file) return false;

				const fileCache = this.app.metadataCache.getFileCache(file);
				if (!fileCache)
					return false;

				const headersWithBulletPoints = getHeadersWithBulletPointsInFile(fileCache, editor);
				if (headersWithBulletPoints === undefined)
					return false;
				if (checking)
					return true;

				const frontmatter = fileCache.frontmatter;
				const powerPointFilePath = frontmatter?.['powerPoint-file'];
				if (!powerPointFilePath || typeof powerPointFilePath !== 'string') {
					new Notice('You have to specify the path for your PowerPoint file in the frontmatter with `powerPoint-file: <path>.pptx`');
					return;
				}
				if (!powerPointFilePath.endsWith('.pptx')) {
					new Notice('The specified PowerPoint file path must end with `.pptx`');
					return;
				}
				async function writeToPowerPoint(headersWithBulletPoints: HeaderBulletPoint[]) {
					const powerPointFile = await PowerPointFile.loadAsync(powerPointFilePath);
					const modifiedPath = powerPointFilePath.replace('.pptx', `_modified.pptx`);
					try {
						const clonedFile = await powerPointFile.copyAsync(modifiedPath);

						for (let i = 0; i < headersWithBulletPoints.length; i++) {
							const header = headersWithBulletPoints[i];
							const bulletPoints = header.bulletPoints;

							await clonedFile.writeNotesFileAsync(i + 1, bulletPoints);
						}
						await clonedFile.saveAsync();
						new Notice(`Wrote notes to ${modifiedPath}`);
					} catch (error) {
						if (error instanceof Error) {
							if ((error as any).code === "EBUSY") {
								new Notice('Failed to save PowerPoint file. It is locked, so you probably need to close it.');
								return;
							}
						}
						new Notice('Failed to save PowerPoint file. Error: ' + error);
					}
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

type HeaderBulletPoint = {
	header: string;
	bulletPoints: string[];
};

function getHeadersWithBulletPointsInFile(fileCache: CachedMetadata, editor: Editor): HeaderBulletPoint[] | undefined {
	const headers = fileCache.headings;
	if (!headers || headers.length === 0) return [];
	const allListItems = fileCache.listItems;
	if (!allListItems || allListItems.length === 0) return [];

	return headers.map((header, index) => {
		const startPosition = header.position.start.line;
		const endPosition = headers.at(index + 1)?.position.start.line;
		return {
			header: header.heading,
			bulletPoints: allListItems.filter(item => isBetweenLines(item, startPosition, endPosition))
				.map(x => getText(x, editor))
		};
	});
}

function isBetweenLines(item: ListItemCache, startLine: number, endLine: number | undefined): boolean {
	const itemLine = item.position.start.line;
	return itemLine > startLine && itemLine < (endLine ?? Infinity);
}


function getText(item: ListItemCache, editor: Editor) {
	const startPosition: EditorPosition = toEditorLineStartPosition(item.position.start);
	const endPosition: EditorPosition = toEditorPosition(item.position.end);
	return editor.getRange(startPosition, endPosition);
}