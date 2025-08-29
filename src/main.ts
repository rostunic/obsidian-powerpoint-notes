import { BlockCache, CachedMetadata, Editor, LinkCache, MarkdownFileInfo, Plugin, TFile, parseLinktext } from 'obsidian';

type Context = {

	editor: Editor,
	fileCache: CachedMetadata,
	currentLink: LinkCache,
	file: TFile,

}

export default class AliasPickerPlugin extends Plugin {

	async onload() {

		this.addCommand({
			id: 'powerpoint-notes:write-notes-to-slides',
			name: 'Write notes to slides',
			editorCheckCallback: (checking: boolean, editor: Editor, markdownFileInfo: MarkdownFileInfo) => {
				// get frontmatter properties of current file:
				const file = markdownFileInfo.file;
				if (!file)
					return;
				const fileCache = this.app.metadataCache.getFileCache(markdownFileInfo.file);
				const frontmatter = fileCache?.frontmatter;
				if (!frontmatter)
					return;
				const powerPointFilePath = frontmatter['powerPoint-file'];
				if (!powerPointFilePath || typeof powerPointFilePath !== 'string' || !powerPointFilePath.endsWith('.pptx'))
					return;

				const getAllHeadersWithBulletPointsInFile = (file: TFile): {header: string, bulletPoints: string[]}[] => {
					const fileCache = this.app.metadataCache.getFileCache(file);
					if (!fileCache) return [];
					const headers = fileCache.headings;
					if (!headers || headers.length === 0) return [];
					const blocks = fileCache.blocks;
					if (!blocks) return [];
					const result: {header: string, bulletPoints: string[]}[] = [];
					for (const header of headers) {
						const nextHeader = headers.find(h => h.position.start.line > header.position.start.line);
						const fileContent = this.app.vault.read(file);
						const bulletPointsPromise = fileContent.then(content => {
							const lines = content.split('\n');
							return Object.values(blocks)
								.filter(block => block.position.start.line > header.position.start.line && block.position.start.line < (nextHeader?.position.start.line ?? Infinity))
								.map(block => {
									const start = block.position.start.line;
									const end = block.position.end.line;
									return lines.slice(start, end + 1).join('\n');
								});
						});
						// Note: bulletPointsPromise is a Promise<string[]>
						// You may need to handle this asynchronously where you use it
						result.push({ header: header.heading, bulletPoints: [] });
					}
					return result;
				};
			}
		});

	}

	onunload() {

	}
}