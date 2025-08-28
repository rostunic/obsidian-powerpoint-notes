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
			}
		});

	}

	onunload() {

	}
}