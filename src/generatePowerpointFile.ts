import * as JSZip from "jszip";

export function generateNotesFile(notes: string[], slideNumber: number): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name="" />
                <p:cNvGrpSpPr />
                <p:nvPr />
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0" />
                    <a:ext cx="0" cy="0" />
                    <a:chOff x="0" y="0" />
                    <a:chExt cx="0" cy="0" />
                </a:xfrm>
            </p:grpSpPr>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="2" name="Folienbildplatzhalter 1" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" noRot="1" noChangeAspect="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="sldImg" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="3" name="Notizenplatzhalter 2" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="body" idx="1" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    ${notes.map(note => generateNote(note)).join('\n')}
                </p:txBody>
            </p:sp>
            <p:sp>
                <p:nvSpPr>
                    <p:cNvPr id="4" name="Foliennummernplatzhalter 3" />
                    <p:cNvSpPr>
                        <a:spLocks noGrp="1" />
                    </p:cNvSpPr>
                    <p:nvPr>
                        <p:ph type="sldNum" sz="quarter" idx="5" />
                    </p:nvPr>
                </p:nvSpPr>
                <p:spPr />
                <p:txBody>
                    <a:bodyPr />
                    <a:lstStyle />
                    <a:p>
                        <a:fld id="{93E155B9-23F8-47C7-B435-FA915C65A956}" type="slidenum">
                            <a:rPr lang="de-DE" smtClean="0" />
                            <a:t>${slideNumber}</a:t>
                        </a:fld>
                        <a:endParaRPr lang="de-DE" />
                    </a:p>
                </p:txBody>
            </p:sp>
        </p:spTree>
        <p:extLst>
            <p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">
                <p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"
                    val="1153082976" />
            </p:ext>
        </p:extLst>
    </p:cSld>
    <p:clrMapOvr>
        <a:masterClrMapping />
    </p:clrMapOvr>
</p:notes>
    `;
}

function getIndentAndContent(note: string): { indent: number; content: string } {
    const indentAndContent = note.match(/^(\s*)- (.*)$/);
    if (!indentAndContent) {
        return { indent: 0, content: note };
    }
    const indent = Math.floor(indentAndContent[1].length / 4);
    const content = indentAndContent[2];
    return { indent, content };
}

export function generateNote(note: string): string {
    const { indent, content } = getIndentAndContent(note) || { indent: 0, content: note };
    const calculatedMargin = 11450 + (150000 * (indent + 1));
    return `<a:p>
                        <a:pPr marL="${calculatedMargin}" indent="-171450">
                            <a:buFontTx />
                            <a:buChar char="-" />
                        </a:pPr>
                        <a:r>
                            <a:rPr lang="de-DE" dirty="0" />
                            <a:t>${content}</a:t>
                        </a:r>
                        <a:endParaRPr lang="de-DE" dirty="0" />
                    </a:p>`;
}

export function writeNotesFile(zip: JSZip, slideNumber: number, notes: string[]): Promise<void> {
    const content = generateNotesFile(notes, slideNumber);
    zip.file(`ppt/notesSlides/notesSlide${slideNumber}.xml`, content);
    return Promise.resolve();
}