import * as fs from 'fs';
import * as JSZip from 'jszip';
import { writeNotesFile } from './generatePowerpointFile';
import { DOMParser } from 'xmldom';


export class PowerPointFile {
  private constructor(public filePath: string, private zipFile: JSZip) { }

  public static async loadAsync(filePath: string): Promise<PowerPointFile> {
    const data = fs.readFileSync(filePath);
    const zip = await JSZip.loadAsync(new Uint8Array(data));
    return new PowerPointFile(filePath, zip);
  }
  public async copyAsync(newFilePath: string): Promise<PowerPointFile> {
    // copy file and reload
    fs.writeFileSync(newFilePath, new Uint8Array(await this.zipFile.generateAsync({ type: 'nodebuffer' })));
    return await PowerPointFile.loadAsync(newFilePath);
  }

  public async getAllNotesFromSlide(slideNumber: number): Promise<string[]> {
    // Directly read the notes XML for the slide
    const notesFile = this.zipFile.file(`ppt/notesSlides/notesSlide${slideNumber}.xml`);
    if (!notesFile) return [];
    const notesXml = await notesFile.async('string');
    const notesDoc = new DOMParser().parseFromString(notesXml, 'application/xml');
    // Find all <p:sp> with <p:ph type="body">
    const spList = Array.from(notesDoc.getElementsByTagName('p:sp'));
    const noteParas: Element[] = [];
    for (const sp of spList) {
      const ph = sp.getElementsByTagName('p:ph')[0];
      if (ph && ph.getAttribute('type') === 'body') {
        const txBody = sp.getElementsByTagName('p:txBody')[0];
        if (txBody) {
          noteParas.push(...Array.from(txBody.getElementsByTagName('a:p')));
        }
      }
    }
    // For each <a:p>, join all <a:t> children
    let notes = noteParas.map(p =>
      Array.from(p.getElementsByTagName('a:t'))
        .map(t => t.textContent)
        .join('')
    ).filter(line => line.trim() !== '');
    // Filter out lines that are only numbers
    notes = notes.filter(line => !/^\d+$/.test(line.trim()));
    return notes;
  }


  public async writeNotesFileAsync(slideNumber: number, notes: string[]): Promise<void> {
    await writeNotesFile(this.zipFile, slideNumber, notes);
  }
  public async saveAsync(): Promise<void> {
    fs.writeFileSync(
      this.filePath,
      new Uint8Array(await this.zipFile.generateAsync({ type: 'nodebuffer' }))
    );
  }
}
