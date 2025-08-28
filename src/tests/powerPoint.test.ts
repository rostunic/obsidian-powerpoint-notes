import * as path from 'path/win32';
import * as fs from 'fs';
import * as JSZip from 'jszip';
import { DOMParser } from 'xmldom';
import { writeNotesFile } from './generatePowerpointFile';


async function getAllNotesFromSlide(zip: JSZip, slideNumber: number): Promise<string[]> {
  // Directly read the notes XML for the slide
  const notesFile = zip.file(`ppt/notesSlides/notesSlide${slideNumber}.xml`);
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

describe('loadPowerPointFile', () => {
  const fullFilePath = path.join(__dirname, 'test-template.pptx');

  it('should extract all notes from the first slide', async () => {
    // Read pptx as zip
    const data = fs.readFileSync(fullFilePath);
    const zip = await JSZip.loadAsync(new Uint8Array(data));
    const notes = await getAllNotesFromSlide(zip, 1);
    expect(notes).toEqual([
      'Test note',
      'Test note 2',
    ]);
  });
});

describe('loadPowerPointFile', () => {

  it('should write notes to the 3rd slide', async () => {
    // Read pptx as zip
    const data = fs.readFileSync(path.join(__dirname, 'test-template.pptx'));
    const zip = await JSZip.loadAsync(new Uint8Array(data));
    const notes = await getAllNotesFromSlide(zip, 3);
    const newNotes = notes.map(note => note + ' (edited)');
    await writeNotesFile(zip, 3, newNotes);
    //write file
    fs.writeFileSync(
      path.join(__dirname, 'test-output.pptx'),
      new Uint8Array(await zip.generateAsync({ type: 'nodebuffer' }))
    );
    
    const data2 = fs.readFileSync(path.join(__dirname, 'test-output.pptx'));
    const zip2 = await JSZip.loadAsync(new Uint8Array(data2));
    const updatedNotes = await getAllNotesFromSlide(zip2, 3);
    expect(updatedNotes).toEqual([
      ' d (edited)'
    ]);
  });
});
