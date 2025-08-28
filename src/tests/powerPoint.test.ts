import * as path from 'path/win32';
import { PowerPointFile } from 'src/PowerPointFile';


const fullFilePath = path.join(__dirname, 'test-template.pptx');

describe('loadPowerPointFile', () => {

  it('should extract all notes from the first slide', async () => {
    // Read pptx as zip
    const powerPointFile = await PowerPointFile.loadAsync(fullFilePath);
    const notes = await powerPointFile.getAllNotesFromSlide(1);
    expect(notes).toEqual([
      'Test note',
      'Test note 2',
    ]);
  });
});

describe('loadPowerPointFile', () => {

  it('should write notes to the 3rd slide', async () => {
    // Read pptx as zip
    const powerPointFile = await PowerPointFile.loadAsync(fullFilePath);
    const notes = await powerPointFile.getAllNotesFromSlide(3);
    const newNotes = notes.map(note => note + ' (edited)');
    const saveFile = await powerPointFile.copyAsync(path.join(__dirname, 'test-output.pptx'));
    await saveFile.writeNotesFileAsync(3, newNotes);
    await saveFile.saveAsync();

    const powerPointFile2 = await PowerPointFile.loadAsync(path.join(__dirname, 'test-output.pptx'));
    const updatedNotes = await powerPointFile2.getAllNotesFromSlide(3);
    expect(updatedNotes).toEqual([
      ' d (edited)'
    ]);
  });
});

