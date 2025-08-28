import Automizer from 'pptx-automizer/dist';


export function loadPowerPointFile(filePath: string): Automizer {
	return new Automizer({ rootTemplate: filePath });
}
