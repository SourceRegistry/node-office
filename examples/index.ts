import OOXML from '../src';

const doc = OOXML.parseFromPath('res/test.docx');

console.log('Type:', doc.type);
console.log('Metadata:', doc.metadata);
console.log('Pages:', doc.pages.length);
console.log('First page content preview:', doc.pages[0].text);
