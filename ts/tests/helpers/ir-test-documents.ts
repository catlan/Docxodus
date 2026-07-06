import { readIrDocument, type IrDocument } from '../../src/index.js';

const enc = new TextEncoder();

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const documentXml = (body: string): string =>
  `<w:document xmlns:w="${W}"><w:body>${body}</w:body></w:document>`;

const escapeXml = (s: string): string =>
  s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');

export const p = (text: string): string =>
  `<w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;

export function parts(body: string): Map<string, Uint8Array> {
  return new Map([['word/document.xml', enc.encode(documentXml(body))]]);
}

export function doc(...paragraphTexts: string[]): IrDocument {
  return readIrDocument(parts(paragraphTexts.map(p).join('')));
}

export function docFromParagraphs(paragraphTexts: Iterable<string>): IrDocument {
  return readIrDocument(parts([...paragraphTexts].map(p).join('')));
}

export function fromBodyXml(body: string): IrDocument {
  return readIrDocument(parts(body));
}
