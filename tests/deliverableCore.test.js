import assert from 'node:assert/strict';
import { deflateRawSync, inflateRawSync } from 'node:zlib';
import {
  buildVersion,
  diffParagraphs,
  extractDocxText,
  getNextVersionName,
  parseDocumentXml,
} from '../src/features/dashboard_web_deliverables/deliverableCore.js';

test('getNextVersionName increments from existing versions', () => {
  assert.equal(getNextVersionName([]), 'V1');
  assert.equal(getNextVersionName([{ versionName: 'V1' }, { versionName: 'V2' }]), 'V3');
});

test('buildVersion creates the next named version with metadata', () => {
  const version = buildVersion({
    versions: [{ versionName: 'V1' }],
    fileName: '可行性分析报告.docx',
    fileSize: 2048,
    extractedText: '第一段',
    changeNote: ' 补充风险 ',
    author: 'tester',
  });

  assert.equal(version.versionName, 'V2');
  assert.equal(version.changeNote, '补充风险');
  assert.equal(version.fileName, '可行性分析报告.docx');
});

test('buildVersion stores source metadata for cloud document links', () => {
  const version = buildVersion({
    versions: [],
    fileName: '方案评审云文档',
    fileSize: 0,
    sourceType: 'lark_doc',
    sourceUrl: 'https://bytedance.larkoffice.com/docx/example',
    extractedText: '云文档正文',
    changeNote: '',
    author: 'tester',
  });

  assert.equal(version.versionName, 'V1');
  assert.equal(version.sourceType, 'lark_doc');
  assert.equal(version.sourceUrl, 'https://bytedance.larkoffice.com/docx/example');
});

test('parseDocumentXml extracts DOCX paragraphs from document.xml', () => {
  const xml = [
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>',
    '<w:p><w:r><w:t>项目背景</w:t></w:r></w:p>',
    '<w:p><w:r><w:t>预算 &amp; 风险</w:t></w:r></w:p>',
    '</w:body></w:document>',
  ].join('');

  assert.equal(parseDocumentXml(xml), '项目背景\n预算 & 风险');
});

test('extractDocxText reads compressed word/document.xml from a DOCX-like zip', async () => {
  const xml = '<w:document><w:body><w:p><w:r><w:t>上线计划</w:t></w:r></w:p></w:body></w:document>';
  const buffer = createZipWithDocumentXml(xml);
  const text = await extractDocxText(buffer, async (data) => inflateRawSync(Buffer.from(data)));

  assert.equal(text, '上线计划');
});

test('diffParagraphs aggregates added, removed and modified paragraphs', () => {
  const result = diffParagraphs(
    ['项目背景', '六月上线', '预算 10 万', '旧风险'].join('\n'),
    ['项目背景', '七月上线', '预算 12 万', '新增验收标准'].join('\n'),
  );

  assert.equal(result.summary.modified, 3);
  assert.equal(result.summary.added, 0);
  assert.equal(result.summary.removed, 0);
  assert.deepEqual(
    result.blocks.filter((block) => block.type !== 'unchanged').map((block) => block.type),
    ['modified', 'modified', 'modified'],
  );
});

function test(name, fn) {
  Promise.resolve()
    .then(fn)
    .then(() => {
      console.log(`ok - ${name}`);
    })
    .catch((error) => {
      console.error(`not ok - ${name}`);
      console.error(error);
      process.exitCode = 1;
    });
}

function createZipWithDocumentXml(xml) {
  const name = Buffer.from('word/document.xml');
  const data = deflateRawSync(Buffer.from(xml));
  const localHeader = Buffer.alloc(30);

  localHeader.writeUInt32LE(0x04034b50, 0);
  localHeader.writeUInt16LE(20, 4);
  localHeader.writeUInt16LE(0, 6);
  localHeader.writeUInt16LE(8, 8);
  localHeader.writeUInt32LE(0, 10);
  localHeader.writeUInt32LE(0, 14);
  localHeader.writeUInt32LE(data.length, 18);
  localHeader.writeUInt32LE(Buffer.byteLength(xml), 22);
  localHeader.writeUInt16LE(name.length, 26);
  localHeader.writeUInt16LE(0, 28);

  const zip = Buffer.concat([localHeader, name, data]);
  return zip.buffer.slice(zip.byteOffset, zip.byteOffset + zip.byteLength);
}
