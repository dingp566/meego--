const WORD_DOCUMENT_PATH = 'word/document.xml';

export function getNextVersionName(versions) {
  const maxVersion = versions.reduce((max, version) => {
    const value = Number(version.versionName.replace(/^V/i, ''));
    return Number.isFinite(value) ? Math.max(max, value) : max;
  }, 0);

  return `V${maxVersion + 1}`;
}

export function normalizeParagraphs(text) {
  return text
    .replace(/\r/g, '')
    .split('\n')
    .map((line) => line.replace(/\s+/g, ' ').trim())
    .filter(Boolean);
}

export function buildVersion({ versions, fileName, fileSize, extractedText, changeNote, author }) {
  return {
    id: `version_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    versionName: getNextVersionName(versions),
    fileName,
    fileSize,
    extractedText,
    changeNote: changeNote.trim(),
    author,
    createdAt: new Date().toISOString(),
  };
}

export function diffParagraphs(oldText, newText) {
  const oldParagraphs = normalizeParagraphs(oldText);
  const newParagraphs = normalizeParagraphs(newText);
  const rows = longestCommonSubsequence(oldParagraphs, newParagraphs);
  const blocks = [];
  let oldIndex = 0;
  let newIndex = 0;

  for (const row of rows) {
    while (oldIndex < row.oldIndex || newIndex < row.newIndex) {
      const removed = oldIndex < row.oldIndex ? oldParagraphs[oldIndex] : undefined;
      const added = newIndex < row.newIndex ? newParagraphs[newIndex] : undefined;

      if (removed && added) {
        blocks.push({
          type: 'modified',
          oldText: removed,
          newText: added,
        });
        oldIndex += 1;
        newIndex += 1;
      } else if (removed) {
        blocks.push({
          type: 'removed',
          oldText: removed,
        });
        oldIndex += 1;
      } else if (added) {
        blocks.push({
          type: 'added',
          newText: added,
        });
        newIndex += 1;
      }
    }

    blocks.push({
      type: 'unchanged',
      oldText: row.value,
      newText: row.value,
    });
    oldIndex = row.oldIndex + 1;
    newIndex = row.newIndex + 1;
  }

  while (oldIndex < oldParagraphs.length || newIndex < newParagraphs.length) {
    const removed = oldParagraphs[oldIndex];
    const added = newParagraphs[newIndex];

    if (removed && added) {
      blocks.push({
        type: 'modified',
        oldText: removed,
        newText: added,
      });
      oldIndex += 1;
      newIndex += 1;
    } else if (removed) {
      blocks.push({
        type: 'removed',
        oldText: removed,
      });
      oldIndex += 1;
    } else if (added) {
      blocks.push({
        type: 'added',
        newText: added,
      });
      newIndex += 1;
    }
  }

  const summary = blocks.reduce(
    (result, block) => {
      result[block.type] += 1;
      return result;
    },
    { added: 0, removed: 0, modified: 0, unchanged: 0 },
  );

  return { summary, blocks };
}

export async function extractDocxText(buffer, inflateRaw) {
  const xml = await readZipEntry(buffer, WORD_DOCUMENT_PATH, inflateRaw || inflateRawWithBrowser);
  return parseDocumentXml(xml);
}

export function parseDocumentXml(xml) {
  const paragraphs = [];
  const paragraphMatches = xml.match(/<w:p[\s\S]*?<\/w:p>/g) || [];

  for (const paragraphXml of paragraphMatches) {
    const textRuns = paragraphXml.match(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g) || [];
    const text = textRuns
      .map((run) => decodeXml(run.replace(/^<w:t(?:\s[^>]*)?>/, '').replace(/<\/w:t>$/, '')))
      .join('');

    if (text.trim()) {
      paragraphs.push(text.trim());
    }
  }

  return paragraphs.join('\n');
}

function longestCommonSubsequence(left, right) {
  const matrix = Array.from({ length: left.length + 1 }, () =>
    Array.from({ length: right.length + 1 }, () => 0),
  );

  for (let i = left.length - 1; i >= 0; i -= 1) {
    for (let j = right.length - 1; j >= 0; j -= 1) {
      matrix[i][j] = left[i] === right[j]
        ? matrix[i + 1][j + 1] + 1
        : Math.max(matrix[i + 1][j], matrix[i][j + 1]);
    }
  }

  const rows = [];
  let i = 0;
  let j = 0;

  while (i < left.length && j < right.length) {
    if (left[i] === right[j]) {
      rows.push({ oldIndex: i, newIndex: j, value: left[i] });
      i += 1;
      j += 1;
    } else if (matrix[i + 1][j] >= matrix[i][j + 1]) {
      i += 1;
    } else {
      j += 1;
    }
  }

  return rows;
}

async function readZipEntry(buffer, targetName, inflateRaw) {
  const bytes = new Uint8Array(buffer);
  let offset = 0;

  while (offset < bytes.length - 30) {
    if (readUint32(bytes, offset) !== 0x04034b50) {
      offset += 1;
      continue;
    }

    const compression = readUint16(bytes, offset + 8);
    const compressedSize = readUint32(bytes, offset + 18);
    const uncompressedSize = readUint32(bytes, offset + 22);
    const nameLength = readUint16(bytes, offset + 26);
    const extraLength = readUint16(bytes, offset + 28);
    const nameStart = offset + 30;
    const dataStart = nameStart + nameLength + extraLength;
    const fileName = new TextDecoder().decode(bytes.slice(nameStart, nameStart + nameLength));
    const data = bytes.slice(dataStart, dataStart + compressedSize);

    if (fileName === targetName) {
      if (compression === 0) {
        return new TextDecoder().decode(data);
      }

      if (compression === 8) {
        const inflated = await inflateRaw(data, uncompressedSize);
        return new TextDecoder().decode(inflated);
      }

      throw new Error('DOCX 使用了暂不支持的压缩方式');
    }

    offset = dataStart + compressedSize;
  }

  throw new Error('未在 DOCX 中找到正文内容');
}

async function inflateRawWithBrowser(data) {
  if (typeof DecompressionStream === 'undefined') {
    throw new Error('当前浏览器不支持 DOCX 解压能力，请使用最新版 Chrome');
  }

  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

function readUint16(bytes, offset) {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readUint32(bytes, offset) {
  return (
    bytes[offset] |
    (bytes[offset + 1] << 8) |
    (bytes[offset + 2] << 16) |
    (bytes[offset + 3] << 24)
  ) >>> 0;
}

function decodeXml(value) {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}
