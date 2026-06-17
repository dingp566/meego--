import React, { useEffect, useMemo, useState } from 'react';
import { AttributeType, FieldType } from '@lark-project/js-sdk';
import './index.css';
import { buildVersion, diffParagraphs, extractDocxText } from './deliverableCore';

type DeliverableVersion = {
  id: string;
  versionName: string;
  fileName: string;
  fileSize: number;
  sourceType?: 'docx' | 'lark_doc';
  sourceUrl?: string;
  extractedText: string;
  changeNote: string;
  author: string;
  createdAt: string;
};

type PendingUpload = {
  sourceType: 'docx';
  file: File;
  extractedText: string;
  versionName: string;
};

type ParsedCloudDocument = {
  title: string;
  content: string;
};

type DiffBlock = {
  type: 'added' | 'removed' | 'modified' | 'unchanged';
  oldText?: string;
  newText?: string;
};

type DiffResult = {
  summary: Record<'added' | 'removed' | 'modified' | 'unchanged', number>;
  blocks: DiffBlock[];
};

const FALLBACK_WORK_ITEM_ID = 'local-preview-work-item';
const FALLBACK_WORK_ITEM_TITLE = '交付物版本管理';

const App: React.FC = () => {
  const [workItemId, setWorkItemId] = useState(FALLBACK_WORK_ITEM_ID);
  const [workItemTitle, setWorkItemTitle] = useState(FALLBACK_WORK_ITEM_TITLE);
  const [runtimeContext, setRuntimeContext] = useState<Record<string, any> | undefined>();
  const [versions, setVersions] = useState<DeliverableVersion[]>([]);
  const [baselineVersionId, setBaselineVersionId] = useState('');
  const [pendingUpload, setPendingUpload] = useState<PendingUpload | null>(null);
  const [sourceMode, setSourceMode] = useState<'docx' | 'lark_doc'>('docx');
  const [docLinkUrl, setDocLinkUrl] = useState('');
  const [docLinkPastedTitle, setDocLinkPastedTitle] = useState('');
  const [error, setError] = useState('');
  const [isUploading, setIsUploading] = useState(false);
  const [isParsingDocLink, setIsParsingDocLink] = useState(false);
  const [previewVersionId, setPreviewVersionId] = useState('');
  const [fromVersionId, setFromVersionId] = useState('');
  const [toVersionId, setToVersionId] = useState('');
  const [viewMode, setViewMode] = useState<'summary' | 'compare'>('compare');

  useEffect(() => {
    let mounted = true;
    let unwatchName: (() => void) | undefined;

    async function loadContext() {
      try {
        const context = await window.JSSDK?.tab?.getContext?.();
        console.log('[getContext]', context);
        const runtimeContext = context as Record<string, any> | undefined;
        const id =
          runtimeContext?.workItemId ||
          runtimeContext?.work_item_id ||
          runtimeContext?.workObjectId ||
          FALLBACK_WORK_ITEM_ID;

        if (mounted) {
          setRuntimeContext(runtimeContext);
          setWorkItemId(String(id));
          const title = await loadWorkItemTitle(runtimeContext);

          if (mounted && title) {
            setWorkItemTitle(title);
          }

          unwatchName = await watchWorkItemName((title) => {
            if (mounted && title) {
              setWorkItemTitle(title);
            }
          });
        }
      } catch {
        setRuntimeContext(undefined);
        setWorkItemId(FALLBACK_WORK_ITEM_ID);
        setWorkItemTitle(FALLBACK_WORK_ITEM_TITLE);
      }
    }

    loadContext();
    return () => {
      mounted = false;
      unwatchName?.();
    };
  }, []);

  useEffect(() => {
    const loadedVersions = loadVersions(workItemId);
    setVersions(loadedVersions);
    setBaselineVersionId(loadBaselineVersionId(workItemId) || loadedVersions[0]?.id || '');
  }, [workItemId]);

  useEffect(() => {
    const baseline = versions.find((item) => item.id === baselineVersionId) || versions[0];
    const latest = versions[0];
    const previous = versions.find((item) => item.id !== baseline?.id);

    if (versions.length >= 2) {
      if (latest?.id === baseline?.id) {
        setFromVersionId(previous?.id || '');
        setToVersionId(baseline?.id || '');
      } else {
        setFromVersionId(baseline?.id || '');
        setToVersionId(latest?.id || '');
      }
    } else {
      setFromVersionId('');
      setToVersionId('');
    }
  }, [baselineVersionId, versions]);

  const currentVersion = versions[0];
  const baselineVersion = versions.find((item) => item.id === baselineVersionId) || currentVersion;
  const previewVersion = versions.find((item) => item.id === previewVersionId);
  const pinnedVersion = baselineVersion;
  const otherVersions = versions.filter((item) => item.id !== pinnedVersion?.id);
  const nextVersionName = `V${versions.length + 1}`;
  const fromVersion = versions.find((item) => item.id === fromVersionId);
  const toVersion = versions.find((item) => item.id === toVersionId);

  const diffResult: DiffResult | null = useMemo(() => {
    if (!fromVersion || !toVersion || fromVersion.id === toVersion.id) {
      return null;
    }

    return diffParagraphs(fromVersion.extractedText, toVersion.extractedText) as DiffResult;
  }, [fromVersion, toVersion]);

  function renderVersionCard(version: DeliverableVersion, extraClass = '') {
    const isBaseline = version.id === baselineVersion?.id;
    const isCurrent = version.id === currentVersion?.id;
    const sourceType = version.sourceType || 'docx';

    return (
      <article
        key={version.id}
        className={['version-card', isBaseline ? 'baseline' : '', extraClass].filter(Boolean).join(' ')}
      >
        <button className="version-main" onClick={() => setPreviewVersionId(version.id)}>
          <SourceIcon sourceType={sourceType} />
          <span className="version-file-meta">
            <strong>{version.fileName}</strong>
            <small>
              {version.versionName} · {isBaseline ? '基线版本 · ' : ''}
              {isCurrent ? '当前版本 · ' : ''}
              {sourceLabel(sourceType)} · {formatVersionSize(version)} · {formatDate(version.createdAt)}
            </small>
          </span>
        </button>
        <div className="version-actions">
          <button
            className="icon-action"
            title={sourceType === 'lark_doc' ? '打开云文档' : '下载原始文档'}
            aria-label={`${sourceType === 'lark_doc' ? '打开' : '下载'} ${version.versionName}`}
            onClick={() => handlePrimaryVersionAction(version)}
          >
            {sourceType === 'lark_doc' ? <OpenLinkIcon /> : <DownloadIcon />}
          </button>
          <button
            className="icon-action"
            disabled={isBaseline}
            title="打为基线"
            aria-label={`打基线 ${version.versionName}`}
            onClick={() => handleSetBaseline(version.id)}
          >
            <BaselineIcon />
          </button>
          <button
            className="icon-action danger"
            title="删除版本"
            aria-label={`删除 ${version.versionName}`}
            onClick={() => handleDeleteVersion(version.id)}
          >
            <TrashIcon />
          </button>
        </div>
      </article>
    );
  }

  async function handleFile(file: File | undefined) {
    setError('');

    if (!file) {
      return;
    }

    if (!file.name.toLowerCase().endsWith('.docx')) {
      setError('当前版本仅支持上传 DOCX 文件');
      return;
    }

    setIsUploading(true);

    try {
      const extractedText = await extractDocxText(await file.arrayBuffer());
      setPendingUpload({
        sourceType: 'docx',
        file,
        extractedText,
        versionName: nextVersionName,
      });
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : 'DOCX 解析失败');
    } finally {
      setIsUploading(false);
    }
  }

  function handleDrop(event: React.DragEvent<HTMLLabelElement>) {
    event.preventDefault();
    handleFile(event.dataTransfer.files[0]);
  }

  async function handlePrepareDocLink() {
    setError('');
    const url = docLinkUrl.trim();

    if (!isFeishuDocUrl(url)) {
      setError('请输入飞书云文档链接，例如 larkoffice.com/docx 或 larkoffice.com/wiki 链接');
      return;
    }

    setIsParsingDocLink(true);

    try {
      const [parsedDocument, fieldTitle] = await Promise.all([
        parseCloudDocumentLink(url),
        resolveCloudDocumentTitleFromWorkItem(url, runtimeContext),
      ]);

      createVersionFromCloudDocument(url, parsedDocument, fieldTitle || docLinkPastedTitle);
    } finally {
      setIsParsingDocLink(false);
    }
  }

  async function handleConfirmUpload() {
    if (!pendingUpload) {
      return;
    }

    const version = buildVersion({
      versions,
      fileName: pendingUpload.file.name,
      fileSize: pendingUpload.file.size,
      sourceType: 'docx',
      sourceUrl: '',
      extractedText: pendingUpload.extractedText,
      changeNote: '',
      author: '当前用户',
    }) as DeliverableVersion;
    const nextVersions = [version, ...versions];
    const nextBaselineId = baselineVersionId || version.id;

    await saveVersionFile(version.id, pendingUpload.file);
    saveVersions(workItemId, nextVersions);
    saveBaselineVersionId(workItemId, nextBaselineId);
    setVersions(nextVersions);
    setBaselineVersionId(nextBaselineId);
    setPendingUpload(null);
  }

  function createVersionFromCloudDocument(
    url: string,
    parsedDocument: ParsedCloudDocument | null,
    pastedTitle: string,
  ) {
    const version = buildVersion({
      versions,
      fileName: parsedDocument?.title || pastedTitle || '飞书云文档链接',
      fileSize: 0,
      sourceType: 'lark_doc',
      sourceUrl: url,
      extractedText: parsedDocument?.content || '',
      changeNote: '',
      author: '当前用户',
    }) as DeliverableVersion;
    const nextVersions = [version, ...versions];
    const nextBaselineId = baselineVersionId || version.id;

    saveVersions(workItemId, nextVersions);
    saveBaselineVersionId(workItemId, nextBaselineId);
    setVersions(nextVersions);
    setBaselineVersionId(nextBaselineId);
    setDocLinkUrl('');
    setDocLinkPastedTitle('');
  }

  function handleSetBaseline(versionId: string) {
    saveBaselineVersionId(workItemId, versionId);
    setBaselineVersionId(versionId);
  }

  async function handleDeleteVersion(versionId: string) {
    const nextVersions = versions.filter((version) => version.id !== versionId);
    const nextBaselineId = baselineVersionId === versionId ? nextVersions[0]?.id || '' : baselineVersionId;

    await deleteVersionFile(versionId);
    saveVersions(workItemId, nextVersions);
    saveBaselineVersionId(workItemId, nextBaselineId);
    setVersions(nextVersions);
    setBaselineVersionId(nextBaselineId);
    if (previewVersionId === versionId) {
      setPreviewVersionId('');
    }
  }

  async function handleDownloadVersion(version: DeliverableVersion) {
    const file = await loadVersionFile(version.id);

    if (!file) {
      setError('未找到原始 DOCX 文件，请重新上传该版本');
      return;
    }

    const url = URL.createObjectURL(file);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = version.fileName;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  function handlePrimaryVersionAction(version: DeliverableVersion) {
    if ((version.sourceType || 'docx') === 'lark_doc' && version.sourceUrl) {
      window.open(version.sourceUrl, '_blank', 'noopener,noreferrer');
      return;
    }

    handleDownloadVersion(version);
  }

  function resetPendingUpload() {
    setPendingUpload(null);
  }

  function handleDocLinkPaste(event: React.ClipboardEvent<HTMLInputElement>) {
    const html = event.clipboardData.getData('text/html');
    const plainText = event.clipboardData.getData('text/plain');
    const pastedLink = extractLinkFromClipboard(html, plainText);

    if (!pastedLink) {
      return;
    }

    event.preventDefault();
    setDocLinkUrl(pastedLink.url);
    setDocLinkPastedTitle(pastedLink.title);
  }

  return (
    <main className="deliverable-page">
      <section className="header-band">
        <div>
          <p className="eyebrow">工作项交付物</p>
          <h1>{workItemTitle}</h1>
        </div>
        <div className="status-panel">
          <span>基线版本</span>
          <strong>{baselineVersion?.versionName || '暂无'}</strong>
        </div>
      </section>

      <section className="page-stack">
          <section className="section-block">
            <div className="section-title">
              <h2>上传新版本</h2>
              <span>上传后将生成 {nextVersionName}</span>
            </div>

            {pendingUpload && (
              <div className="pending-bar">
                <div>
                  <strong>附件编辑中，确认后完成更新</strong>
                  <span>
                    {pendingUpload.versionName} · {pendingUpload.file.name} · {formatFileSize(pendingUpload.file.size)}
                  </span>
                </div>
                <div className="pending-actions">
                  <button onClick={resetPendingUpload}>取消</button>
                  <button className="primary-action" onClick={handleConfirmUpload}>
                    确定
                  </button>
                </div>
              </div>
            )}

            <div className="source-switch" aria-label="交付物来源">
              <button className={sourceMode === 'docx' ? 'active' : ''} onClick={() => setSourceMode('docx')}>
                DOCX 文件
              </button>
              <button className={sourceMode === 'lark_doc' ? 'active' : ''} onClick={() => setSourceMode('lark_doc')}>
                飞书云文档链接
              </button>
            </div>

            {sourceMode === 'docx' && (
              <label
                className="upload-zone"
                onDragOver={(event) => event.preventDefault()}
                onDrop={handleDrop}
              >
                <input
                  type="file"
                  accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                  onChange={(event) => handleFile(event.target.files?.[0])}
                />
                <span className="upload-icon">↑</span>
                <strong>{isUploading ? '正在解析 DOCX...' : '点击上传或拖拽 DOCX 文件到这里'}</strong>
                <small>支持 DOCX，系统会保存正文段落用于版本 Diff</small>
              </label>
            )}

            {sourceMode === 'lark_doc' && (
              <div className="doc-link-form">
                <label>
                  <span>云文档链接</span>
                  <input
                    value={docLinkUrl}
                    placeholder="https://bytedance.larkoffice.com/docx/..."
                    onChange={(event) => setDocLinkUrl(event.target.value)}
                    onPaste={handleDocLinkPaste}
                  />
                </label>
                {docLinkPastedTitle && <p className="doc-link-title-preview">识别到：{docLinkPastedTitle}</p>}
                <p className="doc-link-hint">
                  从飞书复制带标题的链接时会自动识别名称；后端解析接入后可自动读取云文档标题和正文。
                </p>
                <button className="primary-action doc-link-submit" disabled={isParsingDocLink} onClick={handlePrepareDocLink}>
                  {isParsingDocLink ? '解析中...' : '确认链接'}
                </button>
              </div>
            )}

            {error && <p className="error-text">{error}</p>}

            <div className="attachment-area">
              <div className="section-title compact">
                <h2>已上传附件</h2>
                <span>{versions.length} 个版本</span>
              </div>

              {versions.length === 0 && <EmptyState text="上传 DOCX 或添加云文档链接后可在这里管理历史版本" />}
              {versions.length > 0 && (
                <div className="version-rail">
                  {pinnedVersion && renderVersionCard(pinnedVersion, 'pinned')}
                  {otherVersions.length > 0 && (
                    <div className="version-grid">
                      {otherVersions.map((version) => renderVersionCard(version))}
                    </div>
                  )}
                </div>
              )}
            </div>
          </section>

          <section className="section-block compare-section">
            <div className="section-title">
              <h2>版本对比</h2>
              <span>默认以基线版本对比当前版本</span>
            </div>

            <div className="compare-toolbar">
              <select value={fromVersionId} onChange={(event) => setFromVersionId(event.target.value)}>
                {versions.map((version) => (
                  <option key={version.id} value={version.id}>
                    {version.versionName} · {version.fileName}
                  </option>
                ))}
              </select>
              <span className="arrow">→</span>
              <select value={toVersionId} onChange={(event) => setToVersionId(event.target.value)}>
                {versions.map((version) => (
                  <option key={version.id} value={version.id}>
                    {version.versionName} · {version.fileName}
                  </option>
                ))}
              </select>
              <div className="segment">
                <button className={viewMode === 'compare' ? 'active' : ''} onClick={() => setViewMode('compare')}>
                  对比
                </button>
                <button className={viewMode === 'summary' ? 'active' : ''} onClick={() => setViewMode('summary')}>
                  汇总
                </button>
              </div>
            </div>

            {!diffResult && <EmptyState text="至少上传两个版本后可查看 Diff" />}
            {diffResult && (
              <DiffWorkspace
                result={diffResult}
                viewMode={viewMode}
                fromName={fromVersion?.versionName || ''}
                toName={toVersion?.versionName || ''}
              />
            )}
          </section>
      </section>
      {previewVersion && (
        <DocumentPreviewModal
          version={previewVersion}
          onClose={() => setPreviewVersionId('')}
          onPrimaryAction={() => handlePrimaryVersionAction(previewVersion)}
        />
      )}
    </main>
  );
};

function DocumentPreviewModal({
  version,
  onClose,
  onPrimaryAction,
}: {
  version: DeliverableVersion;
  onClose: () => void;
  onPrimaryAction: () => void;
}) {
  const sourceType = version.sourceType || 'docx';
  const paragraphs = version.extractedText
    .split(/\n+/)
    .map((item) => item.trim())
    .filter(Boolean);

  return (
    <div className="preview-overlay" role="dialog" aria-modal="true" aria-label={`${version.fileName} 阅览`}>
      <section className="preview-modal">
        <header className="preview-head">
          <div className="preview-file">
            <SourceIcon sourceType={sourceType} />
            <div>
              <strong>{version.fileName}</strong>
              <small>
                {version.versionName} · {sourceLabel(sourceType)} · {formatVersionSize(version)} ·{' '}
                {formatDate(version.createdAt)}
              </small>
            </div>
          </div>
          <button className="preview-close" aria-label="关闭阅览" onClick={onClose}>
            ×
          </button>
        </header>
        <div className="preview-body">
          <h2>{version.fileName.replace(/\.docx$/i, '')}</h2>
          {paragraphs.length === 0 && <EmptyState text="该版本未保存可阅览正文" />}
          {paragraphs.map((paragraph, index) => (
            <p key={`${version.id}-${index}`}>{paragraph}</p>
          ))}
        </div>
        <footer className="preview-footer">
          <button className="preview-download" onClick={onPrimaryAction}>
            {sourceType === 'lark_doc' ? <OpenLinkIcon /> : <DownloadIcon />}
            {sourceType === 'lark_doc' ? '打开云文档' : '下载附件'}
          </button>
        </footer>
      </section>
    </div>
  );
}

function DiffWorkspace({
  result,
  viewMode,
  fromName,
  toName,
}: {
  result: DiffResult;
  viewMode: 'summary' | 'compare';
  fromName: string;
  toName: string;
}) {
  const changedItems = result.blocks
    .map((block, index) => ({ block, index }))
    .filter((item) => item.block.type !== 'unchanged');

  function jumpTo(index: number) {
    document.getElementById(`diff-change-${index}`)?.scrollIntoView({
      behavior: 'smooth',
      block: 'center',
    });
  }

  return (
    <div className="diff-workspace">
      <div className="diff-main">
        {viewMode === 'summary' && <SummaryView changedItems={changedItems} />}
        {viewMode === 'compare' && <CompareView blocks={result.blocks} />}
      </div>
      <ChangeNavigator
        result={result}
        changedItems={changedItems}
        fromName={fromName}
        toName={toName}
        onJump={jumpTo}
      />
    </div>
  );
}

function SummaryView({ changedItems }: { changedItems: Array<{ block: DiffBlock; index: number }> }) {
  if (changedItems.length === 0) {
    return <EmptyState text="两个版本正文内容一致" />;
  }

  return (
    <div className="summary-view">
      {changedItems.map(({ block, index }, order) => (
        <div id={`diff-change-${index}`} key={`${block.type}-${index}`} className={`diff-row ${block.type}`}>
          <b>
            {order + 1}. {labelOf(block.type)}
          </b>
          {block.oldText && <p className="old-text">{block.oldText}</p>}
          {block.newText && <p className="new-text">{block.newText}</p>}
        </div>
      ))}
    </div>
  );
}

function CompareView({ blocks }: { blocks: DiffBlock[] }) {
  return (
    <div className="compare-view">
      <div className="compare-head">
        <span>旧版本</span>
        <span>新版本</span>
      </div>
      {blocks.map((block, index) => (
        <div
          id={block.type === 'unchanged' ? undefined : `diff-change-${index}`}
          key={`${block.type}-${index}`}
          className={`compare-row ${block.type}`}
        >
          <p>{block.oldText || ''}</p>
          <p>{block.newText || ''}</p>
        </div>
      ))}
    </div>
  );
}

function ChangeNavigator({
  result,
  changedItems,
  fromName,
  toName,
  onJump,
}: {
  result: DiffResult;
  changedItems: Array<{ block: DiffBlock; index: number }>;
  fromName: string;
  toName: string;
  onJump: (index: number) => void;
}) {
  return (
    <aside className="change-navigator">
      <div>
        <p className="eyebrow">变更导航</p>
        <h2>
          {fromName} → {toName}
        </h2>
      </div>

      <div className="change-stats">
        <span>新增 {result.summary.added}</span>
        <span>修改 {result.summary.modified}</span>
        <span>删除 {result.summary.removed}</span>
      </div>

      {changedItems.length === 0 && <EmptyState text="无可跳转变更" />}
      <div className="jump-list">
        {changedItems.map(({ block, index }, order) => (
          <button key={`${block.type}-${index}`} onClick={() => onJump(index)}>
            <b>{order + 1}</b>
            <span>{labelOf(block.type)}</span>
          </button>
        ))}
      </div>
    </aside>
  );
}

function WordFileIcon() {
  return (
    <span className="word-file-icon" aria-hidden="true">
      W
    </span>
  );
}

function LarkDocIcon() {
  return (
    <span className="lark-doc-icon" aria-hidden="true">
      云
    </span>
  );
}

function SourceIcon({ sourceType }: { sourceType: 'docx' | 'lark_doc' }) {
  return sourceType === 'lark_doc' ? <LarkDocIcon /> : <WordFileIcon />;
}

function DownloadIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M12 3v11" />
      <path d="m7 10 5 5 5-5" />
      <path d="M5 20h14" />
    </svg>
  );
}

function OpenLinkIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M14 4h6v6" />
      <path d="M10 14 20 4" />
      <path d="M20 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V5a1 1 0 0 1 1-1h5" />
    </svg>
  );
}

function BaselineIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M12 3.5 14.7 9l6 .9-4.4 4.2 1 6-5.3-2.8-5.3 2.8 1-6-4.4-4.2 6-.9L12 3.5Z" />
    </svg>
  );
}

function TrashIcon() {
  return (
    <svg viewBox="0 0 24 24" aria-hidden="true">
      <path d="M4 7h16" />
      <path d="M10 11v6" />
      <path d="M14 11v6" />
      <path d="M6 7l1 13h10l1-13" />
      <path d="M9 7V4h6v3" />
    </svg>
  );
}

function EmptyState({ text }: { text: string }) {
  return <div className="empty-state">{text}</div>;
}

async function loadWorkItemTitle(context: Record<string, any> | undefined) {
  try {
    if (context?.spaceId && context?.workObjectId && context?.workItemId) {
      const workItem = await window.JSSDK?.WorkItem?.load?.({
        spaceId: context.spaceId,
        workObjectId: context.workObjectId,
        workItemId: Number(context.workItemId),
      });
      console.log('[WorkItem.load]', workItem);

      if (workItem?.name) {
        return workItem.name;
      }
    }
  } catch (reason) {
    console.error('[WorkItem.load]', reason);
  }

  return extractTitleFromContext(context);
}

function extractTitleFromContext(context: Record<string, any> | undefined): string {
  const candidate =
    context?.name ||
    context?.workItemName ||
    context?.work_item_name ||
    context?.attribute?.name?.attributeValue ||
    context?.activeWorkItem?.name ||
    context?.workItem?.name;

  return normalizeWorkItemName(candidate);
}

async function watchWorkItemName(onChange: (title: string) => void): Promise<(() => void) | undefined> {
  try {
    return await window.JSSDK?.tab?.onWorkItemFormValueChanged?.(
      {
        watchKeys: [{ key: AttributeType.name, type: FieldType.text }],
      },
      (changedValue) => {
        console.log('[onWorkItemFormValueChanged]', changedValue);
        const title = normalizeWorkItemName(changedValue?.name);

        if (title) {
          onChange(title);
        }
      },
    );
  } catch (reason) {
    console.error('[onWorkItemFormValueChanged]', reason);
    return undefined;
  }
}

function normalizeWorkItemName(value: any): string {
  const candidate = typeof value === 'string' ? value : value?.attributeValue || value?.value || value?.text;

  return typeof candidate === 'string' ? candidate.trim() : '';
}

function loadVersions(workItemId: string): DeliverableVersion[] {
  try {
    return JSON.parse(localStorage.getItem(storageKey(workItemId)) || '[]');
  } catch {
    return [];
  }
}

function saveVersions(workItemId: string, versions: DeliverableVersion[]) {
  localStorage.setItem(storageKey(workItemId), JSON.stringify(versions));
}

function storageKey(workItemId: string) {
  return `deliverable_versions_${workItemId}`;
}

function loadBaselineVersionId(workItemId: string) {
  return localStorage.getItem(baselineStorageKey(workItemId)) || '';
}

function saveBaselineVersionId(workItemId: string, versionId: string) {
  if (versionId) {
    localStorage.setItem(baselineStorageKey(workItemId), versionId);
  } else {
    localStorage.removeItem(baselineStorageKey(workItemId));
  }
}

function baselineStorageKey(workItemId: string) {
  return `deliverable_baseline_${workItemId}`;
}

async function saveVersionFile(versionId: string, file: File) {
  const db = await openDeliverableDb();
  await requestToPromise(db.transaction('files', 'readwrite').objectStore('files').put(file, versionId));
  db.close();
}

async function loadVersionFile(versionId: string): Promise<Blob | undefined> {
  const db = await openDeliverableDb();
  const file = await requestToPromise<Blob | undefined>(
    db.transaction('files', 'readonly').objectStore('files').get(versionId),
  );
  db.close();
  return file;
}

async function deleteVersionFile(versionId: string) {
  const db = await openDeliverableDb();
  await requestToPromise(db.transaction('files', 'readwrite').objectStore('files').delete(versionId));
  db.close();
}

function openDeliverableDb(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open('deliverable-docx-store', 1);

    request.onupgradeneeded = () => {
      request.result.createObjectStore('files');
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error('无法打开本地文件存储'));
  });
}

function requestToPromise<T = unknown>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error('本地文件存储操作失败'));
  });
}

function labelOf(type: DiffBlock['type']) {
  return {
    added: '新增',
    removed: '删除',
    modified: '修改',
    unchanged: '未变',
  }[type];
}

function formatFileSize(size: number) {
  if (size < 1024 * 1024) {
    return `${Math.max(1, Math.round(size / 1024))} KB`;
  }

  return `${(size / 1024 / 1024).toFixed(1)} MB`;
}

function sourceLabel(sourceType: 'docx' | 'lark_doc') {
  return sourceType === 'lark_doc' ? '飞书云文档' : 'DOCX';
}

function formatVersionSize(version: DeliverableVersion) {
  return (version.sourceType || 'docx') === 'lark_doc' ? '链接' : formatFileSize(version.fileSize);
}

function extractLinkFromClipboard(html: string, plainText: string) {
  const plainUrl = extractFirstFeishuUrl(plainText);

  if (!html) {
    return plainUrl ? { url: plainUrl, title: '' } : null;
  }

  const doc = new DOMParser().parseFromString(html, 'text/html');
  const anchors = Array.from(doc.querySelectorAll('a'));
  const anchor = anchors.find((item) => {
    const href = item.getAttribute('href') || '';
    return isFeishuDocUrl(href);
  });

  if (!anchor) {
    return plainUrl ? { url: plainUrl, title: '' } : null;
  }

  return {
    url: anchor.href,
    title: (anchor.textContent || '').trim(),
  };
}

function extractFirstFeishuUrl(text: string) {
  const matches = text.match(/https?:\/\/[^\s<>"']+/g) || [];
  return matches.find((item) => isFeishuDocUrl(item)) || '';
}

async function parseCloudDocumentLink(url: string): Promise<ParsedCloudDocument | null> {
  try {
    const response = await fetch('/api/proxy/lark-documents/parse', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ url }),
    });

    if (!response.ok) {
      return null;
    }

    const data = (await response.json()) as Partial<ParsedCloudDocument>;
    const title = typeof data.title === 'string' ? data.title.trim() : '';
    const content = typeof data.content === 'string' ? data.content.trim() : '';

    if (!title && !content) {
      return null;
    }

    return {
      title: title || '飞书云文档链接',
      content,
    };
  } catch {
    return null;
  }
}

async function resolveCloudDocumentTitleFromWorkItem(
  url: string,
  context: Record<string, any> | undefined,
): Promise<string> {
  try {
    if (!context?.spaceId || !context?.workObjectId || !context?.workItemId) {
      return '';
    }

    const workItem = await window.JSSDK?.WorkItem?.load?.({
      spaceId: context.spaceId,
      workObjectId: context.workObjectId,
      workItemId: Number(context.workItemId),
    });
    const title = findTitleByUrl(workItem, url);
    console.log('[URL field title lookup]', { url, title, hasWorkItem: Boolean(workItem) });

    return title;
  } catch (reason) {
    console.error('[URL field title lookup]', reason);
    return '';
  }
}

function findTitleByUrl(value: unknown, targetUrl: string, seen = new WeakSet<object>()): string {
  if (!value || typeof value !== 'object') {
    return '';
  }

  if (seen.has(value)) {
    return '';
  }

  seen.add(value);

  const record = value as Record<string, unknown>;
  const url = extractUrlValue(record);

  if (url && isSameDocumentUrl(url, targetUrl)) {
    return extractDisplayTitle(record, targetUrl);
  }

  for (const key of safeObjectKeys(record)) {
    const child = safeRead(record, key);

    if (Array.isArray(child)) {
      for (const item of child) {
        const title = findTitleByUrl(item, targetUrl, seen);

        if (title) {
          return title;
        }
      }
    } else {
      const title = findTitleByUrl(child, targetUrl, seen);

      if (title) {
        return title;
      }
    }
  }

  return '';
}

function extractUrlValue(record: Record<string, unknown>) {
  const directKeys = ['url', 'href', 'link', 'linkUrl', 'targetUrl', 'openUrl', 'value'];

  for (const key of directKeys) {
    const value = safeRead(record, key);

    if (typeof value === 'string' && isFeishuDocUrl(value)) {
      return value;
    }
  }

  return '';
}

function extractDisplayTitle(record: Record<string, unknown>, targetUrl: string) {
  const titleKeys = [
    'title',
    'name',
    'text',
    'displayName',
    'display_name',
    'label',
    'linkText',
    'link_text',
    'displayText',
    'display_text',
    'attributeValue',
  ];

  for (const key of titleKeys) {
    const value = safeRead(record, key);

    if (typeof value === 'string') {
      const title = value.trim();

      if (title && !isSameDocumentUrl(title, targetUrl) && !isFeishuDocUrl(title)) {
        return title;
      }
    }
  }

  return '';
}

function safeObjectKeys(value: Record<string, unknown>) {
  try {
    return Object.keys(value);
  } catch {
    return [];
  }
}

function safeRead(value: Record<string, unknown>, key: string) {
  try {
    return value[key];
  } catch {
    return undefined;
  }
}

function isSameDocumentUrl(left: string, right: string) {
  try {
    const leftUrl = new URL(left);
    const rightUrl = new URL(right);
    const normalize = (url: URL) => `${url.hostname.toLowerCase()}${url.pathname.replace(/\/$/, '')}`;

    return normalize(leftUrl) === normalize(rightUrl);
  } catch {
    return left.trim() === right.trim();
  }
}

function isFeishuDocUrl(value: string) {
  try {
    const url = new URL(value);
    const host = url.hostname.toLowerCase();
    const isFeishuHost =
      host === 'feishu.cn' ||
      host.endsWith('.feishu.cn') ||
      host === 'larkoffice.com' ||
      host.endsWith('.larkoffice.com');
    const isDocPath = /\/(docx|doc|wiki|mindnotes|sheets|base)\//i.test(url.pathname);

    return isFeishuHost && isDocPath;
  } catch {
    return false;
  }
}

function formatDate(value: string) {
  return new Intl.DateTimeFormat('zh-CN', {
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
  }).format(new Date(value));
}

export default App;
