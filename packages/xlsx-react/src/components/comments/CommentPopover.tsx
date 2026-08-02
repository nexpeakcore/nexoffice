/**
 * The popovers behind cell comments: a read-only view of a comment's author and
 * text, and an editor with a textarea plus Save/Delete. Pure chrome — the
 * caller owns the comment list and the wasm call.
 */

import { useEffect, useRef, useState } from 'react';
import type { CSSProperties, ReactNode } from 'react';
import { toolbarColors } from '../ui/ToolbarPrimitives';
import { useTranslation } from '../../i18n';
import { isSavableCommentText } from './commentState';

const popoverStyle: CSSProperties = {
  width: 224,
  padding: 8,
  border: `1px solid ${toolbarColors.border}`,
  borderRadius: 8,
  background: toolbarColors.surface,
  boxShadow: '0 4px 16px rgba(60, 64, 67, 0.24)',
  boxSizing: 'border-box',
  zIndex: 30,
};

const authorStyle: CSSProperties = {
  margin: '0 0 4px',
  font: '600 12px ui-sans-serif, system-ui, sans-serif',
  color: toolbarColors.text,
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap',
};

const textStyle: CSSProperties = {
  margin: 0,
  font: '400 13px ui-sans-serif, system-ui, sans-serif',
  color: toolbarColors.text,
  whiteSpace: 'pre-wrap',
  overflowWrap: 'break-word',
};

const footerButtonStyle: CSSProperties = {
  appearance: 'none',
  minHeight: 26,
  padding: '0 12px',
  border: `1px solid ${toolbarColors.border}`,
  borderRadius: 6,
  background: toolbarColors.surface,
  color: toolbarColors.text,
  font: '500 12px ui-sans-serif, system-ui, sans-serif',
  cursor: 'pointer',
};

function useDismiss(rootRef: React.RefObject<HTMLDivElement | null>, onClose: () => void) {
  useEffect(() => {
    const onPointerDown = (event: MouseEvent) => {
      const target = event.target;
      if (!(target instanceof Node)) return;
      if (rootRef.current?.contains(target)) return;
      if (target instanceof Element && target.closest('[data-xlsx-comment-trigger]')) return;
      onClose();
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') onClose();
    };
    document.addEventListener('mousedown', onPointerDown);
    document.addEventListener('keydown', onKeyDown);
    return () => {
      document.removeEventListener('mousedown', onPointerDown);
      document.removeEventListener('keydown', onKeyDown);
    };
  }, [rootRef, onClose]);
}

function PopoverShell({
  label,
  testId,
  style,
  onClose,
  children,
}: {
  label: string;
  testId: string;
  style?: CSSProperties;
  onClose: () => void;
  children: ReactNode;
}) {
  const rootRef = useRef<HTMLDivElement>(null);
  useDismiss(rootRef, onClose);
  return (
    <div
      ref={rootRef}
      data-testid={testId}
      role="dialog"
      aria-label={label}
      onMouseDown={(e) => e.stopPropagation()}
      onClick={(e) => e.stopPropagation()}
      onDoubleClick={(e) => e.stopPropagation()}
      style={{ ...popoverStyle, ...style }}
    >
      {children}
    </div>
  );
}

export interface CommentViewPopoverProps {
  cellName: string;
  author: string;
  text: string;
  style?: CSSProperties;
  onClose: () => void;
}

export function CommentViewPopover({
  cellName,
  author,
  text,
  style,
  onClose,
}: CommentViewPopoverProps) {
  const { t } = useTranslation();
  return (
    <PopoverShell
      label={t('comment.viewLabel', { cell: cellName })}
      testId="xlsx-comment-popover"
      style={style}
      onClose={onClose}
    >
      <p data-testid="xlsx-comment-author" style={authorStyle}>
        {author}
      </p>
      <p data-testid="xlsx-comment-text" style={textStyle}>
        {text}
      </p>
    </PopoverShell>
  );
}

export interface CommentEditorPopoverProps {
  cellName: string;
  author: string;
  initialText: string;
  /** whether the cell already has a comment the editor can delete. */
  canDelete: boolean;
  style?: CSSProperties;
  onSave: (text: string) => void;
  onDelete: () => void;
  onClose: () => void;
}

export function CommentEditorPopover({
  cellName,
  author,
  initialText,
  canDelete,
  style,
  onSave,
  onDelete,
  onClose,
}: CommentEditorPopoverProps) {
  const { t } = useTranslation();
  const [draft, setDraft] = useState(initialText);
  return (
    <PopoverShell
      label={t('comment.editLabel', { cell: cellName })}
      testId="xlsx-comment-editor"
      style={style}
      onClose={onClose}
    >
      <p style={authorStyle}>{author}</p>
      <textarea
        data-testid="xlsx-comment-input"
        value={draft}
        autoFocus
        rows={3}
        placeholder={t('comment.textPlaceholder')}
        aria-label={t('comment.textPlaceholder')}
        onChange={(e) => setDraft(e.target.value)}
        onKeyDown={(e) => e.stopPropagation()}
        style={{
          width: '100%',
          boxSizing: 'border-box',
          resize: 'vertical',
          border: `1px solid ${toolbarColors.border}`,
          borderRadius: 6,
          padding: 6,
          font: '400 13px ui-sans-serif, system-ui, sans-serif',
          color: toolbarColors.text,
          background: toolbarColors.surface,
          outlineColor: '#217346',
        }}
      />
      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 6, paddingTop: 6 }}>
        {canDelete && (
          <button
            type="button"
            data-testid="xlsx-comment-delete"
            onClick={() => onDelete()}
            style={{ ...footerButtonStyle, marginRight: 'auto', color: '#c5221f' }}
          >
            {t('comment.delete')}
          </button>
        )}
        <button
          type="button"
          data-testid="xlsx-comment-save"
          disabled={!isSavableCommentText(draft)}
          onClick={() => onSave(draft.trim())}
          style={{
            ...footerButtonStyle,
            background: '#217346',
            borderColor: '#217346',
            color: '#ffffff',
            opacity: isSavableCommentText(draft) ? 1 : 0.5,
          }}
        >
          {t('comment.save')}
        </button>
      </div>
    </PopoverShell>
  );
}
