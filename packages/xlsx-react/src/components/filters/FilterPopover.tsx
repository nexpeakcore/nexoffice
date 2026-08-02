/**
 * The value-list popover behind a header filter button: distinct raw column
 * values with checkboxes, a select-all, a blanks entry, and Apply/Clear. Pure
 * chrome — the caller owns the spec and the wasm call.
 */

import { useEffect, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import { toolbarColors } from '../ui/ToolbarPrimitives';
import { useTranslation } from '../../i18n';
import { resolveCriteria } from './filterSpec';

export interface FilterPopoverProps {
  columnName: string;
  values: string[];
  hasBlanks: boolean;
  truncated: boolean;
  /** the column's applied criteria; `null` means unconstrained (all visible). */
  initialValues: string[] | null;
  initialShowBlanks: boolean;
  style?: CSSProperties;
  onApply: (values: string[] | null, showBlanks: boolean) => void;
  onClear: () => void;
  onClose: () => void;
}

const rowStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  minHeight: 26,
  padding: '2px 8px',
  borderRadius: 4,
  cursor: 'pointer',
  font: '400 13px ui-sans-serif, system-ui, sans-serif',
  color: toolbarColors.text,
};

const valueTextStyle: CSSProperties = {
  flex: 1,
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap',
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

export function FilterPopover({
  columnName,
  values,
  hasBlanks,
  truncated,
  initialValues,
  initialShowBlanks,
  style,
  onApply,
  onClear,
  onClose,
}: FilterPopoverProps) {
  const { t } = useTranslation();
  const rootRef = useRef<HTMLDivElement>(null);
  const [checked, setChecked] = useState<ReadonlySet<string>>(
    () => new Set(initialValues ?? values)
  );
  const [showBlanks, setShowBlanks] = useState(
    initialValues === null ? true : initialShowBlanks
  );

  useEffect(() => {
    const onPointerDown = (event: MouseEvent) => {
      const target = event.target;
      if (!(target instanceof Node)) return;
      if (rootRef.current?.contains(target)) return;
      if (target instanceof Element && target.closest('[data-xlsx-filter-trigger]')) return;
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
  }, [onClose]);

  const allChecked =
    values.every((value) => checked.has(value)) && (!hasBlanks || showBlanks);

  const toggleAll = () => {
    if (allChecked) {
      setChecked(new Set());
      setShowBlanks(false);
    } else {
      setChecked(new Set(values));
      setShowBlanks(true);
    }
  };

  const toggleValue = (value: string) => {
    setChecked((previous) => {
      const next = new Set(previous);
      if (next.has(value)) next.delete(value);
      else next.add(value);
      return next;
    });
  };

  return (
    <div
      ref={rootRef}
      data-testid="xlsx-filter-popover"
      role="dialog"
      aria-label={t('filter.columnButton', { column: columnName })}
      style={{
        width: 224,
        padding: 6,
        border: `1px solid ${toolbarColors.border}`,
        borderRadius: 8,
        background: toolbarColors.surface,
        boxShadow: '0 4px 16px rgba(60, 64, 67, 0.24)',
        boxSizing: 'border-box',
        zIndex: 30,
        ...style,
      }}
    >
      <label style={rowStyle}>
        <input
          type="checkbox"
          data-testid="xlsx-filter-select-all"
          checked={allChecked}
          onChange={toggleAll}
        />
        <span style={valueTextStyle}>{t('filter.selectAll')}</span>
      </label>
      <div
        style={{
          maxHeight: 240,
          overflowY: 'auto',
          borderTop: `1px solid ${toolbarColors.border}`,
          borderBottom: `1px solid ${toolbarColors.border}`,
          margin: '4px 0',
          padding: '2px 0',
        }}
      >
        {hasBlanks && (
          <label style={rowStyle}>
            <input
              type="checkbox"
              data-testid="xlsx-filter-blanks"
              checked={showBlanks}
              onChange={() => setShowBlanks((value) => !value)}
            />
            <span style={valueTextStyle}>{t('filter.blanks')}</span>
          </label>
        )}
        {values.map((value) => (
          <label key={value} style={rowStyle}>
            <input
              type="checkbox"
              checked={checked.has(value)}
              onChange={() => toggleValue(value)}
            />
            <span style={valueTextStyle}>{value}</span>
          </label>
        ))}
        {truncated && (
          <div style={{ ...rowStyle, cursor: 'default', color: toolbarColors.muted }}>
            {t('filter.truncated', { count: values.length })}
          </div>
        )}
      </div>
      <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 6, padding: 2 }}>
        <button
          type="button"
          data-testid="xlsx-filter-clear"
          onClick={() => onClear()}
          style={footerButtonStyle}
        >
          {t('filter.clear')}
        </button>
        <button
          type="button"
          data-testid="xlsx-filter-apply"
          onClick={() => {
            const criteria = resolveCriteria(values, checked, showBlanks);
            onApply(criteria.values, criteria.showBlanks);
          }}
          style={{
            ...footerButtonStyle,
            background: '#217346',
            borderColor: '#217346',
            color: '#ffffff',
          }}
        >
          {t('filter.apply')}
        </button>
      </div>
    </div>
  );
}
