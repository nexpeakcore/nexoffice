import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterAll, afterEach, describe, expect, it } from 'bun:test';
import { FilterPopover } from './FilterPopover';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const { cleanup, fireEvent, getByTestId, render } = await import('@testing-library/react');

afterEach(cleanup);
// happy-dom's fetch rejects file: URLs; leave the runner's globals clean for
// later test files that load wasm from disk.
afterAll(() => GlobalRegistrator.unregister());

interface Applied {
  values: string[] | null;
  showBlanks: boolean;
}

function renderPopover(overrides?: Partial<Parameters<typeof FilterPopover>[0]>) {
  const applied: Applied[] = [];
  let cleared = 0;
  let closed = 0;
  const utils = render(
    <FilterPopover
      columnName="B"
      values={['Alpha', 'Beta']}
      hasBlanks={true}
      truncated={false}
      initialValues={null}
      initialShowBlanks={true}
      onApply={(values, showBlanks) => applied.push({ values, showBlanks })}
      onClear={() => {
        cleared += 1;
      }}
      onClose={() => {
        closed += 1;
      }}
      {...overrides}
    />
  );
  return {
    ...utils,
    applied,
    cleared: () => cleared,
    closed: () => closed,
  };
}

describe('FilterPopover', () => {
  it('applies unconstrained criteria when everything stays checked', () => {
    const { applied, container } = renderPopover();
    fireEvent.click(getByTestId(container, 'xlsx-filter-apply'));
    expect(applied).toEqual([{ values: null, showBlanks: true }]);
  });

  it('applies the checked subset after unchecking a value', () => {
    const { applied, getByLabelText, container } = renderPopover();
    fireEvent.click(getByLabelText('Beta'));
    fireEvent.click(getByTestId(container, 'xlsx-filter-apply'));
    expect(applied).toEqual([{ values: ['Alpha'], showBlanks: true }]);
  });

  it('keeps the full value list when only blanks are excluded', () => {
    const { applied, container } = renderPopover();
    fireEvent.click(getByTestId(container, 'xlsx-filter-blanks'));
    fireEvent.click(getByTestId(container, 'xlsx-filter-apply'));
    expect(applied).toEqual([{ values: ['Alpha', 'Beta'], showBlanks: false }]);
  });

  it('select-all toggles every checkbox including blanks', () => {
    const { applied, container } = renderPopover();
    fireEvent.click(getByTestId(container, 'xlsx-filter-select-all'));
    fireEvent.click(getByTestId(container, 'xlsx-filter-apply'));
    expect(applied).toEqual([{ values: [], showBlanks: false }]);
  });

  it('seeds checkbox state from applied criteria', () => {
    const { applied, container } = renderPopover({
      initialValues: ['Beta'],
      initialShowBlanks: false,
    });
    fireEvent.click(getByTestId(container, 'xlsx-filter-apply'));
    expect(applied).toEqual([{ values: ['Beta'], showBlanks: false }]);
  });

  it('clears through the clear button', () => {
    const rendered = renderPopover();
    fireEvent.click(getByTestId(rendered.container, 'xlsx-filter-clear'));
    expect(rendered.cleared()).toBe(1);
  });

  it('closes on Escape', () => {
    const rendered = renderPopover();
    fireEvent.keyDown(document, { key: 'Escape' });
    expect(rendered.closed()).toBe(1);
  });
});
