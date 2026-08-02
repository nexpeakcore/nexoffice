import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { afterEach, describe, expect, it } from 'bun:test';
import { CommentEditorPopover, CommentViewPopover } from './CommentPopover';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
const { cleanup, fireEvent, getByTestId, queryByTestId, render } = await import(
  '@testing-library/react'
);

afterEach(cleanup);

function renderEditor(overrides?: Partial<Parameters<typeof CommentEditorPopover>[0]>) {
  const saved: string[] = [];
  let deleted = 0;
  let closed = 0;
  const utils = render(
    <CommentEditorPopover
      cellName="B2"
      author="Ada"
      initialText=""
      canDelete={false}
      onSave={(text) => saved.push(text)}
      onDelete={() => {
        deleted += 1;
      }}
      onClose={() => {
        closed += 1;
      }}
      {...overrides}
    />
  );
  return {
    ...utils,
    saved,
    deleted: () => deleted,
    closed: () => closed,
  };
}

describe('CommentViewPopover', () => {
  it('shows the author and text', () => {
    let closed = 0;
    const { container } = render(
      <CommentViewPopover
        cellName="B2"
        author="Ada"
        text="Check this"
        onClose={() => {
          closed += 1;
        }}
      />
    );
    expect(getByTestId(container, 'xlsx-comment-author').textContent).toBe('Ada');
    expect(getByTestId(container, 'xlsx-comment-text').textContent).toBe('Check this');
    fireEvent.keyDown(document, { key: 'Escape' });
    expect(closed).toBe(1);
  });
});

describe('CommentEditorPopover', () => {
  it('saves the trimmed draft', () => {
    const { container, saved } = renderEditor();
    fireEvent.change(getByTestId(container, 'xlsx-comment-input'), {
      target: { value: '  a note  ' },
    });
    fireEvent.click(getByTestId(container, 'xlsx-comment-save'));
    expect(saved).toEqual(['a note']);
  });

  it('disables save while the draft is blank', () => {
    const { container, saved } = renderEditor();
    const save = getByTestId(container, 'xlsx-comment-save') as HTMLButtonElement;
    expect(save.disabled).toBe(true);
    fireEvent.click(save);
    expect(saved).toEqual([]);
  });

  it('prefills an existing comment and offers delete', () => {
    const { container, deleted } = renderEditor({ initialText: 'old text', canDelete: true });
    const input = getByTestId(container, 'xlsx-comment-input') as HTMLTextAreaElement;
    expect(input.value).toBe('old text');
    fireEvent.click(getByTestId(container, 'xlsx-comment-delete'));
    expect(deleted()).toBe(1);
  });

  it('hides delete for cells without a comment', () => {
    const { container } = renderEditor();
    expect(queryByTestId(container, 'xlsx-comment-delete')).toBeNull();
  });

  it('closes on Escape', () => {
    const rendered = renderEditor();
    fireEvent.keyDown(document, { key: 'Escape' });
    expect(rendered.closed()).toBe(1);
  });
});
