import { describe, expect, it } from 'bun:test';
import {
  DEFAULT_COMMENT_AUTHOR,
  commentAt,
  isSavableCommentText,
  resolveCommentAuthor,
} from './commentState';

const comments = [
  { row: 0, col: 0, author: 'Ada', text: 'first' },
  { row: 2, col: 3, author: 'Grace', text: 'second' },
];

describe('commentAt', () => {
  it('finds the comment at exact coordinates', () => {
    expect(commentAt(comments, 2, 3)).toEqual(comments[1]);
  });

  it('returns null for uncommented cells', () => {
    expect(commentAt(comments, 2, 0)).toBeNull();
    expect(commentAt([], 0, 0)).toBeNull();
  });
});

describe('resolveCommentAuthor', () => {
  it('keeps the existing author when editing', () => {
    expect(resolveCommentAuthor({ author: 'Ada' }, 'Someone Else')).toBe('Ada');
  });

  it('uses the host user name for new comments', () => {
    expect(resolveCommentAuthor(null, 'Hoang Giang')).toBe('Hoang Giang');
  });

  it('falls back to the default for blank existing and host names', () => {
    expect(resolveCommentAuthor({ author: '  ' }, undefined)).toBe(DEFAULT_COMMENT_AUTHOR);
    expect(resolveCommentAuthor(null, '   ')).toBe(DEFAULT_COMMENT_AUTHOR);
    expect(resolveCommentAuthor(null, undefined)).toBe(DEFAULT_COMMENT_AUTHOR);
  });

  it('trims a padded host name', () => {
    expect(resolveCommentAuthor(null, '  Ada  ')).toBe('Ada');
  });
});

describe('isSavableCommentText', () => {
  it('rejects blank drafts and accepts real text', () => {
    expect(isSavableCommentText('')).toBe(false);
    expect(isSavableCommentText('  \n ')).toBe(false);
    expect(isSavableCommentText('note')).toBe(true);
  });
});
