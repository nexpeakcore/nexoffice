import type { YrsResidentWorkerOpen } from '@betteroffice/docx/yrs';
import { openingBlockPrefix, openingPageWindow } from '../pageWindow';

/**
 * What a display-list build asks the worker for.
 *
 * A document opens on a prefix of its body so its first pages paint without
 * waiting for its last (see `openingBlockPrefix`). That leaves the worker
 * holding a document which stops early, and nothing else will ask for the
 * rest — so the pass that finishes the job is part of opening, not a
 * repagination the host requested. The hook says which it is by clearing the
 * layout input it remembered.
 */
export type ResidentLayoutRequest = 'open' | 'relayout' | 'frame';

/** The open request: the document's first pages, and only those. */
export function openRequestFor(engineOpen: YrsResidentWorkerOpen): YrsResidentWorkerOpen {
  return {
    ...engineOpen,
    pageWindow: openingPageWindow(),
    firstBlocks: openingBlockPrefix(),
  };
}

export function residentLayoutRequest(
  opening: boolean,
  rememberedLayoutInput: string | null,
  layoutInput: string
): ResidentLayoutRequest {
  if (opening) return 'open';
  return rememberedLayoutInput === layoutInput ? 'frame' : 'relayout';
}

/**
 * Whether a relayout is this hook laying out the rest of a document it opened
 * the start of. Every other relayout means the host rebuilt the region request,
 * which is worth saying out loud — per keystroke it is a bug.
 */
export function completesAPartialOpen(rememberedLayoutInput: string | null): boolean {
  return rememberedLayoutInput === null;
}

