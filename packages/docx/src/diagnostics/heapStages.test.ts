import { beforeEach, describe, expect, test } from 'bun:test';
import {
  clearHeapStages,
  heapStages,
  markHeapStage,
  onHeapStage,
  registerHeapProbe,
} from './index';

function countingProbe(): { live: number } {
  const state = { live: 0 };
  registerHeapProbe({
    liveBytes: () => (state.live += 1_000_000),
    peakBytes: () => state.live + 500_000,
    resetPeak: () => {},
  });
  return state;
}

describe('heap stages', () => {
  beforeEach(() => {
    registerHeapProbe(null);
    clearHeapStages();
  });

  test('describes each phase of the open once', () => {
    countingProbe();
    markHeapStage('seed');
    markHeapStage('layout');
    markHeapStage('display');
    expect(heapStages().map((stage) => stage.label)).toEqual(['seed', 'layout', 'display']);
  });

  test('an edit does not append another open phase', () => {
    countingProbe();
    markHeapStage('layout');
    markHeapStage('display');
    const afterOpen = heapStages();

    // Every keystroke relayouts and rebuilds the display. Before this, each one
    // appended another pair, so the list grew without bound and the report's
    // "total held" read the last edit rather than the open it is labelled as.
    for (let edit = 0; edit < 50; edit += 1) {
      markHeapStage('layout');
      markHeapStage('display');
    }
    expect(heapStages()).toEqual(afterOpen);
  });

  test('stays quiet after the open, so hosts stop re-sending per keystroke', () => {
    let notifications = 0;
    const stop = onHeapStage(() => {
      notifications += 1;
    });
    markHeapStage('layout');
    markHeapStage('display');
    expect(notifications).toBe(2);

    for (let edit = 0; edit < 50; edit += 1) {
      markHeapStage('layout');
      markHeapStage('display');
    }
    expect(notifications).toBe(2);
    stop();
  });

  test('notifies for each phase without a probe', () => {
    // A build without the counters still needs the boundary signal.
    let notifications = 0;
    const stop = onHeapStage(() => {
      notifications += 1;
    });
    markHeapStage('seed');
    markHeapStage('layout');
    expect(notifications).toBe(2);
    expect(heapStages()).toEqual([]);
    stop();
  });

  test('the next document starts over', () => {
    countingProbe();
    markHeapStage('layout');
    clearHeapStages();
    markHeapStage('layout');
    expect(heapStages().map((stage) => stage.label)).toEqual(['layout']);
  });
});
