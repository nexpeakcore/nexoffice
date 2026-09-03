import { describe, expect, test } from 'bun:test';

import {
  RESIDENT_WORKER_LAYOUT_BUDGET_MS,
  RESIDENT_WORKER_STARTUP_BUDGET_MS,
  residentWorkerSilenceBudgetMs,
} from './residentWorkerDeadline';

describe('residentWorkerSilenceBudgetMs', () => {
  test('a worker that has said nothing yet is held to the startup budget', () => {
    expect(residentWorkerSilenceBudgetMs(null)).toBe(RESIDENT_WORKER_STARTUP_BUDGET_MS);
  });

  test('the steps before layout stay on the startup budget, where a broken worker fails', () => {
    expect(residentWorkerSilenceBudgetMs('received')).toBe(RESIDENT_WORKER_STARTUP_BUDGET_MS);
    expect(residentWorkerSilenceBudgetMs('sessionReady')).toBe(RESIDENT_WORKER_STARTUP_BUDGET_MS);
    expect(residentWorkerSilenceBudgetMs('stateLoaded')).toBe(RESIDENT_WORKER_STARTUP_BUDGET_MS);
  });

  test('a worker that has proved it is laying out is given room for the document', () => {
    expect(residentWorkerSilenceBudgetMs('layingOut')).toBe(RESIDENT_WORKER_LAYOUT_BUDGET_MS);
    expect(residentWorkerSilenceBudgetMs('laidOut')).toBe(RESIDENT_WORKER_LAYOUT_BUDGET_MS);
  });

  test('proving it is working buys strictly more room than not having started', () => {
    expect(residentWorkerSilenceBudgetMs('layingOut')).toBeGreaterThan(
      residentWorkerSilenceBudgetMs(null)
    );
  });

  test('a broken worker is still given up on, so a wedge is not waited on forever', () => {
    const stages = [null, 'received', 'sessionReady', 'stateLoaded', 'layingOut', 'laidOut'] as const;
    for (const stage of stages) {
      const budget = residentWorkerSilenceBudgetMs(stage);
      expect(Number.isFinite(budget)).toBe(true);
      expect(budget).toBeGreaterThan(0);
    }
  });
});
