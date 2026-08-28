import {
  createDragRange,
  createNewEventPreview,
  withCreateDragPreview,
  CREATE_DRAG_SNAP_SECONDS,
} from '../internal_packages/main-calendar/lib/core/calendar-drag-utils';

const BASE = 1787860800; // 2026-08-27 16:00:00 UTC, a quarter-hour boundary

function drag(overrides: any = {}) {
  return {
    anchorTime: BASE,
    currentTime: BASE + 3600,
    isAllDay: false,
    isDragging: true,
    calendarId: 'cal-1',
    accountId: 'acct-1',
    ...overrides,
  };
}

describe('createDragRange', function () {
  it('returns the dragged span', function () {
    expect(createDragRange(drag())).toEqual({ start: BASE, end: BASE + 3600 });
  });

  it('orders the range when dragging upward', function () {
    const upward = drag({ anchorTime: BASE + 3600, currentTime: BASE });
    expect(createDragRange(upward)).toEqual({ start: BASE, end: BASE + 3600 });
  });

  it('snaps both ends to the quarter hour', function () {
    const messy = drag({ anchorTime: BASE + 100, currentTime: BASE + 3700 });
    const { start, end } = createDragRange(messy);
    expect(start % CREATE_DRAG_SNAP_SECONDS).toBe(0);
    expect(end % CREATE_DRAG_SNAP_SECONDS).toBe(0);
  });

  it('never produces a zero-length event', function () {
    const barely = drag({ currentTime: BASE + 10 });
    const { start, end } = createDragRange(barely);
    expect(end - start).toBe(CREATE_DRAG_SNAP_SECONDS);
  });

  it('keeps a minimum duration when dragged upward by a hair', function () {
    const barely = drag({ anchorTime: BASE + 10, currentTime: BASE });
    const { start, end } = createDragRange(barely);
    expect(end - start).toBe(CREATE_DRAG_SNAP_SECONDS);
  });
});

describe('createNewEventPreview', function () {
  it('describes a timed range for a drag on the hour grid', function () {
    const preview = createNewEventPreview(drag()) as any;
    expect(preview.isAllDay).toBe(false);
    expect(preview.start).toBe(BASE);
    expect(preview.end).toBe(BASE + 3600);
    expect(preview.isDragPreview).toBe(true);
  });

  it('describes an all-day range for a drag on the all-day row', function () {
    const preview = createNewEventPreview(drag({ isAllDay: true })) as any;
    expect(preview.isAllDay).toBe(true);
    expect(preview.start).toBe(undefined);
  });

  it('carries the calendar it will be created on, so it paints in that colour', function () {
    const preview = createNewEventPreview(drag()) as any;
    expect(preview.calendarId).toBe('cal-1');
    expect(preview.accountId).toBe('acct-1');
  });
});

describe('withCreateDragPreview', function () {
  const existing = [{ id: 'a' }, { id: 'b' }] as any[];

  it('adds nothing when no drag is in progress', function () {
    expect(withCreateDragPreview(existing, null)).toBe(existing);
  });

  it('adds nothing until the drag passes the threshold', function () {
    expect(withCreateDragPreview(existing, drag({ isDragging: false }))).toBe(existing);
  });

  it('appends the preview once dragging, leaving the real events alone', function () {
    const result = withCreateDragPreview(existing, drag());
    expect(result.length).toBe(3);
    expect(result.slice(0, 2)).toEqual(existing);
    expect((result[2] as any).isDragPreview).toBe(true);
  });
});

import { detectHitZone } from '../internal_packages/main-calendar/lib/core/calendar-drag-utils';

/** A vertical (week/day view) event box of the given height, at the origin. */
const box = (height: number) =>
  ({ top: 0, bottom: height, left: 0, right: 100, height, width: 100 }) as DOMRect;

describe('detectHitZone', function () {
  const EDGE = 12;

  describe('on a tall event', function () {
    const tall = box(120);

    it('resizes from the top edge', function () {
      expect(detectHitZone(50, 2, tall, EDGE, 'vertical').mode).toBe('resize-start');
    });

    it('resizes from the bottom edge', function () {
      expect(detectHitZone(50, 118, tall, EDGE, 'vertical').mode).toBe('resize-end');
    });

    it('moves from the middle', function () {
      expect(detectHitZone(50, 60, tall, EDGE, 'vertical').mode).toBe('move');
    });
  });

  describe('on an event shorter than two edge zones', function () {
    // A half-hour event is around twenty pixels tall; two 12px zones would cover all of it.
    const short = box(20);

    it('still offers a middle that moves rather than resizes', function () {
      expect(detectHitZone(50, 10, short, EDGE, 'vertical').mode).toBe('move');
    });

    it('keeps the top edge resizable', function () {
      expect(detectHitZone(50, 1, short, EDGE, 'vertical').mode).toBe('resize-start');
    });

    it('keeps the bottom edge resizable', function () {
      expect(detectHitZone(50, 19, short, EDGE, 'vertical').mode).toBe('resize-end');
    });
  });

  it('leaves a movable middle no matter how short the event', function () {
    for (const height of [6, 10, 14, 20, 30, 45]) {
      const mode = detectHitZone(50, height / 2, box(height), EDGE, 'vertical').mode;
      expect(mode).toBe('move');
    }
  });

  describe('horizontally, for the month view', function () {
    const narrow = { top: 0, bottom: 20, left: 0, right: 21, height: 20, width: 21 } as DOMRect;

    it('keeps a movable middle on a narrow event', function () {
      expect(detectHitZone(10, 10, narrow, EDGE, 'horizontal').mode).toBe('move');
    });

    it('keeps the left edge resizable', function () {
      expect(detectHitZone(1, 10, narrow, EDGE, 'horizontal').mode).toBe('resize-start');
    });
  });
});
