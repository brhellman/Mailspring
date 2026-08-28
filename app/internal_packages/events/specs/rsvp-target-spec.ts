import { Calendar, Event } from 'mailspring-exports';
import {
  resolveRSVPTarget,
  resolveDefaultCalendar,
  mayBeAddedToCalendar,
} from '../lib/rsvp-target';

const ACCOUNT_ID = 'acct-1';
const ADDRESSES = ['brian@example.com', 'b.alias@example.com'];
const UID = 'meeting-uid@example.com';

function calendar({
  id,
  name,
  readOnly = false,
  ownership = '',
}: {
  id: string;
  name: string;
  readOnly?: boolean;
  /** What DAV:owner said: 'mine', 'other', or '' when the server didn't answer. */
  ownership?: 'mine' | 'other' | '';
}) {
  return new Calendar({ id, accountId: ACCOUNT_ID, name, readOnly, ownership } as any);
}

function event({
  id,
  calendarId,
  recurrenceId = '',
}: {
  id: string;
  calendarId: string;
  recurrenceId?: string;
}) {
  return new Event({
    id,
    accountId: ACCOUNT_ID,
    calendarId,
    icsuid: UID,
    recurrenceId,
    ics: 'BEGIN:VCALENDAR\r\nEND:VCALENDAR',
  } as any);
}

const MINE = calendar({ id: 'cal-mine', name: 'brian@example.com' });
const ROOM = calendar({ id: 'cal-room', name: '(Conference room) Boardroom' });
const SHARED = calendar({ id: 'cal-shared', name: 'US On Call' });
const HOLIDAYS = calendar({ id: 'cal-holidays', name: 'Holidays', readOnly: true });

describe('resolveRSVPTarget', function () {
  it('picks our own calendar when the meeting is also on a room calendar', function () {
    const mine = event({ id: 'e-mine', calendarId: MINE.id });
    const { target, problem } = resolveRSVPTarget({
      events: [event({ id: 'e-room', calendarId: ROOM.id }), mine],
      calendars: [MINE, ROOM],
      addresses: ADDRESSES,
    });
    expect(problem).toBe(undefined);
    expect(target.event.id).toBe('e-mine');
    expect(target.calendar.id).toBe(MINE.id);
  });

  it("picks our own calendar when the meeting is also on a colleague's shared calendar", function () {
    const { target } = resolveRSVPTarget({
      events: [
        event({ id: 'e-shared', calendarId: SHARED.id }),
        event({ id: 'e-mine', calendarId: MINE.id }),
      ],
      calendars: [SHARED, MINE],
      addresses: ADDRESSES,
    });
    expect(target.calendar.id).toBe(MINE.id);
  });

  it('recognises our calendar by an alias as well as the primary address', function () {
    const aliasCal = calendar({ id: 'cal-alias', name: 'b.alias@example.com' });
    const { target } = resolveRSVPTarget({
      events: [
        event({ id: 'e-room', calendarId: ROOM.id }),
        event({ id: 'e-alias', calendarId: aliasCal.id }),
      ],
      calendars: [ROOM, aliasCal],
      addresses: ADDRESSES,
    });
    expect(target.calendar.id).toBe(aliasCal.id);
  });

  it('takes the only writable copy when none of the calendars is named for us', function () {
    const { target } = resolveRSVPTarget({
      events: [event({ id: 'e-shared', calendarId: SHARED.id })],
      calendars: [SHARED, HOLIDAYS],
      addresses: ADDRESSES,
    });
    expect(target.calendar.id).toBe(SHARED.id);
  });

  it('refuses to guess between two writable calendars that are not ours', function () {
    const { target, problem } = resolveRSVPTarget({
      events: [
        event({ id: 'e-room', calendarId: ROOM.id }),
        event({ id: 'e-shared', calendarId: SHARED.id }),
      ],
      calendars: [ROOM, SHARED],
      addresses: ADDRESSES,
    });
    expect(target).toBe(null);
    expect(problem).toBe('ambiguous');
  });

  it('reports read-only when the only copy is on a calendar we cannot write', function () {
    const { target, problem } = resolveRSVPTarget({
      events: [event({ id: 'e-hol', calendarId: HOLIDAYS.id })],
      calendars: [HOLIDAYS],
      addresses: ADDRESSES,
    });
    expect(target).toBe(null);
    expect(problem).toBe('read-only');
  });

  it('reports not-on-a-calendar when the invitation has not synced anywhere', function () {
    const { target, problem } = resolveRSVPTarget({
      events: [],
      calendars: [MINE],
      addresses: ADDRESSES,
    });
    expect(target).toBe(null);
    expect(problem).toBe('not-on-a-calendar');
  });

  it('ignores a copy whose calendar the account does not have', function () {
    const { target, problem } = resolveRSVPTarget({
      events: [event({ id: 'e-orphan', calendarId: 'cal-that-went-away' })],
      calendars: [MINE],
      addresses: ADDRESSES,
    });
    expect(target).toBe(null);
    expect(problem).toBe('read-only');
  });

  describe('recurring series', function () {
    it('answers the master rather than a modified occurrence', function () {
      const { target } = resolveRSVPTarget({
        events: [
          event({ id: 'e-exception', calendarId: MINE.id, recurrenceId: '20260315T140000Z' }),
          event({ id: 'e-master', calendarId: MINE.id }),
        ],
        calendars: [MINE],
        addresses: ADDRESSES,
      });
      expect(target.event.id).toBe('e-master');
    });

    it('declines when only a modified occurrence has synced', function () {
      const { target, problem } = resolveRSVPTarget({
        events: [
          event({ id: 'e-exception', calendarId: MINE.id, recurrenceId: '20260315T140000Z' }),
        ],
        calendars: [MINE],
        addresses: ADDRESSES,
      });
      expect(target).toBe(null);
      expect(problem).toBe('not-on-a-calendar');
    });
  });

  it('prefers our own calendar even when it sorts last', function () {
    const { target } = resolveRSVPTarget({
      events: [
        event({ id: 'e-room', calendarId: ROOM.id }),
        event({ id: 'e-shared', calendarId: SHARED.id }),
        event({ id: 'e-mine', calendarId: MINE.id }),
      ],
      calendars: [ROOM, SHARED, MINE],
      addresses: ADDRESSES,
    });
    expect(target.calendar.id).toBe(MINE.id);
  });
});

describe('resolveDefaultCalendar', function () {
  it('picks the calendar named for the account', function () {
    const cal = resolveDefaultCalendar([ROOM, SHARED, MINE, HOLIDAYS], ADDRESSES);
    expect(cal.id).toBe(MINE.id);
  });

  it('matches on an alias too', function () {
    const aliasCal = calendar({ id: 'cal-alias', name: 'b.alias@example.com' });
    expect(resolveDefaultCalendar([ROOM, aliasCal], ADDRESSES).id).toBe(aliasCal.id);
  });

  it('takes the only writable calendar when none is named for us', function () {
    expect(resolveDefaultCalendar([SHARED, HOLIDAYS], ADDRESSES).id).toBe(SHARED.id);
  });

  it('refuses to choose between several calendars that are not ours', function () {
    expect(resolveDefaultCalendar([ROOM, SHARED], ADDRESSES)).toBe(null);
  });

  it('never returns a read-only calendar', function () {
    expect(resolveDefaultCalendar([HOLIDAYS], ADDRESSES)).toBe(null);
  });

  it('returns null when the account has no calendars', function () {
    expect(resolveDefaultCalendar([], ADDRESSES)).toBe(null);
  });
});

// ---------------------------------------------------------------------------
// Calendar ownership. A calendar's display name is text its owner chooses, so it can't be
// the only thing that decides where a reply gets written. DAV:owner is the server's answer
// and overrides it in both directions.
// ---------------------------------------------------------------------------

describe('resolveRSVPTarget with server-reported ownership', function () {
  const shared = (over: any = {}) =>
    calendar({ id: 'cal-shared', name: 'Team', ownership: 'other', ...over });

  it("refuses a calendar the server says is someone else's, even named after us", function () {
    // The attack this exists to stop: share a writable calendar, name it after the
    // recipient's own address, and collect the replies they meant for their own copy.
    const impostor = calendar({
      id: 'cal-impostor',
      name: 'brian@example.com',
      ownership: 'other',
    });
    const resolution = resolveRSVPTarget({
      events: [event({ id: 'e1', calendarId: 'cal-impostor' })],
      calendars: [impostor],
      addresses: ADDRESSES,
    });
    expect(resolution.target).toBe(null);
    // Reported as somebody else's rather than as ambiguous: there is nothing ambiguous
    // about it, and the UI says so instead of blaming the user's calendar layout.
    expect(resolution.problem).toBe('not-ours');
  });

  it('accepts a calendar the server says is ours, whatever it is called', function () {
    // The other half: a default calendar named "Calendar" or "Kalender" is still ours, and
    // the name heuristic alone would never have found it.
    const mine = calendar({ id: 'cal-plain', name: 'Calendar', ownership: 'mine' });
    const resolution = resolveRSVPTarget({
      events: [event({ id: 'e1', calendarId: 'cal-plain' })],
      calendars: [mine],
      addresses: ADDRESSES,
    });
    expect(resolution.target && resolution.target.calendar.id).toBe('cal-plain');
  });

  it("never falls back to a foreign calendar just because it's the only writable one", function () {
    const resolution = resolveRSVPTarget({
      events: [event({ id: 'e1', calendarId: 'cal-shared' })],
      calendars: [shared()],
      addresses: ADDRESSES,
    });
    expect(resolution.target).toBe(null);
    expect(resolution.problem).toBe('not-ours');
  });

  it('still uses the name when the server reports no ownership at all', function () {
    // Servers may omit DAV:owner; that must not make every calendar unusable.
    const resolution = resolveRSVPTarget({
      events: [event({ id: 'e1', calendarId: 'cal-mine' })],
      calendars: [MINE],
      addresses: ADDRESSES,
    });
    expect(resolution.target && resolution.target.calendar.id).toBe('cal-mine');
  });
});

describe('resolveDefaultCalendar with server-reported ownership', function () {
  it("won't add an accepted invitation to somebody else's calendar", function () {
    const shared = calendar({ id: 'cal-shared', name: 'Team', ownership: 'other' });
    expect(resolveDefaultCalendar([shared], ADDRESSES)).toBe(null);
  });

  it('prefers the calendar the server says is ours', function () {
    const mine = calendar({ id: 'cal-plain', name: 'Calendar', ownership: 'mine' });
    const shared = calendar({ id: 'cal-shared', name: 'brian@example.com', ownership: 'other' });
    const chosen = resolveDefaultCalendar([mine, shared], ADDRESSES);
    expect(chosen && chosen.id).toBe('cal-plain');
  });
});

// ---------------------------------------------------------------------------
// Storing an emailed invitation. Anyone can send a message; a stored event naming us as
// ORGANIZER makes our own server mail a REQUEST to every ATTENDEE the sender chose.
// ---------------------------------------------------------------------------

describe('mayBeAddedToCalendar', function () {
  it('stores an invitation organized by someone else', function () {
    expect(mayBeAddedToCalendar('mailto:ada@example.com', ADDRESSES)).toBe(true);
  });

  it('refuses one that names us as the organizer', function () {
    expect(mayBeAddedToCalendar('mailto:brian@example.com', ADDRESSES)).toBe(false);
  });

  it('refuses one that names an alias of ours as the organizer', function () {
    expect(mayBeAddedToCalendar('mailto:b.alias@example.com', ADDRESSES)).toBe(false);
  });

  it('is not fooled by capitalisation or a bare address', function () {
    expect(mayBeAddedToCalendar('MAILTO:BRIAN@EXAMPLE.COM', ADDRESSES)).toBe(false);
    expect(mayBeAddedToCalendar('brian@example.com', ADDRESSES)).toBe(false);
  });

  it('refuses an event that names no organizer, which is not a scheduled event', function () {
    expect(mayBeAddedToCalendar('', ADDRESSES)).toBe(false);
    expect(mayBeAddedToCalendar(null as any, ADDRESSES)).toBe(false);
  });
});
