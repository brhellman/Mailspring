import IcalExpander from 'ical-expander';
import { Event } from './flux/models/event';
import { emailFromParticipantURI } from './calendar-utils';

/** One busy occurrence that overlaps the window being checked. */
export interface CalendarConflict {
  /** The stored event the occurrence was expanded from. */
  eventId: string;
  calendarId: string;
  icsuid: string;
  title: string;
  /** Unix seconds. */
  start: number;
  end: number;
  isAllDay: boolean;
}

export interface FindConflictsOptions {
  /** Candidate events, already narrowed by the caller to the account and rough time range. */
  events: Event[];
  /** The window to test, in unix seconds. */
  start: number;
  end: number;
  /** The addresses that count as "me", used to skip meetings this account has declined. */
  addresses: string[];
  /** UID of the event being checked, so it never conflicts with itself. */
  excludeIcsuid?: string;
  /**
   * All-day events are excluded by default. A day marked "Vacation" overlaps everything on
   * that day, so counting it would flag every meeting in the week as conflicted.
   */
  includeAllDay?: boolean;
}

type ICALComponent = any;

/**
 * True when this VEVENT should not be treated as taking up the user's time: cancelled,
 * marked free rather than busy (TRANSP:TRANSPARENT, what Google calls "Free"), or an
 * invitation this account has declined.
 */
function isFreeTime(component: ICALComponent, addresses: string[]): boolean {
  const status = String(component.getFirstPropertyValue('status') || '').toUpperCase();
  if (status === 'CANCELLED') {
    return true;
  }

  const transp = String(component.getFirstPropertyValue('transp') || '').toUpperCase();
  if (transp === 'TRANSPARENT') {
    return true;
  }

  const lowered = addresses.map((a) => a.toLowerCase());
  for (const attendee of component.getAllProperties('attendee')) {
    const email = attendee
      .getValues()
      .map(String)
      .map(emailFromParticipantURI)
      .find((v) => !!v);
    if (email && lowered.includes(email)) {
      const partstat = String(attendee.getParameter('partstat') || '').toUpperCase();
      return partstat === 'DECLINED';
    }
  }

  return false;
}

/**
 * How many occurrences IcalExpander may step through before giving up on one series.
 *
 * It iterates forward from DTSTART with no way to seek, so a cap is a limit on how far back
 * a series may begin rather than on the work the window costs. The library's default of 1000
 * is too low to be safe - a daily meeting that started more than about three years ago never
 * reaches the window and its conflicts go silently unreported - but removing the cap is worse:
 * an invitation is untrusted input, and `RRULE:FREQ=SECONDLY` dated 1970 would spin the
 * renderer forever.
 *
 * So the budget is derived from the series itself: how many steps of its own frequency fit
 * between where it starts and the end of the window, plus slack. Realistic calendars land far
 * below the ceiling - a daily series running since 2000 needs about 9,000 - while a frequency
 * fine enough to be abusive exceeds it and is truncated instead of expanded.
 */
const STEP_SECONDS: { [freq: string]: number } = {
  SECONDLY: 1,
  MINUTELY: 60,
  HOURLY: 3600,
  DAILY: 86400,
  WEEKLY: 604800,
  // Deliberately the shortest month and year. Underestimating the step overestimates the
  // budget, which errs towards expanding a legitimate series rather than truncating it.
  MONTHLY: 28 * 86400,
  YEARLY: 365 * 86400,
};
const MIN_ITERATIONS = 1000;
const MAX_ITERATIONS = 50000;

function iterationBudget(event: Event, windowEnd: number): number {
  const rrule = /^RRULE:(.*)$/im.exec(event.ics);
  if (!rrule) {
    return MIN_ITERATIONS; // not a series; one occurrence is all there is to reach
  }
  const freq = /FREQ=([A-Z]+)/i.exec(rrule[1]);
  const interval = parseInt((/INTERVAL=(\d+)/i.exec(rrule[1]) || [])[1], 10) || 1;
  const step = (STEP_SECONDS[(freq ? freq[1] : '').toUpperCase()] || STEP_SECONDS.DAILY) * interval;
  const steps = Math.ceil(Math.max(0, windowEnd - event.recurrenceStart) / step) + 100;
  return Math.min(MAX_ITERATIONS, Math.max(MIN_ITERATIONS, steps));
}

/**
 * Finds the occurrences on the user's calendars that overlap a window - what Google Calendar
 * shows as "Conflicts with…" on an invitation.
 *
 * Recurring series are expanded over the window rather than compared by their stored
 * `recurrenceStart`/`recurrenceEnd`, which span the whole series and would report a weekly
 * standup as conflicting with every meeting for the next two years.
 *
 * Overlap is half-open: an event that ends exactly when the window starts does not conflict,
 * so back-to-back meetings read as back-to-back rather than as a clash.
 */
export function findConflicts({
  events,
  start,
  end,
  addresses,
  excludeIcsuid,
  includeAllDay = false,
}: FindConflictsOptions): CalendarConflict[] {
  if (!(end > start)) {
    return [];
  }

  const conflicts: CalendarConflict[] = [];
  const seen = new Set<string>();

  for (const event of events) {
    if (excludeIcsuid && event.icsuid === excludeIcsuid) {
      continue;
    }
    // Exceptions are carried inline in the master's ICS, so expanding the master covers
    // them. Expanding an exception row as well would report the occurrence twice.
    if (event.isRecurrenceException()) {
      continue;
    }

    let expanded: { events: any[]; occurrences: any[] };
    try {
      expanded = new IcalExpander({
        ics: event.ics,
        maxIterations: iterationBudget(event, end),
      }).between(new Date(start * 1000), new Date(end * 1000));
    } catch (err) {
      // An unparseable calendar shouldn't fail the whole check; the cost is one conflict
      // going unmentioned.
      continue;
    }

    for (const entry of [...expanded.events, ...expanded.occurrences]) {
      const item = 'item' in entry ? entry.item : entry;
      const component = item.component;
      if (!component || isFreeTime(component, addresses)) {
        continue;
      }

      const isAllDay = !!entry.startDate.isDate;
      if (isAllDay && !includeAllDay) {
        continue;
      }

      const occurrenceStart = Math.round(entry.startDate.toJSDate().getTime() / 1000);
      const occurrenceEnd = Math.round(entry.endDate.toJSDate().getTime() / 1000);
      if (occurrenceEnd <= start || occurrenceStart >= end) {
        continue;
      }

      const key = `${event.id}-${occurrenceStart}`;
      if (seen.has(key)) {
        continue;
      }
      seen.add(key);

      conflicts.push({
        eventId: event.id,
        calendarId: event.calendarId,
        icsuid: event.icsuid,
        title: item.summary || '',
        start: occurrenceStart,
        end: occurrenceEnd,
        isAllDay,
      });
    }
  }

  return conflicts.sort((a, b) => a.start - b.start || a.title.localeCompare(b.title));
}
