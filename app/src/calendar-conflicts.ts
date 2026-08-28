import IcalExpander from 'ical-expander';
import { Event } from './flux/models/event';
import { emailFromParticipantURI } from './calendar-utils';
import { expansionIterationBudget } from './ics-event-helpers';

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
        maxIterations: expansionIterationBudget(event.ics, event.recurrenceStart, end),
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
