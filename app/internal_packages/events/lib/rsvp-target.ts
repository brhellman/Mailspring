import ICAL from 'ical.js';
import { Calendar, CalendarUtils, Event, Utils } from 'mailspring-exports';

/** Where an RSVP will be recorded, once we're certain which copy of the event is ours. */
export interface RSVPTarget {
  event: Event;
  calendar: Calendar;
}

/**
 * Why no copy of the event can be answered. The RSVP still goes to the organizer by email;
 * these values exist so the UI can say what will and won't happen before the user clicks.
 */
export type RSVPTargetProblem =
  /** The event isn't on any calendar we sync. */
  | 'not-on-a-calendar'
  /** Every copy we found is on a read-only calendar. */
  | 'read-only'
  /** Every writable copy is on a calendar the server says belongs to somebody else. */
  | 'not-ours'
  /** Copies on more than one writable calendar, and none is identifiably ours. */
  | 'ambiguous';

export type RSVPTargetResolution =
  | { target: RSVPTarget; problem?: undefined }
  | { target: null; problem: RSVPTargetProblem };

/**
 * Picks the copy of an invitation that this account is entitled to answer.
 *
 * One VEVENT UID can land on several calendars at once. A meeting booked into a room shows
 * up on the room's calendar, a colleague's copy shows up on any calendar of theirs we've
 * subscribed to, and our own copy shows up on our own calendar - all with the same UID, all
 * synced into the same account. Writing our response to the wrong one would edit a resource
 * or a colleague's event, so this returns a target only when it can name one with certainty
 * and reports why it couldn't otherwise.
 *
 * @param events - Every Event row matching the invitation's UID for one account.
 * @param calendars - Every calendar belonging to that account.
 * @param addresses - The account's address and its aliases.
 */
export function resolveRSVPTarget({
  events,
  calendars,
  addresses,
}: {
  events: Event[];
  calendars: Calendar[];
  addresses: string[];
}): RSVPTargetResolution {
  if (!events.length) {
    return { target: null, problem: 'not-on-a-calendar' };
  }

  const calendarsById = new Map(calendars.map((c) => [c.id, c]));
  const writable = events.filter((e) => {
    const calendar = calendarsById.get(e.calendarId);
    return calendar && !calendar.readOnly;
  });
  if (!writable.length) {
    return { target: null, problem: 'read-only' };
  }

  // A calendar the server has named as someone else's is never a candidate, however few
  // others there are - the sole-candidate fallback below exists for calendars whose
  // ownership is simply unknown, not as a reason to write onto a colleague's.
  const calendarIds = new Set(
    writable
      .filter((e) => !CalendarUtils.isSomeoneElsesCalendar(calendarsById.get(e.calendarId)))
      .map((e) => e.calendarId)
  );
  if (!calendarIds.size) {
    return { target: null, problem: 'not-ours' };
  }
  const ownCalendarIds = [...calendarIds].filter((id) =>
    CalendarUtils.isOwnCalendar(calendarsById.get(id), addresses)
  );

  // Our own calendar wins outright. Failing that we'll take the only candidate there is, but
  // never a guess between several: an unanswered invitation is a smaller problem than a
  // response written onto someone else's event.
  let calendarId: string;
  if (ownCalendarIds.length === 1) {
    calendarId = ownCalendarIds[0];
  } else if (ownCalendarIds.length === 0 && calendarIds.size === 1) {
    calendarId = [...calendarIds][0];
  } else {
    return { target: null, problem: 'ambiguous' };
  }

  // A recurring series is answered on its master event, which is what carries the attendee
  // list for every occurrence. Exceptions repeat the UID, so answering one of those would
  // record the response against a single occurrence and leave the rest untouched.
  const onCalendar = writable.filter((e) => e.calendarId === calendarId);
  const event = onCalendar.find((e) => !e.isRecurrenceException());
  if (!event) {
    return { target: null, problem: 'not-on-a-calendar' };
  }

  return { target: { event, calendar: calendarsById.get(calendarId) } };
}

/**
 * Whether an emailed invitation may be stored on one of our calendars.
 *
 * An event whose ORGANIZER is one of our own addresses becomes an organizer resource once it
 * is on our server, and a server implementing scheduling (RFC 6638 section 3.2.1) then mails
 * a REQUEST to every ATTENDEE it carries, from us. That attendee list arrives in a message
 * anyone can send, so storing such an event would turn an inbound message into outbound mail
 * addressed by its sender.
 *
 * We are only ever a guest on an invitation that reaches us by email. One claiming otherwise
 * is answered by email alone and never written to a calendar.
 *
 * @param organizerUri - The invitation's ORGANIZER value, in any form emailFromParticipantURI
 *   accepts. An invitation naming no organizer isn't a scheduled event and is also refused.
 * @param addresses - The account's address and its aliases.
 */
export function mayBeAddedToCalendar(organizerUri: string, addresses: string[]): boolean {
  const organizer = CalendarUtils.emailFromParticipantURI(organizerUri);
  if (!organizer) return false;
  return !addresses.some((a) => Utils.emailIsEquivalent(a, organizer));
}

/** Why a counter-proposal that arrived by email may not be applied to our copy. */
export type CounterProposalProblem =
  /** Our copy of the meeting is not one this account organizes. */
  | 'not-our-meeting'
  /** The message did not come from somebody on the guest list. */
  | 'not-from-a-guest';

/**
 * Whether an emailed COUNTER may move our copy of the meeting.
 *
 * A COUNTER asks the organizer to move a meeting (RFC 5546 section 3.2.7), and accepting one
 * rewrites DTSTART/DTEND on our calendar and re-invites every guest. The proposal itself is
 * just a message: its UID, its ORGANIZER and its ATTENDEE list are all written by whoever
 * sent it, and a UID is not a secret - it travels in the invitation to every guest, and in
 * every reply. So nothing in the attachment can decide this. Both questions are asked of the
 * copy already on our calendar:
 *
 * - do we organize this meeting, since only the organizer may revise one (section 2.1.4), and
 * - is the sender on its guest list, since only an attendee may counter (section 3.2.7).
 *
 * A meeting whose ORGANIZER the server has rewritten to a shared calendar's own address -
 * Google does this for events on a `...@group.calendar.google.com` calendar - reads as
 * somebody else's and is refused. That is a button not offered, not a meeting moved wrongly.
 *
 * @param ics - Our synced copy of the event, never the emailed attachment.
 * @param senderEmail - The From address of the message carrying the COUNTER.
 * @param addresses - The account's address and its aliases.
 */
export function counterProposalProblem({
  ics,
  senderEmail,
  addresses,
}: {
  ics: string;
  senderEmail: string | null;
  addresses: string[];
}): CounterProposalProblem | null {
  let event: ICAL.Event;
  try {
    event = CalendarUtils.parseICSString(ics).event;
  } catch (e) {
    return 'not-our-meeting';
  }

  const organizer = event.organizer ? CalendarUtils.emailFromParticipantURI(event.organizer) : null;
  if (!organizer || !addresses.some((a) => Utils.emailIsEquivalent(a, organizer))) {
    return 'not-our-meeting';
  }

  if (!senderEmail) {
    return 'not-from-a-guest';
  }
  const guests = CalendarUtils.cleanParticipants(event)
    .map((p) => p.email)
    .filter((email): email is string => !!email);
  if (!guests.some((guest) => Utils.emailIsEquivalent(guest, senderEmail))) {
    return 'not-from-a-guest';
  }

  return null;
}

/**
 * The calendars whose events count as the user being busy.
 *
 * Read-only calendars are left out because they hold subscribed feeds - public holidays,
 * sports fixtures - that say nothing about the user's availability. Calendars the user has
 * switched off in the sidebar are left out too, which is what makes the result predictable:
 * conflicts come from the calendars they can see.
 */
export function conflictCalendarIds(
  calendars: Calendar[],
  disabledCalendarIds: string[]
): string[] {
  const disabled = new Set(disabledCalendarIds);
  return calendars.filter((c) => !c.readOnly && !disabled.has(c.id)).map((c) => c.id);
}

/**
 * The calendar an invitation should be added to when it isn't on any of ours yet.
 *
 * Google only puts an invitation on your calendar once you answer it through Google itself -
 * with "Add invitations to my calendar" on its default setting, an emailed invitation you
 * haven't answered exists nowhere but the message. Accepting from the client therefore has
 * to create the event, and this decides where.
 *
 * The rule is the same one resolveRSVPTarget uses, and refuses in the same circumstances:
 * our own calendar if we can name it, the only writable one if there's exactly one, and
 * otherwise nothing, because putting a colleague's meeting on a shared team calendar is
 * worse than leaving it off ours.
 */
export function resolveDefaultCalendar(
  calendars: Calendar[],
  addresses: string[]
): Calendar | null {
  const writable = calendars.filter((c) => !c.readOnly && !CalendarUtils.isSomeoneElsesCalendar(c));
  const own = writable.filter((c) => CalendarUtils.isOwnCalendar(c, addresses));
  if (own.length === 1) {
    return own[0];
  }
  if (own.length === 0 && writable.length === 1) {
    return writable[0];
  }
  return null;
}
