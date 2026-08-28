import {
  AccountStore,
  Actions,
  Calendar,
  CalendarUtils,
  DatabaseStore,
  Event,
  EventRSVPTask,
  ICSEventHelpers,
  ICSParticipantStatus,
  SyncbackEventTask,
  localized,
} from 'mailspring-exports';
import React from 'react';
import { ProposeTimePopover } from 'mailspring-component-kit';
import { EventOccurrence, occurrenceStartUnix, occurrenceEndUnix } from './calendar-data-source';
import { parseEventIdFromOccurrence } from './calendar-drag-utils';

/** Whether this occurrence is an invitation we can answer, rather than one we sent. */
export function canRespondToEvent(occurrence: EventOccurrence): boolean {
  const me = myAttendeeEmail(occurrence);
  if (!me) return false;

  // The organizer doesn't RSVP to their own meeting; they change it.
  const organizer =
    occurrence.organizer && CalendarUtils.emailFromParticipantURI(occurrence.organizer.email);
  return !organizer || organizer.toLowerCase() !== me;
}

/**
 * Our address on this event's guest list, lowercased.
 *
 * Restricted to the account the event is synced under. Any other connected account may also
 * appear as a guest, and answering as them would write a status the event's own account
 * cannot act on - respondToCalendarEvent resolves the same person through selfParticipant,
 * which is account-scoped, and would find nothing.
 */
export function myAttendeeEmail(occurrence: EventOccurrence): string | null {
  for (const attendee of occurrence.attendees || []) {
    if (!attendee.email) continue;
    const account = AccountStore.accountForEmail(attendee.email);
    if (account && account.id === occurrence.accountId) {
      return attendee.email.toLowerCase();
    }
  }
  return null;
}

/** Our current participation status for this occurrence, if we're on the guest list. */
export function myParticipationStatus(occurrence: EventOccurrence): string | null {
  const me = myAttendeeEmail(occurrence);
  if (!me) return null;
  const mine = (occurrence.attendees || []).find((a) => a.email && a.email.toLowerCase() === me);
  return mine ? (mine.partstat || 'NEEDS-ACTION').toUpperCase() : null;
}

/**
 * Answers an invitation from the calendar, doing both halves of an RSVP.
 *
 * A response has to reach two places: the organizer, by an emailed iTIP REPLY, and our own
 * copy of the event, by writing our PARTSTAT back over CalDAV (RFC 6638 section 3.2.5).
 * Doing only the first leaves our calendar showing the invitation as unanswered; doing only
 * the second never tells the organizer. The email is sent even when the calendar write
 * can't happen, because informing the organizer is the part they're waiting on.
 */
export async function respondToCalendarEvent(
  occurrence: EventOccurrence,
  status: ICSParticipantStatus
): Promise<void> {
  const eventId = parseEventIdFromOccurrence(occurrence.id);
  const event = await DatabaseStore.find<Event>(Event, eventId);
  if (!event) {
    console.warn(`Calendar RSVP: could not find event ${eventId}`);
    return;
  }

  let parsed: ReturnType<typeof CalendarUtils.parseICSString>;
  try {
    parsed = CalendarUtils.parseICSString(event.ics);
  } catch (e) {
    AppEnv.showErrorDialog(localized("Sorry, this event's data could not be read."));
    return;
  }

  const me = CalendarUtils.selfParticipant(parsed.event, event.accountId);
  if (!me || !me.email) {
    AppEnv.showErrorDialog(
      localized("You're not on this event's guest list, so there's no RSVP to give.")
    );
    return;
  }

  const calendar = await DatabaseStore.find<Calendar>(Calendar, event.calendarId);
  if (calendar && !calendar.readOnly) {
    const ics = ICSEventHelpers.updateAttendeeStatus(event.ics, me.email, status);
    if (ics && ics !== event.ics) {
      const updated = event.clone();
      updated.ics = ics;
      // Not undoable: an undo would have to retract the reply we email below, and iTIP
      // gives us no way to do that.
      Actions.queueTask(SyncbackEventTask.forUpdating({ event: updated }));
    }
  }

  const organizerEmail = CalendarUtils.emailFromParticipantURI(parsed.event.organizer);
  if (!organizerEmail) return; // nothing to tell - an event with no organizer isn't scheduled

  try {
    Actions.queueTask(
      EventRSVPTask.forReplying({
        accountId: event.accountId,
        icsOriginalData: event.ics,
        icsRSVPStatus: status,
        to: organizerEmail,
      })
    );
  } catch (e) {
    console.warn(`Calendar RSVP: could not build the reply: ${e.message}`);
  }
}

/**
 * Counter-proposes a different time for a meeting we were invited to, from the calendar.
 *
 * Only the organizer may revise a meeting (RFC 5546 section 2.1.4), so an attendee who wants
 * a different slot asks for one instead - iTIP's COUNTER method, section 3.2.7. The message
 * header offers this on the invitation email, but our own accept flow puts the invitation on
 * the calendar, and by the time someone wants to move it the email is long buried. The
 * calendar is where they look.
 *
 * Nothing is written locally: a counter is a request, and the event does not change unless
 * the organizer accepts and sends the update back.
 */
export async function proposeNewTimeForCalendarEvent(
  occurrence: EventOccurrence,
  proposal: { start: Date; end: Date; comment: string }
): Promise<void> {
  const eventId = parseEventIdFromOccurrence(occurrence.id);
  const event = await DatabaseStore.find<Event>(Event, eventId);
  if (!event) {
    console.warn(`Calendar counter-proposal: could not find event ${eventId}`);
    return;
  }

  let parsed: ReturnType<typeof CalendarUtils.parseICSString>;
  try {
    parsed = CalendarUtils.parseICSString(event.ics);
  } catch (e) {
    AppEnv.showErrorDialog(localized("Sorry, this event's data could not be read."));
    return;
  }

  const me = CalendarUtils.selfParticipant(parsed.event, event.accountId);
  if (!me || !me.email) {
    AppEnv.showErrorDialog(
      localized("You're not on this event's guest list, so there's no new time to propose.")
    );
    return;
  }

  const organizerEmail = CalendarUtils.emailFromParticipantURI(parsed.event.organizer);
  if (!organizerEmail) {
    AppEnv.showErrorDialog(
      localized("This event has no organizer, so there's nobody to propose a new time to.")
    );
    return;
  }

  let ics: string;
  try {
    ics = ICSEventHelpers.createCounterProposal(event.ics, {
      email: me.email,
      name: me.component.getParameter('cn') as string,
      start: proposal.start,
      end: proposal.end,
      comment: proposal.comment,
    });
  } catch (e) {
    console.warn(`Calendar counter-proposal: could not build it: ${e.message}`);
    ics = null;
  }
  if (!ics) {
    AppEnv.showErrorDialog(
      localized("Sorry, we couldn't build a counter-proposal for this event.")
    );
    return;
  }

  Actions.queueTask(
    EventRSVPTask.forProposingNewTime({
      accountId: event.accountId,
      to: organizerEmail,
      ics,
      summary: parsed.event.summary,
      comment: proposal.comment,
    })
  );
}

/**
 * Opens the time picker for a counter-proposal, anchored on the event in the grid.
 *
 * Shared by the context menu and the read-only card, so the two offer the same gesture and
 * cannot drift apart. Anchoring falls back to the centre of the window when the event is
 * scrolled out of view, which happens when the card was opened and then the grid moved.
 */
export function openProposeNewTimePopover(occurrence: EventOccurrence): void {
  const eventEl = document.getElementById(occurrence.id);
  Actions.openPopover(
    React.createElement(ProposeTimePopover, {
      start: occurrenceStartUnix(occurrence),
      end: occurrenceEndUnix(occurrence),
      onPropose: (proposal: { start: Date; end: Date; comment: string }) =>
        proposeNewTimeForCalendarEvent(occurrence, proposal),
    }),
    {
      originRect: eventEl
        ? eventEl.getBoundingClientRect()
        : new DOMRect(window.innerWidth / 2, window.innerHeight / 2, 2, 2),
      direction: 'right',
      fallbackDirection: 'left',
    }
  );
}

/** Alias used by the read-only popover, which closes itself before handing over. */
export const proposeNewTimeFromPopover = openProposeNewTimePopover;
