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
import { EventOccurrence } from './calendar-data-source';
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
