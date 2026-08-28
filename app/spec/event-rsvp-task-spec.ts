import { AccountStore, EventRSVPTask } from 'mailspring-exports';

const ACCOUNT_ID = 'acct-1';
const ME = 'brian@example.com';

/**
 * An invitation to a series whose second occurrence was moved, so the calendar carries a
 * master VEVENT and an inline exception - both with the full guest list.
 */
const SERIES_INVITE = [
  'BEGIN:VCALENDAR',
  'VERSION:2.0',
  'PRODID:-//Test//Test//EN',
  'METHOD:REQUEST',
  'BEGIN:VEVENT',
  'UID:series@example.com',
  'DTSTAMP:20260101T000000Z',
  'DTSTART:20260301T140000Z',
  'DTEND:20260301T150000Z',
  'RRULE:FREQ=DAILY;COUNT=5',
  'SUMMARY:Standup',
  'ORGANIZER:mailto:ada@example.com',
  'ATTENDEE;CN=Ada;PARTSTAT=ACCEPTED:mailto:ada@example.com',
  `ATTENDEE;CN=Brian;RSVP=TRUE;PARTSTAT=NEEDS-ACTION:mailto:${ME}`,
  'ATTENDEE;CN=Carol;PARTSTAT=DECLINED:mailto:carol@example.com',
  'END:VEVENT',
  'BEGIN:VEVENT',
  'UID:series@example.com',
  'RECURRENCE-ID:20260302T140000Z',
  'DTSTAMP:20260101T000000Z',
  'DTSTART:20260302T160000Z',
  'DTEND:20260302T170000Z',
  'SUMMARY:Standup (moved)',
  'ORGANIZER:mailto:ada@example.com',
  'ATTENDEE;CN=Ada;PARTSTAT=ACCEPTED:mailto:ada@example.com',
  `ATTENDEE;CN=Brian;RSVP=TRUE;PARTSTAT=NEEDS-ACTION:mailto:${ME}`,
  'ATTENDEE;CN=Carol;PARTSTAT=DECLINED:mailto:carol@example.com',
  'END:VEVENT',
  'END:VCALENDAR',
].join('\r\n');

/** The ATTENDEE lines of each VEVENT, in document order. */
function attendeesPerVevent(ics: string): string[][] {
  const unfolded = ics.replace(/\r\n[ \t]/g, '');
  return unfolded
    .split('BEGIN:VEVENT')
    .slice(1)
    .map((block) =>
      block
        .split(/\r?\n/)
        .filter((l) => l.startsWith('ATTENDEE'))
        .map((l) => l.trim())
    );
}

describe('EventRSVPTask.forReplying', function () {
  beforeEach(function () {
    spyOn(AccountStore, 'accountForEmail').andCallFake((email: string) =>
      email && email.toLowerCase() === ME ? ({ id: ACCOUNT_ID } as any) : null
    );
  });

  const reply = (status: any = 'ACCEPTED') =>
    EventRSVPTask.forReplying({
      accountId: ACCOUNT_ID,
      to: 'ada@example.com',
      icsOriginalData: SERIES_INVITE,
      icsRSVPStatus: status,
    });

  it('leaves exactly one ATTENDEE in every VEVENT, not just the master', function () {
    // RFC 5546 section 3.2.3. A series invitation carries its modified occurrences as
    // further VEVENTs; leaving their guest lists intact ships the organizer a REPLY that
    // also purports to speak for everyone else on the invitation.
    const perVevent = attendeesPerVevent(reply().ics);
    expect(perVevent.length).toBe(2);
    for (const attendees of perVevent) {
      expect(attendees.length).toBe(1);
      expect(attendees[0]).toContain(ME);
    }
  });

  it('records the answer in every VEVENT it kept', function () {
    for (const attendees of attendeesPerVevent(reply('DECLINED').ics)) {
      expect(attendees[0]).toContain('PARTSTAT=DECLINED');
    }
  });

  it('drops RSVP=TRUE, since the answer has now been given', function () {
    for (const attendees of attendeesPerVevent(reply().ics)) {
      expect(attendees[0]).not.toContain('RSVP=TRUE');
    }
  });

  it('keeps the organizer, which is what the reply is matched against', function () {
    expect(reply().ics).toContain('ORGANIZER:mailto:ada@example.com');
  });

  it('declares itself a REPLY', function () {
    const task = reply();
    expect(task.ics).toContain('METHOD:REPLY');
    expect(task.method).toBe('REPLY');
  });

  it('refuses to reply on behalf of an account that is not a guest', function () {
    expect(() =>
      EventRSVPTask.forReplying({
        accountId: 'acct-nobody',
        to: 'ada@example.com',
        icsOriginalData: SERIES_INVITE,
        icsRSVPStatus: 'ACCEPTED' as any,
      })
    ).toThrow();
  });
});
