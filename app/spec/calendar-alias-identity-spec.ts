import { AccountStore, Account, Contact, CalendarUtils, EventRSVPTask } from 'mailspring-exports';
import { Event as MailspringEvent } from '../src/flux/models/event';
import { occurrencesForEvents } from '../internal_packages/main-calendar/lib/core/calendar-data-source';

/*
Identity when the address on the invitation is one of the account's aliases.

Every editing decision in the calendar routes through isMine, and every RSVP routes through
selfParticipant; both resolve identity with AccountStore.accountForEmail. An organizer who
is the user under an alias must come back as the user, or the account loses the ability to
edit its own meetings and a reply is written against nobody.

These stub only the store's data - accounts() and aliases() - so accountForEmail itself, and
the Utils.emailIsEquivalent comparison inside it, are the code under test.
*/

const ACCOUNT_ID = 'acct-alias-1';
const PRIMARY = 'primary@example.com';
const ALIAS = 'the.alias@example.com';
const ALIAS_SPEC = `Alias Name <${ALIAS}>`;
const STRANGER = 'someone.else@example.org';

function icsWith({ organizer, attendees }: { organizer: string; attendees: string[] }): string {
  return [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Test//Test//EN',
    'METHOD:REQUEST',
    'BEGIN:VEVENT',
    'UID:alias-uid@test',
    'DTSTART:20260622T150000Z',
    'DTEND:20260622T160000Z',
    `ORGANIZER;CN=Organizer:mailto:${organizer}`,
    ...attendees.map((a) => `ATTENDEE;PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:${a}`),
    'SUMMARY:Alias Identity Test',
    'DTSTAMP:20260101T000000Z',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');
}

function occurrenceFor(ics: string) {
  const event = new MailspringEvent({
    id: 'event-alias-1',
    accountId: ACCOUNT_ID,
    calendarId: 'calendar-1',
    icsuid: 'alias-uid@test',
    ics,
  } as any);
  const [occ] = occurrencesForEvents([event], {
    startUnix: new Date(2026, 0, 1).getTime() / 1000,
    endUnix: new Date(2027, 0, 1).getTime() / 1000,
  });
  return occ;
}

describe('calendar identity through an account alias', function () {
  beforeEach(function () {
    const account = new Account({
      id: ACCOUNT_ID,
      emailAddress: PRIMARY,
      aliases: [ALIAS_SPEC],
      provider: 'gmail',
    } as any);

    const aliasContact = account.meUsingAlias(ALIAS_SPEC);
    (aliasContact as any).isAlias = true;

    spyOn(AccountStore, 'accounts').andReturn([account]);
    spyOn(AccountStore, 'aliases').andReturn([account.me(), aliasContact] as any);
    spyOn(AccountStore, 'accountForId').andCallFake((id: string) =>
      id === ACCOUNT_ID ? (account as any) : null
    );
  });

  describe('AccountStore.accountForEmail', function () {
    it('resolves the primary address', function () {
      expect(AccountStore.accountForEmail(PRIMARY)?.id).toBe(ACCOUNT_ID);
    });

    it('resolves an alias to the account that owns it', function () {
      // The alias path had never been exercised: this account has no aliases configured in
      // real use, so nothing proved accountForEmail walked past the primary address.
      expect(AccountStore.accountForEmail(ALIAS)?.id).toBe(ACCOUNT_ID);
    });

    it('resolves an alias regardless of the case the organizer wrote it in', function () {
      expect(AccountStore.accountForEmail(ALIAS.toUpperCase())?.id).toBe(ACCOUNT_ID);
    });

    it('does not claim an address that is neither the account nor an alias', function () {
      expect(AccountStore.accountForEmail(STRANGER)).toBe(null);
    });
  });

  describe('Contact.isMe', function () {
    it('is true for an alias', function () {
      expect(new Contact({ email: ALIAS }).isMe()).toBe(true);
    });

    it('is false for a stranger', function () {
      expect(new Contact({ email: STRANGER }).isMe()).toBe(false);
    });
  });

  describe('isMine, which gates editing', function () {
    it('is true when the organizer is the account itself', function () {
      const occ = occurrenceFor(icsWith({ organizer: PRIMARY, attendees: [STRANGER] }));
      expect(occ.isMine).toBe(true);
    });

    it('is true when the organizer is one of the account aliases', function () {
      // The case that would otherwise silently make the user a guest at their own meeting:
      // isMine false means the editor, the context menu and dragging all refuse.
      const occ = occurrenceFor(icsWith({ organizer: ALIAS, attendees: [STRANGER] }));
      expect(occ.isMine).toBe(true);
    });

    it('is false when someone else organised it', function () {
      const occ = occurrenceFor(icsWith({ organizer: STRANGER, attendees: [PRIMARY] }));
      expect(occ.isMine).toBe(false);
    });

    it('is true when the event names no organizer at all', function () {
      const occ = occurrenceFor(
        [
          'BEGIN:VCALENDAR',
          'VERSION:2.0',
          'PRODID:-//Test//Test//EN',
          'BEGIN:VEVENT',
          'UID:alias-uid@test',
          'DTSTART:20260622T150000Z',
          'DTEND:20260622T160000Z',
          'SUMMARY:No organizer',
          'DTSTAMP:20260101T000000Z',
          'END:VEVENT',
          'END:VCALENDAR',
        ].join('\r\n')
      );
      expect(occ.isMine).toBe(true);
    });
  });

  describe('isAwaitingGuests, which only applies to meetings we organise', function () {
    it('is set when we organise under an alias and a guest has not answered', function () {
      const occ = occurrenceFor(icsWith({ organizer: ALIAS, attendees: [ALIAS, STRANGER] }));
      expect(occ.isAwaitingGuests).toBe(true);
    });

    it('is not set for a meeting organised by someone else', function () {
      const occ = occurrenceFor(icsWith({ organizer: STRANGER, attendees: [PRIMARY, STRANGER] }));
      expect(occ.isAwaitingGuests).toBe(false);
    });
  });

  describe('the RSVP reply', function () {
    const icsInvite = icsWith({ organizer: STRANGER, attendees: [ALIAS, 'third@example.org'] });

    it('finds the alias attendee as self', function () {
      const ICAL = require('ical.js');
      const root = new ICAL.Component(ICAL.parse(icsInvite));
      const icsEvent = new ICAL.Event(root.getFirstSubcomponent('vevent'));
      expect(CalendarUtils.selfParticipant(icsEvent, ACCOUNT_ID)?.email).toBe(ALIAS);
    });

    it('replies as the alias and speaks for nobody else', function () {
      // A REPLY names exactly one ATTENDEE (RFC 5546 section 3.2.3). If the alias were not
      // recognised, no attendee would match and every ATTENDEE line would be stripped,
      // leaving a reply that identifies no one.
      const task = EventRSVPTask.forReplying({
        accountId: ACCOUNT_ID,
        to: STRANGER,
        icsOriginalData: icsInvite,
        icsRSVPStatus: 'ACCEPTED',
      });
      const attendeeLines = task.ics
        .replace(/\r\n[ \t]/g, '')
        .split(/\r?\n/)
        .filter((l) => l.startsWith('ATTENDEE'));

      expect(attendeeLines.length).toBe(1);
      expect(attendeeLines[0]).toContain(ALIAS);
      expect(attendeeLines[0]).toContain('PARTSTAT=ACCEPTED');
      expect(task.ics).toContain('METHOD:REPLY');
    });
  });
});
