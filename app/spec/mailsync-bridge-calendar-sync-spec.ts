import MailsyncBridge from '../src/flux/mailsync-bridge';

/*
A calendar write reaches one account's server, so only that account is worth re-reading.
Polling them all puts the write's own re-read behind every other account's CalDAV
round-trip, which is felt as a deleted event lingering on screen.
*/
describe('MailsyncBridge calendar sync', function () {
  let bridge;
  let sent: { [accountId: string]: number };

  beforeEach(function () {
    spyOn(AppEnv, 'isMainWindow').andReturn(true);
    sent = { a: 0, b: 0 };
    bridge = new MailsyncBridge();
    bridge._clients = {
      a: { sendMessage: () => (sent.a += 1) },
      b: { sendMessage: () => (sent.b += 1) },
    };
  });

  it('polls only the account named', function () {
    bridge.sendSyncCalendarNow('a');
    expect(sent).toEqual({ a: 1, b: 0 });
  });

  it('polls every account when none is named, for the Refresh Calendars menu item', function () {
    bridge.sendSyncCalendarNow();
    expect(sent).toEqual({ a: 1, b: 1 });
  });

  it('polls nothing for an account with no sync process', function () {
    bridge.sendSyncCalendarNow('gone');
    expect(sent).toEqual({ a: 0, b: 0 });
  });
});
