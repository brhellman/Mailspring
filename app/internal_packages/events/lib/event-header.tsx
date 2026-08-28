import { RetinaImg } from 'mailspring-component-kit';

import React from 'react';
import fs from 'fs';
import {
  Rx,
  Actions,
  AttachmentStore,
  Account,
  AccountStore,
  Calendar,
  File,
  localized,
  DateUtils,
  CalendarUtils,
  ICSEventHelpers,
  ICSParticipantStatus,
  Message,
  Event,
  EventRSVPTask,
  SyncbackEventTask,
  DatabaseStore,
  RegExpUtils,
} from 'mailspring-exports';
import ICAL from 'ical.js';
import {
  resolveRSVPTarget,
  resolveDefaultCalendar,
  conflictCalendarIds,
  mayBeAddedToCalendar,
  RSVPTargetResolution,
} from './rsvp-target';
import { findConflicts, CalendarConflict } from '../../../src/calendar-conflicts';
import { ProposeTimePopover, formatProposedTime } from './propose-time-popover';
import { findOneIana } from 'windows-iana';

const moment = require('moment-timezone');

const TEL_URI = /tel:\S+?(?=[.,;:]*(?:\s|$))/gi;

/**
 * A LOCATION is usually the video-call URL, most often followed by the rooms booked for it, so
 * each URL or tel: URI in it is linked and the rest stays text.
 */
export function renderLocation(location: string | undefined): React.ReactNode {
  if (!location) return null;

  const links: Array<{ start: number; end: number; href: string }> = [];
  for (const pattern of [RegExpUtils.urlRegex(), TEL_URI]) {
    for (const match of location.matchAll(pattern)) {
      const start = match.index;
      const end = start + match[0].length;
      if (links.some((l) => start < l.end && end > l.start)) continue;
      links.push({ start, end, href: match[0] });
    }
  }
  if (!links.length) return location;

  const nodes: React.ReactNode[] = [];
  let cursor = 0;
  for (const { start, end, href } of links.sort((a, b) => a.start - b.start)) {
    if (start > cursor) nodes.push(location.slice(cursor, start));
    nodes.push(
      <a key={start} href={href}>
        {location.slice(start, end)}
      </a>
    );
    cursor = end;
  }
  if (cursor < location.length) nodes.push(location.slice(cursor));
  return nodes;
}

interface EventHeaderProps {
  message: Message;
  file: File;
}

interface EventHeaderState {
  icsOriginalData?: string;
  /**
   * The invitation exactly as it arrived by email, never replaced by a synced copy.
   *
   * Everything we send back to the organizer is built from this. Google rewrites ORGANIZER
   * on its own copy of an event that lives on a shared calendar, to that calendar's
   * `...@group.calendar.google.com` id - an address no person reads - so a reply addressed
   * from the synced copy is delivered nowhere. RFC 5546 section 3.2.3 is explicit that a
   * REPLY goes to the ORGANIZER of the REQUEST being answered, which is this.
   */
  inviteIcs?: string;
  inviteEvent?: ICAL.Event;
  icsMethod?: 'reply' | 'request' | 'cancel' | 'counter';
  icsEvent?: ICAL.Event;
  /** Which calendar copy of this event, if any, our response will be written to. */
  rsvp?: RSVPTargetResolution;
  /** Everything already on our calendars that overlaps this invitation. */
  conflicts?: CalendarConflict[];
  /** Where the invitation would be added if we accept and it isn't on a calendar yet. */
  addTo?: Calendar;
  /** The calendars that could receive it, so the choice can be changed before answering. */
  addToChoices?: Calendar[];
  /** The slot we last counter-proposed, kept so the row can confirm what was sent. */
  proposed?: { start: number; end: number };
  inflight?: ICSParticipantStatus;
}

/*
The EventHeader allows you to RSVP to a calendar invite embedded in an email. It also
looks to see if a matching event is present on your calendar. In most cases the event
will also be on your calendar, and that version is synced while the email attachment
version gets stale.

We try to show the RSVP status of the event on your calendar if it's present. If not,
we fall back to storing the RSVP status in message metadata (so the "Accept" button is
"sticky", even though we just fire off a RSVP message via email and never hear back.)
*/
export class EventHeader extends React.Component<EventHeaderProps, EventHeaderState> {
  static displayName = 'EventHeader';

  state: EventHeaderState = {
    icsEvent: undefined,
    icsMethod: undefined,
    icsOriginalData: undefined,
    inviteIcs: undefined,
    inviteEvent: undefined,
    rsvp: undefined,
    conflicts: undefined,
    addTo: undefined,
    addToChoices: undefined,
    proposed: undefined,
    inflight: undefined,
  };

  _mounted = false;
  _subscription: Rx.IDisposable;

  componentWillUnmount() {
    this._mounted = false;
    if (this._subscription) {
      this._subscription.dispose();
    }
  }

  componentDidMount() {
    this._mounted = true;
    this._loadICSAttachment();
  }

  componentDidUpdate(prevProps: EventHeaderProps, prevState: EventHeaderState) {
    if (prevState.inflight) {
      this.setState({ inflight: undefined });
    }
    // The attachment lands on disk when mailsync downloads the message body, which can
    // happen after we first render. Retry as long as we have nothing to show, so an
    // invite that arrives while it's open doesn't stay actionless until it's reopened.
    if (!this.state.icsOriginalData && prevProps.message !== this.props.message) {
      this._loadICSAttachment();
    }
  }

  async _loadICSAttachment() {
    const { file, message } = this.props;

    let data: string;
    try {
      // Resolve through the store rather than pathForFile(): mailsync sometimes saves an
      // attachment under a sanitized name, and only the store knows how to find it again.
      const filePath = await AttachmentStore.resolvePathForFile(file);
      data = await fs.promises.readFile(filePath, 'utf8');
    } catch (e) {
      return; // not downloaded yet - componentDidUpdate retries
    }
    if (!this._mounted) return;

    let parsed: ReturnType<typeof CalendarUtils.parseICSString>;
    try {
      parsed = CalendarUtils.parseICSString(data);
    } catch (e) {
      console.warn(
        `EventHeader: Could not parse ICS data from attachment ${file.filename}: ${e.message}`
      );
      return;
    }
    const { event, root } = parsed;

    const method = root.getFirstPropertyValue('method');
    const methodLower = (typeof method === 'string' ? method : 'request').toLowerCase();
    // Normalize to the methods we render. Anything else is treated as an invitation, which
    // is the only method that offers actions the user could get wrong.
    const normalizedMethod = ['reply', 'cancel', 'counter'].includes(methodLower)
      ? methodLower
      : 'request';
    this.setState({
      icsEvent: event,
      icsMethod: normalizedMethod as EventHeaderState['icsMethod'],
      icsOriginalData: data,
      inviteIcs: data,
      inviteEvent: event,
    });

    if (this._subscription) {
      this._subscription.dispose();
    }
    // Bounding by the stored series start/end keeps the query to a handful of rows;
    // findConflicts expands those to the occurrences that actually overlap.
    const windowStart = Math.round(event.startDate.toJSDate().getTime() / 1000);
    const windowEnd = Math.round(event.endDate.toJSDate().getTime() / 1000);

    this._subscription = Rx.Observable.combineLatest(
      // Every calendar copy of this UID: one meeting can sit on our own calendar, a room's
      // and a colleague's at once, and only resolveRSVPTarget can say which is ours.
      Rx.Observable.fromQuery(
        DatabaseStore.findAll<Event>(Event).where({
          icsuid: event.uid,
          accountId: message.accountId,
        })
      ),
      Rx.Observable.fromQuery(
        DatabaseStore.findAll<Calendar>(Calendar).where({ accountId: message.accountId })
      ),
      Rx.Observable.fromQuery(
        DatabaseStore.findAll<Event>(Event).where([
          Event.attributes.accountId.equal(message.accountId),
          Event.attributes.recurrenceStart.lessThan(windowEnd),
          Event.attributes.recurrenceEnd.greaterThan(windowStart),
        ])
      ),
      (calEvents: Event[], calendars: Calendar[], nearby: Event[]) => ({
        calEvents,
        calendars,
        nearby,
      })
    ).subscribe(({ calEvents, calendars, nearby }) => {
      if (!this._mounted) return;

      const addresses = this._accountAddresses();
      const rsvp = resolveRSVPTarget({ events: calEvents, calendars, addresses });
      // When the invitation isn't on a calendar yet, accepting has to create it somewhere.
      // Offer every calendar it could go on and preselect our best guess: a guess the user
      // can see and change is safe in a way a silent one isn't. Once they've chosen, keep
      // their choice rather than recomputing it on the next sync tick.
      let addTo: Calendar;
      let addToChoices: Calendar[];
      if (
        rsvp.problem === 'not-on-a-calendar' &&
        mayBeAddedToCalendar(event.organizer, addresses)
      ) {
        addToChoices = calendars.filter((c) => !c.readOnly);
        addTo = this.state.addTo || resolveDefaultCalendar(calendars, addresses) || addToChoices[0];
      }

      const busyCalendarIds = new Set(
        conflictCalendarIds(calendars, AppEnv.config.get('mailspring.disabledCalendars') || [])
      );
      const conflicts = findConflicts({
        events: nearby.filter((e) => busyCalendarIds.has(e.calendarId)),
        start: windowStart,
        end: windowEnd,
        addresses,
        excludeIcsuid: event.uid,
      });

      // Prefer the synced copy over the emailed attachment - the attachment goes stale as
      // soon as the organizer changes anything.
      const display =
        normalizedMethod === 'counter' ? null : rsvp.target ? rsvp.target.event : calEvents[0];
      if (!display) {
        this.setState({ rsvp, conflicts, addTo, addToChoices });
        return;
      }
      try {
        this.setState({
          icsEvent: CalendarUtils.parseICSString(display.ics).event,
          icsOriginalData: display.ics,
          rsvp,
          conflicts,
          addTo,
          addToChoices,
        });
      } catch (e) {
        console.warn(`EventHeader: Could not parse ICS data from calendar event: ${e.message}`);
        this.setState({ rsvp, conflicts, addTo, addToChoices });
      }
    });
  }

  /** This account's own address plus any aliases, used to recognise our own calendar. */
  _accountAddresses(): string[] {
    const account: Account = AccountStore.accountForId(this.props.message.accountId);
    if (!account) return [];
    return [
      account.emailAddress,
      ...AccountStore.aliases()
        .filter((a) => a.accountId === account.id)
        .map((a) => a.email),
    ];
  }

  render() {
    const { icsEvent, icsMethod } = this.state;
    if (!icsEvent || !icsEvent.startDate) {
      return null;
    }

    // Workaround to convert calendar invites sent out from Microsoft calendars to IANA timezones
    // that can be handled by moments-timezone.
    let startTimezone = findOneIana(icsEvent.startDate.zone.tzid) || icsEvent.startDate.zone.tzid;
    let endTimezone = findOneIana(icsEvent.endDate.zone.tzid) || icsEvent.endDate.zone.tzid;
    // Workaround to convert calendar invites sent out from Google calendar with "Z" timezone
    // to IANA timezone that can be handled by moments-timezone.
    if (startTimezone === 'Z') {
      startTimezone = 'UTC';
    }
    if (endTimezone === 'Z') {
      endTimezone = 'UTC';
    }

    const startMoment = moment
      .tz(icsEvent.startDate.toString(), startTimezone)
      .tz(DateUtils.timeZone);
    const endMoment = moment.tz(icsEvent.endDate.toString(), endTimezone).tz(DateUtils.timeZone);

    const daySeconds = 24 * 60 * 60 * 1000;
    let day = '';
    let time = '';

    if (endMoment.diff(startMoment) < daySeconds) {
      day = startMoment.format('dddd, MMMM Do');
      time = `${startMoment.format(
        DateUtils.getTimeFormat({ timeZone: false })
      )} - ${endMoment.format(DateUtils.getTimeFormat({ timeZone: true }))}`;
    } else {
      day = `${startMoment.format('dddd, MMMM Do')} - ${endMoment.format('MMMM Do')}`;
      if (endMoment.diff(startMoment) % daySeconds === 0) {
        time = localized('All Day');
      } else {
        time = startMoment.format(DateUtils.getTimeFormat({ timeZone: true }));
      }
    }

    return (
      <div className="event-wrapper">
        <div className="event-header">
          <div className="event-download" onClick={() => Actions.fetchAndOpenFile(this.props.file)}>
            <RetinaImg name="icon-attachment-download.png" mode={RetinaImg.Mode.ContentIsMask} />
          </div>
          <RetinaImg name="icon-RSVP-calendar-mini@2x.png" mode={RetinaImg.Mode.ContentPreserve} />
          <span className="event-title-text">{localized('Event')}: </span>
          <span className="event-title">{icsEvent.summary}</span>
        </div>
        <div className="event-body">
          <div className="event-date">
            <div className="event-day">{day}</div>
            <div>
              <div className="event-time">{time}</div>
            </div>
            <div className="event-location">{renderLocation(icsEvent.location)}</div>
            {icsMethod !== 'cancel' && this._renderConflicts()}
            {icsMethod === 'cancel'
              ? this._renderCancellation()
              : icsMethod === 'counter'
                ? this._renderCounterProposal()
                : icsMethod === 'request'
                  ? this._renderRSVP()
                  : this._renderSenderResponse()}
          </div>
        </div>
      </div>
    );
  }

  /*
  Warn about anything already booked over this invitation, the way Google Calendar does, so
  the answer can be given from the email without opening the calendar to check.
  */
  _renderConflicts() {
    const { conflicts } = this.state;
    if (!conflicts || !conflicts.length) return false;

    const label = (conflict: CalendarConflict) => {
      const start = moment.unix(conflict.start).tz(DateUtils.timeZone);
      const end = moment.unix(conflict.end).tz(DateUtils.timeZone);
      const time = `${start.format(
        DateUtils.getTimeFormat({ timeZone: false })
      )} - ${end.format(DateUtils.getTimeFormat({ timeZone: false }))}`;
      return `${conflict.title || localized('(No title)')} (${time})`;
    };

    return (
      <div className="event-conflicts">
        <div className="event-conflicts-title">
          {conflicts.length === 1
            ? localized('Conflicts with an event on your calendar')
            : localized('Conflicts with %@ events on your calendar', conflicts.length)}
        </div>
        {conflicts.map((conflict) => (
          <div className="event-conflict" key={`${conflict.eventId}-${conflict.start}`}>
            {label(conflict)}
          </div>
        ))}
      </div>
    );
  }

  _renderSenderResponse() {
    const { icsEvent } = this.state;
    const from = this.props.message.from[0];
    if (!from) return false;

    const sender = CalendarUtils.cleanParticipants(icsEvent).find((p) => p.email === from.email);
    if (!sender) return false;

    const verb: { [key: string]: string } = {
      DECLINED: localized('declined'),
      ACCEPTED: localized('accepted'),
      TENTATIVE: localized('tentatively accepted'),
      DELEGATED: localized('delegated'),
      COMPLETED: localized('completed'),
    }[sender.status];

    return (
      <div className="event-actions">{localized(`%1$@ has %2$@ this event`, from.email, verb)}</div>
    );
  }

  /*
  A COUNTER is an attendee proposing a different slot for a meeting we organize (RFC 5546
  section 3.2.7), so the actions are the organizer's, not an attendee's - accepting means
  moving the event. The header above already shows the proposed time, because the attachment
  is the only place it exists; our own copy still has the original.
  */
  _renderCounterProposal() {
    const { rsvp } = this.state;
    const from = this.props.message.from[0];
    const proposer = from ? from.displayName() : localized('An attendee');

    if (!rsvp || !rsvp.target) {
      return (
        <div className="event-actions event-counter">
          <div className="event-counter-notice">
            {localized('%@ proposed this new time.', proposer)}
          </div>
          <div className="event-rsvp-destination event-rsvp-email-only">
            {localized("This event isn't on a calendar you can edit, so it can't be moved here.")}
          </div>
        </div>
      );
    }

    return (
      <div className="event-actions event-counter">
        <div className="event-counter-notice">
          {localized('%@ proposed this new time.', proposer)}
        </div>
        <div className="event-rsvp-buttons">
          <div className="btn btn-large" onClick={this._onAcceptProposedTime}>
            {localized('Move event to this time')}
          </div>
        </div>
        <div className="event-rsvp-destination">
          {localized('Your response will be saved to %@', rsvp.target.calendar.name)}
        </div>
      </div>
    );
  }

  _onAcceptProposedTime = () => {
    const { rsvp, icsEvent } = this.state;
    if (!rsvp || !rsvp.target) return;

    const start = Math.round(icsEvent.startDate.toJSDate().getTime() / 1000);
    const end = Math.round(icsEvent.endDate.toJSDate().getTime() / 1000);

    const calEvent = rsvp.target.event;
    let ics: string;
    try {
      ics = ICSEventHelpers.bumpEventSequence(
        ICSEventHelpers.updateEventTimes(calEvent.ics, { start, end, isAllDay: false })
      );
    } catch (e) {
      console.warn(`EventHeader: Could not apply the proposed time: ${e.message}`);
      AppEnv.showErrorDialog(localized("Sorry, we couldn't move this event to the proposed time."));
      return;
    }

    const event = calEvent.clone();
    event.ics = ics;
    event.recurrenceStart = start;
    event.recurrenceEnd = end;

    // The server re-sends the invitation to the guests when the organizer moves the event,
    // so there is no separate REQUEST for us to send.
    Actions.queueTask(
      SyncbackEventTask.forUpdating({
        event,
        undoData: {
          ics: calEvent.ics,
          recurrenceStart: calEvent.recurrenceStart,
          recurrenceEnd: calEvent.recurrenceEnd,
        },
        description: localized('Move event'),
      })
    );
  };

  _renderCancellation() {
    const { icsEvent } = this.state;
    const organizerEmail = CalendarUtils.emailFromParticipantURI(icsEvent.organizer);

    return (
      <div className="event-actions event-cancelled">
        <span className="cancelled-notice">
          {organizerEmail
            ? localized('This event has been cancelled by %@', organizerEmail)
            : localized('This event has been cancelled')}
        </span>
      </div>
    );
  }

  _renderRSVP() {
    const { icsEvent, inflight } = this.state;
    const me = CalendarUtils.selfParticipant(icsEvent, this.props.message.accountId);
    if (!me) {
      // Invitations addressed to a group or a distribution list name the group as the
      // attendee, not us, and iTIP gives us no standing to reply on the group's behalf.
      return (
        <div className="event-actions event-no-rsvp">
          {localized(
            "This invitation was sent to an address that isn't listed as a guest, so there's no RSVP to give."
          )}
        </div>
      );
    }

    let status = me.status;

    const icsTimeProperty = icsEvent.component.getFirstPropertyValue('dtstamp') as ICAL.Time;
    const icsTime = icsTimeProperty ? icsTimeProperty.toJSDate() : new Date(0);

    const metadata = this.props.message.metadataForPluginId('event-rsvp');
    if (metadata && new Date(metadata.time) > icsTime) {
      status = metadata.status;
    }

    const actions: [ICSParticipantStatus, string][] = [
      ['ACCEPTED', localized('Accept')],
      ['TENTATIVE', localized('Maybe')],
      ['DECLINED', localized('Decline')],
    ];

    return (
      <div className="event-actions">
        <div className="event-rsvp-buttons">
          {actions.map(([actionStatus, actionLabel]) => (
            <div
              key={actionStatus}
              className={`btn btn-large btn-rsvp ${status === actionStatus ? actionStatus : ''}`}
              onClick={() => this._onRSVP(actionStatus)}
            >
              {actionStatus === status || actionStatus !== inflight ? (
                actionLabel
              ) : (
                <RetinaImg
                  width={18}
                  name="sending-spinner.gif"
                  mode={RetinaImg.Mode.ContentPreserve}
                />
              )}
            </div>
          ))}
        </div>
        {this._renderProposeNewTime()}
        {this._renderRSVPDestination()}
      </div>
    );
  }

  /*
  Offer the organizer a different slot instead of just accepting or declining - iTIP's
  COUNTER method (RFC 5546 section 3.2.7), which Google Calendar calls "Propose a new time".
  Only the organizer can act on it, so this sends and then says what was sent; there is no
  state on our side to keep until they answer.
  */
  _renderProposeNewTime() {
    const { icsEvent, proposed } = this.state;
    if (proposed) {
      return (
        <div className="event-proposed-time">
          {localized('You proposed %@', formatProposedTime(proposed.start, proposed.end))}
        </div>
      );
    }
    if (!icsEvent.startDate || !icsEvent.endDate) return false;

    return (
      <div
        className="btn btn-link btn-propose-time"
        onClick={(e) => this._onOpenProposeTime(e.currentTarget as HTMLElement)}
      >
        {localized('Propose a new time')}
      </div>
    );
  }

  _onOpenProposeTime(originElement: HTMLElement) {
    const { icsEvent } = this.state;
    const start = Math.round(icsEvent.startDate.toJSDate().getTime() / 1000);
    const end = Math.round(icsEvent.endDate.toJSDate().getTime() / 1000);

    Actions.openPopover(
      <ProposeTimePopover start={start} end={end} onPropose={this._onProposeNewTime} />,
      { originRect: originElement.getBoundingClientRect(), direction: 'up' }
    );
  }

  _onProposeNewTime = ({ start, end, comment }: { start: Date; end: Date; comment: string }) => {
    const { inviteEvent, inviteIcs } = this.state;
    const { message } = this.props;

    const organizerEmail = CalendarUtils.emailFromParticipantURI(inviteEvent.organizer);
    if (!organizerEmail) {
      AppEnv.showErrorDialog(
        localized("This event has no organizer, so there's nobody to propose a new time to.")
      );
      return;
    }

    const me = CalendarUtils.selfParticipant(inviteEvent, message.accountId);
    if (!me || !me.email) return;

    let ics: string;
    try {
      ics = ICSEventHelpers.createCounterProposal(inviteIcs, {
        email: me.email,
        name: me.component.getParameter('cn') as string,
        start,
        end,
        comment,
      });
    } catch (e) {
      console.warn(`EventHeader: Could not build counter-proposal: ${e.message}`);
      ics = null;
    }
    if (!ics) {
      AppEnv.showErrorDialog(
        localized("Sorry, we couldn't build a counter-proposal for this invitation.")
      );
      return;
    }

    Actions.queueTask(
      EventRSVPTask.forProposingNewTime({
        accountId: message.accountId,
        messageId: message.id,
        to: organizerEmail,
        ics,
        summary: inviteEvent.summary,
        comment,
      })
    );

    this.setState({
      proposed: {
        start: Math.round(start.getTime() / 1000),
        end: Math.round(end.getTime() / 1000),
      },
    });
  };

  /*
  Say where the response is going before the user commits to it. Answering an invitation
  does two separate things - it emails the organizer, and it records the answer on our own
  copy of the event - and only the second one is tied to a particular calendar. When we
  can't identify which copy is ours we still send the email, so this has to distinguish
  "recorded on this calendar" from "emailed to the organizer and nothing else".
  */
  _renderRSVPDestination() {
    const { rsvp, addTo } = this.state;
    if (!rsvp) return false;

    if (rsvp.target) {
      return (
        <div className="event-rsvp-destination">
          {localized('Your response will be saved to %@', rsvp.target.calendar.name)}
        </div>
      );
    }

    if (rsvp.problem === 'not-on-a-calendar' && addTo) {
      const choices = this.state.addToChoices || [];
      return (
        <div className="event-rsvp-destination">
          <span>{localized('Accepting will add this event to')}</span>
          {choices.length > 1 ? (
            <select
              className="event-rsvp-calendar-picker"
              value={addTo.id}
              aria-label={localized('Calendar to add this event to')}
              onChange={(e) =>
                this.setState({ addTo: choices.find((c) => c.id === e.target.value) })
              }
            >
              {choices.map((c) => (
                <option key={c.id} value={c.id}>
                  {c.name}
                </option>
              ))}
            </select>
          ) : (
            <span>{addTo.name}</span>
          )}
        </div>
      );
    }

    const explanation = {
      'not-on-a-calendar': localized(
        "This event isn't on any of your calendars, so your response will only be emailed to the organizer."
      ),
      'read-only': localized(
        'This event is only on a read-only calendar, so your response will only be emailed to the organizer.'
      ),
      'not-ours': localized(
        'This event is only on a calendar shared with you by someone else, so your response will only be emailed to the organizer.'
      ),
      ambiguous: localized(
        "This event is on more than one of your calendars, so we can't tell which copy is yours. Your response will only be emailed to the organizer."
      ),
    }[rsvp.problem];

    return <div className="event-rsvp-destination event-rsvp-email-only">{explanation}</div>;
  }

  _onRSVP = (status: ICSParticipantStatus) => {
    const { inviteEvent, inviteIcs, inflight } = this.state;
    if (inflight) return; // prevent double clicks

    const organizerEmail = CalendarUtils.emailFromParticipantURI(inviteEvent.organizer);
    if (!organizerEmail) {
      AppEnv.showErrorDialog(
        localized(
          "Sorry, this event does not have an organizer or the organizer's address is not a valid email address: %@",
          inviteEvent.organizer || '(none)'
        )
      );
      return;
    }

    // EventRSVPTask.forReplying throws if it can't find us as an attendee in the data it's
    // replying with; catch that here instead of letting it crash the click handler.
    let task: EventRSVPTask;
    try {
      task = EventRSVPTask.forReplying({
        accountId: this.props.message.accountId,
        messageId: this.props.message.id,
        icsOriginalData: inviteIcs,
        icsRSVPStatus: status,
        to: organizerEmail,
      });
    } catch (e) {
      console.warn(`EventHeader: Could not build RSVP reply: ${e.message}`);
      AppEnv.showErrorDialog(
        localized(
          "Sorry, we couldn't find your email address in this event's attendee list, so an RSVP reply could not be sent."
        )
      );
      return;
    }

    this.setState({ inflight: status });
    Actions.queueTask(task);
    this._writeRSVPToCalendar(status);
  };

  /*
  Record the response on our own copy of the event as well as emailing the organizer.

  Under CalDAV an attendee answers an invitation by writing their own PARTSTAT back to
  their copy of the event (RFC 6638 section 3.2.5). The emailed REPLY tells the organizer,
  but on its own it leaves our calendar showing the invitation as unanswered - Google only
  reconciles the two when the organizer also happens to be on Google.
  */
  _writeRSVPToCalendar(status: ICSParticipantStatus) {
    const { rsvp, inviteEvent, addTo } = this.state;
    if (!rsvp) return;

    const me = CalendarUtils.selfParticipant(inviteEvent, this.props.message.accountId);
    if (!me || !me.email) return;

    // Deliberately not undoable, in both branches: undoing would have to retract the reply
    // we already emailed the organizer, which iTIP gives us no way to do.
    if (rsvp.target) {
      const { event: calEvent } = rsvp.target;
      const ics = ICSEventHelpers.updateAttendeeStatus(calEvent.ics, me.email, status);
      if (!ics || ics === calEvent.ics) return;

      const event = calEvent.clone();
      event.ics = ics;
      Actions.queueTask(SyncbackEventTask.forUpdating({ event }));
      return;
    }

    // Declining doesn't put the event on our calendar - the whole point of declining is that
    // we won't be there, and Google behaves the same way.
    const attending = status === 'ACCEPTED' || status === 'TENTATIVE';
    if (
      attending &&
      rsvp.problem === 'not-on-a-calendar' &&
      addTo &&
      mayBeAddedToCalendar(inviteEvent.organizer, this._accountAddresses())
    ) {
      this._addInvitationToCalendar(status, me.email, addTo);
    }
  }

  /*
  Put an accepted invitation on our calendar.

  Google only creates the attendee's copy of an event when the attendee answers inside
  Google's own interface. An emailed REPLY updates the organizer's copy and stops there, so
  an invitation answered from a CalDAV client would otherwise be accepted and yet appear on
  no calendar at all. Creating it here is what every calendar client does with an emailed
  invitation, and it carries our response with it so the two never disagree.
  */
  _addInvitationToCalendar(status: ICSParticipantStatus, myEmail: string, calendar: Calendar) {
    const { inviteIcs, inviteEvent } = this.state;

    let ics: string;
    try {
      const answered = ICSEventHelpers.updateAttendeeStatus(inviteIcs, myEmail, status);
      if (!answered) return;
      // The invitation is a scheduling message; a stored event must not carry its METHOD.
      ics = ICSEventHelpers.stripITIPMethod(answered);
    } catch (e) {
      console.warn(
        `EventHeader: Could not build a calendar event from the invitation: ${e.message}`
      );
      return;
    }

    const event = new Event({
      calendarId: calendar.id,
      accountId: this.props.message.accountId,
      ics,
      icsuid: inviteEvent.uid,
      recurrenceStart: Math.round(inviteEvent.startDate.toJSDate().getTime() / 1000),
      recurrenceEnd: Math.round(inviteEvent.endDate.toJSDate().getTime() / 1000),
    });

    Actions.queueTask(
      SyncbackEventTask.forCreating({
        event,
        calendarId: calendar.id,
        accountId: this.props.message.accountId,
      })
    );
  }
}

export default EventHeader;
