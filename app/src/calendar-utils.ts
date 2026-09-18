import { AccountStore, Calendar, RegExpUtils, Utils } from 'mailspring-exports';
import { findOneIana } from 'windows-iana';

type ICAL = typeof import('ical.js').default;
type ICALComponent = InstanceType<ICAL['Component']>;
type ICALProperty = InstanceType<ICAL['Property']>;
type ICALEvent = InstanceType<ICAL['Event']>;

let ICAL: ICAL = null;

export type ICSParticipantStatus =
  | 'NEEDS-ACTION'
  | 'ACCEPTED'
  | 'DECLINED'
  | 'TENTATIVE'
  | 'DELEGATED'
  | 'COMPLETED'
  | 'IN-PROCESS';

export interface ICSParticipant {
  email: string | null;
  role: 'CHAIR' | 'REQ-PARTICIPANT' | 'OPT-PARTICIPANT' | 'NON-PARTICIPANT';
  status: ICSParticipantStatus;
  component: ICALProperty;
}

function fixJCalDatesWithoutTimes(jCal) {
  jCal[1].forEach((property) => {
    if (
      property[0] === 'dtstart' ||
      property[0] === 'dtend' ||
      property[0] === 'exdate' ||
      property[0] === 'rdate'
    ) {
      if (!property[1].value && property[2] === 'date-time' && /T::$/.test(property[3])) {
        property[2] = 'date';
        property[3] = property[3].replace(/T::$/, '');
      }
    }
  });
  jCal[2].forEach(fixJCalDatesWithoutTimes);
}

export function parseICSString(ics: string) {
  if (!ICAL) {
    ICAL = require('ical.js');
  }
  const jcalData = ICAL.parse(ics);

  // workaround https://github.com/mozilla-comm/ical.js/issues/186
  fixJCalDatesWithoutTimes(jcalData);

  const root = new ICAL.Component(jcalData);
  // Before ICAL.Event: relating the exceptions reads each RECURRENCE-ID, and a value read once
  // keeps the zone it was read with.
  registerTimezones(root);
  const event = new ICAL.Event(root.name === 'vevent' ? root : root.getFirstSubcomponent('vevent'));
  return { root, event };
}

const WHOLE_HISTORY = new Date(Date.UTC(1601, 0, 1));

/**
 * Registers a VCALENDAR's VTIMEZONEs with the ICAL.js TimezoneService, so `toJSDate()` on a
 * TZID-relative time gives the instant it names, and describes any zone a VEVENT refers to
 * without one. RFC 7809 lets a server omit the VTIMEZONE for an IANA zone, and ical.js has no
 * zone data of its own, so such a value would otherwise read as floating local time. The
 * synthesised zone describes the rules in force at the property's own date, and the registry is
 * process-wide: the first file to name a zone without its VTIMEZONE fixes those rules for every
 * later file that also omits it. A file that carries the VTIMEZONE replaces it. An identifier
 * no zone can be found for is left alone.
 */
function registerTimezones(vcalendar: ICALComponent): void {
  for (const vtz of vcalendar.getAllSubcomponents('vtimezone')) {
    // The registry is process-wide and ical.js seeds these three names with UTC itself, so a file
    // redefining one would move every time read through that name for the rest of the session.
    if (['UTC', 'GMT', 'Z'].includes(String(vtz.getFirstPropertyValue('tzid')))) continue;
    ICAL.TimezoneService.register(vtz);
  }

  for (const vevent of vcalendar.getAllSubcomponents('vevent')) {
    for (const prop of vevent.getAllProperties()) {
      const tzid = prop.getParameter('tzid');
      if (typeof tzid !== 'string' || ICAL.TimezoneService.has(tzid)) continue;
      // Read the date off the raw value: hydrating it here would cache it as floating.
      const [, y, m, d] = /^(\d{4})(\d{2})(\d{2})/.exec(String(prop.toJSON()[3])) || [];
      const at = y ? new Date(Date.UTC(+y, +m - 1, +d)) : new Date();
      const vtimezone = createVTIMEZONEString(tzid, at);
      if (!vtimezone) continue;
      ICAL.TimezoneService.register(
        new ICAL.Component(
          ICAL.parse(`BEGIN:VCALENDAR\r\nVERSION:2.0\r\n${vtimezone}\r\nEND:VCALENDAR`)
        ).getFirstSubcomponent('vtimezone')
      );
    }
  }
}

let momentDataHorizonYear: number | null = null;

/**
 * The IANA zone whose rules a TZID describes, or null if we can't identify it.
 *
 * A TZID is an opaque name (RFC 5545 section 3.2.19), and Outlook and Exchange write Windows
 * zone names - "Central Standard Time" rather than "America/Chicago". moment-timezone has no
 * data for those: `moment().tz('Central Standard Time')` logs an error and hands back a
 * moment in *the machine's own zone*, so an event authored in Outlook and edited here would
 * be silently shifted by the difference between the two. windows-iana carries the CLDR
 * mapping that closes it.
 *
 * The original TZID is never rewritten, only resolved for the purpose of computing offsets:
 * an Exchange server understands its own names, and RFC 5545 asks only that whatever name is
 * used be defined by a VTIMEZONE in the same object.
 */
export function resolveIanaZone(tzId: string): string | null {
  if (!tzId) return null;
  const momentTz = require('moment-timezone');
  if (momentTz.tz.zone(tzId)) return tzId;
  const mapped = findOneIana(tzId);
  return mapped && momentTz.tz.zone(mapped) ? mapped : null;
}

/**
 * Builds a VTIMEZONE component describing an IANA timezone's actual offset rules.
 *
 * RFC 5545 section 3.2.19 requires a VTIMEZONE for every TZID an object references, and
 * section 3.6.5 defines it as the authority for resolving those times. Servers that hold
 * their own zone database resolve by TZID name and ignore the body, but the ones that do
 * not - and every recipient reading the file directly - compute from what is written here,
 * so a body claiming a single fixed offset puts every event on the other side of a DST
 * transition an hour out.
 *
 * The rules are read out of moment-timezone rather than invented: the two most recent
 * transitions bracketing `referenceDate` give the STANDARD and DAYLIGHT offsets, and the
 * yearly RRULEs are derived from the transition dates themselves. A zone that does not
 * observe DST yields a single STANDARD component, which is correct rather than degraded.
 *
 * @param tzId - Timezone identifier, IANA or a Windows name Outlook wrote (see
 *   resolveIanaZone). It is reproduced verbatim as the component's TZID.
 * @param referenceDate - The era whose rules are described; zones change them over time
 * @returns A VTIMEZONE ICS string (no surrounding VCALENDAR wrapper), or null when the
 *   identifier names no zone we can describe - inventing rules for it would be worse than
 *   leaving the calendar's own component alone.
 */
export function createVTIMEZONEString(
  tzId: string,
  referenceDate: Date,
  enumeratedYears = ENUMERATED_YEARS
): string | null {
  const momentTz = require('moment-timezone');
  const zoneId = resolveIanaZone(tzId);
  if (!zoneId) return null;
  const zone = momentTz.tz.zone(zoneId);

  const formatOffset = (utcOffsetMin: number) => {
    const abs = Math.abs(utcOffsetMin);
    const sign = utcOffsetMin >= 0 ? '+' : '-';
    return `${sign}${String(Math.floor(abs / 60)).padStart(2, '0')}${String(abs % 60).padStart(
      2,
      '0'
    )}`;
  };

  // A subcomponent needs the offset before the transition as well as after it, so each
  // sample carries both. DTSTART is the local wall-clock instant the rule takes effect,
  // which RFC 5545 section 3.6.5 requires to be a floating time.
  const sample = (at: Date, offsetBeforeMin: number) => {
    const m = momentTz(at).tz(zoneId);
    return {
      dtstart: m.format('YYYYMMDD[T]HHmmss'),
      month: m.month() + 1,
      // The nth weekday of the month, which is how these rules are actually written; a
      // fixed date would drift a day every year. The EU writes its transitions as the *last*
      // Sunday of the month, which is the fifth in some years and the fourth in others, so a
      // positive ordinal would stop matching - BYDAY=-1SU is the rule those zones mean.
      nth: m.clone().add(7, 'days').month() !== m.month() ? -1 : Math.ceil(m.date() / 7),
      weekday: ['SU', 'MO', 'TU', 'WE', 'TH', 'FR', 'SA'][m.day()],
      offsetTo: m.utcOffset(),
      offsetFrom: offsetBeforeMin,
      name: m.zoneAbbr(),
    };
  };

  const block = (kind: 'STANDARD' | 'DAYLIGHT', t: ReturnType<typeof sample>) => [
    `BEGIN:${kind}`,
    `DTSTART:${t.dtstart}`,
    `RRULE:FREQ=YEARLY;BYMONTH=${t.month};BYDAY=${t.nth}${t.weekday}`,
    `TZOFFSETFROM:${formatOffset(t.offsetFrom)}`,
    `TZOFFSETTO:${formatOffset(t.offsetTo)}`,
    `TZNAME:${t.name}`,
    `END:${kind}`,
  ];

  // moment-timezone's `untils` are the instants each offset stops applying. The two that
  // bracket the reference date describe the DST rules in force around it.
  const untils: number[] = (zone && zone.untils) || [];
  const refMs = referenceDate.getTime();
  const idx = untils.findIndex((u) => u !== null && u > refMs);
  const transitions: Date[] = [];
  if (zone && idx > 0) {
    for (const u of [untils[idx - 1], untils[idx]]) {
      if (u !== null && isFinite(u)) transitions.push(new Date(u));
    }
  }

  const samples = transitions.map((at) =>
    // One millisecond before the transition is the offset being left behind.
    sample(
      at,
      momentTz(new Date(at.getTime() - 1))
        .tz(zoneId)
        .utcOffset()
    )
  );
  const daylight = samples.find((t) => samples.some((o) => t.offsetTo > o.offsetTo));
  const standard = samples.find((t) => t !== daylight);

  const body: string[] = [];
  if (daylight && standard) {
    body.push(...block('STANDARD', standard), ...block('DAYLIGHT', daylight));
  } else {
    // No DST in this era: one STANDARD with the offset that actually applies, and no RRULE,
    // because there is no recurring transition to describe.
    const m = momentTz(referenceDate).tz(zoneId);
    body.push(
      'BEGIN:STANDARD',
      'DTSTART:19700101T000000',
      `TZOFFSETFROM:${formatOffset(m.utcOffset())}`,
      `TZOFFSETTO:${formatOffset(m.utcOffset())}`,
      `TZNAME:${m.zoneAbbr()}`,
      'END:STANDARD'
    );
  }

  return ['BEGIN:VTIMEZONE', `TZID:${tzId}`, ...body, 'END:VTIMEZONE'].join('\r\n');
}

export function emailFromParticipantURI(uri: string): string | null {
  if (!uri) {
    return null;
  }

  // Normalize to lowercase for comparison
  const uriLower = uri.toLowerCase();

  // Handle mailto: URI format (most common)
  // e.g., "mailto:user@example.com" or "MAILTO:user@example.com"
  if (uriLower.startsWith('mailto:')) {
    const email = uri.slice(7).toLowerCase(); // preserve original then lowercase
    if (email.includes('@')) {
      return email;
    }
    return null;
  }

  // Handle bare email addresses (no mailto: prefix)
  // Some calendar systems just use "user@example.com" directly
  // Use RegExpUtils.emailRegex which supports international characters,
  // and verify the match covers the entire string to reject malformed inputs
  const emailRegex = RegExpUtils.emailRegex();
  const bareMatch = emailRegex.exec(uri);
  if (bareMatch && bareMatch.index === 0 && bareMatch[0].length === uri.length) {
    return uri.toLowerCase();
  }

  // Try to extract an email pattern from the string as a last resort.
  // This handles edge cases like "invalid:user@example.com" or other
  // non-standard formats where the email is embedded in the string.
  emailRegex.lastIndex = 0; // Reset regex state since it has the 'g' flag
  const embeddedMatch = emailRegex.exec(uri);
  if (embeddedMatch) {
    return embeddedMatch[0].toLowerCase();
  }

  return null;
}

export function cleanParticipants(icsEvent: ICALEvent): ICSParticipant[] {
  return icsEvent.attendees.map((a) => ({
    component: a,
    status: (a.getParameter('partstat') || 'NEEDS-ACTION') as ICSParticipantStatus,
    role: (a.getParameter('role') || 'REQ-PARTICIPANT') as ICSParticipant['role'],
    email:
      a
        .getValues()
        .map(emailFromParticipantURI)
        .find((v) => !!v) || null,
  }));
}

export function selfParticipant(
  icsEvent: ICALEvent,
  accountId: string
): ICSParticipant | undefined {
  const me = cleanParticipants(icsEvent).find((a) => {
    const acct = AccountStore.accountForEmail(a.email);
    return acct && acct.id === accountId;
  });
  return me;
}

/**
 * Whether a calendar belongs to the account, rather than being one somebody shared with it.
 *
 * This decides where an RSVP is written, so getting it wrong means editing an event on
 * someone else's calendar.
 *
 * DAV:owner is the server's own answer (RFC 3744 section 5.1) and is taken whenever the
 * server gives one. The fallback - matching the calendar's display name against the
 * account's addresses, which is how Google, Google Workspace and Fastmail name a user's
 * default calendar - only applies when the server said nothing, because a display name is
 * text its owner chooses: somebody who shares a writable calendar named after the
 * recipient's own address would otherwise capture their replies.
 */
export function isOwnCalendar(calendar: Calendar, addresses: string[]): boolean {
  if (calendar.ownership === 'mine') return true;
  if (calendar.ownership === 'other') return false;
  return addresses.some((address) => Utils.emailIsEquivalent(calendar.name, address));
}

/**
 * Whether the server has positively identified this calendar as somebody else's.
 *
 * Distinct from `!isOwnCalendar`: that is also true of a calendar nobody has said anything
 * about, which the "only one candidate" fallbacks are willing to use. This is only true when
 * the server named an owner and it wasn't us, which rules a calendar out entirely.
 *
 * Both paths that write an answer onto an event - the invitation header and the calendar's
 * own RSVP - must agree on this, or the same meeting is answerable from one and not the
 * other, and the disagreement is a write onto a colleague's calendar.
 */
export function isSomeoneElsesCalendar(calendar: Calendar | undefined | null): boolean {
  return !!calendar && calendar.ownership === 'other';
}
