import crypto from 'crypto';
import { findOneIana } from 'windows-iana';
import { parseICSString, emailFromParticipantURI } from './calendar-utils';
import { calendarDateFromUnix, shiftedDayStartUnix, calendarDaysBetween } from './calendar-date';

type ICAL = typeof import('ical.js').default;
type ICALComponent = InstanceType<ICAL['Component']>;
type ICALTime = InstanceType<ICAL['Time']>;
type ICALTimezone = InstanceType<ICAL['Timezone']>;

let ICAL: ICAL = null;

function getICAL(): ICAL {
  if (!ICAL) {
    ICAL = require('ical.js');
  }
  return ICAL;
}

/**
 * The current instant as a UTC ICAL.Time for DTSTAMP, which RFC 5545 section 3.8.7.2 requires
 * in UTC. ICAL.Time.now() is floating local time and serializes without the Z.
 */
function nowUTC(ical: ICAL) {
  return ical.Time.fromJSDate(new Date(), true);
}

/**
 * Options for creating a new ICS event
 */
export interface CreateEventOptions {
  uid?: string;
  summary: string;
  description?: string;
  location?: string;
  start: Date;
  end: Date;
  isAllDay?: boolean;
  timezone?: string; // IANA timezone identifier (e.g., 'America/New_York')
  organizer?: { email: string; name?: string };
  attendees?: Array<{ email: string; name?: string; role?: string; partstat?: string }>;
  recurrenceRule?: string;
}

/**
 * Options for updating event times
 */
export interface UpdateTimesOptions {
  start: number; // Unix timestamp in seconds
  end: number; // Unix timestamp in seconds
  isAllDay?: boolean;
  timezone?: string; // Optional IANA timezone to set (overrides event's original timezone)
}

/**
 * Result of creating a recurrence exception
 */
export interface RecurrenceExceptionResult {
  /** Updated master ICS with the exception VEVENT embedded inline */
  masterIcs: string;
  /** The RECURRENCE-ID value for the exception */
  recurrenceId: string;
}

/**
 * Information about an event's recurrence
 */
export interface RecurrenceInfo {
  isRecurring: boolean;
  rule?: string;
  frequency?: string;
}

/**
 * Generates a unique ID for calendar events.
 *
 * RFC 7986 section 5.3 asks for a UID with the uniqueness properties of a UUID, and warns
 * against deriving one from anything the event itself contains. `Math.random()` is not a
 * source a collision argument can rest on - V8 gives it 128 bits of internal state but no
 * guarantee across contexts - so this takes the platform's CSPRNG-backed generator.
 */
export function generateUID(): string {
  return `${crypto.randomUUID()}@mailspring`;
}

/**
 * How many occurrences an expander may step through before giving up on one series.
 *
 * ical-expander iterates forward from DTSTART with no way to seek, so a cap limits how far
 * back a series may begin rather than how much work a window costs. Too low and a long
 * running series silently expands to nothing: at 100, a weekly meeting that started three
 * years ago never reaches the present and simply disappears from the calendar. Removing the
 * cap is worse - an invitation is untrusted input, and `RRULE:FREQ=SECONDLY` dated 1970
 * would spin forever.
 *
 * So the budget comes from the series itself: how many steps of its own frequency fit
 * between where it starts and the end of the window, plus slack. Real calendars land far
 * below the ceiling - a weekly meeting running since 2020 needs about 300 - while a
 * frequency fine enough to be abusive exceeds it and is truncated instead of expanded.
 *
 * @param ics - The series' calendar object.
 * @param seriesStartUnix - DTSTART of the series, in unix seconds.
 * @param windowEndUnix - The end of the range being expanded, in unix seconds.
 */
export function expansionIterationBudget(
  ics: string,
  seriesStartUnix: number,
  windowEndUnix: number
): number {
  // A master event can reach here with a null or non-finite recurrenceStart - the expansion
  // fallback guards for exactly that. NaN would survive Math.max/Math.min and become the cap
  // itself, and ical-expander's loop is `!this.maxIterations || i < this.maxIterations`
  // (index.js:104), so a NaN cap switches the limit off rather than truncating.
  if (!Number.isFinite(seriesStartUnix) || !Number.isFinite(windowEndUnix)) {
    return MIN_EXPANSION_ITERATIONS;
  }
  const rule = firstVeventRRule(ics);
  if (!rule) {
    return MIN_EXPANSION_ITERATIONS; // not a series; one occurrence is all there is to reach
  }
  const freq = /FREQ=([A-Z]+)/i.exec(rule);
  const interval = parseInt((/INTERVAL=(\d+)/i.exec(rule) || [])[1], 10) || 1;
  const step =
    (EXPANSION_STEP_SECONDS[(freq ? freq[1] : '').toUpperCase()] || EXPANSION_STEP_SECONDS.DAILY) *
    interval;
  const steps = Math.ceil(Math.max(0, windowEndUnix - seriesStartUnix) / step) + 100;
  return Math.min(MAX_EXPANSION_ITERATIONS, Math.max(MIN_EXPANSION_ITERATIONS, steps));
}

/**
 * The RRULE of the first VEVENT, ignoring any that belong to a VTIMEZONE.
 *
 * A VTIMEZONE's STANDARD and DAYLIGHT blocks each carry their own RRULE describing the
 * zone's DST transitions, and they appear before the VEVENT - so a plain search for the
 * first RRULE in the file returns `FREQ=YEARLY;BYMONTH=3;BYDAY=2SU` for a weekly meeting,
 * and any budget derived from it is wrong by a factor of fifty.
 */
function firstVeventRRule(ics: string): string | null {
  const unfolded = ics.replace(/\r\n[ \t]/g, '').replace(/\n[ \t]/g, '');
  const vevent = unfolded.split(/^BEGIN:VEVENT$/m)[1];
  if (!vevent) return null;
  const match = /^RRULE:(.*)$/im.exec(vevent.split(/^END:VEVENT$/m)[0]);
  return match ? match[1] : null;
}

const EXPANSION_STEP_SECONDS: { [freq: string]: number } = {
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
const MIN_EXPANSION_ITERATIONS = 1000;
const MAX_EXPANSION_ITERATIONS = 50000;

/**
 * Formats a Date as an ICS date-only string (YYYYMMDD)
 * Uses LOCAL date components since all-day events represent a day in the user's timezone
 */
function formatDateOnly(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}${month}${day}`;
}

/**
 * Formats a Date as an ICS datetime string in UTC (YYYYMMDDTHHMMSSZ)
 */
function formatDateTimeUTC(date: Date): string {
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');
  const hours = String(date.getUTCHours()).padStart(2, '0');
  const minutes = String(date.getUTCMinutes()).padStart(2, '0');
  const seconds = String(date.getUTCSeconds()).padStart(2, '0');
  return `${year}${month}${day}T${hours}${minutes}${seconds}Z`;
}

/**
 * Creates an ICAL.Time for an all-day event (DATE type, no time component)
 * Uses local date since all-day events represent a calendar day in user's timezone
 */
function createAllDayTime(date: Date, ical: ICAL): ICALTime {
  const time = new ical.Time(
    {
      year: date.getFullYear(),
      month: date.getMonth() + 1,
      day: date.getDate(),
      isDate: true,
    },
    null // timezone parameter (null for floating/all-day)
  );
  return time;
}

/**
 * Creates an ICAL.Time for an all-day event's DTEND, which RFC 5545 defines as exclusive:
 * midnight of the day after the last day covered. Ends don't always arrive on midnight —
 * the all-day toggle passes the event's untouched wall-clock time — and truncating those
 * to a DATE would land on the start's own day.
 * @param start Event start; floors the result so a degenerate end can't precede it
 * @param end Event end; one already at midnight is not pushed out another day
 */
function createAllDayEndTime(start: Date, end: Date, ical: ICAL): ICALTime {
  const lastCovered = new Date(Math.max(end.getTime() - 1, start.getTime()));
  const exclusiveEnd = new Date(
    lastCovered.getFullYear(),
    lastCovered.getMonth(),
    lastCovered.getDate() + 1
  );
  return createAllDayTime(exclusiveEnd, ical);
}

/**
 * Creates an ICAL.Time from a Date, optionally preserving a specific timezone.
 * For timed events, this properly handles timezone conversion.
 *
 * @param date - JavaScript Date (represents a moment in time)
 * @param isAllDay - Whether this is an all-day event
 * @param ical - The ICAL library reference
 * @param preserveZone - Optional timezone to use (from original event)
 */
function createICALTime(
  date: Date,
  isAllDay: boolean,
  ical: ICAL,
  preserveZone?: ICALTimezone | null
): ICALTime {
  if (isAllDay) {
    return createAllDayTime(date, ical);
  }

  // For timed events with a timezone to preserve
  if (
    preserveZone &&
    preserveZone.tzid &&
    preserveZone.tzid !== 'UTC' &&
    preserveZone.tzid !== 'floating'
  ) {
    // Create the time in UTC first, then convert to the target timezone
    // This ensures the moment in time is preserved correctly
    const utcTime = ical.Time.fromJSDate(date, true); // true = use UTC

    // Convert to target timezone
    // Note: This adjusts the wall-clock time to show the same moment in the target zone
    const zonedTime = utcTime.convertToZone(preserveZone);
    return zonedTime;
  }

  // Default: create time in UTC (floating time)
  return ical.Time.fromJSDate(date, true);
}

/**
 * Adds an EXDATE property to a VEVENT, preserving the TZID parameter when needed.
 *
 * ICAL.js's `addPropertyWithValue('exdate', time)` does NOT set the TZID parameter
 * even when the time has a timezone attached. This causes ical-expander to fail to
 * match the EXDATE against zoned occurrences (the EXDATE is serialized as floating
 * time instead of zoned time). We must manually create the property and set TZID.
 */
function addExdateProperty(
  vevent: ICALComponent,
  exdateTime: ICALTime,
  ical: ICAL,
  zone?: ICALTimezone | null
): void {
  const exProp = new ical.Property('exdate', vevent);
  if (zone && zone.tzid && zone.tzid !== 'UTC' && zone.tzid !== 'floating') {
    exProp.setParameter('tzid', zone.tzid);
  }
  exProp.setValue(exdateTime);
  vevent.addProperty(exProp);
}

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
function resolveIanaZone(tzId: string): string | null {
  if (!tzId) return null;
  const momentTz = require('moment-timezone');
  if (momentTz.tz.zone(tzId)) return tzId;
  const mapped = findOneIana(tzId);
  return mapped && momentTz.tz.zone(mapped) ? mapped : null;
}

/**
 * Brings a VCALENDAR's VTIMEZONE components in line with the TZIDs its properties reference.
 *
 * RFC 5545 section 3.2.19 makes a TZID meaningful only by reference to a VTIMEZONE in the
 * same iCalendar object, so the two have to be changed together. Retiming an event into a
 * different zone leaves any inline exception still sitting in the old one, and dropping that
 * zone's VTIMEZONE - as replacing the whole set does - leaves the exception pointing at
 * nothing, which a strict parser is entitled to reject and a lenient one reads as floating.
 *
 * Zones the calendar still refers to are kept as the server wrote them; a server's own
 * VTIMEZONE describes its rules better than one we synthesise. Only a referenced zone with
 * no component at all gets one, and only when moment-timezone recognises the identifier -
 * an unrecognised TZID (a Windows zone name, a private X- identifier) is left exactly as it
 * arrived rather than described wrongly.
 */
function syncVTimezones(vcalendar: ICALComponent, ical: ICAL, referenceDate: Date): void {
  const referenced = new Set<string>();
  for (const component of vcalendar.getAllSubcomponents()) {
    if (component.name === 'vtimezone') continue;
    for (const prop of component.getAllProperties()) {
      const tzid = prop.getParameter('tzid');
      if (tzid) referenced.add(String(tzid));
    }
  }

  const present = new Set<string>();
  for (const vtz of vcalendar.getAllSubcomponents('vtimezone')) {
    const tzid = String(vtz.getFirstPropertyValue('tzid') || '');
    if (!tzid || !referenced.has(tzid)) {
      vcalendar.removeSubcomponent(vtz);
      continue;
    }
    present.add(tzid);
  }

  for (const tzid of referenced) {
    if (present.has(tzid)) continue;
    const vtimezone = createVTIMEZONEString(tzid, referenceDate);
    if (!vtimezone) continue; // unidentifiable zone: leave it exactly as it arrived
    vcalendar.addSubcomponent(
      new ical.Component(
        ical.parse(`BEGIN:VCALENDAR\r\nVERSION:2.0\r\n${vtimezone}\r\nEND:VCALENDAR`)
      ).getFirstSubcomponent('vtimezone')
    );
  }
}

/**
 * Registers all VTIMEZONE subcomponents from a VCALENDAR with the ICAL.js
 * TimezoneService so that subsequent `toJSDate()` calls on TZID-relative times
 * resolve correctly. Duplicate registrations are silently ignored.
 */
function registerTimezones(vcalendar: ICALComponent, ical: ICAL): void {
  for (const vtz of vcalendar.getAllSubcomponents('vtimezone')) {
    try {
      ical.TimezoneService.register(vtz);
    } catch (_) {
      // Ignore duplicate registrations (same TZID registered more than once)
    }
  }
}

/**
 * Removes an existing exception VEVENT from a VCALENDAR that matches the given
 * target time (in UTC milliseconds). Uses `toJSDate().getTime()` for comparison
 * after registering timezones, so TZID-formatted and UTC-formatted RECURRENCE-IDs
 * are both correctly identified as the same moment. Falls back to string comparison
 * if `toJSDate()` throws (e.g., unregistered timezone).
 *
 * This implements the "upsert" behaviour: re-editing an existing exception replaces
 * the old VEVENT rather than adding a duplicate.
 *
 * @param vcalendar - The VCALENDAR component to search within
 * @param targetMs - The expected RECURRENCE-ID moment in UTC milliseconds
 * @param recurrenceId - The formatted RECURRENCE-ID string (used as a string fallback)
 * @param isAllDay - Whether the event is all-day (affects string-fallback comparison)
 * @param ical - The ICAL library reference
 */
function removeExistingExceptionVevent(
  vcalendar: ICALComponent,
  targetMs: number,
  recurrenceId: string,
  isAllDay: boolean,
  ical: ICAL
): void {
  for (const existing of vcalendar.getAllSubcomponents('vevent')) {
    const ridValue = existing.getFirstPropertyValue('recurrence-id') as any;
    if (!ridValue) continue;

    try {
      if (ridValue.toJSDate().getTime() === targetMs) {
        vcalendar.removeSubcomponent(existing);
        break;
      }
    } catch (_) {
      // Fallback: string comparison (e.g., if toJSDate throws for an unknown timezone)
      const ridStr =
        typeof ridValue.toString === 'function' ? ridValue.toString() : String(ridValue);
      const ridFormatted = isAllDay
        ? ridStr.replace(/[^0-9]/g, '').substring(0, 8)
        : ridStr.replace(/[^0-9TZ]/g, '');
      const targetFormatted = recurrenceId.replace(/[^0-9TZ]/g, '');
      if (ridFormatted === targetFormatted || ridStr === recurrenceId) {
        vcalendar.removeSubcomponent(existing);
        break;
      }
    }
  }
}

/**
 * Validates timestamp options and throws if invalid
 */
function validateTimestamps(start: number, end: number): void {
  if (typeof start !== 'number' || typeof end !== 'number') {
    throw new Error('Invalid timestamps: start and end must be numbers');
  }
  if (start < 0 || end < 0) {
    throw new Error('Invalid timestamps: must be positive');
  }
  if (end < start) {
    throw new Error('Invalid timestamps: end time must be after or equal to start time');
  }
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
export function createVTIMEZONEString(tzId: string, referenceDate: Date): string | null {
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

/**
 * Creates a new ICS string for an event
 *
 * @param options - Event creation options including title, times, and optional timezone
 * @returns A valid ICS string representing the event
 */
export function createICSString(options: CreateEventOptions): string {
  const ical = getICAL();
  const isAllDay = options.isAllDay ?? false;

  // Create VCALENDAR component
  const calendar = new ical.Component(['vcalendar', [], []]);
  calendar.updatePropertyWithValue('prodid', '-//Mailspring//Calendar//EN');
  calendar.updatePropertyWithValue('version', '2.0');
  calendar.updatePropertyWithValue('calscale', 'GREGORIAN');

  // Create VEVENT component
  const vevent = new ical.Component('vevent');
  const event = new ical.Event(vevent);

  // Set UID
  event.uid = options.uid || generateUID();

  // Set summary (title)
  event.summary = options.summary;

  // A zone we cannot identify can't be used to derive wall-clock components: moment would
  // quietly substitute the machine's own zone and write the wrong time. Such an event falls
  // through to the UTC path below, which is unambiguous.
  const createZone = options.timezone ? resolveIanaZone(options.timezone) : null;

  if (!isAllDay && createZone) {
    // ical.js TimezoneService only knows UTC/GMT/Z by default — IANA timezone names
    // like "America/Chicago" are never registered, so we can't use it for conversion.
    // Instead, use moment-timezone to extract the correct local time components and
    // create a floating ICAL.Time, then manually stamp the TZID onto the property.
    // This produces: DTSTART;TZID=America/Chicago:20240115T140000
    //
    // The matching VTIMEZONE is added by syncVTimezones once the TZIDs are stamped below;
    // without one, servers that don't carry their own zone database (Yahoo among them)
    // ignore the TZID and read the wall-clock time as UTC.
    const momentTz = require('moment-timezone');
    const startM = momentTz(options.start).tz(createZone);
    const endM = momentTz(options.end).tz(createZone);

    event.startDate = new ical.Time(
      {
        year: startM.year(),
        month: startM.month() + 1, // moment months are 0-indexed
        day: startM.date(),
        hour: startM.hour(),
        minute: startM.minute(),
        second: startM.second(),
        isDate: false,
      },
      ical.Timezone.localTimezone
    );
    event.endDate = new ical.Time(
      {
        year: endM.year(),
        month: endM.month() + 1,
        day: endM.date(),
        hour: endM.hour(),
        minute: endM.minute(),
        second: endM.second(),
        isDate: false,
      },
      ical.Timezone.localTimezone
    );

    vevent.getFirstProperty('dtstart')?.setParameter('tzid', options.timezone);
    vevent.getFirstProperty('dtend')?.setParameter('tzid', options.timezone);
  } else {
    // All-day or no-timezone: use existing path
    const eventTimezone: ICALTimezone | null = null;
    event.startDate = createICALTime(options.start, isAllDay, ical, eventTimezone);
    event.endDate = isAllDay
      ? createAllDayEndTime(options.start, options.end, ical)
      : createICALTime(options.end, false, ical, eventTimezone);
  }

  // Set optional properties
  if (options.description) {
    event.description = options.description;
  }
  if (options.location) {
    event.location = options.location;
  }

  // Set organizer
  if (options.organizer) {
    const organizer = new ical.Property('organizer');
    organizer.setValue(`mailto:${options.organizer.email}`);
    if (options.organizer.name) {
      organizer.setParameter('cn', options.organizer.name);
    }
    vevent.addProperty(organizer);
  }

  // Set attendees
  if (options.attendees) {
    for (const attendee of options.attendees) {
      const prop = new ical.Property('attendee');
      prop.setValue(`mailto:${attendee.email}`);
      if (attendee.name) {
        prop.setParameter('cn', attendee.name);
      }
      prop.setParameter('partstat', attendee.partstat || 'NEEDS-ACTION');
      prop.setParameter('role', attendee.role || 'REQ-PARTICIPANT');
      prop.setParameter('rsvp', 'TRUE');
      vevent.addProperty(prop);
    }
  }

  // Set recurrence rule
  if (options.recurrenceRule) {
    vevent.addPropertyWithValue('rrule', ical.Recur.fromString(options.recurrenceRule));
  }

  // Set timestamp
  vevent.addPropertyWithValue('dtstamp', nowUTC(ical));
  // Guests' clients compare SEQUENCE to decide whether an update supersedes what they hold.
  vevent.addPropertyWithValue('sequence', 0);

  calendar.addSubcomponent(vevent);
  syncVTimezones(calendar, ical, options.start);
  return calendar.toString();
}

/**
 * Records a material change by the organizer, so attendees' clients treat it as an update.
 *
 * RFC 5546 section 2.1.4: the organizer increments SEQUENCE whenever they change something
 * that matters to the guests - the time, the recurrence, whether an occurrence happens at
 * all. A receiving client compares SEQUENCE against the copy it already holds and ignores
 * anything that hasn't advanced, so an update sent without this reaches the guests and then
 * does nothing.
 *
 * An absent SEQUENCE means zero (section 3.7.4), so a first change writes 1 rather than
 * being skipped.
 *
 * This is deliberately not done inside the individual edit helpers. One save from the
 * popover runs several of them - times, guests, recurrence rule - and a bump in each made
 * SEQUENCE jump by three for a single revision. What a revision *is* is only known where
 * the change is assembled, so the callers that publish one call this once, at the end.
 *
 * Never called when an attendee answers an invitation: a REPLY leaves SEQUENCE alone,
 * because the attendee is not changing the event.
 *
 * @param recurrenceId Bump the matching inline exception rather than the master, for an
 *   edit that applies to one occurrence.
 */
export function bumpEventSequence(ics: string, recurrenceId?: string): string {
  const { root } = parseICSString(ics);
  const vevents =
    root.name === 'vevent' ? [root] : (root.getAllSubcomponents('vevent') as ICALComponent[]);
  if (!vevents.length) {
    throw new Error('Invalid ICS: no VEVENT component found');
  }

  const target = recurrenceId
    ? vevents.find((v) => matchesRecurrenceId(v, recurrenceId))
    : vevents.find((v) => !v.getFirstPropertyValue('recurrence-id')) || vevents[0];
  if (!target) {
    throw new Error(`No VEVENT found with RECURRENCE-ID matching ${recurrenceId}`);
  }

  const current = target.getFirstPropertyValue('sequence');
  target.updatePropertyWithValue('sequence', (parseInt(String(current ?? 0), 10) || 0) + 1);
  return root.toString();
}

/** Whether this VEVENT is the inline exception identified by `recurrenceId`. */
function matchesRecurrenceId(vevent: ICALComponent, recurrenceId: string): boolean {
  const rid = vevent.getFirstPropertyValue('recurrence-id');
  if (!rid) return false;
  const ridStr = typeof rid === 'string' ? rid : String((rid as any).toString());
  return (
    ridStr === recurrenceId ||
    ridStr.replace(/[^0-9TZ]/g, '') === recurrenceId.replace(/[^0-9TZ]/g, '')
  );
}

/**
 * Updates the start/end times in an event's ICS data.
 * Preserves all other event properties and properly handles timezone conversion.
 *
 * @param ics - The original ICS string
 * @param options - New start/end times and whether it's an all-day event
 * @returns The modified ICS string
 */
export function updateEventTimes(ics: string, options: UpdateTimesOptions): string {
  // Validate inputs
  validateTimestamps(options.start, options.end);

  const ical = getICAL();
  const { root, event } = parseICSString(ics);

  const startDate = new Date(options.start * 1000);
  const endDate = new Date(options.end * 1000);
  const isAllDay = options.isAllDay ?? false;

  const vevent = root.name === 'vevent' ? root : root.getFirstSubcomponent('vevent');
  if (!vevent) {
    throw new Error('Invalid ICS: no VEVENT component found');
  }

  // As in createICSString: an unidentifiable zone would have moment silently substitute the
  // machine's own, so those events are retimed in UTC rather than at the wrong wall clock.
  const updateZone = options.timezone ? resolveIanaZone(options.timezone) : null;

  if (!isAllDay && updateZone) {
    // User selected a specific timezone — encode wall-clock time in that zone.
    // This mirrors the timezone path in createICSString.
    const momentTz = require('moment-timezone');
    const startM = momentTz(startDate).tz(updateZone);
    const endM = momentTz(endDate).tz(updateZone);

    event.startDate = new ical.Time(
      {
        year: startM.year(),
        month: startM.month() + 1,
        day: startM.date(),
        hour: startM.hour(),
        minute: startM.minute(),
        second: startM.second(),
        isDate: false,
      },
      ical.Timezone.localTimezone
    );
    event.endDate = new ical.Time(
      {
        year: endM.year(),
        month: endM.month() + 1,
        day: endM.date(),
        hour: endM.hour(),
        minute: endM.minute(),
        second: endM.second(),
        isDate: false,
      },
      ical.Timezone.localTimezone
    );

    // Stamp TZID on the date properties
    vevent.getFirstProperty('dtstart')?.setParameter('tzid', options.timezone);
    vevent.getFirstProperty('dtend')?.setParameter('tzid', options.timezone);
  } else {
    // Preserve the original timezone for timed events, or use floating for all-day
    const originalStartZone = event.startDate?.zone;
    const originalEndZone = event.endDate?.zone;
    event.startDate = createICALTime(startDate, isAllDay, ical, originalStartZone);
    event.endDate = isAllDay
      ? createAllDayEndTime(startDate, endDate, ical)
      : createICALTime(endDate, false, ical, originalEndZone);
  }

  // Update DTSTAMP to indicate modification
  vevent.updatePropertyWithValue('dtstamp', nowUTC(ical));

  if (root.name === 'vcalendar') {
    syncVTimezones(root, ical, startDate);
  }

  return root.toString();
}

/**
 * Creates an exception instance for a recurring event by embedding the exception
 * VEVENT inline in the master VCALENDAR (RFC 4791 §4.1 / RFC 5545 compliant).
 *
 * Unlike the old approach (separate VCALENDAR + EXDATE), this embeds the exception
 * VEVENT directly into the master's VCALENDAR so the entire updated master ICS can
 * be PUT to the same resource as a single update task.
 *
 * Upsert semantics: if a VEVENT with the same RECURRENCE-ID already exists in the
 * master VCALENDAR (e.g. re-editing an already-excepted occurrence), it is replaced.
 *
 * @param masterIcs - The master event's ICS data
 * @param originalOccurrenceStart - The original start time of the occurrence being modified (unix seconds)
 * @param newStart - New start time (unix seconds)
 * @param newEnd - New end time (unix seconds)
 * @param isAllDay - Whether this is an all-day event
 * @returns Object with updated master ICS (exception embedded inline) and the RECURRENCE-ID string
 */
export function createRecurrenceException(
  masterIcs: string,
  originalOccurrenceStart: number,
  newStart: number,
  newEnd: number,
  isAllDay: boolean
): RecurrenceExceptionResult {
  // Validate inputs
  validateTimestamps(newStart, newEnd);

  const ical = getICAL();
  const { root: masterRoot, event: masterEvent } = parseICSString(masterIcs);

  // Get the original timezone from the master event to preserve it
  const originalStartZone = masterEvent.startDate?.zone;

  // Create RECURRENCE-ID value from original occurrence start
  const originalDate = new Date(originalOccurrenceStart * 1000);
  const recurrenceId = isAllDay ? formatDateOnly(originalDate) : formatDateTimeUTC(originalDate);

  // masterRoot must be a VCALENDAR (not a bare VEVENT) for inline embedding
  const vcalendar = masterRoot.name === 'vcalendar' ? masterRoot : null;
  const masterVevent = vcalendar
    ? vcalendar.getFirstSubcomponent('vevent')
    : masterRoot.name === 'vevent'
      ? masterRoot
      : null;

  if (!masterVevent) {
    throw new Error('Invalid ICS: no VEVENT component found');
  }

  // Upsert: remove any existing exception VEVENT with this RECURRENCE-ID so that
  // re-editing a previously excepted occurrence replaces the old VEVENT rather than
  // accumulating duplicates. The helper compares by UTC milliseconds so TZID-formatted
  // and UTC-formatted RECURRENCE-IDs are recognised as the same moment.
  const targetMs = originalDate.getTime();
  if (vcalendar) {
    removeExistingExceptionVevent(vcalendar, targetMs, recurrenceId, isAllDay, ical);
  }

  // Deep-clone the master VEVENT for the exception.
  // ical.Component.toJSON() returns a reference to the internal jCal array, NOT a copy.
  // Without JSON.parse/stringify the cloned component shares the same array as the master,
  // so every mutation below (removeAllProperties, updatePropertyWithValue, etc.) silently
  // mutates the master VEVENT too, producing two identical exception VEVENTs and no master.
  const exceptionVevent = new ical.Component(JSON.parse(JSON.stringify(masterVevent.toJSON())));

  // Remove recurrence rule and exclusion dates from the exception (it's a single instance)
  exceptionVevent.removeAllProperties('rrule');
  exceptionVevent.removeAllProperties('rdate');
  exceptionVevent.removeAllProperties('exdate');

  // Set RECURRENCE-ID using UTC format so it is unambiguous and matches the returned
  // recurrenceId string (which is also UTC via formatDateTimeUTC).
  // Using createICALTime with a named timezone produces a floating-time serialization
  // (no TZID parameter on the property) because updatePropertyWithValue does not auto-set TZID.
  const recIdTime = isAllDay
    ? createAllDayTime(originalDate, ical)
    : ical.Time.fromJSDate(originalDate, true); // UTC → serializes as YYYYMMDDTHHMMSSz
  exceptionVevent.updatePropertyWithValue('recurrence-id', recIdTime);

  // Set new times on the exception (preserve timezone)
  const newStartDate = new Date(newStart * 1000);
  const newEndDate = new Date(newEnd * 1000);
  const exceptionICALEvent = new ical.Event(exceptionVevent);
  exceptionICALEvent.startDate = createICALTime(newStartDate, isAllDay, ical, originalStartZone);
  exceptionICALEvent.endDate = isAllDay
    ? createAllDayEndTime(newStartDate, newEndDate, ical)
    : createICALTime(newEndDate, false, ical, originalStartZone);

  // Update DTSTAMP; SEQUENCE is the caller's to advance once per revision.
  const now = nowUTC(ical);
  masterVevent.updatePropertyWithValue('dtstamp', now);
  exceptionVevent.updatePropertyWithValue('dtstamp', now);

  // Embed the exception VEVENT inline in the master VCALENDAR
  if (vcalendar) {
    vcalendar.addSubcomponent(exceptionVevent);
  }

  return {
    masterIcs: masterRoot.toString(),
    recurrenceId,
  };
}

/**
 * Applies property edits (summary, location, description, attendees) to an inline
 * exception VEVENT inside a master VCALENDAR ICS string.
 *
 * This is needed because `updateEventProperty` and `updateAttendees` target the
 * first VEVENT (the master), not a specific exception VEVENT identified by RECURRENCE-ID.
 *
 * @param masterIcs - Master VCALENDAR ICS containing the inline exception VEVENT
 * @param recurrenceId - The RECURRENCE-ID string of the exception to edit
 * @param edits - Property values to apply
 * @returns Updated master ICS string
 */
export function applyEditsToException(
  masterIcs: string,
  recurrenceId: string,
  edits: {
    summary?: string;
    location?: string;
    description?: string;
    attendees?: Array<{ email: string; name?: string | null; partstat?: string }>;
  }
): string {
  const ical = getICAL();
  const { root } = parseICSString(masterIcs);

  const vcalendar = root.name === 'vcalendar' ? root : null;
  if (!vcalendar) {
    throw new Error('Invalid ICS: expected VCALENDAR root');
  }

  const exceptionVevent =
    (vcalendar.getAllSubcomponents('vevent') as ICALComponent[]).find((v) =>
      matchesRecurrenceId(v, recurrenceId)
    ) || null;

  if (!exceptionVevent) {
    throw new Error(`No exception VEVENT found with RECURRENCE-ID matching ${recurrenceId}`);
  }

  const exceptionICALEvent = new ical.Event(exceptionVevent);

  if (edits.summary !== undefined) {
    exceptionICALEvent.summary = edits.summary;
  }
  if (edits.description !== undefined) {
    exceptionICALEvent.description = edits.description;
  }
  if (edits.location !== undefined) {
    exceptionICALEvent.location = edits.location;
  }
  if (edits.attendees !== undefined) {
    reconcileAttendees(exceptionVevent, edits.attendees);
  }

  exceptionVevent.updatePropertyWithValue('dtstamp', nowUTC(ical));

  return root.toString();
}

/**
 * Shifts the RECURRENCE-ID of all inline exception VEVENTs within a master VCALENDAR
 * by the given time delta (in milliseconds). This keeps inline exceptions correctly
 * mapped to their corresponding RRULE-generated slots after the master series is shifted.
 *
 * Exception DTSTART/DTEND are intentionally NOT shifted: preserving the user's explicit
 * exception times (e.g., an exception at 2AM remains at 2AM after shifting the base
 * series from 1AM to 3AM). Only RECURRENCE-ID shifts so ical-expander can still
 * substitute the exception for the correct (now-shifted) occurrence slot.
 *
 * The master's EXDATEs shift too: they name instants the rule does not occur at, so left
 * behind they would exclude nothing.
 *
 * The delta is a fixed number of milliseconds, so a move across a DST change leaves the
 * values on the far side of it an hour off, RECURRENCE-IDs and EXDATEs alike.
 *
 * @param ics - Master VCALENDAR ICS containing inline exception VEVENTs
 * @param deltaMs - Time delta in milliseconds (positive = forward, negative = backward)
 * @returns Updated ICS string with shifted RECURRENCE-IDs
 */
export function shiftInlineExceptions(ics: string, deltaMs: number): string {
  if (deltaMs === 0) return ics;

  const ical = getICAL();
  const { root } = parseICSString(ics);

  const vcalendar = root.name === 'vcalendar' ? root : null;
  if (!vcalendar) return ics;

  // Move an instant the way the master moves: whole days for a DATE value (a 23h or 25h DST
  // delta must not truncate it into the previous day), a plain offset otherwise. A zoned value
  // stays in its zone, because the property keeps its TZID parameter and a UTC value under a
  // TZID is malformed (RFC 5545 section 3.3.5).
  const shiftTime = (value: ICALTime) => {
    const asDate = value.toJSDate();
    if (value.isDate) {
      return createAllDayTime(
        new Date(
          shiftedDayStartUnix(asDate.getTime() / 1000, Math.round(deltaMs / 86400000)) * 1000
        ),
        ical
      );
    }
    const shifted = ical.Time.fromJSDate(new Date(asDate.getTime() + deltaMs), true);
    const zone = value.zone;
    const zoned = zone && zone.tzid && zone.tzid !== 'UTC' && zone.tzid !== 'floating';
    return zoned ? shifted.convertToZone(zone) : shifted;
  };

  for (const vevent of vcalendar.getAllSubcomponents('vevent')) {
    const ridProp = vevent.getFirstProperty('recurrence-id');
    if (!ridProp) {
      const exdateProps = vevent.getAllProperties('exdate');
      if (exdateProps.length) {
        for (const exProp of exdateProps) {
          const values = exProp.getValues() as ICALTime[];
          const shifted = values
            .filter((v) => v && typeof v.toJSDate === 'function')
            .map(shiftTime);
          if (shifted.length === values.length && shifted.length) {
            exProp.setValues(shifted);
          }
        }
        vevent.updatePropertyWithValue('dtstamp', nowUTC(ical));
      }
      continue;
    }

    const ridValue = ridProp.getFirstValue() as ICALTime | null;
    if (!ridValue || typeof ridValue.toJSDate !== 'function') continue;

    const newRidTime = shiftTime(ridValue);

    vevent.updatePropertyWithValue('recurrence-id', newRidTime);
    vevent.updatePropertyWithValue('dtstamp', nowUTC(ical));
  }

  return root.toString();
}

/**
 * Updates times for all occurrences of a recurring event.
 * Shifts the entire series by the delta between the original occurrence and new times.
 *
 * @param ics - The master event's ICS data
 * @param originalOccurrenceStart - The original start time of the dragged occurrence (unix seconds)
 * @param newStart - New start time for the dragged occurrence (unix seconds)
 * @param newEnd - New end time for the dragged occurrence (unix seconds)
 * @param isAllDay - Whether this is an all-day event
 * @returns The modified ICS string with shifted series times
 */
export function updateRecurringEventTimes(
  ics: string,
  originalOccurrenceStart: number,
  newStart: number,
  newEnd: number,
  isAllDay: boolean
): string {
  const { event } = parseICSString(ics);

  const currentStart = event.startDate.toJSDate().getTime();

  if (isAllDay) {
    // Shift the master start by the same whole days the occurrence moved, then span the new
    // duration. Shifting by day count (not a ms delta) keeps the series on midnight across a
    // spring-forward day, where a 23h ms shift would leave the date unchanged and the series
    // silently wouldn't move.
    const days = calendarDaysBetween(
      calendarDateFromUnix(originalOccurrenceStart),
      calendarDateFromUnix(newStart)
    );
    const spanDays = calendarDaysBetween(
      calendarDateFromUnix(newStart),
      calendarDateFromUnix(newEnd)
    );
    const newMasterStart = shiftedDayStartUnix(currentStart / 1000, days);
    return updateEventTimes(ics, {
      start: newMasterStart,
      end: shiftedDayStartUnix(newMasterStart, spanDays),
      isAllDay,
    });
  }

  // Shift the master start by the occurrence's move delta, then apply the new duration, so a
  // resize (which changes newEnd relative to newStart) actually changes the whole series.
  const deltaMs = (newStart - originalOccurrenceStart) * 1000;
  const durationMs = (newEnd - newStart) * 1000;
  const newMasterStart = currentStart + deltaMs;
  return updateEventTimes(ics, {
    start: newMasterStart / 1000,
    end: (newMasterStart + durationMs) / 1000,
    isAllDay,
  });
}

/**
 * Checks if an event has recurrence rules (RRULE or RDATE)
 */
export function isRecurringEvent(ics: string): boolean {
  const { root } = parseICSString(ics);
  const vevent = root.name === 'vevent' ? root : root.getFirstSubcomponent('vevent');

  if (!vevent) {
    return false;
  }

  return !!(vevent.getFirstPropertyValue('rrule') || vevent.getFirstPropertyValue('rdate'));
}

/**
 * Gets information about the recurrence pattern
 */
export function getRecurrenceInfo(ics: string): RecurrenceInfo {
  const { root } = parseICSString(ics);
  const vevent = root.name === 'vevent' ? root : root.getFirstSubcomponent('vevent');

  if (!vevent) {
    return { isRecurring: false };
  }

  const rrule = vevent.getFirstPropertyValue('rrule');
  if (!rrule) {
    return { isRecurring: false };
  }

  // rrule is an ICAL.Recur object when present
  const recur = rrule as InstanceType<ICAL['Recur']>;
  return {
    isRecurring: true,
    rule: recur.toString(),
    frequency: recur.freq,
  };
}

/**
 * Adds an EXDATE to a recurring event to exclude a specific occurrence.
 * Used when deleting a single occurrence of a recurring event.
 * Preserves the original event's timezone.
 *
 * @param ics - The master event's ICS data
 * @param occurrenceStart - The start time of the occurrence to exclude (unix seconds)
 * @param isAllDay - Whether this is an all-day event
 * @returns The modified ICS string with the EXDATE added
 */
export function addExclusionDate(ics: string, occurrenceStart: number, isAllDay: boolean): string {
  const ical = getICAL();
  const { root, event } = parseICSString(ics);

  // Get the VEVENT component
  const vevent = root.name === 'vevent' ? root : root.getFirstSubcomponent('vevent');

  if (!vevent) {
    throw new Error('Invalid ICS: no VEVENT component found');
  }

  // Get the original timezone from the event to preserve it
  const originalZone = event.startDate?.zone;

  // Create EXDATE time from occurrence start (preserve timezone)
  const occurrenceDate = new Date(occurrenceStart * 1000);
  const exdateTime = createICALTime(occurrenceDate, isAllDay, ical, originalZone);
  addExdateProperty(vevent, exdateTime, ical, originalZone);

  // Update DTSTAMP to indicate modification
  vevent.updatePropertyWithValue('dtstamp', nowUTC(ical));

  return root.toString();
}

/**
 * Cancels the occurrence an inline exception VEVENT overrides: removes the VEVENT and excludes
 * its slot on the master, in one write. Removing the VEVENT alone hands the slot back to the
 * RRULE; the EXDATE is what cancels it.
 *
 * This is the only way to remove such an occurrence. The master and every exception share one
 * calendar resource and the sync engine deletes by resource, so DestroyEventTask on an
 * exception row takes the whole series with it (Foundry376/Mailspring-Sync#125 makes the
 * engine refuse that form).
 *
 * @param ics - The master event's ICS data, including its inline exception VEVENTs
 * @param recurrenceId - The RECURRENCE-ID of the exception to cancel, as stored on the row
 * @returns The modified ICS, or the input unchanged when no exception matches
 */
export function removeInlineException(ics: string, recurrenceId: string): string {
  const ical = getICAL();
  const { root } = parseICSString(ics);

  const vcalendar = root.name === 'vcalendar' ? root : null;
  if (!vcalendar) return ics;

  // The row stores the RECURRENCE-ID as written ('20260302T060000Z', '20260302', or wall-clock
  // text whose zone lives only on the property), so a zoned value is matched as text.
  const wanted = recurrenceId.replace(/[-:]/g, '');
  const utc = /^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})Z$/.exec(wanted);
  const wantedMs = utc ? Date.UTC(+utc[1], +utc[2] - 1, +utc[3], +utc[4], +utc[5], +utc[6]) : null;

  let master: ICALComponent | null = null;
  let exception: ICALComponent | null = null;
  let slot: ICALTime | null = null;
  for (const vevent of vcalendar.getAllSubcomponents('vevent')) {
    const ridProp = vevent.getFirstProperty('recurrence-id');
    if (!ridProp) {
      if (!master) master = vevent;
      continue;
    }
    const ridValue = ridProp.getFirstValue() as ICALTime | null;
    if (!ridValue || typeof ridValue.toJSDate !== 'function') continue;
    const sameText = ridValue.toICALString() === wanted;
    const sameInstant = wantedMs !== null && ridValue.toJSDate().getTime() === wantedMs;
    if (sameText || sameInstant) {
      exception = vevent;
      slot = ridValue;
    }
  }

  if (!master || !exception || !slot) return ics;

  vcalendar.removeSubcomponent(exception);

  // The EXDATE is the RECURRENCE-ID itself: the same instant, in the same zone and form, so
  // it names the slot exactly the way the exception did.
  addExdateProperty(master, slot.clone(), ical, slot.zone);
  master.updatePropertyWithValue('dtstamp', nowUTC(ical));

  // Cancelling an occurrence is a change the guests need to see, so advance SEQUENCE the way
  // addExclusionDate does for a plain occurrence.
  const sequence = master.getFirstPropertyValue('sequence');
  if (sequence !== null) {
    master.updatePropertyWithValue('sequence', (parseInt(String(sequence), 10) || 0) + 1);
  }

  return root.toString();
}

/**
 * Sets, updates, or removes the recurrence rule (RRULE) on an event's ICS data.
 *
 * @param ics - The original ICS string
 * @param rruleString - The RRULE string (e.g., 'FREQ=DAILY'), or null/empty to remove
 * @returns The modified ICS string
 */
export function updateRecurrenceRule(ics: string, rruleString: string | null): string {
  const ical = getICAL();
  const { root } = parseICSString(ics);

  const vevent = root.name === 'vevent' ? root : root.getFirstSubcomponent('vevent');
  if (!vevent) {
    throw new Error('Invalid ICS: no VEVENT component found');
  }

  // Remove existing RRULE(s)
  vevent.removeAllProperties('rrule');

  if (rruleString) {
    // Add new RRULE
    vevent.addPropertyWithValue('rrule', ical.Recur.fromString(rruleString));
  } else {
    // Removing recurrence entirely — also clean up EXDATE and RDATE
    // which are meaningless without an RRULE (RFC 5545)
    vevent.removeAllProperties('exdate');
    vevent.removeAllProperties('rdate');
  }

  // Update DTSTAMP to indicate modification
  vevent.updatePropertyWithValue('dtstamp', nowUTC(ical));

  return root.toString();
}

type AttendeeInput = { email: string; name?: string | null; partstat?: string; role?: string };

/**
 * Brings one VEVENT's ATTENDEE properties in line with `attendees`, updating the properties
 * of guests that are already there rather than rebuilding them, so that parameters the
 * caller never sees - CUTYPE=RESOURCE on a meeting room, ROLE=CHAIR on the organizer,
 * DELEGATED-TO chains - survive an unrelated edit to the event's title.
 */
function reconcileAttendees(vevent: ICALComponent, attendees: AttendeeInput[]): void {
  const existingByEmail = new Map<string, ReturnType<typeof vevent.getFirstProperty>>();
  for (const prop of vevent.getAllProperties('attendee')) {
    const email = prop
      .getValues()
      .map(String)
      .map(emailFromParticipantURI)
      .find((v) => !!v);
    if (email) {
      existingByEmail.set(email, prop);
    }
  }

  const keep = new Set<string>();
  for (const attendee of attendees) {
    const email = attendee.email.toLowerCase();
    keep.add(email);

    let prop = existingByEmail.get(email);
    if (prop) {
      prop.setValue(`mailto:${attendee.email}`);
    } else {
      prop = vevent.addPropertyWithValue('attendee' as any, `mailto:${attendee.email}`);
      prop.setParameter('rsvp', 'TRUE');
    }
    if (attendee.name) {
      prop.setParameter('cn', attendee.name);
    }
    prop.setParameter('partstat', attendee.partstat || 'NEEDS-ACTION');
    if (attendee.role) {
      prop.setParameter('role', attendee.role);
    } else if (!prop.getParameter('role')) {
      prop.setParameter('role', 'REQ-PARTICIPANT');
    }
  }

  for (const [email, prop] of existingByEmail) {
    if (!keep.has(email)) {
      vevent.removeProperty(prop);
    }
  }
}

/**
 * Reconciles the event's ATTENDEE properties with the given guest list: guests already on
 * the event are updated in place, new ones are added, and ones no longer listed are removed.
 *
 * Updating in place matters because an ATTENDEE property carries more than an address and a
 * response. Rebuilding the list from scratch would flatten a conference room's
 * CUTYPE=RESOURCE to an ordinary person, demote the organizer from ROLE=CHAIR, and drop
 * DELEGATED-TO chains - none of which the popover that calls this has any way to supply.
 *
 * @param organizer - Used only when the event has guests but names no ORGANIZER yet. An
 *   event with guests is required to name one (RFC 5545 section 3.8.4.3), and CalDAV servers
 *   read it to decide whether to send the invitations (RFC 6638 section 3.2.1). An organizer
 *   already on the event is left alone: someone else's meeting is not ours to take over.
 */
export function updateAttendees(
  ics: string,
  attendees: Array<{ email: string; name?: string | null; partstat?: string; role?: string }>,
  organizer?: { email: string; name?: string }
): string {
  const ical = getICAL();
  const { root } = parseICSString(ics);

  const vevent = root.name === 'vevent' ? root : root.getFirstSubcomponent('vevent');
  if (!vevent) {
    throw new Error('Invalid ICS: no VEVENT component found');
  }

  reconcileAttendees(vevent, attendees);

  if (attendees.length && organizer && !vevent.getFirstProperty('organizer')) {
    const prop = vevent.addPropertyWithValue('organizer' as any, `mailto:${organizer.email}`);
    if (organizer.name) {
      prop.setParameter('cn', organizer.name);
    }
  }

  // Update DTSTAMP
  vevent.updatePropertyWithValue('dtstamp', nowUTC(ical));

  return root.toString();
}

/**
 * Strips the iTIP METHOD from a calendar object so it can be stored as an ordinary event.
 *
 * METHOD is what makes an iCalendar object a scheduling *message* rather than a calendar
 * entry (RFC 5545 section 3.7.2). An invitation arrives as METHOD:REQUEST, and PUTting that
 * to a CalDAV collection is invalid - RFC 4791 section 4.1 requires stored objects to carry
 * no METHOD, and servers are entitled to reject it.
 */
export function stripITIPMethod(ics: string): string {
  const { root } = parseICSString(ics);
  const vcalendar = root.name === 'vcalendar' ? root : null;
  if (!vcalendar) {
    return ics; // a bare VEVENT never had a METHOD to begin with
  }
  vcalendar.removeAllProperties('method');
  return vcalendar.toString();
}

/**
 * Builds the iTIP COUNTER an attendee sends to propose a different time for a meeting they
 * were invited to - Google Calendar's "Propose a new time" (RFC 5546 section 3.2.7).
 *
 * A COUNTER is the original event with the proposed DTSTART/DTEND, not a fresh event: it
 * keeps the UID and ORGANIZER so the organizer's calendar can match it to the invitation,
 * and bumps DTSTAMP so a later proposal supersedes an earlier one. Only the proposing
 * attendee is listed, since the other guests' responses are the organizer's to track and
 * echoing them back would invite the organizer's calendar to overwrite them.
 *
 * Recurrence is deliberately dropped. A counter-proposal names one specific time, so an
 * RRULE inherited from the invitation would read as "move the entire series here".
 *
 * @returns The COUNTER ICS, or null if the proposer isn't an attendee of the event.
 */
export function createCounterProposal(
  ics: string,
  options: { email: string; name?: string; start: Date; end: Date; comment?: string }
): string | null {
  const ical = getICAL();
  const { root } = parseICSString(ics);

  const vcalendar = root.name === 'vcalendar' ? root : null;
  if (!vcalendar) {
    throw new Error('Invalid ICS: expected VCALENDAR root');
  }

  // Counter the master, not one of its modified occurrences.
  const vevents = vcalendar.getAllSubcomponents('vevent') as ICALComponent[];
  const vevent = vevents.find((c) => !c.getFirstPropertyValue('recurrence-id')) || vevents[0];
  if (!vevent) {
    throw new Error('Invalid ICS: no VEVENT component found');
  }

  const target = options.email.toLowerCase();
  const mine = vevent.getAllProperties('attendee').find((prop) =>
    prop
      .getValues()
      .map(String)
      .some((v) => emailFromParticipantURI(v) === target)
  );
  if (!mine) {
    return null;
  }

  const counter = new ical.Component(['vcalendar', [], []]);
  counter.updatePropertyWithValue('prodid', '-//Mailspring//Calendar//EN');
  counter.updatePropertyWithValue('version', '2.0');
  counter.updatePropertyWithValue('calscale', 'GREGORIAN');
  counter.updatePropertyWithValue('method', 'COUNTER');

  const proposed = new ical.Component(ical.parse(vevent.toString())) as ICALComponent;

  // One proposed time, so nothing that would spread it over a series survives.
  for (const name of ['rrule', 'rdate', 'exdate', 'recurrence-id', 'attendee']) {
    proposed.removeAllProperties(name);
  }

  const attendee = proposed.addPropertyWithValue('attendee' as any, `mailto:${options.email}`);
  const cn = mine.getParameter('cn') || options.name;
  if (cn) {
    attendee.setParameter('cn', cn);
  }
  const role = mine.getParameter('role');
  if (role) {
    attendee.setParameter('role', role);
  }
  attendee.setParameter('partstat', 'TENTATIVE');

  const event = new ical.Event(proposed);
  event.startDate = ical.Time.fromJSDate(options.start, true);
  event.endDate = ical.Time.fromJSDate(options.end, true);

  // DURATION and DTEND are mutually exclusive (RFC 5545 section 3.6.1); setting endDate
  // above writes DTEND, so a DURATION carried over from the invitation must go.
  proposed.removeAllProperties('duration');

  proposed.updatePropertyWithValue('dtstamp', nowUTC(ical));
  if (options.comment) {
    proposed.updatePropertyWithValue('comment', options.comment);
  }

  counter.addSubcomponent(proposed);
  return counter.toString();
}

/**
 * Sets the PARTSTAT of a single attendee, leaving every other attendee and every
 * other parameter on the matching ATTENDEE property untouched.
 *
 * This is the CalDAV RSVP mechanism (RFC 6638 §3.2.5): an attendee responds by
 * writing their own PARTSTAT back to their copy of the event. Rewriting the whole
 * attendee list instead would discard the organizer's ROLE/CUTYPE/RSVP parameters
 * and the other attendees' responses.
 *
 * Recurring series are answered as a whole: every VEVENT in the calendar (master
 * and any inline exceptions) gets the new status.
 *
 * @returns The modified ICS string, or null if the attendee isn't in the event.
 */
export function updateAttendeeStatus(ics: string, email: string, partstat: string): string | null {
  const ical = getICAL();
  const { root } = parseICSString(ics);

  const vevents =
    root.name === 'vevent' ? [root] : (root.getAllSubcomponents('vevent') as ICALComponent[]);
  if (!vevents.length) {
    throw new Error('Invalid ICS: no VEVENT component found');
  }

  const target = email.toLowerCase();
  let matched = false;

  for (const vevent of vevents) {
    let changedThisVevent = false;
    for (const attendee of vevent.getAllProperties('attendee')) {
      const isMatch = attendee
        .getValues()
        .some((v) => emailFromParticipantURI(String(v)) === target);
      if (!isMatch) continue;

      attendee.setParameter('partstat', partstat);
      // The response has been given, so the organizer no longer needs to ask for one.
      attendee.removeParameter('rsvp');
      changedThisVevent = true;
    }
    if (changedThisVevent) {
      // Only the components that actually changed advance; a fresh DTSTAMP on an untouched
      // occurrence claims a revision that says nothing, and RFC 5546 section 3.2 breaks ties
      // at equal SEQUENCE on exactly that value.
      vevent.updatePropertyWithValue('dtstamp', nowUTC(ical));
      matched = true;
    }
  }

  return matched ? root.toString() : null;
}

/**
 * Returns the IANA timezone identifier (TZID) from the event's DTSTART, or null
 * if the event uses UTC/floating time.
 */
export function getEventTimezone(ics: string): string | null {
  const { event } = parseICSString(ics);
  const zone = event.startDate?.zone;
  if (zone && zone.tzid && zone.tzid !== 'UTC' && zone.tzid !== 'floating') {
    return zone.tzid;
  }
  return null;
}

/**
 * Updates a specific property in the event's ICS data
 */
export function updateEventProperty(
  ics: string,
  property: 'summary' | 'description' | 'location',
  value: string
): string {
  const ical = getICAL();
  const { root, event } = parseICSString(ics);

  switch (property) {
    case 'summary':
      event.summary = value;
      break;
    case 'description':
      event.description = value;
      break;
    case 'location':
      event.location = value;
      break;
  }

  // Update DTSTAMP
  const vevent = root.name === 'vevent' ? root : root.getFirstSubcomponent('vevent');
  if (!vevent) {
    throw new Error('Invalid ICS: no VEVENT component found');
  }
  vevent.updatePropertyWithValue('dtstamp', nowUTC(ical));

  return root.toString();
}
