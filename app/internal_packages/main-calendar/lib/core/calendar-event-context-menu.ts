import { ICSParticipantStatus, localized } from 'mailspring-exports';
import { EventOccurrence } from './calendar-data-source';
import { canRespondToEvent, myParticipationStatus, respondToCalendarEvent } from './calendar-rsvp';

type TemplateItem =
  | { label: string; click: () => void; type?: 'checkbox'; checked?: boolean; enabled?: boolean }
  | { type: 'separator' };

const isSeparator = (item: TemplateItem) => (item as { type?: string }).type === 'separator';

/**
 * Removes the separators left behind when optional items are dropped from a menu.
 *
 * Filtering in one pass isn't enough: dropping a trailing separator can leave the one before
 * it trailing in turn, and a single predicate evaluated against the original array never
 * sees that. Collapse runs first, then trim both ends.
 */
function trimSeparators(items: TemplateItem[]): TemplateItem[] {
  const collapsed = items.filter(
    (item, i) => !(isSeparator(item) && i > 0 && isSeparator(items[i - 1]))
  );
  let start = 0;
  let end = collapsed.length;
  while (start < end && isSeparator(collapsed[start])) start++;
  while (end > start && isSeparator(collapsed[end - 1])) end--;
  return collapsed.slice(start, end);
}

interface CalendarEventContextMenuOptions {
  occurrence: EventOccurrence;
  /** Whether the event's calendar refuses writes, which hides everything that changes it. */
  readOnly: boolean;
  /**
   * Whether the event's details may be revised here: a writable calendar *and* a meeting we
   * organise. Only the organizer revises a meeting (RFC 5546 section 2.1.4), so an attendee
   * gets "View Event" and answers with the RSVP items instead. Deleting stays governed by
   * `readOnly` alone - removing our own copy of someone else's meeting is ours to do.
   */
  editable: boolean;
  onOpen: () => void;
  onDelete: () => void;
}

/**
 * The right-click menu for an event in the calendar.
 *
 * Everything here is reachable another way - the popover edits, the Delete key removes, the
 * invitation email answers - but only if you already know where to look. Right-clicking the
 * thing you want to change is where people look first.
 */
export class CalendarEventContextMenu {
  private opts: CalendarEventContextMenuOptions;

  constructor(opts: CalendarEventContextMenuOptions) {
    this.opts = opts;
  }

  private rsvpItems(): TemplateItem[] {
    const { occurrence, readOnly } = this.opts;
    if (!canRespondToEvent(occurrence)) {
      return [];
    }

    const current = myParticipationStatus(occurrence);
    const actions: [ICSParticipantStatus, string][] = [
      ['ACCEPTED', localized('Accept')],
      ['TENTATIVE', localized('Maybe')],
      ['DECLINED', localized('Decline')],
    ];

    return actions.map(([status, label]) => ({
      label,
      type: 'checkbox' as const,
      checked: current === status,
      // Answering writes our status onto the event, so a calendar we can't write is a
      // calendar we can't answer on. The reply email alone would leave the two disagreeing.
      enabled: !readOnly,
      click: () => {
        respondToCalendarEvent(occurrence, status);
      },
    }));
  }

  template(): TemplateItem[] {
    const { readOnly, editable, onOpen, onDelete } = this.opts;

    const items: (TemplateItem | null)[] = [
      {
        label: editable ? localized('Edit Event') + '...' : localized('View Event'),
        click: onOpen,
      },
      { type: 'separator' },
      ...this.rsvpItems(),
      { type: 'separator' },
      readOnly ? null : { label: localized('Delete Event'), click: onDelete },
    ];

    return trimSeparators(items.filter(Boolean) as TemplateItem[]);
  }

  displayMenu() {
    require('@electron/remote').Menu.buildFromTemplate(this.template()).popup({});
  }
}
