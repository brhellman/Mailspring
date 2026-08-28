import { Model, AttributeValues } from './model';
import * as Attributes from '../attributes';

/**
Public: The Calendar model represents a Calendar object.

## Attributes

`name`: {AttributeString} The name of the calendar.

`description`: {AttributeString} The description of the calendar.

This class also inherits attributes from {Model}

Section: Models
*/
export class Calendar extends Model {
  static attributes = {
    ...Model.attributes,

    name: Attributes.String({
      modelKey: 'name',
      jsonKey: 'name',
    }),
    description: Attributes.String({
      modelKey: 'description',
      jsonKey: 'description',
    }),
    readOnly: Attributes.Boolean({
      modelKey: 'readOnly',
      jsonKey: 'read_only',
    }),
    ownership: Attributes.String({
      modelKey: 'ownership',
      jsonKey: 'owner',
    }),
    color: Attributes.String({
      modelKey: 'color',
      jsonKey: 'color',
    }),
    order: Attributes.Number({
      modelKey: 'order',
      jsonKey: 'order',
    }),
  };

  public name: string;
  public description: string;
  public readOnly: boolean;
  /**
   * What the server said about who owns this calendar, from DAV:owner (RFC 3744 section
   * 5.1): 'mine' for this account's own principal, 'other' for somebody else's, and empty
   * when the server didn't answer. Used to tell our own copy of an invitation apart from a
   * copy sitting on a calendar someone shared with us.
   */
  public ownership: 'mine' | 'other' | '';
  public color: string;
  public order: number;

  constructor(data: AttributeValues<typeof Calendar.attributes>) {
    super(data);
  }
}
