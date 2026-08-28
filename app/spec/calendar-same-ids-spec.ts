import { sameCalendarIds } from '../internal_packages/main-calendar/lib/core/calendar-helpers';

describe('sameCalendarIds', function () {
  it('treats two freshly-built empty lists as unchanged', function () {
    // This is the case that mattered: `config.get(key) || []` mints a new array each call,
    // so identity comparison reported a change on every update.
    expect(sameCalendarIds([], [])).toBe(true);
  });

  it('is true for the same reference', function () {
    const ids = ['a', 'b'];
    expect(sameCalendarIds(ids, ids)).toBe(true);
  });

  it('is true for equal contents in different arrays', function () {
    expect(sameCalendarIds(['a', 'b'], ['a', 'b'])).toBe(true);
  });

  it('is false when an id is added', function () {
    expect(sameCalendarIds(['a'], ['a', 'b'])).toBe(false);
  });

  it('is false when an id is removed', function () {
    expect(sameCalendarIds(['a', 'b'], ['a'])).toBe(false);
  });

  it('is false when an id is replaced', function () {
    expect(sameCalendarIds(['a', 'b'], ['a', 'c'])).toBe(false);
  });

  it('is false when the order differs, since the caller supplies a stable order', function () {
    expect(sameCalendarIds(['a', 'b'], ['b', 'a'])).toBe(false);
  });

  it('tolerates undefined on either side', function () {
    expect(sameCalendarIds(undefined, [])).toBe(true);
    expect(sameCalendarIds([], undefined)).toBe(true);
    expect(sameCalendarIds(undefined, ['a'])).toBe(false);
  });
});
