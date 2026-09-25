/**
 * Smoke test: the package's modules must survive being required.
 *
 * A throw at import time takes down the whole main window rather than just this package,
 * and nothing else in the suite loads event-header, so a bad import here would otherwise
 * only show up by opening the app.
 */
describe('events package', function () {
  it('loads its entry point', function () {
    const main = require('../lib/main');
    expect(typeof main.activate).toBe('function');
    expect(typeof main.deactivate).toBe('function');
  });

  it('loads the event header', function () {
    expect(require('../lib/event-header').EventHeader).toBeDefined();
  });

  it('loads the propose-time popover from the component kit', function () {
    // Shared with the calendar package, so it lives in src/components rather than here.
    expect(require('mailspring-component-kit').ProposeTimePopover).toBeDefined();
  });
});
