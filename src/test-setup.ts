/* global globalThis, window */
// jsdom has no ResizeObserver, which Fluent's MessageBar needs.
class ResizeObserverStub {
  observe() {}
  unobserve() {}
  disconnect() {}
}
(globalThis as { ResizeObserver?: unknown }).ResizeObserver ??= ResizeObserverStub;
(window as { ResizeObserver?: unknown }).ResizeObserver ??= ResizeObserverStub;
