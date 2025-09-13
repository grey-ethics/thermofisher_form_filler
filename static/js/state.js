/**
 * state.js
 * ----------------------------------------
 * Global application state + pub/sub.
 *
 * Exposes:
 * - AppState: holds { projectLevel, ticks{} }
 * - state.setProjectLevel(v), state.setTick(id, bool)
 * - state.subscribe(fn): notify on changes
 */

export class AppState {
  constructor() {
    this.projectLevel = "";   // "" means placeholder
    this.ticks = {};          // id -> boolean
    this._subs = new Set();
  }
  subscribe(fn) {
    this._subs.add(fn);
    return () => this._subs.delete(fn);
  }
  _emit() {
    for (const fn of this._subs) fn(this);
  }
  setProjectLevel(v) {
    this.projectLevel = v || "";
    this._emit();
  }
  setTick(id, val) {
    this.ticks[id] = !!val;
    this._emit();
  }
  getSnapshot() {
    return {
      projectLevel: this.projectLevel || null,
      ticks: { ...this.ticks }
    };
  }
}
