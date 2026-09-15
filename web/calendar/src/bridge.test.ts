import assert from "node:assert/strict";
import test from "node:test";
import {
  applyHostMessage,
  parseHostMessage,
  selectedMessage,
  type CalendarAdapter,
  type CalendarObservation,
} from "./bridge.ts";

const observation: CalendarObservation = {
  presentationId: "opaque-1",
  title: "Example",
  start: "2026-09-15T10:00:00",
  end: "2026-09-15T11:00:00",
  allDay: false,
  accountLabel: "Work",
  location: "Room 1",
  isRecurring: false,
};

test("render payload replaces synthetic observations", () => {
  let rendered: CalendarObservation[] = [];
  const adapter = fakeAdapter({ replace: (value) => (rendered = value) });
  const message = parseHostMessage(
    JSON.stringify({
      version: 1,
      type: "renderObservations",
      observations: [observation],
    }),
  );
  assert.ok(message);
  applyHostMessage(message, adapter);
  assert.deepEqual(rendered, [observation]);
});

test("navigation commands use the calendar adapter", () => {
  const calls: string[] = [];
  const adapter = fakeAdapter({ navigation: (value) => calls.push(value) });
  for (const action of ["today", "previous", "next"] as const) {
    const message = parseHostMessage(
      JSON.stringify({ version: 1, type: "navigate", action }),
    );
    assert.ok(message);
    applyHostMessage(message, adapter);
  }
  assert.deepEqual(calls, ["today", "previous", "next"]);
});

test("invalid host messages fail closed", () => {
  assert.equal(parseHostMessage("not-json"), null);
  assert.equal(
    parseHostMessage({ version: 2, type: "navigate", action: "next" }),
    null,
  );
  assert.equal(
    parseHostMessage({ version: 1, type: "navigate", action: "sideways" }),
    null,
  );
});

test("selection message contains only protocol data and opaque id", () => {
  assert.deepEqual(JSON.parse(selectedMessage("opaque-1")), {
    version: 1,
    type: "observationSelected",
    presentationId: "opaque-1",
  });
});

function fakeAdapter(options: {
  replace?: (value: CalendarObservation[]) => void;
  navigation?: (value: string) => void;
}): CalendarAdapter {
  return {
    replaceObservations: (value) => options.replace?.(value),
    today: () => options.navigation?.("today"),
    previous: () => options.navigation?.("previous"),
    next: () => options.navigation?.("next"),
  };
}
