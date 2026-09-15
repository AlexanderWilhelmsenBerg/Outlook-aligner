export const protocolVersion = 1;

export interface CalendarObservation {
  presentationId: string;
  title: string;
  start: string;
  end: string;
  allDay: boolean;
  accountLabel: string;
  location: string;
  isRecurring: boolean;
}

export type HostMessage =
  | {
      version: 1;
      type: "renderObservations";
      observations: CalendarObservation[];
    }
  | { version: 1; type: "navigate"; action: "today" | "previous" | "next" };

export interface CalendarAdapter {
  replaceObservations(observations: CalendarObservation[]): void;
  today(): void;
  previous(): void;
  next(): void;
}

export function parseHostMessage(raw: unknown): HostMessage | null {
  let value: unknown = raw;
  if (typeof value === "string") {
    try {
      value = JSON.parse(value) as unknown;
    } catch {
      return null;
    }
  }

  if (
    !isRecord(value) ||
    value.version !== protocolVersion ||
    typeof value.type !== "string"
  ) {
    return null;
  }

  if (value.type === "renderObservations") {
    if (
      !Array.isArray(value.observations) ||
      !value.observations.every(isObservation)
    ) {
      return null;
    }
    return {
      version: 1,
      type: "renderObservations",
      observations: value.observations,
    };
  }

  if (
    value.type === "navigate" &&
    (value.action === "today" ||
      value.action === "previous" ||
      value.action === "next")
  ) {
    return { version: 1, type: "navigate", action: value.action };
  }

  return null;
}

export function applyHostMessage(
  message: HostMessage,
  calendar: CalendarAdapter,
): void {
  if (message.type === "renderObservations") {
    calendar.replaceObservations(message.observations);
    return;
  }

  switch (message.action) {
    case "today":
      calendar.today();
      break;
    case "previous":
      calendar.previous();
      break;
    case "next":
      calendar.next();
      break;
  }
}

export function readyMessage(): string {
  return JSON.stringify({ version: protocolVersion, type: "ready" });
}

export function selectedMessage(presentationId: string): string {
  return JSON.stringify({
    version: protocolVersion,
    type: "observationSelected",
    presentationId,
  });
}

export function rangeChangedMessage(
  title: string,
  start: string,
  end: string,
  visibleCount: number,
): string {
  return JSON.stringify({
    version: protocolVersion,
    type: "rangeChanged",
    title,
    start,
    end,
    visibleCount,
  });
}

function isObservation(value: unknown): value is CalendarObservation {
  return (
    isRecord(value) &&
    typeof value.presentationId === "string" &&
    value.presentationId.length > 0 &&
    typeof value.title === "string" &&
    typeof value.start === "string" &&
    typeof value.end === "string" &&
    typeof value.allDay === "boolean" &&
    typeof value.accountLabel === "string" &&
    typeof value.location === "string" &&
    typeof value.isRecurring === "boolean"
  );
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}
