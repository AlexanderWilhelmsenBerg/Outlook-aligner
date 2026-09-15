import { Calendar } from "fullcalendar";
import dayGridPlugin from "fullcalendar/daygrid";
import classicThemePlugin from "fullcalendar/themes/classic";
import "fullcalendar/skeleton.css";
import "fullcalendar/themes/classic/theme.css";
import "fullcalendar/themes/classic/palette.css";
import {
  applyHostMessage,
  parseHostMessage,
  rangeChangedMessage,
  readyMessage,
  selectedMessage,
  type CalendarAdapter,
  type CalendarObservation,
} from "./bridge";
import "./styles.css";

const calendarElement = document.querySelector<HTMLDivElement>("#calendar");
if (!calendarElement) {
  throw new Error("Calendar host element was not found.");
}

type WebViewMessageListener = (event: MessageEvent<unknown>) => void;

type WebViewBridge = {
  addEventListener(type: "message", listener: WebViewMessageListener): void;
  postMessage(message: string): void;
};

type WebViewWindow = Window & {
  chrome?: { webview?: WebViewBridge };
};

const webview = (window as WebViewWindow).chrome?.webview;
let currentObservations: CalendarObservation[] = [];

const calendar = new Calendar(calendarElement, {
  plugins: [classicThemePlugin, dayGridPlugin],
  initialView: "dayGridMonth",
  height: "100%",
  headerToolbar: false,
  eventClick(info) {
    webview?.postMessage(selectedMessage(info.event.id));
  },
  eventDidMount(info) {
    const element = info.el;
    element.tabIndex = 0;
    element.setAttribute("role", "button");
    element.setAttribute("aria-label", info.event.title);
    element.title = info.event.title;
    element.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        webview?.postMessage(selectedMessage(info.event.id));
      }
    });
  },
  datesSet(info) {
    reportRange(info.view.title, info.start, info.end);
  },
});

const adapter: CalendarAdapter = {
  replaceObservations(observations) {
    currentObservations = observations;
    calendar.removeAllEvents();
    for (const observation of observations) {
      calendar.addEvent({
        id: observation.presentationId,
        title: observation.title,
        start: observation.start,
        end: observation.end,
        allDay: observation.allDay,
        extendedProps: {
          accountLabel: observation.accountLabel,
          location: observation.location,
          isRecurring: observation.isRecurring,
        },
      });
    }

    const view = calendar.view;
    reportRange(view.title, view.currentStart, view.currentEnd);
  },
  today: () => calendar.today(),
  previous: () => calendar.prev(),
  next: () => calendar.next(),
};

function reportRange(title: string, start: Date, end: Date): void {
  const visibleCount = currentObservations.filter((observation) => {
    const observationStart = new Date(observation.start);
    const observationEnd = new Date(observation.end);
    return observationStart < end && observationEnd >= start;
  }).length;
  const startIso = start.toISOString();
  const endIso = end.toISOString();
  const message = rangeChangedMessage(title, startIso, endIso, visibleCount);
  webview?.postMessage(message);
}

webview?.addEventListener("message", (event) => {
  const message = parseHostMessage(event.data);
  if (message) {
    applyHostMessage(message, adapter);
  }
});

calendar.render();
webview?.postMessage(readyMessage());
