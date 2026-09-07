import { Calendar } from "fullcalendar";
import dayGridPlugin from "fullcalendar/daygrid";
import listPlugin from "fullcalendar/list";
import timeGridPlugin from "fullcalendar/timegrid";
import "fullcalendar/skeleton.css";
import "./styles.css";

const calendarElement = document.querySelector<HTMLDivElement>("#calendar");

if (!calendarElement) {
  throw new Error("Calendar host element was not found.");
}

const calendar = new Calendar(calendarElement, {
  plugins: [dayGridPlugin, timeGridPlugin, listPlugin],
  initialView: "dayGridMonth",
  height: "100%",
  headerToolbar: false,
});

calendar.render();
