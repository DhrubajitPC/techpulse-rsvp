#!/usr/bin/env node
// Builds the Teams webhook payload: an Adaptive Card summarizing the
// published digest, wrapped in the shared-secret envelope the Power
// Automate flow checks for. Printed to stdout so the workflow can pipe
// it straight into curl.
import { readFileSync } from "node:fs";
import { evalArray } from "./events.mjs";

const file = process.argv[2] ?? "site/index.html";
const BASE_URL = "https://dhrubajitpc.github.io/techpulse-rsvp";
const ICON_URL = `${BASE_URL}/icon.png`;
const ONE_DAY_MS = 24 * 60 * 60 * 1000;
const WINDOW_DAYS = 14; // two weeks of listed events, from the digest's own updated date

const src = readFileSync(file, "utf8");
const updated = src.match(/data-updated="([^"]+)"/)?.[1] ?? "unknown date";
// Link the notification to this week's own dated snapshot rather than root,
// so it keeps pointing at the right content even after root moves on.
const pageUrl = `${BASE_URL}/${updated}/`;

// Plain YYYY-MM-DD dates, parsed as UTC so calendar-day comparisons below
// aren't sensitive to the runner's local timezone.
function parseISO(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(Date.UTC(y, m - 1, d));
}

function formatDMY(iso) {
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

const WEEKDAY_FORMAT = new Intl.DateTimeFormat("en-GB", { weekday: "short", timeZone: "UTC" });
const DAY_MONTH_FORMAT = new Intl.DateTimeFormat("en-GB", { day: "numeric", month: "short", timeZone: "UTC" });

function formatDateLine(iso) {
  const d = parseISO(iso);
  return `${WEEKDAY_FORMAT.format(d)}, ${DAY_MONTH_FORMAT.format(d)}`;
}

const windowStart = parseISO(updated);
const windowEnd = new Date(windowStart.getTime() + WINDOW_DAYS * ONE_DAY_MS);

const allEvents = evalArray(src, "EVENTS").sort((a, b) => a.date.localeCompare(b.date));
const events = allEvents.filter((e) => {
  const d = parseISO(e.date);
  return d >= windowStart && d <= windowEnd;
});

function headerCell(text) {
  return { type: "TableCell", items: [{ type: "TextBlock", text, wrap: true, size: "Small", weight: "Bolder" }] };
}

function dateCell(lines) {
  return {
    type: "TableCell",
    items: lines.map((text, i) => ({
      type: "TextBlock",
      text,
      wrap: true,
      size: "Small",
      isSubtle: i > 0,
      ...(i > 0 ? { spacing: "None" } : {}),
    })),
  };
}

// A "chip" faked with a colored Container, since the real Badge element
// needs schema 1.6, which this delivery path (Power Automate -> Teams)
// fails to render at all — 1.5 is the confirmed ceiling.
function chip(tag) {
  return {
    type: "Container",
    style: "accent",
    spacing: "None",
    items: [{ type: "TextBlock", text: tag, size: "Small", weight: "Bolder", wrap: false, spacing: "None" }],
  };
}

function eventCell(name, url, tags) {
  return {
    type: "TableCell",
    items: [
      { type: "TextBlock", text: `[${name}](${url})`, wrap: true, size: "Small" },
      { type: "ColumnSet", spacing: "Small", columns: tags.map((t) => ({ type: "Column", width: "auto", items: [chip(t)] })) },
    ],
  };
}

function row(cells) {
  return { type: "TableRow", cells };
}

const table = {
  type: "Table",
  firstRowAsHeaders: true,
  gridStyle: "default",
  columns: [{ width: 1 }, { width: 3 }],
  rows: [
    row([headerCell("Date"), headerCell("Event")]),
    ...events.map((e) =>
      row([
        dateCell([formatDateLine(e.date), formatDMY(e.date)]),
        eventCell(e.name, e.url, e.tags),
      ])
    ),
  ],
};

const summaryText =
  events.length > 0
    ? `${events.length} event${events.length === 1 ? "" : "s"} in the next two weeks — updated ${formatDMY(updated)}`
    : `No events in the next two weeks — updated ${formatDMY(updated)}. See the full digest for what's further out.`;

const card = {
  type: "AdaptiveCard",
  $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
  version: "1.5",
  body: [
    {
      type: "ColumnSet",
      columns: [
        { type: "Column", width: "auto", items: [{ type: "Image", url: ICON_URL, width: "32px", height: "32px" }] },
        {
          type: "Column",
          width: "stretch",
          verticalContentAlignment: "Center",
          items: [{ type: "TextBlock", text: "TechPulse RSVP", weight: "Bolder", size: "Medium" }],
        },
      ],
    },
    { type: "TextBlock", text: summaryText, wrap: true, spacing: "Small" },
    ...(events.length > 0 ? [table] : []),
    { type: "TextBlock", text: `[Full digest](${pageUrl})`, wrap: true, spacing: "Medium" },
  ],
  msteams: { width: "full" }, // use the channel's full width instead of Teams' default narrow card
};

process.stdout.write(JSON.stringify({
  secret: process.env.TEAMS_WEBHOOK_SECRET ?? "",
  type: "message",
  attachments: [{ contentType: "application/vnd.microsoft.card.adaptive", content: card }],
}));
