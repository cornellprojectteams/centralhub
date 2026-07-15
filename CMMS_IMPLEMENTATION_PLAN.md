# Ops CMMS: phased implementation plan

Goal: upgrade the tooling / equipment / purchasing sheets into a lightweight CMMS
(Limble-style: assets, preventive maintenance, parts, work orders, purchasing) on the
existing $0 Google Sheets + Apps Script stack.

## Core idea

The Space Status engine already does the hard part: it turns a row into a tracked,
assignable, notify-and-remind action item with a Mark complete flow, and surfaces open
items in the Operations Team feed and Command Center.

So every new capability below is a **generator that appends a row to the existing
`Form Responses` sheet** and calls the same engine. No new tracking system. A preventive
maintenance task coming due and a restock request are just action items with a different
Type.

## Three new columns on `Form Responses`

Added once, used by all phases. The engine already ignores unknown columns, so this is safe.

| Column | Purpose |
|--------|---------|
| `Type` | Issue (default) / PM / Restock / Purchase. Lets the feed and Command Center group and filter. |
| `Assignee` | An email for items owned by a person (e.g. Noah), not just a team. Added to the recipients and shown on the card. |
| `Source` | Back-link to the row that generated the item, e.g. `PM:14` or `INV:First aid`. Used to update the source when the item is completed. |

## Three new tabs

**Equipment** (the asset registry)

| Asset ID | Name | Category | Location | Owning team | Status | Installed | Notes |
|----------|------|----------|----------|-------------|--------|-----------|-------|

**Inventory** (consumables / parts)

| Item | Location | On hand | Reorder point | Unit | Supplier | Product link | eShop info | Last restocked |
|------|----------|---------|---------------|------|----------|-------------|-----------|----------------|

**PM Schedule** (recurring tasks)

| PM ID | Task | Asset or area | Owning team | Frequency (days) | Last done | Next due | Assignee | Instructions |
|-------|------|---------------|-------------|------------------|-----------|----------|----------|-------------|

## Shared helper (built with Phase 2, reused by Phase 3)

`createActionItem_(fields)` in the Apps Script:
1. Appends a row to `Form Responses` with Team, Type, Required action, Details, Current
   Status (severity, drives the deadline), Assignee, Source.
2. Generates the Issue token and stamps `Notified at` (so it is live).
3. Calls the existing `sendNotification_` so the email, deadline, and reminders all work.

This is the one entry point every generator uses. It is the whole reason the CMMS is cheap.

---

## Phase 1: Asset + inventory registry

What it delivers: clean, normalized data with location and ownership. No automation yet.
This is the foundation the other phases read from.

Steps
1. Create the `Equipment` and `Inventory` tabs with the columns above.
2. Migrate the current "Tools, purchasing & locks" sheet into them (equipment rows into
   Equipment, consumables into Inventory). Fill Location and Owning team for every asset.
3. Set a Reorder point on the consumables that matter (first aid, filters, common spares).
4. Update the hub: point Admin > Inventory at these tabs. Optionally add a read-only web
   view (reuse the `?view=` web app pattern) to browse equipment by team or location.

Done when: every piece of equipment has a location and an owner, and stocked consumables
have a reorder point.

Effort: mostly data cleanup. Little to no new code.

---

## Phase 2: Preventive maintenance scheduler

What it delivers: recurring tasks (monthly filter check, quarterly inspection) create action
items in the Ops feed on their own, and recur when completed.

Steps
1. Create the `PM Schedule` tab. One row per recurring task, with Frequency in days,
   Next due, Assignee, and Instructions.
2. Add `createActionItem_(fields)` (the shared helper).
3. Add `generatePmDue()` on a daily time trigger. For each PM row where `Next due <= today`
   and there is no open action item already pointing at it (checked via `Source = PM:<id>`),
   call `createActionItem_` with Type=PM, Team=owning team, action=Task, details=Instructions,
   Assignee, Source=`PM:<id>`, and an appropriate severity for the deadline.
4. Extend `confirmAddressed`: when a completed item has `Source` starting `PM:`, set that PM
   row's `Last done` = today and `Next due` = today + Frequency. That is the recurrence.

Done when: a due PM task appears in the feed and emails the assignee, and completing it
schedules the next one. The spray-booth-filter monthly check is exactly this.

Effort: one tab, two script functions, one trigger. Reuses the reminder cadence engine.

---

## Phase 3: Purchasing / restock workflow

What it delivers: a restock need becomes a Purchase action item that carries the product and
eShop details and is assigned to a person, completed when the order is placed.

Steps
1. Confirm the `Inventory` tab has Reorder point, Supplier, Product link, and eShop info
   (from Phase 1).
2. Restock from a check: in the on-submit handler, if a submitted survey/check flags a
   restock, look up the item in `Inventory`; if `On hand < Reorder point`, call
   `createActionItem_` with Type=Purchase, the product and eShop info in Details, Assignee,
   and Source=`INV:<item>`. This is the first-aid example.
3. Low-stock scan: add `generateLowStock()` (daily, or on inventory edit) that scans
   `Inventory` and raises a Purchase action item for anything below its reorder point that
   has no open one.
4. On the card and in the email, show the Assignee and the purchasing details. Completing
   the item means the order was placed; optionally bump `Last restocked` and clear the flag.

Done when: a restock need routes to Operations with the product link and "create an eShop
cart, assign to Noah" instructions, and Noah completes it after ordering.

Constraint: Cornell eShop has no API for us, so this stops at the cart instructions and
product link. Everything up to the order is automated; the order itself stays manual.

Effort: extend the on-submit handler, one scan function, and the assignee display.

---

## Command Center integration (small, do alongside Phase 2 and 3)

- The dashboard already lists open action items. With the `Type` column it can group or
  filter: Issues, PM due, Purchases.
- The Command Center stats endpoint can count PM-due and open Purchases as their own
  numbers, so "3 PM tasks due, 2 purchases pending" show on the Ops hero.
- Assignee lets the per-person view (e.g. Noah's queue) filter to `Assignee = me`.

## Constraints and limits

- Everything runs on the current Google + Apps Script stack. No new services.
- Apps Script email quota: 100/day on personal Gmail, 1,500 on cornell.edu. Move to the
  Cornell role account before this scales (already recommended for Space Status).
- eShop stays manual (link + cart instructions).
- Keep the source tabs (Equipment, Inventory, PM) clean; the generators trust their data.

## Build order

1. Phase 1 registry (data cleanup, unblocks everything).
2. Phase 2 PM scheduler (highest ongoing value, proves the generator pattern).
3. Phase 3 purchasing (builds on the same helper).
4. Command Center counts and the per-person queue, folded into 2 and 3.

Last updated 2026-07-14.
