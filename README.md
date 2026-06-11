# password-generator

Here you can insert how ever many letters, symbols and numbers to create a password.

The more characters you use the hard it is to break in your private information.

Thank you for visiting my page

---

## 12-Week Health Plan Calendar

A companion script generates an Apple Calendar import file for a personalised 12-week health and fitness plan.

### Usage

```bash
python generate_health_calendar.py             # starts on next Monday
python generate_health_calendar.py 2026-06-16  # specific start date (Monday recommended)
```

This creates `health_plan_12_week.ics` containing:

- **36 gym sessions** (Mon/Wed/Fri, 10am) with full exercise lists and 30-min alerts — Workouts A, B, C cycling through the week, progressing from 2 sets (weeks 1–2) to full sets with ACL-safe guidelines
- **Daily habit reminders**: protein breakfast (7am), supplements (8:30am), post-dinner walk (7:30pm), magnesium/wind-down (10:30pm) — all recurring for 84 days
- **Bi-weekly target updates**: step goals, takeaway limits, and cardio duration as targets ramp up each phase
- **Phase transition markers** at weeks 1, 5, and 9 (Foundation → Build → Consolidate)
- **Weekly Sunday check-in** at 6pm with a 3-question review prompt
- **One-time medical actions**: message prescriber about Elvanse timing, book HbA1c recheck, get physio sign-off

### Importing into Apple Calendar

**Mac:** File → Import → select `health_plan_12_week.ics`

**iPhone:** AirDrop the `.ics` file to your phone, then tap it to open in Calendar

All events are added under a single **"12-Week Health Plan"** calendar. Delete that calendar to remove all events at once.
