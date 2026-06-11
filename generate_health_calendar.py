#!/usr/bin/env python3
"""
12-Week Health & Fitness Plan — Apple Calendar Generator

Generates health_plan_12_week.ics for import into Apple Calendar / Reminders.

Usage:
    python generate_health_calendar.py             # defaults to next Monday
    python generate_health_calendar.py 2026-06-16  # specific start date (Monday recommended)

Output:
    health_plan_12_week.ics  — import into Apple Calendar via File → Import (Mac)
                               or AirDrop to iPhone and tap to open in Calendar
"""

import sys
import uuid
from datetime import date, datetime, timedelta


# ── Plan data ──────────────────────────────────────────────────────────────

WORKOUTS = {
    'A': {
        'title': 'Workout A — Legs + Push + Core',
        'exercises': [
            ('Bike warm-up',          '5–8 min',     'Easy pace'),
            ('Leg press',             '3 × 10–12',   'Knees to ~90°, feet slightly high on plate, start light'),
            ('Seated hamstring curl', '3 × 10–12',   'Critical for ACL protection — never skip'),
            ('Chest press machine',   '3 × 8–12',    ''),
            ('Seated cable row',      '3 × 8–12',    ''),
            ('Plank',                 '3 × 20–40 s', ''),
        ],
    },
    'B': {
        'title': 'Workout B — Hinge + Glutes + Pull + Shoulders',
        'exercises': [
            ('Bike warm-up',               '5–8 min',   ''),
            ('Dumbbell Romanian deadlift', '3 × 8–10',  'Light, soft knees, hinge at hips'),
            ('Hip abduction machine',      '3 × 12–15', 'Glute med supports the knee'),
            ('Lat pulldown',               '3 × 8–12',  ''),
            ('Machine shoulder press',     '3 × 8–12',  ''),
            ('Glute bridge (floor)',        '3 × 12–15', ''),
            ('Standing calf raise',        '3 × 12–15', ''),
        ],
    },
    'C': {
        'title': 'Workout C — Full Body + Core + Carry',
        'exercises': [
            ('Cross-trainer warm-up',         '5–8 min',   'Low impact'),
            ('Leg press',                     '3 × 10–12', 'Same rules as Workout A'),
            ('Seated hamstring curl',         '3 × 10–12', ''),
            ('Incline dumbbell press',        '3 × 8–12',  ''),
            ('Assisted pull-up or cable row', '3 × 8–12',  ''),
            ('Dead bug',                      '3 × 8/side',''),
            ('Farmer carry',                  '3 × 30 m',  'Grip + core, totally knee-safe'),
        ],
    },
}

# (week_start, week_end, steps_per_day, max_takeaways, cardio_duration)
STEP_SCHEDULE = [
    (1,  2,  6500, 3, '10 min'),
    (3,  4,  7000, 3, '10 min'),
    (5,  6,  7500, 2, '15–20 min'),
    (7,  8,  8500, 2, '15–20 min'),
    (9,  10, 9000, 1, '20–25 min'),
    (11, 12, 9500, 1, '20–25 min'),
]

PHASES = [
    (1,  'Foundation', 'Show up. Light weights, learn form, fix meal structure, anchor wake time.'),
    (5,  'Build',      'Add weight/sets, steps to ~8,500, takeaways down to 2/week, bedtime earlier. Week 5+: add split squats + goblet squats if knee is happy.'),
    (9,  'Consolidate','Progressive overload, steps ~9,500+, takeaways 1/week, lock in routines.'),
]


# ── iCalendar helpers ──────────────────────────────────────────────────────

def _fold(line: str) -> str:
    """Fold a property line to <=75 octets per RFC 5545 (CRLF + SPACE continuation)."""
    parts = []
    current = ''
    current_bytes = 0
    for ch in line:
        ch_bytes = len(ch.encode('utf-8'))
        if current_bytes + ch_bytes > 75:
            parts.append(current)
            current = ' ' + ch
            current_bytes = 1 + ch_bytes
        else:
            current += ch
            current_bytes += ch_bytes
    if current:
        parts.append(current)
    return '\r\n'.join(parts)


def _esc(text: str) -> str:
    """Escape special characters for iCalendar TEXT property values."""
    return (
        text.replace('\\', '\\\\')
            .replace(';', '\\;')
            .replace(',', '\\,')
            .replace('\n', '\\n')
    )


def _uid() -> str:
    return f'{uuid.uuid4()}@12wk-health-plan'


def _fmt(dt: datetime) -> str:
    return dt.strftime('%Y%m%dT%H%M%S')


def _fmt_date(d: date) -> str:
    return d.strftime('%Y%m%d')


def _valarm(minutes_before: int) -> list:
    return [
        'BEGIN:VALARM',
        'ACTION:DISPLAY',
        f'TRIGGER:-PT{minutes_before}M',
        'DESCRIPTION:Reminder',
        'END:VALARM',
    ]


def _event_timed(summary, dtstart, dtend, description='', alarm_min=None, rrule=None):
    lines = [
        'BEGIN:VEVENT',
        f'UID:{_uid()}',
        f'SUMMARY:{_esc(summary)}',
        f'DTSTART:{_fmt(dtstart)}',
        f'DTEND:{_fmt(dtend)}',
    ]
    if rrule:
        lines.append(f'RRULE:{rrule}')
    if description:
        lines.append(f'DESCRIPTION:{_esc(description)}')
    if alarm_min is not None:
        lines.extend(_valarm(alarm_min))
    lines.append('END:VEVENT')
    return lines


def _event_allday(summary, d, description='', alarm_min=None):
    lines = [
        'BEGIN:VEVENT',
        f'UID:{_uid()}',
        f'SUMMARY:{_esc(summary)}',
        f'DTSTART;VALUE=DATE:{_fmt_date(d)}',
        f'DTEND;VALUE=DATE:{_fmt_date(d + timedelta(days=1))}',
    ]
    if description:
        lines.append(f'DESCRIPTION:{_esc(description)}')
    if alarm_min is not None:
        lines.extend(_valarm(alarm_min))
    lines.append('END:VEVENT')
    return lines


# ── Event builders ─────────────────────────────────────────────────────────

def _workout_desc(key: str, week: int) -> str:
    w = WORKOUTS[key]
    lines = []
    if week <= 2:
        lines.append('Weeks 1–2: 2 sets only, deliberately light.')
        lines.append('Goal is attendance and form — not effort. You should finish feeling "I could have done more."')
    else:
        lines.append('3 sets. Increase weight only when you hit the top of the rep range on ALL sets with ~2 reps still in the tank.')
    if week >= 5:
        lines.append('Week 5+: add split squats (shallow depth, hold support) + goblet squats (high box) if knee is happy.')
    lines.append('')
    lines.append('KNEE RULES: pain <=3/10 that settles in 24h = OK. Control the lowering phase: 2–3 sec down every rep.')
    lines.append('')
    lines.append('Exercises:')
    for ex, reps, note in w['exercises']:
        row = f'  {ex}: {reps}'
        if note:
            row += f'  [{note}]'
        lines.append(row)
    return '\n'.join(lines)


def gym_events(start: date) -> list:
    events = []
    rotation = ['A', 'B', 'C']
    for week in range(1, 13):
        week_start = start + timedelta(weeks=week - 1)
        for day_offset, key in zip([0, 2, 4], rotation):
            day = week_start + timedelta(days=day_offset)
            dtstart = datetime(day.year, day.month, day.day, 10, 0)
            dtend = dtstart + timedelta(minutes=55)
            title = WORKOUTS[key]['title']
            events.append(_event_timed(
                f'Week {week}: {title}',
                dtstart, dtend,
                description=_workout_desc(key, week),
                alarm_min=30,
            ))
    return events


def daily_habit_events(start: date) -> list:
    events = []

    def daily(h, m, summary, desc):
        dtstart = datetime(start.year, start.month, start.day, h, m)
        events.append(_event_timed(
            summary, dtstart, dtstart + timedelta(minutes=15),
            description=desc,
            alarm_min=0,
            rrule='FREQ=DAILY;COUNT=84',
        ))

    daily(7, 0,
          'Protein breakfast — before Elvanse kicks in',
          'Target: 40–50g protein before appetite disappears.\n'
          'E.g. 3–4 eggs + 2 wholegrain toast + Greek yoghurt, or oats + protein shake.\n'
          'Bank protein early — mid-morning appetite will drop.')

    daily(8, 30,
          'Supplements: D3+K2 + fish oil (with food)',
          'Take with breakfast or any meal containing fat.\n'
          'Nothing at these doses conflicts with Elvanse.')

    daily(19, 30,
          'Post-dinner walk (10–15 min)',
          'Walk after your two biggest meals — this blunts the glucose spike more than almost anything else.\n'
          'Counts towards your daily step target.\n'
          'Best single habit for your HbA1c.')

    daily(22, 30,
          'Wind-down + magnesium bisglycinate',
          'Take magnesium now (mildly relaxing).\n'
          'Dim lights, screens to night mode.\n'
          'Target bedtime: shift 15–20 min earlier each week toward 12:00–12:30am.')

    return events


def step_target_events(start: date) -> list:
    events = []
    for wk_start, wk_end, steps, takeaways, cardio in STEP_SCHEDULE:
        d = start + timedelta(weeks=wk_start - 1)
        summary = f'Weeks {wk_start}–{wk_end}: {steps:,} steps/day · max {takeaways} takeaways'
        desc = (
            f'Step target: {steps:,}/day\n'
            f'Takeaways: max {takeaways} this block\n'
            f'Cardio (after gym or one off-day): {cardio}\n'
            f'\nKnee-safe cardio options: stationary bike, incline treadmill walk, cross-trainer, rower.'
        )
        events.append(_event_allday(summary, d, description=desc, alarm_min=480))
    return events


def phase_events(start: date) -> list:
    events = []
    for week_num, name, desc in PHASES:
        d = start + timedelta(weeks=week_num - 1)
        events.append(_event_allday(f'Phase: {name} begins (Week {week_num})', d, description=desc, alarm_min=480))
    return events


def sunday_review_events(start: date) -> list:
    events = []
    days_until_sunday = (6 - start.weekday()) % 7
    if days_until_sunday == 0:
        days_until_sunday = 7
    first_sunday = start + timedelta(days=days_until_sunday)
    for i in range(12):
        sunday = first_sunday + timedelta(weeks=i)
        dtstart = datetime(sunday.year, sunday.month, sunday.day, 18, 0)
        events.append(_event_timed(
            f'Week {i + 1} Sunday review (10 min)',
            dtstart, dtstart + timedelta(minutes=10),
            description=(
                'Count your ticks from this week.\n'
                '1. One thing that worked\n'
                '2. One thing to fix next week\n'
                '3. Set gym days in calendar for next week\n'
                '\n80% adherence for 12 weeks beats 100% for 2 weeks then quitting.'
            ),
            alarm_min=0,
        ))
    return events


def medical_action_events(start: date) -> list:
    hba1c_date = start + timedelta(weeks=13)
    events = []

    events.append(_event_allday(
        'ACTION: Message prescriber about Elvanse timing',
        start,
        description=(
            'Ask about moving dose to 8:00–8:30am (currently 10:30am).\n'
            'Reason: medication still active near midnight — disrupting sleep.\n'
            "Don't change timing on your own — just raise it at your next appointment."
        ),
        alarm_min=480,
    ))

    events.append(_event_allday(
        f'ACTION: Book HbA1c recheck (~{hba1c_date.strftime("%d %b %Y")})',
        start,
        description=(
            f'Book a blood test for around {hba1c_date.strftime("%d %B %Y")} (~3 months out).\n'
            'Current HbA1c: 42 mmol/mol (lower edge of prediabetic range — very reversible).\n'
            'Losing 5–7 kg + 3x/week strength training can bring this back to normal.\n'
            'Seeing the number drop is powerful motivation.'
        ),
        alarm_min=480,
    ))

    events.append(_event_allday(
        'ACTION: Get physio green-light on gym programme',
        start,
        description=(
            'Show physio the workout plan and get sign-off.\n'
            'Key exercises to confirm: leg press, hamstring curl, Romanian deadlift.\n'
            'Post-ACL (Feb 2024) — confirm range-of-motion limits.\n'
            'No running, jumping, plyometrics, or pivoting until cleared.'
        ),
        alarm_min=480,
    ))

    events.append(_event_allday(
        'HbA1c recheck — 3-month milestone',
        hba1c_date,
        description=(
            'This is your 12-week reward — seeing the number improve.\n'
            'Book this if not already done.'
        ),
        alarm_min=1440,
    ))

    return events


# ── Calendar assembly ──────────────────────────────────────────────────────

def generate_ics(start: date) -> str:
    all_lines = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//12-Week Health Plan//EN',
        'CALSCALE:GREGORIAN',
        'METHOD:PUBLISH',
        'X-WR-CALNAME:12-Week Health Plan',
    ]

    for group in [
        gym_events(start),
        daily_habit_events(start),
        step_target_events(start),
        phase_events(start),
        sunday_review_events(start),
        medical_action_events(start),
    ]:
        for event_lines in group:
            all_lines.extend(event_lines)

    all_lines.append('END:VCALENDAR')
    return '\r\n'.join(_fold(line) for line in all_lines) + '\r\n'


# ── Entry point ────────────────────────────────────────────────────────────

def _next_monday() -> date:
    today = date.today()
    days = (0 - today.weekday()) % 7
    return today + timedelta(days=days or 7)


def _parse_date(args: list) -> date:
    if len(args) >= 2:
        try:
            d = date.fromisoformat(args[1])
        except ValueError:
            print(f"Error: invalid date '{args[1]}'. Use YYYY-MM-DD format.")
            sys.exit(1)
        if d.weekday() != 0:
            print(f"Note: {args[1]} is not a Monday. The plan works best starting on a Monday.")
        return d
    return _next_monday()


if __name__ == '__main__':
    start = _parse_date(sys.argv)
    print(f'Generating 12-week plan starting {start.strftime("%A %d %B %Y")} ...')
    ics = generate_ics(start)
    out = 'health_plan_12_week.ics'
    with open(out, 'w', newline='') as f:
        f.write(ics)
    n = ics.count('BEGIN:VEVENT')
    print(f'Done. {out} ({n} events)')
    print()
    print('To import into Apple Calendar:')
    print('  Mac:    File -> Import -> select health_plan_12_week.ics')
    print('  iPhone: AirDrop the file to your phone, tap to open in Calendar')
    print()
    print("Delete the '12-Week Health Plan' calendar to remove all events at once.")
