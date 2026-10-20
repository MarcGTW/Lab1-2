# Lab 1–2: Examination Ticket Generator

Extend the existing Python console application into a single project with two launch modes. Preserve Lab 1 console behavior and its journal.xlsx data. Use Tkinter for Lab 2 and reuse atomic Excel persistence. Do not reinstall the development environment.

## Entry points and architecture

- `app.py`: GUI by default; `--console` starts the original console mode.
- `main.py`: original console application and shared Excel persistence helpers.
- `ticket_core.py`: independently testable source readers and result journal.
- `gui.py`: Tkinter interface; no parsing logic in event handlers.
- `start.bat`: find a Python interpreter with required dependencies and forward arguments.

## Console mode (Lab 1)

Prompt for last and first name. Trim surrounding whitespace and reject blanks. Read Unicode characters, support Backspace and immediate Escape exit in either field. Randomly select an integer from 1 to 20. Append the name, number and local timestamp to journal.xlsx immediately. Preserve old rows across restarts. Handle file locks with a user-readable message and retry the same pending record.

## Source files (Lab 2)

Read students.xlsx from the application directory. Each worksheet is a group, its name is the group key. Row 1 is a header, columns A and B contain last and first name from row 2. Skip empty or incomplete entries; log incomplete entries. Preserve empty groups.

Read tickets.docx from the same directory. A heading must match `Билет N`, with a positive integer N. Questions are numbered paragraphs `N. text`; support native Word numbered paragraphs too. Keep exactly three displayed questions. Skip and log tickets with fewer than three questions or malformed blocks. Keep source ticket numbers; sample uniformly from valid ticket objects, including noncontiguous numbers. If no valid tickets remain, show a readable error and keep the application open.

## GUI behavior

Populate a readonly group dropdown on startup. Populate a readonly student dropdown after group selection. Clear student selection when the group changes. Disable Generate until both selections are valid and tickets are loaded. Empty groups must remain selectable without exceptions.

Generate and persist before opening a modal result window. Display the student's group and full name, the source ticket number, and three questions. Escape closes only the result window. Clear the student selection and return focus to the main form, ready for the next student. Provide Reload Files and Open Journal actions. Optional demo sources must never overwrite existing files.

## Assignment and append-only results

Create results.xlsx if absent. Its first sheet has six columns: `Группа`, `Фамилия`, `Имя`, `Номер билета`, `Дата и время`, `Повтор (да/нет)`.

Student identity is the exact trimmed tuple (group, last name, first name). Search the journal from top to bottom. If a matching student exists, reuse the ticket number from the earliest matching row and append a new row with `да`. Otherwise randomly select a valid ticket and append with `нет`. Never update earlier records. The same name in different groups represents different students. If an existing student's original ticket is no longer available, show an error instead of silently assigning a different number.

Store timestamps as native Excel dates. Save after every generation using a temporary file, flush/fsync, then atomically replace the workbook. Close all file handles. If saving fails, retry the same assignment and timestamp without a second random choice. On cancel, do not show a successful result or write a row. Explain file lock and missing source errors. Log parsing errors to app.log. Preserve malformed journal files rather than overwriting them. Assume one running application instance.

## Acceptance and tests

Verify Word parsing (including invalid tickets), group lookup, empty groups, first-assignment randomness, repeated-assignment idempotency across journal reloads, and separation of group identities. Verify append-only history, immediate persistence, retry after file locks, Escape result behavior, missing source recovery, and button state transitions. Preserve all existing Lab 1 regression tests.

Run `python -m unittest discover -s tests -v`. Use temporary directories for tests; do not modify real student records. Retain Lab 1 journal.xlsx unchanged. Version-control commands, when a repository and remote are configured: create branch `lab2-ui-tickets`, review changes, commit as `Lab 2: UI-генератор билетов`, and push the branch.
