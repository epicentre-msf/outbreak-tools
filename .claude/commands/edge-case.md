---
description: Add an edge case to the verify suite — asks what the behaviour is, writes the case, runs it once
argument-hint: [what the behaviour is, in a sentence]
allowed-tools: Read, Write, Edit, Bash, Glob, Grep, PowerShell, AskUserQuestion
---

Add one edge case to the behavioural verify suite.

The behaviour the owner wants covered: **$ARGUMENTS**

Read `.claude/skills/verify/SKILL.md` first for the case contract, then
`scripts/verify/lib/verify-lib.sh` for what a case is given. Read one existing
case under `scripts/verify/cases/` to match its shape — `040` is the one that
pins current behaviour and names its own gap, `030` is the one that measures a
delivered workbook's bytes.

## Work in this order

**1. Settle what the case asserts.** Turn the sentence above into one
observable claim about a generated file. If the behaviour is not yet in the
product, say so and ask the owner which they want:

- assert the **required** behaviour, so the case is red until the product
  changes, and the release is gated on it
- assert **today's** behaviour so a regression shows, with the gap written into
  the case header

Never assert the required behaviour silently when today's differs. A red
nobody expected reads as breakage, and a case that quietly asserts less than
the requirement reads as coverage that is not there.

If the sentence is too vague to produce one observable claim, ask the owner
what a failure would look like on disk. That question usually settles it.

**2. Find out what actually happens now.** Do not guess. Build the thing and
read the output, or read the source that decides it. If the behaviour needs a
fixture that does not exist (a setup with a column emptied, a workbook with a
part removed), write the fixture builder as a `.ps1` beside the case, the way
`030-blank-language-column.ps1` does it, and have it print what it changed so
the case can assert the fixture is what it asked for.

**3. Pick the number.** `NNN-slug.sh`, the next free ten after the last case,
so later insertions have room. Keep cases that share a subject adjacent.

**4. Write the case.** Against the contract in the skill. Its header carries
`# title:` and `# asserts:`, and its opening comment says why the case exists,
what coherent means for it, and what it deliberately does not claim. Follow
the three rules in the skill: allow expected errors by name, judge the output
rather than the exit code, and say what the case does not claim.

Prefer `inspect_linelist` for anything `scripts/devtools/linelist-inspect.R`
already judges, so the suite keeps one definition of a coherent linelist.
Prefer `strings_of` over opening Excel when the question is about a delivered
workbook's content.

**5. Run it once, on its own.**

```bash
bash scripts/verify/verify.sh --case <NNN>
```

A case that has never run is not a case. If it needs Excel, the watchdog
refuses to start while Excel is open — close it first, and tell the owner if
you cannot.

**6. Report.** Say what the case asserts, what it answered on this machine, and
every fixture it needs that this machine lacks. If it came out red by design,
say that in the first line so nobody reads it as new breakage.

## Keep to the suite's shape

- **Nothing ever puts up a dialog.** The suite runs unattended and the owner
  must never be typing a password into it. Never call `Unprotect()` with no
  argument — on a sheet that has a password it prompts; pass a string and it
  raises instead, which a `catch` handles. Read a workbook's own password off
  the copy you are working on (`__pass!A1`) rather than hardcoding or asking.
  Pass `Workbooks.Open` an explicit password argument, and keep
  `DisplayAlerts` and `EnableEvents` off on any Excel you create. The skill's
  "A case never puts up a dialog" section has the detail.
- Preconditions **skip**, they never fail. Every fixture is a gitignored
  binary, so a clone has none of them.
- One case, one behaviour. Two behaviours are two cases.
- Nothing reaches the network. A cached input or a skip.
- Nothing is written outside `$VERIFY_WORK` and the fixtures are read-only.

Do not edit `scripts/verify/verify.sh` to make a case fit. The runner reads
case files and nothing else; a case needing more than the lib gives should grow
the lib, and say so when reporting.
