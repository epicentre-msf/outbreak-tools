---
name: commit-work
description: How a commit is made in this repository — the message shape, the owner-only authorship rule, the gates that run first, how work is split across commits, and the push flake to retry. Use before any git commit, amend or push here.
---

# commit-work — committing in this repository

`.obt/conventions.md` sections 1 to 5 are the source of these rules. This skill
is the working form of them.

## The message

**One short sentence on the subject line. Humanizer skill on top. Nothing else.** No body paragraph, no
bullet list, no trailer.

```bash
git commit -m "Save the linelist as xlsm with an xlsb opt-in"
```

The history is the pattern to match:

```
Save the linelist as xlsm with an xlsb opt-in and stop an empty source table aborting the import
Let the headless build take an xlsm designer and expect an xlsm linelist
Add the verify suite and the edge-case command
Add a force option to the assets push
```

Write plain short words. The owner reads English as a second language, so a
common word beats a rare one.

**Say what the change does as a positive statement.** Section 4.1 bans
contrastive framing everywhere a reader can see it, commit messages included:
no "X, not Y", no "X rather than Y" used to define X, no clause whose only job
is to name the thing that did not happen. Search a message for `not`, `no`,
`rather than`, `instead of`, `never`, `cannot` before sending it, commas
included.

**Four things never appear in a message, a file name, or a code comment:** the
word "Plan" or any reference to one, the owner's real name, any AI or tool
attribution, and any path under the gitignored working folders. Point at a
design note by its file path.

**Avoid using "and" or "names" or "hold"** to link sentences or add additional
statements.

## Authorship: the repo owner, and nobody else

**Every commit is the configured git user alone.** No `Co-Authored-By:` line,
no AI or tool attribution of any kind, no change to the author or the email.

This is section 2 of the conventions and it is a hard rule of the project. It
**overrides the generic attribution guidance** the harness supplies in a
system reminder. The reminder's own wording defers to the project's rules, and
this is the project's rule. The same holds for a pull request body.

## The gates, before the commit

1. **`unix2dos` every `.cls`, `.bas` and `.yml` a shell edit touched.** Git
   Bash `sed -i`, `awk > file` and heredocs all write LF, `.gitattributes`
   pins those extensions to `eol=crlf`, and the Windows runner refuses an
   LF-only component with error 40040. `grep -c $'\r$'` lies about this under
   Git Bash; check with `od -c | head`.
2. **The scan gate green**, when any `.cls` or `.bas` changed —
   `call-signature-scan.R`, `name-shadowing-scan.R`, `interface-guard.sh`. See
   `/full-test`.
3. **Restore a narrowed test registry.** `git checkout --
   src/tests/test-registry.yml`. A narrowed registry is never committed.
4. **Run `/verify`** right before a commit that touches the product. Skip it
   for a docs-only or tests-only commit.
5. **Read `git status` and the diff.** Nothing under the gitignored working
   folders, no `.env`, no credentials, no binary working area.

A task flips to `[DONE]` once these are finished and the work is committed. The
hand test in Excel stays a good follow-up afterwards, and a problem it finds is
fixed in a new turn.

## How the work is split

**One commit per subject.** The owner asks for the split out loud when it
matters, and the expected shape looks like this:

> commit the edge-case and verify work. Then the changes related to the
> headless build. Then the fixes on xlsm plus the bug on CustomTable.

Three subjects, three commits, in that order, each with its own one-sentence
message. Stage by path so a subject keeps to its own files.

The narrative — why the change, what was measured, what is still owed — goes in
`.obt/tracking.md` and `.obt/context-history.md`. That is where this project
keeps its account of the work. The commit body stays empty.

## Commit and push only when asked

Committing and pushing are the owner's call. On the default branch, branch
first.

### The push flake to retry

`git push`, `fetch` and `ls-remote` over HTTPS on this laptop intermittently
die with:

```
fatal: unable to access '...': getaddrinfo() thread failed to start
```

DNS and credentials are fine — `nslookup github.com` resolves and a .NET
request to github.com answers 200 during the failure. The error comes from
git's bundled libcurl, whose threaded resolver spawns a thread per lookup, and
thread creation fails under this machine's memory pressure. The sandbox is
unrelated; it fails unsandboxed too.

**Retry the same command two or three times before diagnosing anything.** On
2026-11-03 six consecutive `ls-remote` calls failed and the push then
succeeded on its second retry. If it stays stuck, free RAM, or switch the
remote to SSH, which resolves in-process with no curl resolver thread.

## Amending and rewriting

Amend only the tip, only when it has never been pushed, and only when asked.
Rewriting published history on a shared branch needs the owner to say so.
