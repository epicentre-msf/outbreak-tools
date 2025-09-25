Look at the Disease classes in src/msetup-codes. 

The suggestions should fit welll with the other existing/rewrited classes in src/classes

How can we factor the code to avoid redundancy and aim for efficiency? _write your response bellow_

- Reuse the new shared linelist pipeline (builder coordinator, layout strategies, worksheet preparer) so disease modules assemble the same workflow instead of duplicating orchestration.
- Centralise disease metadata access through caching adapters so repeated calls no longer scan worksheets or dictionaries multiple times per run.
- Separate disease configuration parsing from worksheet rendering by returning descriptors that downstream layout strategies consume.

What could be the improvement to those classes? _write down your suggestions bellow_

- Encapsulate Excel-specific operations (sheet activation, formatting, VB module wiring) inside dedicated helpers, letting disease logic focus on business rules and validation.
- Provide strategy interfaces per disease variant so adding a new disease becomes an implementation of a clear contract rather than branching inside monolithic procedures.
- Introduce consistent validation/error reporting via `ProjectError`, catching missing or malformed metadata early and surfacing actionable diagnostics.

WRITE DOWN HERE WHAT ARE YOUR SUGGESTIONS. DO NOT HESITATE IF
THERE IS A NEED TO CREATE NEW CLASSES. YOU IMPROVEMENTS SHOULD AIM FOR EFFICIENCY
AND REDUNDANCY, AS THOSE CLASSES ARE CRITICAL.

[OVERALL IMPROVEMENTS] _write your improvements here bellow_

- Adopt a layered architecture mirroring the new linelist design: coordinator + layout strategies + section/disease descriptor builders + worksheet preparers + context caches.
- Standardise shared utilities (normalisation, caching) across disease modules to keep comparisons reliable and maintainable.
- Back new abstractions with focused unit tests, enabling incremental migration away from legacy disease code while maintaining coverage and performance.
