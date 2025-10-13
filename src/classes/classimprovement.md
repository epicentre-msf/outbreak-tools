There are issues with the Event handler in checkingOutput class in src/classes/general.
We want to  only stick to string handler that clears worksheet events and suadd only 
one worksheet_change in checkingOutput worksheet. Clean the checkingOutput to stick to only
this approach. Make sure there are no errors and no double entries of a worksheet_change event.

You should also pay extremely attention to quotes usage and when injecting the handler. The codes handler should not
assume there is a CheckingOutput class present; you should not suppose anything and only inject worksheet logic 
codes for the filtering.

Follow closely the instructions.md, do not use any scripting dictionary, and write down your improvements

Improvements:
- Injected a fully self-contained worksheet filtering handler that no longer depends on `CheckingOutput` members.
- Added code-generation helpers to remove old worksheet procedures before inserting the refreshed handler block.
- Extended `TestCheckingOutput` coverage to confirm the injected code is self-sufficient and remains single-instanced.
