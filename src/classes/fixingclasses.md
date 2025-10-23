
We still have few things missing for the SetupTranslation class.
There is a SetupTranslationOld.cls file in temp/ folder that has the legacy code
that we were trying to port. A few things are still missing; implement them and 
add [done] tag once it is done.

In SetupTranslationsTable:

1-  [done] Use BetterArray instead of collections whenever it is necessary, for coherence in the overall project.

2- [done] Expose a ResetSequence method that will set the target cell for sequence value to 0 (usually when you open the workbook).

3- [done] After updating labels in the translation table, you should delete non existing labels in the translation table, those that have been removed in the corresponding table.

4- [done] Expose a NumberOfMissing() that will print the number of missing labels for ALL the languages except the first language  (key language) in a msgbox. You can test it by using an internal displayPrompts boolean value.


Correct the tests to take in account this new logics. Follow closely instructions.md, in particular,
do not Edit any "TestDictionary" related file, especially DictionaryTestFixture. Never ever touch this file.
Pay also extremely attention to namming convention when updating the class, and focus your update to only 
those new features to avoid breaking stuff. Write compact and efficient code, aim for the smallest number of lines  that achieve
the final desirable results.
