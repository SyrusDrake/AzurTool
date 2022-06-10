# AzurTool

A tool to pull info about all available ships from the wiki and create a Excel file from it. New ships will be appended without changing existing entries, so the user can keep track about the ships in their dock in the "Got?" column.

-----

## Known Issues

* ID column has to be treated as text so can't be sorted correctly by default. 
* If the user uses SoftMaker Office PlanMaker to edit their xlsx files, they cannot be read correctly, due to an incapability with the openpyxl module. A workaround can be found [here](https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1081#note_159870).
