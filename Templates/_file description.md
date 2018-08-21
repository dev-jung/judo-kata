# description of the files 

Contained in here are 5 Excel files which make up the core of this package. 
The PFD documents are here for convenience, the *.vba.txt files are here in order to diff source code - they are simply copied from the excel files.

* __Blank.xlsm__: Contains the blank evaluation sheets. The file contains the sheets for all kata, and an evaluation sheet for each kata. 
                  The excel file is static, contains no significant macro code.
* __KoshikiUra.xlsm__: Just a static file with a large-scale evaluation matrix for Koshiki ura group. 
                       This is handy for the judges because ura is fast ;)
* __List.xlsm__: Contains the template for the starter lists for a kata tournament.
                 The excel file contains the macro code to generate evaluation sheets for each starting pair of the starting lists.
				 It also contains the configuration parameters for the tournament, as well as the I18N/N10 configuration
* __Results.xlsm__: Contains the template for result tables.
                    The excel file contains the macro code to import all evaluation sheets for a kata, sort the starters according to their score, and pretty print the results as a table.
* __Statistics.xlsm__: Contains a template to evaluate the judges and to pinpoint major deviation of scores.
                       The excel file contains the macro code to import all evaluation results of the tournament.
					   This is meant to be used by the supervisors only.

