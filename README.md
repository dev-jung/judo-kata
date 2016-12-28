# judo-kata
Generating evaluation sheets and processing the judges verdicts in judo kata contests

This software package consists of various excel files with embedded macros. 
The files offer the following functions:
* Record starter lists and judge names and generate evaluation sheets from them
* Process the data given by judges. The data are entered manually from the written evaluation sheets. (Automated transfer, e.g. using mobile applications for input, may be added, but not soon)
* Validate the evaluation data and generate the leaderboards
* Provide a statistics package in order to find patterns like "largest point spread in one technique, over all individual kata ratings"

For ease of versioning, the VBA source code of the macros is kept as separate text files next to the excel files. The source code files (named "xyz.vba.txt") are _not necessary_ for the application of the software package.

If a reasonable GUI mechanism for the input can be found, the excel solution will be replaced.