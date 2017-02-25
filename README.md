## TableBuilder ##
TableBuilder is a collection of Microsoft Excel files with Visual Basic macros for generating and analyzing Actuarial Table data.  These were all created in Microsoft Excel 2010, and were originally written to use data retreived from the [TableManager 4.0 plugin](http://www.soa.org/professional-interests/technology/tech-table-manager.aspx).  All projects use initial sample data taken from the [Spring 2008 MLC Tables](http://www.soa.org/Files/Edu/edu-2008-spring-mlc-tables.pdf) published by the Society of Actuaries.
This project contains the following files:

----------
#### Expander.xlsm ####
This file replaces gaps in l<sub>x</sub> data with data automatically generated using uniform, exponential or harmonic assumptions.

#### TableBuilderIV.xlsm ####
This uses a table of ages and q<sub>x</sub> values, and quickly creates a table with the complete set of columns of values used on the actuarial exams.  Version 4 corrects a small error in the <sub>n</sub>E<sub>x</sub> data generating function and sports an improved user interface over version 1 (which is omitted from this repo).

#### TableBuilderV.xlsm ####
This version adds Due and Immediate options to annuities, as well as Joint Life and Last Survivor options on insurances.

#### TableBuilderVII.xlsm ####
Version 7 has significant improvements over version 5.  While the functions are the same, they run much more quickly, are more precise, the code is much more readable, and I have fixed errors with the dual life A<sub>x:x+n</sub> functions.  I rooted out a lot of errors in my programming and rewrote major sections of code.

#### TableBuilderVIII.xlsm ####
Table Builder VIII has only minor improvements over Table Builder VII, namely enabling the Column Auto-Width option, removing the notice of future features, and a few clean-ups to the code.

----------
#### randwalk2.xlsm ####
This program simulates a user chosen number of random walks with a user chosen number of steps each, sorts them in ascending order, and then creates a bar chart that visualizes a resultant curve that takes on the same shape of a histrogram of the same data (which can be visually compared to an histogram of the normal distribution).  Using a larger number of walks than the suggested 100 is enough to create satisfactorily convincing graphs.
To use this program, simply run the macro *RandomWalk*.
For more information on the mathematics simulated by this program, read *Ubbo F. Wiersema's* [Brownian Motion Calculus](http://www.amazon.com/Brownian-Motion-Calculus-Ubbo-Wiersema/dp/0470021705).

#### randwalk3.xlsm ####
Produced in the same afternoon as Random Walk II, this program improves on Random Walk II by finding the theoretical standard deviation and compares the amount of data occurring between one, two and three theoretical standard deviations of the mean (assumed to be zero) and compares it to the proportions of the actual count.  This provides a slightly better means of testing the claim of normality than a simple visual inspection.
This program depends on randwalk2 being in the same directory at time of execution.

