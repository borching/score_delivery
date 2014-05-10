score_delivery
==============

automatically sending score information to individuals in the class

===

2014.5.10.
The current app (score.py) simply reads the information recorded in a google spredsheet in the instructor's account, 
reformulates the data for an individual group (or student) into the HTML format, and then sends the HTML
as an email to the designated group (or students).

The app reads param.py as a parameter file, in which some parameters, such as spreadsheet name, email title, etc., 
can be set up.
