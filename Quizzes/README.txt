WHERE TO SAVE NEW QUIZZES
=========================

EASIEST WAY — drop it right here
--------------------------------
1. Save the new quiz as a .txt file directly into this "Quizzes" folder
   (not into a subject folder).
2. Name it like this:

       Subject_WhatQuiz.txt

   Examples:

       DataComs_Mod14.txt
       AltOS_Quiz2.txt
       AppDev_Quiz3.txt
       OpSy_Mod4.txt
       SocPro_Quiz3.txt
       WebDev_Mod3.txt

3. Double-click "Quiz Toolkit.cmd" in the main folder.
4. Type 2 and press Enter.

The toolkit reads the word before the first "_" (for example "DataComs"
in "DataComs_Mod14.txt"), moves the file into Quizzes\DataComs\, converts
it to Excel, opens the output folder, and copies that folder path to your
clipboard.

SAVE DIRECTLY IN A SUBJECT FOLDER (optional)
--------------------------------------------
You can also save the file into its subject folder yourself:

    Quizzes\DataComs\
    Quizzes\AltOS\
    Quizzes\AppDev\
    Quizzes\OpSy\
    Quizzes\SocPro\
    Quizzes\WebDev\

Then choose 1 in the toolkit and press Enter for the newest quiz.

IF THE FILE IS NOT IN THE RIGHT FORMAT
--------------------------------------
The toolkit checks for "Question:" lines. If it cannot find them, it
moves the file to the "Needs-Fixing" folder in the main directory instead
of converting it. Read Needs-Fixing\README.txt for the fix.

IF THE NAME HAS NO SUBJECT PREFIX
---------------------------------
A file like "Mod14Raw.txt" has no subject before the "_", so the toolkit
moves it to Quizzes\_Unsorted\ and tells you. Rename it like
"DataComs_Mod14.txt" and run option 2 again.
