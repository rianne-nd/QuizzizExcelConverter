FILES THAT NEED FIXING BEFORE CONVERSION
=========================================

OpSy_Mod3.txt
-------------
This file cannot be converted yet. It uses the wrong format:

  - "Question " instead of "Question:"
  - "Option " instead of "Option:"
  - "Explanation " instead of "Explanation:"
  - It has no "*" markers, so the correct answers are not marked.

HOW TO FIX IT
-------------
1. Open OpSy_Mod3.txt in Notepad.
2. Add a colon after every Question, Option, and Explanation label.
3. Put a "*" at the end of each correct option.
4. Save the file.
5. Move it into the matching subject folder under "Quizzes" and convert it.

Easier option: paste the raw text into your AI together with the contents
of "Prompt.txt" and ask it to reformat the quiz. Then save the result as a
new .txt file in the correct subject folder.
