# Chemistry olympiad task generator

## Description

This script allows the user to generate custom MS Word-documents with tasks on different topics: organic, inorganic, physical chemistry etc.

Libraries required: `tkinter`, `python-docx`

Current version is customized and hard-coded to generate files in Russian. 

## How to use this code

Install the libraries required and simply run this code. 

In the first window you are prompted to type in the name of the task and the amount of points for the whole task. You can also select a type of task (organic, inorganic etc., changes the functionality a little).

Below the task name you can print the main body of text for your task and add a picture. 

Next subwindow allows you to generate subtasks: you can add text, picture or generate a table in it. In addition to that, the code will generate an empty field for the student's solution which you can customize. Currently you control the height of the field in centimeters, the width is the width of the page. 

After you're finished with task and subtasks, you must click 'Add task' button, this will add the task to the queue on the second window. In the second window you can see the names of the tasks and the points awarded to each. You can also delete a task from there.

As of now, there is no option to edit the already added tasks, however you can easily do that from MS Word or any other program that works with `.docx` files. Sizes of answer fields are also editable in the final document.


