# Facebook_HTML_EXCEL_converter
This would convert the provided Facebook HTML files into EXCEL file. Depending on the HTML file, it should included some or all comments and replies.\
This could be applied to both classic/old Layout and New Layout of Facebook HTML file.

1. Expand comments and replies. Click all see more. [Expand Facebook Comments and Replies](https://github.com/kionchan7/Facebook_expandCommentReplies_seemore)
2. Save the webpage as `Webpage, Complete (*.html)`
3. Use this python file to convert html file into xlsx file. (currently here)

Here's an example of how to use it:
* Place the html files in the same folder as converter.py
* In Ubuntu: `python3 converter.py`

After converting, the excel sheet will display in this format\
Post Contents(A1)\
Name(A2) Comment or Name (B2) @Name Comment or Comment (C2)\

Upcoming updates:
* The option of extracting names who have commented in a Facebook post (from an HTML file).
* Commenting for this program