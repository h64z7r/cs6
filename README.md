java c
Advanced Spreadsheets using Excel 
Software: Microsoft Excel 
Introduction to Course This course will build on the skills you have developed in the Term 1 Spreadsheets course. It will extend your proficiency using Excel, introducing you to more advanced functions, as well as features to add usability and limit errors. You should adopt a problem-solving approach to explore aspects you have not been explicitly taught. Challenge activities are provided to help you apply and extend your skills. The final exam will be designed so that it is possible to achieve a pass without doing the challenge activities, but will contain a challenge section that provides an opportunity for you to demonstrate your understanding of complex features. 
Resources used in this course Learning materials are in a variety of electronic formats - workbook, tutorials, videos and challenge activities. 






This icon indicates a link to a video, or extra explanations to help you 
This icon indicates a link to an online video. You will need to be online to view it. 

This icon indicates a link to a video that you can view offline. 

This icon indicates a Challenge activity. 
Data files     This icon indicates there is a file you have been provided with to use in the activity. You should have downloaded them all at the beginning of the course. 

Requirements 
If you do   not complete all tasks set   in   your   tutorial   class,   you   must   complete   them   before the   next tutorial.
In some cases, Challenge activities will   be   provided   for   you   to   extend   your   skills.
•               Weekly quizzes
•             Practical tasks -   you   must   complete   all   set   tasks
•             Spreadsheet   Quiz   (assessment)
•             Final   practical   exam
IMPORTANT 
Screen shots used may vary, depending on which version of 
Microsoft Office you are using. 
For this course, you should be using the version of Excel installed on 
your device (not an online version). 
All instructions are written for Windows desktop computers; 
however, instructions are provided for Mac computers where 
necessary. 
If you are using a Mac computer, a laptop computer, or a portable 
device, some features / keyboard shortcuts may be different. 
General instructions 

For each activity: 
Save all files with   your   ZID   in   front   of the   filename.
With every spreadsheet you create,   make sure   that   you   look   at   the   print   preview   :
•             All columns should   fit   within   one   page   (if   not,   change   orientation   to   landscape   OR   readjust size   of columns)
•             Where   possible avoid splitting   any   data   across   more   than   one   page   (move   the   table   OR   resize)
•             Ensure   each   chart   is   entirely visible   on   the   page   (resize   OR   move   the   chart)
•             If the column   shows   monetary values   then   apply   currency   or   accounting   format   (follow   instructions)   .    Make sure you are consistent,   using   the   same   format   across   the spreadsheet.
•             Align column   headings to   match that   of the   data   below,   eg   align   left   over text,   align   right over   numeric values.
•             Set the   number   of decimal   places   consistently   or   as   required   in   the   activity   instructions.
Advanced Excel – Topic 1Nested and Complex Formulas
A nested function is a function inside another function. The IF function   is   often   used   with   nesting. That   is, you can   nest or enclose another function inside an   IF   function.
Activity 1a - Constructing a nested function    Use the file Nested.xlsx for this   activity.
1 In column F,   we   want   to   show   total   sales only for   the South region
2 In Column G, we   want   to   show   total   sales only for   the West region
3 In column H, we   want   to   show   average   sales only for   the South region
4 In Column I, we   want   to   show   average   sales only for   the West region
5 We   will   use   absolute   references.
The logic: 
If B4 contains   the   word “South” then   add   the   values   in   the   range C4:E4.   If   it doesn’t, then   return   (display) 0.
Complete   the   spreadsheet   by   inserting   the   correct   nested   formulas   in   cells F4, G4, H4 and I4 and auto filling down   appropriately.   When   completed,   save your   file.
Note:   In the example above we   have   used the heading as   an   absolute   reference;   however,   it   is best practice to create a separate,   labelled reference   area   if you wish to   use   absolute   references.
Activity 1b - Constructing a nested IF function
Use the file Harwood Sales.xlsx for   this   activity.
25000 -> commission   =   0
40000 -> commission   = 40000*5%         85000 -> commission =   85000*10%
Harwood Sales Company   pays   its sales   people a commission   (an   extra   payment)   depending on their sales   for   the   month.
•             If   a   sales   person   sells   less   than   $30,000   of   items, they   receive   no commission;
•             if   they   sell   from   $30,000 to   $70,000, they   receive   5%   commission   ;   and
•             if   they   sell   more   than   $70,000 they   receive   a   10%   commission.
1 Add   a   column   and   label   it Commission Rate.
2 Using   the   information   above,   use   a   nested IF (an IF inside   an IF)   function   to
calculate each   person’s commission   rate   based on their sales.   Below   is   an
explanation of the syntax of this formula, and the final   formula   using   absolute   cell   reference:
3 Fill   down   this   column.
4 Add another column, and   label   it Commission.   Calculate   the   amount   of   commission   in   dollars.
5 Format the Monthly代 写Advanced Spreadsheets using Excel – Topic 1Java
代做程序编程语言 Sales and Commission columns to   currency   with   no   decimal   places.   Format the Commission Rate column to a   percentage   with   no   decimal
places.
6 Total the Monthly Sales and Commission columns.
7 Format the spreadsheet and   save   your   file.
Activity 1c - Using IF/AND 
Before commencing the following activities,   make sure you   have downloaded   and
reviewed the document IF - IF/AND - IF/OR Overview.pdf from   Moodle. This
explains and   illustrates the use of the AND and   OR   functions   with   the   IF   statement.
   Use the file Overdue.xlsx for   this   activity
An overdue account   is one that   is   late   being   paid. Accounts that are   30   days   or   more   overdue, are charged a   late fee of   $5.   Create   a   labelled   reference   area   with   this information so that you can   use absolute   references   in   your   IF   statements.
1 In the Late Fee column,   create   an   IF function   which   enters   5   for   accounts   that   are   30 days   or   more   overdue, and   0   for   accounts   that   are   less   than   30   days   overdue. Use absolute   referencing.   Format the column to currency with   2   decimal   places.
2 In the Total column,   add   the   amount   due   and   the   late   fee.
3 If   the Total is   more   than   $200 and   the   payment   is   30 days   or   more   overdue, the   account   is Urgent. Add this information to the spreadsheet   so   you   can   use   absolute   referencing.
We can also use an   IF function to display   a   result   when   2   conditions   are   both   true.   To   do   this   we   use the   IF function and the AND function together.
The AND function   returns TRUE if both the   conditions   are true.   For   instance   AND(C4>200,D4>30)   returns TRUE   if both C4   is greater than 200 AND   D4   is greater   than   30.
a         In   the Status column,   use   IF   and   AND   functions.   If   the Total is   more   than   $200   and   if the   payment   is 30 days or   more overdue, then   display   the   word   URGENT.   Otherwise leave   the   cell   blank.
The image   below should   help.   Make sure you   use   absolute   referencing.
4 Format your spreadsheet   and   save.
Activity 1d - Using nested IF or IF/AND 
Use the file Machine Parts.xlsx for   this   activity
A company that sells   machine   parts is   updating their stock   and   calculating   which   products are   most   popular.    The company   buys the   parts at cost   price   and   they   are then   sold   to   customers   plus   a   markup   %.
1 Add a column   labelled Total Cost and   calculate   the   cost   to   the   company   of   buying   the   parts.
2 Insert another column   labelled Unit Selling Price and   calculate the   selling   price   of   each   item   including   the   mark   up   %.
3 Add a column   labelled Total Sales to   calculate   how   much   money   the   company   made   selling each   item at   the Selling Price.
4 Add another column labelled Profit and   subtract   the Total Cost from   the Total Sales. Total the   Profit column.
5 Management   have   decided   that   items   supplied   by   Fender   that   have   less   than   $500         profit are to   be discontinued; all other   items   supplied   by   Fender   should   be   reviewed.   Items supplied   by other suppliers   should   be   restocked.
6 Add a column   labelled Action and   use   a   function   to   show   whether   items   are   to   be   discontinued,   reviewed or   restocked. Use absolute   cell   referencing.
7 Create   a   pie   chart   to   illustrate   the % of Profit for   each   item.
8 Save   your   file.
Activity 1e - IF/OR formula 
Use the file Malley.xlsx for this activity.   Use the Pay worksheet.
The   Malley Organisation wants to show the projected   2021   salaries   of   employees.
A   10%   increase   will   be   given   to   employees   if   they   work   in   the   IT   or   Accounts Departments, or if they   have   been employed   more than   5 years.
All other employees will   receive   an   increase   of   5%.
1 In   the Increase % column,   use   a   function   to   show   the   correct   value.      Make   sure   you use absolute   referencing.
Hint:   Instead of using a nested   If / AND you need to   use   a   nested   If /   OR.
2 Calculate   the 2021 Salary (Column G).
3 Save   your   file.
Challenge activities
Activity 1f 
Use the Bonus 1 worksheet   in the Malley.xlsx workbook
All   employees except those   in   the   Sales   Department   will   get   a   $350   bonus.
1 Construct an IF   formula   in   Column F to   show this.   What   comparison   operator   would   you   use?
2 Use   absolute   referencing.
3 Save   your   file.
Activity 1g 
Use the Bonus 2 worksheet   in the Malley.xlsx workbook
The   Malley   company   will   pay   employees   a   bonus   (extra) amount   of   $350.   However, the following employees should not get   the   bonus   amount:
•             Employees   in the   Sales   Department
•             Employees   with   more   than   10 days   of   absence.
1 Construct   an   IF   formula   in   Column F to   show   this.There are several ways that you can   construct   your   formula,   nesting   either AND   or OR   inside your IF formula. Consider the   appropriate   comparison   operators   to   use with the conditions you   want   to   use.
Tip: Another option   is to   use the   NOT function. This would add another level   of   nesting   inside your   IF formula.
2       Sort your spreadsheet   by column   F,   smallest to   largest.   Compare   your   answer   to   the   screen shot   below, which displays only those who   received   0.   Cells   that   match   the conditions above   have   been   highlighted, to   help you   understand the   logic.
3 Save   your   file.

         
加QQ：99515681  WX：codinghelp
