# Gherkin Syntax Highlight Macro for Microsoft Office Applications (Word and Excel)

**Before you use, it check if [Easy Syntax Highlighter](https://appsource.microsoft.com/en-us/product/office/WA200000011?src=office&tab=Overview) is better dolution for you.**
The plugin is much more powerfull. The macros might be usefull if you dont have the latest versions of MS applications. 

VBA macros intended to: 
- Apply/clear Gherkin Syntax Highlight to/from the selected text in MS Word (or the selected cells in MS EXCEL)
- Highlights non-ascii characters - it is suggested to replace them with regular ASCII chars

![Macro highlighting result](/img/gherkinHighlightMacro.png)

## Tips to use Word Macro:

- Do not include macros into the documents you distribute to others: macros are not  considered to be a safe document content.
- Keep your macros either in Normal.dot or in a dedicated document. If you save it to Normal.dot, you will be able to access the macro anytime . If the macro is stored in a dedicated document, you need to keep it open to access the macro.
- Assign a shortcut to macro (See the instructions below)

##  Word Macro Requirements:

- Make sure you enable macro usage
- Make sure you enable  a reference to "Microsoft VBScript Regular Expressions 5.5" Object Library
Gherkin Macro assumes that each clause is started with a new line.  If highlight is not applied make sure there  is a linefeed character  after each line

 
## Tips to use Excel Macro:

- Do not include macros into the documents you distribute to others: macros are not  considered to be a safe workbook content.
- Keep all your macros in a single workbook. You need to keep it open when you work with spreadsheets that requires macro usage
- Assign a shortcut to macro (See the instructions below)

## Excel Macro Requirements:

- Make sure you enable macro usage
- Make sure you enable  a reference to "Microsoft VBScript Regular Expressions 5.5" Object Library
- Gherkin Macro assumes that each clause is started with a new line.  If highlight is not applied make sure there  is a linefeed character after each line

The Macro Options Window: Shortcut Key
We can use the Macro Options window in Excel to create a shortcut key to call the macro.  Here are the instructions on how to set it up.

Start by going to the Developer tab and clicking on the Macros button.  (If you don't see the Developer tab on your ribbon, you can add it using these instructions.) Alternatively, you can use the keyboard shortcut Alt+F8.

![](https://www.excelcampus.com/wp-content/uploads/2018/09/Macro-Button-on-Developer-Tab.png)

After selecting the macro that you want to assign the shortcut to, click the Options button.
![Assign Keyboard Shortcut to Macro - Open Options Window](https://www.excelcampus.com/wp-content/uploads/2018/09/Assign-Keyboard-Shortcut-to-Macro-Open-Options-Window.png)

In the Macro Options Window, you can create the shortcut you want by adding a letter, number, or symbol.
![Macro shortcut](https://www.excelcampus.com/wp-content/uploads/2018/09/Macro-Options-Window-Choose-Shortcut-Key.png)

Macro Options Window Choose Shortcut Key
Be careful not to override an existing shortcut that you frequently use, such as Ctrl+C to copy. One way to avoid doing this is by adding Shift to the shortcut to make it a bit more complex. In my example, I used Ctrl+Shift+C.

![Assign Keyboard Shortcut to Macro - Ctrl Shift Combination](https://www.excelcampus.com/wp-content/uploads/2018/09/Assign-Keyboard-Shortcut-to-Macro-Ctrl-Shift-Combination.png)

To delete the shortcut, simply repeat the process for accessing the Macro Options Window and then delete the character that you entered to create the shortcut.

