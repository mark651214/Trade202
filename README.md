# Prerequisite
1. MS Visual Studio 2013 or above.
2. Newtonsoft.Json framework.
3. HtmlAgilityPack framework.
4. Install [ImDisk](https://sourceforge.net/projects/imdisk-toolkit/) for speed up file access. 

# Get Start
1. Download historic data from [TWSE](https://www.twse.com.tw/) & [TPEX](https://www.tpex.org.tw/). It's implement at button button_GetMargin_Click , button_Get3Party_Click and button_GetPrice_Click.
2. Merge historic data on virtual disk for later use. It's implement at button button3_Click, button5_Click and button15_Click.
3. Restart the compiled executable file ExamChipTrade.exe. It will reread historic from virtual disk.
4. Select date from calendar.
5. Click a button for calculate statistic at that day your selected. the result will be display at visual studio debug output panel.<p>
 EX: button9_Click is for calculate main 3 big group buy above 1.5% of individual company total stocks at the same day. 