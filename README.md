![QuarkXPress 2018 logo](http://www.quarkforums.com/resources/git/githeader.jpg)
# Mail merge
This JavaScript needs to be installed in QuarkXPress 2018. Feel free to modify this script to your own needs.  
**Please see here on how to install: [**Installation Instructions**](#howinstall)**
## What it does
This script uses a pre-fabricated page and uses it as a template to create as many pages as the number of records found in a VALID csv file.
### Prerequisites
- You need to cretae a document that can be used as a "template"
- The text and images to import data must be on the master page and not on the first page
- The text to be imported must be enclosed in braces as <<tagname>> and the tag name is the data of first tuple of the csv
- You need a csv (comma separated values) file containing your data
- The csv needs to be valid i.e. the CSV data must not contain the separator character
- The separator must be defined in the first line of csv file else the default separator will be taken as comma
- Images will be referenced by paths (relative or absolute)

### Notes
This sample ships with QuarkXPress 2018
### Screenshots
Template (qxp file):  
![QuarkXPress document](http://www.quarkforums.com/resources/git/md_images/mailmerge3.png)  
Data (csv file):     
![CSV file](http://www.quarkforums.com/resources/git/md_images/mailmerge2.png)  
Video:  
[![IMAGE ALT TEXT HERE](http://img.youtube.com/vi/Z5olohxEasg/0.jpg)](http://www.youtube.com/watch?v=Z5olohxEasg)  

### Version History  
May 2018: Original version as supplied with QuarkXPress 2018
## <a name="howinstall"></a>How to install
1. On the GitHub page, download the ZIP by clicking on the green button "Clone or download"  
![Step 1](http://www.quarkforums.com/resources/git/install_images/step1.png)
2. In the popout menu click on "Download ZIP"  
![Step 2](http://www.quarkforums.com/resources/git/install_images/step2.png)
3. Save to your Desktop
4. Unzip (so that you get a folder)
5. Copy the resulting folder to the js folder in your documents folder (**see below**)
6. In QuarkXPress open the "JavaScript" palette  
(via "Window" menu)
7. If you do not see a folder with the name of this JavaScript, click on the little "Home" ("House") symbol.  
![Step 7](http://www.quarkforums.com/resources/git/install_images/step7.png)

Step 5: On MacOS copy to|Step 5: On Windows copy to
---|---
~/Documents/Quark/QuarkXPress 2018/js/|Documents\Quark\QuarkXPress 2018\js\
(so into your "Documents" folder)|(so into your "Documents" folder)

**Run the JavaScript by first double clicking the folder; and then double clicking the Script itself (in the JavaScript palette of QuarkXPress).**

## More Information
More information on QuarkXPress and how to use JavaScript in QuarkXPress can be found here:  
<http://www.quark.com/Support/Documentation/QuarkXPress/>
