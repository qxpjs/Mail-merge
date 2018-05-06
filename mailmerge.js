/*
This JS finds the master page and creates as many pages with that master page as the number of records in a VALID csv file and creates documents by importing csv data
Each page of the document contains data corresponding to a tuple in the csv
-The text and images to import data must be on the master page and not on the first page
-The text to be imported must be enclosed in braces as <<tagname>> and the tag name is the data of first tuple of the csv
-The CSV should be valid i.e. the CSV data must not contain the separator character
-The separator must be defined in the first line of cvs else the default separator will be taken as comma
*/

if(typeof(isLayoutOpen) == "undefined")
{
	//import basic checks
	app.importScript(app.getAppScriptsFolder() + "/Dependencies/qx_validations.js");
	console.log("Loaded library for basic validation checks from application.");
}

var slash= "/";

//copy the required sample files to user scripts folder from app folder
var srcFolder = getPathFromAppFolder("Dependencies","Mail Merge");
var destFolder = app.getUserScriptsFolder().replace("js","Mail Merge Data");
console.log("Copying sample data from application to " + destFolder);
createDirectoryIfNotExists(destFolder);
copyFolderContents(srcFolder + slash , destFolder + slash, slash);

//Load Project
var projPath = getProjPath(destFolder + slash + "Sales-Letter-Template.qpt");
if(projPath != null)
{
	console.log("Loaded Project from: " + projPath);
	app.openProject(projPath);
	if(isLayoutOpen())
	{
		//Read the given CSV
		var CSVPath = getCSVPath(destFolder + slash + "Data.csv");
		if(CSVPath != null)
		{
			console.log("Loaded data file from " + CSVPath);
			var imgFolderPath = CSVPath.replace(/(\\|\/)\w*\.csv/g, "\\");
			var records = GetCSVRecords(CSVPath); // stores all the tuples in an array
			if(records != null)
			{
				var coltags = [];
				//get the separator from csv else set the default value as comma
				if(records[0].search(/sep/i) >=0)
				{
					var separator = records[0].match(/sep\s*=\s*(\S)/i)[1];//get the separator from first line of csv 
					console.log("Found CSV separator: " + separator);
					coltags = records[1].split(separator);
					records.splice(0, 2);//ignore two lines from the top in the CSV
				}
				else
				{
					var separator = ",";
					console.log("Found No CSV separator. setting default separator: comma(,).");
					coltags = records[0].replace(/\n/g,"").split(separator);
					records.splice(0, 1);//ignore one line from the top in the CSV
				}
				console.log("Loaded " + records.length + " records from CSV");
				for(var i =0; i< coltags.length; i++)
				{
					coltags[i] = coltags[i].trim();
					coltags[i] = "<<" + coltags[i] + ">>";//bring the tags in required format
				}	
				//if the CSV is not blank
				if(records[0] != null)
				{
					var columns = [];
					var colnum = records[0].split(separator).length;
					for(var j =0; j< colnum; j++ )
					{
						columns[j] = GetColValues(records, j + 1, separator);// stores column wise data
					}
					if(isTagAvailable(coltags, columns))
					{
						//get the number of records in excel
						var numpages= records.length;

						//create pages
						createPages(numpages)
							.then(makeReplacements(numpages, columns, coltags, imgFolderPath));
					}
					else
					{
						alert("The layout contains no tags to be replaced!");
					}
				}
				else			
					alert("CSV contains no records!");
			}
		}
	}
}

//*****************====================================Functions used in the JavaScript===============================****************//

//This function gets a file from a sub folder in application folder
function getPathFromAppFolder(subFolderName, filename){
	return app.getAppScriptsFolder() + "/" + subFolderName + "/" + filename;
}

//This function copies contents of a folder to a given target folder
function copyFolderContents(srcFolder, destFolder, slash){
	var srcContents = fs.readdirSync(srcFolder);
	for(var i = 0; i< srcContents.length; i++ )
	{
		var srcFilename = srcFolder + srcContents[i];
		var destFilename = destFolder + srcContents[i];
		if(fs.statSync(srcFilename).isFile())
		{
			//copy file if does not exists
			if(!fs.existsSync(destFilename))
			{
				fs.copyFileSync(srcFilename, destFilename);
			}
		}
		else if(fs.statSync(srcFilename).isDirectory())
		{
			createDirectoryIfNotExists(destFilename);
			copyFolderContents(srcFilename + slash, destFilename + slash);
		}
	}
}

//This function creates a directory at a given path if it does not exist already
function createDirectoryIfNotExists(dirPath){
	if(!fs.existsSync(dirPath))
	{
		fs.mkdirSync(dirPath);
		console.log("Created Directory " + dirPath);
	}
}

//function to get the path of project to b opened.
function getProjPath(samplePath){
	var flag = 0;
	while(flag==0)
	{
		var projPath = prompt("Please Enter the Project path: ", samplePath);
		if(projPath == null)
		{
			return null;
		}
		else if(projPath == "" || (projPath.search(".qxp") <0 && projPath.search(".qpt") <0))
		{
			alert("Invalid project path!");
		}
		else
		{
			projPath = projPath.trim();
			projPath = projPath.replace(/\\/g, "\\\\");//replace single slash with double slash in the path
			if(fs.existsSync(projPath)== 0 || fs.existsSync(projPath)== false)
			{
				alert("Project file does not exist!");
			}
			else
			{
				return projPath;
			}
		}
	}
}

//function to get valid input of CSV file path
function getCSVPath(samplePath){
	var flag = 0;
	while(flag==0)
	{
		var CSVPath = prompt("Please Enter the CSV path: ", samplePath);
		if(CSVPath == null)
		{
			return null;
		}
		else if(CSVPath == "" || CSVPath.search(".csv") <0)
		{
			alert("Invalid CSV file path!");
		}
		else
		{
			CSVPath = CSVPath.trim();
			if(fs.existsSync(CSVPath)== 0 || fs.existsSync(CSVPath)== false)
			{
				alert("CSV file does not exist!");
			}
			else
			{
				return CSVPath;
			}
		}
	}
	
}

//returns an array of rows of the records
function GetCSVRecords(filepath){
	var filedata= fs.readFileSync(filepath);//stores a string that contains all the data of a CSV
	if(filedata == "")
	{// if the file does not exist or is blank
		alert("CSV File contains no data!");
		return null;
	}
	else
	{
		var records= filedata.split("\n");//creates array to store tuples of CSV
		for(var i=0; i < records.length; i++ )
		{
			if(records[i] == "")
				break;
		}
		records.splice(i, records.length - i);//deletes extra blank records if created in csv
		return records;
	}
}

//returns teh values of that column
function GetColValues(records, colnumber, separator){
	var col =[];
	var index = [];
	for(var i=0; i < records.length; i++ )
	{
		index = getSeparatorIndices(records[i], separator);//stores the index values of separators in a CSV tuple
		if(colnumber == 1)//for the first column of the tuple
			col[i]= records[i].substring(0, index[0]);
		else if (colnumber == (index.length + 1))//for the last column of the tuple
			col[i]= records[i].substring(index[index.length - 1] + 1, records[i].length);
		else//for all other middle columns
			col[i]= records[i].substring(index[colnumber - 2] + 1, index[colnumber - 1]);
		col[i]= col[i].replace(/"/g, "");//remove extra quotes
	}
	return col;
}

//gets the indices of commas in the record
function getSeparatorIndices(record, separator){
	var index = [];
	var counter = 0;
	var test = record.indexOf(separator);//stores the index of first occurrence of the separator
	while( test >= 0)
	{
		index[counter] = test;
		test = record.indexOf(separator, index[counter] + 1);//stores the next occurrence of the separator
		counter++ ;
	}
	return index;
}

//checks if any tag from the CSV exists on the layout or not
function isTagAvailable(coltags, arr){
	var boxes = app.activeLayoutDOM().querySelectorAll('qx-box');
	//get the blank img box on this spread
	var imgbox= GetEmptyImgBox(boxes);
	var imgregex = /.(jpg|png|bmp|gif|tif)$/g;
	for(var k = 0; k< coltags.length; k++ )
	{
		var isImg = imgregex.exec(arr[k][1]);
		if(isImg != null && imgbox != null)// the tag belongs to an image uri for which image exists
		{
			return true;
		}
		else
		{
			var tag = new RegExp(coltags[k], "gi");
			if(replaceText(boxes, tag, "", "search"))
			{
				return true;
			}
		}
	}
	return false;
}

//creates the necessary number of pages
function createPages(numpages) {
	var promise = new Promise(function(resolve, reject)//promise is used to ensure this task completes and return a promise followed by further execution
	{
		setTimeout(function() 
		{
			console.log("Creating " + numpages + " pages");
			var layoutNode= app.activeLayoutDOM();
			var pages= layoutNode.getElementsByTagName('qx-page');
			var existingPageCount = pages.length;
			console.log(existingPageCount + " pages already exist!");
			//create numpages-existingPageCount pages (since existingPageCount pages are already present) 
			for(var i=existingPageCount + 1; i<= numpages; i++ )
			{
				var qxPage= document.createElement("qx-page");
				qxPage.setAttribute("page-index", i);
				var qxSpread= document.createElement("qx-spread");
				qxSpread.setAttribute("spread-index", i);
				qxSpread.appendChild(qxPage);
				app.activeLayoutDOM().appendChild(qxSpread);
				console.log("Creating page number " + i);
			}
		resolve({});
		}, 0);
	});
	return promise;
}

//function to add images from csv to empty image boxes and replace data as tags
function makeReplacements(numpages, arr, coltags, imgFolderPath) {
	var promise = new Promise(function(resolve, reject)
	{
		setTimeout(function() {
			console.log("Getting the DOM again");//getting the layout DOM again once all the pages are created
			//Get DOM of active Layout
			var layoutNode = app.activeLayoutDOM();
			//get the page node
			var pages= layoutNode.getElementsByTagName('qx-page');
			console.log("Setting values");
			var replacements =0;
			var imagecount=0;
			var imgregex = /.(jpg|png|bmp|gif|tif)$/g;
			//Make replacements on all the pages
			for(var i=0; i< numpages; i++ )
			{
				var spread = pages[i].parentNode;
				//get all boxes on this spread
				var boxes = spread.getElementsByTagName('qx-box');
				
				//get the blank img box on this spread
				var imgbox= GetEmptyImgBox(boxes);
				
				//replace required nodes
				for(var k = 0; k< coltags.length; k++ )
				{
					var imgURI= arr[k][i];
					var isImg = imgregex.exec(arr[k][1]);
					if(isImg != null && imgbox != null)// the tag belongs to an image uri for which image exists
					{
						imgURI= imgFolderPath + imgURI;
						//change the format of image source
						imgURI = convertURIToFilePath(imgURI);
				
						//add the image to box
						AddImageToBox(imgURI, imgbox);
						imagecount++ ;
					}
					else
					{
						var tag = new RegExp(coltags[k], "gi");
						replacements++ ;
						replaceText(boxes, tag, arr[k][i], "");
					}
				}
			}
			
			resolve({});
		}, 0);
	});
	return promise;
}

//function to get empty picture box
function GetEmptyImgBox(boxes){
	for(var i =0; i< boxes.length; i++ )
	{
		if(boxes[i].getAttribute('box-content-type') == "picture" && boxes[i].getElementsByTagName('qx-img')[0].getAttribute('src') == null)
		{
			return boxes[i];
		}
	}
}

//creates a picture box node and add the required picture to it
function AddImageToBox(path, Box){
	var imgtag= Box.getElementsByTagName('qx-img')[0];
	imgtag.setAttribute("src", path);//sets the image path
}

//function to replace/find string in text
function replaceText(Boxes, SourceString, NewString, flag){
	for(var i=0; i< Boxes.length; i++ )
	{
		if(Boxes[i].getAttribute("box-content-type") === "text")
		{
			//Get all the text runs from the box
			var boxTextSpans = Boxes[i].getElementsByTagName("qx-span");
			if (null != boxTextSpans)
			{
				//Iterate through all the Spans and Format the Numbers
				for (var j = 0; j < boxTextSpans.length; j++ )
				{
					var spanChildren = boxTextSpans[j].childNodes;
					if (null != spanChildren)
					{					
						for (var k = 0; k < spanChildren.length; k++ )
						{
							//Check if it is a text node
							if (spanChildren[k].nodeType === 3) 
							{
								if(spanChildren[k].nodeValue.search(SourceString) >= 0)
								{
									if(flag=="search")
									{
										return true;
									}
									else
									{
										spanChildren[k].nodeValue = spanChildren[k].nodeValue.replace(SourceString, NewString);
									}
									
								}
							}
						}
					}
				}
			}
		}
	}
}
