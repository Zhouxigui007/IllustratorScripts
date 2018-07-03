#target illustrator
var doc = app.activeDocument;
var mm = 2.834645; //2.834645 points per mm
var cropMarkSquare = 0.3*mm; //Crop mark size in mm
var cropMarkCircle = 0.5*mm; //Flatbed crop mark size in mm
var yx = new Array();
var weedingBorderSize = mm*2; //weeding border size in mm

d=doc.selection[0]; // first selected elements
xy=d.position; // an Array() with X & Y

// Object size including weeding border
var pWidth = d.width+(weedingBorderSize*2);
var pHeight = d.height+(weedingBorderSize*2);

// Object position including weeding border
var xPosition = xy[0]-weedingBorderSize;
var yPosition = xy[1]+weedingBorderSize;

// Total size
var widthTotal = pWidth+(mm*2)+(cropMarkSquare*2);
var heightTotal = pHeight+(mm*2)+(cropMarkSquare*2);

// Middle crop mark Y position
var yMiddleLeftMark0 = (heightTotal-cropMarkSquare)/1;
var yMiddleLeftMark1 = (heightTotal-cropMarkSquare)/2;
var yMiddleLeftMark2 = (heightTotal-cropMarkSquare)/3;
var yMiddleLeftMark3 = (heightTotal-cropMarkSquare)/4;

//Add 2mm weeding border around selection
function weedingBorder() {
	activeDocument.pathItems.rectangle(yPosition, xPosition, pWidth, pHeight);//weeding border
}

//Add Square corner crop marks
function cornerCropMarkSquares () {
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare, xPosition-mm-cropMarkSquare, cropMarkSquare, cropMarkSquare);//Top Left Mark
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare, xPosition+pWidth+mm, cropMarkSquare, cropMarkSquare);//Top Right Mark
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare-yMiddleLeftMark0, xPosition-mm-cropMarkSquare, cropMarkSquare, cropMarkSquare);//Bottom Left Mark
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare-yMiddleLeftMark0, xPosition+pWidth+mm, cropMarkSquare, cropMarkSquare);//Bottom Right Mark
}

//Add Circle crop marks
function cornerCropMarkCircles () {
	activeDocument.pathItems.ellipse(yPosition+mm+cropMarkCircle-5, xPosition-mm-cropMarkCircle, cropMarkCircle, cropMarkCircle);//Top Left Mark
	activeDocument.pathItems.ellipse(yPosition+mm+cropMarkCircle, xPosition+pWidth+mm-5, cropMarkCircle, cropMarkCircle);//Top Right Mark
	activeDocument.pathItems.ellipse(yPosition+mm+cropMarkSquare-yMiddleLeftMark0, xPosition-mm-cropMarkCircle+5, cropMarkCircle, cropMarkCircle);//Bottom Left Mark
	activeDocument.pathItems.ellipse(yPosition+mm+cropMarkCircle+5-yMiddleLeftMark0, xPosition+pWidth+mm, cropMarkCircle, cropMarkCircle);//Bottom Right Mark
}

//Text function
function text () {
	var text1 = doc.textFrames.pointText( [xPosition, yPosition+1] );
	var fontStyle = text1.textRange.characterAttributes;
	text1.contents = app.activeDocument.name;
	fontStyle.size = 1.5;
	text1.rotate(180, undefined, undefined, true);
}

/*

//Create CutContour Spot Colour
function doesSpotExist(name) {
    for (i=0; i<doc.spots.length; i++) {
        if (doc.spots[i].name==name) {
            return true;
            }
        return false;
        }
    }

if (doesSpotExist("CutContour")==false) {
    var newColor = new CMYKColor();
    newColor.cyan = 0;
    newColor.magenta = 100;
    newColor.yellow = 0;
    newColor.black = 0;

    var newSpot = doc.spots.add();
    newSpot.name = "CutContour";
    newSpot.colorType = ColorModel.SPOT;
    newSpot.color = newColor;
    }

// Set variable to your color  
var newSpotColor = doc.swatches.getByName("CutContour");  
// Set a variable to the document path items collection
var pathList = doc.pathItems;
var pathListHalf = Math.round(doc.pathItems.length/2);

// Loop through this list of objects  
for (var h = 0; h < pathListHalf; h++)     {         
     // Change the property values  
     pathList[h].filled = false;  
     pathList[h].stroked = true;  
     pathList[h].strokeColor = newSpotColor.color;  
     pathList[h].strokeWidth = 0.2834645;     
}     
app.redraw();

*/

//Main
if (d.height > 400*mm) {
	alert("Print panel is too long. Please reduce the length.");
} else if (d.height > 300*mm) {
	weedingBorder();
	cornerCropMarkSquares();
	//cornerCropMarkCircles();
	text();
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare-yMiddleLeftMark3, xPosition-mm-cropMarkSquare, cropMarkSquare, cropMarkSquare);//Middle Square Left Mark 1
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare-yMiddleLeftMark3, xPosition+pWidth+mm, cropMarkSquare, cropMarkSquare);//Middle Square Right Mark 1
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare-yMiddleLeftMark3*2, xPosition-mm-cropMarkSquare, cropMarkSquare, cropMarkSquare);//Middle Square Left Mark 2
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare-yMiddleLeftMark3*2, xPosition+pWidth+mm, cropMarkSquare, cropMarkSquare);//Middle Square Right Mark 2
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare-yMiddleLeftMark3*3, xPosition-mm-cropMarkSquare, cropMarkSquare, cropMarkSquare);//Middle Square Left Mark 3
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare-yMiddleLeftMark3*3, xPosition+pWidth+mm, cropMarkSquare, cropMarkSquare);//Middle Square Right Mark 3
	//activeDocument.pathItems.ellipse(yPosition+mm+cropMarkSquare-5-yMiddleLeftMark3, xPosition-mm-cropMarkCircle, cropMarkCircle, cropMarkCircle);//Middle Square Left Mark 1
	//activeDocument.pathItems.ellipse(yPosition+mm+cropMarkSquare+5-yMiddleLeftMark3, xPosition+pWidth+mm, cropMarkCircle, cropMarkCircle);//Middle Square Right Mark 1
	//activeDocument.pathItems.ellipse(yPosition+mm+cropMarkSquare-5-yMiddleLeftMark3*2, xPosition-mm-cropMarkCircle, cropMarkCircle, cropMarkCircle);//Middle Square Left Mark 2
	//activeDocument.pathItems.ellipse(yPosition+mm+cropMarkSquare+5-yMiddleLeftMark3*2, xPosition+pWidth+mm, cropMarkCircle, cropMarkCircle);//Middle Square Right Mark 2
	//activeDocument.pathItems.ellipse(yPosition+mm+cropMarkSquare-5-yMiddleLeftMark3*3, xPosition-mm-cropMarkCircle, cropMarkCircle, cropMarkCircle);//Middle Square Left Mark 3
	//activeDocument.pathItems.ellipse(yPosition+mm+cropMarkSquare+5-yMiddleLeftMark3*3, xPosition+pWidth+mm, cropMarkCircle, cropMarkCircle);//Middle Square Right Mark 3
	//alert("Crop marks complete.");
} else if (d.height > 200*mm) {
	weedingBorder();
	cornerCropMarkSquares();
	//cornerCropMarkCircles();
	text();
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare-yMiddleLeftMark2, xPosition-mm-cropMarkSquare, cropMarkSquare, cropMarkSquare);//Middle Square Left Mark 1
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare-yMiddleLeftMark2, xPosition+pWidth+mm, cropMarkSquare, cropMarkSquare);//Middle Square Right Mark 1
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare-yMiddleLeftMark2*2, xPosition-mm-cropMarkSquare, cropMarkSquare, cropMarkSquare);//Middle Square Left Mark 2
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare-yMiddleLeftMark2*2, xPosition+pWidth+mm, cropMarkSquare, cropMarkSquare);//Middle Square Right Mark 2
	//activeDocument.pathItems.ellipse(yPosition+mm+cropMarkSquare-5-yMiddleLeftMark2, xPosition-mm-cropMarkCircle, cropMarkCircle, cropMarkCircle);//Middle Square Left Mark 1
	//activeDocument.pathItems.ellipse(yPosition+mm+cropMarkSquare+5-yMiddleLeftMark2, xPosition+pWidth+mm, cropMarkCircle, cropMarkCircle);//Middle Square Right Mark 1
	//activeDocument.pathItems.ellipse(yPosition+mm+cropMarkSquare-5-yMiddleLeftMark2*2, xPosition-mm-cropMarkCircle, cropMarkCircle, cropMarkCircle);//Middle Square Left Mark 2
	//activeDocument.pathItems.ellipse(yPosition+mm+cropMarkSquare+5-yMiddleLeftMark2*2, xPosition+pWidth+mm, cropMarkCircle, cropMarkCircle);//Middle Square Right Mark 2
	//alert("Crop marks complete.");
} else if (d.height > 100*mm) {
	weedingBorder();
	cornerCropMarkSquares();
	//cornerCropMarkCircles();
	text();
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare-yMiddleLeftMark1, xPosition-mm-cropMarkSquare, cropMarkSquare, cropMarkSquare);//Middle Square Left Mark
	activeDocument.pathItems.rectangle(yPosition+mm+cropMarkSquare-yMiddleLeftMark1, xPosition+pWidth+mm, cropMarkSquare, cropMarkSquare);//Middle Square Right Mark
	//activeDocument.pathItems.ellipse(yPosition+mm+cropMarkSquare-5-yMiddleLeftMark1, xPosition-mm-cropMarkCircle, cropMarkCircle, cropMarkCircle);//Middle Circle Left Mark
	//activeDocument.pathItems.ellipse(yPosition+mm+cropMarkSquare+5-yMiddleLeftMark1, xPosition+pWidth+mm, cropMarkCircle, cropMarkCircle);//Middle Circle Right Mark
	alert("Crop marks complete.");
} else if (d.height < 100*mm) {
    weedingBorder();
	cornerCropMarkSquares();
	//cornerCropMarkCircles();
	text();
	//alert("Crop marks complete.");
}

/*
//Debug
alert (
	"Total Width: "+widthTotal+
	"\rTotal Height: "+heightTotal+
	"\rX Position: "+xy[0]+
	"\rY Position: "+xy[1]+
	"\rCalculation: "+Calculation
	);
*/