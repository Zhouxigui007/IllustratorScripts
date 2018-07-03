#target illustrator
var doc = app.activeDocument;
var mm = 2.834645; //2.834645 points per mm

var d=doc.selection[0]; // selected elements
var xy=d.position; // an Array() with X & Y

var weedingBorderSize = 1*mm;

// Object size including weeding border
var pWidth = d.width+weedingBorderSize*2;
var pHeight = d.height+weedingBorderSize*2;

// Object position including weeding border
var xPosition = xy[0]-weedingBorderSize;
var yPosition = xy[1]+weedingBorderSize;

function doesLayerExist(name) {
    for (i=0; i<activeDocument.layers.length; i++) {
        if (activeDocument.layers[i].name==name) {
            return true;
		}
	return false;
	}
}

//////////	
///MAIN///
//////////
	
// Add layer named "Cut Layer"
if (doesLayerExist("Cut Layer")==false) {
	var nl = activeDocument.layers.add();
	nl.name = "Cut Layer";
	// alert("Layer added");
}

// Red CMYK Color
colRed = new CMYKColor();
colRed.cyan = 0;
colRed.magenta = 100;
colRed.yellow = 100;
colRed.black = 0;

// White CMYK Color
colWhite = new CMYKColor();
colWhite.cyan = 0;
colWhite.magenta = 0;
colWhite.yellow = 0;
colWhite.black = 0;

// Move selection to "Cut Layer"
var cutLayer = activeDocument.layers.getByName("Cut Layer");
doc.selection[0].move(cutLayer, ElementPlacement.PLACEATBEGINNING);

// Get colour of selected object
if (doc.pathItems[0].fillColor != "[SpotColor]") {
	var newCMYKColor = new CMYKColor();
	newCMYKColor.cyan = doc.pathItems[0].fillColor.cyan;
	newCMYKColor.magenta = doc.pathItems[0].fillColor.magenta;
	newCMYKColor.yellow = doc.pathItems[0].fillColor.yellow;
	newCMYKColor.black = doc.pathItems[0].fillColor.black;
	var numCMYK = newCMYKColor.cyan + newCMYKColor.magenta + newCMYKColor.yellow + newCMYKColor.black;
} else {
	var newCMYKColor = new CMYKColor();
	newCMYKColor.cyan = doc.pathItems[0].fillColor.spot.color.cyan;
	newCMYKColor.magenta = doc.pathItems[0].fillColor.spot.color.magenta;
	newCMYKColor.yellow = doc.pathItems[0].fillColor.spot.color.yellow;
	newCMYKColor.black = doc.pathItems[0].fillColor.spot.color.black;
	var numCMYK = newCMYKColor.cyan + newCMYKColor.magenta + newCMYKColor.yellow + newCMYKColor.black;
}

// Add weeding border around selection
if (numCMYK != 0) {
	// Add selected objects colour weeding border
	var selectedObject = doc.pathItems.rectangle(yPosition, xPosition, pWidth, pHeight);//weeding border
	selectedObject.stroked = true;
	selectedObject.filled = false;
	selectedObject.strokeColor = newCMYKColor;
	selectedObject.strokeWidth = 0.2834645;
	// Add White background colour
	var newCMYKColor = new CMYKColor();
	newCMYKColor.cyan = 0;
	newCMYKColor.magenta = 0;
	newCMYKColor.yellow = 0;
	newCMYKColor.black = 0;
	selectedObject.fillColor = newCMYKColor;
	// Send selected object to the back
	selectedObject.zOrder(ZOrderMethod.SENDTOBACK);
	doc.selection = null;
} else {
	// Add white weeding border with red background
	var selectedObject = doc.pathItems.rectangle(yPosition, xPosition, pWidth, pHeight);//weeding border
	selectedObject.stroked = true;
	selectedObject.fillColor = colRed;
	selectedObject.strokeColor = colWhite;
	selectedObject.strokeWidth = 0.2834645;
	selectedObject.zOrder(ZOrderMethod.SENDTOBACK);
	doc.selection = null;
}