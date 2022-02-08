#target photoshop
var scriptFolder = (new File($.fileName)).parent; // The location of this script

// Keeping the ruler settings to reset in the end of the script
var startRulerUnits = app.preferences.rulerUnits;
var startTypeUnits = app.preferences.typeUnits;
var startDisplayDialogs = app.displayDialogs;

// Changing ruler settings to pixels for correct image resizing
app.preferences.rulerUnits = Units.PIXELS;
app.preferences.typeUnits = TypeUnits.PIXELS;
app.displayDialogs = DialogModes.NO;

// VARIABLES

var folderMain, folderMatch, folderPSD, filesMatch, filesPSD;
var listMatch = [];
var listPSD = [];
var previousImage, failedToCopy;
var processedFiles = 0;
var customOpt = false;
var promptLoc = false;
/*
listMatch = [
    {
        id: "",
        file: "",
        angle: ""
    }
]
listPSD = [
    {
        id: "",
        file: "",
        number: "",
        match: undefined
    }
]
*/
var d, timeStart;
var errorLog = ["Error log:\n"];

try {
    init();
} catch(e) {
    alert("Error code " + e.number + " (line " + e.line + "):\n" + e);
}

// Reset the ruler
app.preferences.rulerUnits = startRulerUnits;
app.preferences.typeUnits = startTypeUnits;
app.displayDialogs = startDisplayDialogs;

function init() {

    if (app.documents.length != 0) return alert("Please close any open documents before continuing");

    d = new Date();
    timeStart = d.getTime() / 1000;
    
    // This is preperation for server controlled usage. Should probably be a text file instead of CustomOptions. Also, no confirm boxes wtf.
    if (customOpt) {
        var returnBack = false;
        getStoredFolders();
        function getStoredFolders() {
            try {
                var optFolderMatch = app.getCustomOptions("ETON_GarmentsMatch").getString(0);
                if (optFolderMatch != "null" && optFolderPSD != "null") {
                    folderMatch = new Folder(optFolderMatch);
                    if (!folderMatch.exist) {
                        if (confirm("Warning!\nThe folder \"Garments to Colour Match\" is previously stored and does not exist.\n\nThe stored location is:\n" + optFolderMatch + "\n\nDo you want to reset this stored location (will prompt new location input)?")) {
                            var desc = new ActionDescriptor();
                                desc.putString(0, "null");
                            app.putCustomOptions("ETON_GarmentsMatch", desc);
                            getStoredFolders();
                        } else {
                            returnBack = true;
                            return;
                        }
                    }
                    var optFolderPSD = app.getCustomOptions("ETON_GarmentsMatch").getString(1);
                    if (optFolderPSD != "null") {
                        folderPSD = new Folder(optFolderPSD);
                        if (!folderPSD.exist) {
                            if (confirm("Warning!\nThe folder \"Images to work on\" is previously stored and does not exist.\n\nThe stored location is:\n" + optFolderPSD + "\n\nDo you want to reset this stored location (will prompt new location input)?")) {
                                var desc = new ActionDescriptor();
                                    desc.putString(1, "null");
                                app.putCustomOptions("ETON_GarmentsMatch", desc);
                                getStoredFolders();
                            } else {
                                returnBack = true;
                                return;
                            }
                        }
                    }
                } else {
                    return; // promptNewLocations();
                }
            } catch(e) {
                if (e.number == 8500) {
                    return; // promptNewLocations();
                } else {
                    alert("Error code " + e.number + " (line " + e.line + "):\n" + e);
                }
                return;
            }
        }
        if (returnBack) return;
    } else {
        folderMatch = new Folder(scriptFolder + "/Garments to Colour Match");
        // if (!folderMatch.exist) throw new Error("The folder \"Garments to Colour Match\" does not exist in the script location");
        folderPSD = new Folder(scriptFolder + "/Images to work on");
        // if (!folderPSD.exist) throw new Error("The folder \"Images to work on\" does not exist in the script location");
    }

    filesMatch = folderMatch.getFiles("*.jpg");
    if (filesMatch.length == 0) throw new Error("There are no jpg files in the Match folder");
    filesPSD = folderPSD.getFiles("*.psd");
    if (filesPSD.length == 0) throw new Error("There are no psd files in the Work folder");

    for (i = 0; i < filesMatch.length; i++) {
        listMatch.push({
            id: filesMatch[i].name.substring(0,11),
            file: File(filesMatch[i].path + "/" + filesMatch[i].name),
            name: filesMatch[i].name,
            angle: filesMatch[i].name.substring(filesMatch[i].name.indexOf("_") + 1,filesMatch[i].name.length - 4)
        });
    }
    for (i = 0; i < filesPSD.length; i++) {
        listPSD.push({
            id: filesPSD[i].name.substring(0,11),
            file: File(filesPSD[i].path + "/" + filesPSD[i].name),
            name: filesPSD[i].name,
            number: filesPSD[i].name.substring(filesPSD[i].name.indexOf("_") + 1,filesPSD[i].name.length - 4),
            match: undefined
        });
    }
    
    main();

}

function main() {

    for (i = 0; i < listPSD.length; i++) {
        var matchStatus = 0; // 0 is no colour match angle found, low is low satisfaction in angle selection
        for (j = 0; j < listMatch.length; j++) {
            if (listMatch[j].id == listPSD[i].id) {
                switch (listMatch[j].angle) {
                    case "tshov_1": if (matchStatus < 5) listPSD[i].match = listMatch[j].file; matchStatus = 5; break;
                    case "tshgh_1": if (matchStatus < 4) listPSD[i].match = listMatch[j].file; matchStatus = 4; break;
                    case "tshco_1": if (matchStatus < 3) listPSD[i].match = listMatch[j].file; matchStatus = 3; break;
                    case "tshde_1": if (matchStatus < 2) listPSD[i].match = listMatch[j].file; matchStatus = 2; break;
                    case "tshpi_1": if (matchStatus < 1) listPSD[i].match = listMatch[j].file; matchStatus = 1; break;
                }
            }
            if (matchStatus == 5) continue;
        }
    }

    for (i = 0; i < listPSD.length; i++) {
        if (listPSD[i].match != undefined) {

            open(listPSD[i].file);
            try {

                // if (activeDocument.layers.length > 1) throw new Error(activeDocument.name + " has a populated layer structure");
                try {
                    activeDocument.activeLayer = activeDocument.layerSets.getByName("Garment Colour Match");
                    activeDocument.close(SaveOptions.DONOTSAVECHANGES);
                    continue;
                } catch(e) {}
                
                if (!failedToCopy && previousImage != undefined && listPSD[i].match == previousImage.match) {
                    try {
                        var idpast = charIDToTypeID( "past" );
                            var desc209 = new ActionDescriptor();
                            var idinPlace = stringIDToTypeID( "inPlace" );
                            desc209.putBoolean( idinPlace, true );
                        executeAction( idpast, desc209, DialogModes.NO );
                    } catch(e) {
                        var docMatch = open(listPSD[i].match);
                        activeDocument.selection.selectAll();
                        activeDocument.selection.copy();
                        docMatch.close(SaveOptions.DONOTSAVECHANGES);

                        activeDocument.paste();
                    }
                } else {
                    var docMatch = open(listPSD[i].match);
                    activeDocument.selection.selectAll();
                    activeDocument.selection.copy();
                    docMatch.close(SaveOptions.DONOTSAVECHANGES);

                    activeDocument.paste();

                    var lyrMatch = activeDocument.activeLayer;

                    var targetX = activeDocument.width.value/2;
                    var targetY = 0;
                    activeDocument.activeLayer.translate(targetX - activeDocument.activeLayer.bounds[2].value, targetY - activeDocument.activeLayer.bounds[1].value);

                    var lyrSize = activeDocument.activeLayer.bounds[2].value - activeDocument.activeLayer.bounds[0].value;
                    var lengthToCenter = activeDocument.width.value / 2;
                    var resizeAmount = lengthToCenter / lyrSize * 100;
                    activeDocument.activeLayer.resize(resizeAmount, resizeAmount, AnchorPosition.TOPRIGHT);

                    activeDocument.activeLayer.resize(98, 98, AnchorPosition.BOTTOMRIGHT);

                    activeDocument.activeLayer.name = listPSD[i].match.name.substring(listPSD[i].match.name.indexOf("_") + 1,listPSD[i].match.name.length - 4);

                    var grpMatch = activeDocument.layerSets.add();
                    activeDocument.activeLayer.name = "Garment Colour Match";
                    functionLayerColour("Grey");

                    lyrMatch.move(grpMatch, ElementPlacement.INSIDE);
                    grpMatch.visible = false;
                    activeDocument.activeLayer = grpMatch;
                    try { 
                        var idcopy = charIDToTypeID( "copy" );
                        executeAction( idcopy, undefined, DialogModes.NO );
                        failedToCopy = false;
                    } catch(e) {
                        failedToCopy = true;
                    }
                }

                activeDocument.activeLayer = activeDocument.layers[activeDocument.layers.length - 1];

                activeDocument.close(SaveOptions.SAVECHANGES);

            } catch(e) {
                errorLog.push(listPSD[i].name + " @ " + e.line + ": " + e.message);
                activeDocument.close(SaveOptions.DONOTSAVECHANGES);
            }
            previousImage = listPSD[i];
            processedFiles++;

        }
    }

    var d = new Date();
    var timeEnd = d.getTime() / 1000;
    var timeFull = timeEnd - timeStart;
    if (errorLog.length != 1) {
        var date = d.getFullYear() + "/" + (d.getMonth() + 1) + "/" + d.getDate() + " " + d.getHours() + ":" + d.getMinutes();
        saveTxt(date + "\n" + processedFiles + " files processed\n\n" + errorLog.join("\n---\n"), "ERRORLOG", scriptFolder, ".txt");
        alert(processedFiles + " files processed!\nTime elapsed " + formatSeconds(timeFull) + "\n\n" + errorLog.join("\n"));
    } else {
        alert(processedFiles + " files processed!\nTime elapsed " + formatSeconds(timeFull));
    }

}

// FUNCTIONS

function formatSeconds(seconds) {
    var sec_num = parseInt(seconds, 10); // don't forget the second param
    var hours   = Math.floor(sec_num / 3600);
    var minutes = Math.floor((sec_num - (hours * 3600)) / 60);
    var seconds = sec_num - (hours * 3600) - (minutes * 60);

    if (hours   < 10) {hours   = "0"+hours;}
    if (minutes < 10) {minutes = "0"+minutes;}
    if (seconds < 10) {seconds = "0"+seconds;}
    return hours+':'+minutes+':'+seconds;
}

/*
  FUNCTION DESCRIPTION:
    Assign a colour to the active layer
    
  INPUT:
    colour (String) = The colour you want for the layer ("Yellow"). Can be "None" to remove colours.
    
  OUTPUT:
    None
    
  NOTE:
    "Gray" or "grey" for others
*/

function functionLayerColour(colour) {
	switch (colour.toLocaleLowerCase()) {
		case 'red': colour = 'Rd  '; break;
		case 'orange' : colour = 'Orng'; break;
		case 'yellow' : colour = 'Ylw '; break;
		case 'green' : colour = 'Grn '; break;
		case 'blue' : colour = 'Bl  '; break;
		case 'violet' : colour = 'Vlt '; break;
		case 'gray' : colour = 'Gry '; break;
		case 'grey' : colour = 'Gry '; break;
		case 'none' : colour = 'None'; break;
		default : colour = 'None'; break;
	}
	var desc = new ActionDescriptor();
		var ref = new ActionReference();
		ref.putEnumerated( charIDToTypeID('Lyr '), charIDToTypeID('Ordn'), charIDToTypeID('Trgt') );
	desc.putReference( charIDToTypeID('null'), ref );
		var desc2 = new ActionDescriptor();
		desc2.putEnumerated( charIDToTypeID('Clr '), charIDToTypeID('Clr '), charIDToTypeID(colour) );
	desc.putObject( charIDToTypeID('T   '), charIDToTypeID('Lyr '), desc2 );
	executeAction( charIDToTypeID('setd'), desc, DialogModes.NO );
}

function saveTxt(text, name, path, ext) {
    if (!ext) ext = ".txt";
    var saveFile = File(Folder(path) + "/" + name + ext);
    if (saveFile.exists) saveFile.remove();
    saveFile.encoding = "UTF8";
    saveFile.open("e", "TEXT", "????");
    saveFile.writeln(text);
    saveFile.close();
}