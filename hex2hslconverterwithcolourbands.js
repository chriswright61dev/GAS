function hextohsl() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];

    //Set up the header with array
    var values = [
        ["Hexcode", "Hex Number", "Hue", "Saturation", "Luminosity", "Colour Band", "Colour Band Number"]
    ];
    // Set the range of cells for the header
    var headerrange = sheet.getRange("A1:g1");
    // alternatively we can use sheet.getRange(row, column, numRows, numColumns) - numeric

    // Call the setValues method on range and pass in our values
    headerrange.setValues(values);

    var firstRow = 2;
    var lastRow = sheet.getLastRow();
    var hexValues = sheet.getRange(firstRow, 1, lastRow - firstRow + 1).getValues(); // get an 1 dimensional  array of the hex values
    var hslValuesOutput = []; // array that stores all results from the conversion function
    var hslOutputRangeName = ("b" + (firstRow) + ":g" + (lastRow)); // write to 6 cells on each line - hex h s l band bandnumber
    var hslOutputRange = sheet.getRange(hslOutputRangeName); // set range to output
    var swatchRange = ("a" + (firstRow) + ":a" + (lastRow)); // colour these in


    // loop over range
    for (var x = 0; x < (lastRow - firstRow + 1); x++) {
        hslValuesOutput.push(hexhslconvert(hexValues[x][0]));
        // get HSL calculation back as an array and add it to the output array 
        // setValues() and getValues() always use array of arrays ([[]]) as argument
    }

    hslOutputRange.setValues(hslValuesOutput); // output the h s l values etc
    makeswatches(swatchRange); // colour the swatches in
}

function hexhslconvert(hexvalue) {
    // does the conversion #hex to hsl and returns h s l values separated in an array hslvalues
    var hslvalues = new Array(4);
    var hextext = hexvalue.toString().replace("#", ""); // hex value minus #
    var r = (parseInt(hextext.substring(0, 2), 16)) / 255; // separate R in hex and make a fraction of 1
    var g = (parseInt(hextext.substring(2, 4), 16)) / 255; // separate G 
    var b = (parseInt(hextext.substring(4, 6), 16)) / 255; // separate B
    var colourband;
    var colourbandnumber;

    // Find greatest and smallest channel values
    var channelMin = Math.min(r, g, b),
        channelMax = Math.max(r, g, b),
        delta = channelMax - channelMin,
        hue = 0, // hue or colour value
        saturation = 0, // saturation or colour amount
        luminosity = 0; // luma or brightness

    // Calculate the hue
    if (delta === 0) // No difference
        hue = 0;

    else if (channelMax === r) // Red is max
        hue = ((g - b) / delta) % 6;

    else if (channelMax === g) // Green is max
        hue = (b - r) / delta + 2;

    else // Blue is max
        hue = (r - g) / delta + 4;

    hue = Math.round(hue * 60);

    // Give negative hues a correct positive value
    if (hue < 0)
        hue += 360;

    // Calculate luma
    luminosity = (channelMax + channelMin) / 2;

    // Calculate saturation
    saturation = delta == 0 ? 0 : delta / (1 - Math.abs(2 * luminosity - 1));

    // Multiply luma and sat by 100
    saturation = +(saturation * 100).toFixed(0);
    luminosity = +(luminosity * 100).toFixed(0);

    colourband = findcolourband(hue, saturation, luminosity)

    switch (colourband) {
        case "Black":
            colourbandnumber = 1
            break;
        case "Grey":
            colourbandnumber = 2
            break;
        case "White":
            colourbandnumber = 3
            break;
        case "Red":
            colourbandnumber = 4
            break;
        case "Pink":
            colourbandnumber = 5
            break;
        case "Orange":
            colourbandnumber = 6
            break;
        case "Brown":
            colourbandnumber = 7
            break;
        case "Yellow":
            colourbandnumber = 8
            break;
        case "Green":
            colourbandnumber = 9
            break;
        case "Blue":
            colourbandnumber = 10
            break;
        case "Purple":
            colourbandnumber = 11
            break;
    }

    hslvalues[0] = hextext;
    hslvalues[1] = hue;
    hslvalues[2] = saturation;
    hslvalues[3] = luminosity;
    hslvalues[4] = colourband;
    hslvalues[5] = colourbandnumber;
    return hslvalues;
}

function makeswatches(colourhexrange) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    var r = sheet.getRange(colourhexrange);
    r.setBackgrounds(r.getValues())
}

function findcolourband(h, s, l) {
    var hue = h;
    var sat = s;
    var luma = l;
    var colourband = "not set";

    if (luma > 95) {
        colourband = "White";
    }

    else if (luma < 15) {
        colourband = "Black";
    }

    else if (sat < 4) {
        colourband = "Grey";
    }

    // the rest of these are not grey

    else if (hue <= 33 && luma < 30) { // special cases first
        colourband = "Brown";
    }

    else if (hue <= 33) {
        colourband = "Orange";
    }

    else if (hue >= 33 && hue < 40 && luma < 30) {  // special cases first
        colourband = "Brown";
    }

    else if (hue >= 33 && hue < 60) {
        colourband = "Yellow";
    }

    else if (hue >= 60 && hue < 160) {
        colourband = "Green";
    }

    else if (hue >= 160 && hue <= 260) {
        colourband = "Blue";
    }

    else if (hue >= 260 && hue <= 295) {
        colourband = "Purple";
    }

    else if (hue >= 295 && hue <= 345) {
        colourband = "Pink";
    }

    else if (hue >= 345 && hue <= 361 && luma > 60) {
        colourband = "Pink";
    }

    else if (hue >= 345 && hue <= 361) {
        colourband = "Red";
    }



    return (colourband)

}