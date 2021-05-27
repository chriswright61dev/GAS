function hextohsl() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  //Set up the header with array
  var values = [
    ["Hexcode", "Hex Number", "Hue", "Saturation", "Luminosity"]
  ];
  // Set the range of cells for the header
  var headerrange = sheet.getRange("A1:E1");
  // alternatively we can use sheet.getRange(row, column, numRows, numColumns) - numeric

  // Call the setValues method on range and pass in our values
  headerrange.setValues(values);

  var firstRow = 2;
  var lastRow = sheet.getLastRow();
  var hexValues = sheet.getRange(firstRow, 1, lastRow - firstRow + 1).getValues(); // get an 1 dimensional  array of the hex values
  var hslValuesOutput = []; // array that stores all results from the conversion function
  var hslOutputRangeName = ("b" + (firstRow) + ":e" + (lastRow)); // write to 4 cells on each line - hex h s l 
  var hslOutputRange = sheet.getRange(hslOutputRangeName); // set range to output
  var swatchRange = ("a" + (firstRow) + ":a" + (lastRow)); // colour these in
  // loop over range
  for (var x = 0; x < (lastRow - firstRow + 1); x++) {
    hslValuesOutput.push(hexhslconvert(hexValues[x][0]));
    // get HSL calculation back as an array and add it to the output array 
    // setValues() and getValues() always use array of arrays ([[]]) as argument
  }

  hslOutputRange.setValues(hslValuesOutput); // output the h s l values
  makeswatches(swatchRange); // colour the swatches in
}

function hexhslconvert(hexvalue) {
  // does the conversion #hex to hsl and returns h s l values separated in an array hslvalues
  var hslvalues = new Array(4);
  var hextext = hexvalue.toString().replace("#", ""); // hex value minus #
  var r = (parseInt(hextext.substring(0, 2), 16)) / 255; // separate R in hex and make a fraction of 1
  var g = (parseInt(hextext.substring(2, 4), 16)) / 255; // separate G 
  var b = (parseInt(hextext.substring(4, 6), 16)) / 255; // separate B

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

  hslvalues[0] = hextext;
  hslvalues[1] = hue;
  hslvalues[2] = saturation;
  hslvalues[3] = luminosity;
  return hslvalues;
}

function makeswatches(colourhexrange) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var r = sheet.getRange(colourhexrange);
  r.setBackgrounds(r.getValues())
}