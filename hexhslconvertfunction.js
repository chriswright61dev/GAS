function hextohsl() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  //Set the values we want for headers
  /*
     var values = [
     ["Hexcode", "Hex Number", "Hue", "Saturation", "Luminosity", "Notes"]
     ];
    
    // Set the range of cells for the header
    var headerrange = sheet.getRange("A1:F1");
    // Call the setValues method on range and pass in our values
    headerrange.setValues(values);
       */

  var rowOffset = 2;
  var lastRow = sheet.getLastRow();
  var hsloutputrange = "b2:e2";
  var hslvaluesOutput = new Array(4); // stores results from conversion function
  var hexInput = "this is the hexinput";
  var outputvalues;
  // loop over range
  for (var x = 0; x < (lastRow - rowOffset + 1); x++) {
    var inputRange = "a" + (x + rowOffset); // concatenate to give input cell
    var outputRange = "b" + (x + rowOffset) + ":e" + (x + rowOffset); // write to 4 cells hex h s l
    //  grab a hex value
    hexInput = sheet.getRange(inputRange).getValue(); //  get raw hex values - NO QUOTES around input range variable
    // run function to  remove # and generate h s l values
    //hslvaluesOutput = hexhslconvert(hexInput); // array of values to output
    //outputvalues = [hslvaluesOutput];
    outputvalues = [hexhslconvert(hexInput)];
    hsloutputrange = sheet.getRange(outputRange);
    hsloutputrange.setValues(outputvalues);


    //    hsloutputrange.setValues(hslvaluesOutput); // setValues() and getValues() always use 2 dimension arrays

    // the setValues function (note the plural), takes an array of arrays ([[]]) as argument:

    //     In this example, we have a spreadsheet with data from A1 to H8, but we want to update only two cells: D7 and G6. The rest of the cells should be unchanged.

    /*
      
      
      var values = [
  [ "2.000", "1,000,000", "$2.99" ]
];

      
var range = SpreadsheetApp.getActiveSpreadsheet().getRange("A1:H8");
var values = range.getValues();
values[6][3] = "This is D7";
values[5][6] = "This is G6";
range.setValues(values);
Note how we use range.getValues() to get an array of the existing values. Then we set the individual values in that array. Since JavaScript array indexes are 0-based, row 7 will be 6, and column D, being the 4th column, will be 3 - hence values[6][3]. Similarly, G6 will be values[5][6].

Finally, we write the array back to the range. We can update as many values we like in the array, while keeping the number of spreadsheet function calls low - only a single getValues and a single setValues.
      
      */
  }
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

  // Make negative hues positive behind 360°
  if (hue < 0)
    hue += 360;

  // Calculate luma
  luminosity = (channelMax + channelMin) / 2;

  // Calculate saturation
  saturation = delta == 0 ? 0 : delta / (1 - Math.abs(2 * luminosity - 1));

  // Multiply luma and sat by 100
  saturation = +(saturation * 100).toFixed(1);
  luminosity = +(luminosity * 100).toFixed(1);

  hslvalues[0] = hextext;
  hslvalues[1] = hue;
  hslvalues[2] = saturation;
  hslvalues[3] = luminosity;
  return hslvalues;
}