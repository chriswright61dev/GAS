function hextohsl() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  
  // Set the values we want for headers
  var values = [
    ["Hexcode", "Hex Number", "Hue", "Saturation", "Luminosity", "Notes"]
  ];

  // Set the range of cells
  var headerrange = sheet.getRange("A1:F1");
  var lastRow = sheet.getLastRow();
  
  // Call the setValues method on range and pass in our values
  headerrange.setValues(values);
  var hexvalue = sheet.getRange("A2").getValue(); // this is the raw hex value with #
  var hextext = hexvalue.toString().replace("#",""); // hex value minus #

  var r = (parseInt(hextext.substring(0,2),16))/255; // separate R in hex and make a fraction of 1
  var g = (parseInt(hextext.substring(2,4),16))/255; // separate G 
  var b = (parseInt(hextext.substring(4,6),16))/255; // separate B

  // Find greatest and smallest channel values
  var cmin = Math.min(r,g,b),
    cmax = Math.max(r,g,b),
    delta = cmax - cmin,
    hue = 0, // hue
    s = 0, // sat
    l = 0; // luma

  // Calculate hue
   if (delta === 0)  // No difference
    hue = 0;
  
  else if (cmax === r) // Red is max
    hue = ((g - b) / delta) % 6;
    
  else if (cmax === g) // Green is max
    hue = (b - r) / delta + 2;
    
  else // Blue is max
    hue = (r - g) / delta + 4;

  hue = Math.round(h * 60);
    
  // Make negative hues positive behind 360°
  if (hue < 0)
      hue += 360;
  
  // Calculate luma
  l = (cmax + cmin) / 2;

  // Calculate saturation
  s = delta == 0 ? 0 : delta / (1 - Math.abs(2 * l - 1));
    
  // Multiply l and s by 100
  s = +(s * 100).toFixed(1);
  l = +(l * 100).toFixed(1);

  //var hexnumbers = parseInt(hextext);      
  var outputrange = SpreadsheetApp.getActive().getSheetByName('sheet1').getRange("b2:f2");
  var hslValues = [[hextext,hue,s,l,lastRow]]; // setValues() and getValues() always use 2 dimension arrays, even if the range is only 1 row high, so you should use the extra brackets
  outputrange.setValues(hslValues);
  
}