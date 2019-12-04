# VBA-Convert Coordinates
 An Excel function that converts a coordinate (latitude and/or longitude) to DMS (degrees, minutes, seconds) or decimal degrees.  

 The idea was to create a function that could parse latitudes and longitudes in a wide range of formats, allowing it to handle most formats used by programs and websites.

 ## Installation
Import the .bas file into your spreadsheet or use the included sample excel file


 ## Usage
DMS2DEC(coordinate, output_format)

**Coordinate**  Required. The text representation of coordinate that you want to convert.

**Output_format**  Optional. The number -1, 0, or 1. The **output_format** argument specifies the format for the returned value. The default value for this argument is 0.

```
output_format       Result
0 or omitted        Coordinates returned in decimal degrees (d.dddd) format

1                   Coordinates returned in decimal degrees, with Easting and Northing indicators where possible (dd.dddd[NESW])

-1                  Coordinates returned in DMS format, with Easting and Northing indicators where possible (ddº mm' ss.ss"[NESW])
```

 ## Examples

 ```VB.net
 =DMS2DEC("s53 03 47.7")      => -53.06325             ' Defaults to dd.mmmm format
 =DMS2DEC("s53 03 47.7", 1)   =>  53.06325S            ' dd.mmmm with Easting or Northing indicator
 =DMS2DEC("s53 03 47.7", -1)  =>  53° 03' 47.70"S      ' Includes Easting or Northing indicator
 ```

 ## Test cases
 Test cases are included, and the output should look like this:
```
Coordinate		        Decimal degrees
N 53 03 47.7        =>	53.0632500
s53 03 47.7         =>	-53.0632500
s53 03 47.75		=>	-53.0632639
E53 03 47.7		=>	53.0632500
w53 03 47.7		=>	-53.0632500
53 03 47.7 n		=>	53.0632500
53 03 47.7S		=>	-53.0632500
53 03 47.7e		=>	53.0632500
W 53 03 47.7		=>	-53.0632500
N53 03 47.7		=>	53.0632500
s.53.03.47.7		=>	-53.0632500
E53.03.47.7		=>	53.0632500
w53..03..47.7		=>	-53.0632500
n53.03.47.7		=>	53.0632500
53.03.47.7 S		=>	-53.0632500
53.03.47.7e		=>	53.0632500
53..03..47.7W		=>	-53.0632500
N.53.03.47.7		=>	53.0632500
s53.03.47.7		=>	-53.0632500
E-53-03-47.7		=>	53.0632500
w53-03-47.7		=>	-53.0632500
n53--03--47.7		=>	53.0632500
S53-03-47.7		=>	-53.0632500
53-03-47.7 e		=>	53.0632500
53-03-47.7W		=>	-53.0632500
53--03--47.7N		=>	53.0632500
s-53-03-47.7		=>	-53.0632500
E53-03-47.7		=>	53.0632500
53.06325w			=>	-53.0632500
53.06325			=>	53.0632500
53.06325 s		=>	-53.0632500
53.06325 N		=>	53.0632500
-53.06325			=>	-53.0632500
53 03 47.7		=>	53.0632500
-53 03 47.7		=>	-53.0632500
53 03 47.7W		=>	-53.0632500
-53º 03' 47.7"		=>	-53.0632500
53º 03' 47.7"		=>	53.0632500
53º 03' 47.7"s		=>	-53.0632500
53º 03' 47.7''s	=>	-53.0632500
53º 47.7''s		=>	-53.0132500
47.7''s			=>	-0.0132500
N 144 35 26		=>	144.5905556
s144 35 26		=>	-144.5905556
E 144 35 26		=>	144.5905556
144 35 26w		=>	-144.5905556
N144.35.26		=>	144.5905556
144.35.26s		=>	-144.5905556
-144 35 26S		=>	-144.5905556
144º 35' 26"w		=>	-144.5905556
45°43'51''N		=>	45.7308333
009°44'23''E		=>	9.7397222
009°44'23"E		=>	9.7397222
53º 03' 0"		=>	53.0500000
53º 03'			=>	53.0500000
53º 00' 47.7"		=>	53.0132500
53º 47.7"			=>	53.0132500
W-53-03-47.7		=>	-53.0632500
53º 03.55'		=>	53.0591667
  1° 22.011'N		=>	1.3668500
  1°22'0.66"N		=>	1.3668500
53-03-47.71W		=>	-53.0632528
53-03-47.72W		=>	-53.0632556
53.0500° N		=>	53.0647361
53.0500°N			=>	53.0647361
53.0500ºN			=>	53.0647361
-53 03 47.7, E 53 03 47.7	=>	-53.06325, 53.06325
S 53 03 47.7, -53..03..47.7	=>	-53.06325, -53.06325
S 53 03 47.7, E 53 03 47.7	=>	-53.06325, 53.06325
N 53 03 47.7, E 53 03 47.7	=>	53.06325, 53.06325
48.8566° N, 2.3522° E		=>	48.8701713, 2.3528534
48.8566°N, 2.3522°E			=>	48.8701713, 2.3528534
```
 ## License

 MIT
