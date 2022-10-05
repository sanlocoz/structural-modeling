# structural-modeling

## What is this program about?
This program is about automating strucural modeling tasks using CSi software (ETABS).
CSi offers front end framework that can be accessed through many programming languages such as Python, Matlab, C++, C# and VBA.
The documentation of the interface could be found in this relative path (assuming we are in application installation folder):

`.\CSi API ETABS v17.chm`

or for an example of default ETABS 18 installation in windows 10 will result in absolute path as follows:

`C:\Program Files\Computers and Structures\ETABS 18\CSi API ETABS v17.chm`

This is the screenshot of the introduction to CSi API (application programming interface):

<img src="img/1.PNG" alt="CSi API" width="700"/>

*Introduction to CSi API*

## Limitations
The program is just to automate the modeling process of simple 3D space frame as shown in figures below:
<br> <br> <img src="img/2.png" alt="CSi API" width="300"/>
<br> *Plan view (X-Y axis plane)* <br> <br> <img src="img/3.png" alt="CSi API" width="250"/>
<br> *Side view (Y-Z axis plane)* <br> <br> <img src="img/4.png" alt="CSi API" width="250"/>
<br> *Front view (X-Y axis plane)*

Besides the configuration, users also can model many materials and section properties using Excel to later interact with the programming interface.


**NB: Structural optimization is out of the scope of this program.**

## Parameter of the program

The objective of the program is to automate modeling process of 3D simple space frame and their section properties.
Parameter of the program are in the Excel document and retrieved to Python program and will be supplied to API from CSi. This is the list of the parameter used in the automation of the modeling:

- Geometry properties <br>
Jumlah Lantai = 
Tinggi Lantai Tipikal
Tinggi Lantai Bawah
Jumlah Bentang X
Jumlah Bentang Y
Bentang X
Bentang Y

It consists of the configuration in x, y, and z axis, dimensions and number of members in each direction.


- Material properties
Material properties is defined for several parameter such as ultimate compression stress for concrete (The modulus of elasticity is derived by empirical formula for conventional reinforced concrete)"
- Section properties
- Typical dead load
- Typical live load



CSI ETABS front end accessing with Python programming language </br>
Reducing modeling time from 1 hour to 5 minutes, the input data is in spreadsheet

## Future work suggestions