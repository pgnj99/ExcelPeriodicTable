# Microsoft Excel Periodic Table
This is an interactive periodic table with chemistry-themed calculators built using Microsoft Excel, VBA, and macros. I created this as a final project for my Introduction to Engineering Computation in May 2018, and I consider this to be the first real programming project I've written. The assignment was to pick a project to do, and I chose a periodic table because the front-end is easy to visualize and I liked chemistry at the time. The final implementation is certainly flawed but I'm very proud of what I was able to accomplish back then, so I've uploaded it here as a way to preserve it.

To use this, you'll need to open it in Microsoft Excel and enable macros. The temp1 folder contains external assets needed for certain windows. The frm and frx files were exported from Excel, the former of which will show VBA code.

## Array Information
The spreadsheet shows all elements on the periodic table and the data used for its main screen, including name, symbol, atomic number, atomic mass, neutrons, ionic charge, melting point, boiling point, and their default color when first booted up. The button labeled "Open Period Table" will open the front-end when clicked with macros enabled.

Note that, as this project was completed in 2018, values labeled "uncertain" may have become known over time.

## Periodic Table
The main page of the front-end is structured like a periodic table, with each element represented by a button and displaying its atomic number, symbol, and atomic mass. The color of the button represents its class, and the color of the symbol represents its state at the given temperature, which is 273 K by default.

By typing in a new temperature at the top and clicking "Change Temperature", the colors of the symbols will change to reflect their states based on their melting points and boiling points. The list can also be filtered in various ways, such as by clicking on one of the icons at the top to only show elements of that state or class. Additionally, the "Search" button in the upper right corner can be used to search for an element by either its name or its symbol, which will then exclusively show that element. Pressing the "Reset" button will make all elements visible again and restore the temperature back to default.

Each element can be clicked on to show a new window containing information pertaining to the element. This includes its name with phonetic spelling, a button for a voice to pronounce it, its data, a brief description, and an image of something the element can be found in. In the project's current implementation, only three elements have windows like this: hydrogen, helium, and oxygen.

## Calculate
By pressing the "Calculate" button, a new window will open that gives a choice between four different chamistry-related calculators.

### Protons, Neutrons, Electrons
This calculator allows the user to enter the name or symbol of up to four elements with an optional field for charges. The amount of protons, neutrons, or electrons will appear at the bottom when the respective button is pressed.

### Molarity
This calculator contains two sections for molarity-related calculations. The first contains five fields of units: moles, liters, molarity, molar mass, and grams, and pressing upon entering a value, the buttons below allow the user to convert that value to a different unit. The second requires that the user enters both the molarity and liters of one substance and either unit of a second substance to calculate the correct amount of the missing unit for the substances to dilute.

### Ionic Compounds
This calculator accepts two elements and an optional field for their charges and generates the symbol of the compound they would form. If the sum of the charges is not zero, then the calculation will fail.

### Percent Composition
This calculator contains two sections for percent composition. The first accepts up to four elements in an empirical formula and their amounts, if there is more than one, and calculates their percent composition. Alternatively, the second accepts up to four elements and their respective percent compositions and then generates the empirical formula.

## Areas to Improve
As this was written on a strict time limit as the first programming project I've ever done, there are numerous flaws that could be improved on.
1. As previously stated, only three elements got description windows, and I manually made each one. There is likely a better way to automate this using the spreadsheet.
2. The search function and calculators are case-sensitive and require that the first letter of whatever element/symbol you type is capitalized.
3. The buttons on both the main window and the calculators could be organized in a much cleaner way.
4. The reset button does not change the number in the temperature box back to 273.
5. There is no way to seperately reset the temperature and what elements are visible.
6. The calculator often crashes when an invalid element is entered.
7. I haven't taken a chemistry class in years so I can't speak on how accurate the actual calculators are, but even when presenting this project I remember the percent composition calculator did not return the correct empirical formula.
