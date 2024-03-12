# Licence

    SPDX-License-Identifier: Apache-2.0
    
    Copyright 2024 Lou Lerner
    
    This source code is licensed under the Apache License, Version 2.0.
    A copy of the license is included in the root directory of this project as LICENSE.txt.

# CLD() - Compact Letter Display Function for VBA

    CLD() is a VBA function that generates compact letter displays for post hoc tests in statistical
    analysis in an excel sheet. Used after the test, it allows you to easily identify significant
    differences between groups in your data.

    The methodology used for the code functioning is described here:

    Hans-Peter Piepho (2004) An Algorithm for a Letter-Based Representation of All-Pairwise Comparisons,
    Journal of Computational and Graphical Statistics, 13:2, 456-466, DOI: 10.1198/1061860043515 

## /!\ WARNING /!\

    This function has not been thoroughly tested yet, and must still be considered experimental.

## Requirements

    CLD() requires Microsoft Excel and has been tested on Excel 365.

## Installation

    1. In Excel, if the developper tab is not activated, on the File tab, go to
    Options > Customize Ribbon > Main Tabs, select the Developer check box.
    2. Go to the Visual Basic Editor by pressing Alt + F11.
    3. Click Insert > Module to create a new module.
    4. Copy and paste the CLD() function code into the module.

## Usage

    An exemple .xlsm file (.xls with executable code integrated) is included in the repository. One or
    more warning appear before opening these types of files.

    The CLD() function is used in a cell, or more specificaly, in as many cells as there are distinct
    modalities/means which have been compared.

    If n modalites are compared to eachother, then n cells have to be filled with the function, with only
    the "Cell" argument (see below) changing.

    ----------------

    Formula must be structured as such: =CLD(Range1, Range2, Cell, Alpha, Order)

    -Range1 is the adress of a range of cells containing a 2 column table with the modality names in the
    1st column and their values in the 2nd. Example:

    Vineyards       55
    Orchards        28
    Wheat           29
    Maize           67

    -Range2 is the adress of a range of cells containing the comparison matrix, where each cell represents
    a comparison between two means. The matrix should display the test results in the upper right part, and
    the rest must be empty cells. Example:

                        Vineyards    Orchards     Wheat     Maize	
    Vineyards	                  0,05        0,03       1,1
    Orchards                                      1,5        2,1
    Wheat                                                    2,1
    Maize

    -Cell is the adress of a cell (also included in Range1) containing the name of the modality from which
    the letter has to be displayed in the cell containing the formula
            
    -Alpha is the alpha value before which the modalites are considered significantly different. For a pvalue
    of <0,05 Alpha must be set to 5.

    Order has 2 possible values: "normal" and "descending" (keep the quotation marks). Descending puts the
    first letters of the alphabet to the modalities with the highest values, as is the norm. Normal just
    puts the first letters to the first modalities.

    ---------------

    If the comparison matrix has missing (see exception) or unexpected values, the cell containing the fuction
    will apear blank.

    Exception: If both the line and the column of a modality in the matrix have no values, then the modality will
    be considered absent, and the code will still work for the present modalities.
