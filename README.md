Creo_v1
=======

This is a simple webpage designed to provide Creo automation through PTC's integrated Pro/Weblink API. 
The site is hosted using github pages at: [dev.okdane.com](http://www.dev.okdane.com)

-Current Status-
I've worked through several of the non-trivial cases (atleast non-trivial to me). This included:
1. Listing all the *.drw's in your active workspace.
2. Exporting a pdf with custom settings (the settings were no-pdf_viewer and mono-color).
3. Based on the part number the drawing is given a target directory. This is broken out into a dirTarget function (or something like that), so you can change it to suit your needs.
4. A data grid is populated with relevent information using jqGrid. This information contains the drawing description, target directory, and part number.
5. If you click the column header for Color and Target Directory, it will toggle through all the options for every drawing at once.
6. The data grid allows you to toggle between the target directory or just the desktop, toggle whether to override the existing file, and what color option to export it with.n

-Next Areas of Improvement-
1. Expand the function that gets the description for the drawing. A lot of companies don't actally draw the description from the drawing parameneter, they use the Description from the part. My function currently just assumes the part used in the drawing has the same partnumber as the .drw, and otherwise returns Not Available. I plan on eventually expanding this function to determine the active part in the drawing and then pulling the description.
2. Make jqGrid 
3. Change the OverRide/Re-use and Color options into icons instead of text that indicates which option is selected.
4. Make the target directory bold if they have changed it to target the Desktop. This currently works, but it only bolds the column after sorting it.
5. Refactor all the functions into a more appropriate library. This was just kind of a prototype and proof of concept. It would be nice to refactor it and make it work for DXFs, STLs, etc.

Feel free to use any of this, but I would hugely appreciated if you'd teach me anymore you figure out.

D
