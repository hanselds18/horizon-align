# horizon-align
Introducing HorizonAlign, a powerful application designed to shorten the process of horizontal alignment design and calculation. Whether you're an engineer, architect, or construction professional, this app is your go-to solution for precise alignment design, azimuth calculation, and visual representation, all seamlessly integrated with the Unity platform for showing the visual (NOT FOR SALE).

With HorizonAlign, you can design of over 270 predefined points of intersection, ensuring optimal running and reducing the need for manual adjustments. This comprehensive collection covers various scenarios, including road networks, railway tracks, and other linear infrastructure projects.

The app's intuitive interface allows you to effortlessly define your alignment parameters, such as starting and ending points, curve type from an excel file. HorizonAlign automatically calculates the azimuth based on your inputs of coordinate points, ensuring accuracy and efficiency in your designs.

One of the key features of HorizonAlign is its seamless integration with the Unity engine, enabling you to visualize your alignment in a 2D environment. This interactive visual representation provides a clear understanding of how the alignment will look in the real world, allowing you to make informed design decisions and optimize your project's execution.

Furthermore, HorizonAlign empowers you to export your alignment data to a Microsoft Excel file, streamlining your workflow and facilitating collaboration with other stakeholders. The exported file includes all the necessary design parameters, azimuth calculations, and relevant data, ensuring easy sharing and compatibility across platforms.

In addition to its design capabilities, HorizonAlign doubles as a preliminary design tool. By specifying the minimum and maximum range of radius design and speed design, the app provides you with valuable insights into the potential scope of your project. This functionality helps you explore different design possibilities, assess feasibility, and make informed decisions before committing to a final alignment.

HorizonAlign sets a new standard for horizontal alignment design applications, combining robust calculation algorithms, immersive visualizations powered by Unity, and seamless data integration with Microsoft Excel.

# Installation
1. Extract the v1.4-stable.zip
2. Make sure "logo.ico", "UI folder", & "Tabel Rek.Geometrik Jalan.xlsx" are at the same folder as "main.exe"
3. If there are not those 3 files, make sure to show the hidden files by view --> show --> hidden items on Windows 11
4. Run the "main.exe"

# Known Limitation
1. Windows 11 tested (another vers. of Windows yet to test)
2. Azimuth & Horizontal Alignment are not sync
3. Only ".xlsx" format that can be supported for Horizontal Alignment
4. Only support for Urban Roads design
5. Using Indonesia's rule for 2021 Road Geometric Design Guideline for the main guideline & The Green Book 7th Edition of AASTHO for complementary

# How to Use
1. **Azimuth** menu
   1. Fill in the box how many point that want to be calculated
   2. Pop-up window will be shown. Fill in the coordinate points for X & Y points
   3. Click submit button
   4. The results will be shown in the main menu

2. **Horizontal Alignment** menu
   1. Choose ".xlsx" file that want to be calculated
   2. Screening window will be shown
   3. Click submit button
   4. The results will be shown in the form of table
   5. If app cannot find the right answer, a warning message will be shown and shows where the error occured
   6. Click the "Save to .xlsx" button to convert and save it to Microsoft Excel
   7. Click the "Save to .txt" button to save the coordinate point form the calculation

3. **Unity Files**
   1. Extract Editor.zip
   2. Create new 2D scene
   3. Copy "coordinate.txt" & "coordinate_pi.txt" from the app's calculation into the **assets** folder
   4. Copy "editor" folder into the **assets** folder
   5. The coordinates will be drawn in the "Scene" view
