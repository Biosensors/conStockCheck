# conStockCheck
Consignment Stock Check / RGA Form
This application is used to perform stock counts at customer’s location as well as returns. 
How to use it
1)	Install the Excel Add-in from \\mofasa.besa.bsi.corp\Shared\Everyone\IT\BESA\SAP - Cons Stock\CSC\setup.exe
2)	Open the latest Consignment Stock File (current version is 10.13.03).
3)	Click on “ADD-INS” in the Ribbon, then on “Extract Data” in “SAP Extract”. Enter the customer’s 6 digit number (the same as the one in SAP).
4)	The file gets populated with the relevant data. 
5)	Customer Service employee sends it to the sales rep using the “Save and send CS ONLY” button on the “CONSIGMENT STOCK” tab.
6)	When the sales rep received the file, they fill out the “SCANNED QTY” tab with the relevant information. They then start scanning their products starting from cell B8 using their handheld scanner (Motorola CS3070). As this devices acts as a keyboard with an EN-US layout, some users might run into an issue when scanning products containing numbers and/or the letter Z, W, Y, A, Q. Therefore, when the an array which is supposed to have barcodes scanned into it is selected, a VBA code changes the user’s keyboard layout to EN-US, then switches it back to the original one as soon as that given array is left.
7)	Upon completion, the user send the file back to the appropriate department using the “Save file & send mail” button the “CONSIGNMENT STOCK” tab.

Proceeding with an RGA is roughly the same thing:

1)	The sales rep can use any stock check file given that it’s the most recent version
2)	Go to the “SCANNED RETURNS” tab and fill it out relevantly
3)	Scan the item into column B. The same principles apply as point 6 for consignment stock counts.
4)	Click on “Prepare RGA”
5)	Fill out the “RGA FORM” tab relevantly
6)	Click on “Get RGA Number”
7)	Click on “Save file & send mail”
Notes for developers
It is composed of the following components:
-	An SAP Web Service (ZMM_CON)
-	An Excel Add-in (called “SAP CSC”) developed in C#
-	An Excel file full of VBA code
The latest versions are released into production by placing them in 
\\mofasa.besa.bsi.corp\Shared\Customer Service\consignment stock management
from where the CS team will open them and use them.
Upon opening the file, it will try to ping MOSQ04.besa.bsi.corp. Should this server be available, the file will:
-	Update the decode scan logic from this server
-	Update the customer base
Code / project source
\\mofasa.besa.bsi.corp\shared\IT\BESA\Peter's Projects\Cons Stock Check
