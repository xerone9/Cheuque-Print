# Cheuque-Print
Bulk Cheque Printing App (Under Development)

Create Excel File with Party names and amount and press print. The command will directly sent to printer and it will start printing cheques. Only Askari Bank and Al Habib Bank templates are available

**Mechanics:**

1- Upload excel file and press any one template "Askari" or "Al Habib" will add more bank later onwards

2- The app will starts a loop by reading the excel file line by line

3- It'll set date as current date

4- Name will be fetched from the excel file

5- Amount will be fetched also from file

6- Amount translation into word will be done via amount_to_millions.py file

7- Then it'll create a picture with respect to bank layout and that picutre is send for printing

8- Once the picture is printed. The picture will automatically get removed

**Current Issues:**

I am using Tim Golden's approach to print jpg file directly sending command to printer but jpg files has issue that once the picture is created its not crisp as it should be it's a little bit blur and also eventhoug I set white color in the backgroud yet it prints a little greyish shade.

The problem will get resolved if I create svg file instead of jpg file but I am unable to print svg file directly send that command to printer

**Future Update**

Now currently its only working as bulk printing. I will add field in which user can directly enter name and amount for single cheque printing

Will add a UI to create/design your own cheque
