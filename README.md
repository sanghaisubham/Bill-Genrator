# Bill-Genrator
This project is aimed at generating bill in proper invoice format from an excel sheet consisting of the required product and customer details.

## Explaination

The Project contains a Excel file named as "Paavan Dairy-automated-invoice.xlsx" containg few sheets including:

+ Clients: This sheet contains the Details of the Client including :

                + Client Number(Client nr)
                
                + Client Name(Name)
                
                + Client Address(Street)
                
                + Client City(City)
                
                + Cient Contact (Phone Number)
        
+ Products: This sheet contains the Details of the Products in the Paavan Dairy Shop including :

                + Product Number(Product nr)
                
                + Product Name(Description)
                
                + Product Price(Price)
                
+ Sheet1: This sheet contains the Details of the Products bought by a customer during a Particular Week including :

                + Date of Purchase(Date)
                
                + Month of Purchase(Month)
                
                + Client Number(Client Number)
                
                + Client Name(Client Name)
                
                + Product Number(Product)
                
                + Product Quantity(Quantity)
                
 ## Problem 
 
 We have to genrate a weekly invoice for all the  customers containing all the products bought by him during the week with the invoice in the form as given in the Invoice Sheet of "Paavan Dairy-automated-invoice.xlsx" file.
 
 ## Solution Approach
 
 1. The invoice format was first prepared using html and css saved under the filename "Template.html"
 
 2.Using Pythons Pandas and DataFrame knowledge, the Customer Details and the Product Details were brought at a same place,i.e A
   single dataframe was created under the name Merged_Sheet.csv
   
 3.Now the html tags were searched in the html file and the required input was picked from the dataframe created above, to fit 
   in all the required customer and product details in the perfect invoice format.
   
 4.The script was run on an excel sheet which is generated for a particular week and all the generated Invoice in Html Format
   was stored at a new location(i.e A new Folder with the Date of Billing as the Folder Name) with the html file named as
   CustomerName+A Unique Invoice ID.
  
  
  ## Result
  Here ,the script("Invoice_Genration_Script.iypnb") was run on "Paavan Dairy-automated-invoice.xlsx" file which created a 
  folder of the billing date with all the Invoices.
