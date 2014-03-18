The given code firstly converts unstructured data to structured data (by converting it to csv form), then formulated the rules of association based on the algorithm of coherent rule mining (confusion matrix). There are following main functions :

1. conn () : Establishes connection with SQL Server 2008 database with the help of connection string

2. sql_grid_execute() : Fills the adapter and sets the table name. This function basically ensures whether the connection is established appropriately with the database

3. grid () : Sets the data source 

4. Form1_Load () : Calls conn() function and loads the form

5. sql_execute() : Executes the given sql query 

6. Button1_Click() : On click of this button, the value (tablename) as selected from drop down is set for the variable 

7. Button2_Click() : On click of this button, the query is executed for the selected tablename, to display the contents of the table

8. Button3_Click() : On click of this button, number of columns are retrieved

9. name_array() : saves the columns of the selected table into a variable

10. MyCheckboxes_CheckedChanged() :  Applying domain driven concept, it checks the checkboxes whether checked or unchecked

11. csv1() : displays the comma seperated data of the selected columns

12. chckarray() : makes checkboxes in form 2 for the selected checkboxes on form 1

13. Button1_Click_1() : Prepares radio buttons of the selected checkboxes in form2
