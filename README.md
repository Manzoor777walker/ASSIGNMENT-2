# ASSIGNMENT-2
# Data Cleaning And Transformation
Data Cleaning and Transformation in the given set of data by using Excel.										
# 1) Handling Missing Values:		
# If there are products with missing categories, propose a strategy to impute or deal with these missing values effectively.
    • Check for missing values in the 'Price' column.
    • Using conditional formatting to highlight the blank cells
    • Replace missing value by using  IF,ISBLANK, AVERAGE Function
                formula used: [=IF(ISBLANK(J2),AVERAGE($J$2:$J$35),J2)]							
# Correcting Inconsistent Data:									
		• Identify inconsistent text formats present in the "Product Name" column.								
		• Identify typos present in the "Category" column.	here "Electroni"							
		• Using find and replace function to standardize the text formats in the "Product Name" column and fix typos or misspellings in the "Category" column.								                  *By replacing ""Electroni" into "Electronics
                        *Use TRIM,PROPER Function 
                        *implement suitable data type according to the specific data category.
# Removing Duplicates:									
		• No duplicate rows within the dataset based on the entirety of each row, and remove them if any.								
										
# Splitting and Merging Data:									
		• Split the "Product ID" column into two separate columns for " Manufacturing Date" and "Country Code". Remove unnecessary characters, if any.
                                      *Using LEFT and RIGHT Function to splitting the "product ID"
                                      Formula used :[=LEFT(A2,LEN(A2)-3)] and [=RIGHT(A2,2)]
		• Merge the "Brand Name" and "Product Name" columns into one column named "Product Brand".	
                                       *Using "&"function for merging 
                                              Formula used :[=I2& " " &H2]
# Number Formatting:									
		• Formatting the data type of the "Price" column into currency format 								
		• Formatting the "Manufacturing Date" column to display dates in the "DD-MM-YYYY " format(by adding default date)								
										
# Conditional Formatting:									
		• Apply data bar or color scales conditional formatting in the "Price" column.								
		• By creating a custom rule for conditional formatting in the "Category" column to highlight cells where the category is "Electronics."	
                                           


