# Auto-Report-Scrapper


# **How to use ?**

+ open the python file using any text editor
+ Replace  excel_file = r'C:\Users\GAMERXXX\Desktop\stats\trading1.xlsx'  with your path to your excel journal
               sheet_name = 'Sheet1'   Replace with the actual sheet name in your journal
+ Make sure that your excel journal is closed and run the script ! 



in brief 
the code reads HTML data, extracts and filters relevant information, organizes it into a DataFrame, and appends it to an existing Excel file while adjusting formatting.
It provides user-friendly success messages and handles potential errors during execution.



## User Input:
+ Asks the user to enter the name of an HTML file.
       
## HTML Processing:
+ Reads the HTML content from the specified file using BeautifulSoup.
       
## Data Extraction:
+	Extracts relevant data from the HTML, focusing on a specific <td> element with the title "Commission + Swap + Profit + Taxes" and a <td> element with the string "Ticket."
	Appends the extracted data to a list, excluding the last 7 elements.
       
## Data Filtering:
+ Identifies indices of occurrences of 'balance' and 'cancelled' in the extracted data.
	Creates a list of indices to drop based on the identified indices.
	Filters the original data to exclude the identified indices.
       
## Data Formatting:
+	Organizes the filtered data into rows, assuming each row contains 14 elements.
	Defines column names based on the first row of the organized data.
	Creates a DataFrame using pandas with the organized data.
       
## Data Cleaning and Formatting:
+	 Performs various operations on the DataFrame to clean and format the data.
	 Drops unnecessary columns, renames columns, adjusts data types, and converts date and time columns.
       
## Excel File Operations:
+	 Specifies the target Excel file and sheet name.
	 Loads the existing Excel file data into a DataFrame.
	 Determines the starting row for appending new data.
	 Appends the DataFrame with the organized and formatted data to the specified sheet in the Excel file.
	 Opens the Excel file to adjust formatting based on the second row of the existing sheet.
       
## GUI Success Message:
+	 Displays a success message using a tkinter GUI messagebox, indicating that the data has been modified successfully.
       
## Error Handling:
+	 Uses try-except blocks to handle exceptions, specifically FileNotFoundError and a more general Exception. Prints error messages if an exception occurs.

       


