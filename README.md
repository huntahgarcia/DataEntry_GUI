# Bakery Data Entry 

## 1. Overview
This program provides a user-friendly graphical interface for bakery order management. It allows users to input customer details, order information, payment, and order status. The collected data is then automatically saved into an Excel file for further tracking and processing. This tool streamlines the data entry process for bakery owners and staff, ensuring that customer orders are properly recorded and organized.

## 2. Features
- **User Input Forms**: Collect customer details such as name, contact, and address.
- **Order Management**: Allows users to add and remove bakery items with their quantities.
- **Order Status Tracking**: Users can track the status of orders (e.g., Pending, In Progress, Completed).
- **Payment Tracking**: Provides a way to track payment methods (e.g., Cash, Credit Card) and payment status (e.g., Pending, Paid).
- **Calendar Integration**: Includes calendar widgets for selecting order and delivery dates.
- **Excel Integration**: Automatically writes order data to an Excel file, making it easy to maintain and access records.
- **Attention Tracking**: Option to flag orders that need special attention.

## 3. Requirements/Dependencies
1. **Python 3.10.9: Ensure you have Python installed.
2. **PyQt5**: For building the graphical user interface (GUI).
3. **Openpyxl**: For reading and writing Excel files.

To install the dependencies, run:

```bash
pip install PyQt5 openpyxl
```
## 4. Usage Instructions
1. **Launch the Application**:  
   Run the Python script to open the application window.

2. **Enter Customer Information**:  
   Fill out fields for customer name, contact, and address.

3. **Select Bakery Items**:  
   Choose an item from the dropdown (e.g., Sourdough, Cinnamon Rolls) and set the quantity.

4. **Manage Order Status and Payment**:  
   Select the appropriate order status (e.g., Pending) and payment method (e.g., Cash, Credit Card).

5. **Use the Calendar**:  
   Select the start and end dates for the order using the calendar widgets.

6. **Add/Remove Items**:  
   Add the selected items to the order list. You can also remove items if needed.

7. **Submit and Save Data**:  
   Once the form is complete, click "Submit" to save the order data into an Excel file.

8. **Reset the Form**:  
   Use the "Reset" button to clear all fields and start a new order.

9. **Excel File Location**:  
   The order data will be saved in the Excel file located at `/Users/name/Downloads/Bakery_Template.xlsx`.


## 5. Known Issues & Limitations
- **File Path Dependency**:  
  The program uses a hardcoded path for saving and loading the Excel file (`/Users/name/Downloads/Bakery_Template.xlsx`). You may need to adjust this to fit your directory structure.

- **Limited Error Handling**:  
  Basic error handling is in place. Complex scenarios (e.g., invalid input, file access issues) may need further handling.

- **UI Customization**:  
  The application style and look are basic. Users may customize this further according to needs.

## 6. Future Enhancements
- **Dynamic File Path**:  
  Allow the user to select or change the file path for saving Excel data.

- **Enhanced Error Handling**:  
  Implement more comprehensive error handling for cases such as missing fields or file access errors.

- **Multiple Customer Entries**:  
  Add support for multiple orders or entries in one session, and potentially export them in bulk.

- **PDF Export**:  
  Offer an option to export order summaries as PDFs.

## 7. Testing
This program has been tested with basic input scenarios:
- Customer orders with various items were successfully added to the order list.
- Data was successfully written to an Excel file without issues.
- The date selection functionality works correctly with validation for start and end dates.

Testing was conducted using typical bakery items and customer data. The application was able to handle multiple inputs and save them in the provided Excel file.


![Data Entry GUI](https://github.com/huntahgarcia/Data_Entry_GUI/assets/140458716/307911d5-77e9-43f9-b1fd-da043530e75d)
