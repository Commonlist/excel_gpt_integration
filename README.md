# excel_gpt_integration

A VBA function for integrating Excel with the GPT model using OpenAI's API. This function allows users to automate intelligent analysis and data processing directly within Excel spreadsheets.

## Introduction

The `GPT` function is designed to facilitate the integration of Excel with the GPT model from OpenAI. It takes user input and cell data as arguments, sends them to the GPT model, and returns the generated response. This allows for seamless automation of complex data analysis and processing tasks within Excel.

## Features

- **Integrates Excel with GPT:** Uses VBA to send requests to OpenAI's GPT model and retrieve responses.
- **Automates Data Processing:** Automates intelligent data analysis and processing directly within Excel.
- **Handles API Requests:** Manages HTTP requests and responses, including error handling.

## How It Works

1. **Preparing the Request:**  
   The function concatenates the user's prompt with the content of the specified cell to create the request context.

2. **Sending the Request:**  
   An HTTP POST request is sent to the OpenAI API with the prepared context.

3. **Handling the Response:**  
   The function parses the JSON response from the API and extracts the relevant information.

4. **Returning the Result:**  
   The function returns the generated response back to the cell in Excel.

5. **Managing Request Rate:**  
   A delay is introduced between requests to comply with the API rate limit.

## Getting Started

### Prerequisites

To use this script, you need to have the following:

- Microsoft Excel with VBA support
- Access to OpenAI's API (API Key required)
- VPN enabled if you are accessing from Russia

### Installation

1. **Clone the Repository:**

    ```bash
    git clone https://github.com/Commonlist/excel_gpt_integration.git
    ```

2. **Prepare the Excel Workbook:**
   - Open Excel and create a new workbook.
   - Save the workbook as an XLSM file (Excel Macro-Enabled Workbook).

3. **Add the VBA Code:**
   - Press `Alt + F11` to open the VBA editor.
   - Insert a new module: `Insert > Module`.
   - Copy and paste the provided VBA code into the module.

4. **Set Up JSON Converter:**
   - Download `JsonConverter.bas` from [VBA-JSON GitHub](https://github.com/VBA-tools/VBA-JSON).
   - Import `JsonConverter.bas` into your VBA project: `File > Import File`.

5. **Set Up API Credentials:**
   - Replace `$OPENAI_API_KEY` in the VBA code with your actual OpenAI API key.

### Configuration

Update the script `GPT` function with your API credentials and required configurations.

### Running the Script

To use the function in Excel, simply enter a formula like the following in a cell:

```excel
=GPT("Your prompt", A1)
```

This will send the content of cell A1 along with your prompt to the GPT model and display the response in the cell where you entered the formula.

### Example Usage

- **Step 1:** Enter your prompt and cell reference in the formula.
- **Step 2:** The function sends the request to OpenAI's GPT model.
- **Step 3:** The response is displayed in the cell.

## Troubleshooting

If you encounter any issues while using the script, here are some common troubleshooting steps:

- **Check API Credentials:** Ensure that your API key is correctly entered in the VBA code.
- **Enable Macros:** Make sure macros are enabled in Excel.
- **JSON Converter:** Ensure that the JSON converter module is correctly imported into your VBA project.
- **Internet Connection:** Make sure your internet connection is stable.
- **VPN:** Ensure that your VPN is enabled if you are accessing from Russia.
- **Script Errors:** If there are any errors in the script, refer to the error message for more details and fix the mentioned issues.

## Contributing

If you would like to contribute to this project, please fork the repository and submit a pull request with your changes.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.

---

By using this script, you can integrate Excel with OpenAI's GPT model and automate intelligent data analysis and processing directly within your spreadsheets.
