# xlwings-ollama-plugin

Excel plugin leveraging `xlwings` and the Ollama API to generate AI completions. Reads input text from a specified range and writes completions to adjacent cells. Supports model selection via a dedicated cell, enabling seamless integration of AI capabilities into your spreadsheets.

## Installation and Setup

To install the required dependencies and set up the plugin, run the following commands:

```sh
pip install -r requirements.txt
xlwings addin install
```

Then, open Excel and navigate to the sheet named `xlwings.conf`. Add the path of the folder where the Excel file and the script are located next to the `PYTHONPATH` entry.

Your `xlwings.conf` sheet should look something like this:

| Key        | Value                                |
|------------|--------------------------------------|
| PYTHONPATH | C:\path\to\your\folder               |

Make sure to replace `C:\path\to\your\folder` with the actual path to your folder.
