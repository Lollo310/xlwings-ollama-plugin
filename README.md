# xlwings-ollama-plugin

Excel plugin leveraging `xlwings` and the Ollama API to generate AI completions. Reads input text from a specified range and writes completions to adjacent cells. Supports model selection via a dedicated cell, enabling seamless integration of AI capabilities to test your prompts while tracking improvements in an Excel sheet.

## Prerequisites

Before using this plugin, ensure you have Python >= 3.9 installed.

```sh
python --version
```

If not, download and install them from the **Microsoft Store**.

Additionally, ensure you have the Ollama installed and running. To start the Ollama daemon, use the following command:

```sh
ollama serve
```

If not have Ollama, download and install from [Official Site](https://ollama.com/).

To pull the desired model, use the following command:

```sh
ollama pull <model-name>
```

Replace `<model-name>` with the name of the model you want to use.

## Installation and Setup for Windows

To install the required dependencies and set up the plugin, run the following commands:

```sh
pip install -r requirements.txt
```

**Important:** After the installation, if some warings is printed out or the command `xlwings` doesn't work, please add the python Scripts folder path to the PATH environment variable in Windows. This is the path to the `Scripts` folder in your Python installation. An output example:

If you don't know where the Python Scripts folder is, find the Python folder with this command:

```sh
python -c "import os, sys; print(os.path.dirname(sys.executable))"
```

Then look for a `Scripts` folder inside the Python installation directory.

You can follow this guide to help you: [How to Add to the PATH on Windows 10 and Windows 11](https://www.eukhost.com/kb/how-to-add-to-the-path-on-windows-10-and-windows-11/).

Execute the following command in the `ollamaplugin` folder to install the xlwings add-in:

```sh
xlwings addin install
```

Then, open Excel and navigate to the sheet named `xlwings.conf`. Add the path of the folder where the Excel file and the script are located next to the `PYTHONPATH` entry.

Your `xlwings.conf` sheet should look something like this:

| Key              | Value                                              |
|------------------|----------------------------------------------------|
| PYTHONPATH       | C:\path\to\your\folder\ollamaplugin                |
| Interpreter_Win  | python                                             |

Make sure to replace `C:\path\to\your\folder\ollamaplugin` with the actual path to your folder.

> **Note:** If the default value for `Interpreter_Win` does not work, replace it with the path to the `python.exe` file on your PC. For example:
>
> | Key              | Value                                              |
> |------------------|----------------------------------------------------|
> | Interpreter_Win  | C:\path\to\your\python\installation\python.exe     |
