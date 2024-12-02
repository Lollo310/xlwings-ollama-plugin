import json

import requests
import xlwings as xw

OLLAMA_API_URL = "http://localhost:11434/api/generate"


def query_ollama(model, input_text):
    """Queries the Ollama API with the given model and input text.

    Args:
        model (str): The name of the model to query
        input_text (str): The input text to query the model with.

    Raises:
        raise Exception("Failed to query Ollama API")
        raise Exception("Failed to decode JSON response")

    Returns:
        str: The response from the Ollama API as a string.
    """
    payload = {"model": model, "prompt": input_text}
    response = requests.post(OLLAMA_API_URL, json=payload, stream=True)

    if response.status_code != 200:
        raise Exception(f"Error: {response.text}")

    full_response = ""

    for line in response.iter_lines():
        if line:
            try:
                chunk = json.loads(line.decode("utf-8"))

                full_response += chunk.get("response", "")

                if chunk.get("done", False):
                    break
            except json.JSONDecodeError as e:
                raise Exception(f"Error: {e}")

    return full_response.strip()


def main():
    """Fetches model name from the sheet and queries Ollama API for each cell in range A2:A{last_row}."""
    wb = xw.Book.caller()
    sheet = wb.sheets.active
    # Read the model from cell F1
    model = sheet["F1"].value
    last_row = sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row
    for cell in sheet.range(f"A2:A{last_row}"):
        if not cell.value:
            break

        if not cell.offset(0, 1).value:
            completion = query_ollama(model, cell.value)
            cell.offset(0, 1).value = completion


if __name__ == "__main__":
    # Set the mock caller for debugging purposes
    xw.Book("ollamaplugin.xlsm").set_mock_caller()
    main()
