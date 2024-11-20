import json

import requests
import xlwings as xw

OLLAMA_API_URL = "http://localhost:11434/api/generate"


def query_ollama(model, input_text):
    """Function to query Ollama using its CLI."""
    payload = {"model": model, "prompt": input_text}
    response = requests.post(OLLAMA_API_URL, json=payload, stream=True)

    if response.status_code != 200:
        raise f"Error: {response.text}"

    full_response = ""

    for line in response.iter_lines():
        if line:
            try:
                chunk = json.loads(line.decode("utf-8"))

                full_response += chunk.get("response", "")

                if chunk.get("done", False):
                    break
            except json.JSONDecodeError as e:
                raise f"Error: {e}"

    return full_response.strip()


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets.active
    model = sheet["F1"].value

    for cell in sheet.range("A2:A100"):
        if not cell.value:
            break

        if not cell.offset(0, 1).value:
            completion = query_ollama(model, cell.value)
            cell.offset(0, 1).value = completion


if __name__ == "__main__":
    xw.Book("ollamaplugin.xlsm").set_mock_caller()
    main()
