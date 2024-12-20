import xlwings as xw
from AI_langchain import aiLangChain


def query_ollama(model, temperature, system_prompt, num_ctx, input_text):
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
    chatbot = aiLangChain(model, temperature, system_prompt, num_ctx=num_ctx)

    return chatbot.invoke(input_text)


def main():
    """Main function to be called by the Excel plugin."""
    wb = xw.Book.caller()

    sheet_conf = wb.sheets[
        "SpikeTestConfing"
    ]  # This is the first sheet in the excel where the user will input the configuration
    sheet_prompts = wb.sheets[
        "SpikeTestPrompts"
    ]  # This is the second sheet in the excel where the user will input the prompts
    sheet_output = wb.sheets[
        "SpikeTestOutput"
    ]  # This is the third sheet in the excel where the output will be displayed

    # Read the model from cell B1
    model = sheet_conf["B1"].value
    # Read the temperature from cell B2
    temperature = float(sheet_conf["B2"].value)
    # Read the context length from cell B3
    num_ctx = int(sheet_conf["B3"].value)
    system_prompt = ""

    last_row = (
        sheet_prompts.range("B" + str(sheet_prompts.cells.last_cell.row)).end("up").row
    )
    for i, cell in enumerate(sheet_prompts.range(f"B2:B{last_row}"), 2):
        if not cell.value:
            break

        if sheet_prompts["A" + str(i)].value:
            system_prompt = sheet_prompts["A" + str(i)].value

        if not sheet_output["A" + str(i)].value:
            completion = query_ollama(
                model, temperature, system_prompt, num_ctx, cell.value
            )
            sheet_output["A" + str(i)].value = completion


if __name__ == "__main__":
    # Set the mock caller for debugging purposes
    xw.Book("ollamaplugin.xlsm").set_mock_caller()
    main()
