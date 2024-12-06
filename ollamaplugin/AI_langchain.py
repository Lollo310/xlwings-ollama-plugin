from langchain_core.output_parsers import StrOutputParser
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.runnables import (
    RunnablePassthrough,
    RunnableSequence,
)
from langchain_ollama import ChatOllama


class aiLangChain:
    """RBACChatbot is an useful chatbot for working with company documents while respecting RBAC policies."""

    def __init__(
        self,
        model: str,
        temperature: float,
        system_prompt: str,
        output_parser=StrOutputParser(),
        num_ctx=2048,
    ):
        """
        Initializes the RBACChatbot with the given configuration.

        Args:
            model (str): The Ollama-hosted language model to use for chatting.
            temperature (float): The temperature setting for the language model.
            system_prompt (str): The system prompt to use for the chatbot.
            output_parser (StrOutputParser, optional): The output parser to use. Defaults to StrOutputParser().
            num_ctx (int, optional): The maximum number of tokens in the context. Defaults to 2048.

        Raises:
            ValueError: If the temperature is not between 0 and 1.
        """

        if temperature < 0 or temperature > 1:
            raise ValueError("Temperature must be between 0 and 1.")

        self.__model = model
        self.__temperature = temperature
        self.__system_prompt = system_prompt
        self.__output_parser = output_parser
        self.__num_ctx = num_ctx
        self.__chain = self.__init_chain()

    def __init_llm(self) -> ChatOllama:
        """
        Initializes the ollama-hosted language model for the chatbot if it has not been initialized yet.

        Returns:
            ChatOllama: The initialized language model.
        """

        return ChatOllama(
            model=self.__model,
            temperature=self.__temperature,
            num_ctx=self.__num_ctx,  # Max number of tokens in the context (Default: 2048)
        )

    def __create_prompt_template(self) -> ChatPromptTemplate:
        """
        Creates the prompt template for the chatbot.

        Returns:
            ChatPromptTemplate: The created prompt template.
        """
        system_prompt = self.__system_prompt

        return ChatPromptTemplate.from_messages(
            [("system", system_prompt), ("user", "{input}")]
        )

    def __init_chain(self) -> RunnableSequence:
        """
        Initializes the chain of runnable components using LCEL syntax.

        Returns:
            RunnableSequence: The initialized chain of runnable components.
        """
        prompt = self.__create_prompt_template()
        llm = self.__init_llm()

        return {"input": RunnablePassthrough()} | prompt | llm | self.__output_parser

    def invoke(self, input_text: str):
        """
        Invokes the chatbot with the given input text.

        Args:
            input_text (str): The input text from the user.

        Returns:
            Any: The response from the chatbot.
        """

        return self.__chain.invoke({"input": input_text})
