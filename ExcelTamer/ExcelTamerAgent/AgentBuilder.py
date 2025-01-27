from concurrent.futures import ThreadPoolExecutor

from langchain.agents import initialize_agent, AgentType
from langchain_core.language_models import BaseChatModel

from ExcelTamer.ExcelAutomation import ExcelAutomation
from ExcelTamer.ExcelTamerAgent.ExcelTamerTools import ExcelGetStructureTool

executor = None

def create_agent(excel_path: str, llm:BaseChatModel) :
    # We use a global ThreadPoolExecutor to ensure all xlwings calls operate on the same thread.
    # xlwings relies on COM for Excel automation, and Excel typically operates under a
    # Single-Threaded Apartment (STA) model. COM objects in an STA environment must only
    # be accessed by the thread they were created on to avoid threading conflicts. Attempts
    # to use these objects from multiple threads can lead to unpredictable behavior, such as
    # crashes or deadlocks.
    #
    # For more details on STA threading and COM, see:
    # https://devblogs.microsoft.com/oldnewthing/20191125-00/?p=103135
    # https://docs.microsoft.com/en-us/windows/win32/com/using-the-threading

    global executor
    if executor is None:
        executor = ThreadPoolExecutor()
    excel:ExcelAutomation = ExcelAutomation(file_path=excel_path)
    return initialize_agent(
        tools=[ExcelGetStructureTool(excel_automation=excel, executor=executor)],
        llm=llm,
        agent=AgentType.OPENAI_FUNCTIONS,
        verbose=True
    )