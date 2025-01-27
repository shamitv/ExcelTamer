from concurrent.futures import ThreadPoolExecutor

from langchain.agents import initialize_agent, AgentType
from langchain_core.language_models import BaseChatModel

from ExcelTamer.ExcelAutomation import ExcelAutomation
from ExcelTamer.ExcelTamerAgent.ExcelTamerTools import ExcelGetStructureTool

executor = None

def create_agent(excel_path: str, llm:BaseChatModel) :
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