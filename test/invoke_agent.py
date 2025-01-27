from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

from langchain_openai import ChatOpenAI

from ExcelTamer.ExcelTamerAgent.AgentBuilder import create_agent

excel_path = "example.xlsx"

llm = ChatOpenAI(model_name="gpt-4o-mini", temperature=0)

agent = create_agent(excel_path, llm)

response = agent.invoke("Is there a costs sheet, Does this sheet contain time period labels across columns ?")

print(response)

response = agent.invoke("What is the in cell B12 of Expenses?")

print(response)

