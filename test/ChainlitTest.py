import logging

# Set the logging level to DEBUG
logging.basicConfig(level=logging.DEBUG,
    format='%(asctime)s - %(message)s',
    datefmt='%d:%m:%Y %H:%M:%S')

from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

import chainlit as cl

from langchain_openai import ChatOpenAI

from ExcelTamer.ExcelTamerAgent.AgentBuilder import create_agent

llm = ChatOpenAI(model_name="gpt-4o-mini", temperature=0)

#Do not open any specific Excel File, Work with currently open Excel Workbook
excel_path = None
agent = create_agent(excel_path, llm)

 v

@cl.on_chat_start
async def start():
    cl.user_session.set("agent", agent)
    await cl.Message(content="Hello! I am your AI assistant. How can I help you today?").send()

@cl.on_message
async def handle_message(message):
    agent_inst = cl.user_session.get("agent")
    res = await agent_inst.arun(
        message.content, callbacks=[cl.AsyncLangchainCallbackHandler()]
    )
    await cl.Message(content=res).send()

if __name__ == "__main__":
    from chainlit.cli import run_chainlit
    run_chainlit(__file__)